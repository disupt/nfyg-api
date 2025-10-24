"""```
Total_Agatston_Pattern_Extractor.py
====================================

두 번째 모니터에서 Calcium Score Report 창의 두 가지 패턴(레이아웃)을 인식하고,
모니터 전체 화면을 OCR 한 뒤 Total-Agatston Score 값을 추출하는 자동화 스크립트입니다.

기능 개요
---------
1. 사용자 정의 단축키 대기: 지정된 단축키(예: ``F8``, ``Ctrl+Shift+S``) 입력 시 패턴 인식과
   점수 추출 프로세스를 시작합니다.
2. 패턴 인식: 두 번째 모니터 전체 스크린샷을 확보하고, Anchor 포인트 색상값 검증을 통해
   현재 Report 창이 패턴 1인지 패턴 2인지 판별합니다.
3. Score 추출: 식별된 패턴과 무관하게 두 번째 모니터 전체 이미지를 OCR(Tesseract)으로 분석하고
   "Total" 행과 "Agatston" 열 교차점의 숫자 값을 추출합니다.

필수 준비 사항
---------------
* Python 3.9 이상 권장
* ``pip install pyautogui pillow pytesseract pynput``
  - macOS 보안설정에서 PyAutoGUI 및 Pynput이 접근성 권한을 갖도록 허용해야 합니다.
* Tesseract OCR 엔진 설치 (예: macOS ``brew install tesseract``)

구동 예시
---------
```
python Total_Agatston_Pattern_Extractor.py \
    --hotkey "F8" \
    --monitor-region 1920 0 3840 1080 \
    --pattern-file pattern1.json \
    --pattern-file pattern2.json
```

``pattern*.json`` 파일 예시는 다음과 같습니다.

```
{
  "name": "pattern1",
  "priority": 0,
  "anchors": [
    {
      "x": 2200,
      "y": 220,
      "color": [26, 73, 128],
      "tolerance": 25,
      "description": "제목 바 파란색"
    }
  ]
}
```

``monitor_region`` 은 두 번째 모니터의 절대 좌표(left, top, right, bottom)를 의미합니다. macOS 기준으로
메인 모니터 왼쪽에 배치된 보조 모니터는 음수 좌표를 갖습니다. Windows 환경에서도 동일하게 좌표를 지정하면 되며,
Anchor 좌표 역시 전체 화면 기준 절대 좌표로 입력합니다. ``tolerance`` 는 RGB 각 채널의 허용 편차 값입니다.

주의
----
* Retina(HiDPI) 모니터 사용 시 PyAutoGUI의 좌표가 실제 픽셀과 다를 수 있습니다. 이 경우
  ``pyautogui.useImageNotFoundException()`` 을 활용하거나 ``pyautogui.size()`` 로 보정해야 합니다.
* 다중 캡처를 반복하려면 단축키를 여러 번 눌러도 됩니다. 처리 중 다시 단축키를 누르면
  이전 작업이 끝날 때까지 대기합니다.

```
"""

from __future__ import annotations

import argparse
import json
import logging
import os
import threading
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterable, List, Optional, Sequence, Tuple

try:
    import pyautogui
except ImportError as exc:  # pragma: no cover
    raise SystemExit(
        "pyautogui 모듈을 찾을 수 없습니다. `pip install pyautogui`로 설치하세요."
    ) from exc

try:
    from PIL import Image
except ImportError as exc:  # pragma: no cover
    raise SystemExit(
        "Pillow(PIL) 모듈을 찾을 수 없습니다. `pip install pillow`로 설치하세요."
    ) from exc

try:
    from pynput import keyboard
except ImportError as exc:  # pragma: no cover
    raise SystemExit(
        "pynput 모듈을 찾을 수 없습니다. `pip install pynput`으로 설치하세요."
    ) from exc

from Total_Agatston_Extractor import (  # type: ignore
    configure_tesseract,
    extract_total_agatston_score,
    extract_words,
    show_result,
)


LOGGER = logging.getLogger("total_agatston_pattern_extractor")


@dataclass
class AnchorDefinition:
    """패턴 매칭을 위한 Anchor 포인트 정의."""

    x: int
    y: int
    color: Tuple[int, int, int]
    tolerance: int = 20
    description: Optional[str] = None

    @staticmethod
    def from_dict(data: dict) -> "AnchorDefinition":
        try:
            color = tuple(int(c) for c in data["color"])
        except Exception as exc:  # pragma: no cover
            raise ValueError("anchor color 값은 길이 3의 정수 배열이어야 합니다.") from exc

        if len(color) != 3:
            raise ValueError("anchor color 는 [R, G, B] 3요소여야 합니다.")

        return AnchorDefinition(
            x=int(data["x"]),
            y=int(data["y"]),
            color=(int(color[0]), int(color[1]), int(color[2])),
            tolerance=int(data.get("tolerance", 20)),
            description=data.get("description"),
        )


@dataclass
class PatternDefinition:
    """Report 패턴 정의."""

    name: str
    anchors: List[AnchorDefinition] = field(default_factory=list)
    priority: int = 0

    @staticmethod
    def from_dict(data: dict) -> "PatternDefinition":
        anchors = [AnchorDefinition.from_dict(item) for item in data.get("anchors", [])]

        return PatternDefinition(
            name=str(data.get("name", "Unnamed Pattern")),
            anchors=anchors,
            priority=int(data.get("priority", 0)),
        )

    def matches(self, image: Image.Image, origin: Tuple[int, int]) -> bool:
        if not self.anchors:
            LOGGER.debug("패턴 %s 는 Anchor 가 없어 무조건 일치로 판단합니다.", self.name)
            return True

        width, height = image.size
        offset_x, offset_y = origin

        for anchor in self.anchors:
            rel_x = anchor.x - offset_x
            rel_y = anchor.y - offset_y

            if not (0 <= rel_x < width and 0 <= rel_y < height):
                LOGGER.debug(
                    "패턴 %s Anchor(%s) 가 모니터 영역 밖입니다. (rel_x=%s, rel_y=%s)",
                    self.name,
                    anchor.description or f"{anchor.x},{anchor.y}",
                    rel_x,
                    rel_y,
                )
                return False

            pixel = image.getpixel((rel_x, rel_y))
            if not _color_within_tolerance(pixel, anchor.color, anchor.tolerance):
                LOGGER.debug(
                    "패턴 %s Anchor(%s) 색상 불일치: 실제=%s, 기대=%s, tol=%s",
                    self.name,
                    anchor.description or f"{anchor.x},{anchor.y}",
                    pixel,
                    anchor.color,
                    anchor.tolerance,
                )
                return False

        LOGGER.debug("패턴 %s Anchor 검증 통과", self.name)
        return True


@dataclass
class ExtractorConfig:
    hotkey: str = "F8"
    monitor_region: Optional[Tuple[int, int, int, int]] = None
    patterns: List[PatternDefinition] = field(default_factory=list)

    @staticmethod
    def from_sources(
        *,
        config_path: Optional[Path],
        pattern_files: Sequence[Path],
        hotkey_override: Optional[str],
        monitor_region_override: Optional[Sequence[int]],
    ) -> "ExtractorConfig":
        data: dict = {}

        if config_path:
            if not config_path.exists():
                raise SystemExit(f"지정한 설정 파일을 찾을 수 없습니다: {config_path}")
            LOGGER.info("설정 파일 %s 을 로드합니다.", config_path)
            data = json.loads(config_path.read_text(encoding="utf-8"))
        else:
            default_cfg = Path(__file__).with_name("Total_Agatston_Pattern_Config.json")
            if default_cfg.exists():
                LOGGER.info("기본 설정 파일 %s 을 로드합니다.", default_cfg)
                data = json.loads(default_cfg.read_text(encoding="utf-8"))

        patterns: List[PatternDefinition] = []

        for item in data.get("patterns", []):
            patterns.append(PatternDefinition.from_dict(item))

        for pattern_file in pattern_files:
            if not pattern_file.exists():
                raise SystemExit(f"패턴 파일을 찾을 수 없습니다: {pattern_file}")
            LOGGER.info("패턴 정의 %s 을 로드합니다.", pattern_file)
            pattern_data = json.loads(pattern_file.read_text(encoding="utf-8"))
            patterns.append(PatternDefinition.from_dict(pattern_data))

        if not patterns:
            raise SystemExit("패턴 정의가 없습니다. --pattern-file 옵션을 사용하거나 설정 파일에 patterns 배열을 작성하세요.")

        hotkey = hotkey_override or data.get("hotkey") or "F8"

        monitor_region: Optional[Tuple[int, int, int, int]] = None
        if monitor_region_override:
            if len(monitor_region_override) != 4:
                raise SystemExit("--monitor-region 은 x1 y1 x2 y2 4개의 정수를 입력해야 합니다.")
            monitor_region = tuple(int(v) for v in monitor_region_override)
        elif "monitor_region" in data:
            region_seq = data["monitor_region"]
            if len(region_seq) != 4:
                raise SystemExit("설정 파일 monitor_region 은 [x1, y1, x2, y2] 형식이어야 합니다.")
            monitor_region = tuple(int(v) for v in region_seq)

        patterns.sort(key=lambda p: (p.priority, p.name))

        return ExtractorConfig(
            hotkey=hotkey,
            monitor_region=monitor_region,
            patterns=patterns,
        )


def parse_args(argv: Optional[Sequence[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Calcium Score Report 패턴을 인식하고 Total-Agatston Score 를 추출합니다.",
    )
    parser.add_argument(
        "--hotkey",
        type=str,
        default=None,
        help="단축키 문자열 (예: F8, Ctrl+Shift+S)",
    )
    parser.add_argument(
        "--monitor-region",
        type=int,
        nargs=4,
        metavar=("X1", "Y1", "X2", "Y2"),
        help="두 번째 모니터의 영역 (절대 좌표)",
    )
    parser.add_argument(
        "--pattern-config",
        type=Path,
        default=None,
        help="패턴/모니터 정보를 담은 JSON 설정 파일 경로",
    )
    parser.add_argument(
        "--pattern-file",
        type=Path,
        action="append",
        default=[],
        help="개별 패턴 정의 JSON 파일 (여러 번 지정 가능)",
    )
    parser.add_argument(
        "--tesseract-path",
        type=str,
        default=None,
        help="Tesseract 실행 파일 경로 (PATH 미등록 시)",
    )
    parser.add_argument(
        "--no-popup",
        action="store_true",
        help="결과 팝업 표시 생략",
    )
    parser.add_argument(
        "--debug",
        action="store_true",
        help="디버그 로그 출력",
    )
    parser.add_argument(
        "--debug-dir",
        type=Path,
        default=None,
        help="스크린샷/크롭 이미지를 저장할 디렉터리",
    )
    parser.add_argument(
        "--run-once",
        action="store_true",
        help="프로그램 시작 즉시 한 번만 실행하고 종료",
    )
    return parser.parse_args(argv)


def configure_logging(debug: bool) -> None:
    logging.basicConfig(
        level=logging.DEBUG if debug else logging.INFO,
        format="[%(asctime)s] %(levelname)s: %(message)s",
    )


def normalize_hotkey(hotkey: str) -> str:
    tokens = [token.strip().lower() for token in hotkey.split("+") if token.strip()]
    if not tokens:
        raise ValueError("단축키 문자열을 해석할 수 없습니다.")

    key_map = {
        "ctrl": "<ctrl>",
        "control": "<ctrl>",
        "shift": "<shift>",
        "alt": "<alt>",
        "option": "<alt>",
        "command": "<cmd>",
        "cmd": "<cmd>",
        "win": "<cmd>",
    }

    normalized: List[str] = []
    for token in tokens:
        if token in key_map:
            normalized.append(key_map[token])
        elif token.startswith("f") and token[1:].isdigit():
            normalized.append(f"<{token}>")
        elif len(token) == 1:
            normalized.append(token)
        else:
            normalized.append(token)

    return "+".join(normalized)


def _color_within_tolerance(
    actual: Tuple[int, int, int],
    expected: Tuple[int, int, int],
    tolerance: int,
) -> bool:
    return max(abs(int(a) - int(b)) for a, b in zip(actual, expected)) <= tolerance


def _ensure_directory(path: Optional[Path]) -> Optional[Path]:
    if path is None:
        return None
    path.mkdir(parents=True, exist_ok=True)
    return path


class PatternExtractor:
    def __init__(
        self,
        config: ExtractorConfig,
        *,
        use_popup: bool,
        tesseract_path: Optional[str],
        debug_dir: Optional[Path],
    ) -> None:
        self.config = config
        self.use_popup = use_popup
        self.debug_dir = _ensure_directory(debug_dir)
        self._lock = threading.Lock()
        self._last_capture_at: Optional[float] = None
        configure_tesseract(tesseract_path)

    def capture_monitor(self) -> Tuple[Image.Image, Tuple[int, int]]:
        region = self.config.monitor_region
        if region:
            x1, y1, x2, y2 = region
            if x2 <= x1 or y2 <= y1:
                raise ValueError("monitor_region 값이 올바르지 않습니다. (x2 > x1, y2 > y1)")
            width = x2 - x1
            height = y2 - y1
            LOGGER.info("모니터 영역 캡처: (%s,%s) -> (%s,%s)", x1, y1, x2, y2)
            screenshot = pyautogui.screenshot(region=(x1, y1, width, height))
            origin = (x1, y1)
        else:
            LOGGER.info("모니터 영역이 지정되지 않아 전체 화면을 캡처합니다.")
            screenshot = pyautogui.screenshot()
            origin = (0, 0)

        if self.debug_dir:
            timestamp = int(time.time() * 1000)
            screen_path = self.debug_dir / f"screen_{timestamp}.png"
            screenshot.save(screen_path)
            LOGGER.debug("스크린샷을 %s 에 저장했습니다.", screen_path)

        return screenshot, origin

    def detect_pattern(self, screenshot: Image.Image, origin: Tuple[int, int]) -> PatternDefinition:
        LOGGER.info("%d개의 패턴 중 일치 여부를 확인합니다.", len(self.config.patterns))
        for pattern in self.config.patterns:
            if pattern.matches(screenshot, origin):
                LOGGER.info("패턴 '%s' 일치", pattern.name)
                return pattern
        raise ValueError("어느 패턴과도 일치하지 않습니다. Anchor 좌표/색상을 재점검하세요.")

    def process_once(self) -> None:
        if not self._lock.acquire(blocking=False):
            LOGGER.warning("이미 다른 추출 작업이 진행 중입니다. 잠시 후 다시 시도하세요.")
            return

        try:
            self._last_capture_at = time.time()
            screenshot, origin = self.capture_monitor()
            pattern = self.detect_pattern(screenshot, origin)

            words = extract_words(screenshot)
            score = extract_total_agatston_score(words)
            LOGGER.info("패턴 '%s' 일치. 모니터 전체에서 Total Agatston Score=%s 를 추출했습니다.", pattern.name, score)
            show_result(score, self.use_popup)
        except Exception as exc:
            LOGGER.error("Total Agatston Score 추출 실패: %s", exc)
        finally:
            self._lock.release()

    def run_with_hotkey(self) -> None:
        hotkey_display = self.config.hotkey
        try:
            normalized = normalize_hotkey(self.config.hotkey)
        except ValueError as exc:
            raise SystemExit(f"단축키 파싱 실패: {exc}") from exc

        LOGGER.info("단축키 대기 중... (입력: %s / 내부: %s)", hotkey_display, normalized)
        print(f"단축키 대기 중... (입력: {hotkey_display})")

        def on_activate() -> None:
            LOGGER.info("단축키 '%s' 입력 감지", hotkey_display)
            threading.Thread(target=self.process_once, daemon=True).start()

        with keyboard.GlobalHotKeys({normalized: on_activate}) as listener:
            listener.join()


def main(argv: Optional[Sequence[str]] = None) -> int:
    args = parse_args(argv)
    configure_logging(args.debug)

    try:
        config = ExtractorConfig.from_sources(
            config_path=args.pattern_config,
            pattern_files=args.pattern_file,
            hotkey_override=args.hotkey,
            monitor_region_override=args.monitor_region,
        )
    except Exception as exc:
        LOGGER.error("설정 로드 중 오류: %s", exc)
        return 1

    extractor = PatternExtractor(
        config,
        use_popup=not args.no_popup,
        tesseract_path=args.tesseract_path,
        debug_dir=args.debug_dir,
    )

    if args.run_once:
        LOGGER.info("run_once 옵션이 활성화되어 즉시 한 번 실행합니다.")
        extractor.process_once()
        return 0

    try:
        extractor.run_with_hotkey()
    except KeyboardInterrupt:  # pragma: no cover
        LOGGER.info("사용자 인터럽트로 종료합니다.")
        return 130
    except Exception as exc:
        LOGGER.error("핫키 대기 중 오류: %s", exc)
        return 1

    return 0


if __name__ == "__main__":
    raise SystemExit(main())


