"""Total Agatston Score 추출 스크립트.

두 번째 모니터에 띄워둔 Calcium Score Report의 특정 영역을 자동으로 캡처하고,
Tesseract OCR을 통해 "Total" 행과 "Agatston" 열이 만나는 값을 찾아 출력한다.

준비 사항
-----------
1. Python 3.9 이상 권장
2. `pip install pyautogui pillow pytesseract`
   - macOS에서 `pyautogui` 사용 시 시스템 환경설정 > 보안 및 개인 정보 보호에서
     접근성 권한을 허용해야 한다.
3. OS에 Tesseract OCR 엔진 설치
   - macOS(Homebrew): `brew install tesseract`
   - Windows: https://github.com/UB-Mannheim/tesseract/wiki 에서 설치 후 경로 지정

사용 예시
---------
python Total_Agatston_Extractor.py --region 1925 120 2760 970 --screenshot-path report.png

옵션 설명
---------
- --region: 캡처할 영역 좌표 (왼쪽 위 x1 y1, 오른쪽 아래 x2 y2)
- --screenshot-path: 캡처 이미지를 저장할 파일 경로 (미지정 시 메모리만 사용)
- --tesseract-path: PATH에 등록되지 않은 경우 Tesseract 실행 파일 경로
- --no-popup: 결과 팝업(pyautogui.alert) 표시 생략
- --debug: 상세 로그 출력
"""

from __future__ import annotations

import argparse
import logging
import os
import re
import sys
from dataclasses import dataclass
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
    import pytesseract
    from pytesseract import Output
except ImportError as exc:  # pragma: no cover
    raise SystemExit(
        "pytesseract 모듈을 찾을 수 없습니다. `pip install pytesseract`로 설치하세요."
    ) from exc


# 기본 캡처 영역 (필요 시 직접 수정)
DEFAULT_REPORT_REGION: Tuple[int, int, int, int] = (3195, 1830, 3237, 1846)

WINDOWS_TESSERACT_PATHS: Tuple[str, ...] = (
    r"C:\\Program Files\\Tesseract-OCR\\tesseract.exe",
    r"C:\\Program Files (x86)\\Tesseract-OCR\\tesseract.exe",
)

LOGGER = logging.getLogger("total_agatston_extractor")


@dataclass
class OCRWord:
    """OCR 결과의 개별 단어 정보."""

    text: str
    left: int
    top: int
    width: int
    height: int
    conf: float
    line_num: int
    block_num: int
    page_num: int
    par_num: int

    @property
    def center_x(self) -> float:
        return self.left + self.width / 2

    @property
    def center_y(self) -> float:
        return self.top + self.height / 2


def parse_args(argv: Optional[Sequence[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Calcium Score Report의 Total-Agatston 값을 OCR로 추출합니다.",
    )
    parser.add_argument(
        "--region",
        metavar=("X1", "Y1", "X2", "Y2"),
        type=int,
        nargs=4,
        help="보고서 영역의 좌표 (왼쪽 위 x1 y1, 오른쪽 아래 x2 y2)",
    )
    parser.add_argument(
        "--screenshot-path",
        type=str,
        default=None,
        help="캡처 이미지를 저장할 파일 경로",
    )
    parser.add_argument(
        "--tesseract-path",
        type=str,
        default=None,
        help="Tesseract 실행 파일 경로(필요 시)",
    )
    parser.add_argument(
        "--no-popup",
        action="store_true",
        help="결과 팝업(pyautogui.alert) 표시를 생략",
    )
    parser.add_argument(
        "--debug",
        action="store_true",
        help="디버그 로그 출력",
    )
    return parser.parse_args(argv)


def configure_logging(debug: bool) -> None:
    logging.basicConfig(
        level=logging.DEBUG if debug else logging.INFO,
        format="[%(asctime)s] %(levelname)s: %(message)s",
    )


def ensure_region(region: Optional[Sequence[int]]) -> Tuple[int, int, int, int]:
    if region is not None:
        if len(region) != 4:
            raise ValueError("--region 인자는 x1 y1 x2 y2 네 개의 정수가 필요합니다.")
        return tuple(region)  # type: ignore[return-value]

    if any(DEFAULT_REPORT_REGION):
        LOGGER.info("--region 인자가 없어 기본 좌표 %s를 사용합니다.", DEFAULT_REPORT_REGION)
        return DEFAULT_REPORT_REGION

    raise SystemExit(
        "캡처 좌표를 알 수 없습니다. --region 인자를 지정하거나 "
        "DEFAULT_REPORT_REGION을 직접 설정하세요."
    )


def configure_tesseract(path: Optional[str]) -> None:
    if path:
        expanded = os.path.expandvars(path)
        if not os.path.exists(expanded):
            raise SystemExit(f"지정한 Tesseract 경로가 존재하지 않습니다: {expanded}")
        pytesseract.pytesseract.tesseract_cmd = expanded
        LOGGER.debug("Tesseract 경로를 %s 로 설정했습니다.", expanded)
        return

    if os.name == "nt":
        for candidate in WINDOWS_TESSERACT_PATHS:
            if os.path.exists(candidate):
                pytesseract.pytesseract.tesseract_cmd = candidate
                LOGGER.debug("Windows 기본 경로에서 Tesseract를 발견했습니다: %s", candidate)
                return

        LOGGER.warning(
            "Windows 환경에서 Tesseract 경로를 찾지 못했습니다. --tesseract-path 옵션으로 직접 지정하세요."
        )


def capture_region(
    region: Tuple[int, int, int, int],
    screenshot_path: Optional[str] = None,
) -> Image.Image:
    x1, y1, x2, y2 = region
    if x2 <= x1 or y2 <= y1:
        raise ValueError("유효하지 않은 좌표입니다. (x2 > x1, y2 > y1)")

    width = x2 - x1
    height = y2 - y1
    LOGGER.info(
        "영역 캡처: x1=%s, y1=%s, x2=%s, y2=%s (width=%s, height=%s)",
        x1,
        y1,
        x2,
        y2,
        width,
        height,
    )
    image = pyautogui.screenshot(region=(x1, y1, width, height))
    if screenshot_path:
        image.save(screenshot_path)
        LOGGER.info("캡처 이미지를 %s 에 저장했습니다.", screenshot_path)
    return image


def extract_words(image: Image.Image) -> List[OCRWord]:
    LOGGER.debug("OCR 데이터를 추출합니다.")
    data = pytesseract.image_to_data(image, output_type=Output.DICT)
    words: List[OCRWord] = []

    for idx, raw in enumerate(data.get("text", [])):
        text = raw.strip()
        if not text:
            continue

        try:
            conf = float(data.get("conf", ["-1"])[idx])
        except (ValueError, TypeError):
            conf = -1.0

        words.append(
            OCRWord(
                text=text,
                left=int(data.get("left", [0])[idx]),
                top=int(data.get("top", [0])[idx]),
                width=int(data.get("width", [0])[idx]),
                height=int(data.get("height", [0])[idx]),
                conf=conf,
                line_num=int(data.get("line_num", [0])[idx]),
                block_num=int(data.get("block_num", [0])[idx]),
                page_num=int(data.get("page_num", [0])[idx]),
                par_num=int(data.get("par_num", [0])[idx]),
            )
        )

    if not words:
        raise ValueError("OCR 결과에서 텍스트를 찾지 못했습니다.")

    LOGGER.debug("총 %d개의 단어를 인식했습니다.", len(words))
    return words


def _normalize(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip().lower()


def find_total_words(words: Iterable[OCRWord]) -> List[OCRWord]:
    total_words = [w for w in words if _normalize(w.text) == "total"]
    if not total_words:
        raise ValueError("'Total' 단어를 찾지 못했습니다.")
    LOGGER.debug("Total 단어 %d개 발견", len(total_words))
    return total_words


def find_agatston_column_center(words: Iterable[OCRWord]) -> float:
    agatston_words = [w for w in words if _normalize(w.text).startswith("agatston")]
    if not agatston_words:
        raise ValueError("'Agatston' 열을 찾지 못했습니다.")
    center = sum(w.center_x for w in agatston_words) / len(agatston_words)
    LOGGER.debug("Agatston 열 중심 x 좌표: %.2f", center)
    return center


def extract_candidates(
    words: Sequence[OCRWord],
    total_words: Sequence[OCRWord],
    agatston_center: float,
) -> List[Tuple[float, int]]:
    numeric_pattern = re.compile(r"[-+]?\d[\d,\.]*")
    candidates: List[Tuple[float, int]] = []

    for total_word in total_words:
        same_line = [
            w
            for w in words
            if w.line_num == total_word.line_num and w.block_num == total_word.block_num
        ]

        LOGGER.debug(
            "Total(word line=%d, block=%d)과 같은 줄 후보 %d개",
            total_word.line_num,
            total_word.block_num,
            len(same_line),
        )

        closest_value: Optional[int] = None
        closest_distance: Optional[float] = None

        for w in same_line:
            if w is total_word:
                continue

            match = numeric_pattern.search(w.text)
            if not match:
                continue

            digits = re.sub(r"[^0-9]", "", match.group())
            if not digits:
                continue

            value = int(digits)
            distance = abs(w.center_x - agatston_center)

            if closest_distance is None or distance < closest_distance:
                closest_value = value
                closest_distance = distance

        if closest_value is not None and closest_distance is not None:
            candidates.append((closest_distance, closest_value))

    return candidates


def fallback_search(
    words: Sequence[OCRWord],
    agatston_center: float,
    y_threshold: int = 30,
) -> Optional[int]:
    numeric_pattern = re.compile(r"[-+]?\d[\d,\.]*")
    best: Optional[Tuple[float, int]] = None

    for w in words:
        match = numeric_pattern.search(w.text)
        if not match:
            continue

        digits = re.sub(r"[^0-9]", "", match.group())
        if not digits:
            continue

        distance_x = abs(w.center_x - agatston_center)
        distance_y = abs(w.center_y)
        if distance_y > y_threshold:
            continue

        value = int(digits)
        candidate = (distance_x, value)
        if best is None or candidate[0] < best[0]:
            best = candidate

    if best:
        LOGGER.debug("폴백 후보 발견: %s", best)
        return best[1]
    return None


def extract_total_agatston_score(words: Sequence[OCRWord]) -> int:
    total_words = find_total_words(words)
    agatston_center = find_agatston_column_center(words)
    candidates = extract_candidates(words, total_words, agatston_center)

    if candidates:
        candidates.sort(key=lambda item: item[0])
        best_score = candidates[0][1]
        LOGGER.info("Total-Agatston 교차점에서 %d 값을 추출했습니다.", best_score)
        return best_score

    LOGGER.warning("동일 행에서 숫자를 찾지 못해 폴백 검색을 시도합니다.")
    fallback_value = fallback_search(words, agatston_center)
    if fallback_value is not None:
        LOGGER.info("폴백 검색 결과 %d 값을 추출했습니다.", fallback_value)
        return fallback_value

    raise ValueError("Total 행과 Agatston 열의 교차점에서 숫자를 찾지 못했습니다.")


def show_result(score: int, use_popup: bool) -> None:
    message = f"추출된 Score: {score}"
    print(message)
    if use_popup:
        try:
            pyautogui.alert(message)
        except Exception as exc:  # pragma: no cover
            LOGGER.warning("팝업 표시 중 오류가 발생했습니다: %s", exc)


def main(argv: Optional[Sequence[str]] = None) -> int:
    args = parse_args(argv)
    configure_logging(args.debug)

    try:
        region = ensure_region(args.region)
        configure_tesseract(args.tesseract_path)
        image = capture_region(region, args.screenshot_path)
        words = extract_words(image)
        score = extract_total_agatston_score(words)
        show_result(score, not args.no_popup)
        return 0
    except KeyboardInterrupt:  # pragma: no cover
        LOGGER.info("사용자 요청으로 작업을 중단했습니다.")
        return 130
    except Exception as exc:
        LOGGER.error("Agatston Score 추출 중 오류: %s", exc)
        return 1


if __name__ == "__main__":
    sys.exit(main())

