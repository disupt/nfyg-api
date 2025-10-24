from pynput import keyboard
from mss import mss
import pytesseract
from PIL import Image
import pyautogui
try:
    import pyperclip
except ImportError:
    pyperclip = None
import re
import platform
import time

# ----------------- 1. 사용자 설정 영역 (필수) -----------------
# **Tesseract OCR 엔진 경로** (예: r'C:/Program Files/Tesseract-OCR/tesseract.exe')
TESSERACT_PATH = r'C:/Program Files/Tesseract-OCR/tesseract.exe'
RESULT_INPUT_COORD = (970, 886)  # 점수 및 등급을 입력할 대상 좌표
INPUT_CLICK_DELAY = 0.35
INPUT_TYPING_INTERVAL = 0.02

# **1번 패턴 (CT 기기 1)의 점수 영역 좌표**: '1. 화면 좌표 추출기'로 찾은 값
ROI_1 = {
    'monitor': 3,
    'region': {'top': 1967, 'left': 2405, 'width': 94, 'height': 29},
}
# **2번 패턴 (CT 기기 2)의 점수 영역 좌표**: '1. 화면 좌표 추출기'로 찾은 값
ROI_2 = {
    'monitor': 3,
    'region': {'top': 1828, 'left': 3161, 'width': 109, 'height': 24},
}
# -------------------------------------------------------------


class AgatstonScoreMaster:
    def __init__(self, tesseract_path, rois):
        # Tesseract 경로 설정 (사전 준비 필수)
        pytesseract.pytesseract.tesseract_cmd = tesseract_path
        self.rois = rois

    def _prepare_roi(self, roi, monitors):
        """모니터 기준 ROI 정보를 실제 좌표계로 변환합니다."""
        if 'monitor' not in roi:
            return roi

        monitor_idx = roi['monitor']
        if monitor_idx >= len(monitors):
            raise ValueError(
                f"지정된 모니터 번호 {monitor_idx}가 연결된 모니터 수({len(monitors) - 1})를 초과합니다."
            )

        base = monitors[monitor_idx]
        region = roi['region']
        return {
            'top': base['top'] + region['top'],
            'left': base['left'] + region['left'],
            'width': region['width'],
            'height': region['height'],
        }

    def _extract_from_roi(self, roi):
        """단일 ROI에서 스크린샷, 전처리, OCR을 수행하여 점수를 추출합니다."""
        with mss() as sct:
            monitors = sct.monitors
            region = self._prepare_roi(roi, monitors)
            sct_img = sct.grab(region)
        img = Image.frombytes("RGB", sct_img.size, sct_img.bgra, "raw", "BGRX")

        # 이미지 전처리: 그레이스케일 변환 (OCR 정확도 향상)
        processed_img = img.convert('L')

        # OCR 실행: 숫자만 인식하도록 설정 (--psm 7: 단일 텍스트 라인)
        custom_config = r'--oem 3 --psm 7 -c tessedit_char_whitelist=0123456789.'
        text = pytesseract.image_to_string(processed_img, config=custom_config).strip()

        # 추출된 텍스트에서 숫자만 정리하고 정수로 반환
        numbers = re.findall(r'\d+(?:\.\d+)?', text)
        if not numbers:
            return None

        score_value = float(numbers[0])
        if score_value.is_integer():
            return int(score_value)
        return score_value

    def _input_result_to_target(self, score_value, classification_text):
        """지정된 좌표에 점수와 등급을 자동으로 입력합니다."""
        if not RESULT_INPUT_COORD:
            return

        x, y = RESULT_INPUT_COORD
        message = classification_text

        try:
            time.sleep(INPUT_CLICK_DELAY)
            pyautogui.click(x, y)
            time.sleep(0.1)

            if platform.system() == "Darwin":
                pyautogui.hotkey("command", "a")
            else:
                pyautogui.hotkey("ctrl", "a")
            pyautogui.press("backspace")

            if pyperclip is not None:
                try:
                    pyperclip.copy(message)
                    if platform.system() == "Darwin":
                        pyautogui.hotkey("command", "v")
                    else:
                        pyautogui.hotkey("ctrl", "v")
                except Exception as clip_exc:
                    print(f"[WARN] 클립보드 붙여넣기에 실패했습니다: {clip_exc}. 키보드 입력으로 대체합니다.")
                    pyautogui.write(message, interval=INPUT_TYPING_INTERVAL)
            else:
                pyautogui.write(message, interval=INPUT_TYPING_INTERVAL)
        except Exception as exc:
            print(f"[WARN] 자동 입력에 실패했습니다: {exc}")

    def _classify_score(self, score):
        """점수를 요청하신 5등급으로 분류하고 실제 점수를 문자열에 포함합니다."""
        score_value = float(score)
        score_str = f"{score_value:g}"

        if score_value == 0:
            detail = "관상동맥혈관벽에 석회화 침착 없음"
        elif 1 <= score_value <= 10:
            detail = "경도의 관상동맥 석회화"
        elif 11 <= score_value <= 100:
            detail = "중등도의 관상동맥 석회화"
        elif 101 <= score_value <= 400:
            detail = "중고등도의 관상동맥 석회화"
        elif score_value > 400:
            detail = "고도의 관상동맥 석회화"
        else:
            return "알 수 없는 점수 범위"

        return f"*****\n{detail} (coronary arterial calcium scoring : {score_str})\n\n"

    def on_hotkey_press(self):
        """'=' 키 감지 시 실행될 메인 로직입니다."""
        print("'=' 키 감지됨. 점수 추출 시작...")

        # ROI_1 먼저 시도, 실패 시 ROI_2 시도 (두 가지 패턴 처리)
        score = self._extract_from_roi(self.rois[0])
        if score is None:
            score = self._extract_from_roi(self.rois[1])

        # 결과 처리 및 표시
        if score is not None:
            classification = self._classify_score(score)
            self._input_result_to_target(str(score), classification)
            print(f"[INFO] 추출된 점수: {score}")
            print(f"[INFO] 분류 결과:\n{classification}")
        else:
            failure_text = "❌ 점수를 인식하지 못했습니다. CT 사진과 ROI를 확인하세요."
            self._input_result_to_target("N/A", failure_text)
            print("[WARN] 점수 추출 실패. 입력 필드에 실패 메시지를 기록했습니다.")


def start_listener():
    """키보드 리스너를 시작하고 메인 클래스를 실행합니다."""
    master = AgatstonScoreMaster(TESSERACT_PATH, [ROI_1, ROI_2])

    def on_activate():
        master.on_hotkey_press()

    print("'=' 키를 누르면 Agatston 점수 추출을 시도합니다.")
    print("프로그램 종료는 Ctrl+C 또는 창을 닫아주세요.")

    with keyboard.GlobalHotKeys({'=': on_activate}) as listener:
        listener.join()


if __name__ == "__main__":
    start_listener()