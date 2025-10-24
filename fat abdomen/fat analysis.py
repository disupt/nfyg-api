"""
CT 복부지방 자동 판독 프로그램
- 화면에서 복부지방 수치 자동 추출
- 비만도 자동 분류
- 판독지 자동 입력
"""

import pyautogui
import pytesseract
from PIL import Image, ImageGrab
import re
import keyboard
import tkinter as tk
from tkinter import messagebox
import threading


# ==================== 설정 ====================
# Tesseract 경로 설정 (설치 경로에 맞게 수정)
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# 지방값이 출력되는 두 영역의 좌표 (실제 화면에 맞게 조정 필요)
REGION_1 = (1526, 377, 1639, 387)  # (x1, y1, x2, y2)
REGION_2 = (335, 678, 449, 687)


class RegionVisualizer:
    """지정된 영역을 시각적으로 표시하는 클래스"""

    @staticmethod
    def highlight_regions(regions, duration=2000):
        """지정된 영역에 빨간색 테두리를 표시"""
        if not regions:
            return

        def _show():
            root = tk.Tk()
            root.withdraw()

            overlays = []
            for x1, y1, x2, y2 in regions:
                width = max(1, x2 - x1)
                height = max(1, y2 - y1)

                overlay = tk.Toplevel(root)
                overlay.overrideredirect(True)
                overlay.attributes("-topmost", True)
                overlay.attributes("-transparentcolor", "white")
                overlay.configure(bg="white")
                overlay.geometry(f"{width}x{height}+{x1}+{y1}")

                canvas = tk.Canvas(overlay, width=width, height=height, bg="white", highlightthickness=0)
                canvas.pack()
                canvas.create_rectangle(1, 1, width - 1, height - 1, outline="red", width=3)

                overlays.append(overlay)

            root.after(duration, root.destroy)
            root.mainloop()

        threading.Thread(target=_show, daemon=True).start()


# ==================== 비만도 분류 기준 ====================
class ObesityClassifier:
    """비만도 분류 클래스"""

    @staticmethod
    def classify_subcutaneous(value):
        """피하지방 분류"""
        if value < 20000:
            return "정상"
        elif 20000 <= value < 30000:
            return "경증"
        elif 30000 <= value < 40000:
            return "중등도"
        else:
            return "중증"

    @staticmethod
    def classify_visceral(value):
        """내장지방 분류"""
        if value < 20000:
            return "정상"
        elif 20000 <= value < 30000:
            return "경증"
        elif 30000 <= value < 40000:
            return "중등도"
        else:
            return "중증"

    @staticmethod
    def calculate_obesity_ratio(subcutaneous, visceral):
        """비만도 계산 및 분류"""
        if subcutaneous == 0:
            return 0, "계산불가"

        ratio = (visceral / subcutaneous) * 100

        if ratio < 30:
            grade = "정상"
        elif 30 <= ratio < 40:
            grade = "경증"
        elif 40 <= ratio < 50:
            grade = "중등도"
        else:
            grade = "중증"

        return ratio, grade


# ==================== 화면 캡처 및 OCR ====================
class FatValueExtractor:
    """화면에서 지방 수치를 추출하는 클래스"""

    @staticmethod
    def extract_numbers_from_region(region):
        """특정 영역에서 숫자 추출"""
        try:
            # 화면 캡처
            screenshot = ImageGrab.grab(bbox=region)

            # OCR로 텍스트 추출
            text = pytesseract.image_to_string(screenshot, config='--psm 6 digits')

            # 전처리: 불필요한 문자 제거 및 포맷 정리
            cleaned = (
                text.replace(',', '')
                .replace('O', '0')
                .replace('o', '0')
                .replace(':', ' ')
                .replace('㎟', ' ')
            )

            candidates = re.findall(r'\d+(?:\.\d+)?', cleaned)

            values = []
            for candidate in candidates:
                try:
                    value = float(candidate)
                    if value >= 1000:  # 작은 잡음 값 제거
                        values.append(int(round(value)))
                except ValueError:
                    continue

            if len(values) >= 2:
                return values[0], values[1]
            else:
                return None, None

        except Exception as e:
            print(f"영역 {region} 추출 오류: {e}")
            return None, None

    @staticmethod
    def extract_fat_values():
        """두 영역에서 지방값 추출"""
        # 영역 1에서 추출
        sub1, vis1 = FatValueExtractor.extract_numbers_from_region(REGION_1)

        # 영역 2에서 추출
        sub2, vis2 = FatValueExtractor.extract_numbers_from_region(REGION_2)

        # 유효한 값 선택
        subcutaneous = sub1 if sub1 is not None else sub2
        visceral = vis1 if vis1 is not None else vis2

        return subcutaneous, visceral


# ==================== 결과 표시 창 ====================
class ResultWindow:
    """결과를 표시하는 GUI 창"""

    @staticmethod
    def show_result(subcutaneous, visceral, sub_grade, vis_grade, ratio, ratio_grade):
        """결과 창 표시"""
        root = tk.Tk()
        root.title("CT 복부지방 분석 결과")
        root.geometry("400x300")

        # 결과 텍스트
        result_text = f"""
        ===== CT 복부지방 분석 결과 =====

        피하지방: {subcutaneous:,}㎟ ({sub_grade})

        내장지방: {visceral:,}㎟ ({vis_grade})

        비만도: {ratio:.1f}% ({ratio_grade})

        ================================
        """

        label = tk.Label(root, text=result_text, font=("맑은 고딕", 12), justify="left")
        label.pack(pady=20)

        # 판독지 전송 버튼
        def send_to_report():
            ReportGenerator.send_to_report(subcutaneous, visceral, sub_grade, vis_grade, ratio, ratio_grade)
            root.destroy()

        btn_send = tk.Button(
            root,
            text="판독지로 전송 (R)",
            command=send_to_report,
            font=("맑은 고딕", 11),
            bg="#4CAF50",
            fg="white",
            padx=20,
            pady=10,
        )
        btn_send.pack(pady=10)

        btn_close = tk.Button(
            root,
            text="닫기",
            command=root.destroy,
            font=("맑은 고딕", 11),
            padx=20,
            pady=10,
        )
        btn_close.pack()

        root.mainloop()


# ==================== 판독지 작성 ====================
class ReportGenerator:
    """판독지 자동 작성 클래스"""

    @staticmethod
    def generate_report_text(subcutaneous, visceral, sub_grade, vis_grade, ratio, ratio_grade):
        """판독문 생성"""
        report = (
            f"""피하지방 : {sub_grade}({subcutaneous}"㎟)

내장지방 : {vis_grade}({visceral}"㎟)
내장 비만도 : {ratio_grade}({ratio:.1f}%)

*****
피하지방 : {sub_grade}({subcutaneous}"㎟)
내장지방 : {vis_grade}({visceral}"㎟)"""
        )

        return report

    @staticmethod
    def send_to_report(subcutaneous, visceral, sub_grade, vis_grade, ratio, ratio_grade):
        """판독지로 전송"""
        try:
            # 판독창 열기 (r 키)
            pyautogui.press('r')
            pyautogui.sleep(0.5)

            # 판독문 생성
            report_text = ReportGenerator.generate_report_text(
                subcutaneous, visceral, sub_grade, vis_grade, ratio, ratio_grade
            )

            # 판독문 입력
            pyautogui.write(report_text, interval=0.01)

            messagebox.showinfo("완료", "판독지 전송이 완료되었습니다!")

        except Exception as e:
            messagebox.showerror("오류", f"판독지 전송 중 오류 발생:\n{e}")


# ==================== 메인 프로세스 ====================
class CTAnalyzer:
    """메인 분석 클래스"""

    @staticmethod
    def analyze():
        """전체 분석 프로세스 실행"""
        try:
            print("지방값 추출 중...")

            # 추출 영역을 시각적으로 안내
            RegionVisualizer.highlight_regions([REGION_1, REGION_2])

            # 1. 화면에서 지방값 추출
            subcutaneous, visceral = FatValueExtractor.extract_fat_values()

            if subcutaneous is None or visceral is None:
                messagebox.showerror("오류", "지방값을 추출할 수 없습니다.\n화면 영역 설정을 확인해주세요.")
                return

            print(f"추출 완료: 피하지방={subcutaneous}, 내장지방={visceral}")

            # 2. 비만도 분류
            classifier = ObesityClassifier()
            sub_grade = classifier.classify_subcutaneous(subcutaneous)
            vis_grade = classifier.classify_visceral(visceral)
            ratio, ratio_grade = classifier.calculate_obesity_ratio(subcutaneous, visceral)

            print(f"분류 완료: {sub_grade}, {vis_grade}, {ratio_grade}")

            # 3. 결과 창 표시
            ResultWindow.show_result(subcutaneous, visceral, sub_grade, vis_grade, ratio, ratio_grade)

        except Exception as e:
            messagebox.showerror("오류", f"분석 중 오류 발생:\n{e}")


# ==================== 단축키 설정 ====================
def setup_hotkey():
    """단축키 설정 (***를 F12로 대체)"""
    keyboard.add_hotkey('F12', lambda: threading.Thread(target=CTAnalyzer.analyze).start())
    print("CT 복부지방 자동 판독 프로그램 실행 중...")
    print("F12 키를 눌러 분석을 시작하세요.")
    print("종료하려면 Ctrl+C를 누르세요.")
    keyboard.wait()


# ==================== 프로그램 실행 ====================
if __name__ == "__main__":
    setup_hotkey()

