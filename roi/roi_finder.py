import threading

from pynput import keyboard, mouse


coords = {}
stop_event = threading.Event()


def reset_coords():
    coords.clear()
    print("다음 ROI를 위해 다시 '왼쪽 위'를 클릭하세요. (종료: ESC)")


def on_click(x, y, button, pressed):
    """마우스 왼쪽 버튼 클릭으로 전역 좌표를 기록합니다."""
    if button != mouse.Button.left or not pressed:
        return

    if 'top_left' not in coords:
        coords['top_left'] = (x, y)
        print(f"--- 1단계: 왼쪽 위 모서리 (x, y): ({x}, {y}) ---")
    elif 'bottom_right' not in coords:
        coords['bottom_right'] = (x, y)
        print(f"--- 2단계: 오른쪽 아래 모서리 (x, y): ({x}, {y}) ---")

        left, top = coords['top_left']
        width = coords['bottom_right'][0] - left
        height = coords['bottom_right'][1] - top

        if width > 0 and height > 0:
            print("\n✅ 추출된 ROI 좌표 (메인 프로그램에 복사할 값):")
            print(f"{{'top': {top}, 'left': {left}, 'width': {width}, 'height': {height}}}")
            print("-------------------------------------------\n")
        else:
            print("\n⚠️ 오류: 영역이 너무 작거나 마우스 드래그 방향이 잘못되었습니다. 다시 시도하세요.\n")

        reset_coords()


def on_key_press(key):
    """ESC 키를 눌러 프로그램을 종료합니다."""
    if key == keyboard.Key.esc:
        print("프로그램 종료.")
        stop_event.set()
        # False를 반환하면 키보드 리스너가 종료됩니다.
        return False


def main():
    print("--- ROI 추출기 시작 ---")
    print("1. 측정하려는 ROI의 [왼쪽 위] 지점을 클릭하세요.")
    print("2. 같은 ROI의 [오른쪽 아래] 지점을 클릭하면 좌표가 출력됩니다.")
    print("3. 다른 ROI도 동일한 방법으로 반복하세요.")
    print("4. 종료하려면 ESC 키를 누르세요.")
    print("(macOS에서는 처음 실행 시 '손쉬운 사용' 권한 허용이 필요할 수 있습니다.)\n")

    mouse_listener = mouse.Listener(on_click=on_click)
    keyboard_listener = keyboard.Listener(on_press=on_key_press)

    mouse_listener.start()
    keyboard_listener.start()

    try:
        stop_event.wait()
    finally:
        mouse_listener.stop()
        keyboard_listener.stop()
        mouse_listener.join()
        keyboard_listener.join()


if __name__ == "__main__":
    main()

