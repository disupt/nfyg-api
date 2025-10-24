import threading

from pynput import keyboard, mouse


stop_event = threading.Event()


def on_move(x, y):
    """마우스가 움직일 때마다 좌표를 출력합니다."""
    if stop_event.is_set():
        # True/False 대신 None을 반환하면 리스너가 계속 유지됩니다.
        # 여기서는 종료 플래그가 설정된 경우만 정지하도록 False 반환.
        return False

    print(f"마우스 좌표: ({x}, {y})")


def on_key_press(key):
    """ESC 키를 누르면 프로그램을 종료합니다."""
    if key == keyboard.Key.esc:
        print("\n프로그램 종료 요청됨.")
        stop_event.set()
        return False


def main():
    print("--- 마우스 좌표 추적 시작 ---")
    print("마우스를 움직이면 실시간 좌표가 출력됩니다.")
    print("종료하려면 ESC 키를 누르세요.\n")

    mouse_listener = mouse.Listener(on_move=on_move)
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


