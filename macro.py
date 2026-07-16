### macro.py
from pynput import keyboard
import pyautogui
import time
import numpy as np

# ------------------- 설정 -------------------
region = (23, 157, 648, 611)
target_color = (79, 166, 52) # 초록
# target_color = (120, 105, 230) # 보라
tolerance = 5
confirm_button = (750, 655)

def color_match(c1, c2, tol):
    return all(abs(a - b) <= tol for a, b in zip(c1, c2))

def run_macro():
    screenshot = pyautogui.screenshot(region=region)
    img = np.array(screenshot.convert("RGB"))

    found = False
    for y in range(img.shape[0]):
        for x in range(img.shape[1]):
            r, g, b = img[y, x]
            if color_match((r, g, b), target_color, tolerance):
                real_x = region[0] + x
                real_y = region[1] + y
                pyautogui.click(real_x, real_y)

                print(f"좌석 클릭 완료: ({real_x}, {real_y}) / 색상: ({r}, {g}, {b})")

                pyautogui.click(*confirm_button)
                print("좌석 선택 완료 버튼 클릭")
                found = True
                break
        if found:
            break

    if not found:
        print("해당 색상을 찾지 못했습니다.")

# ------------------- 단축키 리스너 -------------------

def on_press(key):
    try:
        if key == keyboard.Key.alt:
            print("option 눌림 → 매크로 실행")
            run_macro()
    except Exception as e:
        print("오류 발생:", e)

print("⌨️ option 키를 누르면 매크로가 실행됩니다. 종료하려면 Ctrl+C.")
with keyboard.Listener(on_press=on_press) as listener:
    listener.join()
