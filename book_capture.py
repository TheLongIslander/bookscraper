import os
import time
import threading
from datetime import datetime

import pyautogui
from pynput import keyboard
from PIL import Image

pyautogui.FAILSAFE = True

def wait_for_key(label: str, target_key):
    """
    Waits until the user presses `target_key`, then returns the current mouse (x, y) as ints.
    """
    print(f"[{label}] Hover your mouse where you want, then press {target_key.name.upper()} …")
    pos_holder = {}
    done = threading.Event()

    def on_press(key):
        if key == target_key:
            x, y = pyautogui.position()
            pos_holder["xy"] = (int(round(x)), int(round(y)))
            print(f"[{label}] Captured at {pos_holder['xy']}")
            done.set()
            return False  # stop listener

    with keyboard.Listener(on_press=on_press) as _:
        done.wait()
    return pos_holder["xy"]

def capture_next_button_xy():
    # Press F8 while hovering the NEXT button
    return wait_for_key("Next Button", keyboard.Key.f8)

def capture_region_by_keys():
    """
    Record a crop rectangle from two key-triggered points:
      - F6 over TOP-RIGHT corner
      - F7 over BOTTOM-LEFT corner
    """
    tr = wait_for_key("Region TOP-RIGHT (press F6)", keyboard.Key.f6)
    bl = wait_for_key("Region BOTTOM-LEFT (press F7)", keyboard.Key.f7)

    (x_tr, y_tr) = tr
    (x_bl, y_bl) = bl

    left   = min(x_tr, x_bl)
    top    = min(y_tr, y_bl)
    right  = max(x_tr, x_bl)
    bottom = max(y_tr, y_bl)

    width  = right - left
    height = bottom - top

    # Clamp to screen just in case
    scr_w, scr_h = pyautogui.size()
    left   = max(0, min(left, scr_w - 1))
    top    = max(0, min(top,  scr_h - 1))
    width  = max(1, min(width,  scr_w - left))
    height = max(1, min(height, scr_h - top))

    print(f"[Region] Using rectangle: left={left}, top={top}, width={width}, height={height}")
    return (left, top, width, height)

def images_to_pdf(img_paths, out_pdf):
    imgs = [Image.open(p).convert("RGB") for p in img_paths]
    if not imgs:
        return
    first, rest = imgs[0], imgs[1:]
    first.save(out_pdf, save_all=True, append_images=rest)
    print(f"[PDF] Wrote {out_pdf}")

def main():
    total_pages = int(input("How many pages to capture? (e.g., 386): ").strip() or "1")
    delay_after_click = float(input("Delay after clicking Next (seconds, e.g., 1.0-1.5): ").strip() or "1.2")
    use_double_click = input("Double-click to advance? (y/n): ").strip().lower() == "y"

    session_name = datetime.now().strftime("book_capture_%Y%m%d_%H%M%S")
    save_dir = os.path.join(os.getcwd(), session_name)
    os.makedirs(save_dir, exist_ok=True)
    print(f"\nSaving images to: {save_dir}\n")

    next_xy = capture_next_button_xy()         # (int, int)
    region  = capture_region_by_keys()         # (int, int, int, int)

    # Give the viewer focus once without clicking near page corners
    pyautogui.click(region[0] + 10, region[1] + 10)
    time.sleep(0.2)

    img_paths = []
    for i in range(1, total_pages + 1):
        # 1) Screenshot region
        img = pyautogui.screenshot(region=region)
        path = os.path.join(save_dir, f"page_{i:04d}.png")
        img.save(path)
        img_paths.append(path)
        print(f"[Capture] Saved {path}")

        # 2) Advance page (skip last)
        if i < total_pages:
            x, y = next_xy
            pyautogui.moveTo(x, y, duration=0.4)
            time.sleep(0.05)
            if use_double_click:
                pyautogui.click(x, y)
                time.sleep(0.05)
                pyautogui.click(x, y)
            else:
                pyautogui.click(x, y)
            time.sleep(delay_after_click)

    # Make a single PDF
    out_pdf = os.path.join(save_dir, "book.pdf")
    images_to_pdf(img_paths, out_pdf)
    print("\nDone!")

if __name__ == "__main__":
    main()
