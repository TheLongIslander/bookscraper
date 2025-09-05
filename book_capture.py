#!/usr/bin/env python3
import os
import time
import glob
import shutil
import threading
from datetime import datetime

import pyautogui
from pynput import keyboard
from PIL import Image

pyautogui.FAILSAFE = True

# ---------- helpers for key-captured points ----------
def wait_for_key(label: str, target_key):
    """Wait until user presses target_key, then return current mouse (x, y) as ints."""
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

    # Clamp to screen
    scr_w, scr_h = pyautogui.size()
    left   = max(0, min(left, scr_w - 1))
    top    = max(0, min(top,  scr_h - 1))
    width  = max(1, min(width,  scr_w - left))
    height = max(1, min(height, scr_h - top))

    print(f"[Region] Using rectangle: left={left}, top={top}, width={width}, height={height}")
    return (left, top, width, height)

# ---------- PDF building (img2pdf preferred; Pillow fallback) ----------
def _img_file_order_ctime_asc(files):
    def times(p):
        st = os.stat(p)
        birth = getattr(st, "st_birthtime", None)
        mtime = st.st_mtime
        return (birth if birth else mtime, mtime, os.path.basename(p).lower())
    return sorted(files, key=times)

def make_pdf_from_folder(folder, out_pdf_path, pattern="page_*.png"):
    files = [f for f in glob.glob(os.path.join(folder, pattern))
             if os.path.splitext(f)[1].lower() in {".png", ".jpg", ".jpeg", ".tif", ".tiff"}]
    if not files:
        print("[PDF] No images found to combine.")
        return None

    files = _img_file_order_ctime_asc(files)  # oldest → newest

    # Try img2pdf (no recompression); fallback to Pillow
    try:
        import img2pdf
        with open(out_pdf_path, "wb") as f_out:
            f_out.write(img2pdf.convert(files))
        print(f"[PDF] Wrote {out_pdf_path} (via img2pdf)")
    except Exception as e:
        print(f"[PDF] img2pdf unavailable/failed ({e}). Falling back to Pillow…")
        imgs = [Image.open(p).convert("RGB") for p in files]
        first, rest = imgs[0], imgs[1:]
        first.save(out_pdf_path, save_all=True, append_images=rest)
        print(f"[PDF] Wrote {out_pdf_path} (via Pillow)")

    return out_pdf_path

# ---------- main capture ----------
def main():
    total_pages = int(input("How many pages to capture? (e.g., 386): ").strip() or "1")
    delay_after_click = float(input("Delay after clicking Next (seconds, e.g., 1.0-1.5): ").strip() or "1.2")
    use_double_click = input("Double-click to advance? (y/n): ").strip().lower() == "y"

    # Timestamp for this session (reuse consistently for run + PDF filename)
    session_dt = datetime.now()

    # Save under ./bookraw/book_capture_YYYYMMDD_HHMMSS
    root_dir = os.path.join(os.getcwd(), "bookraw")
    os.makedirs(root_dir, exist_ok=True)
    session_name = session_dt.strftime("book_capture_%Y%m%d_%H%M%S")
    save_dir = os.path.join(root_dir, session_name)
    os.makedirs(save_dir, exist_ok=True)
    print(f"\nSaving images to: {save_dir}\n")

    next_xy = capture_next_button_xy()
    region  = capture_region_by_keys()

    # Focus the viewer once (click inside region, not near corners)
    pyautogui.click(region[0] + 10, region[1] + 10)
    time.sleep(0.2)

    # Capture loop
    for i in range(1, total_pages + 1):
        img = pyautogui.screenshot(region=region)
        path = os.path.join(save_dir, f"page_{i:04d}.png")
        img.save(path)
        print(f"[Capture] Saved {path}")

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

    # Build PDF inside the run folder
    run_pdf_path = os.path.join(save_dir, "book.pdf")
    made = make_pdf_from_folder(save_dir, out_pdf_path=run_pdf_path, pattern="page_*.png")

    # Also place a human-named copy into ./PDF/
    if made:
        pdf_root = os.path.join(os.getcwd(), "PDF")
        os.makedirs(pdf_root, exist_ok=True)

        # Human-readable name: YYYY-Mon-D-HHMM.pdf  (e.g., 2025-Sep-5-1427.pdf)
        human_name = f"{session_dt:%Y}-{session_dt:%b}-{session_dt.day}-{session_dt:%H%M}.pdf"
        dest_pdf_path = os.path.join(pdf_root, human_name)

        try:
            shutil.copyfile(made, dest_pdf_path)
            print(f"[PDF] Copied to {dest_pdf_path}")
        except Exception as e:
            print(f"[PDF] Could not copy to {dest_pdf_path}: {e}")

    print("\nDone!")

if __name__ == "__main__":
    # Reminder: macOS Privacy & Security → enable Accessibility, Input Monitoring, and Screen Recording
    # for your Terminal/VSCode and the Python binary you are using.
    main()
