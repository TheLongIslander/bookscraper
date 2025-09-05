#!/usr/bin/env python3
import os, sys, time, glob, shutil, subprocess, threading
from datetime import datetime

import pyautogui
from PIL import Image, ImageQt

from PySide6.QtCore import Qt, QTimer, QSize, QRect
from PySide6.QtGui import QPixmap, QPainter, QPen
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QLabel, QPushButton,
    QVBoxLayout, QHBoxLayout, QGroupBox, QCheckBox, QLineEdit,
    QSpinBox, QMessageBox
)

# macOS Quartz for key state polling
from Quartz import CGEventSourceKeyState, kCGEventSourceStateCombinedSessionState

# --- NEW: macOS power/activity controls (PyObjC Foundation) ---
try:
    from Foundation import (
        NSProcessInfo,
        NSActivityUserInitiated,
        NSActivityLatencyCritical,
        NSActivityIdleSystemSleepDisabled,
        NSActivityIdleDisplaySleepDisabled,
    )
    _FOUNDATION_OK = True
except Exception:
    _FOUNDATION_OK = False
# ---------------------------------------------------------------

pyautogui.FAILSAFE = True

# ----- mac virtual keycodes for function keys -----
VK_F6 = 97
VK_F7 = 98
VK_F8 = 100

# ------------------ helpers ------------------
def ensure_dirs():
    root_dir = os.path.join(os.getcwd(), "bookraw")
    os.makedirs(root_dir, exist_ok=True)
    pdf_root = os.path.join(os.getcwd(), "PDF")
    os.makedirs(pdf_root, exist_ok=True)
    return root_dir, pdf_root

def img_list_sorted_ctime(folder, pattern="page_*.png"):
    files = [f for f in glob.glob(os.path.join(folder, pattern))
             if os.path.splitext(f)[1].lower() in {".png", ".jpg", ".jpeg", ".tif", ".tiff"}]
    def times(p):
        st = os.stat(p)
        birth = getattr(st, "st_birthtime", None)
        mtime = st.st_mtime
        return (birth if birth else mtime, mtime, os.path.basename(p).lower())
    return sorted(files, key=times)

def make_pdf_from_folder(folder, out_pdf_path, pattern="page_*.png"):
    files = img_list_sorted_ctime(folder, pattern)
    if not files:
        return None
    try:
        import img2pdf
        with open(out_pdf_path, "wb") as f_out:
            f_out.write(img2pdf.convert(files))
        print(f"[PDF] Wrote {out_pdf_path} (img2pdf)")
    except Exception as e:
        print(f"[PDF] img2pdf failed ({e}); falling back to Pillow…")
        from PIL import Image
        imgs = [Image.open(p).convert("RGB") for p in files]
        first, rest = imgs[0], imgs[1:]
        first.save(out_pdf_path, save_all=True, append_images=rest)
        print(f"[PDF] Wrote {out_pdf_path} (Pillow)")
    return out_pdf_path

# Simple, permission-free activation (no AppleScript automation prompts)
def activate_by_name(app_name: str):
    if not app_name:
        return False
    try:
        subprocess.Popen(["open", "-a", app_name])
        return True
    except Exception as e:
        print(f"[Activate] open -a failed: {e}")
        return False

# ------------------ GUI ------------------
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("BookCap")
        self.resize(900, 640)

        # state
        self.next_xy = None
        self.top_right = None
        self.bottom_left = None
        self.preview_img = None   # PIL image captured at F8
        self.front_app_name = None
        self.session_dt = None
        self.save_dir = None

        # arming which key? one of: None/"F8"/"F6"/"F7"
        self.arm_for = None

        # for edge detection
        self._last_key_down = {VK_F6: False, VK_F7: False, VK_F8: False}

        # --- NEW: track NSActivity handle so we can end it cleanly ---
        self._ns_activity = None

        # preview
        self.preview = QLabel("Press “Record F8” to take a preview screenshot.")
        self.preview.setAlignment(Qt.AlignCenter)
        self.preview.setMinimumSize(QSize(600, 350))
        self.preview.setStyleSheet("QLabel { border: 1px solid #bbb; }")

        # options
        self.chk_auto_stop = QCheckBox("Auto-stop when duplicate page detected"); self.chk_auto_stop.setChecked(True)
        self.chk_double_click = QCheckBox("Double-click to advance")
        self.chk_use_keyboard = QCheckBox("Use Right Arrow instead of mouse click")

        self.delay_edit = QLineEdit("1.2"); self.delay_edit.setFixedWidth(80)
        self.delay_label = QLabel("Delay after advance (seconds):")

        self.chk_fixed = QCheckBox("Capture a specific number of pages")
        self.spin_pages = QSpinBox(); self.spin_pages.setRange(1, 100000); self.spin_pages.setValue(386); self.spin_pages.setEnabled(False)
        self.chk_fixed.toggled.connect(self.spin_pages.setEnabled)

        # buttons
        self.btn_f8 = QPushButton("Record F8 (Next button + take preview)")
        self.btn_f6 = QPushButton("Record F6 (Top-Right)")
        self.btn_f7 = QPushButton("Record F7 (Bottom-Left)")
        self.btn_start = QPushButton("Start Capture")

        self.btn_f8.clicked.connect(lambda: self.arm("F8"))
        self.btn_f6.clicked.connect(lambda: self.arm("F6"))
        self.btn_f7.clicked.connect(lambda: self.arm("F7"))
        self.btn_start.clicked.connect(self.start_capture)

        # layout
        options_box = QGroupBox("Options")
        opt_layout = QVBoxLayout()
        row1 = QHBoxLayout(); row1.addWidget(self.chk_auto_stop); row1.addWidget(self.chk_double_click); row1.addWidget(self.chk_use_keyboard); row1.addStretch()
        row2 = QHBoxLayout(); row2.addWidget(self.delay_label); row2.addWidget(self.delay_edit); row2.addStretch()
        row3 = QHBoxLayout(); row3.addWidget(self.chk_fixed); row3.addWidget(self.spin_pages); row3.addStretch()
        opt_layout.addLayout(row1); opt_layout.addLayout(row2); opt_layout.addLayout(row3)
        options_box.setLayout(opt_layout)

        keys_box = QGroupBox("Capture Keys")
        keys_layout = QHBoxLayout()
        keys_layout.addWidget(self.btn_f8)
        keys_layout.addWidget(self.btn_f6)
        keys_layout.addWidget(self.btn_f7)
        keys_box.setLayout(keys_layout)

        main = QWidget(); v = QVBoxLayout(main)
        v.addWidget(self.preview, 1)
        v.addWidget(keys_box)
        v.addWidget(options_box)
        v.addWidget(self.btn_start)
        self.setCentralWidget(main)

        QMessageBox.information(
            self, "Permissions",
            "macOS: Give your Terminal/VS Code *and* Python Screen Recording + Accessibility in System Settings → Privacy & Security.\n"
            "If F-keys control brightness/volume, hold Fn or enable “Use F1, F2, etc. as standard function keys”."
        )

        # poll F-keys on the main thread (no Carbon handlers = no HIToolbox crashes)
        self.poll_timer = QTimer(self)
        self.poll_timer.timeout.connect(self._poll_hotkeys)
        self.poll_timer.start(30)

    # ------------ arming ------------
    def arm(self, which: str):
        self.arm_for = which
        self.statusBar().showMessage(f"Armed for {which}. Hover target and press {which}.")

    # ------------ preview rendering (fixed DPI + letterboxing) ------------
    def _render_preview(self):
        if not self.preview_img:
            return

        src = self.preview_img  # PIL
        sw, sh = src.size
        lw, lh = self.preview.width(), self.preview.height()

        # scale to fit, preserve aspect
        s = min(lw / sw, lh / sh)
        pw, ph = int(sw * s), int(sh * s)
        xoff, yoff = (lw - pw) // 2, (lh - ph) // 2

        canvas = QPixmap(lw, lh)
        canvas.fill(Qt.white)
        painter = QPainter(canvas)

        qimg = ImageQt.ImageQt(src)
        painter.drawImage(QRect(xoff, yoff, pw, ph), qimg, QRect(0, 0, sw, sh))

        # correct for Retina scaling (coords are in points; screenshot is pixels)
        screen_w, screen_h = pyautogui.size()
        ratio_x = sw / float(screen_w) if screen_w else 1.0
        ratio_y = sh / float(screen_h) if screen_h else 1.0

        # draw selection rect if both corners set
        if self.top_right and self.bottom_left:
            (x_tr, y_tr) = self.top_right
            (x_bl, y_bl) = self.bottom_left
            left, top = min(x_tr, x_bl), min(y_tr, y_bl)
            right, bottom = max(x_tr, x_bl), max(y_tr, y_bl)

            rx1 = int(left * ratio_x * s) + xoff
            ry1 = int(top * ratio_y * s) + yoff
            rx2 = int(right * ratio_x * s) + xoff
            ry2 = int(bottom * ratio_y * s) + yoff

            pen = QPen(Qt.red); pen.setWidth(3)
            painter.setPen(pen)
            painter.drawRect(rx1, ry1, rx2 - rx1, ry2 - ry1)

        painter.end()
        self.preview.setPixmap(canvas)

    # ------------ hotkey polling ------------
    def _is_key_down(self, keycode: int) -> bool:
        try:
            return bool(CGEventSourceKeyState(kCGEventSourceStateCombinedSessionState, keycode))
        except Exception:
            return False

    def _poll_hotkeys(self):
        states = {
            VK_F6: self._is_key_down(VK_F6),
            VK_F7: self._is_key_down(VK_F7),
            VK_F8: self._is_key_down(VK_F8),
        }
        for vk, down in states.items():
            if down and not self._last_key_down[vk]:
                self._handle_key_down(vk)
        self._last_key_down.update(states)

    def _handle_key_down(self, vk: int):
        which = "F8" if vk == VK_F8 else "F6" if vk == VK_F6 else "F7"
        if self.arm_for != which:
            return

        try:
            if which == "F8":
                pos = pyautogui.position()
                self.next_xy = (int(pos.x), int(pos.y))
                self.preview_img = pyautogui.screenshot()  # full screen
                # record current front app name
                try:
                    r = subprocess.run(
                        ['osascript', '-e',
                         'tell application "System Events" to get name of (first process whose frontmost is true)'],
                        capture_output=True, text=True, timeout=2
                    )
                    self.front_app_name = r.stdout.strip() or None
                except Exception:
                    self.front_app_name = None

                self._render_preview()
                self.statusBar().showMessage(f"F8 recorded at {self.next_xy}. Front app: {self.front_app_name or 'unknown'}")

            elif which == "F6":
                pos = pyautogui.position()
                self.top_right = (int(pos.x), int(pos.y))
                self._render_preview()
                self.statusBar().showMessage(f"F6 recorded at {self.top_right}")

            elif which == "F7":
                pos = pyautogui.position()
                self.bottom_left = (int(pos.x), int(pos.y))
                self._render_preview()
                self.statusBar().showMessage(f"F7 recorded at {self.bottom_left}")
        finally:
            self.arm_for = None

    # ------------ capture loop ------------
    def start_capture(self):
        if not self.next_xy or not self.top_right or not self.bottom_left:
            QMessageBox.warning(self, "Missing points", "Please record F8, F6, and F7 first.")
            return

        try:
            delay = float(self.delay_edit.text().strip())
        except ValueError:
            QMessageBox.warning(self, "Delay", "Delay must be a number (seconds).")
            return

        auto_stop = self.chk_auto_stop.isChecked()
        use_double = self.chk_double_click.isChecked()
        use_keyboard = self.chk_use_keyboard.isChecked()
        fixed = self.chk_fixed.isChecked()
        pages = self.spin_pages.value() if fixed else None

        root_dir, _ = ensure_dirs()
        self.session_dt = datetime.now()
        session_name = self.session_dt.strftime("book_capture_%Y%m%d_%H%M%S")
        self.save_dir = os.path.join(root_dir, session_name)
        os.makedirs(self.save_dir, exist_ok=True)

        (x_tr, y_tr) = self.top_right
        (x_bl, y_bl) = self.bottom_left
        left = max(0, min(x_tr, x_bl)); top = max(0, min(y_tr, y_bl))
        right = max(x_tr, x_bl); bottom = max(y_tr, y_bl)
        region = (left, top, right - left, bottom - top)

        self.statusBar().showMessage("Capturing…")

        def worker():
            # ===== BEGIN: Stay-awake stack =====
            # A) Disable App Nap / idle sleeps for this process (if Foundation available)
            if _FOUNDATION_OK:
                try:
                    self._ns_activity = NSProcessInfo.processInfo().beginActivityWithOptions_reason_(
                        NSActivityUserInitiated
                        | NSActivityLatencyCritical
                        | NSActivityIdleSystemSleepDisabled
                        | NSActivityIdleDisplaySleepDisabled,
                        "BookCap capture",
                    )
                except Exception as e:
                    print(f"[NSActivity] beginActivity failed: {e}")

            # B) System-level keep-awake (runs while our PID exists)
            try:
                caffeinate = subprocess.Popen(
                    ["caffeinate", "-dis", "-w", str(os.getpid())],
                    stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL
                )
            except Exception as e:
                caffeinate = None
                print(f"[caffeinate] spawn failed: {e}")

            # C) Touch Bar / UI idle preventer: pulse “user active” every ~30s
            userpulse_stop = threading.Event()
            def _userpulse():
                while not userpulse_stop.is_set():
                    try:
                        subprocess.run(
                            ["caffeinate", "-u", "-t", "2"],
                            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL
                        )
                    except Exception:
                        pass
                    userpulse_stop.wait(30)

            pulse_thread = threading.Thread(target=_userpulse, daemon=True)
            pulse_thread.start()
            # ===== END: Stay-awake stack =====

            # Switch back to target app off the UI thread
            if not activate_by_name(self.front_app_name):
                print("[Activate] Falling back to direct click without activation.")
            time.sleep(0.5)

            # Focus inside region (avoids corner auto-flip)
            try:
                pyautogui.click(region[0] + 10, region[1] + 10)
                time.sleep(0.15)
            except Exception as e:
                print(f"[Focus] click failed: {e}")

            prev_bytes = None
            page_index = 0

            try:
                while True:
                    page_index += 1

                    # tiny jiggle keeps the event stream “hot”
                    pyautogui.moveRel(0, 1, duration=0)
                    pyautogui.moveRel(0, -1, duration=0)

                    img = pyautogui.screenshot(region=region)
                    path = os.path.join(self.save_dir, f"page_{page_index:04d}.png")
                    img.save(path)
                    print(f"[Capture] {path}")

                    if auto_stop:
                        img_bytes = img.convert("RGB").tobytes()
                        if prev_bytes is not None and img_bytes == prev_bytes:
                            try:
                                os.remove(path)
                                print(f"[Auto-Stop] Duplicate detected. Removed: {path}")
                            except Exception:
                                pass
                            break
                        prev_bytes = img_bytes

                    if fixed and page_index >= pages:
                        break

                    # Re-assert focus every 10 pages
                    if page_index % 10 == 0:
                        activate_by_name(self.front_app_name)
                        time.sleep(0.1)

                    if use_keyboard:
                        pyautogui.press("right")
                    else:
                        x, y = self.next_xy
                        pyautogui.moveTo(x, y, duration=0.25)
                        time.sleep(0.05)
                        pyautogui.click(x, y)
                        if use_double:
                            time.sleep(0.05)
                            pyautogui.click(x, y)

                    time.sleep(delay)
            finally:
                # stop user-activity pulses
                try:
                    userpulse_stop.set()
                except Exception:
                    pass

                # end NSActivity so macOS can nap again
                if _FOUNDATION_OK:
                    try:
                        if self._ns_activity is not None:
                            NSProcessInfo.processInfo().endActivity_(self._ns_activity)
                            self._ns_activity = None
                    except Exception as e:
                        print(f"[NSActivity] endActivity failed: {e}")

                # stop caffeinate
                try:
                    if caffeinate is not None:
                        caffeinate.terminate()
                except Exception:
                    pass

            # Make PDFs
            run_pdf_path = os.path.join(self.save_dir, "book.pdf")
            made = make_pdf_from_folder(self.save_dir, out_pdf_path=run_pdf_path)
            if made:
                pdf_root = os.path.join(os.getcwd(), "PDF")
                human_name = f"{self.session_dt:%Y}-{self.session_dt:%b}-{self.session_dt.day}-{self.session_dt:%H%M}.pdf"
                dest_pdf = os.path.join(pdf_root, human_name)
                try:
                    shutil.copyfile(made, dest_pdf)
                    print(f"[PDF] Copied to {dest_pdf}")
                except Exception as e:
                    print(f"[PDF] Copy failed: {e}")

            QTimer.singleShot(0, lambda: (self.statusBar().showMessage("Done!"),
                                          QMessageBox.information(self, "Finished", "Capture complete.")))

        threading.Thread(target=worker, daemon=True).start()


def main():
    app = QApplication(sys.argv)
    w = MainWindow()
    w.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
