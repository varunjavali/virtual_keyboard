import sys
import cv2
import time
import numpy as np
import os
import threading
import webbrowser
import speech_recognition as sr

os.environ["QT_ENABLE_HIGHDPI_SCALING"] = "0"
os.environ["QT_AUTO_SCREEN_SCALE_FACTOR"] = "0"

import pyautogui
from cvzone.HandTrackingModule import HandDetector
from pynput.keyboard import Controller as KeyboardController, Key as PynputKey
from pynput.mouse import Controller as MouseController, Button as MouseButton

import win32gui, win32con
from PySide6.QtWidgets import (QApplication, QWidget, QPushButton,
                                QVBoxLayout, QLabel, QSlider, QHBoxLayout)
from PySide6.QtCore import Qt

# ================= GLOBAL STATES =================
mode          = "keyboard"
finalText     = ""
voice_enabled = True
running       = True
text_lock     = threading.Lock()

DWELL_TIME  = 0.45   # seconds finger must hover before key fires
PINCH_RATIO = 0.28   # pinch dist / hand width threshold

# ================= WIN32 TOPMOST (once per window) =================
_topmost_done = set()

def pin_topmost(title):
    if title in _topmost_done:
        return
    hwnd = win32gui.FindWindow(None, title)
    if hwnd:
        win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0,
                              win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
        _topmost_done.add(title)

# ================= TOOLBAR =================
class ControlToolbar(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Control Panel")
        self.setFixedSize(280, 300)
        self.setWindowFlags(Qt.WindowStaysOnTopHint | Qt.Tool)
        layout = QVBoxLayout()

        self.status = QLabel("MODE: KEYBOARD")
        self.status.setAlignment(Qt.AlignCenter)
        self.status.setStyleSheet("font-weight:bold;font-size:14px;color:#8800ff")

        self.kb  = QPushButton("⌨  KEYBOARD MODE")
        self.ms  = QPushButton("🖱  MOUSE MODE")
        self.clr = QPushButton("🗑  CLEAR TEXT")
        self.kb.clicked.connect(self.setKeyboard)
        self.ms.clicked.connect(self.setMouse)
        self.clr.clicked.connect(self.clearText)

        dwell_row = QHBoxLayout()
        dwell_row.addWidget(QLabel("Dwell:"))
        self.dwell_val    = QLabel(f"{DWELL_TIME:.2f}s")
        self.dwell_slider = QSlider(Qt.Horizontal)
        self.dwell_slider.setRange(20, 100)
        self.dwell_slider.setValue(int(DWELL_TIME * 100))
        self.dwell_slider.valueChanged.connect(self._on_dwell)
        dwell_row.addWidget(self.dwell_slider)
        dwell_row.addWidget(self.dwell_val)

        tip = QLabel("  ☝ Point index finger to hover key\n  Hold still → key fires automatically")
        tip.setStyleSheet("font-size:10px; color:#555")

        layout.addWidget(self.status)
        layout.addWidget(self.kb)
        layout.addWidget(self.ms)
        layout.addWidget(self.clr)
        layout.addLayout(dwell_row)
        layout.addWidget(tip)
        self.setLayout(layout)

    def _on_dwell(self, v):
        global DWELL_TIME
        DWELL_TIME = v / 100.0
        self.dwell_val.setText(f"{DWELL_TIME:.2f}s")

    def setKeyboard(self):
        global mode
        mode = "keyboard"
        self.status.setText("MODE: KEYBOARD")

    def setMouse(self):
        global mode
        mode = "mouse"
        self.status.setText("MODE: MOUSE")

    def clearText(self):
        global finalText
        with text_lock:
            finalText = ""

# ================= VOICE THREAD =================
def voice_loop():
    global mode, finalText, voice_enabled, running
    r = sr.Recognizer()
    r.energy_threshold         = 300
    r.dynamic_energy_threshold = True
    try:
        mic = sr.Microphone()
    except Exception as e:
        print(f"[Voice] Mic init failed: {e}")
        return
    with mic as src:
        r.adjust_for_ambient_noise(src, duration=1)

    while running:
        if not voice_enabled:
            time.sleep(0.5)
            continue
        try:
            with mic as src:
                audio = r.listen(src, timeout=5, phrase_time_limit=4)
            cmd = r.recognize_google(audio).lower()
            print(f"[Voice] {cmd}")
            if   "keyboard mode"    in cmd: mode = "keyboard"
            elif "mouse mode"       in cmd: mode = "mouse"
            elif "clear text"       in cmd:
                with text_lock: finalText = ""
            elif "open chatgpt"     in cmd: webbrowser.open("https://chat.openai.com")
            elif "stop listening"   in cmd: voice_enabled = False
            elif "start listening"  in cmd: voice_enabled = True
            elif "exit application" in cmd: running = False; break
        except sr.WaitTimeoutError:  pass
        except sr.UnknownValueError: pass
        except sr.RequestError as e: print(f"[Voice] API: {e}")
        except Exception as e:       print(f"[Voice] {e}")

# ================= KEY LAYOUT =================
KEY_ROWS = [
    ["Q","W","E","R","T","Y","U","I","O","P"],
    ["A","S","D","F","G","H","J","K","L",";"],
    ["Z","X","C","V","B","N","M",",",".","?"],
    ["SPACE", "BKSP"],
]
KW, KH, KGAP = 88, 72, 6
SX, SY = 35, 25

class Key:
    def __init__(self, pos, text, w=KW, h=KH):
        self.pos         = pos
        self.text        = text
        self.w           = w
        self.h           = h
        self.hover_since = None
        self.last_press  = 0
        self.progress    = 0.0

    def is_hover(self, fx, fy):
        x, y = self.pos
        return x < fx < x + self.w and y < fy < y + self.h

    def update_dwell(self, hovering, now):
        """Returns True exactly once when dwell threshold is crossed."""
        if not hovering:
            self.hover_since = None
            self.progress    = 0.0
            return False
        if self.hover_since is None:
            self.hover_since = now
        elapsed       = now - self.hover_since
        self.progress = min(elapsed / DWELL_TIME, 1.0)
        if self.progress >= 1.0 and (now - self.last_press) > DWELL_TIME + 0.15:
            self.last_press  = now
            self.hover_since = now   # allow repeat if still held
            return True
        return False

def build_keys():
    buttons = []
    for i, row in enumerate(KEY_ROWS):
        if i < 3:
            for j, k in enumerate(row):
                x = SX + j * (KW + KGAP)
                y = SY + i * (KH + KGAP)
                buttons.append(Key([x, y], k))
        else:
            y = SY + 3 * (KH + KGAP)
            buttons.append(Key([SX,       y], "SPACE", w=310, h=KH))
            buttons.append(Key([SX + 320, y], "BKSP",  w=200, h=KH))
    return buttons

# ================= HELPERS =================
class Smoother:
    def __init__(self, a=0.18):
        self.a = a
        self.px = self.py = None

    def smooth(self, x, y):
        if self.px is None:
            self.px, self.py = float(x), float(y)
        else:
            self.px = self.a * x + (1 - self.a) * self.px
            self.py = self.a * y + (1 - self.a) * self.py
        return int(self.px), int(self.py)

    def reset(self):
        self.px = self.py = None


def fingers_up(lm):
    """
    Returns [thumb, index, middle, ring, pinky] True=extended.
    Assumes image has been flipped horizontally.
    """
    tips   = [4,  8, 12, 16, 20]
    joints = [3,  6, 10, 14, 18]
    up = []
    # Thumb: horizontal axis
    up.append(lm[tips[0]][0] > lm[joints[0]][0])
    # Fingers: vertical axis (y increases downward)
    for i in range(1, 5):
        up.append(lm[tips[i]][1] < lm[joints[i]][1])
    return up


def hand_width(lm):
    """Wrist-to-middle-MCP distance as a stable hand size reference."""
    return float(np.hypot(lm[0][0] - lm[9][0], lm[0][1] - lm[9][1])) + 1e-6


def draw_dwell_arc(img, cx, cy, progress, r=24):
    if progress <= 0:
        return
    angle = int(360 * progress)
    g     = int(255 * progress)
    b     = int(255 * (1 - progress))
    cv2.ellipse(img, (cx, cy), (r, r), -90, 0, angle, (0, g, b), 3)
    cv2.circle(img,  (cx, cy), 5, (255, 255, 255), -1)


# ================= CAMERA THREAD =================
def camera_loop():
    global finalText, running

    cap = cv2.VideoCapture(0)
    cap.set(cv2.CAP_PROP_FRAME_WIDTH,  1280)
    cap.set(cv2.CAP_PROP_FRAME_HEIGHT, 720)
    cap.set(cv2.CAP_PROP_FPS, 30)

    # High confidence thresholds for stable tracking
    detector = HandDetector(detectionCon=0.90, maxHands=1)
    try:
        detector.mpHands.min_tracking_confidence = 0.85
    except Exception:
        pass   # some cvzone versions don't expose this

    kb_ctrl  = KeyboardController()
    mouse    = MouseController()
    smoother = Smoother(a=0.18)
    screenW, screenH = pyautogui.size()

    keys         = build_keys()
    was_clicking = False
    flash        = {}
    frames_no_hand = 0

    while running:
        ok, img = cap.read()
        if not ok:
            time.sleep(0.02)
            continue

        img = cv2.flip(img, 1)

        hands, img = detector.findHands(img, draw=True, flipType=False)

        now    = time.time()
        fx, fy = -1, -1
        dist   = 9999.0
        hw     = 1.0
        lm     = None
        up     = [False] * 5

        if not hands:
            frames_no_hand += 1
        else:
            frames_no_hand = 0
            lm   = hands[0]["lmList"]
            hw   = hand_width(lm)
            up   = fingers_up(lm)

            ix, iy   = lm[8][0],  lm[8][1]
            mx2, my2 = lm[12][0], lm[12][1]
            dist     = float(np.hypot(mx2 - ix, my2 - iy))
            fx, fy   = ix, iy

            # ── MOUSE MODE ──────────────────────────────────────────
            if mode == "mouse":
                sx, sy = smoother.smooth(ix, iy)
                msx = int(np.interp(sx, (80, 1200), (0, screenW)))
                msy = int(np.interp(sy, (60, 660),  (0, screenH)))
                msx = max(0, min(msx, screenW - 1))
                msy = max(0, min(msy, screenH - 1))
                mouse.position = (msx, msy)

                # Ratio-based pinch: works at any camera distance
                pinching = (dist / hw) < PINCH_RATIO
                if pinching and not was_clicking:
                    mouse.click(MouseButton.left, 1)
                    print(f"[Mouse] Click ({msx},{msy})  ratio={dist/hw:.2f}")
                was_clicking = pinching
            else:
                smoother.reset()
                was_clicking = False

        # ── KEYBOARD MODE ───────────────────────────────────────────
        if mode == "keyboard":
            # Require pointing gesture: index up, middle+ring+pinky curled
            intentional = bool(lm) and up[1] and not up[2] and not up[3] and not up[4]

            for b in keys:
                hov   = b.is_hover(fx, fy)
                fired = b.update_dwell(hov and intentional, now)

                flashing = flash.get(b.text, 0) > now

                if flashing:
                    color = (255, 255, 255)
                elif hov and intentional:
                    g     = int(60 + 195 * b.progress)
                    color = (0, g, 30)
                elif hov:
                    color = (70, 70, 130)   # hovering but wrong gesture
                else:
                    color = (110, 0, 110)

                x, y = b.pos
                cv2.rectangle(img, (x, y), (x + b.w, y + b.h), color, -1)
                cv2.rectangle(img, (x, y), (x + b.w, y + b.h), (200, 200, 200), 1)
                fs = 0.85 if b.text in ("SPACE", "BKSP") else 1.1
                cv2.putText(img, b.text, (x + 8, y + int(b.h * 0.72)),
                            cv2.FONT_HERSHEY_SIMPLEX, fs, (255, 255, 255), 2)

                if fired:
                    flash[b.text] = now + 0.18
                    if b.text == "BKSP":
                        with text_lock: finalText = finalText[:-1]
                        kb_ctrl.press(PynputKey.backspace)
                        kb_ctrl.release(PynputKey.backspace)
                    elif b.text == "SPACE":
                        with text_lock: finalText += " "
                        kb_ctrl.press(PynputKey.space)
                        kb_ctrl.release(PynputKey.space)
                    else:
                        with text_lock: finalText += b.text
                        kb_ctrl.press(b.text.lower())
                        kb_ctrl.release(b.text.lower())
                    print(f"[Key] '{b.text}'")

            # Dwell arc around fingertip
            if fx > 0:
                best_prog = max((b.progress for b in keys if b.is_hover(fx, fy)), default=0.0)
                draw_dwell_arc(img, fx, fy, best_prog)

        # ── HUD ─────────────────────────────────────────────────────
        with text_lock:
            disp = finalText[-44:]

        cv2.rectangle(img, (28, 590), (1252, 710), (40, 0, 50), -1)
        cv2.rectangle(img, (28, 590), (1252, 710), (180, 0, 180), 2)
        cv2.putText(img, disp, (45, 665),
                    cv2.FONT_HERSHEY_SIMPLEX, 1.4, (255, 255, 255), 2)

        mc = (0, 255, 120) if mode == "keyboard" else (0, 210, 255)
        cv2.putText(img, f"MODE: {mode.upper()}", (970, 46),
                    cv2.FONT_HERSHEY_SIMPLEX, 0.9, mc, 2)

        if lm is not None:
            ratio = dist / hw
            rc    = (0, 255, 0) if ratio < PINCH_RATIO else (0, 140, 255)
            cv2.putText(img, f"Pinch: {ratio:.2f}", (970, 78),
                        cv2.FONT_HERSHEY_SIMPLEX, 0.65, rc, 2)

            if mode == "keyboard":
                gtxt = "POINT ✓ — hover key to type" if intentional else "☝ curl middle/ring/pinky"
                gcol = (0, 255, 80) if intentional else (0, 160, 255)
                cv2.putText(img, gtxt, (970, 110),
                            cv2.FONT_HERSHEY_SIMPLEX, 0.6, gcol, 2)

        if frames_no_hand > 5:
            cv2.putText(img, "⚠  NO HAND DETECTED", (380, 370),
                        cv2.FONT_HERSHEY_SIMPLEX, 1.3, (0, 60, 255), 3)

        cv2.imshow("Gesture Controller", img)
        pin_topmost("Gesture Controller")

        if cv2.waitKey(1) & 0xFF == 27:
            running = False
            break

    cap.release()
    cv2.destroyAllWindows()

# ================= MAIN =================
if __name__ == "__main__":
    app     = QApplication(sys.argv)
    toolbar = ControlToolbar()
    toolbar.show()

    threading.Thread(target=camera_loop, daemon=True).start()
    threading.Thread(target=voice_loop,  daemon=True).start()

    sys.exit(app.exec())