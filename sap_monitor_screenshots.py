#!/usr/bin/env python3
"""
SAP GUI Monitor — System Status, Kernel Information, and SM21.
No GUI Scripting required.

Requirements:
    pip install pywin32 pyautogui python-docx Pillow

Usage:
    python sap_monitor_screenshots.py
    python sap_monitor_screenshots.py --sid PRD
"""

import argparse
import sys
import time
import tempfile
from datetime import datetime
from pathlib import Path


# ═══════════════════════════════════════════════════════════════════
#  CONFIGURATION
# ═══════════════════════════════════════════════════════════════════

DEFAULT_OUTPUT_DIR = r"C:\Users\I543394\OneDrive - SAP SE\SCRIPT_local\saP GUI MON"

WAIT_AFTER_NAV   = 2.5
SCREENSHOT_DELAY = 0.5

# Pixel offset of the SAP command field relative to the SAP window top-left
CMD_FIELD_X = 90
CMD_FIELD_Y = 55

# Pixel offset of the "System" menu item relative to the SAP window top-left
SYSTEM_MENU_X = 322
SYSTEM_MENU_Y = 23
GOTO_MENU_X   = 240   # "Goto" menu item — derived: Edit≈200, System≈322, Goto≈200+40

# ST06 OS Monitor — left-panel tree click offsets (relative to SAP window top-left).
# Measured from a full-screen SAP window; adjust if window layout differs.
ST06_TREE_X       = 168   # x: horizontal centre of the Snapshot tree items
ST06_SYSINFO_Y    = 458   # y: "System information" row under Snapshot
ST06_FILESYSTEM_Y = 569   # y: "Filesystem" row under Snapshot

# DBACOCKPIT — left-panel tree click offsets (relative to SAP window top-left).
DBACP_TREE_X           = 90    # x: on the text/bullet area (short items like "Overview" end before x=151)
DBACP_CURRENT_STATUS_Y = 303   # y: "Current Status"  (confirmed: screen 294)
DBACP_OVERVIEW_Y       = 320   # y: "Overview"        (confirmed: screen 311)
DBACP_ALERTS_Y         = 337   # y: "Alerts"          (confirmed: screen 328)
DBACP_BACKUP_Y         = 452   # y: "Backup"          (screen 443 — Diagnostics confirmed at 421, +22px)
DBACP_BACKUP_CATALOG_Y = 469   # y: "Backup Catalog"  (screen 460 — one row below Backup)

# ═══════════════════════════════════════════════════════════════════


def _check_dependencies():
    missing = []
    for pkg, mod in [
        ("pywin32",     "win32gui"),
        ("pyautogui",   "pyautogui"),
        ("python-docx", "docx"),
        ("Pillow",      "PIL"),
    ]:
        try:
            __import__(mod)
        except ImportError:
            missing.append(pkg)
    if missing:
        print("\n[ERROR] Missing packages. Install with:")
        print(f"  pip install {' '.join(missing)}\n")
        sys.exit(1)


_check_dependencies()

import win32gui
import win32con
import pyautogui
from PIL import ImageGrab
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

pyautogui.PAUSE = 0.05
pyautogui.FAILSAFE = False


# ─────────────────────────────────────────────────────────────────
# SAP window discovery
# ─────────────────────────────────────────────────────────────────

def find_sap_window():
    candidates = []

    def _cb(hwnd, _):
        if not win32gui.IsWindowVisible(hwnd):
            return
        title = win32gui.GetWindowText(hwnd)
        if "SAP" in title and "SAP Logon" not in title:
            candidates.append((hwnd, title))

    win32gui.EnumWindows(_cb, None)

    if not candidates:
        print("\n[ERROR] No SAP GUI window found. Make sure SAP GUI is open and logged in.\n")
        sys.exit(1)

    print(f"  SAP windows found ({len(candidates)}):")
    for h, t in candidates:
        print(f"    hwnd={h}  title={t!r}")

    for h, t in candidates:
        if "Easy Access" in t:
            print(f"  Using: {t!r}")
            return h

    hwnd, title = candidates[0]
    print(f"  Using: {title!r}")
    return hwnd


# ─────────────────────────────────────────────────────────────────
# Window / focus helpers
# ─────────────────────────────────────────────────────────────────

def _focus_and_restore(hwnd):
    placement = win32gui.GetWindowPlacement(hwnd)
    if placement[1] == win32con.SW_SHOWMINIMIZED:
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        time.sleep(0.3)
    win32gui.SetForegroundWindow(hwnd)
    time.sleep(0.3)


def _visible_top_windows():
    hwnds = set()
    def _cb(h, _):
        if win32gui.IsWindowVisible(h) and win32gui.GetWindowText(h):
            hwnds.add(h)
    win32gui.EnumWindows(_cb, None)
    return hwnds


def _wait_new_window(before, timeout=10):
    deadline = time.time() + timeout
    while time.time() < deadline:
        new = _visible_top_windows() - before
        for h in new:
            if win32gui.GetClassName(h) != "#32768" and win32gui.IsWindow(h):
                return h
        time.sleep(0.2)
    return None


def _find_child_by_text(hwnd, text, max_y_from_top=60):
    wr     = win32gui.GetWindowRect(hwnd)
    hit    = []
    needle = text.lower()

    def _cb(h, _):
        # Strip & accelerator markers (SAP stores e.g. "Go&to") before comparing
        raw = win32gui.GetWindowText(h).strip().replace("&", "").lower()
        if raw == needle:
            r = win32gui.GetWindowRect(h)
            if (r[1] - wr[1]) < max_y_from_top:
                hit.append(((r[0] + r[2]) // 2, (r[1] + r[3]) // 2))

    try:
        win32gui.EnumChildWindows(hwnd, _cb, None)
    except Exception:
        pass
    return hit[0] if hit else None


def _wait_title_stable(hwnd, timeout=6):
    prev = ""
    same = 0
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            title = win32gui.GetWindowText(hwnd)
        except Exception:
            break
        if title == prev:
            same += 1
            if same >= 2:
                return
        else:
            same = 0
            prev = title
        time.sleep(0.5)
    time.sleep(0.5)


def _wait_title_changed_then_stable(hwnd, old_title, timeout=10):
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            title = win32gui.GetWindowText(hwnd)
        except Exception:
            break
        if title != old_title:
            break
        time.sleep(0.3)
    _wait_title_stable(hwnd, timeout=max(2.0, deadline - time.time()))


# ─────────────────────────────────────────────────────────────────
# Command field + t-code navigation
# ─────────────────────────────────────────────────────────────────

def _click_command_field(hwnd):
    rect = win32gui.GetWindowRect(hwnd)
    print(f"  [CMD] Window rect: ({rect[0]},{rect[1]},{rect[2]},{rect[3]})")

    # Move to a neutral area (bottom-right of SAP window) first so ALV grid
    # releases any mouse capture before we attempt to click the command field.
    neutral_x = rect[0] + (rect[2] - rect[0]) - 80
    neutral_y = rect[1] + (rect[3] - rect[1]) - 60
    pyautogui.moveTo(neutral_x, neutral_y, duration=0.15)
    time.sleep(0.25)

    x = rect[0] + CMD_FIELD_X
    y = rect[1] + CMD_FIELD_Y
    print(f"  [CMD] Clicking command field at ({x}, {y})")
    pyautogui.click(x, y)
    time.sleep(0.2)


def navigate_tcode(hwnd, tcode):
    _focus_and_restore(hwnd)
    _click_command_field(hwnd)
    old_title = win32gui.GetWindowText(hwnd)
    pyautogui.hotkey("ctrl", "a")
    time.sleep(0.1)
    pyautogui.typewrite(f"/n{tcode}", interval=0.05)
    pyautogui.press("enter")
    _wait_title_changed_then_stable(hwnd, old_title, timeout=WAIT_AFTER_NAV + 5)


# ─────────────────────────────────────────────────────────────────
# Open System > Status...
# ─────────────────────────────────────────────────────────────────

def _open_system_status(hwnd):
    _focus_and_restore(hwnd)
    wr = win32gui.GetWindowRect(hwnd)

    pos = _find_child_by_text(hwnd, "System", max_y_from_top=60)
    if pos:
        sx, sy = pos
        print(f"  [Menu] 'System' child window → ({sx}, {sy})")
    else:
        sx = wr[0] + SYSTEM_MENU_X
        sy = wr[1] + SYSTEM_MENU_Y
        print(f"  [Menu] Pixel fallback → 'System' at ({sx}, {sy})")

    pyautogui.click(sx, sy)
    time.sleep(1.5)

    print("  [Menu] Down×11 + Enter → 'Status...'")
    for _ in range(11):
        pyautogui.press("down")
        time.sleep(0.08)
    pyautogui.press("enter")


# ─────────────────────────────────────────────────────────────────
# Capture System Status + Kernel Information
# ─────────────────────────────────────────────────────────────────

def capture_status_and_kernel(main_hwnd, doc, tmp_dir, img_index):
    print("  Opening System > Status...")
    before = _visible_top_windows()
    try:
        _open_system_status(main_hwnd)
    except Exception as exc:
        print(f"  [WARN] Could not open System > Status: {exc}")
        return img_index

    status_hwnd = _wait_new_window(before, timeout=10)
    if not status_hwnd or not win32gui.IsWindow(status_hwnd):
        print("  [WARN] Status popup did not appear — skipping")
        pyautogui.press("escape")
        return img_index

    time.sleep(1.5)

    # Screenshot 1: Status popup
    img1 = tmp_dir / f"{img_index:02d}_SYSTEM_STATUS.png"
    try:
        _focus_and_restore(status_hwnd)
        time.sleep(SCREENSHOT_DELAY)
        ImageGrab.grab().save(str(img1))
        print("  [OK] Status screenshot saved")
    except Exception as exc:
        print(f"  [WARN] Status screenshot failed: {exc}")
        img1 = None

    _add_block(doc, "System > Status", "System Status", str(img1) if img1 else None)
    img_index += 1

    # Open Kernel info via Shift+F5
    before2 = _visible_top_windows()
    _focus_and_restore(status_hwnd)
    time.sleep(0.4)
    print("  Opening kernel info via Shift+F5...")
    pyautogui.hotkey("shift", "f5")
    time.sleep(2.5)

    # Screenshot 2: Kernel popup
    kernel_hwnd = _wait_new_window(before2, timeout=8)
    if kernel_hwnd and win32gui.IsWindow(kernel_hwnd):
        img2 = tmp_dir / f"{img_index:02d}_SYSTEM_KERNEL.png"
        try:
            _focus_and_restore(kernel_hwnd)
            time.sleep(SCREENSHOT_DELAY)
            ImageGrab.grab().save(str(img2))
            print("  [OK] Kernel screenshot saved")
        except Exception as exc:
            print(f"  [WARN] Kernel screenshot failed: {exc}")
            img2 = None

        _add_block(doc, "System > Status > Navigate", "Kernel Information",
                   str(img2) if img2 else None)
        img_index += 1

        try:
            _focus_and_restore(kernel_hwnd)
            pyautogui.press("enter")
            time.sleep(0.6)
        except Exception:
            pass
    else:
        print("  [WARN] Kernel info popup did not appear")

    # Close Status popup
    try:
        if win32gui.IsWindow(status_hwnd):
            _focus_and_restore(status_hwnd)
            pyautogui.press("enter")
            time.sleep(0.5)
    except Exception:
        pass

    return img_index


# ─────────────────────────────────────────────────────────────────
# SM21 — fill date/time fields and execute
# ─────────────────────────────────────────────────────────────────

def _fill_sap_field(value):
    """Select the entire content of the current SAP input field and replace it.
    Uses Home+Shift+End instead of Ctrl+A because Ctrl+A in SAP GUI selects
    all screen objects, not just the text in the focused field.
    """
    pyautogui.press("home")
    time.sleep(0.05)
    pyautogui.hotkey("shift", "end")
    time.sleep(0.05)
    pyautogui.typewrite(value, interval=0.05)
    time.sleep(0.1)


def capture_sm21(hwnd, doc, tmp_dir, img_index):
    print("  [SM21] Navigating to SM21...")
    navigate_tcode(hwnd, "SM21")
    time.sleep(1.5)   # let the screen fully render

    today = datetime.now().strftime("%d.%m.%y")   # DD.MM.YY as shown in SAP
    print(f"  [SM21] Filling: From/To Date={today}, Time=00:00:00/23:50:00, ExtInst=*")

    _focus_and_restore(hwnd)

    # After /nSM21 the cursor lands on From Date (first input field).
    # Tab order on this screen:
    #   0  From Date
    #   1  From Time
    #   2  To Date
    #   3  To Time
    #   4  Message ID (from)
    #   5  Message ID (to)
    #   6  Client (from)
    #   7  Client (to)
    #   8  User (from)
    #   9  User (to)
    #   10 High severity checkbox
    #   11 Extended Instance Name (from)

    def _tab():
        pyautogui.press("tab")
        time.sleep(0.15)

    _fill_sap_field(today)    # From Date
    _tab()
    _fill_sap_field("000000") # From Time  → SAP formats to 00:00:00
    _tab()
    _fill_sap_field(today)    # To Date
    _tab()
    _fill_sap_field("235000") # To Time    → SAP formats to 23:50:00

    # Tab order after To Time (arrow buttons on each row ARE tab stops):
    #  1 Msg ID(from)  2 Msg ID(to)  3 Msg ID arrow
    #  4 Client(from)  5 Client(to)  6 Client arrow
    #  7 User(from)    8 User(to)    9 User arrow
    # 10 High-severity checkbox
    # 11 Extended Instance Name (from)
    for _ in range(11):
        _tab()

    _fill_sap_field("*")      # Extended Instance Name

    # Execute with F8 — capture title BEFORE pressing so we can detect when
    # the results screen has fully loaded (title may or may not change).
    print("  [SM21] Pressing F8 to execute...")
    _focus_and_restore(hwnd)
    old_title = win32gui.GetWindowText(hwnd)
    pyautogui.press("f8")
    _wait_title_changed_then_stable(hwnd, old_title, timeout=15)
    time.sleep(1.0)

    # Screenshot
    img_path = tmp_dir / f"{img_index:02d}_SM21.png"
    try:
        _focus_and_restore(hwnd)
        time.sleep(SCREENSHOT_DELAY)
        ImageGrab.grab().save(str(img_path))
        print("  [OK] SM21 screenshot saved")
    except Exception as exc:
        print(f"  [WARN] SM21 screenshot failed: {exc}")
        img_path = None

    _add_block(doc, "SM21", "System Log", str(img_path) if img_path else None)

    # Press F3 twice to navigate back to Easy Access so the next tcode
    # navigation lands reliably on the command field.
    _focus_and_restore(hwnd)
    time.sleep(0.3)
    print("  [SM21] Pressing F3 to go back...")
    pyautogui.press("f3")
    time.sleep(1.2)
    print("  [SM21] Pressing F3 again...")
    pyautogui.press("f3")
    time.sleep(1.2)

    return img_index + 1


# ─────────────────────────────────────────────────────────────────
# SM37 — Background Job Overview
# ─────────────────────────────────────────────────────────────────

def capture_sm37(hwnd, doc, tmp_dir, img_index):
    print("  [SM37] Navigating to SM37...")
    navigate_tcode(hwnd, "SM37")
    time.sleep(1.5)

    _focus_and_restore(hwnd)

    def _tab():
        pyautogui.press("tab")
        time.sleep(0.15)

    # Cursor lands on Job Name after fresh SM37 navigation
    print("  [SM37] Setting Job Name = *")
    _fill_sap_field("*")

    _tab()
    print("  [SM37] Setting User Name = *")
    _fill_sap_field("*")

    # ── Job Status checkboxes ───────────────────────────────────
    # SM37 default state: Sched=off, Released=ON, Ready=ON,
    #                     Active=ON, Finished=ON, Cancelled=ON
    # Desired state:      only Active + Cancelled checked.
    # Action: uncheck Released, Ready, Finished (1 Space each).
    #         Leave Sched, Active, Cancelled untouched.
    print("  [SM37] Configuring job status checkboxes...")
    _tab()                          # → Sched.    (default off → leave)
    _tab()                          # → Released  (default ON  → uncheck)
    pyautogui.press("space")
    time.sleep(0.1)
    _tab()                          # → Ready     (default ON  → uncheck)
    pyautogui.press("space")
    time.sleep(0.1)
    _tab()                          # → Active    (default ON  → leave)
    _tab()                          # → Finished  (default ON  → uncheck)
    pyautogui.press("space")
    time.sleep(0.1)
    _tab()                          # → Cancelled (default ON  → leave)

    # Execute with F8
    print("  [SM37] Pressing F8 to execute...")
    _focus_and_restore(hwnd)
    old_title = win32gui.GetWindowText(hwnd)
    pyautogui.press("f8")
    _wait_title_changed_then_stable(hwnd, old_title, timeout=20)
    time.sleep(2.0)

    # Screenshot
    img_path = tmp_dir / f"{img_index:02d}_SM37.png"
    try:
        _focus_and_restore(hwnd)
        time.sleep(SCREENSHOT_DELAY)
        ImageGrab.grab().save(str(img_path))
        print("  [OK] SM37 screenshot saved")
    except Exception as exc:
        print(f"  [WARN] SM37 screenshot failed: {exc}")
        img_path = None

    _add_block(doc, "SM37", "Background Job Overview", str(img_path) if img_path else None)

    # Press F3 twice to return to Easy Access before next tcode navigation.
    _focus_and_restore(hwnd)
    time.sleep(0.3)
    print("  [SM37] Pressing F3 to go back...")
    pyautogui.press("f3")
    time.sleep(1.2)
    print("  [SM37] Pressing F3 again...")
    pyautogui.press("f3")
    time.sleep(1.2)

    return img_index + 1


# ─────────────────────────────────────────────────────────────────
# SM13 — Update Records
# ─────────────────────────────────────────────────────────────────

def capture_sm13(hwnd, doc, tmp_dir, img_index):
    print("  [SM13] Navigating to SM13...")
    navigate_tcode(hwnd, "SM13")
    time.sleep(1.5)

    _focus_and_restore(hwnd)
    today = datetime.now().strftime("%d.%m.%y")

    def _tab():
        pyautogui.press("tab")
        time.sleep(0.15)

    # SM13 tab order (cursor starts on User):
    #   User → [Status radio group — 1 tab stop, skip] → From date
    #   → To date → From time → To time
    # Client is left as-is (* by default). Status radio untouched.
    print("  [SM13] Filling selection criteria...")
    _fill_sap_field("*")       # User
    _tab()                      # → Status radio group (skip — don't fill)
    _tab()                      # → From date
    _fill_sap_field(today)     # From date
    _tab()                      # → To date
    _fill_sap_field(today)     # To date
    _tab()                      # → From time
    _fill_sap_field("000000")  # From time
    _tab()                      # → To time
    _fill_sap_field("235000")  # To time

    print("  [SM13] Pressing F8 to execute...")
    _focus_and_restore(hwnd)
    old_title = win32gui.GetWindowText(hwnd)
    pyautogui.press("f8")
    _wait_title_changed_then_stable(hwnd, old_title, timeout=20)
    time.sleep(2.0)

    img_path = tmp_dir / f"{img_index:02d}_SM13.png"
    try:
        _focus_and_restore(hwnd)
        time.sleep(SCREENSHOT_DELAY)
        ImageGrab.grab().save(str(img_path))
        print("  [OK] SM13 screenshot saved")
    except Exception as exc:
        print(f"  [WARN] SM13 screenshot failed: {exc}")
        img_path = None

    _add_block(doc, "SM13", "Update Records", str(img_path) if img_path else None)

    # Press F3 twice to return to Easy Access before next tcode navigation.
    _focus_and_restore(hwnd)
    time.sleep(0.3)
    print("  [SM13] Pressing F3 to go back...")
    pyautogui.press("f3")
    time.sleep(1.2)
    print("  [SM13] Pressing F3 again...")
    pyautogui.press("f3")
    time.sleep(1.2)

    return img_index + 1


# ─────────────────────────────────────────────────────────────────
# SM12 — Enqueue Administration (Lock Table)
# ─────────────────────────────────────────────────────────────────

def capture_sm12(hwnd, doc, tmp_dir, img_index):
    print("  [SM12] Navigating to SM12...")
    navigate_tcode(hwnd, "SM12")
    time.sleep(1.5)

    _focus_and_restore(hwnd)

    # Cursor lands on Client when SM12 opens — skip it with Tab,
    # then fill only User Name. All other fields left at defaults.
    print("  [SM12] Setting User Name = *")
    pyautogui.press("tab")     # Client → User Name
    time.sleep(0.15)
    _fill_sap_field("*")       # User Name

    print("  [SM12] Pressing F8 to execute...")
    _focus_and_restore(hwnd)
    old_title = win32gui.GetWindowText(hwnd)
    pyautogui.press("f8")
    _wait_title_changed_then_stable(hwnd, old_title, timeout=20)
    time.sleep(2.0)

    img_path = tmp_dir / f"{img_index:02d}_SM12.png"
    try:
        _focus_and_restore(hwnd)
        time.sleep(SCREENSHOT_DELAY)
        ImageGrab.grab().save(str(img_path))
        print("  [OK] SM12 screenshot saved")
    except Exception as exc:
        print(f"  [WARN] SM12 screenshot failed: {exc}")
        img_path = None

    _add_block(doc, "SM12", "Enqueue Administration", str(img_path) if img_path else None)

    # Press F3 twice to return to Easy Access before next tcode navigation.
    _focus_and_restore(hwnd)
    time.sleep(0.3)
    print("  [SM12] Pressing F3 to go back...")
    pyautogui.press("f3")
    time.sleep(1.2)
    print("  [SM12] Pressing F3 again...")
    pyautogui.press("f3")
    time.sleep(1.2)

    return img_index + 1


# ─────────────────────────────────────────────────────────────────
# SM58 — Transactional RFC
# ─────────────────────────────────────────────────────────────────

def capture_sm58(hwnd, doc, tmp_dir, img_index):
    print("  [SM58] Navigating to SM58...")
    navigate_tcode(hwnd, "SM58")
    time.sleep(1.5)

    _focus_and_restore(hwnd)

    # Cursor lands on Display Period (from). Tab order:
    #   from → Display Period to → [→ row button] → User Name from
    # Use End+Shift+Home to select within the FROM field (Shift+End
    # jumps to the TO field in this screen, so we go backwards instead).
    print("  [SM58] Setting User Name = *")
    pyautogui.press("tab")     # → Display Period to
    time.sleep(0.15)
    pyautogui.press("tab")     # → Display Period row arrow button
    time.sleep(0.15)
    pyautogui.press("tab")     # → User Name from
    time.sleep(0.15)
    pyautogui.press("end")
    time.sleep(0.05)
    pyautogui.hotkey("shift", "home")
    time.sleep(0.05)
    pyautogui.typewrite("*", interval=0.05)

    print("  [SM58] Pressing F8 to execute...")
    _focus_and_restore(hwnd)
    old_title = win32gui.GetWindowText(hwnd)
    pyautogui.press("f8")
    _wait_title_changed_then_stable(hwnd, old_title, timeout=20)
    time.sleep(2.0)

    img_path = tmp_dir / f"{img_index:02d}_SM58.png"
    try:
        _focus_and_restore(hwnd)
        time.sleep(SCREENSHOT_DELAY)
        ImageGrab.grab().save(str(img_path))
        print("  [OK] SM58 screenshot saved")
    except Exception as exc:
        print(f"  [WARN] SM58 screenshot failed: {exc}")
        img_path = None

    _add_block(doc, "SM58", "Transactional RFC", str(img_path) if img_path else None)

    # Press F3 twice to return to Easy Access before next tcode navigation.
    _focus_and_restore(hwnd)
    time.sleep(0.3)
    print("  [SM58] Pressing F3 to go back...")
    pyautogui.press("f3")
    time.sleep(1.2)
    print("  [SM58] Pressing F3 again...")
    pyautogui.press("f3")
    time.sleep(1.2)

    return img_index + 1


# ─────────────────────────────────────────────────────────────────
# SMQ1 — qRFC Monitor (Outbound Queue)
# ─────────────────────────────────────────────────────────────────

def capture_smq1(hwnd, doc, tmp_dir, img_index):
    print("  [SMQ1] Navigating to SMQ1...")
    navigate_tcode(hwnd, "SMQ1")
    time.sleep(1.5)

    print("  [SMQ1] Pressing F8 to execute...")
    _focus_and_restore(hwnd)
    old_title = win32gui.GetWindowText(hwnd)
    pyautogui.press("f8")
    _wait_title_changed_then_stable(hwnd, old_title, timeout=20)
    time.sleep(2.0)

    img_path = tmp_dir / f"{img_index:02d}_SMQ1.png"
    try:
        _focus_and_restore(hwnd)
        time.sleep(SCREENSHOT_DELAY)
        ImageGrab.grab().save(str(img_path))
        print("  [OK] SMQ1 screenshot saved")
    except Exception as exc:
        print(f"  [WARN] SMQ1 screenshot failed: {exc}")
        img_path = None

    _add_block(doc, "SMQ1", "qRFC Monitor (Outbound)", str(img_path) if img_path else None)
    return img_index + 1


# ─────────────────────────────────────────────────────────────────
# SMQ2 — qRFC Monitor (Inbound Queue)
# ─────────────────────────────────────────────────────────────────

def capture_smq2(hwnd, doc, tmp_dir, img_index):
    print("  [SMQ2] Navigating to SMQ2...")
    navigate_tcode(hwnd, "SMQ2")
    time.sleep(1.5)

    print("  [SMQ2] Pressing F8 to execute...")
    _focus_and_restore(hwnd)
    old_title = win32gui.GetWindowText(hwnd)
    pyautogui.press("f8")
    _wait_title_changed_then_stable(hwnd, old_title, timeout=20)
    time.sleep(2.0)

    img_path = tmp_dir / f"{img_index:02d}_SMQ2.png"
    try:
        _focus_and_restore(hwnd)
        time.sleep(SCREENSHOT_DELAY)
        ImageGrab.grab().save(str(img_path))
        print("  [OK] SMQ2 screenshot saved")
    except Exception as exc:
        print(f"  [WARN] SMQ2 screenshot failed: {exc}")
        img_path = None

    _add_block(doc, "SMQ2", "qRFC Monitor (Inbound)", str(img_path) if img_path else None)
    return img_index + 1


# ─────────────────────────────────────────────────────────────────
# SMQS — qRFC Scheduler  +  Goto > qRFC Resources
# ─────────────────────────────────────────────────────────────────

def capture_smqs(hwnd, doc, tmp_dir, img_index):
    print("  [SMQS] Navigating to SMQS...")
    navigate_tcode(hwnd, "SMQS")
    time.sleep(2.0)

    # Screenshot 1: SMQS scheduler main screen
    img_path1 = tmp_dir / f"{img_index:02d}_SMQS.png"
    try:
        _focus_and_restore(hwnd)
        time.sleep(SCREENSHOT_DELAY)
        ImageGrab.grab().save(str(img_path1))
        print("  [OK] SMQS screenshot saved")
    except Exception as exc:
        print(f"  [WARN] SMQS screenshot failed: {exc}")
        img_path1 = None

    _add_block(doc, "SMQS", "qRFC Scheduler", str(img_path1) if img_path1 else None)

    # Open Goto menu via Alt+G (direct keyboard accelerator — avoids pixel issues).
    # After the menu opens, Down×2 lands on qRFC Resources (2nd item).
    _focus_and_restore(hwnd)
    time.sleep(0.5)
    print("  [Menu] Alt+G → open Goto menu")
    pyautogui.hotkey("alt", "g")
    time.sleep(1.5)          # wait for dropdown to fully appear

    print("  [Menu] Down×2 + Enter → 'qRFC Resources'")
    pyautogui.press("down")
    time.sleep(0.2)
    pyautogui.press("down")
    time.sleep(0.2)
    pyautogui.press("enter")
    time.sleep(2.0)

    # Screenshot 2: qRFC Resources view
    img_path2 = tmp_dir / f"{img_index + 1:02d}_SMQS_Resources.png"
    try:
        _focus_and_restore(hwnd)
        time.sleep(SCREENSHOT_DELAY)
        ImageGrab.grab().save(str(img_path2))
        print("  [OK] SMQS qRFC Resources screenshot saved")
    except Exception as exc:
        print(f"  [WARN] SMQS qRFC Resources screenshot failed: {exc}")
        img_path2 = None

    _add_block(doc, "SMQS", "qRFC Scheduler – Resources",
               str(img_path2) if img_path2 else None)

    return img_index + 2


# ─────────────────────────────────────────────────────────────────
# ST06 — OS Monitor (3 screenshots: main, Snapshot>SysInfo, Snapshot>Filesystem)
# ─────────────────────────────────────────────────────────────────

def capture_st06(hwnd, doc, tmp_dir, img_index):
    print("  [ST06] Navigating to ST06...")
    navigate_tcode(hwnd, "ST06")
    time.sleep(2.0)

    # Screenshot 1: ST06 main screen
    img_path1 = tmp_dir / f"{img_index:02d}_ST06.png"
    try:
        _focus_and_restore(hwnd)
        time.sleep(SCREENSHOT_DELAY)
        ImageGrab.grab().save(str(img_path1))
        print("  [OK] ST06 main screenshot saved")
    except Exception as exc:
        print(f"  [WARN] ST06 screenshot failed: {exc}")
        img_path1 = None

    _add_block(doc, "ST06", "OS Monitor", str(img_path1) if img_path1 else None)

    # Click "System information" under Snapshot in the left tree panel
    rect = win32gui.GetWindowRect(hwnd)
    _focus_and_restore(hwnd)
    time.sleep(0.3)
    si_x = rect[0] + ST06_TREE_X
    si_y = rect[1] + ST06_SYSINFO_Y
    print(f"  [ST06] Clicking 'System information' at ({si_x}, {si_y})")
    pyautogui.doubleClick(si_x, si_y)
    time.sleep(2.5)

    # Screenshot 2: Snapshot > System information
    img_path2 = tmp_dir / f"{img_index + 1:02d}_ST06_SysInfo.png"
    try:
        _focus_and_restore(hwnd)
        time.sleep(SCREENSHOT_DELAY)
        ImageGrab.grab().save(str(img_path2))
        print("  [OK] ST06 System information screenshot saved")
    except Exception as exc:
        print(f"  [WARN] ST06 System information screenshot failed: {exc}")
        img_path2 = None

    _add_block(doc, "ST06", "OS Monitor – Snapshot > System Information",
               str(img_path2) if img_path2 else None)

    # Click "Filesystem" under Snapshot in the left tree panel
    _focus_and_restore(hwnd)
    time.sleep(0.3)
    fs_x = rect[0] + ST06_TREE_X
    fs_y = rect[1] + ST06_FILESYSTEM_Y
    print(f"  [ST06] Clicking 'Filesystem' at ({fs_x}, {fs_y})")
    pyautogui.doubleClick(fs_x, fs_y)
    time.sleep(2.5)

    # Screenshot 3: Snapshot > Filesystem
    img_path3 = tmp_dir / f"{img_index + 2:02d}_ST06_Filesystem.png"
    try:
        _focus_and_restore(hwnd)
        time.sleep(SCREENSHOT_DELAY)
        ImageGrab.grab().save(str(img_path3))
        print("  [OK] ST06 Filesystem screenshot saved")
    except Exception as exc:
        print(f"  [WARN] ST06 Filesystem screenshot failed: {exc}")
        img_path3 = None

    _add_block(doc, "ST06", "OS Monitor – Snapshot > Filesystem",
               str(img_path3) if img_path3 else None)

    # F3 twice to return to Easy Access before next tcode navigation.
    _focus_and_restore(hwnd)
    time.sleep(0.3)
    print("  [ST06] Pressing F3 to go back...")
    pyautogui.press("f3")
    time.sleep(1.2)
    print("  [ST06] Pressing F3 again...")
    pyautogui.press("f3")
    time.sleep(1.2)

    return img_index + 3


# ─────────────────────────────────────────────────────────────────
# DBACOCKPIT — Database Administration Cockpit
#   Screenshots: Current Status > Overview, Current Status > Alerts,
#                Backup > Backup Catalog
# ─────────────────────────────────────────────────────────────────

def capture_dbacockpit(hwnd, doc, tmp_dir, img_index):
    print("  [DBACOCKPIT] Navigating to DBACOCKPIT...")
    navigate_tcode(hwnd, "DBACOCKPIT")
    time.sleep(3.0)   # web-based cockpit needs extra load time

    rect = win32gui.GetWindowRect(hwnd)

    def _dbl(x_off, y_off, label):
        _focus_and_restore(hwnd)
        time.sleep(0.3)
        sx = rect[0] + x_off
        sy = rect[1] + y_off
        print(f"  [DBACOCKPIT] Double-clicking '{label}' at ({sx}, {sy})")
        pyautogui.doubleClick(sx, sy)
        time.sleep(4.0)   # DBACOCKPIT is web-based — needs time to load content

    # Expand "Current Status", wait for children to render, then select sub-items
    _dbl(DBACP_TREE_X, DBACP_CURRENT_STATUS_Y, "Current Status")
    time.sleep(2.0)
    _dbl(DBACP_TREE_X, DBACP_OVERVIEW_Y, "Overview")

    img_path1 = tmp_dir / f"{img_index:02d}_DBACOCKPIT_Overview.png"
    try:
        _focus_and_restore(hwnd)
        time.sleep(SCREENSHOT_DELAY)
        ImageGrab.grab().save(str(img_path1))
        print("  [OK] DBACOCKPIT Overview screenshot saved")
    except Exception as exc:
        print(f"  [WARN] DBACOCKPIT Overview screenshot failed: {exc}")
        img_path1 = None

    _add_block(doc, "DBACOCKPIT", "Database Administration – Current Status > Overview",
               str(img_path1) if img_path1 else None)

    # Select "Alerts"
    _dbl(DBACP_TREE_X, DBACP_ALERTS_Y, "Alerts")

    img_path2 = tmp_dir / f"{img_index + 1:02d}_DBACOCKPIT_Alerts.png"
    try:
        _focus_and_restore(hwnd)
        time.sleep(SCREENSHOT_DELAY)
        ImageGrab.grab().save(str(img_path2))
        print("  [OK] DBACOCKPIT Alerts screenshot saved")
    except Exception as exc:
        print(f"  [WARN] DBACOCKPIT Alerts screenshot failed: {exc}")
        img_path2 = None

    _add_block(doc, "DBACOCKPIT", "Database Administration – Current Status > Alerts",
               str(img_path2) if img_path2 else None)

    # Expand "Backup", wait for children to render, then select "Backup Catalog"
    _dbl(DBACP_TREE_X, DBACP_BACKUP_Y, "Backup")
    time.sleep(2.0)
    _dbl(DBACP_TREE_X, DBACP_BACKUP_CATALOG_Y, "Backup Catalog")

    img_path3 = tmp_dir / f"{img_index + 2:02d}_DBACOCKPIT_BackupCatalog.png"
    try:
        _focus_and_restore(hwnd)
        time.sleep(SCREENSHOT_DELAY)
        ImageGrab.grab().save(str(img_path3))
        print("  [OK] DBACOCKPIT Backup Catalog screenshot saved")
    except Exception as exc:
        print(f"  [WARN] DBACOCKPIT Backup Catalog screenshot failed: {exc}")
        img_path3 = None

    _add_block(doc, "DBACOCKPIT", "Database Administration – Backup > Backup Catalog",
               str(img_path3) if img_path3 else None)

    # F3 twice to return to Easy Access
    _focus_and_restore(hwnd)
    time.sleep(0.3)
    print("  [DBACOCKPIT] Pressing F3 to go back...")
    pyautogui.press("f3")
    time.sleep(1.2)
    print("  [DBACOCKPIT] Pressing F3 again...")
    pyautogui.press("f3")
    time.sleep(1.2)

    return img_index + 3


# ─────────────────────────────────────────────────────────────────
# Generic helper — navigate, screenshot, F3×2 back
# ─────────────────────────────────────────────────────────────────

def _capture_simple(hwnd, doc, tmp_dir, img_index, tcode, description):
    print(f"  [{tcode}] Navigating to {tcode}...")
    navigate_tcode(hwnd, tcode)
    time.sleep(2.0)

    img_path = tmp_dir / f"{img_index:02d}_{tcode}.png"
    try:
        _focus_and_restore(hwnd)
        time.sleep(SCREENSHOT_DELAY)
        ImageGrab.grab().save(str(img_path))
        print(f"  [OK] {tcode} screenshot saved")
    except Exception as exc:
        print(f"  [WARN] {tcode} screenshot failed: {exc}")
        img_path = None

    _add_block(doc, tcode, description, str(img_path) if img_path else None)

    _focus_and_restore(hwnd)
    time.sleep(0.3)
    print(f"  [{tcode}] Pressing F3 to go back...")
    pyautogui.press("f3")
    time.sleep(1.2)
    pyautogui.press("f3")
    time.sleep(1.2)

    return img_index + 1


def capture_sm51(hwnd, doc, tmp_dir, img_index):
    return _capture_simple(hwnd, doc, tmp_dir, img_index, "SM51", "Application Server List")

def capture_sm50(hwnd, doc, tmp_dir, img_index):
    return _capture_simple(hwnd, doc, tmp_dir, img_index, "SM50", "Work Process Overview")

def capture_st22(hwnd, doc, tmp_dir, img_index):
    return _capture_simple(hwnd, doc, tmp_dir, img_index, "ST22", "ABAP Dump Analysis")

def capture_st02(hwnd, doc, tmp_dir, img_index):
    return _capture_simple(hwnd, doc, tmp_dir, img_index, "ST02", "Tune Summary / Buffer Information")

def capture_rz12(hwnd, doc, tmp_dir, img_index):
    return _capture_simple(hwnd, doc, tmp_dir, img_index, "RZ12", "RFC Server Group Maintenance")

def capture_smicm(hwnd, doc, tmp_dir, img_index):
    return _capture_simple(hwnd, doc, tmp_dir, img_index, "SMICM", "ICM Monitor")

def capture_smgw(hwnd, doc, tmp_dir, img_index):
    return _capture_simple(hwnd, doc, tmp_dir, img_index, "SMGW", "Gateway Monitor")

def capture_smms(hwnd, doc, tmp_dir, img_index):
    return _capture_simple(hwnd, doc, tmp_dir, img_index, "SMMS", "Message Server Monitor")

def capture_sick(hwnd, doc, tmp_dir, img_index):
    return _capture_simple(hwnd, doc, tmp_dir, img_index, "SICK", "Installation Check")

def capture_al08(hwnd, doc, tmp_dir, img_index):
    return _capture_simple(hwnd, doc, tmp_dir, img_index, "AL08", "Users Logged On (All Application Servers)")

def capture_sm04(hwnd, doc, tmp_dir, img_index):
    return _capture_simple(hwnd, doc, tmp_dir, img_index, "SM04", "User Overview")

def capture_sdf_smon(hwnd, doc, tmp_dir, img_index):
    return _capture_simple(hwnd, doc, tmp_dir, img_index, "/SDF/SMON", "SMON – System Monitor")

def capture_aif_err(hwnd, doc, tmp_dir, img_index):
    navigate_tcode(hwnd, "/AIF/ERR")
    time.sleep(2.0)
    pyautogui.press("f8")        # Execute
    time.sleep(2.5)
    img_path = tmp_dir / f"{img_index:02d}_AIF_ERR.png"
    ImageGrab.grab().save(str(img_path))
    _add_block(doc, "/AIF/ERR", "AIF Error Handling", str(img_path))
    pyautogui.press("f3"); time.sleep(1.2)
    pyautogui.press("f3"); time.sleep(1.2)
    return img_index + 1


# ─────────────────────────────────────────────────────────────────
# Word document helpers
# ─────────────────────────────────────────────────────────────────

def _create_document(sid):
    doc = Document()
    for sec in doc.sections:
        sec.top_margin    = Inches(0.75)
        sec.bottom_margin = Inches(0.75)
        sec.left_margin   = Inches(0.75)
        sec.right_margin  = Inches(0.75)

    for _ in range(3):
        doc.add_paragraph()

    title_para = doc.add_paragraph()
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title_para.add_run("SAP GUI Monitoring – Screenshots")
    run.bold = True
    run.font.size = Pt(24)
    run.font.color.rgb = RGBColor(0x1F, 0x4E, 0x79)

    doc.add_paragraph()

    meta_para = doc.add_paragraph()
    meta_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    meta_run = meta_para.add_run(
        f"System / SID:  {sid or '—'}"
        f"     |     Generated:  {datetime.now().strftime('%Y-%m-%d  %H:%M')}"
    )
    meta_run.font.size = Pt(12)
    meta_run.font.color.rgb = RGBColor(0x40, 0x40, 0x40)

    doc.add_page_break()
    return doc


def _add_block(doc, tcode, description, img_path):
    h = doc.add_paragraph()
    h.paragraph_format.space_before = Pt(8)
    h.paragraph_format.space_after  = Pt(2)

    r_code = h.add_run(tcode)
    r_code.bold = True
    r_code.font.size = Pt(14)
    r_code.font.color.rgb = RGBColor(0x1F, 0x4E, 0x79)

    r_sep = h.add_run("   —   ")
    r_sep.font.size = Pt(13)
    r_sep.font.color.rgb = RGBColor(0x70, 0x70, 0x70)

    r_desc = h.add_run(description)
    r_desc.font.size = Pt(13)
    r_desc.font.color.rgb = RGBColor(0x30, 0x30, 0x30)

    p   = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(5)
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"),   "single")
    bottom.set(qn("w:sz"),    "6")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), "1F4E79")
    pBdr.append(bottom)
    pPr.append(pBdr)

    if img_path and Path(img_path).exists():
        try:
            doc.add_picture(img_path, width=Inches(7.0))
            doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
        except Exception as exc:
            doc.add_paragraph(f"[Could not embed screenshot: {exc}]")
    else:
        doc.add_paragraph("[Screenshot not available]")

    doc.add_paragraph()


# ─────────────────────────────────────────────────────────────────
# Main
# ─────────────────────────────────────────────────────────────────

def run(sid="", output_path=None):
    output_dir = Path(DEFAULT_OUTPUT_DIR)
    output_dir.mkdir(parents=True, exist_ok=True)

    if not output_path:
        ts       = datetime.now().strftime("%Y%m%d_%H%M")
        sid_part = f"_{sid}" if sid else ""
        output_path = output_dir / f"SAP_Monitor{sid_part}_{ts}.docx"

    print("\n" + "=" * 62)
    print("  SAP GUI Monitor  (keyboard automation)")
    print("=" * 62)
    print(f"\n  SID    : {sid or '(auto)'}")
    print(f"  Output : {output_path}")
    print()
    print("  NOTE: Do not move the mouse or use the keyboard while")
    print("        the script is running.\n")

    hwnd    = find_sap_window()
    doc     = _create_document(sid)
    tmp_dir = Path(tempfile.mkdtemp(prefix="sap_shots_"))

    img_index = 1

    # Step 1: System Status + Kernel Information
    img_index = capture_status_and_kernel(hwnd, doc, tmp_dir, img_index)

    # Step 2: SM21 System Log with date/time fill + F8
    img_index = capture_sm21(hwnd, doc, tmp_dir, img_index)

    # Step 3: SM37 Background Job Overview
    img_index = capture_sm37(hwnd, doc, tmp_dir, img_index)

    # Step 4: SM13 Update Records
    img_index = capture_sm13(hwnd, doc, tmp_dir, img_index)

    # Step 5: SM12 Enqueue Administration (Lock Table)
    img_index = capture_sm12(hwnd, doc, tmp_dir, img_index)

    # Step 6: SM58 Transactional RFC
    img_index = capture_sm58(hwnd, doc, tmp_dir, img_index)

    # Step 7: SMQ1 qRFC Monitor Outbound
    img_index = capture_smq1(hwnd, doc, tmp_dir, img_index)

    # Step 8: SMQ2 qRFC Monitor Inbound
    img_index = capture_smq2(hwnd, doc, tmp_dir, img_index)

    # Step 9: SMQS qRFC Scheduler + Goto > qRFC Resources
    img_index = capture_smqs(hwnd, doc, tmp_dir, img_index)

    # Step 10: ST06 OS Monitor (main + Snapshot>System info + Snapshot>Filesystem)
    img_index = capture_st06(hwnd, doc, tmp_dir, img_index)

    # Step 11: DBACOCKPIT (Current Status>Overview, Current Status>Alerts, Backup>Backup Catalog)
    img_index = capture_dbacockpit(hwnd, doc, tmp_dir, img_index)

    # Step 12: SM51 Application Server List
    img_index = capture_sm51(hwnd, doc, tmp_dir, img_index)

    # Step 13: SM50 Work Process Overview
    img_index = capture_sm50(hwnd, doc, tmp_dir, img_index)

    # Step 14: ST22 ABAP Dump Analysis
    img_index = capture_st22(hwnd, doc, tmp_dir, img_index)

    # Step 15: ST02 Tune Summary / Buffer Information
    img_index = capture_st02(hwnd, doc, tmp_dir, img_index)

    # Step 16: RZ12 RFC Server Group Maintenance
    img_index = capture_rz12(hwnd, doc, tmp_dir, img_index)

    # Step 17: SMICM ICM Monitor
    img_index = capture_smicm(hwnd, doc, tmp_dir, img_index)

    # Step 18: SMGW Gateway Monitor
    img_index = capture_smgw(hwnd, doc, tmp_dir, img_index)

    # Step 19: SMMS Message Server Monitor
    img_index = capture_smms(hwnd, doc, tmp_dir, img_index)

    # Step 20: SICK Installation Check
    img_index = capture_sick(hwnd, doc, tmp_dir, img_index)

    # Step 21: AL08 Users Logged On
    img_index = capture_al08(hwnd, doc, tmp_dir, img_index)

    # Step 22: SM04 User Overview
    img_index = capture_sm04(hwnd, doc, tmp_dir, img_index)

    # Step 23: /SDF/SMON System Monitor
    img_index = capture_sdf_smon(hwnd, doc, tmp_dir, img_index)

    # Step 24: /AIF/ERR AIF Error Handling
    img_index = capture_aif_err(hwnd, doc, tmp_dir, img_index)

    doc.save(str(output_path))

    for f in tmp_dir.iterdir():
        try:
            f.unlink()
        except Exception:
            pass
    try:
        tmp_dir.rmdir()
    except Exception:
        pass

    print(f"\n{'=' * 62}")
    print(f"  COMPLETE — Word document: {output_path}")
    print(f"{'=' * 62}\n")
    return str(output_path)


def main():
    ap = argparse.ArgumentParser(
        description="SAP GUI Monitor — keyboard automation, no scripting needed"
    )
    ap.add_argument("tcode",    nargs="?", default=None,
                    help="Single t-code to test, e.g. SM12, SM58 (omit to run all)")
    ap.add_argument("--sid",    default="",   help="SAP System ID (e.g. PRD)")
    ap.add_argument("--output", default=None, help="Output .docx file path")
    args = ap.parse_args()

    if args.tcode:
        run_single(args.tcode.upper(), args.sid, args.output)
    else:
        run(args.sid, args.output)


def run_single(tcode, sid="", output_path=None):
    """Run capture for one t-code only — useful for testing."""
    dispatch = {
        "SM21": capture_sm21,
        "SM37": capture_sm37,
        "SM13": capture_sm13,
        "SM12": capture_sm12,
        "SM58": capture_sm58,
        "SMQ1": capture_smq1,
        "SMQ2": capture_smq2,
        "SMQS": capture_smqs,
        "ST06": capture_st06,
        "DBACOCKPIT": capture_dbacockpit,
        "SM51": capture_sm51,
        "SM50": capture_sm50,
        "ST22": capture_st22,
        "ST02": capture_st02,
        "RZ12": capture_rz12,
        "SMICM": capture_smicm,
        "SMGW": capture_smgw,
        "SMMS": capture_smms,
        "SICK": capture_sick,
        "AL08": capture_al08,
        "SM04": capture_sm04,
        "/SDF/SMON": capture_sdf_smon,
        "/AIF/ERR": capture_aif_err,
    }
    if tcode not in dispatch:
        print(f"Unknown t-code '{tcode}'. Available: {', '.join(dispatch)}")
        return

    output_dir = Path(DEFAULT_OUTPUT_DIR)
    output_dir.mkdir(parents=True, exist_ok=True)
    if not output_path:
        ts = datetime.now().strftime("%Y%m%d_%H%M")
        sid_part = f"_{sid}" if sid else ""
        output_path = output_dir / f"SAP_Test_{tcode}{sid_part}_{ts}.docx"

    print("\n" + "=" * 62)
    print(f"  SAP GUI Monitor — single t-code test: {tcode}")
    print("=" * 62)
    print(f"\n  Output : {output_path}")
    print("  NOTE: Do not move the mouse or use the keyboard while")
    print("        the script is running.\n")

    hwnd    = find_sap_window()
    doc     = _create_document(sid)
    tmp_dir = Path(tempfile.mkdtemp(prefix="sap_shots_"))

    dispatch[tcode](hwnd, doc, tmp_dir, 1)

    doc.save(str(output_path))
    for f in tmp_dir.iterdir():
        try:
            f.unlink()
        except Exception:
            pass
    try:
        tmp_dir.rmdir()
    except Exception:
        pass

    print(f"\n{'=' * 62}")
    print(f"  COMPLETE — Word document: {output_path}")
    print(f"{'=' * 62}\n")


if __name__ == "__main__":
    main()
