#!/usr/bin/env python3
"""
SAP GUI Monitor Screenshot Script  (no-scripting edition)
Drives an open SAP GUI window with keyboard automation — no GUI Scripting
permission required.  Navigates each t-code, screenshots the window, and
saves everything into a Word document.

Requirements:
    pip install pywin32 pyautogui python-docx Pillow

Usage:
    python sap_monitor_screenshots.py
    python sap_monitor_screenshots.py --sid PRD
    python sap_monitor_screenshots.py --sid PRD --output "C:/out/monitor.docx"
    python sap_monitor_screenshots.py --tcodes SM50,SM51,ST22
"""

import argparse
import os
import sys
import time
import tempfile
from datetime import datetime
from pathlib import Path


# ═══════════════════════════════════════════════════════════════════
#  CONFIGURATION  —  Edit this section to customise behaviour
# ═══════════════════════════════════════════════════════════════════

DEFAULT_OUTPUT_DIR = r"C:\Users\I543394\OneDrive - SAP SE\SCRIPT_local\saP GUI MON"

# Seconds to wait after sending a t-code before taking the screenshot
WAIT_AFTER_NAV = 2.5

# Extra settle time just before the screenshot
SCREENSHOT_DELAY = 0.5

# Pixel offset of the SAP command field RELATIVE TO THE WINDOW'S top-left
# corner.  These defaults work for SAP GUI 7.x at normal DPI.  If the script
# keeps missing the field, adjust CMD_FIELD_X / CMD_FIELD_Y or run
# `python -c "import pyautogui; pyautogui.displayMousePosition()"` and hover
# over the command box to read the absolute coordinates, then set
# CMD_FIELD_ABSOLUTE = True and put those coordinates below.
CMD_FIELD_X        = 90    # pixels from window left
CMD_FIELD_Y        = 55    # pixels from window top
CMD_FIELD_ABSOLUTE = False  # True → treat X/Y as absolute screen coords

# ── System menu pixel offsets ──────────────────────────────────────
# Used as fallback when SAP doesn't expose menu items as child windows.
# All values are relative to the SAP MAIN window top-left corner.
# The script always prints the exact screen coordinates it clicks —
# use those to calibrate if a click misses.
SYSTEM_MENU_X   = 322   # center-x of the "System" menu item in the menu bar
SYSTEM_MENU_Y   = 23    # center-y of the menu bar row
# "Status..." offset FROM the System menu item click point:
STATUS_OFFSET_X = 60    # pixels right of System → lands in dropdown center
STATUS_OFFSET_Y = 282   # pixels below System → center of Status... row

# ── System Status capture ──────────────────────────────────────────
# Set to False to skip the System > Status section entirely.
CAPTURE_SYSTEM_STATUS = True

# Set to False to skip all t-code navigation (useful while testing Status only).
RUN_TCODES = False

# Pixel offset of the Navigate button INSIDE the Status popup, measured
# from the popup's top-left corner.  The button is the second icon in the
# SAP toolbar (green ✓ | Navigate | icon | ✗).
# Run with CAPTURE_SYSTEM_STATUS = True, look at the first status screenshot,
# hover over the Navigate button and note its screen position, then subtract
# the popup window's top-left to get the relative offset below.
STATUS_NAV_BTN_X = 130   # pixels from popup left edge
STATUS_NAV_BTN_Y = 48    # pixels from popup top  edge

# T-codes to capture: (T-Code, Description, press_F8_to_execute)
TCODES = [
    ("SM51",  "List of SAP Application Servers",            False),
    ("SM50",  "Work Process Overview (Local Server)",       False),
    ("SM66",  "Global Work Process Overview",               False),
    ("SM04",  "User Overview (Local)",                      False),
    ("AL08",  "Users Logged On – All Servers",              False),
    ("SM37",  "Background Job Overview",                    True ),
    ("SM13",  "Update Records",                             False),
    ("SM12",  "Lock Entries",                               False),
    ("SM21",  "System Log",                                 True ),
    ("ST22",  "ABAP Runtime Error Analysis (Dumps)",        False),
    ("DB02",  "Database Performance Monitor",               False),
    ("SMICM", "Internet Communication Manager Monitor",     False),
    ("SM58",  "Transactional RFC Error Log",                False),
    ("RZ20",  "CCMS Monitor",                               False),
]

# ═══════════════════════════════════════════════════════════════════


# ─────────────────────────────────────────────────────────────────
# Dependency check
# ─────────────────────────────────────────────────────────────────

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

# Prevent pyautogui from raising on fast moves
pyautogui.PAUSE = 0.05
pyautogui.FAILSAFE = False


# ─────────────────────────────────────────────────────────────────
# SAP window discovery
# ─────────────────────────────────────────────────────────────────

def find_sap_window():
    """Return the HWND of the best visible SAP GUI session window."""
    candidates = []

    def _cb(hwnd, _):
        if not win32gui.IsWindowVisible(hwnd):
            return
        title = win32gui.GetWindowText(hwnd)
        if "SAP" in title and "SAP Logon" not in title:
            candidates.append((hwnd, title))

    win32gui.EnumWindows(_cb, None)

    if not candidates:
        print("\n[ERROR] No SAP GUI window found.")
        print("  Make sure SAP GUI is open and you are logged in.\n")
        sys.exit(1)

    print(f"  SAP windows found ({len(candidates)}):")
    for h, t in candidates:
        print(f"    hwnd={h}  title={t!r}")

    # Prefer "SAP Easy Access" window; otherwise take first
    for h, t in candidates:
        if "Easy Access" in t:
            print(f"  Using: {t!r}")
            return h

    hwnd, title = candidates[0]
    print(f"  Using: {title!r}")
    return hwnd


# ─────────────────────────────────────────────────────────────────
# System menu navigation (child-window enum + pixel fallback)
# SAP GUI renders its own menu — standard Win32 GetMenu() returns NULL.
# ─────────────────────────────────────────────────────────────────

def _find_child_by_text(hwnd, text, max_y_from_top=60):
    """
    Return (cx, cy) screen coords of the first child window whose
    GetWindowText matches *text* and whose top edge is within
    *max_y_from_top* pixels of the parent window's top edge.
    Returns None if not found.
    """
    wr   = win32gui.GetWindowRect(hwnd)
    hit  = []

    def _cb(h, _):
        if win32gui.GetWindowText(h).strip() == text:
            r = win32gui.GetWindowRect(h)
            if (r[1] - wr[1]) < max_y_from_top:
                hit.append(((r[0] + r[2]) // 2, (r[1] + r[3]) // 2))

    try:
        win32gui.EnumChildWindows(hwnd, _cb, None)
    except Exception:
        pass
    return hit[0] if hit else None


def _open_system_status(hwnd):
    """
    Click 'System' in the SAP menu bar, then 'Status...' in the dropdown.
    Strategy:
      1. Try to locate 'System' as a child window (some SAP GUI builds).
      2. Fall back to pixel offsets SYSTEM_MENU_X / SYSTEM_MENU_Y.
    Then click Status... at SYSTEM_MENU_X + STATUS_OFFSET_X/Y.
    """
    _focus_and_restore(hwnd)
    wr = win32gui.GetWindowRect(hwnd)

    # ── Click "System" in the menu bar ───────────────────────────
    pos = _find_child_by_text(hwnd, "System", max_y_from_top=60)
    if pos:
        sx, sy = pos
        print(f"  [Menu] 'System' found as child window → ({sx}, {sy})")
    else:
        sx = wr[0] + SYSTEM_MENU_X
        sy = wr[1] + SYSTEM_MENU_Y
        print(f"  [Menu] Pixel fallback → clicking 'System' at ({sx}, {sy})")

    pyautogui.click(sx, sy)
    time.sleep(1.5)   # wait for SAP dropdown to fully open

    # ── Navigate to Status... using keyboard ─────────────────────
    # Keyboard navigation is more reliable than pixel clicking in
    # SAP dropdowns.  Down arrows skip separators; Status... is the
    # 11th selectable item in the System menu.
    print(f"  [Menu] Navigating to 'Status...' via keyboard (Down ×11, Enter)")
    for _ in range(11):
        pyautogui.press("down")
        time.sleep(0.08)
    pyautogui.press("enter")



# ─────────────────────────────────────────────────────────────────
# Popup window helpers
# ─────────────────────────────────────────────────────────────────

def _visible_top_windows():
    hwnds = set()
    def _cb(h, _):
        if win32gui.IsWindowVisible(h) and win32gui.GetWindowText(h):
            hwnds.add(h)
    win32gui.EnumWindows(_cb, None)
    return hwnds

def _wait_new_window(before, timeout=6):
    """Wait for a new real dialog window (skips transient menu popups)."""
    deadline = time.time() + timeout
    while time.time() < deadline:
        new = _visible_top_windows() - before
        for h in new:
            # skip the #32768 class used by Windows popup/context menus
            if win32gui.GetClassName(h) != "#32768" and win32gui.IsWindow(h):
                return h
        time.sleep(0.2)
    return None

def _find_navigate_button(popup_hwnd):
    """Try child-window enumeration first; returns (cx, cy) or None."""
    found = []
    def _cb(h, _):
        if "navigate" in win32gui.GetWindowText(h).lower():
            found.append(h)
    try:
        win32gui.EnumChildWindows(popup_hwnd, _cb, None)
    except Exception:
        pass
    if found:
        r = win32gui.GetWindowRect(found[0])
        return ((r[0] + r[2]) // 2, (r[1] + r[3]) // 2)
    return None


# ─────────────────────────────────────────────────────────────────
# System Status capture
# ─────────────────────────────────────────────────────────────────

def capture_system_status(main_hwnd, doc, tmp_dir, img_index):
    """
    1. Open System > Status
    2. Screenshot the status popup → add to doc
    3. Click the Navigate button (kernel info)
    4. Screenshot the kernel popup → add to doc
    5. Close both popups
    Returns the next available img_index.
    """
    print("  [--/--]  System > Status  (menu navigation)")

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

    time.sleep(1.5)   # let the dialog fully render

    # Screenshot 1: Status popup
    img1 = tmp_dir / f"{img_index:02d}_SYSTEM_STATUS.png"
    try:
        _focus_and_restore(status_hwnd)
        time.sleep(SCREENSHOT_DELAY)
        rect = win32gui.GetWindowRect(status_hwnd)
        ImageGrab.grab(bbox=rect).save(str(img1))
    except Exception as exc:
        print(f"  [WARN] Status screenshot failed: {exc}")
        img1 = None

    _add_tcode_block(doc, "System > Status", "System Status", str(img1) if img1 else None)
    img_index += 1

    # Click Navigate button
    before2 = _visible_top_windows()
    _focus_and_restore(status_hwnd)

    btn = _find_navigate_button(status_hwnd)
    if btn:
        pyautogui.click(*btn)
    else:
        # Positional fallback — adjust STATUS_NAV_BTN_X/Y at top of file if needed
        wr = win32gui.GetWindowRect(status_hwnd)
        pyautogui.click(wr[0] + STATUS_NAV_BTN_X, wr[1] + STATUS_NAV_BTN_Y)

    time.sleep(3.0)   # wait for kernel popup to fully load

    # Screenshot 2: Kernel info popup
    kernel_hwnd = _wait_new_window(before2, timeout=8)
    if kernel_hwnd:
        img2 = tmp_dir / f"{img_index:02d}_SYSTEM_KERNEL.png"
        try:
            _focus_and_restore(kernel_hwnd)
            time.sleep(SCREENSHOT_DELAY)
            rect = win32gui.GetWindowRect(kernel_hwnd)
            ImageGrab.grab(bbox=rect).save(str(img2))
        except Exception as exc:
            print(f"  [WARN] Kernel screenshot failed: {exc}")
            img2 = None

        _add_tcode_block(doc, "System > Status > Navigate",
                         "Kernel Information", str(img2) if img2 else None)
        img_index += 1

        _focus_and_restore(kernel_hwnd)
        pyautogui.press("escape")
        time.sleep(0.4)
    else:
        print("  [WARN] Kernel info popup did not appear")

    # Close status popup
    try:
        _focus_and_restore(status_hwnd)
        pyautogui.press("escape")
        time.sleep(0.4)
    except Exception:
        pass

    return img_index


# ─────────────────────────────────────────────────────────────────
# Navigation via keyboard
# ─────────────────────────────────────────────────────────────────

def _focus_and_restore(hwnd):
    """Restore-if-minimised and bring the window to the foreground."""
    placement = win32gui.GetWindowPlacement(hwnd)
    if placement[1] == win32con.SW_SHOWMINIMIZED:
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        time.sleep(0.3)
    win32gui.SetForegroundWindow(hwnd)
    time.sleep(0.3)


def _click_command_field(hwnd):
    """Click the SAP command (OK-code) input field."""
    if CMD_FIELD_ABSOLUTE:
        x, y = CMD_FIELD_X, CMD_FIELD_Y
    else:
        rect = win32gui.GetWindowRect(hwnd)   # (left, top, right, bottom)
        x = rect[0] + CMD_FIELD_X
        y = rect[1] + CMD_FIELD_Y
    pyautogui.click(x, y)
    time.sleep(0.2)


def navigate_tcode(hwnd, tcode, press_f8=False):
    """
    Bring the SAP window to front, click the command field,
    clear it, type /n<tcode>, press Enter, and optionally F8.
    """
    try:
        _focus_and_restore(hwnd)
        _click_command_field(hwnd)

        # Clear any existing text and type the navigation command
        pyautogui.hotkey("ctrl", "a")
        time.sleep(0.1)
        pyautogui.typewrite(f"/n{tcode}", interval=0.05)
        pyautogui.press("enter")
        time.sleep(WAIT_AFTER_NAV)

        # Dismiss any popup (press Enter/Escape)
        pyautogui.press("escape")
        time.sleep(0.2)

        if press_f8:
            _focus_and_restore(hwnd)
            pyautogui.press("f8")
            time.sleep(WAIT_AFTER_NAV)
            pyautogui.press("escape")
            time.sleep(0.2)

        return True

    except Exception as exc:
        print(f"  [WARN] Navigation to {tcode} failed: {exc}")
        return False


# ─────────────────────────────────────────────────────────────────
# Screenshot
# ─────────────────────────────────────────────────────────────────

def capture_window(hwnd):
    """Bring the SAP window to front and capture its bounding rect."""
    try:
        _focus_and_restore(hwnd)
        time.sleep(SCREENSHOT_DELAY)
        rect = win32gui.GetWindowRect(hwnd)
        return ImageGrab.grab(bbox=rect)
    except Exception as exc:
        print(f"  [WARN] Window capture failed ({exc}), falling back to full screen")
        time.sleep(SCREENSHOT_DELAY)
        return ImageGrab.grab()


# ─────────────────────────────────────────────────────────────────
# Word document helpers  (unchanged from original)
# ─────────────────────────────────────────────────────────────────

def _create_document(sid, total):
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
    meta_text = (
        f"System / SID:  {sid or '—'}"
        f"     |     Generated:  {datetime.now().strftime('%Y-%m-%d  %H:%M')}"
        f"     |     T-codes captured:  {total}"
    )
    meta_run = meta_para.add_run(meta_text)
    meta_run.font.size = Pt(12)
    meta_run.font.color.rgb = RGBColor(0x40, 0x40, 0x40)

    doc.add_page_break()
    return doc


def _add_horizontal_rule(doc):
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


def _add_tcode_block(doc, tcode, description, img_path):
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

    _add_horizontal_rule(doc)

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
# Main orchestrator
# ─────────────────────────────────────────────────────────────────

def run(sid="", output_path=None, tcodes=None):
    if tcodes is None:
        tcodes = TCODES

    output_dir = Path(DEFAULT_OUTPUT_DIR)
    output_dir.mkdir(parents=True, exist_ok=True)

    if not output_path:
        ts       = datetime.now().strftime("%Y%m%d_%H%M")
        sid_part = f"_{sid}" if sid else ""
        output_path = output_dir / f"SAP_Monitor{sid_part}_{ts}.docx"

    print("\n" + "=" * 62)
    print("  SAP GUI Monitor Screenshot Script  (keyboard automation)")
    print("=" * 62)

    hwnd = find_sap_window()
    doc  = _create_document(sid, len(tcodes))

    print(f"\n  T-codes  : {len(tcodes)}")
    print(f"  SID      : {sid or '(auto)'}")
    print(f"  Output   : {output_path}")
    print()
    print("  NOTE: Do not move the mouse or use the keyboard while")
    print("        the script is running.\n")

    tmp_dir  = Path(tempfile.mkdtemp(prefix="sap_shots_"))
    captured = 0

    # ── System Status screenshots (before t-codes) ────────────────
    img_index = 1
    if CAPTURE_SYSTEM_STATUS:
        img_index = capture_system_status(hwnd, doc, tmp_dir, img_index)

    for tcode_num, entry in enumerate(tcodes, 1):
        if not RUN_TCODES:
            break
        tcode = entry[0]
        desc  = entry[1] if len(entry) > 1 else tcode
        f8    = entry[2] if len(entry) > 2 else False

        print(f"  [{tcode_num:02d}/{len(tcodes):02d}]  {tcode:<10}  {desc}")

        navigate_tcode(hwnd, tcode, f8)

        img_path = tmp_dir / f"{img_index:02d}_{tcode}.png"
        try:
            img = capture_window(hwnd)
            img.save(str(img_path))
            captured += 1
        except Exception as exc:
            print(f"           [WARN] Screenshot failed: {exc}")
            img_path = None

        _add_tcode_block(doc, tcode, desc, str(img_path) if img_path else None)
        img_index += 1

    # Return to SAP Easy Access
    try:
        _focus_and_restore(hwnd)
        _click_command_field(hwnd)
        pyautogui.hotkey("ctrl", "a")
        pyautogui.typewrite("/n", interval=0.05)
        pyautogui.press("enter")
    except Exception:
        pass

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
    print(f"  COMPLETE")
    print(f"  Screenshots captured : {captured} / {len(tcodes)}")
    print(f"  Word document        : {output_path}")
    print(f"{'=' * 62}\n")
    return str(output_path)


# ─────────────────────────────────────────────────────────────────
# CLI
# ─────────────────────────────────────────────────────────────────

def _run_calibrate():
    """
    Take a screenshot of the SAP window top area and mark the planned
    click positions with coloured circles — no actual clicking.
    Run with:  python sap_monitor_screenshots.py --calibrate
    """
    from PIL import ImageDraw
    hwnd = find_sap_window()
    wr   = win32gui.GetWindowRect(hwnd)

    sx = wr[0] + SYSTEM_MENU_X          # "System" menu click
    sy = wr[1] + SYSTEM_MENU_Y
    dx = sx + STATUS_OFFSET_X            # "Status..." dropdown click
    dy = sy + STATUS_OFFSET_Y

    # Capture the top-left 450×380 px of the SAP window (avoids the big bg)
    crop = (max(wr[0]+8, 0), max(wr[1]+8, 0),
            min(wr[0]+8+450, wr[2]-8), min(wr[1]+8+380, wr[3]-8))
    img  = ImageGrab.grab(bbox=crop)
    draw = ImageDraw.Draw(img)

    def _mark(abs_x, abs_y, colour, label):
        rx = abs_x - crop[0]
        ry = abs_y - crop[1]
        r  = 10
        draw.ellipse((rx-r, ry-r, rx+r, ry+r), outline=colour, width=3)
        draw.text((rx+r+3, ry-8), label, fill=colour)

    _mark(sx, sy, "red",  f"System ({sx},{sy})")
    _mark(dx, dy, "blue", f"Status ({dx},{dy})")

    out = Path(DEFAULT_OUTPUT_DIR) / "calibrate.png"
    img.save(str(out))
    print(f"\n  Calibration image saved: {out}")
    print(f"  RED  dot = 'System' menu click  ({sx}, {sy})")
    print(f"  BLUE dot = 'Status...' click    ({dx}, {dy})")
    print(f"\n  If dots are off-target, adjust SYSTEM_MENU_X/Y and")
    print(f"  STATUS_OFFSET_X/Y at the top of the file, then re-run --calibrate.")


def main():
    ap = argparse.ArgumentParser(
        description="SAP GUI Monitor — keyboard automation, no scripting needed",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python sap_monitor_screenshots.py --calibrate   ← check click positions first
  python sap_monitor_screenshots.py
  python sap_monitor_screenshots.py --sid PRD
  python sap_monitor_screenshots.py --sid PRD --output "C:/out/monitor.docx"
  python sap_monitor_screenshots.py --tcodes SM50,SM51,ST22,DB02
        """,
    )
    ap.add_argument("--sid",       default="",    help="SAP System ID (e.g. PRD)")
    ap.add_argument("--output",    default=None,  help="Output .docx file path")
    ap.add_argument("--calibrate", action="store_true",
                    help="Save a preview image showing where clicks will land (no clicking)")
    ap.add_argument(
        "--tcodes", default=None,
        help="Comma-separated t-codes (overrides built-in list)",
    )
    args = ap.parse_args()

    if args.calibrate:
        _run_calibrate()
        return

    tcodes = None
    if args.tcodes:
        tcodes = [(t.strip(), t.strip(), False) for t in args.tcodes.split(",")]

    run(args.sid, args.output, tcodes)


if __name__ == "__main__":
    main()
