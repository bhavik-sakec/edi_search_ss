"""
Notepad++ Search and Screenshot Tool
Opens a text file in Notepad++, searches for specific segments,
highlights the found text (complete line), and takes a screenshot.

Usage:
    # Single field
    python main.py --file EDI.txt --word "BHT03"
    
    # List of fields
    python main.py --file EDI.txt --list BHT03 BHT04 CLM05
    
    # Text file input (one field per line)
    python main.py --file EDI.txt --txt fields.txt
    
    # Excel input (columns: 'GDF Field' for filename, 'EDI Field' for search term)
    # Supports complex EDI formats like: 2300CLM05-1, 2010AANM109, 2300DTP03 - 434
    python main.py --file EDI.txt --excel Book1.xlsx
"""

import subprocess
import time
import pyautogui
import pyperclip
from datetime import datetime
import os
import argparse
import re
import win32gui
import win32con
from PIL import Image, ImageDraw
import msvcrt  # For keyboard input on Windows

# Disable pyautogui's built-in pause between actions
pyautogui.PAUSE = 0
pyautogui.FAILSAFE = True  # Move mouse to corner to abort

# Optional: pandas for Excel support
try:
    import pandas as pd
    PANDAS_AVAILABLE = True
except ImportError:
    PANDAS_AVAILABLE = False


# Configuration
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
SCREENSHOT_FOLDER = os.path.join(SCRIPT_DIR, "screenshots")

# Global variable to store window handle
NOTEPAD_HWND = None


def parse_field_code(field_code: str) -> tuple:
    """
    Parses a field code into segment name, element number, sub-element number, and search term.
    
    Supported formats:
        BHT03     -> ('BHT', '03', None, 'BHT*')     - Element 03 of BHT segment
        CLM05     -> ('CLM', '05', None, 'CLM*')     - Element 05 of CLM segment
        CLM05-1   -> ('CLM', '05', '1', 'CLM*')      - Sub-element 1 of Element 05
        NM101     -> ('NM1', '01', None, 'NM1*')     - Element 01 of NM1 segment
        NM109     -> ('NM1', '09', None, 'NM1*')     - Element 09 of NM1 segment
        SV101     -> ('SV1', '01', None, 'SV1*')     - Element 01 of SV1 segment
        SV101-1   -> ('SV1', '01', '1', 'SV1*')      - Sub-element 1 of Element 01
    
    Returns: (segment, element_num, sub_element_num, search_term)
    """
    field_code = field_code.strip().upper()
    
    # Known EDI segments that have numbers in their name
    NUMBERED_SEGMENTS = ['NM1', 'N1', 'N2', 'N3', 'N4', 'SV1', 'SV2', 'SV3', 'SV4', 'SV5', 
                         'HI', 'HL', 'K3', 'G1', 'G2', 'G3', 'LX', 'SE', 'ST', 'GS', 'GE',
                         'ISA', 'IEA', 'TA1']
    
    # Pattern 1: Known numbered segment + Element + Sub-element (e.g., NM101-1, SV101-2)
    for seg in NUMBERED_SEGMENTS:
        if field_code.startswith(seg):
            rest = field_code[len(seg):]
            # Check for element-subelement pattern
            match = re.match(r'^(\d+)-(\d+)$', rest)
            if match:
                element_num = match.group(1)
                sub_element = match.group(2)
                return seg, element_num, sub_element, f"{seg}*"
            # Check for element only pattern
            match = re.match(r'^(\d+)$', rest)
            if match:
                element_num = match.group(1)
                return seg, element_num, None, f"{seg}*"
    
    # Pattern 2: Segment + Element + Sub-element (e.g., CLM05-1, DTP01-2)
    match = re.match(r'^([A-Za-z]+)(\d+)-(\d+)$', field_code)
    if match:
        segment = match.group(1)
        element_num = match.group(2)
        sub_element = match.group(3)
        return segment, element_num, sub_element, f"{segment}*"
    
    # Pattern 3: Segment + Element only (e.g., BHT03, CLM05)
    match = re.match(r'^([A-Za-z]+)(\d+)$', field_code)
    if match:
        segment = match.group(1)
        element_num = match.group(2)
        return segment, element_num, None, f"{segment}*"
    
    # Fallback: use as-is for search
    return field_code, '', None, f"{field_code}*"


def parse_edi_field(edi_field: str) -> str:
    """
    Parses complex EDI field formats and returns the search term.
    Based on ASC X12 837 Professional (005010X222A1) implementation guide.
    
    Supported formats and search patterns:
        BHT03/BHT04                 -> 'BHT*'           - Transaction header
        2300CLM05-1                 -> 'CLM*'           - Claim
        2010AANM109                 -> 'NM1*85*'        - Billing provider (2010AA loop)
        2310BNM109                  -> 'NM1*82*'        - Rendering provider (2310B loop)
        2010BANM109                 -> 'NM1*IL*'        - Subscriber (2010BA loop)
        2300DTP03 - 434             -> 'DTP*434*'       - Claim service date
        2300DTP03 - 435             -> 'DTP*435*'       - Discharge date
        2400DTP03 - 472             -> 'DTP*472*'       - Service line date
        2300HI01-2 -- BK/ABK        -> 'HI*ABK:'        - Diagnosis codes
        2410LIN03 -- N4             -> 'LIN**N4'        - NDC drug code
    
    Returns: search_term (e.g., 'NM1*85*', 'DTP*434*', 'HI*ABK:')
    """
    original_field = edi_field.strip()
    edi_field = original_field
    
    # Handle combined fields (e.g., "2300CLM05-1 + 2300CLM05-3") - use first field
    if ' + ' in edi_field:
        edi_field = edi_field.split(' + ')[0].strip()
    
    # Extract qualifier if present
    qualifier = None
    qualifier_type = None  # 'dash' for " - " or 'double_dash' for " -- "
    
    if ' -- ' in edi_field:
        parts = edi_field.split(' -- ')
        edi_field = parts[0].strip()
        qualifier = parts[1].strip() if len(parts) > 1 else None
        qualifier_type = 'double_dash'
    elif ' - ' in edi_field:
        # Handle " - DR" format (space dash space)
        parts = edi_field.split(' - ')
        edi_field = parts[0].strip()
        qualifier = parts[1].strip() if len(parts) > 1 else None
        qualifier_type = 'dash'
    elif ' -' in edi_field:
        # Handle " -BG" format (space before dash, no space after, followed by letters)
        match = re.search(r' -([A-Za-z/]+)', edi_field)
        if match:
            qualifier = match.group(1).strip()
            edi_field = edi_field[:edi_field.find(' -')].strip()
            qualifier_type = 'dash'
    
    # Special handling for HI segments with qualifier attached (e.g., "2300HI01-2-ABJ/BJ")
    # Pattern: ...HI##-#-QUALIFIER
    if not qualifier and 'HI' in edi_field.upper():
        hi_match = re.search(r'HI(\d+)-(\d+)-([A-Za-z/]+)$', edi_field, re.IGNORECASE)
        if hi_match:
            qualifier = hi_match.group(3)
            edi_field = edi_field[:edi_field.rfind('-' + qualifier)]
            qualifier_type = 'attached'
    
    # Extract loop ID if present (e.g., 2010AA, 2310B, 2300, 2400, 2410)
    loop_id = None
    loop_match = re.match(r'^(\d{4}[A-Z]{0,2})', edi_field.upper())
    if loop_match:
        loop_id = loop_match.group(1)
    
    # Known 3-character segments that have numbers in their name
    NUMBERED_SEGMENTS_3 = ['NM1', 'SV1', 'SV2', 'SV3', 'SV4', 'SV5', 'TA1', 'CL1']
    
    # Known 2-character segments
    SEGMENTS_2 = ['N1', 'N2', 'N3', 'N4', 'HI', 'HL', 'K3', 'G1', 'G2', 'G3', 'LX', 'SE', 'ST', 'GS', 'GE']
    
    # Known 3-character segments (no numbers)
    SEGMENTS_3 = ['CLM', 'BHT', 'DTP', 'REF', 'AMT', 'QTY', 'DMG', 'PAT', 'SBR', 'CAS', 'OI', 'MOA', 'LIN', 'CTP', 'PRV', 'CN1', 'PWK', 'CR1', 'CR2', 'CR3', 'CR5', 'CR6', 'CRC', 'HCP', 'TST', 'MEA', 'PER']
    
    segment = None
    
    # First, try to find known 3-char numbered segments (like NM1, SV1)
    for seg in NUMBERED_SEGMENTS_3:
        if seg in edi_field.upper():
            segment = seg
            break
    
    # Try to find known 3-char segments
    if not segment:
        for seg in SEGMENTS_3:
            if seg in edi_field.upper():
                segment = seg
                break
    
    # Try to find known 2-char segments  
    if not segment:
        for seg in SEGMENTS_2:
            idx = edi_field.upper().find(seg)
            if idx >= 0:
                rest = edi_field[idx + len(seg):]
                if rest and rest[0].isdigit():
                    segment = seg
                    break
    
    # Pattern 2: Simple format (e.g., BHT04, CLM05)
    if not segment:
        match = re.match(r'^([A-Za-z]+\d?)(\d+)(?:-\d+)?$', edi_field)
        if match:
            segment = match.group(1).upper()
    
    # Pattern 3: Try to extract alphabetic prefix as segment
    if not segment:
        match = re.match(r'^(\d*)([A-Za-z]+\d?)([A-Za-z]*)(\d+)', edi_field)
        if match:
            seg_candidate = (match.group(2) + match.group(3)).upper()
            for seg in NUMBERED_SEGMENTS_3 + SEGMENTS_3:
                if seg in seg_candidate:
                    segment = seg
                    break
            if not segment:
                for seg in SEGMENTS_2:
                    if seg_candidate.endswith(seg) or seg in seg_candidate:
                        segment = seg
                        break
    
    # Last resort: use parse_field_code for simple formats
    if not segment:
        seg, element_num, sub_element, search_term = parse_field_code(edi_field)
        segment = seg
    
    # ================================================================
    # Build search term based on segment type and loop context
    # Following ASC X12 837P (005010X222A1) patterns
    # ================================================================
    
    # NM1 segments - need entity identifier
    if segment == 'NM1':
        # Determine entity type based on loop ID
        if loop_id:
            if '2010AA' in loop_id:
                return "NM1*85*"  # Billing provider
            elif '2010AB' in loop_id:
                return "NM1*87*"  # Pay-to provider
            elif '2010AC' in loop_id:
                return "NM1*PE*"  # Pay-to plan
            elif '2010BA' in loop_id:
                return "NM1*IL*"  # Subscriber
            elif '2010BB' in loop_id:
                return "NM1*PR*"  # Payer
            elif '2010CA' in loop_id:
                return "NM1*QC*"  # Patient
            elif '2310A' in loop_id:
                return "NM1*DN*"  # Referring provider
            elif '2310B' in loop_id:
                return "NM1*82*"  # Rendering provider
            elif '2310C' in loop_id:
                return "NM1*77*"  # Service facility
            elif '2310D' in loop_id:
                return "NM1*DQ*"  # Supervising provider
            elif '2330' in loop_id:
                return "NM1*"     # Other subscriber (various)
            elif '2420' in loop_id:
                return "NM1*82*"  # Line rendering provider
        return "NM1*"
    
    # HI segments - diagnosis codes with qualifier
    if segment == 'HI':
        if qualifier:
            # Handle "/" separator (e.g., "BK/ABK")
            if '/' in qualifier:
                hi_qualifier = qualifier.split('/')[-1].strip()
            else:
                hi_qualifier = qualifier
            return f"HI*{hi_qualifier}"
        return "HI*"
    
    # DTP segments - date qualifier
    if segment == 'DTP':
        if qualifier:
            # Handle RD8 date range format
            if 'RD8' in qualifier.upper():
                match = re.match(r'(\d+)', qualifier)
                if match:
                    return f"DTP*{match.group(1)}*RD8"
            else:
                # Just the date qualifier (434, 435, 472, etc.)
                num_match = re.match(r'(\d+)', qualifier)
                if num_match:
                    return f"DTP*{num_match.group(1)}*"
        return "DTP*"
    
    # REF segments - need qualifier based on loop
    if segment == 'REF':
        if loop_id:
            if '2010AA' in loop_id:
                return "REF*EI*"  # Billing provider tax ID
            elif '2310B' in loop_id:
                return "REF*"    # Rendering provider secondary ID
        return "REF*"
    
    # LIN segments - drug identification
    if segment == 'LIN':
        if qualifier and 'N4' in qualifier.upper():
            return "LIN**N4*"  # NDC code
        return "LIN*"
    
    # N4 segments - address
    if segment == 'N4':
        return "N4*"
    
    # N3 segments - address line
    if segment == 'N3':
        return "N3*"
    
    # PRV segments - provider info
    if segment == 'PRV':
        return "PRV*"
    
    # SV1/SV2 segments - service line
    if segment in ['SV1', 'SV2']:
        return f"{segment}*"
    
    # LX segments - line counter
    if segment == 'LX':
        return "LX*"
    
    # CTP segments - drug pricing
    if segment == 'CTP':
        return "CTP*"
    
    # DMG segments - demographic info
    if segment == 'DMG':
        return "DMG*"
    
    # CN1 segments - contract info
    if segment == 'CN1':
        return "CN1*"
    
    # CL1 segments - institutional claim code
    if segment == 'CL1':
        return "CL1*"
    
    # HCP segments - repricing
    if segment == 'HCP':
        return "HCP*"
    
    # Default: just segment name
    return f"{segment}*"


def load_from_excel(excel_path: str) -> list:
    """
    Loads search terms from Excel file.
    Expected columns: 'GDF Field' (for screenshot name) and 'EDI Field' (for search term)
    Also supports legacy format with 'File name' and 'Field' columns.
    
    Returns list of tuples: [(screenshot_name, field_code, search_term), ...]
    """
    if not PANDAS_AVAILABLE:
        raise ImportError("pandas is required for Excel support. Install with: pip install pandas openpyxl")
    
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"Excel file not found: {excel_path}")
    
    df = pd.read_excel(excel_path)
    
    # Check for required columns (case-insensitive)
    columns_lower = {col.lower().strip(): col for col in df.columns}
    
    file_name_col = None
    field_col = None
    
    # Check for new format: 'GDF Field' and 'EDI Field'
    for key, col in columns_lower.items():
        if 'gdf' in key and 'field' in key:
            file_name_col = col
        elif 'edi' in key and 'field' in key:
            field_col = col
    
    # Fallback to legacy format: 'File name' and 'Field'
    if file_name_col is None or field_col is None:
        for key, col in columns_lower.items():
            if 'file' in key and 'name' in key:
                file_name_col = col
            elif 'field' in key and file_name_col != col:
                field_col = col
    
    if file_name_col is None or field_col is None:
        raise ValueError(f"Excel must have 'GDF Field' and 'EDI Field' columns (or 'File name' and 'Field'). Found: {list(df.columns)}")
    
    results = []
    for _, row in df.iterrows():
        file_name = str(row[file_name_col]).strip()
        field_code = str(row[field_col]).strip()
        
        if field_code and field_code.lower() != 'nan':
            # Use new parser for complex EDI field formats
            search_term = parse_edi_field(field_code)
            results.append((file_name, field_code, search_term))
    
    return results


def load_from_txt(txt_path: str) -> list:
    """
    Loads search terms from a text file.
    Each line should contain one field code (e.g., BHT03)
    Lines starting with # are treated as comments.
    
    Returns list of tuples: [(field_code, field_code, search_term), ...]
    """
    if not os.path.exists(txt_path):
        raise FileNotFoundError(f"Text file not found: {txt_path}")
    
    results = []
    with open(txt_path, 'r', encoding='utf-8') as f:
        for line_num, line in enumerate(f, 1):
            line = line.strip()
            
            # Skip empty lines and comments
            if not line or line.startswith('#'):
                continue
            
            # Parse the field code
            segment, element_num, sub_element, search_term = parse_field_code(line)
            results.append((line, line, search_term))
    
    return results


def open_notepad_plus_with_file(file_path: str):
    """
    Opens Notepad++ with the specified file, maximizes it, and returns the window handle.
    """
    global NOTEPAD_HWND
    
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"File not found: {file_path}")
    
    notepad_plus_paths = [
        r"C:\Program Files\Notepad++\notepad++.exe",
        r"C:\Program Files (x86)\Notepad++\notepad++.exe",
        "notepad++.exe"
    ]
    
    notepad_exe = None
    for path in notepad_plus_paths:
        if os.path.exists(path):
            notepad_exe = path
            break
    
    if notepad_exe is None:
        notepad_exe = "notepad++.exe"
    
    process = subprocess.Popen([notepad_exe, file_path])
    time.sleep(1)  # Reduced from 2s
    
    NOTEPAD_HWND = find_notepad_window()
    if NOTEPAD_HWND:
        win32gui.ShowWindow(NOTEPAD_HWND, win32con.SW_MAXIMIZE)
        time.sleep(0.1)  # Reduced from 0.3s
        win32gui.SetForegroundWindow(NOTEPAD_HWND)
        time.sleep(0.1)  # Reduced from 0.3s
    
    return process


def find_notepad_window():
    """Finds the Notepad++ window handle."""
    def callback(hwnd, hwnds):
        if win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)
            if 'Notepad++' in title:
                hwnds.append(hwnd)
        return True
    
    hwnds = []
    win32gui.EnumWindows(callback, hwnds)
    return hwnds[0] if hwnds else None


def ensure_notepad_focus():
    """Ensures Notepad++ has focus without changing window state."""
    global NOTEPAD_HWND
    
    if NOTEPAD_HWND is None:
        NOTEPAD_HWND = find_notepad_window()
    
    if NOTEPAD_HWND:
        try:
            win32gui.SetForegroundWindow(NOTEPAD_HWND)
        except:
            pass
        time.sleep(0.05)  # Reduced from 0.1s
        return NOTEPAD_HWND
    return None


def search_and_highlight(search_text: str, dialog_open: bool = False) -> bool:
    """
    Searches for text in Notepad++ and highlights the entire line.
    Returns True if text was found, False otherwise.
    
    Args:
        search_text: Text to search for
        dialog_open: If True, assumes Find dialog is already open (faster)
    """
    if not dialog_open:
        ensure_notepad_focus()
    
    # Open Find dialog only if not already open
    if not dialog_open:
        pyautogui.hotkey('ctrl', 'f')
        time.sleep(0.1)
    
    # Clear and type search text (Ctrl+A to select all in search box)
    pyautogui.hotkey('ctrl', 'a')
    time.sleep(0.02)
    
    pyperclip.copy(search_text)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(0.05)
    
    # Find Next
    pyautogui.press('enter')
    time.sleep(0.1)
    
    # Close Find dialog only if we opened it
    if not dialog_open:
        pyautogui.press('escape')
        time.sleep(0.05)
    
    return True


def select_current_line():
    """Selects the entire current line."""
    pyautogui.press('home')
    time.sleep(0.02)
    pyautogui.hotkey('shift', 'end')
    time.sleep(0.05)


def open_find_dialog():
    """Opens the Find dialog in Notepad++."""
    ensure_notepad_focus()
    pyautogui.hotkey('ctrl', 'f')
    time.sleep(0.1)


def close_find_dialog():
    """Closes the Find dialog."""
    pyautogui.press('escape')
    time.sleep(0.05)


def get_caret_position():
    """Gets the current caret position on screen."""
    import ctypes
    from ctypes import wintypes
    
    class GUITHREADINFO(ctypes.Structure):
        _fields_ = [
            ("cbSize", wintypes.DWORD),
            ("flags", wintypes.DWORD),
            ("hwndActive", wintypes.HWND),
            ("hwndFocus", wintypes.HWND),
            ("hwndCapture", wintypes.HWND),
            ("hwndMenuOwner", wintypes.HWND),
            ("hwndMoveSize", wintypes.HWND),
            ("hwndCaret", wintypes.HWND),
            ("rcCaret", wintypes.RECT),
        ]
    
    gui_info = GUITHREADINFO()
    gui_info.cbSize = ctypes.sizeof(GUITHREADINFO)
    
    user32 = ctypes.windll.user32
    if user32.GetGUIThreadInfo(0, ctypes.byref(gui_info)):
        caret_rect = gui_info.rcCaret
        hwnd_caret = gui_info.hwndCaret
        
        if hwnd_caret:
            caret_height = caret_rect.bottom - caret_rect.top
            point = wintypes.POINT()
            point.x = caret_rect.left
            point.y = caret_rect.top
            user32.ClientToScreen(hwnd_caret, ctypes.byref(point))
            return point.x, point.y, caret_height
    
    pos = pyautogui.position()
    return pos[0], pos[1], 20


def take_screenshot(filename: str, silent: bool = True) -> str:
    """Takes a full screen screenshot and draws a red rectangle around the selected line."""
    if not os.path.exists(SCREENSHOT_FOLDER):
        os.makedirs(SCREENSHOT_FOLDER)
    
    # Clean filename
    safe_filename = re.sub(r'[<>:"/\\|?*]', '_', filename)
    if not safe_filename.endswith('.png'):
        safe_filename += '.png'
    
    screenshot_path = os.path.join(SCREENSHOT_FOLDER, safe_filename)
    
    hwnd = NOTEPAD_HWND
    time.sleep(0.08)  # Reduced from 0.2s
    
    caret_x, caret_y, caret_height = get_caret_position()
    
    if hwnd:
        rect = win32gui.GetWindowRect(hwnd)
        win_left, win_top, win_right, win_bottom = rect
    else:
        win_left, win_top = 0, 0
        win_right = pyautogui.size()[0]
        win_bottom = pyautogui.size()[1]
    
    screenshot = pyautogui.screenshot()
    
    # Draw rectangle around selected line
    draw = ImageDraw.Draw(screenshot)
    padding = 3
    
    rect_left = win_left + 60
    rect_right = caret_x + 10
    rect_top = caret_y - padding
    rect_bottom = caret_y + caret_height + padding
    
    for i in range(3):
        draw.rectangle(
            [rect_left - i, rect_top - i, rect_right + i, rect_bottom + i],
            outline="red"
        )
    
    screenshot.save(screenshot_path)
    
    return screenshot_path


def close_notepad_plus():
    """Closes Notepad++ without saving."""
    global NOTEPAD_HWND
    
    if NOTEPAD_HWND:
        try:
            win32gui.SetForegroundWindow(NOTEPAD_HWND)
        except:
            pass
        time.sleep(0.1)  # Reduced from 0.2s
    
    pyautogui.hotkey('alt', 'F4')
    time.sleep(0.15)  # Reduced from 0.3s
    pyautogui.press('n')
    time.sleep(0.1)  # Reduced from 0.2s
    
    NOTEPAD_HWND = None


def check_quit_key():
    """Check if 'q' key was pressed (non-blocking)."""
    if msvcrt.kbhit():
        key = msvcrt.getch()
        if key.lower() == b'q':
            return True
    return False


def process_search_items(file_path: str, items: list):
    """
    Opens file and processes all search items.
    items: list of tuples (screenshot_name, field_code, search_term)
    Returns: (screenshots_list, not_found_list, was_quit)
    """
    process = open_notepad_plus_with_file(file_path)
    
    try:
        screenshots = []
        not_found = []
        was_quit = False
        
        # TURBO MODE: Open Find dialog once, keep it open for all searches
        open_find_dialog()
        dialog_open = True
        
        for i, (screenshot_name, field_code, search_term) in enumerate(items, 1):
            # Check for quit key
            if check_quit_key():
                print("\n\n*** 'q' pressed - Quitting... ***")
                was_quit = True
                break
            
            # Show simple progress (overwrite same line)
            print(f"\r[{i}/{len(items)}] Processing: {field_code[:30]:<30} [Press 'q' to quit]", end='', flush=True)
            
            # Search with dialog already open (turbo mode)
            found = search_and_highlight(search_term, dialog_open=True)
            
            # Close dialog temporarily to take screenshot, then reopen
            close_find_dialog()
            time.sleep(0.03)
            
            # Select the line for screenshot
            select_current_line()
            
            # Use screenshot_name if provided, otherwise generate from field_code
            if screenshot_name and screenshot_name.lower() != 'nan':
                filename = f"{screenshot_name}"
            else:
                filename = f"search_{i}_{field_code}"
            
            screenshot_path = take_screenshot(filename)
            screenshots.append(screenshot_path)
            
            if not found:
                not_found.append((field_code, search_term))
            
            # Reopen dialog for next search (if not last item and not quitting)
            if i < len(items) and not was_quit:
                open_find_dialog()
        
        return screenshots, not_found, was_quit
        
    finally:
        print()  # New line after progress
        # Make sure dialog is closed before closing Notepad++
        try:
            pyautogui.press('escape')
            time.sleep(0.02)
        except:
            pass
        close_notepad_plus()


def main():
    parser = argparse.ArgumentParser(
        description="Notepad++ Search and Screenshot Tool",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
    python main.py --file EDI.txt --word "BHT03"
    python main.py --file EDI.txt --list BHT03 BHT04 CLM05
    python main.py --file EDI.txt --txt fields.txt
    python main.py --file EDI.txt --excel input.xlsx
    python main.py --file EDI.txt --excel input.xlsx --preview  # Preview only
        """
    )
    
    parser.add_argument('--file', '-f', required=True, help='Path to the file to open in Notepad++')
    parser.add_argument('--word', '-w', help='Single field code to search for. Example: BHT03')
    parser.add_argument('--list', '-l', nargs='+', help='List of field codes. Example: --list BHT03 BHT04 CLM05')
    parser.add_argument('--txt', '-t', help='Text file with field codes (one per line)')
    parser.add_argument('--excel', '-e', help='Excel file with "File name" and "Field" columns')
    parser.add_argument('--preview', '-p', action='store_true', help='Preview search terms without running (dry run)')
    
    args = parser.parse_args()
    
    # Resolve file path
    file_path = args.file
    if not os.path.isabs(file_path):
        file_path = os.path.join(SCRIPT_DIR, file_path)
    
    if not os.path.exists(file_path):
        print(f"Error: File not found: {file_path}")
        return 1
    
    # Parse search items based on input type
    items = []
    
    if args.excel:
        # Excel input
        excel_path = args.excel
        if not os.path.isabs(excel_path):
            excel_path = os.path.join(SCRIPT_DIR, excel_path)
        
        print(f"Loading search terms from Excel: {excel_path}")
        items = load_from_excel(excel_path)
    
    elif args.txt:
        # Text file input
        txt_path = args.txt
        if not os.path.isabs(txt_path):
            txt_path = os.path.join(SCRIPT_DIR, txt_path)
        
        print(f"Loading search terms from text file: {txt_path}")
        items = load_from_txt(txt_path)
        
    elif args.list:
        # List input
        for field_code in args.list:
            segment, element_num, sub_element, search_term = parse_field_code(field_code)
            items.append((field_code, field_code, search_term))
            
    elif args.word:
        # Single field input
        field_code = args.word.strip()
        if field_code:
            segment, element_num, sub_element, search_term = parse_field_code(field_code)
            items.append((field_code, field_code, search_term))
    
    if not items:
        print("Error: No search terms provided. Use --word, --list, --txt, or --excel")
        return 1
    
    print("=" * 80)
    print("  NOTEPAD++ SEARCH AND SCREENSHOT TOOL")
    print("=" * 80)
    print(f"File: {file_path}")
    print(f"Search items: {len(items)}")
    print(f"Screenshot folder: {SCREENSHOT_FOLDER}")
    print("-" * 80)
    print(f"{'#':<5} {'GDF Field':<40} {'EDI Field':<30} {'Search Term':<20}")
    print("-" * 80)
    for i, (name, code, term) in enumerate(items, 1):
        # Truncate long names for display
        name_display = name[:38] + '..' if len(name) > 40 else name
        code_display = code[:28] + '..' if len(code) > 30 else code
        print(f"{i:<5} {name_display:<40} {code_display:<30} {term:<20}")
    print("-" * 80)
    
    # If preview mode, just show the mapping and exit
    if args.preview:
        print("\n*** PREVIEW MODE - No search performed ***")
        print(f"\nTotal items: {len(items)}")
        return 0
    
    print("\nStarting in 1 second... (Don't move the mouse! Press 'q' to quit anytime)")
    time.sleep(1)  # Reduced from 2s
    
    screenshots, not_found, was_quit = process_search_items(file_path, items)
    
    print("\n" + "=" * 60)
    if was_quit:
        print("  STOPPED BY USER")
    else:
        print("  COMPLETE")
    print("=" * 60)
    print(f"Total screenshots: {len(screenshots)}")
    print(f"Found: {len(screenshots) - len(not_found)}")
    print(f"Not found: {len(not_found)}")
    
    # Log not found fields in table format
    if not_found:
        print("\n" + "=" * 70)
        print("  NOT FOUND ITEMS")
        print("=" * 70)
        print(f"{'#':<5} {'Field':<35} {'Searched':<25}")
        print("-" * 70)
        for idx, (field_code, search_term) in enumerate(not_found, 1):
            field_display = field_code[:33] + '..' if len(field_code) > 35 else field_code
            term_display = search_term[:23] + '..' if len(search_term) > 25 else search_term
            print(f"{idx:<5} {field_display:<35} {term_display:<25}")
        print("-" * 70)
        
        # Save to log file
        log_path = os.path.join(SCREENSHOT_FOLDER, "not_found.log")
        with open(log_path, 'w', encoding='utf-8') as f:
            f.write(f"Not Found Fields - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write("=" * 70 + "\n")
            f.write(f"{'#':<5} {'Field':<35} {'Searched':<25}\n")
            f.write("-" * 70 + "\n")
            for idx, (field_code, search_term) in enumerate(not_found, 1):
                f.write(f"{idx:<5} {field_code:<35} {search_term:<25}\n")
        
        print(f"\nLog saved: {log_path}")
    
    return 0


if __name__ == "__main__":
    exit(main())
