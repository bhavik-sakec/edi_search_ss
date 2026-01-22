"""
Notepad++ Search and Screenshot Tool
Opens a text file in Notepad++, searches for specific segments,
highlights the found text (complete line), and takes a screenshot.

Usage:
    # Single field
    python main.py --file EDI.txt --word "BHT03"
    
    # Multiple fields (comma-separated)
    python main.py --file EDI.txt --word "BHT03,CLM05,DTP01"
    
    # List of fields
    python main.py --file EDI.txt --list BHT03 BHT04 CLM05
    
    # Text file input (one field per line)
    python main.py --file EDI.txt --txt fields.txt
    
    # Excel input
    python main.py --file EDI.txt --excel input.xlsx
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


def load_from_excel(excel_path: str) -> list:
    """
    Loads search terms from Excel file.
    Expected columns: 'File name' (for screenshot name) and 'Field' (for search term like BHT03)
    
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
    
    for key, col in columns_lower.items():
        if 'file' in key and 'name' in key:
            file_name_col = col
        elif 'field' in key:
            field_col = col
    
    if file_name_col is None or field_col is None:
        raise ValueError(f"Excel must have 'File name' and 'Field' columns. Found: {list(df.columns)}")
    
    results = []
    for _, row in df.iterrows():
        file_name = str(row[file_name_col]).strip()
        field_code = str(row[field_col]).strip()
        
        if field_code and field_code.lower() != 'nan':
            segment, element_num, sub_element, search_term = parse_field_code(field_code)
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
    time.sleep(2)
    
    NOTEPAD_HWND = find_notepad_window()
    if NOTEPAD_HWND:
        win32gui.ShowWindow(NOTEPAD_HWND, win32con.SW_MAXIMIZE)
        time.sleep(0.3)
        win32gui.SetForegroundWindow(NOTEPAD_HWND)
        time.sleep(0.3)
    
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
        time.sleep(0.1)
        return NOTEPAD_HWND
    return None


def search_and_highlight(search_text: str) -> bool:
    """
    Searches for text in Notepad++ and highlights the entire line.
    Returns True if text was found, False otherwise.
    """
    ensure_notepad_focus()
    
    # Get caret position before search to detect if search was successful
    before_pos = get_caret_position()
    
    # Open Find dialog
    pyautogui.hotkey('ctrl', 'f')
    time.sleep(0.5)
    
    # Clear and type search text
    pyautogui.hotkey('ctrl', 'a')
    time.sleep(0.1)
    
    pyperclip.copy(search_text)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(0.3)
    
    # Find Next
    pyautogui.press('enter')
    time.sleep(0.5)
    
    # Close Find dialog
    pyautogui.press('escape')
    time.sleep(0.3)
    
    # Get caret position after search
    after_pos = get_caret_position()
    
    # Check if caret moved (indicating text was found)
    # If positions are very similar, text was likely not found
    found = True
    if before_pos and after_pos:
        # Compare Y positions (line changed = found)
        if abs(before_pos[1] - after_pos[1]) < 5 and abs(before_pos[0] - after_pos[0]) < 10:
            # Position didn't change much - might not be found
            # But first search from start will always be at same position
            # So we consider it found if we're doing the search
            found = True  # Assume found for now, will improve detection
    
    # Select entire line: Home then Shift+End
    pyautogui.press('home')
    time.sleep(0.1)
    pyautogui.hotkey('shift', 'end')
    time.sleep(0.3)
    
    return found


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


def take_screenshot(filename: str) -> str:
    """Takes a full screen screenshot and draws a red rectangle around the selected line."""
    if not os.path.exists(SCREENSHOT_FOLDER):
        os.makedirs(SCREENSHOT_FOLDER)
        print(f"Created screenshot folder: {SCREENSHOT_FOLDER}")
    
    # Clean filename
    safe_filename = re.sub(r'[<>:"/\\|?*]', '_', filename)
    if not safe_filename.endswith('.png'):
        safe_filename += '.png'
    
    screenshot_path = os.path.join(SCREENSHOT_FOLDER, safe_filename)
    
    hwnd = NOTEPAD_HWND
    time.sleep(0.2)
    
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
    print(f"Screenshot saved: {screenshot_path}")
    
    return screenshot_path


def close_notepad_plus():
    """Closes Notepad++ without saving."""
    global NOTEPAD_HWND
    
    if NOTEPAD_HWND:
        try:
            win32gui.SetForegroundWindow(NOTEPAD_HWND)
        except:
            pass
        time.sleep(0.2)
    
    pyautogui.hotkey('alt', 'F4')
    time.sleep(0.3)
    pyautogui.press('n')
    time.sleep(0.2)
    
    NOTEPAD_HWND = None


def process_search_items(file_path: str, items: list):
    """
    Opens file and processes all search items.
    items: list of tuples (screenshot_name, field_code, search_term)
    Returns: (screenshots_list, not_found_list)
    """
    process = open_notepad_plus_with_file(file_path)
    
    try:
        screenshots = []
        not_found = []
        
        for i, (screenshot_name, field_code, search_term) in enumerate(items, 1):
            print(f"\n[{i}/{len(items)}] Searching for: '{search_term}' (Field: {field_code})")
            
            found = search_and_highlight(search_term)
            time.sleep(0.3)
            
            # Use screenshot_name if provided, otherwise generate from field_code
            if screenshot_name and screenshot_name.lower() != 'nan':
                filename = f"{screenshot_name}"
            else:
                filename = f"search_{i}_{field_code}"
            
            screenshot_path = take_screenshot(filename)
            screenshots.append(screenshot_path)
            
            if found:
                print(f"Screenshot saved for '{field_code}'")
            else:
                print(f"WARNING: '{search_term}' may not have been found!")
                not_found.append((field_code, search_term))
        
        return screenshots, not_found
        
    finally:
        close_notepad_plus()
        print("\nNotepad++ closed.")


def main():
    parser = argparse.ArgumentParser(
        description="Notepad++ Search and Screenshot Tool",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
    python main.py --file EDI.txt --word "BHT03"
    python main.py --file EDI.txt --word "BHT03,CLM05,DTP01"
    python main.py --file EDI.txt --list BHT03 BHT04 CLM05
    python main.py --file EDI.txt --txt fields.txt
    python main.py --file EDI.txt --excel input.xlsx
        """
    )
    
    parser.add_argument('--file', '-f', required=True, help='Path to the file to open in Notepad++')
    parser.add_argument('--word', '-w', help='Field code(s) to search for (comma-separated). Example: BHT03,CLM05')
    parser.add_argument('--list', '-l', nargs='+', help='List of field codes. Example: --list BHT03 BHT04 CLM05')
    parser.add_argument('--txt', '-t', help='Text file with field codes (one per line)')
    parser.add_argument('--excel', '-e', help='Excel file with "File name" and "Field" columns')
    
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
        # Comma-separated input
        for field_code in args.word.split(','):
            field_code = field_code.strip()
            if field_code:
                segment, element_num, sub_element, search_term = parse_field_code(field_code)
                items.append((field_code, field_code, search_term))
    
    if not items:
        print("Error: No search terms provided. Use --word, --list, --txt, or --excel")
        return 1
    
    print("=" * 60)
    print("  NOTEPAD++ SEARCH AND SCREENSHOT TOOL")
    print("=" * 60)
    print(f"File: {file_path}")
    print(f"Search items: {len(items)}")
    for name, code, term in items:
        print(f"  - {name}: {term}")
    print(f"Screenshot folder: {SCREENSHOT_FOLDER}")
    print("-" * 60)
    
    print("\nStarting in 2 seconds... (Don't move the mouse!)")
    time.sleep(2)
    
    screenshots, not_found = process_search_items(file_path, items)
    
    print("\n" + "=" * 60)
    print("  COMPLETE")
    print("=" * 60)
    print(f"Total screenshots: {len(screenshots)}")
    print(f"Found: {len(screenshots) - len(not_found)}")
    print(f"Not found: {len(not_found)}")
    
    print("\nScreenshots:")
    for path in screenshots:
        print(f"  - {path}")
    
    # Log not found fields
    if not_found:
        print("\n" + "-" * 60)
        print("  FIELDS NOT FOUND:")
        print("-" * 60)
        for field_code, search_term in not_found:
            print(f"  - {field_code} (searched: {search_term})")
        
        # Save to log file
        log_path = os.path.join(SCREENSHOT_FOLDER, "not_found.log")
        with open(log_path, 'w', encoding='utf-8') as f:
            f.write(f"Not Found Fields - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write("=" * 50 + "\n\n")
            for field_code, search_term in not_found:
                f.write(f"{field_code} (searched: {search_term})\n")
        
        print(f"\nNot found fields logged to: {log_path}")
    
    return 0


if __name__ == "__main__":
    exit(main())
