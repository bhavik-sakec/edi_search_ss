"""
EDI Segment Search Tool - Notepad++ Version with Screenshot
Takes screenshots for ALL segments, filename from GDF_Field column.

Usage:
    python edi_search_tool.py --file <edi_file> --excel <excel_file>
    python edi_search_tool.py --file <edi_file> --segment BHT03

Examples:
    python edi_search_tool.py --file EDI.txt --excel book.xlsx
    python edi_search_tool.py --file EDI.txt --segment BHT03
"""

import os
import sys
import re
import argparse
import subprocess
import time
from datetime import datetime
from pathlib import Path

# Try to import required libraries
try:
    import pyautogui
    from PIL import Image, ImageGrab, ImageDraw
    import pandas as pd
    import win32gui
    import ctypes
    from ctypes import wintypes
    LIBS_AVAILABLE = True
except ImportError as e:
    print(f"❌ Missing library: {e}")
    print("   Run: pip install pyautogui pillow pandas openpyxl")
    sys.exit(1)

# Disable pyautogui pause for speed
pyautogui.PAUSE = 0.1

# Notepad++ path
NOTEPAD_PATH = r"C:\Program Files\Notepad++\notepad++.exe"

# Screenshot folder
SCREENSHOT_FOLDER = r"C:\Users\bhavi\Downloads\office work\edi\Screenshot"

# Known EDI segments
KNOWN_SEGMENTS = [
    'ISA', 'GS', 'ST', 'BHT', 'NM1', 'N3', 'N4', 'REF', 'PER', 'HL',
    'SBR', 'DMG', 'PAT', 'CLM', 'DTP', 'CL1', 'HI', 'PRV', 'SV1', 'SV2',
    'LX', 'LIN', 'CTP', 'CN1', 'HCP', 'AMT', 'SE', 'GE', 'IEA'
]


def extract_segment_id(edi_ref: str) -> str:
    """
    Extract the segment ID from an EDI reference.
    
    Examples:
        'BHT03' -> 'BHT'
        '2010AANM109' -> 'NM1'
        '2300HI01-2 -- BE' -> 'HI'
        '2400SV202-3' -> 'SV2'
    """
    if not edi_ref or not isinstance(edi_ref, str):
        return None
    
    edi_ref = edi_ref.strip().upper()
    
    # Handle compound fields - take first part
    if '+' in edi_ref:
        edi_ref = edi_ref.split('+')[0].strip()
    
    # Remove qualifiers (-- BE, -BG, etc.)
    edi_ref = re.sub(r'\s*--?\s*[A-Z]{2,3}(?:/[A-Z]{2,3})*\s*$', '', edi_ref)
    
    # Remove parenthetical notes
    edi_ref = re.sub(r'\s*\([^)]*\)\s*$', '', edi_ref)
    
    # Remove "when" conditions
    edi_ref = re.sub(r'\s+when\s+.*$', '', edi_ref, flags=re.IGNORECASE)
    
    # Try to find a known segment in the reference
    for seg in sorted(KNOWN_SEGMENTS, key=len, reverse=True):
        pattern = rf'(?:^|\d{{4}}[A-Z]*)({seg})(?:\d|$)'
        match = re.search(pattern, edi_ref)
        if match:
            return seg
    
    # Fallback: check if starts with 2-3 letter segment
    simple_match = re.match(r'^([A-Z]{2,3})(?:\d|$)', edi_ref)
    if simple_match:
        return simple_match.group(1)
    
    return None


def check_segment_exists(file_path: str, segment_id: str) -> bool:
    """Check if a segment exists in the EDI file."""
    try:
        with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
            content = f.read()
        
        pattern = rf'(?:^|~|\n){segment_id}\*'
        return bool(re.search(pattern, content))
    except Exception:
        return False


def ensure_screenshot_folder():
    """Create screenshot folder if it doesn't exist."""
    if not os.path.exists(SCREENSHOT_FOLDER):
        os.makedirs(SCREENSHOT_FOLDER)


def close_notepad_without_saving():
    """Close the current Notepad++ instance."""
    # Since we use -multiInst and -ro, we can just close the active window safely
    pyautogui.hotkey('alt', 'f4')
    time.sleep(0.5)
    
    # We remove the 'n' press because -ro (Read Only) prevents the save dialog.
    # This keeps the terminal clean.


def get_notepad_hwnd():
    """Find the Notepad++ window handle."""
    return win32gui.FindWindow("Notepad++", None)


def get_caret_position():
    """Gets the current caret position on screen."""
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
    
    # Fallback
    pos = pyautogui.position()
    return pos[0], pos[1], 20


def take_screenshot_with_red_box(filename: str, bounds: tuple = None) -> str:
    """
    Take a screenshot with red box based on caret position.
    The filename is provided, bounds argument is ignored (kept for compatibility).
    """
    ensure_screenshot_folder()
    
    # Get Notepad++ Window info
    hwnd = get_notepad_hwnd()
    time.sleep(0.08)
    
    caret_x, caret_y, caret_height = get_caret_position()
    
    if hwnd:
        try:
            rect = win32gui.GetWindowRect(hwnd)
            win_left, win_top, win_right, win_bottom = rect
        except Exception:
            # Fallback if window not found or error
            win_left, win_top = 0, 0
    else:
        win_left, win_top = 0, 0
    
    screenshot = pyautogui.screenshot()
    
    # Draw rectangle around selected line
    draw = ImageDraw.Draw(screenshot)
    padding = 3
    
    # Logic from user request:
    # Left bound = Window Left + 60 (skipping line numbers)
    # Right bound = Caret X + 10 (end of selection/line)
    rect_left = win_left + 60
    rect_right = caret_x + 10
    rect_top = caret_y - padding
    rect_bottom = caret_y + caret_height + padding
    
    clean_filename = re.sub(r'[^\w\-]', '_', filename)[:50]
    filepath = os.path.join(SCREENSHOT_FOLDER, f"{clean_filename}.png")
    
    for i in range(3):
        draw.rectangle(
            [rect_left - i, rect_top - i, rect_right + i, rect_bottom + i],
            outline="red"
        )
    
    screenshot.save(filepath)
    
    return filepath


def search_and_screenshot(file_path: str, search_term: str, filename: str, notepad_open: bool = False) -> str:
    """
    Search in Notepad++ and take screenshot.
    """
    abs_path = os.path.abspath(file_path)
    
    if not notepad_open:
        subprocess.Popen([NOTEPAD_PATH, "-ro", "-multiInst", "-nosession", abs_path])
        time.sleep(1.5)
    
    # Ensure tool window is center-focused before searching
    pyautogui.click(pyautogui.size().width // 2, pyautogui.size().height // 2)
    time.sleep(0.2)
    
    # Open Find dialog (Ctrl+F)
    pyautogui.hotkey('ctrl', 'f')
    time.sleep(0.5)
    
    # Clear and type search term
    pyautogui.hotkey('ctrl', 'a')
    time.sleep(0.1)
    pyautogui.typewrite(search_term, interval=0.01)
    time.sleep(0.2)
    
    # Search
    pyautogui.press('enter')
    time.sleep(0.3)
    
    # Close Find dialog
    pyautogui.press('escape')
    time.sleep(0.3)
    
    # Select the line (Home, then Shift+End)
    # This places the caret at the END of the line, which is crucial for our new box logic
    pyautogui.press('home')
    time.sleep(0.1)
    pyautogui.hotkey('shift', 'end')
    time.sleep(0.4)

    # Note: We no longer need to find bounds before deselecting.
    # Actually, we don't even need to deselect if we just want the caret position.
    # But the user's logic draws the box. 
    # If the text is highlighted blue, the red box will be around it.
    
    screenshot_path = take_screenshot_with_red_box(filename)
    
    return screenshot_path


def process_excel(file_path: str, excel_path: str, row_range: str = None):
    """
    Process ALL segments from Excel file.
    - Column A (GDF_Field) = Screenshot filename
    - Original_EDI_Field = What to search for
    - row_range: Optional range string (e.g., '1-10', '5-', '-20')
    """
    # Check files exist
    if not os.path.exists(file_path):
        print(f"❌ EDI file not found: {file_path}")
        sys.exit(1)
    
    if not os.path.exists(excel_path):
        print(f"❌ Excel file not found: {excel_path}")
        sys.exit(1)
    
    if not os.path.exists(NOTEPAD_PATH):
        print(f"❌ Notepad++ not found: {NOTEPAD_PATH}")
        sys.exit(1)
    
    # Read Excel
    df = pd.read_excel(excel_path)
    
    # Check required columns
    if 'GDF_Field' not in df.columns:
        print(f"❌ Column 'GDF_Field' not found in Excel")
        print(f"   Available columns: {df.columns.tolist()}")
        sys.exit(1)
    
    if 'Original_EDI_Field' not in df.columns:
        print(f"❌ Column 'Original_EDI_Field' not found in Excel")
        print(f"   Available columns: {df.columns.tolist()}")
        sys.exit(1)
    
    # Apply range filter if provided
    total_rows = len(df)
    if row_range:
        try:
            start_idx = 0
            end_idx = total_rows
            
            if '-' in row_range:
                parts = row_range.split('-')
                if parts[0]:
                    start_idx = int(parts[0]) - 1
                if parts[1]:
                    end_idx = int(parts[1])
            else:
                # Single number
                start_idx = int(row_range) - 1
                end_idx = int(row_range)
            
            # Bounds check
            start_idx = max(0, start_idx)
            end_idx = min(total_rows, end_idx)
            
            if start_idx >= end_idx:
                print(f"⚠️ Invalid range: {row_range}. Processing all rows.")
            else:
                df = df.iloc[start_idx:end_idx]
                print(f"📍 Applied range: {row_range} (Rows {start_idx+1} to {end_idx})")
                
        except ValueError:
            print(f"❌ Invalid range format: {row_range}. Expected format: 'start-end', 'start-', or '-end'.")
            sys.exit(1)
    
    print(f"\n{'='*60}")
    print("EDI SEGMENT SEARCH - BATCH PROCESSING")
    print(f"{'='*60}")
    print(f"📄 EDI File: {file_path}")
    print(f"📊 Excel File: {excel_path}")
    print(f"📋 Rows to process: {len(df)}")
    print(f"📁 Screenshots: {SCREENSHOT_FOLDER}")
    print(f"\n📝 Filename from: GDF_Field (Column A)")
    print(f"🔍 Search from: Original_EDI_Field")
    
    ensure_screenshot_folder()
    
    # Track results
    not_found = []
    found_count = 0
    screenshot_count = 0
    abs_path = os.path.abspath(file_path)
    
    # Open Notepad++ once
    # -ro: Read Only, -multiInst: Separate instance, -nosession: No history
    subprocess.Popen([NOTEPAD_PATH, "-ro", "-multiInst", "-nosession", abs_path])
    time.sleep(1.5)
    
    print(f"\n⏳ Processing {len(df)} rows...")
    
    for idx, row in df.iterrows():
        gdf_field = str(row.get('GDF_Field', '')).strip()
        edi_ref = str(row.get('Original_EDI_Field', '')).strip()
        
        if not gdf_field or gdf_field == 'nan':
            continue
        if not edi_ref or edi_ref == 'nan':
            continue
        
        # Extract segment ID from EDI reference
        segment_id = extract_segment_id(edi_ref)
        
        if not segment_id:
            not_found.append(f"{gdf_field} ({edi_ref})")
            continue
        
        # Check if segment exists in EDI file
        if not check_segment_exists(file_path, segment_id):
            not_found.append(f"{gdf_field} ({edi_ref})")
            continue
        
        # Found - search and screenshot
        search_term = f"{segment_id}*"
        search_and_screenshot(file_path, search_term, gdf_field, notepad_open=True)
        found_count += 1
        screenshot_count += 1
        
        # Progress indicator
        if screenshot_count % 10 == 0:
            print(f"   📸 {screenshot_count} screenshots taken...")
        
        # Small delay between searches
        time.sleep(0.2)
    
    # Print results
    print(f"\n{'='*60}")
    print("RESULTS")
    print(f"{'='*60}")
    print(f"✅ Found: {found_count}")
    print(f"📸 Screenshots taken: {screenshot_count}")
    print(f"❌ Not Found: {len(not_found)}")
    
    if not_found:
        print(f"\n{'='*60}")
        print("❌ NOT FOUND SEGMENTS:")
        print(f"{'='*60}")
        for seg in not_found:
            print(f"   • {seg}")
    
    # Close Notepad++ without saving
    print(f"\n🔒 Closing Notepad++ without saving...")
    close_notepad_without_saving()
    
    print(f"\n📁 Screenshots saved to: {SCREENSHOT_FOLDER}")


def process_single_segment(file_path: str, segment_ref: str):
    """Process a single segment."""
    segment_id = extract_segment_id(segment_ref)
    
    if not segment_id:
        print(f"❌ Could not parse segment: {segment_ref}")
        sys.exit(1)
    
    print(f"\n{'='*50}")
    print("EDI SEGMENT SEARCH - SINGLE")
    print(f"{'='*50}")
    print(f"📄 File: {file_path}")
    print(f"🔍 Reference: {segment_ref}")
    print(f"📍 Segment ID: {segment_id}")
    
    if not check_segment_exists(file_path, segment_id):
        print(f"\n❌ NOT FOUND: '{segment_id}' not in EDI file")
        sys.exit(1)
    
    print(f"\n✅ Segment '{segment_id}' found!")
    
    # Open Notepad++
    abs_path = os.path.abspath(file_path)
    subprocess.Popen([NOTEPAD_PATH, "-ro", "-multiInst", "-nosession", abs_path])
    time.sleep(1.5)
    
    search_term = f"{segment_id}*"
    screenshot_path = search_and_screenshot(file_path, search_term, segment_ref, notepad_open=True)
    
    print(f"\n📸 Screenshot saved: {screenshot_path}")
    
    # Close Notepad++ without saving
    print(f"\n🔒 Closing Notepad++ without saving...")
    close_notepad_without_saving()


def main():
    parser = argparse.ArgumentParser(
        description='EDI Segment Search - Notepad++ with Screenshot',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Process ALL segments from Excel
  # - Filename from GDF_Field (Column A)
  # - Search using Original_EDI_Field
  python edi_search_tool.py --file EDI.txt --excel book.xlsx
  
  # Process single segment
  python edi_search_tool.py --file EDI.txt --segment BHT03
        """
    )
    
    parser.add_argument('--file', '-f', required=True,
                        help='Path to the EDI file')
    parser.add_argument('--excel', '-e',
                        help='Excel file with GDF_Field and Original_EDI_Field columns')
    parser.add_argument('--segment', '-s',
                        help='Single segment to search')
    parser.add_argument('--range', '-r',
                        help='Specific range of rows to process (e.g., 1-10, 5-, -20). 1=A2.')
    
    args = parser.parse_args()
    
    if args.excel:
        process_excel(args.file, args.excel, row_range=args.range)
    elif args.segment:
        process_single_segment(args.file, args.segment)
    else:
        print("❌ Please provide either --excel or --segment")
        parser.print_help()
        sys.exit(1)


if __name__ == "__main__":
    main()
