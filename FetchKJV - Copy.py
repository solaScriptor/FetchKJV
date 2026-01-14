from pynput import keyboard as pynput_keyboard
from pynput.keyboard import Key, KeyCode, Controller
import pyperclip
import pythonbible as bible
import tkinter as tk
from tkinter import colorchooser, scrolledtext
import json
import time
import threading
import os
import traceback
import sys
import re

# System tray support
import pystray
from PIL import Image

# Windows clipboard RTF support
import win32clipboard
import win32con

import win32event
import win32api
import winerror

# -----------------------------
# HOTKEY FORMATTER
# -----------------------------

def format_hotkey(hk):
    parts = []
    if hk["ctrl"]:
        parts.append("Ctrl")
    if hk["alt"]:
        parts.append("Alt")
    if hk["shift"]:
        parts.append("Shift")
    parts.append(hk["key"].upper())
    return " + ".join(parts)

# -----------------------------
# SINGLE INSTANCE CHECK
# -----------------------------

mutex_name = "FetchKJV_SingleInstanceMutex"

# Try to create a named mutex
mutex = win32event.CreateMutex(None, False, mutex_name)

# If the mutex already exists, exit immediately
if win32api.GetLastError() == winerror.ERROR_ALREADY_EXISTS:
    print("Another instance of FetchKJV is already running.")
    sys.exit(0)

# -----------------------------
# MAIN WINDOW
# -----------------------------

hidden_root = tk.Tk()
hidden_root.withdraw()  # Hide the main window

def resource_path(relative_path):
    """Get absolute path to resource, works for dev and PyInstaller."""
    if hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), relative_path)

# -----------------------------
# SETTINGS LOADER
# -----------------------------

SETTINGS_PATH = resource_path("settings.json")

def load_settings():
    """Load settings.json with safe defaults."""
    defaults = {
        "hotkey": {
            "key": "f21",
            "ctrl": False,
            "alt": False,
            "shift": False
        },
        "auto_close_seconds": 3,
        "popup": {
            "bg_large": "light yellow",
            "bg_small": "light green",
            "text_color_small": "dark green",
            "font_large": ["Segoe UI", 11],
            "font_small": ["Segoe UI", 14, "bold"],
            "width_large": 90,
            "height_large": 25
        },
        "show_welcome": True
    }

    try:
        with open(SETTINGS_PATH, "r", encoding="utf-8") as f:
            user_settings = json.load(f)

            # -----------------------------
            # MIGRATE OLD HOTKEY FORMAT
            # -----------------------------
            hk = user_settings.get("hotkey")

            # If the hotkey is a string (old format), replace with default structured dict
            if isinstance(hk, str):
                user_settings["hotkey"] = defaults["hotkey"]

            # If the hotkey is missing keys, also repair it
            if isinstance(hk, dict):
                for key in ("key", "ctrl", "alt", "shift"):
                    if key not in hk:
                        user_settings["hotkey"] = defaults["hotkey"]
                        break

            return merge_settings(defaults, user_settings)

    except Exception as e:
        print("Could not load settings.json, using defaults:", e)
        return defaults


def merge_settings(defaults, user):
    """Recursively merge user settings over defaults."""
    for key, value in defaults.items():
        if key not in user:
            user[key] = value
        elif isinstance(value, dict):
            merge_settings(value, user[key])
    return user

# -----------------------------
# WELCOME POP-UP
# -----------------------------

def show_welcome_popup():
    win = tk.Toplevel(hidden_root)
    win.title("Welcome to FetchKJV")
    win.geometry("420x380")
    win.resizable(False, False)
    win.attributes("-topmost", True)

    frame = tk.Frame(win, padx=20, pady=20)
    frame.pack(expand=True, fill="both")

    message = (
        "Thank you for using FetchKJV!\n\n"
        "Select any text with your mouse,\n"
        "press your hotkey, and the text of any Bible\n"
        "verses will appear in a pop up window.\n"
        "Press the hotkey again to copy the text\n\n"
        "The hotkey can be changed from the settings menu.\n"
        "In future you can access settings from the taskbar.\n\n"
        "Click 'Get Started' to begin."
    )

    label = tk.Label(frame, text=message, justify="left", font=("Segoe UI", 11))
    label.pack(pady=(0, 20))

    # Checkbox variable
    dont_show_var = tk.BooleanVar(value=False)

    chk = tk.Checkbutton(
        frame,
        text="Don't show this again",
        variable=dont_show_var,
        font=("Segoe UI", 10)
    )
    chk.pack(pady=(0, 20))

    # --- BUTTON ROW ---
    btn_frame = tk.Frame(frame)
    btn_frame.pack(pady=(10, 0))

    def close_popup():
        if dont_show_var.get():
            settings["show_welcome"] = False
            with open(SETTINGS_PATH, "w", encoding="utf-8") as f:
                json.dump(settings, f, indent=4)
        win.destroy()

    def open_settings():
        close_popup()
        open_settings_window()

    # Get Started button
    tk.Button(
        btn_frame,
        text="Get Started",
        font=("Segoe UI", 11),
        width=14,
        command=close_popup
    ).pack(side="left", padx=5)

    # Open Settings button
    tk.Button(
        btn_frame,
        text="Open Settings",
        font=("Segoe UI", 11),
        width=14,
        command=open_settings
    ).pack(side="left", padx=5)

    win.protocol("WM_DELETE_WINDOW", close_popup)


settings = load_settings()

# Show welcome popup on first launch
if settings.get("show_welcome", True):
    hidden_root.after(200, show_welcome_popup)

def reload_settings():
    global settings, AUTO_CLOSE_SECONDS, leave_timer, current_root, awaiting_second_press
    print("Reloading settings...")

    # Cancel any pending auto-close timer BEFORE reloading settings
    if current_root and leave_timer:
        try:
            current_root.after_cancel(leave_timer)
        except:
            pass
        leave_timer = None
        awaiting_second_press = False

    settings = load_settings()

    # Reapply key settings
    AUTO_CLOSE_SECONDS = settings["auto_close_seconds"]
    print("Settings reloaded successfully.")

# -----------------------------
# TRAY ICON
# -----------------------------

def create_tray_icon():
    """Creates a Windows system tray icon with an Exit option."""
    try:
        icon_image = Image.open(resource_path("FetchKJV.ico"))
    except Exception as e:
        print("Could not load tray icon:", e)
        return

    def on_exit(icon, item):
        icon.stop()
        os._exit(0)

    menu = pystray.Menu(
        pystray.MenuItem("Settings", lambda icon, item: open_settings_window()),
        pystray.MenuItem("Exit", on_exit)
    )


    icon = pystray.Icon(
        "FetchKJV",
        icon_image,
        "FetchKJV",
        menu
    )

    threading.Thread(target=icon.run, daemon=True).start()

# -----------------------------
# CONFIG
# -----------------------------

AUTO_CLOSE_SECONDS = settings["auto_close_seconds"]
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
KJV_JSON_PATH = resource_path("kjv.json")

last_verses_clean = None
current_root = None
leave_timer = None
awaiting_second_press = False

kb_controller = Controller()

print("Starting FetchKJV...")
create_tray_icon()

# -----------------------------
# LOAD KJV BIBLE
# -----------------------------

try:
    if not os.path.exists(KJV_JSON_PATH):
        raise FileNotFoundError(f"kjv.json not found at: {KJV_JSON_PATH}")
    
    with open(KJV_JSON_PATH, 'r', encoding='utf-8') as f:
        data = json.load(f)
    
    if 'verses' not in data:
        raise ValueError("kjv.json is missing the 'verses' key")
    
    verses = data['verses']
    # Build fast lookup index: (book, chapter, verse) → text
    verse_index = {
        (entry['book_name'], entry['chapter'], entry['verse']): entry['text']
        for entry in verses
    }
    print("KJV Bible loaded successfully (offline mode).")
except Exception as e:
    print("FATAL ERROR - Could not load Bible data:")
    print(str(e))
    print("\nPress Enter to exit...")
    input()
    sys.exit(1)

# -----------------------------
# RTF CLIPBOARD FUNCTIONS
# -----------------------------

def copy_rtf_to_clipboard(rtf_text, plain_text):
    """Places RTF + clean plain text onto the Windows clipboard."""
    win32clipboard.OpenClipboard()
    try:
        win32clipboard.EmptyClipboard()

        CF_RTF = win32clipboard.RegisterClipboardFormat("Rich Text Format")

        # RTF for Word and other rich editors
        win32clipboard.SetClipboardData(CF_RTF, rtf_text.encode('utf-8'))

        # Clean plain text for Notepad and simple editors
        win32clipboard.SetClipboardData(win32con.CF_UNICODETEXT, plain_text)

    finally:
        win32clipboard.CloseClipboard()



def convert_brackets_to_rtf(text):
    r"""Converts [bracketed text] → {\i italic text}."""
    def repl(match):
        inner = match.group(1)
        return r"{\i " + inner + "}"
    return re.sub(r"\[(.*?)\]", repl, text)


def build_rtf_document(rtf_body):
    """Wraps RTF body into a valid RTF document."""
    return r"{\rtf1\ansi " + rtf_body + "}"

def strip_brackets(text):
    """Removes [bracketed] text markers for plain-text clipboard output."""
    return re.sub(r"\[(.*?)\]", r"\1", text)

# -----------------------------
# CORE FUNCTIONS
# -----------------------------

def get_selected_text():
    original_clipboard = pyperclip.paste()

    # ---------------------------------------------------
    # Release Alt before copying (fixes Alt hotkey issue)
    # ---------------------------------------------------
    try:
        kb_controller.release(Key.alt)
    except: pass
    try:
        kb_controller.release(Key.alt_l)
    except: pass
    try:
        kb_controller.release(Key.alt_r)
    except: pass
    try:
        kb_controller.release(Key.alt_gr)
    except: pass

    # Give Windows time to exit menu mode
    time.sleep(0.05)

    # ---------------------------------------------------
    # Perform Ctrl+C to capture selected text
    # ---------------------------------------------------
    with kb_controller.pressed(Key.ctrl):
        kb_controller.press('c')
        kb_controller.release('c')

    time.sleep(0.1)

    selected = pyperclip.paste()

    # Restore original clipboard
    pyperclip.copy(original_clipboard)

    return selected.strip()



def get_verse_text(book_name, chapter, verse_start, verse_end=None):
    texts = []
    end = verse_end if verse_end else verse_start

    for v in range(verse_start, end + 1):
        key = (book_name, chapter, v)
        verse_text = verse_index.get(key)

        if verse_text:
            texts.append(f"{v} {verse_text}")
        else:
            texts.append(f"{v} [Verse not found]")

    return " ".join(texts).strip()



def on_mouse_enter(root):
    global leave_timer
    if leave_timer:
        root.after_cancel(leave_timer)
        leave_timer = None


def on_mouse_leave(root):
    global leave_timer
    leave_timer = root.after(AUTO_CLOSE_SECONDS * 1000, lambda: safe_close(root))


def safe_close(root):
    global current_root, leave_timer, awaiting_second_press

    # Cancel any pending auto-close timer
    if leave_timer:
        try:
            root.after_cancel(leave_timer)
        except:
            pass
        leave_timer = None

    # Now safely destroy the window
    try:
        if root.winfo_exists():
            root.destroy()
    except:
        pass

    current_root = None
    awaiting_second_press = False


def process_text():
    global last_verses_clean, current_root, awaiting_second_press

    try:
        # SECOND PRESS — only valid if popup is still open
        if awaiting_second_press and current_root and current_root.winfo_exists():
            rtf_body = convert_brackets_to_rtf(last_verses_clean)
            rtf_full = build_rtf_document(rtf_body)
            plain_text = strip_brackets(last_verses_clean)
            copy_rtf_to_clipboard(rtf_full, plain_text)


            current_root.after(0, lambda: safe_close(current_root))
            show_popup("Bible verses copied to clipboard!", title="Copied!", small=True)

            awaiting_second_press = False
            return

        # FIRST PRESS — read selection and show popup
        selected = get_selected_text()
        if not selected:
            return

        references = bible.get_references(selected)
        if not references:
            return

        formatted_refs = bible.format_scripture_references(references)
        display_parts = [f"{formatted_refs}\n\n"]
        clean_parts = []

        for ref in references:
            ref_str = bible.format_single_reference(ref)
            book_title = ref.book.title
            
            verse_text = get_verse_text(
                book_title,
                ref.start_chapter,
                ref.start_verse,
                ref.end_verse
            )
            if ref.end_chapter and ref.end_chapter > ref.start_chapter:
                extra = get_verse_text(
                    book_title,
                    ref.end_chapter,
                    1,
                    ref.end_verse
                )
                if "[not found]" not in extra.lower():
                    verse_text += " " + extra
            
            display_parts.append(f"{ref_str}:\n{verse_text}\n\n")
            clean_parts.append(verse_text)

        display_text = "".join(display_parts).strip()
        last_verses_clean = "\n\n".join(clean_parts).strip()

        awaiting_second_press = True
        show_popup(display_text, title="Bible Verses (KJV)", small=False)

    except Exception as e:
        error_msg = traceback.format_exc()
        print("Error in hotkey processing:")
        print(error_msg)
        show_popup(f"Error:\n{str(e)}", title="Error", small=True)


# -----------------------------
# BIBLE POP-UP
# -----------------------------

def show_popup(text, title="Bible Verses (KJV)", small=False):
    global current_root, leave_timer
    
    def run():
        global current_root, leave_timer

        if current_root and not small:
            current_root.after(0, lambda: safe_close(current_root))

        root = tk.Toplevel(hidden_root)
        current_root = root
        root.title(title)
        root.attributes('-topmost', True)

        if small:
            root.resizable(False, False)
            root.configure(bg=settings["popup"]["bg_small"])

            label = tk.Label(
                root,
                text=text,
                font=tuple(settings["popup"]["font_small"]),
                bg=settings["popup"]["bg_small"],
                fg=settings["popup"]["text_color_small"],
                padx=20,
                pady=15
            )
            label.pack()

            root.update_idletasks()
            width = label.winfo_reqwidth() + 40
            height = label.winfo_reqheight() + 30
            root.geometry(f"{width}x{height}")
        else:
            root.resizable(True, True)
            root.configure(bg=settings["popup"]["bg_large"])
            text_widget = scrolledtext.ScrolledText(
                root,
                width=settings["popup"]["width_large"],
                height=settings["popup"]["height_large"],
                font=tuple(settings["popup"]["font_large"]),
                wrap=tk.WORD,
                bg='white'
            )
            text_widget.pack(padx=12, pady=12, expand=True, fill='both')
            text_widget.insert(tk.END, text)
            text_widget.config(state='disabled')

        root.focus_force()

        root.bind('<Escape>', lambda e: safe_close(root))
        root.protocol("WM_DELETE_WINDOW", lambda: safe_close(root))

        root.bind('<Enter>', lambda e: on_mouse_enter(root))
        root.bind('<Leave>', lambda e: on_mouse_leave(root))

        on_mouse_leave(root)

        root.wait_window()

    threading.Thread(target=run, daemon=True).start()


print(f"FetchKJV READY! Select text → press {format_hotkey(settings['hotkey'])}.")
print("- Stays open while mouse is over the window")
print(f"- Closes {AUTO_CLOSE_SECONDS}s after mouse leaves")
print("- Second press copies clean verses as RTF (with italics) and shows confirmation")

# -----------------------------
# SETTINGS MENU
# -----------------------------

def capture_hotkey(update_callback, display_label):
    popup = tk.Toplevel(hidden_root)
    popup.title("Set Hotkey")
    popup.geometry("300x120")
    popup.attributes("-topmost", True)

    tk.Label(popup, text="Press the desired hotkey...", font=("Segoe UI", 11)).pack(pady=20)

    pressed_mods = {"ctrl": False, "alt": False, "shift": False}

    def on_key(event):
        key = event.keysym

        # Track modifiers
        if key in ("Control_L", "Control_R", Key.ctrl_l, Key.ctrl_r, Key.ctrl):
            pressed_mods["ctrl"] = True
            return
        if key in ("Alt_L", "Alt_R", Key.alt_l, Key.alt_r, Key.alt_gr):
            pressed_mods["alt"] = True
            return
        if key in ("Shift_L", "Shift_R", Key.shift_l, Key.shift_r, Key.shift):
            pressed_mods["shift"] = True
            return
                
        # Build hotkey dict
        new_hotkey = {
            "key": key.lower(),
            "ctrl": pressed_mods["ctrl"],
            "alt": pressed_mods["alt"],
            "shift": pressed_mods["shift"]
        }

        update_callback(new_hotkey)

        # Update variable and label
        display_label.config(text=format_hotkey(new_hotkey))

        popup.destroy()

    popup.bind("<KeyPress>", on_key)
    popup.focus_force()

def open_settings_window():
    # Prevent multiple settings windows
    if any(isinstance(w, tk.Toplevel) and w.title() == "Settings" for w in hidden_root.winfo_children()):
        return

    win = tk.Toplevel(hidden_root)
    win.title("Settings")
    win.geometry("400x400")
    win.resizable(False, False)

    # --- HOTKEY ---
    tk.Label(win, text="Hotkey:").pack(anchor="w", padx=10, pady=(10, 0))

    hotkey_frame = tk.Frame(win)
    hotkey_frame.pack(fill="x", padx=10)

    # Store the dict directly
    hotkey_value = settings["hotkey"]  # a normal Python dict

    def update_hotkey(new_hotkey):
        nonlocal hotkey_value
        hotkey_value = new_hotkey

    # Display the hotkey in a readable format
    hotkey_display = tk.Label(hotkey_frame, text=format_hotkey(settings["hotkey"]))
    hotkey_display.pack(side="left", fill="x", expand=True)

    # Button to capture a new hotkey
    tk.Button(
        hotkey_frame,
        text="Set Hotkey",
        command=lambda: capture_hotkey(lambda new: update_hotkey(new), hotkey_display)
    ).pack(side="right", padx=5)

    # --- AUTO CLOSE ---
    tk.Label(win, text="Auto-close pop up after x seconds:").pack(anchor="w", padx=10, pady=(10, 0))
    auto_var = tk.StringVar(value=str(settings["auto_close_seconds"]))
    tk.Entry(win, textvariable=auto_var).pack(fill="x", padx=10)

    # --- SAVE BUTTON ---
    def save_settings():
        new_settings = {
            "hotkey": hotkey_value,
            "auto_close_seconds": int(auto_var.get()),
            "popup": {
                "bg_small": settings["popup"]["bg_small"],
                "bg_large": settings["popup"]["bg_large"],
                "text_color_small": settings["popup"]["text_color_small"],
                "font_large": settings["popup"]["font_large"],
                "font_small": settings["popup"]["font_small"],
                "width_large": settings["popup"]["width_large"],
                "height_large": settings["popup"]["height_large"]
            }
        }

        with open(SETTINGS_PATH, "w", encoding="utf-8") as f:
            json.dump(new_settings, f, indent=4)

        reload_settings()
        show_popup("Settings saved!", title="FetchKJV", small=True)
        win.destroy()


    tk.Button(win, text="Save", command=save_settings).pack(pady=20)

    # Close behaviour
    win.protocol("WM_DELETE_WINDOW", win.destroy)

# Track modifier state
pressed_modifiers = {
    "ctrl": False,
    "alt": False,
    "shift": False
}

# -----------------------------
# NORMALIZED HOTKEY LISTENER
# -----------------------------

def on_press(key):
    print("KEY EVENT:", key, type(key))
    try:
        hk = settings["hotkey"]

        # ---------------------------------------------------
        # 1. Ignore synthetic Ctrl+C events from your own code
        # ---------------------------------------------------
        if isinstance(key, KeyCode) and key.char == 'c' and pressed_modifiers["ctrl"]:
            return

        # ---------------------------------------------------
        # 2. Update modifier state (robust Alt/AltGr handling)
        # ---------------------------------------------------
        # CTRL
        if key in (Key.ctrl_l, Key.ctrl_r, Key.ctrl):
            pressed_modifiers["ctrl"] = True

        # ALT (Left, Right, AltGr)
        if key in (Key.alt_l, Key.alt_r, Key.alt_gr):
            pressed_modifiers["alt"] = True

        # ALT via KeyCode virtual keycodes
        if isinstance(key, KeyCode) and key.vk in (164, 165):
            pressed_modifiers["alt"] = True

        # SHIFT
        if key in (Key.shift_l, Key.shift_r, Key.shift):
            pressed_modifiers["shift"] = True

        # ---------------------------------------------------
        # 3. Normalize repeated AltGr spam
        # ---------------------------------------------------
        if key == Key.alt_gr and pressed_modifiers["alt"]:
            return

        # ---------------------------------------------------
        # 4. Determine main key pressed
        # ---------------------------------------------------
        pressed = None

        # Case A: Special keys (F-keys, arrows, etc.)
        if isinstance(key, Key):
            pressed = key.name.lower()

        # Case B: KeyCode (needed for some F-keys and OEM keys)
        elif isinstance(key, KeyCode):
            # Some keyboards report F-keys as virtual keycodes
            if key.vk and 112 <= key.vk <= 123:
                pressed = f"f{key.vk - 111}"   # 112→F1, 119→F8, etc.
            elif key.char:
                pressed = key.char.lower()

        if not pressed:
            return

        # ---------------------------------------------------
        # 5. Check modifiers
        # ---------------------------------------------------
        if hk["ctrl"] and not pressed_modifiers["ctrl"]:
            return
        if hk["alt"] and not pressed_modifiers["alt"]:
            return
        if hk["shift"] and not pressed_modifiers["shift"]:
            return

        # ---------------------------------------------------
        # 6. Compare with hotkey
        # ---------------------------------------------------
        if pressed == hk["key"]:
            process_text()

    except Exception as e:
        print("Hotkey error:", e)


def on_release(key):
    # ---------------------------------------------------
    # Delay resetting modifiers to avoid premature clearing
    # ---------------------------------------------------
    def clear():
        # CTRL
        if key in (Key.ctrl_l, Key.ctrl_r, Key.ctrl):
            pressed_modifiers["ctrl"] = False

        # ALT (Left, Right, AltGr)
        if key in (Key.alt_l, Key.alt_r, Key.alt_gr):
            pressed_modifiers["alt"] = False

        # ALT via KeyCode virtual keycodes
        if isinstance(key, KeyCode) and key.vk in (164, 165):
            pressed_modifiers["alt"] = False

        # SHIFT
        if key in (Key.shift_l, Key.shift_r, Key.shift):
            pressed_modifiers["shift"] = False

    # Schedule the clear AFTER pynput finishes processing the keypress
    threading.Timer(0.02, clear).start()

listener = pynput_keyboard.Listener(on_press=on_press, on_release=on_release)
listener.start()
hidden_root.mainloop()