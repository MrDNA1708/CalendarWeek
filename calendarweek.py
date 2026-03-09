""" CalendarWeek - A lightweight cross-platform system tray application that displays the current ISO week number and provides a full year calendar.

Supported platforms: Windows, Linux, macOS
Dependencies: pillow, pystray
"""

import datetime
import calendar
import threading
import gc
import sys
import os
import platform

# =============================================================================
# LAZY IMPORTS
# =============================================================================
# These modules are loaded only when needed to reduce initial memory footprint.
# This optimization saves ~5-10MB of RAM when the calendar window is not open.

tk = None          # tkinter module (GUI framework)
Image = None       # PIL.Image for image creation
ImageDraw = None   # PIL.ImageDraw for drawing on images
ImageFont = None   # PIL.ImageFont for text rendering
ImageTk = None     # PIL.ImageTk for tkinter-compatible images

def _init_pil():
    """
    Lazily initialize PIL (Pillow) modules.
    Called only when image creation is needed.
    """
    global Image, ImageDraw, ImageFont
    if Image is None:
        from PIL import Image as _Image, ImageDraw as _ImageDraw, ImageFont as _ImageFont
        Image, ImageDraw, ImageFont = _Image, _ImageDraw, _ImageFont

def _init_tk():
    """
    Lazily initialize tkinter and PIL.ImageTk modules.
    Called only when the calendar window needs to be displayed.
    """
    global tk, ImageTk
    if tk is None:
        import tkinter as _tk
        from PIL import ImageTk as _ImageTk
        tk = _tk
        ImageTk = _ImageTk

# =============================================================================
# GLOBAL STATE
# =============================================================================

import socket
import tempfile

LOCK_SOCKET = None
LOCK_PORT = 47200

calendar_window = None      # Reference to the calendar window (None if closed)
window_lock = threading.Lock()  # Thread lock to prevent race conditions
APP_NAME = "CalendarWeek"   # Application name used for startup registration
SYSTEM = platform.system()  # Operating system: "Windows", "Linux", or "Darwin" (macOS)

# =============================================================================
# CROSS-PLATFORM STARTUP UTILITIES
# =============================================================================
# These functions handle automatic startup registration for all supported OS.
# - Windows: Uses the Registry (HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run)
# - Linux: Creates a .desktop file in ~/.config/autostart/
# - macOS: Creates a .plist file in ~/Library/LaunchAgents/

def get_exe_path():
    """
    Get the absolute path of the current executable or script.

    Returns:
        str: Full path to the executable (if frozen with PyInstaller)
             or the script file (if running as .py)
    """
    if getattr(sys, 'frozen', False):
        # Running as compiled executable (PyInstaller)
        return sys.executable
    # Running as Python script
    return os.path.abspath(__file__)

def _get_startup_path_linux():
    """
    Get the path for the Linux autostart .desktop file.

    Returns:
        str: Full path to ~/.config/autostart/CalendarWeek.desktop
    """
    autostart_dir = os.path.expanduser("~/.config/autostart")
    return os.path.join(autostart_dir, f"{APP_NAME}.desktop")

def _get_startup_path_macos():
    """
    Get the path for the macOS LaunchAgent .plist file.

    Returns:
        str: Full path to ~/Library/LaunchAgents/com.CalendarWeek.plist
    """
    launch_agents = os.path.expanduser("~/Library/LaunchAgents")
    return os.path.join(launch_agents, f"com.{APP_NAME}.plist")

def is_in_startup():
    """
    Check if the application is registered to start automatically at login.

    Returns:
        bool: True if registered for automatic startup, False otherwise
    """
    if SYSTEM == "Windows":
        try:
            import winreg
            key = winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                r"Software\Microsoft\Windows\CurrentVersion\Run",
                0, winreg.KEY_READ
            )
            try:
                winreg.QueryValueEx(key, APP_NAME)
                winreg.CloseKey(key)
                return True
            except FileNotFoundError:
                winreg.CloseKey(key)
                return False
        except Exception:
            return False
    elif SYSTEM == "Linux":
        return os.path.exists(_get_startup_path_linux())
    elif SYSTEM == "Darwin":
        return os.path.exists(_get_startup_path_macos())
    return False

def get_registered_path():
    """
    Get the executable path that is currently registered for automatic startup.

    Returns:
        str or None: The registered path, or None if not registered
    """
    if SYSTEM == "Windows":
        try:
            import winreg
            key = winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                r"Software\Microsoft\Windows\CurrentVersion\Run",
                0, winreg.KEY_READ
            )
            try:
                value, _ = winreg.QueryValueEx(key, APP_NAME)
                winreg.CloseKey(key)
                return value
            except FileNotFoundError:
                winreg.CloseKey(key)
                return None
        except Exception:
            return None
    elif SYSTEM == "Linux":
        desktop_file = _get_startup_path_linux()
        if os.path.exists(desktop_file):
            with open(desktop_file, 'r') as f:
                for line in f:
                    if line.startswith("Exec="):
                        return line.split("=", 1)[1].strip()
        return None
    elif SYSTEM == "Darwin":
        plist_file = _get_startup_path_macos()
        if os.path.exists(plist_file):
            with open(plist_file, 'r') as f:
                content = f.read()
            import re
            match = re.search(r'<string>(/[^<]+)</string>', content)
            if match:
                return match.group(1)
        return None
    return None

def add_to_startup():
    """
    Register the application to start automatically at system login.

    Returns:
        bool: True if registration was successful, False otherwise
    """
    exe_path = get_exe_path()

    if SYSTEM == "Windows":
        try:
            import winreg
            key = winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                r"Software\Microsoft\Windows\CurrentVersion\Run",
                0, winreg.KEY_SET_VALUE
            )
            winreg.SetValueEx(key, APP_NAME, 0, winreg.REG_SZ, exe_path)
            winreg.CloseKey(key)
            return True
        except Exception:
            return False
    elif SYSTEM == "Linux":
        try:
            autostart_dir = os.path.expanduser("~/.config/autostart")
            os.makedirs(autostart_dir, exist_ok=True)
            desktop_content = f"""[Desktop Entry]
Type=Application
Name={APP_NAME}
Exec={exe_path}
Hidden=false
NoDisplay=false
X-GNOME-Autostart-enabled=true
"""
            with open(_get_startup_path_linux(), 'w') as f:
                f.write(desktop_content)
            return True
        except Exception:
            return False
    elif SYSTEM == "Darwin":
        try:
            launch_agents = os.path.expanduser("~/Library/LaunchAgents")
            os.makedirs(launch_agents, exist_ok=True)
            plist_content = f"""<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>Label</key>
    <string>com.{APP_NAME}</string>
    <key>ProgramArguments</key>
    <array>
        <string>{exe_path}</string>
    </array>
    <key>RunAtLoad</key>
    <true/>
</dict>
</plist>
"""
            with open(_get_startup_path_macos(), 'w') as f:
                f.write(plist_content)
            return True
        except Exception:
            return False
    return False

def remove_from_startup():
    """
    Remove the application from automatic startup at system login.

    Returns:
        bool: True if removal was successful, False otherwise
    """
    if SYSTEM == "Windows":
        try:
            import winreg
            key = winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                r"Software\Microsoft\Windows\CurrentVersion\Run",
                0, winreg.KEY_SET_VALUE
            )
            winreg.DeleteValue(key, APP_NAME)
            winreg.CloseKey(key)
            return True
        except Exception:
            return False
    elif SYSTEM == "Linux":
        try:
            os.remove(_get_startup_path_linux())
            return True
        except Exception:
            return False
    elif SYSTEM == "Darwin":
        try:
            os.remove(_get_startup_path_macos())
            return True
        except Exception:
            return False
    return False

def check_startup_path():
    """
    Verify that the registered startup path matches the current executable location.

    Returns:
        tuple: (is_valid: bool, message: str or None)
    """
    if not is_in_startup():
        return True, None

    registered = get_registered_path()
    current = get_exe_path()

    if registered and os.path.normpath(registered) != os.path.normpath(current):
        return False, (
            f"The application has been moved.\n\n"
            f"Registered path:\n{registered}\n\n"
            f"Current path:\n{current}"
        )
    return True, None

def fix_startup_path():
    """
    Update the registered startup path to the current executable location.

    Returns:
        bool: True if the path was successfully updated, False otherwise
    """
    remove_from_startup()
    return add_to_startup()

def toggle_startup(icon, item):
    """
    Toggle the automatic startup setting on or off.
    """
    if is_in_startup():
        remove_from_startup()
    else:
        add_to_startup()

# =============================================================================
# CALENDAR WEEK UTILITIES
# =============================================================================

def current_cw():
    """
    Get the current ISO calendar week number.

    Returns:
        int: Current week number (1-53)
    """
    return datetime.date.today().isocalendar().week

# =============================================================================
# PERIODIC AUTO REFRESH
# =============================================================================

def start_auto_refresh(icon, interval=3600):
    """
    Periodically refresh the tray icon to update the week number.

    Instead of a single long sleep, the loop sleeps in short 30-second steps.
    This makes the refresh robust against system sleep/wake cycles: even if the
    OS freezes or extends a long sleep during standby, the elapsed counter will
    reach the target interval within 30 seconds of the system waking back up.
    Full try/except handling ensures the thread never dies silently.

    Args:
        icon: The pystray Icon object
        interval (int): Refresh interval in seconds (default: 1 hour)
    """
    import time

    def refresh_loop():
        last_cw = current_cw()
        elapsed = 0
        check_step = 30  # Sleep granularity in seconds; robust against OS sleep

        while True:
            try:
                if not icon.visible:
                    break
                time.sleep(check_step)
                elapsed += check_step

                if elapsed >= interval:
                    elapsed = 0
                    new_cw = current_cw()
                    if new_cw != last_cw:
                        icon.icon = create_icon(new_cw)
                        icon.title = f"Week {new_cw:02d}"
                        last_cw = new_cw
            except Exception as e:
                # Thread never dies silently: log the error and keep looping
                print(f"[AutoRefresh] Error: {e}")
                try:
                    time.sleep(check_step)
                except Exception:
                    pass  # Even the fallback sleep failed; just continue looping

    thread = threading.Thread(target=refresh_loop, daemon=True)
    thread.start()

# =============================================================================
# ICON CREATION
# =============================================================================

def create_icon(cw):
    """
    Create the system tray icon displaying the current week number.

    Args:
        cw (int): The week number to display (will be zero-padded to 2 digits)

    Returns:
        PIL.Image: The generated icon image
    """
    _init_pil()
    size = 64
    img = Image.new("RGBA", (size, size), "white")
    d = ImageDraw.Draw(img)
    font = None
    font_names = [
        "segoeuib.ttf",
        "arialbd.ttf",
        "Arial Bold.ttf",
        "DejaVuSans-Bold.ttf",
        "Helvetica-Bold.ttc"
    ]
    for font_name in font_names:
        try:
            font = ImageFont.truetype(font_name, 46)
            break
        except Exception:
            continue
    if font is None:
        font = ImageFont.load_default()
    text = f"{cw:02d}"
    bbox = d.textbbox((0, 0), text, font=font)
    text_width = bbox[2] - bbox[0]
    text_height = bbox[3] - bbox[1]
    x = (size - text_width) / 2
    y = (size - text_height) / 2 - bbox[1]
    d.text((x, y), text, fill="black", font=font)
    d.rectangle([0, 0, size - 1, size - 1], outline="black", width=2)
    return img

def create_window_icon():
    """
    Create the icon displayed in the calendar window's title bar.

    Returns:
        PIL.Image: The generated window icon image
    """
    _init_pil()
    size = 32
    img = Image.new("RGBA", (size, size), "white")
    d = ImageDraw.Draw(img)
    d.rectangle([2, 6, size - 3, size - 3], fill="white", outline="#37352f", width=2)
    d.rectangle([2, 6, size - 3, 14], fill="#ff6b6b", outline="#37352f", width=1)
    d.rectangle([8, 3, 10, 8], fill="#37352f")
    d.rectangle([size - 11, 3, size - 9, 8], fill="#37352f")
    for row in range(2):
        for col in range(3):
            x = 8 + col * 7
            y = 18 + row * 6
            d.rectangle([x, y, x + 4, y + 3], fill="#37352f")
    return img

# =============================================================================
# CALENDAR WINDOW
# =============================================================================

def show_calendar():
    """
    Display the full-year calendar window.

    Features:
    - Shows all 12 months with week numbers
    - Highlights today's date in yellow
    - Scrollable view that auto-scrolls to the current month
    - Single instance (clicking again brings existing window to front)
    - Closes with Escape key or window close button
    """
    _init_tk()
    from tkinter import font as tkfont
    from tkinter import messagebox

    global calendar_window

    with window_lock:
        if calendar_window is not None:
            try:
                calendar_window.lift()
                calendar_window.focus_force()
                return
            except tk.TclError:
                calendar_window = None

    today = datetime.date.today()
    year = today.year
    current_month = today.month

    root = tk.Tk()
    root.title(f"{year} Calendar")
    root.resizable(False, False)
    root.configure(bg="#ffffff")

    icon_image = create_window_icon()
    icon_photo = ImageTk.PhotoImage(icon_image)
    root.iconphoto(True, icon_photo)
    root._icon_ref = icon_photo

    with window_lock:
        calendar_window = root

    colors = {
        "bg": "#ffffff",
        "text": "#37352f",
        "text_muted": "#9b9a97",
        "today_bg": "#ffef3d",
        "today_text": "#37352f",
        "week_num": "#9b9a97",
        "border": "#e9e9e7",
    }

    content_frame = tk.Frame(root, bg=colors["bg"])
    content_frame.pack(fill="both", expand=True)

    canvas = tk.Canvas(content_frame, bg=colors["bg"], highlightthickness=0)
    scrollbar = tk.Scrollbar(content_frame, orient="vertical", command=canvas.yview, width=14)
    scroll_frame = tk.Frame(canvas, bg=colors["bg"])

    scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas_window = canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    def on_mousewheel(event):
        if SYSTEM == "Linux":
            if event.num == 4:
                canvas.yview_scroll(-1, "units")
            elif event.num == 5:
                canvas.yview_scroll(1, "units")
        else:
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    if SYSTEM == "Linux":
        canvas.bind_all("<Button-4>", on_mousewheel)
        canvas.bind_all("<Button-5>", on_mousewheel)
    else:
        canvas.bind_all("<MouseWheel>", on_mousewheel)

    scrollbar.pack(side="right", fill="y")
    canvas.pack(side="left", fill="both", expand=True)

    if SYSTEM == "Windows":
        font_family = "Segoe UI"
        font_family_bold = "Segoe UI Semibold"
    elif SYSTEM == "Darwin":
        font_family = "Helvetica Neue"
        font_family_bold = "Helvetica Neue Medium"
    else:
        font_family = "DejaVu Sans"
        font_family_bold = "DejaVu Sans"

    month_font = tkfont.Font(family=font_family_bold, size=11)
    day_header_font = tkfont.Font(family=font_family, size=9)
    day_font = tkfont.Font(family=font_family, size=10)
    cw_font = tkfont.Font(family=font_family, size=9)

    calendar_container = tk.Frame(scroll_frame, bg=colors["bg"])
    calendar_container.pack(padx=15, pady=(0, 10))

    month_frames = {}
    day_names = ["Mo", "Tu", "We", "Th", "Fr", "Sa", "Su"]

    for month in range(1, 13):
        month_frame = tk.Frame(calendar_container, bg=colors["bg"])
        pady = (0, 5) if month == 1 else (15, 5)
        month_frame.pack(fill="x", pady=pady)
        month_frames[month] = month_frame

        tk.Label(
            month_frame, text=calendar.month_name[month],
            font=month_font, bg=colors["bg"], fg=colors["text"], anchor="w"
        ).pack(fill="x")

        tk.Frame(month_frame, height=1, bg=colors["border"]).pack(fill="x", pady=(5, 8))

        days_header = tk.Frame(month_frame, bg=colors["bg"])
        days_header.pack(fill="x")

        tk.Label(days_header, text="W", font=day_header_font,
                 bg=colors["bg"], fg=colors["text_muted"], width=3).pack(side="left")

        for day_name in day_names:
            tk.Label(days_header, text=day_name, font=day_header_font,
                     bg=colors["bg"], fg=colors["text_muted"], width=3).pack(side="left")

        for week in calendar.monthcalendar(year, month):
            week_frame = tk.Frame(month_frame, bg=colors["bg"])
            week_frame.pack(fill="x", pady=1)

            first_day = next((d for d in week if d != 0), 1)
            cw = datetime.date(year, month, first_day).isocalendar().week

            tk.Label(week_frame, text=f"{cw:02d}", font=cw_font,
                     bg=colors["bg"], fg=colors["week_num"], width=3).pack(side="left")

            for d in week:
                if d == 0:
                    bg, fg, text = colors["bg"], colors["bg"], ""
                elif d == today.day and month == today.month:
                    bg, fg, text = colors["today_bg"], colors["today_text"], str(d)
                else:
                    bg, fg, text = colors["bg"], colors["text"], str(d)

                tk.Label(week_frame, text=text, font=day_font,
                         bg=bg, fg=fg, width=3).pack(side="left")

    root.update_idletasks()
    calendar_width = scroll_frame.winfo_reqwidth()
    window_width = calendar_width + scrollbar.winfo_reqwidth()
    window_height = 480

    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x = screen_width - window_width - 20
    y = (screen_height - window_height) // 2
    root.geometry(f"{window_width}x{window_height}+{x}+{y}")

    canvas.itemconfig(canvas_window, width=calendar_width)

    def scroll_to_current_month():
        root.update_idletasks()
        mf = month_frames[current_month]
        target_y = mf.winfo_y() + (mf.winfo_height() / 2) - (canvas.winfo_height() / 2)
        scroll_height = scroll_frame.winfo_reqheight()
        if scroll_height > canvas.winfo_height():
            scroll_pos = max(0, min(target_y / (scroll_height - canvas.winfo_height()), 1))
            canvas.yview_moveto(scroll_pos)

    root.after(50, scroll_to_current_month)

    def on_close(event=None):
        global calendar_window
        if SYSTEM == "Linux":
            canvas.unbind_all("<Button-4>")
            canvas.unbind_all("<Button-5>")
        else:
            canvas.unbind_all("<MouseWheel>")
        with window_lock:
            calendar_window = None
        root.destroy()
        gc.collect()

    root.bind("<Escape>", on_close)
    root.protocol("WM_DELETE_WINDOW", on_close)
    root.lift()
    root.attributes('-topmost', True)
    root.after(100, lambda: root.attributes('-topmost', False))
    root.focus_force()
    if sys.platform == 'win32':
        root.after(200, lambda: root.focus_force())

    root.mainloop()

# =============================================================================
# TRAY MENU ACTIONS
# =============================================================================

def open_calendar(icon, item):
    threading.Thread(target=show_calendar, daemon=True).start()

def show_path_warning():
    _init_tk()
    from tkinter import messagebox

    is_valid, message = check_startup_path()
    if not is_valid:
        root = tk.Tk()
        root.withdraw()
        result = messagebox.askyesno(
            "Path Changed",
            f"{message}\n\nWould you like to update the automatic startup path?"
        )
        if result:
            if fix_startup_path():
                messagebox.showinfo("Success", "Startup path updated successfully!")
            else:
                messagebox.showerror("Error", "Failed to update the startup path.")
        root.destroy()

# =============================================================================
# SINGLE INSTANCE MECHANISM
# =============================================================================

def acquire_single_instance_lock():
    """
    Attempt to acquire a lock to ensure only one instance is running.

    Uses a socket bound to a local port. If the port is already in use,
    another instance is running.

    Returns:
        bool: True if the lock was acquired (first instance), False otherwise
    """
    global LOCK_SOCKET
    LOCK_SOCKET = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        LOCK_SOCKET.bind(('127.0.0.1', LOCK_PORT))
        LOCK_SOCKET.listen(1)
        return True
    except socket.error:
        LOCK_SOCKET.close()
        LOCK_SOCKET = None
        return False

def release_single_instance_lock():
    """
    Release the single-instance lock. Called automatically on application exit.
    """
    global LOCK_SOCKET
    if LOCK_SOCKET is not None:
        try:
            LOCK_SOCKET.close()
        except:
            pass
        LOCK_SOCKET = None

def notify_existing_instance():
    """
    Notify the existing instance that the user tried to open a new one.
    Signals it to bring the calendar window to the front.
    """
    try:
        client = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        client.connect(('127.0.0.1', LOCK_PORT))
        client.send(b'SHOW')
        client.close()
    except:
        pass

# =============================================================================
# MAIN ENTRY POINT
# =============================================================================

def run():
    """
    Main application entry point.

    Initializes the system tray icon with:
    - Current week number display
    - Menu with: Open Calendar, Start at Login toggle, Exit
    - Left-click to open calendar (default action)
    """
    from pystray import Icon, Menu, MenuItem

    if not acquire_single_instance_lock():
        print("CalendarWeek is already running.")
        notify_existing_instance()
        sys.exit(0)

    import atexit
    atexit.register(release_single_instance_lock)

    threading.Thread(target=show_path_warning, daemon=True).start()

    cw = current_cw()

    menu = Menu(
        MenuItem("Open Calendar", open_calendar, default=True),
        Menu.SEPARATOR,
        MenuItem(
            "Start with Windows" if SYSTEM == "Windows" else "Start at Login",
            toggle_startup,
            checked=lambda item: is_in_startup()
        ),
        Menu.SEPARATOR,
        MenuItem("Exit", lambda icon, item: icon.stop())
    )

    icon = Icon(
        name="CW",
        icon=create_icon(cw),
        title=f"Week {cw:02d}",
        menu=menu
    )

    def listen_for_new_instances():
        """
        Listen for requests from other instances and open the calendar window.
        """
        while True:
            try:
                if LOCK_SOCKET is None:
                    break
                conn, addr = LOCK_SOCKET.accept()
                data = conn.recv(1024)
                conn.close()
                if data == b'SHOW':
                    threading.Thread(target=show_calendar, daemon=True).start()
            except:
                break

    threading.Thread(target=listen_for_new_instances, daemon=True).start()

    start_auto_refresh(icon, interval=3600)

    icon.run()

# =============================================================================
# SCRIPT EXECUTION
# =============================================================================

if __name__ == "__main__":
    run()
