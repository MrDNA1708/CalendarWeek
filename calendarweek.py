"""
CalendarWeek - A lightweight cross-platform system tray application
that displays the current ISO week number and provides a full year calendar.

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
            # Open the Windows Registry key for startup programs
            key = winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                r"Software\Microsoft\Windows\CurrentVersion\Run",
                0, winreg.KEY_READ
            )
            try:
                # Try to read our app's entry
                winreg.QueryValueEx(key, APP_NAME)
                winreg.CloseKey(key)
                return True
            except FileNotFoundError:
                # Entry doesn't exist
                winreg.CloseKey(key)
                return False
        except Exception:
            return False
    
    elif SYSTEM == "Linux":
        # Check if the .desktop file exists
        return os.path.exists(_get_startup_path_linux())
    
    elif SYSTEM == "Darwin":  # macOS
        # Check if the .plist file exists
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
            # Parse the .desktop file to extract the Exec= line
            with open(desktop_file, 'r') as f:
                for line in f:
                    if line.startswith("Exec="):
                        return line.split("=", 1)[1].strip()
        return None
    
    elif SYSTEM == "Darwin":
        plist_file = _get_startup_path_macos()
        if os.path.exists(plist_file):
            # Parse the .plist file to extract the program path
            with open(plist_file, 'r') as f:
                content = f.read()
                import re
                # Find the first <string>/path/...</string> after ProgramArguments
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
            # Open the Registry key with write permission
            key = winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                r"Software\Microsoft\Windows\CurrentVersion\Run",
                0, winreg.KEY_SET_VALUE
            )
            # Write our app's path as a string value
            winreg.SetValueEx(key, APP_NAME, 0, winreg.REG_SZ, exe_path)
            winreg.CloseKey(key)
            return True
        except Exception:
            return False
    
    elif SYSTEM == "Linux":
        try:
            # Ensure the autostart directory exists
            autostart_dir = os.path.expanduser("~/.config/autostart")
            os.makedirs(autostart_dir, exist_ok=True)
            
            # Create a .desktop file following the freedesktop.org specification
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
    
    elif SYSTEM == "Darwin":  # macOS
        try:
            # Ensure the LaunchAgents directory exists
            launch_agents = os.path.expanduser("~/Library/LaunchAgents")
            os.makedirs(launch_agents, exist_ok=True)
            
            # Create a .plist file following Apple's launchd specification
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
            # Delete our app's Registry entry
            winreg.DeleteValue(key, APP_NAME)
            winreg.CloseKey(key)
            return True
        except Exception:
            return False
    
    elif SYSTEM == "Linux":
        try:
            # Delete the .desktop file
            os.remove(_get_startup_path_linux())
            return True
        except Exception:
            return False
    
    elif SYSTEM == "Darwin":
        try:
            # Delete the .plist file
            os.remove(_get_startup_path_macos())
            return True
        except Exception:
            return False
    
    return False


def check_startup_path():
    """
    Verify that the registered startup path matches the current executable location.
    This detects if the user has moved the application after enabling auto-start.
    
    Returns:
        tuple: (is_valid: bool, message: str or None)
               - is_valid: True if paths match or not registered
               - message: Error message if paths don't match, None otherwise
    """
    if not is_in_startup():
        # Not registered for startup, nothing to check
        return True, None
    
    registered = get_registered_path()
    current = get_exe_path()
    
    # Compare normalized paths to handle different path separators
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
    # Remove the old entry and create a new one with the current path
    remove_from_startup()
    return add_to_startup()


def toggle_startup(icon, item):
    """
    Toggle the automatic startup setting on or off.
    This is the callback function for the "Start at Login" menu item.
    
    Args:
        icon: The pystray Icon object (unused but required by pystray)
        item: The MenuItem object (unused but required by pystray)
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
    
    The ISO week number follows the ISO 8601 standard:
    - Weeks start on Monday
    - Week 1 is the week containing the first Thursday of the year
    
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
    
    Args:
        icon: The pystray Icon object
        interval (int): Refresh interval in seconds (default: 1 hour)
    """
    import time
    
    def refresh_loop():
        last_cw = current_cw()
        while icon.visible:
            time.sleep(interval)
            new_cw = current_cw()
            if new_cw != last_cw:
                icon.icon = create_icon(new_cw)
                icon.title = f"Week {new_cw:02d}"
                last_cw = new_cw
    
    thread = threading.Thread(target=refresh_loop, daemon=True)
    thread.start()


# =============================================================================
# ICON CREATION
# =============================================================================

def create_icon(cw):
    """
    Create the system tray icon displaying the current week number.
    
    The icon is a 64x64 white square with a black border,
    displaying the week number in the center.
    
    Args:
        cw (int): The week number to display (will be zero-padded to 2 digits)
    
    Returns:
        PIL.Image: The generated icon image
    """
    _init_pil()
    size = 64
    
    # Create a white square with transparency support
    img = Image.new("RGBA", (size, size), "white")
    d = ImageDraw.Draw(img)
    
    # Try to load a nice font, with cross-platform fallbacks
    font = None
    font_names = [
        "segoeuib.ttf",     # Segoe UI Bold (Windows)
        "arialbd.ttf",      # Arial Bold (Windows)
        "Arial Bold.ttf",   # Arial Bold (macOS)
        "DejaVuSans-Bold.ttf",  # DejaVu Sans Bold (Linux)
        "Helvetica-Bold.ttc"    # Helvetica Bold (macOS)
    ]
    for font_name in font_names:
        try:
            font = ImageFont.truetype(font_name, 46)
            break
        except Exception:
            continue
    
    # Fall back to PIL's built-in bitmap font if no TrueType fonts available
    if font is None:
        font = ImageFont.load_default()
    
    # Format week number with leading zero (e.g., "05", "12")
    text = f"{cw:02d}"
    
    # Calculate text position to center it in the image
    bbox = d.textbbox((0, 0), text, font=font)
    text_width = bbox[2] - bbox[0]
    text_height = bbox[3] - bbox[1]
    x = (size - text_width) / 2
    y = (size - text_height) / 2 - bbox[1]  # Adjust for font baseline offset
    
    # Draw the week number and border
    d.text((x, y), text, fill="black", font=font)
    d.rectangle([0, 0, size - 1, size - 1], outline="black", width=2)
    
    return img


def create_window_icon():
    """
    Create the icon displayed in the calendar window's title bar.
    
    The icon is a 32x32 stylized calendar image with:
    - A red header bar (typical calendar look)
    - Two hanging "rings" at the top
    - A grid pattern representing days
    
    Returns:
        PIL.Image: The generated window icon image
    """
    _init_pil()
    size = 32
    
    img = Image.new("RGBA", (size, size), "white")
    d = ImageDraw.Draw(img)
    
    # Calendar body (white with dark border)
    d.rectangle([2, 6, size - 3, size - 3], fill="white", outline="#37352f", width=2)
    
    # Red header bar (like a real calendar)
    d.rectangle([2, 6, size - 3, 14], fill="#ff6b6b", outline="#37352f", width=1)
    
    # Calendar "rings" (the hanging hooks at the top)
    d.rectangle([8, 3, 10, 8], fill="#37352f")
    d.rectangle([size - 11, 3, size - 9, 8], fill="#37352f")
    
    # Day grid (simplified 3x2 pattern)
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
    
    # Check if window already exists and bring it to front
    with window_lock:
        if calendar_window is not None:
            try:
                calendar_window.lift()
                calendar_window.focus_force()
                return
            except tk.TclError:
                # Window was destroyed but reference wasn't cleared
                calendar_window = None

    # Get current date information
    today = datetime.date.today()
    year = today.year
    current_month = today.month

    # Create the main window
    root = tk.Tk()
    root.title(f"{year} Calendar")
    root.resizable(False, False)
    root.configure(bg="#ffffff")

    # Set the window icon
    icon_image = create_window_icon()
    icon_photo = ImageTk.PhotoImage(icon_image)
    root.iconphoto(True, icon_photo)
    root._icon_ref = icon_photo  # Keep reference to prevent garbage collection

    # Store window reference for single-instance check
    with window_lock:
        calendar_window = root

    # Color scheme (Notion-inspired design)
    colors = {
        "bg": "#ffffff",           # Background color
        "text": "#37352f",         # Primary text color
        "text_muted": "#9b9a97",   # Secondary/muted text color
        "today_bg": "#ffef3d",     # Today highlight background
        "today_text": "#37352f",   # Today text color
        "week_num": "#9b9a97",     # Week number color
        "border": "#e9e9e7",       # Border/separator color
    }

    # Create scrollable container
    content_frame = tk.Frame(root, bg=colors["bg"])
    content_frame.pack(fill="both", expand=True)

    canvas = tk.Canvas(content_frame, bg=colors["bg"], highlightthickness=0)
    scrollbar = tk.Scrollbar(content_frame, orient="vertical", command=canvas.yview, width=14)
    scroll_frame = tk.Frame(canvas, bg=colors["bg"])

    # Configure scroll region to update when content changes
    scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas_window = canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    # Mouse wheel scrolling (cross-platform)
    def on_mousewheel(event):
        """Handle mouse wheel scrolling with cross-platform support."""
        if SYSTEM == "Linux":
            # Linux uses Button-4 (scroll up) and Button-5 (scroll down)
            if event.num == 4:
                canvas.yview_scroll(-1, "units")
            elif event.num == 5:
                canvas.yview_scroll(1, "units")
        else:
            # Windows and macOS use MouseWheel event with delta
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
    
    # Bind appropriate mouse wheel events based on platform
    if SYSTEM == "Linux":
        canvas.bind_all("<Button-4>", on_mousewheel)
        canvas.bind_all("<Button-5>", on_mousewheel)
    else:
        canvas.bind_all("<MouseWheel>", on_mousewheel)

    scrollbar.pack(side="right", fill="y")
    canvas.pack(side="left", fill="both", expand=True)

    # Font configuration (cross-platform with fallbacks)
    if SYSTEM == "Windows":
        font_family = "Segoe UI"
        font_family_bold = "Segoe UI Semibold"
    elif SYSTEM == "Darwin":  # macOS
        font_family = "Helvetica Neue"
        font_family_bold = "Helvetica Neue Medium"
    else:  # Linux
        font_family = "DejaVu Sans"
        font_family_bold = "DejaVu Sans"

    # Create font objects for consistent styling
    month_font = tkfont.Font(family=font_family_bold, size=11)      # Month names
    day_header_font = tkfont.Font(family=font_family, size=9)       # Day name headers (Mo, Tu, etc.)
    day_font = tkfont.Font(family=font_family, size=10)             # Day numbers
    cw_font = tkfont.Font(family=font_family, size=9)               # Week numbers

    # Container for all months
    calendar_container = tk.Frame(scroll_frame, bg=colors["bg"])
    calendar_container.pack(padx=15, pady=(0, 10))

    month_frames = {}  # Store references for auto-scrolling
    day_names = ["Mo", "Tu", "We", "Th", "Fr", "Sa", "Su"]

    # Generate calendar for each month
    for month in range(1, 13):
        month_frame = tk.Frame(calendar_container, bg=colors["bg"])
        # First month has no top padding, others have spacing
        pady = (0, 5) if month == 1 else (15, 5)
        month_frame.pack(fill="x", pady=pady)
        month_frames[month] = month_frame

        # Month name header
        tk.Label(
            month_frame,
            text=calendar.month_name[month],
            font=month_font,
            bg=colors["bg"],
            fg=colors["text"],
            anchor="w"
        ).pack(fill="x")

        # Separator line under month name
        tk.Frame(month_frame, height=1, bg=colors["border"]).pack(fill="x", pady=(5, 8))

        # Day name headers row (W, Mo, Tu, We, Th, Fr, Sa, Su)
        days_header = tk.Frame(month_frame, bg=colors["bg"])
        days_header.pack(fill="x")
        
        # "W" header for week number column
        tk.Label(
            days_header,
            text="W",
            font=day_header_font,
            bg=colors["bg"],
            fg=colors["text_muted"],
            width=3
        ).pack(side="left")
        
        # Day name headers
        for day_name in day_names:
            tk.Label(
                days_header,
                text=day_name,
                font=day_header_font,
                bg=colors["bg"],
                fg=colors["text_muted"],
                width=3
            ).pack(side="left")

        # Generate weeks for this month
        for week in calendar.monthcalendar(year, month):
            week_frame = tk.Frame(month_frame, bg=colors["bg"])
            week_frame.pack(fill="x", pady=1)

            # Get the week number from the first non-zero day in the week
            first_day = next((d for d in week if d != 0), 1)
            cw = datetime.date(year, month, first_day).isocalendar().week

            # Week number label
            tk.Label(
                week_frame,
                text=f"{cw:02d}",
                font=cw_font,
                bg=colors["bg"],
                fg=colors["week_num"],
                width=3
            ).pack(side="left")

            # Day number labels
            for d in week:
                if d == 0:
                    # Empty cell (day belongs to previous/next month)
                    bg, fg, text = colors["bg"], colors["bg"], ""
                elif d == today.day and month == today.month:
                    # Today's date - highlighted
                    bg, fg, text = colors["today_bg"], colors["today_text"], str(d)
                else:
                    # Regular day
                    bg, fg, text = colors["bg"], colors["text"], str(d)
                
                tk.Label(
                    week_frame,
                    text=text,
                    font=day_font,
                    bg=bg,
                    fg=fg,
                    width=3
                ).pack(side="left")

    # Calculate and set window size and position
    root.update_idletasks()
    calendar_width = scroll_frame.winfo_reqwidth()
    window_width = calendar_width + scrollbar.winfo_reqwidth()
    window_height = 480  # Fixed height with scrolling

    # Position window in the bottom-right area of the screen
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x = screen_width - window_width - 20  # 20px from right edge
    y = (screen_height - window_height) // 2  # Vertically centered
    root.geometry(f"{window_width}x{window_height}+{x}+{y}")

    # Set canvas width to match content
    canvas.itemconfig(canvas_window, width=calendar_width)

    def scroll_to_current_month():
        """Auto-scroll to center the current month in the viewport."""
        root.update_idletasks()
        mf = month_frames[current_month]
        
        # Calculate scroll position to center the current month
        target_y = mf.winfo_y() + (mf.winfo_height() / 2) - (canvas.winfo_height() / 2)
        scroll_height = scroll_frame.winfo_reqheight()
        
        if scroll_height > canvas.winfo_height():
            # Clamp scroll position between 0 and 1
            scroll_pos = max(0, min(target_y / (scroll_height - canvas.winfo_height()), 1))
            canvas.yview_moveto(scroll_pos)

    # Scroll to current month after window is fully rendered
    root.after(50, scroll_to_current_month)

    def on_close(event=None):
        """
        Handle window close event.
        Cleans up event bindings and triggers garbage collection.
        """
        global calendar_window
        
        # Unbind mouse wheel events to prevent errors after window destruction
        if SYSTEM == "Linux":
            canvas.unbind_all("<Button-4>")
            canvas.unbind_all("<Button-5>")
        else:
            canvas.unbind_all("<MouseWheel>")
        
        # Clear the window reference
        with window_lock:
            calendar_window = None
        
        root.destroy()
        gc.collect()  # Force garbage collection to free memory

    # Bind close events
    root.bind("<Escape>", on_close)  # Close with Escape key
    root.protocol("WM_DELETE_WINDOW", on_close)  # Close with X button
    
    
    root.lift()                          # Porta sopra le altre finestre
    root.attributes('-topmost', True)    # Forza in primo piano
    root.after(100, lambda: root.attributes('-topmost', False))  # Rimuove topmost dopo 100ms
    root.focus_force()                   # Forza il focus
    
    # Su Windows, un trucco aggiuntivo per garantire il focus
    if sys.platform == 'win32':
        root.after(200, lambda: root.focus_force())

    # Start the window event loop
    root.mainloop()


# =============================================================================
# TRAY MENU ACTIONS
# =============================================================================

def open_calendar(icon, item):
    """
    Open the calendar window in a separate thread.
    
    Running in a separate thread prevents blocking the tray icon's event loop.
    
    Args:
        icon: The pystray Icon object (unused but required by pystray)
        item: The MenuItem object (unused but required by pystray)
    """
    threading.Thread(target=show_calendar, daemon=True).start()


def show_path_warning():
    """
    Check if the executable has been moved and show a warning dialog if so.
    
    This function is called at startup to detect when the user has moved
    the application after enabling auto-start. If the paths don't match,
    it offers to update the registered path.
    """
    _init_tk()
    from tkinter import messagebox
    
    is_valid, message = check_startup_path()
    
    if not is_valid:
        # Create a hidden root window for the dialog
        root = tk.Tk()
        root.withdraw()
        
        # Ask user if they want to fix the path
        result = messagebox.askyesno(
            "Path Changed",
            f"{message}\n\nWould you like to update the automatic startup path?"
        )
        
        if result:
            # User chose to update the path
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
    Tenta di acquisire un lock per garantire che solo un'istanza sia in esecuzione.
    
    Usa un socket su una porta locale specifica. Se la porta è già in uso,
    significa che un'altra istanza è in esecuzione.
    
    Returns:
        bool: True se il lock è stato acquisito (prima istanza),
              False se un'altra istanza è già in esecuzione
    """
    global LOCK_SOCKET
    
    LOCK_SOCKET = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    
    try:
        # Tenta di fare il bind sulla porta locale
        # Se fallisce, un'altra istanza è già in esecuzione
        LOCK_SOCKET.bind(('127.0.0.1', LOCK_PORT))
        LOCK_SOCKET.listen(1)
        return True
    except socket.error:
        # Porta già in uso = altra istanza attiva
        LOCK_SOCKET.close()
        LOCK_SOCKET = None
        return False
def release_single_instance_lock():
    """
    Rilascia il lock dell'istanza singola.
    Chiamata automaticamente quando l'applicazione si chiude.
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
    Notifica l'istanza esistente che l'utente ha tentato di aprire una nuova istanza.
    Opzionale: può essere usato per portare la finestra del calendario in primo piano.
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
    
    # ===== CHECK SINGLE ISTANCE =====
    if not acquire_single_instance_lock():
        # Another istance is already running, notify it and exit
        print("CalendarWeek is already running.")
        notify_existing_instance()
        sys.exit(0)
    
    # Register the cleanup function and relase the lock when the application exits
    import atexit
    atexit.register(release_single_instance_lock)

    # Check for path mismatch at startup (runs in background)
    threading.Thread(target=show_path_warning, daemon=True).start()
    
    # Get current week number for the tray icon
    cw = current_cw()
    
    # Build the context menu
    menu = Menu(
        # Open Calendar - default action (triggered by left-click)
        MenuItem("Open Calendar", open_calendar, default=True),
        
        Menu.SEPARATOR,
        
        # Start at Login toggle with checkbox
        # Label changes based on OS ("Start with Windows" vs "Start at Login")
        MenuItem(
            "Start with Windows" if SYSTEM == "Windows" else "Start at Login",
            toggle_startup,
            checked=lambda item: is_in_startup()  # Checkbox state callback
        ),
        
        Menu.SEPARATOR,
        
        # Exit the application
        MenuItem("Exit", lambda icon, item: icon.stop())
    )
    
    # Create and run the system tray icon
    icon = Icon(
        name="CW",                      # Internal name
        icon=create_icon(cw),           # The icon image
        title=f"Week {cw:02d}",         # Tooltip text
        menu=menu                       # Context menu
    )

    # ===== LISTENER FOR NEW INSTANCES =====
    def listen_for_new_instances():
        """
        Listen to the new requests from other instances and open the calendar.
        """
        while True:
            try:
                if LOCK_SOCKET is None:
                    break
                conn, addr = LOCK_SOCKET.accept()
                data = conn.recv(1024)
                conn.close()
                if data == b'SHOW':
                    # Open the calendar window when a new instance tries to start
                    threading.Thread(target=show_calendar, daemon=True).start()
            except:
                break
    
    # Start the listener in background
    threading.Thread(target=listen_for_new_instances, daemon=True).start()

    # Auto refresh each hour
    start_auto_refresh(icon, interval=3600)
    
    # This blocks and runs the tray icon event loop
    icon.run()


# =============================================================================
# SCRIPT EXECUTION
# =============================================================================

if __name__ == "__main__":

    run()
