# CalendarWeek

A lightweight, cross-platform system tray application that displays the current ISO week number and provides a clean, Notion-inspired full-year calendar view.

![Platform](https://img.shields.io/badge/platform-Windows%20%7C%20Linux%20%7C%20macOS-blue)
![Python](https://img.shields.io/badge/python-3.8%2B-green)
![License](https://img.shields.io/badge/license-MIT-orange)

---

## ✨ Features

- **System Tray Icon** – Always visible week number in your taskbar/menu bar
- **Full Year Calendar** – Clean, scrollable view with week numbers for every week
- **Today Highlighting** – Current date is highlighted in yellow for quick reference
- **Auto-Start** – Option to launch automatically at system login
- **Cross-Platform** – Works on Windows, Linux, and macOS
- **Lightweight** – Minimal memory footprint with lazy loading
- **Single Instance** – Clicking the icon again brings the existing window to front

---

## 📸 Screenshots

### System Tray Icon
The tray icon displays the current ISO week number (01-53).

### Calendar View
A clean, Notion-inspired calendar showing all 12 months with:
- Week numbers in the left column
- Today's date highlighted
- Auto-scroll to current month on open

---

## 🚀 Installation

### Option 1: Download Pre-built Executable (Windows)

1. Go to the [Releases](https://github.com/MrDNA1708/CalendarWeek/releases) page
2. Download `CalendarWeek.exe`
3. Run it – the icon will appear in your system tray

### Option 2: Run from Source

#### Prerequisites

- Python 3.8 or higher
- pip (Python package manager)

#### Steps

```bash
# Clone the repository
git clone https://github.com/yourusername/CalendarWeek.git
cd CalendarWeek

# Install dependencies
pip install -r requirements.txt

# Run the application

python calendarweek.py


