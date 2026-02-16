# YG Crystal Bot

> **v1.2.3** — Automated crystal management for YG game farming windows

A Windows desktop automation bot that monitors game windows and automatically manages crystal items using computer vision (OpenCV template matching) and input automation (PyAutoGUI).

## Features

- **Auto crystal transfer** — Drags crystals from the main inventory to the target slot when the bag has an empty slot
- **Popup handling** — Automatically detects and closes in-game popups (cancel, close buttons, messages)
- **Bag management** — Opens the bag when needed and detects full/empty states
- **Multi-window support** — Monitors all game windows with "FARM" in the title
- **Status tracking** — Tracks windows as `xONLINE`, `xFULL` (bag full), and handles disconnection
- **LINE notifications** — Sends alerts when the bag is full or crystals are depleted

## Requirements

- Windows (uses `win32gui`, screen capture)
- Python 3.x
- Game windows titled with "FARM" (and "xONLINE" suffix when active)

### Dependencies

```
opencv-python
numpy
Pillow
pygetwindow
pyautogui
requests
pywin32
```

## Installation

1. Clone the repository:
   ```bash
   git clone <repository-url>
   cd yg-crystal-bot
   ```

2. Create a virtual environment and install dependencies:
   ```bash
   python -m venv venv
   venv\Scripts\activate
   pip install opencv-python numpy Pillow pygetwindow pyautogui requests pywin32
   ```

3. Set up the `images/` folder with the required template images (see [Images](#images) below).

4. (Optional) Configure LINE notifications in `noti.py` — replace the token with your own LINE Notify token.

## Usage

Run the bot:

```bash
python crystal.py
```

The bot will:

1. Find all windows with "FARM" in the title
2. Cycle through windows marked `xONLINE`
3. Capture the screen, detect popups, open the bag if needed
4. If there is an empty slot and no crystal added yet, drag crystals from the main area to the target slot
5. Send LINE notifications on bag full / out of crystal

## Images

Place template images in `images/`:

| File | Purpose |
|------|---------|
| `main_template_image.png` | Source area for crystal drag |
| `target_template_image.png` | Target slot for crystal |
| `bag.png` | Bag open state |
| `bag_icon.png` | Bag icon to open bag |
| `low_crystal.png`, `high_crystal.png` | Crystal presence check |
| `empty.png` | Empty slot indicator |
| `cancel.png`, `close.png`, `popup_message.png` | Popup elements to close |
| `storage_panel.png`, `shop_panel.png`, `npc_panel.png` | Panels to skip (character busy) |
| `warning.png` | Connection warning |

## Building executable

Build a standalone `.exe` with PyInstaller:

```bash
pip install pyinstaller
pyinstaller crystal.spec
```

Output: `dist/crystal.exe`

## Project structure

```
yg-crystal-bot/
├── crystal.py      # Main bot logic
├── noti.py         # LINE notification helper
├── crystal.spec   # PyInstaller spec
├── images/        # Template images (add your own)
└── README.md
```

## Notes

- Ensure the game windows are visible and not minimized.
- Adjust screen coordinates in the code if your resolution differs from the default (e.g. `1270, 540`, `930, 750`).
- For LINE notifications, create a token at [LINE Notify](https://notify-bot.line.me/).
