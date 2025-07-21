# Windows ChatGPT Rephraser Overlay

A Windows-only Python application that provides a Grammarly-like overlay for rephrasing selected text using ChatGPT.

## Features
- System tray icon for easy access
- Settings window (with API Key, API URL, and Prompt fields)
- Global hotkey to trigger overlay (default: Ctrl+Shift+G)
- Detects selected text in any application
- Overlay UI near cursor with rephrase button
- Uses OpenAI's ChatGPT to rephrase text
- Copy rephrased text to clipboard

## Setup
1. Install Python 3.8+
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
3. Run the app:
   ```bash
   python main.py
   ```

## Usage
- The app runs in the system tray (bottom right of your Windows taskbar).
- Right-click the tray icon to access the menu:
  - **Settings**: Opens a window with two tabs:
    - **General**: (currently empty)
    - **Parameters**: Set your OpenAI API Key, API URL, and the prompt used for rephrasing. These are saved to `settings.json` and used for all requests.
  - **Exit**: Closes the app.
- Select text anywhere in Windows (minimum 100 characters).
- Click the floating button or use the hotkey to rephrase.
- The overlay appears with the rephrased text, which you can click to copy.

## Settings
- Settings are stored in `settings.json` in the app directory.
- You can change the API key, API URL, and prompt at any time via the Settings window.
- Changes take effect immediately after saving.

## HTTP Debugging
If you want to see the full URL and details of API requests (for troubleshooting), HTTP debugging is enabled by default. You will see detailed request logs in your console output.

## Known Limitations
- **Taskbar Icon**: Due to Windows and PyQt5 limitations, the settings window may not always show your custom icon in the taskbar, even though the tray icon and window icon are set. This is a known issue for tray-only apps.
- The "General" tab in settings is currently a placeholder for future options.

## Troubleshooting
- If you get a 404 or authentication error, double-check your API URL and API key in the settings window.
- If you change settings and they do not take effect, restart the app.

## License
MIT 