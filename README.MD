# WRobot Auto Trial

An automated trial manager for WRobot (World of Warcraft bot) that handles trial session management and provides seamless bot operation.

**Supported WoW Versions:**
- Vanilla Classic (1.12.1)
- The Burning Crusade TBC (2.4.3)
- Wrath of the Lich King WOTLK (3.3.5a)
- Cataclysm (4.3.4)
- Mists of Pandaria MOP (5.4.8/5.4.7)
- Warlords of Draenor WOD (6.2.4/6.2.3)
- Legion (7.3.5)
- Battle for Azeroth BFA (8.3.7/8.2.0)
- Shadowlands SL (9.2.7)
- Dragonflight DF (10.2.7/10.1.7)

## Features

- **Automatic Process Monitoring**: Continuously monitors and restarts WRobot.exe if it's not running
- **Trial Session Management**: Automatically handles the 15-minute trial session restarts
- **Captcha Bypass**: Automatically closes captcha windows that appear during trial sessions
- **Bot Auto-Start**: Sends Alt+C hotkey to automatically start the bot after each restart
- **Error Handling**: Robust error handling with automatic retry mechanisms
- **Threading**: Uses separate threads for non-blocking operations

## Requirements

- Python 3.x
- Windows OS
- WRobot.exe (place in the same directory as the script)

### Python Dependencies

```bash
pip install psutil pywin32 pywinauto --upgrade
```

## Installation

1. Download `wrobot_auto_trial.py` from this repository
2. Place the script in the same directory as your `WRobot.exe`
3. Install dependencies:
```bash
pip install psutil pywin32 pywinauto --upgrade
```

## Usage

Simply run the script:

```bash
python wrobot_auto_trial.py
```

The script will:
1. Check if WRobot is running
2. Start WRobot if not found
3. Wait for server connection (up to 2 minutes)
4. Click the "Launch Bot" button
5. Handle the license window and captcha
6. Send Alt+C to start the bot
7. Repeat the process every 5 seconds

## How It Works

### Trial Session Management
WRobot trial version automatically closes every 15 minutes and shows a fake captcha window. This launcher:
- Detects when WRobot closes
- Automatically restarts the application
- Closes the captcha window using Windows API
- Restarts the bot automatically

### Technical Implementation
- Uses `psutil` for process monitoring
- Uses `pywinauto` for UI automation
- Uses `win32gui` for window management
- Implements proper waiting mechanisms for UI elements
- Uses threading for non-blocking operations

## Configuration

The script automatically looks for `WRobot.exe` in the same directory. If you need to change the path, modify the `WROBOT_PATH` variable:

```python
WROBOT_PATH = "C:\\path\\to\\your\\WRobot.exe"
```

## Troubleshooting

**WRobot not starting:**
- Ensure WRobot.exe is in the same directory
- Check if the path in `WROBOT_PATH` is correct
- Run the script as administrator if needed

**UI automation issues:**
- Make sure WRobot window is not minimized
- Check Windows scaling settings (100% recommended)
- Ensure no other automation software is interfering

**Connection timeout:**
- Check your internet connection
- Ensure WRobot servers are accessible
- The script waits up to 2 minutes for server connection

## Legal Notice

This tool is designed for educational purposes and personal use only. Users are responsible for compliance with:
- World of Warcraft Terms of Service
- WRobot licensing terms
- Local laws and regulations

## Contributing

Feel free to submit issues, fork the repository, and create pull requests for any improvements.

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Disclaimer

This software is provided "as is" without warranty of any kind. Use at your own risk.