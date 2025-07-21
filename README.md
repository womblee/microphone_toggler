# Easy Microphone Toggler 🎤🔇

A lightweight Windows utility to toggle microphone mute state with a global hotkey or system tray control. Supports multiple audio devices and customizable configuration.

<img width="256" height="256" alt="mute" src="https://github.com/user-attachments/assets/0d080615-93ad-4f32-b0f2-ecb29d62e720" />

## Features ✨

- 🔥 **Global hotkey** (configurable) for instant mute/unmute
- 🖱️ **System tray control** with visual mute status
- 🎚️ **Per-device control** (default or specific microphone)
- 🔊 **Custom sound effects** for mute/unmute actions
- ⚙️ **Fully configurable** via text file
- 🔄 **Runtime reload** of configuration
- 🛡️ **Low-level keyboard hook** option for better compatibility

## Installation 📥

1. Download the latest release from [Releases page]
2. Extract the ZIP file
3. Run `microphone_toggler.exe`
4. Configure `mic_config.txt` in the same directory
   
*Note: Requires Windows 10/11 (no additional dependencies needed)*
*Default Hotkey: `CTRL + Shift + F1`

## Configuration ⚙️

The application creates `mic_config.txt` on first run. Config showcase:

```ini
# Use system default microphone device
use_default_device = true

# Specific device name (only used if use_default_device = false)
# Run the program to generate 'available_devices.txt' with available devices
# Copy the exact device name from that file
device_name = 

# Hotkey modifier keys (can be combined by adding values):
#   Alt = 1, Control = 2, Shift = 4, Windows Key = 8
#   Examples: Control+Shift = 6, Alt+Control = 3, Shift only = 4
hotkey_mod = 6

# Main key for hotkey (virtual key codes):
#   Function keys: F1=112, F2=113, F3=114, ..., F12=123
#   Letters: A=65, B=66, C=67, ..., Z=90
#   Numbers: 0=48, 1=49, 2=50, ..., 9=57
#   Other: Space=32, Enter=13, Tab=9
hotkey_vk = 112

# Cooldown for toggling the microphone (0-60000, 0=none, 500=half a second, 1000=second)
toggle_cooldown = 1000

# Use low-level keyboard hook for more reliable hotkey detection
# (May work better in some apps like Visual Studio)
use_keyboard_hook = false

# Play notification sounds when muting/unmuting
play_sounds = true

# Volume for notification sounds (0-100, 0=silent, 100=loudest)
sound_volume = 50

# Sound files (must be WAV format, leave empty to disable specific sounds)
# Files should be in the same folder as this program
mute_sound_file = mute.wav
unmute_sound_file = unmute.wav

# Automatically unmute microphone when program exits
# Set to false if you want to keep the mute state when closing
unmute_on_exit = true
```

## Custom Device Selection 🎤
To see available audio devices, right-click tray icon → "List Audio Devices".

To use a custom device:
- Edit the line from `use_default_device = true` to `use_default_device = false` 
- Edit the line `device_name = YOUR DEVICE NAME` in `mic_config.txt`
  
## Building from Source 🛠️
Requirements:
- Visual Studio 2022

Steps:
- Clone repository
- Open `microphone_toggler.sln`
- Build `Release x64`
