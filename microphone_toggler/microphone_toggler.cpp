#include <windows.h>
#include <mmdeviceapi.h>
#include <endpointvolume.h>
#include <audiopolicy.h>
#include <commctrl.h>
#include <shellapi.h>
#include <mmsystem.h>
#include <functiondiscoverykeys_devpkey.h>
#include <fstream>
#include <string>
#include <map>
#include <iostream>
#include <memory>
#include <algorithm>
#include <vector>
#include <chrono>

#include "resource.h"  // Required because (UN)MUTEICON is used below

#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "user32.lib")
#pragma comment(lib, "gdi32.lib")
#pragma comment(lib, "shell32.lib")
#pragma comment(lib, "comctl32.lib")
#pragma comment(lib, "winmm.lib")

// Constants
const int WM_TRAYICON = WM_USER + 1;
const int ID_TRAY_EXIT = 1001;
const int ID_TRAY_TOGGLE = 1002;
const int ID_TRAY_CONFIG = 1003;
const int ID_TRAY_RELOAD_CONFIG = 1004;
const int ID_TRAY_LIST_DEVICES = 1005;
const int HOTKEY_ID = 1;
const int MAX_SOUND_VOLUME = 100;
const int MIN_SOUND_VOLUME = 0;
const int MIN_TOGGLE_COOLDOWN = 0;
const int MAX_TOGGLE_COOLDOWN = 60000;

// Configuration structure
struct Config {
    UINT hotkey_mod = MOD_CONTROL | MOD_SHIFT;
    UINT hotkey_vk = VK_F1;
    bool use_keyboard_hook = true; // Default to true (set to false to use RegisterHotKey method)
    bool play_sounds = true;
    bool unmute_on_exit = true;
    bool use_default_device = true;
    int sound_volume = 50; // 0-100
    int toggle_cooldown = 1000; // Cooldown between each toggle of the program
    std::string device_name = "";  // Specific device name to use
    std::string mute_sound_file = "mute.wav";
    std::string unmute_sound_file = "unmute.wav";
    std::string config_file = "mic_config.txt";
    std::string devices_list_file = "available_devices.txt";
};

// Device information structure
struct AudioDevice {
    std::string id;
    std::string name;
    std::string description;
    bool is_default;
    bool is_enabled;
};

// RAII wrapper for COM interfaces
template<typename T>
class ComPtr {
private:
    T* ptr;

public:
    ComPtr() : ptr(nullptr) {}
    ComPtr(T* p) : ptr(p) {}
    ~ComPtr() { Release(); }

    // Move constructor
    ComPtr(ComPtr&& other) noexcept : ptr(other.ptr) {
        other.ptr = nullptr;
    }

    // Move assignment
    ComPtr& operator=(ComPtr&& other) noexcept {
        if (this != &other) {
            Release();
            ptr = other.ptr;
            other.ptr = nullptr;
        }
        return *this;
    }

    // Disable copy
    ComPtr(const ComPtr&) = delete;
    ComPtr& operator=(const ComPtr&) = delete;

    void Release() {
        if (ptr) {
            ptr->Release();
            ptr = nullptr;
        }
    }

    T* Get() const { return ptr; }
    T** GetAddressOf() { Release(); return &ptr; }
    T* operator->() const { return ptr; }
    operator bool() const { return ptr != nullptr; }
};

class MicrophoneController {
private:
    HWND main_hwnd;
    NOTIFYICONDATA notification_icon_data;
    ComPtr<IMMDeviceEnumerator> device_enumerator;
    ComPtr<IMMDevice> current_device;
    ComPtr<IAudioEndpointVolume> endpoint_volume;
    bool is_muted;
    bool initial_mute_state;
    Config config;
    bool com_initialized;
    bool hotkey_registered;
    bool tray_icon_added;
    HHOOK keyboard_hook = nullptr;
    bool use_keyboard_hook = false;
    DWORD sound_flags;
    std::vector<AudioDevice> available_devices;
    std::string current_device_name;

public:
    MicrophoneController() : main_hwnd(nullptr),
        is_muted(false), initial_mute_state(false),
        com_initialized(false), hotkey_registered(false),
        tray_icon_added(false) {
        memset(&notification_icon_data, 0, sizeof(NOTIFYICONDATA));

        // Pre-calculate sound flags for better performance
        sound_flags = SND_FILENAME | SND_ASYNC | SND_NODEFAULT | SND_NOSTOP;
    }

    ~MicrophoneController() {
        cleanup();
    }

    std::string wstring_to_string(const std::wstring& wstr) {
        if (wstr.empty()) return std::string();

        int size = WideCharToMultiByte(CP_UTF8, 0, wstr.c_str(), -1, nullptr, 0, nullptr, nullptr);
        std::string result(size - 1, 0);
        WideCharToMultiByte(CP_UTF8, 0, wstr.c_str(), -1, &result[0], size, nullptr, nullptr);
        return result;
    }

    std::wstring string_to_wstring(const std::string& str) {
        if (str.empty()) return std::wstring();

        int size = MultiByteToWideChar(CP_UTF8, 0, str.c_str(), -1, nullptr, 0);
        std::wstring result(size - 1, 0);
        MultiByteToWideChar(CP_UTF8, 0, str.c_str(), -1, &result[0], size);
        return result;
    }

    bool initialize_system() {
        // Initialize COM with better error handling
        HRESULT hr = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE);
        if (FAILED(hr) && hr != RPC_E_CHANGED_MODE) {
            return false;
        }
        com_initialized = true;

        // Initialize common controls (only what we need)
        INITCOMMONCONTROLSEX icex;
        icex.dwSize = sizeof(INITCOMMONCONTROLSEX);
        icex.dwICC = ICC_WIN95_CLASSES;
        if (!InitCommonControlsEx(&icex)) {
            return false;
        }

        return true;
    }

    std::vector<AudioDevice> enumerate_audio_devices() {
        std::vector<AudioDevice> devices;

        if (!device_enumerator) {
            HRESULT hr = CoCreateInstance(__uuidof(MMDeviceEnumerator), nullptr, CLSCTX_ALL,
                __uuidof(IMMDeviceEnumerator), (void**)device_enumerator.GetAddressOf());
            if (FAILED(hr)) return devices;
        }

        ComPtr<IMMDeviceCollection> device_collection;
        HRESULT hr = device_enumerator->EnumAudioEndpoints(eCapture, DEVICE_STATE_ACTIVE, device_collection.GetAddressOf());
        if (FAILED(hr)) return devices;

        // Get default device for comparison
        ComPtr<IMMDevice> default_device;
        device_enumerator->GetDefaultAudioEndpoint(eCapture, eConsole, default_device.GetAddressOf());

        UINT count;
        device_collection->GetCount(&count);

        for (UINT i = 0; i < count; i++) {
            ComPtr<IMMDevice> device;
            if (SUCCEEDED(device_collection->Item(i, device.GetAddressOf()))) {
                AudioDevice audio_device;

                // Get device ID
                LPWSTR device_id;
                if (SUCCEEDED(device->GetId(&device_id))) {
                    audio_device.id = wstring_to_string(device_id);
                    CoTaskMemFree(device_id);
                }

                // Get device properties
                ComPtr<IPropertyStore> property_store;
                if (SUCCEEDED(device->OpenPropertyStore(STGM_READ, property_store.GetAddressOf()))) {
                    PROPVARIANT prop_var;
                    PropVariantInit(&prop_var);

                    // Get friendly name
                    if (SUCCEEDED(property_store->GetValue(PKEY_Device_FriendlyName, &prop_var))) {
                        if (prop_var.vt == VT_LPWSTR) {
                            audio_device.name = wstring_to_string(prop_var.pwszVal);
                        }
                        PropVariantClear(&prop_var);
                    }

                    // Get device description
                    if (SUCCEEDED(property_store->GetValue(PKEY_Device_DeviceDesc, &prop_var))) {
                        if (prop_var.vt == VT_LPWSTR) {
                            audio_device.description = wstring_to_string(prop_var.pwszVal);
                        }
                        PropVariantClear(&prop_var);
                    }
                }

                // Check if this is the default device
                audio_device.is_default = false;
                if (default_device) {
                    LPWSTR default_id;
                    if (SUCCEEDED(default_device->GetId(&default_id))) {
                        audio_device.is_default = (audio_device.id == wstring_to_string(default_id));
                        CoTaskMemFree(default_id);
                    }
                }

                // Check device state
                DWORD state;
                audio_device.is_enabled = SUCCEEDED(device->GetState(&state)) && (state == DEVICE_STATE_ACTIVE);

                devices.push_back(audio_device);
            }
        }

        return devices;
    }

    void save_devices_list() {
        available_devices = enumerate_audio_devices();

        std::ofstream file(config.devices_list_file);
        if (file.is_open()) {
            file << "=== AVAILABLE AUDIO INPUT DEVICES ===\n\n";
            file << "Copy the exact device name (including spaces and special characters) to your config file.\n";
            file << "Use the 'device_name' setting in " << config.config_file << "\n\n";

            if (available_devices.empty()) {
                file << "No active audio input devices found!\n";
                file << "Make sure your microphone is connected and enabled.\n";
            }
            else {
                for (size_t i = 0; i < available_devices.size(); i++) {
                    const auto& device = available_devices[i];
                    file << "Device " << (i + 1) << ":\n";
                    file << "  Name: " << device.name << "\n";
                    file << "  Description: " << device.description << "\n";
                    file << "  Status: " << (device.is_enabled ? "Active" : "Inactive") << "\n";
                    file << "  Default: " << (device.is_default ? "Yes" : "No") << "\n";
                    file << "\n";

                    if (device.is_default) {
                        file << "  *** This is your system's default microphone ***\n\n";
                    }
                }

                file << "=== CONFIGURATION INSTRUCTIONS ===\n\n";
                file << "To use a specific device:\n";
                file << "1. Open " << config.config_file << "\n";
                file << "2. Set 'use_default_device = false'\n";
                file << "3. Set 'device_name = [exact device name from above]'\n\n";
                file << "Example:\n";
                file << "use_default_device = false\n";
                if (!available_devices.empty()) {
                    file << "device_name = " << available_devices[0].name << "\n";
                }
            }
        }
    }

    bool find_and_set_target_device() {
        if (!device_enumerator) {
            HRESULT hr = CoCreateInstance(__uuidof(MMDeviceEnumerator), nullptr, CLSCTX_ALL,
                __uuidof(IMMDeviceEnumerator), (void**)device_enumerator.GetAddressOf());
            if (FAILED(hr)) return false;
        }

        current_device.Release();
        endpoint_volume.Release();

        if (config.use_default_device) {
            // Use default device
            HRESULT hr = device_enumerator->GetDefaultAudioEndpoint(eCapture, eConsole, current_device.GetAddressOf());
            if (FAILED(hr)) return false;

            current_device_name = "Default Device";
        }
        else {
            // Find specific device by name
            if (config.device_name.empty()) {
                return false; // No device name specified
            }

            available_devices = enumerate_audio_devices();
            bool device_found = false;

            for (const auto& device : available_devices) {
                if (device.name == config.device_name && device.is_enabled) {
                    // Get the device by ID
                    std::wstring device_id_wide = string_to_wstring(device.id);
                    HRESULT hr = device_enumerator->GetDevice(device_id_wide.c_str(), current_device.GetAddressOf());
                    if (SUCCEEDED(hr)) {
                        current_device_name = device.name;
                        device_found = true;
                        break;
                    }
                }
            }

            if (!device_found) {
                return false;
            }
        }

        // Get endpoint volume interface
        HRESULT hr = current_device->Activate(__uuidof(IAudioEndpointVolume), CLSCTX_ALL,
            nullptr, (void**)endpoint_volume.GetAddressOf());
        if (FAILED(hr)) return false;

        // Get initial mute state and store it
        BOOL muted;
        hr = endpoint_volume->GetMute(&muted);
        if (FAILED(hr)) return false;

        is_muted = muted;
        initial_mute_state = muted;

        return true;
    }

    bool initialize_audio() {
        if (!find_and_set_target_device()) {
            if (!config.use_default_device && !config.device_name.empty()) {
                // Specific device not found, save device list for user
                save_devices_list();

                std::wstring error_msg = L"Could not find the specified microphone device: '";
                error_msg += string_to_wstring(config.device_name);
                error_msg += L"'\n\nA list of available devices has been saved to '";
                error_msg += string_to_wstring(config.devices_list_file);
                error_msg += L"'.\n\nPlease check this file and update your configuration.";

                MessageBox(nullptr, error_msg.c_str(), L"Device Not Found", MB_OK | MB_ICONWARNING);
                return false;
            }
            else if (!config.use_default_device && config.device_name.empty()) {
                // No device name specified but not using default
                save_devices_list();

                std::wstring error_msg = L"No device name specified in configuration.\n\n";
                error_msg += L"A list of available devices has been saved to '";
                error_msg += string_to_wstring(config.devices_list_file);
                error_msg += L"'.\n\nPlease choose a device from the list and update your configuration.";

                MessageBox(nullptr, error_msg.c_str(), L"Configuration Required", MB_OK | MB_ICONINFORMATION);
                return false;
            }
            else {
                // Default device failed
                save_devices_list();

                std::wstring error_msg = L"Could not access the default microphone device.\n\n";
                error_msg += L"Please check:\n";
                error_msg += L"1. Microphone is connected and enabled\n";
                error_msg += L"2. Audio drivers are properly installed\n";
                error_msg += L"3. Windows audio service is running\n\n";
                error_msg += L"A list of available devices has been saved to '";
                error_msg += string_to_wstring(config.devices_list_file);
                error_msg += L"' for reference.";

                MessageBox(nullptr, error_msg.c_str(), L"Audio System Error", MB_OK | MB_ICONSTOP);
                return false;
            }
        }

        return true;
    }

    void load_config() {
        std::ifstream file(config.config_file);
        if (!file.is_open()) {
            save_config(); // Create default config
            return;
        }

        std::string line;
        std::map<std::string, std::string> settings;

        while (std::getline(file, line)) {
            // Trim whitespace
            line.erase(0, line.find_first_not_of(" \t\r\n"));
            line.erase(line.find_last_not_of(" \t\r\n") + 1);

            // Skip comments and empty lines
            if (line.empty() || line[0] == '#' || line[0] == '=') continue;

            size_t pos = line.find('=');
            if (pos != std::string::npos) {
                std::string key = line.substr(0, pos);
                std::string value = line.substr(pos + 1);

                // Trim key and value
                key.erase(key.find_last_not_of(" \t") + 1);
                value.erase(0, value.find_first_not_of(" \t"));
                value.erase(value.find_last_not_of(" \t") + 1);

                settings[key] = value;
            }
        }

        // Parse settings with validation
        try {
            auto parse_bool = [](const std::string& val) {
                std::string lower;
                lower.reserve(val.size());
                std::transform(val.begin(), val.end(), std::back_inserter(lower), ::tolower);
                return (lower == "1" || lower == "true" || lower == "yes" || lower == "on");
                };

            auto parse_uint = [](const std::string& val) { return std::stoul(val); };
            auto parse_int_clamped = [](const std::string& val, int min, int max) {
                return max(min, min(max, std::stoi(val)));
                };

            const auto& s = settings; // alias for brevity

            if (s.find("hotkey_mod") != s.end()) config.hotkey_mod = parse_uint(s.at("hotkey_mod"));
            if (s.find("hotkey_vk") != s.end()) config.hotkey_vk = parse_uint(s.at("hotkey_vk"));
            if (s.find("toggle_cooldown") != s.end()) config.toggle_cooldown = parse_int_clamped(s.at("toggle_cooldown"), MIN_TOGGLE_COOLDOWN, MAX_TOGGLE_COOLDOWN);
            if (s.find("use_keyboard_hook") != s.end()) config.use_keyboard_hook = parse_bool(s.at("use_keyboard_hook"));
            if (s.find("play_sounds") != s.end()) config.play_sounds = parse_bool(s.at("play_sounds"));
            if (s.find("unmute_on_exit") != s.end()) config.unmute_on_exit = parse_bool(s.at("unmute_on_exit"));
            if (s.find("use_default_device") != s.end()) config.use_default_device = parse_bool(s.at("use_default_device"));
            if (s.find("sound_volume") != s.end()) config.sound_volume = parse_int_clamped(s.at("sound_volume"), MIN_SOUND_VOLUME, MAX_SOUND_VOLUME);
            if (s.find("device_name") != s.end()) config.device_name = s.at("device_name");
            if (s.find("mute_sound_file") != s.end()) config.mute_sound_file = s.at("mute_sound_file");
            if (s.find("unmute_sound_file") != s.end()) config.unmute_sound_file = s.at("unmute_sound_file");
        }
        catch (const std::exception&) {
            // If parsing fails, keep default values
        }
    }

    void save_config() {
        std::ofstream file(config.config_file);
        if (file.is_open()) {
            file << "===============================================\n";
            file << "    MICROPHONE CONTROLLER CONFIGURATION\n";
            file << "===============================================\n\n";

            file << "# Lines starting with # are comments\n";
            file << "# Boolean values: true/false, yes/no, 1/0, on/off\n\n";

            file << "=== DEVICE SELECTION ===\n\n";
            file << "# Use system default microphone device\n";
            file << "use_default_device = " << (config.use_default_device ? "true" : "false") << "\n\n";

            file << "# Specific device name (only used if use_default_device = false)\n";
            file << "# Run the program to generate '" << config.devices_list_file << "' with available devices\n";
            file << "# Copy the exact device name from that file\n";
            file << "device_name = " << config.device_name << "\n\n";

            file << "=== HOTKEY CONFIGURATION ===\n\n";
            file << "# Hotkey modifier keys (can be combined by adding values):\n";
            file << "#   Alt = 1, Control = 2, Shift = 4, Windows Key = 8\n";
            file << "#   Examples: Control+Shift = 6, Alt+Control = 3, Shift only = 4\n";
            file << "hotkey_mod = " << config.hotkey_mod << "\n\n";

            file << "# Main key for hotkey (virtual key codes):\n";
            file << "#   Function keys: F1=112, F2=113, F3=114, ..., F12=123\n";
            file << "#   Letters: A=65, B=66, C=67, ..., Z=90\n";
            file << "#   Numbers: 0=48, 1=49, 2=50, ..., 9=57\n";
            file << "#   Other: Space=32, Enter=13, Tab=9\n";
            file << "hotkey_vk = " << config.hotkey_vk << "\n\n";

            file << "# Cooldown for toggling the microphone (0-60000, 0=none, 500=half a second, 1000=second)\n";
            file << "toggle_cooldown = " << config.toggle_cooldown << "\n\n";

            file << "# Use low-level keyboard hook for more reliable hotkey detection\n";
            file << "# (May work better in some apps like Visual Studio)\n";
            file << "use_keyboard_hook = " << (config.use_keyboard_hook ? "true" : "false") << "\n\n";

            file << "=== SOUND SETTINGS ===\n\n";
            file << "# Play notification sounds when muting/unmuting\n";
            file << "play_sounds = " << (config.play_sounds ? "true" : "false") << "\n\n";

            file << "# Volume for notification sounds (0-100, 0=silent, 100=loudest)\n";
            file << "sound_volume = " << config.sound_volume << "\n\n";

            file << "# Sound files (must be WAV format, leave empty to disable specific sounds)\n";
            file << "# Files should be in the same folder as this program, or you can specify the exact path\n";
            file << "mute_sound_file = " << config.mute_sound_file << "\n";
            file << "unmute_sound_file = " << config.unmute_sound_file << "\n\n";

            file << "=== BEHAVIOR SETTINGS ===\n\n";
            file << "# Automatically unmute microphone when program exits\n";
            file << "# Set to false if you want to keep the mute state when closing\n";
            file << "unmute_on_exit = " << (config.unmute_on_exit ? "true" : "false") << "\n\n";

            file << "===============================================\n";
            file << "                QUICK SETUP\n";
            file << "===============================================\n\n";
            file << "1. Run this program to generate device list\n";
            file << "2. Check '" << config.devices_list_file << "' for available microphones\n";
            file << "3. Edit this config file with your preferred settings\n";
            file << "4. Right-click tray icon -> 'Reload Config' to apply changes\n\n";
            file << "Default hotkey: Ctrl+Shift+F1\n";
            file << "Left-click tray icon: Toggle mute\n";
            file << "Right-click tray icon: Show menu\n\n";
        }
    }

    void play_sound(const std::string& sound_file) {
        // Validate configuration and inputs
        if (!config.play_sounds || sound_file.empty()) {
            return;
        }

        // Clamp volume to valid range (0-100)
        int clamped_volume = config.sound_volume;
        if (clamped_volume < 0) {
            clamped_volume = 0;
        }
        else if (clamped_volume > 100) {
            clamped_volume = 100;
        }

        // Early exit if volume is zero (silent)
        if (clamped_volume == 0) {
            return;
        }

        // Check if file exists (more efficient than trying to play)
        DWORD file_attributes = GetFileAttributesA(sound_file.c_str());
        if (file_attributes == INVALID_FILE_ATTRIBUTES) {
            return; // File doesn't exist, silently skip
        }

        // Stop any currently playing sound before starting a new one
        PlaySoundA(nullptr, nullptr, 0); // Stop any currently playing sound

        // Set system volume for sound effects if volume is not 100%
        if (clamped_volume != 100) {
            // Calculate volume (Windows expects 0-0xFFFF range)
            DWORD volume = (DWORD)((clamped_volume / 100.0) * 0xFFFF);
            DWORD stereo_volume = (volume << 16) | volume;
            waveOutSetVolume(nullptr, stereo_volume);
        }

        // Play sound asynchronously with flags to allow interruption
        PlaySoundA(sound_file.c_str(), nullptr, SND_ASYNC | SND_FILENAME | SND_NODEFAULT);
    }

    bool create_main_window() {
        HINSTANCE hinstance = GetModuleHandle(nullptr);
        if (!hinstance) {
            return false;
        }

        // Use a more unique class name
        const wchar_t* class_name = L"MicController_MultiDevice_Enhanced";

        WNDCLASSW wc = { 0 };
        wc.lpfnWndProc = main_window_proc;
        wc.hInstance = hinstance;
        wc.lpszClassName = class_name;
        wc.hCursor = LoadCursor(nullptr, IDC_ARROW);
        wc.hbrBackground = nullptr; // No background needed for hidden window
        wc.style = 0; // Minimal style for hidden window

        if (!RegisterClassW(&wc)) {
            DWORD error = GetLastError();
            if (error != ERROR_CLASS_ALREADY_EXISTS) {
                return false;
            }
        }

        // Create minimal hidden window
        main_hwnd = CreateWindowW(
            class_name,
            L"Microphone Controller",
            WS_POPUP,       // Minimal window style
            0, 0, 1, 1,     // Minimal size
            nullptr,        // No parent
            nullptr,        // No menu
            hinstance,
            this            // Pass this pointer
        );

        return main_hwnd != nullptr;
    }

    bool setup_tray_icon() {
        notification_icon_data.cbSize = sizeof(NOTIFYICONDATA);
        notification_icon_data.hWnd = main_hwnd;
        notification_icon_data.uID = 1;
        notification_icon_data.uFlags = NIF_ICON | NIF_MESSAGE | NIF_TIP;
        notification_icon_data.uCallbackMessage = WM_TRAYICON;

        // Load appropriate icon
        notification_icon_data.hIcon = LoadIcon(GetModuleHandle(NULL), is_muted ? MAKEINTRESOURCE(MUTEICON) : MAKEINTRESOURCE(UNMUTEICON));

        std::wstring tooltip = is_muted ? L"🔇 MUTED" : L"🎤 UNMUTED";
        tooltip += L" - " + string_to_wstring(current_device_name);
        wcscpy_s(notification_icon_data.szTip, tooltip.c_str());

        bool success = Shell_NotifyIcon(NIM_ADD, &notification_icon_data);
        if (success) {
            tray_icon_added = true;
        }
        return success;
    }

    void update_tray_icon() {
        if (!tray_icon_added) return;

        notification_icon_data.hIcon = LoadIcon(GetModuleHandle(NULL), is_muted ? MAKEINTRESOURCE(MUTEICON) : MAKEINTRESOURCE(UNMUTEICON));
        std::wstring tooltip = is_muted ? L"🔇 MUTED" : L"🎤 UNMUTED";
        tooltip += L" - " + string_to_wstring(current_device_name);
        wcscpy_s(notification_icon_data.szTip, tooltip.c_str());
        Shell_NotifyIcon(NIM_MODIFY, &notification_icon_data);
    }

    bool register_global_hotkey() {
        bool success = RegisterHotKey(main_hwnd, HOTKEY_ID, config.hotkey_mod, config.hotkey_vk);
        if (success) {
            hotkey_registered = true;
        }
        return success;
    }

    bool should_handle_hotkey(KBDLLHOOKSTRUCT* kbStruct, WPARAM wParam) {
        // Only react to key down events
        if (wParam != WM_KEYDOWN && wParam != WM_SYSKEYDOWN) {
            return false;
        }

        // Check if the main key matches
        if (kbStruct->vkCode != config.hotkey_vk) {
            return false;
        }

        // Get current modifier states
        BYTE keyboardState[256];
        GetKeyboardState(keyboardState);

        // Check each modifier key state
        bool ctrl_pressed = (keyboardState[VK_CONTROL] & 0x80) != 0;
        bool alt_pressed = (keyboardState[VK_MENU] & 0x80) != 0;
        bool shift_pressed = (keyboardState[VK_SHIFT] & 0x80) != 0;
        bool win_pressed = (keyboardState[VK_LWIN] & 0x80) != 0 ||
            (keyboardState[VK_RWIN] & 0x80) != 0;

        // Determine which modifiers are required by config
        bool ctrl_required = (config.hotkey_mod & MOD_CONTROL) != 0;
        bool alt_required = (config.hotkey_mod & MOD_ALT) != 0;
        bool shift_required = (config.hotkey_mod & MOD_SHIFT) != 0;
        bool win_required = (config.hotkey_mod & MOD_WIN) != 0;

        // Check exact modifier match - no extra keys pressed
        bool exact_match = (ctrl_pressed == ctrl_required) &&
            (alt_pressed == alt_required) &&
            (shift_pressed == shift_required) &&
            (win_pressed == win_required);

        // Special case: If no modifiers are required, we need to verify no modifiers are pressed
        if (config.hotkey_mod == 0) {
            return exact_match && !ctrl_pressed && !alt_pressed &&
                !shift_pressed && !win_pressed;
        }

        return exact_match;
    }

    static LRESULT CALLBACK keyboard_hook_proc(int nCode, WPARAM wParam, LPARAM lParam) {
        if (nCode >= HC_ACTION) {
            KBDLLHOOKSTRUCT* kbStruct = (KBDLLHOOKSTRUCT*)lParam;

            // Get the controller instance (stored in window property)
            HWND hwnd = FindWindow(L"MicController_MultiDevice_Enhanced", nullptr);
            if (hwnd) {
                MicrophoneController* controller = (MicrophoneController*)GetWindowLongPtr(hwnd, GWLP_USERDATA);
                if (controller && controller->should_handle_hotkey(kbStruct, wParam)) {
                    controller->toggle_microphone_mute();
                    return 1; // Block the key from reaching other apps
                }
            }
        }
        return CallNextHookEx(nullptr, nCode, wParam, lParam);
    }

    void toggle_microphone_mute() {
        static std::chrono::steady_clock::time_point last_toggle_time;
        const std::chrono::milliseconds cooldown(config.toggle_cooldown);

        auto now = std::chrono::steady_clock::now();
        if (now - last_toggle_time < cooldown) {
            return; // Still in cooldown, ignore input
        }
        last_toggle_time = now;

        if (!endpoint_volume) return;

        is_muted = !is_muted;
        HRESULT hr = endpoint_volume->SetMute(is_muted, nullptr);

        if (SUCCEEDED(hr)) {
            // Play appropriate sound
            if (is_muted) {
                play_sound(config.mute_sound_file);
            }
            else {
                play_sound(config.unmute_sound_file);
            }

            update_tray_icon();
        }
    }

    void restore_initial_mute_state() {
        if (!endpoint_volume || !config.unmute_on_exit) return;

        // Always unmute on exit if unmute_on_exit is true
        if (is_muted) {
            endpoint_volume->SetMute(FALSE, nullptr);
        }
    }

    void reload_configuration() {
        // Unregister old hotkey
        if (hotkey_registered) {
            UnregisterHotKey(main_hwnd, HOTKEY_ID);
            hotkey_registered = false;
        }

        // Load new config
        load_config();

        // Reinitialize audio with new device settings
        if (!find_and_set_target_device()) {
            std::wstring error_msg = L"Failed to reinitialize audio device after config reload.\n";
            if (!config.use_default_device) {
                error_msg += L"Check your device_name setting in the config file.";
            }
            MessageBox(nullptr, error_msg.c_str(), L"Device Error", MB_OK | MB_ICONWARNING);
        }
        else {
            update_tray_icon(); // Update with new device name
        }

        // Register new hotkey
        if (!register_global_hotkey()) {
            MessageBox(nullptr,
                L"Failed to register new hotkey after config reload.\nThe key combination might be in use.",
                L"Hotkey Registration Failed", MB_OK | MB_ICONWARNING);
        }
    }

    static LRESULT CALLBACK main_window_proc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
        MicrophoneController* controller = nullptr;

        if (msg == WM_NCCREATE) {
            CREATESTRUCT* cs = (CREATESTRUCT*)lParam;
            controller = (MicrophoneController*)cs->lpCreateParams;
            SetWindowLongPtr(hwnd, GWLP_USERDATA, (LONG_PTR)controller);
        }
        else {
            controller = (MicrophoneController*)GetWindowLongPtr(hwnd, GWLP_USERDATA);
        }

        if (controller) {
            return controller->window_procedure(hwnd, msg, wParam, lParam);
        }

        return DefWindowProc(hwnd, msg, wParam, lParam);
    }

    LRESULT window_procedure(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
        switch (msg) {
        case WM_HOTKEY:
            // Only process if we're not using the keyboard hook
            if (!use_keyboard_hook && wParam == HOTKEY_ID) {
                toggle_microphone_mute();
            }
            break;

        case WM_TRAYICON:
            switch (lParam) {
            case WM_LBUTTONUP:
                toggle_microphone_mute();
                break;
            case WM_RBUTTONUP:
                show_context_menu();
                break;
            }
            break;

        case WM_COMMAND:
            handle_menu_command(LOWORD(wParam));
            break;

        case WM_DESTROY:
            PostQuitMessage(0);
            break;

        default:
            return DefWindowProc(hwnd, msg, wParam, lParam);
        }
        return 0;
    }

    void show_context_menu() {
        HMENU menu = CreatePopupMenu();
        if (!menu) return;

        // Add menu items
        std::string toggle_text = is_muted ? "Unmute Microphone" : "Mute Microphone";
        AppendMenuA(menu, MF_STRING, ID_TRAY_TOGGLE, toggle_text.c_str());
        AppendMenuA(menu, MF_SEPARATOR, 0, nullptr);
        AppendMenuA(menu, MF_STRING, ID_TRAY_LIST_DEVICES, "List Audio Devices");
        AppendMenuA(menu, MF_STRING, ID_TRAY_CONFIG, "Open Config File");
        AppendMenuA(menu, MF_STRING, ID_TRAY_RELOAD_CONFIG, "Reload Config");
        AppendMenuA(menu, MF_SEPARATOR, 0, nullptr);
        AppendMenuA(menu, MF_STRING, ID_TRAY_EXIT, "Exit");

        // Get cursor position
        POINT pt;
        GetCursorPos(&pt);

        // Required for proper menu behavior
        SetForegroundWindow(main_hwnd);

        // Show menu
        TrackPopupMenu(menu, TPM_RIGHTBUTTON | TPM_BOTTOMALIGN | TPM_LEFTALIGN,
            pt.x, pt.y, 0, main_hwnd, nullptr);

        DestroyMenu(menu);
    }

    void handle_menu_command(WORD command_id) {
        switch (command_id) {
        case ID_TRAY_TOGGLE:
            toggle_microphone_mute();
            break;

        case ID_TRAY_LIST_DEVICES:
            save_devices_list();
            {
                std::wstring msg = L"Audio device list has been saved to '";
                msg += string_to_wstring(config.devices_list_file);
                msg += L"'.\n\nWould you like to open the file now?";

                if (MessageBox(nullptr, msg.c_str(), L"Device List Generated",
                    MB_YESNO | MB_ICONINFORMATION) == IDYES) {
                    ShellExecuteA(nullptr, "open", config.devices_list_file.c_str(),
                        nullptr, nullptr, SW_SHOW);
                }
            }
            break;

        case ID_TRAY_CONFIG:
            ShellExecuteA(nullptr, "open", config.config_file.c_str(),
                nullptr, nullptr, SW_SHOW);
            break;

        case ID_TRAY_RELOAD_CONFIG:
            reload_configuration();
            MessageBox(nullptr, L"Configuration reloaded successfully!",
                L"Config Reload", MB_OK | MB_ICONINFORMATION);
            break;

        case ID_TRAY_EXIT:
            PostMessage(main_hwnd, WM_CLOSE, 0, 0);
            break;
        }
    }

    void cleanup() {
        // Restore microphone state if needed
        restore_initial_mute_state();

        // Remove tray icon
        if (tray_icon_added) {
            Shell_NotifyIcon(NIM_DELETE, &notification_icon_data);
            tray_icon_added = false;
        }

        // Unregister hotkey
        if (hotkey_registered && main_hwnd) {
            UnregisterHotKey(main_hwnd, HOTKEY_ID);
            hotkey_registered = false;
        }

        if (keyboard_hook) {
            UnhookWindowsHookEx(keyboard_hook);
            keyboard_hook = nullptr;
        }

        // Release COM objects (handled by ComPtr destructors)
        endpoint_volume.Release();
        current_device.Release();
        device_enumerator.Release();

        // Uninitialize COM
        if (com_initialized) {
            CoUninitialize();
            com_initialized = false;
        }

        // Destroy window
        if (main_hwnd) {
            DestroyWindow(main_hwnd);
            main_hwnd = nullptr;
        }
    }

    int run() {
        if (!initialize_system()) {
            MessageBox(nullptr, L"Failed to initialize system components.",
                L"Initialization Error", MB_OK | MB_ICONSTOP);
            return 1;
        }

        load_config();

        if (!initialize_audio()) {
            return 1; // Error already shown in initialize_audio()
        }

        if (!create_main_window()) {
            MessageBox(nullptr, L"Failed to create application window.",
                L"Window Creation Error", MB_OK | MB_ICONSTOP);
            return 1;
        }

        if (!setup_tray_icon()) {
            MessageBox(nullptr, L"Failed to create system tray icon.",
                L"Tray Icon Error", MB_OK | MB_ICONWARNING);
            return 1;
        }

        // Registering hotkey
        const std::wstring hotkey_error_msg =
            L"Failed to register global hotkey.\n"
            L"The key combination might already be in use by another application.\n\n"
            L"You can still use the tray icon to control the microphone.";

        if (config.use_keyboard_hook) {
            keyboard_hook = SetWindowsHookEx(WH_KEYBOARD_LL, keyboard_hook_proc, GetModuleHandle(nullptr), 0);
            if (!keyboard_hook) {
                MessageBox(nullptr,
                    L"Failed to install keyboard hook. Falling back to standard hotkey.",
                    L"Hook Error",
                    MB_OK | MB_ICONWARNING);
                if (!register_global_hotkey()) {
                    MessageBox(nullptr, hotkey_error_msg.c_str(),
                        L"Hotkey Registration Failed",
                        MB_OK | MB_ICONWARNING);
                }
            }
        }
        else if (!register_global_hotkey()) {
            MessageBox(nullptr, hotkey_error_msg.c_str(),
                L"Hotkey Registration Failed",
                MB_OK | MB_ICONWARNING);
        }

        // Message loop
        MSG msg;
        while (GetMessage(&msg, nullptr, 0, 0)) {
            TranslateMessage(&msg);
            DispatchMessage(&msg);
        }

        return 0;
    }
};

// Main entry point
int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow) {
    // Prevent multiple instances
    HANDLE mutex = CreateMutex(nullptr, TRUE, L"MicrophoneController_SingleInstance_Mutex");
    if (GetLastError() == ERROR_ALREADY_EXISTS) {
        MessageBox(nullptr, L"Microphone Controller is already running!\n\nCheck the system tray area.",
            L"Already Running", MB_OK | MB_ICONINFORMATION);
        if (mutex) CloseHandle(mutex);
        return 1;
    }

    MicrophoneController controller;
    int result = controller.run();

    if (mutex) {
        ReleaseMutex(mutex);
        CloseHandle(mutex);
    }

    return result;
}