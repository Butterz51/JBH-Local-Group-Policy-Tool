
import base64
import ctypes
import json
import os
import subprocess
import sys
import threading
import queue
import tkinter as tk
import webbrowser
from datetime import datetime
from tkinter import filedialog, messagebox, ttk

APP_TITLE = "JBH Services Local Group Policy Editor"
DEFAULT_SAVE_NAME = "Config"
BACKEND_SCRIPT_RELATIVE_PATH = os.path.join("Data", "Scripts", "PowerShell", "LocalGroupPolicyTool.ps1")
CONFIGS_RELATIVE_DIR = os.path.join("Data", "Saved Configs")
LOGS_RELATIVE_DIR = "Logs"
APP_UPDATE_URL = "https://github.com/Butterz51/JBH-Local-Group-Policy-Tool"
APP_AUTHOR = "Butterz51 / JBH Services"
APP_VERSION = "2.5.2"
APP_BUILD = "0330.26"

THEMES = {
    "dark": {
        "BG": "#1C1C1C",
        "FG": "#F2F2F2",
        "ACCENT": "#14A8E5",
        "BUTTON_BG": "#14A8E5",
        "BUTTON_FG": "#000000",
        "VCOLOR": "#FDA015",
        "DISABLED_BUTTON_BG": "#3A3A3A",
        "DISABLED_BUTTON_FG": "#9A9A9A",
        "NOTICE_FG": "#9C1515",
        "COMBO_BG": "#2A2A2A",
        "COMBO_DISABLED_BG": "#3A3A3A",
        "PROGRESS_TROUGH": "#1E1E1E",
        "PROGRESS_FILL": "#28C840",
    },
    "light": {
        "BG": "#F3F3F3",
        "FG": "#111111",
        "ACCENT": "#0A84FF",
        "BUTTON_BG": "#0A84FF",
        "BUTTON_FG": "#FFFFFF",
        "VCOLOR": "#C96E00",
        "DISABLED_BUTTON_BG": "#D7D7D7",
        "DISABLED_BUTTON_FG": "#6C6C6C",
        "NOTICE_FG": "#A40000",
        "COMBO_BG": "#FFFFFF",
        "COMBO_DISABLED_BG": "#E6E6E6",
        "PROGRESS_TROUGH": "#D9D9D9",
        "PROGRESS_FILL": "#28A745",
    },
    # Future themes can be added here
}

BG = THEMES["dark"]["BG"]
FG = THEMES["dark"]["FG"]
ACCENT = THEMES["dark"]["ACCENT"]
BUTTON_BG = THEMES["dark"]["BUTTON_BG"]
BUTTON_FG = THEMES["dark"]["BUTTON_FG"]
VCOLOR = THEMES["dark"]["VCOLOR"]
DISABLED_BUTTON_BG = THEMES["dark"]["DISABLED_BUTTON_BG"]
DISABLED_BUTTON_FG = THEMES["dark"]["DISABLED_BUTTON_FG"]
NOTICE_FG = THEMES["dark"]["NOTICE_FG"]
COMBO_BG = THEMES["dark"]["COMBO_BG"]
COMBO_DISABLED_BG = THEMES["dark"]["COMBO_DISABLED_BG"]
PROGRESS_TROUGH = THEMES["dark"]["PROGRESS_TROUGH"]
PROGRESS_FILL = THEMES["dark"]["PROGRESS_FILL"]

AU_OPTIONS = [
    "2 - Notify for download and auto install",
    "3 - Auto download and notify for install",
    "4 - Auto download and schedule the install",
    "5 - Allow local admin to choose setting",
    "6 - Auto download and notify for install and restart",
]

DELIVERY_OPTIMIZATION_MODES = [
    "0 - HTTP only, no peering",
    "1 - HTTP blended with peering behind same NAT",
    "2 - HTTP blended with peering across private group",
    "3 - HTTP blended with Internet peering",
    "99 - Simple download mode",
]

SCHEDULE_MODES = [
    "Every day",
    "Every week",
]

DAYS_OF_WEEK = [
    "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"
]

TIMES = [f"{(hour % 12) or 12}:00 {'AM' if hour < 12 else 'PM'}" for hour in range(24)]

# Widths for combo boxes
OPTIMIZATION = 42  # Delivery Optimization
DEFERRAL = 10  # Feature Updates deferral
AU_OPTION = 45  # AU option
SCHEDULED_INSTALL = 14  # Scheduled install
SCHEDULED_INSTALL_DAY = 12 # Scheduled install day
SCHEDULED_INSTALL_TIME = 11  # Scheduled install time

SECTIONS = [
    {
        "column": 0,
        "title": "System / Sign-in",
        "items": [
            {"key": "sync_foreground_policy", "label": "Always wait for the network at computer startup and logon", "type": "bool", "default": True},
            {"key": "hide_fast_user_switching", "label": "Hide Fast User Switching", "type": "bool", "default": True},
            {"key": "remove_lock_computer", "label": "Remove Lock Computer", "type": "bool", "default": True},
            {"key": "remove_change_password", "label": "Remove Change Password", "type": "bool", "default": True},
            {"key": "remove_logoff", "label": "Remove Logoff", "type": "bool", "default": True},
        ],
    },
    {
        "column": 0,
        "title": "Personalization / Lock Screen",
        "items": [
            {"key": "no_lock_screen", "label": "Do not display the lock screen", "type": "bool", "default": True},
        ],
    },
    {
        "column": 0,
        "title": "File Explorer / Power Menu",
        "items": [
            {"key": "show_sleep_option_disabled", "label": "Do Not Show Sleep in the power options menu", "type": "bool", "default": True},
        ],
    },
    {
        "column": 0,
        "title": "Start Menu and Taskbar / Chat / Explorer",
        "items": [
            {"key": "hide_taskview_button", "label": "Hide the TaskView button", "type": "bool", "default": True},
            {"key": "configure_chat_icon", "label": "Do Not Show the Chat icon on the taskbar", "type": "bool", "default": True},
            {"key": "hide_recommended_sites", "label": "Remove Personalized Website Recommendations from Recommended", "type": "bool", "default": True},
            {"key": "hide_recommended_section", "label": "Remove Recommended section from Start Menu", "type": "bool", "default": True},
            {"key": "hide_most_used_list", "label": 'Do Not Show "Most used" list from Start Menu', "type": "bool", "default": True},
            {"key": "disable_searchbox_suggestions", "label": "Turn off display of recent search entries in the File Explorer search box", "type": "bool", "default": True},
        ],
    },
    {
        "column": 0,
        "title": "Privacy / Telemetry / CEIP / Error Reporting",
        "items": [
            {"key": "disable_feedback_notifications", "label": "Do not show feedback notifications", "type": "bool", "default": True},
            {"key": "allow_telemetry_zero", "label": "Do Not Allow Telemetry data collection", "type": "bool", "default": True},
            {"key": "max_telemetry_zero", "label": "Do Not Allow Max Telemetry", "type": "bool", "default": True},
            {"key": "disable_enterprise_auth_proxy", "label": "Disable Enterprise Auth Proxy", "type": "bool", "default": True},
            {"key": "ceip_disabled", "label": "Disable Customer Experience Improvement Program (CEIP)", "type": "bool", "default": True},
            {"key": "let_apps_run_in_background_deny", "label": "Do Not Let Windows apps run in the background", "type": "bool", "default": True},
            {"key": "disable_windows_error_reporting", "label": "Disable Windows Error Reporting", "type": "bool", "default": True},
        ],
    },
    {
        "column": 1,
        "title": "Account Notifications",
        "items": [
            {"key": "disable_account_notifications", "label": "Turn off account notifications in Start", "type": "bool", "default": True},
        ],
    },
    {
        "column": 1,
        "title": "Cloud Content / Consumer Experience",
        "items": [
            {"key": "disable_cloud_optimized_content", "label": "Turn off cloud optimized content", "type": "bool", "default": True},
            {"key": "disable_consumer_account_state_content", "label": "Turn off cloud consumer account state content", "type": "bool", "default": True},
            {"key": "disable_soft_landing", "label": "Do not show Windows tips", "type": "bool", "default": True},
            {"key": "disable_windows_consumer_features", "label": "Turn off Microsoft consumer experiences", "type": "bool", "default": True},
        ],
    },
    {
        "column": 1,
        "title": "OneDrive",
        "items": [
            {"key": "disable_default_save_to_onedrive", "label": "Do Not Save documents to OneDrive by default", "type": "bool", "default": True},
            {"key": "disable_default_save_to_skydrive", "label": "Do Not Save documents to OneDrive by default (legacy compatibility)", "type": "bool", "default": True},
            {"key": "disable_filesync_ngsc", "label": "Do Not Allow OneDrive sync client", "type": "bool", "default": True},
            {"key": "disable_filesync_legacy", "label": "Do Not Allow legacy file sync", "type": "bool", "default": True},
        ],
    },
    {
        "column": 1,
        "title": "Search / Widgets / Insider",
        "items": [
            {"key": "disable_web_search", "label": "Do not allow web search", "type": "bool", "default": True},
            {"key": "connected_search_use_web_disabled", "label": "Don't search the web or display web results in Search", "type": "bool", "default": True},
            {"key": "allow_widgets_disabled", "label": "Do Not Allow widgets", "type": "bool", "default": True},
            {"key": "hide_windows_insider_pages", "label": "Hide Windows Insider pages", "type": "bool", "default": True},
        ],
    },
    {
        "column": 1,
        "title": "Windows Security / Biometrics",
        "items": [
            {"key": "hide_account_protection", "label": "Hide the Account protection area", "type": "bool", "default": True},
            {"key": "prevent_user_from_modifying_settings", "label": "Prevent user from modifying settings", "type": "bool", "default": True},
            {"key": "hide_family_options", "label": "Hide the Family options area", "type": "bool", "default": True},
            {"key": "hide_non_critical_notifications", "label": "Hide non-critical notifications", "type": "bool", "default": True},
            {"key": "disable_biometrics", "label": "Do Not Allow the use of biometrics", "type": "bool", "default": True},
        ],
    },
    {
        "column": 2,
        "title": "Delivery Optimization",
        "items": [
            {"key": "configure_delivery_optimization_mode", "label": "Configure Delivery Optimization Download Mode", "type": "bool", "default": True},
            {
                "key": "delivery_optimization_mode",
                "label": "Download Mode",
                "type": "combo_required",
                "default": "0 - HTTP only, no peering",
                "values": DELIVERY_OPTIMIZATION_MODES,
                "width": OPTIMIZATION,
                "depends_on": "configure_delivery_optimization_mode",
            },
        ],
    },
    {
        "column": 2,
        "title": "Windows Update > Core",
        "items": [
            {"key": "exclude_drivers_with_wu", "label": "Do not include drivers with Windows Updates", "type": "bool", "default": True},
            {"key": "defer_feature_updates_enabled", "label": "Select when Preview Builds and Feature Updates are received", "type": "bool", "default": True},
            {
                "key": "defer_feature_updates_days",
                "label": "Feature Updates Deferral:",
                "type": "combo_required",
                "default": "93 days",
                "values": [f"{d} days" for d in range(1, 366)],
                "width": DEFERRAL,
                "depends_on": "defer_feature_updates_enabled",
            },
            {"key": "disable_temp_enterprise_feature_control", "label": "Do Not Enable features introduced via servicing that are off by default", "type": "bool", "default": True},
            {"key": "manage_preview_builds_disabled", "label": "Manage preview builds", "type": "bool", "default": True},
        ],
    },
    {
        "column": 2,
        "title": "Windows Update > Automatic Updates Schedule",
        "items": [
            {"key": "configure_automatic_updates", "label": "Configure Automatic Updates", "type": "bool", "default": True},
            {"key": "au_option", "label": "Option", "type": "combo_required", "default": "4 - Auto download and schedule the install", "values": AU_OPTIONS, "width": AU_OPTION, "depends_on": "configure_automatic_updates"},
            {"key": "scheduled_install_mode", "label": "Scheduled install", "type": "combo_required", "default": "Every week", "values": SCHEDULE_MODES, "width": SCHEDULED_INSTALL, "depends_on": "configure_automatic_updates"},
            {"key": "scheduled_install_day", "label": "Scheduled install day", "type": "combo_required", "default": "Saturday", "values": DAYS_OF_WEEK, "width": SCHEDULED_INSTALL_DAY, "depends_on": "configure_automatic_updates"},
            {"key": "scheduled_install_time", "label": "Scheduled install time", "type": "combo_required", "default": "2:00 AM", "values": TIMES, "width": SCHEDULED_INSTALL_TIME, "depends_on": "configure_automatic_updates"},
            {"key": "allow_mu_update_service", "label": "Install updates for other Microsoft products", "type": "bool", "default": True},
            {"key": "no_auto_reboot_with_logged_on_users", "label": "No auto-restart with logged on user for scheduled automatic updates installations", "type": "bool", "default": True},
        ],
    },
    {
        "column": 2,
        "title": "Spotlight / Cloud Content",
        "items": [
            {"key": "disable_windows_spotlight_features", "label": "Turn off all Windows spotlight features", "type": "bool", "default": True},
            {"key": "disable_windows_spotlight_on_settings", "label": "Turn off Windows Spotlight on Settings", "type": "bool", "default": True},
            {"key": "disable_windows_welcome_experience", "label": "Turn off the Windows Welcome Experience", "type": "bool", "default": True},
        ],
    },
]

def is_running_as_admin():
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False


def relaunch_as_admin():
    if getattr(sys, "frozen", False):
        executable = sys.executable
        params = " ".join(f'"{arg}"' for arg in sys.argv[1:])
    else:
        executable = sys.executable
        script_path = os.path.abspath(__file__)
        extra_args = " ".join(f'"{arg}"' for arg in sys.argv[1:])
        params = f'"{script_path}"'
        if extra_args:
            params += f" {extra_args}"

    result = ctypes.windll.shell32.ShellExecuteW(
        None,
        "runas",
        executable,
        params,
        None,
        1,
    )

    if result <= 32:
        raise RuntimeError("Administrator elevation was cancelled or failed.")

    sys.exit(0)

class LocalGroupPolicyGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.theme_mode = "system" # Default to system theme
        self.theme_mode_var = tk.StringVar(value=self.theme_mode)
        self.progress_after_id = None

        self._apply_theme_globals(self._resolve_theme_key(self.theme_mode))

        self.title(APP_TITLE)
        self.configure(bg=BG)
        self.geometry("1460x950") # Set an initial window size that accommodates all content without needing to resize immediately, while allowing for some flexibility on smaller screens 1460x915
        self.minsize(1200, 900) # Set a minimum size to prevent the window from being resized too small to display content properly
        self.resizable(False, False)

        self.values = {}
        self.controls = {}
        self.select_all_button = None
        self.deselect_all_button = None
        self.apply_button = None
        self.save_button = None
        self.load_button = None
        self.progress_bar = None
        self.progress_label = None
        self.restart_notice_label = None
        self.restart_required_in_memory = False
        self.apply_queue = queue.Queue()
        self.apply_worker = None
        self.cancel_button = None

        self.active_dropdown = None
        self.dropdown_owner = None

        self._build_menu()
        self._build_styles()
        self._build_body()
        self._build_footer()
        self._initialize_defaults()

        self.update_idletasks()

        self.after(150, lambda: self._apply_native_titlebar_theme(debug=True))

    def _detect_system_theme(self):
        if os.name != "nt":
            return "dark" # Default to dark on non-Windows platforms since the registry key won't be available and dark is generally easier on the eyes in that case
        
        try:
            import winreg
            with winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"
            ) as key:
                value, _ = winreg.QueryValueEx(key, "AppsUseLightTheme")
            return "light" if int(value) == 1 else "dark"
        except Exception:
            return "dark"


    def _resolve_theme_key(self, theme_mode):
        if theme_mode == "system":
            return self._detect_system_theme()
        return theme_mode


    def _apply_theme_globals(self, theme_key):
        global BG, FG, ACCENT, BUTTON_BG, BUTTON_FG, VCOLOR
        global DISABLED_BUTTON_BG, DISABLED_BUTTON_FG, NOTICE_FG
        global COMBO_BG, COMBO_DISABLED_BG, PROGRESS_TROUGH, PROGRESS_FILL

        theme = THEMES[theme_key]

        BG = theme["BG"]
        FG = theme["FG"]
        ACCENT = theme["ACCENT"]
        BUTTON_BG = theme["BUTTON_BG"]
        BUTTON_FG = theme["BUTTON_FG"]
        VCOLOR = theme["VCOLOR"]
        DISABLED_BUTTON_BG = theme["DISABLED_BUTTON_BG"]
        DISABLED_BUTTON_FG = theme["DISABLED_BUTTON_FG"]
        NOTICE_FG = theme["NOTICE_FG"]
        COMBO_BG = theme["COMBO_BG"]
        COMBO_DISABLED_BG = theme["COMBO_DISABLED_BG"]
        PROGRESS_TROUGH = theme["PROGRESS_TROUGH"]
        PROGRESS_FILL = theme["PROGRESS_FILL"]

    def _hex_to_colorref(self, hex_color: str) -> int:
        value = hex_color.lstrip("#")
        r = int(value[0:2], 16)
        g = int(value[2:4], 16)
        b = int(value[4:6], 16)
        return (b << 16) | (g << 8) | r  # COLORREF = 0x00bbggrr

    def _get_windows_build(self) -> int:
        if os.name != "nt":
            return 0
        try:
            return sys.getwindowsversion().build
        except Exception:
            return 0

    def _get_top_level_hwnd(self, window=None) -> int:
        target = window or self
        target.update_idletasks()

        hwnd = 0
        try:
            hwnd = ctypes.windll.user32.GetParent(target.winfo_id())
        except Exception:
            hwnd = 0

        if not hwnd:
            hwnd = target.winfo_id()

        return hwnd

    def _apply_native_titlebar_theme(self, window=None, debug=False):
        if os.name != "nt":
            return False, "Not running on Windows."

        target = window or self
        target.update_idletasks()

        hwnd = self._get_top_level_hwnd(target)
        if not hwnd:
            return False, "Could not resolve a native window handle."

        build = self._get_windows_build()

        # Since your task says: "Set Title-bar to use System Theme"
        # this follows the actual Windows system theme, not the app override menu.
        resolved_theme = self._detect_system_theme()
        use_dark = resolved_theme == "dark"

        set_attr = ctypes.windll.dwmapi.DwmSetWindowAttribute

        results = []

        # Stage 1: request system dark/light title-bar behavior
        immersive_value = ctypes.c_int(1 if use_dark else 0)

        for attr in (20, 19):
            try:
                hr = set_attr(
                    hwnd,
                    attr,
                    ctypes.byref(immersive_value),
                    ctypes.sizeof(immersive_value),
                )
                results.append(f"DWMWA {attr} -> HRESULT {hr}")
                if hr == 0:
                    break
            except Exception as exc:
                results.append(f"DWMWA {attr} -> EXCEPTION {exc}")

        # Stage 2: on Windows 11+, force caption/text colors directly
        if build >= 22000:
            try:
                caption_color = ctypes.c_int(
                    self._hex_to_colorref(BG if use_dark else THEMES['light']['BG'])
                )
                text_color = ctypes.c_int(
                    self._hex_to_colorref(FG if use_dark else THEMES['light']['FG'])
                )

                hr_caption = set_attr(
                    hwnd,
                    35,  # DWMWA_CAPTION_COLOR
                    ctypes.byref(caption_color),
                    ctypes.sizeof(caption_color),
                )
                results.append(f"DWMWA_CAPTION_COLOR (35) -> HRESULT {hr_caption}")

                hr_text = set_attr(
                    hwnd,
                    36,  # DWMWA_TEXT_COLOR
                    ctypes.byref(text_color),
                    ctypes.sizeof(text_color),
                )
                results.append(f"DWMWA_TEXT_COLOR (36) -> HRESULT {hr_text}")

            except Exception as exc:
                results.append(f"Caption/Text color stage -> EXCEPTION {exc}")

        # Force a visible redraw of the non-client frame
        try:
            previous_state = target.state()
        except Exception:
            previous_state = "normal"

        try:
            target.withdraw()
            target.update_idletasks()

            if previous_state == "zoomed":
                target.deiconify()
                target.state("zoomed")
            elif previous_state == "iconic":
                target.iconify()
            else:
                target.deiconify()

            target.update_idletasks()
        except Exception as exc:
            results.append(f"Redraw stage -> EXCEPTION {exc}")

        if debug:
            print(f"[TitleBarDebug] Windows build: {build}")
            print(f"[TitleBarDebug] Theme: {resolved_theme}")
            print(f"[TitleBarDebug] HWND: {hwnd}")
            for line in results:
                print(f"[TitleBarDebug] {line}")

        success = any("HRESULT 0" in line for line in results)
        return success, " | ".join(results)
    
    def _debug_apply_titlebar_theme(self):
        ok, message = self._apply_native_titlebar_theme(debug=True)
        messagebox.showinfo(
            "Title Bar Debug",
            f"Success: {ok}\n\n{message}"
        )

    def _change_theme(self, theme_mode):
        if self.apply_worker is not None and self.apply_worker.is_alive():
            messagebox.showwarning(
                "Theme Change Blocked",
                "Wait for Apply to finish before changing themes."
            )
            self.theme_mode_var.set(self.theme_mode)
            return

        self.theme_mode = theme_mode
        self.theme_mode_var.set(theme_mode)
        self._rebuild_ui_for_theme()

    def _stop_progress_animation(self):
        if getattr(self, "progress_after_id", None) is not None:
            try:
                self.after_cancel(self.progress_after_id)
            except Exception:
                pass
            self.progress_after_id = None

    def _create_menu_button(self, parent, text):
        return tk.Button(
            parent,
            text=text,
            bg=BG,
            fg=FG,
            activebackground=BG,
            activeforeground=FG,
            highlightthickness=0,
            bd=0,
            relief="flat",
            padx=8,
            pady=3,
            font=("Segoe UI", 9),
            anchor="center",
        )

    def _close_active_dropdown(self):
        if self.active_dropdown is not None:
            try:
                self.active_dropdown.destroy()
            except Exception:
                pass

        self.active_dropdown = None
        self.dropdown_owner = None


    def _create_dropdown_window(self, owner):
        self._close_active_dropdown()

        popup = tk.Toplevel(self)
        popup.withdraw()
        popup.overrideredirect(True)
        popup.configure(bg=ACCENT, bd=0, highlightthickness=0)

        body = tk.Frame(
            popup,
            bg=BG,
            bd=0,
            highlightthickness=0,
        )
        body.pack(fill="both", expand=True, padx=1, pady=1)

        owner.update_idletasks()
        x = owner.winfo_rootx()
        y = owner.winfo_rooty() + owner.winfo_height() + 1

        self.active_dropdown = popup
        self.dropdown_owner = owner

        popup.update_idletasks()
        width = max(210, popup.winfo_reqwidth())
        height = max(10, popup.winfo_reqheight())
        popup.geometry(f"{width}x{height}+{x}+{y}")

        popup.deiconify()
        popup.lift()

        try:
            popup.attributes("-topmost", True)
            popup.after(50, lambda: popup.attributes("-topmost", False))
        except Exception:
            pass

        return popup, body

    def _add_dropdown_row(self, parent, text, command, selected=False):
        row = tk.Frame(parent, bg=BG)
        row.pack(fill="x")

        indicator = tk.Label(
            row,
            text="✓" if selected else " ",
            bg=BG,
            fg=ACCENT if selected else FG,
            width=2,
            anchor="w",
            font=("Segoe UI", 9, "bold"),
        )
        indicator.pack(side="left", padx=(8, 4), pady=4)

        label = tk.Label(
            row,
            text=text,
            bg=BG,
            fg=FG,
            anchor="w",
            font=("Segoe UI", 9),
        )
        label.pack(side="left", fill="x", expand=True, padx=(0, 12), pady=4)

        def on_enter(_event=None):
            row.configure(bg=ACCENT)
            indicator.configure(bg=ACCENT, fg=BUTTON_FG)
            label.configure(bg=ACCENT, fg=BUTTON_FG)

        def on_leave(_event=None):
            row.configure(bg=BG)
            indicator.configure(bg=BG, fg=ACCENT if selected else FG)
            label.configure(bg=BG, fg=FG)

        def on_click(_event=None):
            self._close_active_dropdown()
            self.after(10, command)

        for widget in (row, indicator, label):
            widget.bind("<Enter>", on_enter)
            widget.bind("<Leave>", on_leave)
            widget.bind("<Button-1>", on_click)

        return row

    def _toggle_themes_dropdown(self, owner):
        if self.active_dropdown is not None and self.dropdown_owner == owner:
            self._close_active_dropdown()
            return

        popup, body = self._create_dropdown_window(owner)

        current = self.theme_mode_var.get()

        self._add_dropdown_row(
            body,
            "System Default",
            lambda: self._change_theme("system"),
            selected=(current == "system"),
        )
        self._add_dropdown_row(
            body,
            "Dark Mode",
            lambda: self._change_theme("dark"),
            selected=(current == "dark"),
        )
        self._add_dropdown_row(
            body,
            "Light Mode",
            lambda: self._change_theme("light"),
            selected=(current == "light"),
        )

        popup.update_idletasks()
        x = owner.winfo_rootx()
        y = owner.winfo_rooty() + owner.winfo_height() + 1
        popup.geometry(f"{max(210, popup.winfo_reqwidth())}x{popup.winfo_reqheight()}+{x}+{y}")
        popup.lift()

    def _toggle_help_dropdown(self, owner):
        if self.active_dropdown is not None and self.dropdown_owner == owner:
            self._close_active_dropdown()
            return

        popup, body = self._create_dropdown_window(owner)

        self._add_dropdown_row(
            body,
            "Check For Updates",
            self._show_version_info,
            selected=False,
        )

        self._add_dropdown_row(
            body,
            "Discord",
            self._open_discord_link,
            selected=False,
        )

        self._add_dropdown_row(
            body,
            "Donation link",
            self._open_donation_link,
            selected=False,
        )

        popup.update_idletasks()
        x = owner.winfo_rootx()
        y = owner.winfo_rooty() + owner.winfo_height() + 1
        popup.geometry(f"{max(190, popup.winfo_reqwidth())}x{popup.winfo_reqheight()}+{x}+{y}")
        popup.lift()

    def _build_custom_menu_bar(self, parent):
        left = tk.Frame(parent, bg=BG)
        left.pack(side="left", padx=(6, 0), pady=(2, 2))

        self.themes_menu_button = self._create_menu_button(left, "Themes")
        self.themes_menu_button.configure(command=self._open_themes_dropdown)
        self.themes_menu_button.pack(side="left", padx=(0, 2))

        self.help_menu_button = self._create_menu_button(left, "Help")
        self.help_menu_button.configure(command=self._open_help_dropdown)
        self.help_menu_button.pack(side="left")

    def _open_themes_dropdown(self):
        try:
            self._toggle_themes_dropdown(self.themes_menu_button)
        except Exception as exc:
            messagebox.showerror(
                "Themes Menu Error",
                f"The Themes dropdown could not be opened.\n\n{exc}"
            )

    def _open_help_dropdown(self):
        try:
            self._toggle_help_dropdown(self.help_menu_button)
        except Exception as exc:
            messagebox.showerror(
                "Help Menu Error",
                f"The Help dropdown could not be opened.\n\n{exc}"
            )

    def _rebuild_ui_for_theme(self):
        saved_settings = self.collect_settings() if self.values else {}
        restart_required = getattr(self, "restart_required_in_memory", False)
        current_geometry = self.geometry()

        self._close_active_dropdown()

        self.withdraw()

        try:
            self._stop_progress_animation()
            self._apply_theme_globals(self._resolve_theme_key(self.theme_mode))

            self.config(menu="")
            for child in self.winfo_children():
                child.destroy()

            self.progress_after_id = None

            self.title(APP_TITLE)
            self.configure(bg=BG)

            self.values = {}
            self.controls = {}
            self.select_all_button = None
            self.deselect_all_button = None
            self.apply_button = None
            self.save_button = None
            self.load_button = None
            self.progress_bar = None
            self.progress_label = None
            self.restart_notice_label = None
            self.restart_required_in_memory = False
            self.apply_queue = queue.Queue()
            self.apply_worker = None
            self.cancel_button = None
            self.active_dropdown = None
            self.dropdown_owner = None

            self._build_menu()
            self._build_styles()
            self._build_body()
            self._build_footer()
            self._initialize_defaults()

            for key, value in saved_settings.items():
                if key in self.values:
                    self.values[key].set(value)

            self._apply_control_dependencies()
            self._update_bulk_buttons_state()
            self._set_restart_notice_required(restart_required)

            self.update_idletasks()
            self.geometry(current_geometry)

        except Exception as exc:
            self.deiconify()
            messagebox.showerror(
                "Theme Change Error",
                f"The theme could not be rebuilt correctly.\n\n{exc}"
            )
            return

        self.deiconify()
        self.after(50, lambda: self._apply_native_titlebar_theme(debug=True))

    def _check_for_updates(self):
        try:
            webbrowser.open(APP_UPDATE_URL)
        except Exception as exc:
            messagebox.showerror(
                "Update Check Error",
                f"Could not open the update page.\n\n{exc}"
            )

    def _show_version_info(self):
        win = tk.Toplevel(self)
        win.title("Version Info")
        win.configure(bg=BG)
        win.resizable(False, False)
        win.transient(self)
        win.grab_set()

        outer = tk.Frame(win, bg=BG, padx=22, pady=22)
        outer.pack(fill="both", expand=True)

        tk.Label(
            outer,
            text=APP_TITLE,
            bg=BG,
            fg=FG,
            anchor="w",
            justify="left",
            font=("Segoe UI", 12, "bold"),
        ).pack(fill="x", pady=(0, 16))

        info_text = (
            f"Version: {APP_VERSION}\n"
            f"Build: {APP_BUILD}\n"
            f"Author: {APP_AUTHOR}\n"
            f"\n"
            f"Frontend: Python / Tkinter\n"
            f"Backend: PowerShell"
        )

        tk.Label(
            outer,
            text=info_text,
            bg=BG,
            fg=FG,
            anchor="w",
            justify="left",
            font=("Segoe UI", 10),
        ).pack(fill="x")

        buttons_row = tk.Frame(outer, bg=BG)
        buttons_row.pack(fill="x", pady=(22, 0))

        tk.Button(
            buttons_row,
            text="Check for Updates",
            command=self._check_for_updates,
            bg=BUTTON_BG,
            fg=BUTTON_FG,
            relief="flat",
            font=("Segoe UI", 10),
            width=16,
        ).pack(side="left")

        tk.Button(
            buttons_row,
            text="Close",
            command=win.destroy,
            bg=BUTTON_BG,
            fg=BUTTON_FG,
            relief="flat",
            font=("Segoe UI", 10),
            width=16,
        ).pack(side="right")

        win.after(100, lambda: self._apply_native_titlebar_theme(win, debug=True))
        win.focus_force()

    def _build_menu(self):
        self.config(menu="")

        wrap = tk.Frame(self, bg=BG, bd=0, highlightthickness=0)
        wrap.pack(fill="x", side="top")

        self.top_menu_bar = tk.Frame(
            wrap,
            bg=BG,
            height=30,
            bd=0,
            highlightthickness=0,
        )
        self.top_menu_bar.pack(fill="x", side="top")

        self._build_custom_menu_bar(self.top_menu_bar)

        tk.Frame(wrap, bg=ACCENT, height=1, bd=0, highlightthickness=0).pack(fill="x")

    def _open_discord_link(self):
        try:
            webbrowser.open(DISCORD_URL)
        except Exception as exc:
            messagebox.showerror(
                "Discord Link Error",
                f"Could not open the Discord link.\n\n{exc}"
            )

    def _open_donation_link(self):
        try:
            webbrowser.open(DONATION_URL)
        except Exception as exc:
            messagebox.showerror(
                "Donation Link Error",
                f"Could not open the donation link.\n\n{exc}"
            )

    def _build_styles(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        
        style.configure(
            "JBH.Horizontal.TProgressbar",
            troughcolor=PROGRESS_TROUGH,
            background=PROGRESS_FILL,
            bordercolor=ACCENT,
            lightcolor=PROGRESS_FILL,
            darkcolor=PROGRESS_FILL,
        )

        style.configure(
            "JBH.TCombobox",
            fieldbackground=COMBO_BG,
            background=COMBO_BG,
            foreground=FG,
            arrowcolor=ACCENT,
            bordercolor=ACCENT,
            lightcolor=ACCENT,
            darkcolor=ACCENT,
            insertcolor=FG,
            padding=3,
        )
        style.map(
            "JBH.TCombobox",
            fieldbackground=[
                ("readonly", COMBO_BG),
                ("disabled", COMBO_DISABLED_BG),
            ],
            background=[
                ("readonly", COMBO_BG),
                ("disabled", COMBO_DISABLED_BG),
            ],
            foreground=[
                ("readonly", FG),
                ("disabled", "#909090"),
            ],
            selectforeground=[
                ("readonly", FG),
                ("disabled", "#909090"),
            ],
            selectbackground=[
                ("readonly", COMBO_BG),
                ("disabled", COMBO_DISABLED_BG),
            ],
            arrowcolor=[
                ("readonly", ACCENT),
                ("disabled", "#7F7F7F"),
            ],
        )

        self.option_add("*TCombobox*Listbox.background", COMBO_BG)
        self.option_add("*TCombobox*Listbox.foreground", FG)
        self.option_add("*TCombobox*Listbox.selectBackground", ACCENT)
        self.option_add("*TCombobox*Listbox.selectForeground", BUTTON_FG)

    def _build_body(self):
        body = tk.Frame(self, bg=BG)
        body.pack(fill="both", expand=True, padx=18, pady=(18, 0))

        actions_frame = tk.Frame(body, bg=BG)
        actions_frame.pack(fill="x", anchor="w", pady=(0, 12))

        tk.Label(
            actions_frame,
            text="",
            bg=BG,
            fg=VCOLOR,
            font=("Segoe UI", 9, "bold")
        ).pack(side="right")

        self.select_all_button = tk.Button(
            actions_frame,
            text="Select All",
            command=self._select_all,
            bg=BUTTON_BG,
            fg=BUTTON_FG,
            width=14,
            relief="flat",
            font=("Segoe UI", 10),
        )
        self.select_all_button.pack(side="left", padx=(0, 12), ipadx=4, ipady=4)

        self.deselect_all_button = tk.Button(
            actions_frame,
            text="Deselect All",
            command=self._deselect_all,
            bg=BUTTON_BG,
            fg=BUTTON_FG,
            width=14,
            relief="flat",
            font=("Segoe UI", 10),
        )
        self.deselect_all_button.pack(side="left", ipadx=4, ipady=4)

        columns_frame = tk.Frame(body, bg=BG)
        columns_frame.pack(fill="both", expand=True)

        self.column_frames = []
        for idx in range(3):
            col = tk.Frame(columns_frame, bg=BG)
            col.grid(row=0, column=idx, sticky="n", padx=14)
            columns_frame.grid_columnconfigure(idx, weight=1, uniform="cols")
            self.column_frames.append(col)

        for section in SECTIONS:
            self._build_section(self.column_frames[section["column"]], section)

    def _build_section(self, parent, section):
        wrapper = tk.Frame(parent, bg=BG)
        wrapper.pack(fill="x", pady=(0, 12))

        tk.Label(wrapper, text="=" * 54, bg=BG, fg=FG, anchor="w", font=("Consolas", 9)).pack(fill="x")
        tk.Label(wrapper, text=f"# {section['title']}", bg=BG, fg=FG, anchor="w", font=("Segoe UI", 10, "bold")).pack(fill="x")
        tk.Label(wrapper, text="=" * 54, bg=BG, fg=FG, anchor="w", font=("Consolas", 9)).pack(fill="x", pady=(0, 6))

        content = tk.Frame(wrapper, bg=BG)
        content.pack(fill="x")

        for item in section["items"]:
            if item["type"] == "bool":
                var = tk.BooleanVar(value=item["default"])
                self.values[item["key"]] = var

                cb = tk.Checkbutton(
                    content,
                    text=item["label"],
                    variable=var,
                    onvalue=True,
                    offvalue=False,
                    bg=BG,
                    fg=FG,
                    activebackground=BG,
                    activeforeground=FG,
                    selectcolor=BG,
                    highlightthickness=0,
                    bd=0,
                    anchor="w",
                    font=("Segoe UI", 9),
                    padx=0,
                    pady=1,
                )
                cb.pack(fill="x", anchor="w")
                self.controls[item["key"]] = cb
            else:
                row = tk.Frame(content, bg=BG)
                row.pack(fill="x", pady=(1, 3), anchor="w")

                label_widget = tk.Label(
                    row,
                    text=item["label"],
                    bg=BG,
                    fg=FG,
                    anchor="w",
                    font=("Segoe UI", 9),
                )
                label_widget.pack(side="left")

                var = tk.StringVar(value=item["default"])
                self.values[item["key"]] = var

                combo_width = item.get("width")
                if combo_width is None:
                    combo_width = min(max(len(str(item["default"])) + 2, 10), 28)

                box = ttk.Combobox(
                    row,
                    textvariable=var,
                    values=item["values"],
                    width=combo_width,
                    state="readonly",
                    style="JBH.TCombobox",
                )
                box.pack(side="left", padx=(10, 0))
                self.controls[item["key"]] = box

    def _build_footer(self):
        footer = tk.Frame(self, bg=BG)
        footer.pack(fill="x", padx=18, pady=(10, 18))

        self.apply_button = tk.Button(
            footer,
            text="Apply",
            command=self.apply_settings,
            bg=BUTTON_BG,
            fg=BUTTON_FG,
            width=16,
            relief="flat",
            font=("Segoe UI", 10),
        )
        self.apply_button.pack(side="right", padx=(12, 0), ipadx=6, ipady=6)

        self.save_button = tk.Button(
            footer,
            text="Save Settings",
            command=self.save_settings,
            bg=BUTTON_BG,
            fg=BUTTON_FG,
            width=16,
            relief="flat",
            font=("Segoe UI", 10),
        )
        self.save_button.pack(side="right", padx=(12, 0), ipadx=6, ipady=6)

        self.load_button = tk.Button(
            footer,
            text="Load Saved File",
            command=self.load_settings,
            bg=BUTTON_BG,
            fg=BUTTON_FG,
            width=16,
            relief="flat",
            font=("Segoe UI", 10),
        )
        self.load_button.pack(side="right", padx=(12, 0), ipadx=6, ipady=6)

        self.cancel_button = tk.Button(
            footer,
            text="Cancel",
            command=self.destroy,
            bg=BUTTON_BG,
            fg=BUTTON_FG,
            width=16,
            relief="flat",
            font=("Segoe UI", 10),
        )
        self.cancel_button.pack(side="left", padx=(0, 12), ipadx=6, ipady=6)

        progress_wrap = tk.Frame(footer, bg=BG)
        progress_wrap.pack(side="left", fill="x", expand=True, padx=(10, 10))

        status_row = tk.Frame(progress_wrap, bg=BG)
        status_row.pack(fill="x", pady=(0, 4))

        self.progress_label = tk.Label(
            status_row,
            text="Idle",
            bg=BG,
            fg=FG,
            anchor="w",
            font=("Segoe UI", 9, "bold"),
        )
        self.progress_label.pack(side="left")

        self.restart_notice_label = tk.Label(
            status_row,
            text="",
            bg=BG,
            fg=NOTICE_FG,
            anchor="nw",
            font=("Segoe UI", 10, "bold"),
        )
        self.restart_notice_label.pack(side="left", padx=(150, 0))

        self.progress_bar = ttk.Progressbar(
            progress_wrap,
            mode="determinate",
            length=100,
            style="JBH.Horizontal.TProgressbar",
        )
        self.progress_bar.pack(fill="x", expand=True)

    def _initialize_defaults(self):
        self._apply_control_dependencies()

        for key in (
            "configure_automatic_updates",
            "defer_feature_updates_enabled",
            "configure_delivery_optimization_mode",
        ):
            self.values[key].trace_add("write", lambda *_: self._apply_control_dependencies())

        for var in self.values.values():
            if isinstance(var, tk.BooleanVar):
                var.trace_add("write", lambda *_: self._update_bulk_buttons_state())

        self._update_bulk_buttons_state()
        self._set_progress_idle()
        self._set_restart_notice_required(False)

    def _apply_control_dependencies(self):
        self._set_combo_state(
            "delivery_optimization_mode",
            self.values["configure_delivery_optimization_mode"].get(),
        )
        self._set_combo_state(
            "defer_feature_updates_days",
            self.values["defer_feature_updates_enabled"].get(),
        )

        schedule_enabled = self.values["configure_automatic_updates"].get()
        for key in (
            "au_option",
            "scheduled_install_mode",
            "scheduled_install_day",
            "scheduled_install_time",
        ):
            self._set_combo_state(key, schedule_enabled)

    def _set_combo_state(self, key, enabled):
        control = self.controls.get(key)
        if control:
            control.configure(state="readonly" if enabled else "disabled")

    def _select_all(self):
        for var in self.values.values():
            if isinstance(var, tk.BooleanVar):
                var.set(True)

    def _deselect_all(self):
        for var in self.values.values():
            if isinstance(var, tk.BooleanVar):
                var.set(False)

    def _update_bulk_buttons_state(self):
        bool_vars = [var for var in self.values.values() if isinstance(var, tk.BooleanVar)]
        if not bool_vars:
            return

        all_selected = all(var.get() for var in bool_vars)
        all_deselected = all(not var.get() for var in bool_vars)

        self._set_button_visual_state(self.select_all_button, not all_selected)
        self._set_button_visual_state(self.deselect_all_button, not all_deselected)

        if self.select_all_button is not None:
            self.select_all_button.configure(state="disabled" if all_selected else "normal")

        if self.deselect_all_button is not None:
            self.deselect_all_button.configure(state="disabled" if all_deselected else "normal")

    def collect_settings(self):
        return {key: var.get() for key, var in self.values.items()}

    def save_settings(self):
        data = self.collect_settings()
        configs_dir = self._get_saved_configs_dir()
        os.makedirs(configs_dir, exist_ok=True)

        path = filedialog.asksaveasfilename(
            title="Save GUI Settings",
            defaultextension=".json",
            initialfile=DEFAULT_SAVE_NAME + ".json",
            initialdir=configs_dir,
            filetypes=[("JSON files", "*.json")],
        )
        if not path:
            return

        with open(path, "w", encoding="utf-8") as fh:
            json.dump(data, fh, indent=2)

        messagebox.showinfo("Saved", f"Settings saved to:\n{path}")

    def load_settings(self):
        configs_dir = self._get_saved_configs_dir()
        os.makedirs(configs_dir, exist_ok=True)

        path = filedialog.askopenfilename(
            title="Load GUI Settings",
            initialdir=configs_dir,
            initialfile=DEFAULT_SAVE_NAME + ".json",
            filetypes=[("JSON files", "*.json")],
        )
        if not path:
            return

        with open(path, "r", encoding="utf-8") as fh:
            data = json.load(fh)

        for key, value in data.items():
            if key in self.values:
                self.values[key].set(value)

        self._apply_control_dependencies()

    def _get_runtime_dir(self):
        if getattr(sys, "frozen", False):
            return os.path.dirname(sys.executable)
        return os.path.dirname(os.path.abspath(__file__))

    def _get_source_dir(self):
        return os.path.dirname(os.path.abspath(__file__))

    def _iter_possible_app_roots(self):
        seen = set()
        roots = []

        for base in (
            self._get_runtime_dir(),
            self._get_source_dir(),
            os.getcwd(),
            getattr(sys, "_MEIPASS", None),
        ):
            if not base:
                continue

            current = os.path.abspath(base)
            roots.append(current)
            for _ in range(4):
                parent = os.path.dirname(current)
                if parent == current:
                    break
                roots.append(parent)
                current = parent

        for candidate in roots:
            normalized = os.path.abspath(candidate)
            if normalized in seen:
                continue
            seen.add(normalized)
            yield normalized

    def _get_app_root(self):
        for candidate in self._iter_possible_app_roots():
            expected_backend = os.path.join(candidate, BACKEND_SCRIPT_RELATIVE_PATH)
            if os.path.isfile(expected_backend):
                return candidate

        if getattr(sys, "frozen", False):
            return self._get_runtime_dir()

        return os.path.abspath(os.path.join(self._get_source_dir(), "..", "..", ".."))

    def _get_saved_configs_dir(self):
        return os.path.join(self._get_app_root(), CONFIGS_RELATIVE_DIR)

    def _get_logs_dir(self):
        return os.path.join(self._get_app_root(), LOGS_RELATIVE_DIR)

    def _ensure_backend_script(self):
        app_root = self._get_app_root()
        candidates = [
            os.path.join(app_root, BACKEND_SCRIPT_RELATIVE_PATH),
        ]

        meipass_dir = getattr(sys, "_MEIPASS", None)
        if meipass_dir:
            candidates.extend([
                os.path.join(meipass_dir, "LocalGroupPolicyTool.ps1"),
            ])

        candidates.extend([
            os.path.join(self._get_source_dir(), "LocalGroupPolicyTool.ps1"),
            os.path.join(os.getcwd(), "LocalGroupPolicyTool.ps1"),
        ])

        for candidate in candidates:
            if candidate and os.path.isfile(candidate):
                return candidate

        raise FileNotFoundError(
            "Could not find the backend PowerShell script at the expected path:\n"
            f"{os.path.join(app_root, BACKEND_SCRIPT_RELATIVE_PATH)}"
        )

    def _get_powershell_executable(self):
        system_root = os.environ.get("SystemRoot", r"C:\Windows")
        candidates = [
            os.path.join(system_root, "System32", "WindowsPowerShell", "v1.0", "powershell.exe"),
            os.path.join(system_root, "Sysnative", "WindowsPowerShell", "v1.0", "powershell.exe"),
            "powershell.exe",
        ]

        for candidate in candidates:
            if candidate.lower().endswith("powershell.exe") and os.path.isfile(candidate):
                return candidate
            if candidate == "powershell.exe":
                return candidate

        return "powershell.exe"

    def _extract_backend_summary(self, stdout_text):
        if not stdout_text:
            return {}

        candidates = [stdout_text.strip()]
        candidates.extend(
            line.strip()
            for line in stdout_text.splitlines()
            if line.strip()
        )

        for candidate in reversed(candidates):
            try:
                parsed = json.loads(candidate)
            except json.JSONDecodeError:
                continue

            if isinstance(parsed, dict):
                return parsed

        return {}

    def _load_backend_summary_from_file(self, path):
        if not path or not os.path.isfile(path):
            return {}

        try:
            with open(path, "r", encoding="utf-8") as fh:
                content = fh.read().strip()
        except OSError:
            return {}

        if not content:
            return {}

        try:
            parsed = json.loads(content)
        except json.JSONDecodeError:
            return {}

        return parsed if isinstance(parsed, dict) else {}

    def _set_button_visual_state(self, button, enabled):
        if button is None:
            return

        if enabled:
            button.configure(
                state="normal",
                bg=BUTTON_BG,
                fg=BUTTON_FG,
                activebackground=BUTTON_BG,
                activeforeground=BUTTON_FG,
                disabledforeground=DISABLED_BUTTON_FG,
            )
        else:
            button.configure(
                state="disabled",
                bg=DISABLED_BUTTON_BG,
                fg=DISABLED_BUTTON_FG,
                activebackground=DISABLED_BUTTON_BG,
                activeforeground=DISABLED_BUTTON_FG,
                disabledforeground=DISABLED_BUTTON_FG,
            )

    def _set_restart_notice_visible(self, visible):
        if self.restart_notice_label is None:
            return

        if visible:
            self.restart_notice_label.configure(
                text="A system restart is recommended before relying on all policy changes."
            )
        else:
            self.restart_notice_label.configure(text="")


    def _set_restart_notice_required(self, required):
        self.restart_required_in_memory = bool(required)
        self._set_restart_notice_visible(self.restart_required_in_memory)

    def _set_apply_running_state(self, is_running):
        enabled = not is_running

        self._set_button_visual_state(self.apply_button, enabled)
        self._set_button_visual_state(self.save_button, enabled)
        self._set_button_visual_state(self.load_button, enabled)
        self._set_button_visual_state(self.cancel_button, enabled)
        state = "disabled" if is_running else "normal"

        if self.apply_button is not None:
            self.apply_button.configure(state=state)
        if self.cancel_button is not None:
            self.cancel_button.configure(state=state)
        if self.save_button is not None:
            self.save_button.configure(state=state)
        if self.load_button is not None:
            self.load_button.configure(state=state)
        if self.select_all_button is not None:
            self.select_all_button.configure(state=state if not is_running else "disabled")
        if self.deselect_all_button is not None:
            self.deselect_all_button.configure(state=state if not is_running else "disabled")

        if is_running:
            self._set_button_visual_state(self.select_all_button, False)
            self._set_button_visual_state(self.deselect_all_button, False)
        else:
            self._update_bulk_buttons_state()

    def _set_progress_idle(self):
        if self.progress_label is not None:
            self.progress_label.configure(text="Idle")
        if self.progress_bar is not None:
            self.progress_bar.stop()
            self.progress_bar.configure(mode="determinate", value=0)

    def _set_progress_running(self):
        if self.progress_label is not None:
            self.progress_label.configure(text="Running")
        if self.progress_bar is not None:
            self.progress_bar.configure(mode="determinate")
            self.progress_bar.start(12)

    def _set_progress_finished(self):
        if self.progress_bar is not None:
            self.progress_bar.stop()
            self.progress_bar.configure(mode="determinate", value=100)
        if self.progress_label is not None:
            self.progress_label.configure(text="Finished")

    def _apply_backend_worker(self, data):
        try:
            summary = self.apply_backend(data)
            self.apply_queue.put(("success", summary))
        except Exception as exc:
            self.apply_queue.put(("error", str(exc)))

    def _poll_apply_worker(self):
        if self.apply_worker is not None and self.apply_worker.is_alive():
            self.after(150, self._poll_apply_worker)
            return

        self._set_apply_running_state(False)

        try:
            status, payload = self.apply_queue.get_nowait()
        except queue.Empty:
            self._set_progress_idle()
            messagebox.showerror(
                "Apply Failed",
                "The apply worker ended but no result was returned."
            )
            return

        if status == "error":
            self._set_progress_idle()
            messagebox.showerror(
                "Apply Failed",
                f"The selected policies could not be applied.\n\n{payload}"
            )
            return

        summary = payload
        self._set_progress_finished()
        self.update_idletasks()

        status_text = summary.get("Status", "Success")
        machine_processed = int(summary.get("MachineEntriesProcessed", summary.get("MachineEntriesApplied", 0)) or 0)
        user_processed = int(summary.get("UserEntriesProcessed", summary.get("UserEntriesApplied", 0)) or 0)
        machine_changed = int(summary.get("MachineEntriesChanged", machine_processed) or 0)
        user_changed = int(summary.get("UserEntriesChanged", user_processed) or 0)
        machine_added = summary.get("MachineEntriesAdded")
        machine_updated = summary.get("MachineEntriesUpdated")
        machine_removed = summary.get("MachineEntriesRemoved")
        user_added = summary.get("UserEntriesAdded")
        user_updated = summary.get("UserEntriesUpdated")
        user_removed = summary.get("UserEntriesRemoved")
        restart_recommended = bool(summary.get("RestartRecommended", False))
        self._set_restart_notice_required(restart_recommended)
        log_file = summary.get("LogFile", "")

        result_lines = [
            f"Status: {status_text}",
            "",
            f"Machine policy changes made: {machine_changed}",
            f"User policy changes made: {user_changed}",
            "",
            f"Machine policy entries processed: {machine_processed}",
            f"User policy entries processed: {user_processed}",
        ]

        if machine_added is not None and machine_updated is not None and machine_removed is not None:
            result_lines.extend([
                "",
                f"Machine breakdown: Added: {machine_added} / Updated: {machine_updated} / Removed: {machine_removed}",
            ])

        if user_added is not None and user_updated is not None and user_removed is not None:
            result_lines.extend([
                f"User breakdown: Added: {user_added} / Updated: {user_updated} / Removed: {user_removed}",
            ])

        if restart_recommended:
            result_lines.extend([
                "",
                f"A system restart is recommended\nbefore relying on all policy changes.",
            ])

        if log_file:
            result_lines.extend([
                "",
                f"Log:\n{log_file}",
            ])

        messagebox.showinfo("Apply Complete", "\n".join(result_lines))

        self._set_progress_idle()
        self._update_bulk_buttons_state()

    def apply_backend(self, data):
        backend_path = self._ensure_backend_script()
        log_dir = self._get_logs_dir()
        os.makedirs(log_dir, exist_ok=True)

        log_file = os.path.join(
            log_dir,
            f"{datetime.now().strftime('%m%d%Y')}.log",
        )

        result_file = os.path.join(
            log_dir,
            f"Result.{datetime.now().strftime('%m-%d-%Y_%H-%M-%S_%f')}.json",
        )

        config_json = json.dumps(data, separators=(",", ":"), ensure_ascii=False)
        config_json_base64 = base64.b64encode(config_json.encode("utf-8")).decode("ascii")

        command = [
            self._get_powershell_executable(),
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            backend_path,
            "-ConfigJsonBase64",
            config_json_base64,
            "-LogFilePath",
            log_file,
            "-ResultFilePath",
            result_file,
            "-FromGui",
        ]

        creationflags = 0
        if os.name == "nt":
            creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)

        completed = subprocess.run(
            command,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            cwd=self._get_app_root(),
            creationflags=creationflags,
        )

        summary = {}

        if os.path.isfile(result_file):
            try:
                with open(result_file, "r", encoding="utf-8-sig") as fh:
                    summary = json.load(fh)
            except Exception:
                summary = {}
            finally:
                try:
                    os.remove(result_file)
                except OSError:
                    pass

        if not summary:
            summary = self._extract_backend_summary(completed.stdout)

        if completed.returncode != 0:
            error_parts = [
                f"The PowerShell backend exited with code {completed.returncode}."
            ]

            summary_message = summary.get("Message") if isinstance(summary, dict) else None
            stderr_text = (completed.stderr or "").strip()
            stdout_text = (completed.stdout or "").strip()

            if summary_message:
                error_parts.append(summary_message)
            elif stderr_text:
                error_parts.append(stderr_text)
            elif stdout_text:
                error_parts.append(stdout_text)

            error_parts.append(f"Check the backend log here:\n{log_file}")
            raise RuntimeError("\n\n".join(error_parts))

        if not summary:
            raise RuntimeError(
                "The PowerShell backend completed but did not return a result summary.\n\n"
                f"Check the backend log here:\n{log_file}"
            )

        if "LogFile" not in summary:
            summary["LogFile"] = log_file

        return summary

    def apply_settings(self):
        data = self.collect_settings()

        selected = sum(1 for value in data.values() if value is True)

        message_lines = [
            "Review the selected settings before applying:",
            "",
            f"Checked policy items: {selected}",
        ]

        applied_details = []

        if data.get("configure_delivery_optimization_mode"):
            applied_details.append(f"Delivery Optimization: {data.get('delivery_optimization_mode')}")

        if data.get("defer_feature_updates_enabled"):
            applied_details.append(f"Feature Updates Deferral: {data.get('defer_feature_updates_days')}")

        if data.get("configure_automatic_updates"):
            applied_details.append(f"AU Option: {data.get('au_option')}")
            applied_details.append(f"Scheduled install: {data.get('scheduled_install_mode')}")
            applied_details.append(f"Day: {data.get('scheduled_install_day')}")
            applied_details.append(f"Time: {data.get('scheduled_install_time')}")

        if applied_details:
            message_lines.append("")
            message_lines.append("Applied configured options:")
            message_lines.extend(applied_details)

        confirm = messagebox.askokcancel(
            "Review and Apply",
            "\n".join(message_lines)
        )

        if not confirm:
            return
        
        while not self.apply_queue.empty():
            try:
                self.apply_queue.get_nowait()
            except queue.Empty:
                break
        
        self._set_apply_running_state(True)
        self._set_progress_running()
        self.update_idletasks()

        self.apply_worker = threading.Thread(
            target=self._apply_backend_worker,
            args=(data,),
            daemon=True,
        )
        self.apply_worker.start()
        self.after(150, self._poll_apply_worker)

if __name__ == "__main__":
    try:
        if not is_running_as_admin():
            relaunch_as_admin()
    except Exception as exc:
        messagebox.showerror(
            "Administrator Rights Required",
            f"This tool must be started as Administrator.\n\n{exc}"
        )
        sys.exit(1)

    app = LocalGroupPolicyGUI()
    app.mainloop()