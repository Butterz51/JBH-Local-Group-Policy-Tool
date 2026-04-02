import base64
import ctypes
import json
import os
import subprocess
import sys
import threading
import queue
import tkinter as tk
from datetime import datetime
from tkinter import filedialog, messagebox, ttk

APP_TITLE = "JBH Services Local Group Policy Editor"
APP_VERSION = "2.5.1"
DEFAULT_SAVE_NAME = "Config"
BACKEND_SCRIPT_RELATIVE_PATH = os.path.join("Data", "Scripts", "PowerShell", "LocalGroupPolicyTool.ps1")
CONFIGS_RELATIVE_DIR = os.path.join("Data", "Saved Configs")
LOGS_RELATIVE_DIR = "Logs"

BG = "#1C1C1C" # Main background color
FG = "#F2F2F2" # Main foreground color for text and controls
ACCENT = "#14A8E5" # Accent color for highlights and interactive elements
BUTTON_BG = "#14A8E5" # Background color for buttons
BUTTON_FG = "#000000" # Foreground color for buttons (black for better contrast against the bright accent background)
VCOLOR = "#FDA015"  # Version text color 405CFF
DISABLED_BUTTON_BG = "#3A3A3A" # A darker gray for disabled buttons to indicate they are inactive, while still fitting the overall dark theme
DISABLED_BUTTON_FG = "#9A9A9A" # A lighter gray for disabled button text to ensure it is still readable against the darker disabled button background, while clearly indicating the disabled state

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
##
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
##

class LocalGroupPolicyGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.configure(bg=BG)
        self.geometry("1400x915") # Set an initial window size that accommodates all content without needing to resize immediately, while allowing for some flexibility on smaller screens
        self.minsize(1200, 900) # Set a minimum size to prevent the window from being resized too small to display content properly
        self.resizable(False, False)

        self.values = {}
        self.controls = {}
        self.select_all_button = None
        self.deselect_all_button = None
        self.apply_button = None
        self.save_button = None
        self.load_button = None
        self.apply_queue = queue.Queue()
        self.apply_worker = None
        self.cancel_button = None

        self._build_styles()
        self._build_body()
        self._build_footer()
        self._initialize_defaults()

    def _build_styles(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        
        style.configure(
            "JBH.Horizontal.TProgressbar",
            troughcolor="#1E1E1E",
            background="#28C840",
            bordercolor=ACCENT,
            lightcolor="#28C840",
            darkcolor="#28C840",
        )

        style.configure(
            "JBH.TCombobox",
            fieldbackground="#2A2A2A",
            background="#2A2A2A",
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
                ("readonly", "#2A2A2A"),
                ("disabled", "#3A3A3A"),
            ],
            background=[
                ("readonly", "#2A2A2A"),
                ("disabled", "#3A3A3A"),
            ],
            foreground=[
                ("readonly", FG),
                ("disabled", "#D0D0D0"),
            ],
            selectforeground=[
                ("readonly", FG),
                ("disabled", "#D0D0D0"),
            ],
            selectbackground=[
                ("readonly", "#2A2A2A"),
                ("disabled", "#3A3A3A"),
            ],
            arrowcolor=[
                ("readonly", ACCENT),
                ("disabled", "#7F7F7F"),
            ],
        )

        self.option_add("*TCombobox*Listbox.background", "#2A2A2A")
        self.option_add("*TCombobox*Listbox.foreground", FG)
        self.option_add("*TCombobox*Listbox.selectBackground", ACCENT)
        self.option_add("*TCombobox*Listbox.selectForeground", "#2A2A2A")

    def _build_body(self):
        body = tk.Frame(self, bg=BG)
        body.pack(fill="both", expand=True, padx=18, pady=(18, 0))

        actions_frame = tk.Frame(body, bg=BG)
        actions_frame.pack(fill="x", anchor="w", pady=(0, 12))

        tk.Label(
            actions_frame,
            text=f"Version {APP_VERSION}",
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

                tk.Label(
                    row,
                    text=item["label"],
                    bg=BG,
                    fg=FG,
                    anchor="w",
                    font=("Segoe UI", 9),
                ).pack(side="left")

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
            messagebox.showerror(
                "Apply Failed",
                "The apply worker ended but no result was returned."
            )
            return

        if status == "error":
            messagebox.showerror(
                "Apply Failed",
                f"The selected policies could not be applied.\n\n{payload}"
            )
            return

        summary = payload
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

        completed = subprocess.run(
            command,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            cwd=self._get_app_root(),
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

        self.apply_worker = threading.Thread(
            target=self._apply_backend_worker,
            args=(data,),
            daemon=True,
        )
        self.apply_worker.start()
        self.after(150, self._poll_apply_worker)

if __name__ == "__main__":
    app = LocalGroupPolicyGUI()
    app.mainloop()