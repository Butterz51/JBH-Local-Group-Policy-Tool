
"""
Policy descriptions and UI helpers for the JBH Services Local Group Policy GUI.

Purpose:
- Keep policy explanation text out of the main GUI file.
- Provide richer operator-facing descriptions for each policy.
- Offer a reusable Tkinter information panel that can be plugged into the main GUI.

This module is intentionally conservative with wording. Descriptions focus on the
practical effect that the user should expect in the GUI/tool context rather than
claiming deep operating-system internals for every policy.
"""

from __future__ import annotations

import time
import tkinter as tk
from tkinter import ttk
from typing import Any, Dict, Optional


DEFAULT_TITLE = "Policy Information"
DEFAULT_INTRO = (
    "Move over a policy, click it, or focus a dropdown to see a clearer description "
    "of what that setting affects, why it might be used, and what to watch out for."
)


POLICY_DETAILS: Dict[str, Dict[str, Any]] = {
    "sync_foreground_policy": {
        "title": "Always wait for the network at startup and logon",
        "summary": "Makes Windows wait for network availability before finishing startup and user logon processing.",
        "impact": "Useful when logon scripts, mapped drives, domain resources, or policy processing depend on the network being ready.",
        "consider": "Can increase sign-in time on slow or unstable networks. Usually more helpful on managed business systems than on standalone home PCs.",
    },
    "hide_fast_user_switching": {
        "title": "Hide Fast User Switching",
        "summary": "Removes the Fast User Switching option so another user cannot stay signed in while a different user signs in.",
        "impact": "Reduces shared-session behavior and can simplify accountability on multi-user systems.",
        "consider": "Less convenient on shared computers where multiple people regularly switch between active sessions.",
    },
    "remove_lock_computer": {
        "title": "Remove Lock Computer",
        "summary": "Removes the standard Lock command from the security or sign-out experience.",
        "impact": "Prevents users from leaving a session locked and returning later through the normal lock action.",
        "consider": "This can reduce security on machines where users should be able to briefly step away without fully signing out.",
    },
    "remove_change_password": {
        "title": "Remove Change Password",
        "summary": "Hides the built-in Change Password command from the Windows security options shown to the signed-in user.",
        "impact": "Useful where password changes are handled by another workflow or should not be initiated locally.",
        "consider": "Can frustrate users if local password changes are expected or required by policy.",
    },
    "remove_logoff": {
        "title": "Remove Logoff",
        "summary": "Removes the normal logoff option from the visible user commands.",
        "impact": "Can help enforce a controlled session flow in kiosk-like or tightly managed environments.",
        "consider": "May confuse users on normal desktops because signing out becomes less obvious or may require another method.",
    },
    "no_lock_screen": {
        "title": "Do not display the lock screen",
        "summary": "Skips the decorative lock screen layer and takes the user closer to the sign-in prompt.",
        "impact": "Reduces an extra click or key press and makes sign-in more direct.",
        "consider": "Best for streamlined business systems or lab systems. It removes a familiar part of the normal Windows sign-in experience.",
    },
    "show_sleep_option_disabled": {
        "title": "Hide Sleep from power options",
        "summary": "Removes the Sleep option from the visible power menu choices.",
        "impact": "Helps prevent users from placing the system into a low-power state when shutdown or restart should be used instead.",
        "consider": "May be undesirable on laptops or systems where Sleep is part of the expected workflow.",
    },
    "hide_taskview_button": {
        "title": "Hide Task View button",
        "summary": "Removes the Task View button from the taskbar.",
        "impact": "Cleans up the taskbar and reduces access to virtual desktop and recent-window workflows from that button.",
        "consider": "Users can still expect Task View features on some systems, so this is mainly a UI simplification change.",
    },
    "configure_chat_icon": {
        "title": "Hide Chat icon",
        "summary": "Removes the Chat or Teams-style chat entry point from the taskbar.",
        "impact": "Reduces taskbar clutter and limits access to consumer-style chat entry points.",
        "consider": "Good for business images that do not use the consumer taskbar chat experience.",
    },
    "hide_recommended_sites": {
        "title": "Remove personalized site recommendations",
        "summary": "Stops Windows from surfacing personalized website suggestions in supported recommendation areas.",
        "impact": "Reduces promotional or cloud-personalized content in the Start experience.",
        "consider": "Improves privacy and UI cleanliness, but also removes convenience suggestions some users may actually want.",
    },
    "hide_recommended_section": {
        "title": "Remove Recommended section from Start",
        "summary": "Hides the Recommended area from the Start menu.",
        "impact": "Creates a cleaner Start menu with less recent-file and suggestion content.",
        "consider": "Users lose quick access to recently opened items and suggested content from Start.",
    },
    "hide_most_used_list": {
        "title": 'Hide "Most used" list',
        "summary": "Prevents Windows from showing a frequently used apps list in Start.",
        "impact": "Reduces activity-based personalization in the Start menu.",
        "consider": "Removes a convenience feature for users who rely on Start to surface their common apps.",
    },
    "disable_searchbox_suggestions": {
        "title": "Disable File Explorer search box suggestions",
        "summary": "Stops File Explorer from displaying recent or suggested search entries in its search box.",
        "impact": "Reduces leftover search history visibility and makes Explorer search feel more private.",
        "consider": "Users lose quick recall of previous Explorer searches.",
    },
    "disable_feedback_notifications": {
        "title": "Disable feedback notifications",
        "summary": "Stops Windows from prompting users for feedback notifications.",
        "impact": "Reduces interruptions and removes another consumer-style prompt from the desktop experience.",
        "consider": "Useful on business systems that should avoid Microsoft feedback prompts entirely.",
    },
    "allow_telemetry_zero": {
        "title": "Do not allow telemetry data collection",
        "summary": "Targets the lowest telemetry posture intended by this tool's policy set.",
        "impact": "Reduces the amount of diagnostic or telemetry data sent from the system where supported.",
        "consider": "Telemetry behavior can vary by Windows edition and version. Some Microsoft services and troubleshooting workflows may have less data available.",
    },
    "max_telemetry_zero": {
        "title": "Do not allow max telemetry",
        "summary": "Keeps the device from using a higher diagnostic collection level than intended by this profile.",
        "impact": "Supports a privacy-focused or tightly managed baseline.",
        "consider": "Useful when you want consistency with a low-data diagnostic posture across the device.",
    },
    "disable_enterprise_auth_proxy": {
        "title": "Disable Enterprise Auth Proxy",
        "summary": "Disables the enterprise authentication proxy-related behavior covered by this policy item.",
        "impact": "Can reduce background enterprise authentication handling that is not needed on standalone or non-enterprise systems.",
        "consider": "Avoid on systems that rely on organization-specific sign-in or proxy-backed authentication workflows.",
    },
    "ceip_disabled": {
        "title": "Disable CEIP",
        "summary": "Disables participation in the Customer Experience Improvement Program.",
        "impact": "Reduces optional usage-data sharing intended to improve Microsoft products and experiences.",
        "consider": "Fits privacy-focused or compliance-focused configurations.",
    },
    "let_apps_run_in_background_deny": {
        "title": "Block background app activity",
        "summary": "Prevents supported Windows apps from continuing background activity when they are not actively in use.",
        "impact": "Can reduce background resource usage, notifications, syncing, and incidental data collection.",
        "consider": "Some apps may not update live content, send background notifications, or stay current until opened again.",
    },
    "disable_windows_error_reporting": {
        "title": "Disable Windows Error Reporting",
        "summary": "Turns off Windows Error Reporting behavior covered by this setting.",
        "impact": "Reduces automatic reporting of crash and fault information.",
        "consider": "Helpful for privacy or noise reduction, but it can reduce automatic diagnostics available when troubleshooting issues.",
    },
    "disable_account_notifications": {
        "title": "Turn off account notifications in Start",
        "summary": "Stops account-related prompts and notices from appearing in the Start experience.",
        "impact": "Reduces Microsoft account promotion or account reminder noise.",
        "consider": "Useful on local-account systems or business-managed devices where those prompts are not helpful.",
    },
    "disable_cloud_optimized_content": {
        "title": "Turn off cloud optimized content",
        "summary": "Reduces cloud-personalized or suggested content delivered through supported Windows surfaces.",
        "impact": "Helps keep the OS experience less promotional and more static.",
        "consider": "Users may see fewer personalized suggestions, tips, or cloud-driven recommendations.",
    },
    "disable_consumer_account_state_content": {
        "title": "Turn off cloud consumer account state content",
        "summary": "Prevents Windows from presenting certain consumer account-state messages and related prompts.",
        "impact": "Useful for reducing consumer-oriented messaging on managed or privacy-focused systems.",
        "consider": "Primarily a UI and experience cleanup setting.",
    },
    "disable_soft_landing": {
        "title": "Do not show Windows tips",
        "summary": "Disables Windows tips, hints, and guidance-style prompts.",
        "impact": "Reduces onboarding popups and feature-discovery messages.",
        "consider": "Recommended where users already know the environment or where prompts are seen as noise.",
    },
    "disable_windows_consumer_features": {
        "title": "Turn off Microsoft consumer experiences",
        "summary": "Disables consumer-focused suggestions, app promotions, and similar first-run or recommendation-style experiences covered by this policy.",
        "impact": "Common baseline choice for business images and cleaner desktop builds.",
        "consider": "Reduces promotional content, but also removes some Microsoft-suggested setup experiences.",
    },
    "disable_default_save_to_onedrive": {
        "title": "Do not save to OneDrive by default",
        "summary": "Prevents supported save workflows from preferring OneDrive as the default save target.",
        "impact": "Keeps local storage or other approved locations as the user's normal save path.",
        "consider": "Useful where OneDrive is not used, not licensed, or should not be the default location.",
    },
    "disable_default_save_to_skydrive": {
        "title": "Legacy OneDrive default-save compatibility",
        "summary": "Applies the same general default-save restriction for older or legacy policy paths associated with SkyDrive or OneDrive naming.",
        "impact": "Helps maintain consistency across Windows versions or policy path differences.",
        "consider": "This is a compatibility-oriented setting and pairs well with the main OneDrive default-save policy.",
    },
    "disable_filesync_ngsc": {
        "title": "Block OneDrive sync client",
        "summary": "Prevents use of the newer OneDrive sync client covered by this policy path.",
        "impact": "Stops file syncing through that client and helps keep files local or under another approved sync model.",
        "consider": "Do not enable if the environment depends on OneDrive sync for user documents or shared data access.",
    },
    "disable_filesync_legacy": {
        "title": "Block legacy file sync",
        "summary": "Prevents older file-sync behavior covered by the legacy policy path.",
        "impact": "Supports a stronger no-OneDrive or no-sync posture alongside the newer client restriction.",
        "consider": "Mostly relevant for compatibility coverage across policy variants.",
    },
    "disable_web_search": {
        "title": "Disable web search",
        "summary": "Stops Windows Search from using web-backed search behavior covered by this policy.",
        "impact": "Keeps search focused more on local content and reduces cloud/web result blending.",
        "consider": "Users lose quick web answers from the Windows search experience.",
    },
    "connected_search_use_web_disabled": {
        "title": "Do not display web results in Search",
        "summary": "Prevents Search from showing web results or using connected web search content.",
        "impact": "Improves privacy and avoids Bing-style result blending in the OS search surface.",
        "consider": "Best paired with the other web search restriction for a cleaner local-only search experience.",
    },
    "allow_widgets_disabled": {
        "title": "Disable widgets",
        "summary": "Removes or blocks the widgets experience covered by current Windows builds.",
        "impact": "Reduces background content feeds, taskbar clutter, and consumer-style information surfaces.",
        "consider": "Users lose quick access to weather, news, and widget-based panels from Windows.",
    },
    "hide_windows_insider_pages": {
        "title": "Hide Windows Insider pages",
        "summary": "Hides Windows Insider-related pages from Settings.",
        "impact": "Helps prevent casual access to preview-build enrollment paths on managed systems.",
        "consider": "Recommended for stable production systems that should not expose preview-program options.",
    },
    "hide_account_protection": {
        "title": "Hide Account protection area",
        "summary": "Hides the Account protection section inside Windows Security.",
        "impact": "Reduces access to that Settings or Security area for the signed-in user.",
        "consider": "Useful when the machine should present a simplified or locked-down security UI.",
    },
    "prevent_user_from_modifying_settings": {
        "title": "Prevent user from modifying settings",
        "summary": "Stops users from changing certain Windows Security settings covered by this policy.",
        "impact": "Helps enforce an admin-controlled security baseline.",
        "consider": "Good for managed environments, but it intentionally removes user flexibility.",
    },
    "hide_family_options": {
        "title": "Hide Family options area",
        "summary": "Hides the Family options section within Windows Security.",
        "impact": "Simplifies the Windows Security UI and removes a consumer-oriented section that may not be relevant in business deployments.",
        "consider": "Mainly a presentation and control-surface cleanup setting.",
    },
    "hide_non_critical_notifications": {
        "title": "Hide non-critical security notifications",
        "summary": "Suppresses less important Windows Security notifications while leaving more important ones to surface normally.",
        "impact": "Reduces notification noise and lowers alert fatigue.",
        "consider": "A cleaner experience, but users may miss low-priority reminders that would otherwise be visible.",
    },
    "disable_biometrics": {
        "title": "Disable biometrics",
        "summary": "Prevents the use of biometric sign-in or biometric authentication covered by this policy path.",
        "impact": "Supports environments that require passwords, smart cards, or other non-biometric controls.",
        "consider": "Do not enable if Windows Hello fingerprint or facial recognition is part of the intended sign-in experience.",
    },
    "configure_delivery_optimization_mode": {
        "title": "Configure Delivery Optimization download mode",
        "summary": "Turns on explicit Delivery Optimization configuration so the selected download mode is enforced.",
        "impact": "Lets the tool control whether update content is downloaded only from Microsoft or with peer-assist behavior.",
        "consider": "Best enabled when you want predictable Windows Update download behavior instead of the device default.",
    },
    "delivery_optimization_mode": {
        "title": "Delivery Optimization download mode",
        "summary": "Selects how Windows obtains update content when Delivery Optimization configuration is enabled.",
        "impact": "This affects whether peer sharing is disabled, limited to local scopes, or allowed more broadly.",
        "consider": "Safer and simpler choices reduce peer traffic. Broader peering can save bandwidth in some fleets but changes network behavior.",
    },
    "exclude_drivers_with_wu": {
        "title": "Exclude drivers from Windows Update",
        "summary": "Stops Windows Update from offering driver updates through the normal Windows Update flow.",
        "impact": "Useful when driver control should stay manual or vendor-managed.",
        "consider": "Helps prevent unwanted driver changes, but it also means hardware driver fixes may need to be handled separately.",
    },
    "defer_feature_updates_enabled": {
        "title": "Select when Feature Updates are received",
        "summary": "Turns on explicit control over how long Feature Updates are deferred.",
        "impact": "Useful for delaying major OS feature upgrades until testing or validation is complete.",
        "consider": "This is one of the most important stability settings in the tool for production systems.",
    },
    "defer_feature_updates_days": {
        "title": "Feature Update deferral period",
        "summary": "Sets the number of days Windows should delay Feature Updates after they become available through the applicable servicing logic.",
        "impact": "Longer values favor stability and testing time. Shorter values favor faster access to new platform changes.",
        "consider": "A moderate delay is often a practical balance for business systems.",
    },
    "disable_temp_enterprise_feature_control": {
        "title": "Block features introduced disabled-by-default through servicing",
        "summary": "Prevents Windows from enabling certain newly delivered servicing features that are shipped off by default.",
        "impact": "Helps keep the OS behavior stable after cumulative updates.",
        "consider": "Useful when you want fewer surprise UI or feature changes during normal servicing.",
    },
    "manage_preview_builds_disabled": {
        "title": "Manage preview builds",
        "summary": "Controls preview-build exposure as part of the update posture used by this tool.",
        "impact": "Supports a production-first update baseline that avoids preview or pre-release build channels.",
        "consider": "Recommended for normal business or stable personal systems. Preview exposure is better reserved for testing devices.",
    },
    "configure_automatic_updates": {
        "title": "Configure Automatic Updates",
        "summary": "Turns on explicit control of Windows Automatic Updates so the selected update option and schedule are applied.",
        "impact": "Lets the tool define whether updates are only notified, auto-downloaded, or scheduled for install.",
        "consider": "One of the highest-impact settings in the tool because it directly affects how patching behaves.",
    },
    "au_option": {
        "title": "Automatic Updates option",
        "summary": "Selects the Windows Automatic Updates behavior used when automatic update configuration is enabled.",
        "impact": "Different choices shift control between user prompts, automatic downloads, and scheduled installation.",
        "consider": "Pick the least disruptive option that still matches the maintenance expectations for the device.",
    },
    "scheduled_install_mode": {
        "title": "Scheduled install frequency",
        "summary": "Controls whether scheduled update installation follows a daily or weekly rhythm.",
        "impact": "Used only when the selected Automatic Updates behavior supports scheduled installation.",
        "consider": "Weekly scheduling is usually easier to align with a maintenance window.",
    },
    "scheduled_install_day": {
        "title": "Scheduled install day",
        "summary": "Sets the day used by the scheduled installation behavior.",
        "impact": "Helps align restart and maintenance expectations with when the device is normally idle.",
        "consider": "Best chosen based on real user downtime, not just convenience in the tool.",
    },
    "scheduled_install_time": {
        "title": "Scheduled install time",
        "summary": "Sets the hour used by the scheduled installation behavior.",
        "impact": "Affects when update installation is attempted on schedule-based automatic update modes.",
        "consider": "Choose a time when the device is usually powered on but not actively being used.",
    },
    "allow_mu_update_service": {
        "title": "Install updates for other Microsoft products",
        "summary": "Allows Microsoft Update to include updates for supported Microsoft products beyond core Windows.",
        "impact": "Can help keep Office and other Microsoft components current through the same update channel.",
        "consider": "Useful if you want a broader Microsoft patching posture. Disable only if another management platform owns those products.",
    },
    "no_auto_reboot_with_logged_on_users": {
        "title": "No auto-restart while a user is signed in",
        "summary": "Prevents scheduled automatic update installation from forcing an automatic restart while a user session is active.",
        "impact": "Helps protect against sudden interruption and possible data loss during a live user session.",
        "consider": "Strongly recommended on interactive systems unless enforced reboot timing is more important than user convenience.",
    },
    "disable_windows_spotlight_features": {
        "title": "Turn off all Windows Spotlight features",
        "summary": "Disables Windows Spotlight content, suggestions, rotating backgrounds, and related Spotlight experiences covered by policy.",
        "impact": "Creates a more static and predictable desktop and lock-screen experience.",
        "consider": "Good for business builds or clean minimal setups where promotional content should stay off.",
    },
    "disable_windows_spotlight_on_settings": {
        "title": "Turn off Spotlight on Settings",
        "summary": "Stops Spotlight-style content from appearing inside Settings surfaces that support it.",
        "impact": "Reduces promotional or discoverability content inside Settings.",
        "consider": "Mostly a UI cleanliness choice.",
    },
    "disable_windows_welcome_experience": {
        "title": "Turn off Windows Welcome Experience",
        "summary": "Prevents the post-update or first-run welcome-style experience from showing to the user.",
        "impact": "Reduces interruptions after updates and keeps the desktop workflow direct.",
        "consider": "Recommended for managed systems where users do not need feature-tour messaging.",
    },
}


OPTION_DETAILS: Dict[str, Dict[str, str]] = {
    "delivery_optimization_mode": {
        "0 - HTTP only, no peering": "Downloads content directly without peer sharing. This is the simplest and most predictable choice.",
        "1 - HTTP blended with peering behind same NAT": "Allows peer sharing only with devices that appear behind the same NAT boundary, which can reduce external bandwidth use inside one local site.",
        "2 - HTTP blended with peering across private group": "Allows peering across a broader private grouping than same-NAT only. Useful in some managed fleets but less restrictive than local-only behavior.",
        "3 - HTTP blended with Internet peering": "Allows the broadest peer-sharing posture listed here. This can improve efficiency in some environments but changes how update traffic is sourced.",
        "99 - Simple download mode": "Uses a simpler download behavior without normal peer-sharing participation. A straightforward choice when you want less Delivery Optimization complexity.",
    },
    "au_option": {
        "2 - Notify for download and auto install": "Users are notified that updates are available before download starts, but installation is still handled automatically afterward.",
        "3 - Auto download and notify for install": "Updates download automatically, but the user is notified before installation proceeds.",
        "4 - Auto download and schedule the install": "Downloads happen automatically and installation follows the configured schedule. Best when you want predictable maintenance timing.",
        "5 - Allow local admin to choose setting": "Hands more control to the local administrator for update behavior. Useful only if local admin discretion is part of the support model.",
        "6 - Auto download and notify for install and restart": "Downloads automatically, then notifies for install and restart steps rather than silently pushing through both stages.",
    },
    "scheduled_install_mode": {
        "Every day": "Attempts scheduled installation every day at the selected time. Better for systems that must patch quickly and are frequently available.",
        "Every week": "Attempts scheduled installation once per week on the selected day and time. Easier to align with a planned maintenance window.",
    },
}


def get_policy_detail(key: str, label: str = "") -> Dict[str, Any]:
    detail = POLICY_DETAILS.get(key)
    if detail:
        return detail

    fallback_title = label.strip() or key.replace("_", " ").title()
    return {
        "title": fallback_title,
        "summary": "This setting is available in the current configuration profile, but a dedicated description has not been written yet.",
        "impact": "Review the label and section name to confirm whether it fits the machine's intended management baseline.",
        "consider": "Use caution when enabling unfamiliar policy items until they are documented more fully.",
    }


def get_option_detail(key: str, value: Any) -> str:
    if value is None:
        return ""

    if key == "defer_feature_updates_days":
        return (
            f"The current selection is {value}. Longer deferrals favor stability and controlled rollout. "
            "Shorter deferrals favor getting platform changes sooner."
        )

    if key == "scheduled_install_day":
        return (
            f"The current scheduled day is {value}. Choose a day that aligns with a realistic maintenance window."
        )

    if key == "scheduled_install_time":
        return (
            f"The current scheduled time is {value}. Choose a time when the device is usually powered on but not actively used."
        )

    return OPTION_DETAILS.get(key, {}).get(str(value), "")


def format_policy_text(key: str, label: str = "", current_value: Any = None) -> str:
    detail = get_policy_detail(key, label)
    title = detail.get("title", label or key)

    lines = [
        title,
        "",
        f"What it does: {detail.get('summary', 'No description available.')}",
        f"Why you might use it: {detail.get('impact', 'No additional impact notes available.')}",
        f"What to watch for: {detail.get('consider', 'No caution notes available.')}",
    ]

    if current_value is not None:
        value_text = "Enabled" if current_value is True else "Disabled" if current_value is False else str(current_value)
        lines.extend(["", f"Current selection: {value_text}"])

        option_text = get_option_detail(key, current_value)
        if option_text:
            lines.append(f"Selection detail: {option_text}")

    return "\n".join(lines)


class PolicyInfoPanel(tk.Frame):
    """
    Reusable information panel that shows richer help text for a policy.

    Expected usage from the main GUI:
        panel = PolicyInfoPanel(parent, ...)
        panel.show_default()

        # for each control
        panel.bind_control(control, key=item["key"], label=item["label"], variable=self.values[item["key"]])
    """

    def __init__(
        self,
        parent: tk.Widget,
        *,
        bg: str,
        fg: str,
        accent: str,
        muted_fg: Optional[str] = None,
        title: str = DEFAULT_TITLE,
    ) -> None:
        super().__init__(parent, bg=bg, bd=0, highlightthickness=0)

        self._bg = bg
        self._fg = fg
        self._accent = accent
        self._muted_fg = muted_fg or fg

        self._frame = tk.Frame(
            self,
            bg=bg,
            bd=1,
            highlightthickness=1,
            highlightbackground=accent,
            highlightcolor=accent,
            padx=12,
            pady=10,
        )
        self._frame.pack(fill="x", expand=True)

        self._heading = tk.Label(
            self._frame,
            text=title,
            bg=bg,
            fg=accent,
            anchor="w",
            justify="left",
            font=("Segoe UI", 10, "bold"),
        )
        self._heading.pack(fill="x", pady=(0, 6))

        self._message = tk.Message(
            self._frame,
            text="",
            bg=bg,
            fg=fg,
            anchor="nw",
            justify="left",
            font=("Segoe UI", 9),
            width=1320,
        )
        self._message.pack(fill="x", expand=True)

        self.show_default()

        self.bind("<Configure>", self._handle_resize, add="+")

    def _handle_resize(self, _event: tk.Event) -> None:
        try:
            width = max(self.winfo_width() - 36, 400)
            self._message.configure(width=width)
        except Exception:
            pass

    def show_default(self) -> None:
        self._heading.configure(text=DEFAULT_TITLE)
        self._message.configure(text=DEFAULT_INTRO)

    def show_policy(self, key: str, label: str = "", current_value: Any = None) -> None:
        detail = get_policy_detail(key, label)
        self._heading.configure(text=detail.get("title", DEFAULT_TITLE))
        self._message.configure(text=format_policy_text(key, label, current_value))

    def bind_control(
        self,
        control: tk.Widget,
        *,
        key: str,
        label: str,
        variable: Optional[tk.Variable] = None,
    ) -> None:
        def show_now(_event: Optional[tk.Event] = None) -> None:
            value = variable.get() if variable is not None else None
            self.show_policy(key, label, value)

        def show_after_change(_event: Optional[tk.Event] = None) -> None:
            self.after_idle(show_now)

        for sequence in ("<Enter>", "<FocusIn>", "<Button-1>"):
            control.bind(sequence, show_now, add="+")

        for sequence in ("<ButtonRelease-1>", "<KeyRelease-space>", "<KeyRelease-Return>"):
            control.bind(sequence, show_after_change, add="+")

        if isinstance(control, ttk.Combobox):
            control.bind("<<ComboboxSelected>>", show_now, add="+")

        if variable is not None:
            def traced(*_args: Any) -> None:
                if self.focus_displayof() is not None:
                    self.after_idle(show_now)

            try:
                variable.trace_add("write", traced)
            except Exception:
                pass

def format_policy_hover_text(key: str, label: str = "", current_value: Any = None) -> str:
    detail = get_policy_detail(key, label)

    lines = [
        "Summary:",
        detail.get("summary", "No description available."),
        f"Use Case: {detail.get('impact', 'No additional impact notes available.')}",
        f"Watch For: {detail.get('consider', 'No caution notes available.')}",
    ]

    if current_value is not None:
        value_text = "Enabled" if current_value is True else "Disabled" if current_value is False else str(current_value)
        lines.append(f"Current selection: {value_text}")

        option_text = get_option_detail(key, current_value)
        if option_text:
            lines.append(f"Selection detail: {option_text}")

    return "\n".join(lines)


from Detail_Window_Positioning import position_detail_popup_window


class PolicyHoverDescription:
    """Floating hover popup. Does not reserve any layout space."""

    def __init__(
        self,
        parent: tk.Widget,
        *,
        bg: str,
        fg: str,
        accent: str,
        muted_fg: Optional[str] = None,
        delay_ms: int = 3000,
        lines: int = 6,
        min_width: int = 620,
        min_height: int = 132,
        reference_widget: Optional[tk.Widget] = None,
        left_gap: int = 18,
        top_offset: int = 0,
        right_margin: int = 0,
    ) -> None:
        self._parent_widget = parent
        self._reference_widget = reference_widget
        self._root_window = parent.winfo_toplevel()

        self._bg = bg
        self._fg = fg
        self._accent = accent
        self._muted_fg = muted_fg or fg
        self._delay_ms = max(0, int(delay_ms))
        self._line_count = max(5, int(lines))
        self._min_width = max(500, int(min_width))
        self._min_height = max(120, int(min_height))
        self._left_gap = max(0, int(left_gap))
        self._top_offset = int(top_offset)
        self._right_margin = max(0, int(right_margin))

        self._pending_after_id: Optional[str] = None
        self._pending_group: Optional[str] = None
        self._hide_after_id: Optional[str] = None
        self._ignore_leave_until = 0.0
        self._active_group: Optional[str] = None
        self._popup_visible = False

        self._bindings: Dict[str, Dict[str, Any]] = {}
        self._group_widgets: Dict[str, list[tk.Widget]] = {}
        self._hover_counts: Dict[str, int] = {}

        self._popup = tk.Toplevel(self._root_window)
        self._popup.withdraw()
        self._popup.overrideredirect(True)
        self._popup.transient(self._root_window)
        self._popup.configure(bg=self._accent, bd=0, highlightthickness=0)

        border = tk.Frame(
            self._popup,
            bg=self._accent,
            bd=0,
            highlightthickness=0,
            padx=1,
            pady=1,
        )
        border.pack(fill='both', expand=True)

        body = tk.Frame(border, bg=self._bg, bd=0, highlightthickness=0, padx=12, pady=10)
        body.pack(fill='both', expand=True)
        self._body = body

        heading_font = ('Segoe UI', 9, 'bold')
        value_font = ('Segoe UI', 9)

        self._summary_label = tk.Label(
            body,
            text='Summary:',
            bg=self._bg,
            fg=self._muted_fg,
            anchor='w',
            justify='left',
            font=heading_font,
        )
        self._summary_label.pack(anchor='w')

        self._summary_value = tk.Label(
            body,
            text='',
            bg=self._bg,
            fg=self._fg,
            anchor='w',
            justify='left',
            font=value_font,
            wraplength=self._min_width - 28,
        )
        self._summary_value.pack(fill='x', pady=(2, 0), padx=(16, 0), anchor='w')

        self._use_case_row = tk.Frame(body, bg=self._bg)
        self._use_case_row.pack(fill='x', pady=(8, 0), anchor='w')
        self._use_case_label = tk.Label(
            self._use_case_row,
            text='Use Case:',
            bg=self._bg,
            fg=self._muted_fg,
            anchor='w',
            justify='left',
            font=heading_font,
        )
        self._use_case_label.pack(side='left', anchor='nw')
        self._use_case_value = tk.Label(
            self._use_case_row,
            text='',
            bg=self._bg,
            fg=self._fg,
            anchor='w',
            justify='left',
            font=value_font,
            wraplength=self._min_width - 150,
        )
        self._use_case_value.pack(side='left', fill='x', expand=True, padx=(6, 0), anchor='nw')

        self._watch_for_row = tk.Frame(body, bg=self._bg)
        self._watch_for_row.pack(fill='x', pady=(2, 0), anchor='w')
        self._watch_for_label = tk.Label(
            self._watch_for_row,
            text='Watch For:',
            bg=self._bg,
            fg=self._muted_fg,
            anchor='w',
            justify='left',
            font=heading_font,
        )
        self._watch_for_label.pack(side='left', anchor='nw')
        self._watch_for_value = tk.Label(
            self._watch_for_row,
            text='',
            bg=self._bg,
            fg=self._fg,
            anchor='w',
            justify='left',
            font=value_font,
            wraplength=self._min_width - 150,
        )
        self._watch_for_value.pack(side='left', fill='x', expand=True, padx=(6, 0), anchor='nw')

        self._current_selection_row = tk.Frame(body, bg=self._bg)
        self._current_selection_row.pack(fill='x', pady=(8, 0), anchor='w')
        self._current_selection_label = tk.Label(
            self._current_selection_row,
            text='Current selection:',
            bg=self._bg,
            fg=self._accent,
            anchor='w',
            justify='left',
            font=heading_font,
        )
        self._current_selection_label.pack(side='left', anchor='nw')
        self._current_selection_value = tk.Label(
            self._current_selection_row,
            text='',
            bg=self._bg,
            fg=self._fg,
            anchor='w',
            justify='left',
            font=value_font,
            wraplength=self._min_width - 180,
        )
        self._current_selection_value.pack(side='left', fill='x', expand=True, padx=(6, 0), anchor='nw')

        self._selection_detail_row = tk.Frame(body, bg=self._bg)
        self._selection_detail_row.pack(fill='x', pady=(2, 0), anchor='w')
        self._selection_detail_label = tk.Label(
            self._selection_detail_row,
            text='Selection detail:',
            bg=self._bg,
            fg=self._muted_fg,
            anchor='w',
            justify='left',
            font=heading_font,
        )
        self._selection_detail_label.pack(side='left', anchor='nw')
        self._selection_detail_value = tk.Label(
            self._selection_detail_row,
            text='',
            bg=self._bg,
            fg=self._fg,
            anchor='w',
            justify='left',
            font=value_font,
            wraplength=self._min_width - 180,
        )
        self._selection_detail_value.pack(side='left', fill='x', expand=True, padx=(6, 0), anchor='nw')

        self._root_window.bind('<Configure>', self._on_root_configure, add='+')
        self._parent_widget.bind('<Configure>', self._on_parent_configure, add='+')
        if self._reference_widget is not None:
            self._reference_widget.bind('<Configure>', self._on_parent_configure, add='+')
        self._root_window.bind('<Destroy>', self._on_destroy, add='+')

    def _estimate_height(self) -> int:
        return max(self._min_height, 148 if self._line_count >= 6 else 132)

    def _cancel_pending(self) -> None:
        if self._pending_after_id is not None:
            try:
                self._root_window.after_cancel(self._pending_after_id)
            except Exception:
                pass
        self._pending_after_id = None
        self._pending_group = None

    def _cancel_hide(self) -> None:
        if self._hide_after_id is not None:
            try:
                self._root_window.after_cancel(self._hide_after_id)
            except Exception:
                pass
        self._hide_after_id = None

    def _resolve_group_id(self, key: str, group: Optional[str]) -> str:
        return str(group or key)

    def _get_current_value(self, variable: Optional[tk.Variable]) -> Any:
        if variable is None:
            return None
        try:
            return variable.get()
        except Exception:
            return None

    def _build_group_payload(self, group_id: str) -> Dict[str, str]:
        meta = self._bindings.get(group_id, {})
        key = meta.get('key', group_id)
        label = meta.get('label', '')
        current_value = self._get_current_value(meta.get('variable'))
        detail = get_policy_detail(key, label)

        if current_value is True:
            current_text = 'Enabled'
        elif current_value is False:
            current_text = 'Disabled'
        elif current_value is None:
            current_text = ''
        else:
            current_text = str(current_value)

        option_text = get_option_detail(key, current_value) if current_value is not None else ''

        return {
            'summary': detail.get('summary', 'No description available.'),
            'impact': detail.get('impact', 'No additional impact notes available.'),
            'consider': detail.get('consider', 'No caution notes available.'),
            'current_selection': current_text,
            'selection_detail': option_text,
        }

    def _apply_payload(self, payload: Dict[str, str]) -> None:
        self._summary_value.configure(text=payload.get('summary', ''))
        self._use_case_value.configure(text=payload.get('impact', ''))
        self._watch_for_value.configure(text=payload.get('consider', ''))

        current_selection = payload.get('current_selection', '').strip()
        selection_detail = payload.get('selection_detail', '').strip()

        self._current_selection_value.configure(text=current_selection)
        self._selection_detail_value.configure(text=selection_detail)

        if current_selection:
            if not self._current_selection_row.winfo_manager():
                self._current_selection_row.pack(fill='x', pady=(8, 0), anchor='w')
        else:
            self._current_selection_row.pack_forget()

        if selection_detail:
            if not self._selection_detail_row.winfo_manager():
                self._selection_detail_row.pack(fill='x', pady=(2, 0), anchor='w')
        else:
            self._selection_detail_row.pack_forget()

        self._update_wraplengths()

    def _update_wraplengths(self) -> None:
        try:
            popup_width = max(self._popup.winfo_width(), self._min_width)
        except Exception:
            popup_width = self._min_width

        full_wrap = max(popup_width - 32, 360)
        inline_wrap = max(full_wrap - 120, 220)
        detail_wrap = max(full_wrap - 140, 220)

        self._summary_value.configure(wraplength=full_wrap)
        self._use_case_value.configure(wraplength=inline_wrap)
        self._watch_for_value.configure(wraplength=inline_wrap)
        self._current_selection_value.configure(wraplength=detail_wrap)
        self._selection_detail_value.configure(wraplength=detail_wrap)

    def _position_popup(self) -> None:
        position_detail_popup_window(
            self._popup,
            self._parent_widget,
            reference_widget=self._reference_widget,
            left_gap=self._left_gap,
            top_offset=self._top_offset,
            right_margin=self._right_margin,
            min_width=self._min_width,
            min_height=self._estimate_height(),
        )
        self._popup.update_idletasks()
        self._update_wraplengths()

    def _show_popup(self, group_id: str) -> None:
        self._cancel_pending()
        self._cancel_hide()
        self._apply_payload(self._build_group_payload(group_id))
        self._popup.deiconify()
        try:
            self._popup.lift(self._root_window)
        except Exception:
            self._popup.lift()
        try:
            self._popup.attributes('-topmost', True)
            self._popup.after(50, lambda: self._popup.attributes('-topmost', False))
        except Exception:
            pass
        self._position_popup()
        self._root_window.after_idle(self._position_popup)
        self._popup_visible = True
        self._active_group = group_id
        self._ignore_leave_until = time.monotonic() + 0.15

    def _hide_popup(self) -> None:
        self._cancel_pending()
        self._cancel_hide()
        try:
            self._popup.withdraw()
        except Exception:
            pass
        self._popup_visible = False
        self._active_group = None
        self._summary_value.configure(text='')
        self._use_case_value.configure(text='')
        self._watch_for_value.configure(text='')
        self._current_selection_value.configure(text='')
        self._selection_detail_value.configure(text='')
        self._current_selection_row.pack_forget()
        self._selection_detail_row.pack_forget()

    def clear(self) -> None:
        self._hide_popup()

    def destroy(self) -> None:
        self._on_destroy()

    def _schedule_show(self, group_id: str) -> None:
        self._cancel_pending()
        self._pending_group = group_id
        self._pending_after_id = self._root_window.after(
            self._delay_ms,
            lambda gid=group_id: self._show_if_still_hovered(gid),
        )

    def _show_if_still_hovered(self, group_id: str) -> None:
        self._pending_after_id = None
        self._pending_group = None
        if self._hover_counts.get(group_id, 0) <= 0:
            return
        self._show_popup(group_id)

    def _widget_is_descendant_of(self, widget: tk.Widget, ancestor: tk.Widget) -> bool:
        current = widget
        while current is not None:
            if current == ancestor:
                return True
            try:
                parent_name = current.winfo_parent()
                if not parent_name:
                    return False
                current = current.nametowidget(parent_name)
            except Exception:
                return False
        return False

    def _pointer_is_over_group(self, group_id: str) -> bool:
        try:
            x, y = self._root_window.winfo_pointerxy()
            widget = self._root_window.winfo_containing(x, y)
        except Exception:
            widget = None

        if widget is None and self._popup_visible:
            try:
                widget = self._popup.winfo_containing(x, y)
            except Exception:
                widget = None

        if widget is None:
            return False

        for target in self._group_widgets.get(group_id, []):
            if self._widget_is_descendant_of(widget, target):
                return True

        return self._popup_visible and self._widget_is_descendant_of(widget, self._popup)

    def _increment_hover(self, group_id: str) -> None:
        self._hover_counts[group_id] = self._hover_counts.get(group_id, 0) + 1
        self._cancel_hide()
        if self._active_group == group_id and self._popup_visible:
            return
        if self._active_group and self._active_group != group_id:
            self._hide_popup()
        self._schedule_show(group_id)

    def _decrement_hover(self, group_id: str) -> None:
        current = self._hover_counts.get(group_id, 0)
        self._hover_counts[group_id] = max(0, current - 1)
        self._cancel_hide()
        self._hide_after_id = self._root_window.after(80, lambda gid=group_id: self._confirm_hide(gid))

    def _confirm_hide(self, group_id: str) -> None:
        self._hide_after_id = None
        if time.monotonic() < self._ignore_leave_until:
            self._hide_after_id = self._root_window.after(80, lambda gid=group_id: self._confirm_hide(gid))
            return
        if self._hover_counts.get(group_id, 0) > 0:
            return
        if self._pointer_is_over_group(group_id):
            return
        if self._pending_group == group_id:
            self._cancel_pending()
        if self._active_group == group_id:
            self._hide_popup()

    def _refresh_active_group(self) -> None:
        if not self._popup_visible or not self._active_group:
            return
        self._apply_payload(self._build_group_payload(self._active_group))
        self._position_popup()

    def _trace_variable(self, group_id: str) -> None:
        if self._active_group == group_id and self._popup_visible:
            self._root_window.after_idle(self._refresh_active_group)

    def _on_root_configure(self, _event: Optional[tk.Event] = None) -> None:
        if self._popup_visible and self._active_group:
            self._root_window.after_idle(self._refresh_active_group)

    def _on_parent_configure(self, _event: Optional[tk.Event] = None) -> None:
        if self._popup_visible and self._active_group:
            self._root_window.after_idle(self._refresh_active_group)

    def _on_destroy(self, _event: Optional[tk.Event] = None) -> None:
        try:
            self._cancel_pending()
            self._cancel_hide()
            self._popup.destroy()
        except Exception:
            pass

    def bind_control(
        self,
        control: tk.Widget,
        *,
        key: str,
        label: str,
        variable: Optional[tk.Variable] = None,
        group: Optional[str] = None,
    ) -> None:
        group_id = self._resolve_group_id(key, group)
        self._bindings[group_id] = {'key': key, 'label': label, 'variable': variable}
        self._group_widgets.setdefault(group_id, []).append(control)
        self._hover_counts.setdefault(group_id, 0)

        def on_enter(_event: Optional[tk.Event] = None, gid: str = group_id) -> None:
            self._increment_hover(gid)

        def on_leave(_event: Optional[tk.Event] = None, gid: str = group_id) -> None:
            self._decrement_hover(gid)

        control.bind('<Enter>', on_enter, add='+')
        control.bind('<Leave>', on_leave, add='+')
        control.bind('<FocusIn>', on_enter, add='+')
        control.bind('<FocusOut>', on_leave, add='+')

        if variable is not None:
            def traced(*_args: Any, gid: str = group_id) -> None:
                self._trace_variable(gid)

            try:
                variable.trace_add('write', traced)
            except Exception:
                pass
