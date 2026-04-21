"""Helpers for positioning the policy details popup window within the main GUI."""

from __future__ import annotations

import tkinter as tk


def position_detail_popup_window(
    popup: tk.Toplevel,
    anchor_widget: tk.Widget,
    *,
    reference_widget: tk.Widget | None = None,
    left_gap: int = 18,
    top_offset: int = 0,
    right_margin: int = 0,
    padding: int = 0,
    min_width: int = 420,
    min_height: int = 96,
) -> None:
    """
    Position the hover-details popup inside the top action row without reserving
    any layout space.

    When a reference widget is provided, the popup starts just to the right of it
    and uses the remaining width inside the anchor widget. Otherwise it falls back
    to the anchor widget bounds.
    """
    anchor_widget.update_idletasks()
    popup.update_idletasks()

    anchor_x = anchor_widget.winfo_rootx()
    anchor_y = anchor_widget.winfo_rooty()
    anchor_width = max(anchor_widget.winfo_width(), min_width)
    anchor_height = max(anchor_widget.winfo_height(), min_height)

    req_width = max(popup.winfo_reqwidth(), min_width)
    req_height = max(popup.winfo_reqheight(), min_height)

    x = anchor_x + padding
    y = anchor_y + top_offset + padding
    width = max(anchor_width - (padding * 2), req_width)
    height = req_height

    if reference_widget is not None:
        reference_widget.update_idletasks()
        x = reference_widget.winfo_rootx() + reference_widget.winfo_width() + left_gap
        max_right = anchor_x + anchor_width - right_margin - padding
        width = max_right - x

    width = max(width, min_width)
    height = max(height, min_height)
    popup.geometry(f"{width}x{height}+{x}+{y}")
