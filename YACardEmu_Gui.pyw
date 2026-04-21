import json
import os
import re
import shutil
import socket
import subprocess
import sys
import threading
import time
import configparser
import importlib.util
import ctypes
from ctypes import wintypes
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Tuple
from collections import deque

import requests
import tkinter as tk
from tkinter import ttk, messagebox
from PIL import Image, ImageTk


PYGAME_AVAILABLE = importlib.util.find_spec("pygame") is not None
if PYGAME_AVAILABLE:
    import pygame
else:
    pygame = None


HOST = "127.0.0.1"
EXE_NAME = "YACardEmu.exe"
CONFIG_NAME = "config.ini"
ICON_NAME = "yacardemu.ico"
LINKS_NAME = "card_image_links.json"
INPUT_BINDING_NAME = "input_binding.json"
WINDOW_STATE_NAME = "window_state.json"
TEMPLATE_DIR_NAME = "card_images"

WINDOW_TITLE = "YACardEmu Card Select"
WINDOW_SIZE = "560x660"
MIN_WIDTH = 500
MIN_HEIGHT = 740

PREVIEW_MAX_W = 320
PREVIEW_MAX_H = 480

THUMB_W = 80
THUMB_H = 116
THUMB_COLUMNS = 3

DEFAULT_API_PORT = 8080
DEFAULT_BASEPATH = "./cards/"

SERVER_WAIT_SECONDS = 25.0
SERVER_POLL_INTERVAL = 0.25
HTTP_TIMEOUT = 5.0

SELECTION_WAIT_SECONDS = 2.0
SELECTION_POLL_INTERVAL = 0.10
INSERT_RETRY_DELAY = 0.35

NEW_CARD_POLL_MS = 1500
SELECTED_CARD_POLL_MS = 1000
PREVIEW_REFRESH_DEBOUNCE_MS = 400
READING_OVERLAY_DURATION_MS = 4000
READING_OVERLAY_FLASH_MS = 350
PREVIEW_HOLD_MAX_MS = 30000
MAX_STATUS_LINES = 2
DATA_DIR_NAME = "data"

INPUT_TRIGGER_DEBOUNCE_MS = 700
CONTROLLER_POLL_MS = 30
AXIS_BIND_THRESHOLD = 0.75
TRIGGER_BIND_THRESHOLD = 0.60

XINPUT_GAMEPAD_DPAD_UP = 0x0001
XINPUT_GAMEPAD_DPAD_DOWN = 0x0002
XINPUT_GAMEPAD_DPAD_LEFT = 0x0004
XINPUT_GAMEPAD_DPAD_RIGHT = 0x0008
XINPUT_GAMEPAD_START = 0x0010
XINPUT_GAMEPAD_BACK = 0x0020
XINPUT_GAMEPAD_LEFT_THUMB = 0x0040
XINPUT_GAMEPAD_RIGHT_THUMB = 0x0080
XINPUT_GAMEPAD_LEFT_SHOULDER = 0x0100
XINPUT_GAMEPAD_RIGHT_SHOULDER = 0x0200
XINPUT_GAMEPAD_A = 0x1000
XINPUT_GAMEPAD_B = 0x2000
XINPUT_GAMEPAD_X = 0x4000
XINPUT_GAMEPAD_Y = 0x8000
XINPUT_TRIGGER_THRESHOLD = 180
XINPUT_STICK_THRESHOLD = 24000

XINPUT_BUTTONS = [
    (XINPUT_GAMEPAD_A, "A"),
    (XINPUT_GAMEPAD_B, "B"),
    (XINPUT_GAMEPAD_X, "X"),
    (XINPUT_GAMEPAD_Y, "Y"),
    (XINPUT_GAMEPAD_LEFT_SHOULDER, "LB"),
    (XINPUT_GAMEPAD_RIGHT_SHOULDER, "RB"),
    (XINPUT_GAMEPAD_BACK, "Back"),
    (XINPUT_GAMEPAD_START, "Start"),
    (XINPUT_GAMEPAD_LEFT_THUMB, "L3"),
    (XINPUT_GAMEPAD_RIGHT_THUMB, "R3"),
    (XINPUT_GAMEPAD_DPAD_UP, "D-pad Up"),
    (XINPUT_GAMEPAD_DPAD_DOWN, "D-pad Down"),
    (XINPUT_GAMEPAD_DPAD_LEFT, "D-pad Left"),
    (XINPUT_GAMEPAD_DPAD_RIGHT, "D-pad Right"),
]


class XINPUT_GAMEPAD(ctypes.Structure):
    _fields_ = [
        ("wButtons", wintypes.WORD),
        ("bLeftTrigger", ctypes.c_ubyte),
        ("bRightTrigger", ctypes.c_ubyte),
        ("sThumbLX", ctypes.c_short),
        ("sThumbLY", ctypes.c_short),
        ("sThumbRX", ctypes.c_short),
        ("sThumbRY", ctypes.c_short),
    ]


class XINPUT_STATE(ctypes.Structure):
    _fields_ = [
        ("dwPacketNumber", wintypes.DWORD),
        ("Gamepad", XINPUT_GAMEPAD),
    ]


@dataclass
class CardEntry:
    bin_path: Path
    png_path: Optional[Path]
    name: str


class TemplateGrid(ttk.Frame):
    def __init__(self, parent, app: "App", on_select=None, thumb_w: int = THUMB_W, thumb_h: int = THUMB_H):
        super().__init__(parent)
        self.app = app
        self.on_select = on_select
        self.thumb_w = thumb_w
        self.thumb_h = thumb_h

        self.selected_template: Optional[str] = None
        self.thumb_refs: List[ImageTk.PhotoImage] = []
        self.buttons: Dict[str, tk.Widget] = {}

        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        self.canvas = tk.Canvas(self, highlightthickness=0)
        self.canvas.grid(row=0, column=0, sticky="nsew")

        self.scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.scrollbar.grid(row=0, column=1, sticky="ns")

        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.inner = ttk.Frame(self.canvas)
        self.inner_window = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")

        self.inner.bind("<Configure>", self._on_inner_configure)
        self.canvas.bind("<Configure>", self._on_canvas_configure)

        self.selected_var = tk.StringVar(value="No template selected")
        ttk.Label(self, textvariable=self.selected_var).grid(row=1, column=0, columnspan=2, sticky="w", pady=(6, 0))

    def _on_inner_configure(self, event=None) -> None:
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _on_canvas_configure(self, event=None) -> None:
        self.canvas.itemconfigure(self.inner_window, width=self.canvas.winfo_width())

    def refresh(self) -> None:
        for child in self.inner.winfo_children():
            child.destroy()

        self.thumb_refs.clear()
        self.buttons.clear()

        names = self.app.get_template_names()
        current_selected = self.selected_template if self.selected_template in names else None
        self.selected_template = current_selected

        if not names:
            ttk.Label(self.inner, text="No template PNGs found in card_images").grid(row=0, column=0, sticky="w")
            self.selected_var.set("No template selected")
            return

        row = 0
        col = 0

        for name in names:
            card = ttk.Frame(self.inner, padding=4)
            card.grid(row=row, column=col, padx=4, pady=4, sticky="n")

            preview_box = ttk.Frame(card, relief="solid", borderwidth=1)
            preview_box.grid(row=0, column=0, sticky="n")
            preview_box.grid_propagate(False)
            preview_box.configure(width=self.thumb_w + 8, height=self.thumb_h + 8)

            label = ttk.Label(preview_box, anchor="center", text="Preview")
            label.place(relx=0.5, rely=0.5, anchor="center")

            template_path = self.app.template_dir / name
            try:
                img = Image.open(template_path)
                img.thumbnail((self.thumb_w, self.thumb_h), Image.LANCZOS)
                photo = ImageTk.PhotoImage(img)
                self.thumb_refs.append(photo)
                label.configure(image=photo, text="")
            except Exception:
                label.configure(text="Preview failed")

            btn = ttk.Button(card, text=name, command=lambda n=name: self.select_template(n), width=18)
            btn.grid(row=1, column=0, pady=(4, 0), sticky="ew")

            self.buttons[name] = btn

            if col >= THUMB_COLUMNS - 1:
                row += 1
                col = 0
            else:
                col += 1

        if current_selected:
            self.select_template(current_selected, invoke=False)
        else:
            self.selected_var.set("No template selected")

    def select_template(self, name: str, invoke: bool = True) -> None:
        self.selected_template = name
        self.selected_var.set(f"Selected template: {name}")

        for template_name, btn in self.buttons.items():
            if template_name == name:
                btn.state(["pressed"])
            else:
                btn.state(["!pressed"])

        if invoke and self.on_select:
            self.on_select(name)

    def get_selected(self) -> Optional[str]:
        return self.selected_template

    def set_selected(self, name: Optional[str]) -> None:
        if name and name in self.buttons:
            self.select_template(name)
        else:
            self.selected_template = None
            self.selected_var.set("No template selected")


class AutoSetupDialog(tk.Toplevel):
    def __init__(self, app: "App", card_name: str, signature: str):
        super().__init__(app.root)
        self.app = app
        self.card_name = card_name
        self.signature = signature
        self.result_completed = False

        self.title("New Card Detected")
        self.geometry("560x560")
        self.minsize(520, 520)
        self.transient(app.root)
        self.grab_set()

        self._apply_icon()

        self.name_var = tk.StringVar(value="newcard")
        self.info_var = tk.StringVar(
            value=f"A new card file was detected: {card_name}.\nRename it and choose a clean template image."
        )

        self._build_ui()

    def _apply_icon(self) -> None:
        try:
            if self.app.icon_path.exists():
                self.iconbitmap(str(self.app.icon_path))
        except Exception:
            pass

    def _build_ui(self) -> None:
        self.columnconfigure(0, weight=1)
        self.rowconfigure(2, weight=1)

        top = ttk.Frame(self, padding=10)
        top.grid(row=0, column=0, sticky="ew")
        top.columnconfigure(1, weight=1)

        ttk.Label(top, textvariable=self.info_var, wraplength=500, justify="left").grid(
            row=0, column=0, columnspan=2, sticky="w", pady=(0, 10)
        )

        ttk.Label(top, text="New Name:").grid(row=1, column=0, sticky="w", padx=(0, 8))
        self.name_entry = ttk.Entry(top, textvariable=self.name_var)
        self.name_entry.grid(row=1, column=1, sticky="ew")

        middle = ttk.LabelFrame(self, text="Choose Template Image", padding=8)
        middle.grid(row=2, column=0, sticky="nsew", padx=10, pady=(0, 10))
        middle.columnconfigure(0, weight=1)
        middle.rowconfigure(0, weight=1)

        self.template_grid = TemplateGrid(middle, self.app)
        self.template_grid.grid(row=0, column=0, sticky="nsew")
        self.template_grid.refresh()

        bottom = ttk.Frame(self, padding=(10, 0, 10, 10))
        bottom.grid(row=3, column=0, sticky="ew")
        bottom.columnconfigure(0, weight=1)

        buttons = ttk.Frame(bottom)
        buttons.grid(row=0, column=0)

        ttk.Button(buttons, text="Save", command=self.save, width=16).grid(row=0, column=0, padx=(0, 6))
        ttk.Button(buttons, text="Skip For Now", command=self.skip, width=16).grid(row=0, column=1, padx=(6, 0))

        self.name_entry.focus_set()

    def save(self) -> None:
        template_name = self.template_grid.get_selected()
        if not template_name:
            messagebox.showerror("No template selected", "Choose a template image first.", parent=self)
            return

        new_base = self.name_var.get().strip()
        success, message, new_card_name = self.app.rename_card_entry(self.card_name, new_base)
        if not success:
            messagebox.showerror("Rename failed", message, parent=self)
            return

        self.app.card_links[new_card_name] = template_name
        self.app.save_card_links()

        self.app.load_cards_from_folder(self.app.cards_dir)

        reset_ok, reset_msg = self.app.restore_template_to_card(new_card_name)
        if not reset_ok:
            messagebox.showerror("Template reset failed", reset_msg, parent=self)
            return

        self.app.load_cards_from_folder(self.app.cards_dir)
        self.app.show_current_card_if_name(new_card_name)
        self.app.last_handled_new_card_signature = self.signature
        self.result_completed = True
        self.destroy()

    def skip(self) -> None:
        self.app.last_handled_new_card_signature = self.signature
        self.destroy()


class CardManagerWindow(tk.Toplevel):
    def __init__(self, app: "App"):
        super().__init__(app.root)
        self.app = app
        self.title("Manage Card Links")
        self.geometry("980x620")
        self.minsize(900, 560)
        self.transient(app.root)

        self.selected_card_name: Optional[str] = None

        self._apply_icon()
        self._build_ui()
        self.refresh_lists()

    def _apply_icon(self) -> None:
        try:
            if self.app.icon_path.exists():
                self.iconbitmap(str(self.app.icon_path))
        except Exception:
            pass

    def _build_ui(self) -> None:
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        main = ttk.Frame(self, padding=10)
        main.grid(row=0, column=0, sticky="nsew")
        main.columnconfigure(0, weight=0)
        main.columnconfigure(1, weight=0)
        main.columnconfigure(2, weight=1)
        main.rowconfigure(0, weight=1)

        left = ttk.LabelFrame(main, text="Live Cards", padding=8)
        left.grid(row=0, column=0, sticky="ns", padx=(0, 8))
        left.columnconfigure(0, weight=1)
        left.rowconfigure(1, weight=1)

        ttk.Label(left, text="Cards from YACardEmu's configured cards folder").grid(
            row=0, column=0, sticky="w", pady=(0, 6)
        )

        self.cards_listbox = tk.Listbox(left, exportselection=False, width=28)
        self.cards_listbox.grid(row=1, column=0, sticky="ns")
        self.cards_listbox.bind("<<ListboxSelect>>", lambda e: self.on_card_select())

        middle = ttk.Frame(main, padding=(4, 0, 4, 0))
        middle.grid(row=0, column=1, sticky="ns")

        ttk.Button(middle, text="Refresh", command=self.refresh_lists, width=18).grid(row=0, column=0, pady=(0, 8))
        ttk.Button(middle, text="Rename Card", command=self.rename_card, width=18).grid(row=1, column=0, pady=4)
        ttk.Button(middle, text="Link Template", command=self.link_template, width=18).grid(row=2, column=0, pady=4)
        ttk.Button(middle, text="Unlink Template", command=self.unlink_template, width=18).grid(row=3, column=0, pady=4)
        ttk.Button(middle, text="Reset PNG From Template", command=self.reset_png_from_template, width=18).grid(row=4, column=0, pady=4)
        ttk.Button(middle, text="Auto-Fix card.bin", command=self.auto_fix_new_card, width=18).grid(row=5, column=0, pady=4)
        ttk.Button(middle, text="Open Template Folder", command=self.open_template_folder, width=18).grid(
            row=6, column=0, pady=(12, 4)
        )
        ttk.Button(middle, text="Close", command=self.destroy, width=18).grid(row=7, column=0, pady=(30, 0))

        right = ttk.LabelFrame(main, text="Clean Template Images", padding=8)
        right.grid(row=0, column=2, sticky="nsew", padx=(8, 0))
        right.columnconfigure(0, weight=1)
        right.rowconfigure(1, weight=1)

        ttk.Label(
            right,
            text="Put clean text-free PNG templates in the card_images folder. Filenames can be anything."
        ).grid(row=0, column=0, sticky="w", pady=(0, 6))

        self.template_grid = TemplateGrid(right, self.app)
        self.template_grid.grid(row=1, column=0, sticky="nsew")

        self.link_info_var = tk.StringVar(value="")
        ttk.Label(right, textvariable=self.link_info_var, wraplength=520, justify="left").grid(
            row=2, column=0, sticky="w", pady=(8, 0)
        )

    def refresh_lists(self) -> None:
        self.cards_listbox.delete(0, tk.END)
        for card in self.app.cards:
            self.cards_listbox.insert(tk.END, card.name)

        self.template_grid.refresh()
        self.selected_card_name = None
        self.link_info_var.set("")

    def on_card_select(self) -> None:
        selection = self.cards_listbox.curselection()
        if not selection:
            self.selected_card_name = None
            self.link_info_var.set("")
            return

        self.selected_card_name = self.cards_listbox.get(selection[0])
        linked = self.app.card_links.get(self.selected_card_name, "")
        if linked:
            self.link_info_var.set(f"Linked template for {self.selected_card_name}: {linked}")
            self.template_grid.set_selected(linked)
        else:
            self.link_info_var.set(f"No template linked for {self.selected_card_name}")

    def get_selected_card(self) -> Optional[str]:
        return self.selected_card_name

    def rename_card(self) -> None:
        card_name = self.get_selected_card()
        if not card_name:
            messagebox.showerror("No card selected", "Select a live card first.", parent=self)
            return

        dialog = tk.Toplevel(self)
        dialog.title("Rename Card")
        dialog.transient(self)
        dialog.grab_set()
        try:
            if self.app.icon_path.exists():
                dialog.iconbitmap(str(self.app.icon_path))
        except Exception:
            pass

        dialog.columnconfigure(1, weight=1)

        ttk.Label(dialog, text="New Name:").grid(row=0, column=0, padx=10, pady=10, sticky="w")
        var = tk.StringVar(value=Path(card_name).stem)
        entry = ttk.Entry(dialog, textvariable=var)
        entry.grid(row=0, column=1, padx=(0, 10), pady=10, sticky="ew")
        entry.focus_set()

        def do_rename():
            success, message, _new_name = self.app.rename_card_entry(card_name, var.get().strip())
            if not success:
                messagebox.showerror("Rename failed", message, parent=dialog)
                return
            self.app.load_cards_from_folder(self.app.cards_dir)
            self.refresh_lists()
            dialog.destroy()

        ttk.Button(dialog, text="Save", command=do_rename).grid(row=1, column=0, padx=10, pady=(0, 10))
        ttk.Button(dialog, text="Cancel", command=dialog.destroy).grid(row=1, column=1, padx=10, pady=(0, 10), sticky="e")

    def link_template(self) -> None:
        card_name = self.get_selected_card()
        template_name = self.template_grid.get_selected()

        if not card_name:
            messagebox.showerror("No card selected", "Select a live card first.", parent=self)
            return

        if not template_name:
            messagebox.showerror("No template selected", "Select a clean template image first.", parent=self)
            return

        self.app.card_links[card_name] = template_name
        self.app.save_card_links()
        self.link_info_var.set(f"Linked template for {card_name}: {template_name}")
        messagebox.showinfo("Linked", f"{card_name} is now linked to {template_name}", parent=self)

    def unlink_template(self) -> None:
        card_name = self.get_selected_card()
        if not card_name:
            messagebox.showerror("No card selected", "Select a live card first.", parent=self)
            return

        if card_name in self.app.card_links:
            del self.app.card_links[card_name]
            self.app.save_card_links()

        self.link_info_var.set(f"No template linked for {card_name}")
        messagebox.showinfo("Unlinked", f"Removed template link from {card_name}", parent=self)

    def reset_png_from_template(self) -> None:
        card_name = self.get_selected_card()
        if not card_name:
            messagebox.showerror("No card selected", "Select a live card first.", parent=self)
            return

        success, message = self.app.restore_template_to_card(card_name)
        if not success:
            messagebox.showerror("Reset failed", message, parent=self)
            return

        self.app.load_cards_from_folder(self.app.cards_dir)
        self.app.show_current_card_if_name(card_name)
        messagebox.showinfo("Reset complete", message, parent=self)

    def auto_fix_new_card(self) -> None:
        if "card.bin" not in [c.name for c in self.app.cards]:
            messagebox.showerror("Not found", "card.bin was not found in the live cards folder.", parent=self)
            return

        dialog = AutoSetupDialog(self.app, "card.bin", signature="manual")
        self.wait_window(dialog)
        self.refresh_lists()

    def open_template_folder(self) -> None:
        self.app.ensure_template_dir()
        self.app.open_path(self.app.template_dir)


class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(WINDOW_TITLE)
        self.root.geometry(WINDOW_SIZE)
        self.root.minsize(MIN_WIDTH, MIN_HEIGHT)

        self.base_dir = self.get_app_dir()
        self.data_dir = self.base_dir / DATA_DIR_NAME
        self.exe_path = self.base_dir / EXE_NAME
        self.config_path = self.base_dir / CONFIG_NAME
        self.icon_path = self.resolve_data_path(ICON_NAME, prefer_data=False)
        self.links_path = self.resolve_data_path(LINKS_NAME, prefer_data=True)
        self.input_binding_path = self.resolve_data_path(INPUT_BINDING_NAME, prefer_data=True)
        self.window_state_path = self.resolve_data_path(WINDOW_STATE_NAME, prefer_data=True)
        self.template_dir = self.resolve_data_path(TEMPLATE_DIR_NAME, prefer_data=True, is_dir=True)

        self.migrate_legacy_data_files()

        self.api_port = DEFAULT_API_PORT
        self.cards_dir = self.base_dir / "cards"

        self.base_url = ""
        self.api_inserted_card = ""
        self.api_cards = ""

        self.session = requests.Session()
        self.process: Optional[subprocess.Popen] = None

        self.cards: List[CardEntry] = []
        self.current_index = 0
        self.current_photo: Optional[ImageTk.PhotoImage] = None
        self.current_card_signature: Optional[str] = None
        self.pending_card_refresh_signature: Optional[str] = None
        self.preview_refresh_job: Optional[str] = None

        self.card_links: Dict[str, str] = self.load_card_links()
        self.input_binding: Optional[Dict[str, object]] = self.load_input_binding()
        self.manager_window: Optional[CardManagerWindow] = None

        self.selected_name_var = tk.StringVar(value="")
        self.auto_setup_open = False
        self.last_handled_new_card_signature: Optional[str] = None
        self.status_lines: deque[str] = deque(maxlen=MAX_STATUS_LINES)
        self.status_vars = [tk.StringVar(value="") for _ in range(MAX_STATUS_LINES)]
        self.status_reader_thread: Optional[threading.Thread] = None
        self.reading_overlay_job: Optional[str] = None
        self.reading_flash_job: Optional[str] = None
        self.reading_overlay_visible = False
        self.reading_overlay_active = False
        self.preview_hold_active = False
        self.preview_hold_card_name: Optional[str] = None
        self.preview_hold_job: Optional[str] = None
        self.preview_hold_ignored_change = False
        self.binding_var = tk.StringVar(value="")
        self.binding_listen_var = tk.StringVar(value="")
        self.binding_dialog: Optional[tk.Toplevel] = None
        self.binding_capture_active = False
        self.last_input_trigger_ts = 0.0
        self.insert_in_progress = False
        self.pygame_ready = False
        self.joysticks: Dict[int, object] = {}
        self.xinput = None
        self.xinput_connected: Dict[int, bool] = {}
        self.xinput_prev_state: Dict[int, Dict[str, object]] = {}

        self._apply_window_icon()
        self.apply_saved_window_state()
        self._build_ui()
        self.update_binding_label()
        self.init_controller_input()
        self._set_controls_enabled(False)

        self.root.after(100, self.startup_sequence)

    def get_app_dir(self) -> Path:
        if getattr(sys, "frozen", False):
            return Path(sys.executable).resolve().parent
        return Path(__file__).resolve().parent

    def _apply_window_icon(self) -> None:
        try:
            if self.icon_path.exists():
                self.root.iconbitmap(str(self.icon_path))
        except Exception:
            pass

    def _build_ui(self) -> None:
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        main = ttk.Frame(self.root, padding=10)
        main.grid(row=0, column=0, sticky="nsew")
        main.columnconfigure(0, weight=1)

        viewer_row = ttk.Frame(main)
        viewer_row.grid(row=0, column=0, sticky="n")
        viewer_row.columnconfigure(1, weight=0)
        viewer_row.rowconfigure(0, weight=0)

        self.prev_btn = ttk.Button(
            viewer_row,
            text="◀ Previous",
            command=self.prev_card,
            width=12
        )
        self.prev_btn.grid(row=0, column=0, sticky="ns", padx=(0, 10))

        center = ttk.Frame(viewer_row)
        center.grid(row=0, column=1, sticky="n")
        center.columnconfigure(0, weight=1)

        self.image_frame = ttk.Frame(center, relief="solid", borderwidth=1)
        self.image_frame.grid(row=0, column=0, sticky="n")
        self.image_frame.grid_propagate(False)
        self.image_frame.configure(width=PREVIEW_MAX_W + 12, height=PREVIEW_MAX_H + 12)

        self.image_label = ttk.Label(
            self.image_frame,
            anchor="center",
            text="Starting YACardEmu..."
        )
        self.image_label.place(relx=0.5, rely=0.5, anchor="center")

        info = ttk.Frame(center, padding=(0, 6, 0, 0))
        info.grid(row=1, column=0, sticky="ew")
        info.columnconfigure(1, weight=1)

        ttk.Label(info, text="Selected Card:").grid(row=0, column=0, sticky="w", padx=(0, 8))
        self.selected_entry = ttk.Entry(
            info,
            textvariable=self.selected_name_var,
            state="readonly",
            justify="center",
            width=28
        )
        self.selected_entry.grid(row=0, column=1, sticky="ew")

        self.next_btn = ttk.Button(
            viewer_row,
            text="Next ▶",
            command=self.next_card,
            width=12
        )
        self.next_btn.grid(row=0, column=2, sticky="ns", padx=(10, 0))

        bottom = ttk.Frame(main, padding=(0, 6, 0, 0))
        bottom.grid(row=1, column=0, sticky="ew")
        bottom.columnconfigure(0, weight=1)

        buttons = ttk.Frame(bottom)
        buttons.grid(row=0, column=0)

        self.insert_btn = ttk.Button(buttons, text="Insert Card", command=self.insert_current_card, width=16)
        self.insert_btn.grid(row=0, column=0, padx=(0, 6))

        self.manage_btn = ttk.Button(buttons, text="Manage Cards", command=self.open_manager, width=16)
        self.manage_btn.grid(row=0, column=1, padx=6)

        self.bind_btn = ttk.Button(buttons, text="Insert Keybind", command=self.open_bind_input_dialog, width=18)
        self.bind_btn.grid(row=0, column=2, padx=6)

        self.clear_bind_btn = ttk.Button(buttons, text="Clear Keybind", command=self.clear_input_binding, width=14)
        self.clear_bind_btn.grid(row=0, column=3, padx=(6, 0))

        ttk.Label(bottom, textvariable=self.binding_var, anchor="center", justify="center").grid(
            row=1, column=0, sticky="ew", pady=(8, 2)
        )

        overlay_frame = ttk.Frame(bottom, height=46)
        overlay_frame.grid(row=2, column=0, sticky="ew", pady=(10, 6))
        overlay_frame.columnconfigure(0, weight=1)
        overlay_frame.grid_propagate(False)

        self.reading_overlay_label = tk.Label(
            overlay_frame,
            text="READING CARD...",
            fg="#ff0000",
            bg=self.root.cget("bg"),
            font=("Segoe UI", 14, "bold")
        )
        self.reading_overlay_label.grid(row=0, column=0)
        self.reading_overlay_label.grid_remove()

        status_frame = ttk.Frame(bottom)
        status_frame.grid(row=3, column=0, sticky="ew", pady=(6, 0))
        status_frame.columnconfigure(0, weight=1)

        ttk.Separator(status_frame, orient="horizontal").grid(row=0, column=0, sticky="ew", pady=(0, 6))
        ttk.Label(status_frame, text="CardEmu Status").grid(row=1, column=0, sticky="w")

        self.status_line_labels: List[ttk.Label] = []
        for idx, var in enumerate(self.status_vars, start=2):
            lbl = ttk.Label(status_frame, textvariable=var, anchor="w", justify="left")
            lbl.grid(row=idx, column=0, sticky="ew")
            self.status_line_labels.append(lbl)

        self.root.bind("<Left>", lambda e: self.prev_card() if not self.binding_capture_active else "break")
        self.root.bind("<Right>", lambda e: self.next_card() if not self.binding_capture_active else "break")
        self.root.bind("<Return>", lambda e: self.insert_current_card() if not self.binding_capture_active else "break")
        self.root.bind_all("<KeyPress>", self.on_key_press, add="+")
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        self.root.bind_class("Toplevel", "<Map>", self._on_toplevel_map)

    def _on_toplevel_map(self, event) -> None:
        win = event.widget
        if not isinstance(win, tk.Toplevel):
            return
        if getattr(win, "_app_centered_once", False):
            return
        win._app_centered_once = True
        self.root.after(10, lambda w=win: self.center_child_window(w))

    def center_child_window(self, win: tk.Toplevel) -> None:
        try:
            if not win.winfo_exists():
                return

            win.update_idletasks()

            parent_x = self.root.winfo_rootx()
            parent_y = self.root.winfo_rooty()
            parent_w = self.root.winfo_width()
            parent_h = self.root.winfo_height()

            width = win.winfo_width()
            height = win.winfo_height()

            if width <= 1:
                width = win.winfo_reqwidth()
            if height <= 1:
                height = win.winfo_reqheight()

            pos_x = parent_x + max(0, (parent_w - width) // 2)
            pos_y = parent_y + max(0, (parent_h - height) // 2)

            win.geometry(f"{width}x{height}+{pos_x}+{pos_y}")
        except Exception:
            pass

    def _set_controls_enabled(self, enabled: bool) -> None:
        state = "normal" if enabled else "disabled"
        for widget in (self.prev_btn, self.next_btn, self.insert_btn, self.manage_btn, self.bind_btn, self.clear_bind_btn):
            widget.configure(state=state)

    def resolve_data_path(self, name: str, prefer_data: bool = True, is_dir: bool = False) -> Path:
        data_path = self.data_dir / name
        base_path = self.base_dir / name

        if prefer_data:
            return data_path

        if data_path.exists() or (is_dir and data_path.is_dir()):
            return data_path

        return base_path

    def migrate_legacy_data_files(self) -> None:
        self.data_dir.mkdir(parents=True, exist_ok=True)

        legacy_links = self.base_dir / LINKS_NAME
        if self.links_path == self.data_dir / LINKS_NAME and legacy_links.exists() and not self.links_path.exists():
            try:
                shutil.move(str(legacy_links), str(self.links_path))
            except Exception:
                pass

        legacy_template_dir = self.base_dir / TEMPLATE_DIR_NAME
        if self.template_dir == self.data_dir / TEMPLATE_DIR_NAME and legacy_template_dir.exists() and not self.template_dir.exists():
            try:
                shutil.move(str(legacy_template_dir), str(self.template_dir))
            except Exception:
                pass

    def load_card_links(self) -> Dict[str, str]:
        if not self.links_path.exists():
            return {}
        try:
            with self.links_path.open("r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, dict):
                clean: Dict[str, str] = {}
                for k, v in data.items():
                    if isinstance(k, str) and isinstance(v, str):
                        clean[k] = v
                return clean
        except Exception:
            pass
        return {}

    def save_card_links(self) -> None:
        with self.links_path.open("w", encoding="utf-8") as f:
            json.dump(self.card_links, f, indent=2, ensure_ascii=False)

    def load_input_binding(self) -> Optional[Dict[str, object]]:
        if not self.input_binding_path.exists():
            return None
        try:
            with self.input_binding_path.open("r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, dict) and isinstance(data.get("source"), str):
                return data
        except Exception:
            pass
        return None

    def save_input_binding(self) -> None:
        try:
            if self.input_binding:
                with self.input_binding_path.open("w", encoding="utf-8") as f:
                    json.dump(self.input_binding, f, indent=2, ensure_ascii=False)
            elif self.input_binding_path.exists():
                self.input_binding_path.unlink()
        except Exception:
            pass

    def apply_saved_window_state(self) -> None:
        state = self.load_window_state()
        if not state:
            return

        try:
            width = int(state.get("width", 0))
            height = int(state.get("height", 0))
            x = int(state.get("x", 0))
            y = int(state.get("y", 0))
        except Exception:
            return

        if width < MIN_WIDTH:
            width = MIN_WIDTH
        if height < MIN_HEIGHT:
            height = MIN_HEIGHT

        self.root.geometry(f"{width}x{height}+{x}+{y}")
        try:
            self.root.update_idletasks()
        except Exception:
            pass

    def load_window_state(self) -> Optional[Dict[str, int]]:
        if not self.window_state_path.exists():
            return None
        try:
            with self.window_state_path.open("r", encoding="utf-8") as f:
                data = json.load(f)
            if not isinstance(data, dict):
                return None
            required = ("width", "height", "x", "y")
            if not all(key in data for key in required):
                return None
            return {key: int(data[key]) for key in required}
        except Exception:
            return None

    def save_window_state(self) -> None:
        try:
            self.root.update_idletasks()
            if str(self.root.state()) == "zoomed":
                return

            state = {
                "width": int(self.root.winfo_width()),
                "height": int(self.root.winfo_height()),
                "x": int(self.root.winfo_x()),
                "y": int(self.root.winfo_y()),
            }

            with self.window_state_path.open("w", encoding="utf-8") as f:
                json.dump(state, f, indent=2, ensure_ascii=False)
        except Exception:
            pass

    def update_binding_label(self) -> None:
        self.binding_var.set(f"Insert Button: {self.describe_binding(self.input_binding)}")

    def describe_binding(self, binding: Optional[Dict[str, object]]) -> str:
        if not binding:
            return "None"

        source = str(binding.get("source", "")).lower()
        if source == "keyboard":
            key_name = str(binding.get("key_name") or binding.get("keysym") or "Unknown Key")
            return f"Keyboard - {key_name}"

        if source == "controller":
            device_name = str(binding.get("device_name") or "Controller")
            control_type = str(binding.get("control_type") or "input")
            if control_type == "button":
                button_name = str(binding.get("button_name") or f"Button {binding.get('index', '?')}")
                return f"{device_name} - {button_name}"
            if control_type == "hat":
                hat_name = str(binding.get("hat_name") or binding.get("value") or "D-pad")
                return f"{device_name} - {hat_name}"
            if control_type == "axis":
                axis_name = str(binding.get("axis_name") or f"Axis {binding.get('index', '?')}")
                direction_name = str(binding.get("direction_name") or binding.get("direction") or "+")
                return f"{device_name} - {axis_name} {direction_name}"
            return f"{device_name} - Controller Input"

        if source == "xinput":
            device_name = str(binding.get("device_name") or "XInput Controller")
            control_type = str(binding.get("control_type") or "input")
            if control_type == "button":
                button_name = str(binding.get("button_name") or binding.get("index") or "Button")
                return f"{device_name} - {button_name}"
            if control_type == "trigger":
                trigger_name = str(binding.get("trigger_name") or binding.get("index") or "Trigger")
                return f"{device_name} - {trigger_name}"
            if control_type == "axis":
                axis_name = str(binding.get("axis_name") or binding.get("index") or "Stick")
                direction_name = str(binding.get("direction_name") or binding.get("direction") or "+")
                return f"{device_name} - {axis_name} {direction_name}"
            return f"{device_name} - XInput"

        return "Unknown"

    def clear_input_binding(self) -> None:
        self.input_binding = None
        self.save_input_binding()
        self.update_binding_label()
        self.append_status_line("Insert trigger cleared.")

    def open_bind_input_dialog(self) -> None:
        if self.binding_dialog and self.binding_dialog.winfo_exists():
            self.binding_dialog.lift()
            self.binding_dialog.focus_force()
            return

        self.binding_dialog = tk.Toplevel(self.root)
        dialog = self.binding_dialog
        dialog.title("Bind Insert Trigger")
        dialog.geometry("520x220")
        dialog.minsize(460, 200)
        dialog.transient(self.root)
        dialog.grab_set()

        try:
            if self.icon_path.exists():
                dialog.iconbitmap(str(self.icon_path))
        except Exception:
            pass

        self.binding_listen_var.set(
            "Press the keyboard key or controller input you want to use for Insert.\n"
            "Keyboard, buttons, D-pad, triggers, and stick directions are supported.\n"
            "Press Esc to cancel."
        )
        self.binding_capture_active = True

        frame = ttk.Frame(dialog, padding=12)
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, textvariable=self.binding_listen_var, justify="left", wraplength=480).pack(anchor="w")
        ttk.Label(frame, textvariable=self.binding_var, justify="left", wraplength=480).pack(anchor="w", pady=(12, 0))

        controller_note = []
        if self.xinput is not None:
            controller_note.append("XInput ready")
        if self.pygame_ready:
            controller_note.append("SDL/pygame ready")
        note = "Controller listening available: " + ", ".join(controller_note) if controller_note else "Controller listening unavailable. Install pygame for wider gamepad support; XInput pads may still work on Windows if detected by XInput."
        ttk.Label(frame, text=note, justify="left", wraplength=480).pack(anchor="w", pady=(12, 0))

        button_row = ttk.Frame(frame)
        button_row.pack(anchor="center", pady=(18, 0))

        ttk.Button(button_row, text="Cancel", width=14, command=self.close_bind_input_dialog).grid(row=0, column=0, padx=6)
        ttk.Button(button_row, text="Clear Binding", width=14, command=lambda: [self.clear_input_binding(), self.close_bind_input_dialog()]).grid(row=0, column=1, padx=6)

        dialog.protocol("WM_DELETE_WINDOW", self.close_bind_input_dialog)
        dialog.focus_force()

    def close_bind_input_dialog(self) -> None:
        self.binding_capture_active = False
        if self.binding_dialog and self.binding_dialog.winfo_exists():
            try:
                self.binding_dialog.grab_release()
            except Exception:
                pass
            self.binding_dialog.destroy()
        self.binding_dialog = None

    def set_input_binding(self, binding: Dict[str, object]) -> None:
        self.input_binding = binding
        self.save_input_binding()
        self.update_binding_label()
        self.append_status_line(f"Insert trigger bound: {self.describe_binding(binding)}")
        self.close_bind_input_dialog()

    def normalize_axis_name(self, axis_index: int, value: float) -> str:
        direction = "+" if value >= 0 else "-"
        mapping = {
            0: "Left Stick X",
            1: "Left Stick Y",
            2: "Right Stick X",
            3: "Right Stick Y",
            4: "Trigger/Axis 4",
            5: "Trigger/Axis 5",
        }
        return mapping.get(axis_index, f"Axis {axis_index}")

    def axis_direction_name(self, axis_index: int, value: float) -> str:
        if axis_index == 0:
            return "Right" if value >= 0 else "Left"
        if axis_index == 1:
            return "Down" if value >= 0 else "Up"
        if axis_index == 2:
            return "Right" if value >= 0 else "Left"
        if axis_index == 3:
            return "Down" if value >= 0 else "Up"
        return "+" if value >= 0 else "-"

    def hat_value_name(self, value: Tuple[int, int]) -> str:
        mapping = {
            (0, 1): "D-pad Up",
            (0, -1): "D-pad Down",
            (-1, 0): "D-pad Left",
            (1, 0): "D-pad Right",
            (-1, 1): "D-pad Up-Left",
            (1, 1): "D-pad Up-Right",
            (-1, -1): "D-pad Down-Left",
            (1, -1): "D-pad Down-Right",
        }
        return mapping.get(tuple(value), "D-pad")

    def bindings_match(self, binding: Dict[str, object], event_info: Dict[str, object]) -> bool:
        if str(binding.get("source")) != str(event_info.get("source")):
            return False

        source = str(binding.get("source", ""))
        if source == "keyboard":
            return str(binding.get("keysym", "")).lower() == str(event_info.get("keysym", "")).lower()

        if source not in ("controller", "xinput"):
            return False

        if str(binding.get("guid", "")) != str(event_info.get("guid", "")):
            return False
        if str(binding.get("control_type", "")) != str(event_info.get("control_type", "")):
            return False
        if int(binding.get("index", -1)) != int(event_info.get("index", -2)):
            return False

        if str(binding.get("control_type")) == "hat":
            return tuple(binding.get("value", (0, 0))) == tuple(event_info.get("value", (9, 9)))
        if str(binding.get("control_type")) == "axis":
            return str(binding.get("direction", "")) == str(event_info.get("direction", ""))
        return True

    def trigger_insert_from_binding(self) -> None:
        if self.insert_in_progress or not self.cards:
            return
        self.insert_current_card()

    def can_fire_bound_trigger(self) -> bool:
        if self.binding_capture_active:
            return False
        now = time.monotonic()
        if (now - self.last_input_trigger_ts) * 1000.0 < INPUT_TRIGGER_DEBOUNCE_MS:
            return False
        self.last_input_trigger_ts = now
        return True

    def handle_bound_input_event(self, event_info: Dict[str, object]) -> None:
        if self.binding_capture_active:
            self.set_input_binding(event_info)
            return

        if not self.input_binding:
            return

        if self.bindings_match(self.input_binding, event_info) and self.can_fire_bound_trigger():
            self.root.after(0, self.trigger_insert_from_binding)

    def on_key_press(self, event) -> Optional[str]:
        if self.binding_capture_active:
            if event.keysym == "Escape":
                self.close_bind_input_dialog()
                return "break"

            binding = {
                "source": "keyboard",
                "keysym": str(event.keysym),
                "key_name": str(event.keysym),
                "keycode": int(event.keycode),
            }
            self.set_input_binding(binding)
            return "break"

        event_info = {
            "source": "keyboard",
            "keysym": str(event.keysym),
        }
        self.handle_bound_input_event(event_info)
        return None

    def init_controller_input(self) -> None:
        self.init_xinput()

        if not PYGAME_AVAILABLE or pygame is None:
            self.pygame_ready = False
            self.root.after(CONTROLLER_POLL_MS, self.poll_controller_input)
            return

        try:
            pygame.init()
            pygame.joystick.init()
            self.refresh_controller_devices()
            self.pygame_ready = True
            self.root.after(CONTROLLER_POLL_MS, self.poll_controller_input)
        except Exception as exc:
            self.pygame_ready = False
            self.append_status_line(f"Controller input unavailable: {exc}")
            self.root.after(CONTROLLER_POLL_MS, self.poll_controller_input)

    def init_xinput(self) -> None:
        if os.name != "nt":
            self.xinput = None
            return

        candidates = ("xinput1_4.dll", "xinput1_3.dll", "xinput9_1_0.dll", "xinput1_2.dll", "xinput1_1.dll")
        for dll_name in candidates:
            try:
                dll = ctypes.WinDLL(dll_name)
                dll.XInputGetState.argtypes = [wintypes.DWORD, ctypes.POINTER(XINPUT_STATE)]
                dll.XInputGetState.restype = wintypes.DWORD
                self.xinput = dll
                return
            except Exception:
                continue
        self.xinput = None

    def refresh_controller_devices(self) -> None:
        if pygame is None:
            return

        self.joysticks = {}
        try:
            count = pygame.joystick.get_count()
            for idx in range(count):
                joystick = pygame.joystick.Joystick(idx)
                joystick.init()
                instance_id = joystick.get_instance_id() if hasattr(joystick, "get_instance_id") else idx
                self.joysticks[instance_id] = joystick
        except Exception:
            pass

    def build_controller_event_info(self, joystick, control_type: str, index: int, value) -> Dict[str, object]:
        guid = joystick.get_guid() if hasattr(joystick, "get_guid") else str(joystick.get_id())
        device_name = joystick.get_name()
        event_info: Dict[str, object] = {
            "source": "controller",
            "guid": str(guid),
            "device_name": str(device_name),
            "control_type": control_type,
            "index": int(index),
        }

        if control_type == "button":
            event_info["button_name"] = f"Button {int(index)}"
        elif control_type == "hat":
            event_info["value"] = tuple(value)
            event_info["hat_name"] = self.hat_value_name(tuple(value))
        elif control_type == "axis":
            axis_value = float(value)
            event_info["direction"] = "+" if axis_value >= 0 else "-"
            event_info["axis_name"] = self.normalize_axis_name(index, axis_value)
            event_info["direction_name"] = self.axis_direction_name(index, axis_value)
        return event_info

    def get_xinput_state(self, user_index: int) -> Optional[XINPUT_STATE]:
        if self.xinput is None:
            return None
        state = XINPUT_STATE()
        result = self.xinput.XInputGetState(user_index, ctypes.byref(state))
        if result != 0:
            return None
        return state

    def build_xinput_event_info(self, user_index: int, control_type: str, index: object, value) -> Dict[str, object]:
        event_info: Dict[str, object] = {
            "source": "xinput",
            "guid": f"xinput_{user_index}",
            "device_name": f"XInput Controller {user_index + 1}",
            "control_type": control_type,
            "index": index,
        }
        if control_type == "button":
            event_info["button_name"] = str(value)
        elif control_type == "axis":
            axis_name, axis_value = value
            event_info["index"] = str(axis_name)
            event_info["direction"] = "+" if float(axis_value) >= 0 else "-"
            event_info["axis_name"] = str(axis_name)
            if str(axis_name) == "LX":
                event_info["direction_name"] = "Right" if float(axis_value) >= 0 else "Left"
            elif str(axis_name) == "LY":
                event_info["direction_name"] = "Down" if float(axis_value) >= 0 else "Up"
            elif str(axis_name) == "RX":
                event_info["direction_name"] = "Right" if float(axis_value) >= 0 else "Left"
            elif str(axis_name) == "RY":
                event_info["direction_name"] = "Down" if float(axis_value) >= 0 else "Up"
            else:
                event_info["direction_name"] = "Positive" if float(axis_value) >= 0 else "Negative"
        elif control_type == "trigger":
            trigger_name, trigger_value = value
            event_info["index"] = str(trigger_name)
            event_info["value"] = int(trigger_value)
            event_info["trigger_name"] = str(trigger_name)
        return event_info

    def poll_xinput_controllers(self) -> None:
        if self.xinput is None:
            return

        for user_index in range(4):
            state = self.get_xinput_state(user_index)
            if state is None:
                self.xinput_connected[user_index] = False
                self.xinput_prev_state.pop(user_index, None)
                continue

            gamepad = state.Gamepad
            prev = self.xinput_prev_state.get(user_index, {})

            buttons = int(gamepad.wButtons)
            prev_buttons = int(prev.get("buttons", 0))
            for mask, name in XINPUT_BUTTONS:
                if (buttons & mask) and not (prev_buttons & mask):
                    info = self.build_xinput_event_info(user_index, "button", name, name)
                    self.handle_bound_input_event(info)

            left_trigger = int(gamepad.bLeftTrigger)
            prev_left_trigger = int(prev.get("left_trigger", 0))
            if left_trigger >= XINPUT_TRIGGER_THRESHOLD and prev_left_trigger < XINPUT_TRIGGER_THRESHOLD:
                info = self.build_xinput_event_info(user_index, "trigger", "LT", ("LT", left_trigger))
                self.handle_bound_input_event(info)

            right_trigger = int(gamepad.bRightTrigger)
            prev_right_trigger = int(prev.get("right_trigger", 0))
            if right_trigger >= XINPUT_TRIGGER_THRESHOLD and prev_right_trigger < XINPUT_TRIGGER_THRESHOLD:
                info = self.build_xinput_event_info(user_index, "trigger", "RT", ("RT", right_trigger))
                self.handle_bound_input_event(info)

            stick_axes = [
                ("LX", int(gamepad.sThumbLX)),
                ("LY", int(gamepad.sThumbLY)),
                ("RX", int(gamepad.sThumbRX)),
                ("RY", int(gamepad.sThumbRY)),
            ]
            prev_sticks = prev.get("sticks", {})
            for axis_name, axis_value in stick_axes:
                prev_axis = int(prev_sticks.get(axis_name, 0))
                if axis_value >= XINPUT_STICK_THRESHOLD and prev_axis < XINPUT_STICK_THRESHOLD:
                    info = self.build_xinput_event_info(user_index, "axis", axis_name, (axis_name, axis_value))
                    self.handle_bound_input_event(info)
                elif axis_value <= -XINPUT_STICK_THRESHOLD and prev_axis > -XINPUT_STICK_THRESHOLD:
                    info = self.build_xinput_event_info(user_index, "axis", axis_name, (axis_name, axis_value))
                    self.handle_bound_input_event(info)

            self.xinput_connected[user_index] = True
            self.xinput_prev_state[user_index] = {
                "buttons": buttons,
                "left_trigger": left_trigger,
                "right_trigger": right_trigger,
                "sticks": {name: value for name, value in stick_axes},
            }

    def poll_controller_input(self) -> None:
        try:
            self.poll_xinput_controllers()
            if self.pygame_ready and pygame is not None:
                for event in pygame.event.get():
                    if event.type in (pygame.JOYDEVICEADDED, pygame.JOYDEVICEREMOVED):
                        self.refresh_controller_devices()
                        continue

                    if event.type == pygame.JOYBUTTONDOWN:
                        joystick = self.joysticks.get(getattr(event, "instance_id", getattr(event, "joy", None)))
                        if joystick is not None:
                            info = self.build_controller_event_info(joystick, "button", int(event.button), 1)
                            self.handle_bound_input_event(info)

                    elif event.type == pygame.JOYHATMOTION and tuple(event.value) != (0, 0):
                        joystick = self.joysticks.get(getattr(event, "instance_id", getattr(event, "joy", None)))
                        if joystick is not None:
                            info = self.build_controller_event_info(joystick, "hat", int(event.hat), tuple(event.value))
                            self.handle_bound_input_event(info)

                    elif event.type == pygame.JOYAXISMOTION:
                        axis_value = float(event.value)
                        threshold = TRIGGER_BIND_THRESHOLD if int(event.axis) >= 4 else AXIS_BIND_THRESHOLD
                        if abs(axis_value) >= threshold:
                            joystick = self.joysticks.get(getattr(event, "instance_id", getattr(event, "joy", None)))
                            if joystick is not None:
                                info = self.build_controller_event_info(joystick, "axis", int(event.axis), axis_value)
                                self.handle_bound_input_event(info)
        except Exception:
            pass
        finally:
            self.root.after(CONTROLLER_POLL_MS, self.poll_controller_input)

    def ensure_template_dir(self) -> None:
        self.template_dir.mkdir(parents=True, exist_ok=True)

    def get_template_names(self) -> List[str]:
        self.ensure_template_dir()
        return [p.name for p in sorted(self.template_dir.glob("*.png"))]

    def open_manager(self) -> None:
        self.ensure_template_dir()
        if self.manager_window and self.manager_window.winfo_exists():
            self.manager_window.lift()
            self.manager_window.focus_force()
            self.manager_window.refresh_lists()
            return
        self.manager_window = CardManagerWindow(self)

    def startup_sequence(self) -> None:
        threading.Thread(target=self._startup_worker, daemon=True).start()

    def _startup_worker(self) -> None:
        try:
            self._load_config()

            if not self.exe_path.exists():
                raise FileNotFoundError(f"Could not find {EXE_NAME} next to the app:\n{self.exe_path}")

            popen_kwargs = {
                "args": [str(self.exe_path)],
                "cwd": str(self.base_dir),
                "stdout": subprocess.PIPE,
                "stderr": subprocess.STDOUT,
                "text": True,
                "bufsize": 1,
            }

            if os.name == "nt":
                startupinfo = subprocess.STARTUPINFO()
                startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                creationflags = subprocess.CREATE_NO_WINDOW
                popen_kwargs["startupinfo"] = startupinfo
                popen_kwargs["creationflags"] = creationflags

            self.process = subprocess.Popen(**popen_kwargs)
            self.start_status_reader()

            if not wait_for_port(HOST, self.api_port, SERVER_WAIT_SECONDS, SERVER_POLL_INTERVAL):
                raise TimeoutError(f"{HOST}:{self.api_port} did not become ready within {SERVER_WAIT_SECONDS} seconds.")

            self.root.after(0, lambda: self.append_status_line("YACardEmu ready."))
            self.root.after(0, self.reload_cards_after_startup)

        except Exception as exc:
            self.root.after(0, lambda: self._startup_failed(exc))

    def _load_config(self) -> None:
        parser = configparser.ConfigParser()

        if self.config_path.exists():
            parser.read(self.config_path, encoding="utf-8")
        else:
            parser["config"] = {}

        cfg = parser["config"] if parser.has_section("config") else {}

        api_port_raw = str(cfg.get("apiport", DEFAULT_API_PORT)).strip()
        basepath_raw = str(cfg.get("basepath", DEFAULT_BASEPATH)).strip()

        try:
            self.api_port = int(api_port_raw)
        except ValueError:
            raise ValueError(f"Invalid apiport in {self.config_path}:\n{api_port_raw}")

        self.cards_dir = self.resolve_cards_path(basepath_raw)
        self.base_url = f"http://{HOST}:{self.api_port}"
        self.api_inserted_card = f"{self.base_url}/api/v1/insertedCard"
        self.api_cards = f"{self.base_url}/api/v1/cards"

    def resolve_cards_path(self, basepath_raw: str) -> Path:
        if not basepath_raw:
            return (self.base_dir / "cards").resolve()

        expanded = os.path.expandvars(os.path.expanduser(basepath_raw))
        path_obj = Path(expanded)

        if path_obj.is_absolute():
            return path_obj.resolve()

        return (self.base_dir / path_obj).resolve()

    def _startup_failed(self, exc: Exception) -> None:
        self._set_controls_enabled(False)
        self.image_label.configure(image="", text=f"Startup failed:\n\n{exc}")
        self.current_photo = None
        messagebox.showerror("Startup failed", str(exc))

    def reload_cards_after_startup(self) -> None:
        try:
            self.load_cards_from_folder(self.cards_dir)
            self._set_controls_enabled(True)
            self.root.after(NEW_CARD_POLL_MS, self.poll_for_new_card_bin)
            self.root.after(SELECTED_CARD_POLL_MS, self.poll_selected_card_files)
        except Exception as exc:
            self._set_controls_enabled(False)
            self.image_label.configure(image="", text=f"Card loading failed:\n\n{exc}")
            self.current_photo = None
            messagebox.showerror("Card loading failed", str(exc))

    def is_card_bin_filename(self, name: str) -> bool:
        lower = name.lower()
        return re.fullmatch(r".+\.bin(?:\.\d+)?", lower) is not None

    def list_card_entries_in_folder(self, folder: Path) -> List[CardEntry]:
        cards: List[CardEntry] = []

        for entry in sorted(folder.iterdir(), key=lambda p: p.name.lower()):
            if not entry.is_file():
                continue
            if not self.is_card_bin_filename(entry.name):
                continue

            png_path = folder / f"{entry.name}.png"
            cards.append(CardEntry(
                bin_path=entry,
                png_path=png_path if png_path.exists() else None,
                name=entry.name,
            ))

        return cards

    def get_latest_unhandled_new_card(self) -> Tuple[Optional[CardEntry], Optional[str]]:
        current_names = {c.name for c in self.cards}
        candidates: List[Tuple[float, CardEntry, str]] = []

        for card in self.list_card_entries_in_folder(self.cards_dir):
            if card.name in current_names:
                continue
            signature = self.get_card_signature(card)
            if not signature or signature == self.last_handled_new_card_signature:
                continue

            try:
                stat_bin = card.bin_path.stat()
                newest_mtime = stat_bin.st_mtime_ns
                if card.png_path and card.png_path.exists():
                    newest_mtime = max(newest_mtime, card.png_path.stat().st_mtime_ns)
            except Exception:
                newest_mtime = 0

            candidates.append((newest_mtime, card, signature))

        if not candidates:
            return None, None

        candidates.sort(key=lambda item: (item[0], item[1].name))
        _mtime, card, signature = candidates[-1]
        return card, signature

    def load_cards_from_folder(self, folder: Path) -> None:
        if not folder.exists() or not folder.is_dir():
            raise FileNotFoundError(f"Cards folder not found:\n{folder}")

        previous_name = self.get_selected_card_name()
        self.cards = self.list_card_entries_in_folder(folder)
        self.prune_dead_links()

        if not self.cards:
            self.current_index = 0
            self.current_card_signature = None
            self.pending_card_refresh_signature = None
            self.selected_name_var.set("")
            self.image_label.configure(image="", text="No .bin cards found.")
            self.current_photo = None
            self.image_frame.configure(width=PREVIEW_MAX_W + 12, height=PREVIEW_MAX_H + 12)
            self._fit_window_to_content(PREVIEW_MAX_W + 12, PREVIEW_MAX_H + 12)
            return

        if previous_name:
            matched_index = next((idx for idx, card in enumerate(self.cards) if card.name == previous_name), None)
            if matched_index is not None:
                self.current_index = matched_index
            else:
                self.current_index = min(self.current_index, len(self.cards) - 1)
        else:
            self.current_index = min(self.current_index, len(self.cards) - 1)

        self.show_current_card()

    def prune_dead_links(self) -> None:
        live_names = {c.name for c in self.cards}
        changed = False
        for key in list(self.card_links.keys()):
            if key not in live_names:
                del self.card_links[key]
                changed = True
        if changed:
            self.save_card_links()

    def get_card_signature(self, card: CardEntry) -> Optional[str]:
        try:
            stat_bin = card.bin_path.stat()
            if card.png_path and card.png_path.exists():
                stat_png = card.png_path.stat()
                return f"{int(stat_bin.st_mtime_ns)}:{stat_bin.st_size}:{int(stat_png.st_mtime_ns)}:{stat_png.st_size}"
            return f"{int(stat_bin.st_mtime_ns)}:{stat_bin.st_size}:nopng"
        except Exception:
            return None

    def begin_preview_hold(self, card_name: str) -> None:
        self.end_preview_hold(cancel_only=True)
        self.preview_hold_active = True
        self.preview_hold_card_name = card_name
        self.preview_hold_ignored_change = False
        self.preview_hold_job = self.root.after(PREVIEW_HOLD_MAX_MS, self.end_preview_hold)

    def end_preview_hold(self, cancel_only: bool = False) -> None:
        self.preview_hold_active = False
        self.preview_hold_card_name = None
        self.preview_hold_ignored_change = False
        if self.preview_hold_job:
            try:
                self.root.after_cancel(self.preview_hold_job)
            except Exception:
                pass
            self.preview_hold_job = None
        if cancel_only:
            return

    def should_ignore_preview_refresh(self, latest_signature: Optional[str]) -> bool:
        if not self.preview_hold_active or not self.cards:
            return False

        selected_name = self.get_selected_card_name()
        if selected_name != self.preview_hold_card_name:
            self.end_preview_hold(cancel_only=True)
            return False

        if latest_signature is None or latest_signature == self.current_card_signature:
            return False

        if not self.preview_hold_ignored_change:
            self.preview_hold_ignored_change = True
            self.current_card_signature = latest_signature
            self.pending_card_refresh_signature = latest_signature
            return True

        self.end_preview_hold(cancel_only=True)
        return False

    def poll_selected_card_files(self) -> None:
        try:
            if self.cards:
                card = self.cards[self.current_index]
                latest_signature = self.get_card_signature(card)
                if latest_signature != self.current_card_signature:
                    if self.should_ignore_preview_refresh(latest_signature):
                        pass
                    else:
                        self.end_preview_hold(cancel_only=True)
                        self.schedule_selected_card_refresh(latest_signature)
        except Exception:
            pass
        finally:
            self.root.after(SELECTED_CARD_POLL_MS, self.poll_selected_card_files)

    def schedule_selected_card_refresh(self, latest_signature: Optional[str]) -> None:
        self.pending_card_refresh_signature = latest_signature
        if self.preview_refresh_job:
            try:
                self.root.after_cancel(self.preview_refresh_job)
            except Exception:
                pass
        self.preview_refresh_job = self.root.after(PREVIEW_REFRESH_DEBOUNCE_MS, self.refresh_selected_card_if_changed)

    def refresh_selected_card_if_changed(self) -> None:
        self.preview_refresh_job = None
        if not self.cards:
            return

        selected_name = self.get_selected_card_name()
        if not selected_name:
            return

        try:
            self.cards = self.list_card_entries_in_folder(self.cards_dir)
            self.prune_dead_links()

            if not self.cards:
                self.current_index = 0
                self.current_card_signature = None
                self.pending_card_refresh_signature = None
                self.selected_name_var.set("")
                self.image_label.configure(image="", text="No .bin cards found.")
                self.current_photo = None
                self.image_frame.configure(width=PREVIEW_MAX_W + 12, height=PREVIEW_MAX_H + 12)
                self._fit_window_to_content(PREVIEW_MAX_W + 12, PREVIEW_MAX_H + 12)
                return

            matched_index = next((idx for idx, card in enumerate(self.cards) if card.name == selected_name), None)
            if matched_index is None:
                self.load_cards_from_folder(self.cards_dir)
                return

            self.current_index = matched_index
            latest_signature = self.get_card_signature(self.cards[self.current_index])
            if latest_signature != self.current_card_signature:
                self.show_current_card()
                self.end_preview_hold(cancel_only=True)
        except Exception:
            pass

    def poll_for_new_card_bin(self) -> None:
        try:
            new_card, signature = self.get_latest_unhandled_new_card()

            if new_card is not None:
                self.load_cards_from_folder(self.cards_dir)

            if new_card is not None and signature and not self.auto_setup_open:
                detected_name = new_card.name
                self.auto_setup_open = True

                def open_dialog():
                    try:
                        self.load_cards_from_folder(self.cards_dir)
                        dlg = AutoSetupDialog(self, detected_name, signature)
                        self.root.wait_window(dlg)
                    finally:
                        self.auto_setup_open = False

                self.root.after(50, open_dialog)
        except Exception:
            pass
        finally:
            self.root.after(NEW_CARD_POLL_MS, self.poll_for_new_card_bin)

    def show_current_card(self) -> None:
        if not self.cards:
            self.current_card_signature = None
            return

        card = self.cards[self.current_index]
        self.selected_name_var.set(card.name)
        self.current_card_signature = self.get_card_signature(card)
        self.pending_card_refresh_signature = self.current_card_signature

        try:
            if card.png_path and card.png_path.exists():
                img = Image.open(card.png_path)
                img.thumbnail((PREVIEW_MAX_W, PREVIEW_MAX_H), Image.LANCZOS)
                self.current_photo = ImageTk.PhotoImage(img)

                frame_w = max(self.current_photo.width() + 12, 120)
                frame_h = max(self.current_photo.height() + 12, 120)
                self.image_frame.configure(width=frame_w, height=frame_h)

                self.image_label.configure(image=self.current_photo, text="")
                self.image_label.place(relx=0.5, rely=0.5, anchor="center")

                self._fit_window_to_content(frame_w, frame_h)
            else:
                self.current_photo = None
                self.image_frame.configure(width=PREVIEW_MAX_W + 12, height=PREVIEW_MAX_H + 12)
                self.image_label.configure(image="", text=f"No PNG preview yet for:\n{card.name}")
                self.image_label.place(relx=0.5, rely=0.5, anchor="center")
                self._fit_window_to_content(PREVIEW_MAX_W + 12, PREVIEW_MAX_H + 12)

        except Exception as exc:
            self.current_photo = None
            self.image_frame.configure(width=PREVIEW_MAX_W + 12, height=PREVIEW_MAX_H + 12)
            self.image_label.configure(image="", text=f"Failed to load image:\n{card.name}\n\n{exc}")
            self.image_label.place(relx=0.5, rely=0.5, anchor="center")
            self._fit_window_to_content(PREVIEW_MAX_W + 12, PREVIEW_MAX_H + 12)

    def show_current_card_if_name(self, card_name: str) -> None:
        for idx, card in enumerate(self.cards):
            if card.name == card_name:
                self.current_index = idx
                self.show_current_card()
                return

    def _fit_window_to_content(self, frame_w: int, frame_h: int) -> None:
        content_w = frame_w + 210
        content_h = frame_h + 95

        width = max(MIN_WIDTH, content_w)
        height = max(MIN_HEIGHT, content_h)

        self.root.geometry(f"{width}x{height}")

    def prev_card(self) -> None:
        if not self.cards:
            return
        self.current_index = (self.current_index - 1) % len(self.cards)
        self.show_current_card()

    def next_card(self) -> None:
        if not self.cards:
            return
        self.current_index = (self.current_index + 1) % len(self.cards)
        self.show_current_card()

    def get_selected_card_name(self) -> Optional[str]:
        if not self.cards:
            return None
        return self.cards[self.current_index].name

    def get_card_entry(self, card_name: str) -> Optional[CardEntry]:
        for card in self.cards:
            if card.name == card_name:
                return card
        return None

    def sanitize_base_name(self, name: str) -> str:
        cleaned = name.strip()
        cleaned = re.sub(r"\.bin$", "", cleaned, flags=re.IGNORECASE)
        cleaned = re.sub(r'[<>:"/\\\\|?*\x00-\x1f]', "_", cleaned)
        cleaned = cleaned.strip(" .")
        return cleaned

    def rename_card_entry(self, old_card_name: str, new_base_name: str) -> Tuple[bool, str, str]:
        card = self.get_card_entry(old_card_name)
        if not card:
            return False, f"Card not found: {old_card_name}", ""

        clean_base = self.sanitize_base_name(new_base_name)
        if not clean_base:
            return False, "The new card name is empty or invalid.", ""

        new_card_name = f"{clean_base}.bin"
        new_bin = self.cards_dir / new_card_name
        new_png = self.cards_dir / f"{new_card_name}.png"

        if new_bin.exists() or new_png.exists():
            return False, f"{new_card_name} already exists.", ""

        try:
            card.bin_path.rename(new_bin)
            if card.png_path and card.png_path.exists():
                card.png_path.rename(new_png)
        except Exception as exc:
            return False, f"Rename failed: {exc}", ""

        if old_card_name in self.card_links:
            self.card_links[new_card_name] = self.card_links.pop(old_card_name)
            self.save_card_links()

        return True, f"Renamed {old_card_name} to {new_card_name}", new_card_name

    def restore_template_to_card(self, card_name: str) -> Tuple[bool, str]:
        bin_path = self.cards_dir / card_name
        if not bin_path.exists():
            return False, f"Card not found: {card_name}"

        template_name = self.card_links.get(card_name, "").strip()
        if not template_name:
            return False, f"No template linked for {card_name}"

        template_path = self.template_dir / template_name
        if not template_path.exists():
            return False, f"Linked template not found:\n{template_path}"

        target_png = self.cards_dir / f"{card_name}.png"

        try:
            shutil.copyfile(template_path, target_png)
        except Exception as exc:
            return False, f"Failed copying template: {exc}"

        card = self.get_card_entry(card_name)
        if card:
            card.png_path = target_png

        return True, f"Reset {card_name}.png from template {template_name}"

    def wait_for_selected_card(self, selected: str) -> bool:
        deadline = time.time() + SELECTION_WAIT_SECONDS

        while time.time() < deadline:
            try:
                response = self.session.get(self.api_cards, timeout=HTTP_TIMEOUT)
                response.raise_for_status()
                data = response.json()

                if isinstance(data, list):
                    for item in data:
                        if not isinstance(item, dict):
                            continue
                        name = str(item.get("name", "")).strip()
                        active = bool(item.get("active", False))
                        if name == selected and active:
                            return True
            except Exception:
                pass

            time.sleep(SELECTION_POLL_INTERVAL)

        return False

    def insert_current_card(self) -> None:
        if self.insert_in_progress:
            return

        selected = self.get_selected_card_name()
        if not selected:
            messagebox.showerror("No card", "No card is selected.")
            return

        self.start_reading_overlay()
        self.insert_in_progress = True

        def worker():
            try:
                reset_ok, reset_msg = self.restore_template_to_card(selected)
                if not reset_ok and selected in self.card_links:
                    self.root.after(0, lambda: self.append_status_line(f"Template reset failed: {reset_msg}"))
                    self.root.after(0, lambda: messagebox.showerror("Template reset failed", reset_msg))
                    return

                if reset_ok:
                    self.root.after(0, lambda: self.begin_preview_hold(selected))

                change_payload = {
                    "redirect": "",
                    "cardname": selected,
                }
                change_response = self.session.post(
                    self.api_inserted_card,
                    data=change_payload,
                    timeout=HTTP_TIMEOUT
                )
                change_response.raise_for_status()

                selection_confirmed = self.wait_for_selected_card(selected)

                if not selection_confirmed:
                    time.sleep(INSERT_RETRY_DELAY)

                insert_payload = {
                    "redirect": "",
                    "loadonly": "Insert",
                }
                insert_response = self.session.post(
                    self.api_inserted_card,
                    data=insert_payload,
                    timeout=HTTP_TIMEOUT
                )
                insert_response.raise_for_status()

                if not selection_confirmed:
                    time.sleep(INSERT_RETRY_DELAY)
                    retry_response = self.session.post(
                        self.api_inserted_card,
                        data=insert_payload,
                        timeout=HTTP_TIMEOUT
                    )
                    retry_response.raise_for_status()

                self.root.after(0, lambda: self.append_status_line(f"Inserted {selected}"))
                self.root.after(0, lambda: self.load_cards_from_folder(self.cards_dir))
                self.root.after(0, lambda: self.show_current_card_if_name(selected))

            except Exception as exc:
                self.root.after(0, lambda: self.append_status_line(f"Insert failed: {exc}"))
                self.root.after(0, lambda: messagebox.showerror("Insert failed", str(exc)))
            finally:
                self.root.after(0, lambda: setattr(self, "insert_in_progress", False))

        threading.Thread(target=worker, daemon=True).start()

    def start_status_reader(self) -> None:
        if not self.process or not self.process.stdout or (self.status_reader_thread and self.status_reader_thread.is_alive()):
            return

        def reader() -> None:
            try:
                for raw_line in self.process.stdout:
                    line = raw_line.strip()
                    if line:
                        self.root.after(0, lambda msg=line: self.append_status_line(msg))
            except Exception:
                pass

        self.status_reader_thread = threading.Thread(target=reader, daemon=True)
        self.status_reader_thread.start()

    def append_status_line(self, message: str) -> None:
        self.status_lines.append(message)
        padded = [""] * (MAX_STATUS_LINES - len(self.status_lines)) + list(self.status_lines)
        for var, value in zip(self.status_vars, padded):
            var.set(value)

    def start_reading_overlay(self) -> None:
        self.stop_reading_overlay(clear_only=True)
        self.reading_overlay_active = True
        self.reading_overlay_visible = False
        self._toggle_reading_overlay()
        self.reading_overlay_job = self.root.after(READING_OVERLAY_DURATION_MS, self.stop_reading_overlay)

    def _toggle_reading_overlay(self) -> None:
        if not self.reading_overlay_active:
            return

        self.reading_overlay_visible = not self.reading_overlay_visible
        if self.reading_overlay_visible:
            self.reading_overlay_label.grid()
        else:
            self.reading_overlay_label.grid_remove()

        self.reading_flash_job = self.root.after(READING_OVERLAY_FLASH_MS, self._toggle_reading_overlay)

    def stop_reading_overlay(self, clear_only: bool = False) -> None:
        self.reading_overlay_active = False
        self.reading_overlay_visible = False
        self.reading_overlay_label.grid_remove()

        if self.reading_overlay_job:
            try:
                self.root.after_cancel(self.reading_overlay_job)
            except Exception:
                pass
            self.reading_overlay_job = None

        if self.reading_flash_job:
            try:
                self.root.after_cancel(self.reading_flash_job)
            except Exception:
                pass
            self.reading_flash_job = None

        if clear_only:
            return

    def shutdown_yacardemu(self) -> None:
        if not self.process:
            return

        try:
            if self.process.poll() is None:
                self.process.terminate()
                self.process.wait(timeout=2)
        except Exception:
            try:
                if self.process.poll() is None:
                    self.process.kill()
            except Exception:
                pass
        finally:
            self.process = None

    def on_close(self) -> None:
        try:
            self.save_window_state()
            self.stop_reading_overlay(clear_only=True)
            if self.preview_refresh_job:
                try:
                    self.root.after_cancel(self.preview_refresh_job)
                except Exception:
                    pass
                self.preview_refresh_job = None
            self.end_preview_hold(cancel_only=True)
            self.shutdown_yacardemu()
            if self.pygame_ready and pygame is not None:
                try:
                    pygame.quit()
                except Exception:
                    pass
        finally:
            self.root.destroy()

    def open_path(self, path: Path) -> None:
        try:
            if sys.platform.startswith("win"):
                os.startfile(path)  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(path)])
            else:
                subprocess.Popen(["xdg-open", str(path)])
        except Exception as exc:
            messagebox.showerror("Open failed", str(exc))


def wait_for_port(host: str, port: int, timeout_seconds: float, poll_interval: float) -> bool:
    end_time = time.time() + timeout_seconds
    while time.time() < end_time:
        try:
            with socket.create_connection((host, port), timeout=1.0):
                return True
        except OSError:
            time.sleep(poll_interval)
    return False


def main() -> None:
    root = tk.Tk()

    try:
        style = ttk.Style()
        if "vista" in style.theme_names():
            style.theme_use("vista")
    except Exception:
        pass

    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()