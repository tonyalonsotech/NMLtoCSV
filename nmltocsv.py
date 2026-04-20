import csv
import json
import os
import xml.etree.ElementTree as ET
from urllib.parse import unquote
import tkinter as tk
from tkinter import filedialog, messagebox

# ----------------
# Version History
# ----------------
#
# v1.0 - Initial release
# - Converts NML to CSV
#
# v1.1
# - Added drag-and-drop support
#
# v1.2
# - Refined UI text to reduce redundancy
#
# v1.3
# - Decoupled column selection order from default column order
#
# v1.4
# - Replaced checkboxes with tile-based selection UI
#
# v1.5 - Current Version
# - Added automatic preference saving and loading
# - Preferences are stored in the Windows %APPDATA% directory as nmltocsv_preferences.json 


# Drag and drop support
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    DND_AVAILABLE = True
except ImportError:
    DND_AVAILABLE = False


# Data extraction
def decode_traktor_location(location_elem):
    if location_elem is None:
        return ""

    volume = unquote(location_elem.attrib.get("VOLUME", ""))
    directory = unquote(location_elem.attrib.get("DIR", ""))
    filename = unquote(location_elem.attrib.get("FILE", ""))

    directory = directory.replace(":", "").strip("/")

    if volume:
        if directory:
            return os.path.join(volume + os.sep, directory, filename)
        return os.path.join(volume + os.sep, filename)

    if directory:
        return os.path.join(os.sep, directory, filename)

    return filename


def get_attr(elem, attr_name, default=""):
    if elem is None:
        return default
    return elem.attrib.get(attr_name, default)


def first_non_empty(*values):
    for value in values:
        if value is not None:
            value = str(value).strip()
            if value:
                return value
    return ""


def extract_key(entry):
    info = entry.find("INFO")
    musical_key = entry.find("MUSICAL_KEY")

    return first_non_empty(
        get_attr(entry, "KEY"),
        get_attr(entry, "TONAL_KEY"),
        get_attr(info, "KEY"),
        get_attr(info, "TONAL_KEY"),
        get_attr(musical_key, "VALUE"),
        get_attr(musical_key, "KEY"),
        get_attr(musical_key, "TONAL_KEY"),
    )


def extract_bpm(entry):
    tempo = entry.find("TEMPO")
    info = entry.find("INFO")

    bpm = first_non_empty(
        get_attr(tempo, "BPM"),
        get_attr(tempo, "VALUE"),
        get_attr(entry, "BPM"),
        get_attr(info, "BPM"),
    )

    if not bpm:
        return ""

    try:
        bpm_float = float(bpm)
        if bpm_float.is_integer():
            return str(int(bpm_float))
        return str(round(bpm_float, 2))
    except ValueError:
        return bpm


def extract_label(entry, info, album):
    return first_non_empty(
        get_attr(album, "LABEL"),
        get_attr(info, "LABEL"),
        get_attr(entry, "LABEL"),
        get_attr(album, "PUBLISHER"),
        get_attr(info, "PUBLISHER"),
        get_attr(entry, "PUBLISHER"),
        get_attr(album, "ORGANIZATION"),
        get_attr(info, "ORGANIZATION"),
        get_attr(entry, "ORGANIZATION"),
    )


def extract_year(entry, info, album):
    raw = first_non_empty(
        get_attr(info, "YEAR"),
        get_attr(album, "YEAR"),
        get_attr(entry, "YEAR"),
        get_attr(info, "DATE"),
        get_attr(album, "DATE"),
        get_attr(entry, "DATE"),
        get_attr(info, "RELEASE_DATE"),
        get_attr(album, "RELEASE_DATE"),
        get_attr(entry, "RELEASE_DATE"),
    )

    if not raw:
        return ""

    for sep in ["-", "/", "."]:
        if sep in raw and len(raw) >= 4:
            possible_year = raw[:4]
            if possible_year.isdigit():
                return possible_year

    if raw[:4].isdigit():
        return raw[:4]

    return raw


def extract_time(entry):
    info = entry.find("INFO")

    raw = first_non_empty(
        get_attr(info, "PLAYTIME"),
        get_attr(entry, "PLAYTIME"),
        get_attr(info, "TIME"),
        get_attr(entry, "TIME"),
    )

    if not raw:
        return ""

    try:
        total_seconds = int(float(raw))
        minutes = total_seconds // 60
        seconds = total_seconds % 60
        return f"{minutes}:{seconds:02d}"
    except ValueError:
        return raw


def extract_comment2(entry):
    info = entry.find("INFO")

    return first_non_empty(
        get_attr(info, "RATING"),
        get_attr(entry, "RATING"),
        get_attr(info, "COMMENT2"),
        get_attr(entry, "COMMENT2"),
        get_attr(info, "COMMENT_2"),
        get_attr(entry, "COMMENT_2"),
    )


def extract_entries(nml_path):
    tree = ET.parse(nml_path)
    root = tree.getroot()

    rows = []

    for entry in root.findall(".//COLLECTION/ENTRY"):
        info = entry.find("INFO")
        album = entry.find("ALBUM")
        location = entry.find("LOCATION")

        artist = first_non_empty(
            get_attr(entry, "ARTIST"),
            get_attr(info, "ARTIST")
        )

        track_title = first_non_empty(
            get_attr(entry, "TITLE"),
            get_attr(info, "TITLE")
        )

        album_title = first_non_empty(
            get_attr(album, "TITLE"),
            get_attr(info, "ALBUM"),
            get_attr(entry, "ALBUM")
        )

        genre = first_non_empty(
            get_attr(info, "GENRE"),
            get_attr(entry, "GENRE"),
            get_attr(album, "GENRE")
        )

        label = extract_label(entry, info, album)

        comment = first_non_empty(
            get_attr(info, "COMMENT"),
            get_attr(info, "COMMENTS"),
            get_attr(entry, "COMMENT"),
            get_attr(entry, "COMMENTS")
        )

        comment2 = extract_comment2(entry)
        year = extract_year(entry, info, album)
        musical_key = extract_key(entry)
        bpm = extract_bpm(entry)
        track_time = extract_time(entry)
        file_location = decode_traktor_location(location)

        rows.append({
            "Artist": artist,
            "Track Title": track_title,
            "Album": album_title,
            "Label": label,
            "Genre": genre,
            "Comment": comment,
            "Comment 2": comment2,
            "Year": year,
            "Key": musical_key,
            "BPM": bpm,
            "Time": track_time,
            "File Location": file_location,
        })

    return rows


def write_csv(rows, output_csv, selected_columns):
    with open(output_csv, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=selected_columns)
        writer.writeheader()

        for row in rows:
            filtered_row = {col: row.get(col, "") for col in selected_columns}
            writer.writerow(filtered_row)


# UI
class TraktorExporterApp:
    def toggle_column_tile(self, col):
        self.column_vars[col].set(not self.column_vars[col].get())
        self.update_column_tile(col)
        self.refresh_selected_listbox()
        self.save_preferences()

    def update_column_tile(self, col):
        tile = self.column_tiles.get(col)
        if not tile:
            return

        selected = self.column_vars[col].get()

        bg = self.colors["tile_selected"] if selected else self.colors["tile_unselected"]
        border = self.colors["tile_border_selected"] if selected else self.colors["tile_border_unselected"]
        thickness = 2 if selected else 1

        tile["frame"].configure(bg=bg, highlightbackground=border, highlightthickness=thickness)
        tile["label"].configure(bg=bg)

    def update_all_column_tiles(self):
        for col in self.available_columns:
            self.update_column_tile(col)

    def set_tile_hover(self, col, hovering):
        tile = self.column_tiles.get(col)
        if not tile:
            return

        selected = self.column_vars[col].get()

        if hovering:
            bg = "#2a1740" if not selected else "#2d1548"
        else:
            bg = self.colors["tile_selected"] if selected else self.colors["tile_unselected"]

        tile["frame"].configure(bg=bg)
        tile["label"].configure(bg=bg)

    def __init__(self, root):
        self.root = root
        self.root.title("NML to CSV Converter")
        self.root.geometry("900x745")
        self.root.minsize(900, 745)

        self.colors = {
            "bg": "#1b1026",
            "panel": "#261537",
            "panel_alt": "#2f1b45",
            "entry": "#341f4f",
            "text": "#f3eefe",
            "muted": "#cbb9e8",
            "accent": "#8b5cf6",
            "accent_hover": "#9d73ff",
            "border": "#50306f",
            "success": "#c4b5fd",
            "listbox_bg": "#221331",
            "listbox_select": "#6d48c9",
            "drag_glow": "#7c4dff",
            "tile_selected": "#241238",
            "tile_unselected": "#2f1b45",
            "tile_border_selected": "#a78bfa",
            "tile_border_unselected": "#50306f",
        }

        self.available_columns = sorted([
            "Artist",
            "Track Title",
            "Album",
            "Label",
            "Genre",
            "Comment",
            "Comment 2",
            "Year",
            "Key",
            "BPM",
            "Time",
            "File Location",
        ])

        self.default_selected_order = [
            "Artist",
            "Track Title",
            "Album",
            "Label",
            "Year",
            "Genre",
            "Comment",
            "Comment 2",
            "Key",
            "BPM",
            "Time",
            "File Location",
        ]

        appdata_base = os.getenv("APPDATA") or os.path.expanduser("~")
        appdata_dir = os.path.join(appdata_base, "NMLtoCSV")
        os.makedirs(appdata_dir, exist_ok=True)

        self.preferences_file = os.path.join(
            appdata_dir,
            "nmltocsv_preferences.json"
        )

        self.input_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.status_text = tk.StringVar(
            value="Use Browse or drag & drop your .nml file anywhere. Choose and order your columns. Then hit Export to generate your CSV."
        )

        self.column_vars = {
            col: tk.BooleanVar(value=True)
            for col in self.available_columns
        }

        self.app_icon = None
        self.drag_index = None
        self.dragging_list_item = False
        self.column_tiles = {}
        self.saved_column_order = self.default_selected_order.copy()

        self.load_preferences()

        self.configure_root()
        self.load_app_icon()
        self.build_ui()
        self.refresh_selected_listbox()

        if DND_AVAILABLE:
            self.enable_drag_and_drop()

    def load_preferences(self):
        if not os.path.exists(self.preferences_file):
            return

        try:
            with open(self.preferences_file, "r", encoding="utf-8") as f:
                prefs = json.load(f)

            saved_selected = prefs.get("selected_columns", [])
            if isinstance(saved_selected, list):
                selected_set = set(saved_selected)
                for col in self.available_columns:
                    self.column_vars[col].set(col in selected_set if saved_selected else True)

            saved_order = prefs.get("column_order", [])
            if isinstance(saved_order, list) and saved_order:
                valid_order = [col for col in saved_order if col in self.available_columns]
                for col in self.default_selected_order:
                    if col not in valid_order:
                        valid_order.append(col)
                self.saved_column_order = valid_order

            saved_output = prefs.get("last_output_path", "")
            if isinstance(saved_output, str) and saved_output.strip():
                self.output_path.set(saved_output.strip())

        except Exception:
            pass

    def save_preferences(self):
        try:
            prefs = {
                "selected_columns": [col for col in self.available_columns if self.column_vars[col].get()],
                "column_order": list(self.selected_listbox.get(0, tk.END)) if hasattr(self, "selected_listbox") else self.saved_column_order,
                "last_output_path": self.output_path.get().strip(),
            }

            with open(self.preferences_file, "w", encoding="utf-8") as f:
                json.dump(prefs, f, indent=2)

            self.show_temp_status("Preferences saved.")

        except Exception:
            pass
    def show_temp_status(self, message, duration=2000):
        self.status_text.set(message)
        self.root.after(duration, lambda: self.status_text.set(""))

    def configure_root(self):
        self.root.configure(bg=self.colors["bg"])

    def load_app_icon(self):
        icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "nmltocsvicon.png")
        if os.path.exists(icon_path):
            try:
                self.app_icon = tk.PhotoImage(file=icon_path)
                self.root.iconphoto(True, self.app_icon)
            except Exception:
                pass

    def make_button(self, parent, text, command, width=14, primary=False):
        bg = self.colors["accent"] if primary else self.colors["panel_alt"]

        return tk.Button(
            parent,
            text=text,
            command=command,
            width=width,
            relief="flat",
            bd=0,
            bg=bg,
            fg=self.colors["text"],
            activebackground=self.colors["accent_hover"],
            activeforeground=self.colors["text"],
            font=("Segoe UI", 10, "bold"),
            padx=10,
            pady=8,
            cursor="hand2",
        )

    def make_label(self, parent, text, size=10, bold=False, color=None):
        return tk.Label(
            parent,
            text=text,
            bg=parent.cget("bg"),
            fg=color or self.colors["text"],
            font=("Segoe UI", size, "bold" if bold else "normal"),
            anchor="w",
            justify="left",
        )

    def make_entry(self, parent, textvariable):
        return tk.Entry(
            parent,
            textvariable=textvariable,
            bg=self.colors["entry"],
            fg=self.colors["text"],
            insertbackground=self.colors["text"],
            relief="flat",
            highlightthickness=1,
            highlightbackground=self.colors["border"],
            highlightcolor=self.colors["accent"],
            font=("Segoe UI", 10),
        )

    def build_ui(self):
        outer = tk.Frame(self.root, bg=self.colors["bg"])
        outer.pack(fill="both", expand=True, padx=18, pady=18)
        self.outer_frame = outer

        header = tk.Frame(outer, bg=self.colors["bg"])
        header.pack(fill="x", pady=(0, 14))

        self.make_label(header, "Traktor NML to CSV Converter", size=18, bold=True).pack(anchor="w")

        self.file_panel = tk.Frame(
            outer, bg=self.colors["panel"],
            highlightthickness=1, highlightbackground=self.colors["border"]
        )
        self.file_panel.pack(fill="x", pady=(0, 14))
        self.build_file_section(self.file_panel)

        middle_panel = tk.Frame(outer, bg=self.colors["bg"])
        middle_panel.pack(fill="both", expand=True, pady=(0, 14))

        self.left_panel = tk.Frame(
            middle_panel, bg=self.colors["panel"],
            highlightthickness=1, highlightbackground=self.colors["border"]
        )
        self.left_panel.pack(side="left", fill="both", expand=True, padx=(0, 7))

        self.right_panel = tk.Frame(
            middle_panel, bg=self.colors["panel"],
            highlightthickness=1, highlightbackground=self.colors["border"]
        )
        self.right_panel.pack(side="left", fill="both", expand=True, padx=(7, 0))

        self.build_columns_section(self.left_panel)
        self.build_order_section(self.right_panel)

        self.bottom_panel = tk.Frame(
            outer, bg=self.colors["panel"],
            highlightthickness=1, highlightbackground=self.colors["border"]
        )
        self.bottom_panel.pack(fill="x")

        self.build_bottom_section(self.bottom_panel)

    def build_file_section(self, parent):
        content = tk.Frame(parent, bg=self.colors["panel"])
        content.pack(fill="x", padx=14, pady=14)

        self.make_label(content, "Source NML File", size=11, bold=True).grid(row=0, column=0, sticky="w", pady=(0, 6))
        source_row = tk.Frame(content, bg=self.colors["panel"])
        source_row.grid(row=1, column=0, sticky="ew")
        source_row.grid_columnconfigure(0, weight=1)

        source_entry = self.make_entry(source_row, self.input_path)
        source_entry.grid(row=0, column=0, sticky="ew", padx=(0, 10), ipady=8)

        self.make_button(source_row, "Browse", self.browse_input, width=12).grid(row=0, column=1)

        self.make_label(content, "Target CSV File", size=11, bold=True).grid(row=2, column=0, sticky="w", pady=(14, 6))
        target_row = tk.Frame(content, bg=self.colors["panel"])
        target_row.grid(row=3, column=0, sticky="ew")
        target_row.grid_columnconfigure(0, weight=1)

        target_entry = self.make_entry(target_row, self.output_path)
        target_entry.grid(row=0, column=0, sticky="ew", padx=(0, 10), ipady=8)

        self.make_button(target_row, "Change", self.browse_output, width=12).grid(row=0, column=1)

        self.make_label(
            content,
            "Drag & drop your .nml file anywhere.",
            size=9,
            bold=True,
            color=self.colors["muted"]
        ).grid(row=4, column=0, columnspan=2, pady=(16, 0), sticky="w")

        if not DND_AVAILABLE:
            self.make_label(
                content,
                "Drag-and-drop support is unavailable until tkinterdnd2 is installed.",
                size=9,
                color=self.colors["muted"]
            ).grid(row=6, column=0, columnspan=2, sticky="w", pady=(6, 0))

        content.grid_columnconfigure(0, weight=1)

    def build_columns_section(self, parent):
        top = tk.Frame(parent, bg=self.colors["panel"])
        top.pack(fill="x", padx=14, pady=(8, 10))

        top_inner = tk.Frame(top, bg=self.colors["panel"])
        top_inner.pack(fill="x")

        self.make_label(top_inner, " Column Selection", size=11, bold=True).pack(side="left")

        actions = tk.Frame(top_inner, bg=self.colors["panel"])
        actions.pack(side="right")

        self.make_button(actions, "Select All", self.select_all_columns, width=12).pack(side="left", padx=(0, 8))
        self.make_button(actions, "Clear All", self.clear_all_columns, width=12).pack(side="left")

        container_outer = tk.Frame(parent, bg=self.colors["panel"])
        container_outer.pack(fill="both", expand=True, padx=14, pady=(0, 14))

        container = tk.Frame(container_outer, bg=self.colors["panel"])
        container.pack(fill="both", expand=True)

        columns_per_row = 3
        for idx, col in enumerate(self.available_columns):
            row = idx // columns_per_row
            column = idx % columns_per_row

            card = tk.Frame(
                container,
                bg=self.colors["tile_unselected"],
                highlightthickness=1,
                highlightbackground=self.colors["tile_border_unselected"],
                padx=10,
                pady=8,
                cursor="hand2",
            )
            card.grid(row=row, column=column, sticky="ew", padx=6, pady=6)

            label = tk.Label(
                card,
                text=col,
                bg=self.colors["tile_unselected"],
                fg=self.colors["text"],
                font=("Segoe UI", 9),
                anchor="w",
                justify="left",
                padx=3,
                pady=3,
                cursor="hand2",
            )
            label.pack(fill="x")

            card.bind("<Button-1>", lambda e, c=col: self.toggle_column_tile(c))
            label.bind("<Button-1>", lambda e, c=col: self.toggle_column_tile(c))

            card.bind("<Enter>", lambda e, c=col: self.set_tile_hover(c, True))
            card.bind("<Leave>", lambda e, c=col: self.set_tile_hover(c, False))
            label.bind("<Enter>", lambda e, c=col: self.set_tile_hover(c, True))
            label.bind("<Leave>", lambda e, c=col: self.set_tile_hover(c, False))

            self.column_tiles[col] = {
                "frame": card,
                "label": label,
            }

            self.update_column_tile(col)

        for c in range(columns_per_row):
            container.grid_columnconfigure(c, weight=1)

    def build_order_section(self, parent):
        top = tk.Frame(parent, bg=self.colors["panel"])
        top.pack(fill="x", padx=14, pady=(14, 10))

        self.make_label(top, "Column Order", size=11, bold=True).pack(side="left")

        body = tk.Frame(parent, bg=self.colors["panel"])
        body.pack(fill="both", expand=True, padx=14, pady=(0, 14))

        list_frame = tk.Frame(body, bg=self.colors["panel"])
        list_frame.pack(side="left", fill="both", expand=True)

        self.selected_listbox = tk.Listbox(
            list_frame,
            bg=self.colors["listbox_bg"],
            fg=self.colors["text"],
            selectbackground=self.colors["listbox_select"],
            selectforeground=self.colors["text"],
            relief="flat",
            bd=0,
            highlightthickness=1,
            highlightbackground=self.colors["border"],
            font=("Segoe UI", 10),
            activestyle="none",
            exportselection=False,
        )
        self.selected_listbox.pack(fill="both", expand=True)

        self.selected_listbox.bind("<Button-1>", self.on_listbox_drag_start)
        self.selected_listbox.bind("<B1-Motion>", self.on_listbox_drag_motion)
        self.selected_listbox.bind("<ButtonRelease-1>", self.on_listbox_drag_end)

        side_frame = tk.Frame(body, bg=self.colors["panel"])
        side_frame.pack(side="left", fill="y", padx=(12, 0))

        self.make_button(side_frame, "Reset Order", self.reset_column_order, width=12).pack(pady=(0, 8))

        self.make_label(
            side_frame,
            "Drag items\nto reorder",
            size=10,
            color=self.colors["muted"]
        ).pack(anchor="n", pady=(8, 0))

    def build_bottom_section(self, parent):
        content = tk.Frame(parent, bg=self.colors["panel"])
        content.pack(fill="x", padx=14, pady=14)

        left = tk.Frame(content, bg=self.colors["panel"])
        left.pack(side="left", fill="x", expand=True)

        self.make_label(left, "Status", size=11, bold=True).pack(anchor="w")
        tk.Label(
            left,
            textvariable=self.status_text,
            bg=self.colors["panel"],
            fg=self.colors["success"],
            justify="left",
            wraplength=620,
            anchor="w",
            font=("Segoe UI", 10),
        ).pack(anchor="w", pady=(6, 0))

        right = tk.Frame(content, bg=self.colors["panel"])
        right.pack(side="right")

        self.make_button(right, "Export CSV", self.export_csv, width=16, primary=True).pack()

    def enable_drag_and_drop(self):
        widgets = [
            self.root,
            self.outer_frame,
            self.file_panel,
            self.left_panel,
            self.right_panel,
            self.bottom_panel,
        ]

        for widget in widgets:
            widget.drop_target_register(DND_FILES)
            widget.dnd_bind("<<Drop>>", self.handle_drop)
            widget.dnd_bind("<<DragEnter>>", self.handle_drag_enter)
            widget.dnd_bind("<<DragLeave>>", self.handle_drag_leave)

    def set_drag_highlight(self, active):
        border_color = self.colors["drag_glow"] if active else self.colors["border"]

        panels = [
            getattr(self, "file_panel", None),
            getattr(self, "left_panel", None),
            getattr(self, "right_panel", None),
            getattr(self, "bottom_panel", None),
        ]

        for panel in panels:
            if panel is not None:
                panel.configure(highlightbackground=border_color)

    def handle_drag_enter(self, event):
        self.set_drag_highlight(True)
        return event.action

    def handle_drag_leave(self, event):
        self.set_drag_highlight(False)

    def handle_drop(self, event):
        self.set_drag_highlight(False)

        dropped = event.data.strip()

        if dropped.startswith("{") and dropped.endswith("}"):
            dropped = dropped[1:-1]

        if "} {" in dropped:
            dropped = dropped.split("} {")[0].strip("{}")

        dropped = dropped.strip('"')

        if not dropped.lower().endswith(".nml"):
            messagebox.showerror("Invalid File", "Please drop a .nml file.")
            return

        if not os.path.isfile(dropped):
            messagebox.showerror("Invalid File", "The dropped file does not exist.")
            return

        self.input_path.set(dropped)

        default_name = os.path.splitext(os.path.basename(dropped))[0] + ".csv"
        default_path = os.path.join(os.path.dirname(dropped), default_name)
        self.output_path.set(default_path)
        self.save_preferences()

        self.status_text.set(f"Loaded NML file: {dropped}")

    def browse_input(self):
        file_path = filedialog.askopenfilename(
            title="Select Traktor NML File",
            filetypes=[("Traktor NML Files", "*.nml"), ("All Files", "*.*")]
        )
        if file_path:
            self.input_path.set(file_path)
            if not self.output_path.get():
                default_name = os.path.splitext(os.path.basename(file_path))[0] + ".csv"
                default_path = os.path.join(os.path.dirname(file_path), default_name)
                self.output_path.set(default_path)
            self.save_preferences()

    def browse_output(self):
        file_path = filedialog.asksaveasfilename(
            title="Save CSV As",
            defaultextension=".csv",
            filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")]
        )
        if file_path:
            self.output_path.set(file_path)
            self.save_preferences()

    def select_all_columns(self):
        for var in self.column_vars.values():
            var.set(True)
        self.update_all_column_tiles()
        self.refresh_selected_listbox()
        self.save_preferences()

    def clear_all_columns(self):
        for var in self.column_vars.values():
            var.set(False)
        self.update_all_column_tiles()
        self.refresh_selected_listbox()
        self.save_preferences()

    def get_selected_columns(self):
        return list(self.selected_listbox.get(0, tk.END))

    def refresh_selected_listbox(self):
        current_order = list(self.selected_listbox.get(0, tk.END)) if hasattr(self, "selected_listbox") else []
        checked_columns = [col for col in self.available_columns if self.column_vars[col].get()]

        if not current_order:
            base_order = self.saved_column_order if self.saved_column_order else self.default_selected_order
            new_order = [col for col in base_order if col in checked_columns]
            for col in checked_columns:
                if col not in new_order:
                    new_order.append(col)
        else:
            new_order = [col for col in current_order if col in checked_columns]
            for col in checked_columns:
                if col not in new_order:
                    new_order.append(col)

        if hasattr(self, "selected_listbox"):
            current_selection = self.selected_listbox.curselection()
            selected_name = None
            if current_selection:
                try:
                    selected_name = self.selected_listbox.get(current_selection[0])
                except Exception:
                    selected_name = None

            self.selected_listbox.delete(0, tk.END)
            for col in new_order:
                self.selected_listbox.insert(tk.END, col)

            self.saved_column_order = new_order.copy()

            if selected_name and selected_name in new_order:
                idx = new_order.index(selected_name)
                self.selected_listbox.selection_set(idx)
                self.selected_listbox.activate(idx)

    def on_listbox_drag_start(self, event):
        index = self.selected_listbox.nearest(event.y)
        if index >= 0:
            self.drag_index = index
            self.dragging_list_item = True
            self.selected_listbox.selection_clear(0, tk.END)
            self.selected_listbox.selection_set(index)
            self.selected_listbox.activate(index)

    def on_listbox_drag_motion(self, event):
        if not self.dragging_list_item or self.drag_index is None:
            return

        new_index = self.selected_listbox.nearest(event.y)
        if new_index == self.drag_index or new_index < 0:
            return

        items = list(self.selected_listbox.get(0, tk.END))
        item = items.pop(self.drag_index)
        items.insert(new_index, item)

        self.selected_listbox.delete(0, tk.END)
        for list_item in items:
            self.selected_listbox.insert(tk.END, list_item)

        self.selected_listbox.selection_clear(0, tk.END)
        self.selected_listbox.selection_set(new_index)
        self.selected_listbox.activate(new_index)

        self.drag_index = new_index
        self.saved_column_order = items.copy()
        self.save_preferences()

    def on_listbox_drag_end(self, event):
        self.drag_index = None
        self.dragging_list_item = False

    def reset_column_order(self):
        selected = {col for col in self.available_columns if self.column_vars[col].get()}
        ordered = [col for col in self.default_selected_order if col in selected]

        self.selected_listbox.delete(0, tk.END)
        for item in ordered:
            self.selected_listbox.insert(tk.END, item)

        self.saved_column_order = ordered.copy()
        self.save_preferences()

    def export_csv(self):
        input_file = self.input_path.get().strip()
        output_file = self.output_path.get().strip()
        selected_columns = self.get_selected_columns()

        if not input_file:
            messagebox.showerror("Missing Input", "Please select an NML file.")
            return

        if not os.path.isfile(input_file):
            messagebox.showerror("Invalid Input", "The selected NML file does not exist.")
            return

        if not output_file:
            messagebox.showerror("Missing Output", "Please choose where to save the CSV.")
            return

        if not selected_columns:
            messagebox.showerror("No Columns Selected", "Please select at least one column to export.")
            return

        try:
            self.status_text.set("Reading NML file...")
            self.root.update_idletasks()

            rows = extract_entries(input_file)

            self.status_text.set("Writing CSV...")
            self.root.update_idletasks()

            write_csv(rows, output_file, selected_columns)
            self.save_preferences()

            self.status_text.set(f"Done. Exported {len(rows)} entries to: {output_file}")
            messagebox.showinfo("Export Complete", f"Exported {len(rows)} entries successfully.")
        except Exception as e:
            self.status_text.set("Export failed.")
            messagebox.showerror("Error", f"Failed to export CSV:\n\n{e}")


def create_root():
    if DND_AVAILABLE:
        return TkinterDnD.Tk()
    return tk.Tk()


def main():
    root = create_root()
    app = TraktorExporterApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
