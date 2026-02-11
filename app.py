from __future__ import annotations

from datetime import date, datetime, timedelta
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
import os
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import webbrowser
from typing import Callable, Sequence
from tkinter import font

from constants import ASSET_STATUSES, ASSET_TYPES, BUSINESSES, DEFAULT_LOW_STOCK_LEVEL
from db import (
    add_asset,
    add_asset_acquisition,
    add_product,
    add_user,
    delete_asset,
    duplicate_asset,
    delete_asset_acquisition,
    delete_in_log,
    delete_out_log,
    delete_product,
    delete_user,
    add_asset_status,
    add_in_breakdown,
    delete_asset_status,
    delete_in_breakdown,
    list_assets_for_export,
    list_asset_acquisitions,
    list_asset_acquisitions_report,
    list_asset_statuses_report,
    list_expiry_dates_report,
    list_in_logs_report,
    list_out_logs_report,
    get_perishable_report,
    get_perishable_stock,
    init_db,
    list_assets,
    list_asset_statuses,
    list_in_breakdown,
    list_in_out_logs,
    list_expiry_dates,
    list_products,
    list_users,
    record_in,
    record_out,
    update_asset,
    update_asset_acquisition,
    update_in_log,
    update_out_log,
    update_product,
    update_user,
    verify_user,
)
from export_utils import export_to_excel, export_to_jpg, export_to_pdf

try:
    from tkcalendar import DateEntry  # type: ignore
except Exception:
    DateEntry = None

try:
    from PIL import Image, ImageTk  # type: ignore
except Exception:
    Image = None
    ImageTk = None


def _to_decimal(value: object) -> Decimal:
    if value is None or value == "":
        return Decimal("0")
    if isinstance(value, Decimal):
        return value
    try:
        return Decimal(str(value))
    except (InvalidOperation, ValueError, TypeError):
        return Decimal("0")


def _format_money(value: object) -> str:
    dec = _to_decimal(value).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    return format(dec, ".2f")


def _format_number(value: object) -> str:
    try:
        num = float(value)
    except Exception:
        return str(value)
    if num.is_integer():
        return f"{int(num):,}"
    return f"{num:,.2f}"


def _format_php(value: object) -> str:
    return f"₱{_format_number(value)}"


def _load_preview_image(path: str, size: tuple[int, int] = (240, 180)) -> ImageTk.PhotoImage | None:
    if not path:
        return None
    if Image is None or ImageTk is None:
        return None
    try:
        img = Image.open(path)
        img.thumbnail(size)
        return ImageTk.PhotoImage(img)
    except Exception:
        return None

TAB_LOGOS = {
    "Unica Perishable": "assets/unica_perishable.png",
    "Unica Non-Perishable": "assets/unica_non_perishable.png",
    "HDN Warehouse": "assets/hdn_warehouse.png",
    "HDN Plants": "assets/hdn_plants.png",
}
LOGO_MAX_SIZE = (140, 100)
GITHUB_RELEASES_URL = "https://github.com/amanerdc/aman-inventory/releases"

UI_COLORS = {
    "bg": "#F5F7FB",
    "panel": "#FFFFFF",
    "border": "#D6DEEA",
    "text": "#1F2937",
    "muted": "#6B7280",
    "accent": "#2563EB",
    "accent_dark": "#1E40AF",
    "accent_light": "#E0E7FF",
    "header": "#E9EEF8",
}


def apply_theme(root: tk.Tk) -> None:
    style = ttk.Style(root)
    if "clam" in style.theme_names():
        style.theme_use("clam")

    root.configure(bg=UI_COLORS["bg"])

    default_font = font.nametofont("TkDefaultFont")
    default_font.configure(family="Segoe UI", size=10)
    text_font = font.nametofont("TkTextFont")
    text_font.configure(family="Segoe UI", size=10)

    style.configure("TFrame", background=UI_COLORS["bg"])
    style.configure("Card.TFrame", background=UI_COLORS["panel"])
    style.configure("TLabel", background=UI_COLORS["bg"], foreground=UI_COLORS["text"])
    style.configure("Muted.TLabel", background=UI_COLORS["bg"], foreground=UI_COLORS["muted"])
    style.configure("Panel.TLabel", background=UI_COLORS["panel"], foreground=UI_COLORS["text"])
    style.configure("PanelMuted.TLabel", background=UI_COLORS["panel"], foreground=UI_COLORS["muted"])
    style.configure(
        "Title.TLabel",
        background=UI_COLORS["bg"],
        foreground=UI_COLORS["text"],
        font=("Segoe UI", 14, "bold"),
    )
    style.configure(
        "TLabelFrame",
        background=UI_COLORS["bg"],
        foreground=UI_COLORS["text"],
        padding=8,
    )
    style.configure(
        "TLabelFrame.Label",
        background=UI_COLORS["bg"],
        foreground=UI_COLORS["text"],
        font=("Segoe UI", 10, "bold"),
    )

    style.configure(
        "TButton",
        background=UI_COLORS["panel"],
        foreground=UI_COLORS["text"],
        bordercolor=UI_COLORS["border"],
        borderwidth=1,
        padding=(12, 6),
        relief="flat",
    )
    style.map(
        "TButton",
        background=[("active", "#EEF2F7"), ("pressed", "#E6ECF5")],
        foreground=[("disabled", "#9CA3AF")],
        bordercolor=[("active", "#C7D2E5")],
    )

    style.configure(
        "Accent.TButton",
        background=UI_COLORS["accent"],
        foreground="white",
        borderwidth=0,
        padding=(12, 6),
        relief="flat",
    )
    style.map(
        "Accent.TButton",
        background=[("active", UI_COLORS["accent_dark"]), ("pressed", UI_COLORS["accent_dark"])],
    )

    style.configure(
        "TEntry",
        fieldbackground=UI_COLORS["panel"],
        foreground=UI_COLORS["text"],
        bordercolor=UI_COLORS["border"],
        padding=6,
        relief="flat",
    )
    style.configure(
        "TCombobox",
        fieldbackground=UI_COLORS["panel"],
        foreground=UI_COLORS["text"],
        bordercolor=UI_COLORS["border"],
        padding=6,
        relief="flat",
    )

    style.configure(
        "Treeview",
        background=UI_COLORS["panel"],
        fieldbackground=UI_COLORS["panel"],
        foreground=UI_COLORS["text"],
        rowheight=30,
        bordercolor=UI_COLORS["border"],
        borderwidth=1,
    )
    style.configure(
        "Treeview.Heading",
        background=UI_COLORS["header"],
        foreground=UI_COLORS["text"],
        relief="flat",
        padding=6,
        font=("Segoe UI", 10, "bold"),
    )
    style.map(
        "Treeview",
        background=[("selected", UI_COLORS["accent_light"])],
        foreground=[("selected", UI_COLORS["text"])],
    )

    style.configure(
        "TNotebook",
        background=UI_COLORS["bg"],
        borderwidth=0,
    )
    style.configure(
        "TNotebook.Tab",
        background=UI_COLORS["bg"],
        foreground=UI_COLORS["muted"],
        padding=(12, 6),
    )
    style.map(
        "TNotebook.Tab",
        background=[("selected", UI_COLORS["panel"])],
        foreground=[("selected", UI_COLORS["text"])],
    )


def enable_smooth_resize(root: tk.Tk) -> None:
    def _schedule(_event: tk.Event | None = None) -> None:
        if hasattr(root, "_resize_after"):
            try:
                root.after_cancel(root._resize_after)
            except Exception:
                pass
        root._resize_after = root.after(33, root.update_idletasks)

    root.bind("<Configure>", _schedule, add="+")


def configure_input_shortcuts(root: tk.Tk) -> None:
    def _generate(event: tk.Event, sequence: str) -> str:
        try:
            event.widget.event_generate(sequence)
        except Exception:
            return "break"
        return "break"

    def _undo(event: tk.Event) -> str:
        widget = event.widget
        try:
            widget.event_generate("<<Undo>>")
            return "break"
        except Exception:
            pass
        try:
            widget.tk.call(widget._w, "edit", "undo")
        except Exception:
            return "break"
        return "break"

    def _select_all(event: tk.Event) -> str:
        widget = event.widget
        try:
            widget.select_range(0, "end")
            widget.icursor("end")
            return "break"
        except Exception:
            pass
        try:
            widget.tag_add("sel", "1.0", "end-1c")
            return "break"
        except Exception:
            return "break"

    def _enable_undo(event: tk.Event) -> None:
        try:
            event.widget.configure(undo=True, autoseparators=True, maxundo=50)
        except Exception:
            return

    for class_name in ("Entry", "TEntry", "Text", "TCombobox"):
        root.bind_class(class_name, "<FocusIn>", _enable_undo, add="+")
        root.bind_class(class_name, "<Control-c>", lambda e: _generate(e, "<<Copy>>"), add="+")
        root.bind_class(class_name, "<Control-x>", lambda e: _generate(e, "<<Cut>>"), add="+")
        root.bind_class(class_name, "<Control-v>", lambda e: _generate(e, "<<Paste>>"), add="+")
        root.bind_class(class_name, "<Control-a>", _select_all, add="+")
        root.bind_class(class_name, "<Control-z>", _undo, add="+")


def make_date_entry(parent: tk.Widget, textvariable: tk.StringVar) -> tk.Widget:
    if DateEntry is not None:
        return DateEntry(parent, textvariable=textvariable, date_pattern="yyyy-mm-dd")
    return ttk.Entry(parent, textvariable=textvariable)


PHOTO_THUMBNAIL_SIZE = (64, 64)
PHOTO_ROW_HEIGHT = 72


def configure_photo_treeview_style(widget: tk.Widget) -> None:
    style = ttk.Style(widget)
    style.configure("Photo.Treeview", rowheight=PHOTO_ROW_HEIGHT)


def build_treeview(
    parent: tk.Widget,
    columns: Sequence[str],
    headings: Sequence[str],
    image_heading: str | None = None,
    image_width: int = 80,
    style: str | None = None,
) -> ttk.Treeview:
    container = ttk.Frame(parent, style="Card.TFrame", padding=8)
    container.pack(fill="both", expand=True, padx=8, pady=8)

    show_mode = "headings" if image_heading is None else "tree headings"
    tree = ttk.Treeview(container, columns=columns, show=show_mode, style=style)
    if image_heading is not None:
        tree.heading("#0", text=image_heading)
        tree.column("#0", width=image_width, anchor="center")
    for col, head in zip(columns, headings, strict=False):
        tree.heading(col, text=head)
        tree.column(col, width=130, anchor="w")

    yscroll = ttk.Scrollbar(container, orient="vertical", command=tree.yview)
    xscroll = ttk.Scrollbar(container, orient="horizontal", command=tree.xview)
    tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)

    tree.grid(row=0, column=0, sticky="nsew")
    yscroll.grid(row=0, column=1, sticky="ns")
    xscroll.grid(row=1, column=0, sticky="ew", columnspan=2)

    container.columnconfigure(0, weight=1)
    container.rowconfigure(0, weight=1)
    container.rowconfigure(1, weight=0)

    bind_treeview_copy(tree)
    return tree


def bind_treeview_copy(tree: ttk.Treeview) -> None:
    def _copy_cell(event: tk.Event) -> str | None:
        row_id = tree.identify_row(event.y)
        col_id = tree.identify_column(event.x)
        if not row_id or not col_id:
            return None
        if col_id == "#0":
            value = tree.item(row_id, "text") or ""
        else:
            try:
                idx = int(col_id.replace("#", "")) - 1
            except ValueError:
                return None
            values = tree.item(row_id, "values")
            if idx >= len(values):
                return None
            value = values[idx]
        tree.clipboard_clear()
        tree.clipboard_append(str(value))
        return "break"

    tree.bind("<Double-1>", _copy_cell, add="+")


def configure_context_menu(root: tk.Tk) -> None:
    menu = tk.Menu(root, tearoff=0)

    def _do(widget: tk.Widget, sequence: str) -> None:
        try:
            widget.event_generate(sequence)
        except Exception:
            return

    def _select_all_widget(widget: tk.Widget) -> None:
        try:
            widget.select_range(0, "end")
            widget.icursor("end")
            return
        except Exception:
            pass
        try:
            widget.tag_add("sel", "1.0", "end-1c")
        except Exception:
            return

    def _show(event: tk.Event) -> str:
        widget = event.widget
        menu.delete(0, "end")
        menu.add_command(label="Cut", command=lambda: _do(widget, "<<Cut>>"))
        menu.add_command(label="Copy", command=lambda: _do(widget, "<<Copy>>"))
        menu.add_command(label="Paste", command=lambda: _do(widget, "<<Paste>>"))
        menu.add_separator()
        menu.add_command(label="Select All", command=lambda: _select_all_widget(widget))
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()
        return "break"

    for class_name in ("Entry", "TEntry", "Text", "TCombobox"):
        root.bind_class(class_name, "<Button-3>", _show, add="+")
        root.bind_class(class_name, "<Button-2>", _show, add="+")


def set_tree_image(tree: ttk.Treeview, iid: str, path: str | None) -> None:
    if not path:
        return
    if Image is None or ImageTk is None:
        return
    try:
        img = Image.open(path)
        img.thumbnail(PHOTO_THUMBNAIL_SIZE)
        photo = ImageTk.PhotoImage(img)
    except Exception:
        return
    if not hasattr(tree, "_img_refs"):
        tree._img_refs = {}
    tree._img_refs[iid] = photo
    tree.item(iid, image=photo)


def make_readonly_text(parent: tk.Widget, content: str, height: int = 8) -> tk.Text:
    widget = tk.Text(
        parent,
        height=height,
        wrap="word",
        background=UI_COLORS["panel"],
        foreground=UI_COLORS["text"],
        borderwidth=1,
        relief="solid",
    )
    widget.insert("1.0", content)
    widget.configure(state="disabled")
    widget.pack(fill="x", expand=True)
    return widget


def load_logo_image(path: str) -> ImageTk.PhotoImage | None:
    if Image is None or ImageTk is None:
        return None
    if not path or not os.path.exists(path):
        return None
    try:
        img = Image.open(path)
        img.thumbnail(LOGO_MAX_SIZE)
        return ImageTk.PhotoImage(img)
    except Exception:
        return None


def _safe_date(value: object) -> date | None:
    if isinstance(value, date):
        return value
    if isinstance(value, str):
        try:
            return datetime.strptime(value, "%Y-%m-%d").date()
        except ValueError:
            return None
    return None


def _str_or_empty(value: object, default: str = "") -> str:
    if value is None:
        return default
    return str(value)


class LoginWindow:
    def __init__(self, root: tk.Tk, on_success: Callable[[dict], None]) -> None:
        self.root = root
        self.on_success = on_success
        self.root.title("Aman Inventory - Login")
        self.root.resizable(False, False)

        outer = ttk.Frame(self.root, padding=16)
        outer.pack(fill="both", expand=True)
        self.frame = outer

        card = ttk.Frame(outer, style="Card.TFrame", padding=16)
        card.pack(fill="both", expand=True)

        ttk.Label(card, text="Welcome back", style="Title.TLabel").grid(row=0, column=0, columnspan=2, sticky="w")
        ttk.Label(card, text="Sign in to continue", style="Muted.TLabel").grid(
            row=1, column=0, columnspan=2, sticky="w", pady=(0, 10)
        )

        ttk.Label(card, text="Username").grid(row=2, column=0, sticky="w")
        self.username = ttk.Entry(card)
        self.username.grid(row=2, column=1, sticky="ew", pady=6)

        ttk.Label(card, text="Password").grid(row=3, column=0, sticky="w")
        self.password = ttk.Entry(card, show="*")
        self.password.grid(row=3, column=1, sticky="ew", pady=6)

        card.columnconfigure(1, weight=1)

        login_btn = ttk.Button(card, text="Login", style="Accent.TButton", command=self.login)
        login_btn.grid(row=4, column=0, columnspan=2, pady=(10, 0), sticky="ew")

        self.root.update_idletasks()
        req_w = self.root.winfo_reqwidth()
        req_h = self.root.winfo_reqheight()
        self.root.geometry(f"{req_w}x{req_h}")

    def login(self) -> None:
        username = self.username.get().strip()
        password = self.password.get().strip()
        if not username or not password:
            messagebox.showwarning("Missing", "Please enter username and password.")
            return
        ok, user = verify_user(username, password)
        if not ok or not user:
            messagebox.showerror("Login failed", "Invalid username or password.")
            return
        self.frame.destroy()
        self.on_success(user)


class ProductForm(tk.Toplevel):
    def __init__(self, parent: tk.Widget, title: str, on_save: Callable[[dict], None], initial: dict | None = None) -> None:
        super().__init__(parent)
        self.title(title)
        self.resizable(False, False)
        self.on_save = on_save
        self.initial = initial or {}
        self.transient(parent)
        self.grab_set()

        frame = ttk.Frame(self, padding=12)
        frame.pack(fill="both", expand=True)

        self.vars = {
            "name": tk.StringVar(value=self.initial.get("name", "")),
            "category": tk.StringVar(value=self.initial.get("category", "")),
            "unit": tk.StringVar(value=self.initial.get("unit", "unit")),
            "opening_stock": tk.StringVar(value=_str_or_empty(self.initial.get("opening_stock"), "0")),
            "low_stock_level": tk.StringVar(value=_str_or_empty(self.initial.get("low_stock_level"), str(DEFAULT_LOW_STOCK_LEVEL))),
            "photo_path": tk.StringVar(value=self.initial.get("photo_path", "")),
        }

        self._row(frame, 0, "Name", "name")
        self._row(frame, 1, "Category", "category")
        self._row(frame, 2, "Unit", "unit")
        self._row(frame, 3, "Opening Stock", "opening_stock")
        self._row(frame, 4, "Low Stock Level", "low_stock_level")

        ttk.Label(frame, text="Photo").grid(row=5, column=0, sticky="w", pady=4)
        ttk.Entry(frame, textvariable=self.vars["photo_path"], width=35).grid(row=5, column=1, sticky="ew")
        ttk.Button(frame, text="Browse", command=self._browse_photo).grid(row=5, column=2, padx=4)

        self.preview_label = ttk.Label(frame, text="No image selected")
        self.preview_label.grid(row=6, column=0, columnspan=3, pady=6)
        self.preview_image = None
        self._load_preview(self.vars["photo_path"].get().strip())

        ttk.Button(frame, text="Save", command=self._save).grid(row=7, column=0, columnspan=3, pady=8)

        frame.columnconfigure(1, weight=1)

    def _row(self, frame: ttk.Frame, row: int, label: str, key: str) -> None:
        ttk.Label(frame, text=label).grid(row=row, column=0, sticky="w", pady=4)
        ttk.Entry(frame, textvariable=self.vars[key]).grid(row=row, column=1, columnspan=2, sticky="ew")

    def _load_preview(self, path: str) -> None:
        if not path:
            self.preview_label.configure(text="No image selected", image="")
            return
        if Image is None or ImageTk is None:
            self.preview_label.configure(text="Preview requires Pillow (pip install pillow)", image="")
            return
        try:
            img = Image.open(path)
            img.thumbnail((240, 180))
            self.preview_image = ImageTk.PhotoImage(img)
            self.preview_label.configure(image=self.preview_image, text="")
        except Exception:
            self.preview_label.configure(text="Unable to load preview", image="")

    def _browse_photo(self) -> None:
        path = filedialog.askopenfilename(parent=self, title="Select Photo")
        if path:
            self.vars["photo_path"].set(path)
            self._load_preview(path)

    def _save(self) -> None:
        name = self.vars["name"].get().strip()
        category = self.vars["category"].get().strip()
        unit = self.vars["unit"].get().strip()
        opening = self.vars["opening_stock"].get().strip()
        low_stock = self.vars["low_stock_level"].get().strip()
        if not name or not category or not unit:
            messagebox.showwarning("Missing", "Name, category, and unit are required.")
            return
        try:
            opening_val = float(opening)
            low_stock_val = float(low_stock)
        except ValueError:
            messagebox.showwarning("Invalid", "Opening stock and low stock level must be numbers.")
            return
        self.on_save(
            {
                "name": name,
                "category": category,
                "unit": unit,
                "opening_stock": opening_val,
                "low_stock_level": low_stock_val,
                "photo_path": self.vars["photo_path"].get().strip() or None,
            }
        )
        self.destroy()


class InOutForm(tk.Toplevel):
    def __init__(
        self,
        parent: tk.Widget,
        title: str,
        products: Sequence[tuple[int, str]],
        on_save: Callable[[dict], None],
        default_product: str | None = None,
    ) -> None:
        super().__init__(parent)
        self.title(title)
        self.resizable(False, False)
        self.on_save = on_save
        self.transient(parent)
        self.grab_set()

        frame = ttk.Frame(self, padding=12)
        frame.pack(fill="both", expand=True)

        self.product_var = tk.StringVar()
        self.date_var = tk.StringVar(value=date.today().strftime("%Y-%m-%d"))
        self.time_var = tk.StringVar()
        self.qty_var = tk.StringVar()

        ttk.Label(frame, text="Product").grid(row=0, column=0, sticky="w", pady=4)
        self.all_products = [p[1] for p in products]
        self.product_combo = ttk.Combobox(frame, textvariable=self.product_var, values=self.all_products, state="normal")
        self.product_combo.grid(row=0, column=1, sticky="ew")
        self.product_combo.bind("<KeyRelease>", self._filter_products)
        if default_product:
            self.product_var.set(default_product)

        self.is_in = "IN" in title.upper()
        if self.is_in:
            ttk.Label(frame, text="Delivery Date").grid(row=1, column=0, sticky="w", pady=4)
            make_date_entry(frame, self.date_var).grid(row=1, column=1, sticky="ew")
            ttk.Label(frame, text="Quantity").grid(row=2, column=0, sticky="w", pady=4)
            ttk.Entry(frame, textvariable=self.qty_var).grid(row=2, column=1, sticky="ew")
            row_offset = 2
        else:
            ttk.Label(frame, text="Out Date").grid(row=1, column=0, sticky="w", pady=4)
            make_date_entry(frame, self.date_var).grid(row=1, column=1, sticky="ew")
            ttk.Label(frame, text="Out Time (HH:MM)").grid(row=2, column=0, sticky="w", pady=4)
            ttk.Entry(frame, textvariable=self.time_var).grid(row=2, column=1, sticky="ew")
            row_offset = 2

        if not self.is_in:
            ttk.Label(frame, text="Quantity").grid(row=row_offset + 1, column=0, sticky="w", pady=4)
            ttk.Entry(frame, textvariable=self.qty_var).grid(row=row_offset + 1, column=1, sticky="ew")

        ttk.Button(frame, text="Save", command=self._save).grid(row=row_offset + 2, column=0, columnspan=2, pady=8)
        frame.columnconfigure(1, weight=1)
        self.products = products

    def _filter_products(self, _event: tk.Event) -> None:
        value = self.product_var.get().strip().lower()
        if not value:
            self.product_combo["values"] = self.all_products
            return
        filtered = [p for p in self.all_products if value in p.lower()]
        self.product_combo["values"] = filtered if filtered else self.all_products

    # expiry rows moved to breakdown table

    def _save(self) -> None:
        name = self.product_var.get().strip()
        if not name:
            messagebox.showwarning("Missing", "Please select a product.")
            return
        product_id = next((pid for pid, pname in self.products if pname == name), None)
        if product_id is None:
            messagebox.showwarning("Invalid", "Select a product from the list.")
            return
        date_val = self.date_var.get().strip()
        if not date_val:
            messagebox.showwarning("Missing", "Date is required.")
            return
        payload = {
            "product_id": product_id,
            "date": date_val,
        }
        if self.is_in:
            qty = self.qty_var.get().strip()
            if not qty:
                messagebox.showwarning("Missing", "Quantity is required.")
                return
            try:
                qty_val = float(qty)
            except ValueError:
                messagebox.showwarning("Invalid", "Quantity must be a number.")
                return
            payload["quantity"] = qty_val
        else:
            qty = self.qty_var.get().strip()
            if not qty:
                messagebox.showwarning("Missing", "Quantity is required.")
                return
            try:
                qty_val = float(qty)
            except ValueError:
                messagebox.showwarning("Invalid", "Quantity must be a number.")
                return
            payload["quantity"] = qty_val
            if self.time_var.get().strip():
                payload["time"] = self.time_var.get().strip()
        self.on_save(payload)
        self.destroy()


class LogEditForm(tk.Toplevel):
    def __init__(self, parent: tk.Widget, kind: str, log: dict, on_save: Callable[[dict], None]) -> None:
        super().__init__(parent)
        self.title(f"Edit {kind.upper()} Log")
        self.resizable(False, False)
        self.on_save = on_save
        self.kind = kind
        self.log = log
        self.transient(parent)
        self.grab_set()

        frame = ttk.Frame(self, padding=12)
        frame.pack(fill="both", expand=True)

        self.date_var = tk.StringVar(value=log["delivery_date"] if kind == "in" else log["out_date"])
        self.time_var = tk.StringVar(value=log.get("out_time", ""))
        self.qty_var = tk.StringVar(value=str(log["quantity"]))

        ttk.Label(frame, text="Date").grid(row=0, column=0, sticky="w", pady=4)
        make_date_entry(frame, self.date_var).grid(row=0, column=1, sticky="ew")

        if kind != "in":
            ttk.Label(frame, text="Time (HH:MM)").grid(row=1, column=0, sticky="w", pady=4)
            ttk.Entry(frame, textvariable=self.time_var).grid(row=1, column=1, sticky="ew")
            row_offset = 1
        else:
            row_offset = 0

        ttk.Label(frame, text="Quantity").grid(row=row_offset + 1, column=0, sticky="w", pady=4)
        ttk.Entry(frame, textvariable=self.qty_var).grid(row=row_offset + 1, column=1, sticky="ew")

        ttk.Button(frame, text="Save", command=self._save).grid(row=row_offset + 2, column=0, columnspan=2, pady=8)
        frame.columnconfigure(1, weight=1)

    def _save(self) -> None:
        try:
            qty_val = float(self.qty_var.get().strip())
        except ValueError:
            messagebox.showwarning("Invalid", "Quantity must be a number.")
            return
        payload = {
            "id": self.log["id"],
            "date": self.date_var.get().strip(),
            "quantity": qty_val,
        }
        if self.kind != "in":
            payload["time"] = self.time_var.get().strip()
        self.on_save(payload)
        self.destroy()


class AssetForm(tk.Toplevel):
    def __init__(
        self,
        parent: tk.Widget,
        title: str,
        business: str,
        inventory_type: str,
        on_save: Callable[[dict], None],
        initial: dict | None = None,
    ) -> None:
        super().__init__(parent)
        self.title(title)
        self.resizable(False, False)
        self.on_save = on_save
        self.initial = initial or {}
        self.business = business
        self.inventory_type = inventory_type
        self.is_new = not bool(self.initial.get("id"))
        self.transient(parent)
        self.grab_set()

        frame = ttk.Frame(self, padding=12)
        frame.pack(fill="both", expand=True)

        self.vars = {
            "picture_path": tk.StringVar(value=self.initial.get("picture_path", "")),
            "name": tk.StringVar(value=self.initial.get("name", "")),
            "brand": tk.StringVar(value=self.initial.get("brand", "")),
            "model": tk.StringVar(value=self.initial.get("model", "")),
            "specifications": tk.StringVar(value=self.initial.get("specifications", "")),
            "series_number": tk.StringVar(value=self.initial.get("series_number", "")),
            "quantity": tk.StringVar(value=_str_or_empty(self.initial.get("quantity"), "1")),
            "location": tk.StringVar(value=self.initial.get("location", "")),
            "type": tk.StringVar(value=self.initial.get("type", ASSET_TYPES[0])),
        }
        if self.is_new:
            self.vars.update(
                {
                    "acquisition_date": tk.StringVar(value=date.today().strftime("%Y-%m-%d")),
                    "acquisition_cost": tk.StringVar(value="0"),
                    "delivery_cost": tk.StringVar(value=""),
                    "shop_link": tk.StringVar(value=""),
                }
            )

        self._row(frame, 0, "Name", "name")
        self._row(frame, 1, "Brand", "brand")
        self._row(frame, 2, "Model", "model")
        self._row(frame, 3, "Specifications", "specifications")
        self._row(frame, 4, "Series Number", "series_number")

        self._row(frame, 5, "Quantity", "quantity")
        self._row(frame, 6, "Location", "location")

        ttk.Label(frame, text="Type").grid(row=7, column=0, sticky="w", pady=4)
        ttk.Combobox(frame, textvariable=self.vars["type"], values=ASSET_TYPES, state="readonly").grid(
            row=7, column=1, sticky="ew"
        )

        row_idx = 8
        if self.is_new:
            ttk.Label(frame, text="Initial Acquisition Date").grid(row=row_idx, column=0, sticky="w", pady=4)
            make_date_entry(frame, self.vars["acquisition_date"]).grid(row=row_idx, column=1, columnspan=2, sticky="ew")
            row_idx += 1
            self._row(frame, row_idx, "Initial Acquisition Cost", "acquisition_cost")
            row_idx += 1
            self._row(frame, row_idx, "Initial Delivery Cost", "delivery_cost")
            row_idx += 1
            self._row(frame, row_idx, "Shop", "shop_link")
            row_idx += 1

        ttk.Label(frame, text="Picture").grid(row=row_idx, column=0, sticky="w", pady=4)
        ttk.Entry(frame, textvariable=self.vars["picture_path"], width=35).grid(row=row_idx, column=1, sticky="ew")
        ttk.Button(frame, text="Browse", command=self._browse_picture).grid(row=row_idx, column=2, padx=4)

        self.preview_label = ttk.Label(frame, text="No image selected")
        self.preview_label.grid(row=row_idx + 1, column=0, columnspan=3, pady=6)
        self.preview_image = None
        self._load_preview(self.vars["picture_path"].get().strip())

        ttk.Button(frame, text="Save", command=self._save).grid(row=row_idx + 2, column=0, columnspan=3, pady=8)
        frame.columnconfigure(1, weight=1)

    def _row(self, frame: ttk.Frame, row: int, label: str, key: str) -> None:
        ttk.Label(frame, text=label).grid(row=row, column=0, sticky="w", pady=4)
        ttk.Entry(frame, textvariable=self.vars[key]).grid(row=row, column=1, columnspan=2, sticky="ew")

    def _load_preview(self, path: str) -> None:
        if not path:
            self.preview_label.configure(text="No image selected", image="")
            return
        if Image is None or ImageTk is None:
            self.preview_label.configure(text="Preview requires Pillow (pip install pillow)", image="")
            return
        try:
            img = Image.open(path)
            img.thumbnail((240, 180))
            self.preview_image = ImageTk.PhotoImage(img)
            self.preview_label.configure(image=self.preview_image, text="")
        except Exception:
            self.preview_label.configure(text="Unable to load preview", image="")

    def _browse_picture(self) -> None:
        path = filedialog.askopenfilename(parent=self, title="Select Picture")
        if path:
            self.vars["picture_path"].set(path)
            self._load_preview(path)

    def _save(self) -> None:
        name = self.vars["name"].get().strip()
        if not name:
            messagebox.showwarning("Missing", "Name is required.")
            return
        try:
            quantity = float(self.vars["quantity"].get().strip())
        except ValueError:
            messagebox.showwarning("Invalid", "Quantity must be a number.")
            return
        if self.is_new:
            acquisition_date = self.vars["acquisition_date"].get().strip()
            if not acquisition_date:
                messagebox.showwarning("Missing", "Initial acquisition date is required.")
                return
            try:
                acquisition_cost = float(self.vars["acquisition_cost"].get().strip())
                delivery_cost = self.vars["delivery_cost"].get().strip()
                delivery_cost_val = float(delivery_cost) if delivery_cost else None
            except ValueError:
                messagebox.showwarning("Invalid", "Initial costs must be numbers.")
                return

        payload = {
            "picture_path": self.vars["picture_path"].get().strip() or None,
            "name": name,
            "brand": self.vars["brand"].get().strip() or None,
            "model": self.vars["model"].get().strip() or None,
            "specifications": self.vars["specifications"].get().strip() or None,
            "series_number": self.vars["series_number"].get().strip() or None,
            "quantity": quantity,
            "location": self.vars["location"].get().strip() or None,
            "type": self.vars["type"].get().strip(),
            "business": self.business,
            "inventory_type": self.inventory_type,
        }
        if self.is_new:
            payload.update(
                {
                    "acquisition_date": acquisition_date,
                    "acquisition_cost": acquisition_cost,
                    "delivery_cost": delivery_cost_val,
                    "shop_link": self.vars["shop_link"].get().strip() or None,
                }
            )
        self.on_save(payload)
        self.destroy()


class UserForm(tk.Toplevel):
    def __init__(self, parent: tk.Widget, title: str, on_save: Callable[[dict], None], initial: dict | None = None) -> None:
        super().__init__(parent)
        self.title(title)
        self.resizable(False, False)
        self.on_save = on_save
        self.initial = initial or {}
        self.transient(parent)
        self.grab_set()

        frame = ttk.Frame(self, padding=12)
        frame.pack(fill="both", expand=True)

        self.username_var = tk.StringVar(value=self.initial.get("username", ""))
        self.password_var = tk.StringVar()
        initial_businesses = []
        raw_business = self.initial.get("business", "")
        if isinstance(raw_business, str) and raw_business:
            if raw_business.strip() == "Both":
                initial_businesses = list(BUSINESSES)
            else:
                initial_businesses = [b.strip() for b in raw_business.split(",") if b.strip()]
        self.is_admin_var = tk.BooleanVar(value=bool(self.initial.get("is_admin", False)))

        ttk.Label(frame, text="Username").grid(row=0, column=0, sticky="w", pady=4)
        username_entry = ttk.Entry(frame, textvariable=self.username_var)
        username_entry.grid(row=0, column=1, sticky="ew")
        if self.initial.get("username"):
            username_entry.configure(state="disabled")

        ttk.Label(frame, text="Password").grid(row=1, column=0, sticky="w", pady=4)
        ttk.Entry(frame, textvariable=self.password_var, show="*").grid(row=1, column=1, sticky="ew")

        ttk.Label(frame, text="Businesses").grid(row=2, column=0, sticky="nw", pady=4)
        business_frame = ttk.Frame(frame)
        business_frame.grid(row=2, column=1, sticky="w")
        self.business_vars = {}
        for idx, biz in enumerate(BUSINESSES):
            var = tk.BooleanVar(value=(biz in initial_businesses) or (not initial_businesses and biz == "Unica"))
            self.business_vars[biz] = var
            ttk.Checkbutton(business_frame, text=biz, variable=var).grid(row=idx, column=0, sticky="w")

        ttk.Checkbutton(frame, text="Admin", variable=self.is_admin_var).grid(row=3, column=1, sticky="w", pady=4)

        ttk.Button(frame, text="Save", command=self._save).grid(row=4, column=0, columnspan=2, pady=8)
        frame.columnconfigure(1, weight=1)

    def _save(self) -> None:
        username = self.username_var.get().strip()
        password = self.password_var.get().strip()
        if not username:
            messagebox.showwarning("Missing", "Username is required.")
            return
        if not self.initial.get("username") and not password:
            messagebox.showwarning("Missing", "Password is required for new users.")
            return
        businesses = [biz for biz, var in self.business_vars.items() if var.get()]
        if not businesses:
            messagebox.showwarning("Missing", "Select at least one business.")
            return
        self.on_save(
            {
                "username": username,
                "password": password or None,
                "businesses": businesses,
                "is_admin": self.is_admin_var.get(),
            }
        )
        self.destroy()


class SummaryWindow(tk.Toplevel):
    def __init__(self, parent: tk.Widget, allowed_businesses: list[str]) -> None:
        super().__init__(parent)
        self.title("Summary & Export")
        self.geometry("980x550")
        self.allowed_businesses = allowed_businesses
        configure_photo_treeview_style(self)
        self.image_enabled = False
        self.image_paths: list[str | None] = []

        top = ttk.Frame(self, padding=8)
        top.pack(fill="x")

        ttk.Label(top, text="Business").grid(row=0, column=0, sticky="w")
        self.business_var = tk.StringVar(value=allowed_businesses[0])
        business_combo = ttk.Combobox(top, textvariable=self.business_var, values=allowed_businesses, state="readonly")
        business_combo.grid(row=0, column=1, sticky="w", padx=6)

        ttk.Label(top, text="Type").grid(row=0, column=2, sticky="w")
        self.type_var = tk.StringVar()
        self.type_combo = ttk.Combobox(top, textvariable=self.type_var, state="readonly")
        self.type_combo.grid(row=0, column=3, sticky="w", padx=6)
        self.type_combo.bind("<<ComboboxSelected>>", lambda _evt: self._update_date_controls())

        ttk.Label(top, text="From").grid(row=1, column=0, sticky="w")
        self.start_var = tk.StringVar(value=date.today().replace(day=1).strftime("%Y-%m-%d"))
        self.start_entry = make_date_entry(top, self.start_var)
        self.start_entry.grid(row=1, column=1, sticky="w", padx=6)

        ttk.Label(top, text="To").grid(row=1, column=2, sticky="w")
        self.end_var = tk.StringVar(value=date.today().strftime("%Y-%m-%d"))
        self.end_entry = make_date_entry(top, self.end_var)
        self.end_entry.grid(row=1, column=3, sticky="w", padx=6)

        self.range_buttons = [
            ttk.Button(top, text="Daily", command=lambda: self._set_range("daily")),
            ttk.Button(top, text="Weekly", command=lambda: self._set_range("weekly")),
            ttk.Button(top, text="Monthly", command=lambda: self._set_range("monthly")),
            ttk.Button(top, text="Yearly", command=lambda: self._set_range("yearly")),
        ]
        for idx, btn in enumerate(self.range_buttons):
            btn.grid(row=1, column=4 + idx, padx=4)

        ttk.Button(top, text="Load", command=self.load).grid(row=0, column=4, padx=6)
        ttk.Button(top, text="Export Excel", command=self.export_excel).grid(row=0, column=5, padx=6)
        ttk.Button(top, text="Export PDF", command=self.export_pdf).grid(row=0, column=6, padx=6)
        ttk.Button(top, text="Export JPG", command=self.export_jpg).grid(row=0, column=7, padx=6)

        self.tree = build_treeview(self, columns=(), headings=())
        self.data: list[list[object]] = []
        self.columns: list[str] = []
        self._update_type_options()
        business_combo.bind("<<ComboboxSelected>>", lambda _evt: self._update_type_options())
        self._update_date_controls()

    def _update_type_options(self) -> None:
        business = self.business_var.get()
        if business == "Unica":
            options = [
                "Unica Perishable",
                "Unica Perishable Expiry Dates",
                "Unica Perishable IN Logs",
                "Unica Perishable OUT Logs",
                "Unica Non-Perishable",
                "Unica Non-Perishable Statuses",
                "Unica Non-Perishable Acquisitions",
            ]
        else:
            options = [
                "HDN Warehouse",
                "HDN Warehouse Statuses",
                "HDN Warehouse Acquisitions",
                "HDN Plants",
            ]
        self.type_combo["values"] = options
        if self.type_var.get() not in options:
            self.type_var.set(options[0])
        self._update_date_controls()

    def _set_range(self, kind: str) -> None:
        today = date.today()
        if kind == "daily":
            start = end = today
        elif kind == "weekly":
            start = today - timedelta(days=today.weekday())
            end = start + timedelta(days=6)
        elif kind == "monthly":
            start = today.replace(day=1)
            next_month = (start.replace(day=28) + timedelta(days=4)).replace(day=1)
            end = next_month - timedelta(days=1)
        else:
            start = date(today.year, 1, 1)
            end = date(today.year, 12, 31)
        self.start_var.set(start.strftime("%Y-%m-%d"))
        self.end_var.set(end.strftime("%Y-%m-%d"))

    def _update_date_controls(self) -> None:
        report_type = self.type_var.get()
        perishable = report_type in (
            "Unica Perishable",
            "Unica Perishable Expiry Dates",
            "Unica Perishable IN Logs",
            "Unica Perishable OUT Logs",
        )
        acquisitions = report_type in ("Unica Non-Perishable Acquisitions", "HDN Warehouse Acquisitions")
        state = "normal" if (perishable or acquisitions) else "disabled"
        try:
            self.start_entry.configure(state=state)
        except Exception:
            pass
        try:
            self.end_entry.configure(state=state)
        except Exception:
            pass
        for btn in self.range_buttons:
            btn.configure(state=state)
        if not (perishable or acquisitions):
            self.start_var.set("")
            self.end_var.set("")
        else:
            if not self.start_var.get().strip():
                self.start_var.set(date.today().replace(day=1).strftime("%Y-%m-%d"))
            if not self.end_var.get().strip():
                self.end_var.set(date.today().strftime("%Y-%m-%d"))

    def load(self) -> None:
        business = self.business_var.get()
        inv_type = self.type_var.get()
        if business not in self.allowed_businesses:
            messagebox.showwarning("Access", "You do not have access to that business.")
            return
        if inv_type.startswith("Unica") and business != "Unica":
            messagebox.showwarning("Access", "Unica reports require Unica access.")
            return
        if inv_type.startswith("HDN") and business != "HDN Integrated Farm":
            messagebox.showwarning("Access", "HDN reports require HDN Integrated Farm access.")
            return

        start_date = self.start_var.get().strip()
        end_date = self.end_var.get().strip()

        if inv_type == "Unica Perishable":
            if not start_date or not end_date:
                messagebox.showwarning("Missing", "From and To dates are required for Unica Perishable.")
                return
            rows = get_perishable_report("Unica", start_date, end_date)
            self.columns = ["No.", "Id.", "Product", "Category", "Unit", "In (Range)", "Out (Range)"]
            self.data = [
                [idx, r["product_id"], r["name"], r["category"], r["unit"], r["in_qty"], r["out_qty"]]
                for idx, r in enumerate(rows, start=1)
            ]
            self.image_enabled = False
            self.image_paths = []
        elif inv_type == "Unica Perishable Expiry Dates":
            if not start_date or not end_date:
                messagebox.showwarning("Missing", "From and To dates are required for Expiry Dates.")
                return
            rows = list_expiry_dates_report("Unica", start_date, end_date)
            self.columns = ["No.", "Product Id", "Product", "Delivery Date", "Expiry Date", "Quantity"]
            self.data = [
                [idx, r["product_id"], r["name"], r["delivery_date"], r.get("expiry_date") or "", r["quantity"]]
                for idx, r in enumerate(rows, start=1)
            ]
            self.image_enabled = False
            self.image_paths = []
        elif inv_type == "Unica Perishable IN Logs":
            if not start_date or not end_date:
                messagebox.showwarning("Missing", "From and To dates are required for IN Logs.")
                return
            rows = list_in_logs_report("Unica", start_date, end_date)
            self.columns = ["No.", "Product Id", "Product", "Delivery Date", "Quantity"]
            self.data = [
                [idx, r["product_id"], r["name"], r["delivery_date"], r["quantity"]]
                for idx, r in enumerate(rows, start=1)
            ]
            self.image_enabled = False
            self.image_paths = []
        elif inv_type == "Unica Perishable OUT Logs":
            if not start_date or not end_date:
                messagebox.showwarning("Missing", "From and To dates are required for OUT Logs.")
                return
            rows = list_out_logs_report("Unica", start_date, end_date)
            self.columns = ["No.", "Product Id", "Product", "Out Date", "Out Time", "Quantity"]
            self.data = [
                [idx, r["product_id"], r["name"], r["out_date"], r["out_time"], r["quantity"]]
                for idx, r in enumerate(rows, start=1)
            ]
            self.image_enabled = False
            self.image_paths = []
        elif inv_type in ("Unica Non-Perishable", "HDN Warehouse"):
            biz = "Unica" if inv_type == "Unica Non-Perishable" else "HDN Integrated Farm"
            rows = list_assets_for_export(biz, inv_type, start_date or None, end_date or None)
            self.columns = [
                "No.",
                "Id.",
                "Name",
                "Type",
                "Brand",
                "Model",
                "Specifications",
                "Qty",
                "Total Spent",
                "Location",
            ]
            self.data = [
                [
                    idx,
                    r["id"],
                    r.get("name") or "",
                    r.get("type") or "",
                    r.get("brand") or "",
                    r.get("model") or "",
                    r.get("specifications") or "",
                    r["quantity"],
                    _format_money(r.get("total_spent")),
                    r.get("location") or "",
                ]
                for idx, r in enumerate(rows, start=1)
            ]
            self.image_enabled = True
            self.image_paths = [r.get("picture_path") for r in rows]
        elif inv_type in ("Unica Non-Perishable Statuses", "HDN Warehouse Statuses"):
            biz = "Unica" if inv_type.startswith("Unica") else "HDN Integrated Farm"
            inv_label = "Unica Non-Perishable" if inv_type.startswith("Unica") else "HDN Warehouse"
            rows = list_asset_statuses_report(biz, inv_label)
            self.columns = ["Asset Id", "Name", "Type", "Status", "Quantity"]
            self.data = [[r["asset_id"], r.get("name") or "", r.get("type") or "", r["status"], r["quantity"]] for r in rows]
            self.image_enabled = False
            self.image_paths = []
        elif inv_type in ("Unica Non-Perishable Acquisitions", "HDN Warehouse Acquisitions"):
            biz = "Unica" if inv_type.startswith("Unica") else "HDN Integrated Farm"
            inv_label = "Unica Non-Perishable" if inv_type.startswith("Unica") else "HDN Warehouse"
            rows = list_asset_acquisitions_report(biz, inv_label, start_date or None, end_date or None)
            self.columns = [
                "Asset Id",
                "Name",
                "Type",
                "Acquisition Date",
                "Acquisition Cost",
                "Delivery Cost",
                "Quantity",
                "Total Spent",
                "Shop",
            ]
            self.data = [
                [
                    r["asset_id"],
                    r.get("name") or "",
                    r.get("type") or "",
                    r.get("acquisition_date") or "",
                    _format_money(r.get("acquisition_cost")),
                    _format_money(r.get("delivery_cost")) if r.get("delivery_cost") is not None else "",
                    r["quantity"],
                    _format_money(float(r.get("acquisition_cost") or 0) * float(r.get("quantity") or 0)),
                    r.get("shop_link") or "",
                ]
                for r in rows
            ]
            self.image_enabled = False
            self.image_paths = []
        else:
            self.columns = ["Message"]
            self.data = [["No data yet for HDN Plants."]]
            self.image_enabled = False
            self.image_paths = []

        self._refresh_tree()

    def _refresh_tree(self) -> None:
        self.tree.delete(*self.tree.get_children())
        if self.image_enabled:
            self.tree.configure(show="tree headings", style="Photo.Treeview")
            self.tree.heading("#0", text="Picture")
            self.tree.column("#0", width=80, anchor="center")
            if hasattr(self.tree, "_img_refs"):
                self.tree._img_refs = {}
        else:
            self.tree.configure(show="headings", style="")
        self.tree["columns"] = self.columns
        for col in self.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=140, anchor="w")
        for idx, row in enumerate(self.data):
            iid = str(idx)
            self.tree.insert("", "end", iid=iid, values=row)
            if self.image_enabled:
                path = self.image_paths[idx] if idx < len(self.image_paths) else None
                set_tree_image(self.tree, iid, path)

    def _export(self, kind: str) -> None:
        if not self.data or not self.columns:
            messagebox.showwarning("Empty", "Load a report before exporting.")
            return
        filetypes = [("Excel", "*.xlsx"), ("CSV", "*.csv")] if kind == "excel" else [(kind.upper(), f"*.{kind}")]
        path = filedialog.asksaveasfilename(defaultextension=f".{kind}", filetypes=filetypes)
        if not path:
            return
        start = self.start_var.get().strip()
        end = self.end_var.get().strip()
        range_label = f"{start} to {end}" if start and end else "All Dates"
        title = f"{self.business_var.get()} - {self.type_var.get()} ({range_label})"
        header_lines = [
            f"{self.business_var.get()} Inventory",
            f"{self.type_var.get()}",
            f"As of {date.today().strftime('%Y-%m-%d')}",
        ]
        try:
            if kind == "excel":
                export_to_excel(
                    path,
                    self.columns,
                    self.data,
                    image_paths=self.image_paths if self.image_enabled else None,
                    image_height=PHOTO_THUMBNAIL_SIZE[1],
                    header_lines=header_lines,
                    image_column=1,
                )
            elif kind == "pdf":
                export_to_pdf(
                    path,
                    title,
                    self.columns,
                    self.data,
                    image_paths=self.image_paths if self.image_enabled else None,
                    image_height=PHOTO_THUMBNAIL_SIZE[1],
                    header_lines=header_lines,
                    image_column=1,
                )
            else:
                export_to_jpg(
                    path,
                    title,
                    self.columns,
                    self.data,
                    image_paths=self.image_paths if self.image_enabled else None,
                    image_height=PHOTO_THUMBNAIL_SIZE[1],
                    header_lines=header_lines,
                    image_column=1,
                )
            messagebox.showinfo("Exported", f"Saved to {path}")
        except Exception as exc:
            messagebox.showerror("Export failed", str(exc))

    def export_excel(self) -> None:
        self._export("excel")

    def export_pdf(self) -> None:
        self._export("pdf")

    def export_jpg(self) -> None:
        self._export("jpg")


class InsightsWindow(tk.Toplevel):
    def __init__(self, parent: tk.Tk, allowed_businesses: list[str]) -> None:
        super().__init__(parent)
        self.title("Insights")
        self.geometry("980x600")
        self.allowed_businesses = allowed_businesses
        configure_photo_treeview_style(self)

        top = ttk.Frame(self, padding=8)
        top.pack(fill="x")

        ttk.Label(top, text="Business").grid(row=0, column=0, sticky="w")
        self.business_var = tk.StringVar(value=allowed_businesses[0])
        business_combo = ttk.Combobox(top, textvariable=self.business_var, values=allowed_businesses, state="readonly")
        business_combo.grid(row=0, column=1, sticky="w", padx=6)

        ttk.Button(top, text="Refresh", command=self.load).grid(row=0, column=2, padx=6)
        ttk.Button(top, text="Export Excel", command=self.export_excel).grid(row=0, column=3, padx=6)
        ttk.Button(top, text="Export PDF", command=self.export_pdf).grid(row=0, column=4, padx=6)
        ttk.Button(top, text="Export JPG", command=self.export_jpg).grid(row=0, column=5, padx=6)

        self.tree = build_treeview(self, columns=("metric", "value"), headings=("Metric", "Value"))
        self.columns = ["Metric", "Value"]
        self.data: list[list[object]] = []
        self._status_chart: list[tuple[str, float]] = []
        self._expiry_chart: list[tuple[str, float]] = []

        chart_frame = ttk.Frame(self, padding=8)
        chart_frame.pack(fill="both", expand=False)
        ttk.Label(chart_frame, text="Charts", style="Muted.TLabel").pack(anchor="w")
        self.chart_canvas = tk.Canvas(
            chart_frame,
            height=220,
            background=UI_COLORS["panel"],
            highlightthickness=1,
            highlightbackground=UI_COLORS["border"],
        )
        self.chart_canvas.pack(fill="x", expand=False, pady=(6, 0))
        self.chart_canvas.bind("<Configure>", lambda _e: self._draw_charts(), add="+")
        business_combo.bind("<<ComboboxSelected>>", lambda _evt: self.load())
        self.load()

    def load(self) -> None:
        business = self.business_var.get()
        self.data = self._build_insights(business)
        self._refresh_tree()
        self._draw_charts()

    def _refresh_tree(self) -> None:
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = self.columns
        for col in self.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=220, anchor="w")
        for idx, row in enumerate(self.data, start=1):
            self.tree.insert("", "end", iid=str(idx), values=row)

    def _build_insights(self, business: str) -> list[list[object]]:
        rows: list[list[object]] = []
        asset_business = business
        inventory_type = "Unica Non-Perishable" if business == "Unica" else "HDN Warehouse"
        assets = list_assets(asset_business, inventory_type)
        total_assets = len(assets)
        total_qty = 0.0
        total_spent = 0.0
        for asset in assets:
            try:
                total_qty += float(asset.get("quantity") or 0)
            except ValueError:
                pass
            try:
                total_spent += float(asset.get("total_spent") or 0)
            except ValueError:
                pass

        rows.append(["Asset items", _format_number(total_assets)])
        rows.append(["Asset qty total", _format_number(total_qty)])
        rows.append(["Total spent (assets)", _format_php(total_spent)])

        acquisitions = list_asset_acquisitions_report(asset_business, inventory_type)
        rows.append(["Acquisition entries", _format_number(len(acquisitions))])
        if acquisitions:
            total_acq_qty = 0.0
            dates: list[date] = []
            for acq in acquisitions:
                try:
                    total_acq_qty += float(acq.get("quantity") or 0)
                except ValueError:
                    pass
                acq_date = _safe_date(acq.get("acquisition_date"))
                if acq_date:
                    dates.append(acq_date)
            avg_qty = total_acq_qty / len(acquisitions)
            rows.append(["Avg acquisition qty", _format_number(avg_qty)])
            if dates:
                span_days = (max(dates) - min(dates)).days
                span_months = max(1, round(span_days / 30))
                rows.append(
                    ["Acquisitions per month (approx)", _format_number(len(acquisitions) / span_months)]
                )

        statuses = list_asset_statuses_report(asset_business, inventory_type)
        status_totals: dict[str, float] = {}
        total_status_qty = 0.0
        for entry in statuses:
            status = entry.get("status") or "Unknown"
            try:
                qty = float(entry.get("quantity") or 0)
            except ValueError:
                qty = 0.0
            status_totals[status] = status_totals.get(status, 0.0) + qty
            total_status_qty += qty
        if status_totals:
            rows.append(["Status qty total", _format_number(total_status_qty)])
            self._status_chart = []
            for status, qty in sorted(status_totals.items()):
                pct = (qty / total_status_qty * 100) if total_status_qty else 0
                rows.append([f"Status: {status}", f"{pct:.1f}% (qty {_format_number(qty)})"])
                self._status_chart.append((status, pct))
        else:
            self._status_chart = []

        if business == "Unica":
            products = list_products("Unica")
            rows.append(["Perishable products", _format_number(len(products))])

            expiry_rows = list_expiry_dates_report("Unica")
            rows.append(["Expiry date entries", _format_number(len(expiry_rows))])
            today = date.today()
            expiring_7 = 0
            expired = 0
            for entry in expiry_rows:
                exp = _safe_date(entry.get("expiry_date"))
                if not exp:
                    continue
                delta = (exp - today).days
                if delta < 0:
                    expired += 1
                elif delta <= 7:
                    expiring_7 += 1
            rows.append(["Expired entries", _format_number(expired)])
            rows.append(["Expiring in 7 days", _format_number(expiring_7)])
            self._expiry_chart = [("Expired", float(expired)), ("Expiring <=7d", float(expiring_7))]

            in_logs = list_in_logs_report("Unica")
            out_logs = list_out_logs_report("Unica")
            rows.append(["IN logs", _format_number(len(in_logs))])
            rows.append(["OUT logs", _format_number(len(out_logs))])
            in_qty = sum(float(r.get("quantity") or 0) for r in in_logs)
            out_qty = sum(float(r.get("quantity") or 0) for r in out_logs)
            rows.append(["IN qty total", _format_number(in_qty)])
            rows.append(["OUT qty total", _format_number(out_qty)])
        else:
            self._expiry_chart = []

        return rows

    def _draw_charts(self) -> None:
        self.chart_canvas.delete("all")
        width = int(self.chart_canvas.winfo_width() or 800)
        height = int(self.chart_canvas.winfo_height() or 220)
        margin = 10

        sections = []
        if self._status_chart:
            sections.append(("Status %", self._status_chart))
        if self._expiry_chart:
            sections.append(("Expiry", self._expiry_chart))
        if not sections:
            self.chart_canvas.create_text(
                width // 2,
                height // 2,
                text="No chart data available.",
                fill=UI_COLORS["muted"],
            )
            return

        section_height = max(1, (height - margin * 2) // len(sections))
        for idx, (title, items) in enumerate(sections):
            top = margin + idx * section_height
            self.chart_canvas.create_text(margin, top + 8, text=title, anchor="nw", fill=UI_COLORS["text"])
            if not items:
                continue
            bar_area_top = top + 24
            bar_area_height = section_height - 34
            bar_width = max(10, (width - margin * 2) // len(items))
            max_val = max(v for _n, v in items) or 1
            for i, (name, value) in enumerate(items):
                x0 = margin + i * bar_width + 4
                x1 = x0 + bar_width - 8
                bar_h = int((value / max_val) * max(10, bar_area_height))
                y1 = bar_area_top + bar_area_height
                y0 = y1 - bar_h
                self.chart_canvas.create_rectangle(x0, y0, x1, y1, fill=UI_COLORS["accent"], outline="")
                self.chart_canvas.create_text(
                    (x0 + x1) / 2,
                    y1 + 2,
                    text=name,
                    anchor="n",
                    fill=UI_COLORS["muted"],
                )

    def _export(self, kind: str) -> None:
        if not self.data:
            messagebox.showwarning("Empty", "No insights to export.")
            return
        filetypes = [("Excel", "*.xlsx"), ("CSV", "*.csv")] if kind == "excel" else [(kind.upper(), f"*.{kind}")]
        path = filedialog.asksaveasfilename(defaultextension=f".{kind}", filetypes=filetypes)
        if not path:
            return
        title = f"{self.business_var.get()} - Insights"
        header_lines = [
            f"{self.business_var.get()} Inventory",
            "Insights",
            f"As of {date.today().strftime('%Y-%m-%d')}",
        ]
        try:
            if kind == "excel":
                export_to_excel(path, self.columns, self.data, header_lines=header_lines)
            elif kind == "pdf":
                export_to_pdf(path, title, self.columns, self.data, header_lines=header_lines)
            else:
                export_to_jpg(path, title, self.columns, self.data, header_lines=header_lines)
            messagebox.showinfo("Exported", f"Saved to {path}")
        except Exception as exc:
            messagebox.showerror("Export failed", str(exc))

    def export_excel(self) -> None:
        self._export("excel")

    def export_pdf(self) -> None:
        self._export("pdf")

    def export_jpg(self) -> None:
        self._export("jpg")


class MainWindow:
    def __init__(self, root: tk.Tk, user: dict) -> None:
        self.user = user
        self.username = user["username"]
        self.businesses = user.get("businesses") or [user.get("business")]
        self.is_admin = bool(user["is_admin"])

        self.root = root
        self.root.title("Aman Inventory")
        self.root.geometry("1200x720")
        self.root.resizable(True, True)
        configure_photo_treeview_style(self.root)
        self._logo_images: list[ImageTk.PhotoImage] = []

        self.allowed_businesses = [b for b in self.businesses if b]
        if not self.allowed_businesses:
            self.allowed_businesses = BUSINESSES

        header_font = ("Segoe UI", 11, "bold")
        sub_font = ("Segoe UI", 9)

        self.top = ttk.Frame(self.root, padding=8)
        self.top.pack(fill="x")
        self.top.columnconfigure(0, weight=1)
        self.top.columnconfigure(1, weight=1)
        self.top.columnconfigure(2, weight=0)

        left = ttk.Frame(self.top)
        left.grid(row=0, column=0, sticky="w")
        ttk.Label(
            left,
            text=f"Hello {self.username}!",
            font=header_font,
            padding=(0, 0, 0, 2),
        ).grid(row=0, column=0, sticky="w")
        ttk.Label(
            left,
            text=f"Logged in: {', '.join(self.allowed_businesses)}",
            font=sub_font,
        ).grid(row=1, column=0, sticky="w")

        middle = ttk.Frame(self.top)
        middle.grid(row=0, column=1, sticky="n")
        self.header_summary = ttk.Label(middle, text="Items: 0 | Selected: 0", style="Muted.TLabel")
        self.header_summary.grid(row=0, column=0, sticky="n")

        right = ttk.Frame(self.top)
        right.grid(row=0, column=2, sticky="e")
        ttk.Button(right, text="Summary / Export", command=self.open_summary).pack(side="left", padx=4)
        ttk.Button(right, text="View Insights", command=self.open_insights).pack(side="left", padx=4)
        ttk.Button(right, text="Check for Updates", command=self.open_updates).pack(side="left")


        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True)
        self._tree_tabs: dict[ttk.Treeview, tk.Widget] = {}
        self.notebook.bind("<<NotebookTabChanged>>", self._on_tab_change, add="+")

        if "Unica" in self.allowed_businesses:
            self._build_perishable_tab()
            self._build_assets_tab("Unica", "Unica Non-Perishable")
        if "HDN Integrated Farm" in self.allowed_businesses:
            self._build_assets_tab("HDN Integrated Farm", "HDN Warehouse")
            self._build_blank_tab("HDN Plants")
        if self.is_admin:
            self._build_users_tab()

    def open_summary(self) -> None:
        SummaryWindow(self.root, self.allowed_businesses)

    def open_insights(self) -> None:
        InsightsWindow(self.root, self.allowed_businesses)

    def open_updates(self) -> None:
        if not GITHUB_RELEASES_URL:
            messagebox.showinfo("Updates", "Set GITHUB_RELEASES_URL in app.py to your GitHub releases page.")
            return
        try:
            webbrowser.open(GITHUB_RELEASES_URL)
        except Exception:
            messagebox.showwarning("Updates", "Unable to open the releases page.")

    def _register_tree(self, tree: ttk.Treeview, tab: tk.Widget) -> None:
        self._tree_tabs[tree] = tab

    def _is_tree_active(self, tree: ttk.Treeview) -> bool:
        if not hasattr(self, "notebook"):
            return False
        tab_id = self._tree_tabs.get(tree)
        if tab_id is None:
            return False
        return str(self.notebook.select()) == str(tab_id)

    def _on_tab_change(self, _event: tk.Event) -> None:
        active_tab = self.notebook.select()
        for tree, tab in self._tree_tabs.items():
            if str(tab) == str(active_tab):
                text = getattr(tree, "_summary_text", "Items: 0 | Selected: 0")
                self.header_summary.configure(text=text)
                return
        self.header_summary.configure(text="Items: 0 | Selected: 0")

    def _apply_logo(self, label: ttk.Label, tab_key: str) -> None:
        path = TAB_LOGOS.get(tab_key)
        if not path:
            return
        full_path = path if os.path.isabs(path) else os.path.join(os.path.dirname(__file__), path)
        logo = load_logo_image(full_path)
        if logo is None:
            return
        label.configure(image=logo)
        self._logo_images.append(logo)

    def _bind_dynamic_search(self, entry: ttk.Entry, callback: Callable[[], None]) -> None:
        def _trigger(_event: tk.Event) -> None:
            if hasattr(entry, "_search_after"):
                try:
                    entry.after_cancel(entry._search_after)
                except Exception:
                    pass
            entry._search_after = entry.after(200, callback)

        entry.bind("<KeyRelease>", _trigger, add="+")

    def _update_tree_summary(self, tree: ttk.Treeview, total_items: int, total_qty: float | None = None) -> None:
        if not hasattr(tree, "_summary_label"):
            return
        selected = len(tree.selection())
        parts = [f"Items: {total_items}"]
        if total_qty is not None:
            parts.append(f"Total qty: {total_qty:g}")
        if selected:
            parts.append(f"Selected: {selected}")
        text = " | ".join(parts)
        tree._summary_label.configure(text=text)
        tree._summary_text = text
        if hasattr(self, "header_summary") and self._is_tree_active(tree):
            self.header_summary.configure(text=text)

    def _bind_summary_updates(self, tree: ttk.Treeview) -> None:
        def _on_select(_event: tk.Event) -> None:
            total_items = getattr(tree, "_summary_total_items", 0)
            total_qty = getattr(tree, "_summary_total_qty", None)
            self._update_tree_summary(tree, total_items, total_qty)

        tree.bind("<<TreeviewSelect>>", _on_select, add="+")

    def _get_perishable_categories(self) -> list[str]:
        categories = sorted({row["category"] for row in list_products("Unica") if row.get("category")})
        return ["All"] + categories

    def _refresh_perishable_categories(self) -> None:
        if not hasattr(self, "perishable_category_combo"):
            return
        categories = self._get_perishable_categories()
        self.perishable_category_combo["values"] = categories
        if self.perishable_category.get() not in categories:
            self.perishable_category.set("All")

    def _build_perishable_tab(self) -> None:
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Unica Perishable")

        top = ttk.Frame(tab, padding=(8, 2))
        top.pack(fill="x")

        self.perishable_search = tk.StringVar()
        logo_label = ttk.Label(top)
        logo_label.grid(row=0, column=0, rowspan=2, sticky="w", padx=(0, 10), pady=(0, 4))
        self._apply_logo(logo_label, "Unica Perishable")
        perishable_search_entry = ttk.Entry(top, textvariable=self.perishable_search, width=32)
        perishable_search_entry.grid(row=0, column=1, sticky="ew", padx=4)
        ttk.Label(top, text="Category").grid(row=0, column=2, sticky="w", padx=(8, 2))
        self.perishable_category = tk.StringVar(value="All")
        self.perishable_category_combo = ttk.Combobox(
            top, textvariable=self.perishable_category, values=self._get_perishable_categories(), state="readonly", width=20
        )
        self.perishable_category_combo.grid(row=0, column=3, sticky="w")
        ttk.Button(top, text="Search", command=self.refresh_perishable).grid(row=0, column=4, sticky="w", padx=6)
        ttk.Button(top, text="Clear Search", command=self._clear_perishable_search).grid(row=0, column=5, sticky="w")
        ttk.Button(top, text="Refresh", command=self.refresh_perishable).grid(row=0, column=6, sticky="w", padx=6)

        buttons = ttk.Frame(top)
        buttons.grid(row=1, column=1, columnspan=6, sticky="w", pady=(1, 0))
        ttk.Button(buttons, text="Add Product", command=self.add_product).pack(side="left", padx=2)
        ttk.Button(buttons, text="Edit Product", command=self.edit_product).pack(side="left", padx=2)
        ttk.Button(buttons, text="Delete Product", command=self.delete_product).pack(side="left", padx=2)
        ttk.Button(buttons, text="Record IN", command=self.record_in).pack(side="left", padx=2)
        ttk.Button(buttons, text="Record OUT", command=self.record_out).pack(side="left", padx=2)
        ttk.Button(buttons, text="View Record", command=self.view_perishable_record).pack(side="left", padx=2)
        ttk.Button(buttons, text="Add Expiry Dates", command=self.add_expiry_dates).pack(side="left", padx=2)

        top.columnconfigure(1, weight=1)

        columns = ["no", "id", "name", "category", "unit", "opening", "in_qty", "out_qty", "ending"]
        headings = ["No.", "Id.", "Product", "Category", "Unit", "Opening", "In", "Out", "Ending"]
        self.perishable_tree = build_treeview(tab, columns, headings, image_heading="Picture", style="Photo.Treeview")
        self.perishable_tree.column("no", width=60, anchor="center")
        self.perishable_tree.column("id", width=80, anchor="center")
        self.perishable_tree.tag_configure("low_stock", background="#f8d7da")
        self.perishable_tree.tag_configure("expiry_3", background="#fce5cd")
        self.perishable_tree.tag_configure("expiry_7", background="#fff2cc")
        self._register_tree(self.perishable_tree, tab)

        summary_bar = ttk.Frame(tab, style="Card.TFrame", padding=(8, 4))
        summary_bar.pack(fill="x", padx=8, pady=(0, 8))
        self.perishable_summary = ttk.Label(summary_bar, text="Items: 0 | Total qty: 0", style="PanelMuted.TLabel")
        self.perishable_summary.pack(anchor="w")
        self.perishable_tree._summary_label = self.perishable_summary
        self._bind_summary_updates(self.perishable_tree)

        self.refresh_perishable()

        self._bind_dynamic_search(perishable_search_entry, self.refresh_perishable)
        perishable_search_entry.bind("<Return>", lambda _e: self.refresh_perishable(), add="+")

    def _clear_perishable_search(self) -> None:
        self.perishable_search.set("")
        self.refresh_perishable()

    def refresh_perishable(self) -> None:
        search = self.perishable_search.get().strip()
        category = self.perishable_category.get().strip()
        category_filter = None if not category or category == "All" else category
        rows = get_perishable_stock("Unica", search if search else None, category_filter)
        self.perishable_tree.delete(*self.perishable_tree.get_children())
        if hasattr(self.perishable_tree, "_img_refs"):
            self.perishable_tree._img_refs = {}
        today = date.today()
        total_qty = 0
        for idx, row in enumerate(rows, start=1):
            ending = float(row["opening_stock"]) + float(row["in_qty"]) - float(row["out_qty"])
            total_qty += ending
            expiry_date = _safe_date(row.get("next_expiry"))
            expiring_3_qty = float(row.get("expiring_3_qty") or 0)
            expiring_7_qty = float(row.get("expiring_7_qty") or 0)
            expiry_label = ""
            expiry_tag = ""
            if expiry_date:
                delta = (expiry_date - today).days
                if delta <= 3:
                    suffix = f" ({expiring_3_qty:g} exp)" if expiring_3_qty else ""
                    expiry_label = f"D-{max(delta, 0)}{suffix}"
                    expiry_tag = "expiry_3"
                elif delta <= 7:
                    suffix = f" ({expiring_7_qty:g} exp)" if expiring_7_qty else ""
                    expiry_label = f"D-{delta}{suffix}"
                    expiry_tag = "expiry_7"
                else:
                    expiry_label = expiry_date.strftime("%Y-%m-%d")

            tag = ""
            if ending <= float(row.get("low_stock_level", DEFAULT_LOW_STOCK_LEVEL)):
                tag = "low_stock"
            elif expiry_tag:
                tag = expiry_tag

            self.perishable_tree.insert(
                "",
                "end",
                iid=str(row["id"]),
                text="",
                values=(
                    idx,
                    row["id"],
                    row["name"],
                    row["category"],
                    row["unit"],
                    row["opening_stock"],
                    row["in_qty"],
                    row["out_qty"],
                    ending,
                ),
                tags=(tag,) if tag else (),
            )
            set_tree_image(self.perishable_tree, str(row["id"]), row.get("photo_path"))
        self.perishable_tree._summary_total_items = len(rows)
        self.perishable_tree._summary_total_qty = total_qty
        self._update_tree_summary(self.perishable_tree, len(rows), total_qty)

    def _get_selected_product(self) -> tuple[int, dict] | None:
        sel = self.perishable_tree.selection()
        if not sel:
            return None
        pid = int(sel[0])
        row = next((r for r in list_products("Unica") if r["id"] == pid), None)
        if not row:
            return None
        return pid, dict(row)

    def add_product(self) -> None:
        def on_save(data: dict) -> None:
            add_product(
                data["name"],
                data["category"],
                data["unit"],
                data["opening_stock"],
                data["photo_path"],
                data["low_stock_level"],
                "Unica",
            )
            self.refresh_perishable()
            self._refresh_perishable_categories()

        ProductForm(self.root, "Add Product", on_save)

    def edit_product(self) -> None:
        selected = self._get_selected_product()
        if not selected:
            messagebox.showwarning("Select", "Select a product to edit.")
            return
        pid, row = selected

        def on_save(data: dict) -> None:
            update_product(
                pid,
                data["name"],
                data["category"],
                data["unit"],
                data["opening_stock"],
                data["photo_path"],
                data["low_stock_level"],
            )
            self.refresh_perishable()
            self._refresh_perishable_categories()

        ProductForm(self.root, "Edit Product", on_save, initial=row)

    def delete_product(self) -> None:
        selected = self._get_selected_product()
        if not selected:
            messagebox.showwarning("Select", "Select a product to delete.")
            return
        pid, _ = selected
        if messagebox.askyesno("Confirm", "Delete this product?"):
            delete_product(pid)
            self.refresh_perishable()
            self._refresh_perishable_categories()

    def record_in(self) -> None:
        products = [(r["id"], r["name"]) for r in list_products("Unica")]
        if not products:
            messagebox.showwarning("Missing", "Add a product first.")
            return
        selected = self._get_selected_product()
        default_name = selected[1]["name"] if selected else None

        def on_save(data: dict) -> None:
            record_in(
                data["product_id"],
                data["date"],
                data.get("quantity", 0),
            )
            self.refresh_perishable()

        InOutForm(self.root, "Record IN", products, on_save, default_product=default_name)

    def record_out(self) -> None:
        products = [(r["id"], r["name"]) for r in list_products("Unica")]
        if not products:
            messagebox.showwarning("Missing", "Add a product first.")
            return
        selected = self._get_selected_product()
        default_name = selected[1]["name"] if selected else None

        def on_save(data: dict) -> None:
            time_val = data.get("time", "")
            if not time_val:
                messagebox.showwarning("Missing", "Out time is required for OUT.")
                return
            record_out(data["product_id"], data["date"], time_val, data["quantity"])
            self.refresh_perishable()

        InOutForm(self.root, "Record OUT", products, on_save, default_product=default_name)

    def view_logs(self, kind: str, product_id: int | None = None) -> None:
        if product_id is None:
            selected = self._get_selected_product()
            if not selected:
                messagebox.showwarning("Select", "Select a product first.")
                return
            pid, row = selected
        else:
            pid = product_id
            row = next((r for r in list_products("Unica") if r["id"] == pid), None) or {"name": f"ID {pid}"}
        logs = list_in_out_logs(kind, pid)

        win = tk.Toplevel(self.root)
        win.title(f"{row['name']} - {kind.upper()} Logs")
        top = ttk.Frame(win, padding=6)
        top.pack(fill="x")
        ttk.Button(top, text="Edit", command=lambda: self._edit_log(kind, tree, pid)).pack(side="left", padx=4)
        ttk.Button(top, text="Delete", command=lambda: self._delete_log(kind, tree, pid)).pack(side="left")

        if kind == "in":
            columns = ("id", "date", "qty")
            headings = ("ID", "Delivery Date", "Quantity")
        else:
            columns = ("id", "date", "time", "qty")
            headings = ("ID", "Out Date", "Out Time", "Quantity")

        tree = build_treeview(win, columns, headings)
        for log in logs:
            if kind == "in":
                tree.insert(
                    "",
                    "end",
                    iid=str(log["id"]),
                    values=(log["id"], log["delivery_date"], log["quantity"]),
                )
            else:
                tree.insert("", "end", iid=str(log["id"]), values=(log["id"], log["out_date"], log["out_time"], log["quantity"]))

    def view_perishable_record(self) -> None:
        selected = self._get_selected_product()
        if not selected:
            messagebox.showwarning("Select", "Select a product first.")
            return
        pid, row = selected

        win = tk.Toplevel(self.root)
        win.title(f"{row['name']} - Record")
        header = ttk.Frame(win, padding=8)
        header.pack(fill="x")

        img_label = ttk.Label(header, text="No image")
        img_label.pack(side="left", padx=(0, 10))
        preview = _load_preview_image(row.get("photo_path") or "")
        if preview is not None:
            img_label.configure(image=preview, text="")
            win._preview_image = preview

        info = ttk.Frame(header)
        info.pack(side="left", fill="x", expand=True)
        content = "\n".join(
            [
                f"ID: {row.get('id') or ''}",
                f"Product: {row.get('name') or ''}",
                f"Category: {row.get('category') or ''}",
                f"Unit: {row.get('unit') or ''}",
                f"Opening Stock: {row.get('opening_stock') or ''}",
                f"Low Stock Level: {row.get('low_stock_level') or ''}",
                f"Business: {row.get('business') or ''}",
                f"Photo Path: {row.get('photo_path') or ''}",
            ]
        )
        make_readonly_text(info, content, height=8)

        actions = ttk.Frame(win, padding=6)
        actions.pack(fill="x")
        ttk.Button(actions, text="View IN Logs", command=lambda: self.view_logs("in", pid)).pack(side="left", padx=4)
        ttk.Button(actions, text="View OUT Logs", command=lambda: self.view_logs("out", pid)).pack(side="left", padx=4)
        ttk.Button(actions, text="View Expiry Dates", command=lambda: self.view_expiry_dates(pid)).pack(side="left", padx=4)

    def view_expiry_dates(self, product_id: int | None = None) -> None:
        if product_id is None:
            selected = self._get_selected_product()
            if not selected:
                messagebox.showwarning("Select", "Select a product first.")
                return
            pid, row = selected
        else:
            pid = product_id
            row = next((r for r in list_products("Unica") if r["id"] == pid), None) or {"name": f"ID {pid}"}
        logs = list_expiry_dates(pid)

        win = tk.Toplevel(self.root)
        win.title(f"{row['name']} - Expiry Dates")
        top = ttk.Frame(win, padding=6)
        top.pack(fill="x")
        ttk.Label(top, text="Expiry dates for selected product").pack(side="left")

        columns = ("qty", "expiry", "delivery", "status", "id")
        headings = ("Quantity", "Expiry Date", "Delivery Date", "Status", "ID")
        tree = build_treeview(win, columns, headings)
        tree.tag_configure("expiry_3", background="#fce5cd")
        tree.tag_configure("expiry_7", background="#fff2cc")
        tree.tag_configure("expired", background="#f8d7da")

        today = date.today()
        for log in logs:
            expiry_date = _safe_date(log.get("expiry_date"))
            status = "No Expiry" if not expiry_date else expiry_date.strftime("%Y-%m-%d")
            tag = ""
            if expiry_date:
                delta = (expiry_date - today).days
                if delta < 0:
                    status = "Expired"
                    tag = "expired"
                elif delta <= 3:
                    status = f"D-{max(delta, 0)}"
                    tag = "expiry_3"
                elif delta <= 7:
                    status = f"D-{delta}"
                    tag = "expiry_7"
                else:
                    status = expiry_date.strftime("%Y-%m-%d")
            tree.insert(
                "",
                "end",
                iid=str(log["id"]),
                values=(
                    log["quantity"],
                    log.get("expiry_date") or "",
                    log["delivery_date"],
                    status,
                    log["id"],
                ),
                tags=(tag,) if tag else (),
            )

    def add_expiry_dates(self) -> None:
        selected = self._get_selected_product()
        if not selected:
            messagebox.showwarning("Select", "Select a product first.")
            return
        pid, row = selected
        logs = list_in_out_logs("in", pid)
        if not logs:
            messagebox.showwarning("Missing", "No IN records found for this product.")
            return

        win = tk.Toplevel(self.root)
        win.title(f"{row['name']} - Add Expiry Dates")
        top = ttk.Frame(win, padding=6)
        top.pack(fill="x")

        ttk.Label(top, text="IN Record").grid(row=0, column=0, sticky="w")
        in_var = tk.StringVar()
        in_options = [
            f"{log['id']} | {log['delivery_date']} | Qty {log['quantity']}"
            for log in logs
        ]
        in_combo = ttk.Combobox(top, textvariable=in_var, values=in_options, state="readonly", width=40)
        in_combo.grid(row=0, column=1, sticky="w", padx=6)
        in_var.set(in_options[0])

        list_frame = ttk.Frame(win, padding=6)
        list_frame.pack(fill="both", expand=True)
        columns = ("qty", "expiry", "id")
        headings = ("Quantity", "Expiry Date", "ID")
        tree = build_treeview(list_frame, columns, headings)

        def refresh_tree() -> None:
            tree.delete(*tree.get_children())
            selected_id = int(in_var.get().split("|")[0].strip())
            for item in list_in_breakdown(selected_id):
                tree.insert(
                    "",
                    "end",
                    iid=str(item["id"]),
                    values=(item["quantity"], item.get("expiry_date") or "", item["id"]),
                )

        def add_entry() -> None:
            selected_id = int(in_var.get().split("|")[0].strip())
            log = next((l for l in logs if l["id"] == selected_id), None)
            if not log:
                return
            add_win = tk.Toplevel(win)
            add_win.title("Add Expiry Entry")
            frm = ttk.Frame(add_win, padding=12)
            frm.pack(fill="both", expand=True)
            ttk.Label(frm, text="Qty").grid(row=0, column=0, sticky="w", pady=4)
            qty_var = tk.StringVar()
            ttk.Entry(frm, textvariable=qty_var).grid(row=0, column=1, sticky="ew")
            ttk.Label(frm, text="Expiry Date (optional)").grid(row=1, column=0, sticky="w", pady=4)
            expiry_var = tk.StringVar()
            make_date_entry(frm, expiry_var).grid(row=1, column=1, sticky="ew")

            def save() -> None:
                qty_raw = qty_var.get().strip()
                if not qty_raw:
                    messagebox.showwarning("Missing", "Quantity is required.")
                    return
                try:
                    qty_val = float(qty_raw)
                except ValueError:
                    messagebox.showwarning("Invalid", "Quantity must be a number.")
                    return
                existing = sum(float(r["quantity"]) for r in list_in_breakdown(selected_id))
                if existing + qty_val > float(log["quantity"]):
                    messagebox.showwarning("Limit", "Expiry quantities exceed the IN quantity.")
                    return
                expiry_val = expiry_var.get().strip() or None
                add_in_breakdown(selected_id, expiry_val, qty_val)
                add_win.destroy()
                refresh_tree()

            ttk.Button(frm, text="Save", command=save).grid(row=2, column=0, columnspan=2, pady=8)
            frm.columnconfigure(1, weight=1)

        def delete_entry() -> None:
            sel = tree.selection()
            if not sel:
                messagebox.showwarning("Select", "Select an entry to delete.")
                return
            if not messagebox.askyesno("Confirm", "Delete selected expiry entry?"):
                return
            delete_in_breakdown(int(sel[0]))
            refresh_tree()

        in_combo.bind("<<ComboboxSelected>>", lambda _e: refresh_tree())
        btns = ttk.Frame(win, padding=6)
        btns.pack(fill="x")
        ttk.Button(btns, text="Add", command=add_entry).pack(side="left", padx=2)
        ttk.Button(btns, text="Delete", command=delete_entry).pack(side="left", padx=2)

        refresh_tree()

    def _edit_log(self, kind: str, tree: ttk.Treeview, product_id: int) -> None:
        sel = tree.selection()
        if not sel:
            messagebox.showwarning("Select", "Select a log to edit.")
            return
        log_id = int(sel[0])
        logs = list_in_out_logs(kind, product_id)
        log = next((l for l in logs if l["id"] == log_id), None)
        if not log:
            return

        def on_save(data: dict) -> None:
            if kind == "in":
                update_in_log(data["id"], data["date"], data["quantity"])
            else:
                update_out_log(data["id"], data["date"], data["time"], data["quantity"])
            self.refresh_perishable()
            self._refresh_logs(tree, kind, product_id)

        LogEditForm(self.root, kind, log, on_save)

    def _delete_log(self, kind: str, tree: ttk.Treeview, product_id: int) -> None:
        sel = tree.selection()
        if not sel:
            messagebox.showwarning("Select", "Select a log to delete.")
            return
        log_id = int(sel[0])
        if not messagebox.askyesno("Confirm", "Delete this log?"):
            return
        if kind == "in":
            delete_in_log(log_id)
        else:
            delete_out_log(log_id)
        self.refresh_perishable()
        self._refresh_logs(tree, kind, product_id)

    def _refresh_logs(self, tree: ttk.Treeview, kind: str, product_id: int) -> None:
        tree.delete(*tree.get_children())
        logs = list_in_out_logs(kind, product_id)
        for log in logs:
            if kind == "in":
                tree.insert(
                    "",
                    "end",
                    iid=str(log["id"]),
                    values=(log["id"], log["delivery_date"], log["quantity"]),
                )
            else:
                tree.insert("", "end", iid=str(log["id"]), values=(log["id"], log["out_date"], log["out_time"], log["quantity"]))

    def _build_assets_tab(self, business: str, inventory_type: str) -> None:
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text=inventory_type)

        top = ttk.Frame(tab, padding=8)
        top.pack(fill="x")

        search_var = tk.StringVar()
        logo_label = ttk.Label(top)
        logo_label.grid(row=0, column=0, rowspan=2, sticky="w", padx=(0, 10), pady=(0, 4))
        self._apply_logo(logo_label, inventory_type)
        search_entry = ttk.Entry(top, textvariable=search_var, width=32)
        search_entry.grid(row=0, column=1, sticky="ew", padx=4)
        ttk.Label(top, text="Type").grid(row=0, column=2, sticky="w", padx=(8, 2))
        type_var = tk.StringVar(value="All")
        ttk.Combobox(top, textvariable=type_var, values=["All"] + ASSET_TYPES, state="readonly", width=18).grid(
            row=0, column=3, sticky="w"
        )
        ttk.Button(
            top, text="Search", command=lambda: self.refresh_assets(tree, business, inventory_type, search_var, type_var, sort_var)
        ).grid(row=0, column=4, sticky="w", padx=6)
        ttk.Button(
            top,
            text="Clear Search",
            command=lambda: self._clear_asset_search(tree, business, inventory_type, search_var, type_var, sort_var),
        ).grid(row=0, column=5, sticky="w")
        ttk.Button(
            top,
            text="Refresh",
            command=lambda: self.refresh_assets(tree, business, inventory_type, search_var, type_var, sort_var),
        ).grid(row=0, column=6, sticky="w", padx=6)
        ttk.Label(top, text="Sort").grid(row=0, column=7, sticky="w", padx=(8, 2))
        sort_var = tk.StringVar(value="Alphabetical (A-Z)")
        sort_combo = ttk.Combobox(
            top,
            textvariable=sort_var,
            values=[
                "Alphabetical (A-Z)",
                "Alphabetical (Z-A)",
                "Oldest (Added)",
                "Newest (Added)",
                "Oldest (Acquired)",
                "Newest (Acquired)",
                "Qty Low-High",
                "Qty High-Low",
            ],
            state="readonly",
            width=18,
        )
        sort_combo.grid(row=0, column=8, sticky="w")
        sort_combo.bind(
            "<<ComboboxSelected>>",
            lambda _e: self.refresh_assets(tree, business, inventory_type, search_var, type_var, sort_var),
        )
        buttons = ttk.Frame(top)
        buttons.grid(row=1, column=1, columnspan=8, sticky="w", pady=(1, 0))
        ttk.Button(buttons, text="Add", command=lambda: self.add_asset(tree, business, inventory_type, search_var, type_var)).pack(
            side="left", padx=2
        )
        ttk.Button(buttons, text="Edit", command=lambda: self.edit_asset(tree, business, inventory_type, search_var, type_var)).pack(
            side="left", padx=2
        )
        ttk.Button(buttons, text="Delete", command=lambda: self.delete_asset(tree, business, inventory_type, search_var, type_var)).pack(
            side="left", padx=2
        )
        ttk.Button(
            buttons,
            text="Duplicate",
            command=lambda: self.duplicate_asset(tree, business, inventory_type, search_var, type_var),
        ).pack(side="left", padx=2)
        ttk.Button(buttons, text="View Record", command=lambda: self.view_asset_record(tree, business, inventory_type)).pack(
            side="left", padx=2
        )
        if inventory_type not in ("Unica Non-Perishable", "HDN Warehouse"):
            ttk.Button(buttons, text="View Statuses", command=lambda: self.view_statuses(tree, business, inventory_type)).pack(
                side="left", padx=2
            )
        ttk.Button(buttons, text="Add Statuses", command=lambda: self.add_statuses(tree, business, inventory_type)).pack(
            side="left", padx=2
        )
        if inventory_type in ("Unica Non-Perishable", "HDN Warehouse"):
            ttk.Button(
                buttons,
                text="Add Qty",
                command=lambda: self.add_asset_qty(tree, business, inventory_type, search_var, type_var),
            ).pack(side="left", padx=2)
            ttk.Button(
                buttons,
                text="Add Acquisitions",
                command=lambda: self.add_acquisitions(tree, business, inventory_type),
            ).pack(side="left", padx=2)

        top.columnconfigure(1, weight=1)

        columns = [
            "no",
            "id",
            "name",
            "type",
            "brand",
            "model",
            "specifications",
            "series_number",
            "quantity",
            "total_spent",
            "location",
        ]
        headings = [
            "No.",
            "Id.",
            "Name",
            "Type",
            "Brand",
            "Model",
            "Specifications",
            "Series Number",
            "Qty",
            "Total Spent",
            "Location",
        ]
        tree = build_treeview(tab, columns, headings, image_heading="Picture", style="Photo.Treeview")
        tree.column("no", width=60, anchor="center")
        tree.column("id", width=80, anchor="center")
        tree.configure(selectmode="extended")
        self._register_tree(tree, tab)

        tree._sort_var = sort_var
        self.refresh_assets(tree, business, inventory_type, search_var, type_var, sort_var)

        summary_bar = ttk.Frame(tab, style="Card.TFrame", padding=(8, 4))
        summary_bar.pack(fill="x", padx=8, pady=(0, 8))
        summary = ttk.Label(summary_bar, text="Items: 0 | Total qty: 0", style="PanelMuted.TLabel")
        summary.pack(anchor="w")
        tree._summary_label = summary
        self._bind_summary_updates(tree)

        self._bind_dynamic_search(
            search_entry,
            lambda: self.refresh_assets(tree, business, inventory_type, search_var, type_var, sort_var),
        )
        search_entry.bind(
            "<Return>",
            lambda _e: self.refresh_assets(tree, business, inventory_type, search_var, type_var, sort_var),
        )

    def refresh_assets(
        self,
        tree: ttk.Treeview,
        business: str,
        inventory_type: str,
        search_var: tk.StringVar,
        type_var: tk.StringVar,
        sort_var: tk.StringVar | None = None,
    ) -> None:
        search = search_var.get().strip()
        type_filter = type_var.get().strip()
        if sort_var is None and hasattr(tree, "_sort_var"):
            sort_var = tree._sort_var
        sort_choice = sort_var.get().strip() if sort_var is not None else "Alphabetical (A-Z)"
        type_filter_val = None if not type_filter or type_filter == "All" else type_filter
        rows = list_assets(business, inventory_type, search if search else None, type_filter_val)
        if sort_choice == "Alphabetical (Z-A)":
            rows.sort(key=lambda r: (r.get("name") or "").lower(), reverse=True)
        elif sort_choice == "Oldest (Added)":
            rows.sort(key=lambda r: r.get("id") or 0)
        elif sort_choice == "Newest (Added)":
            rows.sort(key=lambda r: r.get("id") or 0, reverse=True)
        elif sort_choice == "Oldest (Acquired)":
            rows.sort(key=lambda r: r.get("latest_acquisition_date") or r.get("acquisition_date") or "")
        elif sort_choice == "Newest (Acquired)":
            rows.sort(key=lambda r: r.get("latest_acquisition_date") or r.get("acquisition_date") or "", reverse=True)
        elif sort_choice == "Qty Low-High":
            rows.sort(key=lambda r: float(r.get("quantity") or 0))
        elif sort_choice == "Qty High-Low":
            rows.sort(key=lambda r: float(r.get("quantity") or 0), reverse=True)
        else:
            rows.sort(key=lambda r: (r.get("name") or "").lower())
        tree.delete(*tree.get_children())
        if hasattr(tree, "_img_refs"):
            tree._img_refs = {}
        total_qty = 0
        for idx, row in enumerate(rows, start=1):
            try:
                total_qty += float(row.get("quantity") or 0)
            except ValueError:
                pass
            tree.insert(
                "",
                "end",
                iid=str(row["id"]),
                text="",
                values=(
                    idx,
                    row["id"],
                    row.get("name") or "",
                    row.get("type") or "",
                    row.get("brand") or "",
                    row.get("model") or "",
                    row.get("specifications") or "",
                    row.get("series_number") or "",
                    row["quantity"],
                    _format_money(row.get("total_spent")),
                    row.get("location") or "",
                ),
            )
            set_tree_image(tree, str(row["id"]), row.get("picture_path"))
        tree._summary_total_items = len(rows)
        tree._summary_total_qty = total_qty
        self._update_tree_summary(tree, len(rows), total_qty)

    def _clear_asset_search(
        self,
        tree: ttk.Treeview,
        business: str,
        inventory_type: str,
        search_var: tk.StringVar,
        type_var: tk.StringVar,
        sort_var: tk.StringVar | None = None,
    ) -> None:
        search_var.set("")
        self.refresh_assets(tree, business, inventory_type, search_var, type_var, sort_var)

    def add_asset(
        self,
        tree: ttk.Treeview,
        business: str,
        inventory_type: str,
        search_var: tk.StringVar,
        type_var: tk.StringVar,
    ) -> None:
        def on_save(data: dict) -> None:
            asset_id = add_asset(
                data["picture_path"],
                data["name"],
                data["brand"],
                data["model"],
                data["specifications"],
                data["series_number"],
                data["quantity"],
                data["location"],
                None,
                data["business"],
                data["type"],
                data["inventory_type"],
            )
            add_asset_acquisition(
                asset_id,
                data["acquisition_date"],
                data["acquisition_cost"],
                data.get("delivery_cost"),
                data["quantity"],
                data.get("shop_link"),
            )
            self.refresh_assets(tree, business, inventory_type, search_var, type_var)

        AssetForm(self.root, f"Add {inventory_type}", business, inventory_type, on_save)

    def edit_asset(
        self,
        tree: ttk.Treeview,
        business: str,
        inventory_type: str,
        search_var: tk.StringVar,
        type_var: tk.StringVar,
    ) -> None:
        sel = tree.selection()
        if not sel:
            messagebox.showwarning("Select", "Select a record to edit.")
            return
        asset_id = int(sel[0])
        row = next((r for r in list_assets(business, inventory_type) if r["id"] == asset_id), None)
        if not row:
            return

        def on_save(data: dict) -> None:
            update_asset(
                asset_id,
                data["picture_path"],
                data["name"],
                data["brand"],
                data["model"],
                data["specifications"],
                data["series_number"],
                data["quantity"],
                data["location"],
                None,
                data["type"],
            )
            self.refresh_assets(tree, business, inventory_type, search_var, type_var)

        AssetForm(self.root, f"Edit {inventory_type}", business, inventory_type, on_save, initial=dict(row))

    def delete_asset(
        self,
        tree: ttk.Treeview,
        business: str,
        inventory_type: str,
        search_var: tk.StringVar,
        type_var: tk.StringVar,
    ) -> None:
        sel = tree.selection()
        if not sel:
            messagebox.showwarning("Select", "Select a record to delete.")
            return
        ids = [int(item) for item in sel]
        label = "records" if len(ids) > 1 else "record"
        if messagebox.askyesno("Confirm", f"Delete {len(ids)} {label}?"):
            for asset_id in ids:
                delete_asset(asset_id)
            self.refresh_assets(tree, business, inventory_type, search_var, type_var)

    def duplicate_asset(
        self,
        tree: ttk.Treeview,
        business: str,
        inventory_type: str,
        search_var: tk.StringVar,
        type_var: tk.StringVar,
    ) -> None:
        sel = tree.selection()
        if not sel:
            messagebox.showwarning("Select", "Select a record to duplicate.")
            return
        asset_id = int(sel[0])
        new_id = duplicate_asset(asset_id)
        if not new_id:
            messagebox.showwarning("Error", "Unable to duplicate the selected record.")
            return
        self.refresh_assets(tree, business, inventory_type, search_var, type_var)

    def view_asset_record(self, tree: ttk.Treeview, business: str, inventory_type: str) -> None:
        sel = tree.selection()
        if not sel:
            messagebox.showwarning("Select", "Select a record first.")
            return
        asset_id = int(sel[0])
        row = next((r for r in list_assets(business, inventory_type) if r["id"] == asset_id), None)
        if not row:
            return

        win = tk.Toplevel(self.root)
        win.title(f"{row['name']} - Record")
        header = ttk.Frame(win, padding=8)
        header.pack(fill="x")

        img_label = ttk.Label(header, text="No image")
        img_label.pack(side="left", padx=(0, 10))
        preview = _load_preview_image(row.get("picture_path") or "")
        if preview is not None:
            img_label.configure(image=preview, text="")
            win._preview_image = preview

        info = ttk.Frame(header)
        info.pack(side="left", fill="x", expand=True)

        total_spent = _format_money(row.get("total_spent"))
        content = "\n".join(
            [
                f"ID: {row.get('id') or ''}",
                f"Name: {row.get('name') or ''}",
                f"Type: {row.get('type') or ''}",
                f"Brand: {row.get('brand') or ''}",
                f"Model: {row.get('model') or ''}",
                f"Specifications: {row.get('specifications') or ''}",
                f"Series Number: {row.get('series_number') or ''}",
                f"Quantity: {row.get('quantity') or ''}",
                f"Location: {row.get('location') or ''}",
                f"Business: {row.get('business') or ''}",
                f"Inventory Type: {row.get('inventory_type') or ''}",
                f"Latest Acquisition Date: {row.get('latest_acquisition_date') or ''}",
                f"Total Acquired Qty: {row.get('total_acquired_qty') or ''}",
                f"Total Spent: {total_spent}",
                f"Picture Path: {row.get('picture_path') or ''}",
            ]
        )
        make_readonly_text(info, content, height=14)

        status_columns = ("status", "qty")
        status_headings = ("Status", "Quantity")
        status_tree = build_treeview(win, status_columns, status_headings)
        status_tree.configure(height=5)
        for entry in list_asset_statuses(asset_id):
            status_tree.insert("", "end", values=(entry["status"], entry["quantity"]))

        acq_columns = ("date", "acquisition_cost", "delivery_cost", "qty", "shop_link")
        acq_headings = ("Acquisition Date", "Acquisition Cost", "Delivery Cost", "Quantity", "Shop")
        acq_tree = build_treeview(win, acq_columns, acq_headings)
        acq_tree.configure(height=5)
        for entry in list_asset_acquisitions(asset_id):
            acq_tree.insert(
                "",
                "end",
                values=(
                    entry.get("acquisition_date") or "",
                    _format_money(entry.get("acquisition_cost")),
                    _format_money(entry.get("delivery_cost")) if entry.get("delivery_cost") is not None else "",
                    entry.get("quantity") or "",
                    entry.get("shop_link") or "",
                ),
            )

        def _copy_shop(_event: tk.Event | None = None) -> None:
            sel = acq_tree.selection()
            if not sel:
                return
            values = acq_tree.item(sel[0], "values")
            if len(values) < 5:
                return
            shop_val = values[4]
            win.clipboard_clear()
            win.clipboard_append(str(shop_val))

        menu = tk.Menu(win, tearoff=0)
        menu.add_command(label="Copy Shop", command=_copy_shop)

        def _show_menu(event: tk.Event) -> str:
            row_id = acq_tree.identify_row(event.y)
            if row_id:
                acq_tree.selection_set(row_id)
                menu.tk_popup(event.x_root, event.y_root)
            return "break"

        acq_tree.bind("<Button-3>", _show_menu, add="+")
        acq_tree.bind("<Button-2>", _show_menu, add="+")

    def view_statuses(self, tree: ttk.Treeview, business: str, inventory_type: str) -> None:
        sel = tree.selection()
        if not sel:
            messagebox.showwarning("Select", "Select a record first.")
            return
        asset_id = int(sel[0])
        row = next((r for r in list_assets(business, inventory_type) if r["id"] == asset_id), None)
        if not row:
            return

        statuses = list_asset_statuses(asset_id)
        win = tk.Toplevel(self.root)
        win.title(f"{row['name']} - Statuses")
        top = ttk.Frame(win, padding=6)
        top.pack(fill="x")
        ttk.Label(top, text="Status breakdown for selected item").pack(side="left")

        columns = ("status", "qty")
        headings = ("Status", "Quantity")
        status_tree = build_treeview(win, columns, headings)
        for entry in statuses:
            status_tree.insert(
                "",
                "end",
                values=(entry["status"], entry["quantity"]),
            )

    def add_statuses(self, tree: ttk.Treeview, business: str, inventory_type: str) -> None:
        sel = tree.selection()
        if not sel:
            messagebox.showwarning("Select", "Select a record first.")
            return
        asset_id = int(sel[0])
        row = next((r for r in list_assets(business, inventory_type) if r["id"] == asset_id), None)
        if not row:
            return

        win = tk.Toplevel(self.root)
        win.title(f"{row['name']} - Add Statuses")
        top = ttk.Frame(win, padding=6)
        top.pack(fill="x")
        ttk.Label(top, text="Add status breakdown for this item").pack(side="left")

        columns = ("id", "status", "qty")
        headings = ("ID", "Status", "Quantity")
        tree_status = build_treeview(win, columns, headings)

        def refresh() -> None:
            tree_status.delete(*tree_status.get_children())
            for entry in list_asset_statuses(asset_id):
                tree_status.insert("", "end", iid=str(entry["id"]), values=(entry["id"], entry["status"], entry["quantity"]))

        def add_entry() -> None:
            add_win = tk.Toplevel(win)
            add_win.title("Add Status Entry")
            frm = ttk.Frame(add_win, padding=12)
            frm.pack(fill="both", expand=True)
            ttk.Label(frm, text="Qty").grid(row=0, column=0, sticky="w", pady=4)
            qty_var = tk.StringVar()
            ttk.Entry(frm, textvariable=qty_var).grid(row=0, column=1, sticky="ew")
            ttk.Label(frm, text="Status").grid(row=1, column=0, sticky="w", pady=4)
            status_var = tk.StringVar(value=ASSET_STATUSES[0])
            ttk.Combobox(frm, textvariable=status_var, values=ASSET_STATUSES, state="readonly").grid(
                row=1, column=1, sticky="ew"
            )

            def save() -> None:
                qty_raw = qty_var.get().strip()
                if not qty_raw:
                    messagebox.showwarning("Missing", "Quantity is required.")
                    return
                try:
                    qty_val = float(qty_raw)
                except ValueError:
                    messagebox.showwarning("Invalid", "Quantity must be a number.")
                    return
                existing = sum(float(r["quantity"]) for r in list_asset_statuses(asset_id))
                if existing + qty_val > float(row["quantity"]):
                    messagebox.showwarning("Limit", "Status quantities exceed the item quantity.")
                    return
                add_asset_status(asset_id, status_var.get(), qty_val)
                add_win.destroy()
                refresh()

            ttk.Button(frm, text="Save", command=save).grid(row=2, column=0, columnspan=2, pady=8)
            frm.columnconfigure(1, weight=1)

        def delete_entry() -> None:
            sel_row = tree_status.selection()
            if not sel_row:
                messagebox.showwarning("Select", "Select an entry to delete.")
                return
            if not messagebox.askyesno("Confirm", "Delete selected status entry?"):
                return
            delete_asset_status(int(sel_row[0]))
            refresh()

        btns = ttk.Frame(win, padding=6)
        btns.pack(fill="x")
        ttk.Button(btns, text="Add", command=add_entry).pack(side="left", padx=2)
        ttk.Button(btns, text="Delete", command=delete_entry).pack(side="left", padx=2)

        refresh()

    def add_acquisitions(self, tree: ttk.Treeview, business: str, inventory_type: str) -> None:
        sel = tree.selection()
        if not sel:
            messagebox.showwarning("Select", "Select a record first.")
            return
        asset_id = int(sel[0])
        row = next((r for r in list_assets(business, inventory_type) if r["id"] == asset_id), None)
        if not row:
            return

        win = tk.Toplevel(self.root)
        win.title(f"{row['name']} - Add Acquisitions")
        top = ttk.Frame(win, padding=6)
        top.pack(fill="x")
        ttk.Label(top, text="Add acquisition breakdown for this item").pack(side="left")

        columns = ("id", "date", "acquisition_cost", "delivery_cost", "qty", "shop_link")
        headings = ("ID", "Acquisition Date", "Acquisition Cost", "Delivery Cost", "Quantity", "Shop")
        tree_acq = build_treeview(win, columns, headings)

        def refresh() -> None:
            tree_acq.delete(*tree_acq.get_children())
            for entry in list_asset_acquisitions(asset_id):
                tree_acq.insert(
                    "",
                    "end",
                    iid=str(entry["id"]),
                    values=(
                        entry["id"],
                        entry["acquisition_date"],
                        _format_money(entry.get("acquisition_cost")),
                        _format_money(entry.get("delivery_cost")) if entry.get("delivery_cost") is not None else "",
                        entry["quantity"],
                        entry.get("shop_link") or "",
                    ),
                )

        def add_entry() -> None:
            add_win = tk.Toplevel(win)
            add_win.title("Add Acquisition Entry")
            frm = ttk.Frame(add_win, padding=12)
            frm.pack(fill="both", expand=True)

            date_var = tk.StringVar(value=date.today().strftime("%Y-%m-%d"))
            qty_var = tk.StringVar()
            cost_var = tk.StringVar()
            delivery_var = tk.StringVar()
            link_var = tk.StringVar()

            ttk.Label(frm, text="Acquisition Date").grid(row=0, column=0, sticky="w", pady=4)
            make_date_entry(frm, date_var).grid(row=0, column=1, sticky="ew")

            ttk.Label(frm, text="Quantity").grid(row=1, column=0, sticky="w", pady=4)
            ttk.Entry(frm, textvariable=qty_var).grid(row=1, column=1, sticky="ew")

            ttk.Label(frm, text="Acquisition Cost").grid(row=2, column=0, sticky="w", pady=4)
            ttk.Entry(frm, textvariable=cost_var).grid(row=2, column=1, sticky="ew")

            ttk.Label(frm, text="Delivery Cost").grid(row=3, column=0, sticky="w", pady=4)
            ttk.Entry(frm, textvariable=delivery_var).grid(row=3, column=1, sticky="ew")

            ttk.Label(frm, text="Shop").grid(row=4, column=0, sticky="w", pady=4)
            ttk.Entry(frm, textvariable=link_var).grid(row=4, column=1, sticky="ew")

            def save() -> None:
                date_val = date_var.get().strip()
                if not date_val:
                    messagebox.showwarning("Missing", "Acquisition date is required.")
                    return
                qty_raw = qty_var.get().strip()
                cost_raw = cost_var.get().strip()
                if not qty_raw or not cost_raw:
                    messagebox.showwarning("Missing", "Quantity and acquisition cost are required.")
                    return
                try:
                    qty_val = float(qty_raw)
                    cost_val = float(cost_raw)
                    delivery_raw = delivery_var.get().strip()
                    delivery_val = float(delivery_raw) if delivery_raw else None
                except ValueError:
                    messagebox.showwarning("Invalid", "Costs and quantity must be numbers.")
                    return
                existing = sum(float(r["quantity"]) for r in list_asset_acquisitions(asset_id))
                if existing + qty_val > float(row["quantity"]):
                    messagebox.showwarning("Limit", "Acquisition quantities exceed the item quantity.")
                    return
                add_asset_acquisition(asset_id, date_val, cost_val, delivery_val, qty_val, link_var.get().strip() or None)
                add_win.destroy()
                refresh()

            ttk.Button(frm, text="Save", command=save).grid(row=5, column=0, columnspan=2, pady=8)
            frm.columnconfigure(1, weight=1)

        def edit_entry() -> None:
            sel_row = tree_acq.selection()
            if not sel_row:
                messagebox.showwarning("Select", "Select an entry to edit.")
                return
            acq_id = int(sel_row[0])
            current = next((r for r in list_asset_acquisitions(asset_id) if r["id"] == acq_id), None)
            if not current:
                return

            edit_win = tk.Toplevel(win)
            edit_win.title("Edit Acquisition Entry")
            frm = ttk.Frame(edit_win, padding=12)
            frm.pack(fill="both", expand=True)

            date_var = tk.StringVar(value=current.get("acquisition_date") or "")
            qty_var = tk.StringVar(value=str(current.get("quantity") or ""))
            cost_var = tk.StringVar(value=str(current.get("acquisition_cost") or ""))
            delivery_var = tk.StringVar(
                value=str(current.get("delivery_cost") or "") if current.get("delivery_cost") is not None else ""
            )
            link_var = tk.StringVar(value=current.get("shop_link") or "")

            ttk.Label(frm, text="Acquisition Date").grid(row=0, column=0, sticky="w", pady=4)
            make_date_entry(frm, date_var).grid(row=0, column=1, sticky="ew")

            ttk.Label(frm, text="Quantity").grid(row=1, column=0, sticky="w", pady=4)
            ttk.Entry(frm, textvariable=qty_var).grid(row=1, column=1, sticky="ew")

            ttk.Label(frm, text="Acquisition Cost").grid(row=2, column=0, sticky="w", pady=4)
            ttk.Entry(frm, textvariable=cost_var).grid(row=2, column=1, sticky="ew")

            ttk.Label(frm, text="Delivery Cost").grid(row=3, column=0, sticky="w", pady=4)
            ttk.Entry(frm, textvariable=delivery_var).grid(row=3, column=1, sticky="ew")

            ttk.Label(frm, text="Shop").grid(row=4, column=0, sticky="w", pady=4)
            ttk.Entry(frm, textvariable=link_var).grid(row=4, column=1, sticky="ew")

            def save() -> None:
                date_val = date_var.get().strip()
                if not date_val:
                    messagebox.showwarning("Missing", "Acquisition date is required.")
                    return
                qty_raw = qty_var.get().strip()
                cost_raw = cost_var.get().strip()
                if not qty_raw or not cost_raw:
                    messagebox.showwarning("Missing", "Quantity and acquisition cost are required.")
                    return
                try:
                    qty_val = float(qty_raw)
                    cost_val = float(cost_raw)
                    delivery_raw = delivery_var.get().strip()
                    delivery_val = float(delivery_raw) if delivery_raw else None
                except ValueError:
                    messagebox.showwarning("Invalid", "Costs and quantity must be numbers.")
                    return
                existing = sum(float(r["quantity"]) for r in list_asset_acquisitions(asset_id) if r["id"] != acq_id)
                if existing + qty_val > float(row["quantity"]):
                    messagebox.showwarning("Limit", "Acquisition quantities exceed the item quantity.")
                    return
                update_asset_acquisition(acq_id, date_val, cost_val, delivery_val, qty_val, link_var.get().strip() or None)
                edit_win.destroy()
                refresh()

            ttk.Button(frm, text="Save", command=save).grid(row=5, column=0, columnspan=2, pady=8)
            frm.columnconfigure(1, weight=1)

        def delete_entry() -> None:
            sel_row = tree_acq.selection()
            if not sel_row:
                messagebox.showwarning("Select", "Select an entry to delete.")
                return
            if not messagebox.askyesno("Confirm", "Delete selected acquisition entry?"):
                return
            delete_asset_acquisition(int(sel_row[0]))
            refresh()

        btns = ttk.Frame(win, padding=6)
        btns.pack(fill="x")
        ttk.Button(btns, text="Add", command=add_entry).pack(side="left", padx=2)
        ttk.Button(btns, text="Edit", command=edit_entry).pack(side="left", padx=2)
        ttk.Button(btns, text="Delete", command=delete_entry).pack(side="left", padx=2)

        refresh()

    def add_asset_qty(
        self,
        tree: ttk.Treeview,
        business: str,
        inventory_type: str,
        search_var: tk.StringVar,
        type_var: tk.StringVar,
    ) -> None:
        sel = tree.selection()
        if not sel:
            messagebox.showwarning("Select", "Select a record first.")
            return
        asset_id = int(sel[0])
        row = next((r for r in list_assets(business, inventory_type) if r["id"] == asset_id), None)
        if not row:
            return

        win = tk.Toplevel(self.root)
        win.title(f"{row['name']} - Add Quantity")
        frm = ttk.Frame(win, padding=12)
        frm.pack(fill="both", expand=True)

        qty_var = tk.StringVar()

        ttk.Label(frm, text="Additional Quantity").grid(row=0, column=0, sticky="w", pady=4)
        ttk.Entry(frm, textvariable=qty_var).grid(row=0, column=1, sticky="ew")

        def save() -> None:
            raw = qty_var.get().strip()
            if not raw:
                messagebox.showwarning("Missing", "Quantity is required.")
                return
            try:
                add_val = float(raw)
            except ValueError:
                messagebox.showwarning("Invalid", "Quantity must be a number.")
                return
            if add_val <= 0:
                messagebox.showwarning("Invalid", "Quantity must be greater than zero.")
                return
            new_qty = float(row.get("quantity") or 0) + add_val
            update_asset(
                asset_id,
                row.get("picture_path"),
                row.get("name") or "",
                row.get("brand"),
                row.get("model"),
                row.get("series_number"),
                new_qty,
                row.get("location"),
                None,
                row.get("type") or ASSET_TYPES[0],
            )
            win.destroy()
            self.refresh_assets(tree, business, inventory_type, search_var, type_var)

        ttk.Button(frm, text="Save", command=save).grid(row=1, column=0, columnspan=2, pady=8)
        frm.columnconfigure(1, weight=1)

    def _build_blank_tab(self, title: str) -> None:
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text=title)
        msg = ttk.Label(tab, text="Blank screen for future farm inventory.")
        msg.pack(pady=40)

    def _build_users_tab(self) -> None:
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Users")

        top = ttk.Frame(tab, padding=8)
        top.pack(fill="x")
        ttk.Button(top, text="Add User", command=lambda: self.add_user(tree)).pack(side="left")
        ttk.Button(top, text="Edit User", command=lambda: self.edit_user(tree)).pack(side="left", padx=6)
        ttk.Button(top, text="Delete User", command=lambda: self.delete_user(tree)).pack(side="left")

        columns = ["id", "username", "business", "is_admin"]
        headings = ["ID", "Username", "Business", "Admin"]
        tree = build_treeview(tab, columns, headings)
        self.refresh_users(tree)

    def refresh_users(self, tree: ttk.Treeview) -> None:
        tree.delete(*tree.get_children())
        for user in list_users():
            tree.insert("", "end", iid=str(user["id"]), values=(user["id"], user["username"], user["business"], user["is_admin"]))

    def add_user(self, tree: ttk.Treeview) -> None:
        def on_save(data: dict) -> None:
            add_user(data["username"], data["password"], data["businesses"], data["is_admin"])
            self.refresh_users(tree)

        UserForm(self.root, "Add User", on_save)

    def edit_user(self, tree: ttk.Treeview) -> None:
        sel = tree.selection()
        if not sel:
            messagebox.showwarning("Select", "Select a user to edit.")
            return
        user_id = int(sel[0])
        user = next((u for u in list_users() if u["id"] == user_id), None)
        if not user:
            return

        def on_save(data: dict) -> None:
            update_user(user_id, data["password"], data["businesses"], data["is_admin"])
            self.refresh_users(tree)

        UserForm(self.root, "Edit User", on_save, initial=user)

    def delete_user(self, tree: ttk.Treeview) -> None:
        sel = tree.selection()
        if not sel:
            messagebox.showwarning("Select", "Select a user to delete.")
            return
        user_id = int(sel[0])
        if messagebox.askyesno("Confirm", "Delete this user?"):
            delete_user(user_id)
            self.refresh_users(tree)


def main() -> None:
    root = tk.Tk()
    apply_theme(root)
    enable_smooth_resize(root)
    configure_input_shortcuts(root)
    configure_context_menu(root)
    try:
        init_db()
    except Exception as exc:
        messagebox.showerror(
            "Database Error",
            f"Unable to connect to PostgreSQL.\n\n{exc}\n\nCheck db_config.json or AMAN_DB_URL.",
            parent=root,
        )
        root.destroy()
        return

    def launch(user: dict) -> None:
        MainWindow(root, user)

    LoginWindow(root, launch)
    root.mainloop()


if __name__ == "__main__":
    main()
