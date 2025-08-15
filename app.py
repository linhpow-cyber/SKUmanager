#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
TC Premium Tile SKU Manager - single-file app for macOS (Python 3.7) and Windows

- products.xlsx stores product records (Timestamp, BrandCode, BrandName, BrandID, SizeLabel, SizeCode,
  SurfaceLabel, SurfaceCode, MattPolished, SPCode, SKU, CommercialName, Faces, Batch, CountryPrefix, CompanyPrefix, EAN13, ImagePaths, Notes)
- deleted_products.xlsx stores deleted product records (same structure)
- images/<SKU>/ contains saved product images, barcode.png, qrcode.png
- Images > 2000px resized (max dimension 2000), saved as PNG
- SPCode auto-increments per (BrandCode + SizeCode)
- SKU format: [Brand][Size][Matt/Polished(0/1)][SPCode] (e.g., VE606001001, VE606000001)
- Barcode (EAN-13) built using CountryPrefix + CompanyPrefix + BrandID + SPCode (checksum)
- QR contains https://thangcuongtiles.com/<SPCode>
- Delete feature: Confirm deletion, move record to deleted_products.xlsx
- Surface: Checkboxes for White Body, Microcid Glaze, Scratch-Resistant Glaze, Crystal Glaze, Deep Color (optional)
- Matt/Polished: Radio button for Matt (0) or Polished (1), no None option
- Commercial Name: Dialog with 20 Latin Prefixes, 20 Colors, 20 Stone Types, editable, saved as "Gạch porcelain kích thước [Size] [Name]"
- Viewer Preview: CommercialName (large font), full SPCode (smaller), Surface, Matt/Polished, images in vertical stack, barcode/QR code (bottom), scrollable
- SKU List: Scrollable Treeview with proper scrollbars and wider columns
- Image Context Menu: Right-click on images in Preview for Copy Image (to clipboard, PNG, Windows/macOS) and Delete Image (remove from ImagePaths and file)
- Viewer Edit: Add Images button to append images, Add Notes button to append or create Notes
"""

from __future__ import print_function
import os
import sys
import shutil
import datetime as dt
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, Toplevel
from PIL import Image, ImageTk
import pandas as pd
import io

# Optional packages for barcode and qrcode
try:
    import barcode
    from barcode.writer import ImageWriter
except ImportError:
    barcode = None

try:
    import qrcode
except ImportError:
    qrcode = None

# ---------------- Config ----------------
DATA_FILE = "products.xlsx"
DELETED_DATA_FILE = "deleted_products.xlsx"
IMAGES_ROOT = "images"
LOGO_FILE = "logo.png"  # small logo (optional)
MAX_DIM = 2000  # px

# Fixed brands mapping
BRAND_LIST = [
    {"code": "VE", "name": "Vesta", "id": "0"},
    {"code": "OM", "name": "One Max", "id": "1"},
    {"code": "GA", "name": "Granca", "id": "2"},
    {"code": "SA", "name": "STA", "id": "3"},
]
BRAND_CODE_TO_ID = {b["code"]: b["id"] for b in BRAND_LIST}
BRAND_CODE_TO_NAME = {b["code"]: b["name"] for b in BRAND_LIST}
BRAND_CODES = [b["code"] for b in BRAND_LIST]

# Sizes
SIZES = {
    "60x60": "6060",
    "80x80": "8080",
    "40x80": "4080",
    "60x120": "6120",
    "100x100": "1010",
    "120x120": "1212",
    "80x160": "8160",
    "100x200": "1020",
    "120x240": "1224",
}

# Surface options (checkboxes)
SURFACE_OPTIONS = {
    "White Body": "W",
    "Microcid Glaze": "M",
    "Scratch-Resistant Glaze": "S",
    "Crystal Glaze": "C",
    "Deep Color": "D"
}

# Matt/Polished options (radio buttons)
MATT_POLISHED_OPTIONS = {
    "Matt": "0",
    "Polished": "1"
}

# Stone types (20)
STONE_TYPES = [
    "Marble", "Granite", "Quartz", "Porcelain", "Ceramic",
    "Slate", "Travertine", "Limestone", "Sandstone", "Basalt",
    "Onyx", "Schist", "Soapstone", "Terrazzo", "Obsidian",
    "Gneiss", "Tuff", "Breccia", "Porphyry", "Dolomite"
]

# Colors (20)
COLORS = [
    "White", "Black", "Grey", "Beige", "Brown",
    "Ivory", "Cream", "Charcoal", "Slate", "Taupe",
    "Blue", "Green", "Red", "Gold", "Silver",
    "Pearl", "Ebony", "Sand", "Mocha", "Azure"
]

# Latin prefixes (20)
LATIN_PREFIXES = [
    "Lux", "Prima", "Nobilis", "Regal", "Vita",
    "Aurea", "Stella", "Magnus", "Opus", "Elegans",
    "Fortis", "Clarus", "Vera", "Splendid", "Grandis",
    "Purus", "Nexus", "Arca", "Divus", "Optima"
]

# Defaults for EAN building
DEFAULT_COUNTRY_PREFIX = "893"
DEFAULT_COMPANY_PREFIX = "12345"

# Excel columns
COLUMNS = [
    "Timestamp", "BrandCode", "BrandName", "BrandID",
    "SizeLabel", "SizeCode",
    "SurfaceLabel", "SurfaceCode",
    "MattPolished", "SPCode", "SKU", "CommercialName",
    "Faces", "Batch",
    "CountryPrefix", "CompanyPrefix", "EAN13",
    "ImagePaths", "Notes"
]

# Ensure storage
def ensure_storage():
    if not os.path.isdir(IMAGES_ROOT):
        os.makedirs(IMAGES_ROOT, exist_ok=True)
    if not os.path.isfile(DATA_FILE):
        df = pd.DataFrame(columns=COLUMNS)
        df.to_excel(DATA_FILE, index=False)
    if not os.path.isfile(DELETED_DATA_FILE):
        df = pd.DataFrame(columns=COLUMNS)
        df.to_excel(DELETED_DATA_FILE, index=False)

# Load & save DataFrame
def load_df():
    ensure_storage()
    try:
        df = pd.read_excel(DATA_FILE, dtype=str)
    except Exception:
        df = pd.DataFrame(columns=COLUMNS)
    for c in COLUMNS:
        if c not in df.columns:
            df[c] = ""
    if "SPCode" in df.columns:
        df["SPCode"] = df["SPCode"].fillna("").apply(lambda x: str(x).zfill(3) if str(x).strip().isdigit() else x)
    return df

def load_deleted_df():
    ensure_storage()
    try:
        df = pd.read_excel(DELETED_DATA_FILE, dtype=str)
    except Exception:
        df = pd.DataFrame(columns=COLUMNS)
    for c in COLUMNS:
        if c not in df.columns:
            df[c] = ""
    return df

def save_df(df):
    df.to_excel(DATA_FILE, index=False)

def save_deleted_df(df):
    df.to_excel(DELETED_DATA_FILE, index=False)

# EAN-13 utils
def ean13_checkdigit(base12):
    digits = [int(d) for d in base12]
    odd_sum = sum(digits[0::2])
    even_sum = sum(digits[1::2]) * 3
    total = odd_sum + even_sum
    check = (10 - (total % 10)) % 10
    return str(check)

def build_ean13(country_prefix, company_prefix, brand_id, spcode):
    base = "".join(ch for ch in (str(country_prefix) + str(company_prefix) + str(brand_id) + str(int(spcode)).zfill(3)) if ch.isdigit())
    if len(base) < 12:
        base12 = base.rjust(12, "0")
    else:
        base12 = base[-12:]
    return base12 + ean13_checkdigit(base12)

# SKU builder (Brand + Size + Matt/Polished + SPCode)
def build_sku(brand_code, size_code, matt_polished, spcode):
    return "{}{}{}{}".format(brand_code, size_code, matt_polished, str(int(spcode)).zfill(3))

# Build full SPCode for display
def build_full_spcode(brand_code, size_code, matt_polished, spcode):
    return "{}{}{}{}".format(brand_code, size_code, matt_polished, str(int(spcode)).zfill(3))

# Resize & save image
def resize_and_save(src, dst_base):
    try:
        with Image.open(src) as im:
            w, h = im.size
            if max(w, h) > MAX_DIM:
                scale = MAX_DIM / float(max(w, h))
                nw = int(round(w * scale))
                nh = int(round(h * scale))
                im = im.resize((nw, nh), Image.LANCZOS)
                dst = dst_base + ".png"
                im.save(dst, format="PNG", optimize=True)
                return dst
            else:
                ext = os.path.splitext(src)[1].lower()
                dst = dst_base + ext
                shutil.copy2(src, dst)
                return dst
    except Exception:
        try:
            ext = os.path.splitext(src)[1].lower()
            dst = dst_base + ext
            shutil.copy2(src, dst)
            return dst
        except Exception:
            return None

# Copy image to clipboard
def copy_image_to_clipboard(img_path, root):
    try:
        with Image.open(img_path) as img:
            output = io.BytesIO()
            img.convert("RGB").save(output, format="PNG")
            data = output.getvalue()
            output.close()
            root.clipboard_clear()
            if sys.platform == "win32":
                from tkinter import TclError
                try:
                    root.clipboard_append(data, type="image/png")
                except TclError:
                    # Fallback for Windows
                    import win32clipboard
                    output = BytesIO()
                    img.convert("RGB").save(output, format="BMP")
                    bmp_data = output.getvalue()[14:]  # Skip BMP header
                    output.close()
                    win32clipboard.OpenClipboard()
                    win32clipboard.EmptyClipboard()
                    win32clipboard.SetClipboardData(win32clipboard.CF_DIB, bmp_data)
                    win32clipboard.CloseClipboard()
            else:
                # macOS and Linux (requires Tk 8.6+)
                root.clipboard_append(data, type="image/png")
        messagebox.showinfo("Success", "Image copied to clipboard")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to copy image: {str(e)}")

# Delete image from product
def delete_image_from_product(sku, img_path):
    df = load_df()
    row = df[df["SKU"].astype(str) == sku]
    if row.empty:
        return False
    row = row.iloc[0]
    images = str(row.get("ImagePaths","")).split(";") if row.get("ImagePaths","") else []
    if img_path in images:
        images.remove(img_path)
        try:
            os.remove(img_path)
        except Exception:
            pass
        df.loc[df["SKU"] == sku, "ImagePaths"] = ";".join(images)
        save_df(df)
        return True
    return False

# Barcode and QR generation
def generate_barcode_qr(ean13_str, sku_folder, spcode):
    bc_path = None
    qr_path = None
    try:
        if barcode is not None:
            writer = ImageWriter()
            EAN = barcode.get_barcode_class('ean13')
            num = str(ean13_str)
            if len(num) == 13 and num.isdigit():
                bc_obj = EAN(num, writer=writer)
                bc_fname = os.path.join(sku_folder, "{}_barcode".format(spcode))
                bc_obj.save(bc_fname)
                if os.path.isfile(bc_fname + ".png"):
                    bc_path = bc_fname + ".png"
    except Exception:
        bc_path = None
    try:
        if qrcode is not None:
            qr_url = "https://thangcuongtiles.com/{}".format(spcode)
            qr = qrcode.QRCode(box_size=6, border=2)
            qr.add_data(qr_url)
            qr.make(fit=True)
            img = qr.make_image(fill_color="black", back_color="white")
            qr_path = os.path.join(sku_folder, "{}_qrcode.png".format(spcode))
            img.save(qr_path)
    except Exception:
        qr_path = None
    return bc_path, qr_path

# Auto SPCode generation
def next_spcode_for(brand_code, size_code):
    df = load_df()
    if df.empty:
        return "001"
    sel = df[(df["BrandCode"] == brand_code) & (df["SizeCode"] == size_code)]
    if sel.empty:
        return "001"
    try:
        nums = sel["SPCode"].apply(lambda x: int(str(x).strip()) if str(x).strip().isdigit() else 0)
        nxt = int(nums.max()) + 1
        if nxt < 1:
            nxt = 1
        return str(nxt).zfill(3)
    except Exception:
        return "001"

# Open folder helper
def open_folder(path):
    try:
        if sys.platform == "win32":
            os.startfile(path)
        else:
            os.system('open "{}"'.format(path))
    except Exception as e:
        messagebox.showerror("Open folder error", str(e))

# ---------------- Commercial Name Dialog ----------------
class CommercialNameDialog(Toplevel):
    def __init__(self, parent, current_name="", size_label=""):
        super().__init__(parent)
        self.title("Set Commercial Name")
        self.geometry("400x350")
        self.resizable(False, False)
        self.result = None
        self.current_name = current_name
        self.size_label = size_label
        self._build_ui()

    def _build_ui(self):
        pad = 10
        frame = ttk.Frame(self, padding=pad)
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text="Latin Prefix").grid(row=0, column=0, sticky="w", pady=4)
        self.prefix_var = tk.StringVar(value=LATIN_PREFIXES[0])
        ttk.Combobox(frame, textvariable=self.prefix_var, values=LATIN_PREFIXES, state="readonly").grid(row=0, column=1, sticky="w", padx=4)

        ttk.Label(frame, text="Color").grid(row=1, column=0, sticky="w", pady=4)
        self.color_var = tk.StringVar(value=COLORS[0])
        ttk.Combobox(frame, textvariable=self.color_var, values=COLORS, state="readonly").grid(row=1, column=1, sticky="w", padx=4)

        ttk.Label(frame, text="Stone Type").grid(row=2, column=0, sticky="w", pady=4)
        self.stone_var = tk.StringVar(value=STONE_TYPES[0])
        ttk.Combobox(frame, textvariable=self.stone_var, values=STONE_TYPES, state="readonly").grid(row=2, column=1, sticky="w", padx=4)

        ttk.Button(frame, text="Suggest Name", command=self.suggest_name).grid(row=3, column=0, columnspan=2, pady=8)

        ttk.Label(frame, text="Commercial Name").grid(row=4, column=0, sticky="w", pady=4)
        self.name_var = tk.StringVar(value=self.current_name)
        ttk.Entry(frame, textvariable=self.name_var, width=40).grid(row=4, column=1, sticky="w", padx=4)

        btns = ttk.Frame(frame)
        btns.grid(row=5, column=0, columnspan=2, pady=10)
        ttk.Button(btns, text="OK", command=self.on_ok).pack(side="left", padx=5)
        ttk.Button(btns, text="Cancel", command=self.on_cancel).pack(side="left", padx=5)

    def suggest_name(self):
        prefix = self.prefix_var.get()
        color = self.color_var.get()
        stone = self.stone_var.get()
        suggested = f"{prefix} {color} {stone}"
        self.name_var.set(suggested)

    def on_ok(self):
        base_name = self.name_var.get().strip()
        if not base_name:
            messagebox.showerror("Error", "Commercial Name cannot be empty")
            return
        full_name = f"Gạch porcelain kích thước {self.size_label} {base_name}"
        self.result = full_name
        self.destroy()

    def on_cancel(self):
        self.result = None
        self.destroy()

# ---------------- Add Notes Dialog ----------------
class AddNotesDialog(Toplevel):
    def __init__(self, parent, current_notes=""):
        super().__init__(parent)
        self.title("Add Notes")
        self.geometry("400x200")
        self.resizable(False, False)
        self.result = None
        self.current_notes = current_notes
        self._build_ui()

    def _build_ui(self):
        pad = 10
        frame = ttk.Frame(self, padding=pad)
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text="Current Notes:").pack(anchor="w", pady=4)
        ttk.Label(frame, text=self.current_notes or "None", wraplength=350).pack(anchor="w", pady=4)

        ttk.Label(frame, text="Add Note:").pack(anchor="w", pady=4)
        self.note_var = tk.StringVar()
        ttk.Entry(frame, textvariable=self.note_var, width=50).pack(anchor="w", padx=4)

        btns = ttk.Frame(frame)
        btns.pack(pady=10)
        ttk.Button(btns, text="OK", command=self.on_ok).pack(side="left", padx=5)
        ttk.Button(btns, text="Cancel", command=self.on_cancel).pack(side="left", padx=5)

    def on_ok(self):
        new_note = self.note_var.get().strip()
        if not new_note:
            messagebox.showerror("Error", "Note cannot be empty")
            return
        if self.current_notes:
            self.result = f"{self.current_notes}; {new_note}"
        else:
            self.result = new_note
        self.destroy()

    def on_cancel(self):
        self.result = None
        self.destroy()

# ---------------- UI: Entry Tab ----------------
class EntryTab(ttk.Frame):
    def __init__(self, master, viewer_refresh_callback=None):
        super().__init__(master)
        self.viewer_refresh = viewer_refresh_callback
        self.selected_images = []
        self._thumb = None
        self._build_ui()

    def _build_ui(self):
        pad = 6
        top = ttk.Frame(self)
        top.pack(fill="x", padx=pad, pady=pad)
        if os.path.isfile(LOGO_FILE):
            try:
                img = Image.open(LOGO_FILE)
                img.thumbnail((40, 40))
                self.logo_photo = ImageTk.PhotoImage(img)
                ttk.Label(top, image=self.logo_photo).pack(side="left", padx=4)
            except Exception:
                pass
        form = ttk.Frame(self, padding=pad)
        form.pack(fill="x", padx=pad)

        ttk.Label(form, text="Brand").grid(row=0, column=0, sticky="w", padx=4, pady=4)
        self.brand_var = tk.StringVar(value=BRAND_CODES[0])
        self.brand_cb = ttk.Combobox(form, textvariable=self.brand_var, values=BRAND_CODES, state="readonly", width=8)
        self.brand_cb.grid(row=0, column=1, sticky="w", padx=4)

        ttk.Label(form, text="Brand Name").grid(row=0, column=2, sticky="w")
        self.brand_name_var = tk.StringVar(value=BRAND_CODE_TO_NAME[self.brand_var.get()])
        ttk.Label(form, textvariable=self.brand_name_var).grid(row=0, column=3, sticky="w", padx=4)
        self.brand_cb.bind("<<ComboboxSelected>>", self.on_brand_change)

        ttk.Label(form, text="Size").grid(row=1, column=0, sticky="w", padx=4)
        self.size_var = tk.StringVar(value=list(SIZES.keys())[0])
        self.size_cb = ttk.Combobox(form, textvariable=self.size_var, values=list(SIZES.keys()), state="readonly", width=12)
        self.size_cb.grid(row=1, column=1, sticky="w", padx=4)

        ttk.Label(form, text="Surface").grid(row=2, column=0, sticky="w", padx=4, pady=4)
        self.surface_vars = {key: tk.BooleanVar(value=False) for key in SURFACE_OPTIONS}
        self.surface_frame = ttk.Frame(form)
        self.surface_frame.grid(row=2, column=1, columnspan=3, sticky="w", padx=4)
        for i, (label, code) in enumerate(SURFACE_OPTIONS.items()):
            ttk.Checkbutton(self.surface_frame, text=label, variable=self.surface_vars[label]).grid(row=0, column=i, padx=2)

        ttk.Label(form, text="Matt/Polished").grid(row=3, column=0, sticky="w", padx=4, pady=4)
        self.matt_polished_var = tk.StringVar(value="0")
        self.mp_frame = ttk.Frame(form)
        self.mp_frame.grid(row=3, column=1, columnspan=2, sticky="w", padx=4)
        ttk.Radiobutton(self.mp_frame, text="Matt", value="0", variable=self.matt_polished_var).grid(row=0, column=0, padx=4)
        ttk.Radiobutton(self.mp_frame, text="Polished", value="1", variable=self.matt_polished_var).grid(row=0, column=1, padx=4)

        ttk.Label(form, text="Commercial Name").grid(row=4, column=0, sticky="w", padx=4, pady=4)
        self.commercial_name_var = tk.StringVar(value="")
        self.commercial_name_entry = ttk.Entry(form, textvariable=self.commercial_name_var, width=50)
        self.commercial_name_entry.grid(row=4, column=1, columnspan=3, sticky="w", padx=4)
        ttk.Button(form, text="Set Name", command=self.set_commercial_name).grid(row=4, column=4, sticky="w", padx=4)

        ttk.Label(form, text="Faces").grid(row=5, column=0, sticky="w", padx=4, pady=4)
        self.faces_var = tk.StringVar(value="1")
        self.faces_entry = ttk.Entry(form, textvariable=self.faces_var, width=6)
        self.faces_entry.grid(row=5, column=1, sticky="w", padx=4)

        ttk.Label(form, text="SPCode").grid(row=5, column=2, sticky="w", padx=4)
        self.spcode_var = tk.StringVar(value="")
        self.spcode_entry = ttk.Entry(form, textvariable=self.spcode_var, width=8)
        self.spcode_entry.grid(row=5, column=3, sticky="w", padx=4)
        ttk.Button(form, text="Generate SPCode", command=self.on_generate_spcode).grid(row=5, column=4, sticky="w", padx=6)

        ttk.Label(form, text="Batch (opt)").grid(row=6, column=0, sticky="w", padx=4, pady=4)
        self.batch_var = tk.StringVar(value="")
        self.batch_entry = ttk.Entry(form, textvariable=self.batch_var, width=12)
        self.batch_entry.grid(row=6, column=1, sticky="w", padx=4)

        ttk.Label(form, text="CountryPrefix").grid(row=6, column=2, sticky="w", padx=4)
        self.country_var = tk.StringVar(value=DEFAULT_COUNTRY_PREFIX)
        self.country_entry = ttk.Entry(form, textvariable=self.country_var, width=8)
        self.country_entry.grid(row=6, column=3, sticky="w", padx=4)

        ttk.Label(form, text="CompanyPrefix").grid(row=6, column=4, sticky="w", padx=4)
        self.company_var = tk.StringVar(value=DEFAULT_COMPANY_PREFIX)
        self.company_entry = ttk.Entry(form, textvariable=self.company_var, width=12)
        self.company_entry.grid(row=6, column=5, sticky="w", padx=4)

        ttk.Label(form, text="Notes").grid(row=7, column=0, sticky="w", padx=4, pady=4)
        self.notes_var = tk.StringVar(value="")
        self.notes_entry = ttk.Entry(form, textvariable=self.notes_var, width=60)
        self.notes_entry.grid(row=7, column=1, columnspan=5, sticky="w", padx=4)

        btns = ttk.Frame(self)
        btns.pack(fill="x", padx=pad, pady=pad)
        ttk.Button(btns, text="Add Images…", command=self.add_images).pack(side="left", padx=4)
        ttk.Button(btns, text="Clear Images", command=self.clear_images).pack(side="left", padx=4)
        self.save_button = ttk.Button(btns, text="Save Product", command=self.save_product)
        self.save_button.pack(side="left", padx=12)
        ttk.Button(btns, text="Reset", command=self.reset_form).pack(side="left", padx=4)
        ttk.Button(btns, text="Open data folder", command=lambda: open_folder(os.getcwd())).pack(side="right", padx=4)

        mid = ttk.PanedWindow(self, orient="horizontal")
        mid.pack(fill="both", expand=True, padx=pad, pady=pad)
        left = ttk.Frame(mid); mid.add(left, weight=1)
        right = ttk.Frame(mid); mid.add(right, weight=2)

        ttk.Label(left, text="Selected images").pack(anchor="w")
        self.img_listbox = tk.Listbox(left, height=12)
        self.img_listbox.pack(fill="both", expand=True, padx=4, pady=4)
        self.img_listbox.bind("<<ListboxSelect>>", self.on_img_select)

        ttk.Label(right, text="Preview").pack(anchor="w")
        self.preview_label = ttk.Label(right, text="No image", relief="sunken")
        self.preview_label.pack(fill="both", expand=True, padx=4, pady=4)

        self.status_var = tk.StringVar(value="Ready")
        ttk.Label(self, textvariable=self.status_var, relief="sunken", anchor="w").pack(side="bottom", fill="x")

        self.reset_form()

    def on_brand_change(self, event=None):
        b = self.brand_var.get()
        self.brand_name_var.set(BRAND_CODE_TO_NAME.get(b, ""))

    def on_generate_spcode(self):
        brand = self.brand_var.get()
        size_label = self.size_var.get()
        size_code = SIZES.get(size_label)
        sp = next_spcode_for(brand, size_code)
        self.spcode_var.set(sp)
        self.status_var.set("Generated SPCode {}".format(sp))

    def set_commercial_name(self):
        size_label = self.size_var.get()
        dialog = CommercialNameDialog(self, self.commercial_name_var.get(), size_label)
        self.wait_window(dialog)
        if dialog.result is not None:
            self.commercial_name_var.set(dialog.result)
            self.status_var.set(f"Set Commercial Name: {dialog.result}")

    def add_images(self):
        files = filedialog.askopenfilenames(title="Select images", filetypes=(("Images", "*.jpg *.jpeg *.png *.bmp *.webp"), ("All files", "*.*")))
        if not files:
            return
        for f in files:
            if f not in self.selected_images:
                self.selected_images.append(f)
                self.img_listbox.insert("end", os.path.basename(f))
        self.status_var.set("Selected {} images".format(len(self.selected_images)))
        if self.selected_images:
            self.show_preview(self.selected_images[0])

    def clear_images(self):
        self.selected_images = []
        self.img_listbox.delete(0, "end")
        self.preview_label.config(image="", text="No image")
        self.status_var.set("Cleared images")

    def on_img_select(self, event=None):
        sel = self.img_listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        p = self.selected_images[idx]
        self.show_preview(p)

    def show_preview(self, path):
        try:
            im = Image.open(path)
            im.thumbnail((900, 700))
            self._thumb = ImageTk.PhotoImage(im)
            self.preview_label.config(image=self._thumb, text="")
        except Exception as e:
            self.preview_label.config(text="Preview error: {}".format(e))

    def reset_form(self):
        self.brand_var.set(BRAND_CODES[0])
        self.brand_name_var.set(BRAND_CODE_TO_NAME.get(self.brand_var.get(), ""))
        self.size_var.set(list(SIZES.keys())[0])
        for var in self.surface_vars.values():
            var.set(False)
        self.matt_polished_var.set("0")
        self.commercial_name_var.set("")
        self.faces_var.set("1")
        self.spcode_var.set("")
        self.batch_var.set("")
        self.country_var.set(DEFAULT_COUNTRY_PREFIX)
        self.company_var.set(DEFAULT_COMPANY_PREFIX)
        self.notes_var.set("")
        self.clear_images()
        self.status_var.set("Form reset")

    def validate_inputs(self):
        brand = self.brand_var.get()
        if brand not in BRAND_CODES:
            return None, "Invalid Brand"
        size_label = self.size_var.get()
        if size_label not in SIZES:
            return None, "Invalid Size"
        surface_codes = "".join(code for label, code in SURFACE_OPTIONS.items() if self.surface_vars[label].get())
        surface_labels = ", ".join(label for label in SURFACE_OPTIONS if self.surface_vars[label].get())
        matt_polished = self.matt_polished_var.get()
        if matt_polished not in ["0", "1"]:
            return None, "Must select Matt or Polished"
        sp = self.spcode_var.get().strip()
        if not sp.isdigit() or not (1 <= int(sp) <= 999):
            return None, "SPCode must be 001-999"
        faces = self.faces_var.get().strip()
        if not faces.isdigit() or int(faces) < 1:
            return None, "Faces must be positive integer"
        country = self.country_var.get().strip()
        company = self.company_var.get().strip()
        if not country.isdigit() or len(country) not in (2,3):
            return None, "CountryPrefix should be 2-3 digits"
        if not company.isdigit() or not (4 <= len(company) <= 9):
            return None, "CompanyPrefix should be 4-9 digits"
        batch = self.batch_var.get().strip()
        if batch and (not batch.startswith("LOT") or not (len(batch) == 8 and batch[3:5].isdigit() and batch[5:8].isdigit())):
            return None, "Batch must be empty or LOT[Year][Số thứ tự] (e.g., LOT2026001)"
        commercial_name = self.commercial_name_var.get().strip()
        if not commercial_name:
            return None, "Commercial Name is required"
        return {
            "brand_code": brand,
            "brand_name": BRAND_CODE_TO_NAME.get(brand, ""),
            "brand_id": BRAND_CODE_TO_ID.get(brand, "0"),
            "size_label": size_label,
            "size_code": SIZES[size_label],
            "surface_label": surface_labels,
            "surface_code": surface_codes,
            "matt_polished": matt_polished,
            "spcode": str(int(sp)).zfill(3),
            "commercial_name": commercial_name,
            "faces": int(faces),
            "batch": batch,
            "country": country,
            "company": company,
            "notes": self.notes_var.get().strip()
        }, None

    def save_product(self):
        data, err = self.validate_inputs()
        if err:
            messagebox.showerror("Validation", err)
            return
        df = load_df()
        sku = build_sku(data["brand_code"], data["size_code"], data["matt_polished"], data["spcode"])
        ean13 = build_ean13(data["country"], data["company"], data["brand_id"], data["spcode"])

        if sku in df["SKU"].astype(str).values:
            messagebox.showerror("Duplicate", "SKU exists: {}".format(sku))
            return

        sku_dir = os.path.join(IMAGES_ROOT, sku)
        os.makedirs(sku_dir, exist_ok=True)
        saved_paths = []
        face_count = data["faces"]
        idx = 1
        for i, src in enumerate(self.selected_images, start=1):
            if not os.path.isfile(src):
                continue
            face_idx = ((i - 1) % face_count) + 1
            base = os.path.join(sku_dir, "{}_face{:02d}_{:02d}".format(data["spcode"], face_idx, idx))
            dst = resize_and_save(src, base)
            if dst:
                saved_paths.append(dst)
            idx += 1

        bc, qr = generate_barcode_qr(ean13, sku_dir, data["spcode"])
        if bc:
            saved_paths.append(bc)
        if qr:
            saved_paths.append(qr)

        image_paths_str = ";".join(saved_paths)

        new_row = {
            "Timestamp": dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "BrandCode": data["brand_code"],
            "BrandName": data["brand_name"],
            "BrandID": data["brand_id"],
            "SizeLabel": data["size_label"],
            "SizeCode": data["size_code"],
            "SurfaceLabel": data["surface_label"],
            "SurfaceCode": data["surface_code"],
            "MattPolished": data["matt_polished"],
            "SPCode": data["spcode"],
            "SKU": sku,
            "CommercialName": data["commercial_name"],
            "Faces": str(data["faces"]),
            "Batch": data["batch"],
            "CountryPrefix": data["country"],
            "CompanyPrefix": data["company"],
            "EAN13": ean13,
            "ImagePaths": image_paths_str,
            "Notes": data["notes"]
        }

        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
        save_df(df)
        self.status_var.set(f"Saved SKU: {sku}")
        messagebox.showinfo("Saved", f"Saved SKU: {sku}\nEAN-13: {ean13}")
        if callable(self.viewer_refresh):
            try:
                self.viewer_refresh(select_sku=sku)
            except Exception:
                pass
        self.reset_form()

# ---------------- UI: Viewer Tab ----------------
class ViewerTab(ttk.Frame):
    def __init__(self, master):
        super().__init__(master)
        self._thumbs = []
        self._image_paths = []
        self._build_ui()
        self.refresh()

    def _build_ui(self):
        pad = 6
        top = ttk.Frame(self, padding=pad)
        top.pack(fill="x")
        ttk.Label(top, text="Search").grid(row=0, column=0, sticky="w")
        self.q_var = tk.StringVar(value="")
        ttk.Entry(top, textvariable=self.q_var, width=36).grid(row=0, column=1, padx=4)
        ttk.Button(top, text="Filter", command=self.refresh).grid(row=0, column=2, padx=4)

        ttk.Label(top, text="Brand").grid(row=0, column=3, padx=6)
        self.brand_filter = tk.StringVar(value="")
        ttk.Combobox(top, textvariable=self.brand_filter, values=[""] + BRAND_CODES, state="readonly", width=6).grid(row=0, column=4)

        ttk.Label(top, text="Surface").grid(row=0, column=5, padx=6)
        self.surface_filter = tk.StringVar(value="")
        ttk.Combobox(top, textvariable=self.surface_filter, values=[""] + list(SURFACE_OPTIONS.keys()), state="readonly", width=15).grid(row=0, column=6)

        ttk.Label(top, text="Matt/Polished").grid(row=0, column=7, padx=6)
        self.mp_filter = tk.StringVar(value="")
        ttk.Combobox(top, textvariable=self.mp_filter, values=[""] + list(MATT_POLISHED_OPTIONS.keys()), state="readonly", width=10).grid(row=0, column=8)

        ttk.Label(top, text="Size").grid(row=0, column=9, padx=6)
        self.size_filter = tk.StringVar(value="")
        ttk.Combobox(top, textvariable=self.size_filter, values=[""] + list(SIZES.keys()), state="readonly", width=10).grid(row=0, column=10)

        mid = ttk.PanedWindow(self, orient="horizontal")
        mid.pack(fill="both", expand=True, padx=pad, pady=pad)
        left = ttk.Frame(mid); mid.add(left, weight=3)
        right = ttk.Frame(mid); mid.add(right, weight=3)

        # SKU List with proper scrollbars
        left_frame = ttk.Frame(left)
        left_frame.pack(fill="both", expand=True)
        left_canvas = tk.Canvas(left_frame)
        scrollbar_y = ttk.Scrollbar(left_frame, orient="vertical", command=left_canvas.yview)
        scrollbar_x = ttk.Scrollbar(left_frame, orient="horizontal", command=left_canvas.xview)
        scrollable_left = ttk.Frame(left_canvas)
        scrollable_left.bind(
            "<Configure>",
            lambda e: left_canvas.configure(scrollregion=left_canvas.bbox("all"))
        )
        left_canvas.create_window((0, 0), window=scrollable_left, anchor="nw")
        left_canvas.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
        scrollbar_y.pack(side="right", fill="y")
        scrollbar_x.pack(side="bottom", fill="x")
        left_canvas.pack(side="left", fill="both", expand=True)

        cols = ["SKU", "CommercialName", "BrandCode", "SizeLabel", "SurfaceLabel", "MattPolished", "SPCode", "Faces", "Batch", "EAN13", "Timestamp"]
        self.tree = ttk.Treeview(scrollable_left, columns=cols, show="headings")
        self.tree.column("SKU", width=150)
        self.tree.column("CommercialName", width=300)
        self.tree.column("BrandCode", width=80)
        self.tree.column("SizeLabel", width=80)
        self.tree.column("SurfaceLabel", width=120)
        self.tree.column("MattPolished", width=100)
        self.tree.column("SPCode", width=80)
        self.tree.column("Faces", width=60)
        self.tree.column("Batch", width=100)
        self.tree.column("EAN13", width=120)
        self.tree.column("Timestamp", width=120)
        for c in cols:
            self.tree.heading(c, text=c)
        self.tree.pack(fill="both", expand=True)
        self.tree.bind("<<TreeviewSelect>>", self.on_select)
        self.tree.bind("<Double-1>", self.open_folder)

        # Preview with scrollbar
        right_frame = ttk.Frame(right)
        right_frame.pack(fill="both", expand=True, padx=4, pady=4)
        ttk.Label(right_frame, text="Preview").pack(anchor="w")
        canvas = tk.Canvas(right_frame)
        scrollbar = ttk.Scrollbar(right_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        self.preview_frame = scrollable_frame

        # Action buttons
        btns = ttk.Frame(right)
        btns.pack(fill="x", pady=4)
        ttk.Button(btns, text="Add Images", command=self.add_images).pack(side="left", padx=4)
        ttk.Button(btns, text="Add Notes", command=self.add_notes).pack(side="left", padx=4)
        ttk.Button(btns, text="Delete", command=self.delete_product).pack(side="left", padx=4)
        ttk.Button(btns, text="Open folder", command=self.open_folder).pack(side="left", padx=4)

    def _apply_filters(self, df):
        q = str(self.q_var.get()).strip().lower()
        brand = str(self.brand_filter.get()).strip()
        surf = str(self.surface_filter.get()).strip()
        mp = str(self.mp_filter.get()).strip()
        size = str(self.size_filter.get()).strip()
        if brand:
            df = df[df["BrandCode"].astype(str) == brand]
        if surf:
            df = df[df["SurfaceLabel"].str.contains(surf, case=False, na=False)]
        if mp:
            df = df[df["MattPolished"] == MATT_POLISHED_OPTIONS[mp]]
        if size:
            df = df[df["SizeLabel"].astype(str) == size]
        if q:
            df = df[df.apply(lambda r: q in " ".join([str(x).lower() for x in r.values]), axis=1)]
        return df

    def refresh(self, select_sku=None):
        df = load_df()
        df = self._apply_filters(df)
        for r in self.tree.get_children():
            self.tree.delete(r)
        for _, row in df.iterrows():
            vals = [
                row.get("SKU",""), row.get("CommercialName",""), row.get("BrandCode",""), row.get("SizeLabel",""),
                row.get("SurfaceLabel",""), row.get("MattPolished",""), row.get("SPCode",""), row.get("Faces",""),
                row.get("Batch",""), row.get("EAN13",""), row.get("Timestamp","")
            ]
            self.tree.insert("", "end", iid=row.get("SKU",""), values=vals)
        if select_sku and select_sku in self.tree.get_children():
            self.tree.selection_set(select_sku)
            self.tree.see(select_sku)

    def add_images(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Info", "Select a product to add images.")
            return
        sku = sel[0]
        df = load_df()
        row = df[df["SKU"].astype(str) == sku]
        if row.empty:
            messagebox.showerror("Error", "Product not found.")
            return
        row = row.iloc[0]
        files = filedialog.askopenfilenames(title="Select images", filetypes=(("Images", "*.jpg *.jpeg *.png *.bmp *.webp"), ("All files", "*.*")))
        if not files:
            return
        sku_dir = os.path.join(IMAGES_ROOT, sku)
        os.makedirs(sku_dir, exist_ok=True)
        saved_paths = str(row.get("ImagePaths","")).split(";") if row.get("ImagePaths","") else []
        face_count = int(row.get("Faces", 1))
        spcode = row.get("SPCode", "001")
        idx = len([p for p in saved_paths if "_face" in p]) + 1
        for i, src in enumerate(files, start=idx):
            if not os.path.isfile(src):
                continue
            face_idx = ((i - 1) % face_count) + 1
            base = os.path.join(sku_dir, "{}_face{:02d}_{:02d}".format(spcode, face_idx, i))
            dst = resize_and_save(src, base)
            if dst and dst not in saved_paths:
                saved_paths.append(dst)
        df.loc[df["SKU"] == sku, "ImagePaths"] = ";".join(saved_paths)
        save_df(df)
        self.refresh(select_sku=sku)
        messagebox.showinfo("Success", f"Added {len(files)} images to SKU {sku}")

    def add_notes(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Info", "Select a product to add notes.")
            return
        sku = sel[0]
        df = load_df()
        row = df[df["SKU"].astype(str) == sku]
        if row.empty:
            messagebox.showerror("Error", "Product not found.")
            return
        row = row.iloc[0]
        current_notes = row.get("Notes", "")
        dialog = AddNotesDialog(self, current_notes)
        self.wait_window(dialog)
        if dialog.result is not None:
            df.loc[df["SKU"] == sku, "Notes"] = dialog.result
            save_df(df)
            self.refresh(select_sku=sku)
            messagebox.showinfo("Success", f"Added note to SKU {sku}")

    def on_select(self, event=None):
        sel = self.tree.selection()
        if not sel:
            return
        sku = sel[0]
        df = load_df()
        row = df[df["SKU"].astype(str) == sku]
        if row.empty:
            return
        row = row.iloc[0]
        images = str(row.get("ImagePaths","")).split(";") if row.get("ImagePaths","") else []
        faces = int(row.get("Faces") or 1)
        for w in self.preview_frame.winfo_children():
            w.destroy()
        self._thumbs = []
        self._image_paths = []

        # Commercial Name (large font)
        commercial_name = row.get("CommercialName", "No Name")
        ttk.Label(self.preview_frame, text=commercial_name, font=("Helvetica", 20)).pack(anchor="center", pady=10)

        # Full SPCode (smaller font)
        full_spcode = build_full_spcode(
            row.get("BrandCode", ""),
            row.get("SizeCode", ""),
            row.get("MattPolished", ""),
            row.get("SPCode", "")
        )
        ttk.Label(self.preview_frame, text=f"SPCode: {full_spcode}", font=("Helvetica", 14)).pack(anchor="center", pady=5)

        # Surface Label
        surface_label = row.get("SurfaceLabel", "")
        ttk.Label(self.preview_frame, text=f"Surface: {surface_label or 'None'}", font=("Helvetica", 12)).pack(anchor="center", pady=5)

        # Matt/Polished
        mp_label = next((k for k, v in MATT_POLISHED_OPTIONS.items() if v == row.get("MattPolished", "")), "Unknown")
        ttk.Label(self.preview_frame, text=f"Matt/Polished: {mp_label}", font=("Helvetica", 12)).pack(anchor="center", pady=5)

        # Notes
        notes = row.get("Notes", "")
        ttk.Label(self.preview_frame, text=f"Notes: {notes or 'None'}", font=("Helvetica", 12), wraplength=400).pack(anchor="center", pady=5)

        # Images in vertical stack with right-click menu
        img_frame = ttk.Frame(self.preview_frame)
        img_frame.pack(fill="x", padx=10, pady=10)
        if not images:
            ttk.Label(img_frame, text="No images").pack(expand=True)
        else:
            show_count = min(faces, max(1, len(images)))
            for i in range(show_count):
                p = images[i] if i < len(images) else images[0]
                if os.path.isfile(p):
                    try:
                        img = Image.open(p)
                        img.thumbnail((300, 300))
                        photo = ImageTk.PhotoImage(img)
                        self._thumbs.append(photo)
                        self._image_paths.append(p)
                        lbl = ttk.Label(img_frame, image=photo)
                        lbl.pack(pady=8)
                        # Right-click menu
                        menu = tk.Menu(self, tearoff=0)
                        menu.add_command(label="Copy Image", command=lambda path=p: copy_image_to_clipboard(path, self.winfo_toplevel()))
                        menu.add_command(label="Delete Image", command=lambda path=p, s=sku: self.delete_image(path, s))
                        lbl.bind("<Button-3>", lambda e, m=menu: m.post(e.x_root, e.y_root))
                    except Exception:
                        ttk.Label(img_frame, text="Image Error").pack(pady=8)
                else:
                    ttk.Label(img_frame, text="No file").pack(pady=8)

        # Barcode and QR code with right-click menu
        code_frame = ttk.Frame(self.preview_frame)
        code_frame.pack(fill="x", pady=10)
        sku_dir = os.path.join(IMAGES_ROOT, sku)
        bc = os.path.join(sku_dir, f"{row.get('SPCode','')}_barcode.png")
        qr = os.path.join(sku_dir, f"{row.get('SPCode','')}_qrcode.png")
        if os.path.isfile(bc):
            try:
                img = Image.open(bc)
                img.thumbnail((300, 80))
                p = ImageTk.PhotoImage(img)
                self._thumbs.append(p)
                self._image_paths.append(bc)
                lbl = ttk.Label(code_frame, image=p)
                lbl.pack(side="left", padx=10)
                menu = tk.Menu(self, tearoff=0)
                menu.add_command(label="Copy Image", command=lambda path=bc: copy_image_to_clipboard(path, self.winfo_toplevel()))
                menu.add_command(label="Delete Image", command=lambda path=bc, s=sku: self.delete_image(path, s))
                lbl.bind("<Button-3>", lambda e, m=menu: m.post(e.x_root, e.y_root))
            except Exception:
                pass
        if os.path.isfile(qr):
            try:
                img = Image.open(qr)
                img.thumbnail((120, 120))
                q = ImageTk.PhotoImage(img)
                self._thumbs.append(q)
                self._image_paths.append(qr)
                lbl = ttk.Label(code_frame, image=q)
                lbl.pack(side="left", padx=10)
                menu = tk.Menu(self, tearoff=0)
                menu.add_command(label="Copy Image", command=lambda path=qr: copy_image_to_clipboard(path, self.winfo_toplevel()))
                menu.add_command(label="Delete Image", command=lambda path=qr, s=sku: self.delete_image(path, s))
                lbl.bind("<Button-3>", lambda e, m=menu: m.post(e.x_root, e.y_root))
            except Exception:
                pass

    def delete_image(self, img_path, sku):
        if not messagebox.askyesno("Confirm Delete", f"Are you sure you want to delete this image?"):
            return
        if delete_image_from_product(sku, img_path):
            messagebox.showinfo("Success", "Image deleted")
            self.refresh(select_sku=sku)
        else:
            messagebox.showerror("Error", "Failed to delete image")

    def delete_product(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Info", "Select a product to delete.")
            return
        sku = sel[0]
        if not messagebox.askyesno("Confirm Delete", f"Are you sure you want to delete SKU {sku}?"):
            return
        df = load_df()
        row = df[df["SKU"].astype(str) == sku]
        if row.empty:
            messagebox.showerror("Error", "Product not found.")
            return
        row = row.iloc[0]
        deleted_df = load_deleted_df()
        deleted_df = pd.concat([deleted_df, pd.DataFrame([row])], ignore_index=True)
        save_deleted_df(deleted_df)
        df = df[df["SKU"] != sku]
        save_df(df)
        self.refresh()
        messagebox.showinfo("Deleted", f"SKU {sku} deleted and moved to deleted_products.xlsx")

    def open_folder(self, event=None):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Info", "Select a product first.")
            return
        sku = sel[0]
        folder = os.path.join(IMAGES_ROOT, sku)
        if os.path.isdir(folder):
            open_folder(folder)
        else:
            messagebox.showinfo("No images", "No image folder for this SKU yet.")

# ---------------- App ----------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("TC Premium Tile SKU Manager")
        self.geometry("1400x800")
        ensure_storage()
        self.notebook = ttk.Notebook(self)
        self.viewer = ViewerTab(self.notebook)
        self.entry = EntryTab(self.notebook, viewer_refresh_callback=self.viewer.refresh)
        self.notebook.add(self.entry, text="Entry")
        self.notebook.add(self.viewer, text="Viewer")
        self.notebook.pack(fill="both", expand=True)

if __name__ == "__main__":
    app = App()
    app.mainloop()