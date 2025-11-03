# aplikasi_kasir.py
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import datetime
import csv
import os
import json
from tkinter import scrolledtext
import winsound
import threading
import tempfile
import win32print
import win32api
import zipfile
import pandas as pd

# ================== UTILITY: FULLSCREEN & EXIT HELPERS ==================
def toggle_fullscreen(win):
"""Toggle fullscreen state for a given window (F12)."""
try:
is_full = getattr(win, "_is_fullscreen", False)
win.attributes("-fullscreen", not is_full)
win._is_fullscreen = not is_full
except Exception:
# Some platforms or window types might not support attributes, ignore silently
pass

def exit_fullscreen(win):
"""Exit fullscreen for given window (Escape)."""
try:
win.attributes("-fullscreen", False)
win._is_fullscreen = False
except Exception:
pass

def confirm_exit_app(root):
"""Ask user to confirm exit; if yes, quit entire application."""
try:
if messagebox.askyesno("Konfirmasi", "Keluar aplikasi?"):
# Try to close root mainloop and all windows
try:
root.destroy()
except Exception:
try:
root.quit()
except Exception:
pass
except Exception:
# In case messagebox fails (rare), fallback to destroy
try:
root.destroy()
except Exception:
pass

def bind_fullscreen_keys(win, escape_exits_app=False):
"""
Bind F12 to toggle fullscreen for 'win'.
If escape_exits_app=True bind Escape to asking exit for this window (used on main root windows).
Otherwise bind Escape to exit fullscreen for that window (used on dialogs).
"""
try:
# Bind F12 for toggling fullscreen - use add='+' so we don't clobber existing bindings
win.bind("
if escape_exits_app:
# Escape should ask to exit the whole app (confirm)
win.bind("
else:
# Escape should at least exit fullscreen for that specific window
win.bind("
except Exception:
pass

# ================== KONFIGURASI AWAL ==================
ADMIN_USER = "admin"
ADMIN_PASS = "1234"
TOKO_NAMA = "TOKO SIMPANG MOTOR"
TOKO_ALAMAT = "Jl. Poros, Sumber Jaya"

STOK_FILE = "stok.csv"
TRANSAKSI_FILE = "transaksi.csv"
STRUK_DIR = "struk"
SETTINGS_FILE = "pengaturan.json"

# Sistem Tema
THEMES = {
"dark": {
"name": "Dark Mode",
"bg": "#1e1e1e", "fg": "#ffffff", "card_bg": "#2d2d2d",
"accent": "#007acc", "highlight": "#ff6b35", "success": "#28a745",
"error": "#dc3545", "entry_bg": "#404040", "tree_bg": "#2d2d2d",
"tree_fg": "#ffffff", "tree_heading_bg": "#007acc"
},
"light": {
"name": "Light Mode",
"bg": "#f8f9fa", "fg": "#212529", "card_bg": "#ffffff",
"accent": "#007acc", "highlight": "#ff6b35", "success": "#28a745",
"error": "#dc3545", "entry_bg": "#ffffff", "tree_bg": "#ffffff",
"tree_fg": "#212529", "tree_heading_bg": "#007acc"
},
"blue": {
"name": "Blue Ocean",
"bg": "#1a237e", "fg": "#e3f2fd", "card_bg": "#283593",
"accent": "#2979ff", "highlight": "#00e5ff", "success": "#00e676",
"error": "#ff1744", "entry_bg": "#3949ab", "tree_bg": "#283593",
"tree_fg": "#e3f2fd", "tree_heading_bg": "#2979ff"
},
"green": {
"name": "Green Forest",
"bg": "#1b5e20", "fg": "#e8f5e9", "card_bg": "#2e7d32",
"accent": "#00c853", "highlight": "#ffd600", "success": "#00e676",
"error": "#ff3d00", "entry_bg": "#43a047", "tree_bg": "#2e7d32",
"tree_fg": "#e8f5e9", "tree_heading_bg": "#00c853"
}
}

# Buat direktori jika belum ada
if not os.path.exists(STRUK_DIR):
os.makedirs(STRUK_DIR)

# ================== MANAGER PRINTER ==================
class PrinterManager:
def __init__(self):
self.printers = self.get_available_printers()
self.default_printer = self.get_default_printer()
self.receipt_printer = self.default_printer
self.report_printer = self.default_printer
self.label_printer = self.default_printer

def get_available_printers(self):
"""Dapatkan daftar semua printer yang tersedia"""
try:
printers = []
printer_info = win32print.EnumPrinters(
win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
)
for printer in printer_info:
printers.append(printer[2])
return printers
except:
try:
return [win32print.GetDefaultPrinter()]
except:
return []

def get_default_printer(self):
"""Dapatkan printer default"""
try:
return win32print.GetDefaultPrinter()
except:
return ""

def set_receipt_printer(self, printer_name):
"""Set printer untuk struk"""
if printer_name in self.printers:
self.receipt_printer = printer_name
return True
return False

def set_report_printer(self, printer_name):
"""Set printer untuk laporan"""
if printer_name in self.printers:
self.report_printer = printer_name
return True
return False

def set_label_printer(self, printer_name):
"""Set printer untuk label barcode"""
if printer_name in self.printers:
self.label_printer = printer_name
return True
return False

def test_print(self, printer_name, text="Test Printer\n"):
"""Test print ke printer tertentu"""
try:
if printer_name not in self.printers:
return False, "Printer tidak ditemukan"

raw_print(text, printer_name)
return True, "Test print berhasil"
except Exception as e:
return False, f"Error: {str(e)}"

# Initialize printer manager
printer_mgr = PrinterManager()

# ================== FUNGSI PRINT ==================
def raw_print(text, printer_name=None):
"""Print raw text ke printer tertentu"""
if printer_name is None:
printer_name = printer_mgr.receipt_printer

phandle = None
try:
data = text.encode("cp850", errors="replace")
phandle = win32print.OpenPrinter(printer_name)
try:
docinfo = ("Struk", None, "RAW")
job = win32print.StartDocPrinter(phandle, 1, docinfo)
try:
win32print.StartPagePrinter(phandle)
win32print.WritePrinter(phandle, data)
win32print.EndPagePrinter(phandle)
finally:
win32print.EndDocPrinter(phandle)
finally:
win32print.ClosePrinter(phandle)
return True
except Exception as e:
raise

def shell_print_tempfile(tempfile_path, printer_name=None):
"""Print menggunakan shell execute ke printer tertentu"""
if printer_name is None:
printer_name = printer_mgr.receipt_printer

try:
win32api.ShellExecute(0, "print", tempfile_path, f'/d:"{printer_name}"', ".", 0)
except Exception as e:
raise

# ================== MANAGER TEMA - DIPERBAIKI ==================
class ThemeManager:
def __init__(self):
self.current_theme = "dark"
self.themes = THEMES
self.load_theme()

def get_theme(self):
return self.themes[self.current_theme]

def set_theme(self, theme_name):
if theme_name in self.themes:
self.current_theme = theme_name
success = self.save_theme()
if success:
print(f"Theme berhasil diubah ke: {theme_name}")
return success
return False

def get_theme_names(self):
return [theme["name"] for theme in self.themes.values()]

def get_theme_by_name(self, theme_name):
for key, theme in self.themes.items():
if theme["name"] == theme_name:
return key
return "dark"

def load_theme(self):
"""Load tema dari file pengaturan - DIPERBAIKI"""
if os.path.exists(SETTINGS_FILE):
try:
with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
settings = json.load(f)
if "theme" in settings:
# Langsung set current_theme tanpa panggil set_theme()
self.current_theme = settings["theme"]
print(f"Theme loaded: {self.current_theme}")
except Exception as e:
print(f"Error loading theme: {e}")

def save_theme(self):
"""Simpan tema ke file pengaturan - DIPERBAIKI"""
settings = {}
if os.path.exists(SETTINGS_FILE):
try:
with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
settings = json.load(f)
except Exception as e:
print(f"Error reading settings: {e}")

settings["theme"] = self.current_theme

try:
with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
json.dump(settings, f, indent=4, ensure_ascii=False)
print(f"Theme saved successfully: {self.current_theme}")
return True
except Exception as e:
print(f"Error saving theme: {e}")
return False

# Initialize theme manager
theme_mgr = ThemeManager()

# ================== MANAGER SUARA ==================
class SoundManager:
def __init__(self):
self.enabled = True
self.sounds = {
"success": 1000, "error": 400, "add_item": 800,
"delete": 600, "cash": 1200, "theme_change": 700
}

def play_sound(self, sound_type):
if not self.enabled:
return
def play():
try:
if sound_type in self.sounds:
freq = self.sounds[sound_type]
winsound.Beep(freq, 200)
except:
pass
threading.Thread(target=play, daemon=True).start()

sound_mgr = SoundManager()

# ================== FUNGSI UTILITAS ==================
def format_rupiah(n):
try:
n = int(n)
except:
return str(n)
return f"Rp {n:,}".replace(",", ".")

def parse_rupiah(rupiah_str):
if not rupiah_str:
return 0
clean_str = ''.join(filter(str.isdigit, str(rupiah_str)))
return int(clean_str) if clean_str else 0

def load_stok():
stok = {}
if os.path.exists(STOK_FILE):
try:
with open(STOK_FILE, "r", encoding="utf-8") as f:
reader = csv.reader(f)
for row in reader:
if not row or len(row) < 5:
continue
try:
kode, nama = row[0].strip(), row[1].strip()
harga_beli, harga_jual, qty = int(row[2]), int(row[3]), int(row[4])
stok[kode] = {
"nama": nama,
"harga_beli": harga_beli,
"harga_jual": harga_jual,
"qty": qty
}
except:
continue
except:
pass
return stok

def save_stok(stok):
try:
with open(STOK_FILE, "w", newline="", encoding="utf-8") as f:
writer = csv.writer(f)
for kode, d in stok.items():
writer.writerow([kode, d["nama"], d["harga_beli"], d["harga_jual"], d["qty"]])
return True, "Sukses"
except Exception as e:
print(f"Error saving stok: {e}")
return False, str(e)

def append_transaksi(items):
try:
timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
with open(TRANSAKSI_FILE, "a", newline="", encoding="utf-8") as f:
writer = csv.writer(f)
for it in items:
writer.writerow([
timestamp, it["kode"], it["nama"],
it["harga_beli"], it["harga_jual"],
it["qty"], it["subtotal"], it["laba"]
])
except:
pass

def load_settings():
"""Load pengaturan dari file"""
settings = {}
if os.path.exists(SETTINGS_FILE):
try:
with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
settings = json.load(f)
except Exception as e:
print(f"Error loading settings: {e}")
return settings

def save_settings(settings):
"""Simpan pengaturan ke file - DIPERBAIKI"""
try:
with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
json.dump(settings, f, indent=4, ensure_ascii=False)
print("Settings saved successfully")
return True
except Exception as e:
print(f"Error saving settings: {e}")
return False

def get_dashboard_stats():
"""Hitung total penjualan dan jumlah transaksi untuk hari ini."""
today_str = datetime.datetime.now().strftime("%Y-%m-%d")
total_penjualan = 0
transaksi_ids = set()

if not os.path.exists(TRANSAKSI_FILE):
return 0, 0

try:
with open(TRANSAKSI_FILE, "r", encoding="utf-8") as f:
reader = csv.reader(f)
for row in reader:
if row and row[0].startswith(today_str):
try:
timestamp = row[0]
subtotal = int(row[6])
total_penjualan += subtotal
transaksi_ids.add(timestamp)
except (ValueError, IndexError):
continue
return total_penjualan, len(transaksi_ids)
except Exception as e:
print(f"Error reading transaction file for dashboard: {e}")
return 0, 0

def get_stok_kritis(threshold=5):
"""Dapatkan daftar item dengan stok di bawah ambang batas."""
stok = load_stok()
kritis = []
for kode, data in stok.items():
if data['qty'] <= threshold:
kritis.append(f"- {data['nama']} ({data['qty']})")

if not kritis:
return "-"

return "\n".join(kritis[:3]) # Tampilkan hingga 3 item

# ================== FUNGSI PRINT STRUK SATUAN ==================
def print_struk_satuan_laporan(parent, transaksi_data, printer_name=None):
"""Print struk untuk satu transaksi dari laporan"""
theme = theme_mgr.get_theme()

# Dialog preview dan print
win = tk.Toplevel(parent)
win.title("Preview Struk Transaksi")
win.geometry("500x600")
win.configure(bg=theme["bg"])
win.transient(parent)
bind_fullscreen_keys(win) # bind F12 (Escape will exit fullscreen for this dialog)

# Style
style = ttk.Style(win)
style.theme_use("clam")
style.configure(".", background=theme["bg"], foreground=theme["fg"])
style.configure("TLabel", background=theme["bg"], foreground=theme["fg"], font=("Arial", 10))
style.configure("TButton", background=theme["card_bg"], foreground=theme["fg"], padding=6)
style.configure("TFrame", background=theme["bg"])

frm = ttk.Frame(win, padding=20)
frm.pack(expand=True, fill="both")

ttk.Label(frm, text="PREVIEW STRUK TRANSAKSI",
font=("Arial", 14, "bold")).pack(pady=(0, 20))

# Info transaksi
info_frame = ttk.Frame(frm)
info_frame.pack(fill="x", pady=10)

ttk.Label(info_frame, text=f"Waktu: {transaksi_data['timestamp']}",
font=("Arial", 10, "bold")).pack(anchor="w")
ttk.Label(info_frame, text=f"Barang: {transaksi_data['nama']}",
font=("Arial", 10)).pack(anchor="w")
ttk.Label(info_frame, text=f"Kode: {transaksi_data['kode']}",
font=("Arial", 10)).pack(anchor="w")

# Preview struk
preview_frame = ttk.Frame(frm, relief='solid', height=300)
preview_frame.pack(fill="both", expand=True, pady=10)
preview_frame.pack_propagate(False)

# Scrollable text untuk preview
preview_text = scrolledtext.ScrolledText(preview_frame, wrap=tk.WORD,
bg=theme["entry_bg"], fg=theme["fg"],
font=("Courier New", 9), height=15)
preview_text.pack(fill="both", expand=True, padx=5, pady=5)

def build_struk_satuan_laporan(transaksi):
"""Bangun teks struk untuk satu transaksi dari laporan"""
w = 32 # Lebar struk thermal

settings = load_settings()
toko_nama = settings.get("toko_nama", TOKO_NAMA)
toko_alamat = settings.get("toko_alamat", TOKO_ALAMAT)

header = f"{toko_nama:^{w}}\n{toko_alamat:^{w}}\n"
header += "=" * w + "\n"
header += "STRUK TRANSAKSI SATUAN\n"
header += "-" * w + "\n"
header += f"Waktu: {transaksi['timestamp']}\n"
header += "-" * w + "\n"

body = ""
name = transaksi["nama"][:w-16]
body += f"{name}\n"
body += f"Kode : {transaksi['kode']}\n"
body += f"Qty : {transaksi['qty']:>10}\n"
body += f"Harga: {format_rupiah(transaksi['harga_jual']):>12}\n"
body += f"Subtl: {format_rupiah(transaksi['subtotal']):>12}\n"
body += f"Laba : {format_rupiah(transaksi['laba']):>12}\n"

footer = "-" * w + "\n"
footer += "Terima kasih!\n"
footer += "=" * w + "\n\n\n"

receipt = header + body + footer
return receipt

def update_preview():
"""Update preview text"""
struk_text = build_struk_satuan_laporan(transaksi_data)
preview_text.delete(1.0, tk.END)
preview_text.insert(1.0, struk_text)

def print_struk():
"""Print struk ke printer"""
struk_text = build_struk_satuan_laporan(transaksi_data)

try:
current_printer = printer_name or printer_mgr.receipt_printer

raw_print(struk_text, current_printer)
messagebox.showinfo("Sukses", "Struk berhasil dikirim ke printer!")
sound_mgr.play_sound("success")
except Exception as e:
messagebox.showerror("Error", f"Gagal print: {str(e)}")
sound_mgr.play_sound("error")

def print_struk_file():
"""Simpan struk ke file"""
struk_text = build_struk_satuan_laporan(transaksi_data)

timestamp = transaksi_data['timestamp'].replace(":", "").replace(" ", "_")
filename = f"struk_satuan_{transaksi_data['kode']}_{timestamp}.txt"
filepath = os.path.join(STRUK_DIR, filename)

try:
with open(filepath, "w", encoding="utf-8") as f:
f.write(struk_text)
messagebox.showinfo("Sukses", f"Struk disimpan di: {filepath}")
sound_mgr.play_sound("success")
except Exception as e:
messagebox.showerror("Error", f"Gagal simpan file: {str(e)}")
sound_mgr.play_sound("error")

# Button frame
btn_frame = ttk.Frame(frm)
btn_frame.pack(pady=10)

ttk.Button(btn_frame, text="🖨️ Print Struk",
command=print_struk, width=15).pack(side="left", padx=5)

ttk.Button(btn_frame, text="💾 Simpan ke File",
command=print_struk_file, width=15).pack(side="left", padx=5)

ttk.Button(btn_frame, text="👁️ Refresh Preview",
command=update_preview, width=15).pack(side="left", padx=5)

ttk.Button(btn_frame, text="❌ Tutup",
command=win.destroy, width=10).pack(side="left", padx=5)

# Initial preview
update_preview()

# ================== FUNGSI PRINT LAPORAN KESELURUHAN ==================
def print_laporan_keseluruhan(parent, transaksi_list, total_penjualan, total_laba, total_item, printer_name=None):
"""Print laporan transaksi keseluruhan"""
theme = theme_mgr.get_theme()

# Dialog preview dan print
win = tk.Toplevel(parent)
win.title("Preview Laporan Keseluruhan")
win.geometry("600x700")
win.configure(bg=theme["bg"])
win.transient(parent)
bind_fullscreen_keys(win)

# Style
style = ttk.Style(win)
style.theme_use("clam")
style.configure(".", background=theme["bg"], foreground=theme["fg"])
style.configure("TLabel", background=theme["bg"], foreground=theme["fg"], font=("Arial", 10))
style.configure("TButton", background=theme["card_bg"], foreground=theme["fg"], padding=6)
style.configure("TFrame", background=theme["bg"])

frm = ttk.Frame(win, padding=20)
frm.pack(expand=True, fill="both")

ttk.Label(frm, text="PREVIEW LAPORAN KESELURUHAN",
font=("Arial", 14, "bold")).pack(pady=(0, 20))

# Info summary
info_frame = ttk.Frame(frm)
info_frame.pack(fill="x", pady=10)

ttk.Label(info_frame, text=f"Periode: {len(transaksi_list)} transaksi",
font=("Arial", 10, "bold")).pack(anchor="w")
ttk.Label(info_frame, text=f"Total Penjualan: {format_rupiah(total_penjualan)}",
font=("Arial", 10)).pack(anchor="w")
ttk.Label(info_frame, text=f"Total Laba: {format_rupiah(total_laba)}",
font=("Arial", 10)).pack(anchor="w")
ttk.Label(info_frame, text=f"Total Item: {total_item}",
font=("Arial", 10)).pack(anchor="w")

# Preview laporan
preview_frame = ttk.Frame(frm, relief='solid', height=400)
preview_frame.pack(fill="both", expand=True, pady=10)
preview_frame.pack_propagate(False)

# Scrollable text untuk preview
preview_text = scrolledtext.ScrolledText(preview_frame, wrap=tk.WORD,
bg=theme["entry_bg"], fg=theme["fg"],
font=("Courier New", 8), height=20)
preview_text.pack(fill="both", expand=True, padx=5, pady=5)

def build_laporan_keseluruhan(transaksi_list, total_penjualan, total_laba, total_item):
"""Bangun teks laporan keseluruhan"""
w = 80 # Lebar laporan

settings = load_settings()
toko_nama = settings.get("toko_nama", TOKO_NAMA)
toko_alamat = settings.get("toko_alamat", TOKO_ALAMAT)

header = f"{toko_nama:^{w}}\n{toko_alamat:^{w}}\n"
header += "=" * w + "\n"
header += "LAPORAN TRANSAKSI KESELURUHAN\n"
header += f"Dicetak: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
header += f"Periode: {len(transaksi_list)} transaksi\n"
header += "=" * w + "\n"

# Header tabel
body = f"{'NO':<3} {'WAKTU':<16} {'KODE':<8} {'NAMA':<20} {'QTY':>5} {'HARGA':>10} {'SUBTOTAL':>12} {'LABA':>10}\n"
body += "-" * w + "\n"

# Data transaksi
for i, transaksi in enumerate(transaksi_list, 1):
waktu = transaksi['timestamp'][11:16] # Ambil hanya jam:menit
nama = transaksi['nama'][:19] # Potong nama jika terlalu panjang
body += f"{i:<3} {waktu:<16} {transaksi['kode']:<8} {nama:<20} {transaksi['qty']:>5} {format_rupiah(transaksi['harga_jual']):>10} {format_rupiah(transaksi['subtotal']):>12} {format_rupiah(transaksi['laba']):>10}\n"

# Footer
body += "=" * w + "\n"
body += f"TOTAL PENJUALAN: {format_rupiah(total_penjualan):>50}\n"
body += f"TOTAL LABA: {format_rupiah(total_laba):>55}\n"
body += f"TOTAL ITEM TERJUAL: {total_item:>48}\n"
body += "=" * w + "\n\n"

return header + body

def update_preview():
"""Update preview text"""
laporan_text = build_laporan_keseluruhan(transaksi_list, total_penjualan, total_laba, total_item)
preview_text.delete(1.0, tk.END)
preview_text.insert(1.0, laporan_text)

def print_laporan():
"""Print laporan ke printer"""
laporan_text = build_laporan_keseluruhan(transaksi_list, total_penjualan, total_laba, total_item)

try:
current_printer = printer_mgr.report_printer

if not current_printer:
current_printer = printer_mgr.receipt_printer

if not current_printer:
current_printer = win32print.GetDefaultPrinter()

raw_print(laporan_text, current_printer)
messagebox.showinfo("Sukses", f"Laporan berhasil dikirim ke printer!\nPrinter: {current_printer}")
sound_mgr.play_sound("success")
except Exception as e:
messagebox.showerror("Error", f"Gagal print: {str(e)}")
sound_mgr.play_sound("error")

def export_laporan_excel():
"""Export laporan ke Excel"""
try:
# Buat DataFrame dari data transaksi
data = []
for transaksi in transaksi_list:
data.append([
transaksi['timestamp'],
transaksi['kode'],
transaksi['nama'],
transaksi['harga_jual'],
transaksi['qty'],
transaksi['subtotal'],
transaksi['laba']
])

df = pd.DataFrame(data, columns=['Waktu', 'Kode', 'Nama', 'Harga Jual', 'Qty', 'Subtotal', 'Laba'])

# Tambahkan summary
summary_data = {
'Metric': ['Total Transaksi', 'Total Penjualan', 'Total Laba', 'Total Item Terjual', 'Tanggal Export'],
'Value': [len(transaksi_list), total_penjualan, total_laba, total_item, datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")]
}
df_summary = pd.DataFrame(summary_data)

# Export ke Excel
timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
excel_file = f"laporan_transaksi_{timestamp}.xlsx"

with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
df.to_excel(writer, sheet_name="Data Transaksi", index=False)
df_summary.to_excel(writer, sheet_name="Summary", index=False)

messagebox.showinfo("Sukses", f"Laporan berhasil diexport ke: {excel_file}")
sound_mgr.play_sound("success")

except ImportError:
messagebox.showerror("Error", "Library pandas tidak tersedia.\nInstall dengan: pip install pandas openpyxl")
sound_mgr.play_sound("error")
except Exception as e:
messagebox.showerror("Error", f"Gagal export: {str(e)}")
sound_mgr.play_sound("error")

def save_laporan_file():
"""Simpan laporan ke file teks"""
laporan_text = build_laporan_keseluruhan(transaksi_list, total_penjualan, total_laba, total_item)

timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
filename = f"laporan_transaksi_{timestamp}.txt"
filepath = os.path.join(STRUK_DIR, filename)

try:
with open(filepath, "w", encoding="utf-8") as f:
f.write(laporan_text)
messagebox.showinfo("Sukses", f"Laporan disimpan di: {filepath}")
sound_mgr.play_sound("success")
except Exception as e:
messagebox.showerror("Error", f"Gagal simpan file: {str(e)}")
sound_mgr.play_sound("error")

# Button frame
btn_frame = ttk.Frame(frm)
btn_frame.pack(pady=10)

ttk.Button(btn_frame, text="🖨️ Print Laporan",
command=print_laporan, width=15).pack(side="left", padx=5)

ttk.Button(btn_frame, text="📊 Export Excel",
command=export_laporan_excel, width=15).pack(side="left", padx=5)

ttk.Button(btn_frame, text="💾 Simpan ke File",
command=save_laporan_file, width=15).pack(side="left", padx=5)

ttk.Button(btn_frame, text="👁️ Refresh Preview",
command=update_preview, width=15).pack(side="left", padx=5)

ttk.Button(btn_frame, text="❌ Tutup",
command=win.destroy, width=10).pack(side="left", padx=5)

# Initial preview
update_preview()

# ================== GENERATOR BARCODE ==================
def barcode_generator_window(parent):
theme = theme_mgr.get_theme()

win = tk.Toplevel(parent)
win.title("Generator Barcode Sederhana")
win.geometry("500x550")
win.configure(bg=theme["bg"])
win.transient(parent)
win.resizable(True, True)
bind_fullscreen_keys(win)

# Style
style = ttk.Style(win)
style.theme_use("clam")
style.configure(".", background=theme["bg"], foreground=theme["fg"])
style.configure("TLabel", background=theme["bg"], foreground=theme["fg"], font=("Arial", 10))
style.configure("TButton", background=theme["card_bg"], foreground=theme["fg"], padding=6)
style.configure("TEntry", fieldbackground=theme["entry_bg"], foreground=theme["fg"])

frm = ttk.Frame(win, padding=20)
frm.pack(expand=True, fill="both")
frm.rowconfigure(6, weight=1) # Beri bobot pada baris pratinjau (sekarang baris 6)
frm.columnconfigure(0, weight=1)

ttk.Label(frm, text="GENERATOR BARCODE SEDERHANA",
font=("Arial", 14, "bold")).grid(row=0, column=0, pady=(0, 20))

# Input teks untuk barcode
ttk.Label(frm, text="Teks untuk Barcode:", font=("Arial", 10, "bold")).grid(row=1, column=0, sticky="w", pady=(10,5))
entry_teks = ttk.Entry(frm, width=30, font=("Arial", 10))
entry_teks.grid(row=2, column=0, sticky="ew", pady=5)
entry_teks.insert(0, "TEST123")

# Info
info_label = ttk.Label(frm, text="Masukkan teks/kode apapun (angka/huruf)",
foreground=theme["fg"], font=("Arial", 9))
info_label.grid(row=3, column=0, sticky="w", pady=(0,10))

# Pengaturan Ukuran
size_frame = ttk.Frame(frm)
size_frame.grid(row=4, column=0, sticky="ew", pady=(10, 5))

ttk.Label(size_frame, text="Tinggi (mm):", font=('Arial', 9)).grid(row=0, column=0, sticky="w", padx=(0, 5))
entry_height = ttk.Entry(size_frame, width=8)
entry_height.insert(0, "15.0")
entry_height.grid(row=0, column=1, sticky="w")

ttk.Label(size_frame, text="Margin (mm):", font=('Arial', 9)).grid(row=0, column=2, sticky="w", padx=(15, 5))
entry_margin = ttk.Entry(size_frame, width=8)
entry_margin.insert(0, "2.0")
entry_margin.grid(row=0, column=3, sticky="w")

ttk.Label(size_frame, text="Jumlah Cetak:", font=('Arial', 9)).grid(row=0, column=4, sticky="w", padx=(15, 5))
entry_jumlah = ttk.Entry(size_frame, width=8)
entry_jumlah.insert(0, "1")
entry_jumlah.grid(row=0, column=5, sticky="w")

# Opsi Dua Kolom
dual_column_var = tk.BooleanVar()
check_dual = ttk.Checkbutton(frm, text="Cetak Dua Kolom (untuk printer lebar)", variable=dual_column_var)
check_dual.grid(row=5, column=0, sticky="w", pady=5)

# Frame untuk preview
preview_frame = ttk.Frame(frm, relief='solid', height=200)
preview_frame.grid(row=6, column=0, sticky="nsew", pady=10)
preview_frame.pack_propagate(False)

# Status label
status_label = ttk.Label(frm, text="Masukkan teks dan generate barcode",
foreground=theme["fg"], font=("Arial", 9))
status_label.grid(row=7, column=0, pady=5)

def generate_barcode():
teks = entry_teks.get().strip()
if not teks:
messagebox.showerror("Error", "Masukkan teks terlebih dahulu!")
sound_mgr.play_sound("error")
return

try:
# Clear preview lama
for widget in preview_frame.winfo_children():
widget.destroy()

# Import libraries
import barcode
from barcode.writer import ImageWriter
from PIL import Image, ImageTk

# Generate barcode
barcode_class = barcode.get_barcode_class('code128')

# Setup writer options
writer = ImageWriter()
writer.set_options({
'module_height': float(entry_height.get()),
'quiet_zone': float(entry_margin.get()),
'font_size': 0,
'write_text': False
})

# Generate barcode
barcode_obj = barcode_class(teks, writer=writer)

# Buat folder jika belum ada
barcode_dir = 'barcode_produk'
if not os.path.exists(barcode_dir):
os.makedirs(barcode_dir)

# Save barcode
filename = f"barcode_{teks}"
filepath = os.path.join(barcode_dir, filename)
full_path = barcode_obj.save(filepath)

# Tampilkan preview
img = Image.open(full_path)

# Resize untuk preview (max width 400px)
width, height = img.size
if width > 400:
ratio = 400 / width
new_width = 400
new_height = int(height * ratio)
img_resized = img.resize((new_width, new_height), Image.Resampling.LANCZOS)
else:
img_resized = img

photo = ImageTk.PhotoImage(img_resized)

label_preview = ttk.Label(preview_frame, image=photo)
label_preview.image = photo
label_preview.pack(pady=10)

# Tampilkan teks di bawah barcode
label_teks = ttk.Label(preview_frame, text=teks, font=("Arial", 10, "bold"))
label_teks.pack()

status_label.config(text=f"✅ Barcode berhasil dibuat: {filename}.png")

# Simpan path untuk print nanti
win.barcode_path = full_path
win.barcode_teks = teks

sound_mgr.play_sound("success")

except ImportError:
error_msg = "Library barcode tidak tersedia. Install dengan: pip install python-barcode pillow"
status_label.config(text=f"❌ {error_msg}")

for widget in preview_frame.winfo_children():
widget.destroy()

error_label = ttk.Label(preview_frame, text=f"Error:\n{error_msg}",
foreground=theme["error"], wraplength=400, justify="left")
error_label.pack(expand=True)

messagebox.showerror("Error", error_msg)
sound_mgr.play_sound("error")
except Exception as e:
error_msg = f"Gagal generate barcode: {str(e)}"
status_label.config(text=f"❌ {error_msg}")

for widget in preview_frame.winfo_children():
widget.destroy()

error_label = ttk.Label(preview_frame, text=f"Error:\n{str(e)}",
foreground=theme["error"], wraplength=400, justify="left")
error_label.pack(expand=True)

messagebox.showerror("Error", error_msg)
sound_mgr.play_sound("error")

def print_barcode():
if not hasattr(win, 'barcode_path'):
messagebox.showerror("Error", "Generate barcode terlebih dahulu!")
sound_mgr.play_sound("error")
return

try:
jumlah_cetak = int(entry_jumlah.get())
if jumlah_cetak < 1:
messagebox.showerror("Error", "Jumlah cetak harus minimal 1!")
return
except ValueError:
messagebox.showerror("Error", "Jumlah cetak harus berupa angka!")
return

try:
from PIL import Image, ImageWin
import win32ui

# Buka gambar barcode yang sudah digenerate
original_image = Image.open(win.barcode_path)
image_to_print = original_image

if dual_column_var.get():
# Buat gambar baru dengan lebar ganda untuk menampung dua barcode
w, h = original_image.size
combined_image = Image.new('RGB', (w * 2 + 30, h + 30), 'white')
combined_image.paste(original_image, (10, 15))
combined_image.paste(original_image, (w + 20, 15))
image_to_print = combined_image

# Dapatkan printer label dari printer manager
printer_name = printer_mgr.label_printer
hprinter = win32print.OpenPrinter(printer_name)

try:
# Create device context untuk printer
hdc = win32ui.CreateDC()
hdc.CreatePrinterDC(printer_name)

# Penskalaan Cerdas
printable_width = hdc.GetDeviceCaps(110) # Lebar fisik dalam piksel
img_width, img_height = image_to_print.size

if img_width > printable_width:
# Skalakan gambar agar pas dengan lebar kertas
scale_ratio = printable_width / img_width
new_width = printable_width
new_height = int(img_height * scale_ratio)
image_to_print = image_to_print.resize((new_width, new_height), Image.Resampling.LANCZOS)
img_width, img_height = image_to_print.size

# Pusatkan gambar hasil skala
x = (printable_width - img_width) // 2
y = 50 # Margin atas

# Kirim semua salinan dalam satu pekerjaan cetak
hdc.StartDoc("Barcode Print Job")
for i in range(jumlah_cetak):
hdc.StartPage()

# Cetak gambar barcode
dib = ImageWin.Dib(image_to_print)
dib.draw(hdc.GetHandleOutput(), (x, y, x + img_width, y + img_height))

# Tambahkan teks di bawah barcode (hanya jika tidak dua kolom)
if not dual_column_var.get():
hdc.SetMapMode(4) # MM_TEXT
hdc.TextOut(x, y + img_height + 20, win.barcode_teks)

hdc.EndPage()
hdc.EndDoc()

status_label.config(text=f"✅ {jumlah_cetak} barcode '{win.barcode_teks}' berhasil di-print!")
messagebox.showinfo("Sukses", f"{jumlah_cetak} barcode berhasil dikirim ke printer!")
sound_mgr.play_sound("success")

except Exception as e:
raise Exception(f"Gagal dalam proses print: {str(e)}")
finally:
try:
win32print.ClosePrinter(hprinter)
except:
pass

except Exception as e:
error_msg = f"Gagal print barcode: {str(e)}"
status_label.config(text=f"❌ {error_msg}")
messagebox.showerror("Error", error_msg)
sound_mgr.play_sound("error")

# Frame untuk tombol
btn_frame = ttk.Frame(frm)
btn_frame.grid(row=8, column=0, pady=10)

ttk.Button(btn_frame, text="Generate Barcode 🏷️",
command=generate_barcode, width=18).pack(side="left", padx=5)

ttk.Button(btn_frame, text="Print Barcode 🖨️",
command=print_barcode, width=18).pack(side="left", padx=5)

# Tombol buka folder barcode
def open_barcode_folder():
barcode_dir = 'barcode_produk'
if os.path.exists(barcode_dir):
os.startfile(barcode_dir)
else:
messagebox.showinfo("Info", "Folder barcode_produk belum ada")

ttk.Button(btn_frame, text="Buka Folder",
command=open_barcode_folder, width=12).pack(side="left", padx=5)

# Tombol tutup
ttk.Button(frm, text="Tutup", command=win.destroy).grid(row=9, column=0, pady=20)

# Focus ke input
entry_teks.focus_set()
entry_teks.select_range(0, tk.END)

# Binding Enter untuk generate
def on_enter(event):
generate_barcode()

entry_teks.bind('

# ================== BACKUP & RESTORE ==================
def backup_restore_window(parent):
theme = theme_mgr.get_theme()

win = tk.Toplevel(parent)
win.title("Backup & Restore Data")
win.geometry("500x400")
win.configure(bg=theme["bg"])
win.transient(parent)
bind_fullscreen_keys(win)

style = ttk.Style(win)
style.theme_use("clam")
style.configure(".", background=theme["bg"], foreground=theme["fg"])
style.configure("TLabel", background=theme["bg"], foreground=theme["fg"], font=("Arial", 10))
style.configure("TButton", background=theme["card_bg"], foreground=theme["fg"], padding=8)

frm = ttk.Frame(win, padding=20)
frm.pack(expand=True, fill="both")

ttk.Label(frm, text="BACKUP & RESTORE DATA", font=("Arial", 14, "bold")).pack(pady=(0, 20))

def backup_data():
timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
backup_file = f"backup_kasir_{timestamp}.zip"

try:
with zipfile.ZipFile(backup_file, 'w', zipfile.ZIP_DEFLATED) as zipf:
files_to_backup = [STOK_FILE, TRANSAKSI_FILE, SETTINGS_FILE]
for file in files_to_backup:
if os.path.exists(file):
zipf.write(file)

# Backup struk directory
if os.path.exists(STRUK_DIR):
for root_dir, dirs, files in os.walk(STRUK_DIR):
for file in files:
file_path = os.path.join(root_dir, file)
arcname = os.path.relpath(file_path, start=".")
zipf.write(file_path, arcname)

messagebox.showinfo("Backup Sukses", f"Data berhasil dibackup ke:\n{backup_file}")
sound_mgr.play_sound("success")
except Exception as e:
messagebox.showerror("Backup Gagal", f"Gagal membuat backup: {str(e)}")
sound_mgr.play_sound("error")

def restore_data():
file_path = filedialog.askopenfilename(
title="Pilih file backup",
filetypes=[("ZIP files", "*.zip"), ("All files", "*.*")]
)

if not file_path:
return

if not messagebox.askyesno("Konfirmasi Restore",
"Data saat ini akan diganti dengan data dari backup.\nLanjutkan?"):
return

try:
with zipfile.ZipFile(file_path, 'r') as zipf:
zipf.extractall(".")

messagebox.showinfo("Restore Sukses", "Data berhasil di-restore!\nAplikasi akan dimulai ulang.")
sound_mgr.play_sound("success")

# Restart aplikasi
win.destroy()
parent.destroy()
login_window()

except Exception as e:
messagebox.showerror("Restore Gagal", f"Gagal restore data: {str(e)}")
sound_mgr.play_sound("error")

def export_to_excel():
try:
# Export data stok
stok_data = []
stok = load_stok()
for kode, item in stok.items():
stok_data.append([kode, item["nama"], item["harga_beli"], item["harga_jual"], item["qty"]])

df_stok = pd.DataFrame(stok_data, columns=["Kode", "Nama", "Harga Beli", "Harga Jual", "Stok"])

# Export data transaksi
transaksi_data = []
if os.path.exists(TRANSAKSI_FILE):
with open(TRANSAKSI_FILE, "r", encoding="utf-8") as f:
reader = csv.reader(f)
for row in reader:
if len(row) >= 8:
transaksi_data.append(row)

df_transaksi = pd.DataFrame(transaksi_data,
columns=["Timestamp", "Kode", "Nama", "Harga Beli", "Harga Jual",
"Qty", "Subtotal", "Laba"])
else:
df_transaksi = pd.DataFrame(columns=["Timestamp", "Kode", "Nama", "Harga Beli", "Harga Jual",
"Qty", "Subtotal", "Laba"])

# Export ke Excel
timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
excel_file = f"export_data_kasir_{timestamp}.xlsx"

with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
df_stok.to_excel(writer, sheet_name="Data Stok", index=False)
df_transaksi.to_excel(writer, sheet_name="Data Transaksi", index=False)

# Buat sheet summary
summary_data = {
'Metric': ['Total Produk', 'Total Transaksi', 'Tanggal Export'],
'Value': [len(stok_data), len(transaksi_data), datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")]
}
df_summary = pd.DataFrame(summary_data)
df_summary.to_excel(writer, sheet_name="Summary", index=False)

messagebox.showinfo("Export Sukses", f"Data berhasil diexport ke:\n{excel_file}")
sound_mgr.play_sound("success")

except ImportError:
messagebox.showerror("Error", "Library pandas tidak tersedia.\nInstall dengan: pip install pandas openpyxl")
sound_mgr.play_sound("error")
except Exception as e:
messagebox.showerror("Export Gagal", f"Gagal export data: {str(e)}")
sound_mgr.play_sound("error")

def import_from_excel():
file_path = filedialog.askopenfilename(
title="Pilih file Excel",
filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
)

if not file_path:
return

try:
# Baca data stok dari Excel
df_stok = pd.read_excel(file_path, sheet_name="Data Stok")
stok = {}

for _, row in df_stok.iterrows():
kode = str(row['Kode']).strip()
stok[kode] = {
"nama": str(row['Nama']),
"harga_beli": int(row['Harga Beli']),
"harga_jual": int(row['Harga Jual']),
"qty": int(row['Stok'])
}

save_stok(stok)
messagebox.showinfo("Import Sukses", "Data stok berhasil diimport!")
sound_mgr.play_sound("success")

except Exception as e:
messagebox.showerror("Import Gagal", f"Gagal import data: {str(e)}")
sound_mgr.play_sound("error")

# Tombol-tombol fungsi
ttk.Button(frm, text="📂 Backup Data Sekarang", command=backup_data, width=30).pack(pady=10)
ttk.Button(frm, text="🔄 Restore Data dari Backup", command=restore_data, width=30).pack(pady=10)
ttk.Button(frm, text="📊 Export Data ke Excel", command=export_to_excel, width=30).pack(pady=10)
ttk.Button(frm, text="📥 Import Data dari Excel", command=import_from_excel, width=30).pack(pady=10)
ttk.Button(frm, text="Tutup", command=win.destroy, width=30).pack(pady=20)

# ================== WINDOW LOGIN ==================
def login_window():
theme = theme_mgr.get_theme()

root = tk.Tk()
root.title("Login - Aplikasi Kasir")
root.geometry("400x300")
root.resizable(True, True)
root.configure(bg=theme["bg"])
# Bind F12 and Escape (Escape will prompt to exit app here)
bind_fullscreen_keys(root, escape_exits_app=True)

# Buat agar window berada di tengah layar
root.update_idletasks()
width = 400
height = 300
x = (root.winfo_screenwidth() // 2) - (width // 2)
y = (root.winfo_screenheight() // 2) - (height // 2)
root.geometry(f"{width}x{height}+{x}+{y}")

# Style
style = ttk.Style()
style.theme_use('clam')
style.configure('TFrame', background=theme["bg"])
style.configure('TLabel', background=theme["bg"], foreground=theme["fg"], font=('Arial', 10))
style.configure('TButton', background=theme["card_bg"], foreground=theme["fg"], font=('Arial', 10))
style.configure('TEntry', fieldbackground=theme["entry_bg"], foreground=theme["fg"])
style.configure('Accent.TButton', background=theme["accent"], foreground='white', font=('Arial', 10, 'bold'))

main_frame = ttk.Frame(root, padding=30)
main_frame.pack(fill='both', expand=True)

# Header
ttk.Label(main_frame, text="APLIKASI KASIR",
font=('Arial', 18, 'bold'),
foreground=theme["accent"]).pack(pady=10)

settings = load_settings()
toko_nama = settings.get("toko_nama", TOKO_NAMA)
ttk.Label(main_frame, text=toko_nama,
font=('Arial', 12),
foreground=theme["fg"]).pack(pady=5)

# Input fields
ttk.Label(main_frame, text="Username:").pack(anchor='w', pady=(20,5))
username_entry = ttk.Entry(main_frame, width=25, font=('Arial', 11))
username_entry.pack(fill='x', pady=5)

ttk.Label(main_frame, text="Password:").pack(anchor='w', pady=(10,5))
password_entry = ttk.Entry(main_frame, width=25, show='*', font=('Arial', 11))
password_entry.pack(fill='x', pady=5)

def do_login():
username = username_entry.get()
password = password_entry.get()

if username == ADMIN_USER and password == ADMIN_PASS:
sound_mgr.play_sound("success")
root.destroy()
main_window()
else:
messagebox.showerror("Login Gagal", "Username atau password salah!")
sound_mgr.play_sound("error")
username_entry.focus()

def on_username_enter(event):
password_entry.focus()
return "break"

def on_password_enter(event):
do_login()
return "break"

# Bind Enter key
username_entry.bind('
password_entry.bind('

# Login button
login_btn = ttk.Button(main_frame, text="MASUK", command=do_login, style="Accent.TButton")
login_btn.pack(pady=20)

# Focus ke username
username_entry.focus()

# Tampilkan tema yang aktif
current_theme_name = theme_mgr.get_theme()["name"]
ttk.Label(main_frame, text=f"Tema: {current_theme_name}",
font=('Arial', 8),
foreground=theme["fg"]).pack(side='bottom', pady=5)

root.mainloop()

# ================== WINDOW PENGATURAN - DIPERBAIKI ==================
def pengaturan_window(parent, on_theme_changed=None):
theme = theme_mgr.get_theme()
settings = load_settings()

win = tk.Toplevel(parent)
win.title("Pengaturan Aplikasi")
win.geometry("700x700")
win.configure(bg=theme["bg"])
win.transient(parent)
win.resizable(True, True)
bind_fullscreen_keys(win)

# Buat agar window berada di tengah layar
win.update_idletasks()
width = 700
height = 700
x = (win.winfo_screenwidth() // 2) - (width // 2)
y = (win.winfo_screenheight() // 2) - (height // 2)
win.geometry(f"{width}x{height}+{x}+{y}")

# Style
style = ttk.Style(win)
style.theme_use('clam')
style.configure('.', background=theme["bg"], foreground=theme["fg"])
style.configure('TLabel', background=theme["bg"], foreground=theme["fg"])
style.configure('TButton', background=theme["card_bg"], foreground=theme["fg"])
style.configure('TEntry', fieldbackground=theme["entry_bg"], foreground=theme["fg"])
style.configure('TCombobox', fieldbackground=theme["entry_bg"], foreground=theme["fg"])

def apply_theme_to_widgets():
"""Terapkan tema saat ini ke semua widget di jendela pengaturan."""
theme = theme_mgr.get_theme()
win.configure(bg=theme["bg"])

# Konfigurasi ulang gaya
style.configure('.', background=theme["bg"], foreground=theme["fg"])
style.configure('TLabel', background=theme["bg"], foreground=theme["fg"])
style.configure('TButton', background=theme["card_bg"], foreground=theme["fg"])
style.configure('TEntry', fieldbackground=theme["entry_bg"], foreground=theme["fg"])
style.configure('TCombobox', fieldbackground=theme["entry_bg"], foreground=theme["fg"])

# Perbarui frame notebook
try:
notebook.winfo_children()[0].configure(style='TFrame')
notebook.winfo_children()[1].configure(style='TFrame')
notebook.winfo_children()[2].configure(style='TFrame')
except Exception:
pass

print("Tema berhasil diterapkan secara dinamis ke jendela pengaturan.")

win.rowconfigure(0, weight=1)
win.columnconfigure(0, weight=1)

notebook = ttk.Notebook(win)
notebook.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)

# Tab Umum
frm_umum = ttk.Frame(notebook, padding=15)
notebook.add(frm_umum, text="🔧 Umum")

ttk.Label(frm_umum, text="Nama Toko:", font=('Arial', 10, 'bold')).grid(row=0, column=0, sticky="w", pady=8)
nama_toko_entry = ttk.Entry(frm_umum, width=40, font=('Arial', 10))
nama_toko_entry.insert(0, settings.get("toko_nama", TOKO_NAMA))
nama_toko_entry.grid(row=0, column=1, sticky="ew", pady=8, padx=10)

ttk.Label(frm_umum, text="Alamat Toko:", font=('Arial', 10, 'bold')).grid(row=1, column=0, sticky="w", pady=8)
alamat_toko_entry = ttk.Entry(frm_umum, width=40, font=('Arial', 10))
alamat_toko_entry.insert(0, settings.get("toko_alamat", TOKO_ALAMAT))
alamat_toko_entry.grid(row=1, column=1, sticky="ew", pady=8, padx=10)

# Pilihan Tema
ttk.Label(frm_umum, text="Pilih Tema:", font=('Arial', 10, 'bold')).grid(row=2, column=0, sticky="w", pady=(20,5))

theme_names = theme_mgr.get_theme_names()
current_theme_name = theme_mgr.get_theme()["name"]

theme_combo = ttk.Combobox(frm_umum, values=theme_names, state="readonly", width=30)
theme_combo.set(current_theme_name)
theme_combo.grid(row=2, column=1, sticky="ew", pady=8, padx=10)

# Preview tema
preview_frame = ttk.Frame(frm_umum, relief='solid', height=60)
preview_frame.grid(row=3, column=0, columnspan=2, sticky="ew", pady=10)
preview_frame.grid_propagate(False)

def update_preview():
for widget in preview_frame.winfo_children():
widget.destroy()

selected_theme_name = theme_combo.get()
selected_theme_key = theme_mgr.get_theme_by_name(selected_theme_name)
preview_theme = THEMES[selected_theme_key]

preview_frame.configure(style='Preview.TFrame')
style.configure('Preview.TFrame', background=preview_theme["bg"])

preview_label = ttk.Label(preview_frame,
text=f"Preview: {selected_theme_name}",
font=('Arial', 10, 'bold'),
background=preview_theme["bg"],
foreground=preview_theme["fg"])
preview_label.pack(pady=10)

def apply_theme():
"""Terapkan tema - DIPERBAIKI"""
selected_theme_name = theme_combo.get()
selected_theme_key = theme_mgr.get_theme_by_name(selected_theme_name)

if theme_mgr.set_theme(selected_theme_key):
apply_theme_to_widgets()
if on_theme_changed:
on_theme_changed()
update_preview()
sound_mgr.play_sound("theme_change")
else:
messagebox.showerror("Error", "Gagal mengubah tema!")
sound_mgr.play_sound("error")

# Tombol preview dan apply
btn_theme_frame = ttk.Frame(frm_umum)
btn_theme_frame.grid(row=4, column=0, columnspan=2, pady=10)

ttk.Button(btn_theme_frame, text="Preview Tema",
command=update_preview).pack(side='left', padx=5)

ttk.Button(btn_theme_frame, text="Terapkan Tema",
command=apply_theme).pack(side='left', padx=5)

# Tab Multi-Printer
frm_printer = ttk.Frame(notebook, padding=15)
notebook.add(frm_printer, text="🖨️ Multi-Printer")

ttk.Label(frm_printer, text="Pengaturan Multiple Printer",
font=('Arial', 12, 'bold')).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 15))

# Receipt Printer
ttk.Label(frm_printer, text="Printer Struk:", font=('Arial', 10, 'bold')).grid(row=1, column=0, sticky="w", pady=8)
combo_receipt = ttk.Combobox(frm_printer, values=printer_mgr.printers, state="readonly", width=30)
combo_receipt.set(settings.get("receipt_printer", printer_mgr.default_printer))
combo_receipt.grid(row=1, column=1, sticky="w", pady=8, padx=10)

ttk.Button(frm_printer, text="Test",
command=lambda: test_printer(combo_receipt.get(), "Struk")).grid(row=1, column=2, padx=5)

# Report Printer
ttk.Label(frm_printer, text="Printer Laporan:", font=('Arial', 10, 'bold')).grid(row=2, column=0, sticky="w", pady=8)
combo_report = ttk.Combobox(frm_printer, values=printer_mgr.printers, state="readonly", width=30)
combo_report.set(settings.get("report_printer", printer_mgr.default_printer))
combo_report.grid(row=2, column=1, sticky="w", pady=8, padx=10)

ttk.Button(frm_printer, text="Test",
command=lambda: test_printer(combo_report.get(), "Laporan")).grid(row=2, column=2, padx=5)

# Label Printer
ttk.Label(frm_printer, text="Printer Label:", font=('Arial', 10, 'bold')).grid(row=3, column=0, sticky="w", pady=8)
combo_label = ttk.Combobox(frm_printer, values=printer_mgr.printers, state="readonly", width=30)
combo_label.set(settings.get("label_printer", printer_mgr.default_printer))
combo_label.grid(row=3, column=1, sticky="w", pady=8, padx=10)

ttk.Button(frm_printer, text="Test",
command=lambda: test_printer(combo_label.get(), "Label")).grid(row=3, column=2, padx=5)

def test_printer(printer_name, printer_type):
if not printer_name:
messagebox.showwarning("Peringatan", f"Pilih printer {printer_type} terlebih dahulu")
sound_mgr.play_sound("error")
return

success, message = printer_mgr.test_print(printer_name, f"Test {printer_type} Printer\n{datetime.datetime.now()}\n\n")
if success:
messagebox.showinfo("Test Berhasil", f"Printer {printer_type} berhasil di-test")
sound_mgr.play_sound("success")
else:
messagebox.showerror("Test Gagal", message)
sound_mgr.play_sound("error")

# Tab Lainnya
frm_lain = ttk.Frame(notebook, padding=15)
notebook.add(frm_lain, text="⚙️ Lainnya")

ttk.Label(frm_lain, text="Encoding Printer:", font=('Arial', 10, 'bold')).grid(row=0, column=0, sticky="w", pady=8)
combo_encoding = ttk.Combobox(frm_lain, values=["cp850", "cp1252", "utf-8"], state="readonly", width=20)
combo_encoding.set(settings.get("encoding", "cp850"))
combo_encoding.grid(row=0, column=1, sticky="w", pady=8, padx=10)

# Info aplikasi
ttk.Label(frm_lain, text="Info Aplikasi:", font=('Arial', 12, 'bold')).grid(row=2, column=0, sticky="w", pady=(20, 5))

info_text = f"""
Aplikasi Kasir POS v2.0
Dibuat dengan Python Tkinter

Fitur:
✓ Transaksi penjualan
✓ Manajemen stok
✓ Laporan penjualan
✓ Backup & restore data
✓ Cetak struk thermal
✓ Multiple tema
✓ Sound feedback
✓ Multi-printer support
✓ Generator Barcode
✓ Preview & Print Laporan
"""

info_label = ttk.Label(frm_lain, text=info_text, justify="left", font=('Arial', 9))
info_label.grid(row=3, column=0, columnspan=2, sticky="w", pady=10)

def simpan_pengaturan():
"""Simpan pengaturan - DIPERBAIKI"""
# Simpan tema terlebih dahulu
selected_theme_name = theme_combo.get()
selected_theme_key = theme_mgr.get_theme_by_name(selected_theme_name)
theme_mgr.set_theme(selected_theme_key)

new_settings = {
"toko_nama": nama_toko_entry.get().strip(),
"toko_alamat": alamat_toko_entry.get().strip(),
"theme": theme_mgr.current_theme,
# Printer settings
"receipt_printer": combo_receipt.get(),
"report_printer": combo_report.get(),
"label_printer": combo_label.get(),
"encoding": combo_encoding.get()
}

if save_settings(new_settings):
# Update printer manager
printer_mgr.set_receipt_printer(new_settings["receipt_printer"])
printer_mgr.set_report_printer(new_settings["report_printer"])
printer_mgr.set_label_printer(new_settings["label_printer"])

# Terapkan tema secara dinamis
apply_theme_to_widgets()
if on_theme_changed:
on_theme_changed()

messagebox.showinfo("Sukses", "Pengaturan berhasil disimpan!")
sound_mgr.play_sound("success")
win.destroy()
else:
messagebox.showerror("Error", "Gagal menyimpan pengaturan!")
sound_mgr.play_sound("error")

def reset_pengaturan():
if messagebox.askyesno("Reset", "Yakin ingin reset semua pengaturan ke default?"):
default_settings = {
"toko_nama": TOKO_NAMA,
"toko_alamat": TOKO_ALAMAT,
"theme": "dark",
"receipt_printer": printer_mgr.default_printer,
"report_printer": printer_mgr.default_printer,
"label_printer": printer_mgr.default_printer,
"encoding": "cp850"
}
save_settings(default_settings)
# Reset theme manager juga
theme_mgr.set_theme("dark")
messagebox.showinfo("Reset Sukses", "Pengaturan telah direset ke default!")
sound_mgr.play_sound("success")
win.destroy()

# Frame tombol
btn_frame = ttk.Frame(win)
btn_frame.grid(row=1, column=0, pady=10)

ttk.Button(btn_frame, text="💾 Simpan", command=simpan_pengaturan, width=15).pack(side='left', padx=5)
ttk.Button(btn_frame, text="🔄 Reset", command=reset_pengaturan, width=15).pack(side='left', padx=5)
ttk.Button(btn_frame, text="❌ Tutup", command=win.destroy, width=15).pack(side='left', padx=5)

# Initial preview
update_preview()

# ================== FUNGSI PROSES BAYAR ==================
def proses_bayar(parent, items, tree, refresh_total):
total = refresh_total()
if total == 0:
messagebox.showwarning("Peringatan", "Belum ada barang di transaksi!")
sound_mgr.play_sound("error")
return

theme = theme_mgr.get_theme()

bayar_dialog = tk.Toplevel(parent)
bayar_dialog.title("Pembayaran")
bayar_dialog.geometry("500x400")
bayar_dialog.configure(bg=theme["bg"])
bayar_dialog.transient(parent)
bayar_dialog.grab_set()
bind_fullscreen_keys(bayar_dialog)

# Buat agar window berada di tengah layar
bayar_dialog.update_idletasks()
width = 500
height = 400
x = (bayar_dialog.winfo_screenwidth() // 2) - (width // 2)
y = (bayar_dialog.winfo_screenheight() // 2) - (height // 2)
bayar_dialog.geometry(f"{width}x{height}+{x}+{y}")

# Style
style = ttk.Style(bayar_dialog)
style.theme_use('clam')
style.configure('.', background=theme["bg"], foreground=theme["fg"])
style.configure('TLabel', background=theme["bg"], foreground=theme["fg"])
style.configure('TButton', background=theme["card_bg"], foreground=theme["fg"])
style.configure('TEntry', fieldbackground=theme["entry_bg"], foreground=theme["fg"])
style.configure('Accent.TButton', background=theme["accent"], foreground='white')
style.configure('Big.TLabel', font=('Arial', 16, 'bold'))
style.configure('BigMoney.TLabel', font=('Arial', 20, 'bold'))

main_frame = ttk.Frame(bayar_dialog, padding=30)
main_frame.pack(fill='both', expand=True)

# Header besar
ttk.Label(main_frame, text="PEMBAYARAN",
font=('Arial', 20, 'bold'),
foreground=theme["accent"]).pack(pady=10)

# Info total
total_frame = ttk.Frame(main_frame)
total_frame.pack(fill='x', pady=15)

ttk.Label(total_frame, text="TOTAL:",
style='Big.TLabel').pack(side='left')
ttk.Label(total_frame, text=format_rupiah(total),
style='BigMoney.TLabel',
foreground=theme["highlight"]).pack(side='right')

# Input uang bayar
input_frame = ttk.Frame(main_frame)
input_frame.pack(fill='x', pady=20)

ttk.Label(input_frame, text="UANG BAYAR:",
font=('Arial', 14, 'bold')).pack(pady=5)

bayar_entry = ttk.Entry(input_frame, font=('Arial', 16), width=20, justify='center')
bayar_entry.pack(pady=10)
bayar_entry.focus()

# Info kembali
kembali_frame = ttk.Frame(main_frame)
kembali_frame.pack(fill='x', pady=15)

ttk.Label(kembali_frame, text="KEMBALI:",
font=('Arial', 14, 'bold')).pack()

kembali_label = ttk.Label(kembali_frame, text="Rp 0",
font=('Arial', 18, 'bold'),
foreground=theme["success"])
kembali_label.pack(pady=5)

def hitung_kembali():
try:
bayar = parse_rupiah(bayar_entry.get())
if bayar < total:
kembali_label.config(text=f"Kurang: {format_rupiah(total - bayar)}",
foreground=theme["error"])
return False
else:
kembali = bayar - total
kembali_label.config(text=format_rupiah(kembali),
foreground=theme["success"])
return True
except:
kembali_label.config(text="Input tidak valid!", foreground=theme["error"])
return False

def selesaikan_transaksi():
try:
bayar = parse_rupiah(bayar_entry.get())
if bayar < total:
messagebox.showerror("Error", f"Uang bayar kurang {format_rupiah(total - bayar)}!")
return

# Tampilkan popup konfirmasi besar
show_transaksi_sukses(total, bayar, bayar - total)

except Exception as e:
messagebox.showerror("Error", f"Terjadi kesalahan: {str(e)}")

def auto_bayar():
"""Fungsi baru: Hitung kembali + Auto Bayar jika uang cukup"""
if hitung_kembali(): # Jika uang cukup
# Auto klik bayar setelah 100ms (biar user lihat dulu kembaliannya)
bayar_dialog.after(100, selesaikan_transaksi)

def build_receipt_text(items, total, bayar, kembali):
"""Bangun teks struk untuk dicetak"""
w = 32 # Lebar struk thermal
total_items = sum(it['qty'] for it in items)

settings = load_settings()
toko_nama = settings.get("toko_nama", TOKO_NAMA)
toko_alamat = settings.get("toko_alamat", TOKO_ALAMAT)

header = f"{toko_nama:^{w}}\n{toko_alamat:^{w}}\n"
header += "=" * w + "\n"
header += datetime.datetime.now().strftime("Tanggal: %Y-%m-%d %H:%M:%S") + "\n"
header += "-" * w + "\n"

body = ""
for it in items:
name = it["nama"][:w-16]
line1 = f"{name}\n"
line2 = f" {it['qty']} x {format_rupiah(it['harga_jual']):>12} = {format_rupiah(it['subtotal']):>12}\n"
body += line1 + line2

footer = "-" * w + "\n"
footer += f"Total Item : {total_items:>10}\n"
footer += f"TOTAL : Rp {format_rupiah(total):>12}\n"
footer += f"BAYAR : Rp {format_rupiah(bayar):>12}\n"
footer += f"KEMBALI : Rp {format_rupiah(kembali):>12}\n"
footer += "=" * w + "\n"
footer += "Terima kasih atas kunjungannya!\n"
footer += "-" * w + "\n"

receipt = header + body + footer
return receipt

def show_transaksi_sukses(total, bayar, kembali):
# Popup besar untuk transaksi sukses
sukses_window = tk.Toplevel(bayar_dialog)
sukses_window.title("Transaksi Berhasil")
sukses_window.geometry("600x500")
sukses_window.configure(bg=theme["bg"])
sukses_window.transient(bayar_dialog)
sukses_window.grab_set()
bind_fullscreen_keys(sukses_window)

# Posisi di tengah
sukses_window.update_idletasks()
width = 600
height = 500
x = (sukses_window.winfo_screenwidth() // 2) - (width // 2)
y = (sukses_window.winfo_screenheight() // 2) - (height // 2)
sukses_window.geometry(f"{width}x{height}+{x}+{y}")

main_sukses = ttk.Frame(sukses_window, padding=40)
main_sukses.pack(fill='both', expand=True)

# Icon sukses
ttk.Label(main_sukses, text="✅",
font=('Arial', 50)).pack(pady=10)

ttk.Label(main_sukses, text="TRANSAKSI BERHASIL",
font=('Arial', 24, 'bold'),
foreground=theme["success"]).pack(pady=10)

# Detail transaksi
detail_frame = ttk.Frame(main_sukses)
detail_frame.pack(fill='x', pady=30)

ttk.Label(detail_frame, text="Total:",
font=('Arial', 16)).grid(row=0, column=0, sticky='w', pady=8)
ttk.Label(detail_frame, text=format_rupiah(total),
font=('Arial', 16, 'bold')).grid(row=0, column=1, sticky='e', pady=8)

ttk.Label(detail_frame, text="Bayar:",
font=('Arial', 16)).grid(row=1, column=0, sticky='w', pady=8)
ttk.Label(detail_frame, text=format_rupiah(bayar),
font=('Arial', 16, 'bold')).grid(row=1, column=1, sticky='e', pady=8)

ttk.Label(detail_frame, text="Kembali:",
font=('Arial', 16)).grid(row=2, column=0, sticky='w', pady=8)
ttk.Label(detail_frame, text=format_rupiah(kembali),
font=('Arial', 16, 'bold'),
foreground=theme["success"]).grid(row=2, column=1, sticky='e', pady=8)

# Tombol
def tutup_semua():
# Simpan transaksi
append_transaksi(items)

# Cetak struk
rec_text = build_receipt_text(items, total, bayar, kembali)

# Simpan struk ke file
fn = os.path.join(STRUK_DIR, f"struk_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.txt")
try:
with open(fn, "w", encoding="utf-8") as f:
f.write(rec_text)
except:
pass

# Cetak ke printer
try:
raw_print(rec_text + "\n\n\n")
messagebox.showinfo("Sukses", "Transaksi berhasil & struk dicetak!")
sound_mgr.play_sound("cash")
except Exception as e:
try:
temp = tempfile.mktemp(".txt")
with open(temp, "w", encoding="utf-8") as f:
f.write(rec_text)
shell_print_tempfile(temp)
messagebox.showinfo("Sukses", "Transaksi berhasil & struk dicetak!")
sound_mgr.play_sound("cash")
except:
messagebox.showwarning("Cetak Gagal", f"Transaksi berhasil, tapi cetak gagal: {e}")
sound_mgr.play_sound("success")

# Clear transaksi saat ini
items.clear()
for item in tree.get_children():
tree.delete(item)
refresh_total()

# Tutup semua window
sukses_window.destroy()
bayar_dialog.destroy()

# Focus ke input barang untuk transaksi berikutnya
parent.focus_set()

# Bind Enter untuk tutup
sukses_window.bind('

ttk.Button(main_sukses, text="TRANSAKSI BERIKUTNYA (ENTER)",
command=tutup_semua,
style="Accent.TButton").pack(pady=20)

# Auto focus ke window sukses
sukses_window.focus_set()

# Tombol
button_frame = ttk.Frame(main_frame)
button_frame.pack(fill='x', pady=20)

ttk.Button(button_frame, text="💰 BAYAR",
command=selesaikan_transaksi,
style="Accent.TButton").pack(side='left', padx=10)

ttk.Button(button_frame, text="❌ BATAL",
command=bayar_dialog.destroy).pack(side='right', padx=10)

# Bind Enter untuk auto bayar (hitung kembali + bayar otomatis)
bayar_entry.bind('

# Bind key events
def on_key_press(event):
if event.keysym == 'Escape':
# For payment dialog Escape should close dialog (not whole app)
bayar_dialog.destroy()
elif event.keysym == 'Return':
if bayar_entry.get():
auto_bayar()

bayar_dialog.bind('

# ================== WINDOW UTAMA ==================
def main_window():
theme = theme_mgr.get_theme()
settings = load_settings()

root = tk.Tk()
root.title(f"Aplikasi Kasir - {settings.get('toko_nama', TOKO_NAMA)}")
# Bind F12 and Escape (Escape will prompt to exit app here)
bind_fullscreen_keys(root, escape_exits_app=True)

# Gunakan ukuran yang responsif (90% dari layar)
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
width = int(screen_width * 0.9)
height = int(screen_height * 0.9)
root.geometry(f"{width}x{height}")

# Buat agar window berada di tengah layar
x = (screen_width // 2) - (width // 2)
y = (screen_height // 2) - (height // 2)
root.geometry(f"{width}x{height}+{x}+{y}")

root.configure(bg=theme["bg"])
root.resizable(True, True)

# Style
style = ttk.Style()
style.theme_use('clam')

# Configure styles
style.configure('.', background=theme["bg"], foreground=theme["fg"])
style.configure('TLabel', background=theme["bg"], foreground=theme["fg"], font=('Arial', 10))
style.configure('TButton', background=theme["card_bg"], foreground=theme["fg"], font=('Arial', 9))
style.configure('TEntry', fieldbackground=theme["entry_bg"], foreground=theme["fg"])
style.configure('TFrame', background=theme["bg"])

# Treeview style
style.configure('Treeview',
background=theme["tree_bg"],
foreground=theme["tree_fg"],
fieldbackground=theme["tree_bg"])
style.configure('Treeview.Heading',
background=theme["tree_heading_bg"],
foreground='white',
font=('Arial', 10, 'bold'))
style.map('Treeview', background=[('selected', theme["accent"])])

# Data
items = []
stok = load_stok()

def reload_stok_data():
"""Muat ulang data stok dari file dan perbarui variabel stok."""
nonlocal stok
stok = load_stok()
print("Data stok telah dimuat ulang di jendela utama.")
# Mungkin di masa depan, kita bisa menambahkan pembaruan visual di sini jika diperlukan
# seperti memperbarui daftar saran jika sedang ditampilkan.

def apply_theme_to_widgets():
"""Terapkan tema saat ini ke semua widget di jendela utama."""
theme = theme_mgr.get_theme()
root.configure(bg=theme["bg"])

# Konfigurasi ulang gaya
style.configure('.', background=theme["bg"], foreground=theme["fg"])
style.configure('TLabel', background=theme["bg"], foreground=theme["fg"])
style.configure('TButton', background=theme["card_bg"], foreground=theme["fg"])
style.configure('TEntry', fieldbackground=theme["entry_bg"], foreground=theme["fg"])
style.configure('TFrame', background=theme["bg"])
style.configure('Treeview',
background=theme["tree_bg"],
foreground=theme["tree_fg"],
fieldbackground=theme["tree_bg"])
style.configure('Treeview.Heading',
background=theme["tree_heading_bg"],
foreground='white')
style.map('Treeview', background=[('selected', theme["accent"])])
style.configure('Accent.TButton', background=theme["accent"], foreground='white')

# Konfigurasi ulang widget individual
try:
suggestion_listbox_widget = getattr(root, "_suggestion_listbox", None)
if suggestion_listbox_widget:
suggestion_listbox_widget.config(bg=theme["entry_bg"], fg=theme["fg"],
selectbackground=theme["accent"], selectforeground='white')
total_label.config(foreground=theme["highlight"])
item_count_label.config(foreground=theme["fg"])
search_info_label.config(foreground=theme["fg"])

# Perbarui label header
header_frame.winfo_children()[0].config(foreground=theme["accent"])
header_frame.winfo_children()[1].config(foreground=theme["fg"])

# Perbarui label status bar
for child in status_frame.winfo_children():
child.configure(foreground=theme["fg"])
except Exception:
pass

print("Tema berhasil diterapkan secara dinamis ke jendela utama.")

def update_dashboard():
"""Ambil data terbaru dan perbarui label dasbor."""
total_penjualan, jumlah_transaksi = get_dashboard_stats()
stok_kritis = get_stok_kritis()

total_penjualan_label.config(text=f"Total Penjualan Hari Ini:\n{format_rupiah(total_penjualan)}")
jumlah_transaksi_label.config(text=f"Jumlah Transaksi Hari Ini:\n{jumlah_transaksi}")
stok_kritis_label.config(text=f"Stok Kritis:\n{stok_kritis}")
print("Dashboard diperbarui.")

# Main container dengan grid yang responsif
main_container = ttk.Frame(root)
main_container.pack(fill='both', expand=True, padx=10, pady=10)

# Configure grid weights untuk responsif
main_container.columnconfigure(0, weight=1)
main_container.rowconfigure(3, weight=1) # Baris Treeview sekarang adalah 3

# Header
header_frame = ttk.Frame(main_container, padding=15)
header_frame.grid(row=0, column=0, sticky='ew', pady=(0, 10))
header_frame.columnconfigure(0, weight=1)

ttk.Label(header_frame, text=settings.get('toko_nama', TOKO_NAMA),
font=('Arial', 20, 'bold')).grid(row=0, column=0, sticky='w')
ttk.Label(header_frame, text=settings.get('toko_alamat', TOKO_ALAMAT),
font=('Arial', 12),
foreground=theme["fg"]).grid(row=1, column=0, sticky='w')

# Dashboard Frame
dashboard_frame = ttk.Frame(main_container, padding=10, relief="solid")
dashboard_frame.grid(row=1, column=0, sticky="ew", pady=10)
for i in range(4): # 3 untuk stat, 1 untuk refresh
dashboard_frame.columnconfigure(i, weight=1)

# Placeholder untuk statistik
total_penjualan_label = ttk.Label(dashboard_frame, text="Total Penjualan Hari Ini:\nRp 0", font=('Arial', 10))
total_penjualan_label.grid(row=0, column=0, padx=10, pady=5)

jumlah_transaksi_label = ttk.Label(dashboard_frame, text="Jumlah Transaksi Hari Ini:\n0", font=('Arial', 10))
jumlah_transaksi_label.grid(row=0, column=1, padx=10, pady=5)

stok_kritis_label = ttk.Label(dashboard_frame, text="Stok Kritis:\n-", font=('Arial', 10), justify="left")
stok_kritis_label.grid(row=0, column=2, padx=10, pady=5, sticky="w")

# Tombol Refresh
refresh_button = ttk.Button(dashboard_frame, text="🔄 Refresh", command=update_dashboard)
refresh_button.grid(row=0, column=3, padx=10, pady=5, sticky="e")

# Input section
input_frame = ttk.Frame(main_container, padding=10)
input_frame.grid(row=2, column=0, sticky='ew', pady=(0, 10))

# Configure grid untuk input frame
for i in range(3): # Disesuaikan menjadi 3 kolom
input_frame.columnconfigure(i, weight=1)

# Labels
ttk.Label(input_frame, text="Kode / Nama Barang:", font=('Arial', 10, 'bold')).grid(row=0, column=0, sticky='w', padx=5, pady=2)
ttk.Label(input_frame, text="Harga Jual:", font=('Arial', 10, 'bold')).grid(row=0, column=1, sticky='w', padx=5, pady=2)
ttk.Label(input_frame, text="Qty:", font=('Arial', 10, 'bold')).grid(row=0, column=2, sticky='w', padx=5, pady=2)

# Entries
kode_entry = ttk.Entry(input_frame, font=('Arial', 10))
harga_entry = ttk.Entry(input_frame, font=('Arial', 10))
qty_entry = ttk.Entry(input_frame, font=('Arial', 10))

kode_entry.grid(row=1, column=0, padx=5, pady=5, sticky='ew')
harga_entry.grid(row=1, column=1, padx=5, pady=5, sticky='ew')
qty_entry.grid(row=1, column=2, padx=5, pady=5, sticky='ew')

# Info pencarian
search_info_label = ttk.Label(input_frame, text="🔍 Tekan F2 untuk fokus pencarian • ↑↓ untuk navigasi • Enter untuk pilih",
font=('Arial', 8), foreground=theme["fg"])
search_info_label.grid(row=2, column=0, columnspan=3, sticky='w', padx=5, pady=2)

# Suggestion popup variables (we'll use a top-level popup to avoid being covered)
suggestion_win = None
suggestion_listbox = None
suggestion_visible = False
current_suggestions = []

def create_suggestion_popup():
"""Create a Toplevel popup that contains the suggestion listbox.
This prevents it from being covered by other frames/widgets."""
nonlocal suggestion_win, suggestion_listbox
if suggestion_win and tk.Toplevel.winfo_exists(suggestion_win):
return
suggestion_win = tk.Toplevel(root)
suggestion_win.overrideredirect(True)
suggestion_win.transient(root)
# Keep it above other windows in the app
try:
suggestion_win.attributes('-topmost', True)
except Exception:
pass
suggestion_listbox = tk.Listbox(suggestion_win, height=8, bg=theme["entry_bg"], fg=theme["fg"],
selectbackground=theme["accent"], font=('Arial', 9),
selectforeground='white', activestyle='none')
suggestion_listbox.pack(fill='both', expand=True)
# Save reference on root for theme updates
root._suggestion_listbox = suggestion_listbox

# Bind selection and keyboard events for the popup listbox
suggestion_listbox.bind('
suggestion_listbox.bind('
suggestion_listbox.bind('
suggestion_listbox.bind('
suggestion_listbox.bind('

def show_suggestions():
nonlocal suggestion_visible, current_suggestions
query = kode_entry.get().lower().strip()
# If no query, still show some suggestions? Current logic filters with query in code/name.
matches = []
for kode, data in stok.items():
if (query in kode.lower() or query in data['nama'].lower()):
matches.append((kode, data))
# Sort by relevance: exact match first, then partial matches
exact_matches = [m for m in matches if query == m[0].lower() or query == m[1]['nama'].lower()]
partial_matches = [m for m in matches if m not in exact_matches]
matches_sorted = exact_matches + partial_matches
current_suggestions = matches_sorted[:10]

# If no suggestions, hide the popup
if not current_suggestions:
hide_suggestions()
return

# Create popup if not exists
create_suggestion_popup()

# Fill listbox
suggestion_listbox.delete(0, tk.END)
for kode, data in current_suggestions:
display_text = f"{kode} - {data['nama']} (Stok: {data['qty']}) - {format_rupiah(data['harga_jual'])}"
suggestion_listbox.insert(tk.END, display_text)

# Position the popup right under the entry
try:
# Coordinates relative to the screen
x = kode_entry.winfo_rootx()
y = kode_entry.winfo_rooty() + kode_entry.winfo_height()
width = kode_entry.winfo_width()
# Set geometry of the toplevel
suggestion_win.geometry(f"{width}x200+{x}+{y}")
suggestion_win.deiconify()
suggestion_win.lift(aboveThis=root)
suggestion_visible = True
# Select first item
suggestion_listbox.selection_clear(0, tk.END)
suggestion_listbox.selection_set(0)
suggestion_listbox.activate(0)
suggestion_listbox.focus_set()
except Exception:
# Fallback: place within root if geometry fails
suggestion_listbox.place(x=kode_entry.winfo_x(), y=kode_entry.winfo_y() + kode_entry.winfo_height(), width=kode_entry.winfo_width())
suggestion_visible = True

def hide_suggestions():
nonlocal suggestion_visible, suggestion_win, suggestion_listbox
try:
if suggestion_win:
suggestion_win.withdraw()
try:
suggestion_win.destroy()
except Exception:
pass
suggestion_win = None
suggestion_listbox = None
suggestion_visible = False
# cleanup reference on root
if hasattr(root, "_suggestion_listbox"):
try:
delattr = getattr(root, "__delattr__", None)
except Exception:
delattr = None
try:
del root._suggestion_listbox
except Exception:
pass
except Exception:
suggestion_visible = False
suggestion_win = None
suggestion_listbox = None

def on_suggestion_select(event=None):
nonlocal suggestion_listbox, current_suggestions
try:
if suggestion_listbox is None:
return
sel = suggestion_listbox.curselection()
if sel and current_suggestions:
index = sel[0]
if index < len(current_suggestions):
kode, data = current_suggestions[index]
kode_entry.delete(0, tk.END)
kode_entry.insert(0, kode)
harga_entry.delete(0, tk.END)
harga_entry.insert(0, format_rupiah(data['harga_jual']))
qty_entry.delete(0, tk.END)
qty_entry.insert(0, "1")
qty_entry.focus()
qty_entry.select_range(0, tk.END)
hide_suggestions()
except Exception:
hide_suggestions()

def navigate_suggestions(direction):
"""Navigasi suggestion list with arrow keys"""
nonlocal suggestion_listbox
try:
if suggestion_listbox is None or not suggestion_visible:
return
size = suggestion_listbox.size()
if size == 0:
return
cur = suggestion_listbox.curselection()
if not cur:
idx = 0
else:
idx = cur[0]
if direction == 'up':
new_idx = max(0, idx - 1)
else:
new_idx = min(size - 1, idx + 1)
suggestion_listbox.selection_clear(0, tk.END)
suggestion_listbox.selection_set(new_idx)
suggestion_listbox.activate(new_idx)
suggestion_listbox.see(new_idx)
except Exception:
pass

def clear_and_focus_search():
"""Fokus ke pencarian dan clear field"""
kode_entry.delete(0, tk.END)
kode_entry.focus()
show_suggestions()

def quick_search_by_code(code_prefix):
"""Pencarian cepat berdasarkan prefix kode"""
kode_entry.delete(0, tk.END)
kode_entry.insert(0, code_prefix)
kode_entry.focus()
show_suggestions()

# Bind events for search entry
kode_entry.bind('
kode_entry.bind('
kode_entry.bind('
kode_entry.bind('
kode_entry.bind('

# Bind F2 untuk fokus pencarian
root.bind('

# Bind angka untuk quick search
root.bind('
root.bind('
root.bind('

# Treeview untuk transaksi
tree_frame = ttk.Frame(main_container)
tree_frame.grid(row=3, column=0, sticky='nsew', pady=(0, 10))
main_container.rowconfigure(2, weight=1)
tree_frame.rowconfigure(0, weight=1)
tree_frame.columnconfigure(0, weight=1)

columns = ('kode', 'nama', 'harga', 'qty', 'subtotal')
tree = ttk.Treeview(tree_frame, columns=columns, show='headings')

# Define headings
tree.heading('kode', text='KODE')
tree.heading('nama', text='NAMA BARANG')
tree.heading('harga', text='HARGA')
tree.heading('qty', text='QTY')
tree.heading('subtotal', text='SUBTOTAL')

# Define columns dengan lebar yang responsif
tree.column('kode', width=120, anchor='center', minwidth=100)
tree.column('nama', width=300, anchor='w', minwidth=200)
tree.column('harga', width=150, anchor='e', minwidth=100)
tree.column('qty', width=80, anchor='center', minwidth=60)
tree.column('subtotal', width=150, anchor='e', minwidth=100)

# Scrollbar untuk treeview
scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
tree.configure(yscrollcommand=scrollbar.set)

tree.grid(row=0, column=0, sticky='nsew')
scrollbar.grid(row=0, column=1, sticky='ns')

def refresh_total():
total = sum(item['subtotal'] for item in items)
total_label.config(text=f"TOTAL: {format_rupiah(total)}")
item_count = sum(item['qty'] for item in items)
item_count_label.config(text=f"{len(items)} item • {item_count} barang")
return total

def tambah_barang():
kode_input = kode_entry.get().strip()
if not kode_input:
messagebox.showerror("Error", "Masukkan kode atau nama barang!")
sound_mgr.play_sound("error")
kode_entry.focus()
return

barang_ditemukan = None
for kode, data in stok.items():
if (kode_input.lower() == kode.lower() or
kode_input.lower() in data['nama'].lower()):
barang_ditemukan = (kode, data)
break

if not barang_ditemukan:
messagebox.showerror("Error", "Barang tidak ditemukan di stok!")
sound_mgr.play_sound("error")
kode_entry.focus()
show_suggestions()
return

kode, data = barang_ditemukan

try:
qty = int(qty_entry.get() or "1")
harga_jual = parse_rupiah(harga_entry.get()) or data['harga_jual']
except ValueError:
messagebox.showerror("Error", "Qty dan harga harus angka!")
sound_mgr.play_sound("error")
return

if qty <= 0:
messagebox.showerror("Error", "Qty harus lebih dari 0!")
sound_mgr.play_sound("error")
return

if qty > data['qty']:
messagebox.showerror("Stok Tidak Cukup",
f"Stok {data['nama']} hanya {data['qty']}!")
sound_mgr.play_sound("error")
return

subtotal = harga_jual * qty
laba = (harga_jual - data['harga_beli']) * qty

item = {
'kode': kode,
'nama': data['nama'],
'harga_beli': data['harga_beli'],
'harga_jual': harga_jual,
'qty': qty,
'subtotal': subtotal,
'laba': laba
}
items.append(item)

tree.insert('', 'end', values=(
kode,
data['nama'],
format_rupiah(harga_jual),
qty,
format_rupiah(subtotal)
))

stok[kode]['qty'] -= qty
save_stok(stok)

kode_entry.delete(0, 'end')
harga_entry.delete(0, 'end')
qty_entry.delete(0, 'end')
kode_entry.focus()
hide_suggestions()

refresh_total()
sound_mgr.play_sound("add_item")

def hapus_barang():
selected = tree.selection()
if not selected:
messagebox.showwarning("Peringatan", "Pilih barang yang akan dihapus terlebih dahulu!")
sound_mgr.play_sound("error")
return

if messagebox.askyesno("Konfirmasi", "Hapus barang dari transaksi?"):
for item in selected:
values = tree.item(item)['values']
if values:
kode = values[0]
qty = values[3]

if kode in stok:
stok[kode]['qty'] += qty

for i, it in enumerate(items):
if it['kode'] == kode and it['qty'] == qty:
items.pop(i)
break

tree.delete(item)

save_stok(stok)
refresh_total()
sound_mgr.play_sound("delete")

# Button frame
button_frame = ttk.Frame(main_container, padding=10)
button_frame.grid(row=4, column=0, sticky='ew', pady=(0, 10))

# Configure grid untuk button frame
for i in range(3):
button_frame.columnconfigure(i, weight=1)

ttk.Button(button_frame, text="➕ TAMBAH BARANG",
command=tambah_barang).grid(row=0, column=0, padx=2, sticky='ew')

ttk.Button(button_frame, text="🗑️ HAPUS BARANG",
command=hapus_barang).grid(row=0, column=1, padx=2, sticky='ew')

ttk.Button(button_frame, text="💰 PROSES PEMBAYARAN (END)",
command=lambda: proses_bayar(root, items, tree, refresh_total),
style="Accent.TButton").grid(row=0, column=2, padx=2, sticky='ew')

# Total frame
total_frame = ttk.Frame(main_container, padding=10)
total_frame.grid(row=5, column=0, sticky='ew', pady=(0, 10))

total_label = ttk.Label(total_frame, text="TOTAL: Rp 0",
font=('Arial', 18, 'bold'),
foreground=theme["highlight"])
total_label.pack(side='left')

item_count_label = ttk.Label(total_frame, text="Belum ada transaksi",
font=('Arial', 10),
foreground=theme["fg"])
item_count_label.pack(side='right')

# Menu bar
menubar = tk.Menu(root)
root.config(menu=menubar)

# Menu File
file_menu = tk.Menu(menubar, tearoff=0)
menubar.add_cascade(label="File", menu=file_menu)
file_menu.add_command(label="Backup & Restore", command=lambda: backup_restore_window(root))
file_menu.add_separator()
file_menu.add_command(label="Keluar", command=root.quit)

# Menu Data
data_menu = tk.Menu(menubar, tearoff=0)
menubar.add_cascade(label="Data", menu=data_menu)
data_menu.add_command(label="Kelola Stok", command=lambda: kelola_stok_window(root, on_stok_changed=reload_stok_data))
data_menu.add_command(label="Laporan Transaksi", command=lambda: laporan_window(root))
data_menu.add_command(label="Generator Barcode", command=lambda: barcode_generator_window(root))

# Menu Pengaturan
pengaturan_menu = tk.Menu(menubar, tearoff=0)
menubar.add_cascade(label="Pengaturan", menu=pengaturan_menu)
pengaturan_menu.add_command(label="Pengaturan Aplikasi", command=lambda: pengaturan_window(root, on_theme_changed=apply_theme_to_widgets))

# Menu Bantuan
help_menu = tk.Menu(menubar, tearoff=0)
menubar.add_cascade(label="Bantuan", menu=help_menu)
help_menu.add_command(label="Tentang", command=lambda: about_window(root))

# Submenu Shortcut
shortcut_menu = tk.Menu(help_menu, tearoff=0)
help_menu.add_cascade(label="Shortcut Keyboard", menu=shortcut_menu)
shortcut_menu.add_command(label="F2 - Fokus Pencarian", command=lambda: None)
shortcut_menu.add_command(label="↑↓ - Navigasi Suggestion", command=lambda: None)
shortcut_menu.add_command(label="Enter - Pilih/Tambah Barang", command=lambda: None)
shortcut_menu.add_command(label="End - Proses Pembayaran", command=lambda: None)
shortcut_menu.add_command(label="F3-F5 - Pencarian Cepat", command=lambda: None)

# Bind events untuk responsif
def on_resize(event):
# Update popup position when root resizes or moves
try:
if suggestion_visible and suggestion_win and kode_entry:
x = kode_entry.winfo_rootx()
y = kode_entry.winfo_rooty() + kode_entry.winfo_height()
width = kode_entry.winfo_width()
try:
suggestion_win.geometry(f"{width}x200+{x}+{y}")
except Exception:
pass
except Exception:
pass

root.bind('

# Bind Enter key untuk tambah barang
qty_entry.bind('

# Bind End key untuk proses pembayaran
root.bind('

# Click outside to hide suggestions
def on_click_outside(event):
try:
widget = event.widget
# if click is not on kode_entry or suggestion popup, hide suggestions
if widget not in [kode_entry, suggestion_listbox]:
hide_suggestions()
except Exception:
hide_suggestions()

root.bind('

# Status bar dengan info tema dan shortcut
status_frame = ttk.Frame(main_container, padding=5)
status_frame.grid(row=6, column=0, sticky='ew')

ttk.Label(status_frame, text=f"Tema: {theme_mgr.get_theme()['name']}",
font=('Arial', 8),
foreground=theme["fg"]).pack(side='right')

ttk.Label(status_frame, text="F2: Cari • ↑↓: Navigasi • Enter: Pilih • End: Bayar",
font=('Arial', 8),
foreground=theme["fg"]).pack(side='left')

# Auto focus ke input kode saat aplikasi dibuka
kode_entry.focus()

# Muat data dasbor awal
update_dashboard()

root.mainloop()

# ================== WINDOW KELOLA STOK ==================
def kelola_stok_window(parent, on_stok_changed=None):
theme = theme_mgr.get_theme()
stok = load_stok()

win = tk.Toplevel(parent)
win.title("Kelola Stok Barang")
win.geometry("900x600")
win.configure(bg=theme["bg"])
win.transient(parent)
win.resizable(True, True)
bind_fullscreen_keys(win)

# Buat agar window berada di tengah layar
win.update_idletasks()
width = 900
height = 600
x = (win.winfo_screenwidth() // 2) - (width // 2)
y = (win.winfo_screenheight() // 2) - (height // 2)
win.geometry(f"{width}x{height}+{x}+{y}")

# Style
style = ttk.Style(win)
style.theme_use('clam')
style.configure('.', background=theme["bg"], foreground=theme["fg"])
style.configure('TLabel', background=theme["bg"], foreground=theme["fg"])
style.configure('TButton', background=theme["card_bg"], foreground=theme["fg"])
style.configure('TEntry', fieldbackground=theme["entry_bg"], foreground=theme["fg"])
style.configure('Treeview', background=theme["tree_bg"], foreground=theme["tree_fg"])
style.configure('Treeview.Heading', background=theme["tree_heading_bg"], foreground='white')

main_frame = ttk.Frame(win, padding=20)
main_frame.pack(fill='both', expand=True)
main_frame.rowconfigure(2, weight=1)
main_frame.columnconfigure(0, weight=1)

ttk.Label(main_frame, text="KELOLA STOK BARANG",
font=('Arial', 16, 'bold'),
foreground=theme["accent"]).grid(row=0, column=0, sticky="w", pady=10)

# Input frame
input_frame = ttk.Frame(main_frame)
input_frame.grid(row=1, column=0, sticky="ew", pady=10)

ttk.Label(input_frame, text="Kode Barang:").grid(row=0, column=0, padx=5, pady=5, sticky='w')
kode_entry = ttk.Entry(input_frame, width=15)
kode_entry.grid(row=0, column=1, padx=5, pady=5)

ttk.Label(input_frame, text="Nama Barang:").grid(row=0, column=2, padx=5, pady=5, sticky='w')
nama_entry = ttk.Entry(input_frame, width=30)
nama_entry.grid(row=0, column=3, padx=5, pady=5)

ttk.Label(input_frame, text="Harga Beli:").grid(row=1, column=0, padx=5, pady=5, sticky='w')
harga_beli_entry = ttk.Entry(input_frame, width=15)
harga_beli_entry.grid(row=1, column=1, padx=5, pady=5)

ttk.Label(input_frame, text="Harga Jual:").grid(row=1, column=2, padx=5, pady=5, sticky='w')
harga_jual_entry = ttk.Entry(input_frame, width=15)
harga_jual_entry.grid(row=1, column=3, padx=5, pady=5)

ttk.Label(input_frame, text="Stok:").grid(row=1, column=4, padx=5, pady=5, sticky='w')
stok_entry = ttk.Entry(input_frame, width=10)
stok_entry.grid(row=1, column=5, padx=5, pady=5)

# Button frame
button_frame = ttk.Frame(input_frame)
button_frame.grid(row=2, column=0, columnspan=6, pady=10)

def tambah_stok():
kode = kode_entry.get().strip()
nama = nama_entry.get().strip()
try:
harga_beli = parse_rupiah(harga_beli_entry.get())
harga_jual = parse_rupiah(harga_jual_entry.get())
stok_qty = int(stok_entry.get() or "0")
except:
messagebox.showerror("Error", "Input harga dan stok harus angka!")
return

if not kode or not nama:
messagebox.showerror("Error", "Kode dan nama barang harus diisi!")
return

if kode in stok:
messagebox.showerror("Error", f"Kode {kode} sudah ada!")
return

stok[kode] = {
"nama": nama,
"harga_beli": harga_beli,
"harga_jual": harga_jual,
"qty": stok_qty
}

success, msg = save_stok(stok)
if success:
if on_stok_changed:
on_stok_changed()
refresh_tree()
clear_entries()
messagebox.showinfo("Sukses", "Barang berhasil ditambahkan!")
else:
messagebox.showerror("Error", f"Gagal menyimpan: {msg}")

def edit_stok():
selected_iids = tree.selection()
if not selected_iids:
messagebox.showwarning("Peringatan", "Pilih barang yang akan diedit!")
return

kode = selected_iids[0]
if kode not in stok:
messagebox.showerror("Error", "Barang tidak ditemukan!")
return

nama = nama_entry.get().strip()
try:
harga_beli = parse_rupiah(harga_beli_entry.get())
harga_jual = parse_rupiah(harga_jual_entry.get())
stok_qty = int(stok_entry.get() or "0")
except:
messagebox.showerror("Error", "Input harga dan stok harus angka!")
return

if not nama:
messagebox.showerror("Error", "Nama barang harus diisi!")
return

stok[kode] = {
"nama": nama,
"harga_beli": harga_beli,
"harga_jual": harga_jual,
"qty": stok_qty
}

success, msg = save_stok(stok)
if success:
if on_stok_changed:
on_stok_changed()
refresh_tree()
clear_entries()
messagebox.showinfo("Sukses", "Barang berhasil diupdate!")
else:
messagebox.showerror("Error", f"Gagal menyimpan: {msg}")

def hapus_stok():
selected_iids = tree.selection()
if not selected_iids:
messagebox.showwarning("Peringatan", "Pilih barang yang akan dihapus!")
return

kode = selected_iids[0]
if messagebox.askyesno("Konfirmasi", f"Hapus barang {kode}?"):
if kode in stok:
del stok[kode]
success, msg = save_stok(stok)
if success:
if on_stok_changed:
on_stok_changed()
refresh_tree()
clear_entries()
messagebox.showinfo("Sukses", "Barang berhasil dihapus!")
else:
messagebox.showerror("Error", f"Gagal menghapus: {msg}")

def clear_entries():
kode_entry.delete(0, 'end')
nama_entry.delete(0, 'end')
harga_beli_entry.delete(0, 'end')
harga_jual_entry.delete(0, 'end')
stok_entry.delete(0, 'end')

def on_tree_select(event):
selected = tree.selection()
if not selected:
return
values = tree.item(selected[0])['values']
if values:
kode_entry.delete(0, 'end')
kode_entry.insert(0, values[0])
nama_entry.delete(0, 'end')
nama_entry.insert(0, values[1])
harga_beli_entry.delete(0, 'end')
harga_beli_entry.insert(0, format_rupiah(values[2]))
harga_jual_entry.delete(0, 'end')
harga_jual_entry.insert(0, format_rupiah(values[3]))
stok_entry.delete(0, 'end')
stok_entry.insert(0, str(values[4]))

ttk.Button(button_frame, text="➕ TAMBAH",
command=tambah_stok).pack(side='left', padx=5)

ttk.Button(button_frame, text="✏️ EDIT",
command=edit_stok).pack(side='left', padx=5)

ttk.Button(button_frame, text="🗑️ HAPUS",
command=hapus_stok).pack(side='left', padx=5)

ttk.Button(button_frame, text="🧹 BERSIHKAN",
command=clear_entries).pack(side='left', padx=5)

# Treeview
tree_frame = ttk.Frame(main_frame)
tree_frame.grid(row=2, column=0, sticky='nsew', pady=10)
tree_frame.rowconfigure(0, weight=1)
tree_frame.columnconfigure(0, weight=1)

columns = ('kode', 'nama', 'harga_beli', 'harga_jual', 'stok')
tree = ttk.Treeview(tree_frame, columns=columns, show='headings')

tree.heading('kode', text='KODE')
tree.heading('nama', text='NAMA BARANG')
tree.heading('harga_beli', text='HARGA BELI')
tree.heading('harga_jual', text='HARGA JUAL')
tree.heading('stok', text='STOK')

tree.column('kode', width=100)
tree.column('nama', width=300)
tree.column('harga_beli', width=120)
tree.column('harga_jual', width=120)
tree.column('stok', width=80)

scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
tree.configure(yscrollcommand=scrollbar.set)

tree.grid(row=0, column=0, sticky='nsew')
scrollbar.grid(row=0, column=1, sticky='ns')

tree.bind('<

def refresh_tree():
tree.delete(*tree.get_children())
for kode, data in stok.items():
tree.insert('', 'end', iid=kode, values=(
kode,
data['nama'],
format_rupiah(data['harga_beli']),
format_rupiah(data['harga_jual']),
data['qty']
))

refresh_tree()

# ================== WINDOW LAPORAN ==================
def laporan_window(parent):
theme = theme_mgr.get_theme()

win = tk.Toplevel(parent)
win.title("Laporan Transaksi")
win.geometry("1100x700")
win.configure(bg=theme["bg"])
win.transient(parent)
win.resizable(True, True)
bind_fullscreen_keys(win)

# Buat agar window berada di tengah layar
win.update_idletasks()
width = 1100
height = 700
x = (win.winfo_screenwidth() // 2) - (width // 2)
y = (win.winfo_screenheight() // 2) - (height // 2)
win.geometry(f"{width}x{height}+{x}+{y}")

# Style
style = ttk.Style(win)
style.theme_use('clam')
style.configure('.', background=theme["bg"], foreground=theme["fg"])
style.configure('TLabel', background=theme["bg"], foreground=theme["fg"])
style.configure('TButton', background=theme["card_bg"], foreground=theme["fg"])
style.configure('Treeview', background=theme["tree_bg"], foreground=theme["tree_fg"])
style.configure('Treeview.Heading', background=theme["tree_heading_bg"], foreground='white')

main_frame = ttk.Frame(win, padding=20)
main_frame.pack(fill='both', expand=True)
main_frame.rowconfigure(2, weight=1)
main_frame.columnconfigure(0, weight=1)

ttk.Label(main_frame, text="LAPORAN TRANSAKSI",
font=('Arial', 16, 'bold'),
foreground=theme["accent"]).grid(row=0, column=0, sticky="w", pady=10)

# Filter frame
filter_frame = ttk.Frame(main_frame)
filter_frame.grid(row=1, column=0, sticky="ew", pady=10)

ttk.Label(filter_frame, text="Tanggal Mulai:").grid(row=0, column=0, padx=5, pady=5)
start_date = ttk.Entry(filter_frame, width=12)
start_date.grid(row=0, column=1, padx=5, pady=5)
start_date.insert(0, datetime.datetime.now().strftime("%Y-%m-%d"))

ttk.Label(filter_frame, text="Tanggal Akhir:").grid(row=0, column=2, padx=5, pady=5)
end_date = ttk.Entry(filter_frame, width=12)
end_date.grid(row=0, column=3, padx=5, pady=5)
end_date.insert(0, datetime.datetime.now().strftime("%Y-%m-%d"))

def load_transaksi():
if not os.path.exists(TRANSAKSI_FILE):
messagebox.showinfo("Info", "Belum ada data transaksi!")
return []

transaksi = []
try:
with open(TRANSAKSI_FILE, "r", encoding="utf-8") as f:
reader = csv.reader(f)
for row in reader:
if len(row) >= 8:
try:
transaksi.append({
'timestamp': row[0],
'kode': row[1],
'nama': row[2],
'harga_beli': int(row[3]),
'harga_jual': int(row[4]),
'qty': int(row[5]),
'subtotal': int(row[6]),
'laba': int(row[7])
})
except:
continue
except:
pass
return transaksi

def filter_transaksi():
try:
start = datetime.datetime.strptime(start_date.get(), "%Y-%m-%d")
end = datetime.datetime.strptime(end_date.get(), "%Y-%m-%d")
except:
messagebox.showerror("Error", "Format tanggal salah! Gunakan YYYY-MM-DD")
return []

all_transaksi = load_transaksi()
filtered = []
for t in all_transaksi:
try:
t_date = datetime.datetime.strptime(t['timestamp'].split()[0], "%Y-%m-%d")
if start <= t_date <= end:
filtered.append(t)
except:
continue
return filtered

# Variabel untuk menyimpan data transaksi saat ini
current_transaksi = []
current_total_penjualan = 0
current_total_laba = 0
current_total_item = 0

def refresh_laporan():
nonlocal current_transaksi, current_total_penjualan, current_total_laba, current_total_item

transaksi = filter_transaksi()
tree.delete(*tree.get_children())

total_penjualan = 0
total_laba = 0
total_item = 0

for i, t in enumerate(transaksi):
tree.insert('', 'end', iid=i, values=(
t['timestamp'],
t['kode'],
t['nama'],
format_rupiah(t['harga_jual']),
t['qty'],
format_rupiah(t['subtotal']),
format_rupiah(t['laba'])
))
total_penjualan += t['subtotal']
total_laba += t['laba']
total_item += t['qty']

# Simpan data untuk print
current_transaksi = transaksi
current_total_penjualan = total_penjualan
current_total_laba = total_laba
current_total_item = total_item

total_label.config(text=f"Total Penjualan: {format_rupiah(total_penjualan)}")
laba_label.config(text=f"Total Laba: {format_rupiah(total_laba)}")
item_label.config(text=f"Total Item Terjual: {total_item}")

# Tombol filter dan print
button_filter_frame = ttk.Frame(filter_frame)
button_filter_frame.grid(row=0, column=4, columnspan=3, padx=10, pady=5)

ttk.Button(button_filter_frame, text="🔍 MUAT LAPORAN",
command=refresh_laporan).pack(side='left', padx=5)

ttk.Button(button_filter_frame, text="🖨️ PRINT LAPORAN",
command=lambda: print_laporan_keseluruhan(
win, current_transaksi, current_total_penjualan,
current_total_laba, current_total_item
)).pack(side='left', padx=5)

# Treeview dengan context menu
tree_frame = ttk.Frame(main_frame)
tree_frame.grid(row=2, column=0, sticky='nsew', pady=10)
tree_frame.rowconfigure(0, weight=1)
tree_frame.columnconfigure(0, weight=1)

columns = ('waktu', 'kode', 'nama', 'harga', 'qty', 'subtotal', 'laba')
tree = ttk.Treeview(tree_frame, columns=columns, show='headings')

tree.heading('waktu', text='WAKTU')
tree.heading('kode', text='KODE')
tree.heading('nama', text='NAMA BARANG')
tree.heading('harga', text='HARGA JUAL')
tree.heading('qty', text='QTY')
tree.heading('subtotal', text='SUBTOTAL')
tree.heading('laba', text='LABA')

tree.column('waktu', width=150)
tree.column('kode', width=100)
tree.column('nama', width=200)
tree.column('harga', width=120)
tree.column('qty', width=80)
tree.column('subtotal', width=120)
tree.column('laba', width=120)

scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
tree.configure(yscrollcommand=scrollbar.set)

tree.grid(row=0, column=0, sticky='nsew')
scrollbar.grid(row=0, column=1, sticky='ns')

# Context menu untuk print per barang
context_menu = tk.Menu(tree, tearoff=0)
context_menu.add_command(label="📄 Preview & Print Struk",
command=lambda: print_selected_transaksi())
context_menu.add_separator()
context_menu.add_command(label="📋 Copy Data",
command=lambda: copy_selected_data())

def show_context_menu(event):
item = tree.identify_row(event.y)
if item:
tree.selection_set(item)
context_menu.post(event.x_root, event.y_root)

def print_selected_transaksi():
selected_iids = tree.selection()
if not selected_iids:
messagebox.showwarning("Peringatan", "Pilih transaksi terlebih dahulu!")
return

selected_iid = selected_iids[0]
try:
# The iid is the index of the transaction in current_transaksi
transaksi_index = int(selected_iid)
if 0 <= transaksi_index < len(current_transaksi):
transaksi_data = current_transaksi[transaksi_index]
print_struk_satuan_laporan(win, transaksi_data)
else:
messagebox.showerror("Error", "Indeks transaksi tidak valid!")
except (ValueError, IndexError):
messagebox.showerror("Error", "Data transaksi tidak ditemukan!")

def copy_selected_data():
selected = tree.selection()
if not selected:
return

item = tree.item(selected[0])['values']
if item:
data_str = "\t".join(str(x) for x in item)
win.clipboard_clear()
win.clipboard_append(data_str)

tree.bind("
tree.bind("

# Summary frame
summary_frame = ttk.Frame(main_frame)
summary_frame.grid(row=3, column=0, sticky="ew", pady=10)

total_label = ttk.Label(summary_frame, text="Total Penjualan: Rp 0",
font=('Arial', 12, 'bold'),
foreground=theme["highlight"])
total_label.pack(side='left', padx=10)

laba_label = ttk.Label(summary_frame, text="Total Laba: Rp 0",
font=('Arial', 12, 'bold'),
foreground=theme["success"])
laba_label.pack(side='left', padx=10)

item_label = ttk.Label(summary_frame, text="Total Item Terjual: 0",
font=('Arial', 12))
item_label.pack(side='right', padx=10)

# Load initial data
refresh_laporan()

# ================== WINDOW ABOUT ==================
def about_window(parent):
theme = theme_mgr.get_theme()

win = tk.Toplevel(parent)
win.title("Tentang Aplikasi")
win.geometry("400x300")
win.configure(bg=theme["bg"])
win.transient(parent)
win.resizable(False, False)
bind_fullscreen_keys(win)

# Buat agar window berada di tengah layar
win.update_idletasks()
width = 400
height = 300
x = (win.winfo_screenwidth() // 2) - (width // 2)
y = (win.winfo_screenheight() // 2) - (height // 2)
win.geometry(f"{width}x{height}+{x}+{y}")

# Style
style = ttk.Style(win)
style.theme_use('clam')
style.configure('.', background=theme["bg"], foreground=theme["fg"])
style.configure('TLabel', background=theme["bg"], foreground=theme["fg"])

main_frame = ttk.Frame(win, padding=30)
main_frame.pack(fill='both', expand=True)

ttk.Label(main_frame, text="APLIKASI KASIR",
font=('Arial', 18, 'bold'),
foreground=theme["accent"]).pack(pady=10)

ttk.Label(main_frame, text="Versi 2.0",
font=('Arial', 12)).pack(pady=5)

ttk.Label(main_frame, text="Aplikasi kasir dengan fitur lengkap\nuntuk mengelola transaksi dan stok barang",
font=('Arial', 10),
justify='center').pack(pady=10)

ttk.Label(main_frame, text="Fitur:",
font=('Arial', 11, 'bold')).pack(pady=(20,5))

features = [
"✓ Transaksi penjualan",
"✓ Manajemen stok barang",
"✓ Laporan keuangan",
"✓ Multiple tema",
"✓ Responsive design",
"✓ Multi-printer support",
"✓ Generator barcode",
"✓ Preview & Print Laporan"
]

for feature in features:
ttk.Label(main_frame, text=feature,
font=('Arial', 9)).pack()

ttk.Button(main_frame, text="TUTUP",
command=win.destroy).pack(pady=20)

# ================== INISIALISASI APLIKASI ==================
if __name__ == "__main__":
# Initialize files if not exists
if not os.path.exists(TRANSAKSI_FILE):
with open(TRANSAKSI_FILE, 'w', newline='', encoding='utf-8') as f:
pass

if not os.path.exists(STOK_FILE):
stok_contoh = {
"BRG001": {"nama": "Oli Mesin 1L", "harga_beli": 25000, "harga_jual": 35000, "qty": 10},
"BRG002": {"nama": "Ban Dalam", "harga_beli": 45000, "harga_jual": 60000, "qty": 5},
"BRG003": {"nama": "Kampas Rem", "harga_beli": 30000, "harga_jual": 45000, "qty": 8},
"BRG004": {"nama": "Aki Motor", "harga_beli": 150000, "harga_jual": 200000, "qty": 3},
"BRG005": {"nama": "Busi", "harga_beli": 8000, "harga_jual": 12000, "qty": 15},
}
save_stok(stok_contoh)
print("File stok contoh berhasil dibuat!")

# Buat file settings jika belum ada
if not os.path.exists(SETTINGS_FILE):
save_settings({
"toko_nama": TOKO_NAMA,
"toko_alamat": TOKO_ALAMAT,
"theme": "dark",
"receipt_printer": printer_mgr.default_printer,
"report_printer": printer_mgr.default_printer,
"label_printer": printer_mgr.default_printer,
"encoding": "cp850"
})

print("Memulai aplikasi kasir...")
print(f"Tema aktif: {theme_mgr.get_theme()['name']}")
login_window()
