import os, math
import pandas as pd
from pyproj import CRS, Transformer, datadir as pyproj_datadir

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText

# 让打包后的 EXE 能找到 PROJ 数据
os.environ.setdefault("PROJ_LIB", pyproj_datadir.get_data_dir())

PRESETS = [
    ("WGS84 (EPSG:4326)", "EPSG:4326"),
    ("Web Mercator (EPSG:3857)", "EPSG:3857"),
    ("CGCS2000 (EPSG:4490)", "EPSG:4490"),
    ("GCJ-02 (中国火星坐标)", "GCJ-02"),
    ("BD-09 (百度)", "BD-09"),
]

# ---------- GCJ-02 / BD-09 工具 ----------
PI = 3.1415926535897932384626
A = 6378245.0
EE = 0.00669342162296594323

def _out_of_china(lon, lat):
    return not (73.66 <= lon <= 135.05 and 3.86 <= lat <= 53.55)

def _transform_lat(lon, lat):
    ret = -100.0 + 2.0*lon + 3.0*lat + 0.2*lat*lat + 0.1*lon*lat + 0.2*math.sqrt(abs(lon))
    ret += (20.0*math.sin(6.0*lon*PI) + 20.0*math.sin(2.0*lon*PI))*2.0/3.0
    ret += (20.0*math.sin(lat*PI) + 40.0*math.sin(lat/3.0*PI))*2.0/3.0
    ret += (160.0*math.sin(lat/12.0*PI) + 320.0*math.sin(lat*PI/30.0))*2.0/3.0
    return ret

def _transform_lon(lon, lat):
    ret = 300.0 + lon + 2.0*lat + 0.1*lon*lon + 0.1*lon*lat + 0.1*math.sqrt(abs(lon))
    ret += (20.0*math.sin(6.0*lon*PI) + 20.0*math.sin(2.0*lon*PI))*2.0/3.0
    ret += (20.0*math.sin(lon*PI) + 40.0*math.sin(lon/3.0*PI))*2.0/3.0
    ret += (150.0*math.sin(lon/12.0*PI) + 300.0*math.sin(lon/30.0*PI))*2.0/3.0
    return ret

def wgs84_to_gcj02(lon, lat):
    if _out_of_china(lon, lat):
        return lon, lat
    dlat = _transform_lat(lon - 105.0, lat - 35.0)
    dlon = _transform_lon(lon - 105.0, lat - 35.0)
    radlat = lat / 180.0 * PI
    magic = math.sin(radlat)
    magic = 1 - EE * magic * magic
    sqrtmagic = math.sqrt(magic)
    dlat = (dlat * 180.0) / ((A * (1 - EE)) / (magic * sqrtmagic) * PI)
    dlon = (dlon * 180.0) / (A / sqrtmagic * math.cos(radlat) * PI)
    return lon + dlon, lat + dlat

def gcj02_to_wgs84(lon, lat):
    if _out_of_china(lon, lat):
        return lon, lat
    tol = 1e-7
    mlon, mlat = lon, lat
    plon, plat = lon - 0.5, lat - 0.5
    clon, clat = lon + 0.5, lat + 0.5
    for _ in range(30):
        wlon, wlat = (plon + clon) / 2, (plat + clat) / 2
        glon, glat = wgs84_to_gcj02(wlon, wlat)
        dlon, dlat = glon - lon, glat - lat
        if abs(dlon) < tol and abs(dlat) < tol:
            return wlon, wlat
        if dlon > 0: clon = wlon
        else:        plon = wlon
        if dlat > 0: clat = wlat
        else:        plat = wlat
    return wlon, wlat

def gcj02_to_bd09(lon, lat):
    z = math.sqrt(lon*lon + lat*lat) + 0.00002 * math.sin(lat * PI)
    theta = math.atan2(lat, lon) + 0.000003 * math.cos(lon * PI)
    return z * math.cos(theta) + 0.0065, z * math.sin(theta) + 0.006

def bd09_to_gcj02(lon, lat):
    x = lon - 0.0065
    y = lat - 0.006
    z = math.sqrt(x*x + y*y) - 0.00002 * math.sin(y * PI)
    theta = math.atan2(y, x) - 0.000003 * math.cos(x * PI)
    return z * math.cos(theta), z * math.sin(theta)

def wgs84_to_bd09(lon, lat):
    glon, glat = wgs84_to_gcj02(lon, lat)
    return gcj02_to_bd09(glon, glat)

def bd09_to_wgs84(lon, lat):
    glon, glat = bd09_to_gcj02(lon, lat)
    return gcj02_to_wgs84(glon, glat)

# ---------- transform pipeline ----------
def make_proj_transform(src_spec: str, dst_spec: str):
    s = src_spec.upper().strip()
    d = dst_spec.upper().strip()
    def is_wgs84(tag): return tag in ("EPSG:4326", "WGS84")

    if s.startswith("EPSG:") and d.startswith("EPSG:"):
        tr = Transformer.from_crs(CRS.from_user_input(s), CRS.from_user_input(d), always_xy=True)
        return lambda x, y: tr.transform(x, y)

    def to_wgs84(x, y):
        if s == "GCJ-02":   return gcj02_to_wgs84(x, y)
        if s == "BD-09":    return bd09_to_wgs84(x, y)
        if is_wgs84(s):     return x, y
        t = Transformer.from_crs(CRS.from_user_input(s), CRS.from_user_input("EPSG:4326"), always_xy=True)
        return t.transform(x, y)

    def from_wgs84(x, y):
        if d == "GCJ-02":   return wgs84_to_gcj02(x, y)
        if d == "BD-09":    return wgs84_to_bd09(x, y)
        if is_wgs84(d):     return x, y
        t = Transformer.from_crs(CRS.from_user_input("EPSG:4326"), CRS.from_user_input(d), always_xy=True)
        return t.transform(x, y)

    return lambda x, y: from_wgs84(*to_wgs84(x, y))

def resolve_crs(preset_name: str, epsg_text: str) -> str:
    epsg_text = (epsg_text or "").strip()
    if epsg_text:
        return epsg_text if epsg_text.upper().startswith("EPSG:") else f"EPSG:{epsg_text}"
    for n,c in PRESETS:
        if n == preset_name: return c
    return "EPSG:4326"

# ---------- Tk GUI ----------
class ConverterApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("坐标批量转换器By Arope_V1")
        self.geometry("820x520")

        # 文件与sheet
        frm1 = ttk.Frame(self); frm1.pack(fill="x", padx=8, pady=6)
        ttk.Label(frm1, text="Excel 文件").pack(side="left")
        self.var_file = tk.StringVar()
        ttk.Entry(frm1, textvariable=self.var_file, width=70).pack(side="left", padx=6)
        ttk.Button(frm1, text="浏览", command=self.pick_file).pack(side="left")

        frm2 = ttk.Frame(self); frm2.pack(fill="x", padx=8, pady=6)
        ttk.Label(frm2, text="Sheet").pack(side="left")
        self.cmb_sheet = ttk.Combobox(frm2, width=30, state="readonly")
        self.cmb_sheet.pack(side="left", padx=6)
        self.cmb_sheet.bind("<<ComboboxSelected>>", self.load_sheet_preview)

        # 列选择
        frm3 = ttk.Frame(self); frm3.pack(fill="x", padx=8, pady=6)
        ttk.Label(frm3, text="X 列").pack(side="left")
        self.cmb_x = ttk.Combobox(frm3, width=25, state="readonly"); self.cmb_x.pack(side="left", padx=6)
        ttk.Label(frm3, text="Y 列").pack(side="left")
        self.cmb_y = ttk.Combobox(frm3, width=25, state="readonly"); self.cmb_y.pack(side="left", padx=6)

        # 坐标系
        labf1 = ttk.Labelframe(self, text="输入坐标系"); labf1.pack(fill="x", padx=8, pady=6)
        self.cmb_src = ttk.Combobox(labf1, values=[n for n,_ in PRESETS], width=35, state="readonly")
        self.cmb_src.set(PRESETS[0][0]); self.cmb_src.pack(side="left", padx=6, pady=6)
        ttk.Label(labf1, text="自定义 EPSG(选填)").pack(side="left")
        self.var_src_epsg = tk.StringVar(); ttk.Entry(labf1, textvariable=self.var_src_epsg, width=12).pack(side="left", padx=6)

        labf2 = ttk.Labelframe(self, text="输出坐标系"); labf2.pack(fill="x", padx=8, pady=6)
        self.cmb_dst = ttk.Combobox(labf2, values=[n for n,_ in PRESETS], width=35, state="readonly")
        self.cmb_dst.set(PRESETS[1][0]); self.cmb_dst.pack(side="left", padx=6, pady=6)
        ttk.Label(labf2, text="自定义 EPSG(选填)").pack(side="left")
        self.var_dst_epsg = tk.StringVar(); ttk.Entry(labf2, textvariable=self.var_dst_epsg, width=12).pack(side="left", padx=6)

        # 输出文件
        frm4 = ttk.Frame(self); frm4.pack(fill="x", padx=8, pady=6)
        ttk.Label(frm4, text="输出文件").pack(side="left")
        self.var_out = tk.StringVar(); ttk.Entry(frm4, textvariable=self.var_out, width=70).pack(side="left", padx=6)
        ttk.Button(frm4, text="选择", command=self.pick_save).pack(side="left")

        # 按钮 & 日志
        frm5 = ttk.Frame(self); frm5.pack(fill="x", padx=8, pady=6)
        ttk.Button(frm5, text="开始转换", command=self.run).pack(side="left")
        ttk.Button(frm5, text="预览前5行", command=self.preview).pack(side="left", padx=6)

        self.log = ScrolledText(self, height=14); self.log.pack(fill="both", expand=True, padx=8, pady=6)
        ttk.Label(self, text="© 2025 by Arope", anchor="center").pack(side="bottom", pady=4)

        self.df_preview = None

    def log_print(self, *msg):
        self.log.insert("end", " ".join(map(str,msg)) + "\n")
        self.log.see("end")

    def pick_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel","*.xlsx;*.xls")])
        if not path: return
        self.var_file.set(path)
        try:
            xls = pd.ExcelFile(path)
            self.cmb_sheet["values"] = xls.sheet_names
            if xls.sheet_names:
                self.cmb_sheet.set(xls.sheet_names[0])
                self.load_sheet_preview()
            base, _ = os.path.splitext(path)
            if not self.var_out.get():
                self.var_out.set(base + "_converted.xlsx")
            self.log_print(f"已加载：{os.path.basename(path)}；sheets={xls.sheet_names}")
        except Exception as e:
            messagebox.showerror("错误", f"读取Excel失败：{e}")

    def load_sheet_preview(self, *_):
        path = self.var_file.get()
        sheet = self.cmb_sheet.get()
        if not (path and sheet): return
        df = pd.read_excel(path, sheet_name=sheet, nrows=200)
        self.df_preview = df
        cols = list(df.columns)
        self.cmb_x["values"] = cols
        self.cmb_y["values"] = cols
        self.log_print(f"切换到 sheet={sheet}；列={cols}")

    def preview(self):
        if self.df_preview is None:
            self.log_print("请先选择 Excel 文件。"); return
        self.log_print(self.df_preview.head().to_string())

    def pick_save(self):
        path = filedialog.asksaveasfilename(defaultextension=".xlsx",filetypes=[("Excel","*.xlsx")])
        if path: self.var_out.set(path)

    def run(self):
        try:
            path = self.var_file.get()
            sheet = self.cmb_sheet.get()
            xcol = self.cmb_x.get()
            ycol = self.cmb_y.get()
            out_path = self.var_out.get()

            if not (path and os.path.exists(path)):
                self.log_print("[错误] 请选择有效的 Excel 文件。"); return
            if not sheet:
                self.log_print("[错误] 请选择 Sheet。"); return
            if not (xcol and ycol):
                self.log_print("[错误] 请选择 X/Y 列。"); return
            if not out_path:
                base,_ = os.path.splitext(path); out_path = base + "_converted.xlsx"

            src = resolve_crs(self.cmb_src.get(), self.var_src_epsg.get())
            dst = resolve_crs(self.cmb_dst.get(), self.var_dst_epsg.get())
            self.log_print(f"坐标系：{src} -> {dst}")

            df = pd.read_excel(path, sheet_name=sheet)
            if xcol not in df.columns or ycol not in df.columns:
                self.log_print("[错误] X/Y 列不存在于表头。"); return

            x = pd.to_numeric(df[xcol], errors="coerce")
            y = pd.to_numeric(df[ycol], errors="coerce")
            n_total = len(df)
            n_nan = int(x.isna().sum() + y.isna().sum())
            if n_nan > 0:
                self.log_print(f"[提示] 有 {n_nan} 个值无法转换为数值，已按 NaN 处理。")

            f = make_proj_transform(src, dst)
            out_xy = [f(xv, yv) if pd.notna(xv) and pd.notna(yv) else (None, None) for xv,yv in zip(x,y)]
            df["X_out"] = [p[0] for p in out_xy]
            df["Y_out"] = [p[1] for p in out_xy]

            df.to_excel(out_path, index=False)
            self.log_print(f"✅ 完成：共 {n_total} 行 → 保存到：{out_path}")
            messagebox.showinfo("成功", "转换完成！")
        except Exception as e:
            self.log_print("[错误]", e)
            messagebox.showerror("错误", str(e))

if __name__ == "__main__":
    ConverterApp().mainloop()

