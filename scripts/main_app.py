import tkinter as tk
from tkinter import ttk
from app_tk import ConverterApp
from excel_to_vector_tk import VectorApp

class MainApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Arope GIS Toolkit")
        self.geometry("420x220")

        ttk.Label(self, text="请选择工具", font=("Arial", 16)).pack(pady=20)

        ttk.Button(
            self, text="批量坐标转换器 (Excel → Excel)",
            command=self.open_converter, width=35
        ).pack(pady=10)

        ttk.Button(
            self, text="Excel → Shapefile / GeoJSON",
            command=self.open_vector, width=35
        ).pack(pady=10)

        ttk.Label(self, text="© 2025 by Arope", anchor="center").pack(side="bottom", pady=8)

    def open_converter(self):
        # 方案A：如果 ConverterApp 继承自 tk.Tk（你没改成 Toplevel）
        ConverterApp()
        # 如果你按我之前的“方案B”把它改成 tk.Toplevel，就用：
        # ConverterApp(self)

    def open_vector(self):
        # 同上，按你的类基类选择其一
        VectorApp()
        # VectorApp(self)

if __name__ == "__main__":
    MainApp().mainloop()
