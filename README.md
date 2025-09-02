# Arope-GIS-Toolkit

一套用于 GIS 批量处理的小工具集合。  
包含 **坐标批量转换器**（支持 Excel 表格输入/输出，支持 WGS84、CGCS2000、Web Mercator、GCJ-02、BD-09 等常见坐标系），以及未来的更多批量处理脚本。

---

## 功能特色
- Excel 表格批量坐标转换  
- 支持选择输入/输出坐标系：  
  - WGS84 (EPSG:4326)  
  - Web Mercator (EPSG:3857)  
  - CGCS2000 (EPSG:4490)  
  - GCJ-02（国测火星坐标）  
  - BD-09（百度坐标）  
  - 其他 EPSG 编码（手动输入）  
- 自动生成新 Excel，新增 `X_out` / `Y_out` 列  
- 一键打包成 Windows EXE，**无需安装 Python**  

---

## 使用方法

### 方式一：直接运行 EXE
1. 前往 [Releases](https://github.com/AxiArope/Arope-GIS-Toolkit/releases) 页面下载最新的  
   `CoordinateBatchConverter_by_Arope.exe`  
2. 双击运行  
3. 按界面操作：选择 Excel → Sheet → X/Y 列 → 输入/输出坐标系 → 输出路径  
4. 点击 **开始转换**，等待完成  

### 方式二：运行源码
如果你电脑里有 Python 3.10+：  

```bash
git clone https://github.com/AxiArope/Arope-GIS-Toolkit.git
cd Arope-GIS-Toolkit/scripts
pip install -r requirements.txt
python app_tk.py
