# Arope-GIS-Toolkit

一套用于 GIS 批量处理的小工具集合。  
包含 **批量坐标转换器**（Excel→Excel）和 **Excel→矢量导出工具**（Excel→Shapefile/GeoJSON），以及未来的更多批量处理脚本。

---

## ✨ 功能特色

### 🔹 批量坐标转换（Excel → Excel）
- 支持常见坐标系：  
  - WGS84 (EPSG:4326)  
  - Web Mercator (EPSG:3857)  
  - CGCS2000 (EPSG:4490)  
  - GCJ-02（国测火星坐标）  
  - BD-09（百度坐标）  
- 支持自定义 EPSG 编码  
- 自动生成新 Excel，新增 `X_out` / `Y_out` 列  

### 🔹 Excel → Shapefile / GeoJSON
- 支持选择 Sheet、X 列 / Y 列  
- 支持指定 EPSG 坐标系  
- 导出为 ESRI Shapefile (.shp) 或 GeoJSON (.geojson)  

### 🔹 其他特点
- 一键打包成 Windows EXE  
- **无需安装 Python** 即可运行  

---
### 🔹 后续继续追加功能
