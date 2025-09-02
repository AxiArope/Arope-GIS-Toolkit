Arope-GIS-Toolkit

一套用于 GIS 批量处理的小工具集合。
包含 批量坐标转换器（Excel→Excel）和 Excel→矢量导出工具（Excel→Shapefile/GeoJSON），以及未来的更多批量处理脚本。

功能特色

批量坐标转换（Excel→Excel）

支持 WGS84、CGCS2000、Web Mercator、GCJ-02、BD-09 等常见坐标系

支持其他 EPSG 编码（手动输入）

自动生成新 Excel，新增 X_out / Y_out 列

Excel → Shapefile / GeoJSON

支持选择 Sheet、X 列 / Y 列

支持指定 EPSG 坐标系

导出为 ESRI Shapefile (.shp) 或 GeoJSON (.geojson)

一键打包成 Windows EXE，无需安装 Python

使用方法
方式一：直接运行 EXE

前往 Releases
 页面下载最新的

AropeGIS_Toolkit.exe（主菜单版，集成所有工具）

或单独的 CoordinateBatchConverter_by_Arope.exe / ExcelToVector_by_Arope.exe

双击运行

按界面操作：

批量坐标转换器：选择 Excel → Sheet → X/Y 列 → 输入/输出坐标系 → 输出路径

Excel → Shapefile/GeoJSON：选择 Excel → Sheet → X/Y 列 → EPSG → 输出路径

点击 开始，等待完成

方式二：运行源码

如果你电脑里有 Python 3.10+：

git clone https://github.com/AxiArope/Arope-GIS-Toolkit.git
cd Arope-GIS-Toolkit/scripts
pip install -r ../requirements.txt
python main_app.py
