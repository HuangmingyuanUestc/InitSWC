# SWC初始化工具
> 作者：黄明源

# 1 使用方法
1. 下载：[wkhtmltopdf](https://wkhtmltopdf.org/downloads.html) ，或在`asset`文件夹下载win64安装包
2. 配置`main.py`
```python
# 配置：
# 1. 指定人名（要在表内）：
TARGET_HUMAN_NAME = '黄明源'
# 2. 下载：https://wkhtmltopdf.org/downloads.html，并配置wkhtmltopdf的下载路径
WKHTML_PATH = r'C:\Develop\Tools\wkhtmltopdf\bin\wkhtmltopdf.exe'
# 3. 配置是否下载报告和代码，以及是否需要初始化分析表
NEED_DOWNLOAD_REPORT = True
NEED_DOWNLOAD_CODE = True
NEED_DOWNLOAD_SHEET = True
```
3. 下载依赖，运行`main.py`，从`output`文件夹获取结果

# 2 功能列表
## 显性需求功能
- [x] 下载代码
- [x] 下载报告
- [x] 复制示例表格到指定位置
- [ ] 初始化示例表格的finding内容
## 隐性需求
- [x] 显性需求功能的开关配置
- [x] 下载报告
- [x] 包容Exception
