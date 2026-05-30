# 亚马逊广告 Bulk 操作生成器

这是一个本地运行的 Streamlit 工具，用于把中文广告操作需求转换成 Amazon Sponsored Products Bulk 上传表。

工具流程：

1. 上传原始 Amazon Sponsored Products Bulk xlsx 文件。
2. 在输入框填写一个或多个广告操作。
3. 系统解析为标准操作单并做预检查。
4. 生成可上传到 Amazon 的 `bulk_upload.xlsx`。

## 项目路径

当前本地路径：

```text
C:\Users\Admin\Documents\Playground
```

## 主要文件

```text
.
├── app.py
├── bulk_generator.py
├── requirements.txt
├── README.md
└── .gitignore
```

说明：

- `app.py`：Streamlit 页面入口。
- `bulk_generator.py`：Bulk 文件读取、需求解析、预检查和上传表生成逻辑。
- `requirements.txt`：运行依赖。
- `.gitignore`：排除本地输出文件、缓存和临时文件。

## 在新电脑运行

先进入项目目录：

```powershell
cd C:\Users\Admin\Documents\Playground
```

建议创建虚拟环境：

```powershell
python -m venv .venv
.venv\Scripts\activate
```

安装依赖：

```powershell
pip install -r requirements.txt
```

启动工具：

```powershell
python -m streamlit run app.py --server.port 8501
```

启动后在浏览器打开：

```text
http://localhost:8501
```

## 使用提示

- 上传的 Bulk 文件必须是 Amazon Sponsored Products Bulk xlsx。
- 多个操作可以用 `---` 分隔。
- 页面会先显示每个操作块的预检查结果。
- 有些操作块失败时，其他成功操作块仍可合并生成一个 `bulk_upload.xlsx`。
- 生成的 xlsx/csv 文件属于本地输出文件，不建议上传到 GitHub。
