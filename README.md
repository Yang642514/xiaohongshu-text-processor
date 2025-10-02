# 小红书文案转搞定模板工具

该工具用于：
- 输入一段小红书文案
- 按模板拆分提取封面标题、分论点及其内容
- 进行中文标点规范化与排版（每段不超过3行、每行约15字符）
- 将内容写入指定 Excel 套版文件（不改变格式）
- 输出命名为 `小红书标题_时间戳.xlsx` 的文件，并打包为 `小红书标题_时间戳.zip`

## 功能概览
- 文案解析：输入的文案按分论点标题、分论点内容拆分，支持中文序号（一、二、三、）、数字序号（1.、1、）、破折/点列举（-、• 等）识别分论点。
- 标点规范：将英文标点替换为中文标点（，。！？：；“”‘’等）。
- 排版规则：每段最多3行，每行约15个中文字符，自动折行。
- Excel 写入：基于模板首行表头匹配写入目标列，不修改样式。
- 压缩打包：将生成的 Excel 与相关资源一起打包压缩。
- 日志与错误处理：记录运行日志，错误弹窗提示。
- 进度条：展示处理进度。
- 设置界面：可配置模板路径、输出路径、列映射等。
- 关于界面：展示版本、开发者、使用说明。

## 项目结构
```
app/
  core/
    excel_writer.py
    logger.py
    parser.py
    punctuation.py
    utils.py
    zipper.py
  gui/
    about_dialog.py
    main_window.py
    settings_dialog.py
  config/
    default_settings.json
main.py
README.md
requirements.txt
environment.yml
build.ps1
```

## 环境与依赖
建议使用 Conda。

### 创建虚拟环境
```
conda env create -f environment.yml
conda activate xhs-tool
```

或使用 pip：
```
python -m venv .venv
.\.venv\Scripts\activate
pip install -r requirements.txt
```

## 运行
```
python main.py
```

首次运行请在设置界面确认：
- Excel 模板路径（默认读取工作目录下的 `稿定设计-数据上传.xlsx`）
- 输出目录与压缩包目录
- 列映射（匹配模板首行的表头名称）

## 打包为 EXE（PyInstaller）
确保已激活虚拟环境并安装依赖：
```
conda activate xhs-tool
pyinstaller --noconfirm --windowed --name 小红书文案助手 \
  --add-data "稿定设计-数据上传.xlsx;." \
  --add-data "app/config/default_settings.json;app/config" \
  main.py
```

打包完成后的可执行文件位于 `dist/小红书文案助手.exe`。

也可使用提供的脚本：
```
PowerShell ./build.ps1
```

## 使用说明（简要）
- 在主界面文本框粘贴小红书文案，点击“开始处理”。
- 程序将解析标题与分论点，进行标点与排版处理。
- 生成 `标题_时间戳.xlsx`，并在输出目录内打包为 `标题_时间戳.zip`。
- 若模板表头与默认映射不一致，请在设置界面调整“列映射”。

## 注意事项
- 不要修改模板 Excel 的表头、合并单元格等，以免导入失败。
- 图片需与 Excel 同目录并在 Excel 中填写文件名（无后缀），打包时将一并压缩（后续可扩展自动配图）。
- 当前版本仅处理单条文案并写入一行数据；可在后续版本扩展批量处理。

## 后续可扩展方向
- 自动配图：根据分论点内容调用 API 获取图片并插入 Excel。
- 多模板支持：不同套版模板的列映射预设与切换。
- 批量导入：支持一次性处理多条文案形成多行数据。