# Docx -> IDML 工具（中文版说明）

这是一个将 DOCX 转换为 XML / 生成 InDesign JSX 并可（可选）自动导出 IDML 的工具集。

快速上手（Windows PowerShell）：

1. 创建并激活虚拟环境（需 Python 3.10+）：

```powershell
python -m venv .venv
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force; .\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
```

2. 安装依赖：

```powershell
python -m pip install -r .\requirements.txt
# 如果你在 macOS 或不需要 pywin32，可只安装必要的包：
python -m pip install python-docx lxml
```

3. （可选）安装并启用 pre-commit：

```powershell
python -m pip install pre-commit
pre-commit install
pre-commit run --all-files
```

4. 运行脚本（示例）：

```powershell
python .\docx_to_idml.py --password "YourPassword" .\sample.docx
```

仓库包含：
- `docx_to_idml.py`：主入口脚本
- `xml_to_idml.py`：JSX 生成与 InDesign 调用逻辑
- `docx_to_xml_outline_notes_v13.py`：DOCX -> XML 导出器

CI：已添加 GitHub Actions 工作流（`.github/workflows/ci.yml`），会在 push/PR 时运行 lint（black/isort/flake8）、pre-commit 和 pytest。

备注：请根据需要修改 `.pre-commit-config.yaml` 中的工具版本和 `ci.yml` 的 Python 版本矩阵。
