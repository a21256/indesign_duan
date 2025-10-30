# 开发与提交流程说明

这个仓库现在只在版本控制中跟踪三份主脚本：

 - `docx_to_idml.py`
 - `xml_to_idml.py`
 - `docx_to_xml_outline_notes_v13.py`

其它文件（备份、.docx、图片、临时生成的 idml/xml 等）保留在本地磁盘并被 `.gitignore` 忽略，不会出现在仓库或远端。

推荐的日常工作流程（简单、适用于单人或小团队）

1. 获取并同步远端最新改动（始终在开始前执行）：

```powershell
git fetch origin
git pull --rebase origin main
```

2. 创建功能分支（建议）：

```powershell
git checkout -b feat/your-description
```

3. 本地修改并检查变更：

```powershell
git status
git diff
git add <file>
git commit -m "feat: 简短描述你的变更"
```

4. 将本地分支推送到远端并在 GitHub 提交 Pull Request（做代码审查与 CI）：

```powershell
git push -u origin feat/your-description
```

5. 合并后，回到 `main` 并拉取最新：

```powershell
git checkout main
git pull --rebase origin main
```

常用命令小贴士：
- 查看被跟踪文件： `git ls-files`
- 查看工作树是否干净： `git status --porcelain`（无输出表示干净）
- 如果要把某些已跟踪文件停止跟踪但保留本地，使用：
	`git rm --cached <path>` 然后 `git commit`。

如果你希望我把这个 README 做进一步扩展（例如把 CI、pre-commit 和运行测试的步骤加进去），告诉我我会补充。
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
