# simba-image-upload-excel

Codex Skill：批量上传图片到 Simba，并将生成链接/素材 key 回填到 Excel。

## 功能
- `source-col` 模式：按表格图片来源列上传。
- `folder-match` 模式：上传文件夹图片并按文件名匹配回填。
- 每次运行前要求用户提供当次 token（不复用旧 token）。

## 目录
- `SKILL.md`
- `scripts/simba_excel_upload.py`
- `references/quickstart.md`
- `agents/openai.yaml`
