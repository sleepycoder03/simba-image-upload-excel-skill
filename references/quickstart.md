# Quickstart

## 1) 准备
- 本次操作者提供 **当次 token**。
- 有两种输入方式：
  - 已有Excel（按列上传 or 文件夹匹配回填）
  - 仅图片文件夹（自动生成并回填Excel）

## 2) 运行模式

### 模式A：source-col（从表格列读取图片）
```bash
/Users/macbook2/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 \
  /Users/macbook2/.codex/skills/simba-image-upload-excel/scripts/simba_excel_upload.py \
  --mode source-col \
  --input "/abs/path/input.xlsx"
```

### 模式B：folder-match（已有Excel，按文件夹匹配）
```bash
/Users/macbook2/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 \
  /Users/macbook2/.codex/skills/simba-image-upload-excel/scripts/simba_excel_upload.py \
  --mode folder-match \
  --input "/abs/path/input.xlsx" \
  --image-dir "/abs/path/images" \
  --match-col "替代游戏" \
  --rank-col "综合排名"
```

### 模式C：folder-match（仅图片文件夹，自动生成回填表）
```bash
/Users/macbook2/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 \
  /Users/macbook2/.codex/skills/simba-image-upload-excel/scripts/simba_excel_upload.py \
  --mode folder-match \
  --image-dir "/abs/path/images"
```

> 模式C会在图片目录下自动生成并回填Excel。

> 不传 `--token` 会交互输入 token（推荐，避免命令行明文泄露）。

## 3) 查看结果
- 默认输出：
  - source-col：`input_simba已回填.xlsx`
  - folder-match：输出放在 `image-dir` 目录
- 默认日志：`输出文件同名_upload日志.txt`
