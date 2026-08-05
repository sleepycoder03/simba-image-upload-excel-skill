# Quickstart

## 1) 准备
- Excel 表。
- 本次操作者提供 **当次 token**。

## 2) 运行模式

### 模式A：source-col（从表格列读取图片）
```bash
/Users/macbook2/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 \
  /Users/macbook2/.codex/skills/simba-image-upload-excel/scripts/simba_excel_upload.py \
  --mode source-col \
  --input "/abs/path/input.xlsx"
```

### 模式B：folder-match（上传文件夹并按文件名匹配）
```bash
/Users/macbook2/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 \
  /Users/macbook2/.codex/skills/simba-image-upload-excel/scripts/simba_excel_upload.py \
  --mode folder-match \
  --input "/abs/path/input.xlsx" \
  --image-dir "/abs/path/images" \
  --match-col "替代游戏" \
  --rank-col "综合排名"
```

> 不传 `--token` 会交互输入 token（推荐，避免命令行明文泄露）。

## 3) 查看结果
- 默认输出：`input_simba已回填.xlsx`
- 默认日志：`input_simba已回填_upload日志.txt`
