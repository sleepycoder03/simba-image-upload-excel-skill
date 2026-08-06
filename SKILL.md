---
name: simba-image-upload-excel
description: |
  将 Excel 中的图片（URL 或本地路径）批量上传到 Simba 后台，并把 Simba 生成的图片链接和素材 key 回填到表格。
  支持两种模式：按“图片来源列”直接上传，或“上传文件夹图片并按文件名匹配回填”。
  folder-match 模式下如不提供输入Excel，会自动在图片文件夹目录生成并回填表格。
  每次执行前必须向用户索取当次 token，且不得复用旧 token。
---

# Simba 图片上传并回填 Excel

## 何时使用
- 用户要把一批图片上传到 Simba 素材后台。
- 用户要把 Simba 生成的 URL 批量写回 Excel。
- 支持两类输入：
  1. 表格里已有图片 URL/本地路径（`source-col`）
  2. 图片在本地文件夹里，需按文件名匹配行再回填（`folder-match`）

## 强制安全规则
1. **每次运行前必须向用户索取 token**（Bearer 后面的字符串）。
2. **不要默认复用历史 token**，除非用户在当前轮次明确再次提供。
3. **不要把 token 写入文件、脚本、日志或最终回复**。

## 默认执行步骤
1. 确认本次参数（输入文件、sheet、回填列、模式）。
2. 让用户提供 **本次 token**。
3. 运行脚本：`scripts/simba_excel_upload.py`。
4. 返回输出文件路径和日志路径。

## 命令模板
使用 Codex runtime Python：

### A) 按图片来源列上传（source-col）
```bash
/Users/macbook2/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 \
  /Users/macbook2/.codex/skills/simba-image-upload-excel/scripts/simba_excel_upload.py \
  --mode source-col \
  --input "<输入xlsx绝对路径>" \
  --sheet "替代游戏总表" \
  --source-col "缩略图2(512x512)" \
  --url-col "Simba生成链接" \
  --key-col "Simba素材Key" \
  --remark-col "替代游戏"
```

### B) 上传文件夹并按文件名匹配回填（folder-match，已有Excel）
```bash
/Users/macbook2/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 \
  /Users/macbook2/.codex/skills/simba-image-upload-excel/scripts/simba_excel_upload.py \
  --mode folder-match \
  --input "<输入xlsx绝对路径>" \
  --sheet "替代游戏总表" \
  --image-dir "<图片文件夹绝对路径>" \
  --match-col "替代游戏" \
  --rank-col "综合排名" \
  --url-col "Simba生成链接" \
  --key-col "Simba素材Key" \
  --remark-col "替代游戏"
```

### C) 上传文件夹并自动生成/回填表格（folder-match，无输入Excel）
```bash
/Users/macbook2/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 \
  /Users/macbook2/.codex/skills/simba-image-upload-excel/scripts/simba_excel_upload.py \
  --mode folder-match \
  --image-dir "<图片文件夹绝对路径>" \
  --sheet "替代游戏总表"
```

> C模式会自动在**图片文件夹目录**生成输出Excel，并把 Simba 链接回填进去。

> 不传 `--token` 时脚本会交互提示输入 token（推荐）。

## 参数说明（重点）
- `--mode`: `source-col` / `folder-match`
- `--image-dir`: folder-match 必填
- `--match-col`: 用该列值与文件名匹配
- `--rank-col`: 优先按该列数值匹配文件名前缀（如 `01_xxx.jpg`）
- `--allow-fuzzy`: 开启模糊匹配兜底（默认关闭）
- `--overwrite`: 覆盖已有 Simba 链接
- `--max-upload 10`: 仅处理前 10 条（调试）
- `--log "<日志路径>"`: 自定义日志路径

## 输出
- 新 Excel：回填 `Simba生成链接` 与 `Simba素材Key`。
- 日志：成功/失败/跳过统计与失败明细。
