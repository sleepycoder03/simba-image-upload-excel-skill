#!/usr/bin/env python3
"""
Upload image assets to Simba material backend and backfill generated image URLs into Excel.

Security note:
- This script NEVER persists user token to disk.
- If --token is not provided, script prompts interactively each run.
"""

from __future__ import annotations

import argparse
import datetime as dt
import getpass
import importlib
import json
import mimetypes
import re
import subprocess
import sys
import traceback
import urllib.error
import urllib.parse
import urllib.request
import uuid
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple


API_BASE = "https://cms-api.aoneroom.com"
STS_ENDPOINT = "/wefeed-cms-bff/upload/sts-token"
GENERATE_ENDPOINT = "/wefeed-cms-bff/material/image/upload"
IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".webp", ".gif", ".bmp", ".tif", ".tiff"}


class SimbaError(RuntimeError):
    pass


def ensure_import(module_name: str, pip_name: Optional[str] = None):
    dep_dir = Path(__file__).resolve().parent / ".pydeps"
    if dep_dir.exists() and str(dep_dir) not in sys.path:
        sys.path.insert(0, str(dep_dir))

    try:
        return importlib.import_module(module_name)
    except ImportError:
        pip_name = pip_name or module_name
        dep_dir.mkdir(parents=True, exist_ok=True)
        cmd = [
            sys.executable,
            "-m",
            "pip",
            "install",
            "--disable-pip-version-check",
            "--target",
            str(dep_dir),
            pip_name,
        ]
        print(f"[deps] Installing missing dependency: {pip_name}")
        subprocess.check_call(cmd)
        if str(dep_dir) not in sys.path:
            sys.path.insert(0, str(dep_dir))
        return importlib.import_module(module_name)


openpyxl = ensure_import("openpyxl")
oss2 = ensure_import("oss2")


def normalize_text(v: Any) -> str:
    if v is None:
        return ""
    return str(v).strip()


def is_http_url(text: str) -> bool:
    try:
        u = urllib.parse.urlparse(text)
        return u.scheme in ("http", "https")
    except Exception:
        return False


def to_file_name_like(text: str) -> str:
    s = normalize_text(text)
    if not s:
        return ""
    if is_http_url(s):
        p = urllib.parse.urlparse(s)
        return Path(p.path).name or s
    if "/" in s or "\\" in s:
        return Path(s).name
    return s


def normalize_name_key(text: str) -> str:
    s = to_file_name_like(text)
    if not s:
        return ""
    s = Path(s).stem
    s = re.sub(r"^\s*0*\d+[_\-\s]+", "", s)
    s = re.sub(r"(?i)[_\-\s]*(?:512x512|512x384|512x340|340x512|384x512)$", "", s)
    return "".join(ch.lower() for ch in s if ch.isalnum())


def extract_int(value: str) -> Optional[int]:
    s = normalize_text(value)
    if not s:
        return None
    m = re.search(r"\d+", s)
    if not m:
        return None
    try:
        return int(m.group(0))
    except Exception:
        return None


def col_letter_to_index(letter: str) -> int:
    letter = letter.strip().upper()
    if not re.fullmatch(r"[A-Z]+", letter):
        raise ValueError(f"Invalid column letter: {letter}")
    n = 0
    for ch in letter:
        n = n * 26 + (ord(ch) - ord("A") + 1)
    return n


def resolve_column(ws, spec: str, label: str) -> int:
    spec = (spec or "").strip()
    if not spec:
        raise SimbaError(f"{label} 不能为空")

    if spec.isdigit():
        idx = int(spec)
        if idx <= 0:
            raise SimbaError(f"{label} 列号必须 >= 1")
        return idx

    if re.fullmatch(r"[A-Za-z]{1,3}", spec):
        return col_letter_to_index(spec)

    headers = [normalize_text(ws.cell(row=1, column=c).value) for c in range(1, ws.max_column + 1)]
    for i, h in enumerate(headers, start=1):
        if h == spec:
            return i

    lower = spec.lower()
    for i, h in enumerate(headers, start=1):
        if h.lower() == lower:
            return i

    raise SimbaError(f"未找到{label}: {spec}")


def resolve_or_create_column(ws, spec: str, label: str, create_if_missing: bool = False) -> int:
    try:
        return resolve_column(ws, spec, label)
    except SimbaError:
        if not create_if_missing:
            raise
        token = (spec or "").strip()
        if not token or token.isdigit() or re.fullmatch(r"[A-Za-z]{1,3}", token):
            raise
        new_col = ws.max_column + 1
        ws.cell(row=1, column=new_col, value=token)
        return new_col


def cell_string(cell) -> str:
    try:
        if cell.hyperlink and cell.hyperlink.target:
            return normalize_text(cell.hyperlink.target)
    except Exception:
        pass
    return normalize_text(cell.value)


def cell_display_string(cell) -> str:
    return normalize_text(cell.value)


def guess_ext_from_bytes(data: bytes, fallback: str = ".jpg") -> str:
    if data.startswith(b"\xff\xd8\xff"):
        return ".jpg"
    if data.startswith(b"\x89PNG\r\n\x1a\n"):
        return ".png"
    if data.startswith((b"GIF87a", b"GIF89a")):
        return ".gif"
    if data.startswith(b"RIFF") and data[8:12] == b"WEBP":
        return ".webp"
    if data.startswith(b"BM"):
        return ".bmp"
    if data.startswith((b"II*\x00", b"MM\x00*")):
        return ".tiff"
    return fallback


def guess_ext_from_url(url: str) -> str:
    p = urllib.parse.urlparse(url)
    ext = Path(p.path).suffix.lower()
    if ext in IMAGE_EXTS:
        return ".jpg" if ext == ".jpeg" else ext
    mt, _ = mimetypes.guess_type(p.path)
    if mt and mt.startswith("image/"):
        mapped = mimetypes.guess_extension(mt)
        if mapped:
            return ".jpg" if mapped == ".jpe" else mapped
    return ".jpg"


def ensure_is_image(data: bytes, source: str):
    if len(data) < 16:
        raise SimbaError(f"图片内容过短，疑似无效文件: {source}")
    ext = guess_ext_from_bytes(data, fallback="")
    if not ext:
        head = data[:200].decode("utf-8", errors="ignore").lower()
        if "<html" in head or "<!doctype" in head:
            raise SimbaError(f"读取到HTML而非图片: {source}")
        raise SimbaError(f"无法识别为图片格式: {source}")


def read_image_bytes(source: str, timeout: int = 30) -> Tuple[bytes, str]:
    if is_http_url(source):
        req = urllib.request.Request(source, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            data = resp.read()
        ensure_is_image(data, source)
        ext = guess_ext_from_bytes(data, fallback=guess_ext_from_url(source))
        return data, ext

    p = Path(source).expanduser()
    if not p.is_absolute():
        p = Path.cwd() / p
    if not p.exists():
        raise SimbaError(f"本地文件不存在: {p}")
    data = p.read_bytes()
    ensure_is_image(data, str(p))
    ext = guess_ext_from_bytes(data, fallback=p.suffix.lower() or ".jpg")
    if ext == ".jpeg":
        ext = ".jpg"
    return data, ext


@dataclass
class FolderIndex:
    files: List[Path]
    by_rank: Dict[int, List[Path]]
    by_stem: Dict[str, List[Path]]
    by_norm: Dict[str, List[Path]]
    norm_of: Dict[Path, str]


def parse_rank_prefix(stem: str) -> Optional[int]:
    m = re.match(r"^\s*0*([1-9]\d*)[_\-\s]+", stem)
    if not m:
        return None
    try:
        return int(m.group(1))
    except Exception:
        return None


def _add_map_list(target: Dict[Any, List[Path]], key: Any, val: Path):
    if key in (None, ""):
        return
    arr = target.get(key)
    if arr is None:
        target[key] = [val]
    elif val not in arr:
        arr.append(val)


def build_folder_index(image_dir: Path) -> FolderIndex:
    if not image_dir.exists() or not image_dir.is_dir():
        raise SimbaError(f"image-dir 不存在或不是目录: {image_dir}")

    files = sorted([f for f in image_dir.iterdir() if f.is_file() and f.suffix.lower() in IMAGE_EXTS], key=lambda x: x.name.lower())
    if not files:
        raise SimbaError(f"image-dir 下没有图片文件: {image_dir}")

    by_rank: Dict[int, List[Path]] = {}
    by_stem: Dict[str, List[Path]] = {}
    by_norm: Dict[str, List[Path]] = {}
    norm_of: Dict[Path, str] = {}

    for f in files:
        stem = f.stem.strip()
        stem_lower = stem.lower()
        norm = normalize_name_key(stem)
        rank = parse_rank_prefix(stem)

        norm_of[f] = norm
        _add_map_list(by_stem, stem_lower, f)
        _add_map_list(by_norm, norm, f)
        _add_map_list(by_rank, rank, f)

    return FolderIndex(files=files, by_rank=by_rank, by_stem=by_stem, by_norm=by_norm, norm_of=norm_of)


def _unique_paths(paths: Sequence[Path]) -> List[Path]:
    uniq: Dict[str, Path] = {}
    for p in paths:
        uniq[str(p)] = p
    return [uniq[k] for k in sorted(uniq.keys(), key=lambda x: x.lower())]


def _collect_stem_candidates(values: Sequence[str]) -> List[str]:
    out: List[str] = []
    seen = set()
    for v in values:
        s = normalize_text(v)
        if not s:
            continue
        a = s.lower()
        b = to_file_name_like(s).lower()
        c = Path(to_file_name_like(s)).stem.lower()
        for token in (a, b, c):
            if token and token not in seen:
                seen.add(token)
                out.append(token)
    return out


def _collect_norm_candidates(values: Sequence[str]) -> List[str]:
    out: List[str] = []
    seen = set()
    for v in values:
        k = normalize_name_key(v)
        if k and k not in seen:
            seen.add(k)
            out.append(k)
    return out


def match_file_from_folder(
    idx: FolderIndex,
    match_values: Sequence[str],
    rank_value: str,
    allow_fuzzy: bool,
) -> Tuple[Path, str]:
    rank = extract_int(rank_value)
    stem_keys = _collect_stem_candidates(match_values)
    norm_keys = _collect_norm_candidates(match_values)

    rank_matches = idx.by_rank.get(rank, []) if rank is not None else []
    if len(rank_matches) == 1:
        return rank_matches[0], f"rank={rank}"

    if len(rank_matches) > 1 and norm_keys:
        filtered = []
        for p in rank_matches:
            fn = idx.norm_of.get(p, "")
            if any(n and (n == fn or n in fn or fn in n) for n in norm_keys):
                filtered.append(p)
        filtered = _unique_paths(filtered)
        if len(filtered) == 1:
            return filtered[0], f"rank+norm={rank}/{norm_keys[0]}"

    stem_hits: List[Path] = []
    for k in stem_keys:
        stem_hits.extend(idx.by_stem.get(k, []))
    stem_hits = _unique_paths(stem_hits)
    if len(stem_hits) == 1:
        return stem_hits[0], f"stem={stem_keys[0] if stem_keys else ''}"

    norm_hits: List[Path] = []
    for k in norm_keys:
        norm_hits.extend(idx.by_norm.get(k, []))
    norm_hits = _unique_paths(norm_hits)
    if len(norm_hits) == 1:
        return norm_hits[0], f"norm={norm_keys[0] if norm_keys else ''}"

    if allow_fuzzy and norm_keys:
        fuzzy_hits: List[Path] = []
        for p in idx.files:
            fn = idx.norm_of.get(p, "")
            for nk in norm_keys:
                if len(nk) >= 4 and fn and (nk in fn or fn in nk):
                    fuzzy_hits.append(p)
                    break
        fuzzy_hits = _unique_paths(fuzzy_hits)
        if len(fuzzy_hits) == 1:
            return fuzzy_hits[0], f"fuzzy={norm_keys[0]}"

    debug = {
        "rank": rank,
        "stem_keys": stem_keys[:5],
        "norm_keys": norm_keys[:5],
        "rank_candidates": [p.name for p in rank_matches[:5]],
        "stem_candidates": [p.name for p in stem_hits[:5]],
        "norm_candidates": [p.name for p in norm_hits[:5]],
    }
    raise SimbaError(f"文件夹匹配失败或不唯一: {json.dumps(debug, ensure_ascii=False)}")


def sort_files_for_table(files: Sequence[Path]) -> List[Path]:
    def k(p: Path):
        rank = parse_rank_prefix(p.stem)
        if rank is None:
            return (1, 999999, p.name.lower())
        return (0, rank, p.name.lower())

    return sorted(files, key=k)


def create_workbook_from_folder(image_dir: Path, sheet_name: str, rank_col_name: str, match_col_name: str, source_col_name: str):
    idx = build_folder_index(image_dir)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_name

    ws.cell(1, 1, rank_col_name)
    ws.cell(1, 2, match_col_name)
    ws.cell(1, 3, source_col_name)

    files = sort_files_for_table(idx.files)
    for i, p in enumerate(files, start=2):
        rank = parse_rank_prefix(p.stem)
        match_name = re.sub(r"^\s*0*\d+[_\-\s]+", "", p.stem).strip() or p.stem
        ws.cell(i, 1, rank if rank is not None else "")
        ws.cell(i, 2, match_name)
        ws.cell(i, 3, str(p))

    return wb, ws, idx


@dataclass
class StsCred:
    access_key_id: str
    access_key_secret: str
    security_token: str
    endpoint: str
    bucket: str
    expire_time: int


class SimbaClient:
    def __init__(self, token: str, api_base: str = API_BASE, language: str = "zh-CN", timezone: str = "Asia/Shanghai"):
        self.token = token.strip()
        if not self.token:
            raise SimbaError("Token 为空")
        self.api_base = api_base.rstrip("/")
        self.language = language
        self.timezone = timezone
        self._sts: Optional[StsCred] = None

    def _headers(self, content_type: Optional[str] = None) -> Dict[str, str]:
        h = {
            "Authorization": f"Bearer {self.token}",
            "Accept-Language": self.language,
            "Accept-Timezone": self.timezone,
            "User-Agent": "Mozilla/5.0",
        }
        if content_type:
            h["Content-Type"] = content_type
        return h

    def _json_request(self, method: str, path: str, payload: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
        data = None
        headers = self._headers("application/json" if payload is not None else None)
        if payload is not None:
            data = json.dumps(payload, ensure_ascii=False).encode("utf-8")

        url = self.api_base + path
        req = urllib.request.Request(url, data=data, headers=headers, method=method.upper())

        try:
            with urllib.request.urlopen(req, timeout=30) as resp:
                raw = resp.read().decode("utf-8")
        except urllib.error.HTTPError as e:
            body = e.read().decode("utf-8", errors="ignore") if hasattr(e, "read") else ""
            raise SimbaError(f"HTTP {e.code} {path} failed: {body[:500]}")
        except Exception as e:
            raise SimbaError(f"请求失败 {path}: {e}")

        try:
            obj = json.loads(raw)
        except json.JSONDecodeError:
            raise SimbaError(f"接口返回非JSON: {raw[:300]}")

        code = obj.get("code")
        if code not in (0, "0", None):
            raise SimbaError(f"接口返回失败 code={code}, message={obj.get('message') or obj.get('msg')}")
        return obj

    def get_sts(self, force_refresh: bool = False) -> StsCred:
        now = int(dt.datetime.now().timestamp())
        if not force_refresh and self._sts and self._sts.expire_time - now > 120:
            return self._sts

        obj = self._json_request("GET", STS_ENDPOINT)
        data = obj.get("data") or {}

        try:
            sts = StsCred(
                access_key_id=normalize_text(data["accessKeyId"]),
                access_key_secret=normalize_text(data["accessKeySecret"]),
                security_token=normalize_text(data["securityToken"]),
                endpoint=normalize_text(data["endPoint"]),
                bucket=normalize_text(data["bucket"]),
                expire_time=int(str(data.get("expireTime") or "0")),
            )
        except Exception as e:
            raise SimbaError(f"解析 STS 返回失败: {e}, raw={data}")

        if not all([sts.access_key_id, sts.access_key_secret, sts.security_token, sts.endpoint, sts.bucket]):
            raise SimbaError(f"STS 字段不完整: {data}")

        self._sts = sts
        return sts

    def upload_bytes_to_oss(self, image_bytes: bytes, ext: str) -> str:
        sts = self.get_sts()
        ext = ext if ext.startswith(".") else f".{ext}"
        today = dt.datetime.now()
        key = f"simba/{today.year:04d}/{today.month:02d}/{uuid.uuid4().hex}{ext}"

        auth = oss2.StsAuth(sts.access_key_id, sts.access_key_secret, sts.security_token)
        bucket = oss2.Bucket(auth, f"https://{sts.endpoint}", sts.bucket)

        try:
            result = bucket.put_object(key, image_bytes)
        except Exception as e:
            if "InvalidAccessKeyId" in str(e) or "SecurityTokenExpired" in str(e) or "AccessDenied" in str(e):
                sts = self.get_sts(force_refresh=True)
                auth = oss2.StsAuth(sts.access_key_id, sts.access_key_secret, sts.security_token)
                bucket = oss2.Bucket(auth, f"https://{sts.endpoint}", sts.bucket)
                result = bucket.put_object(key, image_bytes)
            else:
                raise SimbaError(f"OSS上传失败: {e}")

        if getattr(result, "status", None) not in (200, 201):
            raise SimbaError(f"OSS上传状态异常: {getattr(result, 'status', None)}")

        return key

    def generate_material_url(self, object_key: str, remark: str) -> str:
        payload = {
            "imageUrl": object_key,
            "remark": remark[:200] if remark else "",
        }
        obj = self._json_request("POST", GENERATE_ENDPOINT, payload)
        data = obj.get("data") or {}
        url = normalize_text(data.get("url"))
        if not url:
            raise SimbaError(f"生成链接返回为空: {obj}")
        return url


def default_output_path(input_path: Optional[Path], mode: str, image_dir: Optional[Path]) -> Path:
    if mode == "folder-match" and image_dir:
        if input_path:
            stem = input_path.stem + "_simba已回填"
        else:
            stem = "Simba图片上传回填结果"
        return image_dir / f"{stem}.xlsx"

    if input_path is None:
        raise SimbaError("source-col 模式必须提供 --input")

    stem = input_path.stem + "_simba已回填"
    return input_path.with_name(stem + input_path.suffix)


def setup_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(description="上传图片到 Simba 并回填 Excel")
    p.add_argument("--input", help="输入Excel路径。folder-match模式可不填：自动在图片目录生成表格")
    p.add_argument("--output", help="输出Excel路径，不填则自动生成")
    p.add_argument("--sheet", default="替代游戏总表", help="工作表名称")

    p.add_argument("--mode", choices=["source-col", "folder-match"], default="source-col", help="上传模式")

    p.add_argument("--source-col", default="缩略图2(512x512)", help="图片来源列（source-col 模式使用）")
    p.add_argument("--image-dir", help="图片目录（folder-match 模式必填）")
    p.add_argument("--match-col", default="替代游戏", help="folder-match：按该列值匹配文件名")
    p.add_argument("--rank-col", default="综合排名", help="folder-match：优先按该列匹配文件名前缀序号")
    p.add_argument("--allow-fuzzy", action="store_true", help="folder-match：开启模糊匹配兜底")

    p.add_argument("--url-col", default="Simba生成链接", help="Simba链接列（列名/列字母/列号）")
    p.add_argument("--key-col", default="Simba素材Key", help="Simba素材key列（列名/列字母/列号）")
    p.add_argument("--remark-col", default="替代游戏", help="备注列（用于上传remark）")

    p.add_argument("--start-row", type=int, default=2, help="起始行")
    p.add_argument("--overwrite", action="store_true", help="覆盖已有 Simba生成链接")
    p.add_argument("--max-upload", type=int, default=0, help="最大上传条数（0=不限制）")
    p.add_argument("--timeout", type=int, default=30, help="下载图片超时秒数")

    p.add_argument("--token", help="Simba Bearer token（不建议明文，留空将交互输入）")
    p.add_argument("--log", help="日志文件路径")
    return p


def run(args: argparse.Namespace) -> Dict[str, Any]:
    token = normalize_text(args.token) if args.token else ""
    if not token:
        token = normalize_text(getpass.getpass("请输入本次 Simba Token（Bearer 后面的字符串）: "))
    if not token:
        raise SimbaError("未提供 token")

    image_dir = Path(args.image_dir).expanduser().resolve() if args.image_dir else None
    if args.mode == "folder-match" and image_dir is None:
        raise SimbaError("folder-match 模式必须提供 --image-dir")

    auto_generated_table = False
    in_path: Optional[Path]
    folder_idx: Optional[FolderIndex] = None

    if args.input:
        in_path = Path(args.input).expanduser().resolve()
        if not in_path.exists():
            raise SimbaError(f"输入文件不存在: {in_path}")
        wb = openpyxl.load_workbook(in_path)
        if args.sheet not in wb.sheetnames:
            raise SimbaError(f"工作表不存在: {args.sheet}, 当前={wb.sheetnames}")
        ws = wb[args.sheet]
    else:
        if args.mode != "folder-match":
            raise SimbaError("source-col 模式必须提供 --input")
        if image_dir is None:
            raise SimbaError("缺少图片目录")
        wb, ws, folder_idx = create_workbook_from_folder(
            image_dir=image_dir,
            sheet_name=args.sheet,
            rank_col_name=args.rank_col,
            match_col_name=args.match_col,
            source_col_name=args.source_col,
        )
        auto_generated_table = True
        in_path = None
        print(f"[INFO] 已自动生成待回填表格（内存）: sheet={args.sheet}")

    out_path = Path(args.output).expanduser().resolve() if args.output else default_output_path(in_path, args.mode, image_dir)

    create_missing_cols = auto_generated_table
    url_col = resolve_or_create_column(ws, args.url_col, "url-col", create_if_missing=True)
    key_col = resolve_or_create_column(ws, args.key_col, "key-col", create_if_missing=True)
    remark_col = resolve_or_create_column(ws, args.remark_col, "remark-col", create_if_missing=create_missing_cols)

    source_col = resolve_or_create_column(ws, args.source_col, "source-col", create_if_missing=create_missing_cols)

    match_col = None
    rank_col = None
    if args.mode == "folder-match":
        match_col = resolve_or_create_column(ws, args.match_col, "match-col", create_if_missing=create_missing_cols)
        rank_col = resolve_or_create_column(ws, args.rank_col, "rank-col", create_if_missing=create_missing_cols)
        if folder_idx is None:
            folder_idx = build_folder_index(image_dir)
        print(f"[INFO] folder-match 已加载图片: {len(folder_idx.files)} 张")

    if args.mode == "folder-match" and match_col is not None:
        for r in range(args.start_row, ws.max_row + 1):
            if not cell_string(ws.cell(r, remark_col)):
                ws.cell(r, remark_col, cell_display_string(ws.cell(r, match_col)))

    client = SimbaClient(token=token)

    total_seen = 0
    success = 0
    skipped = 0
    failed = 0
    uploaded = 0
    errors: List[str] = []

    for r in range(args.start_row, ws.max_row + 1):
        existing = cell_string(ws.cell(r, url_col))
        if existing and not args.overwrite:
            skipped += 1
            continue

        if args.max_upload > 0 and uploaded >= args.max_upload:
            skipped += 1
            continue

        src = ""
        match_debug = ""

        if args.mode == "source-col":
            src = cell_string(ws.cell(r, source_col))
            if not src:
                continue
            total_seen += 1
        else:
            if auto_generated_table:
                src = cell_string(ws.cell(r, source_col))
                if not src:
                    continue
                total_seen += 1
                match_debug = "auto-table"
            else:
                row_match_values = [
                    cell_display_string(ws.cell(r, match_col)),
                    cell_string(ws.cell(r, match_col)),
                    cell_display_string(ws.cell(r, source_col)),
                    cell_string(ws.cell(r, source_col)),
                ]
                if not any(normalize_text(v) for v in row_match_values):
                    continue
                total_seen += 1
                rank_value = cell_display_string(ws.cell(r, rank_col))
                try:
                    matched_file, reason = match_file_from_folder(
                        idx=folder_idx,
                        match_values=row_match_values,
                        rank_value=rank_value,
                        allow_fuzzy=args.allow_fuzzy,
                    )
                    src = str(matched_file)
                    match_debug = reason
                except Exception as e:
                    failed += 1
                    err = f"row={r}, rank={rank_value}, match={normalize_text(row_match_values[0])}, error={e}"
                    errors.append(err)
                    print(f"[FAIL] {err}")
                    continue

        remark = cell_string(ws.cell(r, remark_col))
        if not remark and args.mode == "folder-match" and match_col is not None:
            remark = cell_display_string(ws.cell(r, match_col))
        if not remark:
            remark = f"row-{r}"

        try:
            img, ext = read_image_bytes(src, timeout=args.timeout)
            object_key = client.upload_bytes_to_oss(img, ext)
            simba_url = client.generate_material_url(object_key, remark)

            ws.cell(r, url_col, simba_url)
            ws.cell(r, key_col, object_key)
            success += 1
            uploaded += 1
            extra = f" ({match_debug})" if match_debug else ""
            print(f"[OK] row={r}{extra} -> {simba_url}")
        except Exception as e:
            failed += 1
            err = f"row={r}, source={src}, error={e}"
            errors.append(err)
            print(f"[FAIL] {err}")

    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)

    log_path = Path(args.log).expanduser().resolve() if args.log else out_path.with_name(out_path.stem + "_upload日志.txt")
    with log_path.open("w", encoding="utf-8") as f:
        f.write(f"输入文件: {in_path if in_path else '自动生成'}\n")
        f.write(f"输出文件: {out_path}\n")
        f.write(f"工作表: {args.sheet}\n")
        f.write(f"模式: {args.mode}\n")
        f.write(f"来源列: {args.source_col}, 匹配列: {args.match_col}, 排名列: {args.rank_col}\n")
        if image_dir:
            f.write(f"图片目录: {image_dir}\n")
        f.write(f"链接列: {args.url_col}, Key列: {args.key_col}, 备注列: {args.remark_col}\n")
        f.write(f"自动生成表格: {auto_generated_table}\n")
        f.write(f"扫描条目: {total_seen}\n")
        f.write(f"成功: {success}\n")
        f.write(f"失败: {failed}\n")
        f.write(f"跳过: {skipped}\n")
        if errors:
            f.write("\n失败明细:\n")
            for e in errors:
                f.write(f"- {e}\n")

    summary = {
        "input": str(in_path) if in_path else "自动生成",
        "output": str(out_path),
        "sheet": args.sheet,
        "mode": args.mode,
        "auto_generated_table": auto_generated_table,
        "scanned": total_seen,
        "success": success,
        "failed": failed,
        "skipped": skipped,
        "log": str(log_path),
    }
    print("\nSUMMARY:")
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return summary


def main() -> int:
    parser = setup_parser()
    args = parser.parse_args()
    try:
        run(args)
        return 0
    except Exception as e:
        print(f"\n[ERROR] {e}")
        print(traceback.format_exc())
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
