#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
证据文件按内容日期顺序重命名工具
=================================

功能：读取目标文件夹内每个文件的"内容日期"（优先取 开票日期 / 签订日期 /
签署日期 等标签后的日期，其次取文件中最早的日期），按日期从早到晚排序，
依次重命名为「原告证据-1」「原告证据-2」……（保留原扩展名）。

默认只做 dry-run（干跑预览，不改动任何文件）；加 --apply 才真正重命名。
重命名前会生成一份映射 CSV（旧名 -> 新名），可用 --revert 还原。

依赖：
  - 标准库（自带）
  - pypdf（仅处理 PDF 需要；已装在隔离 venv 中）
  文本型 .docx 用标准库 zipfile+xml 解析，无需额外依赖。
  图片型文件（无文字层）当前不参与日期提取，将排在最后，不影响其余排序。

用法见同目录 README.md
"""

import argparse
import csv
import datetime
import os
import re
import sys
from pathlib import Path

# 纳入排序/重命名的文档类扩展名
DOC_EXTS = {
    ".pdf", ".docx", ".doc", ".txt", ".md", ".csv", ".json",
    ".jpg", ".jpeg", ".png", ".tif", ".tiff",
}

# 优先作为"文件日期"的标签（按列表顺序匹配，命中即用其后首个日期）
DATE_LABELS = [
    "开票日期", "出具日期", "签发日期", "签署日期", "签订日期",
    "成立日期", "创建日期", "制表日期", "日期", "时间",
]

# 日期正则：支持 2026-08-12 / 2026/8/12 / 2026.08.12 / 2026年08月12日 等
DATE_RE = re.compile(
    r"(?P<y>\d{4})\s*[-/年.\s]\s*(?P<m>\d{1,2})\s*[-/月.\s]\s*(?P<d>\d{1,2})\s*日?"
)
# 紧凑型 YYYYMMDD
DATE8_RE = re.compile(r"(?<![\d])(?P<y>\d{4})(?P<m>\d{2})(?P<d>\d{2})(?![\d])")

# 默认目标目录（本机证据文件夹）；可用 --dir 覆盖
DEFAULT_DIR = "/Users/zhoulixuan/Desktop/四明山案-原告证据"
DEFAULT_PREFIX = "原告证据"


# ------------------------- 文本提取 -------------------------
def extract_text(path: Path) -> str:
    """按扩展名提取文件文本内容；失败或无文本返回空串。"""
    ext = path.suffix.lower()
    try:
        if ext == ".pdf":
            import pypdf  # 懒加载，缺依赖时仅 PDF 受影响
            reader = pypdf.PdfReader(str(path))
            return "\n".join((p.extract_text() or "") for p in reader.pages)
        if ext == ".docx":
            return _extract_docx(path)
        if ext in (".txt", ".md", ".csv", ".json"):
            return path.read_text(encoding="utf-8", errors="ignore")
        # 图片：如需 OCR 可在此扩展；当前不参与日期提取
        return ""
    except Exception as e:  # noqa: BLE001
        print(f"[warn] 读取失败 {path.name}: {e}", file=sys.stderr)
        return ""


def _extract_docx(path: Path) -> str:
    import zipfile
    import xml.etree.ElementTree as ET  # noqa: F401

    with zipfile.ZipFile(str(path)) as zf:
        xml = zf.read("word/document.xml").decode("utf-8", "ignore")
    return re.sub(r"<[^>]+>", " ", xml)


# ------------------------- 日期解析 -------------------------
def _to_date(y, m, d) -> datetime.date | None:
    try:
        dt = datetime.date(int(y), int(m), int(d))
        return dt if 1900 <= dt.year <= 2100 else None
    except (ValueError, TypeError):
        return None


def collect_dates(text: str):
    """返回文本中所有合法日期（按出现顺序）。"""
    out = []
    for m in DATE_RE.finditer(text):
        dt = _to_date(m.group("y"), m.group("m"), m.group("d"))
        if dt:
            out.append(dt)
    for m in DATE8_RE.finditer(text):
        dt = _to_date(m.group("y"), m.group("m"), m.group("d"))
        if dt:
            out.append(dt)
    return out


def file_date(text: str) -> datetime.date | None:
    """确定文件代表日期：优先标签后日期，其次全文中最早日期。"""
    for label in DATE_LABELS:
        i = text.find(label)
        if i < 0:
            continue
        after = text[i + len(label):]
        m = DATE_RE.search(after)
        if m:
            dt = _to_date(m.group("y"), m.group("m"), m.group("d"))
            if dt:
                return dt
        m8 = DATE8_RE.search(after)
        if m8:
            dt = _to_date(m8.group("y"), m8.group("m"), m8.group("d"))
            if dt:
                return dt
    ds = collect_dates(text)
    return min(ds) if ds else None


# ------------------------- 主流程 -------------------------
def gather_files(d: Path, prefix: str, log_name: str):
    """收集待处理文件，排除脚本自身、README、映射CSV、已重命名的文件。"""
    pref_re = re.compile(re.escape(prefix) + r"-\d+$")
    exclude = {
        Path(__file__).resolve().name,
        "README.md",
        log_name,
    }
    files = []
    for p in sorted(d.iterdir()):
        if not p.is_file():
            continue
        if p.suffix.lower() not in DOC_EXTS:
            continue
        if p.resolve().name in exclude:
            continue
        if pref_re.match(p.stem):  # 已是 前缀-N 形式，跳过
            continue
        files.append(p)
    return files


def build_mapping(d: Path, prefix: str):
    files = gather_files(d, prefix, "")
    records = []
    for p in files:
        dt = file_date(extract_text(p))
        records.append({"name": p.name, "date": dt, "suffix": p.suffix})
    # 排序：有日期在前(升序)，无日期在后；同级按原名保证确定性
    records.sort(key=lambda r: (
        0 if r["date"] else 1,
        r["date"] or datetime.date(9999, 1, 1),
        r["name"],
    ))
    mapping = []
    for i, r in enumerate(records, 1):
        new_name = f"{prefix}-{i}{r['suffix']}"
        mapping.append({
            "old": r["name"],
            "new": new_name,
            "date": r["date"].isoformat() if r["date"] else "无日期",
        })
    return mapping


def write_log(d: Path, prefix: str, mapping):
    log_path = Path(__file__).with_name(f"{Path(d).name}_重命名映射.csv")
    with open(log_path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["原文件名", "新文件名", "内容日期"])
        for m in mapping:
            w.writerow([m["old"], m["new"], m["date"]])
    return log_path


def apply_rename(d: Path, mapping):
    done, skipped = 0, 0
    for m in mapping:
        src = d / m["old"]
        dst = d / m["new"]
        if dst.exists():
            print(f"  [跳过] 目标已存在：{m['new']}")
            skipped += 1
            continue
        src.rename(dst)
        done += 1
    return done, skipped


def revert(log_csv: str, d: Path):
    """按映射CSV把 前缀-N 还原为原文件名。证据目录由 --dir 指定。"""
    log_path = Path(log_csv).expanduser().resolve()
    if not log_path.exists():
        print("映射文件不存在：", log_path)
        sys.exit(1)
    with open(log_path, encoding="utf-8-sig") as f:
        rows = list(csv.reader(f))
    if not rows or rows[0][0] != "原文件名":
        print("映射文件格式不正确")
        sys.exit(1)
    done = 0
    for old, new, _ in rows[1:]:
        src = d / new
        dst = d / old
        if src.exists() and not dst.exists():
            src.rename(dst)
            done += 1
            print(f"  还原：{new} -> {old}")
        else:
            print(f"  [跳过] {new}（源不存在或目标已存在）")
    print(f"还原完成，共还原 {done} 个文件。")


def main():
    ap = argparse.ArgumentParser(
        description="按文件内容日期排序，批量重命名为 前缀-N",
    )
    ap.add_argument("--dir", default=DEFAULT_DIR, help="目标文件夹（默认：本机证据文件夹）")
    ap.add_argument("--prefix", default=DEFAULT_PREFIX, help="新文件名前缀（默认：原告证据）")
    ap.add_argument("--apply", action="store_true", help="真实执行重命名（默认仅预览）")
    ap.add_argument("--revert", metavar="CSV", help="用映射CSV还原文件名")
    args = ap.parse_args()

    d = Path(args.dir).expanduser().resolve()
    if not d.is_dir():
        print("目录不存在：", d)
        sys.exit(1)

    if args.revert:
        revert(args.revert, d)
        return

    mapping = build_mapping(d, args.prefix)
    if not mapping:
        print("未发现可处理的文档文件。")
        return

    log_path = write_log(d, args.prefix, mapping)

    print("=" * 54)
    print(f"目标目录：{d}")
    print(f"排序依据：文件内容中的日期（升序）")
    print(f"模式：{'真实重命名' if args.apply else '预览(dry-run，未改动文件)'}")
    print("-" * 54)
    print(f"{'序号':<6}{'内容日期':<14}{'原文件名':<22}-> 新文件名")
    for i, m in enumerate(mapping, 1):
        print(f"{i:<6}{m['date']:<14}{m['old']:<22}-> {m['new']}")
    print("-" * 54)
    print(f"映射已写入：{log_path}")

    if args.apply:
        done, skipped = apply_rename(d, mapping)
        print(f"执行完成：重命名 {done} 个，跳过 {skipped} 个。")
    else:
        print("（未改动文件。确认无误后加 --apply 执行真实重命名）")


if __name__ == "__main__":
    main()
