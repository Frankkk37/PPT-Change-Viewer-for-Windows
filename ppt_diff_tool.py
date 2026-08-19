#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PPT Diff MVP v0.9 Windows

基准：v0.9 MVP
目标：Windows 电脑上可直接测试的本地 A/B PPT 比较工具。

功能：
- 人工选择旧版 PPT、 新版 PPT
- 识别新增、删除、移动、修改页
- 短句级文字差异
- HTML / Markdown / JSON 三种报告
- UI 按钮：打开 HTML、打开 Markdown、打开输出文件夹
- 无第三方依赖，仅需 Python 3 标准库

启动：
- Windows 双击：启动PPTDiff.bat
- 命令行：python ppt_diff_tool.py --ui
"""

from __future__ import annotations

import argparse
import datetime as _dt
import difflib
import hashlib
import html
import json
import os
import posixpath
import re
import sys
import threading
import traceback
import zipfile
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Dict, List, Optional, Tuple, Set
import xml.etree.ElementTree as ET


NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}


@dataclass
class SlideInfo:
    position: int
    slide_id: str
    r_id: str
    slide_path: str
    text_blocks: List[str]
    text_segments: List[str]
    text_norm: str
    meaningful_text_hash: str
    image_hashes: List[str]
    image_count: int
    shape_count: int
    graphic_frame_count: int
    semantic_xml_hash: str
    content_hash: str


@dataclass
class SlideMatch:
    old_position: int
    new_position: int
    old_slide_id: str
    new_slide_id: str
    method: str
    similarity: float
    moved: bool


def _sha256_text(text: str) -> str:
    return hashlib.sha256(text.encode("utf-8", errors="ignore")).hexdigest()


def _sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def _clean_text(text: str) -> str:
    return re.sub(r"\s+", " ", text.replace("\u00a0", " ")).strip()


def _normalize_text(text: str) -> str:
    return _clean_text(text).lower()


def _truncate(text: str, max_len: int = 240) -> str:
    text = _clean_text(text)
    return text if len(text) <= max_len else text[:max_len - 1] + "…"


def safe_filename(name: str, max_len: int = 120) -> str:
    name = re.sub(r'[<>:"/\\|?*\n\r\t]+', "_", name)
    name = re.sub(r"\s+", "_", name).strip("_")
    name = name or "ppt_diff"
    return name[:max_len].rstrip("_")


def _read_xml(zf: zipfile.ZipFile, path: str) -> Optional[ET.Element]:
    try:
        data = zf.read(path)
    except KeyError:
        return None
    try:
        return ET.fromstring(data)
    except ET.ParseError:
        return None


def _extract_digit_sequences(text: str) -> List[str]:
    return re.findall(r"\d+", text)


def _mask_digit_sequences(text: str) -> str:
    return re.sub(r"\d+", "#", text)


def _is_numeric_like(text: str) -> bool:
    text = text.strip()
    if not text:
        return False
    return bool(re.fullmatch(r"[\sPpPage第页/\\|:：\-–—().（）\d]+", text))


def _is_probable_page_number_chunk(text: str) -> bool:
    t = _clean_text(text)
    return bool(t and len(t) <= 12 and _is_numeric_like(t) and _extract_digit_sequences(t))


def _meaningful_items(items: List[str]) -> List[str]:
    return [x for x in items if not _is_probable_page_number_chunk(x)]


def _resolve_target(source_xml_path: str, target: str) -> str:
    target = target.replace("\\", "/")
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(posixpath.dirname(source_xml_path), target))


def _rels_path_for(xml_path: str) -> str:
    return posixpath.join(posixpath.dirname(xml_path), "_rels", posixpath.basename(xml_path) + ".rels")


def _read_relationships(zf: zipfile.ZipFile, xml_path: str) -> Dict[str, Dict[str, str]]:
    root = _read_xml(zf, _rels_path_for(xml_path))
    if root is None:
        return {}
    rels: Dict[str, Dict[str, str]] = {}
    for rel in root.findall("rel:Relationship", NS):
        rel_id = rel.attrib.get("Id")
        target = rel.attrib.get("Target")
        rel_type = rel.attrib.get("Type", "")
        target_mode = rel.attrib.get("TargetMode", "")
        if rel_id and target:
            rels[rel_id] = {
                "target": target,
                "target_resolved": _resolve_target(xml_path, target),
                "type": rel_type,
                "target_mode": target_mode,
            }
    return rels


def _extract_slide_order(zf: zipfile.ZipFile) -> List[Tuple[str, str]]:
    root = _read_xml(zf, "ppt/presentation.xml")
    if root is None:
        raise ValueError("无法读取 ppt/presentation.xml。请确认文件是有效的 .pptx 文件。")
    items = []
    for sld in root.findall(".//p:sldIdLst/p:sldId", NS):
        slide_id = sld.attrib.get("id", "")
        r_id = sld.attrib.get(f"{{{NS['r']}}}id", "")
        if slide_id and r_id:
            items.append((slide_id, r_id))
    return items


def _extract_text_blocks(slide_root: ET.Element) -> List[str]:
    blocks: List[str] = []
    for para in slide_root.findall(".//a:p", NS):
        texts = []
        for tnode in para.findall(".//a:t", NS):
            if tnode.text is not None:
                texts.append(tnode.text.replace("\u00a0", " "))
        joined = "".join(texts)
        cleaned = _clean_text(joined)
        if cleaned:
            blocks.append(cleaned)
    return blocks


_SEG_SPLIT_RE = re.compile(r"([，,、；;：:。.!！？?／/|｜])")


def _split_short_segments(text: str, max_segment_len: int = 90) -> List[str]:
    text = _clean_text(text)
    if not text:
        return []
    parts = _SEG_SPLIT_RE.split(text)
    segs: List[str] = []
    buf = ""
    for p in parts:
        if not p:
            continue
        buf += p
        if _SEG_SPLIT_RE.fullmatch(p):
            c = _clean_text(buf)
            if c:
                segs.append(c)
            buf = ""
    if _clean_text(buf):
        segs.append(_clean_text(buf))
    if not segs:
        segs = [text]

    refined: List[str] = []
    for seg in segs:
        if len(seg) <= max_segment_len:
            refined.append(seg)
        else:
            words = seg.split(" ")
            if len(words) <= 1:
                refined.append(seg)
            else:
                cur = ""
                for w in words:
                    if not cur:
                        cur = w
                    elif len(cur) + 1 + len(w) <= max_segment_len:
                        cur += " " + w
                    else:
                        refined.append(cur)
                        cur = w
                if cur:
                    refined.append(cur)
    return refined


def _extract_text_segments(blocks: List[str]) -> List[str]:
    segs: List[str] = []
    for b in blocks:
        segs.extend(_split_short_segments(b))
    return segs


def _extract_slide_counts(slide_root: ET.Element) -> Tuple[int, int, int]:
    return (
        len(slide_root.findall(".//p:sp", NS)),
        len(slide_root.findall(".//p:pic", NS)),
        len(slide_root.findall(".//p:graphicFrame", NS)),
    )


def _extract_image_hashes(zf: zipfile.ZipFile, slide_path: str) -> List[str]:
    rels = _read_relationships(zf, slide_path)
    hashes: List[str] = []
    for rel in rels.values():
        target = rel["target_resolved"]
        rel_type = rel.get("type", "")
        target_mode = rel.get("target_mode", "")
        is_image = "image" in rel_type.lower() or target.lower().endswith(
            (".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff", ".emf", ".wmf", ".svg")
        )
        if not is_image:
            continue
        if target_mode == "External":
            hashes.append("external:" + _sha256_text(rel["target"]))
        else:
            try:
                hashes.append(_sha256_bytes(zf.read(target)))
            except KeyError:
                hashes.append("missing:" + _sha256_text(target))
    return sorted(hashes)


def _semantic_xml_hash(slide_root: ET.Element) -> str:
    cloned = ET.fromstring(ET.tostring(slide_root, encoding="utf-8"))
    for node in cloned.findall(".//a:t", NS):
        if node.text is not None and _is_probable_page_number_chunk(node.text):
            node.text = "__PAGE_NUMBER__"
    return _sha256_bytes(ET.tostring(cloned, encoding="utf-8"))


def extract_pptx(path: Path) -> List[SlideInfo]:
    if not path.exists():
        raise FileNotFoundError(f"文件不存在：{path}")
    if path.suffix.lower() != ".pptx":
        raise ValueError(f"当前只支持 .pptx，不支持：{path.suffix or '(无扩展名)'}。请先另存为 .pptx。")

    try:
        zf_ctx = zipfile.ZipFile(path, "r")
    except zipfile.BadZipFile as e:
        raise ValueError(f"无法打开 PPTX。文件可能损坏，或不是标准 .pptx：{path}") from e

    with zf_ctx as zf:
        presentation_rels = _read_relationships(zf, "ppt/presentation.xml")
        order = _extract_slide_order(zf)
        slides: List[SlideInfo] = []

        for idx, (slide_id, r_id) in enumerate(order, start=1):
            rel = presentation_rels.get(r_id)
            if not rel:
                continue
            slide_path = rel["target_resolved"]
            root = _read_xml(zf, slide_path)
            if root is None:
                continue

            blocks = _extract_text_blocks(root)
            segments = _extract_text_segments(blocks)
            meaningful_blocks = _meaningful_items(blocks)
            meaningful_segments = _meaningful_items(segments)
            text_norm = _normalize_text("\n".join(blocks))
            meaningful_norm = _normalize_text("\n".join(meaningful_blocks))

            shape_count, xml_img_count, graphic_frame_count = _extract_slide_counts(root)
            image_hashes = _extract_image_hashes(zf, slide_path)
            image_count = max(xml_img_count, len(image_hashes))
            semantic_hash = _semantic_xml_hash(root)

            content_basis = json.dumps(
                {
                    "meaningful_blocks": meaningful_norm,
                    "meaningful_segments": _normalize_text("\n".join(meaningful_segments)),
                    "image_hashes": image_hashes,
                    "shape_count": shape_count,
                    "image_count": image_count,
                    "graphic_frame_count": graphic_frame_count,
                },
                ensure_ascii=False,
                sort_keys=True,
            )

            slides.append(
                SlideInfo(
                    position=idx,
                    slide_id=str(slide_id),
                    r_id=r_id,
                    slide_path=slide_path,
                    text_blocks=blocks,
                    text_segments=segments,
                    text_norm=text_norm,
                    meaningful_text_hash=_sha256_text(meaningful_norm),
                    image_hashes=image_hashes,
                    image_count=image_count,
                    shape_count=shape_count,
                    graphic_frame_count=graphic_frame_count,
                    semantic_xml_hash=semantic_hash,
                    content_hash=_sha256_text(content_basis),
                )
            )

        return slides


def _text_similarity(a: str, b: str) -> float:
    if not a and not b:
        return 1.0
    if not a or not b:
        return 0.0
    return difflib.SequenceMatcher(None, a, b).ratio()


def _jaccard(a: List[str], b: List[str]) -> float:
    sa, sb = set(a), set(b)
    if not sa and not sb:
        return 1.0
    if not sa or not sb:
        return 0.0
    return len(sa & sb) / len(sa | sb)


def _count_similarity(a: int, b: int) -> float:
    if a == b:
        return 1.0
    high = max(a, b)
    return 1.0 if high == 0 else 1.0 - abs(a - b) / high


def slide_similarity(old: SlideInfo, new: SlideInfo) -> float:
    old_text = _normalize_text("\n".join(_meaningful_items(old.text_blocks)))
    new_text = _normalize_text("\n".join(_meaningful_items(new.text_blocks)))
    text_sim = _text_similarity(old_text, new_text)
    image_sim = _jaccard(old.image_hashes, new.image_hashes)
    count_sim = (
        _count_similarity(old.shape_count, new.shape_count) * 0.45
        + _count_similarity(old.image_count, new.image_count) * 0.35
        + _count_similarity(old.graphic_frame_count, new.graphic_frame_count) * 0.20
    )
    if len(old_text) >= 20 or len(new_text) >= 20:
        score = 0.76 * text_sim + 0.14 * image_sim + 0.10 * count_sim
    else:
        score = 0.35 * text_sim + 0.45 * image_sim + 0.20 * count_sim
    return round(score, 6)


def match_slides(old_slides: List[SlideInfo], new_slides: List[SlideInfo], threshold: float = 0.78):
    old_by_id = {s.slide_id: s for s in old_slides}
    new_by_id = {s.slide_id: s for s in new_slides}

    matched_old: Set[str] = set()
    matched_new: Set[str] = set()
    matches: List[SlideMatch] = []

    for sid, old in old_by_id.items():
        new = new_by_id.get(sid)
        if not new:
            continue
        matched_old.add(old.slide_id)
        matched_new.add(new.slide_id)
        sim = 1.0 if old.content_hash == new.content_hash else slide_similarity(old, new)
        matches.append(SlideMatch(old.position, new.position, old.slide_id, new.slide_id, "slide_id", sim, old.position != new.position))

    unmatched_old = [s for s in old_slides if s.slide_id not in matched_old]
    unmatched_new = [s for s in new_slides if s.slide_id not in matched_new]

    candidates = []
    for old in unmatched_old:
        for new in unmatched_new:
            sim = slide_similarity(old, new)
            if sim >= threshold:
                candidates.append((sim, old, new))
    candidates.sort(key=lambda x: x[0], reverse=True)

    used_old: Set[str] = set()
    used_new: Set[str] = set()
    for sim, old, new in candidates:
        if old.slide_id in used_old or new.slide_id in used_new:
            continue
        used_old.add(old.slide_id)
        used_new.add(new.slide_id)
        matches.append(SlideMatch(old.position, new.position, old.slide_id, new.slide_id, "content_similarity", sim, old.position != new.position))

    removed = [s for s in unmatched_old if s.slide_id not in used_old]
    added = [s for s in unmatched_new if s.slide_id not in used_new]

    matches.sort(key=lambda m: (m.new_position, m.old_position))
    added.sort(key=lambda s: s.position)
    removed.sort(key=lambda s: s.position)
    return matches, removed, added


def _inline_char_diff(old: str, new: str, max_ops: int = 40) -> List[Dict]:
    sm = difflib.SequenceMatcher(None, old, new)
    ops = []
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            continue
        ops.append({"type": tag, "old": old[i1:i2], "new": new[j1:j2]})
        if len(ops) >= max_ops:
            break
    return ops


def _change_kind(ops: List[Dict]) -> str:
    if not ops:
        return "文字变化"
    has_insert = any(o["type"] == "insert" for o in ops)
    has_delete = any(o["type"] == "delete" for o in ops)
    has_replace = any(o["type"] == "replace" for o in ops)
    if has_insert and not has_delete and not has_replace:
        return "短句内新增"
    if has_delete and not has_insert and not has_replace:
        return "短句内删除"
    if has_replace and not has_insert and not has_delete:
        return "短句内替换"
    return "短句内修改"


def build_text_changes(old: SlideInfo, new: SlideInfo, max_changes: int = 40) -> List[Dict]:
    old_items = _meaningful_items(old.text_segments) or _meaningful_items(old.text_blocks)
    new_items = _meaningful_items(new.text_segments) or _meaningful_items(new.text_blocks)

    sm = difflib.SequenceMatcher(None, old_items, new_items)
    changes: List[Dict] = []

    def add(c):
        if len(changes) < max_changes:
            changes.append(c)

    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            continue
        old_block = old_items[i1:i2]
        new_block = new_items[j1:j2]

        if tag == "delete":
            for txt in old_block:
                add({"type": "removed_segment", "old": _truncate(txt), "new": "", "char_changes": [], "kind": "删除短句"})
        elif tag == "insert":
            for txt in new_block:
                add({"type": "added_segment", "old": "", "new": _truncate(txt), "char_changes": [], "kind": "新增短句"})
        elif tag == "replace":
            max_len = max(len(old_block), len(new_block))
            for k in range(max_len):
                old_txt = old_block[k] if k < len(old_block) else ""
                new_txt = new_block[k] if k < len(new_block) else ""
                if old_txt and new_txt:
                    ops = _inline_char_diff(old_txt, new_txt)
                    add({"type": "changed_segment", "old": _truncate(old_txt), "new": _truncate(new_txt), "char_changes": ops, "kind": _change_kind(ops)})
                elif old_txt:
                    add({"type": "removed_segment", "old": _truncate(old_txt), "new": "", "char_changes": [], "kind": "删除短句"})
                else:
                    add({"type": "added_segment", "old": "", "new": _truncate(new_txt), "char_changes": [], "kind": "新增短句"})

    if len(changes) >= max_changes:
        changes.append({"type": "truncated", "old": "", "new": f"Text diff truncated at {max_changes} changes.", "char_changes": [], "kind": "已截断"})
    return changes


def _is_page_number_like_change(change: Dict) -> bool:
    old = (change.get("old") or "").strip()
    new = (change.get("new") or "").strip()
    if change.get("type") == "truncated":
        return False
    if old and new:
        if old == new:
            return True
        if _extract_digit_sequences(old) and _extract_digit_sequences(new) and _mask_digit_sequences(old) == _mask_digit_sequences(new):
            return True
        if _is_probable_page_number_chunk(old) and _is_probable_page_number_chunk(new):
            return True
        return False
    solo = old or new
    return _is_probable_page_number_chunk(solo)


def build_moved_groups(moved_items: List[Dict]) -> List[Dict]:
    if not moved_items:
        return []
    items = sorted(moved_items, key=lambda x: (x["old_position"], x["new_position"]))
    groups = []
    cur = [items[0]]

    def same_run(prev, item):
        return (
            item["old_position"] == prev["old_position"] + 1
            and item["new_position"] == prev["new_position"] + 1
            and item["delta"] == prev["delta"]
        )

    for item in items[1:]:
        if same_run(cur[-1], item):
            cur.append(item)
        else:
            groups.append(cur)
            cur = [item]
    groups.append(cur)

    result = []
    for g in groups:
        delta = g[0]["delta"]
        result.append({
            "old_start": g[0]["old_position"],
            "old_end": g[-1]["old_position"],
            "new_start": g[0]["new_position"],
            "new_end": g[-1]["new_position"],
            "count": len(g),
            "delta": delta,
            "direction": "后移" if delta > 0 else "前移" if delta < 0 else "未移动",
        })
    result.sort(key=lambda x: (-x["count"], x["old_start"]))
    return result


def build_diff(old_path: Path, new_path: Path, detect_format: bool = False, threshold: float = 0.78) -> Dict:
    old_slides = extract_pptx(old_path)
    new_slides = extract_pptx(new_path)
    matches, removed, added = match_slides(old_slides, new_slides, threshold)

    old_by_id = {s.slide_id: s for s in old_slides}
    new_by_id = {s.slide_id: s for s in new_slides}

    moved_items = []
    modified_details = []
    ignored_page_numbers = []
    format_only = []

    for m in matches:
        old = old_by_id.get(m.old_slide_id)
        new = new_by_id.get(m.new_slide_id)
        if not old or not new:
            continue

        if m.moved:
            moved_items.append({
                "old_position": m.old_position,
                "new_position": m.new_position,
                "old_slide_id": m.old_slide_id,
                "new_slide_id": m.new_slide_id,
                "match_method": m.method,
                "similarity": m.similarity,
                "delta": m.new_position - m.old_position,
            })

        meaningful_text_changed = old.meaningful_text_hash != new.meaningful_text_hash
        images_changed = old.image_hashes != new.image_hashes
        shape_count_changed = old.shape_count != new.shape_count
        image_count_changed = old.image_count != new.image_count
        graphic_frame_count_changed = old.graphic_frame_count != new.graphic_frame_count
        semantic_xml_changed = old.semantic_xml_hash != new.semantic_xml_hash

        text_changes = build_text_changes(old, new) if old.text_norm != new.text_norm else []
        only_page_number = bool(text_changes and all(_is_page_number_like_change(c) for c in text_changes))

        content_changed = any([meaningful_text_changed, images_changed, shape_count_changed, image_count_changed, graphic_frame_count_changed])
        format_changed = bool(detect_format and semantic_xml_changed and not content_changed and not only_page_number)
        modified = content_changed or format_changed

        detail = {
            "old_position": m.old_position,
            "new_position": m.new_position,
            "old_slide_id": m.old_slide_id,
            "new_slide_id": m.new_slide_id,
            "match_method": m.method,
            "similarity": m.similarity,
            "changed": {
                "meaningful_text": meaningful_text_changed,
                "images": images_changed,
                "shape_count": shape_count_changed,
                "image_count": image_count_changed,
                "graphic_frame_count": graphic_frame_count_changed,
                "format_or_layout_xml": format_changed,
                "only_page_number_text_change": only_page_number,
            },
            "text_changes": text_changes if meaningful_text_changed else [],
        }

        if modified:
            modified_details.append(detail)
            if format_changed:
                format_only.append(detail)
        elif only_page_number:
            ignored_page_numbers.append(detail)

    moved_groups = build_moved_groups(moved_items)

    return {
        "schema_version": "0.9-windows",
        "generated_at": _dt.datetime.now().isoformat(timespec="seconds"),
        "from_file": str(old_path),
        "to_file": str(new_path),
        "from_name": old_path.stem,
        "to_name": new_path.stem,
        "settings": {
            "similarity_threshold": threshold,
            "matching": "slide_id_first_then_content_similarity",
            "page_number_filter": True,
            "detect_format_or_layout": detect_format,
            "text_diff_mode": "short_segment",
        },
        "summary": {
            "old_slide_count": len(old_slides),
            "new_slide_count": len(new_slides),
            "added": len(added),
            "removed": len(removed),
            "moved": len(moved_items),
            "moved_groups": len(moved_groups),
            "modified": len(modified_details),
            "format_or_layout_only": len(format_only),
            "matched": len(matches),
            "ignored_page_number_only_changes": len(ignored_page_numbers),
        },
        "added_slides": [{"new_position": s.position, "new_slide_id": s.slide_id, "text_preview": _truncate(s.text_norm, 160)} for s in added],
        "removed_slides": [{"old_position": s.position, "old_slide_id": s.slide_id, "text_preview": _truncate(s.text_norm, 160)} for s in removed],
        "moved_slides": moved_items,
        "moved_groups": moved_groups,
        "modified_slides": modified_details,
        "format_only_slides": format_only,
        "ignored_page_number_only_changes": ignored_page_numbers,
        "matches": [asdict(m) for m in matches],
    }


def _format_range(prefix: str, start: int, end: int) -> str:
    return f"{prefix}{start}" if start == end else f"{prefix}{start}–{prefix}{end}"


def _format_moved_group(g: Dict) -> str:
    old_range = _format_range("P", g["old_start"], g["old_end"])
    new_range = _format_range("P", g["new_start"], g["new_end"])
    delta = abs(g["delta"])
    if g["count"] == 1:
        return f"旧 {old_range} → 新 {new_range}（{g['direction']} {delta} 页）"
    return f"旧 {old_range} → 新 {new_range}（整体{g['direction']} {delta} 页，{g['count']} 页）"


def _format_char_changes_md(ops: List[Dict]) -> str:
    parts = []
    for op in ops[:12]:
        old = op.get("old", "")
        new = op.get("new", "")
        typ = op.get("type")
        if typ == "insert":
            parts.append(f"新增 `{new}`")
        elif typ == "delete":
            parts.append(f"删除 `{old}`")
        elif typ == "replace":
            parts.append(f"`{old}` → `{new}`")
    return "；".join(parts)


def _change_labels(ch: Dict) -> List[str]:
    labels = []
    if ch.get("meaningful_text"):
        labels.append("文字")
    if ch.get("images"):
        labels.append("图片")
    if ch.get("shape_count"):
        labels.append("形状数量")
    if ch.get("image_count"):
        labels.append("图片数量")
    if ch.get("graphic_frame_count"):
        labels.append("表格/图表容器")
    if ch.get("format_or_layout_xml"):
        labels.append("格式/布局")
    return labels or ["内容变化"]


def render_markdown(diff: Dict) -> str:
    s = diff["summary"]
    lines = [
        "# PPT Diff Report",
        "",
        f"- From: `{Path(diff['from_file']).name}`",
        f"- To: `{Path(diff['to_file']).name}`",
        f"- Generated at: `{diff['generated_at']}`",
        f"- Schema version: `{diff['schema_version']}`",
        f"- Text diff mode: `{diff['settings']['text_diff_mode']}`",
        f"- Detect format/layout: `{diff['settings']['detect_format_or_layout']}`",
        "",
        "## Summary",
        "",
        f"- Old slide count: **{s['old_slide_count']}**",
        f"- New slide count: **{s['new_slide_count']}**",
        f"- Added: **{s['added']}**",
        f"- Removed: **{s['removed']}**",
        f"- Moved slides: **{s['moved']}**",
        f"- Modified: **{s['modified']}**",
        f"- Ignored page-number-only changes: **{s['ignored_page_number_only_changes']}**",
        "",
        "## Added Slides",
        "",
    ]

    if diff["added_slides"]:
        for item in diff["added_slides"]:
            lines.append(f"- `+ 新 P{item['new_position']}` ｜ {item.get('text_preview') or '(no text)'}")
    else:
        lines.append("- None")

    lines += ["", "## Removed Slides", ""]
    if diff["removed_slides"]:
        for item in diff["removed_slides"]:
            lines.append(f"- `- 旧 P{item['old_position']}` ｜ {item.get('text_preview') or '(no text)'}")
    else:
        lines.append("- None")

    lines += ["", "## Moved Slides", ""]
    if diff["moved_groups"]:
        for g in diff["moved_groups"]:
            lines.append(f"- {_format_moved_group(g)}")
    else:
        lines.append("- None")

    lines += ["", "## Modified Slides", ""]
    if diff["modified_slides"]:
        for item in diff["modified_slides"]:
            labels = "、".join(_change_labels(item["changed"]))
            lines += [f"### 旧 P{item['old_position']} → 新 P{item['new_position']}", "", f"- Changed: **{labels}**", f"- Match method: `{item['match_method']}` ｜ similarity `{item['similarity']}`"]
            if item.get("text_changes"):
                lines += ["", "Text changes:"]
                for c in item["text_changes"]:
                    if c["type"] == "changed_segment":
                        lines.append(f"- {c.get('kind', '短句内修改')}：")
                        lines.append(f"  - 旧：{c['old']}")
                        lines.append(f"  - 新：{c['new']}")
                        change = _format_char_changes_md(c.get("char_changes", []))
                        if change:
                            lines.append(f"  - 变化：{change}")
                    elif c["type"] == "removed_segment":
                        lines.append(f"- 删除短句：{c['old']}")
                    elif c["type"] == "added_segment":
                        lines.append(f"- 新增短句：{c['new']}")
            lines.append("")
    else:
        lines.append("- None")

    lines += ["", "## Notes", "", "- This is the Windows MVP based on v0.9.", "- Format/layout detection is optional and off by default.", ""]
    return "\n".join(lines)


def _html_text_diff(old: str, new: str) -> Tuple[str, str]:
    sm = difflib.SequenceMatcher(None, old, new)
    old_parts, new_parts = [], []
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        old_seg = html.escape(old[i1:i2])
        new_seg = html.escape(new[j1:j2])
        if tag == "equal":
            old_parts.append(old_seg)
            new_parts.append(new_seg)
        elif tag == "delete":
            old_parts.append(f'<span class="diff-del">{old_seg}</span>')
        elif tag == "insert":
            new_parts.append(f'<span class="diff-add">{new_seg}</span>')
        elif tag == "replace":
            old_parts.append(f'<span class="diff-del">{old_seg}</span>')
            new_parts.append(f'<span class="diff-add">{new_seg}</span>')
    return "".join(old_parts), "".join(new_parts)


def render_html(diff: Dict) -> str:
    s = diff["summary"]
    title = f"PPT Diff Report - {Path(diff['from_file']).name} → {Path(diff['to_file']).name}"
    css = """
body{font-family:Segoe UI,Microsoft YaHei,Arial,sans-serif;background:#f6f7fb;color:#1f2937;margin:0}
.container{max-width:1160px;margin:0 auto;padding:28px}
h1{font-size:24px;margin:0 0 6px}.meta{color:#667085;font-size:13px;margin-bottom:18px}
.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(120px,1fr));gap:12px;margin:16px 0 24px}
.card{background:white;border:1px solid #d9dee8;border-radius:12px;padding:14px}.label{font-size:12px;color:#667085}.value{font-size:26px;font-weight:700}
.section,details{background:white;border:1px solid #d9dee8;border-radius:12px;margin:12px 0;padding:14px}
summary{cursor:pointer;font-weight:700}li{margin:6px 0}.badge{display:inline-block;border-radius:999px;padding:2px 8px;font-size:12px;background:#eef2f6;margin-right:4px}
.add{color:#0f7b3f}.del{color:#b42318}.move{color:#175cd3}.mod{color:#854a0e}
.change{border-left:4px solid #854a0e;background:#fff8ed;padding:10px;border-radius:8px;margin:10px 0}
.oldnew{display:grid;grid-template-columns:1fr 1fr;gap:12px}.textbox{background:white;border:1px solid #d9dee8;border-radius:8px;padding:10px;word-break:break-word}
.tag{display:block;color:#667085;font-size:12px;margin-bottom:4px}.diff-del{background:#fdebea;color:#b42318;text-decoration:line-through}.diff-add{background:#e8f7ee;color:#0f7b3f}
.empty{color:#667085}
@media(max-width:760px){.oldnew{grid-template-columns:1fr}}
"""
    parts = [f"<!doctype html><html><head><meta charset='utf-8'><title>{html.escape(title)}</title><style>{css}</style></head><body><div class='container'>"]
    parts.append(f"<h1>{html.escape(title)}</h1>")
    parts.append(f"<div class='meta'>Generated at: {html.escape(diff['generated_at'])} ｜ Schema: {html.escape(diff['schema_version'])}</div>")
    cards = [
        ("Old", s["old_slide_count"], ""),
        ("New", s["new_slide_count"], ""),
        ("Added", s["added"], "add"),
        ("Removed", s["removed"], "del"),
        ("Moved", s["moved"], "move"),
        ("Modified", s["modified"], "mod"),
        ("Ignored page #", s["ignored_page_number_only_changes"], ""),
    ]
    parts.append("<div class='grid'>")
    for label, value, cls in cards:
        parts.append(f"<div class='card {cls}'><div class='label'>{label}</div><div class='value'>{value}</div></div>")
    parts.append("</div>")

    parts.append("<h2>Added Slides</h2><div class='section'>")
    if diff["added_slides"]:
        parts.append("<ul>")
        for item in diff["added_slides"]:
            parts.append(f"<li><b class='add'>+ 新 P{item['new_position']}</b> ｜ {html.escape(item.get('text_preview') or '(no text)')}</li>")
        parts.append("</ul>")
    else:
        parts.append("<div class='empty'>None</div>")
    parts.append("</div>")

    parts.append("<h2>Removed Slides</h2><div class='section'>")
    if diff["removed_slides"]:
        parts.append("<ul>")
        for item in diff["removed_slides"]:
            parts.append(f"<li><b class='del'>- 旧 P{item['old_position']}</b> ｜ {html.escape(item.get('text_preview') or '(no text)')}</li>")
        parts.append("</ul>")
    else:
        parts.append("<div class='empty'>None</div>")
    parts.append("</div>")

    parts.append("<h2>Moved Slides</h2><div class='section'>")
    if diff["moved_groups"]:
        parts.append("<ul>")
        for g in diff["moved_groups"]:
            parts.append(f"<li><b class='move'>Moved</b> ｜ {html.escape(_format_moved_group(g))}</li>")
        parts.append("</ul>")
    else:
        parts.append("<div class='empty'>None</div>")
    parts.append("</div>")

    parts.append("<h2>Modified Slides</h2>")
    if diff["modified_slides"]:
        for item in diff["modified_slides"]:
            labels = "、".join(_change_labels(item["changed"]))
            parts.append("<details open>")
            parts.append(f"<summary>旧 P{item['old_position']} → 新 P{item['new_position']} ｜ {html.escape(labels)}</summary>")
            parts.append("<div>")
            parts.append(f"<p>Match: {html.escape(item['match_method'])} ｜ Similarity: {item['similarity']}</p>")
            if item.get("text_changes"):
                for c in item["text_changes"]:
                    parts.append("<div class='change'>")
                    parts.append(f"<b>{html.escape(c.get('kind', '文字变化'))}</b>")
                    if c["type"] == "changed_segment":
                        old_h, new_h = _html_text_diff(c["old"], c["new"])
                        parts.append("<div class='oldnew'>")
                        parts.append(f"<div class='textbox'><span class='tag'>旧</span>{old_h}</div>")
                        parts.append(f"<div class='textbox'><span class='tag'>新</span>{new_h}</div>")
                        parts.append("</div>")
                    elif c["type"] == "removed_segment":
                        parts.append(f"<div class='textbox'><span class='tag'>删除</span><span class='diff-del'>{html.escape(c['old'])}</span></div>")
                    elif c["type"] == "added_segment":
                        parts.append(f"<div class='textbox'><span class='tag'>新增</span><span class='diff-add'>{html.escape(c['new'])}</span></div>")
                    parts.append("</div>")
            else:
                parts.append("<p class='empty'>该页变化不是文字变化，或仅检测到图片/结构/格式变化。</p>")
            parts.append("</div></details>")
    else:
        parts.append("<div class='section empty'>None</div>")

    parts.append("</div></body></html>")
    return "\n".join(parts)


def write_outputs(diff: Dict, output_dir: Path) -> Tuple[Path, Path, Path]:
    output_dir.mkdir(parents=True, exist_ok=True)
    base = safe_filename(f"{Path(diff['from_file']).stem}__TO__{Path(diff['to_file']).stem}")
    html_path = output_dir / f"{base}.diff.html"
    md_path = output_dir / f"{base}.diff.md"
    json_path = output_dir / f"{base}.diff.json"
    html_path.write_text(render_html(diff), encoding="utf-8")
    md_path.write_text(render_markdown(diff), encoding="utf-8")
    json_path.write_text(json.dumps(diff, ensure_ascii=False, indent=2), encoding="utf-8")
    return html_path, md_path, json_path


def run_compare(old_pptx: Path, new_pptx: Path, output_dir: Path, detect_format: bool = False):
    diff = build_diff(old_pptx, new_pptx, detect_format=detect_format)
    html_path, md_path, json_path = write_outputs(diff, output_dir)
    return diff, html_path, md_path, json_path


def open_path(path: Optional[Path]) -> None:
    if not path:
        return
    try:
        if sys.platform.startswith("win"):
            os.startfile(str(path))  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            os.system(f'open "{path}"')
        else:
            os.system(f'xdg-open "{path}"')
    except Exception:
        pass


def launch_ui() -> None:
    import tkinter as tk
    from tkinter import filedialog, messagebox
    from tkinter.scrolledtext import ScrolledText

    root = tk.Tk()
    root.title("PPT Diff Tool v0.9 Windows MVP")
    root.geometry("860x600")

    old_var = tk.StringVar()
    new_var = tk.StringVar()
    out_var = tk.StringVar(value=str(Path.cwd() / "diff_output"))
    detect_format_var = tk.BooleanVar(value=False)

    last_html = {"path": None}
    last_md = {"path": None}
    last_out = {"path": None}

    def browse_old():
        p = filedialog.askopenfilename(title="选择旧版 PPT", filetypes=[("PowerPoint .pptx", "*.pptx")])
        if p:
            old_var.set(p)
            if out_var.get().endswith("diff_output"):
                out_var.set(str(Path(p).parent / "diff_output"))

    def browse_new():
        p = filedialog.askopenfilename(title="选择新版 PPT", filetypes=[("PowerPoint .pptx", "*.pptx")])
        if p:
            new_var.set(p)
            if out_var.get().endswith("diff_output"):
                out_var.set(str(Path(p).parent / "diff_output"))

    def browse_out():
        p = filedialog.askdirectory(title="选择输出文件夹")
        if p:
            out_var.set(p)

    def log(msg: str):
        output.insert(tk.END, msg + "\n")
        output.see(tk.END)
        root.update_idletasks()

    def validate():
        if not old_var.get().strip():
            messagebox.showerror("缺少文件", "请选择旧版 PPT。")
            return None
        if not new_var.get().strip():
            messagebox.showerror("缺少文件", "请选择新版 PPT。")
            return None
        old_p = Path(old_var.get().strip())
        new_p = Path(new_var.get().strip())
        out_p = Path(out_var.get().strip() or "diff_output")
        if not old_p.exists():
            messagebox.showerror("文件不存在", f"旧版 PPT 不存在：\n{old_p}")
            return None
        if not new_p.exists():
            messagebox.showerror("文件不存在", f"新版 PPT 不存在：\n{new_p}")
            return None
        if old_p.suffix.lower() != ".pptx" or new_p.suffix.lower() != ".pptx":
            messagebox.showerror("格式不支持", "当前只支持 .pptx。请先另存为 .pptx。")
            return None
        return old_p, new_p, out_p

    def do_compare():
        vals = validate()
        if not vals:
            return
        old_p, new_p, out_p = vals
        btn_compare.config(state=tk.DISABLED)
        btn_open_html.config(state=tk.DISABLED)
        btn_open_md.config(state=tk.DISABLED)
        btn_open_folder.config(state=tk.DISABLED)
        output.delete("1.0", tk.END)
        log("开始比较...")
        log(f"旧版：{old_p.name}")
        log(f"新版：{new_p.name}")
        log(f"输出：{out_p}")
        log(f"检测格式/布局变化：{detect_format_var.get()}")

        def worker():
            try:
                diff, html_path, md_path, json_path = run_compare(old_p, new_p, out_p, detect_format=detect_format_var.get())
                root.after(0, lambda: finish_success(diff, html_path, md_path, json_path, out_p))
            except Exception:
                tb = traceback.format_exc()
                root.after(0, lambda: finish_error(tb))

        threading.Thread(target=worker, daemon=True).start()

    def finish_success(diff, html_path, md_path, json_path, out_p):
        log("")
        log("完成。")
        log(f"HTML：{html_path}")
        log(f"Markdown：{md_path}")
        log(f"JSON：{json_path}")
        log("")
        log("Summary:")
        for k, v in diff["summary"].items():
            log(f"  {k}: {v}")

        last_html["path"] = html_path
        last_md["path"] = md_path
        last_out["path"] = out_p
        btn_open_html.config(state=tk.NORMAL)
        btn_open_md.config(state=tk.NORMAL)
        btn_open_folder.config(state=tk.NORMAL)
        btn_compare.config(state=tk.NORMAL)
        open_path(html_path)

    def finish_error(tb: str):
        log("")
        log("出错：")
        log(tb)
        msg = tb[-1600:]
        if "只支持 .pptx" in tb:
            msg = "当前只支持 .pptx，不支持 .ppt。请先另存为 .pptx。"
        elif "文件不存在" in tb:
            msg = "文件不存在，请重新选择文件。"
        elif "BadZipFile" in tb:
            msg = "无法打开 PPTX。文件可能损坏，或不是标准 .pptx。"
        messagebox.showerror("比较失败", msg)
        btn_compare.config(state=tk.NORMAL)

    frm = tk.Frame(root, padx=14, pady=14)
    frm.pack(fill=tk.BOTH, expand=True)

    tk.Label(frm, text="旧版 PPT（File A）").grid(row=0, column=0, sticky="w")
    tk.Entry(frm, textvariable=old_var, width=90).grid(row=1, column=0, sticky="we", padx=(0, 8))
    tk.Button(frm, text="选择...", command=browse_old).grid(row=1, column=1)

    tk.Label(frm, text="新版 PPT（File B）").grid(row=2, column=0, sticky="w", pady=(10, 0))
    tk.Entry(frm, textvariable=new_var, width=90).grid(row=3, column=0, sticky="we", padx=(0, 8))
    tk.Button(frm, text="选择...", command=browse_new).grid(row=3, column=1)

    tk.Label(frm, text="输出文件夹").grid(row=4, column=0, sticky="w", pady=(10, 0))
    tk.Entry(frm, textvariable=out_var, width=90).grid(row=5, column=0, sticky="we", padx=(0, 8))
    tk.Button(frm, text="选择...", command=browse_out).grid(row=5, column=1)

    tk.Checkbutton(frm, text="检测格式/布局变化（更敏感，默认关闭）", variable=detect_format_var).grid(row=6, column=0, sticky="w", pady=(12, 0))

    btn_frame = tk.Frame(frm)
    btn_frame.grid(row=7, column=0, columnspan=2, sticky="w", pady=(12, 8))
    btn_compare = tk.Button(btn_frame, text="比较两个 PPT", width=16, command=do_compare)
    btn_compare.pack(side=tk.LEFT)

    btn_open_html = tk.Button(btn_frame, text="打开 HTML 报告", width=16, state=tk.DISABLED, command=lambda: open_path(last_html["path"]))
    btn_open_html.pack(side=tk.LEFT, padx=(8, 0))

    btn_open_md = tk.Button(btn_frame, text="打开 Markdown 报告", width=18, state=tk.DISABLED, command=lambda: open_path(last_md["path"]))
    btn_open_md.pack(side=tk.LEFT, padx=(8, 0))

    btn_open_folder = tk.Button(btn_frame, text="打开输出文件夹", width=16, state=tk.DISABLED, command=lambda: open_path(last_out["path"]))
    btn_open_folder.pack(side=tk.LEFT, padx=(8, 0))

    output = ScrolledText(frm, height=18)
    output.grid(row=8, column=0, columnspan=2, sticky="nsew")

    frm.columnconfigure(0, weight=1)
    frm.rowconfigure(8, weight=1)

    log("请选择旧版 PPT、新版 PPT，然后点击“比较两个 PPT”。")
    log("Windows MVP：基于 v0.9；输出 HTML / Markdown / JSON；保留 Markdown 打开按钮。")
    root.mainloop()


def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(description="Compare two .pptx files and generate HTML/Markdown/JSON diff reports.")
    parser.add_argument("old_pptx", nargs="?", type=Path)
    parser.add_argument("new_pptx", nargs="?", type=Path)
    parser.add_argument("-o", "--output-dir", type=Path, default=Path("diff_output"))
    parser.add_argument("--detect-format", action="store_true")
    parser.add_argument("--ui", action="store_true")

    args = parser.parse_args(argv)

    if args.ui or (args.old_pptx is None and args.new_pptx is None):
        launch_ui()
        return 0

    if args.old_pptx is None or args.new_pptx is None:
        parser.error("old_pptx and new_pptx are required unless --ui is used.")

    try:
        diff, html_path, md_path, json_path = run_compare(args.old_pptx, args.new_pptx, args.output_dir, detect_format=args.detect_format)
        print("Done.")
        print(f"HTML:     {html_path}")
        print(f"Markdown: {md_path}")
        print(f"JSON:     {json_path}")
        print("")
        print("Summary:")
        for k, v in diff["summary"].items():
            print(f"  {k}: {v}")
        return 0
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
