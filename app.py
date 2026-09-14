from __future__ import annotations

import csv
import base64
import html
import urllib.request
import urllib.error
import urllib.parse
import uuid
import io
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
import zipfile
import traceback
import logging
import xml.etree.ElementTree as ET
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import List, Tuple, Optional

from flask import Flask, render_template, request, send_file, jsonify
import pymupdf as fitz  # PyMuPDF
from PIL import Image, ImageOps

app = Flask(__name__)

# Runtime log for errors that occur after the server starts.
LOG_DIR = Path(__file__).resolve().parent / 'logs'
LOG_DIR.mkdir(exist_ok=True)
logging.basicConfig(filename=str(LOG_DIR / '运行日志.txt'), level=logging.INFO, encoding='utf-8', format='%(asctime)s %(levelname)s %(message)s')
app.config['MAX_CONTENT_LENGTH'] = 1024 * 1024 * 1024  # 1 GB

WORD_EXTS = {'.doc', '.docx'}
CAPTION_RE_STRICT = re.compile(r'^\s*图\s*([0-9０-９]+)(?:\s+|[：:、.．\-—]\s*)(.{1,60}?)\s*$')
CAPTION_RE_LOOSE = re.compile(r'^\s*图\s*([0-9０-９]+)(.{1,40}?)\s*$')
INVALID_FS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


@dataclass
class FigureInfo:
    figure_no: str
    caption: str
    page: int
    filename: str
    crop_bbox: List[float]
    extraction_method: str = 'unknown'
    confidence: str = 'unknown'
    source_object: str = ''


@dataclass
class ArticleInfo:
    source_file: str
    first_author: str
    title: str
    article_folder: str
    figures: List[FigureInfo]
    warnings: List[str]


def sanitize_filename(s: str, limit: int = 80) -> str:
    s = INVALID_FS.sub('_', s).strip().strip('.')
    s = re.sub(r'\s+', ' ', s)
    return (s[:limit] or '未命名').strip()


def normalize_digits(s: str) -> str:
    table = str.maketrans('０１２３４５６７８９', '0123456789')
    return s.translate(table)


def prepare_word_document(src: Path, out_pdf: Path, normalized_docx: Path) -> List[str]:
    """Render with Microsoft Word and, when possible, create a temporary DOCX copy.

    The normalized DOCX is used only to inspect the real embedded figure object immediately
    before each caption.  This is far more reliable than guessing a crop from the PDF page.
    The source document is never modified.
    """
    warnings: List[str] = []
    if sys.platform.startswith('win'):
        try:
            import pythoncom
            import win32com.client
            pythoncom.CoInitialize()
            word = win32com.client.DispatchEx('Word.Application')
            word.Visible = False
            word.DisplayAlerts = 0
            doc = None
            try:
                doc = word.Documents.Open(str(src.resolve()), ReadOnly=True, AddToRecentFiles=False)
                doc.ExportAsFixedFormat(str(out_pdf.resolve()), 17)  # wdExportFormatPDF
                if src.suffix.lower() == '.docx':
                    shutil.copy2(src, normalized_docx)
                else:
                    # wdFormatDocumentDefault = 16 (.docx). SaveAs2 writes only the temp copy.
                    try:
                        doc.SaveAs2(str(normalized_docx.resolve()), FileFormat=16, AddToRecentFiles=False)
                    except Exception as exc:
                        warnings.append(f'临时 DOC→DOCX 归一化失败，将仅使用 PDF 回退：{exc}')
            finally:
                if doc is not None:
                    doc.Close(False)
                word.Quit()
                pythoncom.CoUninitialize()
            if out_pdf.exists() and out_pdf.stat().st_size > 0:
                return warnings
        except Exception as e:
            raise RuntimeError(f'Microsoft Word 转 PDF 失败：{e}')

    # Non-Windows developer/test fallback.
    soffice = shutil.which('soffice') or shutil.which('libreoffice')
    if soffice:
        result = subprocess.run(
            [soffice, '--headless', '--convert-to', 'pdf', '--outdir', str(out_pdf.parent), str(src)],
            stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, timeout=120
        )
        generated = out_pdf.parent / (src.stem + '.pdf')
        if generated.exists():
            if generated != out_pdf:
                generated.replace(out_pdf)
            if src.suffix.lower() == '.docx':
                shutil.copy2(src, normalized_docx)
            return warnings
        raise RuntimeError('LibreOffice 转 PDF 失败：' + (result.stderr or result.stdout))

    raise RuntimeError('未检测到可用的 Microsoft Word（Windows）或 LibreOffice。')


def clean_line(line: str) -> str:
    line = line.replace('\u3000', ' ').replace('\xa0', ' ')
    line = re.sub(r'\s+', ' ', line).strip()
    return line


def looks_like_author_line(line: str) -> Optional[List[str]]:
    # Remove common affiliation/superscript markers but keep Chinese names.
    if not line or len(line) > 60:
        return None
    if any(k in line for k in ['摘要', '关键词', '大学', '学院', '公司', '研究院', '实验室', '作者简介', '收稿日期']):
        return None
    raw = re.sub(r'[0-9０-９*＊†‡#＃\s]+', '', line)
    parts = [p for p in re.split(r'[，,、;；]+', raw) if p]
    if not (1 <= len(parts) <= 12):
        return None
    cleaned = []
    for p in parts:
        p = p.strip('()（）[]【】^')
        # Common Chinese names: 2-4 Han chars. Allow middle dot for minority names.
        if re.fullmatch(r'[\u3400-\u9fff·]{2,6}', p):
            cleaned.append(p)
        else:
            return None
    return cleaned if cleaned else None


def detect_title_author(pdf: fitz.Document) -> Tuple[str, str, List[str]]:
    warnings = []
    if len(pdf) == 0:
        return '未知作者', '未识别标题', ['PDF 无页面']
    text = pdf[0].get_text('text')
    lines = [clean_line(x) for x in text.splitlines() if clean_line(x)]

    # Keep only the early front-matter region, preferably before 摘要.
    pre = []
    for line in lines[:80]:
        if line.startswith('摘要') or line.startswith('摘 要'):
            break
        pre.append(line)

    author_idx = None
    author_names = None
    for i, line in enumerate(pre[:45]):
        names = looks_like_author_line(line)
        if names:
            # Prefer a multi-author line, but single author is allowed.
            author_idx, author_names = i, names
            if len(names) >= 2:
                break

    if not author_names:
        warnings.append('未可靠识别正文作者，已标记为“未知作者”')
        first_author = '未知作者'
        author_idx = min(8, len(pre))
    else:
        first_author = author_names[0]

    title = None
    skip_words = ['收稿日期', '基金项目', '第', '现代信息科技', 'ISSN', 'DOI']
    if author_idx is not None:
        for j in range(author_idx - 1, -1, -1):
            cand = pre[j].strip()
            if not (5 <= len(cand) <= 100):
                continue
            if any(cand.startswith(k) for k in skip_words):
                continue
            if re.fullmatch(r'[0-9\-—–./年月日 :：]+', cand):
                continue
            # Avoid obvious header/footer fragments.
            if '收稿' in cand or '作者简介' in cand:
                continue
            title = cand
            break

    if not title:
        warnings.append('未可靠识别正文标题')
        title = '未识别标题'

    return first_author, title, warnings


def text_line_entries(page: fitz.Page):
    entries = []
    d = page.get_text('dict')
    for block in d.get('blocks', []):
        if block.get('type') != 0:
            continue
        for line in block.get('lines', []):
            spans = line.get('spans', [])
            if not spans:
                continue
            text = ''.join(s.get('text', '') for s in spans)
            text = clean_line(text)
            if not text:
                continue
            rect = fitz.Rect(spans[0]['bbox'])
            sizes = []
            for sp in spans:
                rect |= fitz.Rect(sp['bbox'])
                if sp.get('size'):
                    sizes.append(float(sp['size']))
            entries.append({'text': text, 'bbox': rect, 'font_size': sum(sizes)/len(sizes) if sizes else 0})
    return entries


def find_captions(pdf: fitz.Document):
    """Locate *standalone* figure captions, not prose references such as “图4展示了…”."""
    found = []
    prose_starts = (
        '所示', '中', '可见', '展示', '显示', '给出', '说明', '表明', '为', '是',
        '采用', '包含', '包括', '反映', '描述', '用于', '对应'
    )
    for pno in range(len(pdf)):
        page = pdf[pno]
        page_rect = page.rect
        for ent in text_line_entries(page):
            txt = ent['text']
            m = CAPTION_RE_STRICT.match(txt)
            strict = m is not None
            if not m:
                # Compatibility fallback for captions typed as “图1总体框架图”.
                # Only accept short, caption-like lines around the central area.
                m = CAPTION_RE_LOOSE.match(txt)
                if not m:
                    continue
            no = normalize_digits(m.group(1))
            rest = clean_line(m.group(2).lstrip('：:、.．-— '))
            if not rest or len(rest) > 60:
                continue
            # Most false positives are ordinary prose beginning with “图N展示了 / 图N所示…”.
            if rest.startswith(prose_starts):
                continue
            if txt.endswith(('。', '；', ';')) and len(txt) > 24:
                continue
            rect = ent['bbox']
            width_ratio = rect.width / max(page_rect.width, 1)
            center_offset = abs((rect.x0 + rect.x1) / 2 - page_rect.width / 2) / max(page_rect.width, 1)
            if width_ratio > 0.72:
                continue
            if not strict and (len(rest) > 24 or center_offset > 0.24):
                continue
            found.append({'page': pno, 'no': no, 'caption': rest, 'bbox': rect, 'text': txt})
    # De-duplicate PDF text fragmentation.
    dedup = []
    for item in found:
        if any(x['page'] == item['page'] and x['no'] == item['no'] and abs(x['bbox'].y0-item['bbox'].y0) < 4 for x in dedup):
            continue
        dedup.append(item)
    return dedup

# OOXML namespaces used for object-first figure extraction.
OOXML_NS = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'v': 'urn:schemas-microsoft-com:vml',
}
REL_NS = {'rel': 'http://schemas.openxmlformats.org/package/2006/relationships'}
RASTER_EXTS = {'.png', '.jpg', '.jpeg', '.bmp', '.gif', '.tif', '.tiff'}
VECTOR_EXTS = {'.emf', '.wmf'}


def _caption_parts(text: str) -> Optional[Tuple[str, str]]:
    """Parse a caption-like paragraph and reject prose references such as 图4展示了……"""
    text = clean_line(text.replace('\u200b', ''))
    prose_starts = (
        '所示', '中', '可见', '展示', '显示', '给出', '说明', '表明', '为', '是',
        '采用', '包含', '包括', '反映', '描述', '用于', '对应'
    )
    m = CAPTION_RE_STRICT.match(text)
    strict = m is not None
    if not m:
        m = CAPTION_RE_LOOSE.match(text)
        if not m:
            return None
    no = normalize_digits(m.group(1))
    rest = clean_line(m.group(2).lstrip('：:、.．-— '))
    if not rest or len(rest) > 60 or rest.startswith(prose_starts):
        return None
    if text.endswith(('。', '；', ';')) and len(text) > 24:
        return None
    if not strict and len(rest) > 24:
        return None
    return no, rest


def _paragraph_text_xml(p) -> str:
    return clean_line(''.join((t.text or '') for t in p.findall('.//w:t', OOXML_NS)).replace('\u200b', ''))


def _paragraph_media_targets(p, rels: dict) -> List[str]:
    targets = []
    for blip in p.findall('.//a:blip', OOXML_NS):
        rid = blip.attrib.get('{' + OOXML_NS['r'] + '}embed')
        target = rels.get(rid)
        if target and target.lower().startswith('media/'):
            targets.append(target)
    for im in p.findall('.//v:imagedata', OOXML_NS):
        rid = im.attrib.get('{' + OOXML_NS['r'] + '}id')
        target = rels.get(rid)
        if target and target.lower().startswith('media/'):
            targets.append(target)
    # Preserve order but remove duplicates.
    return list(dict.fromkeys(targets))


def find_docx_caption_media(docx_path: Path) -> List[dict]:
    """Find the real Word image/OLE preview paragraph immediately before each figure caption.

    Word stores Visio / equation / OLE objects with a preview image (often EMF/WMF).  The
    figure preview immediately before a caption is a much stronger signal than PDF geometry.
    """
    if not docx_path.exists():
        return []
    try:
        with zipfile.ZipFile(docx_path) as z:
            root = ET.fromstring(z.read('word/document.xml'))
            relroot = ET.fromstring(z.read('word/_rels/document.xml.rels'))
            rels = {
                rel.attrib.get('Id'): rel.attrib.get('Target', '')
                for rel in relroot.findall('.//rel:Relationship', REL_NS)
            }
            paragraphs = root.findall('.//w:body//w:p', OOXML_NS)
            results = []
            for i, p in enumerate(paragraphs):
                cap = _caption_parts(_paragraph_text_xml(p))
                if not cap:
                    continue
                no, caption = cap
                chosen = None
                # Normally the preceding paragraph is exactly the figure object.  Ignore only
                # truly blank spacer paragraphs; never jump across prose, which would risk
                # selecting an equation or an earlier unrelated object.
                scan_indices = [i] + list(range(i - 1, max(-1, i - 5), -1))
                for j in scan_indices:
                    q = paragraphs[j]
                    refs = _paragraph_media_targets(q, rels)
                    qtext = _paragraph_text_xml(q)
                    if refs:
                        refs2 = []
                        for target in refs:
                            member = 'word/' + target.lstrip('/')
                            try:
                                size = z.getinfo(member).file_size
                            except KeyError:
                                size = 0
                            refs2.append((size, target, member))
                        refs2.sort(reverse=True)
                        if refs2:
                            _, target, member = refs2[0]
                            chosen = {
                                'figure_no': no,
                                'caption': caption,
                                'target': target,
                                'member': member,
                                'ext': Path(target).suffix.lower(),
                                'paragraph_index': i,
                                'object_paragraph_index': j,
                            }
                        break
                    if qtext and j != i:
                        break
                if chosen:
                    results.append(chosen)
            return results
    except Exception:
        logging.exception('解析 DOCX 图对象关系失败: %s', docx_path)
        return []


def trim_and_validate_png(path: Path, padding: int = 8) -> bool:
    """Trim outer white margins and reject genuinely blank / tiny results."""
    try:
        with Image.open(path) as im0:
            if 'A' in im0.getbands():
                rgba = im0.convert('RGBA')
                bg = Image.new('RGBA', rgba.size, (255, 255, 255, 255))
                bg.alpha_composite(rgba)
                im = bg.convert('RGB')
            else:
                im = im0.convert('RGB')
            gray = ImageOps.grayscale(im)
            # Anything darker than near-white counts as visible content.
            mask = gray.point(lambda p: 255 if p < 250 else 0)
            bbox = mask.getbbox()
            if not bbox:
                return False
            x0, y0, x1, y1 = bbox
            if (x1 - x0) < 35 or (y1 - y0) < 25:
                return False
            x0 = max(0, x0 - padding); y0 = max(0, y0 - padding)
            x1 = min(im.width, x1 + padding); y1 = min(im.height, y1 + padding)
            cropped = im.crop((x0, y0, x1, y1))
            cropped.save(path, 'PNG')
            return True
    except Exception:
        logging.exception('PNG 裁白/有效性检查失败: %s', path)
        return False


def render_vector_media_via_word(media_path: Path, out_png: Path, work_dir: Path) -> bool:
    """Render an exact EMF/WMF preview through Word, then trim the white page."""
    if not sys.platform.startswith('win'):
        return False
    pdf_path = work_dir / (media_path.stem + '_vector_render.pdf')
    raw_png = work_dir / (media_path.stem + '_vector_render.png')
    try:
        import pythoncom
        import win32com.client
        pythoncom.CoInitialize()
        word = win32com.client.DispatchEx('Word.Application')
        word.Visible = False
        word.DisplayAlerts = 0
        doc = None
        try:
            doc = word.Documents.Add()
            # Small margins maximize usable page area.
            doc.PageSetup.TopMargin = 18
            doc.PageSetup.BottomMargin = 18
            doc.PageSetup.LeftMargin = 18
            doc.PageSetup.RightMargin = 18
            rng = doc.Range(0, 0)
            ils = doc.InlineShapes.AddPicture(str(media_path.resolve()), False, True, rng)
            max_w, max_h = 540.0, 760.0
            w, h = float(ils.Width), float(ils.Height)
            scale = min(1.0, max_w / max(w, 1.0), max_h / max(h, 1.0))
            if scale < 1.0:
                ils.Width = w * scale
                ils.Height = h * scale
            doc.ExportAsFixedFormat(str(pdf_path.resolve()), 17)
        finally:
            if doc is not None:
                doc.Close(False)
            word.Quit()
            pythoncom.CoUninitialize()
        if not pdf_path.exists():
            return False
        d = fitz.open(pdf_path)
        if len(d) < 1:
            d.close(); return False
        pix = d[0].get_pixmap(matrix=fitz.Matrix(3.0, 3.0), alpha=False)
        pix.save(str(raw_png))
        d.close()
        shutil.copy2(raw_png, out_png)
        return trim_and_validate_png(out_png, padding=10)
    except Exception:
        logging.exception('Word 渲染 EMF/WMF 失败: %s', media_path)
        return False
    finally:
        for q in (pdf_path, raw_png):
            try:
                q.unlink()
            except Exception:
                pass


def extract_docx_media_as_png(docx_path: Path, media: dict, out_png: Path, work_dir: Path) -> Tuple[bool, str]:
    """Extract the exact preview object tied to a caption, rasterizing vector previews if needed."""
    try:
        with zipfile.ZipFile(docx_path) as z:
            data = z.read(media['member'])
    except Exception:
        logging.exception('读取 DOCX 媒体失败: %s %s', docx_path, media)
        return False, 'docx_media_read_failed'

    ext = media.get('ext', '').lower()
    media_file = work_dir / f"direct_{media['figure_no']}_{abs(hash(media['member']))}{ext or '.bin'}"
    media_file.write_bytes(data)
    try:
        if ext in RASTER_EXTS:
            try:
                with Image.open(media_file) as im:
                    im = im.convert('RGB')
                    im.save(out_png, 'PNG')
                if trim_and_validate_png(out_png, padding=8):
                    return True, 'docx_exact_raster'
            except Exception:
                logging.exception('直接读取 DOCX 栅格图失败: %s', media_file)
        elif ext in VECTOR_EXTS:
            if render_vector_media_via_word(media_file, out_png, work_dir):
                return True, 'docx_exact_vector'
        # Unknown media types get one final Word-render attempt.
        elif render_vector_media_via_word(media_file, out_png, work_dir):
            return True, 'docx_exact_word_render'
    finally:
        try:
            media_file.unlink()
        except Exception:
            pass
    try:
        out_png.unlink()
    except Exception:
        pass
    return False, 'docx_media_unrenderable'


def _body_ceiling_y(page: fitz.Page, cap_rect: fitz.Rect) -> Optional[float]:
    """Bottom of the nearest clear body-text line above the caption.

    Used only as a safety clamp for PDF fallbacks so vector clustering cannot swallow an
    explanatory paragraph above a figure (the v0.7 图2 failure mode).
    """
    page_rect = page.rect
    candidates = []
    for ent in text_line_entries(page):
        r = ent['bbox']
        if r.y1 >= cap_rect.y0 - 6:
            continue
        t = re.sub(r'\s+', '', ent['text'])
        width_ratio = r.width / max(page_rect.width, 1)
        if len(t) >= 24 and width_ratio >= 0.43:
            candidates.append(r.y1)
    return max(candidates) if candidates else None



def is_body_like(text: str, rect: fitz.Rect, page_rect: fitz.Rect) -> bool:
    t = re.sub(r'\s+', '', text)
    width_ratio = rect.width / max(page_rect.width, 1)
    # Long full-width paragraph lines are likely body text, not labels inside a figure.
    return len(t) >= 28 and width_ratio >= 0.48


def _clip_rect(r: fitz.Rect, page_rect: fitz.Rect, cap_y0: float) -> fitz.Rect:
    return fitz.Rect(
        max(page_rect.x0, r.x0), max(page_rect.y0, r.y0),
        min(page_rect.x1, r.x1), min(cap_y0 - 0.8, r.y1)
    )


def _nearby_raster_bbox(page: fitz.Page, cap_rect: fitz.Rect) -> Optional[fitz.Rect]:
    """Prefer the actual embedded image immediately above a caption.

    Word-exported PDF preserves ordinary screenshots, diagrams and many pasted figures as
    one raster rectangle. Cropping that rectangle directly avoids swallowing body text.
    """
    candidates = []
    seen = set()
    try:
        for img in page.get_images(full=True):
            xref = img[0]
            for r0 in page.get_image_rects(xref):
                r = fitz.Rect(r0)
                key = tuple(round(x, 2) for x in r)
                if key in seen:
                    continue
                seen.add(key)
                gap = cap_rect.y0 - r.y1
                # Caption typically touches the image or is separated by a small blank gap.
                if gap < -3 or gap > 30:
                    continue
                if r.width < 55 or r.height < 38 or r.get_area() < 3500:
                    continue
                # Ignore tiny header/logo images spanning implausible locations.
                candidates.append((gap, -r.get_area(), r))
    except Exception:
        return None
    if not candidates:
        return None
    candidates.sort(key=lambda x: (x[0], x[1]))
    return candidates[0][2]


def _vector_cluster_bbox(page: fitz.Page, cap_rect: fitz.Rect) -> Optional[fitz.Rect]:
    """Fallback for Visio / Word shapes exported as PDF vectors.

    Start from vector objects nearest the caption, then absorb nearby vector/text objects.
    This intentionally excludes long body-text lines.
    """
    page_rect = page.rect
    entries = text_line_entries(page)
    items = []
    try:
        for dr in page.get_drawings():
            r = fitz.Rect(dr['rect'])
            if r.y1 <= cap_rect.y0 + 2 and cap_rect.y0 - r.y1 <= 420 and r.width > 1 and r.height > 1:
                items.append(('draw', r))
    except Exception:
        pass
    if not items:
        return None

    # Seed = nearest substantial drawing to the caption.
    substantial = [(cap_rect.y0-r.y1, -r.get_area(), r) for _, r in items if r.width >= 18 or r.height >= 18]
    if not substantial:
        return None
    substantial.sort(key=lambda x:(x[0],x[1]))
    seed = fitz.Rect(substantial[0][2])
    if substantial[0][0] > 55:
        return None

    cluster = fitz.Rect(seed)
    # Iteratively absorb drawings close to the current cluster. Large padding allows
    # disconnected boxes/arrows within one flowchart to become one cluster.
    for _ in range(8):
        old = tuple(cluster)
        expanded = fitz.Rect(cluster.x0-28, cluster.y0-28, cluster.x1+28, cluster.y1+28)
        for _, r in items:
            if r.intersects(expanded):
                cluster |= r
        # Include short text labels that sit inside / immediately around the vector cluster.
        expanded2 = fitz.Rect(cluster.x0-18, cluster.y0-18, cluster.x1+18, cluster.y1+18)
        for ent in entries:
            r = ent['bbox']
            if r.y1 > cap_rect.y0 + 1:
                continue
            if is_body_like(ent['text'], r, page_rect):
                continue
            if r.intersects(expanded2):
                cluster |= r
        if tuple(cluster) == old:
            break

    if cluster.width < 55 or cluster.height < 45 or cluster.get_area() < 3500:
        return None
    # Reject full-width table/body-rule clusters.
    if cluster.width > page_rect.width * 0.92 and cluster.height < 80:
        return None
    return cluster


def _heuristic_region_bbox(page: fitz.Page, cap, prev_caption_y1: Optional[float]) -> fitz.Rect:
    """Last-resort page-region heuristic. Marked low-confidence in the report."""
    page_rect = page.rect
    cap_rect = cap['bbox']
    cap_y0 = cap_rect.y0
    entries = text_line_entries(page)
    body_candidates = []
    for ent in entries:
        r = ent['bbox']
        if r.y1 <= cap_y0 - 3 and is_body_like(ent['text'], r, page_rect):
            body_candidates.append(r)
    top = page_rect.y0 + 18
    if body_candidates:
        nearest = max(body_candidates, key=lambda r: r.y1)
        top = nearest.y1 + 5
    if prev_caption_y1 is not None:
        top = max(top, prev_caption_y1 + 8)
    if cap_y0 - top < 70:
        top = max(page_rect.y0 + 18, cap_y0 - min(page_rect.height * 0.38, 300))
    if cap_y0 - top > page_rect.height * 0.62:
        top = cap_y0 - page_rect.height * 0.55
    return fitz.Rect(page_rect.x0 + 18, top, page_rect.x1 - 18, cap_y0 - 3)


def crop_figure(page: fitz.Page, cap, prev_caption_y1: Optional[float], out_path: Path) -> Tuple[fitz.Rect, str, str]:
    """PDF fallback only. v0.8 first tries the exact Word object before calling this."""
    page_rect = page.rect
    cap_rect = cap['bbox']
    cap_y0 = cap_rect.y0
    ceiling = _body_ceiling_y(page, cap_rect)

    def render_and_check(crop: fitz.Rect) -> bool:
        if crop.width < 20 or crop.height < 20:
            return False
        pix = page.get_pixmap(matrix=fitz.Matrix(3.0, 3.0), clip=crop, alpha=False)
        pix.save(str(out_path))
        return trim_and_validate_png(out_path, padding=8)

    # 1) Real raster immediately above the caption.
    r = _nearby_raster_bbox(page, cap_rect)
    if r is not None:
        pad = 1.5
        crop = _clip_rect(fitz.Rect(r.x0-pad, r.y0-pad, r.x1+pad, r.y1+pad), page_rect, cap_y0)
        # Some Office PDFs contain a blank/white preview layer. Never trust geometry alone.
        if render_and_check(crop):
            return crop, 'pdf_nearby_raster', 'high'
        try:
            out_path.unlink()
        except Exception:
            pass

    # 2) Visio / grouped vector shapes. Clamp its top below the nearest body paragraph so
    # explanatory text above the figure cannot be swallowed into the result.
    r = _vector_cluster_bbox(page, cap_rect)
    if r is not None:
        pad = 7
        crop = _clip_rect(fitz.Rect(r.x0-pad, r.y0-pad, r.x1+pad, r.y1+pad), page_rect, cap_y0)
        if ceiling is not None and cap_y0 - ceiling >= 45 and crop.y0 < ceiling + 3:
            crop.y0 = ceiling + 3
        if render_and_check(crop):
            return crop, 'pdf_vector_cluster', 'medium'
        try:
            out_path.unlink()
        except Exception:
            pass

    # 3) Last-resort broad region; explicitly low confidence.
    crop = _heuristic_region_bbox(page, cap, prev_caption_y1)
    if not render_and_check(crop):
        # Keep a visible diagnostic instead of silently emitting a totally white PNG.
        # A slightly broader region may reveal what Word rendered and is easier to debug.
        crop = fitz.Rect(page_rect.x0 + 12, max(page_rect.y0 + 12, cap_y0 - 320), page_rect.x1 - 12, cap_y0 - 2)
        pix = page.get_pixmap(matrix=fitz.Matrix(3.0, 3.0), clip=crop, alpha=False)
        pix.save(str(out_path))
        trim_and_validate_png(out_path, padding=8)
    return crop, 'pdf_heuristic_fallback', 'low'

def process_one_word(src: Path, out_root: Path, work_dir: Path) -> ArticleInfo:
    warnings = []
    stamp = str(int(time.time()*1000))
    safe_stem = sanitize_filename(src.stem, 60)
    pdf_path = work_dir / (safe_stem + '_' + stamp + '.pdf')
    normalized_docx = work_dir / (safe_stem + '_' + stamp + '_normalized.docx')
    warnings.extend(prepare_word_document(src, pdf_path, normalized_docx))
    pdf = fitz.open(pdf_path)

    first_author, title, detect_warn = detect_title_author(pdf)
    warnings.extend(detect_warn)

    folder_name = sanitize_filename(f'{first_author}_{title}', 90)
    article_dir = out_root / folder_name
    if article_dir.exists():
        i = 2
        while (out_root / f'{folder_name}_{i}').exists():
            i += 1
        article_dir = out_root / f'{folder_name}_{i}'
        warnings.append('出现同名作者+标题，输出目录已自动加序号')
    article_dir.mkdir(parents=True, exist_ok=True)

    caps = find_captions(pdf)
    if not caps:
        warnings.append('未检测到“图N”格式的独立图题')

    # Strongest extraction signal: the Word object preview in the paragraph immediately
    # before the caption. This fixes OLE/Visio/EMF figures and avoids body-text contamination.
    direct_media = find_docx_caption_media(normalized_docx)
    media_by_no = {}
    for m in direct_media:
        media_by_no.setdefault(m['figure_no'], []).append(m)
    used_media_ids = set()

    caps_by_page = {}
    for c in caps:
        caps_by_page.setdefault(c['page'], []).append(c)
    for pno in caps_by_page:
        caps_by_page[pno].sort(key=lambda x: x['bbox'].y0)

    figures = []
    used_names = {}
    for pno in sorted(caps_by_page):
        prev_cap_y1 = None
        for cap in caps_by_page[pno]:
            no = cap['no']
            base = f'图{no}'
            used_names[base] = used_names.get(base, 0) + 1
            suffix = '' if used_names[base] == 1 else f'_{used_names[base]}'
            if suffix:
                warnings.append(f'检测到重复图号：图{no}')
            filename = f'{base}{suffix}.png'
            out_img = article_dir / filename

            crop = None
            method = ''
            confidence = ''
            source_object = ''

            candidates = media_by_no.get(no, [])
            # Prefer caption text match if a malformed document contains duplicate figure numbers.
            ranked = sorted(
                enumerate(candidates),
                key=lambda kv: (0 if clean_line(kv[1].get('caption','')) == clean_line(cap['caption']) else 1, kv[0])
            )
            for idx, media in ranked:
                ident = (no, media.get('paragraph_index'), media.get('member'))
                if ident in used_media_ids:
                    continue
                ok, direct_method = extract_docx_media_as_png(normalized_docx, media, out_img, work_dir)
                used_media_ids.add(ident)
                if ok:
                    method = direct_method
                    confidence = 'high'
                    source_object = media.get('target', '')
                    crop = fitz.Rect(0, 0, 0, 0)
                    break
                warnings.append(f'图{no} 已找到 Word 原始图对象（{media.get("target", "?")}），但直接渲染失败，已回退到 PDF 裁图')

            if not method:
                crop, method, confidence = crop_figure(pdf[pno], cap, prev_cap_y1, out_img)

            if confidence == 'low':
                warnings.append(f'图{no} 未能直接提取 Word 图对象，且 PDF 中也未找到可靠边界，使用低置信度区域裁切，请人工检查')

            figures.append(FigureInfo(
                figure_no=no,
                caption=cap['caption'],
                page=pno+1,
                filename=filename,
                crop_bbox=[] if method.startswith('docx_exact') else [round(crop.x0,1), round(crop.y0,1), round(crop.x1,1), round(crop.y1,1)],
                extraction_method=method,
                confidence=confidence,
                source_object=source_object,
            ))
            prev_cap_y1 = cap['bbox'].y1

    article = ArticleInfo(
        source_file=src.name,
        first_author=first_author,
        title=title,
        article_folder=article_dir.name,
        figures=figures,
        warnings=warnings,
    )
    with open(article_dir / 'manifest.json', 'w', encoding='utf-8') as f:
        json.dump(asdict(article), f, ensure_ascii=False, indent=2)
    pdf.close()
    return article


def make_report(out_root: Path, articles: List[ArticleInfo]):
    with open(out_root / '提取报告.csv', 'w', newline='', encoding='utf-8-sig') as f:
        w = csv.writer(f)
        w.writerow(['来源文件', '第一作者', '正文标题', '识别图数', '图号', '提取方式', '低置信度图', '异常/提醒'])
        for a in articles:
            w.writerow([
                a.source_file, a.first_author, a.title, len(a.figures),
                '、'.join('图'+x.figure_no for x in a.figures),
                '；'.join(f'图{x.figure_no}:{x.extraction_method}' for x in a.figures),
                '、'.join('图'+x.figure_no for x in a.figures if x.confidence == 'low'),
                '；'.join(a.warnings)
            ])


def zip_dir(folder: Path, zip_path: Path):
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as z:
        for p in folder.rglob('*'):
            if p.is_file():
                z.write(p, p.relative_to(folder))


@app.get('/')
def index():
    return render_template('index.html')


@app.post('/extract')
def extract():
    logging.info('收到提取请求')
    files = request.files.getlist('files')
    word_files = [f for f in files if Path(f.filename).suffix.lower() in WORD_EXTS]
    if not word_files:
        return jsonify({'error': '所选文件夹中没有 .doc 或 .docx 文件。'}), 400

    temp_root = Path(tempfile.mkdtemp(prefix='journal_fig_'))
    incoming = temp_root / 'incoming'
    out_root = temp_root / '图件提取结果'
    work = temp_root / 'work'
    incoming.mkdir(); out_root.mkdir(); work.mkdir()

    articles = []
    failures = []
    try:
        for f in word_files:
            # Browsers may send paths; only keep basename for safety.
            name = Path(f.filename.replace('\\', '/')).name
            src = incoming / sanitize_filename(name, 150)
            f.save(src)
            try:
                articles.append(process_one_word(src, out_root, work))
            except Exception as e:
                logging.exception('处理文件失败: %s', name)
                failures.append({'file': name, 'error': str(e), 'traceback': traceback.format_exc()})

        if failures:
            with open(out_root / '处理失败.json', 'w', encoding='utf-8') as fp:
                json.dump(failures, fp, ensure_ascii=False, indent=2)
        make_report(out_root, articles)
        with open(out_root / '批次摘要.json', 'w', encoding='utf-8') as fp:
            json.dump({
                'word_count': len(word_files),
                'success_count': len(articles),
                'failure_count': len(failures),
                'figure_count': sum(len(a.figures) for a in articles),
                'articles': [asdict(a) for a in articles],
                'failures': failures,
            }, fp, ensure_ascii=False, indent=2)

        zip_path = temp_root / '图件提取结果.zip'
        zip_dir(out_root, zip_path)
        data = zip_path.read_bytes()
        mem = io.BytesIO(data)
        mem.seek(0)
        return send_file(mem, as_attachment=True, download_name='图件提取结果.zip', mimetype='application/zip')
    finally:
        # send_file reads BytesIO, so temp files can be deleted immediately.
        shutil.rmtree(temp_root, ignore_errors=True)


# ---------------------------------------------------------------------------
# v1.2: strict two-pass text audit + multi-provider semantic comparison
# Providers: Alibaba Bailian/Qwen (primary), Groq/Qwen, OpenRouter Free.
# Gemini is intentionally removed.
# ---------------------------------------------------------------------------
JOBS_DIR = Path(__file__).resolve().parent / 'runtime' / 'jobs'
JOBS_DIR.mkdir(parents=True, exist_ok=True)
REVISED_EXTS = {'.pdf', '.jpg', '.jpeg', '.png'}
FIG_NO_RE = re.compile(r'图\s*([0-9０-９]+)', re.I)

PROVIDERS = {
    'bailian': {
        'label': '阿里云百炼 Qwen',
        'cred_target': 'JournalFigureExtractor/BailianAPIKey',
        'env': 'DASHSCOPE_API_KEY',
        'base_url': 'https://dashscope.aliyuncs.com/compatible-mode/v1',
        'default_model': 'qwen3.7-plus',
    },
    'groq': {
        'label': 'Groq Qwen',
        'cred_target': 'JournalFigureExtractor/GroqAPIKey',
        'env': 'GROQ_API_KEY',
        'base_url': 'https://api.groq.com/openai/v1',
        'default_model': 'qwen/qwen3.8-27b',
    },
    'openrouter': {
        'label': 'OpenRouter Free',
        'cred_target': 'JournalFigureExtractor/OpenRouterAPIKey',
        'env': 'OPENROUTER_API_KEY',
        'base_url': 'https://openrouter.ai/api/v1',
        'default_model': 'openrouter/free',
    },
}
_SESSION_PROVIDER_KEYS: dict[str, str] = {}


class ProviderAPIError(RuntimeError):
    def __init__(self, provider: str, message: str, *, status: Optional[int] = None, transient: bool = False):
        super().__init__(message)
        self.provider = provider
        self.status = status
        self.transient = transient


def _cleanup_old_jobs(max_age_hours: int = 72) -> None:
    cutoff = time.time() - max_age_hours * 3600
    try:
        for d in JOBS_DIR.iterdir():
            if d.is_dir() and d.stat().st_mtime < cutoff:
                shutil.rmtree(d, ignore_errors=True)
    except Exception:
        logging.exception('清理旧任务失败')


def _job_dir(job_id: str) -> Path:
    if not re.fullmatch(r'[0-9a-f]{16,40}', job_id or ''):
        raise ValueError('无效任务 ID')
    d = JOBS_DIR / job_id
    if not d.exists():
        raise FileNotFoundError('任务不存在或已清理')
    return d


def _job_json_path(job_id: str) -> Path:
    return _job_dir(job_id) / 'job.json'


def _load_job(job_id: str) -> dict:
    return json.loads(_job_json_path(job_id).read_text(encoding='utf-8'))


def _save_job(job_id: str, data: dict) -> None:
    path = _job_json_path(job_id)
    tmp = path.with_suffix('.tmp')
    tmp.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding='utf-8')
    tmp.replace(path)


def _decode_cred_blob(blob) -> str:
    if isinstance(blob, bytes):
        for enc in ('utf-8', 'utf-16-le'):
            try:
                val = blob.decode(enc).rstrip('\x00')
                if val:
                    return val
            except Exception:
                pass
        return ''
    return str(blob or '')


def _provider_cfg(provider: str) -> dict:
    if provider not in PROVIDERS:
        raise ValueError(f'不支持的 AI 服务：{provider}')
    return PROVIDERS[provider]


def load_provider_key(provider: str) -> Tuple[Optional[str], str]:
    cfg = _provider_cfg(provider)
    if _SESSION_PROVIDER_KEYS.get(provider):
        return _SESSION_PROVIDER_KEYS[provider], 'session'
    env_key = os.environ.get(cfg['env'], '').strip()
    if env_key:
        return env_key, 'environment'
    if sys.platform.startswith('win'):
        try:
            import win32cred
            cred = win32cred.CredRead(cfg['cred_target'], win32cred.CRED_TYPE_GENERIC, 0)
            key = _decode_cred_blob(cred.get('CredentialBlob')).strip()
            if key:
                return key, 'windows_credential'
        except Exception:
            pass
    return None, 'none'


def save_provider_key(provider: str, key: str, persist: bool = True) -> str:
    cfg = _provider_cfg(provider)
    key = (key or '').strip()
    if not key:
        raise ValueError('API Key 为空')
    _SESSION_PROVIDER_KEYS[provider] = key
    if persist and sys.platform.startswith('win'):
        try:
            import win32cred
            win32cred.CredWrite({
                'Type': win32cred.CRED_TYPE_GENERIC,
                'TargetName': cfg['cred_target'],
                'UserName': cfg['label'],
                'CredentialBlob': key,
                'Persist': win32cred.CRED_PERSIST_LOCAL_MACHINE,
                'Comment': f'科技期刊图件提取器 {cfg["label"]} API Key',
            }, 0)
            return 'windows_credential'
        except Exception as exc:
            logging.warning('%s 保存到 Windows 凭据失败，仅保留本次会话: %s', provider, exc)
            return 'session'
    return 'session'


def clear_provider_key(provider: str) -> None:
    cfg = _provider_cfg(provider)
    _SESSION_PROVIDER_KEYS.pop(provider, None)
    if sys.platform.startswith('win'):
        try:
            import win32cred
            win32cred.CredDelete(cfg['cred_target'], win32cred.CRED_TYPE_GENERIC, 0)
        except Exception:
            pass


def _provider_http_json(provider: str, path: str, api_key: str, payload: Optional[dict] = None,
                        timeout: int = 120, extra_headers: Optional[dict] = None) -> dict:
    cfg = _provider_cfg(provider)
    url = cfg['base_url'].rstrip('/') + '/' + path.lstrip('/')
    headers = {
        'Authorization': f'Bearer {api_key}',
        'Accept': 'application/json',
        'User-Agent': 'JournalFigureExtractor/1.1',
    }
    if extra_headers:
        headers.update(extra_headers)
    data = None
    if payload is not None:
        headers['Content-Type'] = 'application/json; charset=utf-8'
        data = json.dumps(payload, ensure_ascii=False).encode('utf-8')
    req = urllib.request.Request(url, data=data, headers=headers, method='POST' if data is not None else 'GET')
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            raw = resp.read().decode('utf-8', errors='replace')
            return json.loads(raw)
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode('utf-8', errors='replace')
        try:
            j = json.loads(detail)
            err = j.get('error')
            if isinstance(err, dict):
                msg = err.get('message') or err.get('code') or detail
            else:
                msg = str(err or detail)
        except Exception:
            msg = detail
        label = cfg['label']
        if exc.code == 429:
            raise ProviderAPIError(provider, f'{label} 触发频率/额度限制（429）：{msg}', status=429, transient=True)
        if exc.code in (401, 403):
            raise ProviderAPIError(provider, f'{label} API Key 无效、权限不足或模型未授权（{exc.code}）：{msg}', status=exc.code)
        transient = exc.code in (408, 409, 425) or exc.code >= 500
        raise ProviderAPIError(provider, f'{label} HTTP {exc.code}: {msg}', status=exc.code, transient=transient)
    except urllib.error.URLError as exc:
        raise ProviderAPIError(provider, f'无法连接 {cfg["label"]}：{exc.reason}', transient=True)
    except TimeoutError:
        raise ProviderAPIError(provider, f'{cfg["label"]} 请求超时', transient=True)


def test_provider_key(provider: str, api_key: str, model: str) -> dict:
    cfg = _provider_cfg(provider)
    if provider == 'bailian':
        payload = {
            'model': model or cfg['default_model'],
            'messages': [{'role': 'user', 'content': '只回复 OK'}],
            'temperature': 0,
            'max_tokens': 8,
            'enable_thinking': False,
        }
        data = _provider_http_json(provider, 'chat/completions', api_key, payload, timeout=35)
        text = str((((data.get('choices') or [{}])[0].get('message') or {}).get('content') or '')).strip()
        return {'ok': True, 'model_available': True, 'test_reply': text[:80]}
    data = _provider_http_json(provider, 'models', api_key, timeout=30)
    ids = [str(x.get('id', '')) for x in data.get('data', []) if isinstance(x, dict)]
    wanted = model or cfg['default_model']
    return {'ok': True, 'model_available': wanted in ids if ids else None, 'model_count': len(ids)}


def _api_image_bytes(path: Path, max_side: int = 2600) -> Tuple[bytes, str]:
    with Image.open(path) as im0:
        if 'A' in im0.getbands():
            rgba = im0.convert('RGBA')
            bg = Image.new('RGBA', rgba.size, (255, 255, 255, 255))
            bg.alpha_composite(rgba)
            im = bg.convert('RGB')
        else:
            im = im0.convert('RGB')
        im.thumbnail((max_side, max_side), Image.Resampling.LANCZOS)
        buf = io.BytesIO()
        im.save(buf, format='PNG', optimize=True)
        data = buf.getvalue()
        if len(data) > 8 * 1024 * 1024:
            buf = io.BytesIO()
            im.save(buf, format='JPEG', quality=94, optimize=True)
            return buf.getvalue(), 'image/jpeg'
        return data, 'image/png'


def _extract_json_text(text: str) -> dict:
    text = (text or '').strip()
    if text.startswith('```'):
        text = re.sub(r'^```(?:json)?\s*', '', text, flags=re.I)
        text = re.sub(r'\s*```$', '', text)
    try:
        return json.loads(text)
    except Exception:
        i, j = text.find('{'), text.rfind('}')
        if i >= 0 and j > i:
            return json.loads(text[i:j+1])
        raise ValueError('AI 返回内容不是有效 JSON')


def _normalize_comparison_result(obj: dict) -> dict:
    result = str(obj.get('result', 'REVIEW')).upper().strip()
    if result not in {'PASS', 'REVIEW', 'FAIL'}:
        result = 'REVIEW'
    diffs = obj.get('differences') if isinstance(obj.get('differences'), list) else []
    clean_diffs = []
    for d in diffs[:30]:
        if not isinstance(d, dict):
            continue
        clean_diffs.append({
            'type': str(d.get('type', '其他'))[:40],
            'location': str(d.get('location', ''))[:160],
            'original': str(d.get('original', ''))[:500],
            'revised': str(d.get('revised', ''))[:500],
            'severity': str(d.get('severity', '中'))[:20],
            'confidence': d.get('confidence', None),
            'reason': str(d.get('reason', ''))[:500],
        })
    ignored = obj.get('ignored_changes') if isinstance(obj.get('ignored_changes'), list) else []
    uncertainties = obj.get('uncertainties') if isinstance(obj.get('uncertainties'), list) else []
    return {
        'result': result,
        'summary': str(obj.get('summary', '')).strip()[:1200],
        'differences': clean_diffs,
        'ignored_changes': [str(x)[:300] for x in ignored[:20]],
        'uncertainties': [str(x)[:300] for x in uncertainties[:20]],
    }



def _normalize_text_for_inventory(value: str) -> str:
    value = str(value or '')
    value = re.sub(r'\s+', '', value)
    value = value.translate(str.maketrans({
        '“':'"','”':'"','‘':"'",'’':"'",'（':'(', '）':')', '：':':', '；':';', '，':',', '。':'.',
    }))
    return value.strip()


def _normalize_text_audit(obj: dict) -> dict:
    def clean_items(key):
        raw = obj.get(key) if isinstance(obj.get(key), list) else []
        out=[]
        for x in raw[:120]:
            if isinstance(x, dict):
                text=str(x.get('text','')).strip()
                if not text:
                    continue
                out.append({'text':text[:500], 'location':str(x.get('location',''))[:180], 'confidence':x.get('confidence')})
            elif str(x).strip():
                out.append({'text':str(x).strip()[:500], 'location':'', 'confidence':None})
        return out
    diffs=[]
    raw_diffs=obj.get('text_differences') if isinstance(obj.get('text_differences'), list) else []
    for d in raw_diffs[:60]:
        if not isinstance(d,dict):
            continue
        status=str(d.get('status','changed')).lower().strip()
        if status not in {'changed','missing','added','uncertain'}:
            status='changed'
        diffs.append({
            'status':status,
            'location':str(d.get('location',''))[:180],
            'original':str(d.get('original',''))[:500],
            'revised':str(d.get('revised',''))[:500],
            'confidence':d.get('confidence'),
            'reason':str(d.get('reason',''))[:400],
        })
    out={
        'original_items':clean_items('original_items'),
        'revised_items':clean_items('revised_items'),
        'text_differences':diffs,
        'uncertainties':[str(x)[:300] for x in (obj.get('uncertainties') or [])[:30] if str(x).strip()],
    }
    from collections import Counter
    a=Counter(_normalize_text_for_inventory(x['text']) for x in out['original_items'] if _normalize_text_for_inventory(x['text']))
    b=Counter(_normalize_text_for_inventory(x['text']) for x in out['revised_items'] if _normalize_text_for_inventory(x['text']))
    only_a=[]; only_b=[]
    for k,n in (a-b).items():
        only_a.extend([k]*min(n,10))
    for k,n in (b-a).items():
        only_b.extend([k]*min(n,10))
    out['backend_inventory_match'] = not only_a and not only_b
    out['backend_only_original'] = only_a[:30]
    out['backend_only_revised'] = only_b[:30]
    return out


def _text_audit_prompt(*, author: str, figure_no: str, caption: str, title: str) -> str:
    return f"""你是科技期刊生产流程中的“逐字机械校对员”。这一步不是判断两张图语义是否相同，而是做严格文字清点。

图A=作者原稿；图B=重制图。作者：{author}；图号：图{figure_no}；图题：{caption}；文章：{title}。

强制规则：
1. 先分别、独立抄录图A和图B里所有有语义的可读文字。一个文本框/标签/公式/数字串作为一个 item。必须原样抄录，禁止自动纠错、补字、同义改写或根据另一张图猜测。
2. 中文必须逐字核对。少一个字、多一个字、错一个字、数字/英文/符号不同，都属于文字差异。比如“资源管理”与“源管理”绝不能视为等价。
3. 允许忽略纯空格、换行、字体、字号、字重以及排版导致的断行；但不得忽略标点、数字、单位、上下标、希腊字母、公式符号。
4. 如果某个字看不清，原样写你能确认的内容，并在 uncertainties 说明，不要脑补。
5. 不要在这一阶段评价框线、箭头、层级或版式。只做文字库存和文字差异。
6. 只能输出 JSON 对象，不要 Markdown。

JSON：
{{
  "original_items":[{{"text":"原样文字","location":"大致位置","confidence":0.0}}],
  "revised_items":[{{"text":"原样文字","location":"大致位置","confidence":0.0}}],
  "text_differences":[{{"status":"changed|missing|added|uncertain","location":"位置","original":"图A原文","revised":"图B原文","confidence":0.0,"reason":"逐字差异说明"}}],
  "uncertainties":["看不清的地方"]
}}"""

def _comparison_prompt(*, author: str, figure_no: str, caption: str, title: str) -> str:
    return f'''你是科技期刊生产流程中的“重制图内容等价性校对员”。

任务：比较两张图。图A是作者原稿，图B是编辑/美编重制图。文章第一作者：{author}；图号：图{figure_no}；图题：{caption}；文章标题：{title}。

判定原则：
1. 只检查信息/语义是否被改变，不做审美评价。
2. 必须忽略：字体、字号、字重、线宽、间距、框大小、版式重排、留白、黑白/彩色转换、无语义的颜色变化、清晰度与抗锯齿差异。
3. 必须逐项核对：可读文字、中文错字、英文/缩写、数字、小数点、正负号、上下标、希腊字母、公式、单位、百分号、坐标轴/刻度/图例、节点/模块数量、层级、箭头与连接关系及方向、流程顺序、曲线/柱形/数据点表达的数据关系，以及任何增删内容。
4. 若颜色承担分类、状态、图例等信息，颜色关系改变属于内容错误；纯装饰颜色变化忽略。
5. 不能因为布局变化就推断内容错误。先分别读懂两张图，再对照。
6. 看不清、被裁切、分辨率不足或无法确定时，不要猜，result 必须用 REVIEW，并在 uncertainties 写清楚。
7. PASS = 未发现实质内容差异；FAIL = 存在明确、可定位的实质差异；REVIEW = 有疑点但无法可靠确认。
8. 尽量少报假阳性。若只是同一内容重新绘制，应 PASS。
9. 请按 JSON 格式输出，且只能输出一个 JSON 对象，不要 Markdown，不要 JSON 之外的文字。

JSON 结构：
{{
  "result": "PASS|REVIEW|FAIL",
  "summary": "一句或两句中文结论",
  "differences": [
    {{"type":"文字|数字|公式|单位|节点|连接|方向|图例|数据|缺失|新增|其他", "location":"具体位置", "original":"原稿内容", "revised":"重制内容", "severity":"高|中|低", "confidence":0.0, "reason":"为什么构成实质差异"}}
  ],
  "ignored_changes": ["确认属于纯样式变化的项目"],
  "uncertainties": ["无法确认的内容"]
}}'''


def _data_url(data: bytes, mime: str) -> str:
    return f'data:{mime};base64,' + base64.b64encode(data).decode('ascii')


def _text_audit_with_provider_once(provider: str, original: Path, revised: Path, *, api_key: str,
                                   model: str, author: str, figure_no: str, caption: str, title: str) -> dict:
    a_bytes, a_mime = _api_image_bytes(original, max_side=3400)
    b_bytes, b_mime = _api_image_bytes(revised, max_side=3400)
    prompt = _text_audit_prompt(author=author, figure_no=figure_no, caption=caption, title=title)
    content = [
        {'type': 'text', 'text': prompt + '\n\n【图A：作者原稿】'},
        {'type': 'image_url', 'image_url': {'url': _data_url(a_bytes, a_mime)}},
        {'type': 'text', 'text': '【图B：重制图】'},
        {'type': 'image_url', 'image_url': {'url': _data_url(b_bytes, b_mime)}},
    ]
    payload = {
        'model': model,
        'messages': [{'role': 'user', 'content': content}],
        'temperature': 0,
        'max_tokens': 4200,
        'response_format': {'type': 'json_object'},
    }
    if provider == 'bailian':
        payload['enable_thinking'] = False
    elif provider == 'groq':
        payload['reasoning_effort'] = 'none'
    elif provider == 'openrouter':
        payload['provider'] = {'require_parameters': True}
    try:
        response = _provider_http_json(provider, 'chat/completions', api_key, payload, timeout=180)
    except ProviderAPIError as exc:
        if exc.status == 400 and ('response_format' in str(exc).lower() or 'structured' in str(exc).lower()):
            payload.pop('response_format', None)
            if provider == 'openrouter':
                payload.pop('provider', None)
            response = _provider_http_json(provider, 'chat/completions', api_key, payload, timeout=180)
        else:
            raise
    choices = response.get('choices') or []
    if not choices:
        raise ProviderAPIError(provider, f'{_provider_cfg(provider)["label"]} 文字清点未返回候选结果')
    message = choices[0].get('message') or {}
    text = message.get('content')
    if isinstance(text, list):
        text = ''.join(str(x.get('text', '')) for x in text if isinstance(x, dict))
    text = str(text or '').strip()
    if not text:
        raise ProviderAPIError(provider, f'{_provider_cfg(provider)["label"]} 文字清点返回空结果')
    audit = _normalize_text_audit(_extract_json_text(text))
    usage = response.get('usage') or {}
    audit['usage'] = {
        'prompt_tokens': usage.get('prompt_tokens'),
        'output_tokens': usage.get('completion_tokens'),
        'total_tokens': usage.get('total_tokens'),
    }
    return audit


def _merge_strict_text_audit(result: dict, audit: dict) -> dict:
    result['text_audit'] = audit
    definite=[]; uncertain=[]
    for d in audit.get('text_differences', []) or []:
        try:
            conf=float(d.get('confidence')) if d.get('confidence') is not None else 0.8
        except Exception:
            conf=0.8
        if d.get('status') in {'changed','missing','added'} and conf >= 0.72:
            definite.append(d)
        else:
            uncertain.append(d)
    # Promote definite text findings into the main difference list so the editor cannot miss them.
    existing={(str(x.get('original','')),str(x.get('revised',''))) for x in (result.get('differences') or []) if isinstance(x,dict)}
    for d in definite[:20]:
        key=(str(d.get('original','')),str(d.get('revised','')))
        if key in existing:
            continue
        result.setdefault('differences', []).insert(0, {
            'type':'文字', 'location':d.get('location',''), 'original':d.get('original',''), 'revised':d.get('revised',''),
            'severity':'高', 'confidence':d.get('confidence'), 'reason':d.get('reason') or '逐字清点发现文字不一致',
        })
    if definite:
        result['result']='FAIL'
        result['summary'] = f'逐字文字清点发现 {len(definite)} 处明确差异。' + ((' ' + result.get('summary','')) if result.get('summary') else '')
    elif (not audit.get('backend_inventory_match', True)) or uncertain or audit.get('uncertainties'):
        if result.get('result') == 'PASS':
            result['result']='REVIEW'
            result['summary'] = '逐字文字清点存在未能完全对齐的项目，禁止自动判 PASS；请人工复核。 ' + result.get('summary','')
    return result


def _compare_with_provider_once(provider: str, original: Path, revised: Path, *, api_key: str,
                                model: str, author: str, figure_no: str, caption: str, title: str, text_audit: Optional[dict] = None) -> dict:
    a_bytes, a_mime = _api_image_bytes(original)
    b_bytes, b_mime = _api_image_bytes(revised)
    prompt = _comparison_prompt(author=author, figure_no=figure_no, caption=caption, title=title)
    if text_audit:
        prompt += '\n\n【第一阶段逐字清点结果】\n' + json.dumps(text_audit, ensure_ascii=False) + '\n硬性要求：第一阶段若列出明确文字 changed/missing/added，本阶段不得判 PASS；必须在 differences 中逐项体现。'
    content = [
        {'type': 'text', 'text': prompt + '\n\n【图A：作者原稿】'},
        {'type': 'image_url', 'image_url': {'url': _data_url(a_bytes, a_mime)}},
        {'type': 'text', 'text': '【图B：重制图】'},
        {'type': 'image_url', 'image_url': {'url': _data_url(b_bytes, b_mime)}},
    ]
    payload = {
        'model': model,
        'messages': [{'role': 'user', 'content': content}],
        'temperature': 0.1,
        'max_tokens': 2500,
        'response_format': {'type': 'json_object'},
    }
    if provider == 'bailian':
        payload['enable_thinking'] = False
    elif provider == 'groq':
        payload['reasoning_effort'] = 'none'
    elif provider == 'openrouter':
        payload['provider'] = {'require_parameters': True}

    try:
        response = _provider_http_json(provider, 'chat/completions', api_key, payload, timeout=150)
    except ProviderAPIError as exc:
        if exc.status == 400 and ('response_format' in str(exc).lower() or 'structured' in str(exc).lower()):
            payload.pop('response_format', None)
            if provider == 'openrouter':
                payload.pop('provider', None)
            response = _provider_http_json(provider, 'chat/completions', api_key, payload, timeout=150)
        else:
            raise

    choices = response.get('choices') or []
    if not choices:
        raise ProviderAPIError(provider, f'{_provider_cfg(provider)["label"]} 未返回候选结果')
    message = choices[0].get('message') or {}
    text = message.get('content')
    if isinstance(text, list):
        text = ''.join(str(x.get('text', '')) for x in text if isinstance(x, dict))
    text = str(text or '').strip()
    if not text:
        raise ProviderAPIError(provider, f'{_provider_cfg(provider)["label"]} 返回结果中没有文本')
    parsed = _normalize_comparison_result(_extract_json_text(text))
    usage = response.get('usage') or {}
    parsed['provider'] = provider
    parsed['provider_label'] = _provider_cfg(provider)['label']
    parsed['model'] = str(response.get('model') or model)
    parsed['usage'] = {
        'prompt_tokens': usage.get('prompt_tokens'),
        'output_tokens': usage.get('completion_tokens'),
        'total_tokens': usage.get('total_tokens'),
    }
    return parsed


def _model_for(provider: str, models: dict) -> str:
    cfg = _provider_cfg(provider)
    return str((models or {}).get(provider) or cfg['default_model']).strip() or cfg['default_model']


def compare_images_multi_provider(original: Path, revised: Path, *, mode: str, models: dict, precision: str = 'strict',
                                  author: str, figure_no: str, caption: str, title: str) -> dict:
    mode = (mode or 'auto').strip().lower()
    if mode == 'auto':
        chain = ['bailian', 'groq', 'openrouter']
    elif mode in PROVIDERS:
        chain = [mode]
    else:
        raise ValueError('无效 AI 服务模式')

    available = []
    for provider in chain:
        key, source = load_provider_key(provider)
        if key:
            available.append((provider, key, source))
    if not available:
        if mode == 'auto':
            raise RuntimeError('尚未配置任何 AI API Key。至少配置百炼、Groq 或 OpenRouter 其中一个。')
        raise RuntimeError(f'尚未配置 {_provider_cfg(mode)["label"]} API Key。')

    attempts = []
    first_provider = available[0][0]
    last_exc: Optional[Exception] = None
    for provider, key, _source in available:
        model = _model_for(provider, models)
        for attempt, delay in enumerate((0, 2), start=1):
            if delay:
                time.sleep(delay)
            try:
                text_audit = None
                if (precision or 'strict').strip().lower() != 'fast':
                    text_audit = _text_audit_with_provider_once(
                        provider, original, revised, api_key=key, model=model,
                        author=author, figure_no=figure_no, caption=caption, title=title,
                    )
                result = _compare_with_provider_once(
                    provider, original, revised, api_key=key, model=model,
                    author=author, figure_no=figure_no, caption=caption, title=title, text_audit=text_audit,
                )
                if text_audit is not None:
                    result = _merge_strict_text_audit(result, text_audit)
                    au = (text_audit.get('usage') or {})
                    ru = (result.get('usage') or {})
                    def addnum(a,b):
                        try: return (int(a or 0) + int(b or 0)) or None
                        except Exception: return b or a
                    result['usage'] = {
                        'prompt_tokens': addnum(au.get('prompt_tokens'), ru.get('prompt_tokens')),
                        'output_tokens': addnum(au.get('output_tokens'), ru.get('output_tokens')),
                        'total_tokens': addnum(au.get('total_tokens'), ru.get('total_tokens')),
                    }
                    result['audit_usage'] = au
                result['precision'] = (precision or 'strict').strip().lower()
                result['requested_mode'] = mode
                result['fallback_used'] = provider != first_provider or bool(attempts)
                if result['fallback_used']:
                    history = '；'.join(f"{_provider_cfg(x['provider'])['label']}：{x['error']}" for x in attempts[-4:])
                    result['fallback_note'] = f'已自动切换到 {result["provider_label"]} 完成校对。' + (f' 前序失败：{history}' if history else '')
                result['retry_history'] = attempts[-10:]
                return result
            except Exception as exc:
                last_exc = exc
                transient = isinstance(exc, ProviderAPIError) and exc.transient
                attempts.append({'provider': provider, 'model': model, 'attempt': attempt, 'error': str(exc)[:500]})
                logging.warning('AI provider failed provider=%s model=%s attempt=%s transient=%s: %s', provider, model, attempt, transient, exc)
                if transient and attempt == 1:
                    continue
                break
        if mode != 'auto':
            break

    detail = '\n'.join(f"- {_provider_cfg(x['provider'])['label']} / {x['model']}: {x['error']}" for x in attempts[-8:])
    raise RuntimeError('所有可用 AI 服务都未能完成本次校对。\n' + detail) from last_exc

def render_revised_file(src: Path, out_png: Path) -> List[str]:
    warnings = []
    ext = src.suffix.lower()
    if ext == '.pdf':
        d = fitz.open(src)
        try:
            if len(d) < 1:
                raise RuntimeError('PDF 无页面')
            if len(d) > 1:
                warnings.append(f'{src.name} 有 {len(d)} 页；当前只使用第 1 页作为重制图')
            pix = d[0].get_pixmap(matrix=fitz.Matrix(3.0, 3.0), alpha=False)
            pix.save(str(out_png))
        finally:
            d.close()
        trim_and_validate_png(out_png, padding=10)
        return warnings
    if ext in {'.jpg', '.jpeg', '.png'}:
        with Image.open(src) as im0:
            if 'A' in im0.getbands():
                rgba = im0.convert('RGBA')
                bg = Image.new('RGBA', rgba.size, (255,255,255,255))
                bg.alpha_composite(rgba)
                im = bg.convert('RGB')
            else:
                im = im0.convert('RGB')
            im.save(out_png, 'PNG')
        trim_and_validate_png(out_png, padding=6)
        return warnings
    raise RuntimeError(f'不支持的重制图格式：{ext}')


def _relative_upload_paths(field: str, count: int) -> List[str]:
    try:
        vals = json.loads(request.form.get(field, '[]'))
        if isinstance(vals, list):
            vals = [str(x) for x in vals]
        else:
            vals = []
    except Exception:
        vals = []
    if len(vals) < count:
        vals += [''] * (count - len(vals))
    return vals[:count]


def _figure_no_from_name(text: str) -> Optional[str]:
    m = FIG_NO_RE.search(text or '')
    return normalize_digits(m.group(1)) if m else None


def _pair_revised(articles: List[ArticleInfo], original_root: Path, revised_items: List[dict]) -> Tuple[List[dict], List[dict], List[dict]]:
    originals = []
    for ai, a in enumerate(articles):
        for fi, fig in enumerate(a.figures):
            originals.append({
                'article_index': ai,
                'figure_index': fi,
                'author': a.first_author,
                'title': a.title,
                'article_folder': a.article_folder,
                'figure_no': str(fig.figure_no),
                'caption': fig.caption,
                'original_path': str((original_root / a.article_folder / fig.filename).resolve()),
                'original_method': fig.extraction_method,
                'original_confidence': fig.confidence,
            })
    authors = sorted({x['author'] for x in originals if x['author'] and x['author'] != '未知作者'}, key=len, reverse=True)
    used_originals = set()
    pairs = []
    unmatched_revised = []
    for r in revised_items:
        label = (r.get('relative_path') or r.get('source_name') or '').replace('\\', '/')
        no = _figure_no_from_name(label)
        if not no:
            r['reason'] = '文件名/路径中未识别到“图N”'
            unmatched_revised.append(r)
            continue
        author_hits = [a for a in authors if a in label]
        candidate_originals = [x for x in originals if x['figure_no'] == no]
        if author_hits:
            author = author_hits[0]
            candidate_originals = [x for x in candidate_originals if x['author'] == author]
        elif len({x['article_index'] for x in originals}) == 1:
            pass
        else:
            r['reason'] = '重制图路径中未找到正文第一作者，且本批包含多篇文章，无法安全自动配对'
            unmatched_revised.append(r)
            continue
        available = [x for x in candidate_originals if (x['article_index'], x['figure_index']) not in used_originals]
        if len(available) != 1:
            r['reason'] = '找到多个或零个可能的原稿图，需人工确认'
            unmatched_revised.append(r)
            continue
        o = available[0]
        used_originals.add((o['article_index'], o['figure_index']))
        pair_id = f'p{len(pairs)+1:04d}'
        pair = dict(o)
        pair.update({
            'id': pair_id,
            'revised_path': r['rendered_path'],
            'revised_source': r.get('relative_path') or r.get('source_name'),
            'revised_warnings': r.get('warnings', []),
            'status': 'PENDING',
            'comparison': None,
        })
        pairs.append(pair)
    unmatched_originals = [x for x in originals if (x['article_index'], x['figure_index']) not in used_originals]
    return pairs, unmatched_originals, unmatched_revised


@app.get('/api/providers/status')
def providers_status():
    out = {}
    for provider, cfg in PROVIDERS.items():
        key, source = load_provider_key(provider)
        out[provider] = {
            'label': cfg['label'],
            'configured': bool(key),
            'source': source if key else 'none',
            'default_model': cfg['default_model'],
        }
    return jsonify({'providers': out})


@app.post('/api/providers/key')
def provider_key_set():
    data = request.get_json(silent=True) or {}
    provider = str(data.get('provider', '')).strip().lower()
    try:
        cfg = _provider_cfg(provider)
    except Exception as exc:
        return jsonify({'error': str(exc)}), 400
    key = str(data.get('api_key', '')).strip()
    model = str(data.get('model', cfg['default_model'])).strip() or cfg['default_model']
    persist = bool(data.get('persist', True))
    if not key:
        return jsonify({'error': f'请输入 {cfg["label"]} API Key'}), 400
    try:
        test = test_provider_key(provider, key, model)
        source = save_provider_key(provider, key, persist=persist)
        test.update({'saved': True, 'source': source, 'provider': provider, 'label': cfg['label'], 'model': model})
        return jsonify(test)
    except Exception as exc:
        logging.exception('%s API Key 测试失败', provider)
        return jsonify({'error': str(exc)}), 400


@app.delete('/api/providers/key/<provider>')
def provider_key_clear(provider: str):
    try:
        clear_provider_key(provider)
        return jsonify({'ok': True})
    except Exception as exc:
        return jsonify({'error': str(exc)}), 400

@app.post('/api/prepare-compare')
def prepare_compare():
    _cleanup_old_jobs()
    word_files = [f for f in request.files.getlist('word_files') if Path(f.filename).suffix.lower() in WORD_EXTS]
    revised_files = [f for f in request.files.getlist('revised_files') if Path(f.filename).suffix.lower() in REVISED_EXTS]
    if not word_files:
        return jsonify({'error': '没有选择 .doc / .docx 原稿。'}), 400
    if not revised_files:
        return jsonify({'error': '没有选择 PDF / JPG 重制图。'}), 400
    revised_relpaths = _relative_upload_paths('revised_relpaths', len(revised_files))
    job_id = uuid.uuid4().hex[:20]
    job = JOBS_DIR / job_id
    incoming_words = job / 'incoming_words'
    incoming_revised = job / 'incoming_revised'
    original_root = job / 'originals'
    revised_root = job / 'revised_rendered'
    work = job / 'work'
    for d in (incoming_words, incoming_revised, original_root, revised_root, work):
        d.mkdir(parents=True, exist_ok=True)
    articles: List[ArticleInfo] = []
    failures = []
    for idx, f in enumerate(word_files):
        name = Path(f.filename.replace('\\', '/')).name
        src = incoming_words / f'{idx:03d}_{sanitize_filename(name, 140)}'
        f.save(src)
        try:
            articles.append(process_one_word(src, original_root, work))
        except Exception as exc:
            logging.exception('对比准备：处理 Word 失败 %s', name)
            failures.append({'file': name, 'error': str(exc)})
    if not articles:
        shutil.rmtree(job, ignore_errors=True)
        return jsonify({'error': '所有 Word 原稿处理失败。', 'failures': failures}), 500
    revised_items = []
    for idx, f in enumerate(revised_files):
        orig_name = Path(f.filename.replace('\\', '/')).name
        ext = Path(orig_name).suffix.lower()
        src = incoming_revised / f'{idx:04d}_{sanitize_filename(Path(orig_name).stem, 100)}{ext}'
        f.save(src)
        out = revised_root / f'r{idx:04d}.png'
        rel = revised_relpaths[idx] or orig_name
        try:
            warns = render_revised_file(src, out)
            revised_items.append({
                'source_name': orig_name,
                'relative_path': rel,
                'rendered_path': str(out.resolve()),
                'warnings': warns,
            })
        except Exception as exc:
            revised_items.append({
                'source_name': orig_name,
                'relative_path': rel,
                'rendered_path': '',
                'warnings': [str(exc)],
                'render_failed': True,
                'reason': '重制图无法渲染：' + str(exc),
            })
    render_ok = [x for x in revised_items if x.get('rendered_path')]
    render_bad = [x for x in revised_items if not x.get('rendered_path')]
    pairs, unmatched_originals, unmatched_revised = _pair_revised(articles, original_root, render_ok)
    unmatched_revised.extend(render_bad)
    job_data = {
        'job_id': job_id,
        'created_at': time.strftime('%Y-%m-%d %H:%M:%S'),
        'articles': [asdict(a) for a in articles],
        'word_failures': failures,
        'pairs': pairs,
        'unmatched_originals': unmatched_originals,
        'unmatched_revised': unmatched_revised,
    }
    (job / 'job.json').write_text(json.dumps(job_data, ensure_ascii=False, indent=2), encoding='utf-8')
    public_pairs = []
    for p0 in pairs:
        public_pairs.append({k: p0.get(k) for k in [
            'id','author','title','figure_no','caption','original_method','original_confidence',
            'revised_source','revised_warnings','status','comparison'
        ]})
    return jsonify({
        'job_id': job_id,
        'article_count': len(articles),
        'original_figure_count': sum(len(a.figures) for a in articles),
        'revised_count': len(revised_files),
        'matched_count': len(pairs),
        'unmatched_original_count': len(unmatched_originals),
        'unmatched_revised_count': len(unmatched_revised),
        'pairs': public_pairs,
        'unmatched_originals': [{k:x.get(k) for k in ['author','title','figure_no','caption']} for x in unmatched_originals],
        'unmatched_revised': [{k:x.get(k) for k in ['relative_path','source_name','reason','warnings']} for x in unmatched_revised],
        'word_failures': failures,
    })


@app.get('/api/job/<job_id>/pair/<pair_id>/image/<side>')
def pair_image(job_id: str, pair_id: str, side: str):
    try:
        data = _load_job(job_id)
        pair = next((x for x in data.get('pairs', []) if x.get('id') == pair_id), None)
        if not pair:
            return jsonify({'error': '找不到该图对'}), 404
        if side == 'original':
            path = Path(pair['original_path'])
        elif side == 'revised':
            path = Path(pair['revised_path'])
        else:
            return jsonify({'error': '无效 side'}), 400
        if not path.exists():
            return jsonify({'error': '图片不存在'}), 404
        return send_file(path, mimetype='image/png', max_age=0)
    except Exception as exc:
        return jsonify({'error': str(exc)}), 404


@app.post('/api/job/<job_id>/pair/<pair_id>/compare')
def compare_pair(job_id: str, pair_id: str):
    body = request.get_json(silent=True) or {}
    mode = str(body.get('mode', 'auto')).strip().lower() or 'auto'
    models = body.get('models') if isinstance(body.get('models'), dict) else {}
    precision = str(body.get('precision', 'strict')).strip().lower() or 'strict'
    try:
        data = _load_job(job_id)
        pair = next((x for x in data.get('pairs', []) if x.get('id') == pair_id), None)
        if not pair:
            return jsonify({'error': '找不到该图对'}), 404
        pair['status'] = 'RUNNING'
        _save_job(job_id, data)
        result = compare_images_multi_provider(
            Path(pair['original_path']), Path(pair['revised_path']), mode=mode, models=models, precision=precision,
            author=pair.get('author',''), figure_no=pair.get('figure_no',''), caption=pair.get('caption',''), title=pair.get('title','')
        )
        pair['comparison'] = result
        pair['status'] = result['result']
        _save_job(job_id, data)
        return jsonify({'pair_id': pair_id, 'status': pair['status'], 'comparison': result})
    except Exception as exc:
        logging.exception('AI 图件校对失败 %s/%s', job_id, pair_id)
        try:
            data = _load_job(job_id)
            pair = next((x for x in data.get('pairs', []) if x.get('id') == pair_id), None)
            if pair:
                pair['status'] = 'ERROR'
                pair['comparison'] = {'error': str(exc)}
                _save_job(job_id, data)
        except Exception:
            pass
        msg = str(exc)
        code = 429 if ('429' in msg or '频率' in msg or '额度' in msg) else 500
        return jsonify({'error': msg}), code

def _report_html(data: dict) -> str:
    cards = []
    for pair in data.get('pairs', []):
        comp = pair.get('comparison') or {}
        result = comp.get('result') or pair.get('status') or 'PENDING'
        diff_html = ''
        for d in comp.get('differences', []) or []:
            diff_html += '<div class="diff"><b>{}</b> · {}<br>原稿：{}<br>重制：{}<br>{}</div>'.format(
                html.escape(str(d.get('type',''))), html.escape(str(d.get('location',''))),
                html.escape(str(d.get('original',''))), html.escape(str(d.get('revised',''))),
                html.escape(str(d.get('reason',''))),
            )
        if not diff_html:
            diff_html = '<p>未列出实质差异。</p>'
        base = sanitize_filename(f"{pair.get('author','')}_图{pair.get('figure_no','')}_{pair.get('id','')}", 80)
        cards.append(f'<section class="pair"><h2>{html.escape(pair.get("author", ""))} · 图{html.escape(str(pair.get("figure_no", "")))}　{html.escape(pair.get("caption", ""))}</h2><div class="grid"><div><h3>原稿</h3><img src="原稿图/{base}.png"></div><div><h3>重制图</h3><img src="重制图/{base}.png"></div><div><h3>{html.escape(result)}</h3><p>{html.escape(str(comp.get("summary", "尚未 AI 校对")))}</p>{diff_html}</div></div></section>')
    return '<!doctype html><meta charset="utf-8"><title>图件 AI 校对报告</title><style>body{font-family:"Microsoft YaHei",sans-serif;background:#f3f3f1;margin:24px;color:#222}.pair{background:white;border:1px solid #ddd;border-radius:14px;padding:18px;margin:16px auto;max-width:1500px}.grid{display:grid;grid-template-columns:1fr 1fr 1fr;gap:16px}.grid>div{min-width:0}.grid img{width:100%;max-height:680px;object-fit:contain;background:#fafafa;border:1px solid #eee}.diff{padding:10px 0;border-top:1px solid #eee;line-height:1.6}@media(max-width:900px){.grid{grid-template-columns:1fr}}</style><h1>科技期刊图件 AI 校对报告</h1>' + ''.join(cards)


@app.get('/api/job/<job_id>/export')
def export_job(job_id: str):
    try:
        data = _load_job(job_id)
        mem = io.BytesIO()
        csv_buf = io.StringIO()
        cw = csv.writer(csv_buf)
        cw.writerow(['第一作者','文章标题','图号','图题','重制文件','AI结果','结论','差异数','差异摘要','AI服务','模型','总Token'])
        report_data = json.loads(json.dumps(data, ensure_ascii=False))
        with zipfile.ZipFile(mem, 'w', zipfile.ZIP_DEFLATED) as z:
            for pair in data.get('pairs', []):
                comp = pair.get('comparison') or {}
                base = sanitize_filename(f"{pair.get('author','')}_图{pair.get('figure_no','')}_{pair.get('id','')}", 80)
                op = Path(pair['original_path'])
                rp = Path(pair['revised_path'])
                if op.exists():
                    z.write(op, f'原稿图/{base}.png')
                if rp.exists():
                    z.write(rp, f'重制图/{base}.png')
                ds = '；'.join(f"{d.get('location','')}：{d.get('original','')} → {d.get('revised','')}" for d in (comp.get('differences') or []))
                cw.writerow([
                    pair.get('author',''), pair.get('title',''), pair.get('figure_no',''), pair.get('caption',''), pair.get('revised_source',''),
                    comp.get('result') or pair.get('status',''), comp.get('summary',''), len(comp.get('differences') or []), ds,
                    comp.get('provider_label',''), comp.get('model',''), (comp.get('usage') or {}).get('total_tokens','')
                ])
            for p0 in report_data.get('pairs', []):
                p0.pop('original_path', None)
                p0.pop('revised_path', None)
            z.writestr('AI校对结果.json', json.dumps(report_data, ensure_ascii=False, indent=2))
            z.writestr('AI校对结果.csv', '\ufeff' + csv_buf.getvalue())
            z.writestr('对比报告.html', _report_html(data))
        mem.seek(0)
        return send_file(mem, as_attachment=True, download_name='图件AI校对结果.zip', mimetype='application/zip')
    except Exception as exc:
        return jsonify({'error': str(exc)}), 404


APP_ID = 'journal-figure-extractor'
APP_VERSION = '1.2'

@app.get('/health')
def health():
    return jsonify({'ok': True, 'app_id': APP_ID, 'version': APP_VERSION, 'platform': sys.platform, 'python': sys.version.split()[0]})

@app.get('/diagnostics')
def diagnostics():
    result = {'platform': sys.platform, 'python': sys.version, 'word_com': False, 'word_version': None, 'errors': []}
    try:
        import pymupdf as _fitz
        result['pymupdf'] = getattr(_fitz, '__doc__', '')[:80]
    except Exception as e:
        result['errors'].append('PyMuPDF: ' + str(e))
    if sys.platform.startswith('win'):
        try:
            import pythoncom
            import win32com.client
            pythoncom.CoInitialize()
            w = win32com.client.DispatchEx('Word.Application')
            result['word_version'] = str(w.Version)
            w.Quit()
            pythoncom.CoUninitialize()
            result['word_com'] = True
        except Exception as e:
            result['errors'].append('Word COM: ' + str(e))
    return jsonify(result)


if __name__ == '__main__':
    port = int(os.environ.get('JFE_PORT', '8765'))
    app.run(host='127.0.0.1', port=port, debug=False, threaded=False)
