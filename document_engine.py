# -*- coding: utf-8 -*-
"""Document parsing/rendering core for Document Object Studio."""
from __future__ import annotations

import hashlib
import math
import mimetypes
import os
import shutil
import subprocess
import sys
import tempfile
import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from pathlib import Path, PurePosixPath
from typing import Any


@dataclass
class DocObject:
    uid: str
    label: str
    kind: str
    page: int | None = None
    bbox: tuple[float, float, float, float] | None = None
    data: bytes | None = None
    mime_type: str | None = None
    text: str = ""
    html: str = ""
    children: list["DocObject"] = field(default_factory=list)
    metadata: dict[str, Any] = field(default_factory=dict)


class DocumentEngine:
    VERSION = "0.2.0"

    def __init__(self):
        self.path = ""
        self.kind = ""
        self.temp_dir = tempfile.mkdtemp(prefix="doc_object_studio_")
        self.preview_paths: list[str] = []
        self.pdf_doc = None
        self.preview_pdf_doc = None
        self.pptx = None
        self.docx = None
        self.docx_objects: list[DocObject] = []
        self.docx_table_objects: list[DocObject] = []
        self.docx_paragraph_objects: list[DocObject] = []
        self._docx_preview_map: dict[tuple, DocObject] = {}
        self.slide_width = 0
        self.slide_height = 0
        self.render_scale = 3.0
        self.progress_callback = None

    def _report_progress(self, percent: float, message: str):
        callback = self.progress_callback
        if callback is not None:
            try:
                callback(max(0, min(100, int(percent))), message)
            except Exception:
                pass

    def close(self):
        for attr in ("pdf_doc", "preview_pdf_doc"):
            doc = getattr(self, attr, None)
            if doc is not None:
                try:
                    doc.close()
                except Exception:
                    pass
                setattr(self, attr, None)
        if self.temp_dir and os.path.isdir(self.temp_dir):
            shutil.rmtree(self.temp_dir, ignore_errors=True)
        self.temp_dir = ""
        self.preview_paths = []

    def open(self, path: str) -> int:
        self._report_progress(0, "正在准备文档")
        self.close()
        self.temp_dir = tempfile.mkdtemp(prefix="doc_object_studio_")
        self.pptx = None
        self.docx = None
        self.docx_objects = []
        self.docx_table_objects = []
        self.docx_paragraph_objects = []
        self._docx_preview_map = {}
        self.slide_width = 0
        self.slide_height = 0
        self.path = os.path.abspath(path)
        suffix = Path(path).suffix.lower()
        if suffix == ".pdf":
            self.kind = "pdf"
            self._open_pdf(self.path)
        elif suffix == ".pptx":
            self.kind = "pptx"
            self._open_pptx(self.path)
        elif suffix == ".docx":
            self.kind = "docx"
            self._open_docx(self.path)
        else:
            raise ValueError("仅支持 PDF、PPTX、DOCX 文件。")
        self._report_progress(100, "页面预览已就绪")
        return len(self.preview_paths)

    @staticmethod
    def _run_soffice(args: list[str], timeout: int = 180):
        executable = shutil.which("soffice") or shutil.which("libreoffice")
        if not executable and sys.platform.startswith("win"):
            for p in (
                r"C:\Program Files\LibreOffice\program\soffice.exe",
                r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
            ):
                if os.path.exists(p):
                    executable = p
                    break
        if not executable:
            raise RuntimeError("未找到 LibreOffice。请安装 Microsoft Office 或 LibreOffice。")
        return subprocess.run([executable, *args], capture_output=True, text=True, timeout=timeout)

    def _render_pdf(
        self,
        pdf_path: str,
        prefix: str = "page",
        scale: float | None = None,
        progress_start: int = 10,
        progress_end: int = 96,
    ) -> list[str]:
        import fitz
        doc = fitz.open(pdf_path)
        scale = float(scale or self.render_scale)
        outputs = []
        total = max(1, len(doc))
        for i, page in enumerate(doc):
            pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
            out = os.path.join(self.temp_dir, f"{prefix}_{i+1:04d}.png")
            pix.save(out)
            outputs.append(out)
            percent = progress_start + (progress_end - progress_start) * (i + 1) / total
            self._report_progress(percent, f"正在生成高清预览 {i + 1}/{total} 页")
        doc.close()
        return outputs

    def _office_to_pdf(self, path: str) -> str:
        out_dir = os.path.join(self.temp_dir, "office_pdf")
        os.makedirs(out_dir, exist_ok=True)
        # On Windows prefer Office COM to preserve the native Office layout.
        if sys.platform.startswith("win"):
            try:
                import pythoncom
                import win32com.client
                pythoncom.CoInitialize()
                if Path(path).suffix.lower() == ".pptx":
                    app = win32com.client.DispatchEx("PowerPoint.Application")
                    # PowerPoint rejects Visible=False on some Office builds;
                    # WithWindow=False below keeps the presentation window hidden.
                    pres = app.Presentations.Open(path, ReadOnly=True, WithWindow=False)
                    pdf_path = os.path.join(out_dir, "office_export.pdf")
                    pres.SaveAs(pdf_path, 32)  # ppSaveAsPDF
                    pres.Close()
                    app.Quit()
                    if os.path.exists(pdf_path):
                        return pdf_path
                else:
                    app = win32com.client.DispatchEx("Word.Application")
                    app.Visible = False
                    app.DisplayAlerts = 0
                    doc = app.Documents.Open(path, ReadOnly=True, AddToRecentFiles=False, Visible=False)
                    pdf_path = os.path.join(out_dir, "office_export.pdf")
                    doc.ExportAsFixedFormat(pdf_path, 17)  # wdExportFormatPDF
                    doc.Close(False)
                    app.Quit()
                    if os.path.exists(pdf_path):
                        return pdf_path
            except Exception:
                try:
                    app.Quit()
                except Exception:
                    pass
        result = self._run_soffice(["--headless", "--convert-to", "pdf", "--outdir", out_dir, path])
        candidate = os.path.join(out_dir, Path(path).stem + ".pdf")
        if result.returncode != 0 or not os.path.exists(candidate):
            raise RuntimeError(
                "Office 文件转 PDF 预览失败。请安装 Microsoft Office 或 LibreOffice。\n"
                + (result.stderr or result.stdout or "无转换器输出")
            )
        return candidate

    def _open_pdf(self, path: str):
        import fitz
        self._report_progress(5, "正在读取 PDF 页面")
        self.pdf_doc = fitz.open(path)
        self.preview_paths = self._render_pdf(path, "pdf_page", progress_start=8, progress_end=97)

    def _open_pptx(self, path: str):
        import fitz
        from pptx import Presentation
        self._report_progress(5, "正在解析 PowerPoint 对象")
        self.pptx = Presentation(path)
        self.slide_width = int(self.pptx.slide_width)
        self.slide_height = int(self.pptx.slide_height)
        self._report_progress(20, "正在调用 PowerPoint 导出页面")
        pdf_path = self._office_to_pdf(path)
        self.preview_pdf_doc = fitz.open(pdf_path)
        self._report_progress(42, "正在绘制高清幻灯片预览")
        self.preview_paths = self._render_pdf(pdf_path, "pptx_page", progress_start=44, progress_end=97)

    def _open_docx(self, path: str):
        import fitz
        from docx import Document
        self._report_progress(5, "正在解析 Word 对象和图片")
        self.docx = Document(path)
        self.docx_objects = []
        self.docx_table_objects = []
        self.docx_paragraph_objects = []
        self._docx_preview_map = {}
        self._parse_docx_package(path)
        self._report_progress(20, "正在调用 Word 导出分页预览")
        pdf_path = self._office_to_pdf(path)
        self.preview_pdf_doc = fitz.open(pdf_path)
        self._report_progress(42, "正在绘制高清页面预览")
        self.preview_paths = self._render_pdf(pdf_path, "docx_page", progress_start=44, progress_end=94)
        self._map_docx_preview_images()
        self._report_progress(97, "正在匹配页面对象")

    def page_count(self) -> int:
        return len(self.preview_paths)

    def objects_for_page(self, page_index: int) -> list[DocObject]:
        if self.kind == "pdf":
            return self._pdf_objects(self.pdf_doc[page_index], page_index, self.pdf_doc, "pdf")
        if self.kind == "pptx":
            return self._pptx_objects(page_index)
        if self.kind == "docx":
            return self._docx_page_objects(page_index)
        return []

    @staticmethod
    def _span_html(span: dict) -> str:
        text = (span.get("text") or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        size = float(span.get("size", 10))
        font = span.get("font", "Arial").replace("'", "")
        color = span.get("color", 0)
        rgb = f"#{(color >> 16) & 255:02x}{(color >> 8) & 255:02x}{color & 255:02x}"
        flags = int(span.get("flags", 0))
        weight = "bold" if flags & 16 else "normal"
        style = f"font-family:'{font}';font-size:{size:.1f}pt;color:{rgb};font-weight:{weight};"
        if flags & 2:
            style += "font-style:italic;"
        return f'<span style="{style}">{text}</span>'

    def _pdf_objects(self, page, page_index: int, source_doc=None, source_name="pdf") -> list[DocObject]:
        source_doc = source_doc or self.pdf_doc
        result: list[DocObject] = []
        # Text is split by line, making copy more granular than a whole page block.
        text_no = 0
        for block in page.get_text("dict").get("blocks", []):
            if block.get("type") != 0:
                continue
            for line in block.get("lines", []):
                spans = line.get("spans", [])
                text = "".join(s.get("text", "") for s in spans).strip()
                if not text:
                    continue
                bbox = tuple(float(v) for v in line.get("bbox", block.get("bbox", (0, 0, 0, 0))))
                text_no += 1
                html = "".join(self._span_html(s) for s in spans)
                result.append(DocObject(
                    uid=f"pdf-{page_index}-text-{text_no}",
                    label=f"文字行 {text_no} · {text[:42]}", kind="text", page=page_index,
                    bbox=bbox, text=text, html=f"<div>{html}</div>",
                    metadata={"source": source_name + "-text", "source_doc": source_name}
                ))

        # Use image xrefs + placements to preserve the embedded image resource.
        seen = set()
        seen_bboxes = set()
        image_no = 0
        for item in page.get_images(full=True):
            xref = item[0]
            if not xref:
                continue
            try:
                extracted = source_doc.extract_image(xref)
                data = extracted.get("image", b"")
                ext = extracted.get("ext", "bin").lower()
                mime = mimetypes.types_map.get("." + ext, "application/octet-stream")
                placements = page.get_image_rects(xref, transform=True) or []
                for placement in placements:
                    rect = placement[0] if isinstance(placement, tuple) else placement
                    bbox_key = tuple(round(v, 2) for v in (rect.x0, rect.y0, rect.x1, rect.y1))
                    key = (xref, *bbox_key)
                    if bbox_key in seen_bboxes:
                        continue
                    seen.add(key)
                    seen_bboxes.add(bbox_key)
                    image_no += 1
                    result.append(DocObject(
                        uid=f"pdf-{page_index}-image-{image_no}",
                        label=f"图片 {image_no} · {ext.upper()} · {len(data)//1024} KB",
                        kind="image", page=page_index,
                        bbox=(rect.x0, rect.y0, rect.x1, rect.y1), data=data,
                        mime_type=mime, metadata={"xref": xref, "source": source_name + "-xref", "source_doc": source_name, "width": extracted.get("width"), "height": extracted.get("height")}
                    ))
            except Exception:
                continue

        # Inline images / images that PyMuPDF did not expose through xref placement.
        for block in page.get_text("dict").get("blocks", []):
            if block.get("type") != 1 or not block.get("image"):
                continue
            bbox = tuple(float(v) for v in block.get("bbox", (0, 0, 0, 0)))
            bbox_key = tuple(round(v, 2) for v in bbox)
            key = (0, *bbox_key)
            if bbox_key in seen_bboxes:
                continue
            seen.add(key)
            seen_bboxes.add(bbox_key)
            image_no += 1
            data = block["image"]
            ext = block.get("ext", "png")
            result.append(DocObject(
                uid=f"pdf-{page_index}-image-{image_no}",
                label=f"图片 {image_no} · {ext.upper()}", kind="image", page=page_index,
                bbox=bbox, data=data, mime_type=mimetypes.types_map.get("." + ext, "image/png"),
                metadata={"width": block.get("width"), "height": block.get("height"), "source": source_name + "-inline", "source_doc": source_name}
            ))

        # Expose PDF tables as whole objects. The text extraction is only used to
        # build a portable HTML clipboard fallback; bbox remains in PDF points so
        # the selection frame uses the same coordinates as the page preview.
        try:
            # For original PDF pages, use the vector/text table detector without
            # PyMuPDF's optional layout gate. This keeps the selectable table set
            # identical between the source environment and the packaged EXE.
            tables = (
                page.find_tables(use_layout=False).tables
                if source_name == "pdf"
                else page.find_tables().tables
            )
        except Exception:
            tables = []
        for table_no, table in enumerate(tables, 1):
            try:
                rows = [[cell or "" for cell in row] for row in table.extract()]
                if not rows or not any(cell.strip() for row in rows for cell in row):
                    continue
                bbox = tuple(float(value) for value in table.bbox)
                html_rows = "".join(
                    "<tr>" + "".join(f"<td>{self._html_escape(cell)}</td>" for cell in row) + "</tr>"
                    for row in rows
                )
                text_rows = ["\t".join(row) for row in rows]
                result.append(DocObject(
                    uid=f"pdf-{page_index}-table-{table_no}",
                    label=f"表格 {table_no} · {len(rows)} 行 × {max((len(row) for row in rows), default=0)} 列",
                    kind="table", page=page_index, bbox=bbox,
                    text="\n".join(text_rows),
                    html=f"<table border='1' cellspacing='0' cellpadding='4'>{html_rows}</table>",
                    metadata={"native_kind": "pdf-table", "source_doc": source_name, "rows": rows},
                ))
            except Exception:
                continue

        # Group connected PDF path fragments into logical vector regions.
        drawings = page.get_drawings()
        clusters = self._cluster_drawings(drawings, tolerance=3.0)
        for i, cluster in enumerate(clusters, 1):
            rect = cluster["bbox"]
            if rect.width < 1 and rect.height < 1:
                continue
            result.append(DocObject(
                uid=f"pdf-{page_index}-vector-{i}",
                label=f"矢量图形组 {i} · {len(cluster['indices'])} 个路径",
                kind="vector", page=page_index,
                bbox=(rect.x0, rect.y0, rect.x1, rect.y1),
                metadata={"drawing_indices": cluster["indices"], "source": source_name + "-vector", "source_doc": source_name, "path_count": len(cluster["indices"])}
            ))
        result.sort(key=lambda o: (o.bbox[1] if o.bbox else 0, o.bbox[0] if o.bbox else 0, o.kind))
        return result

    @staticmethod
    def _cluster_drawings(drawings: list[dict], tolerance: float = 3.0) -> list[dict]:
        import fitz
        rects = []
        for i, drawing in enumerate(drawings):
            rect = drawing.get("rect")
            if rect is not None:
                rects.append((i, fitz.Rect(rect)))
        n = len(rects)
        parent = list(range(n))

        def find(x):
            while parent[x] != x:
                parent[x] = parent[parent[x]]
                x = parent[x]
            return x

        def union(a, b):
            ra, rb = find(a), find(b)
            if ra != rb:
                parent[rb] = ra

        for i in range(n):
            _, a = rects[i]
            expanded = fitz.Rect(a.x0-tolerance, a.y0-tolerance, a.x1+tolerance, a.y1+tolerance)
            for j in range(i+1, n):
                _, b = rects[j]
                if expanded.intersects(b) or expanded.contains(b) or b.contains(expanded):
                    union(i, j)
        groups: dict[int, list[tuple[int, Any]]] = {}
        for i, item in enumerate(rects):
            groups.setdefault(find(i), []).append(item)
        output = []
        for values in groups.values():
            bbox = fitz.Rect(values[0][1])
            for _, rect in values[1:]:
                bbox |= rect
            output.append({"bbox": bbox, "indices": [idx for idx, _ in values]})
        output.sort(key=lambda x: (x["bbox"].y0, x["bbox"].x0))
        return output

    @staticmethod
    def _svg_color(value):
        if value is None:
            return "none"
        try:
            r, g, b = value[:3]
            return "#%02x%02x%02x" % (round(r*255), round(g*255), round(b*255))
        except Exception:
            return "none"

    @staticmethod
    def _fmt_point(point):
        return f"{float(point.x):.3f} {float(point.y):.3f}"

    def pdf_vector_svg(self, obj: DocObject) -> bytes:
        """Build a true vector-only SVG from the selected PDF drawing paths."""
        import fitz
        import xml.etree.ElementTree as ET
        source_doc = self.preview_pdf_doc if obj.metadata.get("source_doc") == "docx-preview" else self.pdf_doc
        page = source_doc[obj.page]
        drawings = page.get_drawings()
        indices = obj.metadata.get("drawing_indices", list(range(len(drawings))))
        selected = [drawings[i] for i in indices if 0 <= i < len(drawings)]
        if not selected:
            raise ValueError("该 PDF 区域没有可导出的矢量路径。")
        rect = fitz.Rect(obj.bbox) & page.rect
        width, height = max(1.0, rect.width), max(1.0, rect.height)
        root = ET.Element("svg", {
            "xmlns": "http://www.w3.org/2000/svg",
            "version": "1.1",
            "width": f"{width:.3f}pt",
            "height": f"{height:.3f}pt",
            "viewBox": f"{rect.x0:.3f} {rect.y0:.3f} {width:.3f} {height:.3f}",
        })
        for drawing in selected:
            commands = []
            for item in drawing.get("items", []):
                op = item[0]
                if op == "l":
                    commands.append(f"M {self._fmt_point(item[1])} L {self._fmt_point(item[2])}")
                elif op == "c":
                    commands.append(
                        f"M {self._fmt_point(item[1])} C {self._fmt_point(item[2])} "
                        f"{self._fmt_point(item[3])} {self._fmt_point(item[4])}"
                    )
                elif op == "re":
                    r = fitz.Rect(item[1])
                    commands.append(
                        f"M {r.x0:.3f} {r.y0:.3f} L {r.x1:.3f} {r.y0:.3f} "
                        f"L {r.x1:.3f} {r.y1:.3f} L {r.x0:.3f} {r.y1:.3f} Z"
                    )
                elif op == "qu":
                    q = item[1]
                    commands.append(
                        f"M {self._fmt_point(q.ul)} L {self._fmt_point(q.ur)} "
                        f"L {self._fmt_point(q.lr)} L {self._fmt_point(q.ll)} Z"
                    )
            if not commands:
                continue
            if drawing.get("closePath"):
                commands[-1] += " Z"
            attrs = {
                "d": " ".join(commands),
                "fill": self._svg_color(drawing.get("fill")),
                "stroke": self._svg_color(drawing.get("color")),
                "stroke-width": f"{float(drawing.get('width') or 0.6):.3f}",
                "stroke-linecap": "round",
                "stroke-linejoin": "round",
            }
            fill_opacity = drawing.get("fill_opacity")
            stroke_opacity = drawing.get("stroke_opacity")
            if fill_opacity is not None:
                attrs["fill-opacity"] = str(float(fill_opacity))
            if stroke_opacity is not None:
                attrs["stroke-opacity"] = str(float(stroke_opacity))
            if drawing.get("even_odd"):
                attrs["fill-rule"] = "evenodd"
            root.append(ET.Element("path", attrs))
        return ET.tostring(root, encoding="utf-8", xml_declaration=True)

    def pdf_region_png(self, objects: list[DocObject], scale: float = 3.0) -> bytes:
        import fitz
        if not objects:
            raise ValueError("未选择 PDF 对象")
        page_index = objects[0].page
        source_doc = self.preview_pdf_doc if objects[0].metadata.get("source_doc") == "docx-preview" else self.pdf_doc
        page = source_doc[page_index]
        rects = [fitz.Rect(o.bbox) for o in objects if o.bbox]
        if not rects:
            raise ValueError("所选对象没有边界框")
        rect = rects[0]
        for r in rects[1:]:
            rect |= r
        pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), clip=rect & page.rect, alpha=False)
        return pix.tobytes("png")

    def _pptx_objects(self, page_index: int) -> list[DocObject]:
        from pptx.enum.shapes import MSO_SHAPE_TYPE
        slide = self.pptx.slides[page_index]

        # PowerPoint stores shapes inside a group in that group's child coordinate
        # space.  The coordinates exposed by python-pptx are therefore not slide
        # coordinates for descendants.  Carry the group transforms down the tree so
        # the preview overlay and the rendered slide use the same coordinate space.
        identity = (1.0, 0.0, 0.0, 1.0, 0.0, 0.0)

        def compose(a, b):
            """Return affine transform a(b(point)); matrices use x'=ax+cy+e."""
            aa, ab, ac, ad, ae, af = a
            ba, bb, bc, bd, be, bf = b
            return (
                aa * ba + ac * bb, ab * ba + ad * bb,
                aa * bc + ac * bd, ab * bc + ad * bd,
                aa * be + ac * bf + ae, ab * be + ad * bf + af,
            )

        def translation(x, y):
            return (1.0, 0.0, 0.0, 1.0, float(x), float(y))

        def scale(x, y):
            return (float(x), 0.0, 0.0, float(y), 0.0, 0.0)

        def point(matrix, x, y):
            a, b, c, d, e, f = matrix
            return (a * x + c * y + e, b * x + d * y + f)

        def transform_bbox(bbox, matrix):
            x0, y0, x1, y1 = bbox
            points = [point(matrix, x, y) for x, y in (
                (x0, y0), (x1, y0), (x1, y1), (x0, y1)
            )]
            xs, ys = zip(*points)
            return (min(xs), min(ys), max(xs), max(ys))

        def shape_rotation(shape, x0, y0, width, height):
            degrees = float(getattr(shape, "rotation", 0.0) or 0.0)
            if not degrees:
                return identity
            radians = math.radians(degrees)
            cosine, sine = math.cos(radians), math.sin(radians)
            cx, cy = x0 + width / 2.0, y0 + height / 2.0
            rotate = (cosine, sine, -sine, cosine, 0.0, 0.0)
            return compose(
                translation(cx, cy),
                compose(rotate, translation(-cx, -cy)),
            )

        def group_transform(shape):
            xfrm = shape._element.grpSpPr.xfrm
            off, ext = xfrm.off, xfrm.ext
            child_off, child_ext = xfrm.chOff, xfrm.chExt
            # Malformed or empty group extents are rare; keeping identity is safer
            # than turning all descendant coordinates into infinities.
            if not ext.cx or not ext.cy or not child_ext.cx or not child_ext.cy:
                return identity
            sx, sy = ext.cx / child_ext.cx, ext.cy / child_ext.cy
            base = compose(
                translation(off.x, off.y),
                compose(scale(sx, sy), translation(-child_off.x, -child_off.y)),
            )
            cx, cy = off.x + ext.cx / 2.0, off.y + ext.cy / 2.0
            flip = (
                -1.0 if xfrm.flipH else 1.0, 0.0,
                0.0, -1.0 if xfrm.flipV else 1.0, 0.0, 0.0,
            )
            flip_transform = compose(
                translation(cx, cy),
                compose(flip, translation(-cx, -cy)),
            )
            degrees = float(getattr(shape, "rotation", 0.0) or 0.0)
            if degrees:
                radians = math.radians(degrees)
                cosine, sine = math.cos(radians), math.sin(radians)
                rotate = (cosine, sine, -sine, cosine, 0.0, 0.0)
                rotation_transform = compose(
                    translation(cx, cy),
                    compose(rotate, translation(-cx, -cy)),
                )
                transform = compose(rotation_transform, flip_transform)
            else:
                transform = flip_transform
            return compose(transform, base)

        def walk(shapes, prefix: list[int] | None = None, parent_name: str = "", parent_transform=identity) -> list[DocObject]:
            items = []
            for index, shape in enumerate(shapes, 1):
                path = list(prefix or []) + [index]
                name = getattr(shape, "name", f"对象 {index}")
                left, top = float(shape.left), float(shape.top)
                width, height = float(shape.width), float(shape.height)
                local_bbox = (left, top, left + width, top + height)
                # Rotation belongs to the shape itself; the inherited matrix then
                # maps all four corners through any enclosing group transforms.
                bbox = transform_bbox(local_bbox, compose(parent_transform, shape_rotation(shape, left, top, width, height)))
                label_prefix = f"{parent_name}/" if parent_name else ""
                if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                    obj = DocObject(
                        uid=f"pptx-{page_index}-shape-{shape.shape_id}", label=f"组合：{label_prefix}{name}",
                        kind="group", page=page_index, bbox=bbox,
                        metadata={"shape_id": int(shape.shape_id), "shape_path": path, "native_kind": "pptx-shape"}
                    )
                    child_transform = compose(parent_transform, group_transform(shape))
                    obj.children = walk(shape.shapes, path, name, child_transform)
                    items.append(obj)
                elif getattr(shape, "has_table", False):
                    table = shape.table
                    rows = [[cell.text or "" for cell in row.cells] for row in table.rows]
                    html_rows = "".join(
                        "<tr>" + "".join(f"<td>{self._html_escape(cell)}</td>" for cell in row) + "</tr>"
                        for row in rows
                    )
                    text_rows = ["\t".join(row) for row in rows]
                    items.append(DocObject(
                        uid=f"pptx-{page_index}-shape-{shape.shape_id}",
                        label=f"表格：{label_prefix}{name} · {len(rows)} 行 × {max((len(row) for row in rows), default=0)} 列",
                        kind="table", page=page_index, bbox=bbox,
                        text="\n".join(text_rows),
                        html=f"<table border='1' cellspacing='0' cellpadding='4'>{html_rows}</table>",
                        metadata={"shape_id": int(shape.shape_id), "shape_path": path, "native_kind": "pptx-table", "rows": rows},
                    ))
                elif shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                    try:
                        blob = shape.image.blob
                        ext = shape.image.ext
                        mime = mimetypes.types_map.get("." + ext, "application/octet-stream")
                        image_size_px = [int(shape.image.size[0]), int(shape.image.size[1])]
                        size = f"{image_size_px[0]}×{image_size_px[1]}"
                    except Exception:
                        blob, mime, image_size_px, size = None, None, None, ""
                    items.append(DocObject(
                        uid=f"pptx-{page_index}-shape-{shape.shape_id}", label=f"图片：{label_prefix}{name} {size}",
                        kind="image", page=page_index, bbox=bbox, data=blob, mime_type=mime,
                        metadata={"shape_id": int(shape.shape_id), "shape_path": path, "native_kind": "pptx-shape",
                                  "image_size_px": image_size_px}
                    ))
                elif getattr(shape, "has_chart", False):
                    items.append(DocObject(
                        uid=f"pptx-{page_index}-shape-{shape.shape_id}", label=f"图表：{label_prefix}{name}",
                        kind="image", page=page_index, bbox=bbox,
                        metadata={"shape_id": int(shape.shape_id), "shape_path": path, "native_kind": "pptx-chart"}
                    ))
                elif getattr(shape, "has_text_frame", False) and shape.text.strip():
                    items.append(DocObject(
                        uid=f"pptx-{page_index}-shape-{shape.shape_id}", label=f"文字：{label_prefix}{name} · {shape.text[:32].replace(chr(10),' ')}",
                        kind="text", page=page_index, bbox=bbox, text=shape.text,
                        metadata={"shape_id": int(shape.shape_id), "shape_path": path, "native_kind": "pptx-shape"}
                    ))
                else:
                    items.append(DocObject(
                        uid=f"pptx-{page_index}-shape-{shape.shape_id}", label=f"形状：{label_prefix}{name}",
                        kind="shape", page=page_index, bbox=bbox,
                        metadata={"shape_id": int(shape.shape_id), "shape_path": path, "native_kind": "pptx-shape"}
                    ))
            return items
        return walk(slide.shapes)

    @staticmethod
    def _rel_target(target: str) -> str:
        if target.startswith("/"):
            return target.lstrip("/")
        return str(PurePosixPath("word") / PurePosixPath(target))

    @staticmethod
    def _xml_text(node: ET.Element, ns: dict) -> str:
        return "".join(t.text or "" for t in node.findall(".//w:t", ns)).strip()

    @staticmethod
    def _placement_key(page_index: int, rect) -> tuple:
        return (page_index, round(rect.x0, 1), round(rect.y0, 1), round(rect.x1, 1), round(rect.y1, 1))

    def _map_docx_preview_images(self):
        """Map rendered PDF image placements back to original DOCX media/OLE assets.

        This is reliable for ordinary inline images and OLE preview objects when the
        export keeps one placement per source object. Unmatched placements retain the
        PDF-extracted resource rather than being guessed.
        """
        if self.preview_pdf_doc is None:
            return
        sources = [o for o in self.docx_objects if o.kind in ("image", "ole") and o.metadata.get("doc_order")]
        sources.sort(key=lambda o: int(o.metadata.get("doc_order", 0)))
        placements = []
        for page_index, page in enumerate(self.preview_pdf_doc):
            for image in page.get_images(full=True):
                xref = image[0]
                if not xref:
                    continue
                for rect in page.get_image_rects(xref) or []:
                    placements.append((page_index, rect))
        placements.sort(key=lambda item: (item[0], item[1].y0, item[1].x0))
        # Source document order and visual PDF order are normally identical for inline
        # media. Only map when the counts agree, avoiding silent wrong-image matches.
        if len(placements) == len(sources):
            for (page_index, rect), source in zip(placements, sources):
                self._docx_preview_map[self._placement_key(page_index, rect)] = source

    def _docx_page_objects(self, page_index: int) -> list[DocObject]:
        page = self.preview_pdf_doc[page_index]
        visual = self._pdf_objects(page, page_index, self.preview_pdf_doc, "docx-preview")
        mapped = []
        for obj in visual:
            if obj.kind == "image" and obj.bbox:
                import fitz
                source = self._docx_preview_map.get(self._placement_key(page_index, fitz.Rect(obj.bbox)))
                if source is not None:
                    metadata = dict(source.metadata)
                    metadata.update({"source_doc": "docx-preview", "preview_bbox": obj.bbox})
                    obj = DocObject(
                        uid=f"docx-page-{page_index}-{source.uid}", label=source.label,
                        kind=source.kind, page=page_index, bbox=obj.bbox,
                        data=source.data, mime_type=source.mime_type, text=source.text,
                        html=source.html, metadata=metadata
                    )
            elif obj.kind == "table":
                source = self._match_docx_table(obj)
                if source is not None:
                    metadata = dict(source.metadata)
                    metadata.update({
                        "source_doc": "docx-preview",
                        "preview_bbox": obj.bbox,
                        "native_table_uid": source.uid,
                        "native_kind": "word-table",
                    })
                    obj = DocObject(
                        uid=obj.uid, label=source.label, kind="table", page=page_index, bbox=obj.bbox,
                        text=source.text, html=source.html, metadata=metadata,
                    )
            mapped.append(obj)
        if page_index == 0 and (self.docx_table_objects or self.docx_paragraph_objects):
            mapped.append(DocObject(
                uid="docx-document-structure", label="Word 原生段落与表格（文档级，可保留格式复制）", kind="group",
                children=self.docx_paragraph_objects + self.docx_table_objects,
                metadata={"source": "docx-structure"}
            ))
        return mapped

    def _match_docx_table(self, preview_table: DocObject) -> DocObject | None:
        """Match one rendered page fragment to its source Word table, if possible."""
        if not self.docx_table_objects:
            return None
        import difflib
        import unicodedata

        def compact(value: str) -> str:
            normalized = unicodedata.normalize("NFKC", value).casefold()
            return "".join(char for char in normalized if char.isalnum())

        preview_rows = preview_table.metadata.get("rows", [])
        preview_cols = max((len(row) for row in preview_rows), default=0)
        preview_text = compact(preview_table.text)
        if not preview_text or not preview_cols:
            return None

        candidates = []
        for source in self.docx_table_objects:
            source_rows = source.metadata.get("rows", [])
            source_cols = max((len(row) for row in source_rows), default=0)
            source_text = compact(source.text)
            if not source_text or not source_cols or abs(source_cols - preview_cols) > 1:
                continue
            text_score = difflib.SequenceMatcher(None, source_text, preview_text, autojunk=False).ratio()
            column_score = 1.0 if source_cols == preview_cols else 0.65
            candidates.append((text_score * column_score, source))
        if not candidates:
            return None
        score, source = max(candidates, key=lambda entry: entry[0])
        return source if score >= 0.28 else None

    def _parse_docx_package(self, path: str):
        ns = {
            "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
            "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
            "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
            "o": "urn:schemas-microsoft-com:office:office",
            "v": "urn:schemas-microsoft-com:vml",
        }
        with zipfile.ZipFile(path) as zf:
            rels = ET.fromstring(zf.read("word/_rels/document.xml.rels"))
            rel_ns = "{http://schemas.openxmlformats.org/package/2006/relationships}"
            relmap = {}
            reltypes = {}
            for rel in rels:
                rid = rel.attrib.get("Id")
                target = rel.attrib.get("Target", "")
                typ = rel.attrib.get("Type", "")
                if rid and target:
                    if target.startswith("/"):
                        resolved = target.lstrip("/")
                    else:
                        resolved = str(PurePosixPath("word") / PurePosixPath(target))
                    relmap[rid] = resolved
                    reltypes[rid] = typ
            root = ET.fromstring(zf.read("word/document.xml"))
            body = root.find("w:body", ns)
            objects: list[DocObject] = []
            paragraph_com_indices = {id(p): i + 1 for i, p in enumerate(root.findall(".//w:p", ns))}
            paragraph_index = 0
            table_index = 0
            inline_index = 0
            shape_index = 0
            visual_order = 0
            image_ref_count = 0
            ole_ref_count = 0
            referenced_media = set()

            def media_obj(rid: str, name: str, bbox=None, native_index=None):
                target = relmap.get(rid)
                if not target or target not in zf.namelist():
                    return None
                data = zf.read(target)
                referenced_media.add(target)
                ext = Path(target).suffix.lower()
                mime = mimetypes.types_map.get(ext, "application/octet-stream")
                return DocObject(
                    uid=f"docx-media-{len(objects)+1}", label=f"{name} · {Path(target).name}",
                    kind="image", data=data, mime_type=mime, bbox=bbox,
                    metadata={"package_path": target, "native_kind": "word-inline", "native_index": native_index, "extension": ext}
                )

            if body is not None:
                for child in list(body):
                    tag = child.tag
                    if tag == f"{{{ns['w']}}}p":
                        paragraph_index += 1
                        text = self._xml_text(child, ns)
                        if text:
                            objects.append(DocObject(
                                uid=f"docx-paragraph-{paragraph_index}",
                                label=f"段落 {paragraph_index} · {text[:52].replace(chr(10),' ')}",
                                kind="text", text=text,
                                metadata={"native_kind": "word-paragraph", "native_index": paragraph_com_indices.get(id(child), paragraph_index)}
                            ))
                        # Walk visual elements in document order. Inline objects and
                        # floating shapes use different Word COM collections.
                        for visual_node in child.iter():
                            if visual_node.tag == f"{{{ns['wp']}}}inline":
                                inline_index += 1
                                visual_order += 1
                                docpr = visual_node.find("wp:docPr", ns)
                                name = docpr.attrib.get("name", f"图片 {inline_index}") if docpr is not None else f"图片 {inline_index}"
                                blip = visual_node.find(".//a:blip", ns)
                                if blip is not None:
                                    rid = blip.attrib.get(f"{{{ns['r']}}}embed")
                                    obj = media_obj(rid, name, None, inline_index)
                                    if obj:
                                        obj.metadata.update({"native_kind": "word-inline", "doc_order": visual_order})
                                        objects.append(obj)
                            elif visual_node.tag == f"{{{ns['wp']}}}anchor":
                                shape_index += 1
                                visual_order += 1
                                docpr = visual_node.find("wp:docPr", ns)
                                name = docpr.attrib.get("name", f"浮动图片 {shape_index}") if docpr is not None else f"浮动图片 {shape_index}"
                                blip = visual_node.find(".//a:blip", ns)
                                if blip is not None:
                                    rid = blip.attrib.get(f"{{{ns['r']}}}embed")
                                    obj = media_obj(rid, name, None, shape_index)
                                    if obj:
                                        obj.metadata.update({"native_kind": "word-shape", "doc_order": visual_order})
                                        objects.append(obj)
                            elif visual_node.tag == f"{{{ns['w']}}}object":
                                inline_index += 1
                                visual_order += 1
                                ole_el = visual_node.find(".//o:OLEObject", ns)
                                ole_rid = ole_el.attrib.get(f"{{{ns['r']}}}id") if ole_el is not None else None
                                preview = visual_node.find(".//v:imagedata", ns)
                                preview_rid = preview.attrib.get(f"{{{ns['r']}}}id") if preview is not None else None
                                ole_path = relmap.get(ole_rid) if ole_rid else None
                                preview_path = relmap.get(preview_rid) if preview_rid else None
                                ole_data = zf.read(ole_path) if ole_path in zf.namelist() else None
                                preview_data = zf.read(preview_path) if preview_path in zf.namelist() else None
                                if ole_path:
                                    referenced_media.add(ole_path)
                                if preview_path:
                                    referenced_media.add(preview_path)
                                objects.append(DocObject(
                                    uid=f"docx-ole-{inline_index}", label=f"嵌入对象/公式 {inline_index} · {Path(ole_path or preview_path or 'OLE').name}",
                                    kind="ole", data=ole_data, mime_type="application/vnd.openxmlformats-officedocument.oleObject",
                                    metadata={"native_kind": "word-inline", "native_index": inline_index, "doc_order": visual_order,
                                              "ole_path": ole_path, "preview_path": preview_path,
                                              "preview_data": preview_data}
                                ))
                    elif tag == f"{{{ns['w']}}}tbl":
                        table_index += 1
                        rows = []
                        for tr in child.findall("w:tr", ns):
                            row = []
                            for tc in tr.findall("w:tc", ns):
                                cell_text = "\n".join(self._xml_text(p, ns) for p in tc.findall("w:p", ns)).strip()
                                row.append(cell_text)
                            rows.append(row)
                        text_rows = ["\t".join(r) for r in rows]
                        html_rows = "".join("<tr>" + "".join(f"<td>{self._html_escape(c)}</td>" for c in r) + "</tr>" for r in rows)
                        table_obj = DocObject(
                            uid=f"docx-table-{table_index}", label=f"表格 {table_index} · {len(rows)} 行 × {max((len(r) for r in rows), default=0)} 列",
                            kind="table", text="\n".join(text_rows),
                            html=f"<table border='1' cellspacing='0' cellpadding='4'>{html_rows}</table>",
                            metadata={"native_kind": "word-table", "native_index": table_index, "rows": rows}
                        )
                        # Cell children can be copied individually; the parent remains
                        # a real table object when copied in Word on Windows.
                        for ri, tr in enumerate(child.findall("w:tr", ns), 1):
                            for ci, tc in enumerate(tr.findall("w:tc", ns), 1):
                                cell_text = "\n".join(self._xml_text(p, ns) for p in tc.findall("w:p", ns)).strip()
                                if cell_text:
                                    table_obj.children.append(DocObject(
                                        uid=f"docx-table-{table_index}-r{ri}-c{ci}",
                                        label=f"第 {ri} 行第 {ci} 列：{cell_text[:30]}",
                                        kind="text", text=cell_text,
                                        metadata={"native_kind": "word-table-cell", "table_index": table_index, "row": ri, "col": ci}
                                    ))
                        objects.append(table_obj)

            # Include all remaining package media; WMF preview assets are often
            # linked to embedded OLE equations and should not silently disappear.
            for name in zf.namelist():
                if name.startswith("word/media/") and name not in referenced_media:
                    data = zf.read(name)
                    ext = Path(name).suffix.lower()
                    objects.append(DocObject(
                        uid=f"docx-asset-{len(objects)+1}", label=f"未映射媒体资源 · {Path(name).name}",
                        kind="image", data=data, mime_type=mimetypes.types_map.get(ext, "application/octet-stream"),
                        metadata={"package_path": name, "extension": ext, "unmapped": True}
                    ))
            self.docx_objects = objects
            self.docx_table_objects = [o for o in objects if o.kind == "table"]
            self.docx_paragraph_objects = [o for o in objects if o.kind == "text" and o.metadata.get("native_kind") == "word-paragraph"]

    @staticmethod
    def _html_escape(value: str) -> str:
        return value.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")

    def count_objects(self, objects: list[DocObject] | None = None) -> dict[str, int]:
        objects = self.docx_objects if objects is None and self.kind == "docx" else (objects or [])
        counts: dict[str, int] = {}
        def walk(items):
            for obj in items:
                counts[obj.kind] = counts.get(obj.kind, 0) + 1
                walk(obj.children)
        walk(objects)
        return counts

    def svg_for_docx_wmf(self, obj: DocObject) -> bytes | None:
        """Convert a WMF preview to a tightly-cropped SVG.

        LibreOffice's direct WMF->SVG export uses an A4 drawing page. We therefore
        use Inkscape's geometry query (when available) to crop the SVG viewBox to
        the actual drawing. Without a crop-capable tool we return None rather than
        putting a tiny equation on an entire A4-sized clipboard canvas.
        """
        import shutil
        import xml.etree.ElementTree as ET
        data = obj.data
        if obj.metadata.get("extension") != ".wmf" or not data:
            return None
        src = os.path.join(self.temp_dir, f"{obj.uid}.wmf")
        Path(src).write_bytes(data)
        outdir = os.path.join(self.temp_dir, "wmf_svg")
        os.makedirs(outdir, exist_ok=True)
        try:
            self._run_soffice(["--headless", "--convert-to", "svg", "--outdir", outdir, src], timeout=90)
            svg_path = os.path.join(outdir, Path(src).stem + ".svg")
            if not os.path.exists(svg_path):
                return None
            inkscape = shutil.which("inkscape")
            if not inkscape:
                return None
            query = subprocess.run([inkscape, "--query-all", svg_path], capture_output=True, text=True, timeout=45)
            if query.returncode != 0 or not query.stdout.strip():
                return None
            query_rows = []
            for line in query.stdout.splitlines():
                fields = line.strip().split(",")
                if len(fields) >= 5:
                    try:
                        query_rows.append((fields[0], *map(float, fields[-4:])))
                    except Exception:
                        pass
            if not query_rows:
                return None
            root = ET.parse(svg_path).getroot()
            root_box = query_rows[0][1:]
            root_area = max(1.0, root_box[2] * root_box[3])
            candidates = []
            drawable_prefixes = ("text", "tspan", "path", "ellipse", "polygon", "polyline", "circle", "line", "rect", "image")
            for element_id, ex, ey, ew, eh in query_rows[1:]:
                if not element_id.startswith(drawable_prefixes) or ew <= 0 or eh <= 0:
                    continue
                if element_id.startswith("rect") and ew * eh >= root_area * 0.90:
                    continue  # ignore the full-canvas background rectangle
                candidates.append((ex, ey, ew, eh))
            if candidates:
                x0 = min(v[0] for v in candidates)
                y0 = min(v[1] for v in candidates)
                x1 = max(v[0] + v[2] for v in candidates)
                y1 = max(v[1] + v[3] for v in candidates)
                x, y, w, h = x0, y0, x1-x0, y1-y0
            else:
                _, x, y, w, h = query_rows[0]
            if w <= 0 or h <= 0:
                return None
            viewbox = [float(v) for v in root.attrib.get("viewBox", "0 0 21000 29700").replace(",", " ").split()]
            if len(viewbox) != 4:
                return None
            vb_x, vb_y, vb_w, vb_h = viewbox
            width_attr = root.attrib.get("width", "210mm")
            height_attr = root.attrib.get("height", "297mm")
            def css_px(value, fallback):
                import re
                m = re.match(r"^([0-9.]+)([a-zA-Z]*)$", value)
                if not m:
                    return fallback
                number, unit = float(m.group(1)), m.group(2).lower()
                factors = {"": 1.0, "px": 1.0, "pt": 96/72, "in": 96, "mm": 96/25.4, "cm": 96/2.54}
                return number * factors.get(unit, 1.0)
            css_w = css_px(width_attr, 793.7)
            css_h = css_px(height_attr, 1122.5)
            sx = vb_w / css_w
            sy = vb_h / css_h
            pad = 2.0
            crop_x = vb_x + max(0.0, x-pad) * sx
            crop_y = vb_y + max(0.0, y-pad) * sy
            crop_w = min(vb_x+vb_w-crop_x, (w+2*pad)*sx)
            crop_h = min(vb_y+vb_h-crop_y, (h+2*pad)*sy)
            root.set("viewBox", f"{crop_x:.3f} {crop_y:.3f} {crop_w:.3f} {crop_h:.3f}")
            root.set("width", f"{w+2*pad:.3f}px")
            root.set("height", f"{h+2*pad:.3f}px")
            return ET.tostring(root, encoding="utf-8", xml_declaration=True)
        except Exception:
            return None

    @staticmethod
    def png_from_svg(svg: bytes | None) -> bytes | None:
        """Optional PNG fallback rendered from an already-cropped SVG."""
        if not svg:
            return None
        try:
            import cairosvg
            return cairosvg.svg2png(bytestring=svg)
        except Exception:
            return None

    def png_for_docx_wmf(self, obj: DocObject) -> bytes | None:
        return self.png_from_svg(self.svg_for_docx_wmf(obj))

    def payload_for_object(self, obj: DocObject) -> dict[str, bytes | str]:
        """Return MIME payloads for clipboard. Does not touch OS clipboard."""
        if obj.kind == "image" and obj.data:
            ext = obj.metadata.get("extension", "")
            if ext == ".wmf":
                svg = self.svg_for_docx_wmf(obj)
                if svg:
                    png = self.png_from_svg(svg)
                    payload = {"image/svg+xml": svg, "text/plain": f"矢量图像：{obj.label}"}
                    if png:
                        payload["image/png"] = png
                    return payload
                return {"image/x-wmf": obj.data}
            mime = obj.mime_type or mimetypes.guess_type(obj.metadata.get("package_path", "image.png"))[0] or "application/octet-stream"
            return {mime: obj.data}
        if obj.kind == "text":
            return {"text/plain": obj.text, "text/html": obj.html or f"<div>{self._html_escape(obj.text)}</div>"}
        if obj.kind == "table":
            return {"text/plain": obj.text, "text/html": obj.html}
        if obj.kind == "vector":
            svg = self.pdf_vector_svg(obj)
            png = self.pdf_region_png([obj], 3.0)
            return {"image/svg+xml": svg, "image/png": png}
        if obj.kind == "ole":
            preview = obj.metadata.get("preview_data")
            preview_path = obj.metadata.get("preview_path", "")
            if preview and str(preview_path).lower().endswith(".wmf"):
                temp_obj = DocObject(obj.uid + "-preview", obj.label, "image", data=preview,
                                     metadata={"extension": ".wmf"})
                svg = self.svg_for_docx_wmf(temp_obj)
                if svg:
                    png = self.png_from_svg(svg)
                    payload = {"image/svg+xml": svg, "text/plain": f"OLE 对象预览：{obj.label}"}
                    if png:
                        payload["image/png"] = png
                    return payload
                return {"image/x-wmf": preview}
            if preview:
                return {"image/png": preview}
            raise RuntimeError("该 OLE 对象没有可用预览；在 Windows + Word 中可尝试原生复制。")
        raise RuntimeError(f"当前对象类型暂不支持复制：{obj.kind}")

    def pdf_region_payload(self, objects: list[DocObject]) -> dict[str, bytes]:
        if not objects or any(o.page != objects[0].page for o in objects):
            raise ValueError("PDF 多选对象必须来自同一页。")
        # The cropped SVG retains vector content; PNG is included as compatibility fallback.
        rect_objects = [o for o in objects if o.bbox]
        if not rect_objects:
            raise ValueError("所选对象没有可复制边界。")
        import fitz, xml.etree.ElementTree as ET
        source_doc = self.preview_pdf_doc if objects[0].metadata.get("source_doc") == "docx-preview" else self.pdf_doc
        page = source_doc[objects[0].page]
        rects = [fitz.Rect(o.bbox) for o in rect_objects]
        rect = rects[0]
        for r in rects[1:]: rect |= r
        rect &= page.rect
        root = ET.fromstring(page.get_svg_image(text_as_path=False))
        root.set("viewBox", f"{rect.x0:.3f} {rect.y0:.3f} {max(1,rect.width):.3f} {max(1,rect.height):.3f}")
        root.set("width", f"{max(1,rect.width):.3f}pt")
        root.set("height", f"{max(1,rect.height):.3f}pt")
        svg = ET.tostring(root, encoding="utf-8", xml_declaration=True)
        png = self.pdf_region_png(rect_objects, 3.0)
        return {"image/svg+xml": svg, "image/png": png}
