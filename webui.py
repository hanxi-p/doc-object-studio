# -*- coding: utf-8 -*-
"""Local browser UI for 文档对象工坊."""
from __future__ import annotations

import html
import io
import json
import mimetypes
import os
import base64
import secrets
import shutil
import sys
import tempfile
import threading
import time
import urllib.parse
import webbrowser
from collections import Counter
from email.parser import BytesParser
from email.policy import default
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path, PurePosixPath
from urllib.parse import parse_qs, urlsplit
from uuid import uuid4

from document_engine import DocumentEngine, DocObject

APP_NAME = "文档对象工坊"
MAX_UPLOAD_BYTES = 512 * 1024 * 1024
WEB_FILES = {
    "/": "index.html",
    "/app.css": "app.css",
    "/app.js": "app.js",
    "/favicon.svg": "favicon.svg",
    "/favicon.ico": "app-icon.ico",
}
SUPPORTED = {".pdf", ".pptx", ".docx"}


def resource_path(relative: str) -> Path:
    root = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return root / relative


def flatten(objects):
    for obj in objects:
        yield obj
        yield from flatten(obj.children)


def serialize_object(obj: DocObject, display_transform=None) -> dict:
    display_bbox = list(obj.bbox) if obj.bbox else None
    if display_bbox and display_transform is not None:
        try:
            import fitz

            display_bbox = list(fitz.Rect(obj.bbox) * display_transform)
        except Exception:
            pass
    return {
        "uid": obj.uid,
        "label": obj.label,
        "kind": obj.kind,
        "page": obj.page,
        "bbox": list(obj.bbox) if obj.bbox else None,
        "display_bbox": display_bbox,
        "text": obj.text or "",
        "data_size": len(obj.data) if obj.data else 0,
        "mime_type": obj.mime_type or "",
        "image_size_px": obj.metadata.get("image_size_px"),
        "native_table_uid": obj.metadata.get("native_table_uid"),
        "z_path": obj.metadata.get("shape_path", []),
        "children": [serialize_object(child, display_transform) for child in obj.children],
    }


def serialize_selectable_object(obj: DocObject) -> dict:
    """Only expose whole-document picture/table entries to the selection UI."""
    return {
        "uid": obj.uid,
        "label": obj.label,
        "kind": obj.kind,
        "page": obj.page,
        "data_size": len(obj.data) if obj.data else 0,
        "mime_type": obj.mime_type or "",
        "image_size_px": obj.metadata.get("image_size_px"),
        "z_path": obj.metadata.get("shape_path", []),
    }


def payload_for_selected_objects(engine: DocumentEngine, objects: list[DocObject]) -> dict[str, bytes | str]:
    """Build a lossless HTML collection when several whole-document objects are checked."""
    if len(objects) == 1:
        return engine.payload_for_object(objects[0])

    html_parts = []
    text_parts = []
    for obj in objects:
        page_label = f"（第 {obj.page + 1} 页）" if obj.page is not None else ""
        heading = f"<h3>{html.escape(obj.label + page_label)}</h3>"
        if obj.kind == "image" and obj.data:
            mime = obj.mime_type or mimetypes.guess_type(obj.label)[0] or "application/octet-stream"
            source = "data:" + mime + ";base64," + base64.b64encode(obj.data).decode("ascii")
            html_parts.append(heading + f'<img alt="{html.escape(obj.label)}" src="{source}">')
            text_parts.append(obj.label + page_label)
        elif obj.kind == "table":
            html_parts.append(heading + (obj.html or f"<pre>{html.escape(obj.text)}</pre>"))
            text_parts.append(obj.text or obj.label)
        elif obj.kind == "image" and obj.metadata.get("native_kind") == "pptx-chart":
            raise ValueError("PowerPoint 原生图表请单独勾选复制，以保留可编辑图表格式。")
        else:
            raise ValueError("当前所选对象无法组成多对象剪贴板内容，请逐个复制。")
    return {"text/plain": "\n\n".join(text_parts), "text/html": "<div>" + "".join(html_parts) + "</div>"}


def _cf_html(value: str) -> bytes:
    prefix = "<html><body><!--StartFragment-->"
    suffix = "<!--EndFragment--></body></html>"
    fragment = value
    if "<" not in fragment:
        fragment = f"<div>{html.escape(fragment)}</div>"
    prefix_bytes = prefix.encode("utf-8")
    fragment_bytes = fragment.encode("utf-8")
    suffix_bytes = suffix.encode("utf-8")
    header_template = (
        "Version:0.9\r\nStartHTML:{:010d}\r\nEndHTML:{:010d}\r\n"
        "StartFragment:{:010d}\r\nEndFragment:{:010d}\r\n"
    )
    header_length = len(header_template.format(0, 0, 0, 0).encode("ascii"))
    start_html = header_length
    start_fragment = start_html + len(prefix_bytes)
    end_fragment = start_fragment + len(fragment_bytes)
    end_html = end_fragment + len(suffix_bytes)
    header = header_template.format(start_html, end_html, start_fragment, end_fragment).encode("ascii")
    return header + prefix_bytes + fragment_bytes + suffix_bytes


def _image_dib(data: bytes) -> bytes | None:
    try:
        from PIL import Image

        with Image.open(io.BytesIO(data)) as image:
            output = io.BytesIO()
            image.convert("RGB").save(output, format="BMP")
        return output.getvalue()[14:]
    except Exception:
        return None


def set_windows_clipboard(payload: dict[str, bytes | str]) -> None:
    if not sys.platform.startswith("win"):
        raise RuntimeError("复制到系统剪贴板需要 Windows。")
    import win32clipboard
    import win32con

    opened = False
    last_error = None
    for attempt in range(8):
        try:
            win32clipboard.OpenClipboard()
            opened = True
            break
        except Exception as exc:
            last_error = exc
            time.sleep(0.08 * (attempt + 1))
    if not opened:
        raise RuntimeError(f"Windows 剪贴板暂时被占用：{last_error}")

    try:
        win32clipboard.EmptyClipboard()
        dib = None
        for mime, value in payload.items():
            if mime == "text/plain":
                win32clipboard.SetClipboardText(str(value), win32con.CF_UNICODETEXT)
            elif mime == "text/html":
                win32clipboard.SetClipboardData(
                    win32clipboard.RegisterClipboardFormat("HTML Format"), _cf_html(str(value))
                )
            elif mime in ("image/png", "image/jpeg", "image/bmp"):
                raw = value.encode("utf-8") if isinstance(value, str) else value
                name = "PNG" if mime == "image/png" else "JFIF" if mime == "image/jpeg" else "image/bmp"
                win32clipboard.SetClipboardData(win32clipboard.RegisterClipboardFormat(name), raw)
                if dib is None:
                    dib = _image_dib(raw)
            elif mime.startswith("image/"):
                raw = value.encode("utf-8") if isinstance(value, str) else value
                win32clipboard.SetClipboardData(win32clipboard.RegisterClipboardFormat(mime), raw)
            else:
                raw = value.encode("utf-8") if isinstance(value, str) else value
                win32clipboard.SetClipboardData(win32clipboard.RegisterClipboardFormat(mime), raw)
        if dib:
            win32clipboard.SetClipboardData(win32con.CF_DIB, dib)
    finally:
        win32clipboard.CloseClipboard()


def office_native_copy(engine: DocumentEngine, obj: DocObject) -> None:
    if not sys.platform.startswith("win"):
        raise RuntimeError("Office 原生复制需要 Windows + Microsoft Office。")
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    app = None
    document = None
    try:
        if engine.kind == "pptx":
            app = win32com.client.DispatchEx("PowerPoint.Application")
            document = app.Presentations.Open(os.path.abspath(engine.path), ReadOnly=True, WithWindow=False)
            slide = document.Slides.Item(int(obj.page) + 1)
            shape_path = obj.metadata.get("shape_path", [])
            if not shape_path:
                raise RuntimeError("没有找到 PowerPoint 对象路径。")
            shape = slide.Shapes.Item(int(shape_path[0]))
            for child_index in shape_path[1:]:
                shape = shape.GroupItems.Item(int(child_index))
            shape.Copy()
        elif engine.kind == "docx":
            app = win32com.client.DispatchEx("Word.Application")
            app.Visible = False
            app.DisplayAlerts = 0
            document = app.Documents.Open(
                os.path.abspath(engine.path), ReadOnly=True, AddToRecentFiles=False, Visible=False
            )
            native_kind = obj.metadata.get("native_kind")
            if native_kind == "word-paragraph":
                document.Paragraphs.Item(int(obj.metadata["native_index"])).Range.Copy()
            elif native_kind == "word-table":
                document.Tables.Item(int(obj.metadata["native_index"])).Range.Copy()
            elif native_kind == "word-table-cell":
                table = document.Tables.Item(int(obj.metadata["table_index"]))
                table.Cell(int(obj.metadata["row"]), int(obj.metadata["col"])).Range.Copy()
            elif native_kind == "word-inline":
                document.InlineShapes.Item(int(obj.metadata["native_index"])).Range.Copy()
            elif native_kind == "word-shape":
                document.Shapes.Item(int(obj.metadata["native_index"])).Copy()
            else:
                raise RuntimeError("该 Word 对象没有可用的原生复制索引。")
        else:
            raise RuntimeError("不支持的 Office 文档类型。")
    finally:
        if document is not None:
            try:
                if engine.kind == "pptx":
                    document.Close()
                else:
                    document.Close(False)
            except Exception:
                pass
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


class Studio:
    def __init__(self):
        self.engine = DocumentEngine()
        self.engine.progress_callback = self._engine_progress
        self.upload_dir = Path(tempfile.mkdtemp(prefix="doc_object_studio_upload_"))
        self.uploaded_path: Path | None = None
        self.file_name = ""
        self.page_cache: dict[int, list[DocObject]] = {}
        self.thumbnail_cache: dict[int, bytes] = {}
        self.document_object_cache: list[DocObject] | None = None
        self.document_object_map: dict[str, DocObject] = {}
        self._progress_lock = threading.Lock()
        self.progress = {"active": False, "percent": 0, "message": "准备就绪"}

    def _engine_progress(self, percent: int, message: str):
        with self._progress_lock:
            self.progress = {"active": True, "percent": percent, "message": message}

    def progress_state(self) -> dict:
        with self._progress_lock:
            return dict(self.progress)

    def open_upload(self, file_name: str, data: bytes) -> dict:
        safe_name = PurePosixPath(file_name.replace("\\", "/")).name
        suffix = Path(safe_name).suffix.lower()
        if suffix not in SUPPORTED:
            raise ValueError("目前只支持 PDF、PPTX 和 DOCX 文件。")
        if not data:
            raise ValueError("文件内容为空。")

        path = self.upload_dir / f"{uuid4().hex}{suffix}"
        path.write_bytes(data)
        with self._progress_lock:
            self.progress = {"active": True, "percent": 3, "message": "文件已上传，正在解析"}
        try:
            pages = self.engine.open(str(path))
        except Exception:
            with self._progress_lock:
                self.progress = {"active": False, "percent": 0, "message": "文档打开失败"}
            path.unlink(missing_ok=True)
            raise

        old_path = self.uploaded_path
        self.uploaded_path = path
        self.file_name = safe_name
        self.page_cache.clear()
        self.thumbnail_cache.clear()
        self.document_object_cache = None
        self.document_object_map.clear()
        with self._progress_lock:
            self.progress = {"active": False, "percent": 100, "message": "页面预览已就绪"}
        if old_path and old_path != path:
            old_path.unlink(missing_ok=True)
        return {"file_name": self.file_name, "kind": self.engine.kind, "pages": pages}

    def page_objects(self, page_index: int) -> list[DocObject]:
        if page_index not in self.page_cache:
            self.page_cache[page_index] = self.engine.objects_for_page(page_index)
        return self.page_cache[page_index]

    def document_objects(self) -> list[DocObject]:
        if self.document_object_cache is not None:
            return self.document_object_cache
        objects = []
        seen = set()
        docx_preview_tables: dict[str, DocObject] = {}
        page_count = self.engine.page_count()
        with self._progress_lock:
            self.progress = {"active": True, "percent": 0, "message": "正在索引整份文档的图片与表格"}
        try:
            for page_index in range(page_count):
                for obj in flatten(self.page_objects(page_index)):
                    if obj.kind not in ("image", "table"):
                        continue
                    if self.engine.kind == "docx" and obj.kind == "table":
                        native_uid = obj.metadata.get("native_table_uid")
                        if native_uid:
                            docx_preview_tables.setdefault(native_uid, obj)
                        # The DOCX package contains the complete source table list.
                        # Ignore rendered PDF table fragments here so split tables
                        # and unmatched layout detections do not appear as duplicates.
                        continue
                    if obj.uid in seen:
                        continue
                    if self.engine.kind == "pdf" and obj.kind == "image" and obj.bbox:
                        # Some PDFs encode individual text strokes as thousands of
                        # sub-2.5pt image strips. They are page-rendering fragments,
                        # not useful whole-document figures for the copy list.
                        placed_width = abs(obj.bbox[2] - obj.bbox[0])
                        placed_height = abs(obj.bbox[3] - obj.bbox[1])
                        if min(placed_width, placed_height) < 2.5:
                            continue
                    # A picture must contain original bytes, unless it is an editable
                    # PowerPoint chart copied through Office's native object route.
                    if obj.kind == "image" and not obj.data and obj.metadata.get("native_kind") != "pptx-chart":
                        continue
                    seen.add(obj.uid)
                    objects.append(obj)
                with self._progress_lock:
                    self.progress = {
                        "active": True,
                        "percent": int((page_index + 1) / max(page_count, 1) * 100),
                        "message": f"正在检查第 {page_index + 1} / {page_count} 页的图片与表格",
                    }

            # DOCX page fragments that match the same source table share one
            # document-wide checkbox. The page-level IDs remain aliases for preview
            # outlines; the copy route uses Word's complete native table.
            if self.engine.kind == "docx":
                for table in self.engine.docx_table_objects:
                    preview = docx_preview_tables.get(table.uid)
                    metadata = dict(table.metadata)
                    if preview is not None:
                        metadata.update({
                            "source_doc": "docx-preview",
                            "native_table_uid": table.uid,
                            "preview_bbox": preview.bbox,
                            "preview_uid": preview.uid,
                        })
                        obj = DocObject(
                            uid=table.uid, label=table.label, kind="table", page=preview.page,
                            bbox=preview.bbox, text=table.text, html=table.html, metadata=metadata,
                        )
                    else:
                        obj = table
                    if obj.uid not in seen:
                        seen.add(obj.uid)
                        objects.append(obj)

                # DOCX package images can exist outside the rendered page (for
                # example in headers). Keep original assets when no placement mapped.
                mapped_media = {
                    obj.metadata.get("package_path")
                    for obj in objects
                    if obj.kind == "image" and obj.metadata.get("package_path")
                }
                for obj in self.engine.docx_objects:
                    if obj.kind == "image" and obj.data and obj.uid not in seen and obj.metadata.get("package_path") not in mapped_media:
                        seen.add(obj.uid)
                        objects.append(obj)
        except Exception:
            with self._progress_lock:
                self.progress = {"active": False, "percent": 0, "message": "全文对象索引失败"}
            raise

        self.document_object_cache = objects
        self.document_object_map = {obj.uid: obj for obj in objects}
        with self._progress_lock:
            self.progress = {"active": False, "percent": 100, "message": "图片与表格清单已就绪"}
        return objects

    def selected_document_objects(self, uids: list[str]) -> list[DocObject]:
        available = self.document_object_map if self.document_object_cache is not None else {
            obj.uid: obj for obj in self.document_objects()
        }
        if not uids or any(uid not in available for uid in uids):
            raise ValueError("请选择文档中的图片或表格对象。")
        objects = [available[uid] for uid in uids]
        if any(obj.kind not in ("image", "table") for obj in objects):
            raise ValueError("仅支持复制图片和表格。")
        return objects

    def state(self, page_index: int) -> dict:
        if not self.engine.path:
            return {"opened": False, "version": self.engine.VERSION}
        page_count = self.engine.page_count()
        page_index = max(0, min(page_index, page_count - 1))
        preview_path = Path(self.engine.preview_paths[page_index])
        from PIL import Image

        with Image.open(preview_path) as preview:
            preview_width, preview_height = preview.size
        if self.engine.kind == "pptx":
            coordinate_width = self.engine.slide_width
            coordinate_height = self.engine.slide_height
            coordinate_space = "EMU"
            display_transform = None
        else:
            source = self.engine.pdf_doc if self.engine.kind == "pdf" else self.engine.preview_pdf_doc
            page = source[page_index]
            page_rect = page.rect
            coordinate_width = float(page_rect.width)
            coordinate_height = float(page_rect.height)
            coordinate_space = "PDF point"
            display_transform = page.rotation_matrix
        objects = self.page_objects(page_index)
        return {
            "opened": True,
            "version": self.engine.VERSION,
            "file_name": self.file_name or Path(self.engine.path).name,
            "kind": self.engine.kind,
            "page_count": page_count,
            "page_index": page_index,
            "slide_width": self.engine.slide_width,
            "coordinate_width": coordinate_width,
            "coordinate_height": coordinate_height,
            "coordinate_space": coordinate_space,
            "page_rotation": 0 if display_transform is None else int(page.rotation),
            "preview_width": preview_width,
            "preview_height": preview_height,
            "render_scale": self.engine.render_scale,
            "objects": [serialize_object(obj, display_transform) for obj in objects],
        }

    def thumbnail(self, page_index: int) -> bytes:
        if not 0 <= page_index < self.engine.page_count():
            raise ValueError("页码无效。")
        cached = self.thumbnail_cache.get(page_index)
        if cached is not None:
            return cached
        from PIL import Image

        source = Path(self.engine.preview_paths[page_index])
        with Image.open(source) as image:
            image.thumbnail((260, 150), Image.Resampling.LANCZOS)
            if image.mode != "RGB":
                image = image.convert("RGB")
            output = io.BytesIO()
            image.save(output, format="JPEG", quality=90, optimize=True)
        data = output.getvalue()
        self.thumbnail_cache[page_index] = data
        return data

    def selected_objects(self, page_index: int, uids: list[str]) -> list[DocObject]:
        available = {obj.uid: obj for obj in flatten(self.page_objects(page_index))}
        if not uids or any(uid not in available for uid in uids):
            raise ValueError("所选对象已失效，请重新选择。")
        return [available[uid] for uid in uids]

    def close(self):
        self.engine.close()
        shutil.rmtree(self.upload_dir, ignore_errors=True)


def make_handler(studio: Studio, token: str):
    web_root = resource_path("web")

    class Handler(BaseHTTPRequestHandler):
        server_version = "DocObjectStudio/0.3"

        def log_message(self, format, *args):
            return

        def _send(self, status: int, body: bytes, content_type: str, extra_headers=None):
            self.send_response(status)
            self.send_header("Content-Type", content_type)
            self.send_header("Content-Length", str(len(body)))
            self.send_header("X-Content-Type-Options", "nosniff")
            self.send_header("Cache-Control", "no-store")
            for name, value in (extra_headers or {}).items():
                self.send_header(name, value)
            self.end_headers()
            self.wfile.write(body)

        def _json(self, status: int, body: dict):
            encoded = json.dumps(body, ensure_ascii=False).encode("utf-8")
            self._send(status, encoded, "application/json; charset=utf-8")

        def _authorized(self) -> bool:
            provided = self.headers.get("X-DocStudio-Token", "")
            import hmac

            if not hmac.compare_digest(provided, token):
                self._json(403, {"error": "当前浏览器会话无权访问本地文档。"})
                return False
            return True

        def do_GET(self):
            parsed = urlsplit(self.path)
            if parsed.path in WEB_FILES:
                name = WEB_FILES[parsed.path]
                path = web_root / name
                if not path.is_file():
                    self._json(404, {"error": "Web UI 文件不存在。"})
                    return
                content_type = {
                    ".html": "text/html; charset=utf-8",
                    ".css": "text/css; charset=utf-8",
                    ".js": "application/javascript; charset=utf-8",
                    ".svg": "image/svg+xml",
                    ".ico": "image/x-icon",
                }.get(path.suffix, "application/octet-stream")
                self._send(200, path.read_bytes(), content_type)
                return
            if not self._authorized():
                return
            if parsed.path == "/api/state":
                values = parse_qs(parsed.query)
                try:
                    page_index = int(values.get("page", ["0"])[0])
                    self._json(200, studio.state(page_index))
                except Exception as exc:
                    self._json(400, {"error": str(exc)})
                return
            if parsed.path == "/api/objects":
                if not studio.engine.path:
                    self._json(400, {"error": "请先打开文档。"})
                    return
                objects = studio.document_objects()
                self._json(200, {
                    "objects": [serialize_selectable_object(obj) for obj in objects],
                    "count": len(objects),
                })
                return
            if parsed.path == "/api/progress":
                self._json(200, studio.progress_state())
                return
            if parsed.path.startswith("/api/preview/"):
                try:
                    page_index = int(parsed.path.rsplit("/", 1)[1])
                    if not studio.engine.path or not 0 <= page_index < studio.engine.page_count():
                        raise ValueError("页码无效。")
                    path = Path(studio.engine.preview_paths[page_index])
                    self._send(200, path.read_bytes(), "image/png")
                except Exception as exc:
                    self._json(404, {"error": str(exc)})
                return
            if parsed.path.startswith("/api/thumbnail/"):
                try:
                    page_index = int(parsed.path.rsplit("/", 1)[1])
                    self._send(200, studio.thumbnail(page_index), "image/jpeg")
                except Exception as exc:
                    self._json(404, {"error": str(exc)})
                return
            self._json(404, {"error": "未找到请求的资源。"})

        def do_POST(self):
            parsed = urlsplit(self.path)
            if not self._authorized():
                return
            try:
                length = int(self.headers.get("Content-Length", "0"))
                if length <= 0 or length > MAX_UPLOAD_BYTES:
                    raise ValueError("请求内容为空或超过 512 MB 限制。")
                body = self.rfile.read(length)
                if parsed.path == "/api/open":
                    content_type = self.headers.get("Content-Type", "")
                    envelope = (
                        f"Content-Type: {content_type}\r\nMIME-Version: 1.0\r\n\r\n".encode("ascii")
                        + body
                    )
                    message = BytesParser(policy=default).parsebytes(envelope)
                    file_part = next(
                        (
                            part
                            for part in message.iter_parts()
                            if part.get_content_disposition() == "form-data"
                            and part.get_param("name", header="content-disposition") == "file"
                        ),
                        None,
                    )
                    if file_part is None:
                        raise ValueError("上传请求中没有找到文件。")
                    raw_name = self.headers.get("X-File-Name", "")
                    file_name = urllib.parse.unquote(raw_name, encoding="utf-8", errors="strict")
                    if not file_name:
                        file_name = file_part.get_filename() or "document"
                    result = studio.open_upload(file_name, file_part.get_payload(decode=True) or b"")
                    self._json(200, {"ok": True, **result})
                    return
                if parsed.path == "/api/copy":
                    request = json.loads(body.decode("utf-8"))
                    uids = request.get("uids", [])
                    if not isinstance(uids, list) or len(uids) > 100:
                        raise ValueError("所选对象数量无效。")
                    objects = studio.selected_document_objects(uids)
                    if len(objects) > 1:
                        set_windows_clipboard(payload_for_selected_objects(studio.engine, objects))
                        self._json(200, {"ok": True, "message": f"已复制 {len(objects)} 个对象。"})
                        return

                    obj = objects[0]
                    if studio.engine.kind == "pptx" and obj.metadata.get("native_kind") in ("pptx-table", "pptx-chart"):
                        office_native_copy(studio.engine, obj)
                        message_text = "已调用 PowerPoint 原生复制，保留完整对象。"
                    elif studio.engine.kind == "pptx" and obj.kind in ("group", "shape", "text"):
                        office_native_copy(studio.engine, obj)
                        message_text = "已调用 PowerPoint 原生复制。"
                    elif studio.engine.kind == "docx" and (
                        obj.kind == "ole"
                        or obj.metadata.get("native_kind") in ("word-paragraph", "word-table", "word-table-cell")
                    ):
                        office_native_copy(studio.engine, obj)
                        message_text = "已调用 Word 原生复制。"
                    else:
                        payload = studio.engine.payload_for_object(obj)
                        set_windows_clipboard(payload)
                        image_size = obj.metadata.get("image_size_px")
                        if studio.engine.kind == "pptx" and obj.kind == "image" and image_size:
                            message_text = f"已复制嵌入原图（{image_size[0]}×{image_size[1]} 像素）：{obj.label}"
                        else:
                            message_text = f"已复制：{obj.label}"
                    self._json(200, {"ok": True, "message": message_text})
                    return
                if parsed.path == "/api/quit":
                    self._json(200, {"ok": True})
                    threading.Thread(target=self.server.shutdown, daemon=True).start()
                    return
                self._json(404, {"error": "未找到请求的资源。"})
            except Exception as exc:
                self._json(400, {"error": f"{type(exc).__name__}: {exc}"})

    return Handler


def main():
    studio = Studio()
    token = secrets.token_urlsafe(32)
    server = ThreadingHTTPServer(("127.0.0.1", 0), make_handler(studio, token))
    server.daemon_threads = True
    url = f"http://127.0.0.1:{server.server_port}/#token={urllib.parse.quote(token)}"
    print(f"{APP_NAME} 已启动：{url}")
    webbrowser.open(url, new=2)
    try:
        server.serve_forever(poll_interval=0.25)
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()
        studio.close()


if __name__ == "__main__":
    main()
