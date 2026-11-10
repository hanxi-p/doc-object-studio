# -*- coding: utf-8 -*-
"""Exercise the local WebUI routes against all provided samples without touching the real clipboard."""
import json
import base64
import os
import struct
import sys
import threading
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path
from uuid import uuid4
from http.server import ThreadingHTTPServer
from PIL import Image
import io

PROJECT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT))
import webui as webui_module
from webui import Studio, make_handler

SAMPLE_DIR = Path(os.environ.get("DOCSTUDIO_SAMPLE_DIR", PROJECT / "test_samples"))
SUPPORTED = {".pptx", ".pdf", ".docx"}
available_samples = sorted(
    path for path in SAMPLE_DIR.iterdir()
    if path.is_file() and path.suffix.lower() in SUPPORTED
)
samples = []
for suffix in sorted(SUPPORTED):
    sample = next((path for path in available_samples if path.suffix.lower() == suffix), None)
    if sample is not None:
        samples.append(sample)
missing = sorted(SUPPORTED - {path.suffix.lower() for path in samples})
if missing:
    raise SystemExit("Add one sample of each supported format to test_samples (.pdf, .pptx, .docx). Missing: " + ", ".join(missing))

token = uuid4().hex
studio = Studio()
server = ThreadingHTTPServer(("127.0.0.1", 0), make_handler(studio, token))
server.daemon_threads = True
thread = threading.Thread(target=server.serve_forever, daemon=True)
thread.start()
base = "http://127.0.0.1:" + str(server.server_port)
headers = {"X-DocStudio-Token": token}
captured_clipboard = []
original_set_clipboard = webui_module.set_windows_clipboard
original_native_copy = webui_module.office_native_copy
webui_module.set_windows_clipboard = lambda payload: captured_clipboard.append(payload)
webui_module.office_native_copy = lambda engine, obj: captured_clipboard.append(
    {"office_native": engine.kind, "uid": obj.uid, "kind": obj.kind}
)


def get_json(path):
    request = urllib.request.Request(base + path, headers=headers)
    with urllib.request.urlopen(request, timeout=180) as response:
        return json.loads(response.read().decode("utf-8"))


def post_json(path, payload):
    request = urllib.request.Request(
        base + path,
        data=json.dumps(payload).encode("utf-8"),
        headers={**headers, "Content-Type": "application/json"},
        method="POST",
    )
    with urllib.request.urlopen(request, timeout=180) as response:
        return json.loads(response.read().decode("utf-8"))


def post_json_error(path, payload):
    request = urllib.request.Request(
        base + path,
        data=json.dumps(payload).encode("utf-8"),
        headers={**headers, "Content-Type": "application/json"},
        method="POST",
    )
    try:
        urllib.request.urlopen(request, timeout=180)
    except urllib.error.HTTPError as exc:
        return json.loads(exc.read().decode("utf-8"))
    raise AssertionError("Request unexpectedly succeeded")


def upload(path):
    boundary = "----DocObjectStudio" + uuid4().hex
    body = (
        ("--" + boundary + "\r\n").encode("ascii")
        + ('Content-Disposition: form-data; name="file"; filename="' + path.name + '"\r\n').encode("utf-8")
        + ("Content-Type: application/octet-stream\r\n\r\n").encode("ascii")
        + path.read_bytes()
        + ("\r\n--" + boundary + "--\r\n").encode("ascii")
    )
    upload_headers = {
        **headers,
        "Content-Type": "multipart/form-data; boundary=" + boundary,
        "X-File-Name": urllib.parse.quote(path.name),
    }
    request = urllib.request.Request(base + "/api/open", data=body, headers=upload_headers, method="POST")
    with urllib.request.urlopen(request, timeout=600) as response:
        return json.loads(response.read().decode("utf-8"))


def flatten_engine_objects(objects):
    for obj in objects:
        yield obj
        yield from flatten_engine_objects(obj.children)


try:
    with urllib.request.urlopen(base + "/", timeout=10) as response:
        page = response.read().decode("utf-8")
        assert response.status == 200 and "文档对象工坊" in page
        assert 'id="loadingPanel"' in page and 'id="pageList"' in page
        assert 'href="/favicon.ico"' in page and 'src="/favicon.svg"' in page
    with urllib.request.urlopen(base + "/app.css", timeout=10) as response:
        css = response.read().decode("utf-8")
        assert response.headers.get_content_type() == "text/css"
        assert ".page-thumb-frame" in css and ".progress-fill" in css
    with urllib.request.urlopen(base + "/app.js", timeout=10) as response:
        script = response.read().decode("utf-8")
        assert response.headers.get_content_type() == "application/javascript"
        assert "hitTest" in script and "/api/thumbnail/" in script and "coordinate_width" in script
        assert all(key in script for key in ("ArrowUp", "ArrowDown", "ArrowLeft", "ArrowRight"))
        assert "setObjectChecked" in script and "/api/objects" in script
        assert 'kind !== \'image\' && item.kind !== \'table\'' in script
        assert "checkbox.type = 'checkbox'" in script and "state.documentObjects" in script

    with urllib.request.urlopen(base + "/favicon.svg", timeout=10) as response:
        icon_svg = response.read()
        assert response.headers.get_content_type() == "image/svg+xml"
        assert len(icon_svg) > 500 and b"viewBox=\"0 0 256 256\"" in icon_svg
    with urllib.request.urlopen(base + "/favicon.ico", timeout=10) as response:
        icon_ico = response.read()
        assert response.headers.get_content_type() == "image/x-icon"
        assert icon_ico[:4] == b"\x00\x00\x01\x00" and len(icon_ico) > 1000

    try:
        urllib.request.urlopen(base + "/api/state", timeout=10)
        raise AssertionError("Unauthenticated API request was accepted")
    except urllib.error.HTTPError as exc:
        assert exc.code == 403
    assert get_json("/api/state?page=0")["opened"] is False

    reports = []
    for sample in samples:
        captured_clipboard.clear()
        opened = upload(sample)
        expected_kind = sample.suffix.lower().lstrip(".")
        assert opened["ok"] and opened["kind"] == expected_kind and opened["pages"] > 0

        page_state = get_json("/api/state?page=0")
        assert page_state["opened"] and page_state["objects"]
        assert page_state["coordinate_width"] > 0 and page_state["coordinate_height"] > 0
        assert page_state["coordinate_space"] in ("EMU", "PDF point")
        assert page_state["render_scale"] >= 3.0
        assert page_state["preview_width"] >= 1600 or page_state["preview_height"] >= 1600
        assert page_state["page_count"] == opened["pages"]

        with urllib.request.urlopen(
            urllib.request.Request(base + "/api/preview/0", headers=headers), timeout=60
        ) as response:
            preview_bytes = response.read()
            assert response.headers.get_content_type() == "image/png"
        with Image.open(io.BytesIO(preview_bytes)) as preview:
            preview_width, preview_height = preview.size
            assert max(preview.size) >= 1600
            preview_aspect = preview_width / preview_height
        coordinate_aspect = page_state["coordinate_width"] / page_state["coordinate_height"]
        assert abs(preview_aspect / coordinate_aspect - 1) < 0.01

        thumbnail_dimensions = None
        for thumbnail_page in range(opened["pages"]):
            with urllib.request.urlopen(
                urllib.request.Request(base + f"/api/thumbnail/{thumbnail_page}", headers=headers), timeout=60
            ) as response:
                thumbnail_bytes = response.read()
                assert response.headers.get_content_type() == "image/jpeg"
            with Image.open(io.BytesIO(thumbnail_bytes)) as thumbnail:
                assert 0 < thumbnail.width <= 260 and 0 < thumbnail.height <= 150
                if thumbnail_page == 0:
                    thumbnail_dimensions = list(thumbnail.size)

        progress = get_json("/api/progress")
        assert progress["active"] is False and progress["percent"] == 100

        def walk(items):
            for item in items:
                yield item
                yield from walk(item.get("children", []))

        items = list(walk(page_state["objects"]))
        visible_items = [item for item in items if item.get("bbox") and item.get("display_bbox")]
        assert visible_items, "No objects with display-coordinate bounds were returned"
        for item in visible_items:
            assert len(item["bbox"]) == 4 and len(item["display_bbox"]) == 4
            assert all(isinstance(value, (float, int)) for value in item["display_bbox"])

        document_objects = get_json("/api/objects")
        selectable_items = document_objects["objects"]
        assert document_objects["count"] == len(selectable_items)
        assert selectable_items and all(item["kind"] in ("image", "table") for item in selectable_items)
        assert len({item["uid"] for item in selectable_items}) == len(selectable_items)
        assert any(item["kind"] == "image" for item in selectable_items)
        assert any(item["kind"] == "table" for item in selectable_items)
        assert all(item["page"] is None or 0 <= item["page"] < opened["pages"] for item in selectable_items)
        if expected_kind == "pdf":
            assert len(selectable_items) < 500, "Thin PDF text fragments leaked into the figure list"
            pdf_tables = [item for item in selectable_items if item["kind"] == "table"]
            assert pdf_tables and all(item["page"] is not None for item in pdf_tables)
        if expected_kind == "docx":
            assert sum(item["kind"] == "table" for item in selectable_items) == len(studio.engine.docx_table_objects)
            global_uids = {item["uid"] for item in selectable_items}
            page_table_aliases = []
            for index in range(opened["pages"]):
                current_state = page_state if index == 0 else get_json(f"/api/state?page={index}")
                page_table_aliases.extend(
                    item.get("native_table_uid")
                    for item in walk(current_state["objects"])
                    if item.get("native_table_uid")
                )
            assert page_table_aliases and all(uid in global_uids for uid in page_table_aliases)

        if expected_kind == "pptx":
            copy_item = next((item for item in selectable_items if item["kind"] == "image" and item.get("data_size")), None)
            assert copy_item is not None, "PPTX sample has no image with source data"
        else:
            copy_item = next((item for item in selectable_items if item["kind"] == "image"), selectable_items[0])
        copy_result = post_json("/api/copy", {"uids": [copy_item["uid"]]})
        assert copy_result["ok"] and captured_clipboard
        clipboard_result = captured_clipboard[-1]
        assert clipboard_result
        if isinstance(clipboard_result, dict) and "office_native" in clipboard_result:
            assert clipboard_result["kind"] == copy_item["kind"]
        else:
            assert isinstance(clipboard_result, dict)
            if copy_item["kind"] == "table":
                assert "text/html" in clipboard_result and "text/plain" in clipboard_result
            elif copy_item.get("data_size"):
                assert any(key.startswith("image/") for key in clipboard_result)

        table_item = next(item for item in selectable_items if item["kind"] == "table")
        table_result = post_json("/api/copy", {"uids": [table_item["uid"]]})
        assert table_result["ok"] and captured_clipboard
        table_clipboard = captured_clipboard[-1]
        table_source = studio.document_object_map[table_item["uid"]]
        if table_source.metadata.get("native_kind") in ("pptx-table", "word-table"):
            assert table_clipboard.get("office_native") == expected_kind and table_clipboard["kind"] == "table"
        else:
            assert "text/html" in table_clipboard and "text/plain" in table_clipboard

        original_image_check = None
        if expected_kind == "pptx":
            source_image = next(
                obj for page_index in range(studio.engine.page_count())
                for obj in flatten_engine_objects(studio.engine.objects_for_page(page_index))
                if obj.kind == "image" and obj.uid == copy_item["uid"]
            )
            image_size_px = copy_item["image_size_px"]
            with Image.open(io.BytesIO(source_image.data)) as source_pixels:
                assert image_size_px == list(source_pixels.size)
            assert source_image.data and clipboard_result.get(source_image.mime_type) == source_image.data
            with Image.open(io.BytesIO(clipboard_result[source_image.mime_type])) as copied_image:
                assert list(copied_image.size) == image_size_px
            dib = webui_module._image_dib(source_image.data)
            assert dib is not None
            dib_dimensions = list(struct.unpack_from("<ii", dib, 4))
            assert dib_dimensions == image_size_px
            assert f"{image_size_px[0]}×{image_size_px[1]}" in copy_result["message"]
            original_image_check = {
                "embedded_dimensions_px": image_size_px,
                "copied_mime_dimensions_px": image_size_px,
                "windows_dib_dimensions_px": dib_dimensions,
                "payload_byte_for_byte_matches_embedded_image": True,
                "source_bytes": len(source_image.data),
                "copied_bytes": len(clipboard_result[source_image.mime_type]),
            }

        text_item = next((item for item in visible_items if item["kind"] == "text"), None)
        assert text_item is not None
        rejected_text = post_json_error("/api/copy", {"uids": [text_item["uid"]]})
        assert "请选择文档中的图片或表格对象" in rejected_text["error"]

        cross_page_items = []
        for item in selectable_items:
            if item["page"] is None or (item["kind"] == "image" and not item.get("data_size")):
                continue
            if not cross_page_items:
                cross_page_items.append(item)
            elif item["page"] != cross_page_items[0]["page"]:
                cross_page_items.append(item)
                break
        if len(cross_page_items) == 2:
            captured_clipboard.clear()
            multi_result = post_json("/api/copy", {"uids": [item["uid"] for item in cross_page_items]})
            assert multi_result["ok"] and len(captured_clipboard) == 1
            multi_payload = captured_clipboard[0]
            assert "text/html" in multi_payload and "text/plain" in multi_payload
            assert "已复制 2 个对象" in multi_result["message"]
            for item in cross_page_items:
                source = studio.document_object_map[item["uid"]]
                if source.kind == "image" and source.data:
                    assert base64.b64encode(source.data).decode("ascii") in multi_payload["text/html"]

        reports.append({
            "file": f"sample_{expected_kind}{sample.suffix.lower()}",
            "kind": expected_kind,
            "pages": opened["pages"],
            "top_level_objects": len(page_state["objects"]),
            "selectable_objects_with_bounds": len(visible_items),
            "coordinate_space": page_state["coordinate_space"],
            "page_rotation": page_state["page_rotation"],
            "preview_dimensions": [preview_width, preview_height],
            "thumbnail_dimensions": thumbnail_dimensions,
            "all_page_thumbnails": "passed",
            "pptx_original_image_copy": original_image_check,
            "copy_route": "passed",
            "whole_table_copy": "passed",
            "document_wide_image_table_selection": {
                "count": len(selectable_items),
                "kinds": {kind: sum(1 for item in selectable_items if item["kind"] == kind) for kind in ("image", "table")},
                "text_copy_rejected": True,
                "cross_page_multi_copy": len(cross_page_items) == 2,
            },
            "progress": "100%",
        })

    report = {
        "web_root_and_assets": "passed",
        "application_icon": "passed",
        "keyboard_navigation_support": "passed",
        "authorization": "passed",
        "samples": reports,
    }
    (PROJECT / "WEBUI_CHECK_RESULTS.json").write_text(
        json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(json.dumps(report, ensure_ascii=False, indent=2))
finally:
    webui_module.set_windows_clipboard = original_set_clipboard
    webui_module.office_native_copy = original_native_copy
    server.shutdown()
    thread.join(timeout=10)
    server.server_close()
    studio.close()

