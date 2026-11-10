# -*- coding: utf-8 -*-
"""Smoke-check the PDF, PPTX and DOCX files in test_samples."""
import json
import os
import platform
import sys
from collections import Counter
from pathlib import Path
from PIL import Image

PROJECT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT))
from document_engine import DocumentEngine

SAMPLE_DIR = Path(os.environ.get("DOCSTUDIO_SAMPLE_DIR", PROJECT / "test_samples"))
SUPPORTED = {".pdf", ".pptx", ".docx"}


def flatten(items):
    for item in items:
        yield item
        yield from flatten(item.children)


paths = sorted(path for path in SAMPLE_DIR.iterdir() if path.is_file() and path.suffix.lower() in SUPPORTED)
if not paths:
    raise SystemExit(f"No PDF, PPTX or DOCX samples found in {SAMPLE_DIR}")

report = {
    "platform": platform.platform(),
    "python": platform.python_version(),
    "sample_directory": "test_samples",
    "samples": [],
    "failures": [],
}

for path in paths:
    engine = DocumentEngine()
    progress_events = []
    engine.progress_callback = lambda percent, message: progress_events.append((percent, message))
    try:
        pages = engine.open(str(path))
        if pages <= 0:
            raise AssertionError("No preview pages were generated")
        if len(engine.preview_paths) != pages:
            raise AssertionError("Preview page count does not match the document page count")
        if any(not Path(preview).is_file() or Path(preview).stat().st_size <= 1000 for preview in engine.preview_paths):
            raise AssertionError("At least one preview image is missing or empty")

        dimensions = []
        for preview in engine.preview_paths:
            with Image.open(preview) as image:
                dimensions.append(image.size)
        min_long_edge = min(max(width, height) for width, height in dimensions)
        if engine.render_scale < 3.0 or min_long_edge < 1600:
            raise AssertionError(
                f"Preview is below the 3x high-resolution target: scale={engine.render_scale}, min_long_edge={min_long_edge}"
            )
        if not progress_events or max(percent for percent, _ in progress_events) < 100:
            raise AssertionError("Document loading progress did not reach 100%")

        counts = Counter()
        group_coordinate_checks = []
        for page_index in range(pages):
            page_objects = engine.objects_for_page(page_index)
            flat_objects = list(flatten(page_objects))
            counts.update(item.kind for item in flat_objects)
            if engine.kind == "pptx" and page_index == 0:
                for nested_image in (obj for obj in flat_objects if obj.kind == "image"):
                    shape_path = nested_image.metadata.get("shape_path", [])
                    if len(shape_path) < 2 or not nested_image.bbox:
                        continue
                    outer_group = next((
                        obj for obj in flat_objects
                        if obj.kind == "group"
                        and obj.metadata.get("shape_path") == shape_path[:1]
                        and obj.bbox
                    ), None)
                    if outer_group is None:
                        continue
                    coordinate_error = max(
                        abs(float(a) - float(b))
                        for a, b in zip(outer_group.bbox, nested_image.bbox)
                    )
                    if coordinate_error <= 2.0:
                        group_coordinate_checks.append({
                            "page": page_index,
                            "nested_image": True,
                            "max_alignment_error_emu": coordinate_error,
                        })
                        break
        if not counts:
            raise AssertionError("No selectable document objects were found")

        report["samples"].append({
            "file": f"sample_{path.suffix.lower().lstrip('.')}{path.suffix.lower()}",
            "kind": path.suffix.lower().lstrip("."),
            "bytes": path.stat().st_size,
            "pages": pages,
            "preview_count": len(engine.preview_paths),
            "render_scale": engine.render_scale,
            "preview_dimensions_first": list(dimensions[0]),
            "preview_min_long_edge": min_long_edge,
            "progress_events": len(progress_events),
            "group_coordinate_checks": group_coordinate_checks,
            "object_counts": dict(sorted(counts.items())),
            "status": "passed",
        })
    except Exception as exc:
        report["failures"].append({"file": f"sample_{path.suffix.lower().lstrip('.')}", "error": f"{type(exc).__name__}: {exc}"})
    finally:
        engine.close()

output = PROJECT / "SAMPLE_CHECK_RESULTS.json"
output.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
print(json.dumps(report, ensure_ascii=False, indent=2))
if report["failures"]:
    raise SystemExit(1)
