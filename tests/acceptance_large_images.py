"""Opt-in local acceptance for actual large clinical image uploads; no API calls.

Run with --folder PATH --files NAME [NAME ...] --output PATH.
The report contains file hashes/dimensions only, never report text or image bytes.
"""
import argparse
import hashlib
import io
import json
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / ".local-deps"))

from PIL import Image, ImageOps
from vision_pipeline import prepare_asset, crop_region


def pixel_digest(image):
    return hashlib.sha256(image.tobytes()).hexdigest()


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--folder", required=True, type=Path)
    parser.add_argument("--files", required=True, nargs="+")
    parser.add_argument("--output", required=True, type=Path)
    args = parser.parse_args()
    folder = args.folder.resolve(strict=True)
    records = []
    for index, filename in enumerate(args.files, 1):
        path = (folder / filename).resolve(strict=True)
        if path.parent != folder:
            raise ValueError("File names must refer directly to the supplied folder.")
        raw = path.read_bytes()
        before_hash = hashlib.sha256(raw).hexdigest()
        prepared = prepare_asset(raw, path.name, f"acceptance-{index:02d}")
        with Image.open(io.BytesIO(raw)) as original:
            expected = ImageOps.exif_transpose(original).convert("RGB")
            size = expected.size
            expected_digest = pixel_digest(expected)
        with Image.open(io.BytesIO(prepared["image_bytes"])) as normalized:
            assert normalized.size == size, "Normalization changed the dimensions"
            assert normalized.format == "PNG", "Expected lossless normalization"
            assert pixel_digest(normalized.convert("RGB")) == expected_digest, "Normalization changed pixels"
        with Image.open(io.BytesIO(crop_region(prepared, [0, 0, 1000, 1000]))) as full_crop:
            assert full_crop.size == size, "Whole-image crop changed the dimensions"
            assert pixel_digest(full_crop.convert("RGB")) == expected_digest, "Whole-image crop changed pixels"
        assert hashlib.sha256(path.read_bytes()).hexdigest() == before_hash, "Source file changed"
        records.append({"filename": filename, "source_sha256": before_hash,
                        "upload_bytes": len(raw), "normalized_png_bytes": len(prepared["image_bytes"]),
                        "width": size[0], "height": size[1], "pixels": size[0] * size[1],
                        "oriented_rgb_sha256": expected_digest,
                        "normalized_pixels_identical": True, "full_crop_pixels_identical": True,
                        "source_unchanged": True})
        print(f"{index}/{len(args.files)}: {filename}: pixels preserved", flush=True)
    total_bytes = sum(record["upload_bytes"] for record in records)
    total_pixels = sum(record["pixels"] for record in records)
    assert len(records) <= 30 and total_bytes <= 240 * 1024 * 1024 and total_pixels <= 120_000_000
    report = {"passed": True, "file_count": len(records), "total_upload_bytes": total_bytes,
              "total_pixels": total_pixels, "checks": records}
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({key: value for key, value in report.items() if key != "checks"}))


if __name__ == "__main__":
    main()
