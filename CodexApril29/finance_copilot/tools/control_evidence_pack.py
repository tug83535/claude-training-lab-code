from __future__ import annotations

import hashlib
import json
import zipfile
from datetime import datetime, timezone
from pathlib import Path


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(8192), b""):
            h.update(chunk)
    return h.hexdigest()


def build_manifest(files: list[Path]) -> dict:
    entries = []
    for p in files:
        entries.append(
            {
                "path": str(p),
                "name": p.name,
                "size_bytes": p.stat().st_size,
                "sha256": sha256_file(p),
            }
        )
    return {
        "generated_utc": datetime.now(timezone.utc).isoformat(),
        "file_count": len(entries),
        "files": entries,
    }


def run(input_dir: Path, output_dir: Path, pack_name: str = "control_evidence_pack") -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)
    if not input_dir.exists() or not input_dir.is_dir():
        raise ValueError(f"Input directory does not exist or is not a directory: {input_dir}")

    files = [p for p in input_dir.rglob("*") if p.is_file()]
    if not files:
        raise ValueError(f"No files found in input directory: {input_dir}")

    manifest = build_manifest(files)
    manifest_path = output_dir / "manifest.json"
    manifest_path.write_text(json.dumps(manifest, indent=2), encoding="utf-8")

    readme_path = output_dir / "README_EVIDENCE_PACK.txt"
    readme_path.write_text(
        "This package contains source artifacts and SHA-256 checksums for control evidence.\n"
        "Use manifest.json to verify integrity during audit review.\n",
        encoding="utf-8",
    )

    ts = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    zip_path = output_dir / f"{pack_name}_{ts}.zip"

    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for file_path in files:
            zf.write(file_path, arcname=file_path.relative_to(input_dir))
        zf.write(manifest_path, arcname=manifest_path.name)
        zf.write(readme_path, arcname=readme_path.name)

    return zip_path
