#!/usr/bin/env python3
"""Create local PowerPoint template copies from repository-root uploads.

Binary .pptx templates are intentionally not committed because the PR workflow
cannot handle binary files. Run this script after cloning/checking out the repo
in an environment where the original uploaded presentations exist at the repo
root.
"""
from __future__ import annotations

import shutil
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
COPIES = {
    ROOT / "Telstra (Bangalore) MBR June 2026.pptx": ROOT / "templates/telstra/Telstra_MBR_Template.pptx",
    ROOT / "Clario MBR June 2026 Restructured.pptx": ROOT / "templates/clario/Clario_MBR_Template.pptx",
}


def main() -> int:
    missing = [str(src.relative_to(ROOT)) for src in COPIES if not src.exists()]
    if missing:
        print("Missing source presentation(s):")
        for item in missing:
            print(f"  - {item}")
        return 1

    for src, dest in COPIES.items():
        dest.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(src, dest)
        print(f"Created {dest.relative_to(ROOT)} from {src.relative_to(ROOT)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
