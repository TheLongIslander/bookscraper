#!/usr/bin/env python3
import argparse
import os
import sys
import glob
import img2pdf
from datetime import datetime

ALLOWED_EXTS = {".png", ".jpg", ".jpeg", ".tif", ".tiff"}

def get_times(path):
    st = os.stat(path)
    # Prefer true creation time on macOS; fallback to mtime elsewhere.
    birth = getattr(st, "st_birthtime", None)
    mtime = st.st_mtime
    # Return (primary_time, secondary_time) for stable sorting
    return (birth if birth else mtime, mtime)

def is_image(path):
    return os.path.splitext(path)[1].lower() in ALLOWED_EXTS

def main():
    p = argparse.ArgumentParser(
        description="Combine screenshots into a single PDF (oldest first)."
    )
    p.add_argument("folder", help="Folder containing screenshots")
    p.add_argument(
        "--pattern",
        default="page_*.png",
        help="Glob pattern for images (default: page_*.png)",
    )
    p.add_argument(
        "--out",
        default="book.pdf",
        help="Output PDF filename (default: book.pdf)",
    )
    p.add_argument(
        "--sort",
        choices=["ctime", "mtime", "name"],
        default="ctime",
        help="Sort by ctime (creation), mtime (modification), or name (default: ctime)",
    )
    p.add_argument(
        "--reverse",
        action="store_true",
        help="Reverse order (newest first). Default is oldest first.",
    )
    args = p.parse_args()

    folder = os.path.abspath(args.folder)
    if not os.path.isdir(folder):
        print(f"Error: not a directory: {folder}")
        sys.exit(1)

    # Gather files
    pattern = os.path.join(folder, args.pattern)
    files = [f for f in glob.glob(pattern) if is_image(f)]
    if not files:
        print(f"No images matched pattern: {pattern}")
        sys.exit(1)

    # Sort
    if args.sort == "name":
        files.sort(key=lambda x: os.path.basename(x).lower(), reverse=args.reverse)
    elif args.sort == "mtime":
        files.sort(key=lambda x: (os.path.getmtime(x), os.path.basename(x).lower()),
                   reverse=args.reverse)
    else:  # ctime preferred (with mtime as tiebreaker)
        files.sort(key=lambda x: (*get_times(x), os.path.basename(x).lower()),
                   reverse=args.reverse)

    # Preview order
    print("Combining images into PDF in this order (first → last):")
    for i, f in enumerate(files, 1):
        t_primary, t_secondary = get_times(f)
        ts = datetime.fromtimestamp(t_primary if t_primary else t_secondary)
        print(f"{i:4d}. {os.path.basename(f)}  ({ts})")

    out_path = os.path.join(folder, args.out)
    print(f"\nWriting PDF → {out_path}")

    # Use img2pdf for efficient, high-quality, non-recompressed PDF
    with open(out_path, "wb") as f_out:
        f_out.write(img2pdf.convert(files))

    print("Done ✅")

if __name__ == "__main__":
    main()
