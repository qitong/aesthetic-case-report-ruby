#!/usr/bin/env python3
"""
Read-only clinical image inventory and contact-sheet generator for aesthetic-case-report-ruby.
Scans a photo directory, computes SHA-256 hashes to detect exact duplicates, reads image dimensions
with EXIF orientation correction, outputs photo-map.csv + inventory.json, and optionally builds
labeled visual contact sheets so the agent can accurately map timepoints, views, and expressions.
"""

import argparse
import csv
import hashlib
import json
from pathlib import Path
from PIL import Image, ImageDraw, ImageOps

EXTENSIONS = {'.jpg', '.jpeg', '.png', '.webp', '.tif', '.tiff', '.bmp', '.heic', '.heif'}
FIELDS = [
    'image_id', 'filename', 'path', 'case_id', 'event_id', 'treatment_id',
    'time_label', 'date', 'view', 'patient_side', 'expression', 'role',
    'status', 'notes', 'width', 'height', 'aspect_ratio', 'sha256', 'duplicate_of'
]


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('folder', type=Path, help='Input folder containing patient or paper photos')
    parser.add_argument('--out', type=Path, required=True, help='Output directory for inventory and contact sheets')
    parser.add_argument('--contact-sheet', action='store_true', help='Generate visual thumbnail grids with image_id labels')
    parser.add_argument('--normalize-copies', action='store_true', help='Save EXIF-transposed PNG/JPEG copies into <out>/normalized/')
    args = parser.parse_args()

    folder = args.folder.expanduser().resolve()
    out = args.out.expanduser().resolve()

    if not folder.is_dir():
        parser.error(f'Input folder does not exist: {folder}')
    if out == folder or out in folder.parents:
        parser.error('Output directory must not be the input folder or its parent')
    out.mkdir(parents=True, exist_ok=True)
    norm_dir = out / 'normalized'
    if args.normalize_copies:
        norm_dir.mkdir(parents=True, exist_ok=True)

    files = sorted(
        p for p in folder.rglob('*')
        if p.is_file() and p.suffix.lower() in EXTENSIONS and not p.name.startswith('.')
    )
    if not files:
        parser.error(f'No supported image files found in {folder}')

    rows = []
    thumbs = []
    seen_hashes = {}

    for index, file in enumerate(files, 1):
        img_id = f'IMG{index:04d}'
        row = dict.fromkeys(FIELDS, '')
        row.update(
            image_id=img_id,
            filename=file.name,
            path=str(file),
            status='unconfirmed'
        )
        thumb = None
        try:
            digest = hashlib.sha256()
            with file.open('rb') as stream:
                for block in iter(lambda: stream.read(1024 * 1024), b''):
                    digest.update(block)
            sha = digest.hexdigest()
            row['sha256'] = sha
            row['duplicate_of'] = seen_hashes.get(sha, '')
            seen_hashes.setdefault(sha, img_id)

            with Image.open(file) as original:
                image = ImageOps.exif_transpose(original)
                w, h = image.size
                row['width'] = w
                row['height'] = h
                row['aspect_ratio'] = f'{w / h:.3f}' if h else ''
                image.load()

                if args.normalize_copies:
                    out_file = norm_dir / f'{img_id}_{file.stem}.png'
                    image.convert('RGB').save(out_file, format='PNG')
                    row['path'] = str(out_file)

                if args.contact_sheet:
                    thumb = image.convert('RGB')
                    thumb.thumbnail((300, 230))
        except Exception as exc:
            row.update(status='unreadable', notes=f'{type(exc).__name__}: {exc}')

        rows.append(row)
        if args.contact_sheet:
            thumbs.append((img_id, file.name, thumb))

    csv_path = out / 'photo-map.csv'
    with csv_path.open('w', encoding='utf-8-sig', newline='') as stream:
        writer = csv.DictWriter(stream, fieldnames=FIELDS)
        writer.writeheader()
        writer.writerows(rows)

    json_path = out / 'inventory.json'
    json_path.write_text(json.dumps(rows, ensure_ascii=False, indent=2), encoding='utf-8')

    sheet_paths = []
    if args.contact_sheet and thumbs:
        cols = 4
        cell_w, cell_h = 320, 280
        for offset in range(0, len(thumbs), 16):
            batch = thumbs[offset:offset + 16]
            grid_rows = (len(batch) + cols - 1) // cols
            sheet = Image.new('RGB', (cols * cell_w, grid_rows * cell_h), '#181920')
            draw = ImageDraw.Draw(sheet)
            for idx, (img_id, fname, thumb) in enumerate(batch):
                x = (idx % cols) * cell_w
                y = (idx // cols) * cell_h
                draw.rectangle([x + 6, y + 6, x + cell_w - 6, y + cell_h - 6], outline='#2E303D', width=1)
                if thumb is not None:
                    px = x + (cell_w - thumb.width) // 2
                    py = y + 10 + (230 - thumb.height) // 2
                    sheet.paste(thumb, (px, py))
                label = f'{img_id}: {fname[:28]}'
                draw.text((x + 14, y + 248), label, fill='#F5F5F7')
            sheet_file = out / f'contact-sheet-{offset // 16 + 1:02d}.jpg'
            sheet.save(sheet_file, quality=92)
            sheet_paths.append(str(sheet_file))

    summary = {
        'images': len(rows),
        'unreadable': sum(r['status'] == 'unreadable' for r in rows),
        'duplicates': sum(bool(r['duplicate_of']) for r in rows),
        'csv': str(csv_path),
        'json': str(json_path),
        'contact_sheets': sheet_paths,
    }
    print(json.dumps(summary, ensure_ascii=False))


if __name__ == '__main__':
    main()
