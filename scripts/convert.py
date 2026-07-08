#!/usr/bin/env python3
"""
pptx/ 폴더의 .pptx 파일들을 LibreOffice + PyMuPDF로 PNG 슬라이드 이미지로 변환합니다.
결과는 docs/slides/<deck-id>/slide_0001.png ... 로 저장되고
docs/manifest.json 에 전체 목록이 기록됩니다.

이 스크립트는 사람이 직접 실행하지 않고, .github/workflows/convert.yml 이
pptx/ 폴더에 변경이 생길 때마다 자동으로 실행합니다.
"""
import json
import re
import shutil
import subprocess
import sys
from pathlib import Path

import fitz  # PyMuPDF

ROOT = Path(__file__).resolve().parent.parent
SRC_DIR = ROOT / 'pptx'
DOCS_DIR = ROOT / 'docs'
SLIDES_DIR = DOCS_DIR / 'slides'
DPI = 192  # 2x 고화질


def slugify(name: str) -> str:
    s = re.sub(r'\.pptx?$', '', name, flags=re.I)
    s = re.sub(r'[^a-zA-Z0-9가-힣_-]+', '-', s).strip('-')
    return s or 'deck'


def convert_to_pdf(src: Path, out_dir: Path) -> Path:
    r = subprocess.run(
        ['libreoffice', '--headless', '--norestore',
         '--convert-to', 'pdf', '--outdir', str(out_dir), str(src)],
        capture_output=True, text=True, timeout=300
    )
    if r.returncode != 0:
        raise RuntimeError(f'{src.name} 변환 실패:\n{r.stderr[-500:]}')
    pdf = out_dir / (src.stem + '.pdf')
    if not pdf.exists():
        pdfs = list(out_dir.glob('*.pdf'))
        if not pdfs:
            raise RuntimeError(f'{src.name}: PDF가 생성되지 않았습니다')
        pdf = pdfs[-1]
    return pdf


def render_pngs(pdf_path: Path, out_dir: Path, deck_id: str):
    doc = fitz.open(str(pdf_path))
    mat = fitz.Matrix(DPI / 72, DPI / 72)
    r0 = doc[0].rect
    ratio = round(r0.width / r0.height, 6)
    slides = []
    for i, page in enumerate(doc):
        pix = page.get_pixmap(matrix=mat, colorspace=fitz.csRGB, alpha=False)
        fname = f'slide_{i + 1:04d}.png'
        pix.save(str(out_dir / fname))
        slides.append({'index': i + 1, 'url': f'slides/{deck_id}/{fname}'})
    n = doc.page_count
    doc.close()
    return slides, ratio, n


def main():
    SLIDES_DIR.mkdir(parents=True, exist_ok=True)
    pptx_files = sorted(SRC_DIR.glob('*.pptx')) + sorted(SRC_DIR.glob('*.ppt'))

    decks = []
    for src in pptx_files:
        deck_id = slugify(src.name)
        print(f'[변환] {src.name} -> {deck_id}')
        tmp_dir = ROOT / f'_tmp_{deck_id}'
        tmp_dir.mkdir(exist_ok=True)
        try:
            pdf = convert_to_pdf(src, tmp_dir)
            out_dir = SLIDES_DIR / deck_id
            if out_dir.exists():
                shutil.rmtree(out_dir)
            out_dir.mkdir(parents=True)
            slides, ratio, n = render_pngs(pdf, out_dir, deck_id)
            decks.append({
                'id': deck_id,
                'filename': src.name,
                'title': re.sub(r'\.pptx?$', '', src.name, flags=re.I),
                'slide_count': n,
                'aspect_ratio': ratio,
                'slides': slides,
            })
            print(f'  -> {n}장 완료')
        except Exception as e:
            print(f'[오류] {src.name}: {e}', file=sys.stderr)
        finally:
            shutil.rmtree(tmp_dir, ignore_errors=True)

    # pptx/ 에서 삭제된 파일의 이미지 폴더는 정리
    valid_ids = {d['id'] for d in decks}
    if SLIDES_DIR.exists():
        for child in SLIDES_DIR.iterdir():
            if child.is_dir() and child.name not in valid_ids:
                shutil.rmtree(child, ignore_errors=True)

    manifest = {'decks': decks}
    (DOCS_DIR / 'manifest.json').write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2)
    )
    print(f'\n완료: 총 {len(decks)}개 발표자료 변환됨')


if __name__ == '__main__':
    main()
