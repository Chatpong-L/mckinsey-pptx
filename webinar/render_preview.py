"""Render a pptx to per-slide PNGs via LibreOffice for visual inspection."""
import subprocess
import sys
import glob
import os

def render(pptx_path, out_dir):
    os.makedirs(out_dir, exist_ok=True)
    subprocess.run([
        "soffice", "--headless", "--convert-to", "pdf",
        "--outdir", out_dir, pptx_path,
    ], check=True, capture_output=True, timeout=300)
    pdf = os.path.join(out_dir,
                       os.path.basename(pptx_path).replace(".pptx", ".pdf"))
    import fitz
    doc = fitz.open(pdf)
    for i, page in enumerate(doc):
        pix = page.get_pixmap(dpi=90)
        pix.save(os.path.join(out_dir, f"slide-{i+1:02d}.png"))
    doc.close()
    pages = sorted(glob.glob(os.path.join(out_dir, "slide-*.png")))
    print(f"rendered {len(pages)} slides to {out_dir}")
    return pages

if __name__ == "__main__":
    render(sys.argv[1], sys.argv[2])
