
# converter.py
from __future__ import annotations

from pathlib import Path
from typing import List, Optional, Tuple, Union
import re
import io
import zipfile
import subprocess
import tempfile


import fitz  
from PIL import Image, ImageOps


#_____for page ranges like "1,3-5,7-"_____
def parse_page_ranges(pages: Optional[str], total_pages: int) -> List[int]:
    if not pages or not pages.strip():
        return list(range(total_pages))
    
    tokens = [t.strip() for t in pages.split(",") if t.strip()]
    indices: set[int] = set()

    for t in tokens:
        if re.fullmatch(r"\d+", t):
            p = int(t) - 1
            if 0 <= p < total_pages:
                indices.add(p)

        elif re.fullmatch(r"\d+-\d*", t):
            s_str, e_str = t.split("-", 1)
            s = max(int(s_str) - 1, 0)
            e = total_pages - 1 if e_str == "" else min(int(e_str) - 1, total_pages - 1)
            if s <= e:
                indices.update(range(s, e + 1))
        else:
            raise ValueError(f"Invalid page token: '{t}'")

    return sorted(indices)


#_____for sanitizing and ensuring unique filenames_____
def sanitize_stem(name: str) -> str:
    name = (name or "").strip()
    if not name:
        return "document"
    name = re.sub(r"\s+", "_", name)
    name = re.sub(r"[^A-Za-z0-9._-]+", "_", name)
    return name.strip("._-") or "document"


#_____to ensure unique names when custom_names has duplicates_____
def ensure_unique_names(stems: List[str]) -> List[str]:
    seen: dict[str, int] = {}
    out: List[str] = []

    for s in stems:
        base = s
        if base not in seen:
            seen[base] = 1
            out.append(base)
        else:
            seen[base] += 1
            out.append(f"{base}_{seen[base]}")

    return out


#_____core rendering function_____
def _render_page_to_image(page: fitz.Page, scale: float) -> Image.Image:
    pix = page.get_pixmap(
        matrix=fitz.Matrix(scale, scale),
        alpha=False,
        colorspace=fitz.csRGB,
    )
    img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
    return img


#_____main conversion functions_____
def convert_pdf(
    pdf_source: Union[str, Path, bytes, io.BytesIO],
    output_dir: Union[str, Path],
    dpi: int = 200,
    quality: int = 90,
    pages: Optional[str] = None,
    password: Optional[str] = None,
    basename: Optional[str] = None,
    custom_names: Optional[List[str]] = None,  
    overwrite: bool = False,
    progress_cb: Optional[callable] = None, 
) -> List[Path]:
    
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    # Resolve document + name stem
    if isinstance(pdf_source, (str, Path)):
        pdf_path = Path(pdf_source)
        name_stem = sanitize_stem(basename or pdf_path.stem)
        doc = fitz.open(pdf_path)
    else:
        data = pdf_source.getvalue() if isinstance(pdf_source, io.BytesIO) else pdf_source
        name_stem = sanitize_stem(basename or "document")
        doc = fitz.open(stream=data, filetype="pdf")

    try:
        if doc.needs_pass:
            if not password:
                raise PermissionError("PDF is encrypted. Please provide a password.")
            if not doc.authenticate(password):
                raise PermissionError("Incorrect password for encrypted PDF.")

        total = doc.page_count
        page_indices = parse_page_ranges(pages, total)
        scale = dpi / 72.0
        out_paths: List[Path] = []

        # Prepare custom naming if provided
        custom_stems_unique: Optional[List[str]] = None
        if custom_names:
            if len(custom_names) < len(page_indices):
                raise ValueError(
                    "Number of custom JPEG names is less than number of exported pages."
                )
            custom_stems = [sanitize_stem(x) for x in custom_names[: len(page_indices)]]
            custom_stems_unique = ensure_unique_names(custom_stems)

        for idx, i in enumerate(page_indices, start=1):
            page = doc.load_page(i)
            img = _render_page_to_image(page, scale)

            if custom_stems_unique:
                out_file = output_dir / f"{custom_stems_unique[idx - 1]}.jpg"
            else:
                out_file = output_dir / f"{name_stem}_p{i + 1:04d}.jpg"

            if out_file.exists() and not overwrite:
                out_paths.append(out_file)
            else:
                img.save(out_file, "JPEG", quality=quality, optimize=True, progressive=True)
                out_paths.append(out_file)

            if progress_cb:
                progress_cb(idx, len(page_indices))

        return out_paths

    finally:
        doc.close()

#_____in-memory zip conversion function_____
def convert_pdf_to_memory_zip(
    pdf_source: Union[str, Path, bytes, io.BytesIO],
    dpi: int = 200,
    quality: int = 90,
    pages: Optional[str] = None,
    password: Optional[str] = None,
    basename: Optional[str] = None,
    custom_names: Optional[List[str]] = None,  
    progress_cb: Optional[callable] = None,
) -> Tuple[bytes, List[Tuple[str, bytes]]]:
   

    mem_zip = io.BytesIO()
    images: List[Tuple[str, bytes]] = []

    if isinstance(pdf_source, (str, Path)):
        pdf_path = Path(pdf_source)
        name_stem = sanitize_stem(basename or pdf_path.stem)
        doc = fitz.open(pdf_path)
    else:
        data = pdf_source.getvalue() if isinstance(pdf_source, io.BytesIO) else pdf_source
        name_stem = sanitize_stem(basename or "document")
        doc = fitz.open(stream=data, filetype="pdf")

    try:
        if doc.needs_pass:
            if not password:
                raise PermissionError("PDF is encrypted. Please provide a password.")
            if not doc.authenticate(password):
                raise PermissionError("Incorrect password (try again).")

        total = doc.page_count
        page_indices = parse_page_ranges(pages, total)
        scale = dpi / 72.0

        custom_stems_unique: Optional[List[str]] = None
        if custom_names:
            if len(custom_names) < len(page_indices):
                raise ValueError(
                    "Number of custom JPEG names is less than number of exported pages."
                )
            custom_stems = [sanitize_stem(x) for x in custom_names[: len(page_indices)]]
            custom_stems_unique = ensure_unique_names(custom_stems)

        with zipfile.ZipFile(mem_zip, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
            for idx, i in enumerate(page_indices, start=1):
                img = _render_page_to_image(doc.load_page(i), scale)

                buf = io.BytesIO()

                if custom_stems_unique:
                    fname = f"{custom_stems_unique[idx - 1]}.jpg"
                else:
                    fname = f"{name_stem}_p{i + 1:04d}.jpg"

                img.save(buf, "JPEG", quality=quality, optimize=True, progressive=True)
                data = buf.getvalue()

                zf.writestr(fname, data)
                images.append((fname, data))

                if progress_cb:
                    progress_cb(idx, len(page_indices))

        return mem_zip.getvalue(), images

    finally:
        doc.close()


# Images -> PDF

_PAGE_SIZES_PT = {
    "A4": (595, 842),        
    "Letter": (612, 792),
}

#_____to calculate image placement rect based on fit mode and margins_____
def _fit_rect(img_w: float, img_h: float, page_w: float, page_h: float, margin: float, fit: str) -> fitz.Rect:
    fit = (fit or "contain").lower()
    margin = max(0.0, float(margin))

    # content area inside margins
    x0, y0 = margin, margin
    x1, y1 = page_w - margin, page_h - margin
    cw, ch = max(1.0, x1 - x0), max(1.0, y1 - y0)

    if fit == "stretch":
        return fitz.Rect(x0, y0, x1, y1)

    # contain or cover
    if fit == "cover":
        scale = max(cw / img_w, ch / img_h)
    else:  # "contain"
        scale = min(cw / img_w, ch / img_h)

    tw, th = img_w * scale, img_h * scale
    left = x0 + (cw - tw) / 2
    top = y0 + (ch - th) / 2
    return fitz.Rect(left, top, left + tw, top + th)

#______Convert word to pdf________


def convert_docx_to_pdf_bytes(docx_bytes: bytes, input_filename: str = "input.docx") -> bytes:
    """
    Convert DOCX bytes to PDF bytes using LibreOffice (soffice) headless.
    Requires LibreOffice installed and 'soffice' available in PATH.
    """
    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)
        in_path = tmp_path / input_filename
        out_dir = tmp_path / "out"
        out_dir.mkdir(parents=True, exist_ok=True)

        in_path.write_bytes(docx_bytes)

        # LibreOffice conversion command
        cmd = [
            "soffice",
            "--headless",
            "--nologo",
            "--nofirststartwizard",
            "--convert-to", "pdf",
            "--outdir", str(out_dir),
            str(in_path),
        ]

        # Run conversion
        result = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)

        if result.returncode != 0:
            raise RuntimeError(
                "LibreOffice conversion failed.\n"
                f"STDOUT: {result.stdout}\n"
                f"STDERR: {result.stderr}"
            )

        pdf_path = out_dir / (in_path.stem + ".pdf")
        if not pdf_path.exists():
            raise FileNotFoundError("PDF output not created by LibreOffice.")

        return pdf_path.read_bytes()

#_____main function to convert images to PDF bytes_____
def convert_images_to_pdf_bytes(
    images: List[Tuple[str, bytes]],
    page_size: str = "A4",          
    fit: str = "contain",           
    margin: float = 18,             
    assume_dpi: int = 96,           
    jpeg_quality: int = 85,
    sort_by_name: bool = True,
) -> bytes:
    
    if not images:
        raise ValueError("No images provided.")

    imgs = images[:]
    if sort_by_name:
        imgs.sort(key=lambda x: (x[0] or "").lower())

    doc = fitz.open()

    try:
        for fname, data in imgs:
            if not data:
                continue

            pil = Image.open(io.BytesIO(data))
            pil = ImageOps.exif_transpose(pil) 

            if (page_size or "").lower() == "auto":
                dpi = max(1, int(assume_dpi))
                page_w = pil.width * 72.0 / dpi
                page_h = pil.height * 72.0 / dpi
            else:
                ps = page_size if page_size in _PAGE_SIZES_PT else "A4"
                page_w, page_h = _PAGE_SIZES_PT[ps]

            page = doc.new_page(width=page_w, height=page_h)

            has_alpha = pil.mode in ("RGBA", "LA") or ("transparency" in pil.info)

            img_buf = io.BytesIO()
            if has_alpha:
                pil_rgba = pil.convert("RGBA")
                pil_rgba.save(img_buf, format="PNG", optimize=True)
            else:
                pil_rgb = pil.convert("RGB")
                pil_rgb.save(img_buf, format="JPEG", quality=int(jpeg_quality), optimize=True, progressive=True)

            img_stream = img_buf.getvalue()

            rect = _fit_rect(pil.width, pil.height, page_w, page_h, margin, fit)
            keep_prop = (fit or "contain").lower() != "stretch"

            page.insert_image(rect, stream=img_stream, keep_proportion=keep_prop)

        return doc.tobytes(deflate=True, garbage=4)

    finally:
        doc.close()


def convert_images_to_pdf(
    images: List[Tuple[str, bytes]],
    output_path: Union[str, Path],
    **kwargs,
) -> Path:
    """File-saving wrapper (optional)."""
    out = Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    pdf_bytes = convert_images_to_pdf_bytes(images, **kwargs)
    out.write_bytes(pdf_bytes)
    return out
