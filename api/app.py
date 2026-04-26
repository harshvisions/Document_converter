from __future__ import annotations

import re
from pathlib import Path
from typing import Optional, List

import fitz
import pandas as pd
import streamlit as st

from converter import (
    convert_pdf_to_memory_zip,
    convert_images_to_pdf_bytes,
    ensure_unique_names,
    parse_page_ranges,
    sanitize_stem,
)

APP_NAME = "PDF Tools"
TAGLINE = "Fast, private conversions for students and office work."
VERSION = "1.0.0"

BASE_DIR = Path(__file__).parent
ASSETS_DIR = BASE_DIR / "assets"
CSS_COMMON = ASSETS_DIR / "styles_common.css"
CSS_LIGHT = ASSETS_DIR / "styles_light.css"
CSS_DARK = ASSETS_DIR / "styles_dark.css"


def load_css(path: Path) -> None:
    """Inject a local CSS file into the Streamlit app."""
    if path.exists():
        st.markdown(f"<style>{path.read_text(encoding='utf-8')}</style>", unsafe_allow_html=True)


def sanitize_zip_name(name: str) -> str:
    name = (name or "").strip()
    name = re.sub(r"\s+", "_", name)
    name = re.sub(r"[^A-Za-z0-9._-]+", "_", name)
    return name.strip("._-") or "images"


def parse_custom_names(text: str) -> Optional[List[str]]:
    """Parse custom names (newline or comma-separated)."""
    if not text or not text.strip():
        return None
    clean = text.replace(',', '\n')
    names = [x.strip() for x in clean.splitlines() if x.strip()]
    return names or None



def auth_pdf(doc: fitz.Document, password: Optional[str]) -> None:
    if doc.needs_pass:
        if not password:
            raise PermissionError("This PDF is encrypted. Please provide a password.")
        if not doc.authenticate(password):
            raise PermissionError("Incorrect password for encrypted PDF.")


@st.cache_data(show_spinner=False)
def first_page_thumbnail(pdf_bytes: bytes, password: Optional[str]) -> bytes:
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    try:
        auth_pdf(doc, password)
        page = doc.load_page(0)
        pix = page.get_pixmap(matrix=fitz.Matrix(0.35, 0.35), alpha=False, colorspace=fitz.csRGB)
        return pix.tobytes("png")
    finally:
        doc.close()


def preview_pdf_to_images(
    pdf_bytes: bytes,
    file_stem: str,
    pages_spec: Optional[str],
    naming_mode: str,
    custom_names_text: Optional[str],
    password: Optional[str],
) -> pd.DataFrame:
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    try:
        auth_pdf(doc, password)
        total = doc.page_count
        page_indices = parse_page_ranges(pages_spec, total)

        if naming_mode == "Auto":
            stems = [f"{sanitize_stem(file_stem)}_page_{i+1:03d}" for i in page_indices]
        else:
            names = parse_custom_names(custom_names_text) or []
            if len(names) != len(page_indices):
                raise ValueError(
                    f"Custom names count must match selected pages. Expected {len(page_indices)}, got {len(names)}."
                )
            stems = ensure_unique_names([sanitize_stem(n) for n in names])

        rows = [{"Page": i + 1, "Filename": f"{stem}.jpg"} for i, stem in zip(page_indices, stems)]
        return pd.DataFrame(rows)
    finally:
        doc.close()


TOOLS = [
    {"id": "home", "title": "Home", "desc": "Dashboard", "icon": "🏠", "status": "ready"},
    {"id": "pdf_to_images", "title": "PDF to Images", "desc": "Export PDF pages as high-quality JPG.", "icon": "📄", "status": "ready"},
    {"id": "images_to_pdf", "title": "Images to PDF", "desc": "Combine images into a single PDF.", "icon": "🖼️", "status": "ready"},
    {"id": "merge_pdfs", "title": "Merge PDFs", "desc": "Combine multiple PDFs into one file.", "icon": "🧩", "status": "soon"},
    {"id": "compress_pdf", "title": "Compress PDF", "desc": "Reduce PDF size (coming soon).", "icon": "🗜️", "status": "soon"},
]


def set_tool(tool_id: str) -> None:
    st.session_state.active_tool = tool_id
    try:
        st.query_params["tool"] = tool_id
    except Exception:
        st.experimental_set_query_params(tool=tool_id)


def get_tool_from_query() -> Optional[str]:
    try:
        return st.query_params.get("tool")
    except Exception:
        qp = st.experimental_get_query_params()
        return (qp.get("tool") or [None])[0]


# ----------------- Page config -----------------
st.set_page_config(page_title=APP_NAME, page_icon="🧰", layout="wide", initial_sidebar_state="expanded")

if "theme" not in st.session_state:
    st.session_state.theme = "Light"

# Load CSS
load_css(CSS_COMMON)
load_css(CSS_DARK if st.session_state.theme == "Dark" else CSS_LIGHT)

if "active_tool" not in st.session_state:
    st.session_state.active_tool = "home"

query_tool = get_tool_from_query()
known_ids = {t["id"] for t in TOOLS}
if query_tool in known_ids:
    st.session_state.active_tool = query_tool


# ----------------- Sidebar -----------------
with st.sidebar:
    st.markdown(f"### 🧰 {APP_NAME}")
    st.caption(TAGLINE)

    st.session_state.theme = st.radio("Theme", ["Light", "Dark"], horizontal=True)

    st.markdown("---")

    labels = []
    label_to_id = {}
    for t in TOOLS:
        label = f"{t['icon']}  {t['title']}"
        labels.append(label)
        label_to_id[label] = t["id"]

    current_label = next((lbl for lbl, tid in label_to_id.items() if tid == st.session_state.active_tool), labels[0])
    choice = st.radio("Navigation", labels, index=labels.index(current_label))
    set_tool(label_to_id[choice])

    st.markdown("---")
    st.markdown(
        f"<div class='small-note'>Version {VERSION}<br/>Files are processed in memory during this session.</div>",
        unsafe_allow_html=True,
    )


def topbar() -> None:
    st.markdown(
        f"""
        <div class="topbar">
          <div class="brand">
            <h1>🧰 {APP_NAME}</h1>
            <p>{TAGLINE}</p>
          </div>
          <div class="small-note">Theme: <b>{st.session_state.theme}</b> • v{VERSION}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_home() -> None:
    topbar()
    st.write("")

    st.markdown(
        """
        <div class="note">
          <b>Quick tips</b>
          <p>
            • For sharper text (notes/assignments), use <b>DPI 200–300</b>.<br/>
            • Use <b>page ranges</b> like <code>1,3-5,7-</code> to export specific pages.<br/>
            • Use <b>Auto naming</b> for consistent results.
          </p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.write("")
    q = st.text_input("Search tools", placeholder="Type: images, pdf, merge …")

    tiles = [t for t in TOOLS if t["id"] != "home"]
    if q:
        ql = q.lower().strip()
        tiles = [t for t in tiles if ql in (t["title"] + " " + t["desc"] + " " + t["id"]).lower()]

    st.markdown("<div class='grid'>", unsafe_allow_html=True)
    for t in tiles:
        badge_class = "ready" if t["status"] == "ready" else "soon"
        st.markdown(
            f"""
            <div class="card">
              <div class="title">
                <h3>{t['icon']} {t['title']}</h3>
                <span class="badge {badge_class}">{t['status'].upper()}</span>
              </div>
              <p>{t['desc']}</p>
            </div>
            """,
            unsafe_allow_html=True,
        )
        st.button(
            "Open",
            type="primary",
            use_container_width=True,
            disabled=(t["status"] != "ready"),
            on_click=set_tool,
            args=(t["id"],),
            key=f"open_{t['id']}",
        )
        st.write("")
    st.markdown("</div>", unsafe_allow_html=True)


def render_pdf_to_images() -> None:
    topbar()
    st.markdown("## 📄 PDF → Images (JPG)")
    st.caption("Export selected pages as JPG images. Great for notes, handouts, and office documents.")

    uploaded = st.file_uploader("Upload PDF files", type=["pdf"], accept_multiple_files=True)

    colA, colB, colC, colD = st.columns(4)
    with colA:
        dpi = st.slider("DPI (quality)", 72, 400, 220, help="Higher DPI = sharper images but slower.")
    with colB:
        quality = st.slider("JPG quality", 50, 95, 90)
    with colC:
        pages_spec = st.text_input("Pages (e.g. 1,3-5,7-)", value="")
    with colD:
        password = st.text_input("Password (if encrypted)", value="", type="password")

    naming_mode = st.radio("Naming", ["Auto", "Custom"], horizontal=True)
    custom_names_text = None
    if naming_mode == "Custom":
        custom_names_text = st.text_area(
            "Custom names (one per page; newline or comma-separated)",
            placeholder="intro\nchapter_1\nchapter_2",
            height=120,
        )

    st.write("")
    st.subheader("Preview")
    st.caption("Review filenames before converting.")

    if uploaded:
        for f in uploaded:
            pdf_bytes = f.getvalue()
            stem = Path(f.name).stem

            try:
                thumb = first_page_thumbnail(pdf_bytes, password.strip() or None)
                st.image(thumb, width=160, caption=f"{f.name} (page 1)")
            except Exception as e:
                st.warning(f"Thumbnail unavailable for {f.name}: {e}")

            try:
                df = preview_pdf_to_images(
                    pdf_bytes=pdf_bytes,
                    file_stem=stem,
                    pages_spec=pages_spec.strip() or None,
                    naming_mode=naming_mode,
                    custom_names_text=custom_names_text,
                    password=password.strip() or None,
                )
                st.dataframe(df, use_container_width=True, hide_index=True)
            except Exception as e:
                st.error(f"Preview error for {f.name}: {e}")
    else:
        st.info("Upload at least one PDF to see preview.")

    st.write("")
    st.subheader("Convert & download")

    run = st.button("Convert to JPG (ZIP)", type="primary", use_container_width=True, disabled=not uploaded)

    if run and uploaded:
        for f in uploaded:
            pdf_bytes = f.getvalue()
            stem = Path(f.name).stem
            zip_name = sanitize_zip_name(stem) + "_images.zip"

            bar = st.progress(0, text="Starting…")

            def progress_cb(done: int, total: int):
                pct = 0 if total == 0 else int((done / total) * 100)
                bar.progress(min(max(pct, 0), 100), text=f"Rendering {done}/{total} pages…")

            try:
                custom_names = parse_custom_names(custom_names_text) if naming_mode == "Custom" else None

                zip_bytes, _members = convert_pdf_to_memory_zip(
                    pdf_source=pdf_bytes,
                    dpi=int(dpi),
                    quality=int(quality),
                    pages=pages_spec.strip() or None,
                    password=password.strip() or None,
                    basename=stem,
                    custom_names=custom_names,
                    progress_cb=progress_cb,
                )

                bar.progress(100, text="Done ✅")
                st.success(f"Converted: {f.name}")

                st.download_button(
                    "Download ZIP",
                    data=zip_bytes,
                    file_name=zip_name,
                    mime="application/zip",
                    use_container_width=True,
                    key=f"dl_{f.name}",
                )
                st.write("")
            except Exception as e:
                bar.empty()
                st.error(f"Conversion failed for {f.name}: {e}")


def render_images_to_pdf() -> None:
    topbar()
    st.markdown("## 🖼️ Images → PDF")
    st.caption("Upload images and export one PDF (one image per page).")

    uploaded = st.file_uploader(
        "Upload images",
        type=["png", "jpg", "jpeg", "webp"],
        accept_multiple_files=True,
    )

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        page_size = st.selectbox("Page size", ["A4", "Letter", "auto"], index=0)
    with col2:
        fit = st.selectbox("Fit mode", ["contain", "cover", "stretch"], index=0)
    with col3:
        margin = st.slider("Margin (pt)", 0, 72, 18, help="72pt = 1 inch")
    with col4:
        jpg_quality = st.slider("JPG quality", 50, 95, 85)

    assume_dpi = 96
    if page_size == "auto":
        assume_dpi = st.selectbox("Assume DPI (auto mode)", [72, 96, 150, 300], index=1)

    out_name = st.text_input("Output PDF name", value="images.pdf")

    if uploaded:
        st.subheader("Preview")
        thumbs = [f.getvalue() for f in uploaded[:12]]
        st.image(thumbs, width=140, caption=[f.name for f in uploaded[:12]])
        if len(uploaded) > 12:
            st.info(f"Preview shows first 12 images of {len(uploaded)}.")

    run = st.button("Convert to PDF", type="primary", use_container_width=True, disabled=not uploaded)

    if run and uploaded:
        try:
            images = [(f.name, f.getvalue()) for f in uploaded]
            pdf_bytes = convert_images_to_pdf_bytes(
                images,
                page_size=page_size,
                fit=fit,
                margin=float(margin),
                assume_dpi=int(assume_dpi),
                jpeg_quality=int(jpg_quality),
                sort_by_name=True,
            )

            st.success("PDF created ✅")
            st.download_button(
                "Download PDF",
                data=pdf_bytes,
                file_name=out_name if out_name.lower().endswith(".pdf") else out_name + ".pdf",
                mime="application/pdf",
                use_container_width=True,
            )
        except Exception as e:
            st.error(f"Conversion failed: {e}")


# ----------------- Router -----------------
active = st.session_state.active_tool
if active == "home":
    render_home()
elif active == "pdf_to_images":
    render_pdf_to_images()
elif active == "images_to_pdf":
    render_images_to_pdf()
else:
    topbar()
    st.warning("This tool is not implemented yet (coming soon).")
    st.button("Back to Home", use_container_width=True, on_click=set_tool, args=("home",))
