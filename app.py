"""
CV Reformatter — aplikasi lokal (Streamlit)

Cara jalanin:
    pip install -r requirements.txt
    streamlit run app.py

Alur pemakaian:
    1. Upload TEMPLATE (contoh CV yang formatnya mau ditiru) -> harus file .docx
    2. Upload CV SUMBER (CV orang yang mau diformat ulang) -> boleh .docx atau .pdf
    3. Klik "Ekstrak Data CV" -> aplikasi otomatis membaca & memilah isi CV
    4. Cek & koreksi hasil ekstraksi di layar review (opsional tapi disarankan,
       karena ekstraksi berbasis aturan/kata kunci, bukan AI, jadi tidak 100% sempurna)
    5. Klik "Buat CV Sesuai Template" -> unduh hasil .docx
"""

import copy
import os
import tempfile

import streamlit as st

from cv_parser import (
    ENTRY_SECTIONS,
    LIST_SECTIONS,
    _load_keywords,
    parse_cv,
)
from cv_builder import DEFAULT_LABELS, analyze_template, build_cv

st.set_page_config(page_title="CV Reformatter", page_icon="📄", layout="centered")

SECTION_LABELS_ID = {
    "summary": "Ringkasan",
    "experience": "Pengalaman Kerja",
    "education": "Pendidikan",
    "skills": "Keahlian",
    "certifications": "Sertifikasi",
    "languages": "Bahasa",
    "organizations": "Organisasi",
    "projects": "Proyek",
    "awards": "Penghargaan",
    "references": "Referensi",
}


def _save_upload_to_tmp(uploaded_file):
    suffix = os.path.splitext(uploaded_file.name)[1].lower()
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    tmp.write(uploaded_file.getvalue())
    tmp.close()
    return tmp.name


def _blank_entry():
    return {"title": "", "subtitle": "", "date": "", "bullets": []}


def _check_password():
    """
    Gerbang password sederhana & OPSIONAL.

    Kalau aplikasi dijalankan lokal (tanpa apa-apa disetel), fungsi ini tidak melakukan apa-apa
    (langsung lolos). Kalau di-deploy ke web (misal Streamlit Community Cloud) dan diberi CV
    asli orang lain, sebaiknya aktifkan ini supaya link-nya tidak bisa dibuka sembarang orang:
    di Streamlit Cloud, buka menu App -> Settings -> Secrets, lalu isi:

        APP_PASSWORD = "password-rahasia-kamu"

    Setelah itu, aplikasi akan minta password ini sebelum bisa dipakai.
    """
    try:
        app_password = st.secrets.get("APP_PASSWORD")
    except Exception:
        app_password = None
    if not app_password:
        return True
    if st.session_state.get("_authed"):
        return True
    st.title("📄 CV Reformatter")
    st.info("Aplikasi ini dilindungi password. Minta password ke pemilik aplikasi kalau belum punya.")
    pwd = st.text_input("Password akses", type="password")
    if pwd:
        if pwd == app_password:
            st.session_state["_authed"] = True
            st.rerun()
        else:
            st.error("Password salah.")
    return False


if not _check_password():
    st.stop()

st.title("📄 CV Reformatter")
st.caption(
    "Upload template (contoh format yang diinginkan) + CV sumber (CV orang lain). "
    "Aplikasi ini otomatis mengisi ulang pengalaman, pendidikan, skill, dll ke dalam format template — "
    "berjalan sepenuhnya lokal, tanpa API AI berbayar."
)

with st.expander("ℹ️ Catatan penting sebelum mulai", expanded=False):
    st.markdown(
        "- **Template** wajib file **.docx** (dipakai untuk meniru gaya tulisan & urutan bagian).\n"
        "- **CV sumber** boleh **.docx** atau **.pdf**.\n"
        "- Ekstraksi memakai aturan & kata kunci (bilingual ID/EN), **bukan AI**, jadi untuk CV dengan "
        "format tidak umum, hasilnya mungkin perlu dikoreksi manual di langkah review sebelum di-generate.\n"
        "- Kalau bagian tertentu di CV sumber tidak ada di template, bagian itu tetap ditambahkan di akhir "
        "dokumen hasil (datanya tidak dibuang begitu saja)."
    )

col1, col2 = st.columns(2)
with col1:
    template_file = st.file_uploader("Template CV (.docx)", type=["docx"], key="template")
with col2:
    source_file = st.file_uploader("CV Sumber (.docx / .pdf)", type=["docx", "pdf"], key="source")

if "parsed" not in st.session_state:
    st.session_state.parsed = None

extract_clicked = st.button("🔍 Ekstrak Data CV", disabled=source_file is None, type="primary")

if extract_clicked and source_file is not None:
    with st.spinner("Membaca & memilah isi CV..."):
        src_path = _save_upload_to_tmp(source_file)
        try:
            keywords = _load_keywords()
            st.session_state.parsed = parse_cv(src_path, keywords)
            st.success("Ekstraksi selesai. Cek & koreksi hasilnya di bawah kalau perlu.")
        except Exception as e:
            st.error(f"Gagal mengekstrak CV: {e}")
        finally:
            os.unlink(src_path)

parsed = st.session_state.parsed

if parsed:
    st.divider()
    st.subheader("✏️ Review & koreksi hasil ekstraksi")

    with st.expander("Kontak / Data Diri", expanded=True):
        c = parsed["contact"]
        c["name"] = st.text_input("Nama", c.get("name", ""))
        cc1, cc2 = st.columns(2)
        c["email"] = cc1.text_input("Email", c.get("email", ""))
        c["phone"] = cc2.text_input("Telepon", c.get("phone", ""))
        cc3, cc4 = st.columns(2)
        c["linkedin"] = cc3.text_input("LinkedIn", c.get("linkedin", ""))
        c["location"] = cc4.text_input("Lokasi / info tambahan", c.get("location", ""))

    for section_id in list(parsed["order"]):
        label = SECTION_LABELS_ID.get(section_id, section_id.title())
        content = parsed["sections"].get(section_id)

        if section_id == "summary":
            with st.expander(label, expanded=False):
                parsed["sections"][section_id] = st.text_area(
                    "Isi ringkasan", content or "", key=f"summary_text", height=100
                )

        elif section_id in ENTRY_SECTIONS:
            with st.expander(f"{label} ({len(content)} entri)", expanded=False):
                to_delete = []
                for i, entry in enumerate(content):
                    st.markdown(f"**Entri {i + 1}**")
                    e1, e2 = st.columns(2)
                    entry["title"] = e1.text_input("Judul / Posisi", entry.get("title", ""), key=f"{section_id}_{i}_title")
                    entry["subtitle"] = e2.text_input(
                        "Sub-judul / Institusi", entry.get("subtitle", ""), key=f"{section_id}_{i}_subtitle"
                    )
                    entry["date"] = st.text_input("Periode / Tanggal", entry.get("date", ""), key=f"{section_id}_{i}_date")
                    bullets_text = st.text_area(
                        "Poin-poin (1 baris = 1 poin)",
                        "\n".join(entry.get("bullets", [])),
                        key=f"{section_id}_{i}_bullets",
                        height=90,
                    )
                    entry["bullets"] = [b.strip() for b in bullets_text.split("\n") if b.strip()]
                    if st.checkbox("Hapus entri ini", key=f"{section_id}_{i}_delete"):
                        to_delete.append(i)
                    st.markdown("---")
                for i in reversed(to_delete):
                    content.pop(i)
                if st.button(f"+ Tambah entri {label}", key=f"add_{section_id}"):
                    content.append(_blank_entry())
                    st.rerun()

        else:  # list sederhana: skills, languages, awards, references, dll
            with st.expander(label, expanded=False):
                items_text = st.text_area(
                    "Daftar item (1 baris = 1 item)",
                    "\n".join(content or []),
                    key=f"{section_id}_list",
                    height=100,
                )
                parsed["sections"][section_id] = [x.strip() for x in items_text.split("\n") if x.strip()]

    st.divider()
    generate_clicked = st.button(
        "🛠️ Buat CV Sesuai Template", type="primary", disabled=template_file is None
    )
    if template_file is None:
        st.info("Upload file template (.docx) di atas dulu untuk bisa generate hasil akhirnya.")

    if generate_clicked and template_file is not None:
        with st.spinner("Menyusun CV sesuai gaya template..."):
            tpl_path = _save_upload_to_tmp(template_file)
            out_path = None
            try:
                keywords = _load_keywords()
                analysis = analyze_template(tpl_path, keywords)
                out_fd, out_path = tempfile.mkstemp(suffix=".docx")
                os.close(out_fd)
                build_cv(copy.deepcopy(parsed), analysis, out_path)
                with open(out_path, "rb") as f:
                    result_bytes = f.read()
                st.success("CV berhasil dibuat!")
                base_name = os.path.splitext(source_file.name)[0] if source_file else "cv"
                st.download_button(
                    "⬇️ Unduh CV Hasil Reformat (.docx)",
                    data=result_bytes,
                    file_name=f"{base_name}_reformatted.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )
            except Exception as e:
                st.error(f"Gagal membuat CV: {e}")
            finally:
                if os.path.exists(tpl_path):
                    os.unlink(tpl_path)
                if out_path and os.path.exists(out_path):
                    os.unlink(out_path)
else:
    st.info("Upload CV sumber lalu klik **Ekstrak Data CV** untuk mulai.")
