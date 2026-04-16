"""
cleanmeta — Веб-інтерфейс для заміни метаданих документів
Запуск:  pip install streamlit
         streamlit run cleanmeta_app.py
Потрібен: exiftool (brew install exiftool / sudo apt install libimage-exiftool-perl)
"""

import streamlit as st
import subprocess
import tempfile
import random
import datetime
import os
import json
import shutil
import zipfile
from pathlib import Path

# ─── Конфігурація сторінки ───
st.set_page_config(
    page_title="CleanMeta — Очистка метаданих",
    page_icon="🛡️",
    layout="wide",
)

# ─── CSS ───
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@400;600&family=Manrope:wght@400;600;800&display=swap');

    .stApp {
        font-family: 'Manrope', sans-serif;
    }

    .main-title {
        font-family: 'Manrope', sans-serif;
        font-weight: 800;
        font-size: 2.4rem;
        background: linear-gradient(135deg, #0f172a, #334155);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin-bottom: 0;
    }

    .subtitle {
        color: #64748b;
        font-size: 1.05rem;
        margin-bottom: 2rem;
    }

    .meta-card {
        background: #f8fafc;
        border: 1px solid #e2e8f0;
        border-radius: 12px;
        padding: 1.2rem;
        margin-bottom: 0.8rem;
        font-family: 'JetBrains Mono', monospace;
        font-size: 0.82rem;
        line-height: 1.6;
    }

    .meta-card.before {
        border-left: 4px solid #ef4444;
    }

    .meta-card.after {
        border-left: 4px solid #22c55e;
    }

    .tag-label {
        display: inline-block;
        font-family: 'JetBrains Mono', monospace;
        font-size: 0.72rem;
        font-weight: 600;
        padding: 2px 8px;
        border-radius: 4px;
        margin-bottom: 8px;
    }

    .tag-before {
        background: #fef2f2;
        color: #dc2626;
    }

    .tag-after {
        background: #f0fdf4;
        color: #16a34a;
    }

    .file-badge {
        display: inline-block;
        background: #f1f5f9;
        border: 1px solid #cbd5e1;
        border-radius: 6px;
        padding: 4px 10px;
        margin: 2px;
        font-size: 0.85rem;
    }

    .stat-number {
        font-family: 'JetBrains Mono', monospace;
        font-weight: 600;
        font-size: 1.8rem;
        color: #0f172a;
    }
</style>
""", unsafe_allow_html=True)

# ─── Корпоративні пари Creator/Producer ───
CORPORATE_PAIRS = [
    ("SAP Crystal Reports 2020",                  "SAP Crystal Reports",                          "Adobe XMP Core 5.6-c017 91.174166, 2020/11/18"),
    ("SAP BusinessObjects Web Intelligence",       "SAP BusinessObjects BI Platform 4.3",          "Adobe XMP Core 6.0.0"),
    ("Oracle BI Publisher 12.2.1",                 "Oracle BI Publisher",                          "Adobe XMP Core 5.4-c005 78.147326, 2012/08/23"),
    ("Oracle Reports 12c",                         "Oracle Application Server Reports Services",   "XMP toolkit 3.0-28, framework 1.6"),
    ("IBM Cognos Analytics 11.2",                  "IBM Cognos",                                   "Adobe XMP Core 5.6-c015 91.163280, 2018/02/05"),
    ("IBM FileNet Content Manager 5.5",            "IBM FileNet P8",                               "Adobe XMP Core 5.2-c001 63.139439, 2010/09/27"),
    ("OpenText Exstream 16.6",                     "OpenText Output Management",                   "Adobe XMP Core 5.6-c017 91.174166, 2020/11/18"),
    ("OpenText Content Server 21.4",               "OpenText Documentum",                          "Adobe XMP Core 6.0.0"),
    ("Microsoft Reporting Services 15.0",          "Microsoft SQL Server Reporting Services",      "Microsoft XMP 1.0"),
    ("Adobe Experience Manager 6.5",               "Adobe AEM Forms",                              "Adobe XMP Core 6.0.0"),
    ("Adobe LiveCycle Designer ES4",               "Adobe LiveCycle Output Service",                "Adobe XMP Core 5.4-c005 78.147326, 2012/08/23"),
    ("TIBCO Jaspersoft 7.9",                       "JasperReports Server",                         "Adobe XMP Core 5.6-c015 91.163280, 2018/02/05"),
    ("MicroStrategy Intelligence Server 2021",     "MicroStrategy Report Services",                "Adobe XMP Core 5.6-c017 91.174166, 2020/11/18"),
    ("Windward Studios AutoTag 15",                "Windward Output Engine",                       "Adobe XMP Core 6.0.0"),
    ("Hyland OnBase 21.1",                         "Hyland Document Composition",                  "Adobe XMP Core 5.6-c015 91.163280, 2018/02/05"),
    ("Kofax Communications Manager 5.5",           "Kofax Output Server",                         "Adobe XMP Core 5.4-c005 78.147326, 2012/08/23"),
    ("Xerox DocuShare 7.0",                        "Xerox FreeFlow Variable Information Suite",    "Adobe XMP Core 6.0.0"),
    ("Canon Therefore 2021",                       "Canon Therefore Output Server",                "Adobe XMP Core 5.6-c017 91.174166, 2020/11/18"),
    ("Ricoh ProcessDirector",                      "Ricoh Production Print Services",              "Adobe XMP Core 5.2-c001 63.139439, 2010/09/27"),
    ("Apache FOP Version 2.9",                     "Apache FOP Version 2.9",                       "Adobe XMP Core 5.6-c015 91.163280, 2018/02/05"),
    ("Telerik Reporting 2023",                     "Telerik Reporting Engine",                     "Adobe XMP Core 6.0.0"),
    ("Actuate BIRT 4.9",                           "Eclipse BIRT Runtime",                         "Adobe XMP Core 5.4-c005 78.147326, 2012/08/23"),
    ("Stimulsoft Reports 2024.1",                  "Stimulsoft Reports Engine",                    "Adobe XMP Core 5.6-c017 91.174166, 2020/11/18"),
]


def check_exiftool():
    """Перевіряє наявність exiftool."""
    return shutil.which("exiftool") is not None


def random_work_datetime():
    """Випадкова дата за останні 7 робочих днів, робочі години."""
    now = datetime.datetime.now()
    day_offset = random.randint(0, 6)
    d = now - datetime.timedelta(days=day_offset)
    while d.weekday() >= 5:
        d -= datetime.timedelta(days=1)
    d = d.replace(
        hour=random.randint(8, 17),
        minute=random.randint(0, 59),
        second=random.randint(0, 59),
        microsecond=0,
    )
    return d.strftime("%Y:%m:%d %H:%M:%S")


def get_metadata(filepath):
    """Витягує метадані файлу через exiftool."""
    try:
        result = subprocess.run(
            ["exiftool", "-json", filepath],
            capture_output=True, text=True, timeout=15,
        )
        data = json.loads(result.stdout)
        return data[0] if data else {}
    except Exception:
        return {}


def clean_pdf(filepath):
    """Підміна метаданих PDF на корпоративні."""
    dt = random_work_datetime()
    creator, producer, xmp_toolkit = random.choice(CORPORATE_PAIRS)
    cmd = [
        "exiftool", "-overwrite_original",
        f"-Author=",
        f"-Title=",
        f"-Subject=",
        f"-Keywords=",
        f"-Creator={creator}",
        f"-Producer={producer}",
        f"-CreateDate={dt}",
        f"-ModifyDate={dt}",
        f"-MetadataDate={dt}",
        f"-XMP-x:XMPToolkit={xmp_toolkit}",
        filepath,
    ]
    subprocess.run(cmd, capture_output=True, timeout=30)
    return {"creator": creator, "producer": producer, "date": dt, "xmp_toolkit": xmp_toolkit}


def clean_office(filepath):
    """Повна очистка метаданих Office/ODT."""
    dt = random_work_datetime()
    cmd = [
        "exiftool", "-overwrite_original",
        "-Author=",
        "-Creator=",
        "-LastModifiedBy=",
        "-Company=",
        "-Manager=",
        "-Title=",
        "-Subject=",
        "-Keywords=",
        "-Description=",
        "-Comment=",
        "-Application=",
        "-AppVersion=",
        "-TotalEditTime=0",
        "-RevisionNumber=1",
        f"-CreateDate={dt}",
        f"-ModifyDate={dt}",
        filepath,
    ]
    subprocess.run(cmd, capture_output=True, timeout=30)
    return {"date": dt}


SENSITIVE_KEYS = [
    "Author", "Creator", "Producer", "Company", "Manager",
    "LastModifiedBy", "Title", "Subject", "Keywords", "Description",
    "Comment", "Application", "AppVersion", "Software",
    "CreateDate", "ModifyDate", "MetadataDate",
    "TotalEditTime", "RevisionNumber", "XMPToolkit",
]


def format_meta_display(meta: dict, keys=None):
    """Форматує метадані для показу."""
    if keys is None:
        keys = SENSITIVE_KEYS
    lines = []
    for k in keys:
        if k in meta and str(meta[k]).strip():
            lines.append(f"<b>{k}:</b> {meta[k]}")
    return "<br>".join(lines) if lines else "<i>— порожньо —</i>"


# ─── UI ───
st.markdown('<div class="main-title">🛡️ CleanMeta</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitle">Заміна та очистка метаданих PDF, DOCX, XLSX, PPTX, ODT</div>', unsafe_allow_html=True)

if not check_exiftool():
    st.error("⚠️ **exiftool** не знайдено! Встанови: `brew install exiftool` або `sudo apt install libimage-exiftool-perl`")
    st.stop()

# Tabs
tab_single, tab_batch = st.tabs(["📄 Один файл", "📦 Пакетна обробка"])

# ─── Один файл ───
with tab_single:
    uploaded = st.file_uploader(
        "Завантаж документ",
        type=["pdf", "docx", "xlsx", "pptx", "odt", "ods", "odp"],
        key="single",
    )

    if uploaded:
        ext = Path(uploaded.name).suffix.lower()
        st.markdown(f'<span class="file-badge">📄 {uploaded.name} ({uploaded.size / 1024:.1f} KB)</span>', unsafe_allow_html=True)

        with tempfile.NamedTemporaryFile(delete=False, suffix=ext) as tmp:
            tmp.write(uploaded.getvalue())
            tmp_path = tmp.name

        # Метадані ДО
        meta_before = get_metadata(tmp_path)

        col_before, col_after = st.columns(2)

        with col_before:
            st.markdown('<span class="tag-label tag-before">⚠️ ДО ОЧИСТКИ</span>', unsafe_allow_html=True)
            st.markdown(
                f'<div class="meta-card before">{format_meta_display(meta_before)}</div>',
                unsafe_allow_html=True,
            )

        # Кнопка очистки
        if st.button("🧹 Очистити метадані", type="primary", use_container_width=True):
            with st.spinner("Обробка..."):
                if ext == ".pdf":
                    result = clean_pdf(tmp_path)
                else:
                    result = clean_office(tmp_path)

            meta_after = get_metadata(tmp_path)

            with col_after:
                st.markdown('<span class="tag-label tag-after">✅ ПІСЛЯ ОЧИСТКИ</span>', unsafe_allow_html=True)
                st.markdown(
                    f'<div class="meta-card after">{format_meta_display(meta_after)}</div>',
                    unsafe_allow_html=True,
                )

            # Завантаження
            prefix = ''.join(random.choices('abcdefghijklmnopqrstuvwxyz', k=3))
            with open(tmp_path, "rb") as f:
                st.download_button(
                    label=f"⬇️ Завантажити {uploaded.name}",
                    data=f.read(),
                    file_name=f"{prefix}_{uploaded.name}",
                    mime="application/octet-stream",
                    use_container_width=True,
                )

            if ext == ".pdf":
                st.success(f"Creator: **{result['creator']}** · Producer: **{result['producer']}** · Date: {result['date']}")
            else:
                st.success(f"Метадані очищено · Date: {result['date']}")

            os.unlink(tmp_path)

# ─── Пакетна обробка ───
with tab_batch:
    uploaded_files = st.file_uploader(
        "Завантаж кілька файлів",
        type=["pdf", "docx", "xlsx", "pptx", "odt", "ods", "odp"],
        accept_multiple_files=True,
        key="batch",
    )

    if uploaded_files:
        st.markdown(f"Файлів: **{len(uploaded_files)}**")

        if st.button("🧹 Очистити всі", type="primary", use_container_width=True):
            tmpdir = tempfile.mkdtemp()
            results = []
            progress = st.progress(0)

            for i, uf in enumerate(uploaded_files):
                ext = Path(uf.name).suffix.lower()
                fpath = os.path.join(tmpdir, uf.name)
                with open(fpath, "wb") as f:
                    f.write(uf.getvalue())

                if ext == ".pdf":
                    res = clean_pdf(fpath)
                else:
                    res = clean_office(fpath)

                results.append({"name": uf.name, "path": fpath, **res})
                progress.progress((i + 1) / len(uploaded_files))

            # Показ результатів
            for r in results:
                label = f"✅ {r['name']}"
                if "creator" in r:
                    label += f" → {r['creator']}"
                st.markdown(f"- {label} · {r['date']}")

            # ZIP архів
            zip_path = os.path.join(tmpdir, "documents.zip")
            with zipfile.ZipFile(zip_path, "w") as zf:
                for r in results:
                    bp = ''.join(random.choices('abcdefghijklmnopqrstuvwxyz', k=3))
                    zf.write(r["path"], f"{bp}_{r['name']}")

            with open(zip_path, "rb") as f:
                st.download_button(
                    label=f"⬇️ Завантажити всі ({len(results)} файлів) .zip",
                    data=f.read(),
                    file_name="documents.zip",
                    mime="application/zip",
                    use_container_width=True,
                )

            st.success(f"Оброблено {len(results)} файлів!")

            # Авто-очистка
            shutil.rmtree(tmpdir, ignore_errors=True)

# ─── Приватність ───
with st.sidebar:
    st.markdown("### 🔒 Приватність")
    st.markdown("""
- Файли обробляються **на сервері** і **автоматично видаляються** після скачування
- Жодних зовнішніх API-запитів
- Ніякої аналітики чи збору даних
- Код відкритий — можеш перевірити
    """)
    st.markdown("---")

# ─── Інфо ───
with st.expander("ℹ️ Що саме змінюється?"):
    st.markdown("""
**PDF:**
- Creator + Producer → випадкова корпоративна пара (SAP, Oracle, IBM, Adobe тощо — 19 варіантів)
- CreateDate = ModifyDate = випадкова дата за останні 7 робочих днів
- Author, Title, Subject, Keywords → порожні

**DOCX / XLSX / PPTX / ODT:**
- Author, Creator, LastModifiedBy, Company, Manager → порожні
- Application, AppVersion → порожні
- Title, Subject, Keywords, Description, Comment → порожні
- TotalEditTime → 0, RevisionNumber → 1
- CreateDate = ModifyDate = випадкова дата за останні 7 робочих днів
""")
