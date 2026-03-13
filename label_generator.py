import streamlit as st
import pandas as pd
import datetime
import re
from docx import Document
from docx.shared import Cm, Pt
from pystrich.code128 import Code128Encoder
from PIL import Image, ImageDraw, ImageFont
from io import BytesIO

BARCODE_WIDTH     = 3.6
BARCODE_HEIGHT    = 1.9
BARCODE_TEXT_SIZE = 18
MAX_ROWS          = 8


# -------------------------------------------------------
# Hulpfuncties
# -------------------------------------------------------

def strip_spaces(value):
    """Verwijder alle spaties uit een veld."""
    return str(value).replace(" ", "").strip()


def generate_word_from_dataframe(df):
    """Genereer Word-document vanuit een DataFrame met interne kolomnamen:
    containertype, straat, huisnummer, toevoeging, postcode, woonplaats
    """
    output_doc = Document()
    section = output_doc.sections[-1]
    section.page_width    = Cm(10.0)
    section.page_height   = Cm(8.0)
    section.left_margin   = Cm(0.8)
    section.right_margin  = Cm(0.5)
    section.top_margin    = Cm(2)
    section.bottom_margin = Cm(0.5)

    try:
        font = ImageFont.truetype("arial.ttf", BARCODE_TEXT_SIZE)
    except IOError:
        font = ImageFont.load_default()

    dpi              = 96
    pixels_per_cm    = dpi / 2.54
    horizontal_shift = int(0.75 * pixels_per_cm)
    text_area_height = 30

    for label_idx, (idx, row) in enumerate(df.iterrows()):
        containertype = str(row.get('containertype', ''))
        straat        = str(row.get('straat', ''))
        huisnummer    = str(row.get('huisnummer', ''))
        toevoeging    = str(row.get('toevoeging', ''))
        postcode      = str(row.get('postcode', ''))
        woonplaats    = str(row.get('woonplaats', ''))
        barcode_value = f"{postcode}{huisnummer}{toevoeging}"

        encoder       = Code128Encoder(barcode_value)
        barcode_img   = encoder.get_imagedata()
        barcode_image = Image.open(BytesIO(barcode_img))

        bbox = barcode_image.getbbox()
        if bbox:
            left, _, right, _ = bbox
            barcode_image = barcode_image.crop((left, 0, right, barcode_image.height))

        draw = ImageDraw.Draw(barcode_image)
        width, height = barcode_image.size
        draw.rectangle([0, height - text_area_height, width, height], fill="white")

        text      = ""
        bbox_text = draw.textbbox((0, 0), text, font=font)
        text_y    = height - text_area_height + ((text_area_height - (bbox_text[3] - bbox_text[1])) / 2)
        draw.text((horizontal_shift, text_y), text, fill="black", font=font)

        barcode_buf = BytesIO()
        barcode_image.save(barcode_buf, format="PNG")
        barcode_buf.seek(0)

        # Voeg page break toe vóór elk label (behalve het eerste)
        if label_idx > 0:
            output_doc.add_page_break()

        p_title = output_doc.add_paragraph(containertype)
        for run in p_title.runs:
            run.font.name = 'Arial'
            run.bold = True
        p_title.style.font.size = Pt(12)
        p_title.paragraph_format.space_before = Pt(0)
        p_title.paragraph_format.space_after = Pt(4)

        p_img   = output_doc.add_paragraph()
        run_img = p_img.add_run()
        run_img.add_picture(barcode_buf, width=Cm(BARCODE_WIDTH), height=Cm(BARCODE_HEIGHT))
        p_img.paragraph_format.space_before = Pt(0)
        p_img.paragraph_format.space_after = Pt(4)

        p_info1 = output_doc.add_paragraph(f"{straat} {huisnummer} {toevoeging}".strip())
        for run in p_info1.runs:
            run.font.name = 'Arial'
            run.bold = True
        p_info1.style.font.size = Pt(12)
        p_info1.paragraph_format.space_before = Pt(0)
        p_info1.paragraph_format.space_after = Pt(2)

        p_info2 = output_doc.add_paragraph(f"{postcode} {woonplaats}")
        for run in p_info2.runs:
            run.font.name = 'Arial'
            run.bold = True
        p_info2.style.font.size = Pt(12)
        p_info2.paragraph_format.space_before = Pt(0)
        p_info2.paragraph_format.space_after = Pt(0)

    docx_buffer = BytesIO()
    output_doc.save(docx_buffer)
    docx_buffer.seek(0)
    return docx_buffer


def get_category_counts(df_raw):
    """Bereken tellingen per categorie en validatiefouten op basis van het originele DataFrame.
    REMOVE-rijen worden gefilterd voor de telling. Overgeslagen = rijen met ontbrekende/ongeldige velden.
    """
    # Filter REMOVE-rijen eerst weg voor de tellingen
    if 'CATEGORYNAME' in df_raw.columns:
        cats_all = df_raw['CATEGORYNAME'].astype(str).str.strip().str.upper()
        df_telling = df_raw[cats_all != 'REMOVE'].copy()
    else:
        df_telling = df_raw.copy()

    # Tellingen op gefilterd DataFrame
    if 'CATEGORYNAME' in df_telling.columns:
        cats = df_telling['CATEGORYNAME'].astype(str).str.strip().str.upper()
        wissel    = int((cats == 'CHANGE').sum())
        uitzetten = int(((cats == 'NEW') | (cats == 'EXTRA')).sum())
    else:
        wissel    = 0
        uitzetten = len(df_telling)

    # Overgeslagen: rijen met ontbrekende/ongeldige velden (op het totale df incl. REMOVE-filter)
    overgeslagen_rows = []
    for idx, row in df_telling.iterrows():
        redenen = []
        containercode = str(row.get('CONTAINERCODE', '')).strip()
        streetname    = str(row.get('STREETNAME', '')).strip()
        zipcode       = str(row.get('ZIPCODE', '')).strip()
        city          = str(row.get('CITY', '')).strip()

        if len(containercode) < 5:
            redenen.append(f"ContainerCode te kort of leeg ('{containercode}')")
        if not streetname:
            redenen.append("StreetName leeg")
        if not zipcode:
            redenen.append("ZipCode leeg")
        if not city:
            redenen.append("City leeg")

        if redenen:
            adres = f"{streetname} {str(row.get('HOUSENUMBER', '')).strip()}".strip()
            overgeslagen_rows.append({
                'rij':       idx + 2,
                'adres':     adres or '—',
                'postcode':  zipcode or '—',
                'container': containercode or '—',
                'reden':     ' · '.join(redenen),
            })

    return {
        'wissel':            wissel,
        'uitzetten':         uitzetten,
        'overgeslagen':      len(overgeslagen_rows),
        'overgeslagen_rows': overgeslagen_rows,
    }


XIMMIO_EXPORT_COLUMNS = {'Stad', 'Straat', 'Huisnummer', 'Postcode', 'SubTaskDesc'}


def is_ximmio_export(df):
    """Detecteer of het bestand een Ximmio bakwagen export is."""
    return XIMMIO_EXPORT_COLUMNS.issubset(set(df.columns))


def parse_subtaskdesc(value):
    """Extraheer CategoryName en ContainerType uit SubTaskDesc.
    - CHANGE: containertype staat na '>' tussen ()
    - NEW/EXTRA: containertype staat tussen () aan het einde
    - REMOVE: geen containertype nodig
    Geeft (category, containertype) terug.
    """
    s = str(value)
    m_cat = re.search(r'-\s*(CHANGE|NEW|EXTRA|REMOVE)\b', s, re.IGNORECASE)
    if not m_cat:
        return None, None
    cat = m_cat.group(1).upper()

    if cat == 'CHANGE':
        # Pak de laatste (...), dan daarin alles na de laatste >
        m_paren = re.search(r'\(([^(]+)\)\s*$', s)
        inner = m_paren.group(1) if m_paren else s
        m_arrow = re.search(r'>\s*(\S+)\s*$', inner)
        container = m_arrow.group(1).strip() if m_arrow else None
    elif cat in ('NEW', 'EXTRA'):
        m = re.search(r'\(([^)]+)\)\s*$', s)
        container = m.group(1).strip() if m else None
    else:
        container = None

    return cat, container


def dataframe_from_ximmio_export(df):
    """Map Ximmio bakwagen export naar intern DataFrame formaat."""
    rows = []
    for _, row in df.iterrows():
        subtask = str(row.get('SubTaskDesc', ''))
        cat, container = parse_subtaskdesc(subtask)

        if cat == 'REMOVE':
            continue  # Overslaan

        hl_raw     = row.get('Huisletter', '')
        tv_raw     = row.get('Huisnummer toevoeging', '')
        huisletter = '' if pd.isna(hl_raw) else str(hl_raw).strip()
        toevoeging = '' if pd.isna(tv_raw) else str(tv_raw).strip()
        # Underscore in toevoeging behouden zoals het is (Ximmio gebruikt dit als separator)

        rows.append({
            'containertype': strip_spaces(container or ''),
            'straat':        str(row.get('Straat', '')).strip(),
            'huisnummer':    str(row.get('Huisnummer', '')).strip(),
            'toevoeging':    (huisletter + toevoeging).strip(),
            'postcode':      strip_spaces(str(row.get('Postcode', ''))),
            'woonplaats':    str(row.get('Stad', '')).strip(),
            '_cat':          cat or '',
            '_zipcode_raw':  str(row.get('Postcode', '')).strip(),
            '_huisnummer_raw': str(row.get('Huisnummer', '')).strip(),
            '_huisletter_raw': huisletter,
            '_toevoeging_raw': toevoeging,
        })

    result_df = pd.DataFrame(rows) if rows else pd.DataFrame()
    return result_df


def dataframe_from_file(file):
    """Lees CSV/XLSX/CSV en detecteer automatisch het formaat.
    Geeft een tuple terug: (df, category_counts)
    """
    df_raw = pd.read_excel(file) if file.name.endswith(".xlsx") else pd.read_csv(file)

    if is_ximmio_export(df_raw):
        # ── Ximmio bakwagen export ──────────────────────────────
        # Tellingen uit SubTaskDesc voor de rapportage
        cats_series = df_raw['SubTaskDesc'].apply(lambda v: parse_subtaskdesc(v)[0])
        cats_upper  = cats_series.fillna('').str.upper()
        overgeslagen_rows = []
        for idx, row in df_raw.iterrows():
            cat, container = parse_subtaskdesc(str(row.get('SubTaskDesc', '')))
            if cat == 'REMOVE':
                continue
            def leeg(val):
                """Geeft True als de waarde leeg, NaN of de string 'nan' is."""
                import pandas as pd
                if val is None:
                    return True
                if isinstance(val, float) and pd.isna(val):
                    return True
                return str(val).strip().lower() in ('', 'nan', 'none')

            redenen = []
            containercode = strip_spaces(container or '')
            streetname    = row.get('Straat', '')
            zipcode_raw   = row.get('Postcode', '')
            city_raw      = row.get('Stad', '')
            huisnummer_v  = row.get('Huisnummer', '')
            subtask_v     = row.get('SubTaskDesc', '')

            streetname_s  = '' if leeg(streetname)  else str(streetname).strip()
            zipcode_s     = '' if leeg(zipcode_raw)  else strip_spaces(str(zipcode_raw))
            city_s        = '' if leeg(city_raw)     else str(city_raw).strip()
            huisnummer_s  = '' if leeg(huisnummer_v) else str(huisnummer_v).strip()
            subtask_s     = '' if leeg(subtask_v)    else str(subtask_v).strip()

            if len(containercode) < 5:
                redenen.append(f"ContainerCode te kort of leeg ('{containercode}')")
            if not streetname_s:
                redenen.append("Straat leeg")
            if not zipcode_s:
                redenen.append("Postcode leeg")
            if not city_s:
                redenen.append("Stad leeg")
            if not huisnummer_s:
                redenen.append("Huisnummer leeg")
            if not subtask_s:
                redenen.append("SubTaskDesc leeg")
            if redenen:
                huisnummer = str(row.get('Huisnummer', '')).strip()
                overgeslagen_rows.append({
                    'rij':       idx + 2,
                    'adres':     f"{streetname_s} {huisnummer_s}".strip() or '—',
                    'postcode':  zipcode_s or '—',
                    'container': containercode or '—',
                    'reden':     ' · '.join(redenen),
                })

        counts = {
            'wissel':            int((cats_upper == 'CHANGE').sum()),
            'uitzetten':         int(((cats_upper == 'NEW') | (cats_upper == 'EXTRA')).sum()),
            'overgeslagen':      len(overgeslagen_rows),
            'overgeslagen_rows': overgeslagen_rows,
        }

        result_df = dataframe_from_ximmio_export(df_raw)

        # Sorteer oplopend
        if not result_df.empty:
            result_df['_hn_int'] = pd.to_numeric(result_df['_huisnummer_raw'], errors='coerce').fillna(0).astype(int)
            result_df = result_df.sort_values(
                by=['_zipcode_raw', '_hn_int', '_huisletter_raw', '_toevoeging_raw'],
                ascending=True, na_position='last'
            ).reset_index(drop=True)
            result_df = result_df.drop(columns=['_cat', '_zipcode_raw', '_huisnummer_raw',
                                                 '_huisletter_raw', '_toevoeging_raw', '_hn_int'])

        return result_df, counts

    else:
        # ── Standaard Ximmio CSV/XLSX export ───────────────────
        df = df_raw.copy()
        df.columns = df.columns.str.upper()

        counts = get_category_counts(df)

        # Sorteer oplopend op ZipCode, HouseNumber, HouseLetter, HouseNumberAddition
        sort_cols = []
        for col in ['ZIPCODE', 'HOUSENUMBER', 'HOUSELETTER', 'HOUSENUMBERADDITION']:
            if col in df.columns:
                sort_cols.append(col)
        if sort_cols:
            df['HOUSENUMBER'] = pd.to_numeric(df['HOUSENUMBER'], errors='coerce').fillna(0).astype(int)
            df = df.sort_values(by=sort_cols, ascending=True, na_position='last').reset_index(drop=True)

        if 'CATEGORYNAME' in df.columns:
            df = df[df['CATEGORYNAME'].astype(str).str.strip().str.upper() != 'REMOVE']

        houseletter          = df['HOUSELETTER'].fillna('').astype(str).str.strip()
        housenumber_addition = df['HOUSENUMBERADDITION'].fillna('').astype(str).str.strip()

        result_df = pd.DataFrame({
            'containertype': df['CONTAINERCODE'].apply(strip_spaces),
            'straat':        df['STREETNAME'].astype(str),
            'huisnummer':    df['HOUSENUMBER'].astype(str),
            'toevoeging':    (houseletter + housenumber_addition).str.strip(),
            'postcode':      df['ZIPCODE'].apply(strip_spaces),
            'woonplaats':    df['CITY'].astype(str),
        })

        return result_df, counts


# -------------------------------------------------------
# Streamlit UI
# -------------------------------------------------------

st.set_page_config(page_title="Labelgenerator", page_icon="📦")
st.title("📦 Containerlabelgenerator")

tab_xlsx, tab_manual = st.tabs(["📂 XLSX uploaden", "✏️ Handmatig invoeren"])

# ── Tab 1: XLSX upload ─────────────────────────────────
with tab_xlsx:
    st.write("Upload een XLSX bestand om labels te genereren met barcodes.")

    uploaded_file = st.file_uploader("Sleep je .xlsx bestand hiernaartoe", type=["xlsx"])

    if uploaded_file:
        if st.button("Verwerken", key="btn_xlsx"):
            with st.spinner("Bezig met verwerken..."):
                df, counts = dataframe_from_file(uploaded_file)
                docx_file = generate_word_from_dataframe(df)
                timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M")
                st.success(f"✅ {len(df)} label(s) gegenereerd!")

                col1, col2, col3 = st.columns(3)
                col1.metric("🔄 Wissel",       counts['wissel'],      help="CategoryName = CHANGE")
                col2.metric("📦 Uitzetten",    counts['uitzetten'],   help="CategoryName = NEW of EXTRA")
                col3.metric("⛔ Overgeslagen", counts['overgeslagen'], help="Ontbrekende/ongeldige velden")

                if counts['overgeslagen_rows']:
                    with st.expander(f"⛔ {counts['overgeslagen']} overgeslagen rij(en) — klik om te bekijken"):
                        overgeslagen_df = pd.DataFrame(counts['overgeslagen_rows'])
                        overgeslagen_df.columns = ['Rij', 'Adres', 'Postcode', 'Containertype', 'Reden']
                        st.dataframe(overgeslagen_df, hide_index=True, use_container_width=True)

                st.download_button(
                    label="📥 Download Word-bestand",
                    data=docx_file,
                    file_name=f"containerlabels_{timestamp}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key="dl_xlsx"
                )

# ── Tab 2: Handmatig invoeren ──────────────────────────
with tab_manual:
    st.write(f"Vul hieronder handmatig de gegevens in (maximaal {MAX_ROWS} labels).")

    if 'num_rows' not in st.session_state:
        st.session_state.num_rows = 1

    col_add, col_remove = st.columns([1, 1])
    with col_add:
        if st.button("➕ Rij toevoegen", disabled=st.session_state.num_rows >= MAX_ROWS):
            st.session_state.num_rows += 1
    with col_remove:
        if st.button("➖ Rij verwijderen", disabled=st.session_state.num_rows <= 1):
            st.session_state.num_rows -= 1

    st.markdown("---")

    rows = []
    for i in range(st.session_state.num_rows):
        st.markdown(f"**Label {i + 1}**")
        c1, c2, c3, c4, c5, c6 = st.columns([2, 2, 1, 1, 1.5, 2])
        with c1:
            containertype = st.text_input("Containertype", key=f"ct_{i}", placeholder="OPK_140L")
        with c2:
            straat = st.text_input("Straat", key=f"st_{i}", placeholder="Teststraat")
        with c3:
            huisnummer = st.text_input("Nr.", key=f"hn_{i}", placeholder="9")
        with c4:
            toevoeging = st.text_input("Toev.", key=f"tv_{i}", placeholder="A")
        with c5:
            postcode = st.text_input("Postcode", key=f"pc_{i}", placeholder="1234AA")
        with c6:
            woonplaats = st.text_input("Woonplaats", key=f"wp_{i}", placeholder="Rijswijk")

        rows.append({
            'containertype': strip_spaces(containertype),
            'straat':        straat.strip(),
            'huisnummer':    huisnummer.strip(),
            'toevoeging':    toevoeging.strip(),
            'postcode':      strip_spaces(postcode),
            'woonplaats':    woonplaats.strip(),
        })

    st.markdown("---")

    if st.button("Verwerken", key="btn_manual"):
        df_manual = pd.DataFrame(rows)
        # Alleen rijen met minimaal postcode + huisnummer
        df_manual = df_manual[
            (df_manual['postcode'] != '') & (df_manual['huisnummer'] != '')
        ].reset_index(drop=True)

        if df_manual.empty:
            st.warning("Vul minimaal postcode en huisnummer in voor één label.")
        else:
            with st.spinner("Bezig met verwerken..."):
                docx_file = generate_word_from_dataframe(df_manual)
                timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M")
                st.success(f"{len(df_manual)} label(s) gegenereerd!")
                st.download_button(
                    label="📥 Download Word-bestand",
                    data=docx_file,
                    file_name=f"containerlabels_{timestamp}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key="dl_manual"
                )