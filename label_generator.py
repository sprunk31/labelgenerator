import streamlit as st
import pandas as pd
import datetime
from docx import Document
from docx.shared import Cm, Pt
from pystrich.code128 import Code128Encoder
from PIL import Image, ImageDraw, ImageFont
from io import BytesIO

BARCODE_WIDTH     = 3.6
BARCODE_HEIGHT    = 1.9
BARCODE_TEXT_SIZE = 18
MAX_ROWS          = 10

REQUIRED_COLUMNS = ['CONTAINERCODE', 'STREETNAME', 'HOUSENUMBER', 'HOUSELETTER',
                    'HOUSENUMBERADDITION', 'ZIPCODE', 'CITY']


# -------------------------------------------------------
# Hulpfuncties
# -------------------------------------------------------

def strip_spaces(value):
    """Verwijder alle spaties uit een veld."""
    return str(value).replace(" ", "").strip()


def validate_dataframe(df):
    """
    Valideer het DataFrame op ontbrekende kolommen en lege verplichte velden.
    Geeft een lijst van foutmeldingen terug. Lege lijst = geen fouten.
    """
    errors = []

    # Controleer verplichte kolommen
    missing_cols = [col for col in REQUIRED_COLUMNS if col not in df.columns]
    if missing_cols:
        errors.append(f"❌ Ontbrekende kolom(men): **{', '.join(missing_cols)}**")
        return errors  # Verdere validatie heeft geen zin zonder de kolommen

    # Controleer lege ContainerCode per rij
    empty_containercode = df[df['CONTAINERCODE'].astype(str).str.strip() == '']
    if not empty_containercode.empty:
        rijen = ', '.join(str(i + 2) for i in empty_containercode.index.tolist())  # +2: header + 0-index
        errors.append(f"❌ **ContainerCode** is leeg op rij(en): {rijen}")

    # Controleer lege ZipCode per rij
    empty_zipcode = df[df['ZIPCODE'].astype(str).str.strip() == '']
    if not empty_zipcode.empty:
        rijen = ', '.join(str(i + 2) for i in empty_zipcode.index.tolist())
        errors.append(f"❌ **ZipCode** is leeg op rij(en): {rijen}")

    # Controleer lege HouseNumber per rij
    empty_housenumber = df[df['HOUSENUMBER'].astype(str).str.strip() == '']
    if not empty_housenumber.empty:
        rijen = ', '.join(str(i + 2) for i in empty_housenumber.index.tolist())
        errors.append(f"❌ **HouseNumber** is leeg op rij(en): {rijen}")

    # Controleer lege StreetName per rij
    empty_streetname = df[df['STREETNAME'].astype(str).str.strip() == '']
    if not empty_streetname.empty:
        rijen = ', '.join(str(i + 2) for i in empty_streetname.index.tolist())
        errors.append(f"❌ **StreetName** is leeg op rij(en): {rijen}")

    # Controleer lege City per rij
    empty_city = df[df['CITY'].astype(str).str.strip() == '']
    if not empty_city.empty:
        rijen = ', '.join(str(i + 2) for i in empty_city.index.tolist())
        errors.append(f"❌ **City** is leeg op rij(en): {rijen}")

    return errors


def generate_word_from_dataframe(df):
    """Genereer Word-document vanuit een DataFrame met interne kolomnamen."""
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

    for idx, row in df.iterrows():
        containertype = str(row.get('containertype', ''))
        straat        = str(row.get('straat', ''))
        huisnummer    = str(row.get('huisnummer', ''))
        toevoeging    = str(row.get('toevoeging', ''))
        postcode      = str(row.get('postcode', ''))
        woonplaats    = str(row.get('woonplaats', ''))
        barcode_value = f"{postcode}{huisnummer}{toevoeging}"

        try:
            encoder       = Code128Encoder(barcode_value)
            barcode_img   = encoder.get_imagedata()
            barcode_image = Image.open(BytesIO(barcode_img))
        except Exception as e:
            raise ValueError(
                f"Barcode genereren mislukt voor rij {idx + 1} "
                f"(postcode: '{postcode}', huisnummer: '{huisnummer}', toevoeging: '{toevoeging}'): {e}"
            )

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

        p_title = output_doc.add_paragraph(containertype)
        for run in p_title.runs:
            run.font.name = 'Arial'
            run.bold = True
        p_title.style.font.size = Pt(12)
        p_title.paragraph_format.space_after = Pt(6)

        p_img   = output_doc.add_paragraph()
        run_img = p_img.add_run()
        run_img.add_picture(barcode_buf, width=Cm(BARCODE_WIDTH), height=Cm(BARCODE_HEIGHT))
        p_img.paragraph_format.space_after = Pt(6)

        p_info1 = output_doc.add_paragraph(f"{straat} {huisnummer} {toevoeging}".strip())
        for run in p_info1.runs:
            run.font.name = 'Arial'
            run.bold = True
        p_info1.style.font.size = Pt(12)
        p_info1.paragraph_format.space_after = Pt(2)

        p_info2 = output_doc.add_paragraph(f"{postcode} {woonplaats}")
        for run in p_info2.runs:
            run.font.name = 'Arial'
            run.bold = True
        p_info2.style.font.size = Pt(12)
        p_info2.paragraph_format.space_after = Pt(0)

        if idx < len(df) - 1:
            output_doc.add_page_break()

    docx_buffer = BytesIO()
    output_doc.save(docx_buffer)
    docx_buffer.seek(0)
    return docx_buffer


def dataframe_from_file(file):
    """Lees CSV/XLSX en map naar interne kolomnamen. Hoofdletter-onafhankelijk."""
    try:
        df = pd.read_excel(file) if file.name.endswith(".xlsx") else pd.read_csv(file)
    except Exception as e:
        raise ValueError(f"Bestand kon niet worden gelezen: {e}")

    if df.empty:
        raise ValueError("Het bestand is leeg — er zijn geen rijen gevonden.")

    df.columns = df.columns.str.upper()

    # Valideer kolommen en inhoud
    validation_errors = validate_dataframe(df)
    if validation_errors:
        raise ValueError("\n\n".join(validation_errors))

    houseletter          = df['HOUSELETTER'].fillna('').astype(str).str.strip()
    housenumber_addition = df['HOUSENUMBERADDITION'].fillna('').astype(str).str.strip()

    return pd.DataFrame({
        'containertype': df['CONTAINERCODE'].apply(strip_spaces),
        'straat':        df['STREETNAME'].astype(str),
        'huisnummer':    df['HOUSENUMBER'].astype(str),
        'toevoeging':    (houseletter + housenumber_addition).str.strip(),
        'postcode':      df['ZIPCODE'].apply(strip_spaces),
        'woonplaats':    df['CITY'].astype(str),
    })


# -------------------------------------------------------
# Streamlit UI
# -------------------------------------------------------

st.set_page_config(page_title="Labelgenerator", page_icon="📦")
st.title("📦 Containerlabelgenerator")

tab_csv, tab_manual = st.tabs(["📂 CSV/XLSX uploaden", "✏️ Handmatig invoeren"])


# ── Tab 1: CSV upload ──────────────────────────────────
with tab_csv:
    st.write("Upload een CSV of XLSX bestand om labels te genereren met barcodes.")

    st.markdown("#### Vereiste CSV/XLSX headers")
    voorbeeld_df = pd.DataFrame([{
        "ContainerCode": "OPK_140L",
        "StreetName": "Teststraat",
        "HouseNumber": 9,
        "HouseLetter": "A",
        "HouseNumberAddition": "",
        "ZipCode": "1234AA",
        "City": "Rijswijk",
    }])
    st.dataframe(voorbeeld_df, width="stretch", hide_index=True)

    uploaded_file = st.file_uploader("Sleep je .csv of .xlsx bestand hiernaartoe", type=["csv", "xlsx"])

    if uploaded_file:
        if st.button("Verwerken", key="btn_csv"):
            with st.spinner("Bezig met verwerken..."):
                try:
                    df = dataframe_from_file(uploaded_file)
                except ValueError as e:
                    st.error(f"**Fout bij het inlezen van het bestand:**\n\n{e}")
                    st.stop()
                except Exception as e:
                    st.error(f"**Onverwachte fout bij het inlezen:**\n\n{e}")
                    st.stop()

                try:
                    docx_file = generate_word_from_dataframe(df)
                except ValueError as e:
                    st.error(f"**Fout bij het genereren van labels:**\n\n{e}")
                    st.stop()
                except Exception as e:
                    st.error(f"**Onverwachte fout bij het genereren van labels:**\n\n{e}")
                    st.stop()

                timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M")
                st.success(f"✅ {len(df)} label(s) gegenereerd!")
                st.download_button(
                    label="📥 Download Word-bestand",
                    data=docx_file,
                    file_name=f"containerlabels_{timestamp}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key="dl_csv"
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
            containertype = st.text_input("Containertype *", key=f"ct_{i}", placeholder="OPK_140L")
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

        # Validatie: zoek rijen met fouten en rapporteer ze allemaal
        validation_errors = []

        for i, row in df_manual.iterrows():
            label_num = i + 1
            rij_fouten = []

            if not row['containertype']:
                rij_fouten.append("Containertype is verplicht")
            if not row['postcode']:
                rij_fouten.append("Postcode ontbreekt")
            if not row['huisnummer']:
                rij_fouten.append("Huisnummer ontbreekt")

            if rij_fouten:
                validation_errors.append(f"**Label {label_num}:** " + " · ".join(rij_fouten))

        if validation_errors:
            st.error("**Vul de verplichte velden in voordat je verder gaat:**\n\n" + "\n\n".join(validation_errors))
        else:
            # Filter rijen zonder postcode of huisnummer (extra zekerheid)
            df_manual = df_manual[
                (df_manual['postcode'] != '') & (df_manual['huisnummer'] != '')
            ].reset_index(drop=True)

            if df_manual.empty:
                st.warning("Er zijn geen geldige labels om te verwerken.")
            else:
                with st.spinner("Bezig met verwerken..."):
                    try:
                        docx_file = generate_word_from_dataframe(df_manual)
                    except ValueError as e:
                        st.error(f"**Fout bij het genereren van labels:**\n\n{e}")
                        st.stop()
                    except Exception as e:
                        st.error(f"**Onverwachte fout bij het genereren van labels:**\n\n{e}")
                        st.stop()

                    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M")
                    st.success(f"✅ {len(df_manual)} label(s) gegenereerd!")
                    st.download_button(
                        label="📥 Download Word-bestand",
                        data=docx_file,
                        file_name=f"containerlabels_{timestamp}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key="dl_manual"
                    )