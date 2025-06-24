import streamlit as st
import pandas as pd
import datetime
import os
from docx import Document
from docx.shared import Cm, Pt
from pystrich.code128 import Code128Encoder
from PIL import Image, ImageDraw, ImageFont
from io import BytesIO, StringIO


def generate_word_from_excel(file, barcode_text_size=12, barcode_width_cm=4, barcode_height_cm=None):
    df = pd.read_excel(file)

    output_doc = Document()
    section = output_doc.sections[-1]
    section.page_width = Cm(10.0)
    section.page_height = Cm(8.0)
    section.left_margin = Cm(0.8)
    section.right_margin = Cm(0.5)
    section.top_margin = Cm(2)
    section.bottom_margin = Cm(0.5)

    try:
        font = ImageFont.truetype("arial.ttf", barcode_text_size)
    except IOError:
        font = ImageFont.load_default()

    dpi = 96
    pixels_per_cm = dpi / 2.54
    horizontal_shift = int(0.75 * pixels_per_cm)
    text_area_height = 30

    for idx, row in df.iterrows():
        containertype = str(row.get('containertype', ''))
        straat = str(row.get('straat', ''))
        huisnummer = str(row.get('huisnummer', ''))
        toevoeging_value = row.get('toevoeging', '')
        toevoeging = '' if pd.isna(toevoeging_value) else str(toevoeging_value)
        postcode = str(row.get('postcode', ''))
        woonplaats = str(row.get('woonplaats', ''))
        barcode_value = f"{postcode}{huisnummer}{toevoeging}"

        # Barcode genereren
        encoder = Code128Encoder(barcode_value)
        barcode_img = encoder.get_imagedata()
        barcode_image = Image.open(BytesIO(barcode_img))

        # Crop whitespace
        bbox = barcode_image.getbbox()
        if bbox:
            left, _, right, _ = bbox
            barcode_image = barcode_image.crop((left, 0, right, barcode_image.height))

        # Tekstvlak
        draw = ImageDraw.Draw(barcode_image)
        width, height = barcode_image.size
        draw.rectangle([0, height - text_area_height, width, height], fill="white")

        text = ""
        bbox_text = draw.textbbox((0, 0), text, font=font)
        text_y = height - text_area_height + ((text_area_height - (bbox_text[3] - bbox_text[1])) / 2)
        draw.text((horizontal_shift, text_y), text, fill="black", font=font)

        # Barcode tijdelijk opslaan in memory
        barcode_buf = BytesIO()
        barcode_image.save(barcode_buf, format="PNG")
        barcode_buf.seek(0)

        # Voeg toe aan Word
        p_title = output_doc.add_paragraph(containertype)
        for run in p_title.runs:
            run.font.name = 'Arial'
            run.bold = True
        p_title.style.font.size = Pt(8)

        p_img = output_doc.add_paragraph()
        run_img = p_img.add_run()
        if barcode_height_cm:
            run_img.add_picture(barcode_buf, width=Cm(barcode_width_cm), height=Cm(barcode_height_cm))
        else:
            run_img.add_picture(barcode_buf, width=Cm(barcode_width_cm))

        p_info1 = output_doc.add_paragraph(f"{straat} {huisnummer} {toevoeging}")
        for run in p_info1.runs:
            run.font.name = 'Arial'
            run.bold = True
        p_info1.style.font.size = Pt(10)

        p_info2 = output_doc.add_paragraph(f"{postcode} {woonplaats}")
        for run in p_info2.runs:
            run.font.name = 'Arial'
            run.bold = True
        p_info2.style.font.size = Pt(10)

        output_doc.add_page_break()

    # Document opslaan in memory
    docx_buffer = BytesIO()
    output_doc.save(docx_buffer)
    docx_buffer.seek(0)
    return docx_buffer


# -----------------------
# Streamlit UI
# -----------------------
st.set_page_config(page_title="Labelgenerator", page_icon="📦")
st.title("📦 Containerlabelgenerator")
st.write("Upload een Excel-bestand om labels te genereren met barcodes.")

uploaded_file = st.file_uploader("Sleep je `.xlsx` bestand hiernaartoe", type=["xlsx"])

if uploaded_file:
    barcode_width = st.slider("Barcode breedte (cm)", 2.0, 6.0, 3.0, 0.1)
    barcode_height = st.slider("Barcode hoogte (cm)", 0.5, 3.0, 1.5, 0.1)

    if st.button("Verwerken"):
        with st.spinner("Bezig met verwerken..."):
            docx_file = generate_word_from_excel(
                uploaded_file,
                barcode_text_size=18,
                barcode_width_cm=barcode_width,
                barcode_height_cm=barcode_height,
            )
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M")
            filename = f"containerlabels_{timestamp}.docx"

            st.success("Labels gegenereerd!")
            st.download_button(
                label="📥 Download Word-bestand",
                data=docx_file,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
