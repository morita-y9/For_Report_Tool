import pdfplumber

from docx import Document
import spacy
import streamlit as st

def extract_pdf_elements(pdf_path):
    elements = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text(x_tolerance=1)
            tables = page.extract_tables()
            images = page.images
            elements.append({
                "text": text,
                "tables": tables,
                "images": images
            })
    return elements

def write_to_word(elements, output_path):
    # 表紙の挿入
    doc = Document("report_format.docx")

    for page_num, content in enumerate(elements):

        if content['text']:
            for line in content['text'].split('\n'):
                doc.add_paragraph(line)

        for table in content['tables']:
            rows, cols = len(table), len(table[0])
            t = doc.add_table(rows=rows, cols=cols)
            for i, row in enumerate(table):
                for j, cell in enumerate(row):
                    t.cell(i, j).text = cell or ''

    doc.save(output_path)
    
nlp = spacy.load("en_core_web_sm")


st.title("word用実験レポート作成ツール")

uploaded_file = st.file_uploader("PDFファイルをアップロード", type=["pdf"])

if uploaded_file:
    st.success("PDFファイルが正常に読み込まれました！")
    elements = extract_pdf_elements(uploaded_file)

    st.write("抽出されたページ数:", len(elements))
    
    if st.button("Wordに変換"):
        output_path = "output.docx"
        write_to_word(elements, output_path)
        with open(output_path, "rb") as f:
            st.download_button("Wordファイルをダウンロードする", f, file_name="converted.docx")

