import streamlit as st
import os
import camelot
import pandas as pd
from io import BytesIO
import os

os.environ['OPENCV_IO_MAX_IMAGE_PIXELS'] = '2200000000'  # ou outro valor adequado para suas imagens


def pdf_to_xlsx(pdf_path, xlsx_path):
    # Extrair dados da tabela do PDF usando camelot
    tables = camelot.read_pdf(pdf_path, flavor='stream', pages='all')

    # Concatenar os DataFrames resultantes
    df = pd.concat([table.df for table in tables], ignore_index=True)

    # Ajustar número de colunas para 6
    df = df.iloc[:, :6]

    # Salvar os dados em um arquivo XLSX
    df.to_excel(xlsx_path, index=False, sheet_name='Sheet1')

def main():
    st.title("Conversor de PDF para Excel")

    # Upload do arquivo PDF
    uploaded_file = st.file_uploader("Escolha um arquivo PDF", type="pdf")

    if uploaded_file is not None:
        # Salvar o arquivo temporariamente
        pdf_path = "temp_pdf.pdf"
        xlsx_path = "output_excel.xlsx"

        with open(pdf_path, "wb") as f:
            f.write(uploaded_file.getbuffer())

        # Converter PDF para Excel
        pdf_to_xlsx(pdf_path, xlsx_path)

        # Download do arquivo Excel gerado
        st.markdown("### Download do arquivo Excel gerado")
        with open(xlsx_path, "rb") as f:
            bytes_data = f.read()
            st.download_button(label="Baixar Excel", data=BytesIO(bytes_data), file_name="output_excel.xlsx", key="download_button")

        # Limpar arquivos temporários
        os.remove(pdf_path)
        os.remove(xlsx_path)

if __name__ == "__main__":
    main()
