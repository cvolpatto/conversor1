import streamlit as st
import os
import tabula
import pandas as pd
from io import BytesIO
import time
import camelot.io as camelot

def pdf_to_xlsx(pdf_path, xlsx_path, num_columns, progress_bar):
    # Extrair dados da tabela do PDF usando tabula
    tables = tabula.read_pdf(pdf_path, pages='all', multiple_tables=True)

    # Concatenar os DataFrames resultantes
    df = pd.concat([table for table in tables], ignore_index=True)

    # Ajustar número de colunas
    df = df.iloc[:, :num_columns]

    # Salvar os dados em um arquivo XLSX
    df.to_excel(xlsx_path, index=False, sheet_name='Sheet1')

    # Atualizar barra de progresso
    progress_bar.progress(100)

def main():
    st.title("Conversor de PDF para Excel")

    # Upload do arquivo PDF
    uploaded_file = st.file_uploader("Escolha um arquivo PDF", type="pdf")

    if uploaded_file is not None:
        # Input do número de colunas
        num_columns = st.number_input("Digite o número de colunas para a conversão:", min_value=1, value=8)

        # Botão para iniciar a conversão
        if st.button("Iniciar Conversão"):
            # Barra de progresso
            progress_bar = st.progress(0)
            
            # Salvar o arquivo temporariamente
            pdf_path = "temp_pdf.pdf"
            xlsx_path = "output_excel.xlsx"

            with open(pdf_path, "wb") as f:
                f.write(uploaded_file.getbuffer())

            # Converter PDF para Excel
            pdf_to_xlsx(pdf_path, xlsx_path, num_columns, progress_bar)

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
