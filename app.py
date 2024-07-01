import streamlit as st
import os
import camelot
import pandas as pd
from io import BytesIO
import time

os.environ['OPENCV_IO_MAX_IMAGE_PIXELS'] = '2200000000'  # ou outro valor adequado para suas imagens

def pdf_to_xlsx(pdf_path, xlsx_path, num_columns, progress_bar):
    # Extrair dados da tabela do PDF usando camelot
    tables = camelot.read_pdf(pdf_path, flavor='stream', pages='all')

    # Atualizar a barra de progresso
    progress_bar.progress(30)

    # Concatenar os DataFrames resultantes
    df = pd.concat([table.df for table in tables], ignore_index=True)

    # Atualizar a barra de progresso
    progress_bar.progress(60)

    # Ajustar número de colunas de acordo com o especificado
    df = df.iloc[:, :num_columns]

    # Atualizar a barra de progresso
    progress_bar.progress(90)

    # Salvar os dados em um arquivo XLSX
    df.to_excel(xlsx_path, index=False, sheet_name='Sheet1')

    # Atualizar a barra de progresso para 100%
    progress_bar.progress(100)

def main():
    st.title("Conversor de PDF para Excel")

    # Upload do arquivo PDF
    uploaded_file = st.file_uploader("Escolha um arquivo PDF", type="pdf")

    if uploaded_file is not None:
        # Entrada para o número de colunas
        num_columns = st.number_input("Digite o número de colunas para a conversão", min_value=1, value=8)

        # Botão para iniciar a conversão
        if st.button("Iniciar Conversão"):
            # Salvar o arquivo temporariamente
            pdf_path = "temp_pdf.pdf"
            xlsx_path = "output_excel.xlsx"

            with open(pdf_path, "wb") as f:
                f.write(uploaded_file.getbuffer())

            # Barra de progresso
            progress_bar = st.progress(0)

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