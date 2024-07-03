import streamlit as st
import os
import pandas as pd
from io import BytesIO
from PyPDF2 import PdfReader
import camelot

def pdf_to_xlsx(pdf_path, xlsx_path, num_columns, progress_bar):
    all_data = []
    tables = camelot.read_pdf(pdf_path, pages='all', strip_text='\n', flavor='stream')
    num_pages = len(tables)

    for i, table in enumerate(tables):
        df = table.df
        if df.shape[1] < num_columns:
            df = df.reindex(columns=range(num_columns), fill_value="")
        elif df.shape[1] > num_columns:
            df = df.iloc[:, :num_columns]
        all_data.append(df)
        progress_bar.progress((i + 1) / num_pages)

    if all_data:
        df = pd.concat(all_data, ignore_index=True)
        df.columns = [f'Col{i+1}' for i in range(num_columns)]
        df.to_excel(xlsx_path, index=False, sheet_name='Sheet1')
    else:
        raise ValueError("Nenhuma tabela foi extraída do PDF.")

def main():
    st.title("Conversor de PDF para Excel")
    
    uploaded_file = st.file_uploader("Escolha um arquivo PDF", type="pdf")
    
    if uploaded_file is not None:
        num_columns = st.number_input("Digite o número de colunas para a conversão:", min_value=1, value=8)
        
        if st.button("Iniciar Conversão"):
            progress_bar = st.progress(0)
            
            pdf_path = "temp_pdf.pdf"
            xlsx_path = "output_excel.xlsx"
            
            with open(pdf_path, "wb") as f:
                f.write(uploaded_file.getbuffer())
            
            try:
                pdf_to_xlsx(pdf_path, xlsx_path, num_columns, progress_bar)
                
                st.markdown("### Download do arquivo Excel gerado")
                with open(xlsx_path, "rb") as f:
                    bytes_data = f.read()
                    st.download_button(label="Baixar Excel", data=BytesIO(bytes_data), file_name="output_excel.xlsx", key="download_button")
            
            except ValueError as e:
                st.error(str(e))
            finally:
                if os.path.exists(pdf_path):
                    os.remove(pdf_path)
                if os.path.exists(xlsx_path):
                    os.remove(xlsx_path)

if __name__ == "__main__":
    main()
