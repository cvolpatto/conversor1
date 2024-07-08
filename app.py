import streamlit as st
import pandas as pd
import fitz  # PyMuPDF
from io import BytesIO
import os
import tabula

def extract_tables_from_pdf(pdf_path):
    # Use tabula to extract tables from the PDF
    tables = tabula.read_pdf(pdf_path, pages='all', multiple_tables=True)
    return tables

def merge_dataframes(dataframes):
    # Concatenate all dataframes into one
    full_df = pd.concat(dataframes, ignore_index=True)
    return full_df

def pdf_to_xlsx(pdf_path, xlsx_path, progress_bar):
    tables = extract_tables_from_pdf(pdf_path)
    if not tables:
        raise ValueError("Nenhuma tabela foi extraída do PDF.")
    
    all_dataframes = []
    for i, df in enumerate(tables):
        all_dataframes.append(df)
        progress_bar.progress((i + 1) / len(tables))

    full_df = merge_dataframes(all_dataframes)
    full_df.to_excel(xlsx_path, index=False, sheet_name='Sheet1')

def main():
    st.title("Conversor de PDF para Excel")
    
    uploaded_file = st.file_uploader("Escolha um arquivo PDF", type="pdf")
    
    if uploaded_file is not None:
        if st.button("Iniciar Conversão"):
            progress_bar = st.progress(0)
            
            pdf_path = "temp_pdf.pdf"
            xlsx_path = "output_excel.xlsx"
            
            with open(pdf_path, "wb") as f:
                f.write(uploaded_file.getbuffer())
            
            try:
                pdf_to_xlsx(pdf_path, xlsx_path, progress_bar)
                
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
