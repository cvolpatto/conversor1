import streamlit as st
import pandas as pd
import fitz  # PyMuPDF
from io import BytesIO
import os

def extract_tables_from_pdf(pdf_path):
    pdf_document = fitz.open(pdf_path)
    all_text = []
    for page_num in range(len(pdf_document)):
        page = pdf_document.load_page(page_num)
        text = page.get_text("text")
        all_text.append(text)
    return all_text

def text_to_dataframe(text, num_columns):
    rows = text.split('\n')
    data = []
    for row in rows:
        columns = row.split()
        data.append(columns)

    # Convert list of lists to DataFrame
    df = pd.DataFrame(data)
    
    # Adjust DataFrame to have the specified number of columns
    df = df.apply(lambda x: x.str.split(expand=True).stack()).unstack()
    df = df.iloc[:, :num_columns]
    df.columns = [f'Col{i+1}' for i in range(num_columns)]
    
    return df

def pdf_to_xlsx(pdf_path, xlsx_path, num_columns, progress_bar):
    all_text = extract_tables_from_pdf(pdf_path)
    all_dataframes = []
    for i, text in enumerate(all_text):
        df = text_to_dataframe(text, num_columns)
        all_dataframes.append(df)
        progress_bar.progress((i + 1) / len(all_text))

    if all_dataframes:
        full_df = pd.concat(all_dataframes, ignore_index=True)
        full_df.to_excel(xlsx_path, index=False, sheet_name='Sheet1')
    else:
        raise ValueError("Nenhum texto foi extraído do PDF.")

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
