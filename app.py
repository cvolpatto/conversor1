import streamlit as st
import os
import camelot
import pandas as pd
from io import BytesIO
import logging

# Configurar o logger
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def pdf_to_xlsx(pdf_path, xlsx_path, num_columns, progress_bar):
    logging.info(f"Iniciando a extração de tabelas do arquivo PDF: {pdf_path}")

    try:
        # Extrair dados da tabela do PDF usando Camelot
        tables = camelot.read_pdf(pdf_path, flavor='stream', pages='all')
        logging.info(f"Número total de tabelas extraídas: {len(tables)}")

        all_data = []
        for i, table in enumerate(tables):
            logging.info(f"Processando tabela {i+1} de {len(tables)}.")
            df = table.df
            logging.debug(f"Tabela original extraída com {df.shape[0]} linhas e {df.shape[1]} colunas.")

            # Verificar e ajustar o número de colunas do DataFrame
            if df.shape[1] < num_columns:
                df = df.reindex(columns=range(num_columns), fill_value="")
                logging.debug(f"Tabela ajustada para {num_columns} colunas (completada com valores vazios).")
            elif df.shape[1] > num_columns:
                df = df.iloc[:, :num_columns]
                logging.debug(f"Tabela ajustada para {num_columns} colunas (colunas extras removidas).")

            all_data.append(df)
            logging.info(f"Tabela {i+1} com {df.shape[0]} linhas extraída e ajustada.")
            progress_bar.progress((i + 1) / len(tables))
            logging.info(f"Barra de progresso atualizada para {(i + 1) / len(tables):.2%}")

        if all_data:
            # Concatenar todos os DataFrames em um único DataFrame
            df = pd.concat(all_data, ignore_index=True)
            logging.info(f"Todas as tabelas concatenadas em um único DataFrame com {df.shape[0]} linhas no total.")
            
            # Salvar o DataFrame concatenado em um arquivo Excel
            df.to_excel(xlsx_path, index=False, sheet_name='Sheet1')
            logging.info(f"Dados salvos no arquivo {xlsx_path}.")
        else:
            logging.error("Nenhuma tabela foi extraída do PDF.")
            raise ValueError("Nenhuma tabela foi extraída do PDF.")

    except Exception as e:
        logging.exception("Ocorreu um erro durante a extração e conversão das tabelas.")
        raise e

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
