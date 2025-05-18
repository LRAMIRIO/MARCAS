import streamlit as st
import pandas as pd
import openpyxl
import re
import tempfile
from io import BytesIO

def extrair_marcas_convertidas(files):
    marcas_por_item = {}

    for nome_arquivo, file in files.items():
        if not nome_arquivo.endswith(".xlsx") or "convertido" not in nome_arquivo:
            continue

        try:
            df = pd.read_excel(file)
            linhas = df['Linha'].astype(str).tolist()

            # Pegar número do item no nome do arquivo
            item_match = re.search(r"item-(\d+)", nome_arquivo)
            item = int(item_match.group(1)) if item_match else None

            for i in range(len(linhas) - 2):
                linha1 = linhas[i].lower()
                linha2 = linhas[i+1].lower()
                linha3 = linhas[i+2].strip()

                if "(total)" in linha1 and "fornecedor" in linha1 and "habilitado" in linha2:
                    if "marca/fabricante:" in linha3.lower():
                        marca = linha3.split(":", 1)[-1].strip()
                        if item:
                            marcas_por_item[item] = marca
                        break

        except Exception as e:
            st.warning(f"Erro ao processar {nome_arquivo}: {e}")

    return marcas_por_item

def preencher_planilha_modelo(modelo_file, marcas_dict):
    wb = openpyxl.load_workbook(modelo_file)
    ws = wb.active

    for row in ws.iter_rows(min_row=4):
        cell_item = row[0].value  # coluna A
        if cell_item:
            item = int(str(cell_item).replace(".", "").strip())
            if item in marcas_dict:
                row[3].value = marcas_dict[item]  # coluna D: MARCA

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# Streamlit App
st.title("Preenchimento Automático de Marcas por Item")
st.markdown("Envie os arquivos `.xlsx` convertidos dos PDFs (com 'item-x_convertido.xlsx') e a planilha modelo preenchida.")

uploaded_files = st.file_uploader("Envie os arquivos convertidos e a planilha modelo:", accept_multiple_files=True, type="xlsx")

if uploaded_files:
    modelo_file = None
    arquivos_dict = {}

    for file in uploaded_files:
        if "modelo" in file.name.lower():
            modelo_file = file
        else:
            arquivos_dict[file.name] = file

    if modelo_file and arquivos_dict:
        marcas = extrair_marcas_convertidas(arquivos_dict)
        output_excel = preencher_planilha_modelo(modelo_file, marcas)
        st.success("Planilha preenchida com sucesso!")

        st.download_button(
            label="📄 Baixar planilha com marcas preenchidas",
            data=output_excel,
            file_name="Planilha_com_marcas_preenchidas_FINAL.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("✉️ Envie pelo menos uma planilha modelo e arquivos convertidos.")
