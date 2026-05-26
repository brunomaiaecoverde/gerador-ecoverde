import streamlit as st
from docxtpl import DocxTemplate
import requests
import re
import subprocess
import os
import datetime

# LINKS DOS SEUS MODELOS NO GOOGLE DRIVE (Configurados para exportação .docx)
# Certifique-se de que estes links sejam de "Qualquer pessoa com o link pode ler"
URL_PROCURACAO = "https://docs.google.com/document/d/1LUtJrayUBGobMQ1jy2QFIeL0WoflNXNe/export?format=docx"
URL_DISTRATO = "https://docs.google.com/document/d/1Tf6rRbCnBTBcWFr-pVDw7sBV0yZRqeqa/export?format=docx"

# 1. BASE DE DADOS
RESPONSAVEIS_DB = {
    "Gabriela Freitas": {"nome": "Gabriela Ferreira de Freitas", "info": ", brasileira, solteira, líder de projetos júnior 2, CPF 121.142.656-41"},
    "Bruno Maia": {"nome": "Bruno Maia", "info": ", brasileiro, solteiro, estagiário de projetos viários, CPF: 142.172.726-93"},
    "Gustavo Costa": {"nome": "Gustavo de Souza Costa", "info": ", brasileiro, solteiro, Analista de Projetos Viários Júnior II, CPF 105.159.886-93"}
}

SERVICOS_DB = {
    "PSCIP (Incêndio)": "Projeto de segurança e combate contra incêndio e pânico perante ao órgão competente;",
    "Aprovação Arquitetônica": "Processo de aprovação de projetos arquitetônicos perante ao órgão competente;",
    "RIC (Trânsito)": "Elaboração e aprovação de Relatório de Impacto na Circulação, perante ao órgão competente;"
}

# 2. CONFIGURAÇÃO DA PÁGINA
st.set_page_config(page_title="Gerador Ecoverde", page_icon="📄", layout="centered")
st.title("📄 Sistema de Documentos Ecoverde")

# 3. INTERFACE DE SELEÇÃO
tipo_doc = st.sidebar.selectbox("Selecione o Documento:", ["Procuração", "Termo de Distrato"])

# 4. CAPTURA DE DADOS (Notion)
query_params = st.query_params
empresa = st.text_input("Cliente / Razão Social", value=query_params.get("empresa", ""))
cnpj = st.text_input("CNPJ", value=query_params.get("cnpj", ""))
end_cont = st.text_input("Endereço do Contratante", value=query_params.get("endereco_contratante", ""))
end_obra = st.text_input("Endereço da Obra", value=query_params.get("endereco_obra", ""))

if not end_obra: end_obra = end_cont

# Campos específicos por doc
if tipo_doc == "Procuração":
    responsaveis = st.multiselect("Funcionários:", list(RESPONSAVEIS_DB.keys()))
    servicos = st.multiselect("Serviços:", list(SERVICOS_DB.keys()))
    url_modelo = URL_PROCURACAO
else:
    info_adicional = st.text_input("Dados adicionais do contrato (Manual):")
    url_modelo = URL_DISTRATO

if st.button("🚀 Gerar e Baixar PDF", type="primary"):
    with st.spinner("Processando documento..."):
        try:
            # Baixar modelo
            response = requests.get(url_modelo)
            with open("modelo.docx", "wb") as f: f.write(response.content)
            
            # Preparar dados
            ctx = {
                "empresa": empresa.upper(), "cnpj": cnpj, "endereco": end_cont, 
                "endereco_contratante": end_cont, "endereco_obra": end_obra,
                "data": datetime.date.today().strftime("%d de %B de %Y").replace("May", "maio")
            }
            if tipo_doc == "Procuração":
                lista = [RESPONSAVEIS_DB[k]["nome"] + RESPONSAVEIS_DB[k]["info"] for k in responsaveis]
                ctx["responsaveis"] = "e funcionários: " + "; ".join(lista) + "." if lista else "."
                ctx["servicos"] = "\n".join([f"• {SERVICOS_DB[s]}" for s in servicos])
            else:
                ctx["info_adicional"] = info_adicional

            # Gerar
            doc = DocxTemplate("modelo.docx")
            doc.render(ctx)
            doc.save("temp.docx")
            subprocess.run(["libreoffice", "--headless", "--convert-to", "pdf", "temp.docx"], check=True)
            
            with open("temp.pdf", "rb") as f:
                st.download_button(f"📥 Baixar {tipo_doc}.pdf", f, file_name=f"{tipo_doc}_{empresa}.pdf")
        except Exception as e:
            st.error(f"Erro: {e}")
