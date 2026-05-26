import streamlit as st
from docxtpl import DocxTemplate
import requests
import subprocess
import os
import datetime

# LINKS DOS MODELOS (Google Docs configurados como exportação .docx)
URL_PROCURACAO = "https://docs.google.com/document/d/1LUtJrayUBGobMQ1jy2QFIeL0WoflNXNe/export?format=docx"
URL_DISTRATO = "https://docs.google.com/document/d/1Tf6rRbCnBTBcWFr-pVDw7sBV0yZRqeqa/export?format=docx"

# BASE DE DADOS
RESPONSAVEIS_DB = {
    "Gabriela Freitas": {"nome": "Gabriela Ferreira de Freitas", "info": ", brasileira, solteira, líder de projetos júnior 2, CPF 121.142.656-41"},
    "Bruno Maia": {"nome": "Bruno Maia", "info": ", brasileiro, solteiro, estagiário de projetos viários, CPF: 142.172.726-93"},
    "Gustavo Costa": {"nome": "Gustavo de Souza Costa", "info": ", brasileiro, solteiro, Analista de Projetos Viários Júnior II, CPF 105.159.886-93"},
    "Shérida Silva": {"nome": "Shérida Patrícia Milanez Silva", "info": ", brasileira, casada, Analista de integração de projetos júnior, CPF 072.958.656-16"}
}

SERVICOS_DB = {
    "PSCIP (Incêndio)": "Projeto de segurança e combate contra incêndio e pânico perante ao órgão competente;",
    "Aprovação Arquitetônica": "Processo de aprovação de projetos arquitetônicos perante ao órgão competente;",
    "RIC (Trânsito)": "Elaboração e aprovação de Relatório de Impacto na Circulação, perante ao órgão competente;"
}

# INTERFACE
st.set_page_config(page_title="Sistema Ecoverde", layout="wide")
st.title("📄 Sistema Central de Documentos Ecoverde")

# Cards de Setores
setor = st.radio("Selecione o Setor:", ["Projetos", "Meio Ambiente", "Gestão Ambiental", "Execução"], horizontal=True)

if setor == "Projetos":
    tipo_doc = st.selectbox("Documento:", ["Procuração", "Termo de Distrato", "Termo de Finalização"])
    
    # Inputs (Recebendo os dados com blindagem)
    col1, col2 = st.columns(2)
    with col1:
        empresa = st.text_input("Cliente")
        end_cont = st.text_input("Endereço Contratante")
    with col2:
        cnpj = st.text_input("CNPJ")
        end_obra = st.text_input("Endereço da Obra")
        
    # Lógica de seleção específica
    if tipo_doc == "Procuração":
        resp_sel = st.multiselect("Responsáveis:", list(RESPONSAVEIS_DB.keys()))
        serv_sel = st.multiselect("Serviços:", list(SERVICOS_DB.keys()))
        url_modelo = URL_PROCURACAO
    else:
        info_adicional = st.text_input("Dados adicionais / Contrato:")
        url_modelo = URL_DISTRATO

    if st.button("🚀 Gerar Documento"):
        with st.spinner("Gerando..."):
            try:
                # Baixar e preparar modelo
                resp = requests.get(url_modelo)
                with open("modelo.docx", "wb") as f: f.write(resp.content)
                
                # Contexto dinâmico
                ctx = {
                    "empresa": empresa.upper(), "cnpj": cnpj,
                    "endereco_contratante": end_cont,
                    "endereco_obra": end_obra if end_obra else end_cont,
                    "data": datetime.date.today().strftime("%d de %B de %Y").replace("May", "maio")
                }
                
                if tipo_doc == "Procuração":
                    lista = [RESPONSAVEIS_DB[k]["nome"] + RESPONSAVEIS_DB[k]["info"] for k in resp_sel]
                    ctx["responsaveis"] = "e funcionários: " + "; ".join(lista) + "." if lista else "."
                    ctx["servicos"] = "\n".join([f"• {SERVICOS_DB[s]}" for s in serv_sel])
                else:
                    ctx["info_adicional"] = info_adicional
                
                # Renderizar
                doc = DocxTemplate("modelo.docx")
                doc.render(ctx)
                doc.save("final.docx")
                subprocess.run(["libreoffice", "--headless", "--convert-to", "pdf", "final.docx"], check=True)
                
                with open("final.pdf", "rb") as f:
                    st.download_button(f"📥 Baixar {tipo_doc}.pdf", f, file_name=f"{tipo_doc}_{empresa}.pdf")
            except Exception as e:
                st.error(f"Erro: {e}")
else:
    st.info("Setor em desenvolvimento...")
