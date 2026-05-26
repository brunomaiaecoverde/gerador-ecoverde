import streamlit as st
from docxtpl import DocxTemplate
import requests
import re
import subprocess
import os
import datetime

# LINK DO SEU GOOGLE DOCS CONFIGURADO
GOOGLE_DOCS_URL = "https://docs.google.com/document/d/1LUtJrayUBGobMQ1jy2QFIeL0WoflNXNe/edit?usp=sharing&ouid=106316936646315398659&rtpof=true&sd=true"

# 1. BASE DE DADOS INTEGRADA
RESPONSAVEIS_DB = {
    "Gabriela Freitas": {"nome": "Gabriela Ferreira de Freitas", "info": ", brasileira, solteira, líder de projetos júnior 2, CPF 121.142.656-41"},
    "Shérida Silva": {"nome": "Shérida Patrícia Milanez Silva", "info": ", brasileira, casada, Analista de integração de projetos júnior, CPF 072.958.656-16"},
    "Eduarda Ferreira": {"nome": "Eduarda Pimenta Ferreira", "info": ", brasileira, solteira, desenhista, CPF 703.540.986-67"},
    "Caroline Monte Mor": {"nome": "Caroline Candido Oliveira Monte Mor", "info": ", brasileira, casada, desenhista projetista 2, CPF 124.910.106-90"},
    "Maria Luiza Santos": {"nome": "Maria Luiza Cardozo Estevão dos Santos", "info": ", brasileira, solteira, desenhista projetista 2, CPF 144.200.216-67"},
    "Bruno Maia": {"nome": "Bruno Maia", "info": ", brasileiro, solteiro, estagiário de projetos viários, CPF: 142.172.726-93"},
    "Gustavo Costa": {"nome": "Gustavo de Souza Costa", "info": ", brasileiro, solteiro, Analista de Projetos Viários Júnior II, CPF 105.159.886-93"}
}

SERVICOS_DB = {
    "PSCIP (Incêndio)": "Projeto de segurança e combate contra incêndio e pânico perante ao órgão competente;",
    "Aprovação Arquitetônica": "Processo de aprovação de projetos arquitetônicos perante ao órgão competente;",
    "RIC (Trânsito)": "Elaboração e aprovação de Relatório de Impacto na Circulação, perante ao órgão competente;",
    "EIV (Vizinhança)": "Elaboração e aprovação de Estudo de Impacto na Vizinhança, perante ao órgão competente;",
    "Projeto Viário": "Elaboração e aprovação de Projeto Viário, perante ao órgão competente;"
}

# CONFIGURAÇÃO DA PÁGINA WEB
st.set_page_config(page_title="Gerador de Procurações", page_icon="📄", layout="centered")

st.title("📄 Gerador Automático de Procurações")
st.caption("Ecoverde Projetos e Consultoria Ambiental - Sistema Cloud")

# 2. CAPTURA DOS DADOS DA URL (Com blindagem contra erros de link)
query_params = st.query_params
url_empresa = query_params.get("empresa", "<<NOME DA EMPRESA>>")
url_cnpj = query_params.get("cnpj", "<<CNPJ>>")
url_endereco_contratante = query_params.get("endereco_contratante", "<<ENDEREÇO DO CONTRATANTE>>")
url_endereco_obra = query_params.get("endereco_obra", "")

# Lógica inteligente: se a obra vier vazia, usa a do contratante
if not url_endereco_obra or url_endereco_obra == "<<ENDEREÇO DA OBRA>>":
    url_endereco_obra = url_endereco_contratante

st.subheader("1. Dados do Cliente e Processo")
col1, col2 = st.columns(2)
with col1:
    empresa = st.text_input("Cliente / Razão Social", value=url_empresa)
    endereco_contratante = st.text_input("Endereço do Contratante (Sede)", value=url_endereco_contratante)
with col2:
    cnpj = st.text_input("CNPJ", value=url_cnpj)
    endereco_obra = st.text_input("Endereço da Obra (Imóvel)", value=url_endereco_obra)

st.write("---")

st.subheader("2. Seleção de Responsáveis (Outorgados)")
responsaveis_selecionados = st.multiselect("Selecione os funcionários:", options=list(RESPONSAVEIS_DB.keys()), default=[])

st.write("---")

st.subheader("3. Escopo de Poderes (Serviços)")
servicos_selecionados = st.multiselect("Selecione os serviços:", options=list(SERVICOS_DB.keys()), default=[])

# 4. GERAÇÃO
if st.button("🚀 Gerar e Baixar Procuração em PDF", type="primary"):
    with st.spinner("Processando..."):
        try:
            match = re.search(r'/document/d/([a-zA-Z0-9-_]+)', GOOGLE_DOCS_URL)
            doc_id = match.group(1)
            response = requests.get(f"https://docs.google.com/document/d/{doc_id}/export?format=docx")
            
            arquivo_base = "base_modelo.docx"
            with open(arquivo_base, "wb") as f: f.write(response.content)
            
            # Responsáveis (Texto puro para evitar bug de fonte na nuvem)
            if responsaveis_selecionados:
                lista_resp = [RESPONSAVEIS_DB[k]["nome"] + RESPONSAVEIS_DB[k]["info"] for k in responsaveis_selecionados]
                texto_responsaveis = "e funcionários: " + "; ".join(lista_resp) + "."
            else:
                texto_responsaveis = "."
                
            # Serviços
            lista_servicos = [f"• {SERVICOS_DB[s]}" for s in servicos_selecionados]
            lista_servicos.append("• Outras autarquias cabíveis ao município.")
            texto_servicos = "\n".join(lista_servicos)
            
            # Contexto
            context = {
                "empresa": empresa.upper(),
                "cnpj": cnpj,
                "endereco_contratante": endereco_contratante,
                "endereco_obra": endereco_obra,
                "responsaveis": texto_responsaveis,
                "servicos": texto_servicos,
                "data": datetime.date.today().strftime("%d de %B de %Y").replace("May", "maio")
            }

            doc = DocxTemplate(arquivo_base)
            doc.render(context)
            doc.save("temp.docx")
            
            subprocess.run(["libreoffice", "--headless", "--convert-to", "pdf", "temp.docx"], check=True)
            
            with open("temp.pdf", "rb") as f:
                st.download_button("📄 Baixar Procuração (.pdf)", f, file_name=f"Procuracao_{empresa.replace(' ', '_')}.pdf")
                    
        except Exception as e:
            st.error(f"Erro: {e}")
