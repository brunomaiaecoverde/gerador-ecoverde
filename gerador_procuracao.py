import streamlit as st
from docxtpl import DocxTemplate, RichText
import requests
import re
import subprocess
import os
import datetime

# COLE O SEU LINK DE PARTILHA DO GOOGLE DOCS AQUI:
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
    "PSCIP (Incêndio)": "Projeto de segurança e combate contra incêndio e pânico através do CBMMG, órgão competente;",
    "Aprovação Arquitetônica": "Processo de aprovação de projetos arquitetônicos perante a Prefeitura Municipal de {cidade};",
    "RIC (Trânsito)": "Elaboração e aprovação de Relatório de Impacto na Circulação, perante a ECOS - Transbetim;",
    "EIV (Vizinhança)": "Elaboração e aprovação de Estudo de Impacto na Vizinhança, perante a SEDURB;",
    "Projeto Viário": "Elaboração e aprovação de Projeto Viário, perante a TRANSCON;"
}

# CONFIGURAÇÃO DA PÁGINA WEB
st.set_page_config(page_title="Gerador de Procurações - Ecoverde", page_icon="📄", layout="centered")

st.title("📄 Gerador Automático de Procurações")
st.caption("Ecoverde Projetos e Consultoria Ambiental - Sistema Cloud")

# 2. CAPTURA DOS DADOS DA URL
query_params = st.query_params
url_empresa = query_params.get("empresa", "<<NOME DA EMPRESA>>")
url_cnpj = query_params.get("cnpj", "<<CNPJ>>")
url_endereco = query_params.get("endereco", "<<ENDEREÇO>>")

st.subheader("1. Dados do Cliente e Processo")
col1, col2 = st.columns(2)
with col1:
    empresa = st.text_input("Cliente / Razão Social", value=url_empresa)
    cnpj = st.text_input("CNPJ", value=url_cnpj)
with col2:
    endereco = st.text_input("Endereço do Contratante", value=url_endereco)
    cidade = st.text_input("Cidade do Processo (Prefeitura e Data)", value="Betim")

st.write("---")

st.subheader("2. Seleção de Responsáveis (Outorgados)")
st.info("A proprietária Karen Kolansky já foi incluída automaticamente no modelo.")
responsaveis_selecionados = st.multiselect(
    "Selecione os funcionários adicionais que farão parte desta procuração:",
    options=list(RESPONSAVEIS_DB.keys()),
    default=["Gabriela Freitas", "Bruno Maia"]
)

st.write("---")

st.subheader("3. Escopo de Poderes (Serviços)")
servicos_selecionados = st.multiselect(
    "Selecione os serviços que constarão nos poderes da procuração:",
    options=list(SERVICOS_DB.keys()),
    default=[]
)

st.write("---")

# 4. DOWNLOAD DO DOCS, PREENCHIMENTO E CONVERSÃO EM PDF
if st.button("🚀 Gerar e Baixar Procuração em PDF", type="primary"):
    
    if not servicos_selecionados:
        st.error("Por favor, selecione ao menos um serviço para compor os poderes.")
    elif GOOGLE_DOCS_URL == "SUA_URL_DO_GOOGLE_DOCS_AQUI":
        st.error("Por favor, configure o link do seu Google Docs na linha 10 do código.")
    else:
        with st.spinner("Buscando modelo no Google Docs e a converter para PDF..."):
            try:
                match = re.search(r'/document/d/([a-zA-Z0-9-_]+)', GOOGLE_DOCS_URL)
                if not match:
                    st.error("Link do Google Docs inválido.")
                    st.stop()
                
                doc_id = match.group(1)
                export_url = f"https://docs.google.com/document/d/{doc_id}/export?format=docx"
                
                response = requests.get(export_url)
                if response.status_code != 200:
                    st.error("Não foi possível aceder ao Google Doc. Verifique as permissões.")
                    st.stop()
                
                arquivo_base = "base_modelo.docx"
                with open(arquivo_base, "wb") as f:
                    f.write(response.content)
                
                rt_responsaveis = RichText()
                # Fonte ajustada para a nuvem: Arial tamanho 11 (22 meias-pontos)
                font_padrao = {'font': 'Arial', 'size': 22}
                
                if responsaveis_selecionados:
                    # Retirei o espaço do início porque você vai dar o espaço no Google Docs
                    rt_responsaveis.add("e funcionários: ", **font_padrao)
                    for i, func_key in enumerate(responsaveis_selecionados):
                        dados = RESPONSAVEIS_DB[func_key]
                        rt_responsaveis.add(dados["nome"], bold=True, **font_padrao)
                        rt_responsaveis.add(dados["info"], **font_padrao)
                        if i < len(responsaveis_selecionados) - 1:
                            rt_responsaveis.add("; ", **font_padrao)
                    rt_responsaveis.add(".", **font_padrao)
                else:
                    rt_responsaveis.add("", **font_padrao)
                    
                lista_servicos = [f"• {SERVICOS_DB[s].format(cidade=cidade)}" for s in servicos_selecionados]
                lista_servicos.append("• Outras autarquias cabíveis ao município.")
                texto_servicos = "\n".join(lista_servicos)
                
                data_atual = datetime.date.today()
                meses = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"]
                texto_data = f"{data_atual.day} de {meses[data_atual.month - 1]} de {data_atual.year}."

                context = {
                    "empresa": empresa.upper(),
                    "cnpj": cnpj,
                    "endereco": endereco,
                    "cidade": cidade,
                    "responsaveis": rt_responsaveis,
                    "servicos": texto_servicos,
                    "data": texto_data
                }

                doc = DocxTemplate(arquivo_base)
                doc.render(context)
                
                arquivo_temp_docx = "temp_procuracao.docx"
                doc.save(arquivo_temp_docx)
                
                subprocess.run(["libreoffice", "--headless", "--convert-to", "pdf", arquivo_temp_docx], check=True)
                
                arquivo_temp_pdf = "temp_procuracao.pdf"
                
                with open(arquivo_temp_pdf, "rb") as pdf_file:
                    pdf_bytes = pdf_file.read()
                    
                st.success("🎉 Procuração em PDF gerada com sucesso!")
                
                st.download_button(
                    label="📄 Baixar Procuração (.pdf)",
                    data=pdf_bytes,
                    file_name=f"Procuracao_{empresa.replace(' ', '_')}.pdf",
                    mime="application/pdf"
                )
                
                for arq in [arquivo_base, arquivo_temp_docx, arquivo_temp_pdf]:
                    if os.path.exists(arq):
                        os.remove(arq)
                        
            except Exception as e:
                st.error(f"Erro ao processar o documento: {e}")
