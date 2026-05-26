import streamlit as st
from docxtpl import DocxTemplate
import requests
import subprocess
import os
import datetime

# LINKS DOS MODELOS NO DRIVE (Exportação DOCX)
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

# INTERFACE PRINCIPAL
st.set_page_config(page_title="Sistema Ecoverde", page_icon="📄", layout="wide")
st.title("📄 Sistema Central de Documentos Ecoverde")
st.write("---")

# 1. CAPTURA DOS DADOS DA URL (Vindos do Notion)
query_params = st.query_params
url_empresa = query_params.get("empresa", "")
url_cnpj = query_params.get("cnpj", "")
url_endereco_contratante = query_params.get("endereco_contratante", "")
url_endereco_obra = query_params.get("endereco_obra", "")

# Lógica inteligente: se a obra vier vazia ou for o texto padrão, usa a do contratante
if not url_endereco_obra or url_endereco_obra == "<<ENDEREÇO DA OBRA>>":
    url_endereco_obra = url_endereco_contratante

# 2. ESCOLHA DE SETOR (CARDS / ABAS)
st.subheader("1. Seleção de Setor e Documento")
aba_projetos, aba_meio_ambiente, aba_gestao, aba_execucao = st.tabs(["🏗️ Projetos", "🌱 Meio Ambiente", "📊 Gestão Ambiental", "👷 Execução"])

# --- ABA DE PROJETOS ---
with aba_projetos:
    tipo_doc = st.selectbox("Qual documento deseja gerar?", ["Procuração", "Termo de Distrato", "Termo de Finalização"], key="doc_projetos")
    
    st.write("---")
    st.subheader("2. Dados do Cliente e Processo")
    
    # 3. INPUTS PREENCHIDOS AUTOMATICAMENTE (O ERRO ESTAVA AQUI, AGORA CORRIGIDO)
    col1, col2 = st.columns(2)
    with col1:
        empresa = st.text_input("Cliente / Razão Social", value=url_empresa)
        endereco_contratante = st.text_input("Endereço do Contratante (Sede)", value=url_endereco_contratante)
    with col2:
        cnpj = st.text_input("CNPJ", value=url_cnpj)
        endereco_obra = st.text_input("Endereço da Obra (Imóvel)", value=url_endereco_obra)

    st.write("---")
    
    # 4. CAMPOS ESPECÍFICOS POR DOCUMENTO
    if tipo_doc == "Procuração":
        st.subheader("3. Escopo da Procuração")
        resp_sel = st.multiselect("Responsáveis (Outorgados):", list(RESPONSAVEIS_DB.keys()))
        serv_sel = st.multiselect("Serviços (Poderes):", list(SERVICOS_DB.keys()))
        url_modelo = URL_PROCURACAO
    elif tipo_doc == "Termo de Distrato":
        st.subheader("3. Dados do Distrato")
        info_adicional = st.text_input("Número do Contrato ou Informação Adicional:")
        url_modelo = URL_DISTRATO
    else: # Termo de Finalização
        st.subheader("3. Dados da Finalização")
        info_adicional = st.text_input("Número do Contrato ou Informação Adicional:")
        url_modelo = "" # Deixado em branco até você me mandar o link desse modelo

    st.write("---")
    
    # 5. GERAÇÃO DO PDF
    if st.button("🚀 Gerar e Baixar PDF", type="primary"):
        if tipo_doc == "Termo de Finalização" and not url_modelo:
            st.warning("⚠️ O link do modelo do Termo de Finalização ainda não foi configurado no código.")
        else:
            with st.spinner(f"Gerando {tipo_doc}..."):
                try:
                    # Baixar modelo do Google Docs
                    resp = requests.get(url_modelo)
                    if resp.status_code != 200:
                        st.error("Erro ao baixar o modelo. Verifique se o link está público.")
                        st.stop()
                        
                    with open("modelo.docx", "wb") as f: 
                        f.write(resp.content)
                    
                    # Data formatada em português (Garante que nunca saia 'May')
                    meses_pt = {
                        "January": "janeiro", "February": "fevereiro", "March": "março", 
                        "April": "abril", "May": "maio", "June": "junho", 
                        "July": "julho", "August": "agosto", "September": "setembro", 
                        "October": "outubro", "November": "novembro", "December": "dezembro"
                    }
                    data_hoje = datetime.date.today()
                    mes_nome = meses_pt[data_hoje.strftime("%B")]
                    data_formatada = f"{data_hoje.day} de {mes_nome} de {data_hoje.year}"

                    # Dicionário de substituição (Contexto)
                    ctx = {
                        "empresa": empresa.upper(), 
                        "cnpj": cnpj,
                        "endereco_contratante": endereco_contratante,
                        "endereco_obra": endereco_obra,
                        "data": data_formatada
                    }
                    
                    # Lógica exata e perfeita da Procuração
                    if tipo_doc == "Procuração":
                        if resp_sel:
                            lista_resp = [RESPONSAVEIS_DB[k]["nome"] + RESPONSAVEIS_DB[k]["info"] for k in resp_sel]
                            ctx["responsaveis"] = "e funcionários: " + "; ".join(lista_resp) + "."
                        else:
                            ctx["responsaveis"] = "."
                            
                        if serv_sel:
                            lista_servicos = [f"• {SERVICOS_DB[s]}" for s in serv_sel]
                            lista_servicos.append("• Outras autarquias cabíveis ao município.")
                            ctx["servicos"] = "\n".join(lista_servicos)
                        else:
                            ctx["servicos"] = "• Outras autarquias cabíveis ao município."
                            
                    # Lógica do Distrato e Finalização
                    else:
                        ctx["info_adicional"] = info_adicional
                    
                    # Preencher o documento
                    doc = DocxTemplate("modelo.docx")
                    doc.render(ctx)
                    doc.save("temp.docx")
                    
                    # Converter para PDF
                    subprocess.run(["libreoffice", "--headless", "--convert-to", "pdf", "temp.docx"], check=True)
                    
                    # Disponibilizar download
                    with open("temp.pdf", "rb") as f:
                        st.download_button(
                            label=f"📄 Baixar {tipo_doc} (.pdf)", 
                            data=f, 
                            file_name=f"{tipo_doc}_{empresa.replace(' ', '_')}.pdf",
                            mime="application/pdf"
                        )
                except Exception as e:
                    st.error(f"Erro ao processar o documento: {e}")

# --- OUTROS SETORES (Prontos para receber automações no futuro) ---
with aba_meio_ambiente:
    st.info("O Setor de Meio Ambiente está em desenvolvimento...")
with aba_gestao:
    st.info("O Setor de Gestão Ambiental está em desenvolvimento...")
with aba_execucao:
    st.info("O Setor de Execução está em desenvolvimento...")
