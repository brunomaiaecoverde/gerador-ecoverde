import streamlit as st
from docxtpl import DocxTemplate
import requests
import re
import subprocess
import os
import datetime

# LINKS DOS MODELOS (Google Docs configurados para exportação)
URL_PROCURACAO = "https://docs.google.com/document/d/1LUtJrayUBGobMQ1jy2QFIeL0WoflNXNe/export?format=docx"
URL_DISTRATO = "https://docs.google.com/document/d/1Tf6rRbCnBTBcWFr-pVDw7sBV0yZRqeqa/export?format=docx"

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

st.set_page_config(page_title="Sistema Ecoverde", page_icon="📄", layout="centered")

# --- CONTROLE DE NAVEGAÇÃO (MEMÓRIA DO SITE) ---
if 'setor_selecionado' not in st.session_state:
    st.session_state.setor_selecionado = None
if 'doc_selecionado' not in st.session_state:
    st.session_state.doc_selecionado = None

def alterar_setor(nome_setor):
    st.session_state.setor_selecionado = nome_setor
    st.session_state.doc_selecionado = None # Reseta o documento ao trocar de setor

def alterar_doc(nome_doc):
    st.session_state.doc_selecionado = nome_doc

# --- PÁGINA INICIAL (MENU DE CARDS GRANDES) ---
if st.session_state.setor_selecionado is None:
    st.title("🏢 Sistema Central Ecoverde")
    st.write("Selecione um setor abaixo para iniciar:")
    
    st.write("---")
    
    # Usando colunas e botões primários para dar destaque visual
    col1, col2 = st.columns(2, gap="large")
    
    with col1:
        st.button("🏗️ PROJETOS", on_click=alterar_setor, args=("Projetos",), use_container_width=True, type="primary")
        st.button("🌱 MEIO AMBIENTE", on_click=alterar_setor, args=("Meio Ambiente",), use_container_width=True)
        
    with col2:
        st.button("📊 GESTÃO AMBIENTAL", on_click=alterar_setor, args=("Gestão Ambiental",), use_container_width=True)
        st.button("👷 EXECUÇÃO", on_click=alterar_setor, args=("Execução",), use_container_width=True)

# --- PÁGINA DO SETOR: PROJETOS ---
elif st.session_state.setor_selecionado == "Projetos":
    
    # NÍVEL 1: SELEÇÃO DE DOCUMENTO (Cards Menores)
    if st.session_state.doc_selecionado is None:
        st.button("⬅️ Voltar ao Menu Principal", on_click=alterar_setor, args=(None,))
        st.write("---")
        st.title("🏗️ Setor de Projetos")
        st.write("Selecione o documento que deseja gerar:")
        
        col1, col2, col3 = st.columns(3)
        with col1:
            st.button("📄 Procuração", on_click=alterar_doc, args=("Procuração",), use_container_width=True)
        with col2:
            st.button("❌ Termo de Distrato", on_click=alterar_doc, args=("Termo de Distrato",), use_container_width=True)
        with col3:
            st.button("✅ Termo de Finalização", on_click=alterar_doc, args=("Termo de Finalização",), use_container_width=True)
            
    # NÍVEL 2: FORMULÁRIO DO DOCUMENTO
    else:
        st.button("⬅️ Voltar aos Documentos", on_click=alterar_doc, args=(None,))
        st.write("---")
        
        st.title(f"📄 Gerador - {st.session_state.doc_selecionado}")
        
        # 2. CAPTURA DOS DADOS DA URL (Com blindagem)
        query_params = st.query_params
        url_empresa = query_params.get("empresa", "")
        url_cnpj = query_params.get("cnpj", "")
        url_endereco_contratante = query_params.get("endereco_contratante", "")
        url_endereco_obra = query_params.get("endereco_obra", "")

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
        
        # Lógica Específica da Procuração
        if st.session_state.doc_selecionado == "Procuração":
            st.subheader("2. Seleção de Responsáveis (Outorgados)")
            responsaveis_selecionados = st.multiselect("Selecione os funcionários:", options=list(RESPONSAVEIS_DB.keys()), default=[])

            st.write("---")

            st.subheader("3. Escopo de Poderes (Serviços)")
            servicos_selecionados = st.multiselect("Selecione os serviços:", options=list(SERVICOS_DB.keys()), default=[])
            
            url_modelo = URL_PROCURACAO

        # Lógica do Distrato e Finalização
        else:
            st.subheader("2. Dados Adicionais")
            info_adicional = st.text_input("Número do Contrato ou Informação Adicional:")
            if st.session_state.doc_selecionado == "Termo de Distrato":
                url_modelo = URL_DISTRATO
            else:
                url_modelo = "" # O link da finalização será colocado aqui depois
                
        st.write("---")

        # 4. GERAÇÃO
        if st.button(f"🚀 Gerar e Baixar {st.session_state.doc_selecionado}", type="primary"):
            if st.session_state.doc_selecionado == "Termo de Finalização" and not url_modelo:
                st.warning("⚠️ O link do modelo do Termo de Finalização ainda não foi configurado.")
            else:
                with st.spinner("Processando..."):
                    try:
                        match = re.search(r'/document/d/([a-zA-Z0-9-_]+)', url_modelo)
                        doc_id = match.group(1)
                        response = requests.get(f"https://docs.google.com/document/d/{doc_id}/export?format=docx")
                        
                        arquivo_base = "base_modelo.docx"
                        with open(arquivo_base, "wb") as f: f.write(response.content)
                        
                        # Data sempre em PT-BR
                        meses_pt = {
                            "January": "janeiro", "February": "fevereiro", "March": "março", 
                            "April": "abril", "May": "maio", "June": "junho", 
                            "July": "julho", "August": "agosto", "September": "setembro", 
                            "October": "outubro", "November": "novembro", "December": "dezembro"
                        }
                        data_hoje = datetime.date.today()
                        mes_nome = meses_pt[data_hoje.strftime("%B")]
                        data_formatada = f"{data_hoje.day} de {mes_nome} de {data_hoje.year}"
                        
                        # Contexto Compartilhado
                        context = {
                            "empresa": empresa.upper(),
                            "cnpj": cnpj,
                            "endereco_contratante": endereco_contratante,
                            "endereco_obra": endereco_obra,
                            "data": data_formatada
                        }
                        
                        # Preenche contexto específico
                        if st.session_state.doc_selecionado == "Procuração":
                            if responsaveis_selecionados:
                                lista_resp = [RESPONSAVEIS_DB[k]["nome"] + RESPONSAVEIS_DB[k]["info"] for k in responsaveis_selecionados]
                                context["responsaveis"] = "e funcionários: " + "; ".join(lista_resp) + "."
                            else:
                                context["responsaveis"] = "."
                                
                            lista_servicos = [f"• {SERVICOS_DB[s]}" for s in servicos_selecionados]
                            lista_servicos.append("• Outras autarquias cabíveis ao município.")
                            context["servicos"] = "\n".join(lista_servicos)
                        else:
                            context["info_adicional"] = info_adicional

                        doc = DocxTemplate(arquivo_base)
                        doc.render(context)
                        doc.save("temp.docx")
                        
                        subprocess.run(["libreoffice", "--headless", "--convert-to", "pdf", "temp.docx"], check=True)
                        
                        with open("temp.pdf", "rb") as f:
                            st.download_button("📄 Baixar PDF", f, file_name=f"{st.session_state.doc_selecionado.replace(' ', '_')}_{empresa.replace(' ', '_')}.pdf")
                                
                    except Exception as e:
                        st.error(f"Erro: {e}")

# --- OUTROS SETORES ---
else:
    st.button("⬅️ Voltar ao Menu Principal", on_click=alterar_setor, args=(None,))
    st.write("---")
    st.info(f"O setor de {st.session_state.setor_selecionado} ainda está em desenvolvimento.")
