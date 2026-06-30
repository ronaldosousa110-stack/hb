import os
import zipfile
from io import BytesIO
import pandas as pd
import streamlit as st
from docx import Document

# 1. Configura o visual da página do site
st.set_page_config(page_title="Gerador de Certificados NR", page_icon="🎓", layout="centered")

# Cria as caixas de memória do Streamlit se elas não existirem
if "conectado" not in st.session_state:
    st.session_state["conectado"] = False

# ==========================================================
# SISTEMA DE LOGIN
# ==========================================================
def tela_login():
    st.subheader("🔒 Acesso Restrito")
    usuario = st.text_input("Utilizador", key="user_input")
    senha = st.text_input("Senha", type="password", key="pass_input")
    
    if st.button("Entrar", type="primary"):
        try:
            # Puxa o utilizador e senha configurados nos Secrets do site
            usuario_correto = st.secrets["credenciais"]["usuario"]
            senha_correta = st.secrets["credenciais"]["senha"]
            
            if usuario == usuario_correto and senha == senha_correta:
                st.session_state["conectado"] = True
                st.success("Acesso concedido! A carregar...")
                st.rerun() 
            else:
                st.error("Utilizador ou senha incorretos.")
        except KeyError:
            st.error("Erro técnico: As credenciais ainda não foram configuradas nos 'Secrets' do Streamlit Cloud.")

# ==========================================================
# SEU DICIONÁRIO DE MODELOS
# ==========================================================
MODELOS = {
    "NR-01 Integração": "modelo_certificadoNR01.docx",
    "NR-06 EPI": "modelo_certificadoNR06.docx",
    "NR-10 Segurança em Eletricidade": "modelo_certificadoNR10.docx",
    "NR-11 Transporte e Movimentação": "modelo_certificadoNR11.docx",
    "NR-12 Máquinas e Equipamentos": "modelo_certificadoNR12.docx",
    "NR-12 Operador de Motosserra": "modelo_certificadoNR12MOTOSSERRA.docx",
    "NR-18 Construção Civil": "modelo_certificadoNR18.docx",
    "NR-23 Combate a Incêndio": "modelo_certificadoNR23.docx",
    "NR-31.7 Segurança na Aplicação de Agrotóxicos": "modelo_certificadoNR31.7.docx",
    "NR-32 Serviços de Saúde": "modelo_certificadoNR32.docx",
    "NR-34 Trabalho a Quente": "modelo_certificadoNR34.docx",
    "NR-35 Trabalho em Altura": "modelo_certificadoNR35.docx"
}

# Lógica de substituição de texto original
def processar_substituicao(paragrafo, todas_tags, dados_com_negrito):
    if not any(tag in paragrafo.text for tag in todas_tags):
        return
    texto_completo = paragrafo.text
    paragrafo.text = ""
    while True:
        proxima_tag = None
        menor_posicao = len(texto_completo)
        for tag in todas_tags:
            posicao = texto_completo.find(tag)
            if posicao != -1 and posicao < menor_posicao:
                menor_posicao = posicao
                proxima_tag = tag
        if proxima_tag:
            texto_antes = texto_completo[:menor_posicao]
            run_antes = paragrafo.add_run(texto_antes)
            valor_real = todas_tags[proxima_tag]
            run_dados = paragrafo.add_run(valor_real)
            if proxima_tag in dados_com_negrito:
                run_dados.bold = True
            texto_completo = texto_completo[menor_posicao + len(proxima_tag):]
        else:
            paragrafo.add_run(texto_completo)
            break

# ==========================================================
# CONTROLO DE TELAS (LOGIN VS PAINEL)
# ==========================================================

# SE NÃO ESTIVER CONECTADO: Mostra apenas o Login E a Planilha Modelo embaixo
if not st.session_state["conectado"]:
    tela_login()
    st.write("---")
    
    # Botão da planilha disponível logo na tela inicial
    caminho_planilha_modelo = os.path.join("static", "modelo_dados.xlsx")
    if os.path.exists(caminho_planilha_modelo):
        with open(caminho_planilha_modelo, "rb") as f:
            st.download_button(
                label="ℹ️ Descarregar Planilha Modelo de Inserção",
                data=f,
                file_name="modelo_dados.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.warning("Aviso: O ficheiro 'modelo_dados.xlsx' não foi encontrado na pasta 'static'. Certifique-se de que fez o upload dele no GitHub.")
    
    st.stop() # Bloqueia o restante do painel para quem não está logado

# ==========================================================
# PAINEL PRINCIPAL (SÓ APARECE APÓS LOGIN CORRETO)
# ==========================================================
st.title("🎓 Gerador Web de Certificados NR")

if st.sidebar.button("Sair / Logout"):
    st.session_state["conectado"] = False
    st.rerun()

st.write("---")

nr_escolhida = st.selectbox("1. Selecione o Treinamento (NR):", ["Clique para escolher..."] + list(MODELOS.keys()))
arquivo_excel = st.file_uploader("2. Envie a planilha de dados preenchida (.xlsx):", type=["xlsx", "xls"])

if nr_escolhida != "Clique para escolher..." and arquivo_excel is not None:
    if st.button("Processar e Gerar Certificados", type="primary"):
        try:
            df = pd.read_excel(arquivo_excel)
            caminho_modelo = os.path.join('modelos_docx', MODELOS[nr_escolhida])
            
            if not os.path.exists(caminho_modelo):
                st.error(f"Erro: O modelo '{MODELOS[nr_escolhida]}' não foi encontrado na pasta 'modelos_docx'.")
            else:
                memoria_zip = BytesIO()
                barra_progresso = st.progress(0)
                total_linhas = len(df)
                
                with zipfile.ZipFile(memoria_zip, 'a', zipfile.ZIP_DEFLATED) as zip_file:
                    for idx, linha in df.iterrows():
                        doc = Document(caminho_modelo)
                        data_formatada = pd.to_datetime(linha["Data_Final"]).strftime('%d/%m/%Y') if pd.notna(linha.get("Data_Final")) else ""
                        
                        dados_com_negrito = {
                            "[NOME]": str(linha["Nome"]) if pd.notna(linha.get("Nome")) else "",
                            "[CPF]": str(linha["CPF"]) if pd.notna(linha.get("CPF")) else "",
                            "[EMPRESA]": str(linha["Empresa"]) if pd.notna(linha.get("Empresa")) else "",
                            "[CNPJ]": str(linha["CNPJ"]) if pd.notna(linha.get("CNPJ")) else "",
                            "[PERIODO]": str(linha["Periodo"]) if pd.notna(linha.get("Periodo")) else ""
                        }
                        dados_sem_negrito = {"[DATA_FINAL]": data_formatada}
                        todas_tags = {**dados_com_negrito, **dados_sem_negrito}
                        
                        for p in doc.paragraphs:
                            processar_substituicao(p, todas_tags, dados_com_negrito)
                        for t in doc.tables:
                            for row in t.rows:
                                for cell in row.cells:
                                    for p in cell.paragraphs:
                                        processar_substituicao(p, todas_tags, dados_com_negrito)
                                        
                        nome_limpo = str(linha["Nome"]).strip().replace(" ", "_")
                        nome_final_arquivo = f"{nr_escolhida.replace(' ', '_')}_{nome_limpo}.docx"
                        
                        memoria_doc = BytesIO()
                        doc.save(memoria_doc)
                        memoria_doc.seek(0)
                        zip_file.writestr(nome_final_arquivo, memoria_doc.getvalue())
                        barra_progresso.progress((idx + 1) / total_linhas)
                
                memoria_zip.seek(0)
                st.success("✨ Todos os certificados foram gerados!")
                
                st.download_button(
                    label="📥 Descarregar Todos os Certificados (.ZIP)",
                    data=memoria_zip,
                    file_name=f"Certificados_{nr_escolhida.replace(' ', '_')}.zip",
                    mimetype="application/zip"
                )
        except Exception as e:
            st.error(f"Ocorreu um erro ao ler a planilha: {e}")

st.write("---")
# Mostra também o modelo dentro do painel logado para garantir dupla disponibilidade
caminho_planilha_modelo = os.path.join("static", "modelo_dados.xlsx")
if os.path.exists(caminho_planilha_modelo):
    with open(caminho_planilha_modelo, "rb") as f:
        st.download_button(
            label="ℹ️ Descarregar Planilha Modelo de Inserção",
            data=f,
            file_name="modelo_dados.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="modelo_logado"
        )
