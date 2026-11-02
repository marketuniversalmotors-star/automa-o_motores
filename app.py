import streamlit as st
import openpyxl
import io
import pandas as pd
import json
import os
import glob

st.set_page_config(page_title="Universal Motors - Registo de Ensaio", page_icon="⚡", layout="wide")

# --- CSS PERSONALIZADO PARA RESPONSIVIDADE NO TELEMÓVEL ---
st.markdown("""
    <style>
    /* Ajustar o título principal para telemóveis */
    @media (max-width: 768px) {
        h1 {
            font-size: 1.8rem !important;
        }
        h2 {
            font-size: 1.4rem !important;
        }
        h3 {
            font-size: 1.2rem !important;
        }
    }
    </style>
""", unsafe_allow_html=True)

# --- DIRETÓRIOS E PERSISTÊNCIA NA NUVEM ---
DB_FILE = "dados_motores.json"
EXCEL_DIR = "excel_gerados"

if not os.path.exists(EXCEL_DIR):
    os.makedirs(EXCEL_DIR)

def carregar_dados():
    if os.path.exists(DB_FILE):
        try:
            with open(DB_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return []
    return []

def guardar_dados(dados):
    with open(DB_FILE, "w", encoding="utf-8") as f:
        json.dump(dados, f, ensure_ascii=False, indent=4)

# Inicializar os valores dos campos de input no session_state para limpeza automática
if "val_u" not in st.session_state: st.session_state.val_u = ""
if "val_v" not in st.session_state: st.session_state.val_v = ""
if "val_w" not in st.session_state: st.session_state.val_w = ""
if "val_serie" not in st.session_state: st.session_state.val_serie = ""
if "val_fuga" not in st.session_state: st.session_state.val_fuga = ""

# --- BARRA LATERAL ---
st.sidebar.subheader("⚙️ Gestão de Registos")
if st.sidebar.button("🗑️ Limpar Todos os Registos"):
    guardar_dados([])
    st.rerun()

st.sidebar.markdown("---")

# --- ÁREA DE MÚLTIPLOS FICHEIROS GERADOS ---
st.sidebar.subheader("📥 Ficheiros Gerados")
ficheiros_excel = glob.glob(os.path.join(EXCEL_DIR, "*.xlsx"))

if ficheiros_excel:
    ficheiros_ordenados = sorted(ficheiros_excel, key=os.path.getmtime, reverse=True)
    
    for ficheiro in ficheiros_ordenados:
        nome_fich = os.path.basename(ficheiro)
        with open(ficheiro, "rb") as f:
            excel_bytes_guardado = f.read()
        
        st.sidebar.download_button(
            label=f"📥 {nome_fich}",
            data=excel_bytes_guardado,
            file_name=nome_fich,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=ficheiro,
            use_container_width=True
        )
    
    if st.sidebar.button("🗑️ Limpar Todos os Excel Gerados"):
        for f in ficheiros_excel:
            try:
                os.remove(f)
            except:
                pass
        st.rerun()
else:
    st.sidebar.info("Nenhum Excel gerado ainda.")

st.sidebar.markdown("---")
st.sidebar.info("💡 **Dica de Utilização:**\n\n* Utilize o teclado do telemóvel em cada campo e clique em **Ok (Adicionar)**.\n\n* O Colaborador preenche a nossa Ordem de Produção e os dados dos motores pelo telemóvel e no final clica em Processar e Gerar Excel Final.")

# -------------------------------------------------------------

st.title("⚡ Registo de Ensaio de Motores")

# Campos de Cabeçalho (Ordem de Compra e Ordem de Produção)
col_h1, col_h2 = st.columns(2)
with col_h1:
    ordem_compra = st.text_input("Ordem de Compra:", value="EF001EF25/02868")
with col_h2:
    ordem_producao = st.text_input("Ordem de Produção:", value="2026/452")

st.markdown("---")
st.subheader("Introdução de Dados do Motor")

# Campos de Entrada limpos (sem microfones)
col_u, col_v, col_w, col_ser, col_fuga = st.columns(5)

with col_u:
    st.session_state.val_u = st.text_input("U", value=st.session_state.val_u, placeholder="Ex: 8,82")
with col_v:
    st.session_state.val_v = st.text_input("V", value=st.session_state.val_v, placeholder="Ex: 9,0")
with col_w:
    st.session_state.val_w = st.text_input("W", value=st.session_state.val_w, placeholder="Ex: 8,65")
with col_ser:
    st.session_state.val_serie = st.text_input("No série", value=st.session_state.val_serie, placeholder="Ex: 261933-001")
with col_fuga:
    st.session_state.val_fuga = st.text_input("Ifuga (mA)", value=st.session_state.val_fuga, placeholder="Ex: 3,9")

# Botões de Ação lado a lado
col_btn1, col_btn2 = st.columns([1, 4])

with col_btn1:
    btn_ok = st.button("✅ Ok (Adicionar)", type="primary", use_container_width=True)

with col_btn2:
    btn_gerar = st.button("✨ Processar e Gerar Excel Final", use_container_width=True)

# Lógica ao clicar em "Ok" (Guarda no ficheiro partilhado)
if btn_ok:
    if not st.session_state.val_serie.strip() and not st.session_state.val_fuga.strip():
        st.warning("Por favor, preencha pelo menos o Número de Série ou a Fuga.")
    else:
        novo_motor = {
            "u": st.session_state.val_u.strip(),
            "v": st.session_state.val_v.strip(),
            "w": st.session_state.val_w.strip(),
            "no_serie": st.session_state.val_serie.strip(),
            "ifuga": st.session_state.val_fuga.strip()
        }
        
        medicoes = carregar_dados()
        medicoes.append(novo_motor)
        guardar_dados(medicoes)
        
        # Limpar os valores dos inputs instantaneamente
        st.session_state.val_u = ""
        st.session_state.val_v = ""
        st.session_state.val_w = ""
        st.session_state.val_serie = ""
        st.session_state.val_fuga = ""
        
        st.success(f"Motor adicionado com sucesso! ({len(medicoes)} registados)")
        st.rerun()

# --- TABELA DE PRÉ-VISUALIZAÇÃO EM TEMPO REAL ---
medicoes_atuais = carregar_dados()

st.markdown("---")
st.subheader(f"📊 Pré-visualização da Tabela ({len(medicoes_atuais)} motores adicionados)")

if medicoes_atuais:
    dados_exibicao = []
    for idx, m in enumerate(medicoes_atuais):
        bloco = "Bloco 1 (1-20)" if idx < 20 else "Bloco 2 (21-40)"
        dados_exibicao.append({
            "Nº Registo": idx + 1,
            "Bloco": bloco,
            "U": m["u"],
            "V": m["v"],
            "W": m["w"],
            "No série": m["no_serie"],
            "Ifuga (mA)": m["ifuga"]
        })
    
    df_preview = pd.DataFrame(dados_exibicao)
    st.dataframe(df_preview, use_container_width=True, hide_index=True)
else:
    st.info("A tabela está vazia. Assim que o colaborador adicionar um motor, ele aparecerá aqui instantaneamente.")

# --- GERAR EXCEL FINAL ---
if btn_gerar:
    medicoes_para_excel = carregar_dados()
    if not medicoes_para_excel:
        st.warning("Adicione pelo menos um motor antes de gerar o Excel.")
    else:
        try:
            wb = openpyxl.load_workbook("modelo_registo.xlsx")
            ws = wb.active
            
            # Preencher Ordem de Compra (C5) e Ordem de Produção (C6)
            if ordem_compra:
                ws["C5"] = ordem_compra
            if ordem_producao:
                ws["C6"] = ordem_producao
            
            for idx, item in enumerate(medicoes_para_excel):
                if idx < 20:
                    row = 8 + idx
                    if item["u"]: ws[f"A{row}"] = item["u"]
                    if item["v"]: ws[f"B{row}"] = item["v"]
                    if item["w"]: ws[f"C{row}"] = item["w"]
                    if item["no_serie"]: ws[f"D{row}"] = item["no_serie"]
                    if item["ifuga"]: ws[f"E{row}"] = item["ifuga"]
                elif idx < 40:
                    row = 8 + (idx - 20)
                    if item["u"]: ws[f"F{row}"] = item["u"]
                    if item["v"]: ws[f"G{row}"] = item["v"]
                    if item["w"]: ws[f"H{row}"] = item["w"]
                    if item["no_serie"]: ws[f"I{row}"] = item["no_serie"]
                    if item["ifuga"]: ws[f"J{row}"] = item["ifuga"]
            
            output = io.BytesIO()
            wb.save(output)
            excel_bytes = output.getvalue()
            
            oc_limpa = ordem_compra.replace('/', '_').replace(' ', '_')
            op_limpa = ordem_producao.replace('/', '_').replace(' ', '_')
            nome_ficheiro = f"REGISTO_OC_{oc_limpa}_OP_{op_limpa}.xlsx"
            
            caminho_completo = os.path.join(EXCEL_DIR, nome_ficheiro)
            with open(caminho_completo, "wb") as f:
                f.write(excel_bytes)
            
            st.success(f"Ficheiro '{nome_ficheiro}' gerado e adicionado à lista lateral com sucesso!")
            st.rerun()
            
        except Exception as e:
            st.error(f"Erro ao processar Excel: {e}")
