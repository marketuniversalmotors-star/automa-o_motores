import streamlit as st
import openpyxl
import io
import pandas as pd
import json
import os

st.set_page_config(page_title="Universal Motors - Registo de Ensaio", page_icon="⚡", layout="wide")

# --- FICHEIRO DE PERSISTÊNCIA LOCAL / NUVEM ---
DB_FILE = "dados_motores.json"

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
st.sidebar.info("💡 **Dica de Utilização:**\n- O operador preenche os motores pelo telemóvel.\n- A colega acompanha a tabela no PC, define as Ordens e descarrega o Excel final.")

# -------------------------------------------------------------

st.title("⚡ Registo de Ensaio de Motores - Nuvem (Tempo Real)")
st.markdown("Aplicação sincronizada em tempo real. Utilize o microfone (voz) em cada campo e clique em **Ok (Adicionar)**.")

# Campos de Cabeçalho (Ordem de Compra e Ordem de Produção)
col_h1, col_h2 = st.columns(2)
with col_h1:
    ordem_compra = st.text_input("Ordem de Compra:", value="EF001EF25/02868")
with col_h2:
    ordem_producao = st.text_input("Ordem de Produção:", value="2026/452")

st.markdown("---")
st.subheader("🎤 Introdução de Dados do Motor")

# Campos de Entrada com suporte a ditado por voz
col_u, col_v, col_w, col_ser, col_fuga = st.columns(5)

with col_u:
    st.session_state.val_u = st.text_input("U 🎤", value=st.session_state.val_u, placeholder="Ex: 8,82")
with col_v:
    st.session_state.val_v = st.text_input("V 🎤", value=st.session_state.val_v, placeholder="Ex: 9,0")
with col_w:
    st.session_state.val_w = st.text_input("W 🎤", value=st.session_state.val_w, placeholder="Ex: 8,65")
with col_ser:
    st.session_state.val_serie = st.text_input("No série 🎤", value=st.session_state.val_serie, placeholder="Ex: 261933-001")
with col_fuga:
    st.session_state.val_fuga = st.text_input("Ifuga (mA) 🎤", value=st.session_state.val_fuga, placeholder="Ex: 3,9")

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
st.subheader(f"📊 Pré-visualização da Tabela em Tempo Real ({len(medicoes_atuais)} motores adicionados)")

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
            
            nome_ficheiro = f"REGISTO_ENSAIO_{ordem_compra.replace('/', '_')}.xlsx"
            
            st.success("Ficheiro Excel oficial gerado com sucesso!")
            st.download_button(
                label="📥 Descarregar Ficheiro Excel Oficial",
                data=excel_bytes,
                file_name=nome_ficheiro,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"Erro ao processar Excel: {e}")