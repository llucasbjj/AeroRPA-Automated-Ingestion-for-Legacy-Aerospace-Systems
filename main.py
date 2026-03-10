import streamlit as st
import re
import time
import pyautogui
import pdfplumber
import io
import sys
import os
import pandas as pd
import uuid
import threading

# --- CUSTOM MODULES ---
sys.path.append(os.path.join(os.path.dirname(__file__), 'src'))
try:
    from config import *
    from danfe_crawler import DanfeCrawler
    from validator import PlateValidator
except ImportError:
    # Fallbacks for standalone execution
    COMPANY_CNPJ_ROOTS = ('******', '******')
    PAGE_TITLE = "RPA AeroSys"
    PAGE_ICON = "🚀"
    DanfeCrawler = None
    PlateValidator = None

# =========================================================================
# 🚀 AeroRPA: Automated Ingestion for Legacy Aerospace Systems
# =========================================================================

# Page Config
st.set_page_config(page_title=PAGE_TITLE, page_icon=PAGE_ICON, layout="wide")

@st.cache_resource
def get_validator():
    if PlateValidator:
        try: return PlateValidator()
        except: return None
    return None

def valida_chave_nfe(chave: str) -> bool:
    if len(chave) != 44 or not chave.isdigit(): return False
    soma, peso = 0, 2
    for d in reversed(chave[:-1]):
        soma += int(d) * peso
        peso = 2 if peso >= 9 else peso + 1
    dv = 11 - (soma % 11)
    return str(0 if dv >= 10 else dv) == chave[-1]

def smart_parse_float(s):
    if not s or str(s).lower() == 'nan': return 0.0
    s = str(s).strip()
    if ',' in s and '.' in s:
        if s.rfind(',') > s.rfind('.'): return float(s.replace('.', '').replace(',', '.'))
        else: return float(s.replace(',', ''))
    if ',' in s: return float(s.replace(',', '.'))
    return float(s) if '.' in s else float(s)

def extract_details_from_obs(text):
    if not isinstance(text, str): return None, None, None
    text = text.upper()
    p_m = re.search(r'(?:MATRICULA|PREFIXO|AERONAVE|PLACA|VEICULO)[\s:.-]*([A-Z0-9-]{5,8})', text)
    if not p_m: p_m = re.search(r'([A-Z]{2,3}-?[0-9A-Z]{3,5})', text)
    t = p_m.group(1).replace('-', '') if p_m else None
    h_m = re.search(r'(?:HORAS|CICLOS|HORIMETRO|ODOMETRO|KM)[\s:.-]*(\d+)', text)
    return t, None, (h_m.group(1) if h_m else None)

def extrair_dados_nota_individual(texto_nota) -> dict:
    d = {"cnpj_posto": "", "filial": "", "documento": "", "data_emissao": "", "hora_emissao": "", "serie": "", "placa": "", "produto": "", "quantidade": "", "valor": "", "km": "", "itens": [], "valor_total_nota": "", "alertas": [], "chave_nfe": ""}
    
    cn = re.findall(r'\d{2}\.\d{3}\.\d{3}\/\d{4}\-\d{2}', texto_nota)
    c_b = next((c for c in cn if c.startswith(COMPANY_CNPJ_ROOTS)), None)
    if c_b:
        m_f = re.search(r'\/0*(\d+)\-', c_b)
        if m_f: d["filial"] = m_f.group(1)
            
    c_f = next((c for c in cn if not c.startswith(COMPANY_CNPJ_ROOTS)), None)
    if c_f: d["cnpj_posto"] = re.sub(r'[^0-9]', '', c_f)

    # Identifiers
    m_d = re.search(r'(?:N[ºo\.]*|MANIFESTO|FATURA)\s*([\d\.]+)', texto_nota, re.IGNORECASE)
    if m_d: d["documento"] = m_d.group(1).replace('.', '').lstrip('0') or '0'
    m_s = re.search(r'S[ÉE]RIE[^\d]*(\d+)', texto_nota, re.IGNORECASE)
    if m_s: d["serie"] = m_s.group(1).lstrip('0') or '0'

    # Date/Time
    m_p = re.search(r'PROTOCOLO.*?(\d{2}/\d{2}/\d{4})\s+(\d{2}:\d{2})', texto_nota, re.IGNORECASE | re.DOTALL)
    if m_p: d["data_emissao"], d["hora_emissao"] = m_p.group(1).replace('/', ''), m_p.group(2).replace(':', '')
    else:
        m_dt = re.search(r'(?:DATA|EMISSAO|ABASTECIMENTO)\s*[\r\n]*\s*(\d{2}[/-]\d{2}[/-]\d{4})', texto_nota, re.IGNORECASE)
        if m_dt: d["data_emissao"] = m_dt.group(1).replace('/', '').replace('-', '')
            
    # Tail Number
    m_t = re.search(r'(?:MATRICULA|PREFIXO|AERONAVE|PLACA)[:\s-]*([A-Z0-9-]{5,8})\b', texto_nota, re.IGNORECASE)
    if not m_t: m_t = re.search(r'\b([A-Z]{2,3}-?[0-9A-Z]{3,5})\b', texto_nota, re.IGNORECASE)
    if m_t:
        p_b = re.sub(r'[^A-Z0-9]', '', m_t.group(1).upper())
        d["placa"] = p_b
        val = get_validator()
        if val:
            res = val.validate_plate_fuzzy(p_b)
            if res.get('is_match'):
                d["placa"] = res['best_match']
                if res.get('driver'): d["motorista"] = res['driver']

    # Metrics
    m_h = re.search(r'\b(?:HORAS|CICLOS|HORIMETRO|KM)\s*[:=]?\s*([\d\.,]+)', texto_nota, re.IGNORECASE)
    if m_h: d["km"] = re.sub(r'[^\d]', '', m_h.group(1).split(',')[0].split('.')[0])
    
    # Access Key
    ck = re.search(r'CHAVE DE ACESSO[\s:]*([\d\s]{44,60})', texto_nota, re.IGNORECASE)
    if ck: d["chave_nfe"] = re.sub(r'\s', '', ck.group(1))

    # Analytical Item Extraction
    def parse_v(t):
        m = list(re.finditer(r'\s(L|LT|LTS|UN|GL|KG|M3|GAL)\s+((?:[\d\.,%]+\s*){3,12})', t, re.IGNORECASE))
        if not m: return None, None, ""
        nu = re.findall(r'[\d\.,]+', m[0].group(2))
        return (nu[0], (nu[3] if len(nu) >= 4 else nu[2]), m[0].group(0)) if len(nu) >= 3 else (None, None, "")

    tb = texto_nota 
    pf = sorted(list(re.finditer(r'(QAV|AVGAS|ARLA|DIESEL|ABASTECIMENTO|JET\s*A1)', texto_nota, re.IGNORECASE)), key=lambda x: x.start())
    it_c = {}
    for p in pf:
        np = p.group(1).upper()
        q, v, sm = parse_v(tb[p.start():p.start()+150])
        if q and v:
            nm = "QAV" if "QAV" in np or "JET" in np else "AVGAS"
            if nm not in it_c: it_c[nm] = {"nome": nm, "quantidade": q, "valor": v, "pos_t": p.start()}
            else:
                it_c[nm]["quantidade"] = f"{smart_parse_float(it_c[nm]['quantidade']) + smart_parse_float(q):.2f}".replace('.', ',')
                it_c[nm]["valor"] = f"{smart_parse_float(it_c[nm]['valor']) + smart_parse_float(v):.2f}".replace('.', ',')
            if sm: tb = tb[:p.start()] + tb[p.start():].replace(sm, " "*len(sm), 1)
    
    d["itens"] = sorted(list(it_c.values()), key=lambda k: k.get('pos_t', 0))
    mt = re.search(r'(?:V\.|VALOR)\s*TOTAL\s*(?:DA\s*NOTA)?(?:R\$?\s*)?([\d\.,]+)', texto_nota, re.IGNORECASE)
    if mt: d["valor_total_nota"] = mt.group(1).replace('.', '')
    return d

def open_pdf_bg(chave):
    if DanfeCrawler: DanfeCrawler().visualize_note(chave)

def executar_rpa(dados, mp):
    pyautogui.PAUSE = 0.1
    time.sleep(1)
    # Hardware injection routine
    for key, val in [('filial', 4), ('cnpj_posto', 1), ('documento', 1)]:
        pyautogui.write(str(dados[key]), interval=0.08); pyautogui.press('tab', presses=val, interval=0.1)
    pyautogui.write(f"{dados['data_emissao']}{dados.get('hora_emissao', '0000')}", interval=0.01); pyautogui.press('tab', presses=3, interval=0.1)
    pyautogui.write(dados['placa']); pyautogui.press('tab')
    if "Interativo" in mp: pyautogui.alert("Foque no Cockpit e clique OK.")
    pyautogui.press('tab'); pyautogui.write(dados['produto']); pyautogui.press('tab', presses=5)
    pyautogui.write(dados['quantidade']); pyautogui.press('tab', presses=2); pyautogui.write(dados['valor'])
    pyautogui.press('tab', presses=3); pyautogui.write(dados['km']); pyautogui.press('tab')

# --- UI LAYER ---
st.title("🚀 RPA AeroSys: Ingestão de Suprimentos")
for k in ['arquivos_pdf', 'fila_notas', 'indice_nota_atual']:
    if k not in st.session_state: st.session_state[k] = [] if 'indice' not in k else 0

up = st.file_uploader("1. Manifestos (PDF/Excel/CSV)", type=["pdf", "xlsx", "xls", "csv"], accept_multiple_files=True)

with st.expander("🛠️ Console de OCR & Recuperação"):
    ta = st.text_area("Bloco de Texto:", height=100)
    if st.button("🔎 Detectar Manifestos", use_container_width=True):
        f = re.findall(r'\b\d{44}\b', ta)
        b4 = re.finditer(r'(?<!\d)((?:\d{4}[\s\n]+){9}\d{4}|\d{40})(?!\d)', ta)
        for b in b4:
            nx = re.finditer(r'(?<=[\s\n])(\d{4})(?=[^\d]|$)', ta[b.end():b.end()+300])
            for fr in nx:
                ck = b.group(0).replace(" ", "").replace("\n", "") + fr.group(1)
                if valida_chave_nfe(ck): f.append(ck)
        for ch in set(f): st.session_state.fila_notas.append({"arquivo_origem": "Painel OCR", "texto": f"CHAVE {ch}", "is_image": True, "dados_pre_extraidos": {'origem': 'OCR', 'chave_nfe': ch, 'itens': [], 'placa': '', 'km': '', 'filial': '', 'documento': ''}})
        st.success(f"{len(set(f))} Documentos identificados.")

if up:
    if [a.name for a in up] != [a.name for a in st.session_state.arquivos_pdf]:
        st.session_state.arquivos_pdf, st.session_state.fila_notas, st.session_state.indice_nota_atual = up, [], 0
        for a in up:
            if a.name.lower().endswith(('.xlsx', '.xls', '.csv')):
                df = pd.read_csv(a, dtype=str) if a.name.endswith('.csv') else pd.read_excel(a, dtype=str)
                def gv(r, *ns):
                    for n in ns:
                        for c in df.columns:
                            if n.lower() in str(c).lower(): return r[c]
                    return ''
                for _, r in df.iterrows():
                    st.session_state.fila_notas.append({"arquivo_origem": a.name, "dados_pre_extraidos": {'origem': 'EXCEL', 'filial': gv(r, 'Base', 'FBO'), 'cnpj_posto': gv(r, 'Fornec'), 'documento': gv(r, 'Manifesto'), 'data_emissao': gv(r, 'Data'), 'placa': gv(r, 'Ma'), 'km': gv(r, 'Hor'), 'serie': '1', 'itens': [{'nome': gv(r, 'Prod'), 'quantidade': gv(r, 'Vol'), 'valor': gv(r, 'Tot')}]}})
            else:
                with pdfplumber.open(a) as pdf:
                    for p in pdf.pages:
                        t = p.extract_text()
                        if t: st.session_state.fila_notas.append({"arquivo_origem": a.name, "texto": t, "uid": str(uuid.uuid4())})

if st.session_state.fila_notas:
    idx = st.session_state.indice_nota_atual
    nt = st.session_state.fila_notas[idx]
    d = nt.get('dados_pre_extraidos') or extrair_dados_nota_individual(nt['texto'])
    nt['dados_pre_extraidos'] = d
    uk = str(uuid.uuid4())[:6]
    st.markdown("---")
    l, r = st.columns([3, 1])
    l.info(f"📁 **Manifesto {idx + 1} de {len(st.session_state.fila_notas)}** | `{nt['arquivo_origem']}`")
    if d.get('chave_nfe') and r.button("👁️ Abrir PDF", use_container_width=True): threading.Thread(target=open_pdf_bg, args=(d['chave_nfe'],), daemon=True).start(); st.toast("PDF enviado ao navegador.")
    cf, ci = st.columns([5, 1])
    with cf:
        c1, c2, c3, c4, c5 = st.columns([1, 1, 1, 1, 1])
        fb = st.text_input("FBO", d.get('filial', ''), key=f"f_{uk}")
        cp = st.text_input("CNPJ Forn.", d.get('cnpj_posto', ''), key=f"c_{uk}")
        mn = st.text_input("Manifesto", d.get('documento', ''), key=f"m_{uk}")
        da = st.text_input("Data", d.get('data_emissao', ''), key=f"d_{uk}")
        mt = st.text_input("Matrícula", d.get('placa', ''), key=f"a_{uk}")
        hr = st.text_input("Horas", d.get('km', ''), key=f"r_{uk}")
        it = d['itens'][0] if d.get('itens') else {'nome': 'QAV', 'quantidade': '', 'valor': ''}
        pn, vq, vt = st.text_input("Prop.", it['nome'], key=f"p_{uk}"), st.text_input("Vol.", it['quantidade'], key=f"q_{uk}"), st.text_input("Total", it['valor'], key=f"t_{uk}")
    with ci:
        with st.expander("👤 Log", expanded=True): st.text_input("Cmdte", d.get('motorista', ''), disabled=True)
    st.markdown("---")
    ms = st.radio("Sync:", ("Interativo", "Agressivo", "Lento"), horizontal=True)
    if st.button("🚀 INICIAR INGESTÃO", type="primary", use_container_width=True):
        st.warning("EM 5 SEGUNDOS...")
        time.sleep(5)
        try: executar_rpa({'filial': fb, 'cnpj_posto': cp, 'documento': mn, 'data_emissao': da, 'placa': mt, 'produto': pn, 'quantidade': vq, 'valor': vt, 'km': hr}, ms); st.success("OK.")
        except Exception as e: st.error(str(e))
    nv = st.columns([1, 1, 1])
    if idx > 0 and nv[0].button("⬅️ Ant."): st.session_state.indice_nota_atual -= 1; st.rerun()
    nv[1].markdown(f"<center>{idx+1}/{len(st.session_state.fila_notas)}</center>", unsafe_allow_html=True)
    if idx < len(st.session_state.fila_notas)-1 and nv[2].button("Próx. ➡️"): st.session_state.indice_nota_atual += 1; st.rerun()
