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

# Importa módulos internos
sys.path.append(os.path.join(os.path.dirname(__file__), 'src'))
try:
    from validator import PlateValidator
except ImportError:
    pass

try:
    from danfe_crawler import DanfeCrawler
except ImportError:
    DanfeCrawler = None

@st.cache_resource
def get_validator():
    try:
        return PlateValidator()
    except Exception as e:
        print(f"Erro ao instanciar PlateValidator: {e}")
        return None

def extract_details_from_obs(text):
    """Extrai Matrícula, Comandante e Horas de Voo do texto de Observação."""
    if not isinstance(text, str): return None, None, None
    text = text.upper()
    
    plate_match = re.search(r'(?:MATRICULA|PREFIXO|AERONAVE|PLACA|VEICULO)[\s:.-]*([A-Z]{3}-?[0-9][A-Z0-9][0-9]{2})', text)
    if not plate_match:
         plate_match = re.search(r'([A-Z]{3}-?[0-9][A-Z0-9][0-9]{2})', text)
    plate = plate_match.group(1).replace('-', '') if plate_match else None
    if plate_match and not plate:
         plate = plate_match.group(0).replace('-', '')

    driver = None # Extração de comandante via DB interno
    
    km = None
    km_match = re.search(r'(?:HORAS|CICLOS|HORIMETRO|ODOMETRO|KM)[\s:.-]*(\d+)', text)
    if km_match: km = km_match.group(1)

    return plate, driver, km

# Configuração da página
st.set_page_config(page_title="RPA AeroSys - Ingestão de Logs", page_icon="🚀", layout="wide")

def smart_parse_float(s):
    """Lógica robusta para converter strings numéricas de CSV/Excel (Ponto ou Vírgula)."""
    if not s or str(s).lower() == 'nan':
        return 0.0
    s = str(s).strip()
    
    if ',' in s and '.' in s:
        if s.rfind(',') > s.rfind('.'): 
            return float(s.replace('.', '').replace(',', '.'))
        else: 
            return float(s.replace(',', ''))
            
    if ',' in s:
        return float(s.replace(',', '.'))
    if '.' in s:
        return float(s)
        
    try:
        return float(s)
    except:
        return 0.0

def extrair_dados_nota_individual(texto_nota) -> dict:
    dados = {
        "cnpj_posto": "", "filial": "", "documento": "",
        "data_emissao": "", "hora_emissao": "", "serie": "",
        "placa": "", "produto": "",
        "quantidade": "", "valor": "", "km": "",
        "itens": [], "lista_valores": [], "valor_total_nota": "",
        "alertas": []
    }
    
    cnpjs = re.findall(r'\d{2}\.\d{3}\.\d{3}\/\d{4}\-\d{2}', texto_nota)
    
    # CNPJs base da operação
    raizes_empresas = ('******', '******')
    
    cnpj_empresa = next((c for c in cnpjs if c.startswith(raizes_empresas)), None)
    if cnpj_empresa:
        match_filial = re.search(r'\/0*(\d+)\-', cnpj_empresa)
        if match_filial: dados["filial"] = match_filial.group(1)
            
    cnpj_posto = next((c for c in cnpjs if not c.startswith(raizes_empresas)), None)
    if cnpj_posto:
        dados["cnpj_posto"] = re.sub(r'[^0-9]', '', cnpj_posto)

    nota_match = re.search(r'N[ºo\.]*\s*([\d\.]+)', texto_nota)
    if nota_match:
        nota_raw = nota_match.group(1).replace('.', '')
        dados["documento"] = nota_raw.lstrip('0') or '0'
    
    serie_match = re.search(r'S[ÉE]RIE[^\d]*(\d+)', texto_nota, re.IGNORECASE)
    if serie_match:
        dados["serie"] = serie_match.group(1).lstrip('0') or '0'

    protocolo_match = re.search(r'PROTOCOLO DE AUTORIZA[ÇC][ÃA]O DE USO\n?.*?(\d{2}/\d{2}/\d{4})\s+(\d{2}:\d{2})', texto_nota, re.IGNORECASE | re.DOTALL)
    if protocolo_match:
        dados["data_emissao"] = protocolo_match.group(1).replace('/', '') 
        dados["hora_emissao"] = protocolo_match.group(2).replace(':', '') 
    else:
        data_match = re.search(r'(?:DATA\s+DA\s+EMISSÃO|Data\s+Emissão)\s*[\r\n]*\s*(\d{2}[/-]\d{2}[/-]\d{4})', texto_nota, re.IGNORECASE)
        if not data_match:
            data_match = re.search(r'(?<!Impresso em:\s)(?<!Impresso em:)(\d{2}[/-]\d{2}[/-]\d{4})', texto_nota, re.IGNORECASE)
            
        if data_match:
            dados["data_emissao"] = data_match.group(1).replace('/', '')
            
    placa_match = re.search(r'(?:MATRICULA|PREFIXO|AERONAVE|PLACA)[:\s-]*([A-Z0-9]{7,8})\b', texto_nota, re.IGNORECASE)
    if not placa_match:
        placa_match = re.search(r'\b([A-Z]{3}[0-9O][A-Z0-9][0-9O]{2})\b', texto_nota, re.IGNORECASE)
        
    if placa_match:
        placa_bruta = re.sub(r'[^A-Z0-9]', '', placa_match.group(1).upper())
        if len(placa_bruta) == 7 and placa_bruta[3] == 'O': 
            placa_bruta = placa_bruta[:3] + '0' + placa_bruta[4:]
        dados["placa"] = placa_bruta
        
        val = get_validator()
        if val:
            try:
                res = val.validate_plate_fuzzy(placa_bruta)
                if res.get('is_match') and res.get('best_match'):
                    dados["placa"] = res['best_match']
                    if res.get('driver'):
                        dados["comandante"] = res['driver']
            except Exception as e:
                print(f"Erro Fuzzy Matrícula: {e}")

    chave_match = re.search(r'CHAVE DE ACESSO[\s:]*([\d\s]{44,60})', texto_nota, re.IGNORECASE)
    if chave_match:
        dados["chave_nfe"] = re.sub(r'\s', '', chave_match.group(1))
    else:
        dados["chave_nfe"] = ""

    km_match = re.search(r'\b(?:OD[OÓ]METRO|HOD[OÓ]METRO|OD[OÓ]M|HOD)\s*[:=]?\s*([\d\.,]+)', texto_nota, re.IGNORECASE)
    if not km_match:
        km_match = re.search(r'\bKM\s*[:=-]?\s*(\d{4,7})\b', texto_nota, re.IGNORECASE)
        if not km_match:
             km_match = re.search(r'\bKM\s*[:=]\s*([\d\.,]+)', texto_nota, re.IGNORECASE)
        
    if km_match:
        km_limpo = re.sub(r'[^\d]', '', km_match.group(1).split(',')[0].split('.')[0])
        dados["km"] = km_limpo
    elif dados.get("placa"):
        placa_bruta_pesquisa = re.search(r'\b([A-Z]{3}[0-9O][A-Z0-9][0-9O]{2})\b', texto_nota, re.IGNORECASE)
        if placa_bruta_pesquisa:
            grid_km_match = re.search(rf'{placa_bruta_pesquisa.group(1)}\s+(\d{{4,7}})\b', texto_nota, re.IGNORECASE)
            if grid_km_match:
                dados["km"] = grid_km_match.group(1)
                
    if "comandante" not in dados:
        dados["comandante"] = ""

    dados["itens"] = []
    posicoes_usadas = set()

    def extract_product_values(contexto_txt):
        matches = list(re.finditer(r'\s(L|LT|LTS|UN|KG|GL|M3|LI|L\s*I|L\.I\.?|LITROS?|CX|PC|PCT)\s+((?:[\d\.,%]+\s*){3,12})', contexto_txt, re.IGNORECASE))
        if not matches: return None, None, "", "", ""
            
        melhor_match = None
        melhor_nums = []
        
        for v_match in matches:
            temp_nums = re.findall(r'[\d\.,]+', v_match.group(2))
            if not temp_nums or len(temp_nums) < 3: continue
                
            tem_ncm = len(temp_nums[0]) >= 7 and ',' not in temp_nums[0]
            virgulas = sum(1 for n in temp_nums[:4] if ',' in n)
            
            if melhor_match is None or (not tem_ncm and virgulas >= 1):
                melhor_match = v_match
                melhor_nums = temp_nums
                if not tem_ncm and virgulas >= 2: break
                    
        if not melhor_match or not melhor_nums or len(melhor_nums) < 3:
            return None, None, "", "", ""
            
        v_match = melhor_match
        nums = melhor_nums
            
        q_raw = nums[0].replace('.', '')
        if len(nums) >= 4:
            v_raw = nums[3].replace('.', '')
            if v_raw in ['0,00', '0'] and nums[2] not in ['0,00', '0']:
                v_raw = nums[2].replace('.', '')
        else:
            v_raw = nums[2].replace('.', '')
            
        aliq_icms = ""
        if len(nums) >= 7:
            valid_rates = ['12,00', '17,00', '18,00', '19,50', '19,00', '20,00']
            for idx in [-1, -2, -3, -4]:
                if len(nums) >= abs(idx):
                    val_c = str(nums[idx]).replace('.', ',')
                    if val_c in valid_rates:
                        aliq_icms = val_c
                        break
            
            if not aliq_icms:
                 for idx in [-1, -2, -3]:
                      if len(nums) >= abs(idx):
                          if str(nums[idx]).replace('.', ',') == '0,00':
                              aliq_icms = '0,00'
                              break
            
            if not aliq_icms and '%' in v_match.group(2):
                 val_c2 = str(nums[-2]).replace('.', ',')
                 if val_c2.startswith(('12', '17', '18', '19', '20')):
                     aliq_icms = val_c2

        if ',' in q_raw:
            qi, qd = q_raw.split(',')
            q_raw = f"{qi},{qd[:3]}"
            if q_raw.endswith(',000'): q_raw = q_raw[:-3] + '00'
            
        if ',' in v_raw:
            vi, vd = v_raw.split(',')
            v_raw = f"{vi},{vd[:2]}"
            
        return q_raw, v_raw, aliq_icms, v_match.group(0), v_match.group(1).upper().strip()

    texto_busca = texto_nota 
    produtos_encontrados = list(re.finditer(r'(QAV|AVGAS|JET\s*A1)', texto_nota, re.IGNORECASE))
    produtos_encontrados.sort(key=lambda m: m.start())
    
    for m in produtos_encontrados:
        termo_original = m.group(1).upper()
        pos_busca = m.start()
            
        context = texto_busca[pos_busca:min(len(texto_busca), pos_busca + 150)]
        q_raw, v_raw, a_icms, str_matched, unit_matched = extract_product_values(context)
        
        if q_raw and v_raw:
            if 'AVGAS' in termo_original and unit_matched in ['UN', 'GL', 'CX', 'PC', 'PCT']:
                desc_match = re.search(r'\b(\d{1,3})\s*(?:L|LT|LTS|LITRO|LITROS)\b', context, re.IGNORECASE)
                if desc_match:
                    try:
                        multiplicador = int(desc_match.group(1))
                        q_float = float(q_raw.replace(',', '.'))
                        q_litros = q_float * multiplicador
                        q_raw = f"{q_litros:.2f}".replace('.', ',')
                    except ValueError:
                        pass
                        
            dados["itens"].append({
                "produto_cod": "8" if 'QAV' in termo_original or 'JET' in termo_original else "6",
                "nome": "QAV" if 'QAV' in termo_original or 'JET' in termo_original else "AVGAS",
                "quantidade": q_raw,
                "valor": v_raw,
                "aliq_icms": a_icms,
                "posicao_texto": pos_busca
            })
            
            if str_matched:
                mascara = " " * len(str_matched)
                parte_frente = texto_busca[pos_busca:].replace(str_matched, mascara, 1)
                texto_busca = texto_busca[:pos_busca] + parte_frente
         
    dados["itens"] = sorted(dados["itens"], key=lambda k: int(k.get('posicao_texto', 0)))
         
    itens_consolidados = {}
    for item in dados["itens"]:
        nome = item["nome"]
        if nome not in itens_consolidados:
            itens_consolidados[nome] = item.copy()
        else:
            qtde_atual = float(itens_consolidados[nome]["quantidade"].replace('.', '').replace(',', '.'))
            qtde_nova = float(item["quantidade"].replace('.', '').replace(',', '.'))
            nova_qtde = qtde_atual + qtde_nova
            
            qtde_str = f"{nova_qtde:.3f}".replace('.', ',')
            if qtde_str.endswith(',000'): qtde_str = qtde_str[:-3] + ',00'
            itens_consolidados[nome]["quantidade"] = qtde_str
            
            alerta_msg = f"⚠️ Atenção: Múltiplos lançamentos de '{nome}' foram detectados e somados neste manifesto."
            alerta_lista = dados.get("alertas", [])
            if not isinstance(alerta_lista, list): alerta_lista = []
            if alerta_msg not in alerta_lista:
                alerta_lista.append(alerta_msg)
                dados["alertas"] = alerta_lista
            
            valor_atual = float(itens_consolidados[nome]["valor"].replace('.', '').replace(',', '.'))
            valor_novo = float(item["valor"].replace('.', '').replace(',', '.'))
            novo_valor = valor_atual + valor_novo
            
            itens_consolidados[nome]["valor"] = f"{novo_valor:.2f}".replace('.', ',')
            
    dados["itens"] = list(itens_consolidados.values())
         
    dados["valor_total_nota"] = ""
    match_tabela = re.search(r'(?:V\.|VALOR)\s*TOTAL\s*DA\s*NOTA[^\n]*\n([^\n]+)', texto_nota, re.IGNORECASE)
    if match_tabela:
        linha_valores = match_tabela.group(1)
        numeros = re.findall(r'[\d\.,]+', linha_valores)
        if numeros: dados["valor_total_nota"] = numeros[-1].replace('.', '')
            
    if not dados["valor_total_nota"]:
        match_exato = re.search(r'(?:V\.|VALOR)\s*TOTAL\s*DA\s*NOTA[^\d\n]*(?:R\$?\s*)?([\d\.,]+)', texto_nota, re.IGNORECASE)
        if match_exato: dados["valor_total_nota"] = match_exato.group(1).replace('.', '')
             
    if not dados["valor_total_nota"]:
         match_generico = re.search(r'Valor\s*Total[\s:]*(?:dos\s*produtos\s*)?(?:R\$?\s*)?([\d\.,]+)', texto_nota, re.IGNORECASE)
         if match_generico: dados["valor_total_nota"] = match_generico.group(1).replace('.', '')
        
    if dados["itens"] and dados["valor_total_nota"]:
        try:
            total_produtos = sum(float(item["valor"].replace('.', '').replace(',', '.')) for item in dados["itens"])
            total_nota = float(dados["valor_total_nota"].replace('.', '').replace(',', '.'))
            
            if total_produtos > total_nota + 0.01:
                desconto = total_produtos - total_nota
                item_para_desconto = next((item for item in dados["itens"] if item["nome"] == "QAV"), None)
                if not item_para_desconto:
                    item_para_desconto = next((item for item in dados["itens"] if item["nome"] == "AVGAS"), None)
                    
                if item_para_desconto:
                    valor_atual_produto = float(item_para_desconto["valor"].replace('.', '').replace(',', '.'))
                    novo_valor_produto = valor_atual_produto - desconto
                    if novo_valor_produto < 0: novo_valor_produto = 0
                         
                    item_para_desconto["valor"] = f"{novo_valor_produto:.2f}".replace('.', ',')
                    alerta_msg = f"💡 Desconto de R$ {desconto:.2f} detectado no manifesto e deduzido do item {item_para_desconto['nome']} para nivelar o total."
                    alerta_lista = dados.get("alertas", [])
                    if alerta_msg not in alerta_lista:
                        alerta_lista.append(alerta_msg)
                        dados["alertas"] = alerta_lista
                        
        except Exception as e:
            print(f"Erro ao calcular desconto: {e}")

    return dados

def executar_rpa_aerosys(dados, buscar_posto_f2, modo_popup):
    """Executa a sequência de teclado exata na máquina virtual."""
    pyautogui.PAUSE = 0.1 
    
    time.sleep(0.5) 
    pyautogui.write(str(dados['filial']), interval=0.08)
    pyautogui.press('tab', presses=4, interval=0.1)
    
    if buscar_posto_f2:
        pyautogui.press('f2')
        time.sleep(1.5) 
        pyautogui.press('tab', presses=2, interval=0.1)
        
        cnpj = dados['cnpj_posto']
        if len(cnpj) == 14:
            pyautogui.write(cnpj[:-2])
            time.sleep(0.3)
            pyautogui.write(cnpj[-2:])
        else:
            pyautogui.write(cnpj) 
            
        pyautogui.press('tab', presses=6, interval=0.1)
        pyautogui.press('enter')
        time.sleep(1) 
        pyautogui.press('tab')
        pyautogui.press('enter')
    else:
        codigo_posto = str(dados.get('cnpj_posto', '')).strip()
        pyautogui.write(codigo_posto, interval=0.08)
        
    pyautogui.press('tab', interval=0.1)
    pyautogui.write(dados['documento'])
    pyautogui.press('tab', interval=0.1)
    
    dt_hora = dados['data_emissao']
    hora = dados.get('hora_emissao', '').strip()
    if not hora: hora = "0000"
    dt_hora = f"{dt_hora}{hora}"
        
    pyautogui.write(dt_hora, interval=0.01) 
    pyautogui.press('tab', presses=3, interval=0.1) 
    pyautogui.write(dados['placa'])
    pyautogui.press('tab', interval=0.1) 
    
    if modo_popup == "Modo Agressivo (Rápido) - Dar ENTER/ESPAÇO para tentar fechar tudo automaticamente":
        time.sleep(1.5)
        pyautogui.press('space') 
        time.sleep(0.3)
        pyautogui.press('space') 
        time.sleep(0.3)
        pyautogui.press('space') 
    elif modo_popup == "Pausar 4 segundos (Lento) - Sem perguntas, apenas pausa":
        time.sleep(4)
    elif modo_popup == "Modo Interativo (Perfeito) - O robô pausa, eu fecho o popup, e dou OK pro robô continuar":
        try:
            import tkinter as tk
            from tkinter import messagebox
            root = tk.Tk()
            root.attributes("-topmost", True)
            root.withdraw()
            resposta = messagebox.askokcancel("🛑 Robô Pausado", "1. Feche os avisos do ERP VDI (se houver).\n\n2. Clique em OK para continuar a Ingestão.\n\n(Ou Cancelar se houver falha grave)")
            root.quit()
            root.destroy()
            if not resposta:
                 raise Exception("Ingestão abortada pelo usuário após preencher a Matrícula.")
        except Exception as e:
             if str(e).startswith("Ingestão abortada"):
                 raise e
             escolha = pyautogui.confirm(
                 text="O robô pausou!\n\n1. Resolva/Feche os alertas no ERP.\n2. Depois, clique em Continuar.",
                 title="🛑 Robô Esperando seu OK",
                 buttons=['Continuar Ingestão', 'Parar Tudo (Falha Grave)']
             )
             if escolha == 'Parar Tudo (Falha Grave)':
                 raise Exception("Ingestão abortada pelo usuário durante verificação de Matrícula.")
             
        time.sleep(0.5)

    time.sleep(0.5) 
    pyautogui.press('tab', interval=0.1) 
    
    pyautogui.write(dados['produto'])
    pyautogui.press('tab', presses=5, interval=0.1)
    pyautogui.write(dados['quantidade'])
    pyautogui.press('tab', presses=2, interval=0.1)
    pyautogui.write(dados['valor'])
    pyautogui.press('tab', presses=3, interval=0.1)
    
    pyautogui.write(dados['km'])
    pyautogui.press('tab', interval=0.1) 
    time.sleep(0.5)


# ===================== UI STREAMLIT =====================
st.title("🚀 RPA AeroSys: Ingestão Automática em Lote")

if 'arquivos_pdf' not in st.session_state: st.session_state.arquivos_pdf = []
if 'fila_notas' not in st.session_state: st.session_state.fila_notas = []
if 'indice_nota_atual' not in st.session_state: st.session_state.indice_nota_atual = 0
if 'rejeicoes_excel' not in st.session_state: st.session_state.rejeicoes_excel = []

arquivos_upados = st.file_uploader("1. Faça o Upload dos arquivos PDF, Excel ou CSV aqui", type=["pdf", "xlsx", "xls", "csv"], accept_multiple_files=True)

with st.expander("📸 Tem Manifestos que são só imagem escaneada? Clique aqui para processar em Lote!"):
    st.info("Abra seus PDFs de imagem, aperte **Ctrl+A** (selecionar tudo) e **Ctrl+C** (copiar), e cole o texto todo aqui.")
    texto_colado = st.text_area("Cole todo o texto solto aqui:", height=150)
    if st.button("Buscar Chaves NFe no Texto Colado", use_container_width=True):
        if texto_colado:
            chaves_unicas = []
            def valida_chave_nfe(chave: str) -> bool:
                if len(chave) != 44 or not chave.isdigit(): return False
                soma = 0
                peso = 2
                for digito in reversed(chave[:-1]):
                    soma += int(digito) * peso
                    peso += 1
                    if peso > 9: peso = 2
                resto = soma % 11
                dv = 11 - resto
                if dv >= 10: dv = 0
                return str(dv) == chave[-1]

            texto_limpo = re.sub(r'[\s\.\-]', '', texto_colado)
            chaves = re.findall(r'\b\d{44}\b', texto_colado)
            if not chaves: chaves = re.findall(r'(?=(?:\D|^)(\d{44})(?:\D|$))', texto_limpo)
            for ch in chaves:
                if valida_chave_nfe(ch) and ch not in chaves_unicas: chaves_unicas.append(ch)
            
            blocos_40 = re.finditer(r'(?<!\d)((?:\d{4}[\s\n]+){9}\d{4}|\d{40})(?!\d)', texto_colado)
            for b in blocos_40:
                start, end = b.span()
                chunk_40 = re.sub(r'\D', '', b.group(0))
                if len(chunk_40) == 40:
                    texto_frente = texto_colado[end:end+400]
                    cands_4 = re.finditer(r'(?<=[\s\n])(\d{4})(?=[^\d]|$)', texto_frente)
                    for c in cands_4:
                        chave_teste = chunk_40 + c.group(1)
                        if valida_chave_nfe(chave_teste):
                            if chave_teste not in chaves_unicas: chaves_unicas.append(chave_teste)
                            break 

            if chaves_unicas:
                for ch in chaves_unicas:
                    st.session_state.fila_notas.append({
                        "arquivo_origem": f"Texto Colado ({ch[:6]}...)",
                        "texto": f"CHAVE DE ACESSO {ch}",
                        "is_image": True,
                        "dados_pre_extraidos": {
                            'origem': 'PDF_IMAGEM_LOTE', 'chave_nfe': ch,
                            'itens': [], 'filial': '', 'cnpj_posto': '', 'documento': '',
                            'data_emissao': '', 'hora_emissao': '', 'serie': '', 'placa': '',
                            'comandante': '', 'km': '', 'valor_total_nota': '', 'alertas': []
                        }
                    })
                st.success(f"✅ {len(chaves_unicas)} Chaves NFe válidas processadas!")
            else:
                st.error("Nenhuma chave válida encontrada no texto colado.")
        else:
            st.warning("Cole algum texto antes de buscar.")

if arquivos_upados:
    nomes_atuais = [f.name for f in arquivos_upados]
    nomes_sessao = [f.name for f in st.session_state.arquivos_pdf]
    if set(nomes_atuais) != set(nomes_sessao):
        st.session_state.arquivos_pdf = arquivos_upados
        st.session_state.fila_notas = []
        st.session_state.indice_nota_atual = 0
        st.session_state.rejeicoes_excel = []
        
        with st.spinner("Estruturando pipeline de ingestão de dados..."):
            for arquivo in arquivos_upados:
                name_lower = arquivo.name.lower()
                if name_lower.endswith(('.xlsx', '.xls', '.csv')):
                    try:
                        if name_lower.endswith('.csv'):
                            pos = arquivo.tell()
                            primeira_linha = arquivo.readline().decode('utf-8', errors='ignore')
                            arquivo.seek(pos)
                            separador = ';' if primeira_linha.count(';') > primeira_linha.count(',') else ','
                            df = pd.read_csv(arquivo, dtype=str, sep=separador, encoding='utf-8')
                        else:
                            df = pd.read_excel(arquivo, dtype=str)
                            
                        notas_agrupadas_excel = {}
                        for _, row in df.iterrows():
                            def get_val(*possible_names):
                                for name in possible_names:
                                    candidates = [name, name.replace('.', ''), name.replace(' ', ''), name.lower()]
                                    for cand in candidates:
                                        for df_col in df.columns:
                                            if str(df_col).lower().strip() == cand.lower().strip(): return row[df_col]
                                    for cand in candidates:
                                        for df_col in df.columns:
                                            if cand.lower().strip() in str(df_col).lower().strip(): return row[df_col]
                                return ''
                                
                            produto_raw = str(get_val('Produto', 'Prod')).upper()
                            if "QAV" not in produto_raw and "AVGAS" not in produto_raw:
                                st.session_state.rejeicoes_excel.append(f"Ignorado (Combustível Inválido): {produto_raw} | Matrícula: {get_val('Placa', 'Veiculo', 'Matricula')}")
                                continue 
                                
                            raw_plate = str(get_val('Placa', 'Veiculo', 'Matricula')).strip()
                            obs_text = str(get_val('Observ NF', 'Observ. NF', 'Observações', 'Obs'))
                            ext_plate, ext_driver, ext_km = extract_details_from_obs(obs_text)
                            
                            final_plate = None
                            comandante_db = ""
                            val_sys = get_validator()
                            
                            if val_sys:
                                val1 = val_sys.validate_plate_fuzzy(raw_plate)
                                if val1.get('is_match'):
                                    final_plate = val1['best_match']
                                    comandante_db = val1.get('driver', '')
                                elif ext_plate:
                                    val2 = val_sys.validate_plate_fuzzy(ext_plate)
                                    if val2.get('is_match'):
                                        final_plate = val2['best_match']
                                        comandante_db = val2.get('driver', '')
                                        
                            if not final_plate:
                                prt_plate = ext_plate if ext_plate else raw_plate
                                st.session_state.rejeicoes_excel.append(f"Ignorado (Matrícula não cadastrada): {prt_plate} | Obs: {obs_text}")
                                continue 
                                
                            n_doc = str(get_val('NF', 'Numero', 'Fatura', 'Manifesto'))
                            if not n_doc or n_doc == 'nan': n_doc = '1-1-1'
                            partes_doc = n_doc.split('-')
                            documento = partes_doc[0] if len(partes_doc) > 0 else n_doc
                            documento = documento.lstrip('0') or '0'
                            
                            serie_val = str(get_val('Série', 'Serie')).strip()
                            if serie_val in ['nan', 'None', '', '0.0', '0']: serie_val = '001'
                            serie = partes_doc[1] if len(partes_doc) > 1 else serie_val
                            serie = str(serie).split('.')[0].zfill(3) 
                            
                            data_raw = str(get_val('Emissão', 'Data Emissão', 'Data')).strip()
                            match_iso = re.search(r'(\d{4})-(\d{2})-(\d{2})', data_raw)
                            match_br = re.search(r'(\d{2})[/-](\d{2})[/-](\d{4})', data_raw)
                            if match_iso:
                                data_em = f"{match_iso.group(3)}{match_iso.group(2)}{match_iso.group(1)}"
                            elif match_br:
                                data_em = f"{match_br.group(1)}{match_br.group(2)}{match_br.group(3)}"
                            else:
                                data_em = ''.join([c for c in data_raw[:10] if c.isdigit()])
                                
                            match_hora = re.search(r'(\d{2}):?(\d{2})', data_raw[10:])
                            if match_hora:
                                hora_em = f"{match_hora.group(1)}{match_hora.group(2)}"
                            else:
                                hora_em = ''.join([c for c in data_raw[10:] if c.isdigit()])[:4]
                            
                            produto = "QAV" if "QAV" in produto_raw else "AVGAS"
                            prod_cod = "6" if "AVGAS" in produto else "8"
                            
                            km_val = ext_km if ext_km else str(get_val('Km atual', 'KM', 'Horas', 'Ciclos'))
                            
                            raw_chave_obj = str(get_val('Chave', 'Chave NFe', 'Chave de Acesso', 'Danfe', 'Chave Nfe')).strip()
                            if 'E+' in raw_chave_obj or 'e+' in raw_chave_obj:
                                try:
                                    chave_nfe_raw = "{:.0f}".format(float(raw_chave_obj))
                                except:
                                    chave_nfe_raw = raw_chave_obj
                            else:
                                chave_nfe_raw = raw_chave_obj
                                
                            chave_nfe = ''.join([c for c in str(chave_nfe_raw) if c.isdigit()])
                            chave_grupo = chave_nfe if chave_nfe else str(documento).strip()
                            if not chave_grupo: chave_grupo = f"linha_{_}"
                                
                            quantidade_str = str(get_val('Quant', 'Quant.', 'Quantidade', 'Volume')).replace('.', ',')
                            if "AVGAS" in produto:
                                desc_match = re.search(r'\b(\d{1,3})\s*(?:L|LT|LTS|LITRO|LITROS)\b', produto_raw, re.IGNORECASE)
                                if desc_match:
                                    try:
                                        multiplicador = int(desc_match.group(1))
                                        q_float = smart_parse_float(quantidade_str)
                                        q_litros = q_float * multiplicador
                                        quantidade_str = f"{q_litros:.3f}".replace('.', ',')
                                        if quantidade_str.endswith(',000'): 
                                            quantidade_str = quantidade_str[:-3] + ',00'
                                    except ValueError:
                                        pass

                            valor_bruto_str = str(get_val('Vlr. Prod.', 'ValorNF', 'Valor NF', 'Valor', 'Total')).replace('.', '').replace(',', '.')
                            desconto_str = str(get_val('Vlr.Desc.', 'Vlr Desc', 'Vlr. Desc.', 'Desconto', 'Desc', 'Desc.')).replace('.', '').replace(',', '.')

                            try:
                                valor_bruto = smart_parse_float(valor_bruto_str)
                                desconto_val = smart_parse_float(desconto_str)
                                valor_final = valor_bruto - desconto_val
                                valor_final_str = f"{valor_final:.2f}".replace('.', ',')
                            except Exception:
                                valor_final_str = valor_bruto_str.replace('.', ',')

                            item_atual = {
                                "produto_cod": prod_cod,
                                "nome": produto,
                                "quantidade": quantidade_str,
                                "valor": valor_final_str,
                                "aliq_icms": ""
                            }
                            
                            if chave_grupo not in notas_agrupadas_excel:
                                alerta_excel = []
                                placa_limpa_raw = raw_plate.replace('-', '').upper()
                                if raw_plate and final_plate and placa_limpa_raw != final_plate.replace('-', '').upper():
                                    alerta_excel.append(f"🛠️ Matrícula corrigida via IA (Original: '{raw_plate}' ➔ Assumida: '{final_plate}')")
                                    
                                d_pre = {
                                    "origem": "EXCEL",
                                    "chave_nfe": chave_nfe,
                                    "cnpj_posto": str(get_val('Cód.Posto', 'CódPosto', 'Código do Posto', 'Fornecedor')).strip(),
                                    "filial": str(get_val('Cód. Filial (Dest.)', 'Filial', 'Base')).strip(),
                                    "documento": str(documento).strip(),
                                    "data_emissao": data_em.strip(),
                                    "hora_emissao": hora_em.strip(),
                                    "serie": str(serie).strip(),
                                    "placa": final_plate,
                                    "comandante": comandante_db,
                                    "produto": produto,
                                    "km": km_val.replace('.0', '').strip(),
                                    "alertas": alerta_excel,
                                    "itens": [item_atual],
                                    "lista_valores": [],
                                    "valor_total_nota": f"{smart_parse_float(get_val('Valor NF', 'Total', 'Valor', 'ValorNF')):.2f}".replace('.', ',')
                                }
                                
                                notas_agrupadas_excel[chave_grupo] = {
                                    "arquivo_origem": arquivo.name,
                                    "dados_pre_extraidos": d_pre
                                }
                            else:
                                notas_agrupadas_excel[chave_grupo]["dados_pre_extraidos"]["itens"].append(item_atual)
                                
                        for nota_dict in notas_agrupadas_excel.values():
                            st.session_state.fila_notas.append(nota_dict)
                            
                    except Exception as e:
                        st.error(f"Erro ao processar base de dados {arquivo.name}: {e}")
                else:
                    try:
                        with pdfplumber.open(arquivo) as pdf:
                            agrupadas_arquivo = {}
                            for page in pdf.pages:
                                txt = page.extract_text()
                                if not txt: continue
                                
                                chave_match = re.search(r'CHAVE DE ACESSO[\s:]*([\d\s]{44,60})', txt, re.IGNORECASE)
                                if chave_match:
                                    chave = re.sub(r'\s', '', chave_match.group(1))
                                else:
                                    nota_match = re.search(r'N[ºo\.]*\s*([\d\.]+)', txt)
                                    if nota_match:
                                        chave = nota_match.group(1).replace('.', '').lstrip('0')
                                    else:
                                        chave = 'continuação'
                                        
                                if chave == 'continuação' and agrupadas_arquivo:
                                    uid = list(agrupadas_arquivo.keys())[-1]
                                    agrupadas_arquivo[uid] += "\n" + str(txt)
                                else:
                                    if chave not in agrupadas_arquivo: agrupadas_arquivo[chave] = ""
                                    agrupadas_arquivo[chave] += "\n" + str(txt)
                                    
                            for texto_nota in agrupadas_arquivo.values():
                                if len(texto_nota.strip()) > 50:
                                    st.session_state.fila_notas.append({
                                        "uid": str(uuid.uuid4()),
                                        "arquivo_origem": arquivo.name,
                                        "texto": texto_nota,
                                        "is_image": False
                                    })
                    except Exception as e:
                        st.error(f"Erro na extração de texto do artefato {arquivo.name}: {e}")
                        
        if st.session_state.rejeicoes_excel:
            with st.expander(f"⚠️ {len(st.session_state.rejeicoes_excel)} registros incompatíveis ignorados"):
                for rej in st.session_state.rejeicoes_excel: st.write(f"- {rej}")

st.markdown("---")

if st.session_state.fila_notas:
    total_notas = len(st.session_state.fila_notas)
    indice_tela = st.session_state.indice_nota_atual
    nota_atual = st.session_state.fila_notas[indice_tela]
    
    st.info(f"📁 **Fila de Processamento:** Analisando artefato **{indice_tela + 1} de {total_notas}** | Source: `{nota_atual['arquivo_origem']}`")
    
    d = nota_atual.get('dados_pre_extraidos')
    if not d:
        if nota_atual.get('is_image'):
            d = {
                'origem': 'PDF_IMAGEM', 'chave_nfe': '', 'itens': [], 'filial': '', 'cnpj_posto': '', 'documento': '',
                'data_emissao': '', 'hora_emissao': '', 'serie': '', 'placa': '',
                'comandante': '', 'km': '', 'valor_total_nota': '', 'alertas': []
            }
        else:
            d = extrair_dados_nota_individual(nota_atual['texto'])
            d['origem'] = 'PDF'
        nota_atual['dados_pre_extraidos'] = d
            
    if 'uid' not in nota_atual: nota_atual['uid'] = str(uuid.uuid4())
    nota_uid = nota_atual['uid']
    
    st.markdown("---")
    
    col_hdr1, col_hdr2 = st.columns([3, 1])
    with col_hdr1:
        st.subheader("Auditoria de Dados do Voo")
        st.info("Painel interativo para intervenção manual em caso de falha na inferência.")
        for alerta in d.get("alertas", []): st.warning(alerta, icon="⚠️")
        
    with col_hdr2:
        if d.get('chave_nfe') and len(d['chave_nfe']) == 44:
            if st.button("👁️ Abrir PDF Nativo", use_container_width=True):
                if DanfeCrawler:
                    try:
                        import threading
                        def _abrir_pdf_bg(chave):
                            try: DanfeCrawler().visualize_note(chave)
                            except: pass
                        threading.Thread(target=_abrir_pdf_bg, args=(d['chave_nfe'],), daemon=True).start()
                        st.toast("Enviado ao visualizador web.", icon="✅")
                    except: pass
                else: st.error("Módulo DanfeCrawler ausente.")
        elif d['origem'] == 'EXCEL':
             st.warning("Sem Chave Válida (Fonte: DB/XLSX)")
    
    if d['itens']:
        if len(d['itens']) > 1:
            st.warning("🚨 **Atenção:** Manifesto com múltiplos propelentes detectados. Ingestão deve ser sequencial.", icon="🚨")
            
        nomes_itens = [f"Item {i+1} ({item['nome']} - {item['quantidade']}L)" for i, item in enumerate(d['itens'])]
        item_selecionado = st.selectbox("Selecione a carga para ingestão atual:", nomes_itens, key=f"sb_item_{nota_uid}")
        index_item = nomes_itens.index(item_selecionado)
        
        dados_item_atual = d['itens'][index_item]
        sugestao_produto = dados_item_atual['produto_cod']
        sugestao_qtde = dados_item_atual['quantidade']
        sugestao_valor = dados_item_atual['valor']
        sugestao_aliq = dados_item_atual.get('aliq_icms', '')
        
        num_item_doc = index_item + 1
        serie_doc = str(d.get('serie', '001')).split('.')[0].zfill(3)
        base_doc = d['documento'].split('-')[0] 
        sugestao_doc = f"{base_doc}-{serie_doc}-{num_item_doc}"
    else:
        index_item = 0
        if d.get('origem') == 'PDF_IMAGEM':
            st.error("📸 **Artefato Scaneado:** Texto de máquina não encontrado.", icon="🚨")
            chave_manual = st.text_input("🔑 Inserir chave de acesso manualmente:", value=d.get('chave_nfe', ''), max_chars=44)
            d['chave_nfe'] = chave_manual.strip()
        else:
            st.warning("⚠️ Algoritmo não classificou propelente. Intervenção exigida.")
            
        sugestao_produto, sugestao_qtde, sugestao_valor = "", "", d.get('valor_total_nota', '')
        sugestao_doc, sugestao_aliq = d.get('documento', ''), ""
        
    st.markdown("---")
    
    col_fields, col_info = st.columns([5, 1])
    with col_fields:
        col1, col2, col3, col4, col5 = st.columns([1, 1, 1, 1, 1])
        with col1:
            def update_filial(): d['filial'] = st.session_state[f"f_filial_{nota_uid}"]
            f_filial = st.text_input("Base Operacional (FBO)", d['filial'], key=f"f_filial_{nota_uid}", on_change=update_filial)
            def update_cnpj(): d['cnpj_posto'] = st.session_state[f"f_cnpj_{nota_uid}"]
            f_cnpj = st.text_input("CNPJ Fornecedor", d['cnpj_posto'], key=f"f_cnpj_{nota_uid}", on_change=update_cnpj)
        with col2:
            def update_doc(): d['documento'] = st.session_state[f"f_doc_{nota_uid}_{index_item}"]
            f_doc = st.text_input("Nº Fatura/Manifesto", sugestao_doc, key=f"f_doc_{nota_uid}_{index_item}", on_change=update_doc)
            def update_serie(): d['serie'] = st.session_state[f"f_serie_{nota_uid}"]
            sugestao_serie = st.text_input("Série", d['serie'], key=f"f_serie_{nota_uid}", on_change=update_serie)
        with col3:
            sugestao_data = st.text_input("Data do Abastecimento", d['data_emissao'], key=f"f_dtem_{nota_uid}")
            sugestao_hora = st.text_input("Hora (Local/ZULU)", d.get('hora_emissao', ''), key=f"f_hrem_{nota_uid}")
        with col4:
            sugestao_placa = st.text_input("Matrícula da Aeronave", d['placa'], key=f"f_placa_{nota_uid}")
            sugestao_km = st.text_input("Horas de Voo / Ciclos", d['km'], key=f"f_km_{nota_uid}")
        with col5:
            f_produto = st.text_input("Propelente (8=QAV, 6=AVGAS)", sugestao_produto, key=f"f_prod_{nota_uid}_{index_item}")
            f_qtde = st.text_input("Volume (Litros)", sugestao_qtde, key=f"f_qtde_{nota_uid}_{index_item}")
            f_valor = st.text_input("Valor Transação", sugestao_valor, key=f"f_valor_{nota_uid}_{index_item}")

    with col_info:
        with st.expander("ℹ️ Info Complementar", expanded=True):
            f_aliq = st.text_input("Aliq Imposto", sugestao_aliq, disabled=True, key=f"f_aliq_{nota_uid}_{index_item}")
            st.text_input("Comandante", d.get('comandante', ''), disabled=True)
            
    st.markdown("---")
    with st.expander("⚙️ Interface com VDI / Configurações de Ingestão"):
        default_f2_index = 1 if d['origem'] == 'EXCEL' else 0
        buscar_f2 = st.radio("Ativar busca de Fornecedor/F2 no ERP?", ("Sim (Recomendado)", "Não (Digitar Direto)"), index=default_f2_index)
        modo_popup = st.radio("Defesa contra alertas nativos do VDI:", ("Modo Interativo (Pausar e aguardar Humano)", "Modo Agressivo (Bypass automático)", "Pausar Timeout cego (Lento)"))
            
    st.warning("⚠️ **Atenção:** Ao iniciar, foque no campo 'Base Operacional' do ERP. O controle de hardware será assumido pela pipeline.")
    col_bt1, col_bt2, col_bt3, col_bt4 = st.columns([2, 1, 1, 1])
    
    with col_bt1:
        if st.button("🚀 INICIAR INGESTÃO NO ERP", type="primary", use_container_width=True):
            dados_finais = {
                'filial': f_filial, 'cnpj_posto': f_cnpj, 'documento': sugestao_doc, 'serie': sugestao_serie,
                'data_emissao': sugestao_data, 'hora_emissao': sugestao_hora, 'placa': sugestao_placa,
                'produto': f_produto, 'quantidade': f_qtde, 'valor': f_valor, 'km': sugestao_km
            }
            aviso = st.empty()
            for i in range(5, 0, -1):
                aviso.error(f"⏳ Posicione o cursor no ERP! Ingestão em {i} segundos...")
                time.sleep(1)
            aviso.success("🤖 Robô assumiu o controle de Hardware.")
            try:
                usar_f2 = buscar_f2.startswith("Sim")
                executar_rpa_aerosys(dados_finais, usar_f2, modo_popup)
                st.success("✅ Ingestão finalizada.")
            except Exception as e:
                st.error(f"❌ Intervenção Operacional: {e}")
                
    with col_bt2:
        if indice_tela > 0 and st.button("⬅️ Anterior", use_container_width=True):
            st.session_state.indice_nota_atual -= 1
            st.rerun()
            
    with col_bt3:
        st.markdown(f"<div style='text-align: center; padding-top: 10px; color: #aaa; font-weight: bold;'>Registro {indice_tela+1} / {total_notas}</div>", unsafe_allow_html=True)

    with col_bt4:
        if indice_tela < total_notas - 1 and st.button("Próximo ➡️", use_container_width=True):
            st.session_state.indice_nota_atual += 1
            st.rerun()
        elif total_notas > 1 and indice_tela == total_notas - 1:
             st.success("🏁 Fila concluída!")
