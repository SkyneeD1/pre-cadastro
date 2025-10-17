# -*- coding: utf-8 -*-
"""
Automação eLaw - Cadastro / Atualização com planilha
VERSÃO: V3.3.13 (FLUXO PRÉ-CADASTRO - COM QR CODE)
"""

import os
import re
import time
import math
import traceback
from datetime import datetime, timedelta

import pandas as pd

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException

from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# =====================
# CONFIGURAÇÕES
# =====================
EXCEL_PATH = "PLANILHA CADASTRO NOVA AÇÃO.xlsx"
CHROMEDRIVER_PATH = "C:/chromedriver/chromedriver.exe"
SITE_URL = "https://vtal.elaw.com.br/"
YELLOW_HEX = "FFF200"
WAIT_SHORT = 4  # REDUZIDO
WAIT_MEDIUM = 12  # REDUZIDO
WAIT_LONG = 25  # REDUZIDO

# =====================
# LER PLANILHA
# =====================
df = pd.read_excel(EXCEL_PATH)
if "STATUS" not in df.columns:
    df["STATUS"] = ""
df["STATUS"] = df["STATUS"].astype("object")

def set_status(idx, text):
    try:
        df.at[idx, "STATUS"] = str(text)
    except Exception:
        df.loc[idx, "STATUS"] = str(text)

# Colunas
COL_NUM_PROCESSO         = "Número do processo"
COL_RITO                 = "Localização do Processo"
COL_ESTADO               = "Estado"
COL_COMARCA              = "Comarca"
COL_FORO                 = "Foro/Tribunal"
COL_VARA                 = "Vara"
COL_CLASSIFICACAO        = "Classificação Interna"
COL_INSTANCIA            = "Instância"
COL_FASE                 = "Fase"
COL_JUIZ                 = "Juiz"
COL_CLIENTE_EMPRESA      = "Empresa e Forma de participação"
COL_CPF_PARTE_CONTR      = "CPF DA PARTE CONTRARIA"
COL_ADV_CONTR            = "Advogado da Parte Contrária"
COL_DATA_DISTR           = "Data de Distribuição"
COL_DATA_CITACAO         = "Data de Citação"
COL_TIPO_ACAO            = "Tipo de Ação"
COL_VALOR_CAUSA          = "Valor da Causa"
COL_ADV_RESP             = "Advogado Responsável"
COL_GESTOR_JURIDICO      = "Gestor Jurídico"
COL_TIPO_DOC             = "Tipo de Documento"
COL_RECLAMADA_1          = "1 Reclamada"
COL_RECLAMADA_2          = "2 Reclamada" 
COL_RECLAMADA_3          = "3 Reclamada"
COL_RECLAMADA_4          = "4 Reclamada"
COL_RECLAMADA_5          = "5 Reclamada"
COL_RECLAMADA_6          = "6 Reclamada"
COL_RECLAMADA_7          = "7 Reclamada"

# Dicionário para armazenar quais colunas falharam por linha
colunas_com_erro = {}

# =====================
# NORMALIZAÇÃO DE DATAS
# =====================
EXCEL_EPOCH = datetime(1899, 12, 30)

def as_ddmmyyyy(raw):
    if raw is None:
        return ""
    try:
        if pd.isna(raw):
            return ""
    except Exception:
        pass

    if isinstance(raw, (datetime, pd.Timestamp)):
        return raw.strftime("%d/%m/%Y")

    if isinstance(raw, (int, float)) and not isinstance(raw, bool):
        if math.isfinite(raw):
            try:
                dt = EXCEL_EPOCH + timedelta(days=float(raw))
                if 1900 <= dt.year <= 2100:
                    return dt.strftime("%d/%m/%Y")
            except Exception:
                pass

    s = str(raw).strip()
    if not s:
        return ""

    for dayfirst in (True, False):
        try:
            dt = pd.to_datetime(s, dayfirst=dayfirst, errors="raise")
            if 1900 <= dt.year <= 2100:
                return dt.strftime("%d/%m/%Y")
        except Exception:
            pass

    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%m/%d/%Y"):
        try:
            dt = datetime.strptime(s, fmt)
            if 1900 <= dt.year <= 2100:
                return dt.strftime("%d/%m/%Y")
        except Exception:
            continue

    digits = re.sub(r"\D", "", s)
    if len(digits) == 8:
        try:
            dt = datetime.strptime(digits, "%d%m%Y")
            return dt.strftime("%d/%m/%Y")
        except Exception:
            pass
        try:
            dt = datetime.strptime(digits, "%Y%m%d")
            return dt.strftime("%d/%m/%Y")
        except Exception:
            pass

    return ""

# =====================
# SELENIUM SETUP - VERSÃO OTIMIZADA COM IMAGENS PARA QR CODE
# =====================
options = webdriver.ChromeOptions()

# OTIMIZAÇÕES DE DESEMPENHO (mantendo imagens para QR Code)
options.add_argument("--start-maximized")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--disable-gpu")  # Melhora performance
options.add_argument("--disable-extensions")  # Remove extensões
options.add_argument("--disable-plugins")  # Desabilita plugins
# REMOVIDO: --disable-images (para QR Code funcionar)
# REMOVIDO: --blink-settings=imagesEnabled=false (para QR Code funcionar)
options.add_argument("--disable-background-timer-throttling")
options.add_argument("--disable-backgrounding-occluded-windows")
options.add_argument("--disable-renderer-backgrounding")

# OTIMIZAÇÕES DE REDE
options.add_argument("--aggressive-cache-discard")
options.add_argument("--disable-features=VizDisplayCompositor")
options.add_argument("--disable-ipc-flooding-protection")

# SILENCIAR LOGS
options.add_argument("--log-level=3")  # Só erros críticos
options.add_argument("--silent")
options.add_argument("--disable-logging")

# PREFERÊNCIAS DE PERFORMANCE (PERMITINDO IMAGENS PARA QR CODE)
prefs = {
    'profile.default_content_setting_values': {
        'images': 1,  # ✅ PERMITE IMAGENS (para QR Code)
        'javascript': 1,  # Mantém JavaScript (necessário para o eLaw)
        'plugins': 2,  # Bloqueia plugins
        'popups': 2,  # Bloqueia popups
        'geolocation': 2,  # Bloqueia geolocalização
        'notifications': 2,  # Bloqueia notificações
        'auto_select_certificate': 2,
        'fullscreen': 2,
    },
    'disk-cache-size': 4096,  # Cache de disco
}

options.add_experimental_option("prefs", prefs)

# SERVICE CONFIG
service = Service(CHROMEDRIVER_PATH)

# CONFIGURAÇÕES ADICIONAIS DO DRIVER
service.creationflags = 0x08000000  # Flags para melhor performance no Windows

driver = webdriver.Chrome(service=service, options=options)

# CONFIGURAÇÕES DE TIMEOUT OTIMIZADAS
driver.set_page_load_timeout(30)  # Timeout reduzido para carregamento
driver.implicitly_wait(5)  # Wait implícito reduzido

wait = WebDriverWait(driver, WAIT_LONG)

# =====================
# FUNÇÃO PARA AGUARDAR QR CODE E FAZER LOGIN MANUAL
# =====================
def aguardar_login_qrcode():
    """Aguarda o usuário fazer login via QR Code manualmente"""
    print("📱 AGUARDANDO LOGIN VIA QR CODE...")
    print("👀 Por favor, escaneie o QR Code que aparecerá na tela")
    
    try:
        # Aguardar o elemento do QR Code aparecer
        qr_code_element = WebDriverWait(driver, 120).until(
            EC.presence_of_element_located((By.ID, "qrcode-access"))
        )
        print("✅ QR Code detectado na tela!")
        print("📱 Use o app do eLaw no celular para escanear o código")
        
        # Aguardar o usuário fazer login (timeout de 3 minutos)
        print("⏳ Aguardando você escanear o QR Code e fazer login...")
        WebDriverWait(driver, 180).until(
            EC.url_contains("/homePage.elaw")
        )
        print("✅ Login realizado com sucesso!")
        return True
        
    except Exception as e:
        print(f"❌ Erro no login via QR Code: {e}")
        print("💡 Dica: Certifique-se de escanear o QR Code dentro de 3 minutos")
        return False

# =====================
# HELPERS FLEXÍVEIS
# =====================
def safe_text(val):
    try:
        if pd.isna(val):
            return ""
    except Exception:
        pass
    return str(val).strip()

def to_amount_str(val):
    """
    Converte valor para formato numérico simples
    """
    if val is None or (isinstance(val, float) and math.isnan(val)) or (isinstance(val, str) and not val.strip()):
        return ""
    
    try:
        # Se já for número, formata sem casas decimais se for inteiro
        if isinstance(val, (int, float)):
            if val == int(val):
                return str(int(val))
            else:
                return str(val)
        
        valor_str = str(val).strip()
        
        # Remove R$, espaços e outros caracteres
        valor_str = re.sub(r'[R$\s]', '', valor_str)
        
        # Remove pontos de milhar e converte vírgula para ponto
        if '.' in valor_str and ',' in valor_str:
            # Formato: 1.234,56 → 1234.56
            valor_str = valor_str.replace('.', '').replace(',', '.')
        elif ',' in valor_str:
            # Formato: 1234,56 → 1234.56
            valor_str = valor_str.replace(',', '.')
        
        # Converte para float e formata
        valor_float = float(valor_str)
        
        # Formata sem casas decimais se for inteiro
        if valor_float == int(valor_float):
            return str(int(valor_float))
        else:
            return str(valor_float)
            
    except Exception as e:
        print(f"⚠️ Erro ao converter valor '{val}': {e}")
        # Fallback: remove tudo que não é número ou ponto
        try:
            cleaned = re.sub(r'[^\d.]', '', str(val))
            return str(float(cleaned))
        except:
            return str(val)

def wait_element_by_id_suffix(suffix, tag="*", timeout=WAIT_LONG, condition=None):
    """Localiza elemento pelo final do ID (funciona com 4c, 4g, 45, etc)"""
    selector = f"{tag}[id$='{suffix}']"
    locator = (By.CSS_SELECTOR, selector)
    expected = condition(locator) if condition else EC.presence_of_element_located(locator)
    return WebDriverWait(driver, timeout).until(expected)

def attempt_twice(action_desc, func, *args, **kwargs):
    for tent in range(1, 3):
        try:
            r = func(*args, **kwargs)
            print(f"✅ {action_desc} (tentativa {tent})")
            return True if r is None else r
        except Exception as e:
            print(f"⚠️ Falha em '{action_desc}' (tentativa {tent}): {e}")
            if tent == 1:
                time.sleep(1.0)  # REDUZIDO
    return False

def _xpath_literal(texto):
    if "'" not in texto:
        return f"'{texto}'"
    if '"' not in texto:
        return f'"{texto}"'
    partes = texto.split("'")
    pedacos = []
    for idx, parte in enumerate(partes):
        if parte:
            pedacos.append(f"'{parte}'")
        if idx != len(partes) - 1:
            pedacos.append("\"'\"")
    return "concat(" + ",".join(pedacos) + ")"

def clicar_elemento_por_sufixo(suffix, tag="button"):
    """Clica em elemento pelo sufixo do ID"""
    elem = wait_element_by_id_suffix(suffix, tag, condition=EC.element_to_be_clickable)
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", elem)
    try:
        elem.click()
    except Exception:
        driver.execute_script("arguments[0].click();", elem)
    time.sleep(0.3)  # REDUZIDO

def preencher_input_por_sufixo(suffix, valor, tag="input"):
    """Preenche input pelo sufixo do ID com o mínimo de esperas possíveis"""
    if valor == "" and valor != 0:
        return

    elem = wait_element_by_id_suffix(suffix, tag, condition=EC.element_to_be_clickable)
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", elem)

    value_str = str(valor)

    try:
        # Usa JavaScript para limpar e preencher rapidamente o campo
        driver.execute_script(
            "arguments[0].focus();"
            "arguments[0].value = '';"
            "arguments[0].dispatchEvent(new Event('input', {bubbles: true}));",
            elem,
        )
        driver.execute_script(
            "arguments[0].value = arguments[1];"
            "arguments[0].dispatchEvent(new Event('input', {bubbles: true}));"
            "arguments[0].dispatchEvent(new Event('change', {bubbles: true}));",
            elem,
            value_str,
        )

        WebDriverWait(driver, 4).until(
            lambda d: driver.execute_script("return arguments[0].value;", elem) == value_str
        )
    except Exception:
        # Fallback: usa send_keys caso o JavaScript falhe
        try:
            elem.click()
            elem.send_keys(Keys.CONTROL, "a")
            elem.send_keys(Keys.BACKSPACE)
            elem.send_keys(value_str)
            WebDriverWait(driver, 4).until(
                lambda d: driver.execute_script("return arguments[0].value;", elem) == value_str
            )
        except StaleElementReferenceException:
            # Reobtém o elemento e tenta novamente apenas uma vez
            elem = wait_element_by_id_suffix(suffix, tag, condition=EC.element_to_be_clickable)
            elem.clear()
            elem.send_keys(value_str)
        finally:
            try:
                WebDriverWait(driver, 2).until(
                    lambda d: driver.execute_script("return arguments[0].value;", elem) == value_str
                )
            except Exception:
                pass

def digitar_data_humano_por_sufixo(suffix, data_valor):
    """Digita data em modo humano pelo sufixo do ID"""
    try:
        if not data_valor:
            return True
        campo = wait_element_by_id_suffix(suffix, "input", condition=EC.element_to_be_clickable)
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", campo)
        campo.click()
        time.sleep(0.2)  # REDUZIDO
        campo.send_keys(Keys.CONTROL, "a")
        campo.send_keys(Keys.BACKSPACE)
        time.sleep(0.1)  # REDUZIDO
        for ch in data_valor:
            campo.send_keys(ch)
            time.sleep(0.04)  # REDUZIDO
        time.sleep(0.08)  # REDUZIDO
        campo.send_keys(Keys.ENTER)
        print(f"✅ Data '{data_valor}' digitada (modo humano)")
        time.sleep(0.25)  # REDUZIDO
        return True
    except Exception as e:
        print(f"❌ Erro ao digitar data: {e}")
        return False

def anexar_arquivo_por_input(file_path):
    upload_input = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@type='file']")))
    upload_input.send_keys(os.path.abspath(file_path))
    time.sleep(0.6)  # REDUZIDO

def marcar_erro(idx, etapa, err):
    msg = f"ERRO {etapa}: {err}"
    print(f"❌ {msg}")
    set_status(idx, f"⚠️ {msg}")

def esperar_texto_em_tabela_outras_partes(texto, timeout=WAIT_MEDIUM):
    if not texto:
        return False
    literal = _xpath_literal(texto.strip())
    xpath = f"//table[contains(@id,'outrasParte')]//span[contains(normalize-space(.), {literal})]"
    try:
        WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.XPATH, xpath)))
        return True
    except Exception as e:
        print(f"⚠️ Não encontrei '{texto}' na lista de Outras Partes: {e}")
        return False

# =====================
# FUNÇÕES AUTCOMPLETE SIMPLIFICADAS
# =====================
def preencher_autocomplete_simples(suffix, valor, tag="input"):
    """Preenche autocomplete aguardando a lista abrir em vez de usar sleeps longos"""
    if not valor:
        return True

    panel_selector = "div.ui-autocomplete-panel[style*='display: block']"

    try:
        elem = wait_element_by_id_suffix(suffix, tag, condition=EC.element_to_be_clickable)
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", elem)
        elem.click()

        elem.send_keys(Keys.CONTROL, "a")
        elem.send_keys(Keys.BACKSPACE)

        elem.send_keys(valor)
        print(f"✍️ Digitando '{valor}' no autocomplete...")

        panel = WebDriverWait(driver, 5).until(
            EC.visibility_of_element_located((By.CSS_SELECTOR, panel_selector))
        )

        literal = _xpath_literal(valor)
        option_locator = (
            By.XPATH,
            f"//div[contains(@class,'ui-autocomplete-panel') and contains(@style,'display: block')]//li[normalize-space(.)={literal} and not(contains(@class,'ui-state-disabled'))]",
        )

        try:
            option = WebDriverWait(driver, 3).until(EC.element_to_be_clickable(option_locator))
        except TimeoutException:
            option = WebDriverWait(driver, 3).until(
                EC.element_to_be_clickable(
                    (
                        By.XPATH,
                        "//div[contains(@class,'ui-autocomplete-panel') and contains(@style,'display: block')]//li[not(contains(@class,'ui-state-disabled'))]",
                    )
                )
            )

        driver.execute_script("arguments[0].scrollIntoView({block:'nearest'});", option)
        driver.execute_script("arguments[0].click();", option)

        WebDriverWait(driver, 4).until(
            lambda d: driver.execute_script("return arguments[0].value;", elem).strip() != ""
        )

        print(f"✅ Autocomplete preenchido: {valor}")
        return True

    except Exception as e:
        print(f"❌ Erro no autocomplete '{valor}': {e}")
        return False

# =====================
# FUNÇÕES PRIMEFACES FLEXÍVEIS - USANDO A MESMA LÓGICA DOS OUTROS DROPDOWNS
# =====================
_SIGLA_ESTADO_RE = re.compile(r"^[A-Z]{2}$")

def _ajusta_valor_para_estado(label_id, valor):
    if not valor:
        return valor
    id_lower = label_id.lower()
    pode_ser_estado = ("comboestadovara" in id_lower) or ("estado" in id_lower)
    if pode_ser_estado and _SIGLA_ESTADO_RE.match(valor.strip().upper()):
        return valor.strip().upper() + " -"
    return valor

def selecionar_primefaces_por_sufixo(suffix, valor, timeout=WAIT_LONG):
    """Seleciona em dropdown PrimeFaces com foco em velocidade."""
    valor = _ajusta_valor_para_estado(suffix, (valor or "").strip())
    alvo_normalizado = valor.split(" -")[0].strip().lower() if valor else ""

    label = wait_element_by_id_suffix(suffix, "span", condition=EC.element_to_be_clickable)
    texto_atual = (label.text or "").strip().lower()
    if alvo_normalizado and alvo_normalizado in texto_atual:
        return True

    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", label)
    driver.execute_script("arguments[0].click();", label)

    panel = WebDriverWait(driver, timeout).until(
        EC.visibility_of_element_located((By.CSS_SELECTOR, "div.ui-selectonemenu-panel[style*='display: block']"))
    )

    filtro = None
    try:
        filtro = panel.find_element(By.XPATH, ".//input[contains(@id,'_filter')]")
    except Exception:
        filtro = None

    if filtro is not None:
        filtro.click()
        filtro.send_keys(Keys.CONTROL, "a")
        filtro.send_keys(Keys.BACKSPACE)
        if valor:
            filtro.send_keys(valor)

        try:
            WebDriverWait(driver, 2).until(
                EC.presence_of_element_located(
                    (
                        By.CSS_SELECTOR,
                        "div.ui-selectonemenu-panel[style*='display: block'] li:not(.ui-state-disabled)",
                    )
                )
            )
        except TimeoutException:
            pass

        filtro.send_keys(Keys.ENTER)
    else:
        option = None
        if valor:
            literal = _xpath_literal(valor)
            for locator in (
                (
                    By.XPATH,
                    f"//div[contains(@class,'ui-selectonemenu-panel') and contains(@style,'display: block')]//li[@data-label and normalize-space(@data-label)={literal} and not(contains(@class,'ui-state-disabled'))]",
                ),
                (
                    By.XPATH,
                    f"//div[contains(@class,'ui-selectonemenu-panel') and contains(@style,'display: block')]//li[normalize-space(.)={literal} and not(contains(@class,'ui-state-disabled'))]",
                ),
            ):
                try:
                    option = WebDriverWait(driver, 2).until(EC.element_to_be_clickable(locator))
                    break
                except TimeoutException:
                    continue

        if option is None:
            option = WebDriverWait(driver, 2).until(
                EC.element_to_be_clickable(
                    (
                        By.CSS_SELECTOR,
                        "div.ui-selectonemenu-panel[style*='display: block'] li:not(.ui-state-disabled)",
                    )
                )
            )

        driver.execute_script("arguments[0].scrollIntoView({block:'nearest'});", option)
        driver.execute_script("arguments[0].click();", option)

    try:
        WebDriverWait(driver, 3).until(EC.invisibility_of_element(panel))
    except TimeoutException:
        pass

    if valor:
        try:
            atualizado = wait_element_by_id_suffix(suffix, "span")
            WebDriverWait(driver, 2).until(
                lambda d: alvo_normalizado in ((atualizado.text or "").strip().lower())
            )
        except TimeoutException:
            pass

    return True

# =====================
# FUNÇÃO ESPECÍFICA PARA PRÉ-CADASTRO - BASEADA NO TESTE QUE FUNCIONOU
# =====================
def selecionar_area_pre_cadastro():
    """Função específica para selecionar área e subárea no pré-cadastro - BASEADA NO TESTE JS"""
    print("📋 Selecionando área do processo...")
    
    def _selecionar_opcao(dropdown_id, valor_desejado):
        """Seleciona opção em dropdown - mesma lógica do teste JS"""
        print(f"🎯 Tentando selecionar: {valor_desejado}")
        
        try:
            # 1. Clicar no dropdown para abrir
            dropdown = driver.find_element(By.ID, dropdown_id)
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", dropdown)
            dropdown.click()
            print(f"✅ Dropdown {dropdown_id} clicado")
            time.sleep(1.0)  # REDUZIDO
            
            # 2. Buscar todas as opções abertas
            panels = driver.find_elements(By.CSS_SELECTOR, 'div.ui-selectonemenu-panel[style*="display: block"]')
            print(f"📋 Panels encontrados: {len(panels)}")
            
            if len(panels) == 0:
                print("❌ Nenhum panel aberto encontrado")
                return False
            
            opcao_encontrada = None
            
            # 3. Buscar em todos os panels abertos
            for i, panel in enumerate(panels):
                print(f"🔍 Buscando no panel {i + 1}...")
                
                # Tentar por data-label
                try:
                    por_data_label = panel.find_element(By.XPATH, f".//li[@data-label='{valor_desejado}']")
                    opcao_encontrada = por_data_label
                    print(f"✅ Encontrado por data-label: {valor_desejado}")
                    break
                except:
                    pass
                
                # Tentar por texto
                try:
                    opcoes = panel.find_elements(By.CSS_SELECTOR, 'li.ui-selectonemenu-item')
                    for opcao in opcoes:
                        if opcao.text.strip() == valor_desejado:
                            opcao_encontrada = opcao
                            print(f"✅ Encontrado por texto: {valor_desejado}")
                            break
                    if opcao_encontrada:
                        break
                except:
                    pass
            
            if opcao_encontrada:
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", opcao_encontrada)
                driver.execute_script("arguments[0].click();", opcao_encontrada)
                print(f"🎉 {valor_desejado} selecionado com sucesso!")
                time.sleep(0.8)  # REDUZIDO
                return True
            else:
                print(f"❌ Opção '{valor_desejado}' não encontrada em nenhum panel")
                return False
                
        except Exception as e:
            print(f"❌ Erro ao selecionar {valor_desejado}: {e}")
            return False
    
    # EXECUTAR SELEÇÕES - MESMA SEQUÊNCIA DO TESTE JS
    
    print("=" * 50)
    print("1️⃣ TESTANDO SELEÇÃO DE TRABALHISTA")
    
    # Testar Área - Trabalhista
    if not attempt_twice("Selecionar Trabalhista", _selecionar_opcao, "comboArea_label", "Trabalhista"):
        raise Exception("Não foi possível selecionar Trabalhista")
    
    print("⏳ Aguardando 1.5 segundos para carregar subáreas...")  # REDUZIDO
    time.sleep(1.5)
    
    print("=" * 50)
    print("2️⃣ TESTANDO SELEÇÃO DE DIREITO INDIVIDUAL")
    
    # Testar Subárea - Direito Individual
    if not attempt_twice("Selecionar Direito Individual", _selecionar_opcao, "comboAreaSub_label", "Direito Individual"):
        raise Exception("Não foi possível selecionar Direito Individual")
    
    print("⏳ Aguardando 0.8 segundos...")  # REDUZIDO
    time.sleep(0.8)
    
    print("=" * 50)
    print("3️⃣ TESTANDO BOTÃO CONTINUAR")
    
    # Testar botão Continuar
    def _clicar_continuar():
        btn_continuar = driver.find_element(By.ID, "btnContinuar")
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn_continuar)
        btn_continuar.click()
        return True
    
    if not attempt_twice("Clicar em Continuar", _clicar_continuar):
        raise Exception("Botão Continuar indisponível")
    
    print("✅ Botão Continuar clicado!")
    time.sleep(2.0)  # REDUZIDO
    print("🎉 TESTE CONCLUÍDO COM SUCESSO!")
    print("✅ Área e subárea selecionadas com sucesso!")

# =====================
# FUNÇÃO PARA ADICIONAR RECLAMADAS (SIMPLIFICADA)
# =====================
def adicionar_reclamadas_automaticamente(reclamadas_nomes, idx):
    """Adiciona automaticamente as reclamadas das colunas Q a W"""
    if not any(reclamadas_nomes):
        print("ℹ️ Nenhuma reclamada para adicionar")
        return True

    reclamadas_adicionadas = 0
    colunas_falhas = []  # Armazena as colunas que falharam
    
    # Mapeamento de índice para nome da coluna
    colunas_map = {
        0: COL_RECLAMADA_1,
        1: COL_RECLAMADA_2,
        2: COL_RECLAMADA_3,
        3: COL_RECLAMADA_4,
        4: COL_RECLAMADA_5,
        5: COL_RECLAMADA_6,
        6: COL_RECLAMADA_7
    }
    
    for i, parte_nome in enumerate(reclamadas_nomes):
        # Verificar se é vazio de forma mais robusta
        if (not parte_nome or 
            str(parte_nome).strip() == "" or 
            str(parte_nome).lower() == "nan" or
            str(parte_nome).strip() == "None"):
            continue

        print(f"➕ Adicionando reclamada {i+1}: {parte_nome}")

        try:
            # 1. PREENCHER AUTCOMPLETE - MÉTODO SIMPLIFICADO
            if not preencher_autocomplete_simples(":autocompleteOutraParte_input", parte_nome):
                print(f"❌ Falha ao preencher autocomplete para '{parte_nome}'")
                colunas_falhas.append(colunas_map[i])
                continue

            # 2. SELECIONAR "RÉU" - USANDO A MESMA LÓGICA DOS OUTROS DROPDOWNS
            try:
                # Usar a mesma função que já funciona nos outros lugares
                if not attempt_twice(f"Selecionar Réu para reclamada {i+1}", 
                                   selecionar_primefaces_por_sufixo, ":processoParteSelect_label", "Réu"):
                    print(f"⚠️ Não foi possível selecionar 'Réu' para reclamada {i+1}")
                    colunas_falhas.append(colunas_map[i])
                    continue
                    
                print(f"✅ Papel 'Réu' selecionado para reclamada {i+1}")
                    
            except Exception as e:
                print(f"❌ Erro ao selecionar Réu: {e}")
                colunas_falhas.append(colunas_map[i])
                continue

            # 3. CLICAR EM "ADICIONAR"
            try:
                botao_adicionar = wait_element_by_id_suffix(
                    ":outrasParteAddButtom", "button", condition=EC.element_to_be_clickable
                )
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", botao_adicionar)
                botao_adicionar.click()
                time.sleep(2.0)  # REDUZIDO
                
                # Verificar se foi adicionada (com timeout menor)
                try:
                    if esperar_texto_em_tabela_outras_partes(parte_nome, timeout=4):  # REDUZIDO
                        print(f"✅ Reclamada '{parte_nome}' adicionada com sucesso!")
                        reclamadas_adicionadas += 1
                    else:
                        print(f"⚠️ Reclamada '{parte_nome}' pode não ter sido adicionada")
                        colunas_falhas.append(colunas_map[i])
                except Exception:
                    print(f"⚠️ Reclamada '{parte_nome}' pode não ter sido adicionada")
                    colunas_falhas.append(colunas_map[i])
                    
            except Exception as e:
                print(f"❌ Erro ao clicar em Adicionar: {e}")
                colunas_falhas.append(colunas_map[i])
                continue

        except Exception as e:
            print(f"❌ Erro ao adicionar reclamada '{parte_nome}': {e}")
            colunas_falhas.append(colunas_map[i])
            continue

    print(f"📊 Total de reclamadas adicionadas: {reclamadas_adicionadas}")
    
    # Se houve falhas, armazena quais colunas falharam
    if colunas_falhas:
        print(f"❌ Colunas com problema: {colunas_falhas}")
        colunas_com_erro[idx] = colunas_falhas
        return False
    
    return reclamadas_adicionadas > 0

# =====================
# MODAIS COM IFRAME (FLEXÍVEIS)
# =====================
def _get_visible_dialogs():
    dialogs = driver.find_elements(By.CSS_SELECTOR, "div.ui-dialog.ui-overlay-visible, div.ui-dialog[style*='display: block']")
    return [d for d in dialogs if d.is_displayed()]

def _find_dialog_iframe(dialog):
    try:
        return dialog.find_element(By.CSS_SELECTOR, "iframe")
    except Exception:
        return None

def _switch_into_dialog_iframe_by_hint(id_hint_contains, timeout=WAIT_LONG):
    t0 = time.time()
    while time.time() - t0 < timeout:
        dialogs = _get_visible_dialogs()
        for d in dialogs:
            try:
                dlg_id = d.get_attribute("id") or ""
                if id_hint_contains and id_hint_contains in dlg_id:
                    ifr = _find_dialog_iframe(d)
                    if ifr:
                        print(f"🔎 Dialog encontrado, entrando no iframe...")
                        driver.switch_to.frame(ifr)
                        return d
            except Exception:
                pass
        for d in dialogs:
            ifr = _find_dialog_iframe(d)
            if ifr:
                driver.switch_to.frame(ifr)
                return d
        time.sleep(0.15)  # REDUZIDO
    raise Exception("Timeout ao localizar iframe")

def _leave_iframe():
    try:
        driver.switch_to.default_content()
    except Exception:
        pass

def criar_juiz_modal_js(juiz_nome):
    print("➡️ Abrindo modal Juiz...")
    clicar_elemento_por_sufixo(":juizBtnNovo")
    
    print("⏳ Aguardando dialog + iframe do Juiz...")
    dialog = _switch_into_dialog_iframe_by_hint("juizBtnNovo_dlg", timeout=WAIT_LONG)

    try:
        print("✍️ Preenchendo nome do Juiz...")
        input_elem = WebDriverWait(driver, WAIT_LONG).until(EC.presence_of_element_located((By.ID, "j_id_w")))
        driver.execute_script("arguments[0].focus();", input_elem)
        driver.execute_script("arguments[0].value = arguments[1];", input_elem, juiz_nome)
        driver.execute_script("arguments[0].dispatchEvent(new Event('input',{bubbles:true}));", input_elem)
        time.sleep(0.25)  # REDUZIDO

        print("💾 Clicando salvar do Juiz...")
        salvar_btn = driver.find_element(By.ID, "btnSalvarjuiz")
        driver.execute_script("arguments[0].click();", salvar_btn)
        time.sleep(0.6)  # REDUZIDO
    except Exception as e:
        raise Exception(f"Erro ao preencher/salvar Juiz: {e}")
    finally:
        _leave_iframe()

def incluir_parte_contraria_modal_js(cpf_cnpj):
    print("➡️ Abrindo modal Parte Contrária...")
    clicar_elemento_por_sufixo(":parteContrariaMainGridBtnNovo")

    print("⏳ Aguardando dialog + iframe da Parte Contrária...")
    dialog = _switch_into_dialog_iframe_by_hint("parteContrariaMainGridBtnNovo_dlg", timeout=WAIT_LONG)

    try:
        print("✍️ Preenchendo CPF/CNPJ...")
        input_elem = WebDriverWait(driver, WAIT_LONG).until(EC.presence_of_element_located((By.ID, "j_id_1e")))
        driver.execute_script("arguments[0].focus();", input_elem)
        driver.execute_script("arguments[0].value = arguments[1];", input_elem, cpf_cnpj)
        driver.execute_script("arguments[0].dispatchEvent(new Event('input',{bubbles:true}));", input_elem)
        time.sleep(0.25)  # REDUZIDO

        print("➡️ Clicando 'Continuar'...")
        cont_btn = driver.find_element(By.ID, "j_id_1i")
        driver.execute_script("arguments[0].click();", cont_btn)
        time.sleep(0.6)  # REDUZIDO

        print("💾 Clicando 'Salvar'...")
        save_btn = WebDriverWait(driver, WAIT_LONG).until(EC.visibility_of_element_located((By.ID, "parteContrariaButtom")))
        driver.execute_script("arguments[0].click();", save_btn)
        time.sleep(0.6)  # REDUZIDO
    except Exception as e:
        raise Exception(f"Erro ao incluir Parte Contrária: {e}")
    finally:
        _leave_iframe()

def colorir_colunas_amarelo_no_excel(excel_path, colunas_erro_dict, header_rows=1):
    """Pinta apenas as colunas específicas que falharam"""
    try:
        wb = load_workbook(excel_path)
        ws = wb.active
        
        # Mapeamento de nomes de colunas para índices
        colunas_indices = {
            COL_RECLAMADA_1: 'Q',
            COL_RECLAMADA_2: 'R', 
            COL_RECLAMADA_3: 'S',
            COL_RECLAMADA_4: 'T',
            COL_RECLAMADA_5: 'U',
            COL_RECLAMADA_6: 'V',
            COL_RECLAMADA_7: 'W'
        }
        
        fill = PatternFill(start_color=YELLOW_HEX, end_color=YELLOW_HEX, fill_type="solid")
        
        for idx, colunas_falhas in colunas_erro_dict.items():
            excel_row = idx + 1 + header_rows
            
            for coluna_nome in colunas_falhas:
                if coluna_nome in colunas_indices:
                    col_letter = colunas_indices[coluna_nome]
                    celula = f"{col_letter}{excel_row}"
                    ws[celula].fill = fill
                    print(f"🎨 Coluna {col_letter} da linha {excel_row} colorida de amarelo")
        
        wb.save(excel_path)
        print(f"🎨 Colunas com erro coloridas de amarelo")
        
    except Exception as e:
        print(f"⚠️ Falha ao colorir colunas: {e}")

# =====================
# FLUXO PRINCIPAL MODIFICADO - PRÉ-CADASTRO
# =====================
try:
    driver.get(SITE_URL)
    print("👀 Aguardando login... (até 180s)")
    try:
        WebDriverWait(driver, 180).until(EC.url_contains("/homePage.elaw"))
        print("✅ Login detectado, iniciando automação...")
    except:
        print("⚠️ Login não detectado. Faça login e pressione ENTER.")
        input("👉 Pressione ENTER após logar...")

    # NOVO FLUXO: ACESSAR PÁGINA DE PRÉ-CADASTRO
    print("➡️ Acessando página de pré-cadastro...")
    driver.get("https://vtal.elaw.com.br/processoPreCadastroNew.elaw")
    time.sleep(2)  # REDUZIDO

    for idx, row in df.iterrows():
        processo = safe_text(row.get(COL_NUM_PROCESSO, ""))
        if not processo:
            continue

        print("\n" + "="*86)
        print(f"🔎 Linha {idx+1} | Processo: {processo}")
        set_status(idx, "EM ANDAMENTO...")

        # Extrair campos
        rito            = safe_text(row.get(COL_RITO, ""))
        estado_vara     = safe_text(row.get(COL_ESTADO, ""))
        comarca_vara    = safe_text(row.get(COL_COMARCA, ""))
        foro_tribunal   = safe_text(row.get(COL_FORO, ""))
        vara_especifica = safe_text(row.get(COL_VARA, ""))
        classificacao   = safe_text(row.get(COL_CLASSIFICACAO, ""))
        instancia       = safe_text(row.get(COL_INSTANCIA, ""))
        fase_processo   = safe_text(row.get(COL_FASE, ""))
        juiz_nome       = safe_text(row.get(COL_JUIZ, ""))
        cliente_empresa = safe_text(row.get(COL_CLIENTE_EMPRESA, ""))
        cpf_cnpj_contr  = safe_text(row.get(COL_CPF_PARTE_CONTR, ""))
        advogado_contr  = safe_text(row.get(COL_ADV_CONTR, ""))
        tipo_processo   = safe_text(row.get(COL_TIPO_ACAO, ""))
        valor_causa     = to_amount_str(row.get(COL_VALOR_CAUSA, ""))
        adv_resp        = safe_text(row.get(COL_ADV_RESP, ""))
        gestor_juridico = safe_text(row.get(COL_GESTOR_JURIDICO, ""))

        # DATAS
        data_distrib    = as_ddmmyyyy(row.get(COL_DATA_DISTR, ""))
        data_receb      = as_ddmmyyyy(row.get(COL_DATA_CITACAO, ""))

        tipo_doc_val    = safe_text(row.get(COL_TIPO_DOC, "")) or "Petição Inicial"

        # RECLAMADAS
        reclamadas_nomes = [
            safe_text(row.get(COL_RECLAMADA_1, "")),
            safe_text(row.get(COL_RECLAMADA_2, "")),
            safe_text(row.get(COL_RECLAMADA_3, "")),
            safe_text(row.get(COL_RECLAMADA_4, "")),
            safe_text(row.get(COL_RECLAMADA_5, "")),
            safe_text(row.get(COL_RECLAMADA_6, "")),
            safe_text(row.get(COL_RECLAMADA_7, ""))
        ]

        pdf_filename = f"ATOrd_{processo}.pdf"
        pdf_path = os.path.join(os.getcwd(), pdf_filename)

        try:
            # NOVO FLUXO: SELECIONAR ÁREA E SUBÁREA - USANDO MESMA LÓGICA QUE JÁ FUNCIONA
            if not attempt_twice("Selecionar área e subárea", selecionar_area_pre_cadastro):
                raise Exception("Falha ao selecionar área e subárea")

            # INSERIR NÚMERO DO PROCESSO DIRETAMENTE (sem pesquisar)
            print("🔢 Inserindo número do processo...")
            def _inserir_numero_processo():
                # Localizar o campo pelo ID parcial (flexível)
                numero_input = wait_element_by_id_suffix(":inputTxtNumeroMask", "input")
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", numero_input)
                numero_input.click()
                time.sleep(0.15)  # REDUZIDO
                numero_input.send_keys(Keys.CONTROL, "a")
                numero_input.send_keys(Keys.BACKSPACE)
                time.sleep(0.15)  # REDUZIDO
                numero_input.send_keys(processo)
                time.sleep(0.3)  # REDUZIDO
                print(f"✅ Número do processo '{processo}' inserido")
            
            if not attempt_twice("Inserir número do processo", _inserir_numero_processo):
                raise Exception("Não foi possível inserir número do processo")

            # O RESTO DO CÓDIGO PERMANECE IGUAL:
            # DROPDOWNS FLEXÍVEIS
            if rito:
                attempt_twice("Selecionar Rito", selecionar_primefaces_por_sufixo, ":comboRito_label", rito)
            if estado_vara:
                attempt_twice("Selecionar Estado", selecionar_primefaces_por_sufixo, ":comboEstadoVara_label", estado_vara)
            if comarca_vara:
                attempt_twice("Selecionar Comarca", selecionar_primefaces_por_sufixo, ":comboComarcaVara_label", comarca_vara)
            if foro_tribunal:
                attempt_twice("Selecionar Foro/Tribunal", selecionar_primefaces_por_sufixo, ":comboForoTribunal_label", foro_tribunal)
            if vara_especifica:
                attempt_twice("Selecionar Vara", selecionar_primefaces_por_sufixo, ":comboVara_label", vara_especifica)
            if classificacao:
                attempt_twice("Selecionar Classificação", selecionar_primefaces_por_sufixo, ":processoClassificacaoCombo_label", classificacao)
            if instancia:
                attempt_twice("Selecionar Instância", selecionar_primefaces_por_sufixo, ":j_id_4c_5_2_2_3_9_19_1_label", instancia)
            if fase_processo:
                attempt_twice("Selecionar Fase", selecionar_primefaces_por_sufixo, ":processoFaseCombo_label", fase_processo)
            if cliente_empresa:
                attempt_twice("Selecionar Empresa", selecionar_primefaces_por_sufixo, ":comboClientProcessoParte_label", cliente_empresa)

            # PAPEL = RÉU
            attempt_twice("Selecionar Papel = Réu", selecionar_primefaces_por_sufixo, ":j_id_4c_5_2_2_9_9_2_6_label", "Réu")

            # TIPO DOCUMENTO
            if tipo_doc_val:
                attempt_twice("Selecionar Tipo de Documento", selecionar_primefaces_por_sufixo, ":eFileTipoCombo_label", tipo_doc_val)

            # PARTE DOCUMENTO = AUTOR
            attempt_twice("Selecionar Parte = Autor", selecionar_primefaces_por_sufixo, ":j_id_4c_5_2_2_b_9_8_5_2_n_label", "Autor")

            # MODAIS
            if juiz_nome:
                if not attempt_twice("Criar Juiz", criar_juiz_modal_js, juiz_nome):
                    raise Exception("Juiz não pôde ser criado")

            if cpf_cnpj_contr:
                if not attempt_twice("Incluir Parte Contrária", incluir_parte_contraria_modal_js, cpf_cnpj_contr):
                    raise Exception("Falha ao incluir parte contrária")

            # RECLAMADAS AUTOMÁTICAS
            print("👥 Processando reclamadas automáticas...")
            if not attempt_twice("Adicionar reclamadas", adicionar_reclamadas_automaticamente, reclamadas_nomes, idx):
                print("⚠️ Algumas reclamadas podem não ter sido adicionadas - colunas serão marcadas como amarelas")

            # ADVOGADO PARTE CONTRÁRIA - MÉTODO SIMPLIFICADO
            if advogado_contr:
                if not attempt_twice("Selecionar Advogado Contrário", 
                                   preencher_autocomplete_simples, 
                                   ":autocompleteAdvogadoParteContrariaNome_input", advogado_contr):
                    print("⚠️ Não foi possível selecionar Advogado Contrário")

            # DATAS
            if data_distrib:
                attempt_twice("Data Distribuição", digitar_data_humano_por_sufixo, ":dataDistribuicao_input", data_distrib)
            if data_receb:
                attempt_twice("Data Citação", digitar_data_humano_por_sufixo, ":dataRecebimento_input", data_receb)

            # TIPO AÇÃO
            if tipo_processo:
                attempt_twice("Tipo de Ação", selecionar_primefaces_por_sufixo, ":comboProcessoTipo_label", tipo_processo)

            # VALOR CAUSA
            if valor_causa:
                attempt_twice("Valor da Causa", preencher_input_por_sufixo, ":amountCase_input", valor_causa)

            # ADVOGADO RESPONSÁVEL - MÉTODO SIMPLIFICADO
            if adv_resp:
                if not attempt_twice("Advogado Responsável", 
                                   preencher_autocomplete_simples, 
                                   ":autoCompleteLawyer_input", adv_resp):
                    print("⚠️ Não foi possível selecionar Advogado Responsável")

            # GESTOR JURÍDICO - MÉTODO SIMPLIFICADO
            if gestor_juridico:
                if not attempt_twice("Gestor Jurídico", 
                                   preencher_autocomplete_simples, 
                                   ":j_id_4c_5_2_2_l_9_45_3_1_2_2_1_2g_input", gestor_juridico):
                    print("⚠️ Não foi possível selecionar Gestor Jurídico")

            # UPLOAD PDF
            if not os.path.exists(pdf_path):
                print(f"⚠️ PDF não encontrado: {pdf_path}")
            attempt_twice("Anexar PDF", anexar_arquivo_por_input, pdf_path)

            # SALVAR
            if not attempt_twice("Salvar alterações", clicar_elemento_por_sufixo, "btnSalvarOpen"):
                raise Exception("Falha ao salvar")

            set_status(idx, "OK")
            print(f"✅ Finalizado com sucesso: {processo}")

        except Exception as e_row:
            marcar_erro(idx, "PROCESSAMENTO LINHA", e_row)
            traceback.print_exc()

        time.sleep(0.4)  # REDUZIDO

    # SALVAR EXCEL
    df.to_excel(EXCEL_PATH, index=False)
    print("📁 Excel atualizado com STATUS.")

    if colunas_com_erro:
        colorir_colunas_amarelo_no_excel(EXCEL_PATH, colunas_com_erro, header_rows=1)
        try:
            print("⚠️ Erros encontrados. Abrindo planilha...")
            os.startfile(EXCEL_PATH)
        except Exception as e:
            print(f"ℹ️ Não foi possível abrir a planilha: {e}")

except Exception as e_main:
    print(f"❌ ERRO GERAL: {e_main}")
    traceback.print_exc()
finally:
    try:
        driver.quit()
    except:
        pass
    print("🧹 Navegador encerrado.")