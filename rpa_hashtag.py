import os
import re
import csv
import traceback
import pandas as pd
from time import sleep
from threading import Thread
from datetime import datetime, timedelta, time as TimeObject
from selenium.webdriver.common.by import By # type: ignore
from selenium.webdriver.support.ui import WebDriverWait # type: ignore
from selenium.webdriver.support import expected_conditions as EC # type: ignore
from selenium.common.exceptions import TimeoutException, NoSuchElementException, StaleElementReferenceException, JavascriptException, ElementClickInterceptedException, ElementNotInteractableException # type: ignore
from selenium.webdriver.common.keys import Keys # type: ignore
from dotenv import load_dotenv
from colorama import init, Fore, Style
from math import ceil
import glob

# ==============================
# CONFIGURAÇÃO GERAL
# ==============================
init(autoreset=True)
load_dotenv()

URL       = "https://app.agilizone.com/login"
USER      = "davi.duarte@hashtagentrega.com"
PASSWORD  = "Hashtag#2022"
DATA_DIR  = "base"
LOG_DIR   = "logs"
MODO_DEBUG = ("False").lower() == "true"

# Variáveis para a lógica de transição de data
DATA_FORMATO = "%d/%m/%Y" 

CSV_BASE_FILENAME = "resumo_entregadores"
EXCEL_PATH = "Informes Lojas Hashtag.xlsx"
DATA_PARAM_PATH = "data_param.txt"

os.makedirs(LOG_DIR, exist_ok=True)
os.makedirs(DATA_DIR, exist_ok=True)

LOG_FILE = os.path.join(LOG_DIR , "agilizone.log")
LOG_TODAY_FILE = os.path.join(LOG_DIR , "resumo_diario.log")

# ==============================
# SISTEMA DE LOG 
# ==============================

def log(msg, level="info"):
    """Exibe mensagens coloridas e registra em arquivo."""
    cores = {
        "info": Fore.CYAN + "ℹ️ ",
        "success": Fore.GREEN + "✅ ",
        "warn": Fore.YELLOW + "⚠️ ",
        "error": Fore.RED + "❌ ",
        "debug": Fore.MAGENTA + "🔍 [DEBUG] ",
        "trace": Fore.LIGHTBLACK_EX + "⋯ ",
        "critical": Fore.RED + Style.BRIGHT + "🔥 "
    }
    prefix = cores.get(level, "")
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    message = f"[{timestamp}] [{level.upper()}] {msg}"

    print(f"{prefix}[{level.upper()}] {msg}{Style.RESET_ALL}")

    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(message + "\n")


def log_resume(msg, level="info"):
    """Exibe mensagens coloridas e registra em arquivo - resumo diário."""
    cores = {
        "info": Fore.CYAN + "ℹ️ ",
        "success": Fore.GREEN + "✅ ",
        "warn": Fore.YELLOW + "⚠️ ",
        "error": Fore.RED + "❌ ",
        "debug": Fore.MAGENTA + "🔍 [DEBUG] ",
        "trace": Fore.LIGHTBLACK_EX + "⋯ ",
        "critical": Fore.RED + Style.BRIGHT + "🔥 "
    }
    prefix = cores.get(level, "")
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    message = f"[{timestamp}] [{level.upper()}] {msg}"

    with open(LOG_TODAY_FILE, "a", encoding="utf-8") as f:
        f.write(message + "\n")


def save_artifacts(tag: str):
    """Salva HTML e screenshot apenas em modo debug."""
    if not MODO_DEBUG:
        return
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    html_path = f"{LOG_DIR}/{tag}_{timestamp}.html"
    img_path = f"{LOG_DIR}/{tag}_{timestamp}.png"
    try:
        with open(html_path, "w", encoding="utf-8") as f:
            f.write(driver.page_source)
        driver.save_screenshot(img_path)
        log(f"Artifacts salvos: {html_path}, {img_path}", "debug")
    except Exception as e:
        log(f"Falha ao salvar artifacts: {e}", "warn")


# ==============================
# FUNÇÕES AUXILIARES E SELENIUM
# ==============================

def esperar_ui_estavel(driver, timeout=10):
    """
    Aguarda o fim do backdrop e o carregamento completo da página.
    Verifica também se o <body> não está bloqueado por 'overflow:hidden',
    comum em transições do Material-UI.
    """
    try:
        wait = WebDriverWait(driver, timeout)

        # Etapa 1 — Aguarda desaparecimento do backdrop
        wait.until_not(
            EC.presence_of_element_located((By.CSS_SELECTOR, ".MuiBackdrop-root"))
        )

        # Etapa 2 — Aguarda estado 'complete' do documento
        wait.until(lambda d: d.execute_script("return document.readyState") == "complete")

        # Etapa 3 — Aguarda liberação do body (sem overflow:hidden)
        try:
            wait.until(lambda d: "overflow: hidden" not in d.find_element(By.TAG_NAME, "body").get_attribute("style"))
        except Exception:
            pass  # Em algumas telas, o estilo pode não estar disponível

        log("Interface estabilizada (body liberado e backdrop ausente).", "debug")
        return True

    except TimeoutException:
        log("⚠️ Timeout ao aguardar estabilização de interface.", "warn")
        return False


def limpar_campo_completo(driver, campo_data):
    """Limpa o campo de input usando JavaScript para desconsiderar máscaras."""
    driver.execute_script("arguments[0].value = ''; arguments[0].dispatchEvent(new Event('input'));", campo_data)
    sleep(0.5)


def clicar_botao_fechar(driver, timeout=5):
    """Tenta fechar popups/modais usando aria-label='close' ou ESC."""
    try:
        wait = WebDriverWait(driver, timeout)

        # Usa o fallback robusto para clicar no botão de fechar
        clicar_com_fallback(driver, By.CSS_SELECTOR, "button[aria-label='close']", timeout=3)

        # Espera o modal desaparecer (o container da tabela)
        wait.until(EC.invisibility_of_element_located(
            (By.XPATH, "//div[@id='table-deliverymen-container']"))
        )
        log("Popup fechado com sucesso.", "info")

    except NoSuchElementException:
        log("Botão de fechar popup não encontrado, tentando ESC.", "warn")
        driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.ESCAPE)
        sleep(1)

    except TimeoutException:
        log("Tempo excedido ao tentar fechar popup — tentando ESC.", "warn")
        driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.ESCAPE)
        sleep(1)

    except Exception as e:
        log(f"Erro ao tentar fechar popup: {e}", "error")


def limpar_texto(texto, coluna_index):
    """Limpa e formata o texto da célula com base na coluna."""
    texto_limpo = texto.strip()
    
    if coluna_index == 0:
        return texto_limpo.replace('#', '').strip()
    
    elif coluna_index == 2:
        return texto_limpo.replace('R$', '').strip()
        
    return texto_limpo


def xpath_literal(s):
    """Retorna um literal XPath seguro para strings que contêm aspas simples ou duplas."""
    if "'" not in s:
        return f"'{s}'"
    if '"' not in s:
        return f'"{s}"'
    # Se contiver ambas, usa concat()
    parts = s.split("'")
    return "concat(" + ", \"'\", ".join(f"'{p}'" for p in parts) + ")"


def proximo_dia(data_str: str, formato=DATA_FORMATO) -> str:
    """Calcula e retorna a string do dia seguinte à data fornecida."""
    try:
        data_obj = datetime.strptime(data_str, formato)
        proxima_data_obj = data_obj + timedelta(days=1)
        return proxima_data_obj.strftime(formato)
    except ValueError:
        log(f"Formato de data inválido para cálculo: {data_str}", "error")
        return data_str


def clicar_com_fallback(driver, by, seletor, timeout=10):
    """
    Tenta clicar em um elemento com fallback JavaScript e validações extras.
    Corrige erro comum: 'JavascriptException: arguments[0].click is not a function'
    """
    try:
        wait = WebDriverWait(driver, timeout)
        elemento = wait.until(EC.element_to_be_clickable((by, seletor)))
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", elemento)
        sleep(0.2)
        elemento.click()
        return True

    except (ElementClickInterceptedException, ElementNotInteractableException) as e:
        log(f"⚠️ Clique interceptado ({type(e).__name__}) — aguardando backdrop e tentando novamente.", "warn")
        esperar_ui_estavel(driver, timeout=5)
        try:
            elemento = wait.until(EC.element_to_be_clickable((by, seletor)))
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", elemento)
            elemento.click()
            return True
        except Exception as e2:
            log(f"⚠️ Segunda tentativa falhou ({type(e2).__name__}) — aplicando fallback JS.", "warn")

    except TimeoutException:
        log(f"⏳ Timeout: elemento não ficou clicável ({by}, {seletor})", "warn")
        return False

    # === Fallback via JavaScript (revisado) ===
    try:
        elemento = driver.find_element(by, seletor)
        driver.execute_script("""
            if (arguments[0] && typeof arguments[0].click === 'function') {
                arguments[0].scrollIntoView({block: 'center'});
                arguments[0].click();
            } else {
                throw new Error('Elemento inválido para clique JS.');
            }
        """, elemento)
        log(f"✅ Clique JS executado com sucesso: {seletor}", "debug")
        return True

    except JavascriptException as e:
        log(f"⚠️ Erro JS no clique: {e.msg} — recriando referência e tentando novamente.", "warn")
        try:
            sleep(1)
            elemento = wait.until(EC.element_to_be_clickable((by, seletor)))
            driver.execute_script("arguments[0].click();", elemento)
            return True
        except Exception as e3:
            log(f"❌ Falha final no fallback JS: {type(e3).__name__} - {e3}", "error")
            return False

    except Exception as e:
        log(f"❌ Erro inesperado ao tentar clicar ({type(e).__name__}): {e}", "error")
        return False

# ==============================
# ETAPAS DO FLUXO (Mantidas)
# ==============================

def login(driver, user, pwd, timeout=30):
    log("Acessando portal...", "info")
    wait = WebDriverWait(driver, timeout)
    driver.maximize_window()
    driver.get(URL)
    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[name='email']"))).send_keys(user)
    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[name='password']"))).send_keys(pwd)
    # Clique com fallback
    clicar_com_fallback(driver, By.CSS_SELECTOR, "button.MuiButton-containedPrimary")
    wait.until(EC.url_contains("app.agilizone.com"))
    sleep(1)
    esperar_ui_estavel(driver)
    log("Login realizado e interface carregada.", "success")
    save_artifacts("apos_login")


def selecionar_loja(driver, nome_loja):
    log(f"Selecionando loja inicial: {nome_loja}", "info")

    # Etapa 1 — Abre o seletor de lojas
    clicar_com_fallback(driver, By.CSS_SELECTOR, "button[title='Open']")

    # Etapa 2 — Seleciona a loja desejada no menu
    xpath_loja = f"//li[normalize-space(text())='{nome_loja}']"
    clicar_com_fallback(driver, By.XPATH, xpath_loja)

    # Etapa 3 — Aguarda a troca efetiva da loja
    wait = WebDriverWait(driver, 10)
    wait.until(EC.text_to_be_present_in_element((By.TAG_NAME, "body"), nome_loja))
    sleep(1)
    esperar_ui_estavel(driver)

    log(f"✅ Loja '{nome_loja}' selecionada e estável.", "success")


def mudar_loja(driver, nova_loja, timeout=15):
    """
    Tenta mudar para a nova loja. Em caso de falha ou interface travada, recarrega o portal e tenta novamente 1x.
    Mantém o fluxo e os logs detalhados existentes.
    """
    for tentativa in range(2):  # 🔁 Máximo de 2 tentativas: original + fallback
        log(f"Mudando loja para: {nova_loja} (tentativa {tentativa+1}/2)", "info")

        try:
            # ==========================================================
            # 🧩 Etapa 1 — Clicar no botão "Mudar loja"
            # ==========================================================
            clicar_com_fallback(driver, By.XPATH, "//button[normalize-space(text())='Mudar loja']")
            log("Botão 'Mudar loja' clicado.", "debug")

            # ==========================================================
            # 🧩 Etapa 2 — Aguardar estabilização inicial da interface
            # ==========================================================
            try:
                WebDriverWait(driver, 2).until_not(
                    EC.presence_of_element_located((By.CSS_SELECTOR, ".MuiBackdrop-root"))
                )
                WebDriverWait(driver, 2).until(
                    EC.text_to_be_present_in_element(
                        (By.CSS_SELECTOR, "span.storeName__GjA9R"),
                        nova_loja.split("/")[-1].strip()
                    )
                )
            except TimeoutException:
                log("⚠️ Timeout ao aguardar interface pós-troca de loja. Tentando estabilizar novamente...", "warn")
                esperar_ui_estavel(driver, timeout=2)

            # ==========================================================
            # 🧩 Etapa 3 — Abrir o dropdown de lojas
            # ==========================================================
            wait = WebDriverWait(driver, timeout)

            wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "button[aria-label='Open']")))
            clicar_com_fallback(driver, By.CSS_SELECTOR, "button[aria-label='Open']")
            sleep(1)

            # ==========================================================
            # 🧩 Etapa 4 — Selecionar a nova loja no menu
            # ==========================================================
            xpath_loja = f"//li[normalize-space(text())={xpath_literal(nova_loja)}]"
            clicar_com_fallback(driver, By.XPATH, xpath_loja)
            log(f"Loja '{nova_loja}' selecionada no menu.", "debug")
            sleep(2)

            # ==========================================================
            # 🧩 Etapa 5 — Verificação de erro de página offline
            # ==========================================================
            try:
                body = driver.find_element(By.TAG_NAME, "body").get_attribute("class")
                if "neterror" in body or driver.current_url.startswith("chrome-error://"):
                    log("🌐 Conexão perdida após clicar em 'Selecionar loja'. Recarregando portal...", "warn")
                    raise TimeoutException("Página offline detectada")
            except Exception:
                log("Conexão normal após troca de loja.", "debug")

            # ==========================================================
            # 🧩 Etapa 6 — Confirmar que a loja foi realmente carregada
            # ==========================================================
            wait.until(EC.text_to_be_present_in_element((By.TAG_NAME, "body"), nova_loja))
            sleep(1)
            esperar_ui_estavel(driver)
            log(f"Loja '{nova_loja}' agora está ativa e interface estabilizada.", "success")

            # ==========================================================
            # 🧩 Etapa 7 — Verificar se há alerta de loja inativa
            # ==========================================================
            try:
                WebDriverWait(driver, 3).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, ".MuiAlert-message"))
                )
                alert_elem = driver.find_element(By.CSS_SELECTOR, ".MuiAlert-message")
                alert_text = alert_elem.text.strip().lower()

                if "sua loja se encontra inativa" in alert_text or \
                   "você não tem permissão para acessar esta página" in alert_text:
                    log(f"⚠️ Loja '{nova_loja}' está inativa e será ignorada.", "warn")
                    return "inativa"

            except TimeoutException:
                pass  # Nenhum alerta encontrado, segue o fluxo normal
            except NoSuchElementException:
                pass

            # ✅ Se tudo deu certo, sai do loop
            return

        except TimeoutException:
            if tentativa == 0:
                log("⚠️ Interface travada após troca de loja. Recarregando portal e tentando novamente...", "warn")
                driver.get("https://app.agilizone.com/pedidos")
                esperar_ui_estavel(driver, timeout=15)
                continue  # tenta novamente
            else:
                log("❌ Segunda tentativa também falhou. Abortando troca de loja.", "error")
                return
        except Exception as e:
            log(f"❌ Erro inesperado ao mudar loja: {type(e).__name__} - {e}", "error")
            return


def verificar_loja_inativa(driver, nome_loja):
    """Verifica se a loja exibiu alerta de inatividade após o carregamento."""
    try:
        WebDriverWait(driver, 3).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, ".MuiAlert-message"))
        )
        alert_elem = driver.find_element(By.CSS_SELECTOR, ".MuiAlert-message")
        alert_text = alert_elem.text.strip().lower()
        if "sua loja se encontra inativa" in alert_text:
            log(f"⚠️ Loja '{nome_loja}' está inativa e será ignorada.", "warn")
            return True
    except (TimeoutException, NoSuchElementException):
        pass
    return False


def selecionar_menu_inicial(driver):
    log("Selecionando menu inicial...", "info")
    clicar_com_fallback(driver, By.CSS_SELECTOR, "svg[data-testid='MenuIcon']")
    sleep(1)


def selecionar_relatorio_pedidos(driver, max_tentativas=2):
    """
    Abre o menu lateral e seleciona o item 'Relatório de Pedidos'.
    Evita TimeoutException por renderização lenta (Material-UI/React).
    """
    for tentativa in range(max_tentativas):
        try:
            log("Navegando para o relatório de pedidos...", "info")

            # Garante que o menu lateral está presente
            WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.XPATH, "//li[normalize-space()='Relatório de Pedidos']"))
            )

            # Clica no menu com fallback
            clicar_com_fallback(driver, By.XPATH, "//li[normalize-space()='Relatório de Pedidos']")
            esperar_ui_estavel(driver, timeout=10)
            log("✅ Relatório de Pedidos acessado com sucesso.", "success")
            return True

        except TimeoutException:
            log(f"⚠️ Tentativa {tentativa + 1}: Menu ainda não disponível. Recarregando interface...", "warn")
            try:
                driver.refresh()
                esperar_ui_estavel(driver, timeout=15)
                sleep(2)
            except Exception as e:
                log(f"Erro ao recarregar interface: {e}", "warn")

        except Exception as e:
            log(f"❌ Falha ao selecionar 'Relatório de Pedidos' ({type(e).__name__}): {e}", "error")

    log(f"❌ Falha ao encontrar 'Relatório de Pedidos' após {max_tentativas} tentativas.", "error")
    return False


def aplicar_filtros(driver, data_inicio_completa: str, data_final_completa: str, timeout=15):
    """Aplica o filtro de data combinando a data base e os horários."""
    log(f"Aplicando filtros: {data_inicio_completa} a {data_final_completa}", "info")
    
    wait = WebDriverWait(driver, timeout)

    # Campo Data Inicial
    data_inicial = wait.until(EC.visibility_of_element_located(
        (By.XPATH, "//label[text()='Data inicial']/following-sibling::div//input"))
    )
    limpar_campo_completo(driver, data_inicial)
    data_inicial.send_keys(Keys.CONTROL, 'a')
    data_inicial.send_keys(Keys.BACKSPACE)
    data_inicial.send_keys(data_inicio_completa)
    
    # Campo Data Final
    data_final = wait.until(EC.visibility_of_element_located(
        (By.XPATH, "//label[text()='Data final']/following-sibling::div//input"))
    )
    limpar_campo_completo(driver, data_final)
    data_final.send_keys(Keys.CONTROL, 'a')
    data_final.send_keys(Keys.BACKSPACE)
    data_final.send_keys(data_final_completa)

    # ==========================================================
    # 🧩 Clique com fallback robusto no botão "Aplicar filtros"
    # ==========================================================
    clicar_com_fallback(
        driver,
        By.XPATH,
        "//button[contains(text(), 'Aplicar filtros')]",
        timeout=5
    )
    
    sleep(4)  # Mantém o mesmo delay para estabilidade pós-filtro
    
    log("Filtros aplicados. Aguardando a recarga dos dados.", "success")
    save_artifacts("apos_aplicar_filtros")


def clicar_resumo_entregas(driver, timeout):
    """Tenta clicar no botão 'Resumo de entregas' e aguarda o modal. Retorna True/False."""
    log("Clicando no botão 'Resumo de entregas por entregador'...", "info")
    
    # Cria um WebDriverWait local com o timeout especificado
    wait = WebDriverWait(driver, timeout)

    # ==========================================================
    # 🧩 Etapa 1 — Garante que nenhum backdrop esteja ativo antes do clique
    # ==========================================================
    try:
        WebDriverWait(driver, 5).until_not(
            EC.presence_of_element_located((By.CSS_SELECTOR, ".MuiBackdrop-root"))
        )
    except TimeoutException:
        log("⚠️ Backdrop ainda presente antes de clicar em Resumo. Tentando forçar estabilização...", "warn")
        esperar_ui_estavel(driver, timeout=5)

    # ==========================================================
    # 🧩 Etapa 2 — Tenta clicar no botão usando fallback robusto
    # ==========================================================
    try:
        clicar_com_fallback(
            driver,
            By.XPATH,
            "//button[contains(text(), 'Resumo de entregas por entregador')]",
            timeout=timeout
        )

        # Aguarda o modal abrir
        wait.until(
            EC.presence_of_element_located((By.XPATH, "//div[@id='table-deliverymen-container']"))
        )
        log("Modal Resumo de entregas carregado.", "success")
        save_artifacts("modal_resumo_aberto")
        return True

    # ==========================================================
    # 🧩 Etapa 3 — Tratamentos de exceção e fluxos alternativos
    # ==========================================================
    except TimeoutException:
        log("Timeout: Botão 'Resumo de entregas' não ficou clicável ou modal não carregou. Pulando extração.", "warn")
        return False
    except NoSuchElementException:
        log("Botão 'Resumo de entregas por entregador' não encontrado (Provavelmente sem entregas). Pulando extração.", "warn")
        return False
    except Exception as e:
        log(f"Erro inesperado ao tentar clicar em Resumo: {type(e).__name__} - {e}", "error")
        return False


def extrair_dados_tabela(driver, table_xpath, timeout=15):
    """Extrai os dados da tabela no modal e os retorna, sem exportar CSV."""
    dados_tabela = []
    
    try:
        wait = WebDriverWait(driver, timeout)

        table_locator = (By.XPATH, table_xpath)
        table = wait.until(EC.presence_of_element_located(table_locator))
        
        headers = [th.text.strip() for th in table.find_element(By.TAG_NAME, "thead").find_element(By.TAG_NAME, "tr").find_elements(By.TAG_NAME, "th")]
        headers[0] = 'Entregador'
        
        body = table.find_element(By.TAG_NAME, "tbody")
        rows = body.find_elements(By.TAG_NAME, "tr")
        log(f"Linhas encontradas na tabela: {len(rows)}", "debug")

        for idx, row in enumerate(rows):
            cells = row.find_elements(By.XPATH, "./*")
            row_data = []
            is_total_row = len(cells) < len(headers)

            for i, cell in enumerate(cells):
                try:
                    # ==========================================================
                    # ##### ALTERADO: Lógica robusta para Col 0 e 1 #####
                    # ==========================================================
                    if (i == 0 or i == 1) and not is_total_row: # Col 0 (Entregador) ou Col 1 (Pedidos)
                        try:
                            # Tenta ler como um botão primeiro (ex: Ruan Pablo)
                            cell_text = cell.find_element(By.TAG_NAME, "button").text.strip()
                        except NoSuchElementException:
                            # Se falhar, usa 'textContent' que é mais robusto que '.text'
                            # (ex: Antônio Weslley)
                            cell_text = cell.get_attribute('textContent').strip()
                        row_data.append(limpar_texto(cell_text, i))
                    
                    # Coluna 3 (Pix) pode ter 'u' (underline)
                    elif i == 3 and not is_total_row: 
                        try:
                            cell_text = cell.find_element(By.TAG_NAME, "u").text.strip()
                        except NoSuchElementException:
                            cell_text = cell.text.strip() or "-"
                        row_data.append(limpar_texto(cell_text, i))
                    
                    else: # Colunas 2 (Soma Taxas) e outras
                        cell_text = cell.text.strip()
                        row_data.append(limpar_texto(cell_text, i))
                    # ==========================================================

                except Exception as e:
                    log(f"Erro ao processar célula [{idx}, {i}]: {e}", "warn")
                    row_data.append("-")

            while len(row_data) < len(headers):
                row_data.append("-")
            
            if len(row_data) == len(headers):
                 dados_tabela.append(row_data)

        return headers, dados_tabela

    except TimeoutException:
        log("⏳ Timeout: Tabela de entregadores não carregada no modal. Retornando vazio.", "error")
        save_artifacts("timeout_tabela_modal")
        return None, None
    except Exception as e:
        log(f"❌ Erro na extração da tabela: {type(e).__name__} - {e}. Retornando vazio.", "error")
        return None, None


def exportar_dados_finais(dados_finais: list, csv_base_filename: str):
    """Exporta todos os dados coletados (incluindo cabeçalho) para um único CSV."""
    if not dados_finais:
        log("Não há dados coletados para exportar.", "warn")
        return False

    try:
        df = pd.DataFrame(dados_finais[1:], columns=dados_finais[0])

        # === BLOCO 1: Classificação do Turno ===
        def classificar_periodo(data_str):
            try:
                # Normaliza meses em português -> formato numérico
                meses_pt = {
                    "jan": "01", "fev": "02", "mar": "03", "abr": "04",
                    "mai": "05", "jun": "06", "jul": "07", "ago": "08",
                    "set": "09", "out": "10", "nov": "11", "dez": "12"
                }
                data_str = data_str.strip().lower()
                for k, v in meses_pt.items():
                    if k in data_str:
                        data_str = data_str.replace(k, v)
                # Agora formato: 14/10/2025 17h09 -> 14/10/2025 17:09
                data_str = data_str.replace("h", ":")
                data = datetime.strptime(data_str, "%d/%m/%Y %H:%M")

                hora = data.time()

                # Define as duas fronteiras dos turnos
                HORA_MANHA = TimeObject(7, 0)  # 07:00:00
                HORA_NOITE = TimeObject(18, 0) # 18:00:00

                if hora < HORA_MANHA:
                    # Antes das 07:00 (00:00 - 06:59)
                    return "madrugada"
                elif hora < HORA_NOITE:
                    # Entre 07:00 e 17:59
                    return "manhã"
                else:
                    # De 18:00 em diante (18:00 - 23:59)
                    return "noite"

            except Exception:
                return ""

        if "Data de Criação" in df.columns:
            df["Classificação do Turno"] = df["Data de Criação"].apply(classificar_periodo)
            log("Coluna 'Classificação do Turno' adicionada com sucesso.", "debug")

        # === BLOCO 2: Classificação do Dia e Pagamento por Turno ===
        def obter_dia_semana(data_str):
            try:
                data_base = datetime.strptime(data_str.strip(), "%d/%m/%Y")
                dias_semana = ["segunda", "terça", "quarta", "quinta", "sexta", "sábado", "domingo"]
                return dias_semana[data_base.weekday()]
            except Exception:
                return ""

        def calcular_pagamento(dia_semana, classificacao_turno):
            dia_semana = dia_semana.lower().strip()
            classificacao_turno = classificacao_turno.lower().strip()

            if dia_semana in ["sábado", "domingo"]:
                if classificacao_turno == "manhã" or classificacao_turno == "noite":
                    return 80.00
                elif classificacao_turno == "madrugada":
                    return 100.00
            else:  # Segunda a sexta
                if classificacao_turno == "manhã" or classificacao_turno == "noite":
                    return 70.00
                elif classificacao_turno == "madrugada":
                    return 90.00
            return 0.00

        if "Data Filtro" in df.columns:
            df["Classificação do Dia"] = df["Data Filtro"].apply(obter_dia_semana)

        if "Classificação do Turno" in df.columns and "Classificação do Dia" in df.columns:
            df["Pagamento por Turno"] = df.apply(
                lambda x: calcular_pagamento(x["Classificação do Dia"], x["Classificação do Turno"]),
                axis=1
            )
            log("Colunas 'Classificação do Dia' e 'Pagamento por Turno' adicionadas com sucesso.", "debug")

        # Remove linhas com Entregador == "Total"
        if "Entregador" in df.columns:
            df = df[df["Entregador"].str.strip().str.lower() != "total"]

        # Atualiza lista consolidada para exportar
        dados_finais = [df.columns.tolist()] + df.values.tolist()

    except Exception as e:
        log(f"Falha ao processar e classificar dados: {e}", "warn")

    # === BLOCO 3: Exportação final ===
    timestamp_str = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = os.path.join(DATA_DIR, f"{csv_base_filename}_{timestamp_str}.csv")

    with open(filename, 'w', newline='', encoding='utf-8') as csvfile:
        writer = csv.writer(csvfile, delimiter='\t')
        writer.writerow(dados_finais[0])
        for row in dados_finais[1:]:
            writer.writerow(row)

    log(f"Exportação CONSOLIDADA concluída com sucesso: {filename}", "success")
    return True

# ==============================
# LEITURA E CONVERSÃO DA PLANILHA
# ==============================

def ler_datas_parametro(file_path):
    """Lê as datas do arquivo de texto, uma por linha."""
    log(f"Lendo datas do arquivo: {file_path}", "info")
    datas = []
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            for line in f:
                date_str = line.strip()
                if date_str:
                    datas.append(date_str)
        log(f"Total de {len(datas)} datas carregadas para processamento.", "success")
        return datas
    except FileNotFoundError:
        log(f"❌ Erro: O arquivo de datas não foi encontrado no caminho: {file_path}", "error")
        return []
    except Exception as e:
        log(f"❌ Erro ao ler o arquivo de datas: {e}", "error")
        return []


def ler_dados_da_planilha(excel_path):
    """
    Lê a planilha Excel com a nova estrutura simplificada:
    Nome | Trabalha por Expediente (S/N) | Hora início | Hora fim | Loja ativa (S/N)

    Retorna uma lista de dicionários com:
    - nome (str)
    - expediente (str)
    - h_ini_obj (datetime.time)
    - h_fim_obj (datetime.time)
    - ativa (str)
    """

    log(f"Lendo planilha: {excel_path}", "info")

    try:
        df = pd.read_excel(excel_path)

        # --- Define as colunas esperadas na nova estrutura ---
        df.columns = ["nome", "expediente", "h_ini", "h_fim", "ativa"]

        def robust_parse_time(value):
            """Converte HH:MM ou datetime.time em objeto time seguro."""
            if pd.isna(value):
                return None
            if isinstance(value, TimeObject):
                return value
            if isinstance(value, datetime):
                return value.time()
            if isinstance(value, str):
                for fmt in ("%H:%M:%S", "%H:%M"):
                    try:
                        return datetime.strptime(value.strip(), fmt).time()
                    except ValueError:
                        continue
            return None

        dados_formatados = []

        for _, row in df.iterrows():
            h_ini = robust_parse_time(row["h_ini"])
            h_fim = robust_parse_time(row["h_fim"])
            ativa = str(row["ativa"]).upper().strip()
            expediente = str(row["expediente"]).upper().strip()

            if not (h_ini and h_fim):
                log(f"Ignorando loja {row['nome']} — horários inválidos ({row['h_ini']} / {row['h_fim']}).", "warn")
                continue

            dados_formatados.append({
                "nome": str(row["nome"]).strip(),
                "expediente": expediente,
                "h_ini_obj": h_ini,
                "h_fim_obj": h_fim,
                "ativa": ativa
            })

        log(f"Planilha lida com sucesso. Total de {len(dados_formatados)} lojas processadas.", "success")
        return dados_formatados

    except Exception as e:
        log(f"❌ Erro ao ler ou formatar a planilha: {type(e).__name__} - {e}", "error")
        raise

# -----------------------------
# HELPERS: normalização, coleta MUI e cruzamento
# -----------------------------

def _normalize_name_for_key(s: str) -> str:
    """
    Normaliza nomes para chaves de índice, removendo caracteres especiais,
    prefixos comuns (MP, números) e todos os espaços.
    """
    if s is None:
        return ""

    s = str(s).strip().lower()
        # Remove '#', '-', "'", '(', ')' etc.
    s = re.sub(r"[#\-'\(\)]", '', s)
    # Usa regex para remover 'mp' seguido de espaço, no início da string
    s = re.sub(r'^\s*mp\s*', '', s)
    # Usa regex para remover 'gg' seguido de espaço, no início da string
    s = re.sub(r'^\s*gg\s*', '', s)
    # Usa regex para remover qualquer número seguido de espaço, no início da string
    s = re.sub(r'^\s*\d+\s*', '', s)
    # Etapa 4: Remover TODOS os espaços restantes
    s = re.sub(r'\s+', '', s)

    return s


def _normalize_name_for_csv(s: str) -> str:
    """Remove '#' e aplica strip para apresentação no CSV."""
    if s is None:
        return ""
    s = str(s)
    s = re.sub(r'^\s*#\s*', '', s)
    s = s.strip()
    s = re.sub(r'\s+', ' ', s)
    return s


def coletar_tabela_mui(driver, timeout=15):
    """
    Coleta a tabela paginada do Material-UI (tabela principal visível).
    Força 100 linhas por página quando possível, itera todas as páginas,
    e retorna lista de dicts com chaves:
      ['Status','Valor do pedido','Taxa de entrega','Taxa do entregador','Data de Criação','Entregador','Pagamento']
    """
    dados = []
    try:
        wait = WebDriverWait(driver, timeout)

        # encontra o grid MUI visível
        grid = wait.until(EC.presence_of_element_located(
            (By.XPATH, "//div[@role='grid' and contains(@class,'MuiDataGrid-root')]")
        ))
        log("Grid MUI detectado para coleta.", "debug")

        # ==========================================================
        # ##### NOVO: Loop para resetar a paginação para a Página 1 #####
        # ==========================================================
        log("Resetando grid MUI para a primeira página...", "debug")
        btn_prev_xpath = (".//div[contains(@class,'MuiTablePagination-root')]"
                          "//button[contains(@aria-label,'anterior') or contains(@aria-label,'Previous') "
                          "or contains(@title,'anterior') or contains(@title,'Previous')]")
        
        # Tenta voltar para a página 1 (máx 10 cliques para evitar loop infinito)
        for _ in range(10): 
            try:
                btn_prev = grid.find_element(By.XPATH, btn_prev_xpath)
                
                disabled_attr = (btn_prev.get_attribute("disabled") or "").lower()
                classes = btn_prev.get_attribute("class") or ""
                
                if (not btn_prev.is_enabled()) or (disabled_attr in ("true", "disabled")) or ("Mui-disabled" in classes):
                    log("Grid MUI resetado (Página 1 alcançada).", "debug")
                    break # Botão "Anterior" está desabilitado, estamos na página 1

                # Se estiver habilitado, clica para voltar
                clicar_com_fallback(driver, By.XPATH, btn_prev_xpath)
                sleep(0.5)
                esperar_ui_estavel(driver, 0.5)

            except NoSuchElementException:
                log("Controle de paginação 'Anterior' não encontrado (página única).", "debug")
                break # Sai do loop de reset
            except Exception as e:
                log(f"Erro ao tentar resetar paginação: {e}. Interrompendo reset.", "warn")
                break
        # ==========================================================
        # ##### NOVO: Fim do reset de paginação #####
        # ==========================================================

        # Tentar setar linhas por página = 100
        try:
            combobox = grid.find_element(
                By.XPATH,
                ".//div[contains(@class,'MuiTablePagination-root')]//div[@role='combobox' or contains(@class,'MuiSelect-select')]"
            )

            driver.execute_script("arguments[0].scrollIntoView(true);", combobox)
            clicar_com_fallback(driver, By.XPATH,
                ".//div[contains(@class,'MuiTablePagination-root')]//div[@role='combobox' or contains(@class,'MuiSelect-select')]"
            )

            opt100_xpath = "//li[normalize-space()='100']"
            clicar_com_fallback(driver, By.XPATH, opt100_xpath)

            sleep(0.5)
            esperar_ui_estavel(driver, timeout=0.5)
            log("Setado 'Linhas por página' para 100 (quando disponível).", "debug")
        except Exception as e:
            log(f"Não foi possível forçar 100 linhas/página ou combobox não disponível: {e}", "debug")

        # ==========================================================
        # Função interna para extrair a página atual
        # ==========================================================
        def extrair_pagina():
            pagina = []
            try:
                rows = grid.find_elements(By.XPATH, ".//div[contains(@class,'MuiDataGrid-row')]")
                log(f"Extraindo {len(rows)} linhas da página atual do MUI.", "debug")
                for row in rows:
                    try:
                        def cell_text(field):
                            try:
                                cel = row.find_element(By.XPATH, f".//div[@role='cell' and @data-field='{field}']")
                                txt = cel.text.strip()
                                return re.sub(r'\s+', ' ', txt.replace('\n', ' ')).strip()
                            except Exception:
                                return ""
                        rec = {
                            "Status": cell_text("status"),
                            "Valor do pedido": cell_text("amount"),
                            "Taxa de entrega": cell_text("deliveryFee"),
                            "Taxa do entregador": cell_text("deliverymanFee"),
                            "Data de Criação": cell_text("date"),
                            "Entregador": _normalize_name_for_csv(cell_text("deliveryman")),
                            "Pagamento": cell_text("paymentType"),
                        }
                        pagina.append(rec)
                    except StaleElementReferenceException:
                        log("StaleElement ao extrair linha; pulando e continuando.", "warn")
                        continue
                    except Exception as ex_row:
                        log(f"Erro ao processar linha do MUI: {ex_row}", "warn")
                return pagina
            except Exception as e:
                log(f"Erro extraindo página do MUI: {e}", "warn")
                return []

        # ==========================================================
        # Loop principal de paginação
        # ==========================================================
        pagina_idx = 0
        while True:
            pagina_idx += 1
            sleep(0.5)
            esperar_ui_estavel(driver)
            pagina = extrair_pagina()
            if pagina:
                dados.extend(pagina)
            else:
                log("Nenhuma linha extraída nesta página do MUI.", "debug")

            try:
                btn_next = grid.find_element(By.XPATH,
                    ".//div[contains(@class,'MuiTablePagination-root')]"
                    "//button[contains(@aria-label,'próxima') or contains(@aria-label,'Next') "
                    "or contains(@title,'próxima') or contains(@title,'Next')]"
                )

                disabled_attr = (btn_next.get_attribute("disabled") or "").lower()
                classes = btn_next.get_attribute("class") or ""
                if (not btn_next.is_enabled()) or (disabled_attr in ("true", "disabled")) or ("Mui-disabled" in classes):
                    log("Botão 'Próxima' desabilitado — fim da paginação MUI.", "debug")
                    break

                driver.execute_script("arguments[0].scrollIntoView(true);", btn_next)
                clicar_com_fallback(driver, By.XPATH,
                    ".//div[contains(@class,'MuiTablePagination-root')]"
                    "//button[contains(@aria-label,'próxima') or contains(@aria-label,'Next') "
                    "or contains(@title,'próxima') or contains(@title,'Next')]"
                )

                sleep(0.5)
                esperar_ui_estavel(driver)
                sleep(0.6)
                log(f"Navegando para próxima página MUI (passo {pagina_idx}).", "debug")
                continue
            except NoSuchElementException:
                log("Controle de paginação MUI não encontrado — assumindo página única.", "debug")
                break
            except Exception as e:
                log(f"Erro ao avançar paginação MUI: {e} — encerrando paginação.", "warn")
                break

        log(f"Coleta MUI finalizada. Total registros coletados: {len(dados)}", "success")
        return dados

    except TimeoutException:
        log("Timeout: grid MUI não encontrado.", "error")
        return []
    except Exception as e:
        log(f"Erro inesperado em coletar_tabela_mui: {e}", "error")
        return []


def cruzar_modal_com_mui(dados_modal, dados_mui, data_base, h_ini_str, h_fim_str, nome_loja):
    """
    LEFT JOIN: para cada entregador na modal (dados_modal - lista de listas),
    gera 0..N linhas a partir de correspondências em dados_mui (lista de dicts).
    Mantém os campos MUI vazios ('') quando não houver correspondência (F1).
    
    ##### NOVO: Retorna uma tupla: (linhas_saida, houve_falha_qualidade) #####
    """
    linhas_saida = []
    houve_falha_qualidade = False # ##### NOVO: Flag de controle de qualidade
    index = {}
    for rec in dados_mui:
        chave = _normalize_name_for_key(rec.get("Entregador", ""))
        index.setdefault(chave, []).append(rec)
    log(f"Índice MUI criado com {len(index)} chaves únicas.", "debug")

    for linha_modal in dados_modal:
        try:
            entregador_raw = linha_modal[0]
            chave_pix = linha_modal[3] if len(linha_modal) > 3 else ""
        except Exception:
            entregador_raw = ""
            chave_pix = ""

        chave = _normalize_name_for_key(entregador_raw)
        encontrados = index.get(chave, []) # Obter a lista de pedidos reais UMA VEZ
        pedidos_reais = len(encontrados)

        # ==========================================================
        # ##### NOVO: INÍCIO DA VERIFICAÇÃO DE QUALIDADE #####
        # ==========================================================

        if chave == "total":
            log("Ignorando verificação de qualidade para a linha 'Total'.", "debug")
        else:
            try:
                # Tenta ler a contagem de pedidos da coluna 1 do modal
                pedidos_esperados_str = linha_modal[1].strip() 
                pedidos_esperados = int(pedidos_esperados_str)
            except (ValueError, IndexError, TypeError):
                log(f"Não foi possível ler a contagem de pedidos (Coluna 1) para '{entregador_raw}'. Pulando verificação.", "warn")
                pedidos_esperados = -1 # Marcar como inválido para não comparar
        
            # Compara o esperado (Modal) vs. o real (MUI)
            if pedidos_esperados >= 0 and pedidos_esperados != pedidos_reais:
                houve_falha_qualidade = True # Seta a flag principal da função
                log(f"❌ FALHA NA QUALIDADE DE DADOS ({nome_loja})", "critical")
                log(f"  Entregador: {entregador_raw}", "critical")
                log(f"  Pedidos Esperados (Modal): {pedidos_esperados}", "critical")
                log(f"  Pedidos Encontrados (MUI): {pedidos_reais}", "critical")
            
            elif pedidos_esperados >= 0:
                log(f"✅ Verificação de qualidade OK para {entregador_raw} ({pedidos_esperados} pedidos).", "debug")

        # ==========================================================
        # ##### NOVO: FIM DA VERIFICAÇÃO DE QUALIDADE #####
        # ==========================================================

        # A lógica de JOIN (cruzamento) continua a mesma
        if encontrados:
            for rec in encontrados:
                linha = [
                    _normalize_name_for_csv(entregador_raw),
                    chave_pix or "",
                    data_base or "",
                    h_ini_str or "",
                    h_fim_str or "",
                    nome_loja or "",
                    rec.get("Status", "") or "",
                    rec.get("Valor do pedido", "") or "",
                    rec.get("Taxa de entrega", "") or "",
                    rec.get("Taxa do entregador", "") or "",
                    rec.get("Data de Criação", "") or "",
                    rec.get("Pagamento", "") or ""
                ]
                linhas_saida.append(linha)
        else:
            # Mantém o entregador mesmo sem pedidos na MUI (LEFT JOIN)
            linha = [
                _normalize_name_for_csv(entregador_raw),
                chave_pix or "",
                data_base or "",
                h_ini_str or "",
                h_fim_str or "",
                nome_loja or "",
                "", "", "", "", "", ""
            ]
            linhas_saida.append(linha)

    log(f"Cruzamento concluído. Linhas produzidas: {len(linhas_saida)}", "debug")
    return linhas_saida, houve_falha_qualidade # ##### NOVO: Retorna a tupla

# -----------------------------
# FUNÇÃO REESCRITA: processar_lojas_e_turnos_por_data
# -----------------------------

def processar_lojas_e_turnos_por_data(driver, lista_lojas, data_base, csv_base_filename, is_first_run):
    """
    Processa todas as lojas ativas para uma data base específica,
    aplicando os filtros de Data D + Hora início e Data D+1 + Hora fim.
    """

    MAX_QUALIDADE_TENTATIVAS = 3 # Número de tentativas para a verificação Modal vs. MUI

    # ==========================================================
    # 🧩 Etapa 0: Inicialização e filtragem
    # ==========================================================
    dados_consolidados = []
    status_lojas = {}  # Controle de sucesso/pendente por loja

    lojas_ativas = [l for l in lista_lojas if l['ativa'] == 'S']
    if not lojas_ativas:
        log("Nenhuma loja ativa para processar nesta rodada.", "warn")
        return False

    loja_atual = ""
    primeira_loja_nome = lojas_ativas[0]['nome']

    # ==========================================================
    # 🧩 Etapa 1: Login e seleção da primeira loja
    # ==========================================================
    if is_first_run:
        log(f"🧩 Etapa 1: Iniciando login e configuração da primeira loja ({primeira_loja_nome})...", "info")

        login(driver, USER, PASSWORD)
        selecionar_loja(driver, primeira_loja_nome)
        esperar_ui_estavel(driver, timeout=10)

        # --- Verificação de loja inativa logo após login ---
        if verificar_loja_inativa(driver, primeira_loja_nome):
            status_lojas[primeira_loja_nome] = "PENDENTE"
            log(f"⚠️ Primeira loja '{primeira_loja_nome}' ignorada por estar inativa.", "warn")

            # Remove a loja inativa da lista
            lojas_ativas = [l for l in lojas_ativas if l['nome'] != primeira_loja_nome]
            loja_atual = None

        else:
            loja_atual = primeira_loja_nome
            selecionar_menu_inicial(driver)
            selecionar_relatorio_pedidos(driver)
            log("Sincronizando interface e fechando popups antes da 1ª interação com filtros.", "info")
            clicar_botao_fechar(driver)
            sleep(1)
            esperar_ui_estavel(driver)

    # ==========================================================
    # 🧩 Etapa 2: Loop principal sobre lojas
    # ==========================================================
    for loja_info in lojas_ativas:
        nome_loja = loja_info['nome']
        h_ini_obj = loja_info['h_ini_obj']
        h_fim_obj = loja_info['h_fim_obj']

        log(f"\n======================================", "info")
        log(f"🧩 Etapa 2: INICIANDO PROCESSAMENTO → {nome_loja} | Data: {data_base}", "info")

        # ----------------------------------------------------------
        # Troca de loja, se necessário
        # ----------------------------------------------------------
        if nome_loja != loja_atual:
            log(f"🔄 Mudando de loja: {loja_atual or 'N/D'} → {nome_loja}", "info")
            status_mudanca = mudar_loja(driver, nome_loja)

            # --- Tratamento de loja inativa ---
            if status_mudanca == "inativa":
                status_lojas[nome_loja] = "PENDENTE"
                log(f"⚠️ Loja '{nome_loja}' ignorada por estar inativa.", "warn")
                continue  # pula para a próxima loja

            # --- Fluxo normal após troca ---
            esperar_ui_estavel(driver, timeout=30)
            sleep(1)
            esperar_ui_estavel(driver, timeout=30)
            loja_atual = nome_loja

            # Garante que estamos no menu correto antes de aplicar filtros
            selecionar_menu_inicial(driver)
            selecionar_relatorio_pedidos(driver)
            clicar_botao_fechar(driver)
            sleep(1)
            esperar_ui_estavel(driver)
            log("✅ Sincronização pós-troca de loja realizada.", "debug")

        # ==========================================================
        # 🧩 Etapa 3: Aplicar filtros e abrir modal de entregas
        # ==========================================================
        data_filtro_ini = data_base
        data_filtro_fim = proximo_dia(data_base)

        h_ini_str = h_ini_obj.strftime("%H:%M")
        h_fim_str = h_fim_obj.strftime("%H:%M")

        data_inicio_completa = f"{data_filtro_ini} {h_ini_str}"
        data_final_completa = f"{data_filtro_fim} {h_fim_str}"

        log(f"📅 Aplicando filtros: {data_inicio_completa} → {data_final_completa}", "info")

        aplicar_filtros(driver, data_inicio_completa, data_final_completa)
        clicar_botao_fechar(driver)

        # 🔁 Tentativas de abrir o modal
        tentativas = 0
        max_tentativas = 3
        modal_aberto = False

        while tentativas < max_tentativas and not modal_aberto:
            tentativas += 1
            log(f"Tentativa {tentativas} de abrir o modal de entregas...", "info")
            modal_aberto = clicar_resumo_entregas(driver, timeout=2)
            if not modal_aberto:
                sleep(2)

        # ==========================================================
        # 🧩 Etapa 4: Extração e cruzamento dos dados (COM RETRY)
        # ==========================================================
        
        ##### Define o número de tentativas #####
        MAX_QUALIDADE_TENTATIVAS = 3 # Ou use a constante definida anteriormente

        if modal_aberto:
            headers, dados_brutos = extrair_dados_tabela(
                driver, "//div[@id='table-deliverymen-container']//table"
            )

            if dados_brutos:
                log(f"  → Dados do modal (Controle Mestre) coletados (linhas: {len(dados_brutos)}).", "debug")
            else:
                log("  → Modal abriu mas não houve linhas extraídas.", "debug")
                dados_brutos = [] # Garante que não é None

            clicar_botao_fechar(driver)

            ##### Variáveis para armazenar o resultado do loop #####
            linhas_cruzadas_final = []  # Armazena os dados da *melhor* tentativa
            falha_qualidade_final = True  # Começa assumindo que vai falhar

            ##### Início do loop de tentativas de qualidade #####
            for tentativa_qualidade in range(MAX_QUALIDADE_TENTATIVAS):
                log(f"  Iniciando tentativa de coleta e verificação de qualidade {tentativa_qualidade + 1}/{MAX_QUALIDADE_TENTATIVAS}...", "info")
                
                try:
                    # ETAPA A (Repetível): Coleta os dados detalhados da MUI
                    dados_mui_tentativa = coletar_tabela_mui(driver)

                    # ETAPA B (Repetível): Cruza e Verifica a Qualidade
                    linhas_cruzadas_tentativa, falha_qualidade_tentativa = cruzar_modal_com_mui(
                        dados_brutos, 
                        dados_mui_tentativa, 
                        data_base, h_ini_str, h_fim_str, nome_loja
                    )

                    # ETAPA C (Garante NÃO DUPLICAÇÃO):
                    # SUBSTITUÍDO o resultado anterior. Não soma.
                    # Se T1 deu 11 e T2 deu 12, 'linhas_cruzadas_final' será 12.
                    linhas_cruzadas_final = linhas_cruzadas_tentativa
                    falha_qualidade_final = falha_qualidade_tentativa

                    # ETAPA D (Condição de Sucesso):
                    if not falha_qualidade_tentativa:
                        log(f"  ✅ SUCESSO na verificação de qualidade (Tentativa {tentativa_qualidade + 1}). Dados reconciliados.", "success")
                        break # Sai do loop de tentativas de qualidade
                    else:
                        log(f"  ⚠️ Falha na verificação de qualidade (Tentativa {tentativa_qualidade + 1}). Retentando...", "warn")
                        sleep(2) # Pequena pausa antes de re-coletar a MUI

                except Exception as e:
                    # Captura erro na coleta (ex: coletar_tabela_mui falhou)
                    log(f"❌ Erro durante a tentativa de coleta/cruzamento {tentativa_qualidade + 1}: {e}", "error")
                    falha_qualidade_final = True # Marca esta tentativa como falha
                    sleep(2)
                    # O loop continua para a próxima tentativa

            # Processa o resultado da *última* (ou única bem-sucedida) tentativa
            headers_finais = [
                'Entregador', 'Chave Pix', 'Data Filtro', 'Hora Início Filtro',
                'Hora Fim Filtro', 'Loja', 'Status', 'Valor do pedido',
                'Taxa de entrega', 'Taxa do entregador', 'Data de Criação', 'Pagamento'
            ]
            if not dados_consolidados:
                dados_consolidados.append(headers_finais)

            # Adiciona os dados da tentativa final (seja ela 11 ou 12)
            for linha in linhas_cruzadas_final:
                dados_consolidados.append(linha)

            # Define o status da loja com base no resultado *final* do loop
            if falha_qualidade_final:
                log(f"  ⚠️ Dados de {nome_loja} adicionados, mas COM AVISO DE QUALIDADE (divergência Modal/MUI) após {MAX_QUALIDADE_TENTATIVAS} tentativas.", "warn")
                status_lojas[nome_loja] = "SUCESSO (AVISO QUALIDADE)"
            else:
                log(f"  ✅ Dados cruzados (modal × MUI) adicionados ao consolidado.", "success")
                status_lojas[nome_loja] = "SUCESSO"

        else:
            log(f"⚠️ Falha após {max_tentativas} tentativas de abrir o modal. Nenhuma extração realizada.", "warn")
            status_lojas[nome_loja] = "PENDENTE"

        sleep(1)
        log(f"🏁 Fim do processamento de {nome_loja} na data {data_base}.", "success")

    # ==========================================================
    # 🧩 Etapa 5: Resumo final e exportação
    # ==========================================================
    total_sucesso = sum(1 for s in status_lojas.values() if s == "SUCESSO")
    total_aviso_qualidade = sum(1 for s in status_lojas.values() if s == "SUCESSO (AVISO QUALIDADE)")
    total_pendente = sum(1 for s in status_lojas.values() if s == "PENDENTE")

    log("======================================", "info")
    log("📊 RESUMO FINAL DE CAPTURA DE DADOS:", "info")
    log_resume("======================================", "info")
    log_resume("📊 RESUMO FINAL DE CAPTURA DE DADOS:", "info")

    for loja, status in status_lojas.items():
        log(f"  - {loja}: {status}", "info")
        log_resume(f"  - {loja}: {status}", "info")
    
    log(f"Resumo geral → SUCESSO: {total_sucesso}; AVISO QUALIDADE: {total_aviso_qualidade}; PENDENTE: {total_pendente}", "success")
    log_resume(f"Resumo geral → SUCESSO: {total_sucesso}; AVISO QUALIDADE: {total_aviso_qualidade}; PENDENTE: {total_pendente}", "success")

    return exportar_dados_finais(dados_consolidados, csv_base_filename)

# ==========================================================
# 🚀 EXECUÇÃO MULTITHREAD (DIVISÃO DE LOJAS EM LOTES)
# ==========================================================

def iniciar_driver_economico():
    """Cria um driver otimizado para baixa performance (sem GPU e imagens)."""
    from selenium import webdriver # type: ignore
    from selenium.webdriver.chrome.options import Options # type: ignore

    chrome_options = Options()
    #chrome_options.add_argument("--headless=new")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--window-size=800,600")
    chrome_options.add_experimental_option("prefs", {
        "profile.managed_default_content_settings.images": 2,
        "profile.default_content_setting_values.notifications": 2,
    })
    driver = webdriver.Chrome(options=chrome_options)
    driver.set_page_load_timeout(60)
    return driver


def dividir_em_lotes(lojas, qtd_lotes, limite_por_lote=30):
    """Divide a lista de lojas em N lotes equilibrados respeitando o limite máximo por lote."""
    if qtd_lotes <= 1:
        return [lojas[i:i + limite_por_lote] for i in range(0, len(lojas), limite_por_lote)]
    tam = min(ceil(len(lojas) / qtd_lotes), limite_por_lote)
    return [lojas[i:i + tam] for i in range(0, len(lojas), tam)]


def sufixo_thread(thread_id):
    """Gera sufixo padronizado para identificação por thread."""
    return f"_T{thread_id}"


def worker_thread(thread_id, lojas_subset, data_base, csv_base_filename):
    """Executa o processamento de um lote de lojas em uma thread independente."""
    try:
        log(f"[THREAD {thread_id}] Iniciando driver otimizado...", "info")
        driver = iniciar_driver_economico()
        is_first_run = True
        csv_thread_name = f"{csv_base_filename}{sufixo_thread(thread_id)}"

        log(f"[THREAD {thread_id}] Iniciando lote ({len(lojas_subset)} lojas)...", "info")
        processar_lojas_e_turnos_por_data(
            driver=driver,
            lista_lojas=lojas_subset,
            data_base=data_base,
            csv_base_filename=csv_thread_name,
            is_first_run=is_first_run
        )

        log(f"[THREAD {thread_id}] Lote concluído com sucesso.", "success")

    except Exception as e:
        erro_completo = traceback.format_exc()
        log(f"[THREAD {thread_id}] ERRO não tratado: {type(e).__name__} - {e}\n--- STACK TRACE ---\n{erro_completo}", "error")
    finally:
        try:
            driver.quit()
        except:
            pass
        log(f"[THREAD {thread_id}] Driver encerrado.", "info")


def mesclar_csvs_parciais(csv_base_filename, output_prefix=None):
    """Mescla todos os CSVs gerados por threads (_T1, _T2, etc.) em um único arquivo final."""
    pattern = os.path.join(DATA_DIR, f"{csv_base_filename}_T*_*.csv")
    arquivos = sorted(glob.glob(pattern))
    if not arquivos:
        log("Nenhum CSV parcial encontrado para mesclagem.", "warn")
        return False

    dfs = []
    for arq in arquivos:
        try:
            df = pd.read_csv(arq, sep="\t", dtype=str)
            dfs.append(df)
            log(f"Mesclagem: carregado {arq} ({len(df)} linhas).", "debug")
        except Exception as e:
            log(f"Falha ao ler {arq}: {e}", "warn")

    if not dfs:
        log("Nenhum CSV válido encontrado para mesclagem.", "warn")
        return False

    df_final = pd.concat(dfs, ignore_index=True)
    timestamp_str = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_prefix = output_prefix or csv_base_filename
    arquivo_final = os.path.join(DATA_DIR, f"{out_prefix}_{timestamp_str}.csv")

    df_final.to_csv(arquivo_final, sep="\t", index=False, encoding="utf-8")
    log(f"✅ CSV final mesclado gerado: {arquivo_final} (total {len(df_final)} linhas)", "success")
    log_resume(f"✅ CSV final mesclado gerado: {arquivo_final} (total {len(df_final)} linhas)", "success")
    return True


def executar_multithread(lojas_todas, datas_a_processar, csv_base_filename, num_threads=2, stagger_seg=5):
    """
    Divide as lojas em N threads, executa cada lote em paralelo e mescla os CSVs.
    Ideal para máquinas modestas (2 threads, Chrome headless).
    """
    lojas_ativas = [l for l in lojas_todas if l.get("ativa") == "S"]
    if not lojas_ativas:
        log("Nenhuma loja ativa encontrada para processamento.", "warn")
        log_resume("Nenhuma loja ativa encontrada para processamento.", "warn")
        return

    lotes = dividir_em_lotes(lojas_ativas, num_threads, limite_por_lote=30)

    log(f"Preparando execução multithread ({num_threads} threads)...", "info")
    log(f"Tamanhos dos lotes: {[len(l) for l in lotes]}", "debug")
    log_resume(f"Preparando execução multithread ({num_threads} threads)...", "info")
    log_resume(f"Tamanhos dos lotes: {[len(l) for l in lotes]}", "debug")

    for data_base in datas_a_processar:
        log(f">>>> EXECUÇÃO MULTITHREAD PARA A DATA: {data_base} <<<<", "info")
        log_resume(f">>>> EXECUÇÃO MULTITHREAD PARA A DATA: {data_base} <<<<", "info")
        threads = []

        for idx, subset in enumerate(lotes, start=1):
            t = Thread(
                target=worker_thread,
                args=(idx, subset, data_base, csv_base_filename),
                daemon=True
            )
            threads.append(t)
            t.start()
            sleep(stagger_seg)  # evita sobrecarga simultânea no login

        for t in threads:
            t.join()

        mesclar_csvs_parciais(
            csv_base_filename,
            output_prefix=f"{csv_base_filename}_FINAL_{data_base.replace('/', '-')}"
        )

    log("✅ Execução multithread concluída com sucesso.", "success")
    log_resume("✅ Execução multithread concluída com sucesso.", "success")

# ==========================================================
# EXECUÇÃO PRINCIPAL
# ==========================================================
if __name__ == "__main__":
    try:
        tempo_inicio = datetime.now()
        log("XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX", "info")
        log("Iniciando automação Agilizone Hashtag...", "info")        
        log(f"Modo atual: {'DEBUG' if MODO_DEBUG else 'PRODUÇÃO'}", "debug" if MODO_DEBUG else "info")

        log_resume("XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX", "info")
        log_resume("Iniciando automação Agilizone Hashtag...", "info")        
        log_resume(f"Modo atual: {'DEBUG' if MODO_DEBUG else 'PRODUÇÃO'}", "debug" if MODO_DEBUG else "info")

        log("Limpando arquivos CSV temporários de execuções anteriores...", "debug")
        padrao_limpeza = os.path.join(DATA_DIR, "resumo_entregadores_T*.csv")
        arquivos_antigos = glob.glob(padrao_limpeza)
        
        if arquivos_antigos:
            for f in arquivos_antigos:
                try:
                    os.remove(f)
                    log(f"Removido CSV temporário antigo: {f}", "debug")
                except Exception as e:
                    log(f"Não foi possível remover o arquivo {f}: {e}", "warn")
        else:
            log("Nenhum CSV temporário antigo encontrado.", "debug")

        datas_a_processar = ler_datas_parametro(DATA_PARAM_PATH)
        dados_das_lojas = ler_dados_da_planilha(EXCEL_PATH)

        if not datas_a_processar or not dados_das_lojas:
            log("❌ Nenhuma data ou loja disponível para processamento.", "error")
            raise Exception("Dados de entrada insuficientes.")

        # ⚙️ CONFIGURAÇÃO DE EXECUÇÃO
        USAR_MULTITHREAD = True      # ← alternar para False para modo tradicional
        NUM_THREADS = 2              # ← ideal para PCs modestos
        STAGGER_SEGUNDOS = 10        # ← atraso entre logins

        if USAR_MULTITHREAD:
            executar_multithread(
                lojas_todas=dados_das_lojas,
                datas_a_processar=datas_a_processar,
                csv_base_filename=CSV_BASE_FILENAME,
                num_threads=NUM_THREADS,
                stagger_seg=STAGGER_SEGUNDOS
            )
        else:
            driver = iniciar_driver_economico()
            primeira_execucao = True
            for data_parametro in datas_a_processar:
                log(f"\n>>>> INICIANDO CICLO PARA A DATA: {data_parametro} <<<<", "info")
                log_resume(f"\n>>>> INICIANDO CICLO PARA A DATA: {data_parametro} <<<<", "info")
                processar_lojas_e_turnos_por_data(
                    driver, dados_das_lojas, data_parametro, CSV_BASE_FILENAME, primeira_execucao
                )
                primeira_execucao = False
            driver.quit()

        duracao = datetime.now() - tempo_inicio
        log(f"Duração total: {duracao}", "info")
        log("XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX", "info")
        log_resume(f"Duração total: {duracao}", "info")
        log_resume("XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX", "info")

    except Exception as e:
        log(f"Erro FATAL no fluxo principal: {type(e).__name__} - {e}", "error")
        save_artifacts("erro_fatal")
