import shutil
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from datetime import datetime, timedelta
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options

# ---------------- FUNÇÕES AUXILIARES ----------------
def get_previous_month_dates():
    today = datetime.today()
    first_day_of_current_month = today.replace(day=1)
    last_day_of_previous_month = first_day_of_current_month - timedelta(days=1)
    first_day_of_previous_month = last_day_of_previous_month.replace(day=1)
    return first_day_of_previous_month.strftime("%d/%m/%Y"), last_day_of_previous_month.strftime("%d/%m/%Y")

def verificar_download_concluido(caminho_downloads, empresa_id):
    relatorio_final = f"relatorio ({empresa_id}).xlsx"
    caminho_relatorio_final = os.path.join(caminho_downloads, relatorio_final)
    arquivos = os.listdir(caminho_downloads)
    for arquivo in arquivos:
        if f"relatorio ({empresa_id})" in arquivo.lower() and not arquivo.endswith(".crdownload"):
            return os.path.join(caminho_downloads, arquivo)
    for arquivo in arquivos:
        nome_lower = arquivo.lower()
        if "relatorio -" in nome_lower and not arquivo.endswith(".crdownload"):
            caminho_antigo = os.path.join(caminho_downloads, arquivo)
            os.rename(caminho_antigo, caminho_relatorio_final)
            return caminho_relatorio_final
    return None

def verificar_erro(driver):
    try:
        erro_elemento = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located(
                (By.XPATH, "//div[@class='ui-widget ui-widget-content ui-corner-all ui-state-highlight']")
            )
        )
        if erro_elemento.is_displayed():
            return True
    except:
        return False

def copiar_e_renomear_arquivo(caminho_origem, caminho_destino, nome_arquivo_destino):
    if os.path.exists(caminho_origem):
        shutil.copy(caminho_origem, os.path.join(caminho_destino, nome_arquivo_destino))

def aguardar_url_conter(driver, texto, timeout=30):
    """Espera até a URL conter determinado texto."""
    WebDriverWait(driver, timeout).until(EC.url_contains(texto))

def preencher_campo(driver, by, locator, valor, timeout=20):
    """Espera campo ficar visível, limpa e digita."""
    campo = WebDriverWait(driver, timeout).until(
        EC.visibility_of_element_located((by, locator))
    )
    campo.clear()
    campo.send_keys(valor)
    return campo

# ---------------- CONFIGURAÇÕES ----------------
data_inicio, data_fim = get_previous_month_dates()
data_limite = "31/12/2021"
chrome_driver_path = r"C:\Users\caio.silva\Desktop\CAIO\Aplicativos\chromedriver-win32\chromedriver.exe"
caminho_origem_relatorio = r"C:\Users\caio.silva\Desktop\CAIO\Aplicativos\relatorio.xlsx"
caminho_downloads = r"C:\Users\caio.silva\Downloads"
empresas_excluidas = {50, 51, 53, 66, 74, 124}

EMAIL = "caio.silva@cury.net"
SENHA = "@Negro10..."

# ---------------- DRIVER ----------------
options = Options()
options.add_argument("--disable-notifications")
options.add_argument("--log-level=3")          # suprime erros GCM/PHONE_REGISTRATION no console
options.add_argument("--disable-background-networking")
options.add_argument("--disable-sync")
options.add_argument("--no-first-run")
options.add_experimental_option("excludeSwitches", ["enable-logging"])  # remove DevTools spam

service = Service(chrome_driver_path)
driver = webdriver.Chrome(service=service, options=options)
wait = WebDriverWait(driver, 30)

driver.get("https://curyempreendimentos.sienge.com.br/sienge/index.jsp")

try:
    # ── Passo 1: botão "Entrar com Sienge ID" ──────────────────────────────
    print("Clicando em 'Entrar com Sienge ID'...")
    wait.until(EC.element_to_be_clickable((By.ID, "btnEntrarComSiengeID"))).click()

    # ── Passo 2: campo de e-mail na página do Sienge (antes do redirect MS) ─
    # Aguarda o campo aparecer — pode estar na mesma página ou em redirect
    print("Preenchendo e-mail no Sienge...")
    preencher_campo(driver, By.NAME, "email", EMAIL)

    wait.until(EC.element_to_be_clickable(
        (By.XPATH, "//button[contains(translate(text(),'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ'), 'CONTINUAR')]")
    )).click()

    # ── Passo 3: Microsoft login — campo e-mail (ID i0116) ──────────────────
    print("Preenchendo e-mail na Microsoft...")
    preencher_campo(driver, By.ID, "i0116", EMAIL, timeout=30)
    wait.until(EC.element_to_be_clickable((By.ID, "idSIButton9"))).click()

    # ── Passo 4: Microsoft — campo senha ────────────────────────────────────
    print("Preenchendo senha...")
    preencher_campo(driver, By.ID, "i0118", SENHA, timeout=30)
    wait.until(EC.element_to_be_clickable((By.ID, "idSIButton9"))).click()

    # ── Passo 5: "Manter conectado?" — clica em Sim se aparecer ─────────────
    try:
        btn_sim = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "idSIButton9"))
        )
        btn_sim.click()
        print("Clicou em 'Sim' (manter conectado).")
    except:
        print("Tela 'manter conectado' não apareceu, continuando...")

    # ── Passo 6: Aguarda redirect de volta ao Sienge ─────────────────────────
    print("Aguardando redirect para o Sienge...")
    WebDriverWait(driver, 60).until(EC.url_contains("sienge.com.br"))
    print("Login concluído!")
    time.sleep(3)

    # ── Navega para o relatório ───────────────────────────────────────────────
    driver.get("https://curyempreendimentos.sienge.com.br/sienge/8/index.html#/common/page/1579")
    time.sleep(5)
    driver.refresh()
    time.sleep(2)

    iframe = WebDriverWait(driver, 60).until(
        EC.presence_of_element_located((By.TAG_NAME, 'iframe'))
    )
    driver.switch_to.frame(iframe)

    # ---------------- FILTROS ----------------
    botao_filtro = WebDriverWait(driver, 60).until(
        EC.element_to_be_clickable((By.NAME, "toggleFiltro"))
    )
    driver.execute_script("arguments[0].scrollIntoView();", botao_filtro)
    botao_filtro.click()
    time.sleep(2)

    checkbox_juros = WebDriverWait(driver, 30).until(
        EC.presence_of_element_located((By.ID, 'entity.flCalcularJurosMulta'))
    )
    if not checkbox_juros.is_selected():
        driver.execute_script("arguments[0].click();", checkbox_juros)

    checkbox_congelar = WebDriverWait(driver, 30).until(
        EC.presence_of_element_located((By.ID, 'entity.flCongelarAcrescimos'))
    )
    if not checkbox_congelar.is_selected():
        driver.execute_script("arguments[0].click();", checkbox_congelar)

    preencher_campo(driver, By.ID, 'entity.dtLimiteParaNaoCongelarAcrescimos', data_limite)
    preencher_campo(driver, By.ID, 'entity.nuLimiteDiasCalculoEncargos', "0")
    preencher_campo(driver, By.ID, 'entity.dtCorrecao', data_fim)
    preencher_campo(driver, By.ID, 'entity.dtPosicaoEm', data_fim)
    preencher_campo(driver, By.ID, 'entity.dtAcrescimo', data_fim)
    print("Filtros configurados.")

    # ---------------- LOOP DAS EMPRESAS ----------------
    for empresa_id in range(1, 343):
        if empresa_id in empresas_excluidas:
            continue
        try:
            print(f"\nProcessando empresa {empresa_id}...")
            campo_empresa = WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.ID, 'cdEmpresaView'))
            )
            campo_empresa.clear()
            campo_empresa.send_keys(str(empresa_id))
            WebDriverWait(driver, 30).until(
                EC.element_to_be_clickable((By.ID, 'nmEmpresa'))
            ).click()   
            time.sleep(2)

            if verificar_erro(driver):
                print(f"Empresa {empresa_id} retornou erro, pulando...")
                continue

            botao_exportar = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.NAME, 'btExportarParaXlsx'))
            )
            botao_exportar.click()
            print("Aguardando sistema liberar...")

            tempo_inicio = time.time()
            while time.time() - tempo_inicio < 300:
                try:
                    campo_teste = driver.find_element(By.ID, "cdEmpresaView")
                    campo_teste.click()
                    break
                except:
                    time.sleep(5)

            tempo_inicio_download = time.time()
            arquivo = None
            while time.time() - tempo_inicio_download < 60:
                arquivo = verificar_download_concluido(caminho_downloads, empresa_id)
                if arquivo:
                    break
                print("Aguardando download finalizar...")
                time.sleep(7)

            if not arquivo:
                print("Download não encontrado, usando fallback...")
                novo_nome = f"relatorio ({empresa_id}).xlsx"
                copiar_e_renomear_arquivo(caminho_origem_relatorio, caminho_downloads, novo_nome)
                continue

            print(f"Download concluído: {arquivo}")

        except Exception as e:
            print(f"Erro empresa {empresa_id}: {e}")
            continue

    print("Processo finalizado.")

except Exception as e:
    print(f"Erro fatal: {e}")

finally:
    driver.quit()
    print("Navegador fechado.")