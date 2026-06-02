# Automação Unificada - Empregados + Estagiários
from datetime import datetime
import re  # Para extrair a data (dd/mm/aaaa) do texto do detalhe
import time
import os
import sys
import pandas as pd
from dotenv import load_dotenv
import logging
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import smtplib
from email.mime.text import MIMEText

# ================= CONFIG =================
load_dotenv()

email = os.getenv("EMAIL")
senha = os.getenv("SENHA")
url = os.getenv("URL_ENTRADA")

# Antecedência (em dias) para criar a tarefa com base na data de início do gozo.
# Ex.: se hoje é dia 1 e a antecedência é 10, cria tarefas cujo início do gozo
# seja até o dia 11 (inclusive). Datas mais distantes ficam pendentes e são
# reavaliadas nas próximas execuções.
DIAS_ANTECEDENCIA = 10

TIPOS = [
    {
    "nome": "aviso previo ferias",
    "arquivo": "dados_aviso_previo.xlsx",
    "grid_id": "gridVacationList",
    "aba_calculo_ferias": False,
    "titulo_tarefa": "DP - FERIAS",
    "xpath_expandir": '//*[@id="items-per-page-0"]',
    "id_minimo": 1290
},
{
    "nome": "calculo",
    "arquivo": "dados_calculo_de_ferias.xlsx",
    "grid_id": "gridVacationCalculationList",
    "aba_calculo_ferias": True,
    "titulo_tarefa": "DP - FERIAS",
    "xpath_expandir": '//*[@id="items-per-page-1"]',
    "id_minimo": 800
}
]

# ================= LOG =================
def configurar_log():

    if getattr(sys, 'frozen', False):
        base_dir = os.path.dirname(sys.executable)
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))

    MESES_PT = {
        1: "janeiro", 2: "fevereiro", 3: "março", 4: "abril",
        5: "maio", 6: "junho", 7: "julho", 8: "agosto",
        9: "setembro", 10: "outubro", 11: "novembro", 12: "dezembro"
    }

    agora = datetime.now()
    nome_projeto = "log_Ferias_calculo"
    data_hoje = agora.strftime("%Y-%m-%d")
    mes_pasta = f"{MESES_PT[agora.month]}_{agora.year}"

    logs_dir = os.path.join(base_dir, "logs", "ferias", mes_pasta)
    os.makedirs(logs_dir, exist_ok=True)

    arquivo_log = os.path.join(
        logs_dir,
        f"{nome_projeto}_{data_hoje}.log"
    )

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[
            logging.FileHandler(
                arquivo_log,
                encoding="utf-8"
            ),
            logging.StreamHandler()
        ]
    )

# ================= EXCEL =================
def carregar_empresas():
    df = pd.read_excel("Empresas_1.xlsx")

    df["Apelido"] = df["Apelido"].astype(str).str.strip().str.lower()

    mapa = dict(zip(df["Apelido"], df["Código"]))

    return mapa


def montar_apelido_com_codigo(apelido, mapa):
    apelido_limpo = str(apelido).strip().lower()

    codigo = mapa.get(apelido_limpo)

    if codigo:
        apelido_curto = apelido_limpo[:2]
        return f"{codigo} - {apelido_curto}"
    else:
        logging.warning(f"Apelido não encontrado: {apelido}")
        return apelido


def carregar_base(arquivo):
    try:
        df = pd.read_excel(arquivo)

        if "status" not in df.columns:
            df["status"] = ""

        if "protocolo" not in df.columns:
            df["protocolo"] = ""

        if "data_inicio_gozo" not in df.columns:
            df["data_inicio_gozo"] = ""

        if "descricao" not in df.columns:
            df["descricao"] = ""

        df["status"] = df["status"].fillna("")
        df["protocolo"] = df["protocolo"].fillna("")
        df["data_inicio_gozo"] = df["data_inicio_gozo"].fillna("")
        df["descricao"] = df["descricao"].fillna("")

        return df

    except Exception:
        return pd.DataFrame(
            columns=[
                "id",
                "empregado",
                "apelido",
                "cliente",
                "admissao",
                "data_inicio_gozo",
                "descricao",
                "status",
                "protocolo"
            ]
        )


def salvar_base(df, arquivo):
    try:
        df.to_excel(arquivo, index=False)
    except Exception:
        logging.exception("Erro ao salvar Excel")
        raise

# ================= EMAIL =================
ULTIMO_EMAIL = 0


def pode_enviar_email():
    global ULTIMO_EMAIL

    agora = time.time()

    if agora - ULTIMO_EMAIL > 300:
        ULTIMO_EMAIL = agora
        return True

    return False


def enviar_email_erro(assunto, mensagem):
    try:
        remetente = os.getenv("EMAIL_REMETENTE")
        senha_email = os.getenv("EMAIL_SENHA")
        destinatario = os.getenv("EMAIL_DESTINO")

        msg = MIMEText(mensagem)
        msg["Subject"] = assunto
        msg["From"] = remetente
        msg["To"] = destinatario

        servidor = smtplib.SMTP("smtp.gmail.com", 587)
        servidor.starttls()
        servidor.login(remetente, senha_email)

        servidor.sendmail(remetente, destinatario, msg.as_string())
        servidor.quit()

        logging.info("📧 Email de erro enviado")

    except Exception:
        logging.exception("Erro ao enviar email")

# ================= SELENIUM =================
def clicar(driver, xpath, tempo=15):
    elemento = WebDriverWait(driver, tempo).until(
        EC.element_to_be_clickable((By.XPATH, xpath))
    )

    time.sleep(1)
    elemento.click()
    time.sleep(1)


def clicar_1(driver, xpath, tempo=15):
    elemento = WebDriverWait(driver, tempo).until(
        EC.element_to_be_clickable((By.XPATH, xpath))
    )

    time.sleep(1)
    elemento.click()
    time.sleep(1)
    elemento.send_keys("1")
    time.sleep(1)
    elemento.send_keys(Keys.ENTER)


def escrever(driver, xpath, texto, tempo=15):
    campo = WebDriverWait(driver, tempo).until(
        EC.presence_of_element_located((By.XPATH, xpath))
    )

    time.sleep(1)
    campo.clear()
    campo.send_keys(texto)
    time.sleep(1)
    campo.send_keys(Keys.ENTER)
    time.sleep(1)


def escrever_sem_enter(driver, xpath, texto, tempo=10):
    campo = WebDriverWait(driver, tempo).until(
        EC.presence_of_element_located((By.XPATH, xpath))
    )

    time.sleep(1)
    campo.clear()
    campo.send_keys(texto)
    time.sleep(1)

# ================= SCRAPING =================
def pegar_dados(driver, grid_id):
    dados = []
    linha = 2

    # Se a grid não carregar, normalmente significa que não há dados neste mês —
    # nesse caso retornamos lista vazia e seguimos o fluxo, sem tratar como erro.
    try:
        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.ID, grid_id))
        )
    except Exception:
        logging.info("📭 Nenhum dado encontrado (grid não carregou) — seguindo sem registros")
        return dados

    while True:
        try:
            base = f'//*[@id="{grid_id}"]/div[1]/div[1]/div[1]/div[{linha}]'

            id_ = driver.find_element(By.XPATH, base + '/div[1]').text
            empregado = driver.find_element(By.XPATH, base + '/div[5]').text
            apelido = driver.find_element(By.XPATH, base + '/div[6]').text
            cliente = driver.find_element(By.XPATH, base + '/div[7]').text
            admissao = driver.find_element(By.XPATH, base + '/div[9]').text

            if not id_:
                break

            dados.append({
                "id": id_,
                "empregado": empregado,
                "apelido": apelido,
                "cliente": cliente,
                "admissao": admissao
            })

            linha += 1

        except:
            logging.info("Todos IDs lidos!")
            break

    return dados

# ================= DATA DE INÍCIO DO GOZO =================
def extrair_data_inicio_gozo(driver):
    """
    Extrai o texto da "Data de início do gozo" do detalhe que está aberto.

    A data não tem um atributo estável (data-qe-id), então ancoramos no rótulo
    "Data de início do gozo" e pegamos o valor da coluna ao lado. Isso é mais
    seguro do que depender de posição ou da classe CSS (knowledge-ultralight).

    Returns:
        str: Texto bruto encontrado (ex.: "08/06/2026") ou "" se não encontrar.
    """
    xpaths = [
        # 1ª opção: pelo rótulo -> valor na coluna irmã (col-sm-7).
        '//span[contains(normalize-space(.), "Data de início do gozo")]'
        '/ancestor::div[contains(@class,"col-sm-2")][1]'
        '/following-sibling::div[1]//span',
        # Reserva: pelo rótulo -> qualquer span seguinte na mesma linha.
        '//span[contains(normalize-space(.), "Data de início do gozo")]'
        '/ancestor::div[contains(@class,"row")][1]'
        '//div[contains(@class,"col-sm-7")]//span',
    ]
    for xp in xpaths:
        try:
            for el in driver.find_elements(By.XPATH, xp):
                texto = el.text.strip()
                if texto:
                    return texto
        except Exception:
            continue
    return ""


def extrair_descricao(driver):
    """
    Extrai o texto da "Descrição" do detalhe que está aberto.

    Ancoramos no rótulo "Descrição" (classe detail-label) e pegamos o valor da
    coluna ao lado (span.detail-data). Ex.: "João: 15/06/2026 a 29/06/2026".

    Returns:
        str: Texto da descrição ou "" se não encontrar.
    """
    xpaths = [
        # 1ª opção: pelo rótulo exato -> valor na coluna irmã.
        '//span[contains(@class,"detail-label") and normalize-space(.)="Descrição"]'
        '/ancestor::div[contains(@class,"col-sm-2")][1]'
        '/following-sibling::div[1]//span',
        # Reserva: pelo rótulo -> span de dados na mesma linha.
        '//span[normalize-space(.)="Descrição"]'
        '/ancestor::div[contains(@class,"row")][1]'
        '//span[contains(@class,"detail-data")]',
    ]
    for xp in xpaths:
        try:
            for el in driver.find_elements(By.XPATH, xp):
                texto = el.text.strip()
                if texto:
                    return texto
        except Exception:
            continue
    return ""


def parse_data(texto):
    """
    Procura uma data no formato dd/mm/aaaa dentro do texto e devolve um date.

    Returns:
        datetime.date ou None se não encontrar/parsear.
    """
    if not texto:
        return None
    m = re.search(r'(\d{2})/(\d{2})/(\d{4})', str(texto))
    if not m:
        return None
    try:
        return datetime.strptime(m.group(0), "%d/%m/%Y").date()
    except ValueError:
        return None


def deve_criar_agora(data_inicio_texto):
    """
    Decide se a tarefa deve ser criada hoje com base na data de início do gozo.

    Regra: cria se o início do gozo estiver a até DIAS_ANTECEDENCIA dias (ou se
    já passou). Datas mais distantes ficam pendentes para as próximas execuções.
    Se a data não puder ser lida, cria mesmo assim (não bloqueia o processo).

    Returns:
        bool: True para criar a tarefa agora, False para adiar.
    """
    data_inicio = parse_data(data_inicio_texto)

    if data_inicio is None:
        logging.warning(
            f"⚠️ Data de início do gozo ilegível ('{data_inicio_texto}') — "
            f"criando tarefa mesmo assim"
        )
        return True

    hoje = datetime.now().date()
    dias_restantes = (data_inicio - hoje).days
    data_fmt = data_inicio.strftime('%d/%m/%Y')

    if dias_restantes <= DIAS_ANTECEDENCIA:
        logging.info(
            f"📅 Início do gozo {data_fmt} (faltam {dias_restantes} dia(s)) — "
            f"dentro do prazo de {DIAS_ANTECEDENCIA} dias, criar tarefa"
        )
        return True

    logging.info(
        f"📅 Início do gozo {data_fmt} (faltam {dias_restantes} dia(s)) — "
        f"fora do prazo de {DIAS_ANTECEDENCIA} dias, adiando criação"
    )
    return False


def buscar_data_inicio_gozo(driver, linha, grid_id):
    """
    Abre o detalhe (modal) do funcionário para extrair a data de início do gozo.

    O detalhe é aberto pelo ícone de info da linha, cujo data-qe-id começa com
    "generic-icon-details-" (o sufixo é um hash dinâmico). Depois fecha o modal
    pelo botão com data-qe-id="header-close-button".

    Args:
        driver: Instância do WebDriver
        linha: Número da linha na tabela
        grid_id: ID do grid

    Returns:
        tuple: (data_inicio_gozo, descricao) — ambos str. Podem vir vazios se
               não forem encontrados no detalhe.
    """
    base = f'//*[@id="{grid_id}"]/div[1]/div[1]/div[1]/div[{linha}]'

    # Ícone de info da linha (sufixo do data-qe-id é dinâmico).
    xpath_icone = base + '//span[starts-with(@data-qe-id, "generic-icon-details-")]'

    clicar(driver, xpath_icone)
    time.sleep(2)

    # Aguarda o modal abrir e extrai a data.
    data_inicio_gozo = extrair_data_inicio_gozo(driver)
    if data_inicio_gozo:
        logging.info(f"📆 Data de início do gozo: {data_inicio_gozo}")
    else:
        logging.warning("⚠️ Não foi possível extrair a data de início do gozo")

    # Extrai a descrição (vai para a observação ao criar a tarefa).
    descricao = extrair_descricao(driver)
    if descricao:
        logging.info(f"📝 Descrição: {descricao}")
    else:
        logging.warning("⚠️ Não foi possível extrair a descrição")

    # Fecha o modal de detalhe (botão estável por data-qe-id).
    try:
        clicar(driver, '//button[@data-qe-id="header-close-button"]')
        time.sleep(1)
    except Exception:
        logging.warning("⚠️ Não foi possível fechar o modal de detalhe")

    return data_inicio_gozo, descricao

# ================= LOGIN =================
def iniciar_driver():
    options = Options()

    options.add_argument("--headless=new")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-gpu")
    options.add_argument("--remote-debugging-port=9222")

    driver = webdriver.Chrome(options=options)

    logging.info("Iniciando Driver")

    return driver


def fazer_login(driver):
    driver.get(url)

    time.sleep(2)

    clicar(driver,'//*[@id="trauth-continue-signout-btn"]')

    time.sleep(2)

    escrever(driver,'//*[@id="username"]', email)
    logging.info('Email validado')

    time.sleep(2)

    escrever(driver,'//*[@id="password"]', senha)
    logging.info('Senha validada')

    time.sleep(2)


def navegar_ate_funcionarios(driver):
    logging.info('Logado')

    time.sleep(1)

    clicar(driver,'//*[@id="bm-header-app-menu-toggle"]')

    time.sleep(1)

    clicar(driver,'//*[@id="bm-header-app-menu"]/ul/li[2]/a/span')

    time.sleep(1)

    driver.switch_to.window(driver.window_handles[-1])

    menu = WebDriverWait(driver, 20).until(
        EC.visibility_of_element_located((By.XPATH, '//*[@id="custom-header"]/div[5]/div/on-nav/nav/div[2]'))
    )

    ActionChains(driver).move_to_element(menu).perform()

    opcao = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.XPATH, '//*[contains(text(), "Férias")]'))
    )

    driver.execute_script("arguments[0].click();", opcao)

# ================= TAREFA =================
def criar_tarefa(driver,empregado,apelido_formatado,titulo_tarefa,main_window,observacao=""):


    time.sleep(2)

    clicar(driver,'//*[@id="custom-header"]/div[1]/button')

    time.sleep(2)

    clicar(driver,'//*[@id="custom-header"]/div[2]/div/div[2]/div[2]/bm-header-app-menu/ul/bm-header-app-menu-item[1]/li/a')

    time.sleep(3)

    driver.switch_to.window(driver.window_handles[-1])

    clicar(driver,'//*[@id="gestta-menu"]/div/div[5]/div/div/div[6]/div/div[1]/div')

    time.sleep(2)

    clicar(driver,'//*[@id="modal-body"]/fieldset/div[1]/div/div/div[1]/span/span[2]/span')

    escrever(
        driver,
        '//*[@id="modal-body"]/fieldset/div[1]/div/div/input[1]',
        titulo_tarefa
    )

    escrever(driver, '//*[@id="complement"]', empregado)

    clicar(driver,'//*[@id="gestta-multiselect-dropdown-11-p"]')

    escrever_sem_enter(
        driver,
        '//*[@id="gestta-multiselect-dropdown-11"]/div/div/ul/li[1]/input',
        apelido_formatado
    )

    clicar(driver,'//*[@id="gestta-multiselect-dropdown-11"]/div/div/ul/li[4]/a')

    clicar(driver,'//*[@id="gestta-multiselect-dropdown-11-p"]')

    clicar(driver,'//*[@id="modal-body"]/fieldset/div[5]/div/span')#clicar em customizar

    # Escreve a descrição no campo de observação (textarea#note), se houver.
    if observacao:
        try:
            escrever_sem_enter(driver, '//*[@id="note"]', observacao)
            logging.info(f"📝 Observação preenchida: {observacao}")
        except Exception:
            logging.warning("⚠️ Não foi possível preencher a observação")

    clicar_1(driver, '//*[@id="modal-body"]/fieldset/div[9]/div/div/span')

    clicar(driver, '//*[@id="modal-body"]/fieldset/div[9]/div/div/div/ul/li[2]/span/button[1]')


    clicar(driver, '//*[@id="modal-body"]/div/button[2]')

    logging.info("⏳ Aguardando geração do número da tarefa...")

    time.sleep(5)

    try:
        xpath_protocolo = '//*[@id="mixed-task-details-view-content"]/div[1]/div[1]/div[2]/div[1]/div[2]/div/div/div[2]/div[1]/label/span[2]'

        elemento_protocolo = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, xpath_protocolo))
        )

        numero_tarefa = elemento_protocolo.text

        logging.info(f"🔢 Tarefa gerada: {numero_tarefa}")

    except Exception:
        logging.warning(
            "⚠️ Não foi possível capturar o número da tarefa, mas ela foi salva."
        )

        numero_tarefa = "Não capturado"

    driver.close()

    driver.switch_to.window(main_window)

    time.sleep(2)

    return numero_tarefa

# ================= PROCESSAMENTO =================
def processar_tipo(
    driver,
    tipo,
    mapa_empresas,
    main_window
):

    logging.info(f"🚀 Processando: {tipo['nome'].upper()}")

    driver.switch_to.window(main_window)

    time.sleep(2)

    driver.refresh()

    time.sleep(4)

    # ================= ABA =================
    if tipo["aba_calculo_ferias"]:

        logging.info("📂 Abrindo aba calculo de ferias")

        WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable(
                (By.XPATH, '//*[@id="ngb-nav-1"]')
            )
        )

        clicar(driver, '//*[@id="ngb-nav-1"]') #clica na aba calculo de ferias

        time.sleep(4)

    else:
        time.sleep(2)

    time.sleep(2)

    # Prepara a grid: espera carregar e expande os itens por página.
    # Se a grid não aparecer, normalmente significa que não há dados neste mês —
    # nesse caso seguimos o fluxo normalmente, sem tratar como erro.
    try:
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.ID, tipo["grid_id"]))
        )

        clicar_1(driver, tipo["xpath_expandir"])
        time.sleep(1)

        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.ID, tipo["grid_id"]))
        )
    except Exception:
        logging.info("📭 Grid não carregou (provavelmente sem dados neste mês) — seguindo sem coletar novos registros")

    # ================= CARREGAR DADOS EXISTENTES =================
    df_total = carregar_base(tipo["arquivo"])
    ids_na_base = set(df_total["id"].astype(str))

    # IDs que estão na base mas ainda sem a data de início do gozo
    ids_sem_data = set(
        df_total[
            (df_total["data_inicio_gozo"] == "") &
            (df_total["status"] != "Concluído")
        ]["id"].astype(str)
    )

    # ================= COLETA DE DADOS DO SITE =================
    logging.info("📊 Coletando dados do site (inline com data de início do gozo)...")
    novos_registros = []
    linha = 2

    try:
        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.ID, tipo["grid_id"]))
        )
    except Exception:
        logging.info("📭 Nenhum dado encontrado (grid não carregou) — seguindo sem registros")

    # Loop para ler todas as linhas da tabela
    while True:
        try:
            base = f'//*[@id="{tipo["grid_id"]}"]/div[1]/div[1]/div[1]/div[{linha}]'
            id_ = driver.find_element(By.XPATH, base + '/div[1]').text

            if not id_:
                break

            empregado = driver.find_element(By.XPATH, base + '/div[5]').text
            apelido   = driver.find_element(By.XPATH, base + '/div[6]').text
            cliente   = driver.find_element(By.XPATH, base + '/div[7]').text
            admissao  = driver.find_element(By.XPATH, base + '/div[9]').text

            logging.info(f"🔎 Lendo linha {linha}: {empregado}")

            # Ignora IDs abaixo do mínimo configurado no tipo
            if int(id_) <= tipo["id_minimo"]:
                linha += 1
                continue

            eh_novo = str(id_) not in ids_na_base
            precisa_data = eh_novo or str(id_) in ids_sem_data

            data_inicio_gozo = ""
            descricao = ""

            # Abre o detalhe para pegar a data de início do gozo e a descrição
            if precisa_data:
                data_inicio_gozo, descricao = buscar_data_inicio_gozo(driver, linha, tipo["grid_id"])

                # Atualiza registro existente com a data e a descrição
                if not eh_novo:
                    mask = df_total["id"].astype(str) == str(id_)
                    df_total.loc[mask, "data_inicio_gozo"] = data_inicio_gozo
                    df_total.loc[mask, "descricao"] = descricao

            # Adiciona novo registro se necessário
            if eh_novo:
                novos_registros.append({
                    "id": id_, "empregado": empregado, "apelido": apelido,
                    "cliente": cliente, "admissao": admissao,
                    "data_inicio_gozo": data_inicio_gozo, "descricao": descricao,
                    "status": "", "protocolo": ""
                })

            linha += 1

        except Exception:
            logging.info("Todos os IDs lidos!")
            break

    # ================= ADICIONAR NOVOS REGISTROS =================
    if novos_registros:
        df_novo = pd.DataFrame(novos_registros)
        df_total = pd.concat([df_total, df_novo], ignore_index=True)
        df_total = df_total.drop_duplicates(subset="id")
        logging.info(f"✅ {len(novos_registros)} novos registros adicionados")
    else:
        logging.info("😴 Nenhum ID novo detectado")

    # Salva a base atualizada
    salvar_base(df_total, tipo["arquivo"])

    # ================= REESTABELECER PÁGINA =================
    logging.info("🔄 Re-estabelecendo página antes de criar tarefas...")
    driver.switch_to.window(main_window)
    driver.refresh()
    time.sleep(4)

    pendentes = df_total[df_total["status"] != "Concluído"]

    if not pendentes.empty:

        logging.info(
            f"📂 Há {len(pendentes)} tarefa(s) pendente(s)"
        )

        for index, registro in pendentes.iterrows():

            apelido_formatado = montar_apelido_com_codigo(
                registro["apelido"],
                mapa_empresas
            )

            try:

                # Regra da data de início do gozo: só cria se estiver dentro do
                # prazo de DIAS_ANTECEDENCIA dias. Caso contrário, deixa pendente
                # para ser reavaliado nas próximas execuções.
                data_gozo_txt = str(registro.get("data_inicio_gozo") or "")
                if not deve_criar_agora(data_gozo_txt):
                    logging.info(f"⏭️ Adiando {registro['empregado']} — início do gozo fora do prazo")
                    continue

                logging.info(
                    f"🛠️ Criando tarefa para {registro['empregado']}"
                )

                numero_tarefa = criar_tarefa(
                    driver,
                    registro["empregado"],
                    apelido_formatado,
                    tipo["titulo_tarefa"],
                    main_window,
                    str(registro.get("descricao") or "")
                )

                df_total.at[index, "status"] = "Concluído"

                df_total.at[index, "protocolo"] = numero_tarefa

                salvar_base(df_total, tipo["arquivo"])

                logging.info(
                    f"✔️ Concluído: {registro['empregado']} | Prot: {numero_tarefa}"
                )

                driver.switch_to.window(main_window)

                time.sleep(2)

            except Exception as e:

                logging.error(
                    f"❌ Erro ao processar {registro['empregado']}"
                )

                raise e

    else:
        logging.info("⭐ Nenhuma tarefa pendente")

# ================= AUTOMAÇÃO =================
def rodar_automacao():

    mapa_empresas = carregar_empresas()

    driver = iniciar_driver()

    try:

        fazer_login(driver)

        navegar_ate_funcionarios(driver)

        time.sleep(2)

        main_window = driver.current_window_handle

        for tipo in TIPOS:

            try:

                processar_tipo(
                    driver,
                    tipo,
                    mapa_empresas,
                    main_window
                )

            except Exception:

                logging.exception(
                    f"Erro ao processar tipo: {tipo['nome']}"
                )

                raise

    finally:

        driver.quit()

        logging.info("Driver finalizado")

# ================= MAIN =================
MAX_TENTATIVAS = 3


def main():

    configurar_log()

    tentativa = 0

    while tentativa < MAX_TENTATIVAS:

        tentativa += 1

        logging.info(f"🚀 Tentativa {tentativa}/{MAX_TENTATIVAS} — Iniciando automação unificada")

        try:

            rodar_automacao()

            logging.info("✅ Automação finalizada com sucesso")

            return

        except Exception:

            logging.exception(f"❌ Falha na tentativa {tentativa}/{MAX_TENTATIVAS}")

            if tentativa < MAX_TENTATIVAS:

                logging.info("🔁 Aguardando 10 segundos antes de nova tentativa...")

                time.sleep(10)

    # Esgotou todas as tentativas
    mensagem = (
        f"O aplicativo de Férias tentou executar {MAX_TENTATIVAS} vezes "
        f"e o erro persistiu em todas as tentativas.\n\n"
        f"Data/Hora: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n\n"
        f"Verifique o arquivo de log para mais detalhes."
    )

    logging.critical(
        f"🚨 {MAX_TENTATIVAS} tentativas esgotadas. Encerrando aplicativo."
    )

    enviar_email_erro(
        f"🚨 FALHA CRÍTICA — Automação Férias encerrada após {MAX_TENTATIVAS} tentativas",
        mensagem
    )

    sys.exit(1)

# ================= RUN =================
if __name__ == "__main__":
    main()
