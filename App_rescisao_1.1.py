# ================= AUTOMAÇÃO DE RESCISÃO =================
# Script para automatizar a criação de tarefas de cálculo de rescisão no sistema Onvio

# ================= IMPORTAÇÕES =================
from datetime import datetime  # Para trabalhar com datas e horários atuais
import time  # Para pausas (time.sleep) - necessário para aguardar carregamento
import os  # Para acessar variáveis de ambiente e manipular caminhos
import sys  # Para verificar se está rodando como .exe e controlar saída
import pandas as pd  # Para manipular DataFrames e arquivos Excel
from dotenv import load_dotenv  # Para carregar credenciais do arquivo .env
import logging  # Para gerar logs (registros) da execução
from selenium import webdriver  # Framework para automatizar navegador Chrome
from selenium.webdriver.common.by import By  # Para localizar elementos
from selenium.webdriver.chrome.options import Options  # Configurações do Chrome
from selenium.webdriver.support.ui import WebDriverWait  # Para esperar elementos
from selenium.webdriver.support import expected_conditions as EC  # Condições de espera
from selenium.webdriver.common.keys import Keys  # Para enviar teclas especiais
from selenium.webdriver.common.action_chains import ActionChains  # Ações de mouse
import smtplib  # Para envio de emails
from email.mime.text import MIMEText  # Formatar emails

# ================= CONFIGURAÇÕES INICIAIS =================
# Carrega variáveis de ambiente (EMAIL, SENHA do arquivo .env)
load_dotenv()

# Obtém credenciais de acesso
email = os.getenv("EMAIL")  # Email para login no Onvio
senha = os.getenv("SENHA")  # Senha para login no Onvio
url = os.getenv("URL_ENTRADA")

# Configuração do tipo de processamento (cálculo de rescisão)
TIPOS = [
{
    "nome": "calculo de rescisao",  # Nome do tipo de processamento
    "arquivo": "dados_calculo_de_rescisao.xlsx",  # Arquivo Excel para armazenar dados
    "grid_id": "gridTerminationCalculationList",  # ID da tabela no site
    "aba_calculo_rescisao": True,  # Indica que deve acessar aba de cálculo de rescisão
    "titulo_tarefa": "",  # Título vazio - será definido dinamicamente baseado no motivo
    "xpath_expandir": '//*[@id="items-per-page-1"]',  # XPATH para expandir itens
    "usar_motivo": True  # Indica que deve buscar o motivo da rescisão
}
]

# ================= CONFIGURAÇÃO DE LOGS =================
def configurar_log():
    """
    Configura o sistema de logging (registros) da aplicação.
    Organiza os logs em pastas por mês/ano.
    Funciona tanto como script Python quanto como executável (.exe).
    """
    
    # Detecta se está rodando como executável ou script
    if getattr(sys, 'frozen', False):
        base_dir = os.path.dirname(sys.executable)  # Diretório do .exe
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))  # Diretório do script

    # Dicionário para converter número do mês em nome em português
    MESES_PT = {
        1: "janeiro", 2: "fevereiro", 3: "março", 4: "abril",
        5: "maio", 6: "junho", 7: "julho", 8: "agosto",
        9: "setembro", 10: "outubro", 11: "novembro", 12: "dezembro"
    }

    # Obtém data e hora atuais
    agora = datetime.now()
    nome_projeto = "log_rescisao"  # Nome base do arquivo de log
    data_hoje = agora.strftime("%Y-%m-%d")  # Data no formato YYYY-MM-DD
    mes_pasta = f"{MESES_PT[agora.month]}_{agora.year}"  # Ex: "janeiro_2025"

    # Cria a estrutura de pastas
    logs_dir = os.path.join(base_dir, "logs", "rescisao", mes_pasta)
    os.makedirs(logs_dir, exist_ok=True)

    # Define o caminho completo do arquivo de log
    arquivo_log = os.path.join(logs_dir, f"{nome_projeto}_{data_hoje}.log")

    # Configura o logging
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[
            logging.FileHandler(arquivo_log, encoding="utf-8"),  # Salva em arquivo
            logging.StreamHandler()  # Exibe no console
        ]
    )

# ================= FUNÇÕES PARA MANIPULAÇÃO DE EXCEL =================
def carregar_empresas():
    """
    Carrega o arquivo Empresas_1.xlsx com mapeamento de empresas.
    
    Returns:
        dict: Dicionário com apelido -> código da empresa
    """
    df = pd.read_excel("Empresas_1.xlsx")
    df["Apelido"] = df["Apelido"].astype(str).str.strip().str.lower()  # Padroniza apelidos
    mapa = dict(zip(df["Apelido"], df["Código"]))
    return mapa

def carregar_motivos():
    """
    Carrega o arquivo motivos_rescisao.xlsx que mapeia motivos de rescisão
    para títulos de tarefas correspondentes.
    
    Returns:
        dict: Dicionário com motivo -> título da tarefa
              Ex: {"demissão sem justa causa": "DP - RESCISÃO SEM JUSTA CAUSA"}
    """
    try:
        # Lê o arquivo Excel com os motivos
        df = pd.read_excel("motivos_rescisao.xlsx")

        # Limpa os nomes das colunas (remove espaços extras)
        df.columns = df.columns.str.strip()

        # Padroniza o campo "Motivo da rescisão" (minúsculas, sem espaços)
        df["Motivo da rescisão"] = (
            df["Motivo da rescisão"]
            .astype(str)
            .str.strip()
            .str.lower()
        )

        # Encontra a coluna que contém "templates" (case insensitive)
        coluna_template = next(
            c for c in df.columns if c.lower() == "templates"
        )

        # Padroniza os templates
        df[coluna_template] = df[coluna_template].astype(str).str.strip()

        # Cria dicionário: motivo -> template
        mapa = dict(zip(df["Motivo da rescisão"], df[coluna_template]))

        logging.info("✅ Motivos de rescisão carregados")
        return mapa

    except Exception:
        logging.exception("Erro ao carregar motivos_rescisao.xlsx")
        return {}  # Retorna dicionário vazio em caso de erro


def montar_apelido_com_codigo(apelido, mapa):
    """
    Monta apelido formatado com código da empresa.
    Exemplo: "empresa x" + código "123" -> "123 - em"
    """
    apelido_limpo = str(apelido).strip().lower()
    codigo = mapa.get(apelido_limpo)

    if codigo:
        apelido_curto = apelido_limpo[:2]  # Primeiras 2 letras
        return f"{codigo} - {apelido_curto}"
    else:
        logging.warning(f"Apelido não encontrado: {apelido}")
        return apelido


def carregar_base(arquivo):
    """
    Carrega o arquivo Excel com os dados dos funcionários.
    Cria estrutura padrão se o arquivo não existir.
    
    Colunas:
        - id: ID do funcionário
        - empregado: Nome do funcionário
        - apelido: Apelido da empresa
        - cliente: Nome do cliente
        - admissao: Data de admissão
        - motivo: Motivo da rescisão
        - titulo_tarefa: Título da tarefa a ser criada
        - status: Status do processamento
        - protocolo: Número do protocolo gerado
    """
    try:
        df = pd.read_excel(arquivo)

        # Adiciona colunas padrão se não existirem
        if "status" not in df.columns:
            df["status"] = ""

        if "protocolo" not in df.columns:
            df["protocolo"] = ""

        if "motivo" not in df.columns:
            df["motivo"] = ""

        if "titulo_tarefa" not in df.columns:
            df["titulo_tarefa"] = ""

        # Substitui valores vazios por string vazia
        df["status"] = df["status"].fillna("")
        df["protocolo"] = df["protocolo"].fillna("")
        df["motivo"] = df["motivo"].fillna("")
        df["titulo_tarefa"] = df["titulo_tarefa"].fillna("")

        return df

    except Exception:
        # Retorna DataFrame vazio com estrutura definida
        return pd.DataFrame(
            columns=[
                "id", "empregado", "apelido", "cliente", "admissao",
                "motivo", "titulo_tarefa", "status", "protocolo"
            ]
        )


def salvar_base(df, arquivo):
    """
    Salva o DataFrame em arquivo Excel.
    """
    try:
        df.to_excel(arquivo, index=False)
    except Exception:
        logging.exception("Erro ao salvar Excel")
        raise

# ================= FUNÇÕES DE ENVIO DE EMAIL =================
ULTIMO_EMAIL = 0  # Controla último envio de email (evita spam)

def pode_enviar_email():
    """
    Verifica se pode enviar email (máximo 1 a cada 5 minutos).
    """
    global ULTIMO_EMAIL
    agora = time.time()
    if agora - ULTIMO_EMAIL > 300:  # 300 segundos = 5 minutos
        ULTIMO_EMAIL = agora
        return True
    return False

def enviar_email_erro(assunto, mensagem):
    """
    Envia email de erro via SMTP do Gmail.
    """
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

# ================= FUNÇÕES DE INTERAÇÃO COM SELENIUM =================
def clicar(driver, xpath, tempo=15):
    """
    Clica em um elemento identificado por XPATH.
    Aguarda até o elemento estar clicável.
    """
    elemento = WebDriverWait(driver, tempo).until(
        EC.element_to_be_clickable((By.XPATH, xpath))
    )
    time.sleep(1)
    elemento.click()
    time.sleep(1)


def clicar_1(driver, xpath, tempo=15):
    """
    Clica em um elemento, digita '100' e pressiona ENTER.
    Usado para selecionar quantidade "100 itens por página".
    """
    elemento = WebDriverWait(driver, tempo).until(
        EC.element_to_be_clickable((By.XPATH, xpath))
    )
    time.sleep(1)
    elemento.click()
    time.sleep(1)
    elemento.send_keys("100")  # Digita 100 (quantidade de itens por página)
    time.sleep(1)
    elemento.send_keys(Keys.ENTER)  # Confirma


def escrever(driver, xpath, texto, tempo=15):
    """
    Escreve texto em um campo e pressiona ENTER.
    """
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
    """
    Escreve texto em um campo sem pressionar ENTER.
    Usado em campos de autocomplete.
    """
    campo = WebDriverWait(driver, tempo).until(
        EC.presence_of_element_located((By.XPATH, xpath))
    )
    time.sleep(1)
    campo.clear()
    campo.send_keys(texto)
    time.sleep(1)

# ================= FUNÇÕES DE SCRAPING =================
def pegar_dados(driver, grid_id):
    """
    Extrai dados da tabela (grid) do site.
    Retorna lista de dicionários com dados dos funcionários.
    """
    dados = []
    linha = 2  # Começa na linha 2 (linha 1 é cabeçalho)

    WebDriverWait(driver, 15).until(
        EC.presence_of_element_located((By.ID, grid_id))
    )

    while True:
        try:
            # Constrói XPATH para a linha atual
            base = f'//*[@id="{grid_id}"]/div[1]/div[1]/div[1]/div[{linha}]'

            id_ = driver.find_element(By.XPATH, base + '/div[1]').text
            empregado = driver.find_element(By.XPATH, base + '/div[5]').text
            apelido = driver.find_element(By.XPATH, base + '/div[6]').text
            cliente = driver.find_element(By.XPATH, base + '/div[7]').text
            admissao = driver.find_element(By.XPATH, base + '/div[9]').text

            print(empregado)  # Mostra no console (debug)

            if not id_:
                break

            dados.append({
                "id": id_,
                "empregado": empregado,
                "apelido": apelido,
                "cliente": cliente,
                "admissao": admissao,
                "linha": linha  # Guarda número da linha para referência
            })

            linha += 1

        except Exception:
            logging.info("Todos IDs lidos!")
            break

    return dados

# ================= FUNÇÃO PARA BUSCAR MOTIVO DA RESCISÃO =================
def buscar_motivo_rescisao(driver, linha):
    """
    Abre o detalhe do funcionário para extrair o motivo da rescisão.
    
    Args:
        driver: Instância do WebDriver
        linha: Número da linha na tabela
    
    Returns:
        str: Motivo da rescisão encontrado
    """
    # Salva as janelas abertas antes de clicar
    janelas_antes = set(driver.window_handles)

    # Constrói XPATH do botão de detalhe (ícone de olho ou similar)
    xpath_click = (
        f'//*[@id="gridTerminationCalculationList"]'
        f'/div[1]/div[1]/div[1]/div[{linha}]/div[2]/div/span[2]'
    )

    clicar(driver, xpath_click)  # Clica para abrir detalhes

    time.sleep(2)

    # Verifica se abriu uma nova janela
    janelas_depois = set(driver.window_handles)
    nova_janela = janelas_depois - janelas_antes

    if nova_janela:
        # Se abriu nova janela, muda o foco para ela
        driver.switch_to.window(nova_janela.pop())
        logging.info("🪟 Detalhe abriu em nova janela — switch realizado")

    # Aguarda e localiza o elemento que contém o motivo da rescisão
    elemento = WebDriverWait(driver, 15).until(
        EC.presence_of_element_located(
            (By.CSS_SELECTOR, '[data-qe-id="terminate-motive"]')
        )
    )

    motivo = elemento.text.strip()  # Extrai o texto do motivo
    logging.info(f"📋 Motivo da rescisão: {motivo}")

    # Fecha o modal de detalhes (clica no botão "Voltar" ou "Fechar")
    clicar(
        driver,
        '//*[@id="ngb-nav-1-panel"]/app-termination-calculation/div[2]/app-generic-detail/div/div/div/div[1]/div[2]/button'
    )

    time.sleep(1)

    # Se abriu nova janela, fecha e volta para a original
    if nova_janela:
        driver.close()  # Fecha a janela de detalhe
        driver.switch_to.window(list(janelas_antes)[0])  # Volta para janela principal

    return motivo

# ================= FUNÇÕES DE LOGIN E NAVEGAÇÃO =================
def iniciar_driver():
    """
    Inicializa o Chrome WebDriver em modo headless.
    """
    options = Options()
    options.add_argument("--headless=new")  # Modo sem interface gráfica
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-gpu")
    options.add_argument("--remote-debugging-port=9222")

    driver = webdriver.Chrome(options=options)
    logging.info("Iniciando Driver")
    return driver


def fazer_login(driver):
    """
    Realiza login no sistema Onvio.
    """
    driver.get(url) # Logout primeiro
    time.sleep(2)
    clicar(driver, '//*[@id="trauth-continue-signout-btn"]')
    time.sleep(2)
    escrever(driver, '//*[@id="username"]', email)
    logging.info('Email validado')
    time.sleep(2)
    escrever(driver, '//*[@id="password"]', senha)
    logging.info('Senha validada')
    time.sleep(2)


def navegar_ate_funcionarios(driver):
    """
    Navega pelos menus até chegar na página de Rescisão.
    """
    logging.info('Logado')
    time.sleep(1)
    clicar(driver, '//*[@id="bm-header-app-menu-toggle"]')  # Abre menu
    time.sleep(1)
    clicar(driver, '//*[@id="bm-header-app-menu"]/ul/li[2]/a/span')  # Funcionários
    time.sleep(1)
    driver.switch_to.window(driver.window_handles[-1])  # Muda para nova aba

    # Menu de navegação superior (hover)
    menu = WebDriverWait(driver, 20).until(
        EC.visibility_of_element_located((By.XPATH, '//*[@id="custom-header"]/div[5]/div/on-nav/nav/div[2]'))
    )
    ActionChains(driver).move_to_element(menu).perform()  # Hover

    # Clica na opção "Rescisão"
    opcao = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.XPATH, '//*[contains(text(), "Rescisão")]'))
    )
    driver.execute_script("arguments[0].click();", opcao)  # Clique via JavaScript

# ================= FUNÇÃO DE CRIAÇÃO DE TAREFAS =================
def criar_tarefa(driver, empregado, apelido_formatado, titulo_tarefa, main_window):
    """
    Cria uma tarefa de rescisão no sistema para um funcionário.
    
    Args:
        driver: Instância do WebDriver
        empregado: Nome do funcionário
        apelido_formatado: Apelido da empresa formatado com código
        titulo_tarefa: Título da tarefa (baseado no motivo da rescisão)
        main_window: Handle da janela principal
    
    Returns:
        str: Número do protocolo da tarefa criada
    """
    # Garante que está na janela principal
    driver.switch_to.window(main_window)
    time.sleep(2)

    # Abre menu de tarefas
    clicar(driver, '//*[@id="custom-header"]/div[1]/button')
    time.sleep(2)

    # Seleciona "Nova Tarefa"
    clicar(driver, '//*[@id="custom-header"]/div[2]/div/div[2]/div[2]/bm-header-app-menu/ul/bm-header-app-menu-item[1]/li/a')
    time.sleep(3)

    # Muda para nova aba de criação de tarefa
    driver.switch_to.window(driver.window_handles[-1])

    # Seleciona tipo de tarefa
    clicar(driver, '//*[@id="gestta-menu"]/div/div[5]/div/div/div[6]/div/div[1]/div')
    time.sleep(2)

    # Abre campo de título
    clicar(driver, '//*[@id="modal-body"]/fieldset/div[1]/div/div/div[1]/span/span[2]/span')
    
    # Escreve título da tarefa
    escrever(driver, '//*[@id="modal-body"]/fieldset/div[1]/div/div/input[1]', titulo_tarefa)

    # Escreve nome do empregado
    escrever(driver, '//*[@id="complement"]', empregado)

    # Seleciona empresa
    clicar(driver, '//*[@id="gestta-multiselect-dropdown-11-p"]')
    escrever_sem_enter(
        driver,
        '//*[@id="gestta-multiselect-dropdown-11"]/div/div/ul/li[1]/input',
        apelido_formatado
    )
    clicar(driver, '//*[@id="gestta-multiselect-dropdown-11"]/div/div/ul/li[4]/a')
    clicar(driver, '//*[@id="gestta-multiselect-dropdown-11-p"]')

    # Confirma seleção do funcionário
    clicar(driver, '//*[@id="modal-body"]/fieldset/div[5]/div/span')

    # Define quantidade (100 itens)
    clicar_1(driver, '//*[@id="modal-body"]/fieldset/div[9]/div/div/span')
    clicar(driver, '//*[@id="modal-body"]/fieldset/div[9]/div/div/div/ul/li[2]/span/button[1]')

    # Salva a tarefa
    clicar(driver, '//*[@id="modal-body"]/div/button[2]')

    logging.info("⏳ Aguardando fechamento do modal...")

    # Aguarda o modal fechar (até 30 segundos)
    try:
        WebDriverWait(driver, 30).until(
            EC.invisibility_of_element_located((By.XPATH, '//*[@id="modal-body"]'))
        )
        logging.info("✅ Modal fechado — aguardando número da tarefa...")
    except Exception:
        logging.warning("⚠️ Modal não fechou no tempo esperado, tentando capturar protocolo mesmo assim...")

    # Captura o número da tarefa gerada
    try:
        xpath_protocolo = '//*[@id="mixed-task-details-view-content"]/div[1]/div[1]/div[2]/div[1]/div[2]/div/div/div[2]/div[1]/label/span[2]'
        elemento_protocolo = WebDriverWait(driver, 40).until(
            EC.presence_of_element_located((By.XPATH, xpath_protocolo))
        )
        numero_tarefa = elemento_protocolo.text
        logging.info(f"🔢 Tarefa gerada: {numero_tarefa}")
    except Exception:
        logging.warning(f"⚠️ Não foi possível capturar o número da tarefa. URL atual: {driver.current_url}")
        numero_tarefa = "Não capturado"

    # Fecha aba da tarefa e volta para principal
    driver.close()
    driver.switch_to.window(main_window)
    time.sleep(2)

    return numero_tarefa

# ================= FUNÇÃO DE PROCESSAMENTO PRINCIPAL =================
def processar_tipo(driver, tipo, mapa_empresas, main_window, mapa_motivos=None):
    """
    Processa o tipo de rescisão:
    1. Acessa aba de cálculo de rescisão
    2. Coleta dados do site (incluindo motivos)
    3. Compara com base existente
    4. Cria tarefas com títulos baseados nos motivos
    5. Atualiza status no Excel
    
    Args:
        driver: Instância do WebDriver
        tipo: Configurações do tipo de processamento
        mapa_empresas: Mapeamento de apelidos para códigos
        main_window: Handle da janela principal
        mapa_motivos: Mapeamento de motivos para títulos de tarefa
    """
    logging.info(f"🚀 Processando: {tipo['nome'].upper()}")

    # Volta para janela principal e recarrega
    driver.switch_to.window(main_window)
    time.sleep(2)
    driver.refresh()
    time.sleep(4)

    # ================= ABRIR ABA CORRETA =================
    if tipo["aba_calculo_rescisao"]:
        logging.info("📂 Abrindo aba calculo de rescisao")
        WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="ngb-nav-1"]'))
        )
        clicar(driver, '//*[@id="ngb-nav-1"]')  # Clica na aba de cálculo de rescisão
        time.sleep(2)
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.ID, tipo["grid_id"]))
        )

    time.sleep(2)
    clicar_1(driver, tipo["xpath_expandir"])  # Expande para 100 itens por página
    time.sleep(1)
    WebDriverWait(driver, 15).until(
        EC.presence_of_element_located((By.ID, tipo["grid_id"]))
    )

    # ================= CARREGAR DADOS EXISTENTES =================
    df_total = carregar_base(tipo["arquivo"])
    ids_na_base = set(df_total["id"].astype(str))
    
    # IDs que estão na base mas sem título definido
    ids_sem_titulo = set(
        df_total[
            (df_total["titulo_tarefa"] == "") &
            (df_total["status"] != "Concluído")
        ]["id"].astype(str)
    )

    # ================= COLETA DE DADOS DO SITE =================
    logging.info("📊 Coletando dados do site (inline com motivo)...")
    novos_registros = []
    linha = 2

    WebDriverWait(driver, 15).until(
        EC.presence_of_element_located((By.ID, tipo["grid_id"]))
    )

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

            # Verifica se é novo ou precisa de motivo
            eh_novo = str(id_) not in ids_na_base
            precisa_motivo = eh_novo or str(id_) in ids_sem_titulo

            motivo = ""
            titulo_tarefa = ""

            # Busca o motivo se necessário
            if precisa_motivo and tipo.get("usar_motivo") and mapa_motivos:
                motivo = buscar_motivo_rescisao(driver, linha)
                titulo_tarefa = mapa_motivos.get(motivo.lower(), "DP - RESCISAO")
                
                if mapa_motivos.get(motivo.lower()) is None:
                    logging.warning(f"⚠️ Motivo não encontrado no mapeamento: '{motivo}' — usando título padrão: DP - RESCISAO")

                # Atualiza registro existente com motivo e título
                if not eh_novo:
                    mask = df_total["id"].astype(str) == str(id_)
                    df_total.loc[mask, "motivo"] = motivo
                    df_total.loc[mask, "titulo_tarefa"] = titulo_tarefa

            # Adiciona novo registro se necessário
            if eh_novo:
                novos_registros.append({
                    "id": id_, "empregado": empregado, "apelido": apelido,
                    "cliente": cliente, "admissao": admissao, "motivo": motivo,
                    "titulo_tarefa": titulo_tarefa, "status": "", "protocolo": ""
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

    # ================= PROCESSAR TAREFAS PENDENTES =================
    pendentes = df_total[df_total["status"] != "Concluído"]

    if not pendentes.empty:
        logging.info(f"📂 Há {len(pendentes)} tarefa(s) pendente(s)")

        for index, registro in pendentes.iterrows():
            apelido_formatado = montar_apelido_com_codigo(registro["apelido"], mapa_empresas)

            try:
                titulo = str(registro.get("titulo_tarefa") or "").strip()

                if not titulo:
                    logging.warning(f"⚠️ Título vazio para {registro['empregado']}, pulando...")
                    continue

                logging.info(f"🛠️ Criando tarefa para {registro['empregado']} | Título: {titulo}")

                # Cria a tarefa
                numero_tarefa = criar_tarefa(
                    driver, registro["empregado"], apelido_formatado, titulo, main_window
                )

                # Atualiza status e protocolo
                df_total.at[index, "status"] = "Concluído"
                df_total.at[index, "protocolo"] = numero_tarefa

                salvar_base(df_total, tipo["arquivo"])

                logging.info(f"✔️ Concluído: {registro['empregado']} | Prot: {numero_tarefa}")

                driver.switch_to.window(main_window)
                time.sleep(2)

            except Exception as e:
                logging.error(f"❌ Erro ao processar {registro['empregado']}")
                raise e

    else:
        logging.info("⭐ Nenhuma tarefa pendente")

# ================= AUTOMAÇÃO PRINCIPAL =================
def rodar_automacao():
    """
    Função principal que orquestra toda a automação de rescisão.
    """
    # Carrega mapeamentos
    mapa_empresas = carregar_empresas()  # Mapeamento de empresas
    mapa_motivos = carregar_motivos()    # Mapeamento de motivos de rescisão

    # Inicia o navegador
    driver = iniciar_driver()

    try:
        # Login e navegação
        fazer_login(driver)
        navegar_ate_funcionarios(driver)
        time.sleep(2)
        main_window = driver.current_window_handle  # Guarda janela principal

        # Processa cada tipo de rescisão
        for tipo in TIPOS:
            try:
                processar_tipo(driver, tipo, mapa_empresas, main_window, mapa_motivos)
            except Exception:
                logging.exception(f"Erro ao processar tipo: {tipo['nome']}")
                raise

    finally:
        # Garante que o driver seja fechado
        driver.quit()
        logging.info("Driver finalizado")

# ================= FUNÇÃO PRINCIPAL COM TENTATIVAS =================
MAX_TENTATIVAS = 3  # Número máximo de tentativas

def main():
    """
    Função principal com sistema de tentativas.
    Tenta executar até MAX_TENTATIVAS vezes em caso de falha.
    Se todas falharem, envia email de alerta.
    """
    configurar_log()

    tentativa = 0

    while tentativa < MAX_TENTATIVAS:
        tentativa += 1
        logging.info(f"🚀 Tentativa {tentativa}/{MAX_TENTATIVAS} — Iniciando automação unificada")

        try:
            rodar_automacao()
            logging.info("✅ Automação finalizada com sucesso")
            return  # Sai se tudo deu certo

        except Exception:
            logging.exception(f"❌ Falha na tentativa {tentativa}/{MAX_TENTATIVAS}")

            if tentativa < MAX_TENTATIVAS:
                logging.info("🔁 Aguardando 10 segundos antes de nova tentativa...")
                time.sleep(10)

    # ================= ESGOTOU AS