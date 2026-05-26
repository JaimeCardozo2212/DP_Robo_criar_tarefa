# ================= IMPORTAÇÕES =================
from datetime import datetime  # Para trabalhar com datas e horas
import time  # Para pausas (time.sleep)
import os  # Para acessar variáveis de ambiente e manipular arquivos/caminhos
import sys  # Para verificar se está rodando como .exe e sair do programa
import pandas as pd  # Para manipular DataFrames e arquivos Excel
from dotenv import load_dotenv  # Para carregar variáveis de ambiente do arquivo .env
import logging  # Para gerar logs (registros) da execução
from selenium import webdriver  # Para automatizar o navegador Chrome
from selenium.webdriver.common.by import By  # Para localizar elementos por XPATH, ID, etc.
from selenium.webdriver.chrome.options import Options  # Para configurar opções do Chrome
from selenium.webdriver.support.ui import WebDriverWait  # Para esperar elementos carregarem
from selenium.webdriver.support import expected_conditions as EC  # Condições de espera
from selenium.webdriver.common.keys import Keys  # Para enviar teclas especiais (Enter, Tab)
from selenium.webdriver.common.action_chains import ActionChains  # Para ações como mouse hover
import smtplib  # Para enviar emails (SMTP)
from email.mime.text import MIMEText  # Para formatar emails em texto

# ================= CONFIGURAÇÕES INICIAIS =================
# Carrega variáveis de ambiente do arquivo .env (EMAIL, SENHA, etc.)
load_dotenv()

# Obtém email e senha das variáveis de ambiente
email = os.getenv("EMAIL")
senha = os.getenv("SENHA")
url = os.getenv("URL_ENTRADA")

# Configuração dos tipos de admissão (empregados e estagiários)
TIPOS = [
    {
        "nome": "empregado",  # Nome do tipo
        "arquivo": "dados_empregados.xlsx",  # Arquivo Excel para salvar dados
        "grid_id": "gridPayrollRegistrationEmployeeList",  # ID da tabela no site
        "aba_estagiario": False,  # Não é estagiário
        "titulo_tarefa": "DP - ADMISSÃO EMPREGADO(A)",  # Título da tarefa no sistema
        "xpath_expandir": '//*[@id="items-per-page-0"]',  # XPATH para expandir lista
        "id_minimo": 2310
    },
    {
        "nome": "estagiario",
        "arquivo": "dados_estagiarios.xlsx",
        "grid_id": "gridPayrollEntriesInternList",
        "aba_estagiario": True,  # É estagiário
        "titulo_tarefa": "DP - ADMISSÃO EMPREGADO(A)",
        "xpath_expandir": '//*[@id="items-per-page-1"]',
        "id_minimo": 2240
    }
]

# ================= CONFIGURAÇÃO DE LOGS =================
def configurar_log():
    """
    Configura o sistema de logging (registros) da aplicação.
    Cria pastas organizadas por mês/ano e arquivos de log por dia.
    """
    
    # Verifica se o programa está rodando como executável (.exe) ou como script Python
    if getattr(sys, 'frozen', False):
        # Se for .exe, usa o diretório do executável
        base_dir = os.path.dirname(sys.executable)
    else:
        # Se for script Python, usa o diretório do script
        base_dir = os.path.dirname(os.path.abspath(__file__))

    # Dicionário para converter número do mês em nome em português
    MESES_PT = {
        1: "janeiro", 2: "fevereiro", 3: "março", 4: "abril",
        5: "maio", 6: "junho", 7: "julho", 8: "agosto",
        9: "setembro", 10: "outubro", 11: "novembro", 12: "dezembro"
    }

    # Obtém data e hora atual
    agora = datetime.now()
    nome_projeto = "log_admissao_unificada"  # Nome base do arquivo de log
    data_hoje = agora.strftime("%Y-%m-%d")  # Data no formato YYYY-MM-DD
    mes_pasta = f"{MESES_PT[agora.month]}_{agora.year}"  # Ex: "janeiro_2025"

    # Cria o caminho completo para a pasta de logs
    logs_dir = os.path.join(base_dir, "logs", "admissao", mes_pasta)
    os.makedirs(logs_dir, exist_ok=True)  # Cria a pasta se não existir

    # Define o caminho completo do arquivo de log
    arquivo_log = os.path.join(
        logs_dir,
        f"{nome_projeto}_{data_hoje}.log"
    )

    # Configura o logging
    logging.basicConfig(
        level=logging.INFO,  # Nível INFO (mostra informações, avisos e erros)
        format="%(asctime)s [%(levelname)s] %(message)s",  # Formato: data/hora [nível] mensagem
        handlers=[
            logging.FileHandler(arquivo_log, encoding="utf-8"),  # Salva em arquivo
            logging.StreamHandler()  # Também exibe no console
        ]
    )

# ================= FUNÇÕES PARA ARQUIVOS EXCEL =================
def carregar_empresas():
    """
    Carrega o arquivo Empresas_1.xlsx que contém o mapeamento
    entre apelidos das empresas e seus códigos.
    Retorna um dicionário onde a chave é o apelido e o valor é o código.
    """
    try:
        # Lê o arquivo Excel
        df = pd.read_excel("Empresas_1.xlsx")

        # Limpa os dados: converte para string, remove espaços e converte para minúsculas
        df["Apelido"] = df["Apelido"].astype(str).str.strip().str.lower()

        # Cria um dicionário: chave = Apelido, valor = Código
        mapa = dict(zip(df["Apelido"], df["Código"]))

        return mapa
    
    except Exception:
        logging.exception("Erro ao carregar Empresas_1.xlsx")
        return {}  # Retorna dicionário vazio em caso de erro

def montar_apelido_com_codigo(apelido, mapa):
    """
    Monta um apelido formatado com código da empresa.
    Exemplo: "123 - em" (código + primeiras letras do apelido)
    """
    apelido_limpo = str(apelido).strip().lower()  # Limpa o apelido

    codigo = mapa.get(apelido_limpo)  # Busca o código no mapa

    if codigo:
        apelido_curto = apelido_limpo[:2]  # Pega as 2 primeiras letras
        return f"{codigo} - {apelido_curto}"  # Retorna formato com código
    else:
        logging.warning(f"Apelido não encontrado: {apelido}")
        return apelido  # Retorna apelido original se não encontrar

def carregar_base(arquivo):
    """
    Carrega o arquivo Excel que armazena os dados dos funcionários.
    Se o arquivo não existir ou estiver vazio, cria um DataFrame vazio com as colunas necessárias.
    """
    try:
        df = pd.read_excel(arquivo)

        # Adiciona colunas se não existirem
        if "status" not in df.columns:
            df["status"] = ""

        if "protocolo" not in df.columns:
            df["protocolo"] = ""

        # Preenche valores vazios (NaN) com string vazia
        df["status"] = df["status"].fillna("")
        df["protocolo"] = df["protocolo"].fillna("")

        return df

    except Exception:
        # Se der erro, retorna DataFrame vazio com a estrutura correta
        return pd.DataFrame(
            columns=[
                "id", "empregado", "apelido", "cliente",
                "admissao", "status", "protocolo"
            ]
        )

def salvar_base(df, arquivo):
    """
    Salva o DataFrame em um arquivo Excel.
    """
    try:
        df.to_excel(arquivo, index=False)  # index=False para não salvar o índice
    except Exception:
        logging.exception("Erro ao salvar Excel")
        raise  # Propaga a exceção

# ================= FUNÇÕES DE EMAIL =================
ULTIMO_EMAIL = 0  # Variável global para controlar o último envio de email

def pode_enviar_email():
    """
    Controla o limite de envio de emails (máximo 1 a cada 5 minutos / 300 segundos).
    Previne spam de emails em caso de erros repetidos.
    """
    global ULTIMO_EMAIL

    agora = time.time()

    if agora - ULTIMO_EMAIL > 300:  # Se passaram mais de 5 minutos
        ULTIMO_EMAIL = agora
        return True

    return False

def enviar_email_erro(assunto, mensagem):
    """
    Envia email de erro para o destinatário configurado.
    Usa SMTP do Gmail.
    """
    try:
        # Obtém configurações do email das variáveis de ambiente
        remetente = os.getenv("EMAIL_REMETENTE")
        senha_email = os.getenv("EMAIL_SENHA")
        destinatario = os.getenv("EMAIL_DESTINO")

        # Cria a mensagem do email
        msg = MIMEText(mensagem)
        msg["Subject"] = assunto
        msg["From"] = remetente
        msg["To"] = destinatario

        # Conecta ao servidor SMTP do Gmail
        servidor = smtplib.SMTP("smtp.gmail.com", 587)
        servidor.starttls()  # Inicia conexão segura
        servidor.login(remetente, senha_email)  # Faz login

        # Envia o email
        servidor.sendmail(remetente, destinatario, msg.as_string())
        servidor.quit()

        logging.info("📧 Email de erro enviado")

    except Exception:
        logging.exception("Erro ao enviar email")

# ================= FUNÇÕES DE INTERAÇÃO COM SELENIUM =================
def clicar(driver, xpath, tempo=15):
    """
    Clica em um elemento da página identificado por XPATH.
    Aguarda até o elemento estar clicável.
    """
    # Espera o elemento ficar clicável
    elemento = WebDriverWait(driver, tempo).until(
        EC.element_to_be_clickable((By.XPATH, xpath))
    )

    time.sleep(1)  # Pequena pausa para estabilidade
    elemento.click()
    time.sleep(1)

def clicar_1(driver, xpath, tempo=15):
    """
    Clica em um elemento e depois digita '1' e ENTER.
    Usado para selecionar a opção "1" em campos de quantidade.
    """
    # Espera o elemento ficar clicável
    elemento = WebDriverWait(driver, tempo).until(
        EC.element_to_be_clickable((By.XPATH, xpath))
    )

    time.sleep(1)
    elemento.click()  # Clica
    time.sleep(1)
    elemento.send_keys("1")  # Digita 1
    time.sleep(1)
    elemento.send_keys(Keys.ENTER)  # Pressiona Enter

def escrever(driver, xpath, texto, tempo=15):
    """
    Escreve texto em um campo de input e pressiona ENTER.
    """
    # Aguarda o campo estar presente
    campo = WebDriverWait(driver, tempo).until(
        EC.presence_of_element_located((By.XPATH, xpath))
    )

    time.sleep(1)
    campo.clear()  # Limpa o campo
    campo.send_keys(texto)  # Digita o texto
    time.sleep(1)
    campo.send_keys(Keys.ENTER)  # Pressiona Enter
    time.sleep(1)

def escrever_sem_enter(driver, xpath, texto, tempo=10):
    """
    Escreve texto em um campo de input sem pressionar ENTER.
    Usado em campos de autocomplete/pesquisa.
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
    Percorre linha por linha até não encontrar mais dados.
    Retorna uma lista de dicionários com os dados de cada funcionário.
    """
    dados = []
    linha = 2  # Começa na linha 2 (linha 1 é o cabeçalho)

    # Aguarda a tabela carregar
    WebDriverWait(driver, 15).until(
        EC.presence_of_element_located((By.ID, grid_id))
    )

    while True:
        try:
            # Constrói o XPATH para a linha atual
            base = f'//*[@id="{grid_id}"]/div[1]/div[1]/div[1]/div[{linha}]'

            # Extrai cada campo da linha
            id_ = driver.find_element(By.XPATH, base + '/div[1]').text
            empregado = driver.find_element(By.XPATH, base + '/div[5]').text
            apelido = driver.find_element(By.XPATH, base + '/div[6]').text
            cliente = driver.find_element(By.XPATH, base + '/div[7]').text
            admissao = driver.find_element(By.XPATH, base + '/div[9]').text

            # Se ID estiver vazio, termina a leitura
            if not id_:
                break

            # Adiciona os dados à lista
            dados.append({
                "id": id_,
                "empregado": empregado,
                "apelido": apelido,
                "cliente": cliente,
                "admissao": admissao
            })

            linha += 1  # Próxima linha

        except Exception:
            logging.info("Todos IDs lidos!")
            break

    return dados

# ================= FUNÇÕES DE NAVEGAÇÃO =================
def iniciar_driver():
    """
    Inicializa o driver do Chrome em modo headless (sem interface gráfica).
    Configura opções para melhor desempenho e estabilidade.
    """
    options = Options()

    # Configurações para modo headless (roda sem abrir janela do navegador)
    options.add_argument("--headless=new")  # Modo headless moderno
    options.add_argument("--window-size=1920,1080")  # Tamanho da janela virtual
    options.add_argument("--disable-dev-shm-usage")  # Para sistemas com pouco espaço
    options.add_argument("--no-sandbox")  # Necessário em alguns ambientes
    options.add_argument("--disable-gpu")  # Desabilita GPU (melhora performance)
    options.add_argument("--remote-debugging-port=9222")  # Para debug remoto

    # Cria o driver
    driver = webdriver.Chrome(options=options)

    logging.info("Iniciando Driver")
    return driver

def fazer_login(driver):
    """
    Realiza login no sistema Onvio usando email e senha do arquivo .env
    """
    # Acessa página de logout primeiro para garantir estado limpo
    driver.get(url)

    time.sleep(2)

    # Confirma logout se necessário
    clicar(driver, '//*[@id="trauth-continue-signout-btn"]')

    time.sleep(2)

    # Digita email
    escrever(driver, '//*[@id="username"]', email)
    logging.info('Email validado')

    time.sleep(2)

    # Digita senha
    escrever(driver, '//*[@id="password"]', senha)
    logging.info('Senha validada')

    time.sleep(2)

def navegar_ate_funcionarios(driver):
    """
    Navega pelo menu até chegar na página de Cadastro de Funcionários.
    """
    logging.info('Logado')

    time.sleep(1)

    # Abre o menu principal
    clicar(driver, '//*[@id="bm-header-app-menu-toggle"]')

    time.sleep(1)

    # Clica na opção "Funcionários"
    clicar(driver, '//*[@id="bm-header-app-menu"]/ul/li[2]/a/span')

    time.sleep(1)

    # Muda para a nova aba que foi aberta
    driver.switch_to.window(driver.window_handles[-1])

    # Localiza o menu de navegação (hover)
    menu = WebDriverWait(driver, 20).until(
        EC.visibility_of_element_located((By.XPATH, '//*[@id="custom-header"]/div[5]/div/on-nav/nav/div[2]'))
    )

    # Move o mouse sobre o menu (hover)
    ActionChains(driver).move_to_element(menu).perform()

    # Localiza e clica em "Cadastro de Funcionários"
    opcao = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.XPATH, '//*[contains(text(), "Cadastro de Funcionários")]'))
    )

    driver.execute_script("arguments[0].click();", opcao)

# ================= CRIAÇÃO DE TAREFAS =================
def criar_tarefa(
    driver,
    empregado,
    apelido_formatado,
    cliente,
    titulo_tarefa,
    main_window
):
    """
    Cria uma tarefa de admissão no sistema para um funcionário específico.
    Preenche todos os campos necessários e retorna o número da tarefa gerada.
    """
    time.sleep(2)

    # Abre o menu de tarefas
    clicar(driver, '//*[@id="custom-header"]/div[1]/button')

    time.sleep(2)

    # Seleciona "Nova Tarefa"
    clicar(driver, '//*[@id="custom-header"]/div[2]/div/div[2]/div[2]/bm-header-app-menu/ul/bm-header-app-menu-item[1]/li/a')

    time.sleep(3)

    # Muda para a nova aba de criação de tarefa
    driver.switch_to.window(driver.window_handles[-1])

    # Seleciona o tipo de tarefa
    clicar(driver, '//*[@id="gestta-menu"]/div/div[5]/div/div/div[6]/div/div[1]/div')

    time.sleep(2)

    # Abre o campo de título
    clicar(driver, '//*[@id="modal-body"]/fieldset/div[1]/div/div/div[1]/span/span[2]/span')

    # Escreve o título da tarefa
    escrever(driver, '//*[@id="modal-body"]/fieldset/div[1]/div/div/input[1]', titulo_tarefa)

    # Escreve o nome do empregado
    escrever(driver, '//*[@id="complement"]', empregado)

    # Seleciona a empresa no campo de autocomplete
    clicar(driver, '//*[@id="gestta-multiselect-dropdown-11-p"]')

    # Digita o apelido formatado
    escrever_sem_enter(
        driver,
        '//*[@id="gestta-multiselect-dropdown-11"]/div/div/ul/li[1]/input',
        apelido_formatado
    )

    # Clica na empresa na lista
    clicar(driver, '//*[@id="gestta-multiselect-dropdown-11"]/div/div/ul/li[4]/a')

    # Fecha o dropdown
    clicar(driver, '//*[@id="gestta-multiselect-dropdown-11-p"]')

    # Clica no botão de salvar/confirmar
    clicar(driver, '//*[@id="modal-body"]/fieldset/div[5]/div/span')

    # Define quantidade (digita 1)
    clicar_1(driver, '//*[@id="modal-body"]/fieldset/div[9]/div/div/span')

    # Confirma a quantidade
    clicar(driver, '//*[@id="modal-body"]/fieldset/div[9]/div/div/div/ul/li[2]/span/button[1]')

    # Salva a tarefa (clica no botão Salvar)
    clicar(driver, '//*[@id="modal-body"]/div/button[2]')

    logging.info("⏳ Aguardando geração do número da tarefa...")

    time.sleep(5)

    # Tenta capturar o número da tarefa gerada
    try:
        xpath_protocolo = '//*[@id="mixed-task-details-view-content"]/div[1]/div[1]/div[2]/div[1]/div[2]/div/div/div[2]/div[1]/label/span[2]'

        elemento_protocolo = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, xpath_protocolo))
        )

        numero_tarefa = elemento_protocolo.text
        logging.info(f"🔢 Tarefa gerada: {numero_tarefa}")

    except Exception:
        logging.warning("⚠️ Não foi possível capturar o número da tarefa, mas ela foi salva.")
        numero_tarefa = "Não capturado"

    # Fecha a aba da tarefa e volta para a janela principal
    driver.close()
    driver.switch_to.window(main_window)

    time.sleep(2)

    return numero_tarefa

# ================= PROCESSAMENTO PRINCIPAL =================
def processar_tipo(
    driver,
    tipo,
    mapa_empresas,
    main_window
):
    """
    Processa um tipo específico (empregado ou estagiário):
    1. Acessa a aba correta
    2. Coleta dados do site
    3. Compara com dados salvos
    4. Cria tarefas para novos registros
    5. Atualiza status no Excel
    """
    logging.info(f"🚀 Processando: {tipo['nome'].upper()}")

    # Volta para janela principal e recarrega
    driver.switch_to.window(main_window)
    time.sleep(2)
    driver.refresh()
    time.sleep(4)

    # ================= ABA CORRETA =================
    if tipo["aba_estagiario"]:
        # Se for estagiário, clica na aba de estagiários
        logging.info("📂 Abrindo aba de estagiários")

        WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="ngb-nav-2"]'))
        )

        clicar(driver, '//*[@id="ngb-nav-2"]')
        time.sleep(4)

        # Espera a grid de estagiários carregar
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.ID, tipo["grid_id"]))
        )

    else:
        # Se for empregado, espera a grid de empregados carregar
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.ID, tipo["grid_id"]))
        )

    time.sleep(2)

    # Expande para mostrar mais itens por página
    clicar_1(driver, tipo["xpath_expandir"])

    time.sleep(1)

    # Aguarda a grid estar presente
    WebDriverWait(driver, 15).until(
        EC.presence_of_element_located((By.ID, tipo["grid_id"]))
    )

    # ================= COLETA DE DADOS =================
    logging.info("📊 Coletando dados do site...")
    dados_site = pegar_dados(driver, tipo["grid_id"])

    # Carrega dados já salvos
    df_total = carregar_base(tipo["arquivo"])

    # Identifica IDs já existentes na base
    ids_na_base = set(df_total["id"].astype(str))

    # Filtra apenas registros novos com ID acima do mínimo definido por tipo
    novos_registros = [
        d for d in dados_site
        if str(d["id"]) not in ids_na_base
        and int(d["id"]) > tipo["id_minimo"]
    ]

    # ================= ADICIONA NOVOS REGISTROS =================
    if novos_registros:
        df_novo = pd.DataFrame(novos_registros)
        df_novo["status"] = ""  # Status inicial vazio

        # Concatena com dados existentes
        df_total = pd.concat([df_total, df_novo], ignore_index=True)
        df_total = df_total.drop_duplicates(subset="id")  # Remove duplicatas

        salvar_base(df_total, tipo["arquivo"])
        logging.info(f"✅ {len(novos_registros)} novos registros adicionados")
    else:
        logging.info("😴 Nenhum ID novo detectado")

    # ================= PROCESSAMENTO DE PENDENTES =================
    # Filtra registros não concluídos
    pendentes = df_total[df_total["status"] != "Concluído"]

    if not pendentes.empty:
        logging.info(f"📂 Há {len(pendentes)} tarefa(s) pendente(s)")

        # Processa cada registro pendente
        for index, registro in pendentes.iterrows():
            # Formata o apelido com código da empresa
            apelido_formatado = montar_apelido_com_codigo(
                registro["apelido"],
                mapa_empresas
            )

            try:
                logging.info(f"🛠️ Criando tarefa para {registro['empregado']}")

                # Cria a tarefa
                numero_tarefa = criar_tarefa(
                    driver,
                    registro["empregado"],
                    apelido_formatado,
                    registro["cliente"],
                    tipo["titulo_tarefa"],
                    main_window
                )

                # Atualiza status e protocolo
                df_total.at[index, "status"] = "Concluído"
                df_total.at[index, "protocolo"] = numero_tarefa

                # Salva no Excel
                salvar_base(df_total, tipo["arquivo"])

                logging.info(f"✔️ Concluído: {registro['empregado']} | Prot: {numero_tarefa}")

                # Volta para janela principal
                driver.switch_to.window(main_window)
                time.sleep(2)

            except Exception as e:
                logging.error(f"❌ Erro ao processar {registro['empregado']}")
                raise e  # Propaga o erro para ser tratado no nível superior

    else:
        logging.info("⭐ Nenhuma tarefa pendente")

# ================= AUTOMAÇÃO PRINCIPAL =================
def rodar_automacao():
    """
    Função principal que orquestra toda a automação:
    1. Carrega mapa de empresas
    2. Inicia driver
    3. Faz login
    4. Navega até funcionários
    5. Processa cada tipo (empregados e estagiários)
    """
    # Carrega o mapeamento de empresas
    mapa_empresas = carregar_empresas()

    # Inicia o navegador
    driver = iniciar_driver()

    try:
        # Realiza login no sistema
        fazer_login(driver)

        # Navega até a página de funcionários
        navegar_ate_funcionarios(driver)

        time.sleep(2)

        # Salva o identificador da janela principal
        main_window = driver.current_window_handle

        # Processa cada tipo (empregado e estagiário)
        for tipo in TIPOS:
            try:
                processar_tipo(
                    driver,
                    tipo,
                    mapa_empresas,
                    main_window
                )
            except Exception:
                logging.exception(f"Erro ao processar tipo: {tipo['nome']}")
                raise  # Interrompe em caso de erro

    finally:
        # Garante que o driver seja fechado mesmo se houver erro
        driver.quit()
        logging.info("Driver finalizado")

# ================= FUNÇÃO PRINCIPAL COM TENTATIVAS =================
MAX_TENTATIVAS = 3  # Número máximo de tentativas em caso de falha

def main():
    """
    Função principal do programa.
    Tenta executar a automação até 3 vezes em caso de falha.
    Se todas falharem, envia email de alerta e encerra.
    """
    # Configura o sistema de logs
    configurar_log()

    tentativa = 0

    # Loop de tentativas
    while tentativa < MAX_TENTATIVAS:
        tentativa += 1

        logging.info(f"🚀 Tentativa {tentativa}/{MAX_TENTATIVAS} — Iniciando automação unificada")

        try:
            # Executa a automação
            rodar_automacao()

            logging.info("✅ Automação finalizada com sucesso")
            return  # Sai da função se tudo deu certo

        except Exception:
            logging.exception(f"❌ Falha na tentativa {tentativa}/{MAX_TENTATIVAS}")

            # Se ainda tem tentativas restantes, aguarda antes de tentar novamente
            if tentativa < MAX_TENTATIVAS:
                logging.info("🔁 Aguardando 10 segundos antes de nova tentativa...")
                time.sleep(10)

    # ================= ESGOTOU TENTATIVAS =================
    # Prepara mensagem de erro para email
    mensagem = (
        f"O aplicativo de Admissão tentou executar {MAX_TENTATIVAS} vezes "
        f"e o erro persistiu em todas as tentativas.\n\n"
        f"Data/Hora: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n\n"
        f"Verifique o arquivo de log para mais detalhes."
    )

    logging.critical(f"🚨 {MAX_TENTATIVAS} tentativas esgotadas. Encerrando aplicativo.")

    # Envia email de alerta
    enviar_email_erro(
        f"🚨 FALHA CRÍTICA — Automação Admissão encerrada após {MAX_TENTATIVAS} tentativas",
        mensagem
    )

    # Encerra o programa com código de erro
    sys.exit(1)

# ================= PONTO DE ENTRADA =================
# Verifica se o script está sendo executado diretamente (não importado)
if __name__ == "__main__":
    main()