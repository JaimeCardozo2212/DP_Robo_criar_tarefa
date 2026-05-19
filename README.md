================================================================================
                    DOCUMENTAÇÃO TÉCNICA
          SISTEMA DE AUTOMAÇÃO DE ADMISSÃO ONVIO
                          Versão 1.1
================================================================================

ÍNDICE
--------------------------------------------------------------------------------
1. VISÃO GERAL DO SISTEMA
2. REQUISITOS DO SISTEMA
3. ESTRUTURA DE ARQUIVOS
4. CONFIGURAÇÃO INICIAL
5. CONFIGURAÇÃO DO CHROMEDRIVER
6. COMO EXECUTAR
7. FUNCIONAMENTO DETALHADO
8. SISTEMA DE TOLERÂNCIA A FALHAS
9. FORMATO DOS LOGS
10. TROUBLESHOOTING
11. MONITORAMENTO E MANUTENÇÃO
12. PERSONALIZAÇÕES POSSÍVEIS
13. SEGURANÇA
14. SUPPORT E CONTATO
15. CHANGELOG
16. CHECKLIST DE IMPLANTAÇÃO


================================================================================
1. VISÃO GERAL DO SISTEMA
================================================================================

O Sistema de Automação de Admissão Onvio é uma ferramenta desenvolvida em Python 
para automatizar o processo de criação de tarefas de admissão de funcionários 
(empregados e estagiários) no sistema web Onvio.

1.1 OBJETIVOS PRINCIPAIS
--------------------------------------------------------------------------------
✅ Extrair automaticamente dados de novos funcionários do sistema Onvio
✅ Criar tarefas de admissão de forma automatizada
✅ Manter um histórico local em arquivos Excel
✅ Gerar logs detalhados para auditoria e debug
✅ Enviar alertas por email em caso de falhas críticas

1.2 FLUXO DE TRABALHO
--------------------------------------------------------------------------------
Início → Login Onvio → Acessar Funcionários → Coletar dados → 
Comparar com Excel → Criar tarefas → Atualizar Excel → Registrar logs → Fim


================================================================================
2. REQUISITOS DO SISTEMA
================================================================================

2.1 HARDWARE RECOMENDADO
--------------------------------------------------------------------------------
- Processador: 2.0 GHz ou superior
- RAM: 4 GB mínimo (8 GB recomendado)
- Espaço em disco: 500 MB para o programa + espaço para logs

2.2 SOFTWARE NECESSÁRIO
--------------------------------------------------------------------------------
- Sistema Operacional: Windows 10/11, Linux ou macOS
- Python: Versão 3.8 ou superior
- Google Chrome: Última versão estável
- ChromeDriver: Compatível com a versão do Chrome

2.3 DEPENDÊNCIAS PYTHON
--------------------------------------------------------------------------------
Para instalar as dependências necessárias, execute:

pip install selenium pandas python-dotenv openpyxl

Caso queira usar o WebDriver Manager (recomendado):

pip install webdriver-manager


================================================================================
3. ESTRUTURA DE ARQUIVOS
================================================================================

Projeto/
│
├── App_ferias_1.0.py           # Script principal
├── Empresas_1.xlsx              # Mapeamento de empresas
├── dados_empregados.xlsx        # Base de empregados (gerado)
├── dados_estagiarios.xlsx       # Base de estagiários (gerado)
├── .env                         # Variáveis de ambiente (criar manualmente)
├── .gitignore                   # Arquivos ignorados pelo Git
│
├── logs/                        # Pasta de logs (criada automaticamente)
│   └── admissao/
│       └── mes_ano/
│           └── log_admissao_unificada_YYYY-MM-DD.log
│
└── README.md                    # Este documento


================================================================================
4. CONFIGURAÇÃO INICIAL
================================================================================

4.1 ARQUIVO .env (OBRIGATÓRIO)
--------------------------------------------------------------------------------
Crie um arquivo .env na raiz do projeto com o seguinte conteúdo:

# Credenciais de acesso ao Onvio
EMAIL=seu_email@onvio.com.br
SENHA=sua_senha

# Configurações de email para alertas
EMAIL_REMETENTE=seu_email@gmail.com
EMAIL_SENHA=senha_app_gmail
EMAIL_DESTINO=destinatario@empresa.com

⚠️ IMPORTANTE: Nunca compartilhe ou faça commit do arquivo .env!

4.2 ARQUIVO Empresas_1.xlsx
--------------------------------------------------------------------------------
Crie um arquivo Excel com a seguinte estrutura:

| Apelido      | Código |
|--------------|--------|
| empresa x    | 123    |
| empresa y    | 456    |

COLUNAS OBRIGATÓRIAS:
- Apelido: Nome curto da empresa (será padronizado em minúsculas)
- Código: Código numérico da empresa no sistema

4.3 ESTRUTURA DOS ARQUIVOS DE DADOS
--------------------------------------------------------------------------------
Os arquivos dados_empregados.xlsx e dados_estagiarios.xlsx terão a estrutura:

| id    | empregado    | apelido   | cliente   | admissao    | status      | protocolo |
|-------|--------------|-----------|-----------|-------------|-------------|-----------|
| 12345 | João Silva   | empresa x | Cliente A | 01/01/2025  | Concluído   | TAR-001   |


================================================================================
5. CONFIGURAÇÃO DO CHROMEDRIVER
================================================================================

5.1 MÉTODO AUTOMÁTICO (RECOMENDADO)
--------------------------------------------------------------------------------
Instale o webdriver-manager:

pip install webdriver-manager

Modifique a função iniciar_driver() no código:

from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service

def iniciar_driver():
    options = Options()
    options.add_argument("--headless=new")
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    return driver

5.2 MÉTODO MANUAL
--------------------------------------------------------------------------------
1. Baixe o ChromeDriver em: https://chromedriver.chromium.org/
2. Coloque o executável na pasta do projeto ou no PATH do sistema


================================================================================
6. COMO EXECUTAR
================================================================================

6.1 EXECUÇÃO NORMAL
--------------------------------------------------------------------------------
python App_ferias_1.0.py

6.2 COMO EXECUTÁVEL (.EXE)
--------------------------------------------------------------------------------
# Instalar PyInstaller
pip install pyinstaller

# Gerar executável
pyinstaller --onefile --console --name "AdmissaoOnvio" App_ferias_1.0.py

6.3 AGENDAMENTO NO WINDOWS (TASK SCHEDULER)
--------------------------------------------------------------------------------
1. Abra o Agendador de Tarefas
2. Criar tarefa básica
3. Configurar gatilho (ex: diário às 08:00)
4. Ação: Iniciar programa → caminho do .exe
5. Concluir


================================================================================
7. FUNCIONAMENTO DETALHADO
================================================================================

7.1 FLUXO DE EXECUÇÃO
--------------------------------------------------------------------------------
FASE 1 - INICIALIZAÇÃO
- Carrega configurações, logs e mapeamento de empresas

FASE 2 - LOGIN
- Acessa o Onvio com credenciais do .env

FASE 3 - NAVEGAÇÃO
- Navega até a página de funcionários

FASE 4 - COLETA
- Extrai dados da tabela (grid) do site

FASE 5 - COMPARAÇÃO
- Identifica novos IDs não processados

FASE 6 - CRIAÇÃO DE TAREFAS
- Cria tarefa para cada novo funcionário

FASE 7 - ATUALIZAÇÃO
- Salva status e protocolo no Excel

FASE 8 - FINALIZAÇÃO
- Fecha driver e registra conclusão

7.2 TIPOS DE PROCESSAMENTO
--------------------------------------------------------------------------------
| TIPO       | ABA          | GRID ID                              | ARQUIVO                  |
|------------|--------------|--------------------------------------|--------------------------|
| Empregado  | Empregados   | gridPayrollRegistrationEmployeeList | dados_empregados.xlsx    |
| Estagiário | Estagiários  | gridPayrollEntriesInternList        | dados_estagiarios.xlsx   |

7.3 SISTEMA DE LOGS
--------------------------------------------------------------------------------
ESTRUTURA DE PASTAS:
logs/admissao/janeiro_2025/log_admissao_unificada_2025-01-15.log

NÍVEIS DE LOG:
- INFO: Informações gerais da execução
- WARNING: Avisos (ex: apelido não encontrado)
- ERROR: Erros recuperáveis
- CRITICAL: Erros fatais


================================================================================
8. SISTEMA DE TOLERÂNCIA A FALHAS
================================================================================

8.1 TENTATIVAS AUTOMÁTICAS
--------------------------------------------------------------------------------
O sistema tenta executar 3 vezes em caso de falha (configurável):

MAX_TENTATIVAS = 3

8.2 CONTROLE DE EMAIL
--------------------------------------------------------------------------------
- Máximo de 1 email a cada 5 minutos
- Evita spam em loops de erro

8.3 TIMEOUTS
--------------------------------------------------------------------------------
- Clique em elementos: 15 segundos
- Carregamento de páginas: 20 segundos
- Captura de protocolo: 20 segundos


================================================================================
9. FORMATO DOS LOGS
================================================================================

9.1 EXEMPLO DE LOG DE SUCESSO
--------------------------------------------------------------------------------
2025-01-15 10:30:45 [INFO] Iniciando Driver
2025-01-15 10:30:48 [INFO] Email validado
2025-01-15 10:30:51 [INFO] Senha validada
2025-01-15 10:30:55 [INFO] Logado
2025-01-15 10:31:20 [INFO] 🚀 Processando: EMPREGADO
2025-01-15 10:31:25 [INFO] 📊 Coletando dados do site...
2025-01-15 10:31:30 [INFO] 😴 Nenhum ID novo detectado
2025-01-15 10:31:30 [INFO] ⭐ Nenhuma tarefa pendente
2025-01-15 10:31:35 [INFO] Driver finalizado
2025-01-15 10:31:35 [INFO] ✅ Automação finalizada com sucesso

9.2 EXEMPLO DE LOG COM ERRO
--------------------------------------------------------------------------------
2025-01-15 10:30:45 [INFO] Iniciando Driver
2025-01-15 10:30:48 [ERROR] Erro ao carregar Empresas_1.xlsx
2025-01-15 10:30:48 [ERROR] ❌ Falha na tentativa 1/3
2025-01-15 10:30:58 [INFO] 🔁 Aguardando 10 segundos antes de nova tentativa...


================================================================================
10. TROUBLESHOOTING (SOLUÇÃO DE PROBLEMAS)
================================================================================

10.1 ERRO: CHROMEDRIVER NÃO ENCONTRADO
--------------------------------------------------------------------------------
SOLUÇÃO: Use o webdriver-manager
pip install webdriver-manager

10.2 ERRO: SESSÃO NÃO CRIADA
--------------------------------------------------------------------------------
SOLUÇÃO: Atualize o Chrome e o ChromeDriver
Baixar ChromeDriver mais recente: https://chromedriver.chromium.org/

10.3 ERRO: ARQUIVO .ENV NÃO CARREGADO
--------------------------------------------------------------------------------
SOLUÇÃO: Verifique se o arquivo está na raiz e tem o formato correto
# Verificar se o arquivo existe
dir .env  # Windows
ls -la | grep .env  # Linux/Mac

10.4 ERRO: ELEMENTO NÃO ENCONTRADO (XPATH)
--------------------------------------------------------------------------------
CAUSAS POSSÍVEIS:
- O site mudou a estrutura HTML
- Timeout insuficiente
- Elemento ainda não carregado

SOLUÇÃO: Aumentar timeout ou atualizar XPATHs

10.5 LOGS NÃO SENDO GERADOS
--------------------------------------------------------------------------------
VERIFIQUE:
- Permissões de escrita na pasta logs/
- Se a pasta foi criada automaticamente
- Encoding UTF-8 disponível


================================================================================
11. MONITORAMENTO E MANUTENÇÃO
================================================================================

11.1 O QUE MONITORAR
--------------------------------------------------------------------------------
| ITEM                     | FREQUÊNCIA | AÇÃO EM CASO DE ANOMALIA      |
|--------------------------|------------|-------------------------------|
| Logs de erro             | Diária     | Investigar causa raiz         |
| Tamanho da pasta logs    | Mensal     | Limpar logs antigos           |
| Versão do Chrome         | Semanal    | Atualizar ChromeDriver        |
| Arquivos Excel           | Diária     | Verificar integridade         |

11.2 LIMPEZA DE LOGS (RECOMENDADO)
--------------------------------------------------------------------------------
Crie um script para limpar logs antigos:

def limpar_logs(dias=30):
    """Remove logs mais antigos que X dias"""
    import os, time
    pasta_logs = "logs"
    agora = time.time()
    
    for root, dirs, files in os.walk(pasta_logs):
        for arquivo in files:
            if arquivo.endswith('.log'):
                caminho = os.path.join(root, arquivo)
                if os.path.getmtime(caminho) < agora - (dias * 86400):
                    os.remove(caminho)
                    print(f"Removido: {caminho}")


================================================================================
12. PERSONALIZAÇÕES POSSÍVEIS
================================================================================

12.1 ALTERAR NÚMERO DE TENTATIVAS
--------------------------------------------------------------------------------
MAX_TENTATIVAS = 5  # Altere conforme necessidade

12.2 ADICIONAR NOVO TIPO DE FUNCIONÁRIO
--------------------------------------------------------------------------------
TIPOS.append({
    "nome": "aprendiz",
    "arquivo": "dados_aprendizes.xlsx",
    "grid_id": "gridApprenticeList",
    "aba_estagiario": False,
    "titulo_tarefa": "DP - ADMISSÃO APRENDIZ",
    "xpath_expandir": '//*[@id="items-per-page-2"]'
})

12.3 ALTERAR INTERVALO DE EMAIL
--------------------------------------------------------------------------------
def pode_enviar_email():
    # Mudar de 300 (5min) para 600 (10min)
    if agora - ULTIMO_EMAIL > 600:
        ...


================================================================================
13. SEGURANÇA
================================================================================

13.1 BOAS PRÁTICAS
--------------------------------------------------------------------------------
✅ Nunca commitar arquivo .env
✅ Usar token de acesso em vez de senha quando possível
✅ Limitar acesso à pasta de logs
✅ Rotacionar credenciais periodicamente

13.2 .GITIGNORE RECOMENDADO
--------------------------------------------------------------------------------
# Secrets
.env
*.env

# Python
__pycache__/
*.pyc

# Logs
*.log
logs/

# Virtual Environment
.venv/
venv/

# IDE
.vscode/
.idea/

# Excel temporários
~$*.xlsx


================================================================================
14. SUPORTE E CONTATO
================================================================================

14.1 LOGS PARA SUPORTE
--------------------------------------------------------------------------------
Ao reportar problemas, inclua:
1. Último arquivo de log gerado
2. Versão do Python (python --version)
3. Versão do Chrome e ChromeDriver
4. Mensagem de erro completa

14.2 INFORMAÇÕES PARA REPORTAR
--------------------------------------------------------------------------------
# Coletar informações do sistema
python --version
pip list | grep -E "selenium|pandas|dotenv"
# Versão do Chrome: chrome://settings/help


================================================================================
15. CHANGELOG
================================================================================

| VERSÃO | DATA       | ALTERAÇÕES                                      |
|--------|------------|-------------------------------------------------|
| 1.0    | 2025-01-15 | Versão inicial - Suporte a empregados e estagiários |


================================================================================
16. CHECKLIST DE IMPLANTAÇÃO
================================================================================

☐ Python 3.8+ instalado
☐ Dependências instaladas (pip install -r requirements.txt)
☐ Chrome atualizado
☐ ChromeDriver configurado
☐ Arquivo .env criado com credenciais
☐ Arquivo Empresas_1.xlsx criado
☐ Pasta logs/ com permissão de escrita
☐ Teste de execução manual bem-sucedido
☐ Agendamento configurado (se necessário)
☐ Monitoramento de logs ativado
☐ Backup dos arquivos Excel configurado


================================================================================
17. CONCLUSÃO
================================================================================

Este sistema automatiza completamente o processo de criação da tarefade admissão 
de funcionários no Onvio, reduzindo trabalho manual e eliminando erros de digitação. 
Com logs detalhados, sistema de tentativas e alertas por email, a ferramenta é robusta 
e confiável para uso em produção.

--------------------------------------------------------------------------------
Desenvolvido para Departamento de TI
Jaime Cardozo | Data: Maio/2026
================================================================================