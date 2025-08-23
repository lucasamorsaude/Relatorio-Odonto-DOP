# Instale as bibliotecas necessárias:
# pip install selenium webdriver-manager

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import time
import pandas as pd
import smtplib # Para enviar e-mail
import os 
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders



def enviar_email(nome_arquivo_anexo):
    """Envia o e-mail com o relatório em anexo."""
    de_email = os.environ.get('EMAIL_USER')
    para_email = os.environ.get('EMAIL_TO')
    senha_email = os.environ.get('EMAIL_PASSWORD')
    
    if not all([de_email, para_email, senha_email]):
        print("❌ Variáveis de ambiente para e-mail não configuradas. E-mail não enviado.")
        return

    msg = MIMEMultipart()
    msg['From'] = de_email
    msg['To'] = para_email
    msg['Subject'] = "Relatório Diário de Crédito PF Odonto"

    corpo = "Olá,\n\nSegue em anexo o relatório de indicadores gerado hoje.\n\nAtenciosamente,\nBot do Lucas M."
    msg.attach(MIMEText(corpo, 'plain'))

    # Anexar o arquivo
    try:
        with open(nome_arquivo_anexo, "rb") as attachment:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f"attachment; filename= {nome_arquivo_anexo}")
        msg.attach(part)
        print(f"📄 Anexando arquivo '{nome_arquivo_anexo}'...")
    except FileNotFoundError:
        print(f"❌ Arquivo '{nome_arquivo_anexo}' não encontrado para anexar.")
        return

    # Conectar e enviar
    try:
        server = smtplib.SMTP('smtp.gmail.com', 587) # Exemplo para Gmail
        server.starttls()
        server.login(de_email, senha_email)
        text = msg.as_string()
        server.sendmail(de_email, para_email, text)
        server.quit()
        print(f"✅ E-mail enviado com sucesso para {para_email}!")
    except Exception as e:
        print(f"❌ Falha ao enviar o e-mail: {e}")








# --- Configurações do Navegador (Chrome) ---
options = webdriver.ChromeOptions()
options.add_argument('--headless')  # Roda o navegador em segundo plano
options.add_argument('--no-sandbox')
options.add_argument('--disable-dev-shm-usage')
options.add_argument('--window-size=1920,1080') # Define um tamanho de janela para evitar problemas de layout



# --- Inicialização do WebDriver ---
# O webdriver-manager cuida do chromedriver pra você, sem dor de cabeça.
try:
    servico = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=servico, options=options)
    # Define uma espera máxima para os elementos aparecerem na página
    wait = WebDriverWait(driver, 60) 
    print("Bot iniciado em modo headless.")

    # --- Início da Automação ---

    # 1. Acessar a URL do sistema
    print("Acessando o sistema...")
    driver.get("https://maistodos.reportload.com/signin")

    # 2. Fazer login
    # Use 'wait.until' para garantir que a página carregou antes de interagir
    print("Realizando login...")
    time.sleep(5)
    usuario_maistodos = os.environ.get('MAISTODOS_USER')
    campo_usuario = wait.until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div/div/div/main/div/div/div/div/div/form/div/div[1]/label/div/div[1]/div[2]/input')))
    campo_usuario.send_keys(usuario_maistodos)


    senha_maistodos = os.environ.get('MAISTODOS_PASSWORD')
    campo_senha = driver.find_element(By.XPATH, "/html/body/div[1]/div/div/div/main/div/div/div/div/div/form/div/div[2]/label/div/div[1]/div[2]/input")
    campo_senha.send_keys(senha_maistodos)

    time.sleep(10)

    botao_login = driver.find_element(By.XPATH, "/html/body/div[1]/div/div/div/main/div/div/div/div/div/form/div/div[5]/button")
    botao_login.click()
    print("Login efetuado com sucesso!")

    time.sleep(2)

    driver.get("https://maistodos.reportload.com/d535f1b9-73a6-4c25-b264-66ac14813cd0/reports/view")
    

    time.sleep(15)

    btn_bookmark = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="q-app"]/div/div/div/header/div/div[1]/div/button[8]')))
    btn_bookmark.click()

    seletor_css = "button i.mdi-reload"
    botao_carregar = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, seletor_css)))
    botao_carregar.click()

    # --- Mudar para o iframe ---
    print("Aguardando e mudando para o iframe do Power BI...")
    # Melhoria 1: Usar uma espera explícita para o iframe estar pronto
    iframe = wait.until(EC.frame_to_be_available_and_switch_to_it((By.TAG_NAME, 'iframe')))
    print("Foco alterado para o iframe.")
    

    seletor_xpath = "//input[starts-with(@aria-label, 'Data de início')]"
    campo_data_inicio = wait.until(EC.presence_of_element_located((By.XPATH, seletor_xpath)))
    campo_data_inicio.clear()
    campo_data_inicio.send_keys("19/08/2025")

    seletor_xpath = "//input[starts-with(@aria-label, 'Data de término')]"
    campo_data_inicio = wait.until(EC.presence_of_element_located((By.XPATH, seletor_xpath)))
    campo_data_inicio.clear()
    campo_data_inicio.send_keys("19/08/2025")


    try:
        print("Iniciando a extração dos dados da tabela...")

        # 1. Espera o contêiner da tabela carregar
        wait.until(EC.presence_of_element_located((By.XPATH, "//div[@role='grid']")))
        time.sleep(3) 

        # 2. Define os cabeçalhos da tabela manualmente
        cabecalhos = ['Rank', 'Unidade', 'Simulações', 'Simulações únicas', 'Taxa de simulações', 'Total contratado']
        
        # 3. Pega todas as linhas de dados
        linhas_elements = driver.find_elements(By.XPATH, "//div[@role='row' and @aria-rowindex]")
        tabela_dados = []
        for linha in linhas_elements:
            celulas = linha.find_elements(By.XPATH, ".//div[@role='gridcell' and @column-index]")
            dados_da_linha = [celula.text for celula in celulas[1:]]
            if len(dados_da_linha) == len(cabecalhos):
                tabela_dados.append(dados_da_linha)
        print(f"Extração concluída. {len(tabela_dados)} linhas de dados encontradas.")

        # Cria o DataFrame inicial
        df = pd.DataFrame(tabela_dados, columns=cabecalhos)

        # 4. Filtra o DataFrame
        df_filtrado = df[df['Unidade'].str.startswith('amor saude', na=False)].copy()
        
        # 5. Ordena o DataFrame
        df_final_ordenado = df_filtrado.sort_values(by='Unidade').reset_index(drop=True)

        print("\n--- Tabela Final, Filtrada e Ordenada ---")
        print(df_final_ordenado)

        # 6. --- NOVO PASSO: SALVAR COM AUTOAJUSTE DE COLUNAS ---
        # Substituímos a linha to_excel simples por este bloco
        
        arquivo_saida = "relatorio_credito_pf_odonto.xlsx"
        writer = pd.ExcelWriter(arquivo_saida, engine='xlsxwriter')

        # Escreve o DataFrame no arquivo Excel
        df_final_ordenado.to_excel(writer, sheet_name='Ranking', index=False)

        # Pega os objetos do workbook e da planilha para poder manipulá-los
        workbook  = writer.book
        worksheet = writer.sheets['Ranking']

        # Loop para percorrer cada coluna
        for i, col in enumerate(df_final_ordenado.columns):
            # Encontra o tamanho máximo do texto na coluna (incluindo o cabeçalho)
            tamanho_coluna = max(df_final_ordenado[col].astype(str).map(len).max(), len(col))
            # Define a largura da coluna no Excel (com um pequeno espaço extra)
            worksheet.set_column(i, i, tamanho_coluna + 2)

        # Salva o arquivo Excel
        writer.close()
        print(f"\nArquivo '{arquivo_saida}' salvo com colunas ajustadas automaticamente!")

        enviar_email(arquivo_saida)



    except Exception as e:
        print(f"Ocorreu um erro ao extrair a tabela: {e}")
    
    print("Tarefa concluída!")


finally:
    # --- Finalização ---
    if 'driver' in locals():
        driver.quit()
        print("Navegador fechado.")