import requests
import json
import pandas as pd
import pprint
import traceback
import numpy as np
import os # Para ler variáveis de ambiente
import smtplib # Para enviar e-mail
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
# XlsxWriter não precisa ser importado aqui, o Pandas o usa através do 'engine'.

# As funções auxiliares (carregar_config, autenticar_e_pegar_token, obter_ranking_indicador)
# continuam exatamente as mesmas.

def carregar_config():
    """Carrega as configurações do arquivo config.json."""
    try:
        with open("config.json", "r", encoding="utf-8") as f:
            config = json.load(f)
        required_keys = ["usuario", "senha", "mes_referencia", "ano_referencia", "unidades_filtro"]
        if not all(key in config for key in required_keys):
            missing_key = next((key for key in required_keys if key not in config), "desconhecida")
            raise KeyError(f"A chave '{missing_key}' está faltando no config.json")
        return config
    except Exception as e:
        print(f"❌ Erro ao carregar 'config.json': {e}")
        return None

def autenticar_e_pegar_token(usuario, senha, mes_num, ano_ref):
    """Faz o login e retorna o token de autenticação."""
    url = "https://apibi.webdentalsolucoes.io/api/autenticar"
    payload = {
        "username": usuario, "password": senha, "mes_escolhido": mes_num,
        "ano_escolhido": ano_ref, "base": "sistemaw_clinicatodos"
    }
    print("🔐 Autenticando para obter o token...")
    response = requests.post(url, json=payload, timeout=20)
    response.raise_for_status()
    token = response.json().get("token")
    if not token: raise ValueError("Token não encontrado na resposta da autenticação.")
    print("✅ Token recebido com sucesso!")
    return token

def obter_ranking_indicador(token, mes_ano, indicador_nome):
    """Busca o ranking completo para um indicador específico."""
    url = "https://apibi.webdentalsolucoes.io/api/indicadores/indicador_mensal"
    headers = {
        "Authorization": f"Bearer {token}", "Content-Type": "application/json",
        "base": "sistemaw_clinicatodos", "origin": "https://novobi.webdentalsolucoes.io",
        "referer": "https://novobi.webdentalsolucoes.io/"
    }
    payload = { "indicador": indicador_nome, "mesanoatual": mes_ano, "periodo": 12, "unidade": "" }
    print(f"📊 Buscando dados do indicador '{indicador_nome}'...")
    response = requests.post(url, headers=headers, json=payload, timeout=30)
    response.raise_for_status()
    print(f"✅ Dados de '{indicador_nome}' recebidos!")
    return response.json().get("lista_rank_todas_unidades_mensal", [])

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
    msg['Subject'] = "Relatório Diário de Indicadores Odonto"

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


def main():
    # Carregue as configurações do seu JSON normalmente
    with open("config.json", "r", encoding="utf-8") as f:
        config = json.load(f)

    # Obtenha os dados sensíveis das variáveis de ambiente
    usuario_api = os.environ.get('API_USER')
    senha_api = os.environ.get('API_PASSWORD')
    
    if not all([usuario_api, senha_api]):
        print("❌ Variáveis de ambiente API_USER ou API_PASSWORD não definidas.")
        return

    meses_map = {
        "Janeiro": "01", "Fevereiro": "02", "Março": "03", "Abril": "04",
        "Maio": "05", "Junho": "06", "Julho": "07", "Agosto": "08",
        "Setembro": "09", "Outubro": "10", "Novembro": "11", "Dezembro": "12"
    }
    mes_num = meses_map.get(config["mes_referencia"])
    mes_ano_formatado = f"{config['ano_referencia']}-{mes_num}"
    UNIDADES_PARA_FILTRAR = config["unidades_filtro"]
    ano_ref = config["ano_referencia"]

    try:
        token_acesso = autenticar_e_pegar_token(usuario_api, senha_api, mes_num, ano_ref)
        
        ranking_cadastrados = obter_ranking_indicador(token_acesso, mes_ano_formatado, "pacientes_cadastrados")
        ranking_orcado = obter_ranking_indicador(token_acesso, mes_ano_formatado, "total_orcados")
        ranking_efetivacao = obter_ranking_indicador(token_acesso, mes_ano_formatado, "finan_total_efetivacoes")
        ranking_cadeira = obter_ranking_indicador(token_acesso, mes_ano_formatado, "finan_total_efetivacoes_cadeira")
        
        print("\n🛠️  Processando e juntando os dados...")
        
        df_cadastrados = pd.DataFrame(ranking_cadastrados)
        df_orcado = pd.DataFrame(ranking_orcado)
        df_efetivacao = pd.DataFrame(ranking_efetivacao)
        df_cadeira = pd.DataFrame(ranking_cadeira)
        
        df_base = df_cadastrados[['unidade', 'nm_unidade_atendimento', 'valor']].rename(columns={'nm_unidade_atendimento': 'Unidade', 'valor': 'N. Cadastrados'})
        df_orcado_temp = df_orcado[['unidade', 'valor']].rename(columns={'valor': 'Orçado Total'})
        df_efetivacao_temp = df_efetivacao[['unidade', 'valor']].rename(columns={'valor': 'Efetivação Total'})
        df_cadeira_temp = df_cadeira[['unidade', 'valor']].rename(columns={'valor': 'Efetivação Cadeira'})

        df_final = pd.merge(df_base, df_orcado_temp, on='unidade', how='left')
        df_final = pd.merge(df_final, df_efetivacao_temp, on='unidade', how='left')
        df_final = pd.merge(df_final, df_cadeira_temp, on='unidade', how='left')

        df_final['N. Cadastrados'] = pd.to_numeric(df_final['N. Cadastrados']).astype(int)
        df_final['Orçado Total'] = pd.to_numeric(df_final['Orçado Total']).astype(float)
        df_final['Efetivação Total'] = pd.to_numeric(df_final['Efetivação Total']).astype(float)
        df_final['Efetivação Cadeira'] = pd.to_numeric(df_final['Efetivação Cadeira']).astype(float)
        
        # --- CÁLCULO DA CONVERSÃO (AGORA COMO NÚMERO) ---
        df_final['Conversão'] = df_final['Efetivação Total'] / df_final['Orçado Total']
        df_final['Conversão'] = df_final['Conversão'].replace([np.inf, -np.inf], 0).fillna(0) # Trata divisão por zero
        
        print(f"Filtrando por {len(UNIDADES_PARA_FILTRAR)} unidades especificadas no config.json...")
        df_filtrado = df_final[df_final['Unidade'].isin(UNIDADES_PARA_FILTRAR)]
        
        if df_filtrado.empty:
            print("\n⚠️ Nenhuma das unidades especificadas foi encontrada nos resultados da API.")
            return
            
        df_filtrado = df_filtrado.sort_values(by="Unidade")
        
        colunas_finais = ['Unidade', 'N. Cadastrados', 'Orçado Total', 'Efetivação Total', 'Efetivação Cadeira', 'Conversão']
        df_para_salvar = df_filtrado[colunas_finais]

        # --- SALVANDO E FORMATANDO O EXCEL ---
        nome_arquivo = f"relatorio_indicadores_odonto.xlsx"

        writer = pd.ExcelWriter(nome_arquivo, engine='xlsxwriter')
        df_para_salvar.to_excel(writer, index=False, sheet_name='Relatório')

        # Pega os objetos do workbook e da planilha para aplicar a formatação
        workbook  = writer.book
        worksheet = writer.sheets['Relatório']

        # Cria o formato de porcentagem que será usado
        percent_format = workbook.add_format({'num_format': '0.00%'})

        # --- LÓGICA DE AUTO-AJUSTE ---
        # Itera sobre cada coluna do DataFrame para calcular e aplicar a largura
        for idx, col in enumerate(df_para_salvar.columns):
            # Encontra o tamanho máximo entre o cabeçalho e os dados da coluna
            # O .astype(str) é importante para garantir que números também sejam medidos como texto
            max_len = max(
                len(str(col)),  # Comprimento do nome do cabeçalho
                df_para_salvar[col].astype(str).str.len().max()  # Comprimento do maior dado na coluna
            ) + 2  # Adiciona um pequeno espaço extra para não ficar colado

            # Define a formatação da célula (por padrão, nenhuma)
            cell_format = None
            
            # Se a coluna atual for a 'Conversão', usamos o formato de porcentagem
            if col == 'Conversão':
                cell_format = percent_format
                # Podemos garantir uma largura mínima para a coluna de porcentagem se quisermos
                max_len = max(max_len, 12)

            # Aplica a largura e a formatação (se houver) à coluna
            worksheet.set_column(idx, idx, max_len, cell_format)

        # Salva o arquivo
        writer.close()
        
        print(f"\n🎉 Relatório FINAL '{nome_arquivo}' gerado com sucesso!")
        print("Dados gerados:")
        print(df_para_salvar.to_string())

        enviar_email(nome_arquivo)
        
    except requests.exceptions.HTTPError as e:
        print(f"\n❌ Erro na requisição HTTP: {e.response.status_code}")
        try: pprint.pprint(e.response.json())
        except json.JSONDecodeError: print(e.response.text)
            
    except Exception as e:
        print(f"\n❌ Ocorreu um erro inesperado no processo:")
        traceback.print_exc()
    finally:
        print("\n▶️ Automação finalizada.")

if __name__ == "__main__":
    main()