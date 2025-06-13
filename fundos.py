import pandas as pd
import oracledb
import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# === Oracle Client ===
oracledb.init_oracle_client(lib_dir=r"Caminho do client Oracle")

# === Conexão Oracle ===
usuario = 'User'
senha = 'Senha'
host = 'Hostname'
porta = "Porta"
service_name = 'service_name'
nome_tabela = 'nome da tabela/view'

# === Configuração de envio de e-mail ===
EMAIL_REMETENTE = 'remetente'
EMAIL_SENHA = 'senha'
EMAIL_DESTINATARIO = 'destinatario'

def enviar_email_erro(fundo, mensagem_erro):
    assunto = f"[ERRO] Falha ao carregar fundo: {fundo}"
    corpo = f"""
Ocorreu um erro ao tentar carregar os dados do fundo: {fundo}

Detalhes do erro:
{mensagem_erro}
    """
    msg = MIMEMultipart()
    msg['From'] = EMAIL_REMETENTE
    msg['To'] = EMAIL_DESTINATARIO
    msg['Subject'] = assunto
    msg.attach(MIMEText(corpo, 'plain'))

    try:
        with smtplib.SMTP('smtp.office365.com', 587) as servidor:
            servidor.starttls()
            servidor.login(EMAIL_REMETENTE, EMAIL_SENHA)
            servidor.send_message(msg)
        print(f"📧 E-mail de erro enviado para {EMAIL_DESTINATARIO}")
    except Exception as e:
        print(f"⚠️ Falha ao enviar e-mail de erro: {e}")
        
#Renomear os campos da planilha para os nomes do campo do banco de dados
planilhas_fundos = {
    'KS': {
        'arquivo': r'Caminho do .xlsx',
        'sheet_name': 'Dados',
        'mapping': {
            'tipo': 'PAR_ST_TIPO',
            'Nome Sacado': 'CLI_ST_NOME',
            'duplicata': 'PAR_ST_CODIGO',
            'cedente': 'EMP_ST_NOME',
            'total': 'PAR_RE_VALOR',
            'Subtipo': 'PAR_ST_SUBTIPO',
        }
    },
    'SIFRA': {
        'arquivo': r'Caminho do .xlsx',
        'sheet_name': 'Planilha1',
        'mapping': {
            'Tipo': 'PAR_ST_TIPO',
            'Sacado': 'CLI_ST_NOME',
            'Nº Doc.': 'PAR_ST_CODIGO',
            'Empreendimento': 'EMP_ST_NOME',
            'Valor': 'PAR_RE_VALOR',
            'Modalidade': 'PAR_ST_SUBTIPO',
            'Emissão': 'PAR_DT_EMISSAO',
            'Vencimento': 'PAR_DT_VENCIMENTO',
        }
    },
    'RED': {
        'arquivo': r'Caminho do .xlsx',
        'sheet_name': 'Planilha1',
        'mapping': {
            'Sacado': 'CLI_ST_NOME',
            'Tipo': 'PAR_ST_CODIGO',
            'Empreendimento': 'EMP_ST_NOME',
            'Valor ': 'PAR_RE_VALOR',
            'Subtipo': 'PAR_ST_SUBTIPO',
            'Vencimento': 'PAR_DT_VENCIMENTO',
        }
    },
    'PAULISTA': {
        'arquivo': r'Caminho do .xlsx',
        'sheet_name': 'Planilha1',
        'mapping': {
            'Sacado': 'CLI_ST_NOME',
            'Tipo': 'PAR_ST_CODIGO',
            'Empreendimento': 'EMP_ST_NOME',
            'Valor ': 'PAR_RE_VALOR',
            'Subtipo': 'PAR_ST_SUBTIPO',
            'Vencimento': 'PAR_DT_VENCIMENTO',
        }
    },
    'CONTINENTAL': {
        'arquivo': r'Caminho do .xlsx',
        'sheet_name': 'Planilha1',
        'mapping': {
            'Sacado': 'CLI_ST_NOME',
            'Tipo': 'PAR_ST_CODIGO',
            'Empreendimento': 'EMP_ST_NOME',
            'Valor ': 'PAR_RE_VALOR',
            'Subtipo': 'PAR_ST_SUBTIPO',
            'Vencimento': 'PAR_DT_VENCIMENTO',
        }
    },
}

def processar_fundo(nome_fundo, config):
    arquivo = config['arquivo']
    sheet = config['sheet_name']
    mapping = config['mapping']

    if not os.path.exists(arquivo):
        print(f"❌ Planilha '{arquivo}' não encontrada.")
        enviar_email_erro(nome_fundo, f"Planilha '{arquivo}' não encontrada.")
        return

    print(f"\n📥 Processando fundo: {nome_fundo}...")

    conn = None
    cursor = None

    try:
        df = pd.read_excel(arquivo, sheet_name=sheet)

        # Substituir strings em branco por None
        df = df.replace(r'^\s*$', None, regex=True)

        # Renomear apenas colunas que existem
        df = df.rename(columns={k: v for k, v in mapping.items() if k in df.columns})
        df['BAN_ST_NOME'] = nome_fundo

        colunas_oracle = list(mapping.values()) + ['BAN_ST_NOME']
        
        # Garantir que todas as colunas existam no DataFrame
        for col in colunas_oracle:
            if col not in df.columns:
                df[col] = None

        # Tipagem automática
        colunas_varchar = [col for col in colunas_oracle if 'ST_' in col or col == 'BAN_ST_NOME']
        colunas_float = [col for col in colunas_oracle if 'RE_' in col]
        colunas_date = [col for col in colunas_oracle if 'DT_' in col]

        for col in colunas_varchar:
            df[col] = df[col].astype(str)

        for col in colunas_float:
            df[col] = pd.to_numeric(df[col], errors='coerce')

        for col in colunas_date:
            df[col] = pd.to_datetime(df[col], errors='coerce').dt.date

        # Finalização
        df = df[colunas_oracle]
        df = df.where(pd.notnull(df), None)

        dados = [tuple(x) for x in df.values]

        if not dados:
            print("⚠️ Nenhum dado encontrado para inserção.")
            return

        dsn = f"{host}:{porta}/{service_name}"
        conn = oracledb.connect(user=usuario, password=senha, dsn=dsn)
        cursor = conn.cursor()

        print(f"🧹 Limpando dados antigos de '{nome_fundo}'...")
        cursor.execute(f"DELETE FROM {nome_tabela} WHERE BAN_ST_NOME = :fundo", {'fundo': nome_fundo})
        conn.commit()
        print("✅ Dados antigos removidos.")

        placeholders = ", ".join([f":{i+1}" for i in range(len(colunas_oracle))])
        sql_insert = f"INSERT INTO {nome_tabela} ({', '.join(colunas_oracle)}) VALUES ({placeholders})"
        cursor.executemany(sql_insert, dados)
        conn.commit()
        print(f"✅ Inseridos {len(dados)} registros para '{nome_fundo}'.")

    except Exception as e:
        erro_msg = f"{e}"
        print(f"❌ Erro ao processar fundo '{nome_fundo}': {erro_msg}")
        enviar_email_erro(nome_fundo, erro_msg)

    finally:
        if cursor:
            try: cursor.close()
            except: pass
        if conn:
            try: conn.close()
            except: pass
        print("🔚 Conexão encerrada.")

# === Execução para todos os fundos ===
if __name__ == "__main__":
    for fundo, config in planilhas_fundos.items():
        processar_fundo(fundo, config)
