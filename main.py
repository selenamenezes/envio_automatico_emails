from datetime import datetime
import os
import pandas as pd
import time
import win32com.client as win32
import re
import config as c
import traceback
from typing import Optional, Iterable


# MANTER O COMPUTADOR ATIVO DURANTE A EXECUÇÃO DO SCRIPT
ES_CONTINUOUS = 0x80000000
ES_SYSTEM_REQUIRED = 0x00000001
ES_DISPLAY_REQUIRED = 0x00000002

#IMAGEM PARA O CORPO DO E-MAIL
CID_IMAGEM = "caminhao1"

#CONFIGURAÇÕES DE DATA E HORA
agora = datetime.now().date()
timestamp_transito = os.path.getmtime(c.PLANILHA_TRANSITO)
ultima_att = datetime.fromtimestamp(timestamp_transito).date()
data_hoje = datetime.today().strftime('%d/%m/%Y')

SETORES = ["CENTRO-OESTE", "METROPOLITANO", "NORTE", "OESTE", "SUL", "CENTRO", "BLINDAGEM"]
SETOR_EXPEDIDOR = ["CENTRO-OESTE", "METROPOLITANO", "NORTE", "OESTE", "SUL", "CENTRO", "BLINDAGEM", "AMARA"]

DF = pd.read_excel(c.PLANILHA_TRANSITO, sheet_name="RELAÇÃO_PEDIDOS")
TIPO = ["UTD", "COMERCIAL", "TECNICA"]

#DF_OM = pd.read_excel(c.PLANILHA_TRANSITO, sheet_name="OM")

# ARQUIVOS E CONTATOS
UTD = c.CONTATO_UTD
TECNICA = c.CONTATO_TECNICA
COMERCIAL = c.CONTATO_COMERCIAL
ARQUIVOS_UTD = c.ARQUIVOS_UTD
ARQUIVOS_TECNICA = c.ARQUIVOS_TECNICA
ARQUIVOS_COMERCIAL = c.ARQUIVOS_COMERCIAL
ARQUIVOS_UTD_SEM_DEP = c.ARQUIVOS_UTD_SEM_DEP
ARQUIVOS_COMERCIAL_SEM_DEP = c.ARQUIVOS_COMERCIAL_SEM_DEP
ARQUIVOS_TECNICA_SEM_DEP = c.ARQUIVOS_TECNICA_SEM_DEP
ARQUIVOS_OM = c.ARQUIVOS_OM

def limpar_xlsx_na_pasta(pasta: str) -> None:
    if not pasta:
        return

    os.makedirs(pasta, exist_ok=True)

    try:
        for nome in os.listdir(pasta):
            if nome.lower().endswith(".xlsx"):
                caminho = os.path.join(pasta, nome)
                try:
                    os.remove(caminho)
                except PermissionError:
                    print(f"[AVISO] nao foi possível remover: {caminho}")
                except Exception as e:
                    print(f"[AVISO] falha ao remover {caminho}: {e}")
    except FileNotFoundError:
        return


def atualizar_transito_zero_2025():

    consultas = ["Consulta - ZMM94", "Consulta - ZMM94 (2)", "Consulta - DATA BASE", "Consulta - RELAÇÃO_PEDIDOS"]

    excel = win32.DispatchEx("Excel.Application")
    excel.DisplayAlerts = False
    excel.EnableEvents = False
    excel.AskToUpdateLinks = False
    wb = None
    try:
        wb = excel.Workbooks.Open(c.PLANILHA_TRANSITO)
        for consulta in consultas:
            print(f"Atualizando Consulta: {consulta}")
            try:
                connection = next((c for c in wb.Connections if c.Name == consulta), None)
                if connection and connection.Type == 1:
                    oledb = connection.OLEDBConnection
                    oledb.BackgroundQuery = False
                    oledb.Refresh()
                    print(f"\tConsulta '{consulta}' atualizada!")
                else:
                    print(f"Conexão '{consulta}' não encontrada ou inválida.")
            except Exception as e:
                print(f"Erro ao atualizar '{consulta}': {e}")
                print(traceback.format_exc())
        wb.Close(SaveChanges=1)
    except Exception as e:
        print(f"Erro ao abrir/atualizar planilha {c.PLANILHA_TRANSITO}: {e}")
        print(traceback.format_exc())
    finally:
        if wb:
            wb = None
        excel.Quit()
        excel = None

def filtro_transito(df, tipos, setores):

    df["SETOR"] = df["SETOR"].astype(str).str.strip().str.upper()
    df["SETOR_EXPEDIDOR"] = df["SETOR_EXPEDIDOR"].astype(str).str.strip().str.upper()
    df["TIPO"] = df["TIPO"].astype(str).str.strip().str.upper()
    df["Tipo de TR"] = df["Tipo de TR"].astype(str).str.strip().str.upper()
    df["ATIVO"] = df["ATIVO"].astype(str).str.strip().str.upper()

    pasta_1 = c.PASTA_TRANSITO_REDE
    pasta_2 = c.PASTA_TRANSITO_BD

    os.makedirs(pasta_1, exist_ok=True)
    os.makedirs(pasta_2, exist_ok=True)

    for tipo in tipos:
        for setor in setores:
            dfs_validos = []

            filtros = [
                (df["SETOR"] == setor) & ((df["Tipo de TR"] == "DE-PARA") | (df["Tipo de TR"] == "EXPEDIÇÃO - MEDIDOR") | (df["Tipo de TR"] == "EXPEDIÇÃO")) & (df["ATIVO"] == "SIM") & (df["TIPO"] == tipo) & ((df["Status"] == "10 DIAS") | (df["Status"] == "+ 10 DIAS")),
                (df["SETOR_EXPEDIDOR"] == setor) & ((df["Tipo de TR"] == "REVERSA") | (df["Tipo de TR"] == "SUCATA") | (df["Tipo de TR"] == "REFORMA") | (df["Tipo de TR"] == "SUCATA - ETM")) & (df["ATIVO"] == "SIM") & (df["TIPO"] == tipo) & ((df["Status"] == "10 DIAS") | (df["Status"] == "+ 10 DIAS"))
            ]

            for f in filtros:
                df_filtrado = df[f]
                if not df_filtrado.empty:
                    dfs_validos.append(df_filtrado)

            if dfs_validos:
                df_final = pd.concat(dfs_validos)
                nome_arquivo = f"{tipo}_{setor}.xlsx"
                caminho_1 = os.path.join(pasta_1, nome_arquivo)
                caminho_2 = os.path.join(pasta_2, nome_arquivo)

                df_final.to_excel(caminho_1, index=False)
                df_final.to_excel(caminho_2, index=False)

                print(f"Arquivos salvos em:\n- {caminho_1}\n- {caminho_2}")
            else:
                print(f"Nenhum dado para {tipo} - {setor}")
                

def filtro_transito_sem_recebedor(df, tipos, setores):
    pasta_1 = c.PASTA_SEM_DEP_REDE
    pasta_2 = c.PASTA_SEM_DEP
    os.makedirs(pasta_1, exist_ok=True)
    os.makedirs(pasta_2, exist_ok=True)

    for tipo in tipos:
        for setor in setores:
            dfs_validos = []

            filtros = [
                (df["SETOR"] == setor) & ((df["Tipo de TR"] == "DE-PARA") | (df["Tipo de TR"] == "EXPEDIÇÃO - MEDIDOR") | (df["Tipo de TR"] == "EXPEDIÇÃO")) & (df["ATIVO"] == "SIM") & (df["TIPO"] == tipo) & ((df["Status"] == "10 DIAS") | (df["Status"] == "+ 10 DIAS")) & (df["SEM DEP"] == "SIM"),
                (df["SETOR_EXPEDIDOR"] == setor) & ((df["Tipo de TR"] == "REVERSA") | (df["Tipo de TR"] == "SUCATA") | (df["Tipo de TR"] == "REFORMA") | (df["Tipo de TR"] == "SUCATA - ETM")) & (df["ATIVO"] == "SIM") & (df["TIPO"] == tipo) & ((df["Status"] == "10 DIAS") | (df["Status"] == "+ 10 DIAS")) & (df["SEM DEP"] == "NAO")
            ]

            for f in filtros:
                df_filtrado = df[f]
                if not df_filtrado.empty:
                    dfs_validos.append(df_filtrado)

            if dfs_validos:
                df_final = pd.concat(dfs_validos)
                nome_arquivo = f"{tipo}_{setor}_SEM_RECEBEDOR.xlsx"
                caminho_1 = os.path.join(pasta_1, nome_arquivo)
                caminho_2 = os.path.join(pasta_2, nome_arquivo)

                df_final.to_excel(caminho_1, index=False)
                df_final.to_excel(caminho_2, index=False)

                print(f"Arquivos salvos")
            else:
                print(f"Nenhum dado para {tipo} - {setor}")


def filtro_om_pendente_encerramento(df, setores):

    pasta_om = c.PASTA_OM
    os.makedirs(pasta_om, exist_ok=True)

    colunas_necessarias = {'STATUS BI', 'ZMM94', 'SETOR', 'Tipo de Nota Fiscal'}
    faltantes = colunas_necessarias - set(df.columns)
    if faltantes:
        raise KeyError(f"Colunas ausentes no DataFrame: {sorted(faltantes)}")

    
    df['STATUS BI'] = df['STATUS BI'].astype(str).str.strip()
    df['ZMM94'] = df['ZMM94'].astype(str).str.strip()
    df['SETOR'] = df['SETOR'].astype(str).str.strip()
    df['Tipo de Nota Fiscal'] = df['Tipo de Nota Fiscal'].astype(str).str.strip()

    def nome_seguro(texto):

        return re.sub(r'[\\/:*?"<>|]+', '_', str(texto))

    for setor in setores:
        filtro = (
            (df['STATUS BI'] == "Em andamento") &
            (df['ZMM94'] == 'TRANSITO ENCERRADO') &
            (df['SETOR'] == setor) &
            (df['Tipo de Nota Fiscal'] == "Neoenergia")
        )

        df_filtrado = df.loc[filtro].copy()

        if not df_filtrado.empty:

            df_filtrado = df_filtrado.drop_duplicates().reset_index(drop=True)

            nome_arquivo = f"ENCERRAR_OM_{nome_seguro(setor)}.xlsx"
            caminho = os.path.join(pasta_om, nome_arquivo)

            df_filtrado.to_excel(caminho, index=False, engine="openpyxl")

            print(f"Arquivo salvo: {caminho} ({len(df_filtrado)} linhas)")
        else:
            print(f"Nenhum dado para o setor: {setor}")


def enviar_emails(tipo_nome, emails_dict, arquivos_dict, cc):
    outlook = win32.Dispatch("Outlook.Application")

    for setor, lista_emails in emails_dict.items():
        anexos = []

        caminho_principal = arquivos_dict.get(setor)
        if caminho_principal and os.path.exists(caminho_principal):
            anexos.append(caminho_principal)

        if anexos and lista_emails:
            mail = outlook.CreateItem(0)
            mail.To = lista_emails[0]
            mail.Subject = f"[TRANSITO] {tipo_nome} - {setor.upper()} - {data_hoje}"
            mail.HTMLBody = c.CORPO_TRANSITO
            mail.CC = cc

            for anexo in anexos:
                mail.Attachments.Add(anexo)

            if os.path.exists(c.CAMINHO_IMAGEM):
                imagem = mail.Attachments.Add(c.CAMINHO_IMAGEM)
                imagem.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", CID_IMAGEM)

            mail.Send()
            time.sleep(300)
            print(f"E-mail enviado para {tipo_nome} - {setor.upper()} com {len(anexos)} anexo(s).")
        else:
            print(f"Nenhum e-mail enviado para {tipo_nome} - {setor.upper()} (sem anexos ou destinatários).")

#def enviar_emails_om(emails_dict, arquivos_dict, cc):
#    outlook = win32.Dispatch("Outlook.Application")
#
#    for setor,lista_emails in emails_dict.items():
#        anexos = []
#
#        caminho_principal = arquivos_dict.get(setor)
#        if caminho_principal and os.path.exists(caminho_principal):
#            anexos.append(caminho_principal)
#
#        if anexos and lista_emails:
#            mail = outlook.CreateItem(0)
#            mail.To = lista_emails[0]
#            mail.Subject = f"[OM]: PENDÊNCIAS DE ENCERRAMENTO - {setor.upper()} | {data_hoje}"
#            mail.HTMLBody = c.CORPO_OM
#            mail.CC = cc

#            for anexo in anexos:
#                mail.Attachments.Add(anexo)


#            mail.Send()
#            time.sleep(300)
#            print(f"Alerta OM enviado para {setor.upper()} com {len(anexos)} anexo")
#        else:
#            print(f"Alerta OM enviado para {setor} (sem anexos ou destinatários)")


#limpar_xlsx_na_pasta(c.PASTA_TRANSITO_REDE)
#limpar_xlsx_na_pasta(c.PASTA_TRANSITO_BD)

#filtro_transito(DF, TIPO, SETORES)
#time.sleep(50)

enviar_emails("UTD", UTD, ARQUIVOS_UTD, c.CC)
time.sleep(600)
enviar_emails("TECNICA", TECNICA, ARQUIVOS_TECNICA, c.CC)
time.sleep(600)
enviar_emails("COMERCIAL", COMERCIAL, ARQUIVOS_COMERCIAL, c.CC)

#filtro_om_pendente_encerramento(DF_OM, SETORES)
#enviar_emails_om(UTD, ARQUIVOS_OM, c.CC)
#enviar_emails_om(TECNICA, ARQUIVOS_OM, c.CC)
#enviar_emails_om(COMERCIAL, ARQUIVOS_OM, c.CC)
