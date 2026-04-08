import re
import time
from datetime import datetime
from pathlib import Path

import pandas as pd
import win32com.client as win32

# --- Configurações de Caminhos e Constantes ---
PASTA_DOCUMENTOS = Path.home() / "Documents" / "Fechamento"
ARQUIVO_LANCAMENTOS = PASTA_DOCUMENTOS / "lancamento.xlsx"
ARQUIVO_ASSOCIADOS = PASTA_DOCUMENTOS / "ACIP.xlsx"

ASSUNTO_EMAIL = "Fechamento"
ENVIAR_DE_VERDADE = False  # Se False, apenas salva no rascunho do Outlook
ASSINATURA_NOME = "William - ACIP"


def normalizar_nome_empresa(texto: str) -> str:
    """Limpa e padroniza o nome da empresa para busca no mapa de e-mails."""
    if texto is None:
        return ""
    texto = str(texto).strip().lower()
    texto = re.sub(r"\s+", " ", texto)
    return texto


def mapear_emails_dos_associados(df_associados: pd.DataFrame) -> dict:
    """Cria um dicionário vinculando nomes (fantasia ou razão social) aos e-mails de cobrança."""
    mapa_emails = {}

    for _, linha in df_associados.iterrows():
        email = str(linha.get("Email cobrança", "")).strip()
        
        if not email or email.lower() == "nan":
            continue

        nome_fantasia = normalizar_nome_empresa(linha.get("Nome Fantasia", ""))
        razao_social = normalizar_nome_empresa(linha.get("Razão Social", ""))

        if nome_fantasia:
            mapa_emails[nome_fantasia] = email
        if razao_social:
            mapa_emails[razao_social] = email

    return mapa_emails


def formatar_data(valor) -> str:
    """Converte valores de data para o formato brasileiro DD/MM/AAAA."""
    if pd.isna(valor):
        return ""
    try:
        data_dt = pd.to_datetime(valor, dayfirst=True, errors="coerce")
        if not pd.isna(data_dt):
            return data_dt.strftime("%d/%m/%Y")
    except Exception:
        pass
    return str(valor).strip()


def formatar_valor_moeda(valor) -> str:
    """Formata números para o padrão de moeda brasileiro (R$ 1.234,56)."""
    if pd.isna(valor):
        return ""
    
    texto_valor = str(valor).strip()
    if not texto_valor:
        return ""

    limpo = texto_valor.replace("R$", "").strip()

    try:
        if "," in limpo and "." in limpo:
            numero = float(limpo.replace(".", "").replace(",", "."))
        elif "," in limpo:
            numero = float(limpo.replace(",", "."))
        else:
            numero = float(limpo)

        saida = f"{numero:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        return f"R$ {saida}"
    except Exception:
        return texto_valor


def obter_aplicativo_outlook():
    """Tenta conectar ao Outlook de várias formas para evitar erro de classe inválida."""
    try:
        # Tenta pegar a instância que já está aberta
        return win32.GetActiveObject("Outlook.Application")
    except Exception:
        try:
            # Tenta abrir uma nova instância
            return win32.Dispatch("Outlook.Application")
        except Exception as e:
            print(f"Erro crítico: Não foi possível encontrar o Outlook. Verifique se ele está instalado.")
            print(f"Detalhes do erro: {e}")
            raise

def enviar_email_outlook(app_outlook, destinatarios: str, assunto: str, corpo_html: str, enviar: bool):
    """Cria e envia (ou salva) o e-mail através do Outlook."""
    email_obj = app_outlook.CreateItem(0)
    destinatarios = str(destinatarios).replace(";", ",")
    lista_de_emails = [e.strip() for e in destinatarios.split(",") if e.strip()]

    if not lista_de_emails:
        return

    email_obj.Subject = assunto
    email_obj.HTMLBody = corpo_html
    email_obj.To = lista_de_emails[0]
    if len(lista_de_emails) > 1:
        email_obj.CC = "; ".join(lista_de_emails[1:])

    if enviar:
        email_obj.Send()
    else:
        email_obj.Save()


def carregar_dados_lancamentos(caminho_arquivo):
    """Lê o Excel de lançamentos e mapeia as colunas conforme a estrutura da planilha."""
    df_bruto = pd.read_excel(caminho_arquivo, header=None)
    df_bruto = df_bruto.dropna(how="all")

    linhas_processadas = []

    for _, linha in df_bruto.iterrows():
        # Captura as colunas por índice físico do Excel
        nome_empresa = linha.iloc[0] if len(linha) > 0 else None
        nome_funcionario = linha.iloc[1] if len(linha) > 1 else None
        data_lanc = linha.iloc[2] if len(linha) > 2 else None
        num_recibo = linha.iloc[3] if len(linha) > 3 else None
        tratamento = linha.iloc[4] if len(linha) > 4 else None
        valor_total = linha.iloc[5] if len(linha) > 5 else None

        texto_validacao = str(nome_empresa).strip().lower() if pd.notna(nome_empresa) else ""

        # Ignora cabeçalhos
        if not texto_validacao or "relatório" in texto_validacao or texto_validacao == "empresa":
            continue

        # Validação mínima: precisa ter empresa e funcionário
        if pd.notna(nome_empresa) and pd.notna(nome_funcionario):
            linhas_processadas.append({
                "empresa": str(nome_empresa).strip(),
                "funcionario": str(nome_funcionario).strip(),
                "data": data_lanc,
                "recibo": str(num_recibo).strip() if pd.notna(num_recibo) else "",
                "tratamento": str(tratamento).strip() if pd.notna(tratamento) else "",
                "valor": valor_total,
            })

    return pd.DataFrame(linhas_processadas)


def construir_corpo_html(mes_referencia: str, dados_empresa: pd.DataFrame) -> str:
    """Gera o corpo do e-mail em HTML com a tabela de lançamentos da empresa."""
    linhas_tabela = []

    for _, linha in dados_empresa.iterrows():
        # Organiza os dados para as células da tabela
        valores_celulas = [
            linha.get("empresa"),
            linha.get("funcionario"),
            formatar_data(linha.get("data")),
            linha.get("recibo"),
            linha.get("tratamento"),
            formatar_valor_moeda(linha.get("valor"))
        ]

        html_celulas = ""
        for i, conteudo in enumerate(valores_celulas):
            alinhamento = "left"
            if i in [2, 3]: alinhamento = "center"
            if i == 5: alinhamento = "right"

            html_celulas += f'<td style="border:1px solid #000; padding:6px; font-size:11pt; text-align:{alinhamento};">{conteudo}</td>'

        linhas_tabela.append(f"<tr>{html_celulas}</tr>")

    tabela_html = f"""
    <table style="border-collapse:collapse; margin-top:10px;">
      <tbody>
        {''.join(linhas_tabela)}
      </tbody>
    </table>
    """

    corpo_final = f"""
    <div style="font-family:Calibri, Arial, sans-serif; font-size:11pt;">
      <p>Bom dia!</p>
      <p>Segue fechamento referente a {mes_referencia}:</p>
      {tabela_html}
      <p style="margin-top:16px;">Att,<br/>{ASSINATURA_NOME}</p>
    </div>
    """
    return corpo_final


def main():
    """Função principal para execução do fechamento."""
    if not ARQUIVO_LANCAMENTOS.exists():
        raise FileNotFoundError(f"Erro: {ARQUIVO_LANCAMENTOS} não encontrado.")
    if not ARQUIVO_ASSOCIADOS.exists():
        raise FileNotFoundError(f"Erro: {ARQUIVO_ASSOCIADOS} não encontrado.")

    df_lancamentos = carregar_dados_lancamentos(ARQUIVO_LANCAMENTOS)
    df_associados = pd.read_excel(ARQUIVO_ASSOCIADOS, sheet_name="Associados")

    if df_lancamentos.empty:
        print("Nenhum dado válido para processar.")
        return

    mapa_de_emails = mapear_emails_dos_associados(df_associados)
    mes_atual = datetime.now().strftime("%m/%Y")
    app_outlook = obter_aplicativo_outlook()

    sem_email = []
    enviados = 0

    for nome_empresa, grupo in df_lancamentos.groupby("empresa"):
        email_destino = mapa_de_emails.get(normalizar_nome_empresa(nome_empresa))

        if not email_destino:
            sem_email.append(nome_empresa)
            continue

        corpo = construir_corpo_html(mes_atual, grupo)
        enviar_email_outlook(app_outlook, email_destino, ASSUNTO_EMAIL, corpo, enviar=ENVIAR_DE_VERDADE)

        enviados += 1
        print(f"[OK] {nome_empresa} -> {email_destino}")

    print(f"\nConcluído! Total enviados/salvos: {enviados}")
    if sem_email:
        print("Empresas sem e-mail encontrado:", ", ".join(sem_email))


if __name__ == "__main__":
    main()