# 🚀 Automação de Fechamento Mensal - ACIP

Este projeto automatiza o envio de extratos de utilização mensal para os associados da ACIP, integrando dados de planilhas Excel diretamente com o Outlook(CLASSICO).

## 📋 Como funciona?
O sistema utiliza duas planilhas principais:
1. **ACIP.xlsx**: Contém o cadastro completo de associados (Nome Fantasia, Razão Social e E-mail de cobrança).
2. **lancamento.xlsx**: Contém os registros de uso do mês (Empresa, Funcionário, Data, Recibo e Valor).

O código cruza essas informações, agrupa todos os gastos de uma mesma empresa e gera um e-mail formatado.

## ✉️ Modos de Envio
O script possui uma trava de segurança no topo do arquivo (`ENVIAR_DE_VERDADE`):
* **Modo Conferência (False)**: Os e-mails são gerados e salvos na pasta **Rascunhos** do Outlook. Ideal para revisão antes do envio real.
* **Modo Produção (True)**: Os e-mails são enviados automaticamente para os destinatários.

## 🛠️ Tecnologias Utilizadas
* Python 3.13
* Pandas (Processamento de dados)
* PyWin32 (Integração com Outlook Clássico, não funciona com o novo Outlook)

