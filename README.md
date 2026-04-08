# 🚀 Automação de Fechamento Mensal - ACIP

Este projeto automatiza o processo de fechamento financeiro mensal da ACIP, realizando o cruzamento de dados entre relatórios de utilização e o cadastro de associados para o envio automatizado de extratos via e-mail.

## 📋 Como o Projeto Funciona

O sistema utiliza duas fontes de dados principais em formato Excel:

1.  **`ACIP.xlsx` (Cadastro)**: Funciona como o banco de dados de associados, contendo nomes fantasias, razões sociais e e-mails de cobrança.
2.  **`lancamento.xlsx` (Utilização)**: Contém os registros detalhados de uso do mês, incluindo funcionário, data, número do recibo e valor.

O script Python lê esses arquivos, agrupa todos os gastos de uma mesma empresa em uma única tabela e gera um e-mail formatado profissionalmente.

## ⚠️ Requisito Crítico: Outlook Clássico

Esta automação foi desenvolvida utilizando a biblioteca `pywin32` para integração com o ecossistema Windows. **O funcionamento é exclusivo para o Outlook Clássico (Desktop)**. 
* O software deve estar instalado e com uma conta de e-mail configurada e ativa no computador onde o script será executado.
* È recomendado abrir e deixar o outlook aberto para rodar o código.

## 🛠️ Tecnologias Utilizadas

* **Python 3.13**: Linguagem base do projeto.
* **Pandas**: Para manipulação, limpeza e cruzamento de grandes volumes de dados.
* **PyWin32**: Para a comunicação direta e automação do Microsoft Outlook.
* **Openpyxl**: Para leitura e escrita de arquivos Excel (.xlsx).

## ⚙️ Configuração e Instalação

1.  **Instale as dependências**:
    ```bash
    pip install pandas openpyxl pywin32
    ```

2.  **Prepare as Planilhas**:
    Certifique-se de que os arquivos `ACIP.xlsx` e `lancamento.xlsx` estejam na mesma pasta do script `fechamento.py`.

3.  **Modos de Envio**:
    No topo do arquivo `fechamento.py`, você encontrará a variável `ENVIAR_DE_VERDADE`:
    * `False` (Padrão): O script gera os e-mails e os salva na pasta **Rascunhos** do Outlook para revisão.
    * `True`: O script realiza o envio imediato após o processamento.

## 🛡️ Tratamento de Erros e Robustez

O código foi "blindado" para lidar com cenários reais de planilhas mal formatadas:
* **Normalização**: Os nomes das empresas são padronizados para garantir o vínculo com o e-mail, mesmo com diferenças de espaços ou letras maiúsculas/minúsculas.
* **Segurança de Dados**: O script valida a existência de dados em cada coluna. Caso encontre uma célula vazia ou erro de data, ele pula a linha e avisa no terminal, evitando a interrupção da automação.
* **Valores Financeiros**: Trata automaticamente a conversão de moedas e formatos de texto para garantir que a tabela final no e-mail esteja correta.

---
*Desenvolvido por William Santos Trettin como parte de soluções de automação empresarial.*
