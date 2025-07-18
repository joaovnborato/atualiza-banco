# 📊 Sistema de Importação de Indicadores Operacionais via Excel e SQL Server

Este projeto em Python automatiza a importação de dados operacionais a partir de planilhas Excel diretamente para um banco de dados SQL Server (Azure).  
Ideal para contextos logísticos ou operacionais onde é necessário consolidar indicadores recorrentes como **Devolução**, **Refugo**, **Rating**, **Dispersão** e **Reposição**.

---

## ✅ Funcionalidades

- 🔄 Leitura automática de planilhas `.xlsx`
- ⚙️ Conexão com banco de dados SQL Server (pymssql)
- 🧠 Verificação de duplicidade por `ID_Motorista` e `Data`
- 📥 Inserção nas tabelas correspondentes com controle de erro
- 📊 Relatório de inserções, duplicações e falhas ao final da execução

---

## 🧾 Planilhas suportadas

O sistema espera os seguintes arquivos (na mesma pasta do script):

- `devolucao.xlsx`
- `refugo.xlsx`
- `rating.xlsx`
- `dispersao.xlsx`
- `reposicao.xlsx`

Cada arquivo deve conter ao menos as colunas:

- `ID` – Código do motorista  
- `Nome` – Nome abreviado  
- `Data` – Data da medição  
- `Valor` – Valor do indicador (porcentagem, nota, KM etc.)

> 🛠️ **Atenção**: Os **nomes das tabelas no banco** (`Devolucao`, `Refugo`, `Rating`, etc.) e os **nomes das colunas no código** devem ser **ajustados conforme a estrutura do seu banco de dados** e os campos reais nas suas planilhas.

---

## ⚙️ Tecnologias utilizadas

- Python 3.x
- Pandas
- OpenPyXL
- PyMSSQL
- SQL Server (Azure)

---

## 🚀 Como usar

1. Instale as dependências:

```bash
pip install pandas openpyxl pymssql

2. Configure sua string de conexão no arquivo:
conn = pymssql.connect(
    server='seu-servidor.database.windows.net',
    user='seu-usuario',
    password='sua-senha',
    database='seu-banco'
)

3. Coloque os arquivos .xlsx esperados na mesma pasta do script.

4. Execute o script.

## 📦 Resultado esperado
Ao final, será exibido um relatório como este:

-----------------------------------------RESUMO-------------------------------------------
Sistema encerrado às 15:42:18
TOTAL     >>>>> INSERIDOS = 128 - DUPLICADOS = 27 - ERROS = 2

DEVOLUÇÃO >>>>> 32 inseridos, 5 duplicados, e 0 erros.
REFUGO    >>>>> 21 inseridos, 2 duplicados, e 0 erros.
RATING    >>>>> 19 inseridos, 4 duplicados, e 1 erro.
DISPERSÃO >>>>> 28 inseridos, 9 duplicados, e 1 erro.
REPOSIÇÃO >>>>> 28 inseridos, 7 duplicados, e 0 erros.
