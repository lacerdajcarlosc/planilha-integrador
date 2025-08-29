# 📊 Processador de Planilhas

Este script em Python automatiza o processamento de planilhas de prestação de serviço (PJES), consolidando dados, calculando valores, classificando verbas e gerando tabelas dinâmicas em um único arquivo Excel.

## 🧩 Funcionalidades

- Leitura de múltiplos arquivos `.xlsx` de uma pasta específica
- Tratamento e padronização de dados
- Cálculo de valores com base no cargo e quantidade de cotas
- Classificação de verbas conforme a data de término
- Consolidação por matrícula
- Geração de tabelas dinâmicas:
  - Por local de prestação de serviço
  - Por tipo de verba
  - Exclusiva para verba 223
- Exportação para um arquivo Excel com múltiplas abas
- Registro de erros em um log separado

## 📁 Estrutura Esperada

- Os arquivos devem estar na pasta definida pela variável `pasta`, por padrão:

- Cada arquivo deve conter uma aba chamada **"Planilha de PJES"** com os dados iniciando na linha 10 (por isso o `skiprows=9`).

## 🧮 Regras de Negócio

- **Cargos com cota de R$300**:
- Oficial da PM, Tenente, Capitão, Major, Tenente Coronel, Coronel, Delegado de Polícia, Perito Criminal, Médico Legista
- **Demais cargos**: cota de R$200
- **Verba**:
- `223` → se a data de término for no mesmo mês/ano da competência
- `423` → se a data de término for anterior à competência
- `""` → caso contrário ou data ausente

## 📦 Requisitos

- Python 3.8+
- Bibliotecas:
- `pandas`
- `openpyxl`
- `locale`
- `datetime`
- `os`

Instale com:
```bash
pip install pandas openpyxl

🚀 Como Executar
Atualize o caminho da pasta na variável pasta se necessário.

Execute o script:

bash
python processador_pjes.py
O script irá:

Processar todos os arquivos .xlsx da pasta

Gerar um arquivo consolidado chamado TABELA_BASE_YYYYMMDD_HHMMSS.xlsx

Criar abas com os seguintes nomes:

AGOSTO (base consolidada)

RESUMO POR LOCAL

VERBA 223

TOTAL POR VERBA

🛠️ Log de Erros
Caso algum arquivo não possa ser processado, o erro será registrado em um arquivo log_erros_YYYYMMDD_HHMMSS.txt na mesma pasta.

Se nenhum erro for encontrado, o log será automaticamente excluído.

📌 Observações
O script utiliza a localidade pt_BR.UTF-8 para formatar nomes de meses em português.

Certifique-se de que os arquivos estejam corretamente nomeados e estruturados conforme o padrão esperado.

