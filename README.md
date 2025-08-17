# Gerenciador de Produtos em Planilha Excel

Este é um script em Python de linha de comando desenvolvido para facilitar a adição de novos produtos a uma planilha Excel já existente, garantindo a organização e a integridade dos dados.

## Funcionalidades

- **Leitura de Dados Existentes:** O script lê uma planilha Excel (`produtos.xls`) para carregar o catálogo de produtos atual.
- **Coleta Interativa de Dados:** Um sistema de entrada de dados simples e interativo guia o usuário para adicionar informações sobre novos produtos.
- **Dados Padronizados:** Atributos como "Estoque", "Situação" e "Categoria" são automaticamente preenchidos com valores padrão, simplificando o processo.
- **Geração de Backup Automática:** Em vez de sobrescrever o arquivo original, o script salva a nova planilha com um carimbo de data e hora (`YYYYMMDD_HHMMSS`), criando um backup automático e rastreável.

## Requisitos

Certifique-se de ter as seguintes dependências instaladas em seu ambiente Python:

- **Python 3.6+**
- `pandas`
- `openpyxl` (necessário para ler e escrever arquivos .xlsx)

Para instalar as bibliotecas, execute o seguinte comando no seu terminal:

```bash
pip install pandas openpyxl
```
