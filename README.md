
## Descrição do Projeto

Este projeto tem como objetivo realizar a comparação e validação de dados entre duas planilhas: uma extraída do sistema WMS (Warehouse Management System) e outra do SAP. A principal análise envolve verificar a validade de lotes de produtos com base nas datas de vencimento e gerar relatórios sobre inconsistências encontradas entre os dois sistemas.

## Funcionalidades

- **Leitura das Planilhas**: Carrega dados das planilhas do SAP e WMS.
- **Renomeação das Colunas**: As colunas das planilhas são renomeadas para nomes mais legíveis e compreensíveis.
- **Filtragem e Preparo dos Dados**: Seleção de colunas relevantes e transformação de dados (como datas e IDs únicos).
- **Junção de Dados**: Realiza o `merge` das duas planilhas com base em um identificador único de produto e lote.
- **Comparação de Lotes**: Verifica a consistência entre as datas de validade dos produtos nos sistemas SAP e WMS.
- **Ação de Correção**: Define ações de correção para os lotes que apresentarem inconsistência (por exemplo, "Validar Físico").
- **Exportação dos Resultados**: Gera um novo arquivo Excel contendo o relatório final com as inconsistências e ações recomendadas.

## Como Usar

1. **Instalação das Dependências**:  
   O projeto depende das bibliotecas `pandas` e `openpyxl` para leitura e manipulação de arquivos Excel. Caso ainda não as tenha instalado, execute:
   
   ```bash
   pip install pandas openpyxl xlrd
   ```

2. **Carregue os Arquivos**:  
   As planilhas do SAP e WMS devem ser localizadas nos caminhos definidos no código. Certifique-se de ajustar os caminhos para os seus arquivos.

3. **Executar o Script**:  
   Após ajustar os caminhos, execute o script Python para gerar o relatório final. O arquivo Excel com o resultado será salvo no local especificado.

4. **Analisar os Resultados**:  
   O arquivo gerado conterá as informações sobre os lotes verificados e suas respectivas ações (como "Lote OK" ou "Validar Físico" para inconsistências).

## Exemplo de Uso

```python
import pandas as pd

# Carregar as planilhas
wms_df = pd.read_excel('path_to_wms.xlsx')
sap_df = pd.read_excel('path_to_sap.xlsx')

# Realizar as operações de comparação e análise...
```

## Contribuições

Se você deseja contribuir com melhorias ou novas funcionalidades, sinta-se à vontade para abrir um `pull request`.

---
