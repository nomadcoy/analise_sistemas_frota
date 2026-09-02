# Comparativo e Auditoria Integrada de Frotas (SGTM | SGTP | SIMETRO)

Este projeto consiste em um script em Python desenvolvido para automatizar o cruzamento, higienização e auditoria de bases de dados de frotas de transporte urbano provenientes de três sistemas distintos: **SGTM**, **SGTP** e **SIMETRO**.

O objetivo principal é identificar a **frota operacional comum e ativa**, catalogar divergências de cadastros e analisar o histórico legado de veículos.

---

## Funcionalidades

- **Higienização de Dados Automatizada**: Tratamento de caracteres especiais (pontos, hífens, espaços), padronização em caixa alta (`uppercase`) e deduplicação de chaves primárias (placas).
- **Leitura Resiliente e Multi-formato**: Suporte a diferentes encodings (`utf-8`, `latin1`), delimitadores (vírgula e ponto e vírgula) e tratamento de arquivos sem extensão ou com cabeçalhos secundários (`skiprows`).
- **Análise de Conjuntos (Set Theory)**: Cruzamento ágil utilizando operações matemáticas de conjuntos (`intersection`, `difference`) para mapear:
  - Frota ativa e convergente nos 3 sistemas.
  - Veículos exclusivos de cada base.
  - Interseções parciais entre sistemas.
- **Estudo Dedicado de Frota Legada**: Diagnóstico específico dos veículos exclusivos do SGTM, classificando-os entre ativos sem baixa cadastrada e veículos baixados com histórico de saída.
- **Geração de Reports Automatizados**:
  - **Excel (`.xlsx`)**: Planilha estruturada com múltiplas abas (`Resumo_Geral`, `Presente_nos_3`, `Estudo_Exclusivas_SGTM`, etc.).
  - **Relatório TXT (`.txt`)**: Documento procedimental detalhado com métricas, logs e estatísticas do processamento.

---

## Tecnologias Utilizadas

- **Python 3.x**
- **Pandas**: Manipulação, limpeza e cruzamento de DataFrames.
- **OpenPyXL**: Engine para exportação formatada em arquivos Excel.
- **OS & DateTime**: Gerenciamento de caminhos de arquivos e carimbos de data/hora.

---

## Estrutura dos Arquivos Gerados

Ao executar o script, a seguinte estrutura de relatórios é consolidada:

```text
├── relatorio_comparativo_frotas.xlsx
│   ├── Resumo_Geral              # Visão executiva das métricas
│   ├── Presente_nos_3            # Base operacional convergente (5.833 placas)
│   ├── Estudo_Exclusivas_SGTM    # Análise das 1.320 placas exclusivas
│   ├── Apenas_SGTP               # Divergências pontuais SGTP
│   └── Apenas_SIMETRO            # Divergências pontuais SIMETRO
└── relatorio_procedimento_frotas.txt # Log procedural detalhado