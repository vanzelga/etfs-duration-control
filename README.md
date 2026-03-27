# ETFs Duration Control

Automação com front em Excel para calcular o duration médio ponderado por fundo, usando:

- posição do Metabase
- carteira IMA-B da ANBIMA
- substituição do Duration por PMR para NTN-Bs elegíveis
- validação de desenquadramento
- snapshots diários das bases consumidas

---

# 1. Objetivo

O projeto foi criado para entregar, de forma simples e operacional:

1. leitura da posição da carteira na data informada
2. leitura da carteira IMA-B da mesma data
3. substituição do Duration por PMR em NTN-Bs elegíveis
4. cálculo do duration ponderado por fundo
5. identificação de desenquadramento
6. entrega do resultado em um Excel de front (.xlsm)

---

# 2. Regra de negócio

## 2.1. Base Meta

A base do Metabase traz a posição dos ativos por fundo na data escolhida.

Campos mais importantes:
- CgePortfolio
- CodTipoAgrupamento
- VlPosicao
- Duration
- NuIsin

## 2.2. Base ANBIMA

A base da ANBIMA traz a carteira do índice IMA-B da data escolhida.

Campos mais importantes:
- C_Cod_ISIN
- C_PMR
- C_Duration

## 2.3. Regra de substituição

Para ativos elegíveis:

- se CodTipoAgrupamento = NTNB
- e Duration <> 1

então:

- o Duration original do Meta é substituído por PMR da ANBIMA

## 2.4. Regra de erro

Se existir uma NTNB elegível no Meta e o ISIN não for encontrado na ANBIMA para a data informada:

- o processo deve ser tratado como erro de negócio
- isso evita entregar cálculo inconsistente

## 2.5. Cálculo final

O duration ponderado por fundo é calculado assim:

sum(VlPosicao * Duration_Final) / sum(VlPosicao)

Agrupamento:
- DataCarteira
- CgePortfolio

## 2.6. Desenquadramento

Atualmente, a regra considerada é:

- desenquadrado_720 = duration_ponderado < 720

---

# 3. Visão geral do fluxo

## Fluxo operacional

1. usuário informa a data no Excel
2. usuário clica no botão
3. o VBA chama o Python
4. o Python busca:
   - Meta
   - ANBIMA
5. o Python salva snapshots
6. o Python calcula o resultado final
7. o VBA abre o arquivo gerado
8. o VBA copia os dados para o .xlsm
9. o usuário analisa o resumo na aba controle

---

# 4. Estrutura do projeto

etfs_duration/  
├── main.py                     — Orquestra a execução completa  
├── config.py                   — Centraliza configurações do projeto  
├── README.md                   — Documentação do projeto  
├── requirements.txt            — Dependências Python  
├── .gitignore                  — Regras de arquivos ignorados no Git  
├── src/  
│   ├── __init__.py             — Marca `src` como pacote Python  
│   ├── anbima_client.py        — Consulta e trata base da ANBIMA  
│   ├── metabase_client.py      — Consulta e trata base do Metabase  
│   ├── calculator.py           — Aplica regras e calcula duration  
│   ├── excel_exporter.py       — Gera o Excel final  
│   └── utils.py                — Funções auxiliares  
└── vba/  
    ├── modMain.bas             — Macro principal do front Excel  
    └── modButtons.bas          — Rotinas auxiliares de botões  

## Estrutura de saída em disco

C:\composicao\  
├── debug\                      — Arquivos de debug  
├── meta_raw\                   — Snapshot bruto do Metabase  
├── meta_parsed\                — Snapshot tratado do Metabase  
├── anbima_raw\                 — Snapshot bruto da ANBIMA  
├── anbima_parsed\              — Snapshot tratado da ANBIMA  
└── duration_final_YYYYMMDD.xlsx — Arquivo final do cálculo  

---

# 5. Responsabilidade de cada arquivo

## main.py

Arquivo principal do projeto.

Responsável por:
- receber a data de execução
- chamar Meta
- chamar ANBIMA
- salvar snapshots
- executar cálculo
- gerar Excel final

## src/metabase_client.py

Responsável por:
- montar o request para o Metabase
- buscar a posição da data informada
- devolver:
  - DataFrame tratado
  - resposta bruta
- salvar snapshots do Meta

Saídas geradas:
- meta_raw/meta_YYYYMMDD.json
- meta_parsed/meta_YYYYMMDD.xlsx

## src/anbima_client.py

Responsável por:
- montar o request para a ANBIMA
- aquecer a sessão
- baixar a carteira IMA-B
- parsear a resposta texto/XML-like
- devolver:
  - DataFrame tratado
  - resposta bruta
- salvar snapshots da ANBIMA

Saídas geradas:
- anbima_raw/anbima_YYYYMMDD.txt
- anbima_parsed/anbima_YYYYMMDD.xlsx

## src/calculator.py

Responsável por:
- aplicar a regra de negócio
- identificar NTN-B elegível
- substituir Duration por PMR
- construir detalhe dos ativos
- gerar tabela final por fundo
- marcar desenquadramento

Saídas lógicas:
- detalhe_ativos
- resultado_fundos
- erros

## src/excel_exporter.py

Responsável por:
- gerar o arquivo final .xlsx
- escrever as abas:
  - base_meta
  - base_anbima
  - detalhe_ativos
  - resultado_fundos
  - erros

## src/utils.py

Responsável por helpers simples:
- criação de diretórios
- parse de data
- conversão de número BR
- formatação de datas

## vba/modMain.bas

Responsável por:
- ler a data da aba controle
- chamar python main.py
- abrir o arquivo resultado
- copiar dados para o front Excel
- formatar abas
- atualizar resumo
- registrar log simples na aba log

---

# 6. Como rodar manualmente

## Instalar dependências

pip install -r requirements.txt

## Rodar por terminal

python main.py 2026-03-25

Formato da data:
- YYYY-MM-DD

Exemplo:
- python main.py 2026-03-25

Se nenhuma data for informada:
- o script usa o dia anterior como fallback

---

# 7. Como rodar pelo Excel

## Arquivo de front

O front é um arquivo .xlsm.

## Abas esperadas no .xlsm

O Excel precisa ter estas abas com esses nomes exatos:

- controle
- resultado_fundos
- detalhe_ativos
- erros
- log

## Entrada

Na aba controle:
- célula B2 = data da carteira

## Ação

O botão chama a macro:
- RodarETFsDuration

## O que a macro faz

1. lê a data de B2
2. executa o Python
3. espera o processo terminar
4. abre duration_final_YYYYMMDD.xlsx
5. copia os dados para o .xlsm
6. formata as abas
7. escreve resumo na aba controle
8. registra linha na aba log

---

# 8. Saídas geradas

## 8.1. Arquivo resultado final

Arquivo:
- C:\composicao\duration_final_YYYYMMDD.xlsx

Abas:
- base_meta
- base_anbima
- detalhe_ativos
- resultado_fundos
- erros

## 8.2. Snapshots

### Meta
- bruto: meta_raw
- tratado: meta_parsed

### ANBIMA
- bruto: anbima_raw
- tratado: anbima_parsed

## 8.3. Debug

A pasta debug guarda respostas auxiliares da ANBIMA para troubleshooting.

---

# 9. Resumo da aba controle

A macro escreve um resumo executivo em D2:E8, com:

- última execução
- data da carteira
- status
- quantidade de fundos
- quantidade de desenquadrados
- quantidade de erros
- caminho do arquivo gerado

Isso foi pensado para o usuário bater o olho e entender o status sem abrir outras abas.

---

# 10. Principais pontos de manutenção

## 10.1. Se o Metabase mudar

Verificar:
- URL do card público
- id do parâmetro DataCarteira
- nomes das colunas retornadas

Arquivo afetado:
- src/metabase_client.py

## 10.2. Se a ANBIMA mudar

Verificar:
- payload
- comportamento do Dt_Ref
- estrutura da resposta
- nomes dos campos:
  - C_Cod_ISIN
  - C_PMR
  - C_Duration

Arquivo afetado:
- src/anbima_client.py

## 10.3. Se a regra de negócio mudar

Exemplos:
- trocar critério de elegibilidade
- trocar PMR por outro campo
- mudar regra de desenquadramento
- mudar agrupamento final

Arquivo afetado:
- src/calculator.py

## 10.4. Se a saída Excel mudar

Exemplos:
- novas abas
- nova ordem de colunas
- novas regras de formatação

Arquivos afetados:
- src/excel_exporter.py
- vba/modMain.bas

---

# 11. Regras importantes do sistema

## 11.1. Não confiar só no arquivo final

O projeto salva snapshots justamente para permitir:
- auditoria
- reprocessamento
- troubleshooting

## 11.2. Não mudar nome das abas do front sem ajustar VBA

Se o nome da aba mudar no .xlsm, a macro pode quebrar.

## 11.3. Não mudar nome das colunas sem revisar cálculo

O projeto depende de nomes de colunas específicos para:
- merge
- regra de substituição
- agrupamento
- destaque visual

---

# 12. Troubleshooting rápido

## Erro: Subscript out of range

Normalmente significa:
- aba não encontrada no .xlsm
- aba não encontrada no .xlsx gerado

Verificar:
- nomes das abas

## Erro de import no Python

Exemplo:
- função não encontrada
- módulo não encontrado

Verificar:
- se o conteúdo dos arquivos em src/ está atualizado
- se existe src/__init__.py

## Arquivo final não encontrado

Verificar:
- se o Python rodou até o fim
- se houve erro de negócio
- se existe arquivo antigo sendo usado indevidamente
- pasta C:\composicao

## ANBIMA sem retorno válido

Verificar:
- pasta debug
- snapshots anbima_raw
- possíveis mudanças no site/payload

## Resultado parece estranho

Verificar:
1. snapshot do Meta
2. snapshot da ANBIMA
3. merge por ISIN
4. regra de elegibilidade
5. cálculo do ponderado

---

# 13. Próximas evoluções possíveis

Itens que podem ser adicionados depois do MVP:

- log em arquivo .log
- botão para limpar log
- botão para abrir pasta de saída
- validação visual mais forte no front
- tratamento de múltiplas datas
- histórico consolidado
- versionamento/export dos módulos VBA
- parametrização por arquivo config.py

---

# 14. Resumo executivo

Em uma frase:

Este projeto automatiza a leitura de posições do Metabase e da carteira IMA-B da ANBIMA, aplica a regra de substituição de Duration por PMR para NTN-Bs elegíveis, calcula o duration ponderado por fundo e entrega o resultado em um front Excel com snapshots diários para rastreabilidade.
