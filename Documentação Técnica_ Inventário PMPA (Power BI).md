# Documentação Técnica: Inventário de Máquinas PMPA (Power BI)

## 1. Objetivo

Cruzar a relação patrimonial de notebooks/máquinas com o inventário de hardware coletado nas estações Windows, gerando uma tabela única por bem patrimonial, com os dados técnicos da máquina quando ela foi localizada no inventário.

## 2. Arquitetura

```
Inventário_HW_Windows_PMPA_AAAAMMDD.xlsx ──► PMPA ──────────────┐
                                                                ▼
Inventario Notebook PMPA.xlsx ──► CBP-Relação Mobiliário ──► Merge (PK) ──► Tabela final
CBP-Relação Mobiliário (2) ────────► Append
```

| Consulta | Fonte | Aba | Papel |
| --- | --- | --- | --- |
| PMPA | Inventário_HW_Windows_PMPA_20260828.xlsx | PMPA | Dados de hardware coletados |
| CBP-Relação Mobiliário | Inventario Notebook PMPA.xlsx | CBP-Relação Mobiliário | Base patrimonial (tabela principal) |
| CBP-Relação Mobiliário (2) | Consulta auxiliar (não enviada) | - | Linhas adicionais anexadas à base patrimonial |

Pasta de origem: `C:\Users\leonardo.meneghini\Documents\Inventario Maquinas\fontes\`

## 3. Chave de relacionamento

- **PK** em ambas as consultas.
- No patrimônio: cópia de `SÉRIE`.
- No inventário: cópia de `Bios Serial`.
- Tipo de junção: **Left Outer** (patrimônio à esquerda).

## 4. Consulta PMPA: etapas

| # | Etapa | Função M | Descrição |
| --- | --- | --- | --- |
| 1 | Source | `Excel.Workbook` | Abre o arquivo de inventário |
| 2 | PMPA_Sheet | Navegação | Seleciona a aba PMPA |
| 3 | Promoted Headers | `Table.PromoteHeaders` | Primeira linha vira cabeçalho |
| 4 | Duplicated Column | `Table.DuplicateColumn` | Copia `Bios Serial` |
| 5 | Reordered Columns | `Table.ReorderColumns` | Coloca a cópia na primeira posição |
| 6 | Renamed Columns | `Table.RenameColumns` | Renomeia a cópia para `PK` |

## 5. Consulta CBP-Relação Mobiliário: etapas

| # | Etapa | Função M | Descrição |
| --- | --- | --- | --- |
| 1 | Source / Tabela1 | `Excel.Workbook` | Abre o arquivo e seleciona a aba |
| 2 | Tabela1ComCabecalho | `Table.PromoteHeaders` | Promove cabeçalho |
| 3 | Appended Query | `Table.Combine` | Une com a consulta (2) |
| 4 | Duplicated/Renamed Columns | `Table.DuplicateColumn` | Cria `SÉRIE_c` (cópia de `SÉRIE`) |
| 5 | Changed Type | `Table.TransformColumnTypes` | `DATA AQUISIÇÃO` para `date` |
| 6 | Duplicated/Renamed Columns1 | `Table.DuplicateColumn` | Cria `DATA AQUISIÇÃO_copia` |
| 7 | Split Column by Delimiter | `Table.SplitColumn` | `UNIDADE` em `UNIDADE.1` e `UNIDADE.2` (delimitador "-") |
| 8 | Duplicated/Renamed Columns2 | `Table.DuplicateColumn` | Cria `PK` (cópia de `SÉRIE`) |
| 9 | Reordered Columns | `Table.ReorderColumns` | Define a ordem final |
| 10 | Merged Queries | `Table.NestedJoin` | Left Outer com PMPA por `PK` |
| 11 | Expanded | `Table.ExpandTableColumn` | Expande colunas com prefixo `PMPA.` |

## 6. Dicionário de dados

### 6.1 Colunas patrimoniais

| Coluna | Descrição | Observação |
| --- | --- | --- |
| PK | Chave de junção (cópia de SÉRIE) | Texto |
| CONTASUBCONTA | Conta contábil / subconta |  |
| UNIDADE.1 | Parte 1 da unidade (antes do hífen) | Normalmente o código |
| UNIDADE.2 | Parte 2 da unidade (depois do hífen) | Normalmente a descrição |
| SUBUNIDADE | Subunidade |  |
| LOTAÇÃO | Local de lotação do bem |  |
| NÚMERO | Número do patrimônio |  |
| DESCRIÇÃO | Descrição do bem |  |
| DENOMINAÇÃO | Denominação do bem |  |
| MARCA / MODELO | Fabricante e modelo |  |
| SÉRIE | Número de série | Origem da PK |
| DATA AQUISIÇÃO | Data de aquisição | Tipo date |
| DATA LANÇAMENTO | Data de lançamento contábil | Sem tipagem |
| DATA GARANTIA | Fim da garantia | Sem tipagem |
| EMPENHO / PROCESSO | Empenho e processo de compra |  |
| VALOR / VALOR ATUAL | Valor de aquisição e valor atual | Sem tipagem |
| CPF/CNPJ / FORNECEDOR | Dados do fornecedor |  |
| NRFID | Identificador RFID |  |
| FOTOBEM / FOTOETQT | Referências de foto do bem e da etiqueta |  |
| SÉRIE_c | Cópia de segurança de SÉRIE |  |
| DATA AQUISIÇÃO_copia | Cópia de segurança de DATA AQUISIÇÃO |  |

### 6.2 Colunas de hardware (prefixo `PMPA.`)

| Coluna | Descrição |
| --- | --- |
| PMPA.PK | Chave vinda do inventário. **Nulo = máquina não inventariada** |
| PMPA.Nome do computador | Hostname |
| PMPA.IP | Endereço IP |
| PMPA.Rede | Rede/segmento |
| PMPA.Sistema Operacional | Versão do Windows |
| PMPA.CPU | Processador |
| PMPA.Bios Serial | Serial lido da BIOS |
| PMPA.Memória (Mbytes) | Memória total |
| PMPA.Informação da BIOS | Versão/dados da BIOS |
| PMPA.Memory.Bank #1 a #4 | Módulos de memória por slot |
| PMPA.Disco Primário / Secundário | Discos instalados |
| PMPA.Data do inventário | Data da coleta |

## 7. Regras de negócio

- Todo bem patrimonial permanece na saída (Left Outer).
- `PMPA.PK` nulo indica bem sem coleta de hardware, útil para medir cobertura do inventário.
- As colunas `SÉRIE_c` e `DATA AQUISIÇÃO_copia` preservam os valores originais.

## 8. Riscos e melhorias recomendadas

| # | Ponto | Risco | Recomendação |
| --- | --- | --- | --- |
| 1 | Caminhos fixos no perfil do usuário | Falha em outra máquina ou no Serviço Power BI | Usar parâmetro de pasta, SharePoint/rede e gateway |
| 2 | Nome do arquivo com data | Troca manual a cada novo inventário | Ler a pasta e escolher o arquivo mais recente |
| 3 | PK sem normalização | Falsos "não encontrados" (espaços, caixa) | `Text.Upper(Text.Trim([PK]))` nas duas consultas |
| 4 | PK duplicada no PMPA | Multiplica linhas do patrimônio | Manter a coleta mais recente por PK |
| 5 | Split de UNIDADE por todos os hífens | Perda de partes se houver mais de um hífen | Dividir só no primeiro hífen |
| 6 | Poucas colunas tipadas | Somas e filtros incorretos | Tipar VALOR, VALOR ATUAL, datas e memória (cultura pt-BR) |
| 7 | Append exige colunas idênticas | Colunas extras com nulos | Conferir cabeçalhos da consulta (2) |
| 8 | Nome das abas fixo | Erro de atualização se renomeadas | Manter nomes ou tratar com `try ... otherwise` |

### Trechos sugeridos

Normalizar a PK (aplicar antes do join, nas duas consultas):

```m
Table.TransformColumns(Fonte, {{"PK", each Text.Upper(Text.Trim(_)), type text}})
```

Manter o registro mais recente por PK no PMPA:

```m
let
    Ordenado = Table.Sort(#"Renamed Columns", {{"Data do inventário", Order.Descending}}),
    Unico = Table.Distinct(Ordenado, {"PK"})
in
    Unico
```

Dividir UNIDADE apenas no primeiro hífen:

```m
Table.SplitColumn(Fonte, "UNIDADE",
    Splitter.SplitTextByEachDelimiter({"-"}, QuoteStyle.None, false),
    {"UNIDADE.1", "UNIDADE.2"})
```

## 9. Procedimento de atualização

1. Salvar o novo arquivo de inventário de hardware na pasta `fontes`.
2. Ajustar o caminho na consulta PMPA (ou parâmetro), se o nome mudou.
3. Conferir se o cabeçalho das duas planilhas não mudou.
4. Em **Página Inicial > Atualizar** no Power BI Desktop.
5. Validar: contagem de linhas igual à do patrimônio e percentual de `PMPA.PK` nulo.
6. Publicar no Serviço Power BI, se aplicável (exige gateway para arquivos locais).

## 10. Testes de validação

- O total de linhas da tabela final deve ser igual ao total do patrimônio (após o append). Se for maior, há PK duplicada no PMPA.
- Contar `PMPA.PK = null` para obter a quantidade de máquinas não inventariadas.
- Verificar linhas com `UNIDADE.2` nulo, que indicam unidades sem hífen.