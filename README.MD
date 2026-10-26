# Extração de Inventário em PDFs → CSV (com conferência de totais)

Script em Python que lê relatórios de inventário patrimonial em PDF, extrai os dados de cada bem e gera um arquivo **CSV**. Ao final, **confere a soma dos valores extraídos com o "Total Geral"** impresso em cada relatório, apontando quais PDFs precisam de revisão manual.

> **Versão:** 2.0 · **Última atualização:** _26/10/2026_

---

## Sumário

- [Mudanças em relação à versão anterior](#mudanças-em-relação-à-versão-anterior)
- [O que o script faz](#o-que-o-script-faz)
- [Campos extraídos](#campos-extraídos)
- [Pré-requisitos e instalação](#pré-requisitos-e-instalação)
- [Configuração](#configuração)
- [Como usar](#como-usar)
- [Como funciona](#como-funciona)
- [Conferência com o Total Geral](#conferência-com-o-total-geral)
- [Estrutura da saída](#estrutura-da-saída)
- [Ajustando o padrão de leitura (regex)](#ajustando-o-padrão-de-leitura-regex)
- [Solução de problemas](#solução-de-problemas)
- [Limitações conhecidas](#limitações-conhecidas)
- [Roadmap](#roadmap)
- [Histórico de alterações](#histórico-de-alterações)

---

## Mudanças em relação à versão anterior

| Aspecto | Versão 1 | Versão 2 |
|---|---|---|
| Leitura do PDF | `pdfplumber` + OCR (Tesseract) para PDFs em imagem | `pypdf` apenas (**sem OCR**) |
| Estratégia de extração | Um regex por campo, procurando rótulos (`Patrimônio:`, `Status Bem:`…) | **Um único regex por linha**, baseado na posição das colunas do relatório |
| Formato de saída | Excel `.xlsx` formatado (`pandas` + `openpyxl`) | **CSV** com separador `;` e codificação `utf-8-sig` (abre direto no Excel) |
| Campos | Item (texto), Patrimônio, Tombo, Localização, Status Bem, Sit. Física | **Item (número)**, Patrimônio, Localização, Status Bem, Sit. Física e **Vl. Aquisição**; coluna **Página** adicionada; **Tombo removido** |
| Validação | Contagem de valores vazios por coluna | **Soma dos valores × "Total Geral"** de cada PDF, com motivo detalhado |
| Dependências | pandas, pdfplumber, pymupdf, pytesseract, pillow, openpyxl | **pypdf** (o resto é biblioteca padrão) |
| Configuração | Várias variáveis no topo | Apenas `pastaEntrada` e `pastaSaida` |
| Subpastas | Opcional (`INCLUIR_SUBPASTAS`) | Não suportado (só a pasta informada) |

## O que o script faz

1. Lista todos os `*.pdf` da pasta de entrada, em ordem alfabética.
2. Lê o texto de cada página com `pypdf`.
3. Procura, linha a linha, as linhas que iniciam um bem (padrão com número do item, patrimônio de 8 dígitos, localização, status, situação física e valor).
4. Captura também o **"Total Geral"** impresso no relatório.
5. Grava todos os itens em um único CSV.
6. Compara a soma dos valores extraídos com o Total Geral e imprime um **resumo final** com os PDFs que precisam de conferência.

Um PDF com erro **não interrompe** os demais; ele é registrado no resumo final.

## Campos extraídos

| Coluna | Descrição | Observações |
|---|---|---|
| `Arquivo` | Nome do PDF de origem | |
| `Página` | Página do PDF onde o item foi encontrado | Útil para localizar o item no documento |
| `Item` | Número sequencial do item | Gravado como inteiro |
| `Patrimônio` | Número de patrimônio | Exatamente **8 dígitos**; gravado como texto (preserva zeros à esquerda) |
| `Localização` | Local do bem | Formato `NNNNN − NNNNNNNN − descrição`; espaços repetidos são normalizados |
| `Status Bem` | Situação cadastral | Uma palavra (sem espaços) |
| `Sit. Física` | Situação física | Uma palavra (sem espaços) |
| `Vl. Aquisição` | Valor de aquisição | Formato brasileiro (`1.180,00`) |

> A descrição do bem é reconhecida pelo padrão, mas **não é gravada** no CSV.

## Pré-requisitos e instalação

- **Python** 3.x
- Biblioteca **pypdf**:

```bash
pip install pypdf
```

Não é necessário instalar o Tesseract nem outras bibliotecas.

## Configuração

No topo do script, as duas linhas de caminho vêm **comentadas** e precisam ser ativadas:

```python
pastaEntrada = r"C:\caminho\da\pasta\com\os\pdfs"
pastaSaida   = r"C:\caminho\da\pasta\de\saida"
```

1. Substitua os textos de exemplo pelos caminhos reais.
2. **Remova o `#` do início das duas linhas.** Sem isso, o script falha com `NameError`.

O arquivo gerado se chama `relatorio_saida.csv` e a pasta de saída é criada automaticamente, se não existir.

## Como usar

```bash
python nome_do_script.py
```

Durante a execução, cada PDF gera uma linha de status:

```
relatorio_01.pdf: 42 itens | soma 51.380,00 | total do relatório 51.380,00 -> OK
```

Ao final aparece o resumo com o total de PDFs, itens salvos e a lista de arquivos a conferir.

## Como funciona

```
pasta de PDFs ─► pypdf (texto por página) ─► regex de linha de item ─► registros
                                         └─► regex "Total Geral" ──────┐
                                                                       ▼
                                           soma dos valores × Total Geral ─► OK / CONFERIR
registros ─► relatorio_saida.csv
```

## Conferência com o Total Geral

Para cada PDF, o script soma a coluna `Vl. Aquisição` dos itens extraídos e compara com o "Total Geral" do relatório (tolerância de R$ 0,01). Um PDF entra na lista **CONFERIR** quando:

| Situação | Mensagem no resumo | Causa provável |
|---|---|---|
| Nenhum item encontrado | `nenhum item encontrado (layout diferente ou PDF sem texto)` | PDF escaneado (imagem) ou layout diferente do padrão |
| "Total Geral" ausente | `'Total Geral' não encontrado no PDF` | Rótulo diferente ou PDF sem esse total |
| Soma **menor** que o total | `soma extraída … menor que o Total Geral … possíveis itens não capturados` | Alguma linha de item não casou com o regex |
| Soma **maior** que o total | `soma extraída … maior que o Total Geral … possíveis linhas duplicadas` | Linhas lidas a mais ou duplicadas |
| Erro de leitura | `erro ao ler o PDF (…)` | Arquivo corrompido ou protegido |

## Estrutura da saída

- **Arquivo:** `relatorio_saida.csv`
- **Separador:** `;`
- **Codificação:** `utf-8-sig` (acentos corretos ao abrir no Excel)
- **Cabeçalho:** `Arquivo;Página;Item;Patrimônio;Localização;Status Bem;Sit. Física;Vl. Aquisição`

> O valor de aquisição é gravado como **texto** no formato brasileiro (`1.180,00`). Para somar no Excel, pode ser necessário convertê-lo para número.

## Ajustando o padrão de leitura (regex)

Toda a extração depende do regex `padrao_item`, que descreve uma linha de item:

```
Item  Patrimônio(8 dígitos)  descrição  Localização  Status  Sit.Física  Valor
```

Pontos de atenção:

- A **Localização** usa o caractere `−` (sinal de menos Unicode, U+2212) como separador, **não** o hífen comum `-`. Se o seu PDF usar outro caractere, o regex não vai casar.
- O padrão exige que o **valor esteja no fim da linha** (`$`).
- Status e situação física são lidos como **uma única palavra** (`\S+`). Valores com espaço (ex.: "Em uso") não casam.
- O número de dígitos do patrimônio (`\d{8}`) é fixo.

Se o relatório mudar de layout, ajuste os grupos correspondentes e rode de novo: o campo de conferência com o Total Geral mostra rapidamente se algum item deixou de ser capturado.

## Solução de problemas

| Sintoma | Causa provável | O que fazer |
|---|---|---|
| `NameError: name 'pastaSaida' is not defined` | Linhas de caminho ainda comentadas | Remover o `#` e informar os caminhos |
| `0 PDF(s) encontrado(s)` | Caminho errado ou extensão diferente de `.pdf` minúsculo | Conferir `pastaEntrada`; arquivos `.PDF` em maiúsculas podem não ser listados em sistemas que diferenciam maiúsculas |
| `nenhum item encontrado` | PDF escaneado ou layout diferente | Verificar se o texto é selecionável; revisar o regex |
| Soma menor que o Total Geral | Itens não capturados | Comparar o PDF com o CSV (use a coluna `Página`) |
| Soma maior que o Total Geral | Linhas duplicadas ou lidas a mais | Procurar itens repetidos no CSV do mesmo arquivo |
| CSV com acentos quebrados | Aberto com codificação errada | Abrir pelo Excel (o arquivo já usa `utf-8-sig`) |
| `PermissionError` ao salvar | CSV aberto no Excel | Fechar o arquivo e rodar novamente |

## Limitações conhecidas

- **Não faz OCR:** PDFs escaneados (imagem) não são lidos. Para isso, a versão 1 ainda é a alternativa.
- Regex rígido: depende do layout exato do relatório (8 dígitos, separador `−`, status em uma palavra).
- Busca PDFs apenas na pasta informada (sem subpastas).
- O CSV é sobrescrito a cada execução.
- Itens que ocupam mais de uma linha no PDF podem não ser capturados.
- _Adicionar outras limitações aqui._

## Roadmap

- [ ] Parametrizar as pastas por linha de comando (`argparse`)
- [ ] Opção de gerar `.xlsx` formatado além do CSV
- [ ] Reintroduzir OCR como alternativa para PDFs em imagem
- [ ] Gravar o valor de aquisição como número
- [ ] Gerar relatório das conferências em arquivo
- [ ] _Novas ideias…_

## Histórico de alterações

| Data | Versão | Alteração |
|---|---|---|
| _AAAA-MM-DD_ | 1.0 | Extração por rótulos com OCR e exportação para Excel |
| _AAAA-MM-DD_ | 2.0 | Migração para `pypdf`, regex por linha, saída em CSV, novo campo Vl. Aquisição e conferência com o Total Geral |