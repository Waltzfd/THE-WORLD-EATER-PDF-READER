# Extração de Inventário em PDFs → Excel

Script em Python que lê PDFs de inventário patrimonial (com texto selecionável **ou** escaneados como imagem), extrai os campos de cada bem e gera uma planilha `.xlsx` organizada.

> **Status:** em desenvolvimento · **Última atualização:** _preencher_

---

## Sumário

- [O que o script faz](#o-que-o-script-faz)
- [Campos extraídos](#campos-extraídos)
- [Pré-requisitos](#pré-requisitos)
- [Instalação](#instalação)
- [Configuração](#configuração)
- [Como usar](#como-usar)
- [Como funciona](#como-funciona)
- [Estrutura da saída](#estrutura-da-saída)
- [Ajustando os padrões (regex)](#ajustando-os-padrões-regex)
- [Solução de problemas](#solução-de-problemas)
- [Limitações conhecidas](#limitações-conhecidas)
- [Roadmap](#roadmap)
- [Histórico de alterações](#histórico-de-alterações)

---

## O que o script faz

1. Percorre uma pasta de PDFs (opcionalmente incluindo subpastas).
2. Tenta ler o texto direto do PDF; se o arquivo for uma imagem (quase sem texto), aplica **OCR** com Tesseract.
3. Divide o texto em blocos, um por bem (cada bloco começa numa linha com `Item`).
4. Extrai os campos de cada bloco usando expressões regulares.
5. Remove linhas vazias/duplicadas e salva tudo em um Excel formatado (cabeçalho congelado, filtros e colunas ajustadas).
6. Um PDF com erro **não interrompe** o processamento dos demais; as falhas são listadas no final.

## Campos extraídos

| Coluna | Descrição | Observações |
|---|---|---|
| `arquivo` | Nome do PDF de origem | Adicionado automaticamente |
| `Item` | Descrição do item | Também marca o início de cada bem |
| `Patrimonio` | Número de patrimônio | Lido como texto para preservar zeros à esquerda (`000123`) |
| `Tombo` | Número de tombamento | Aceita "Tombo" e "Tombamento" |
| `Localização` | Local do bem | |
| `Status Bem` | Situação cadastral do bem | |
| `Sit.Fisica` | Situação física do bem | |

Campos não encontrados ficam vazios (`None`) e geram um aviso no log.

## Pré-requisitos

- **Python** 3.10 ou superior (o código usa `list[str]`, `tuple[...]` e `X | Y`).
- **Tesseract OCR** com o idioma **Portuguese** instalado:
  - Windows: <https://github.com/UB-Mannheim/tesseract/wiki>
  - Conferir a instalação: `tesseract --list-langs` (deve listar `por`).

## Instalação

```bash
pip install pandas pdfplumber pymupdf pytesseract pillow openpyxl
```

> Sugestão: usar um ambiente virtual (`python -m venv .venv`).

## Configuração

Todas as configurações ficam no topo do script, na seção **2. CONFIGURAÇÕES**.

| Variável | Padrão | Função |
|---|---|---|
| `PASTA_ENTRADA` | _(caminho local)_ | Pasta onde estão os PDFs |
| `PASTA_SAIDA` | _(caminho local)_ | Pasta onde o Excel será salvo |
| `ARQUIVO_SAIDA` | `saida_inventario.xlsx` | Nome do arquivo gerado |
| `INCLUIR_SUBPASTAS` | `False` | `True` busca PDFs também em subpastas |
| `VARIOS_ITENS_POR_PDF` | `True` | `True`: vários bens por PDF; `False`: um bem por PDF |
| `TESSERACT_EXE` | `C:\Program Files\Tesseract-OCR\tesseract.exe` | Caminho do executável do Tesseract |
| `IDIOMA_OCR` | `por` | Idioma do OCR |
| `DPI_OCR` | `300` | Resolução do OCR (maior = mais preciso e mais lento) |
| `MIN_CARACTERES_TEXTO` | `20` | Abaixo disso, o PDF é tratado como imagem |

> Nos caminhos do Windows, use `r"..."` (string crua) para evitar problemas com `\`.

## Como usar

1. Coloque os PDFs na pasta definida em `PASTA_ENTRADA`.
2. Ajuste as configurações, se necessário.
3. Execute:

```bash
python nome_do_script.py
```

4. Ao final, abra o arquivo gerado em `PASTA_SAIDA`.

No console você verá o progresso (`Processando 3/20: arquivo.pdf`), as falhas, um resumo da tabela e a contagem de **valores vazios por coluna** — se uma coluna tiver muitos vazios, o regex dela provavelmente precisa de ajuste.

## Como funciona

```
PDF ──► tem texto? ──sim──► pdfplumber ─┐
            │                           ├─► texto ─► blocos por "Item" ─► regex por campo ─► DataFrame ─► .xlsx
            └──não──► PyMuPDF (página→imagem) ─► Tesseract (OCR) ─┘
```

| Etapa | Biblioteca |
|---|---|
| Leitura de texto selecionável | `pdfplumber` |
| Conversão de página em imagem | `PyMuPDF (fitz)` |
| OCR | `pytesseract` + `Pillow` |
| Tabela e exportação | `pandas` + `openpyxl` |

## Estrutura da saída

Planilha com uma aba chamada **Inventario**, uma linha por bem:

| arquivo | Item | Patrimonio | Tombo | Localização | Status Bem | Sit.Fisica |
|---|---|---|---|---|---|---|
| exemplo.pdf | … | … | … | … | … | … |

Recursos do Excel: cabeçalho congelado, filtros automáticos e largura de coluna ajustada (máx. 60 caracteres).

## Ajustando os padrões (regex)

Os campos são definidos na lista `CAMPOS`, usando a função auxiliar `campo(nome, rotulo, valor)`:

```python
campo("Patrimonio", r"Patrim[oô]nio", r"[\w./-]+")
```

- `rotulo`: texto que aparece antes do valor no PDF.
- `valor`: formato do que deve ser capturado (padrão: até o fim da linha).
- Classes como `[oô]`, `[cç]` e `[ií]` toleram erros comuns do OCR.
- A **primeira** entrada de `CAMPOS` define o início de cada bem.

Para **adicionar um novo campo**, inclua uma linha em `CAMPOS`; a coluna aparece automaticamente na planilha.

## Solução de problemas

| Sintoma | Causa provável | O que fazer |
|---|---|---|
| `Nenhum PDF processado` | Caminho da pasta errado | Conferir `PASTA_ENTRADA` |
| Tabela vazia | Regex não casou com o texto | O script imprime os primeiros 500 caracteres do 1º PDF; compare com os rótulos |
| `TesseractNotFoundError` | Caminho do Tesseract incorreto | Ajustar `TESSERACT_EXE` |
| Erro de idioma no OCR | Pacote `por` não instalado | Reinstalar o Tesseract marcando Portuguese |
| `Sem permissão para gravar` | Excel aberto com o arquivo | Fechar a planilha e rodar novamente |
| Muitos campos vazios | OCR de baixa qualidade ou rótulo diferente | Aumentar `DPI_OCR` ou revisar o regex |

## Limitações conhecidas

- Depende de os rótulos nos PDFs seguirem o padrão esperado (`Item`, `Patrimônio`, `Tombo`, etc.).
- A qualidade do OCR varia conforme a qualidade do escaneamento.
- A remoção de duplicatas considera todas as colunas, inclusive `arquivo`.
- _Adicionar outras limitações aqui._

## Roadmap

- [ ] Parametrizar pastas por linha de comando (`argparse`)
- [ ] Gerar relatório das falhas em arquivo
- [ ] Validação dos campos (ex.: formato do patrimônio)
- [ ] Testes automatizados com PDFs de exemplo
- [ ] _Novas ideias…_

## Histórico de alterações

| Data | Versão | Alteração |
|---|---|---|
| _AAAA-MM-DD_ | 0.1 | Versão inicial: extração de 6 campos, suporte a OCR e exportação para Excel |