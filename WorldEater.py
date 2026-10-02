# ============================================================
# EXTRAÇÃO DE INVENTÁRIO EM PDFs (TEXTO OU IMAGEM) -> EXCEL (.xlsx)
# Campos: Item, Patrimonio, Tombo, Localização, Status Bem, Sit.Fisica
# ============================================================
#
# PRÉ-REQUISITOS
#   1) pip install pandas pdfplumber pymupdf pytesseract pillow openpyxl
#   2) Instalar o Tesseract (https://github.com/UB-Mannheim/tesseract/wiki)
#      marcando o idioma "Portuguese". Conferir com: tesseract --list-langs

# --- 1. IMPORTAÇÕES ---
import io                              # io: trata bytes da imagem como se fossem um arquivo
import logging                         # logging: mensagens com nível (INFO, WARNING...)
import re                              # re: expressões regulares
from dataclasses import dataclass      # dataclass: classe simples para guardar dados
from pathlib import Path               # Path: caminhos de pastas e arquivos

import fitz                            # fitz (PyMuPDF): transforma páginas de PDF em imagens
import pandas as pd                    # pandas: tabelas (DataFrames)
import pdfplumber                      # pdfplumber: lê texto selecionável dos PDFs
import pytesseract                     # pytesseract: usa o Tesseract para ler imagens
from PIL import Image                  # Pillow: abre e manipula imagens


# --- 2. CONFIGURAÇÕES ---

# Caminhos completos do Windows: r"..." evita que "\U" seja lido como caractere especial.
PASTA_ENTRADA = Path(r"r")
PASTA_SAIDA = Path(r"C:\Users\claudio.mota\Documents\Nuceria_Resultados")
ARQUIVO_SAIDA = PASTA_SAIDA / "saida_inventario.xlsx"

INCLUIR_SUBPASTAS = False     # True = procura PDFs também nas subpastas

# True  = cada PDF pode ter VÁRIOS bens (uma linha no CSV por bem).
#         O texto é dividido em blocos, um a cada linha que começa com "Item".
# False = um bem por PDF (pega a primeira ocorrência de cada campo).
VARIOS_ITENS_POR_PDF = True

# OCR (usado só quando o PDF é imagem)
TESSERACT_EXE = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
IDIOMA_OCR = "por"            # idioma do texto (por = português)
DPI_OCR = 300                 # mais alto = mais preciso, porém mais lento
MIN_CARACTERES_TEXTO = 20     # abaixo disso, o PDF é tratado como imagem

pytesseract.pytesseract.tesseract_cmd = TESSERACT_EXE

# force=True faz as mensagens aparecerem também em Jupyter, Spyder e IDLE.
logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s", force=True)
log = logging.getLogger(__name__)


# --- 3. DEFINIÇÃO DOS CAMPOS ---

@dataclass(frozen=True)
class Campo:
    nome: str            # nome da coluna no CSV
    padrao: re.Pattern   # regex compilado; o valor fica no grupo 1


def campo(nome: str, rotulo: str, valor: str = r".+") -> Campo:
    """Monta um Campo a partir do rótulo do PDF e do formato do valor.

    rotulo: regex do texto que aparece antes do valor (ex.: 'Status[ \\t]*Bem').
    valor : regex do que deve ser capturado (padrão: até o fim da linha).

    [ \\t]* aceita só espaços e tabs (não quebra de linha), então um campo vazio
    não "rouba" o texto da linha de baixo.
    """
    return Campo(nome, re.compile(rf"{rotulo}[ \t]*:?[ \t]*({valor})", re.I))


# Os acentos tolerantes ([oô], [cç], [aã], [ií]) ajudam quando o OCR erra.
# O Patrimônio é capturado como texto para não perder zeros à esquerda (000123).
# A PRIMEIRA linha da lista é o campo que inicia cada bem (usado para dividir blocos).
CAMPOS = [
    campo("Item",        r"\bItem\b"),
    campo("Patrimonio",  r"Patrim[oô]nio", r"[\w./-]+"),
    campo("Tombo",       r"\bTomb(?:o|amento)\b", r"[\w./-]+"),   # aceita "Tombo" e "Tombamento"
    campo("Localização", r"Localiza[cç][aã]o"),
    campo("Status Bem",  r"Status[ \t]*Bem"),
    campo("Sit.Fisica",  r"Sit\.?[ \t]*F[ií]sica"),
]

# Divide o texto a cada linha que começa com "Item" (lookahead: mantém o "Item" no bloco).
DIVISOR_BLOCOS = re.compile(r"^(?=[ \t]*Item\b)", re.I | re.M)


# --- 4. LEITURA DO PDF (TEXTO DIRETO OU OCR) ---

def ler_texto_ocr(caminho_pdf: Path) -> str:
    """Converte cada página em imagem e lê o texto com OCR."""
    paginas = []
    with fitz.open(caminho_pdf) as doc:
        for pagina in doc:
            pix = pagina.get_pixmap(dpi=DPI_OCR)                      # página -> imagem
            imagem = Image.open(io.BytesIO(pix.tobytes("png")))       # bytes PNG -> Pillow
            paginas.append(pytesseract.image_to_string(imagem, lang=IDIOMA_OCR))
    return "\n".join(paginas)


def ler_texto_pdf(caminho_pdf: Path) -> str:
    """Lê o texto direto; se o PDF for imagem (quase sem texto), usa OCR."""
    with pdfplumber.open(caminho_pdf) as pdf:
        texto = "\n".join(p.extract_text() or "" for p in pdf.pages)

    if len(texto.strip()) < MIN_CARACTERES_TEXTO:
        log.info("%s: sem texto selecionável, usando OCR", caminho_pdf.name)
        texto = ler_texto_ocr(caminho_pdf)
    return texto


# --- 5. EXTRAÇÃO DOS CAMPOS ---

def dividir_em_blocos(texto: str) -> list[str]:
    """Separa o texto em um bloco por bem. Sem nenhum 'Item', devolve o texto inteiro."""
    if not VARIOS_ITENS_POR_PDF:
        return [texto]
    # Só mantém blocos que realmente têm o campo inicial (descarta cabeçalhos).
    blocos = [b for b in DIVISOR_BLOCOS.split(texto) if CAMPOS[0].padrao.search(b)]
    return blocos or [texto]


def extrair_bloco(bloco: str, arquivo: str) -> dict:
    """Extrai todos os campos de um bloco de texto. Campo não encontrado vira None."""
    registro = {"arquivo": arquivo}
    for c in CAMPOS:
        achado = c.padrao.search(bloco)
        if achado:
            registro[c.nome] = " ".join(achado.group(1).split())   # limpa espaços repetidos
        else:
            registro[c.nome] = None
            log.warning("%s: campo '%s' não encontrado", arquivo, c.nome)
    return registro


def extrair_registros(caminho_pdf: Path) -> list[dict]:
    """Extrai todos os bens de um PDF (uma linha da tabela por bem)."""
    texto = ler_texto_pdf(caminho_pdf)
    return [extrair_bloco(b, caminho_pdf.name) for b in dividir_em_blocos(texto)]


def listar_pdfs(pasta: Path) -> list[Path]:
    """Lista os PDFs da pasta em ordem alfabética (aceita .pdf e .PDF)."""
    busca = pasta.rglob if INCLUIR_SUBPASTAS else pasta.glob
    return sorted(p for p in busca("*") if p.suffix.lower() == ".pdf")


def processar_pasta(pasta: Path) -> tuple[list[dict], list[tuple[str, str]]]:
    """Processa todos os PDFs. Devolve (registros, falhas)."""
    arquivos = listar_pdfs(pasta)
    log.info("%d PDFs encontrados em %s", len(arquivos), pasta)

    registros, falhas = [], []
    for i, caminho in enumerate(arquivos, start=1):
        log.info("Processando %d/%d: %s", i, len(arquivos), caminho.name)
        try:
            registros.extend(extrair_registros(caminho))
        except Exception as e:                  # um PDF ruim não derruba os outros
            falhas.append((caminho.name, str(e)))
    return registros, falhas


# --- 6. MONTAGEM DA TABELA ---

def montar_dataframe(registros: list[dict]) -> pd.DataFrame:
    """Cria o DataFrame e remove linhas sem dados ou repetidas."""
    colunas = [c.nome for c in CAMPOS]
    df = pd.DataFrame(registros, columns=["arquivo", *colunas])
    df = df.dropna(how="all", subset=colunas).drop_duplicates()
    # kind="stable" mantém a ordem original dos bens dentro de cada arquivo.
    return df.sort_values("arquivo", kind="stable").reset_index(drop=True)


def salvar_excel(df: pd.DataFrame, caminho: Path) -> None:
    """Salva o DataFrame em .xlsx com cabeçalho congelado, filtro e colunas ajustadas."""
    with pd.ExcelWriter(caminho, engine="openpyxl") as escritor:
        df.to_excel(escritor, index=False, sheet_name="Inventario")
        planilha = escritor.sheets["Inventario"]

        planilha.freeze_panes = "A2"               # mantém o cabeçalho visível ao rolar
        planilha.auto_filter.ref = planilha.dimensions   # botões de filtro no cabeçalho

        # Ajusta a largura de cada coluna ao maior texto dela (limite de 60 caracteres).
        for coluna in planilha.columns:
            maior = max(len(str(celula.value or "")) for celula in coluna)
            planilha.column_dimensions[coluna[0].column_letter].width = min(maior + 2, 60)


# --- 7. FUNÇÃO PRINCIPAL ---

def main() -> None:
    print(f"Pasta de entrada: {PASTA_ENTRADA} (existe: {PASTA_ENTRADA.exists()})")

    registros, falhas = processar_pasta(PASTA_ENTRADA)
    for nome, erro in falhas:
        print(f"FALHA em {nome}: {erro}")

    if not registros:
        print("Nenhum PDF processado. Confira o caminho da pasta.")
        return

    df = montar_dataframe(registros)

    # Tabela vazia = nenhum regex casou. Mostramos o texto lido para ajustar os padrões.
    if df.empty:
        print("Nenhum campo foi extraído. Os regex não casaram com o texto lido.")
        print("Texto do primeiro PDF (primeiros 500 caracteres):")
        print(ler_texto_pdf(listar_pdfs(PASTA_ENTRADA)[0])[:500])
        return

    print(f"\n--- Resultado: {len(df)} linhas ---")
    print(df.head())
    print("\nValores vazios por coluna:")
    print(df.isna().sum())    # muitos vazios = regex dessa coluna precisa de ajuste

    PASTA_SAIDA.mkdir(parents=True, exist_ok=True)
    try:
        salvar_excel(df, ARQUIVO_SAIDA)
        print(f"\nSalvo em {ARQUIVO_SAIDA}")
    except PermissionError:
        print("Sem permissão para gravar. Feche a planilha se ela estiver aberta no Excel.")


if __name__ == "__main__":
    main()