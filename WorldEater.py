from pypdf import PdfReader  # pip install pypdf
import re
import csv
import os
import glob


#pastaEntrada = r"INFORMAR A PASTA ONDE ESTÁ LOCALIZADA OS PDFS "
#pastaSaida = r"INFORMAR A PASTA ONDE O ARQUIVO CSV SERÁ CRIADO "
#APOS INSERIR A PASTA DE ENTRADA E SAIDA RETIRAR OS #
arquivoSaida = os.path.join(pastaSaida, "relatorio_saida.csv")
os.makedirs(pastaSaida, exist_ok=True)

# Padrão da linha que inicia cada item do relatório
padrao_item = re.compile(
    r"^(\d+)\s+"                                    # 1: Item
    r"(\d{8})\s+"                                   # 2: Patrimônio (8 dígitos)
    r".*?\s+"                                       # descrição (ignorada aqui)
    r"(\d{5}\s+−\s+\d{8}\s+−\s+.*?)\s+"             # 3: Localização
    r"(\S+)\s+"                                     # 4: Status do bem
    r"(\S+)\s+"                                     # 5: Situação física
    r"(\d{1,3}(?:\.\d{3})*,\d{2})$"                 # 6: Valor de aquisição
)

def para_numero(valor):
    """Converte '1.180,00' em 1180.0"""
    return float(valor.replace(".", "").replace(",", "."))

def formatar(valor):
    """Converte 1180.0 em '1.180,00' (formato brasileiro)"""
    return f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

lista_pdfs = sorted(glob.glob(os.path.join(pastaEntrada, "*.pdf")))
print(f"{len(lista_pdfs)} PDF(s) encontrado(s) em {pastaEntrada}\n")

conferir = []   # lista de (nome do PDF, motivo) para o resumo final
totalItens = 0

with open(arquivoSaida, "w", newline="", encoding="utf-8-sig") as csvfile:
    escritor = csv.writer(csvfile, delimiter=";")
    escritor.writerow(["Arquivo", "Página", "Item", "Patrimônio",
                       "Localização", "Status Bem", "Sit. Física", "Vl. Aquisição"])

    for arquivoInput in lista_pdfs:
        nomeArquivo = os.path.basename(arquivoInput)
        registros = []
        total_relatorio = None

        try:
            with open(arquivoInput, "rb") as pdf:
                reader = PdfReader(pdf)

                for numero, pagina in enumerate(reader.pages, start=1):
                    texto = pagina.extract_text() or ""

                    for linha in texto.split("\n"):
                        m = padrao_item.match(linha.strip())
                        if m:
                            item, patrimonio, localizacao, status, sit_fisica, valor = m.groups()
                            registros.append([
                                nomeArquivo,
                                numero,
                                int(item),
                                patrimonio,
                                re.sub(r"\s+", " ", localizacao),
                                status,
                                sit_fisica,
                                valor,
                            ])

                    m_total = re.search(r"Total Geral:\s*([\d.]+,\d{2})", texto)
                    if m_total:
                        total_relatorio = m_total.group(1)

        except Exception as erro:
            print(f"[ERRO] {nomeArquivo}: {erro}")
            conferir.append((nomeArquivo, f"erro ao ler o PDF ({erro})"))
            continue

        escritor.writerows(registros)
        totalItens += len(registros)
        soma = sum(para_numero(r[7]) for r in registros)

        # --- Verifica se este PDF precisa de conferência e por quê ---
        motivos = []
        if not registros:
            motivos.append("nenhum item encontrado (layout diferente ou PDF sem texto)")
        if total_relatorio is None:
            motivos.append("'Total Geral' não encontrado no PDF")
        elif registros:
            diferenca = para_numero(total_relatorio) - soma
            if abs(diferenca) >= 0.01:
                if diferenca > 0:
                    motivos.append(
                        f"soma extraída ({formatar(soma)}) menor que o Total Geral "
                        f"({total_relatorio}) por {formatar(diferenca)}: possíveis itens não capturados"
                    )
                else:
                    motivos.append(
                        f"soma extraída ({formatar(soma)}) maior que o Total Geral "
                        f"({total_relatorio}) por {formatar(-diferenca)}: possíveis linhas duplicadas ou lidas a mais"
                    )

        if motivos:
            conferir.append((nomeArquivo, "; ".join(motivos)))
            situacao = "CONFERIR"
        else:
            situacao = "OK"

        print(f"{nomeArquivo}: {len(registros)} itens | soma {formatar(soma)} | "
              f"total do relatório {total_relatorio} -> {situacao}")

# --- Resumo final ---
print("\n" + "=" * 60)
print(f"RESUMO: {len(lista_pdfs)} PDF(s) lido(s), {totalItens} itens salvos em:")
print(arquivoSaida)
print("=" * 60)

if conferir:
    print(f"\n{len(conferir)} PDF(s) precisam ser conferidos:\n")
    for nome, motivo in conferir:
        print(f" - {nome}")
        print(f"   Motivo: {motivo}\n")
else:
    print("\nTodos os PDFs bateram com o Total Geral. Nenhuma conferência necessária.")