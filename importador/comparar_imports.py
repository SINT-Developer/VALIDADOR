"""
Salva snapshot do banco e compara dois snapshots (tradicional vs importador).

Uso:
    python comparar_imports.py salvar tradicional
    python comparar_imports.py salvar importador
    python comparar_imports.py comparar
"""

import sys
import json
import pyodbc
from decimal import Decimal
import datetime

CONN_STR = (
    "DRIVER={ODBC Driver 17 for SQL Server};"
    "SERVER=127.0.0.1;"
    "DATABASE=SRPP;"
    "UID=sa;"
    "PWD=M4573R;"
    "TrustServerCertificate=yes;"
)

SNAPSHOT_DIR = "."

QUERIES = {
    "Produto": """
        SELECT CodProduto, CodAuxiliarProduto, Produto,
               PrecoTabela1, PrecoTabela2, PrecoTabela3,
               QtdeEstoqueAtual, QtdeEstoqueFuturo,
               LimiteDescIndividual, AliquotaIPI,
               MultiploGrade, DescontoGrade, PathFotografia,
               CodFilial, PrecoPromocional, TipoVendaSemEstoque,
               CodFamilia, CodEstilo, QtdeMultipla, QtdeMinima
        FROM Produto ORDER BY CodProduto, CodAuxiliarProduto
    """,
    "Cliente": """
        SELECT CodCliente, CodRepresentante, NomeFantasia, RazaoSocial,
               CNPJCPF, IERG, Logradouro, Bairro, Cidade, UF, CEP,
               DDD, Telefone1, Telefone2, FAX, NomeContato,
               Observacao, EMail, PrecoTabela, CodTransportadora
        FROM Cliente ORDER BY CodCliente
    """,
    "Representante": """
        SELECT CodRepresentante, Representante
        FROM Representante ORDER BY CodRepresentante
    """,
    "CondicaoPagamento": """
        SELECT CodCondPagamento, CondPagamento, Desconto1, Desconto2, Desconto3,
               VlrMinimoPedido, TipoCondPagamento, CondPagamentoPadrao
        FROM CondicaoPagamento ORDER BY CodCondPagamento
    """,
    "Transportadora": """
        SELECT CodTransportadora, Transportadora, TransportadoraPadrao
        FROM Transportadora ORDER BY CodTransportadora
    """,
    "Familia": """
        SELECT CodFamilia, Familia, MultiploFamilia, MinimoFamilia, DescontoFamilia
        FROM Familia ORDER BY CodFamilia
    """,
    "Estilo": """
        SELECT CodEstilo, Estilo
        FROM Estilo ORDER BY CodEstilo
    """,
    "Filial": """
        SELECT CodFilial, Filial, TituloAdicional1, TituloAdicional2
        FROM Filial ORDER BY CodFilial
    """,
    "RestricaoCondicaoPagamento": """
        SELECT CodCondPagamento, CodFilial, VlrMinimoPedido
        FROM RestricaoCondicaoPagamento ORDER BY CodCondPagamento, CodFilial
    """,
}


def serialize(val):
    if isinstance(val, Decimal):
        return float(val)
    if isinstance(val, (datetime.date, datetime.datetime)):
        return str(val)
    return val


def tirar_snapshot(nome):
    conn = pyodbc.connect(CONN_STR)
    snapshot = {}

    for tabela, query in QUERIES.items():
        cursor = conn.cursor()
        cursor.execute(query)
        cols = [d[0] for d in cursor.description]
        rows = []
        for row in cursor.fetchall():
            rows.append({col: serialize(val) for col, val in zip(cols, row)})
        snapshot[tabela] = rows
        print(f"  {tabela}: {len(rows)} registros")

    conn.close()

    path = f"{SNAPSHOT_DIR}/snapshot_{nome}.json"
    with open(path, "w", encoding="utf-8") as f:
        json.dump(snapshot, f, ensure_ascii=False, indent=2)

    print(f"\nSnapshot '{nome}' salvo em {path}")


def comparar():
    with open(f"{SNAPSHOT_DIR}/snapshot_tradicional.json", encoding="utf-8") as f:
        trad = json.load(f)
    with open(f"{SNAPSHOT_DIR}/snapshot_importador.json", encoding="utf-8") as f:
        imp = json.load(f)

    total_diffs = 0

    for tabela in QUERIES:
        rows_t = trad.get(tabela, [])
        rows_i = imp.get(tabela, [])

        print(f"\n{'='*60}")
        print(f"TABELA: {tabela}  |  Tradicional: {len(rows_t)}  |  Importador: {len(rows_i)}")

        if len(rows_t) != len(rows_i):
            print(f"  !! CONTAGEM DIFERENTE !!")
            total_diffs += 1

        # Indexar pelo primeiro campo (PK)
        pk = list(rows_t[0].keys())[0] if rows_t else (list(rows_i[0].keys())[0] if rows_i else None)
        if not pk:
            continue

        idx_t = {str(r[pk]): r for r in rows_t}
        idx_i = {str(r[pk]): r for r in rows_i}

        so_trad = set(idx_t) - set(idx_i)
        so_imp  = set(idx_i) - set(idx_t)

        if so_trad:
            print(f"  Registros so no TRADICIONAL ({len(so_trad)}): {sorted(so_trad)[:10]}")
            total_diffs += len(so_trad)
        if so_imp:
            print(f"  Registros so no IMPORTADOR ({len(so_imp)}): {sorted(so_imp)[:10]}")
            total_diffs += len(so_imp)

        # Comparar campos dos registros em comum
        diffs_campos = 0
        for chave in sorted(set(idx_t) & set(idx_i)):
            rt = idx_t[chave]
            ri = idx_i[chave]
            for campo in rt:
                vt = rt[campo]
                vi = ri.get(campo)
                # Comparar como string para evitar divergencias de tipo float
                if str(vt) != str(vi):
                    if diffs_campos == 0:
                        print(f"  Diferencas de campo:")
                    print(f"    PK={chave} | {campo}: TRAD={repr(vt)}  IMP={repr(vi)}")
                    diffs_campos += 1
                    if diffs_campos >= 20:
                        print(f"    ... (limitado a 20 diferencas por tabela)")
                        break
            if diffs_campos >= 20:
                break

        total_diffs += diffs_campos

        if not so_trad and not so_imp and diffs_campos == 0:
            print(f"  OK - identicos")

    print(f"\n{'='*60}")
    if total_diffs == 0:
        print("RESULTADO: IDENTICOS - Importador produziu exatamente o mesmo resultado que o metodo tradicional.")
    else:
        print(f"RESULTADO: {total_diffs} diferenca(s) encontrada(s). Revisar acima.")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)

    cmd = sys.argv[1]

    if cmd == "salvar" and len(sys.argv) == 3:
        nome = sys.argv[2]
        print(f"Tirando snapshot '{nome}'...")
        tirar_snapshot(nome)

    elif cmd == "comparar":
        comparar()

    else:
        print(__doc__)
        sys.exit(1)
