import logging
import sqlite3
from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# ============================================================
# CONFIGURAÇÃO
# ============================================================
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
)
logger = logging.getLogger("pipeline")

ARQUIVO = "vendas.csv"
TABELA = "clientes"
PK = ["customer_id"]
COLUNAS_ESPERADAS = {
    "customer_id",
    "customer_unique_id",
    "customer_zip_code_prefix",
    "customer_city",
    "customer_state",
}
UFS_VALIDAS = {
    "AC", "AL", "AP", "AM", "BA", "CE", "DF", "ES", "GO", "MA", "MT", "MS",
    "MG", "PA", "PB", "PR", "PE", "PI", "RJ", "RN", "RS", "RO", "RR", "SC",
    "SP", "SE", "TO",
}


def find_project_root() -> Path:
    """Sobe a partir deste arquivo até achar uma pasta que contenha data/raw."""
    for parent in Path(__file__).resolve().parents:
        if (parent / "data" / "raw").is_dir():
            return parent
    raise FileNotFoundError(
        f"Pasta 'data/raw' não encontrada acima de {Path(__file__).resolve()}"
    )


BASE_DIR = find_project_root()
RAW_DIR = BASE_DIR / "data" / "raw"
PROCESSED_DIR = BASE_DIR / "data" / "processed"
WAREHOUSE_DIR = BASE_DIR / "data" / "warehouse"
OUTPUT_DIR = BASE_DIR / "data" / "output"
DB_PATH = WAREHOUSE_DIR / "vendas.db"
REPORT_PATH = OUTPUT_DIR / "relatorio_clientes.xlsx"


# ============================================================
# EXTRACT
# ============================================================
def extract_csv(file_name: str) -> pd.DataFrame:
    file_path = RAW_DIR / file_name
    if not file_path.exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {file_path}")

    logger.info("Iniciando extração: %s", file_name)

    for encoding in ("utf-8", "latin-1"):
        try:
            df = pd.read_csv(
                file_path,
                encoding=encoding,
                dtype={"customer_zip_code_prefix": "str"},  # CEP como texto
            )
            logger.info("Lido com encoding: %s", encoding)
            break
        except UnicodeDecodeError:
            logger.warning("Falha com encoding %s, tentando próximo...", encoding)
    else:
        raise ValueError(f"Não foi possível decodificar {file_name}")

    if df.empty:
        raise ValueError(f"{file_name} está vazio")

    logger.info("Registros: %d | Colunas: %d", len(df), len(df.columns))
    return df


# ============================================================
# TRANSFORM
# ============================================================
def normalize_text(series: pd.Series) -> pd.Series:
    """Minúsculas, sem acentos, sem espaços extras."""
    return (
        series.str.strip()
        .str.lower()
        .str.normalize("NFKD")
        .str.encode("ascii", errors="ignore")
        .str.decode("ascii")
        .str.replace(r"\s+", " ", regex=True)
    )


def transform_customers(df: pd.DataFrame) -> pd.DataFrame:
    logger.info("Iniciando transformação (%d registros)", len(df))
    df = df.copy()
    df.columns = df.columns.str.strip().str.lower()

    faltando = COLUNAS_ESPERADAS - set(df.columns)
    if faltando:
        raise ValueError(f"Colunas ausentes: {sorted(faltando)}")

    # CEP com 5 dígitos
    df["customer_zip_code_prefix"] = (
        df["customer_zip_code_prefix"].astype(str).str.strip().str.zfill(5)
    )

    # Cidade padronizada
    antes = df["customer_city"].nunique()
    df["customer_city"] = normalize_text(df["customer_city"])
    logger.info("Cidades distintas: %d -> %d", antes, df["customer_city"].nunique())

    # Estado em maiúsculas e validado
    df["customer_state"] = df["customer_state"].str.strip().str.upper()
    invalidas = ~df["customer_state"].isin(UFS_VALIDAS)
    if invalidas.any():
        logger.warning("%d registros com UF inválida", invalidas.sum())

    # Duplicados
    antes = len(df)
    df = df.drop_duplicates()
    logger.info("Duplicados removidos: %d", antes - len(df))

    if df.duplicated(subset=PK).any():
        logger.warning("Existem chaves repetidas em %s", PK)

    logger.info("Transformação concluída (%d registros)", len(df))
    return df


def save_processed(df: pd.DataFrame, file_name: str) -> None:
    PROCESSED_DIR.mkdir(parents=True, exist_ok=True)
    path = PROCESSED_DIR / file_name
    df.to_csv(path, index=False, encoding="utf-8")
    logger.info("Arquivo salvo em: %s", path)


# ============================================================
# LOAD
# ============================================================
def load_to_sqlite(df: pd.DataFrame, table: str) -> None:
    WAREHOUSE_DIR.mkdir(parents=True, exist_ok=True)

    with sqlite3.connect(DB_PATH) as conn:
        df.to_sql(table, conn, if_exists="replace", index=False)
        conn.execute(
            f"CREATE INDEX IF NOT EXISTS idx_{table}_unique_id "
            f"ON {table} (customer_unique_id)"
        )
        total = conn.execute(f"SELECT COUNT(*) FROM {table}").fetchone()[0]

    if total != len(df):
        raise ValueError(f"Contagem diferente: df={len(df)} | banco={total}")

    logger.info("Carga concluída: %d registros na tabela '%s'", total, table)
    logger.info("Banco: %s", DB_PATH)


# ============================================================
# ANALYSIS
# ============================================================
QUERIES = {
    "clientes_por_estado": """
        SELECT customer_state AS estado,
               COUNT(*) AS total,
               ROUND(100.0 * COUNT(*) / (SELECT COUNT(*) FROM clientes), 2) AS percentual
        FROM clientes
        GROUP BY customer_state
        ORDER BY total DESC
    """,
    "top_10_cidades": """
        SELECT customer_city AS cidade,
               customer_state AS estado,
               COUNT(*) AS total
        FROM clientes
        GROUP BY customer_city, customer_state
        ORDER BY total DESC
        LIMIT 10
    """,
    "clientes_recorrentes": """
        SELECT CASE WHEN compras > 1 THEN 'recorrente' ELSE 'unico' END AS tipo,
               COUNT(*) AS clientes
        FROM (
            SELECT customer_unique_id, COUNT(*) AS compras
            FROM clientes
            GROUP BY customer_unique_id
        )
        GROUP BY tipo
    """,
    "top_10_clientes_mais_compras": """
        SELECT customer_unique_id, COUNT(*) AS compras
        FROM clientes
        GROUP BY customer_unique_id
        ORDER BY compras DESC
        LIMIT 10
    """,
    "cidades_por_estado": """
        SELECT customer_state AS estado,
               COUNT(DISTINCT customer_city) AS cidades
        FROM clientes
        GROUP BY customer_state
        ORDER BY cidades DESC
    """,
}


def run_analysis() -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    with sqlite3.connect(DB_PATH) as conn:
        for nome, sql in QUERIES.items():
            resultado = pd.read_sql(sql, conn)
            resultado.to_csv(OUTPUT_DIR / f"{nome}.csv", index=False, encoding="utf-8")
            print(f"\n=== {nome} ===")
            print(resultado.to_string(index=False))

    logger.info("Resultados salvos em: %s", OUTPUT_DIR)


# ============================================================
# RELATÓRIO EXCEL
# ============================================================
def _format_sheet(ws, header_row: int = 1) -> None:
    """Cabeçalho colorido, painel congelado e largura de colunas automática."""
    fill = PatternFill("solid", fgColor="1F4E78")
    for cell in ws[header_row]:
        if cell.value is not None:
            cell.fill = fill
            cell.font = Font(bold=True, color="FFFFFF")
            cell.alignment = Alignment(horizontal="center")

    ws.freeze_panes = ws.cell(row=header_row + 1, column=1)

    for col in ws.columns:
        largura = max(len(str(c.value)) if c.value is not None else 0 for c in col)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(largura + 3, 45)


def generate_excel_report() -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    with sqlite3.connect(DB_PATH) as conn:
        kpis = pd.read_sql(
            """
            SELECT COUNT(*) AS registros,
                   COUNT(DISTINCT customer_unique_id) AS unicos,
                   COUNT(DISTINCT customer_state) AS estados,
                   COUNT(DISTINCT customer_city) AS cidades
            FROM clientes
            """,
            conn,
        ).iloc[0]
        resultados = {nome: pd.read_sql(sql, conn) for nome, sql in QUERIES.items()}

    recompras = int(kpis["registros"] - kpis["unicos"])
    resumo = pd.DataFrame(
        {
            "Indicador": [
                "Total de registros",
                "Clientes únicos",
                "Compras de clientes recorrentes",
                "Taxa de recompra",
                "Estados atendidos",
                "Cidades distintas",
            ],
            "Valor": [
                int(kpis["registros"]),
                int(kpis["unicos"]),
                recompras,
                f"{100 * recompras / kpis['registros']:.2f}%",
                int(kpis["estados"]),
                int(kpis["cidades"]),
            ],
        }
    )

    with pd.ExcelWriter(REPORT_PATH, engine="openpyxl") as writer:
        # Aba Resumo
        resumo.to_excel(writer, sheet_name="Resumo", index=False, startrow=3)
        ws = writer.sheets["Resumo"]
        ws["A1"] = "Relatório de Clientes - Olist"
        ws["A1"].font = Font(bold=True, size=16)
        ws["A2"] = f"Gerado em {datetime.now():%d/%m/%Y %H:%M}"
        ws["A2"].font = Font(italic=True, color="666666")
        _format_sheet(ws, header_row=4)

        # Uma aba por análise
        for nome, df in resultados.items():
            df.to_excel(writer, sheet_name=nome[:31], index=False)
            _format_sheet(writer.sheets[nome[:31]])

        # Gráfico: top 10 estados
        ws_estado = writer.sheets["clientes_por_estado"]
        n = min(len(resultados["clientes_por_estado"]), 10)
        grafico = BarChart()
        grafico.title = "Top 10 estados por número de clientes"
        grafico.y_axis.title = "Clientes"
        grafico.x_axis.title = "Estado"
        grafico.add_data(
            Reference(ws_estado, min_col=2, min_row=1, max_row=n + 1),
            titles_from_data=True,
        )
        grafico.set_categories(Reference(ws_estado, min_col=1, min_row=2, max_row=n + 1))
        grafico.height, grafico.width = 8, 16
        ws_estado.add_chart(grafico, "F2")

    logger.info("Relatório Excel salvo em: %s", REPORT_PATH)


# ============================================================
# MAIN
# ============================================================
def run_pipeline() -> None:
    logger.info("=== Pipeline iniciada ===")
    logger.info("Raiz do projeto: %s", BASE_DIR)

    raw = extract_csv(ARQUIVO)
    clean = transform_customers(raw)
    save_processed(clean, "clientes_limpo.csv")
    load_to_sqlite(clean, TABELA)

    logger.info("=== Gerando análises ===")
    run_analysis()

    logger.info("=== Gerando relatório Excel ===")
    generate_excel_report()

    logger.info("=== Pipeline concluída com sucesso ===")


if __name__ == "__main__":
    try:
        run_pipeline()
    except Exception:
        logger.exception("Pipeline falhou")
        raise