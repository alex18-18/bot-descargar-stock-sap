import os
import sys
from pathlib import Path
from datetime import datetime

import pandas as pd
import pyodbc

BASE_DIR = Path(__file__).resolve().parent.parent
RESULT_DIR = BASE_DIR / "output" / "RESULTADO_DB"

QUERY = """
SELECT bodega AS BODEGA, referencia AS REFERENCIA, codigo_barras AS CODIGO_BARRAS,
       descripcion AS DESCRIPCION, stock AS STOCK, proveedor AS PROVEEDOR
FROM stock_articulos
WHERE stock <> 0
"""


def get_driver():
    # Usa el driver ODBC de SQL Server más reciente instalado
    drivers = [d for d in pyodbc.drivers() if "SQL Server" in d]
    if not drivers:
        raise RuntimeError("No hay drivers ODBC de SQL Server instalados.")
    preferred = [d for d in drivers if d.startswith("ODBC Driver")]
    return sorted(preferred)[-1] if preferred else drivers[0]


def main():
    server, database = os.environ.get("DB_SERVER"), os.environ.get("DB_NAME")
    if not server or not database:
        raise RuntimeError("Define las variables de entorno DB_SERVER y DB_NAME.")

    driver = get_driver()
    conn_str = f"DRIVER={{{driver}}};SERVER={server};DATABASE={database};"
    user, password = os.environ.get("DB_USER"), os.environ.get("DB_PASSWORD")
    conn_str += f"UID={user};PWD={password};" if user and password else "Trusted_Connection=yes;"

    print(f"Conectando a {server}/{database}...")
    with pyodbc.connect(conn_str, timeout=30) as conn:
        df = pd.read_sql(QUERY, conn)
    print(f"Filas obtenidas: {len(df):,}")

    # La bodega llega con prefijo; solo interesan los últimos 4 dígitos
    df["BODEGA"] = df["BODEGA"].astype(str).str.strip().str[-4:]

    RESULT_DIR.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    # Excel admite como máximo 1,048,575 filas de datos; si se excede se guarda en CSV
    if len(df) < 1_048_576:
        out_file = RESULT_DIR / f"stock_{stamp}.xlsx"
        df.to_excel(out_file, index=False)
    else:
        out_file = RESULT_DIR / f"stock_{stamp}.csv"
        df.to_csv(out_file, index=False, sep=";", encoding="utf-8-sig")
    print(f"Resultado guardado en: {out_file}")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"ERROR: {e}")
        sys.exit(1)