from pathlib import Path

import numpy as np
import pandas as pd

BASE_DIR = Path(__file__).resolve().parent.parent
MAPPING_FILE = BASE_DIR / "mapping" / "tiendas_ejemplo.csv"
N_BLOQUES = 3


def split_in_blocks(centros, n_bloques=N_BLOQUES):
    return [bloque.tolist() for bloque in np.array_split(centros, n_bloques) if len(bloque) > 0]


def get_centros(data, marca):
    return (
        data.loc[data["MARCA"] == marca, "CENTRO"]
        .astype(str)
        .str.strip()
        .loc[lambda serie: serie != ""]
        .drop_duplicates()
        .tolist()
    )


def main():
    if not MAPPING_FILE.is_file():
        print(f"ERROR: no existe {MAPPING_FILE}. Ejecuta src/generar_datos_ejemplo.py.")
        return 1

    data = pd.read_csv(MAPPING_FILE, dtype=str, keep_default_na=False)
    faltantes = {"MARCA", "CENTRO"} - set(data.columns)
    if faltantes:
        print(f"ERROR: faltan columnas en el archivo: {', '.join(sorted(faltantes))}")
        return 1

    data["MARCA"] = data["MARCA"].astype(str).str.strip()
    data["CENTRO"] = data["CENTRO"].astype(str).str.strip()
    data = data.loc[data["MARCA"] != ""].copy()

    marcas = data["MARCA"].drop_duplicates().tolist()
    if not marcas:
        print("ERROR: el archivo de mapping no contiene marcas.")
        return 1

    total_centros = 0
    total_en_bloques = 0
    todo_ok = True

    for marca in marcas:
        centros = get_centros(data, marca)
        bloques = split_in_blocks(centros)
        suma_bloques = sum(len(bloque) for bloque in bloques)
        ok = suma_bloques == len(centros)
        todo_ok = todo_ok and ok

        total_centros += len(centros)
        total_en_bloques += suma_bloques
        estado = "OK" if ok else "ERROR"
        print(
            f"{marca}: centros={len(centros)}, "
            f"bloques={[len(bloque) for bloque in bloques]}, "
            f"suma={suma_bloques} [{estado}]"
        )

    print(f"TOTAL: centros={total_centros}, suma_en_bloques={total_en_bloques}")

    if not todo_ok or total_en_bloques != total_centros:
        print("ERROR: la suma de los bloques no coincide con el total de centros.")
        return 1

    print("Verificación correcta: la suma de los bloques coincide con el total de centros.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
