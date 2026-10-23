import random
from pathlib import Path

import pandas as pd

MAPPING_DIR = Path(__file__).resolve().parent.parent / "mapping"

marcas = {"MARCA_A": 36, "MARCA_B": 62, "MARCA_C": 3, "MARCA_D": 90}
filas = []
for i, (marca, n) in enumerate(marcas.items(), start=1):
    for j in range(n):
        filas.append({"MARCA": marca, "CENTRO": f"{i}{j:03d}", "TIENDA": f"Tienda {marca[-1]}{j + 1}"})

random.seed(42)
random.shuffle(filas)
MAPPING_DIR.mkdir(parents=True, exist_ok=True)
pd.DataFrame(filas).to_csv(MAPPING_DIR / "tiendas_ejemplo.csv", index=False)
print(f"Generado: {len(filas)} filas")