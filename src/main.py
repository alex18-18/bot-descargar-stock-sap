import getpass
import json
import logging
import os
import shutil
import sys
import time
from datetime import datetime
from pathlib import Path

import numpy as np

# ---RUTAS---------------------------------------------------------------------

BASE_DIR = Path(__file__).resolve().parent.parent
LOGS_DIR = BASE_DIR / "logs"
CONFIG_DIR = BASE_DIR / "config"
CONFIG_FILE = CONFIG_DIR / "config.json"
MAPPING_DIR = BASE_DIR / "mapping"
MAPPING_FILE = MAPPING_DIR / "tiendas_ejemplo.csv"
SAP_DIR = BASE_DIR / "output" / "SAP"

# ---PARAMETROS----------------------------------------------------------------

N_BLOQUES = 3
TIMEOUT_EXCEL = 120   # segundos de espera por la ventana de Excel
TIMEOUT_SAP = 60      # segundos de espera para que SAP GUI exponga su API
DRY_RUN = True        # True: solo simula la division en bloques, sin abrir SAP


# ---LOGGING-------------------------------------------------------------------

class InfoFilter(logging.Filter):
    def filter(self, record):
        return record.levelno < logging.ERROR


class UserFilter(logging.Filter):
    def __init__(self):
        super().__init__()
        try:
            self.username = getpass.getuser()
        except Exception:
            self.username = os.environ.get("USERNAME", "system")

    def filter(self, record):
        record.username = self.username
        return True


def clean_old_logs(logs_path, days=30):
    now = datetime.now()
    for subfolder in ("execution_info", "execution_errors"):
        folder = Path(logs_path) / subfolder
        if not folder.is_dir():
            continue
        for file in folder.iterdir():
            try:
                age_days = (now - datetime.fromtimestamp(file.stat().st_mtime)).days
                if file.is_file() and age_days > days:
                    file.unlink()
            except OSError:
                continue


def setup_logging(logs_path, name="BotStockSAP"):
    logging.raiseExceptions = False
    dir_info = Path(logs_path) / "execution_info"
    dir_error = Path(logs_path) / "execution_errors"
    dir_info.mkdir(parents=True, exist_ok=True)
    dir_error.mkdir(parents=True, exist_ok=True)
    clean_old_logs(logs_path, days=30)
    timestamp = datetime.now().strftime("%d_%m_%Y_%H_%M_%S")

    logger = logging.getLogger(name)
    logger.propagate = False
    logger.setLevel(logging.INFO)
    logger.handlers.clear()

    formatter = logging.Formatter(
        "%(asctime)s | %(levelname)-8s | USER = %(username)s | MOD = %(module)s | FUNC = %(funcName)s | LN = %(lineno)d | MSG = %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )
    user_filter = UserFilter()

    info_handler = logging.FileHandler(dir_info / f"log_info_{timestamp}.log", encoding="utf-8")
    info_handler.setLevel(logging.INFO)
    info_handler.addFilter(InfoFilter())

    error_handler = logging.FileHandler(dir_error / f"log_errores_{timestamp}.log", encoding="utf-8")
    error_handler.setLevel(logging.ERROR)

    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(logging.INFO)

    for handler in (info_handler, error_handler, console_handler):
        handler.setFormatter(formatter)
        handler.addFilter(user_filter)
        logger.addHandler(handler)
    return logger


# ---UTILIDADES----------------------------------------------------------------

def load_config():
    if not CONFIG_FILE.is_file():
        raise FileNotFoundError(
            "No existe config/config.json. Copia config/config.example.json "
            "a config/config.json y ajusta los valores."
        )
    try:
        with open(CONFIG_FILE, encoding="utf-8") as file:
            return json.load(file)
    except json.JSONDecodeError as exc:
        raise ValueError(f"config/config.json no contiene un JSON válido: {exc}") from exc


def clean_folder():
    # Solo se limpia la salida de SAP para no borrar el resultado de la consulta SQL
    folder = Path(SAP_DIR)
    folder.mkdir(parents=True, exist_ok=True)
    for item in folder.iterdir():
        if item.is_dir():
            shutil.rmtree(item)
        else:
            item.unlink()


def split_in_blocks(centros, n_bloques=N_BLOQUES):
    # Reparte los centros en lotes parejos, sin lotes vacios
    return [bloque.tolist() for bloque in np.array_split(centros, n_bloques) if len(bloque) > 0]


def read_mapping():
    import pandas as pd

    if not MAPPING_FILE.is_file():
        raise FileNotFoundError(
            "No existe mapping/tiendas_ejemplo.csv. "
            "Ejecuta: python src/generar_datos_ejemplo.py"
        )

    data = pd.read_csv(MAPPING_FILE, dtype=str, keep_default_na=False)
    faltantes = {"MARCA", "CENTRO"} - set(data.columns)
    if faltantes:
        raise ValueError(f"Faltan columnas en {MAPPING_FILE.name}: {', '.join(sorted(faltantes))}")

    data["MARCA"] = data["MARCA"].astype(str).str.strip()
    data["CENTRO"] = data["CENTRO"].astype(str).str.strip()
    return data.loc[data["MARCA"] != ""].copy()


def get_centros(data, marca):
    return (
        data.loc[data["MARCA"] == marca, "CENTRO"]
        .astype(str)
        .str.strip()
        .loc[lambda serie: serie != ""]
        .drop_duplicates()
        .tolist()
    )


# ---SAP-----------------------------------------------------------------------

def stablish_sap_connection(sap_cfg):
    import win32com.client

    logger.info("Iniciando conexión con SAP...")
    os.startfile(sap_cfg["logon_path"])

    app = None
    deadline = time.monotonic() + TIMEOUT_SAP
    while time.monotonic() < deadline:
        try:
            app = win32com.client.GetObject("SAPGUI").GetScriptingEngine
            break
        except Exception:
            time.sleep(1)

    if app is None:
        raise RuntimeError(f"SAP GUI no respondió en {TIMEOUT_SAP} segundos.")

    connection = None
    for name in sap_cfg["connection_names"]:
        try:
            connection = app.OpenConnection(name, True)
            if connection is not None:
                logger.info(f"Conexión establecida con: {name}.")
                break
        except Exception as exc:
            logger.warning(f"No fue posible conectar con {name}: {exc}")

    if connection is None:
        raise RuntimeError("No se pudo abrir ninguna conexión de SAP.")
    return connection.Children(0)


def get_window_by_title(window_title, exact_match=False):
    import win32con
    import win32gui

    found = []

    def enum_windows(window, _):
        title = win32gui.GetWindowText(window)
        match = title == window_title if exact_match else window_title.lower() in title.lower()
        if match:
            found.append((window, title))

    win32gui.EnumWindows(enum_windows, None)
    if not found:
        return None
    handler, title = found[0]
    win32gui.ShowWindow(handler, win32con.SW_RESTORE)
    win32gui.ShowWindow(handler, win32con.SW_MAXIMIZE)
    return handler, title


def wait_and_close_excel(nombre):
    import win32con
    import win32gui

    deadline = time.monotonic() + TIMEOUT_EXCEL
    while time.monotonic() < deadline:
        result = get_window_by_title(nombre)
        if result:
            win32gui.PostMessage(result[0], win32con.WM_CLOSE, 0, 0)
            return
        time.sleep(1)
    raise TimeoutError(f"No apareció la ventana de Excel para {nombre}")


def download_block(session, marca, idx, total, bloque, marca_dir):
    import pyperclip

    nombre = f"export_{marca}_bloque_{idx}"
    logger.info(f"Descargando {marca} - bloque {idx}/{total} ({len(bloque)} centros)...")
    pyperclip.copy("\r\n".join(bloque))

    session.findById("wnd[0]/usr/btn%_WERKS_%_APP_%-VALU_PUSH").Press()
    session.findById("wnd[1]/tbar[0]/btn[16]").Press()
    session.findById("wnd[1]/tbar[0]/btn[24]").Press()
    session.findById("wnd[1]/tbar[0]/btn[8]").Press()
    session.findById("wnd[0]/tbar[1]/btn[8]").Press()

    grid = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell")
    grid.ContextMenu()
    grid.SelectContextMenuItem("&XXL")
    session.findById("wnd[1]/usr/ssubSUB_CONFIGURATION:SAPLSALV_GUI_CUL_EXPORT_AS:0512/txtGS_EXPORT-FILE_NAME").Text = nombre
    session.findById("wnd[1]/tbar[0]/btn[20]").Press()
    session.findById("wnd[1]/usr/ctxtDY_PATH").Text = str(marca_dir)
    session.findById("wnd[1]/tbar[0]/btn[11]").Press()

    wait_and_close_excel(nombre)
    logger.info(f"Informe {nombre} descargado con exito.")
    session.findById("wnd[0]/tbar[0]/btn[3]").Press()


def get_user_credentials(config):
    user = getpass.getuser()
    if user not in config.get("users", {}):
        raise KeyError(
            f"El usuario de Windows '{user}' no existe en config['users']. "
            "Agrégalo en config/config.json."
        )
    return user, config["users"][user]["SAP"]


def download_sap():
    data = read_mapping()
    marcas = data["MARCA"].drop_duplicates().tolist()

    if DRY_RUN:
        logger.info("Modo simulacro activo: no se abrirá SAP.")
        for marca in marcas:
            centros = get_centros(data, marca)
            bloques = split_in_blocks(centros)
            logger.info(
                f"[SIMULACRO] {marca}: {len(centros)} centros -> "
                f"bloques de {[len(bloque) for bloque in bloques]}"
            )
        return []

    config = load_config()
    user, sap_user = get_user_credentials(config)
    sap_cfg = config["sap"]

    logger.info(f"Iniciando descarga de informes de SAP para el usuario {user}...")
    clean_folder()
    session = stablish_sap_connection(sap_cfg)
    session.findById("wnd[0]").Maximize()
    session.findById("wnd[0]/usr/txtRSYST-BNAME").Text = sap_user["USER"]
    session.findById("wnd[0]/usr/pwdRSYST-BCODE").Text = sap_user["PASSWORD"]
    session.findById("wnd[0]").SendVKey(0)
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nMB52"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/tbar[1]/btn[17]").Press()
    session.findById("wnd[1]/usr/txtV-LOW").Text = sap_cfg["layout_variant"]
    session.findById("wnd[1]/usr/txtENAME-LOW").Text = sap_cfg["layout_owner"]
    session.findById("wnd[1]/tbar[0]/btn[8]").Press()

    fallidos = []
    for marca in marcas:
        marca_dir = SAP_DIR / marca
        marca_dir.mkdir(parents=True, exist_ok=True)
        centros = get_centros(data, marca)
        bloques = split_in_blocks(centros)
        if not bloques:
            logger.warning(f"{marca} no tiene centros; se omite.")
            continue

        for idx, bloque in enumerate(bloques, start=1):
            try:
                download_block(session, marca, idx, len(bloques), bloque, marca_dir)
            except Exception as exc:
                # Un bloque fallido no detiene el resto del proceso
                logger.error(f"Falló {marca} bloque {idx}: {exc}", exc_info=True)
                fallidos.append((marca, idx))

    if fallidos:
        logger.warning(f"Bloques fallidos: {fallidos}")
    logger.info("Proceso de descarga de los informes finalizado")
    return fallidos


# ---CONFIGURACION Y LOGGER----------------------------------------------------

try:
    logger = setup_logging(LOGS_DIR)
except Exception as e:
    logging.basicConfig(level=logging.INFO)
    logger = logging.getLogger("fallback")
    logger.warning(f"Usando logger de respaldo (solo consola): {e}")


# ---FLUJO PRINCIPAL-----------------------------------------------------------

def main():
    try:
        fallidos = download_sap()
        return 1 if fallidos else 0
    except Exception as e:
        logger.error(f"Error en el flujo principal: {e}", exc_info=True)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
