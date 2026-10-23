# Bot de descarga de stock SAP

Automatización en Python que descarga inventarios desde SAP GUI y consulta
stock en SQL Server. El proyecto está preparado para ejecutarse en modo
simulacro sin SAP, sin base de datos y sin credenciales reales.

## Problema que resuelve

La descarga manual obliga a entrar a SAP, cargar los centros de cada marca y
exportar los resultados uno por uno. Con marcas de decenas de centros, la
transacción también puede saturarse si se envían listas demasiado largas. Este
bot reparte los centros en bloques, automatiza las exportaciones y deja un
registro de cada ejecución.

## Qué hace

- Lee marcas y centros desde `mapping/tiendas_ejemplo.csv`.
- Divide los centros de cada marca en 3 bloques equilibrados, sin duplicados.
- En modo real, carga los centros en la transacción MB52 mediante el
  portapapeles y exporta cada bloque desde SAP.
- Guarda una carpeta por marca dentro de `output/SAP` y archivos con el nombre
  `export_<marca>_bloque_<n>`.
- Si un bloque falla, registra el error, continúa con los demás y advierte al
  final cuáles quedaron pendientes.
- Espera la ventana de Excel con un tiempo máximo y falla de forma controlada
  si no aparece.
- Consulta stock en SQL Server y guarda el resultado en `output/RESULTADO_DB`.
- Registra información y errores en carpetas separadas dentro de `logs/`, con
  limpieza automática de logs de más de 30 días.

## Tecnologías

Python · pandas · numpy · pyodbc · pywin32 (SAP GUI Scripting) · pyperclip ·
openpyxl

## Estructura

```text
.
├── config/
│   └── config.example.json
├── mapping/
│   ├── .gitkeep
│   └── tiendas_ejemplo.csv
├── output/
│   └── .gitkeep
├── logs/
│   └── .gitkeep
├── src/
│   ├── main.py
│   ├── consulta_stock_y2.py
│   └── generar_datos_ejemplo.py
├── test/
│   └── prueba_bloques.py
├── .gitignore
├── README.md
└── requirements.txt
```

## Uso

### 1. Instalación

Se requiere Windows para la automatización de SAP GUI.

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

### 2. Generar datos de ejemplo

```powershell
python src/generar_datos_ejemplo.py
```

Esto crea `mapping/tiendas_ejemplo.csv` con marcas y centros ficticios.

### 3. Verificar el reparto en bloques

```powershell
python test/prueba_bloques.py
```

El script muestra, por marca, la cantidad de centros, el tamaño de cada bloque
y comprueba que la suma coincida con el total.

### 4. Ejecutar el modo simulacro

Con `DRY_RUN = True` en `src/main.py`:

```powershell
python src/main.py
```

El modo simulacro solo necesita el archivo de mapping. No abre SAP, no lee
credenciales y no requiere `config/config.json`.

### 5. Configurar una ejecución real

1. Copia `config/config.example.json` a `config/config.json`.
2. Reemplaza los valores de ejemplo por los de tu entorno local.
3. Ajusta la clave de `users` para que coincida con tu usuario de Windows.
4. Cambia `DRY_RUN = False` en `src/main.py`.

El archivo `config/config.json` está ignorado por Git y no debe publicarse.

### 6. Consulta SQL

La conexión se configura con variables de entorno:

```powershell
$env:DB_SERVER = "TU_SERVIDOR"
$env:DB_NAME = "TU_BASE_DE_DATOS"
python src/consulta_stock_y2.py
```

Si no se definen `DB_USER` y `DB_PASSWORD`, se usa autenticación de Windows
mediante `Trusted_Connection=yes`. La consulta de ejemplo usa tablas y campos
ficticios.

## Decisiones de diseño

- **Bloques:** repartir los centros evita enviar listas demasiado largas a la
  transacción y permite continuar si un lote falla.
- **Timeout de Excel:** la espera de la ventana tiene un máximo de 120 segundos
  para evitar bloqueos indefinidos.
- **Timeout de SAP:** la conexión espera como máximo 60 segundos a que SAP GUI
  exponga su API; si no ocurre, se lanza un error controlado.
- **Configuración externa:** credenciales, variante, conexiones y ruta de
  acceso se leen desde `config/config.json`, fuera del código.
- **Modo simulacro:** permite probar la división de centros sin instalar SAP,
  sin base de datos y sin configurar credenciales.
- **Salida separada:** la limpieza automática solo afecta a `output/SAP`, nunca
  al resto de resultados.
- **Formato de salida SQL:** se usa `.xlsx` mientras el volumen sea compatible
  con Excel; si se supera el límite de filas, se guarda en `.csv`.

> Las credenciales, los nombres de servidores, las marcas y cualquier dato de
> la empresa no forman parte de este repositorio.
