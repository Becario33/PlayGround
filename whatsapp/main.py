# PlayGround / whatsapp: consulta SQL → WhatsApp Web (grupo Data Analytics).
#   python main.py
# Primera vez: QR. Cierra Chrome antes de volver a correr.

import argparse
import os
import shutil
import sys
import time

URL_WA = "https://web.whatsapp.com/"
PERFIL = "chrome_whatsapp_perfil"
GRUPO_DEFAULT = "Data Analytics"
TEXTO_DEFAULT = "Prueba de persistencia: la sesión sigue abierta; el QR solo se escanea una vez."
PLANTILLA_NOMBRE = "Plantilla.xlsx"
CONSULTA = """
SELECT TOP 10
    NombrePelicula,
    SUM(CAST(Asistencia AS bigint)) AS Asistencia,
    SUM(CAST(Taquilla AS bigint)) AS Taquilla
FROM [Programacion].[dbo].[ComscoreMPAMexico]
WHERE FechaComscore = DATEADD(day, -1, CAST(GETDATE() AS date))
GROUP BY NombrePelicula
ORDER BY Taquilla DESC
""".strip()
DRIVERS_ODBC = (
    "ODBC Driver 17 for SQL Server",
    "ODBC Driver 18 for SQL Server",
)


def _dir_script():
    try:
        return os.path.dirname(os.path.abspath(__file__))
    except NameError:
        return os.getcwd()


def _cargar_env():
    ruta = os.path.join(_dir_script(), ".env")
    if not os.path.isfile(ruta):
        return False
    with open(ruta, encoding="utf-8") as f:
        for linea in f:
            linea = linea.strip()
            if not linea or linea.startswith("#") or "=" not in linea:
                continue
            k, v = linea.split("=", 1)
            # .env manda: no dejar un SQL_* vacío del sistema tapando la clave
            os.environ[k.strip()] = v.strip().strip('"').strip("'")
    return True


def _conectar_sql():
    try:
        import pyodbc
    except ImportError as e:
        raise RuntimeError(
            "Falta pyodbc. En whatsapp:\n"
            "  python -m pip install -r requirements.txt"
        ) from e
    server = os.environ.get("SQL_SERVER", "").strip()
    database = os.environ.get("SQL_DATABASE", "").strip()
    user = os.environ.get("SQL_USER", "").strip()
    password = os.environ.get("SQL_PASSWORD", "")
    if not server or not database or not user:
        raise RuntimeError(
            "Falta .env (SQL_SERVER, SQL_DATABASE, SQL_USER, SQL_PASSWORD).\n"
            "  copy .env.example .env"
        )
    if not password:
        raise RuntimeError(
            "SQL_PASSWORD vacío en .env. Edita whatsapp\\.env y vuelve a correr."
        )

    instalados = [d for d in pyodbc.drivers()]
    preferidos = [d for d in DRIVERS_ODBC if d in instalados]
    if not preferidos:
        raise RuntimeError(
            "No hay driver ODBC de SQL Server. Instala 'ODBC Driver 17 for SQL Server'.\n"
            "Drivers vistos: " + (", ".join(instalados) if instalados else "(ninguno)")
        )

    print("SQL: " + server + " / " + database + " (user " + user + ")")
    print("Drivers a probar: " + ", ".join(preferidos))
    errores = []
    for driver in preferidos:
        extra = "TrustServerCertificate=yes;Encrypt=no;" if "18" in driver else ""
        conn_str = (
            f"DRIVER={{{driver}}};SERVER={server};DATABASE={database};"
            f"UID={user};PWD={password};{extra}"
        )
        try:
            conn = pyodbc.connect(conn_str, timeout=15)
            print("  OK con driver:", driver)
            return conn
        except Exception as e:
            msg = str(e).split("\n")[0]
            errores.append(driver + " → " + msg)
            print("  Falló", driver + ":", msg)

    detalle = " | ".join(errores)
    raise RuntimeError(
        "No conectó a SQL Server (" + server + "). "
        "Revisa red corporativa / VPN / firewall y que el .env tenga el server bien.\n"
        + detalle
    )


def _celda(valor):
    if valor is None:
        return "NULL"
    if hasattr(valor, "strftime"):
        return valor.strftime("%Y-%m-%d")
    return str(valor)


def consulta_asistencia():
    conn = _conectar_sql()
    try:
        cur = conn.cursor()
        cur.execute(CONSULTA)
        cols = [d[0] for d in cur.description]
        filas = [tuple(f) for f in cur.fetchall()]
    finally:
        conn.close()
    lineas = [" | ".join(cols)]
    if not filas:
        lineas.append("(sin filas para ayer)")
    else:
        for i, fila in enumerate(filas, 1):
            lineas.append(f"{i} | " + " | ".join(_celda(v) for v in fila))
    return "\n".join(lineas), cols, filas


def _dir_resultados():
    ruta = os.path.join(_dir_script(), "resultados")
    os.makedirs(ruta, exist_ok=True)
    return ruta


def _etiqueta_ayer():
    from datetime import date, timedelta

    return (date.today() - timedelta(days=1)).strftime("%Y-%m-%d")


def _etiqueta_antier():
    from datetime import date, timedelta

    return (date.today() - timedelta(days=2)).strftime("%Y-%m-%d")


DIR_LAYOUT = "layout"


def _dir_layout():
    ruta = os.path.join(_dir_script(), DIR_LAYOUT)
    os.makedirs(ruta, exist_ok=True)
    return ruta


def plantilla_oficial():
    """Ruta fija de layout/Plantilla.xlsx (no se modifica)."""
    ruta = os.path.join(_dir_layout(), PLANTILLA_NOMBRE)
    if not os.path.isfile(ruta):
        raise RuntimeError(
            "Falta la plantilla oficial:\n  "
            + ruta
            + "\nPon Plantilla.xlsx en layout/ y vuelve a correr."
        )
    return ruta


def excel_en_layout():
    """Compat: plantilla oficial fija."""
    return plantilla_oficial()


def llenar_plantilla_top10(cols, filas):
    """
    Copia Plantilla.xlsx a resultados/ y escribe solo .value
    (como tecleo humano): B=Película, C=Asistencia, D=Taquilla, Total C15/D15.
    No toca font, fill, number_format ni formato condicional.
    """
    from openpyxl import load_workbook

    dia = _etiqueta_ayer()
    destino = os.path.join(_dir_resultados(), f"top10_taquilla_{dia}.xlsx")
    shutil.copy2(plantilla_oficial(), destino)

    por_nombre = {str(c).strip().lower(): i for i, c in enumerate(cols)}
    for clave in ("nombrepelicula", "asistencia", "taquilla"):
        if clave not in por_nombre:
            raise RuntimeError(
                "La consulta no trajo columna '" + clave + "'. Columnas: " + ", ".join(cols)
            )

    wb = load_workbook(destino)
    ws = wb.active

    total_asist = 0
    total_taq = 0
    for i in range(10):
        fila_excel = 3 + i
        if i < len(filas):
            fila = filas[i]
            pelicula = fila[por_nombre["nombrepelicula"]]
            asist = fila[por_nombre["asistencia"]]
            taq = fila[por_nombre["taquilla"]]
            ws.cell(fila_excel, 2).value = pelicula
            ws.cell(fila_excel, 3).value = int(asist) if asist is not None else None
            ws.cell(fila_excel, 4).value = int(taq) if taq is not None else None
            if asist is not None:
                total_asist += int(asist)
            if taq is not None:
                total_taq += int(taq)
        else:
            ws.cell(fila_excel, 2).value = None
            ws.cell(fila_excel, 3).value = None
            ws.cell(fila_excel, 4).value = None

    ws.cell(15, 3).value = total_asist if filas else None
    ws.cell(15, 4).value = total_taq if filas else None
    wb.save(destino)
    wb.close()
    print("Copia con datos (Plantilla intacta):")
    print(" ", destino)
    return destino


def _fmt_celda_excel(valor, num_fmt):
    if valor is None or valor == "":
        return ""
    if isinstance(valor, str):
        return valor
    fmt = (num_fmt or "General").lower()
    try:
        if "%" in fmt:
            return f"{float(valor) * 100:.1f}%"
        if "$" in fmt:
            if ".00" in fmt or "0.00" in fmt:
                return f"${float(valor):,.2f}"
            return f"${float(valor):,.0f}"
        if isinstance(valor, float) and not valor.is_integer():
            if ".00" in fmt or "0.00" in fmt:
                return f"{valor:,.2f}"
            return f"{valor:,.1f}"
        if isinstance(valor, (int, float)):
            return f"{int(round(float(valor))):,}"
    except Exception:
        pass
    return str(valor)


def _excel_tint_rgb(rgb, tint):
    """Aproxima el tint de Excel sobre un RGB 0–255."""
    if tint is None:
        return rgb
    r, g, b = rgb

    def one(c):
        c = c / 255.0
        if tint < 0:
            # aclarar (hacia blanco): típico Dark1 Lighter 95% ~ tint -0.05 en UI
            # En práctica header gris ≈ F2F2F2
            if abs(tint) < 0.08 and rgb == (0, 0, 0):
                return 0xF2
            c = c * (1.0 + tint) + (1.0 - (1.0 + tint)) * 1.0 * abs(tint) / max(abs(tint), 1e-9)
            # fallback lighten
            c = min(1.0, c + (1.0 - c) * abs(tint) * 18)
        else:
            c = c * (1.0 - tint)
        return int(round(max(0, min(1, c)) * 255))

    if tint < 0 and rgb == (0, 0, 0) and abs(tint) < 0.08:
        return (0xF2, 0xF2, 0xF2)
    return (one(r), one(g), one(b))


def _color_openpyxl(color):
    if color is None:
        return None
    try:
        if color.type == "rgb" and color.rgb and str(color.rgb) not in ("00000000", "None"):
            rgb = str(color.rgb)
            if len(rgb) == 8:
                rgb = rgb[2:]
            if len(rgb) == 6 and rgb.upper() != "000000":
                return "#" + rgb.upper()
            if len(rgb) == 6 and rgb.upper() == "000000":
                return "#000000"
        if color.type == "theme":
            theme = color.theme
            tint = float(color.tint or 0)
            base = {
                0: (0, 0, 0),
                1: (255, 255, 255),
                2: (31, 73, 125),
                3: (238, 236, 225),
                4: (79, 129, 189),
                5: (192, 80, 77),
                6: (155, 187, 89),
                7: (128, 100, 162),
                8: (75, 172, 198),
                9: (247, 150, 70),
            }.get(theme, (255, 255, 255))
            r, g, b = _excel_tint_rgb(base, tint)
            return f"#{r:02X}{g:02X}{b:02X}"
    except Exception:
        return None
    return None


def _fill_celda(cell):
    fill = cell.fill
    if not fill or fill.patternType not in ("solid",):
        return "#FFFFFF"
    hx = _color_openpyxl(fill.fgColor)
    if not hx or hx.upper() in ("#00000000",):
        return "#FFFFFF"
    # openpyxl a veces marca fill “vacío” como negro transparente
    if hx.upper() == "#000000" and fill.fgColor and getattr(fill.fgColor, "type", None) == "rgb":
        rgb = str(fill.fgColor.rgb or "")
        if rgb in ("00000000", "0"):
            return "#FFFFFF"
    return hx


def _fuente_celda(cell):
    hx = _color_openpyxl(cell.font.color) if cell.font else None
    return hx or "#404040"


def _borde_color(cell):
    for side in (cell.border.left, cell.border.right, cell.border.top, cell.border.bottom):
        if side and side.style:
            hx = _color_openpyxl(side.color) if side.color else None
            return hx or "#A6A6A6"
    return "#A6A6A6"


def _imagen_casi_blanca(ruta_o_img, umbral=0.995):
    """True si la imagen casi no tiene contenido (blanco/gris vacío)."""
    from PIL import Image

    if hasattr(ruta_o_img, "size"):
        img = ruta_o_img
    else:
        if not os.path.isfile(ruta_o_img):
            return True
        img = Image.open(ruta_o_img)
    if img.mode != "RGB":
        img = img.convert("RGB")
    # muestreo rápido
    small = img.resize((min(200, img.size[0]), min(120, img.size[1])))
    gris = small.convert("L")
    px = list(gris.getdata())
    if not px:
        return True
    # casi todo muy claro
    claros = sum(1 for p in px if p > 245)
    if (claros / float(len(px))) >= umbral:
        return True
    # poco contraste / casi sin tinta (COM a veces da un rectángulo gris vacío)
    mn, mx = gris.getextrema()
    if (mx - mn) < 50:
        return True
    oscuros = sum(1 for p in px if p < 200)
    if oscuros < max(3, int(len(px) * 0.005)):
        return True
    return False


def _rango_datos_xlsx(xlsx_path):
    """Hoja activa + dirección A1 del bloque con datos (sin celdas vacías de más)."""
    from openpyxl import load_workbook
    from openpyxl.utils import get_column_letter

    wb = load_workbook(xlsx_path, data_only=True)
    ws = wb.active
    min_r = min_c = max_r = max_c = None
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is not None and cell.value != "":
                r, c = cell.row, cell.column
                min_r = r if min_r is None else min(min_r, r)
                max_r = r if max_r is None else max(max_r, r)
                min_c = c if min_c is None else min(min_c, c)
                max_c = c if max_c is None else max(max_c, c)
    if min_r is None:
        raise RuntimeError("El Excel no tiene datos.")
    addr = (
        f"{get_column_letter(min_c)}{min_r}:"
        f"{get_column_letter(max_c)}{max_r}"
    )
    return ws.title, addr


def _captura_excel_com(xlsx_path, jpg_path):
    """
    Windows: captura visual EXACTA del rango con datos (Excel real).
    Visible=True evita el JPEG blanco típico de Excel oculto.
    """
    if sys.platform != "win32":
        return False
    try:
        import pythoncom
        from win32com.client import Dispatch
        from PIL import Image, ImageGrab
    except Exception as e:
        print("  Excel COM no disponible:", str(e).split("\n")[0])
        print("  Instala: python -m pip install pywin32")
        return False

    try:
        hoja, addr = _rango_datos_xlsx(xlsx_path)
    except Exception as e:
        print("  No leí el rango:", str(e).split("\n")[0])
        return False

    pythoncom.CoInitialize()
    excel = None
    wb = None
    chart_obj = None
    try:
        excel = Dispatch("Excel.Application")
        # Visible=True es clave: sin ventana Excel a menudo exporta blanco
        excel.Visible = True
        excel.DisplayAlerts = False
        excel.ScreenUpdating = True
        excel.AskToUpdateLinks = False
        ruta = os.path.abspath(xlsx_path)
        print(f"  Excel COM: abriendo rango {hoja}!{addr}")
        wb = excel.Workbooks.Open(ruta, ReadOnly=True, UpdateLinks=0)
        try:
            ws = wb.Worksheets(hoja)
        except Exception:
            ws = wb.ActiveSheet
        ws.Activate()
        rng = ws.Range(addr)
        try:
            excel.ActiveWindow.Zoom = 100
            excel.ActiveWindow.ScrollRow = rng.Row
            excel.ActiveWindow.ScrollColumn = rng.Column
        except Exception:
            pass
        time.sleep(0.8)

        def _guardar_img(img, etiqueta):
            if img is None:
                return False
            if img.mode != "RGB":
                img = img.convert("RGB")
            if _imagen_casi_blanca(img):
                print(f"  {etiqueta}: casi blanca, descarto")
                return False
            img.save(jpg_path, format="JPEG", quality=95, optimize=True)
            print(f"  Captura Excel ({etiqueta}):", img.size[0], "x", img.size[1])
            return True

        # 1) CopyPicture → portapapeles (foto tal cual pantalla)
        for intento in range(3):
            try:
                rng.CopyPicture(Appearance=1, Format=2)  # xlScreen, xlBitmap
            except Exception:
                rng.CopyPicture(Appearance=1, Format=-4147)
            time.sleep(0.5 + intento * 0.3)
            img = ImageGrab.grabclipboard()
            if _guardar_img(img, f"clipboard {intento + 1}"):
                return True

        # 2) Chart.Export del mismo rango
        try:
            width = max(float(rng.Width), 10.0)
            height = max(float(rng.Height), 10.0)
            chart_obj = ws.ChartObjects().Add(
                float(rng.Left), float(rng.Top), width, height
            )
            chart = chart_obj.Chart
            rng.CopyPicture(Appearance=1, Format=2)
            time.sleep(0.5)
            chart.Paste()
            time.sleep(0.5)
            png_tmp = os.path.abspath(jpg_path.rsplit(".", 1)[0] + "_com.png")
            chart.Export(png_tmp)
            chart_obj.Delete()
            chart_obj = None
            if os.path.isfile(png_tmp):
                img = Image.open(png_tmp).convert("RGB")
                try:
                    os.remove(png_tmp)
                except Exception:
                    pass
                if _guardar_img(img, "Chart.Export"):
                    return True
        except Exception as e:
            print("  Chart.Export:", str(e).split("\n")[0])
            try:
                if chart_obj is not None:
                    chart_obj.Delete()
                    chart_obj = None
            except Exception:
                pass

        print("  Excel COM no logró una captura con contenido")
        return False
    except Exception as e:
        print("  Excel COM falló:", str(e).split("\n")[0])
        return False
    finally:
        try:
            if chart_obj is not None:
                chart_obj.Delete()
        except Exception:
            pass
        try:
            if wb is not None:
                wb.Close(False)
        except Exception:
            pass
        try:
            if excel is not None:
                excel.Quit()
        except Exception:
            pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def captura_xlsx_como_imagen(xlsx_path):
    """
    Excel con datos (resultados/) → foto EXACTA (Windows COM).
    Pillow solo si no hay Excel/COM (Mac o fallo total).
    """
    out = os.path.join(
        _dir_resultados(),
        os.path.splitext(os.path.basename(xlsx_path))[0] + "_wa.jpg",
    )
    print("Excel → captura (tal cual se ve):")
    print(" ", os.path.abspath(xlsx_path))
    if not os.path.isfile(xlsx_path):
        raise RuntimeError("No existe el Excel: " + xlsx_path)

    # Windows: foto real del Excel (CF, formatos, merges tal cual)
    if sys.platform == "win32":
        if _captura_excel_com(xlsx_path, out) and not _imagen_casi_blanca(out):
            print("  OK captura Excel (réplica visual)")
            print(" ", out)
            return out
        print("  AVISO: Excel COM falló; Pillow NO es idéntico al Excel.")
        print("  Cierra TODO Excel y reintenta.")

    # Fallback (Mac / si COM no está)
    _captura_xlsx_pillow(xlsx_path, out)
    if _imagen_casi_blanca(out):
        raise RuntimeError(
            "No pude capturar el Excel.\n"
            "En la PC de trabajo: cierra Excel, confirma pywin32, y vuelve a correr."
        )
    print("  OK Pillow (aproximado — no es captura exacta)")
    print(" ", out)
    return out


def _captura_xlsx_pillow(xlsx_path, jpg_path):
    """Fallback: pinta desde openpyxl respetando anchos, altos y fills."""
    from openpyxl import load_workbook
    from openpyxl.utils import get_column_letter
    from PIL import Image, ImageDraw, ImageFont

    wb_vals = load_workbook(xlsx_path, data_only=True)
    wb = load_workbook(xlsx_path, data_only=False)
    ws_v = wb_vals.active
    ws = wb[ws_v.title]

    min_r = min_c = max_r = max_c = None
    for row in ws_v.iter_rows():
        for cell in row:
            if cell.value is not None and cell.value != "":
                r, c = cell.row, cell.column
                min_r = r if min_r is None else min(min_r, r)
                max_r = r if max_r is None else max(max_r, r)
                min_c = c if min_c is None else min(min_c, c)
                max_c = c if max_c is None else max(max_c, c)
    if min_r is None:
        raise RuntimeError("El Excel está vacío.")

    escala = 2
    default_row = float(ws.sheet_format.defaultRowHeight or 15)
    default_col = float(ws.sheet_format.defaultColWidth or 8.43)

    def col_w(c):
        letter = get_column_letter(c)
        dim = ws.column_dimensions.get(letter)
        wch = dim.width if dim and dim.width is not None else default_col
        return max(int(round((float(wch) + 0.75) * 7 * escala)), int(18 * escala))

    def row_h(r):
        dim = ws.row_dimensions.get(r)
        if dim is not None and dim.height is not None:
            pt = float(dim.height)
        else:
            pt = default_row
        return max(int(round(pt * 96 / 72 * escala)), int(12 * escala))

    anchos = [col_w(c) for c in range(min_c, max_c + 1)]
    altos = [row_h(r) for r in range(min_r, max_r + 1)]
    w = sum(anchos) + 1
    h = sum(altos) + 1
    img = Image.new("RGB", (w, h), "white")
    draw = ImageDraw.Draw(img)
    pad = max(2, int(2 * escala))
    med = ImageDraw.Draw(Image.new("RGB", (8, 8)))

    def font(pt, bold=False):
        px = max(1, int(round(float(pt) * 96 / 72 * escala)))
        for ruta in (
            ("calibrib.ttf" if bold else "calibri.ttf"),
            ("/Windows/Fonts/calibrib.ttf" if bold else "/Windows/Fonts/calibri.ttf"),
            ("arialbd.ttf" if bold else "arial.ttf"),
            ("/Windows/Fonts/arialbd.ttf" if bold else "/Windows/Fonts/arial.ttf"),
            (
                "/System/Library/Fonts/Supplemental/Arial Bold.ttf"
                if bold
                else "/System/Library/Fonts/Supplemental/Arial.ttf"
            ),
        ):
            try:
                return ImageFont.truetype(ruta, px)
            except Exception:
                continue
        return ImageFont.load_default()

    def wrap(texto, fnt, max_w):
        palabras = str(texto).split()
        if not palabras:
            return [""]
        lineas, actual = [], palabras[0]
        for p in palabras[1:]:
            prueba = actual + " " + p
            bb = med.textbbox((0, 0), prueba, font=fnt)
            if (bb[2] - bb[0]) <= max_w - pad * 2:
                actual = prueba
            else:
                lineas.append(actual)
                actual = p
        lineas.append(actual)
        return lineas

    y = 0
    for ir, r in enumerate(range(min_r, max_r + 1)):
        x = 0
        ah = altos[ir]
        for ic, c in enumerate(range(min_c, max_c + 1)):
            aw = anchos[ic]
            cell = ws.cell(r, c)
            cell_v = ws_v.cell(r, c)
            fill = _fill_celda(cell)
            borde = _borde_color(cell)
            draw.rectangle([x, y, x + aw - 1, y + ah - 1], fill=fill, outline=borde)
            texto = _fmt_celda_excel(cell_v.value, cell.number_format)
            if texto != "":
                size = cell.font.size or 11
                fnt = font(size, bool(cell.font.bold))
                color = _fuente_celda(cell)
                if isinstance(cell_v.value, (int, float)) and cell_v.value < 0:
                    # respeta rojo de fuente si ya viene; si no, marca negativo
                    if color.upper() in ("#404040", "#000000", "#595959"):
                        color = "#C00000"
                wrap_txt = bool(cell.alignment and cell.alignment.wrap_text) or (r == min_r)
                lineas = wrap(texto, fnt, aw) if wrap_txt else [texto]
                lhs = [med.textbbox((0, 0), ln, font=fnt)[3] - med.textbbox((0, 0), ln, font=fnt)[1] for ln in lineas]
                th = sum(lhs) + max(0, len(lineas) - 1)
                y0 = y + max(pad, (ah - th) // 2)
                align = (cell.alignment.horizontal if cell.alignment else None) or (
                    "left" if ic <= 1 else "right"
                )
                for ln, lh in zip(lineas, lhs):
                    tw = med.textbbox((0, 0), ln, font=fnt)[2] - med.textbbox((0, 0), ln, font=fnt)[0]
                    if align == "center":
                        tx = x + (aw - tw) // 2
                    elif align == "right":
                        tx = x + aw - tw - pad
                    else:
                        tx = x + pad
                    draw.text((tx, y0), ln, font=fnt, fill=color)
                    y0 += lh
            x += aw
        y += ah

    img.save(jpg_path, format="JPEG", quality=95, optimize=True)
    print("  Captura Pillow (fallback):", img.size[0], "x", img.size[1])
    return True


# Layout vacío clonado de: Películas Semana 37.xlsx → "Top Fin de Semana" AY4:BK16
# Colores/anchos/bordes leídos del Excel (hair + fill #F2F2F2).
ENCABEZADOS_LAYOUT_FS = (
    "Rank",
    "Película",
    "Asistencia Nacional",
    "Taquilla Nacional",
    "MS CNMX Asistencia",
    "MS SA CNMX Asistencia",
    "Dif SA",
    "MS CNMX Taquilla",
    "MS SA Cnmx Taquilla",
    "Dif SA",
    "PPB CNMX",
    "PPB SA",
    "Dif SA",
)
# Anchos Excel (character width); None → default 8.43
_ANCHOS_EXCEL_FS = (
    8.43, 32.0, 16.0, 16.53125, 9.19921875, 12.19921875, 9.19921875,
    8.43, 11.796875, 9.19921875, 8.43, 9.19921875, 9.19921875,
)
_ALTO_HEADER_PT = 32.55
_ALTO_FILA_PT = 15.75
_FILL_GRIS = "#F2F2F2"  # solo header y fila Total (rgb FFF2F2F2 del xlsx)
_BORDE_HAIR = "#A6A6A6"  # theme lt1 tint≈0.35
_TXT_HEADER = "#595959"
_TXT_CUERPO = "#404040"


def _excel_width_px(width_chars, escala):
    return max(int(round((float(width_chars) + 0.75) * 7 * escala)), int(20 * escala))


def _excel_height_px(points, escala):
    return max(int(round(float(points) * 96 / 72 * escala)), int(12 * escala))


def captura_layout_vacio():
    """
    Imagen layout AY4:BK16 vacía, visual como el Excel (bordes hair, grises, wrap).
    Solo textos: encabezados fila 4 + 'Total' en Película. Sin números ni ranks.
    """
    from PIL import Image, ImageDraw, ImageFont

    escala = 2
    headers = list(ENCABEZADOS_LAYOUT_FS)
    anchos = [_excel_width_px(w, escala) for w in _ANCHOS_EXCEL_FS]
    alto_h = _excel_height_px(_ALTO_HEADER_PT, escala)
    alto_f = _excel_height_px(_ALTO_FILA_PT, escala)
    filas_cuerpo = 12  # 5–14 datos, 15 vacía, 16 Total
    w = sum(anchos) + 1
    h = alto_h + alto_f * filas_cuerpo + 1

    def _font(size_pt, bold=False):
        px = max(1, int(round(size_pt * 96 / 72 * escala)))
        nombres = (
            ("calibrib.ttf" if bold else "calibri.ttf"),
            ("Calibri Bold.ttf" if bold else "Calibri.ttf"),
            ("/Windows/Fonts/calibrib.ttf" if bold else "/Windows/Fonts/calibri.ttf"),
            ("arialbd.ttf" if bold else "arial.ttf"),
            ("/Windows/Fonts/arialbd.ttf" if bold else "/Windows/Fonts/arial.ttf"),
            (
                "/System/Library/Fonts/Supplemental/Arial Bold.ttf"
                if bold
                else "/System/Library/Fonts/Supplemental/Arial.ttf"
            ),
        )
        for ruta in nombres:
            try:
                return ImageFont.truetype(ruta, px)
            except Exception:
                continue
        return ImageFont.load_default()

    font_h = _font(11, bold=True)
    font_t = _font(11, bold=True)
    draw_m = ImageDraw.Draw(Image.new("RGB", (8, 8)))

    def _wrap(texto, fnt, max_w, pad):
        palabras = str(texto).split()
        if not palabras:
            return [""]
        lineas, actual = [], palabras[0]
        for p in palabras[1:]:
            prueba = actual + " " + p
            bb = draw_m.textbbox((0, 0), prueba, font=fnt)
            if (bb[2] - bb[0]) <= max_w - pad * 2:
                actual = prueba
            else:
                lineas.append(actual)
                actual = p
        lineas.append(actual)
        return lineas

    img = Image.new("RGB", (w, h), "white")
    draw = ImageDraw.Draw(img)
    pad = max(2, int(3 * escala))

    draw.rectangle([0, 0, w - 1, alto_h - 1], fill=_FILL_GRIS)
    x = 0
    for i, titulo in enumerate(headers):
        aw = anchos[i]
        lineas = _wrap(titulo, font_h, aw, pad)
        alt_lin = []
        for ln in lineas:
            bb = draw_m.textbbox((0, 0), ln, font=font_h)
            alt_lin.append(bb[3] - bb[1])
        th = sum(alt_lin) + max(0, len(lineas) - 1) * int(1 * escala)
        y0 = max(pad, (alto_h - th) // 2)
        for ln, lh in zip(lineas, alt_lin):
            bb = draw_m.textbbox((0, 0), ln, font=font_h)
            tw = bb[2] - bb[0]
            tx = x + pad if i <= 1 else x + (aw - tw) // 2
            draw.text((tx, y0), ln, font=font_h, fill=_TXT_HEADER)
            y0 += lh + int(1 * escala)
        x += aw

    x = 0
    for aw in anchos:
        draw.line([(x, 0), (x, alto_h)], fill=_BORDE_HAIR, width=1)
        x += aw
    draw.line([(0, 0), (w - 1, 0)], fill=_BORDE_HAIR, width=1)
    draw.line([(w - 1, 0), (w - 1, alto_h)], fill=_BORDE_HAIR, width=1)
    draw.line([(0, alto_h - 1), (w - 1, alto_h - 1)], fill=_BORDE_HAIR, width=1)

    for r in range(filas_cuerpo):
        y1 = alto_h + r * alto_f
        y2 = y1 + alto_f
        es_sep = r == 10
        es_total = r == 11
        x = 0
        for i, aw in enumerate(anchos):
            if es_total:
                fill = _FILL_GRIS
            else:
                fill = "white"
            draw.rectangle([x, y1, x + aw - 1, y2 - 1], fill=fill)
            if es_total and i == 1:
                draw.text((x + pad, y1 + pad), "Total", font=font_t, fill=_TXT_HEADER)
            # Rejilla hair en datos (5–14) y en Total (16); la fila 15 queda en blanco
            if not es_sep:
                draw.rectangle(
                    [x, y1, x + aw - 1, y2 - 1],
                    outline=_BORDE_HAIR,
                    width=1,
                )
            x += aw

    ruta = os.path.join(_dir_resultados(), "layout_fs_vacio.jpg")
    img.save(ruta, format="JPEG", quality=95, optimize=True)
    print("Layout vacío (clon visual Excel AY4:BK16):")
    print(" ", ruta)
    print(" ", img.size[0], "x", img.size[1])
    return ruta


def armar_excel(cols, filas):
    from datetime import datetime
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

    dia = _etiqueta_antier()
    wb = Workbook()
    ws = wb.active
    ws.title = "Comscore"
    ws["A1"] = f"Top 10 películas por taquilla Comscore ({dia})"
    ws["A1"].font = Font(bold=True, size=14)
    # # | Película | Taquilla | Asistencia
    headers = ["#"] + list(cols)
    n = max(len(headers), 1)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n)
    encabezado = PatternFill("solid", fgColor="217346")
    fuente = Font(bold=True, color="FFFFFF")
    borde = Border(
        left=Side(style="thin", color="B0B0B0"),
        right=Side(style="thin", color="B0B0B0"),
        top=Side(style="thin", color="B0B0B0"),
        bottom=Side(style="thin", color="B0B0B0"),
    )
    for i, col in enumerate(headers, 1):
        cel = ws.cell(2, i, col)
        cel.fill = encabezado
        cel.font = fuente
        cel.alignment = Alignment(horizontal="center")
        cel.border = borde
    if not filas:
        ws.cell(3, 1, "(sin filas para antier)").border = borde
        for c in range(2, n + 1):
            ws.cell(3, c, "").border = borde
    else:
        for r, fila in enumerate(filas, 3):
            cel_n = ws.cell(r, 1, r - 2)
            cel_n.alignment = Alignment(horizontal="center")
            cel_n.border = borde
            for c, valor in enumerate(fila, 2):
                cel = ws.cell(r, c)
                if hasattr(valor, "strftime"):
                    cel.value = valor.strftime("%Y-%m-%d")
                    cel.alignment = Alignment(horizontal="center")
                else:
                    cel.value = valor
                    if isinstance(valor, (int, float)):
                        if c == 3:
                            cel.number_format = '"$"#,##0'
                        else:
                            cel.number_format = "#,##0"
                        cel.alignment = Alignment(horizontal="right")
                cel.border = borde
    ws.column_dimensions["A"].width = 6
    ws.column_dimensions["B"].width = 42
    ws.column_dimensions["C"].width = 16
    ws.column_dimensions["D"].width = 14
    nombre = datetime.now().strftime("top10_taquilla_%Y-%m-%d.xlsx")
    ruta = os.path.join(_dir_resultados(), nombre)
    wb.save(ruta)
    return ruta


def _fuente_tabla(escala):
    from PIL import ImageFont

    candidatos = (
        ("arial.ttf", "arialbd.ttf"),
        ("Arial.ttf", "Arial Bold.ttf"),
        ("/Windows/Fonts/arial.ttf", "/Windows/Fonts/arialbd.ttf"),
        ("/System/Library/Fonts/Supplemental/Arial.ttf",
         "/System/Library/Fonts/Supplemental/Arial Bold.ttf"),
    )
    for regular, bold in candidatos:
        try:
            return (
                ImageFont.truetype(regular, 16 * escala),
                ImageFont.truetype(bold, 18 * escala),
                ImageFont.truetype(bold, 16 * escala),
            )
        except Exception:
            continue
    font = ImageFont.load_default()
    return font, font, font


def _captura_pillow(xlsx_path, jpg_path):
    from openpyxl import load_workbook
    from PIL import Image, ImageDraw

    wb = load_workbook(xlsx_path)
    ws = wb.active
    n_cols = max(ws.max_column or 1, 1)
    filas = []
    for row in ws.iter_rows(min_row=1, max_col=n_cols, max_row=ws.max_row, values_only=True):
        celdas = []
        for j, v in enumerate(row):
            if v is None:
                celdas.append("")
            elif isinstance(v, (int, float)) and not isinstance(v, bool):
                # Columna Taquilla (por encabezado) con $
                encabezados = filas[1] if len(filas) > 1 else []
                if j < len(encabezados) and "taquilla" in str(encabezados[j]).lower():
                    celdas.append(f"${int(v):,}")
                else:
                    celdas.append(f"{int(v):,}")
            else:
                celdas.append(str(v))
        if any(c.strip() for c in celdas):
            filas.append(celdas)
    if not filas:
        filas = [["(sin filas)"]]

    escala = 3
    pad_x, pad_y = 10 * escala, 3 * escala
    font, font_t, font_b = _fuente_tabla(escala)
    medidor = ImageDraw.Draw(Image.new("RGB", (8, 8)))

    def _font_fila(i):
        if i == 0:
            return font_t
        if i == 1:
            return font_b
        return font

    def _recorte_texto(txt, fnt, fill, bg):
        from PIL import ImageChops

        txt = txt or " "
        b = medidor.textbbox((0, 0), txt, font=fnt)
        w = max(int(b[2] - b[0]) + 80, 80)
        h = max(int(b[3] - b[1]) + 80, 80)
        im = Image.new("RGB", (w, h), bg)
        ImageDraw.Draw(im).text((40, 40), txt, font=fnt, fill=fill)
        diff = ImageChops.difference(im, Image.new("RGB", im.size, bg))
        mask = diff.convert("L").point(lambda p: 255 if p > 18 else 0)
        box = mask.getbbox()
        if not box:
            return Image.new("RGB", (1, 1), bg)
        return im.crop(box)

    ncols = max((len(f) for f in filas), default=1)
    recortes = []
    for i, fila in enumerate(filas):
        fnt = _font_fila(i)
        if i == 0:
            recortes.append([_recorte_texto(fila[0], fnt, "#217346", "white")])
        elif i == 1:
            recortes.append([
                _recorte_texto(fila[j] if j < len(fila) else "", fnt, "white", "#217346")
                for j in range(ncols)
            ])
        else:
            recortes.append([
                _recorte_texto(fila[j] if j < len(fila) else "", fnt, "#222222", "white")
                for j in range(ncols)
            ])

    anchos = []
    for c in range(ncols):
        mx = 1
        for i, fila_r in enumerate(recortes):
            if i == 0 or c >= len(fila_r):
                continue
            mx = max(mx, fila_r[c].size[0] + pad_x * 2)
        anchos.append(mx)
    titulo_w = recortes[0][0].size[0] + pad_x * 2
    if titulo_w > sum(anchos):
        anchos[-1] += titulo_w - sum(anchos)

    altos = [max(im.size[1] for im in fila_r) + pad_y * 2 for fila_r in recortes]
    w, h = sum(anchos), sum(altos)
    tabla = Image.new("RGB", (w, h), "white")
    y = 0
    for i, fila_r in enumerate(recortes):
        ah = altos[i]
        x = 0
        if i == 0:
            tabla.paste(fila_r[0], (pad_x, y + (ah - fila_r[0].size[1]) // 2))
        elif i == 1:
            tabla.paste(Image.new("RGB", (w, ah), "#217346"), (0, y))
            for j, im in enumerate(fila_r):
                tabla.paste(im, (x + pad_x, y + (ah - im.size[1]) // 2))
                x += anchos[j]
        else:
            for j, im in enumerate(fila_r):
                tabla.paste(im, (x + pad_x, y + (ah - im.size[1]) // 2))
                x += anchos[j]
        y += ah

    draw = ImageDraw.Draw(tabla)
    draw.rectangle([0, 0, w - 1, h - 1], outline="#C8C8C8")
    y = 0
    for ah in altos[:-1]:
        y += ah
        draw.line([(0, y), (w - 1, y)], fill="#C8C8C8")
    x = 0
    for c in range(ncols - 1):
        x += anchos[c]
        draw.line([(x, altos[0]), (x, h - 1)], fill="#C8C8C8")
    # Solo la tabla (sin lienzo gris). El anti-sticker ya va por paste/Document.
    tabla = tabla.convert("RGB")
    tabla.save(jpg_path, format="JPEG", quality=95, optimize=True)
    print("  JPEG foto:", tabla.size[0], "x", tabla.size[1])
    return jpg_path


def captura_excel(xlsx_path):
    jpg_path = os.path.splitext(xlsx_path)[0] + ".jpg"
    _captura_pillow(xlsx_path, jpg_path)
    print("Imagen (JPEG, no sticker):")
    print(" ", jpg_path)
    return jpg_path


def _perfil_en_uso(perfil):
    for nombre in ("SingletonLock", "lockfile", "DevToolsActivePort"):
        if os.path.exists(os.path.join(perfil, nombre)):
            return True
    return False


def _lanzar(pw, perfil):
    comunes = dict(
            headless=False,
            locale="es-MX",
            no_viewport=True,
        args=["--disable-dev-shm-usage"],
    )
    extras = dict(chromium_sandbox=True, ignore_default_args=["--no-sandbox"])
    ultimo = None
    for extra in (extras, {}):
        kwargs = {**comunes, **extra}
        try:
            return pw.chromium.launch_persistent_context(
                perfil, channel="chrome", **kwargs
            )
        except TypeError:
            ultimo = None
            continue
        except Exception as e:
            ultimo = e
            try:
                return pw.chromium.launch_persistent_context(perfil, **kwargs)
            except TypeError:
                continue
            except Exception as e2:
                ultimo = e2
    if ultimo:
        raise ultimo
    return pw.chromium.launch_persistent_context(perfil, **comunes)


def _abrir_whatsapp(page):
    url = page.url or ""
    if "web.whatsapp.com" in url and url != "about:blank":
        print("WhatsApp ya estaba en esta ventana.")
        return True
    ultimo = None
    for n in range(1, 4):
        try:
            page.goto(URL_WA, wait_until="domcontentloaded", timeout=90000)
            return True
        except Exception as e:
            ultimo = e
            print(f"Intento {n}/3: no cargó WhatsApp Web.")
            page.wait_for_timeout(2000)
    print("No cargó https://web.whatsapp.com/")
    print("  1) Cierra el Chrome del experimento (la ventana de la corrida anterior) y Enter en esa consola.")
    print("  2) En el corporativo WhatsApp a veces se corta; reintenta en un rato.")
    if ultimo is not None:
        print(" ", str(ultimo).split("\n")[0])
    return False


def _hay_sesion(page):
    if page.locator('[data-testid="chat-list-search-container"]').count():
        return True
    if page.get_by_role("textbox", name="Search or start a new chat").count():
        return True
    if page.get_by_role("textbox", name="Buscar un chat o iniciar uno nuevo").count():
        return True
    if page.locator("#pane-side").count():
        return True
    return page.locator('[aria-label="Lista de chats"]').count() > 0


def _hay_qr(page):
    if page.locator("canvas[aria-label*='QR']").count():
        return True
    if page.locator("div[data-ref]").count():
        return True
    if page.get_by_text("Escanea el código QR", exact=False).count():
        return True
    return page.get_by_text("Scan the QR code", exact=False).count() > 0


def esperar_sesion(page, segundos=180):
    fin = time.time() + segundos
    aviso_qr = False
    while time.time() < fin:
        if _hay_sesion(page):
            print("Sesión lista (sin QR).")
            return True
        if _hay_qr(page) and not aviso_qr:
            print("Escanea el QR con el teléfono. Esta carpeta de Chrome se reusa la próxima vez.")
            aviso_qr = True
        page.wait_for_timeout(1000)
    print("No entró a WhatsApp a tiempo.")
    return False


def _caja_busqueda(page):
    loc = page.locator('[data-testid="chat-list-search-container"] input[data-tab="3"]')
    if loc.count():
        return loc.first
    loc = page.get_by_role("textbox", name="Search or start a new chat")
    if loc.count():
        return loc.first
    loc = page.get_by_role("textbox", name="Buscar un chat o iniciar uno nuevo")
    if loc.count():
        return loc.first
    loc = page.locator('input[aria-label="Search or start a new chat"]')
    if loc.count():
        return loc.first
    loc = page.locator('input[data-tab="3"]')
    if loc.count():
        return loc.first
    return None


def _caja_mensaje(page):
    nombres = (
        "Type a message",
        "Escribe un mensaje",
        "Mensaje",
    )
    for nombre in nombres:
        loc = page.locator(f'[contenteditable="true"][aria-label="{nombre}"]')
        if loc.count():
            return loc.last
        loc = page.get_by_role("textbox", name=nombre)
        if loc.count():
            return loc.last
    loc = page.locator("footer [contenteditable='true']")
    if loc.count():
        return loc.last
    loc = page.locator('#main footer [contenteditable="true"]')
    if loc.count():
        return loc.last
    return None


def abrir_grupo(page, nombre):
    caja = _caja_busqueda(page)
    if caja is None:
        print("No encontré la búsqueda de chats.")
        return False
    caja.click()
    page.wait_for_timeout(300)
    caja.fill(nombre)
    print(f'Busqué: {nombre}')
    page.wait_for_timeout(2000)

    chat = page.locator(f'span[title="{nombre}"]')
    if chat.count() == 0:
        chat = page.locator(f'div[title="{nombre}"]')
    if chat.count() == 0:
        chat = page.get_by_title(nombre, exact=True)
    try:
        chat.first.click(timeout=15000)
    except Exception:
        print(f'No encontré el grupo "{nombre}" en la lista filtrada.')
        return False

    try:
        page.get_by_role("banner").get_by_text(nombre, exact=False).first.wait_for(timeout=8000)
    except Exception:
        try:
            page.locator("#main header").get_by_text(nombre, exact=False).first.wait_for(timeout=8000)
        except Exception:
            page.wait_for_timeout(1500)
    print(f'Grupo abierto: {nombre}')
    return True


def _boton_enviar(page):
    for sel in (
        'div[role="button"][aria-label="Send"]',
        'div[role="button"][aria-label="Enviar"]',
        'button[aria-label="Send"]',
        'button[aria-label="Enviar"]',
        'span[data-icon="send"]',
    ):
        loc = page.locator(sel)
        if loc.count():
            return loc.last
    return page.locator("button[aria-label='Send']")


def _js_clic_enviar_media(page):
    return page.evaluate(
        """() => {
            const el = document.querySelector('div[role="button"][aria-label="Send 1 selected"]')
                || document.querySelector('div[role="button"][aria-label="Send"]')
                || document.querySelector('div[role="button"][aria-label="Enviar"]')
                || document.querySelector('[data-icon="wds-ic-send-filled"]')?.closest('[role="button"]');
            if (!el) return false;
            el.click();
            return true;
        }"""
    )


def _hay_preview_media(page):
    for sel in (
        'div[role="button"][aria-label="Send 1 selected"]',
        'div[role="button"][aria-label="Send"]',
        'div[role="button"][aria-label="Enviar"]',
    ):
        if page.locator(sel).count():
            return True
    return False


def _mantener_seleccion(page):
    txt = page.evaluate(
        """() => (document.body && document.body.innerText || '').toLowerCase()"""
    )
    if not (
        "discard selection" in txt
        or "descartar selección" in txt
        or "descartar seleccion" in txt
    ):
        return False
    print("  diálogo Discard → Cancel")
    for sel in (
        'button:has-text("Cancel")',
        'button:has-text("Cancelar")',
        'div[role="button"]:has-text("Cancel")',
        'div[role="button"]:has-text("Cancelar")',
    ):
        loc = page.locator(sel)
        if loc.count():
            try:
                loc.first.click(timeout=3000)
                page.wait_for_timeout(400)
                return True
            except Exception:
                pass
    return False


def _abrir_clip(page):
    btn = page.locator('button[aria-label="Attach"], button[aria-label="Adjuntar"]')
    if not btn.count():
        return False
    try:
        if btn.first.get_attribute("aria-expanded") == "true":
            return True
        btn.first.click(timeout=8000)
    except Exception:
        return False
    for _ in range(20):
        if btn.first.get_attribute("aria-expanded") == "true":
            return True
        if page.locator('button[role="menuitem"]').count():
            return True
        page.wait_for_timeout(200)
    return True


def _pulsar_enviar_media(page):
    _mantener_seleccion(page)
    media = page.locator('div[role="button"][aria-label="Send 1 selected"]')
    if media.count() == 0:
        media = page.locator(
            'div[role="button"][aria-label="Send"], '
            'div[role="button"][aria-label="Enviar"]'
        )
    try:
        if media.count():
            media.last.click(timeout=10000, force=True)
        elif not _js_clic_enviar_media(page):
            return False
    except Exception:
        if not _js_clic_enviar_media(page):
            return False
    page.wait_for_timeout(2000)
    _mantener_seleccion(page)
    if _hay_preview_media(page):
        if not _js_clic_enviar_media(page):
            return False
        page.wait_for_timeout(1500)
    return True


def _adjuntar_pegando(page, ruta_abs):
    """
    Pegar el JPEG en la caja del chat. WhatsApp lo trata como foto
    (no usa el input de sticker / image/* del clip).
    """
    caja = _caja_mensaje(page)
    if caja is None:
        print("  paste: no hay caja de mensaje")
        return False
    try:
        caja.click(timeout=8000)
    except Exception:
        return False
    page.wait_for_timeout(200)
    with open(ruta_abs, "rb") as f:
        b64 = __import__("base64").b64encode(f.read()).decode("ascii")
    nombre = os.path.basename(ruta_abs)
    ok = page.evaluate(
        """async ({ b64, nombre }) => {
            const caja = document.querySelector('#main footer [contenteditable="true"]')
                || document.querySelector('footer [contenteditable="true"]');
            if (!caja) return false;
            const bin = atob(b64);
            const bytes = new Uint8Array(bin.length);
            for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
            const file = new File([bytes], nombre, { type: 'image/jpeg' });
            const dt = new DataTransfer();
            dt.items.add(file);
            caja.focus();
            const ev = new ClipboardEvent('paste', { bubbles: true, cancelable: true, clipboardData: dt });
            // Chromium a veces ignora clipboardData del constructor; forzar DataTransfer
            try {
                Object.defineProperty(ev, 'clipboardData', { get: () => dt });
            } catch (e) {}
            caja.dispatchEvent(ev);
            return true;
        }""",
        {"b64": b64, "nombre": nombre},
    )
    if not ok:
        print("  paste: evaluate falló")
        return False
    print("  paste JPEG en caja de mensaje")
    for _ in range(32):
        page.wait_for_timeout(250)
        _mantener_seleccion(page)
        if _hay_preview_media(page):
            print("  Preview OK tras paste.")
            return True
    print("  paste: no abrió preview")
    return False


def _descartar_preview(page):
    """Cierra un preview a medias antes de otro intento."""
    if not _hay_preview_media(page) and not page.evaluate(
        """() => {
            const t = (document.body && document.body.innerText || '').toLowerCase();
            return t.includes('discard selection') || t.includes('descartar');
        }"""
    ):
        return
    print("  cierro preview anterior...")
    page.keyboard.press("Escape")
    page.wait_for_timeout(500)
    for sel in (
        'button:has-text("Discard")',
        'button:has-text("Descartar")',
        'div[role="button"]:has-text("Discard")',
        'div[role="button"]:has-text("Descartar")',
    ):
        loc = page.locator(sel)
        if loc.count():
            try:
                loc.last.click(timeout=3000)
                page.wait_for_timeout(500)
                break
            except Exception:
                pass
    page.wait_for_timeout(400)


def _adjuntar_documento(page, ruta_abs):
    """Document: WhatsApp NUNCA lo convierte en sticker."""
    if not _abrir_clip(page):
        print("  document: no abrí Attach")
        return False
    page.wait_for_timeout(400)
    doc = page.locator(
        'button[role="menuitem"][aria-label*="Document"], '
        'button[role="menuitem"][aria-label*="Documento"]'
    )
    if not doc.count():
        print("  document: no vi menuitem Document")
        return False
    with open(ruta_abs, "rb") as f:
        raw = f.read()
    payload = {
        "name": os.path.basename(ruta_abs),
        "mimeType": "image/jpeg",
        "buffer": raw,
    }
    try:
        with page.expect_file_chooser(timeout=6000) as fc_info:
            doc.first.click(timeout=8000)
        fc_info.value.set_files(payload)
        print("  Document + file_chooser OK")
    except Exception as e:
        print("  document file_chooser:", str(e).split("\n")[0])
        infos = page.evaluate(
            """() => [...document.querySelectorAll('input[type="file"]')].map((el, i) => ({
                i, accept: el.accept || '', multiple: !!el.multiple
            }))"""
        )
        print("  inputs:", infos)
        puesto = False
        for info in infos:
            a = (info.get("accept") or "").lower().strip()
            # Solo input de documentos (*), nunca image/* ni png/webp
            if "webp" in a and "png" in a:
                continue
            if a in ("image/*",) or a.startswith("image/*"):
                continue
            if a in ("", "*", "*/*") or a.startswith("*") or (
                info.get("multiple") and "image" not in a
            ):
                try:
                    page.locator('input[type="file"]').nth(info["i"]).set_input_files(payload)
                    print(f"  set_input_files DOC input[{info['i']}] accept={a[:40]!r}")
                    puesto = True
                    break
                except Exception as e2:
                    print(f"  input[{info['i']}]:", str(e2).split("\n")[0])
        if not puesto:
            return False
    for _ in range(32):
        page.wait_for_timeout(250)
        _mantener_seleccion(page)
        if _hay_preview_media(page):
            print("  Preview OK (documento).")
            return True
    return False


def enviar_imagen(page, ruta, pie=""):
    """
    Por qué salía sticker: el input footer accept=image/* en WA actual a veces
    entra al pipeline de sticker aunque el preview se vea como foto.

    Orden:
      1) Pegar JPEG en la caja → foto real
      2) Si no, Attach → Document → nunca sticker (archivo con preview)
    """
    if not os.path.isfile(ruta):
        print("No está el JPEG de la captura.")
        return False
    ruta_abs = os.path.abspath(ruta)

    print("Intento 1: pegar imagen en el chat (foto)...")
    if _adjuntar_pegando(page, ruta_abs):
        if _pulsar_enviar_media(page):
            print("Imagen enviada (foto por paste).")
            return True
        print("  paste abrió preview pero no envió.")
    _descartar_preview(page)

    print("Intento 2: Document (garantiza que NO sea sticker)...")
    if _adjuntar_documento(page, ruta_abs):
        if _pulsar_enviar_media(page):
            print("Imagen enviada como documento (no sticker).")
            return True
        print("  documento en preview pero no envió.")

    print("No pude enviar la imagen.")
    return False


def escribir_y_enviar(page, texto):
    caja = _caja_mensaje(page)
    if caja is None:
        print("No encontré la caja de mensaje.")
        return False
    caja.click()
    page.wait_for_timeout(300)
    page.keyboard.insert_text(texto)
    page.wait_for_timeout(300)
    btn = _boton_enviar(page)
    if btn.count():
        try:
            btn.first.click(timeout=8000)
        except Exception:
            page.keyboard.press("Enter")
    else:
        page.keyboard.press("Enter")
    page.wait_for_timeout(2000)
    print("Mensaje enviado.")
    return True


def enviar_grupo(page, nombre, texto, imagen=None):
    if not abrir_grupo(page, nombre):
        return False
    if imagen:
        return enviar_imagen(page, imagen)
    return escribir_y_enviar(page, texto)


def enviar(page, telefono, texto):
    tel = "".join(c for c in telefono if c.isdigit())
    if not tel:
        print("Número vacío.")
        return False
    from urllib.parse import quote
    page.goto(
        f"https://web.whatsapp.com/send?phone={tel}&text={quote(texto)}",
        wait_until="domcontentloaded",
        timeout=120000,
    )
    if not esperar_sesion(page, 90):
        return False
    enviar_btn = page.locator('button[aria-label="Enviar"]')
    if enviar_btn.count() == 0:
        enviar_btn = page.locator('span[data-icon="send"]')
    try:
        enviar_btn.first.click(timeout=20000)
    except Exception:
        print("No encontré el botón Enviar. Revisa el número o si WhatsApp cambió la web.")
        return False
    page.wait_for_timeout(2000)
    print("Mensaje enviado (o al menos se pulsó Enviar).")
    return True


def main():
    if sys.platform == "win32":
        try:
            sys.stdout.reconfigure(encoding="utf-8")
        except Exception:
            pass

    ap = argparse.ArgumentParser(description="Consulta SQL → WhatsApp (perfil persistente).")
    ap.add_argument("--grupo", default=GRUPO_DEFAULT, help="Nombre exacto del grupo")
    ap.add_argument("--para", default="", help="Número destino con lada país, sin + (opcional)")
    ap.add_argument("--texto", default=None, help="Texto fijo; si omites, manda el resultado SQL")
    ap.add_argument("--prueba", action="store_true", help="Manda el texto de persistencia, sin SQL")
    ap.add_argument(
        "--top10",
        action="store_true",
        help="Manda el Top 10 de ayer por taquilla (SQL → Plantilla en resultados/).",
    )
    ap.add_argument("--solo-abrir", action="store_true", help="No enviar, solo abrir sesión")
    args = ap.parse_args()

    _cargar_env()
    texto = args.texto
    imagen = None
    if args.prueba:
        texto = TEXTO_DEFAULT
    elif texto is None and not args.solo_abrir:
        try:
            # Default y --top10: SQL Top10 ayer → copia Plantilla en resultados/ → captura
            print("Top 10 ayer por taquilla → Plantilla (copia en resultados/).")
            texto, cols, filas = consulta_asistencia()
            xlsx = llenar_plantilla_top10(cols, filas)
            imagen = captura_xlsx_como_imagen(xlsx)
            print("Resultado SQL:")
            print(texto)
            print()
        except Exception as e:
            print("No salió la consulta o la captura.")
            print(" ", str(e))
            return 1
        print("Se mandará la captura del Excel en resultados/.")
        print()

    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        print("Falta Playwright. En la carpeta del repo:")
        print("  python -m pip install -r requirements.txt")
        print("  python -m playwright install chromium")
        return 1

    perfil = os.path.join(_dir_script(), PERFIL)
    os.makedirs(perfil, exist_ok=True)
    print("Perfil Chrome (no es tu Chrome de diario):")
    print(" ", perfil)
    if _perfil_en_uso(perfil):
        print("Ese perfil parece ocupado. Cierra el Chrome del experimento (corrida anterior) y vuelve a correr.")

    with sync_playwright() as pw:
        try:
            ctx = _lanzar(pw, perfil)
        except Exception as e:
            print("No se abrió Chrome.")
            print("  Cierra la ventana Chrome del experimento si sigue abierta.")
            print("  python -m pip install -r requirements.txt")
            print("  python -m playwright install chromium")
            print(" ", str(e).split("\n")[0])
            return 1
        page = ctx.pages[0] if ctx.pages else ctx.new_page()
        if not _abrir_whatsapp(page):
            ctx.close()
            return 2
        ok = esperar_sesion(page)
        if ok and not args.solo_abrir:
            try:
            if args.para:
                    ok = enviar(page, args.para, texto)
            else:
                    ok = enviar_grupo(page, args.grupo, texto, imagen=imagen)
            except Exception as e:
                print("Falló el envío (Chrome sigue abierto).")
                print(" ", str(e).split("\n")[0])
                ok = False
        print("Cierra la ventana de Chrome cuando termines (o Enter aquí).")
        try:
            input()
        except EOFError:
            page.wait_for_timeout(8000)
        ctx.close()
    return 0 if ok else 2


if __name__ == "__main__":
    raise SystemExit(main())
