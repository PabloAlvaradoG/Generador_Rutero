# ============================================================
#  RUTERO DE DESPACHO - Streamlit App
#  Replica exacta del ModDespacho_v12.bas
#  Usuario sube PLANTILLA_DESPACHO.xlsm y descarga el rutero
# ============================================================

import streamlit as st
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, numbers
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.hyperlink import Hyperlink
import io
import datetime
from collections import defaultdict

# ── Colores (mismo que VBA) ───────────────────────────────────
C_AZO = "1F3864"
C_AZM = "2E75B6"
C_AZC = "D6E4F6"
C_BLN = "FFFFFF"
C_GRS = "F2F2F2"
C_BRD = "C0C0C0"

# ── Helpers de estilo ─────────────────────────────────────────
def thin_border():
    s = Side(style="thin", color=C_BRD)
    return Border(left=s, right=s, top=s, bottom=s)

def medium_bottom(color=C_AZO):
    s_thin = Side(style="thin", color=C_BRD)
    s_med  = Side(style="medium", color=color)
    return Border(left=s_thin, right=s_thin, top=s_thin, bottom=s_med)

def set_cell(ws, row, col, value=None, bold=False, size=9, color="000000",
             bg=None, halign="center", valign="center", fmt=None,
             italic=False, border=True):
    c = ws.cell(row=row, column=col)
    if value is not None:
        c.value = value
    c.font = Font(name="Arial", bold=bold, size=size, color=color, italic=italic)
    if bg:
        c.fill = PatternFill("solid", fgColor=bg)
    c.alignment = Alignment(horizontal=halign, vertical=valign)
    if border:
        c.border = thin_border()
    if fmt:
        c.number_format = fmt
    return c

def merge_set(ws, r1, c1, r2, c2, value=None, bold=False, size=9,
              color="000000", bg=None, halign="center", fmt=None, italic=False):
    ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)
    c = ws.cell(r1, c1)
    if value is not None:
        c.value = value
    c.font = Font(name="Arial", bold=bold, size=size, color=color, italic=italic)
    if bg:
        c.fill = PatternFill("solid", fgColor=bg)
    c.alignment = Alignment(horizontal=halign, vertical="center")
    c.border = thin_border()
    if fmt:
        c.number_format = fmt
    return c

def lbl_cab(ws, row, c1, c2, text):
    """Label de cabecera: fondo blanco, texto azul oscuro, negrita"""
    if c1 < c2:
        ws.merge_cells(start_row=row, start_column=c1, end_row=row, end_column=c2)
    c = ws.cell(row, c1)
    c.value = text
    c.font = Font(name="Arial", bold=True, size=8, color=C_AZO)
    c.fill = PatternFill("solid", fgColor=C_BLN)
    c.alignment = Alignment(horizontal="left", vertical="center")
    c.border = thin_border()

def val_cab(ws, row, c1, c2, value, bold=False, fmt=None):
    """Valor de cabecera: fondo blanco, texto gris oscuro"""
    if c1 < c2:
        ws.merge_cells(start_row=row, start_column=c1, end_row=row, end_column=c2)
    c = ws.cell(row, c1)
    c.value = value
    c.font = Font(name="Arial", bold=bold, size=9, color="323232")
    c.fill = PatternFill("solid", fgColor=C_BLN)
    c.alignment = Alignment(horizontal="left", vertical="center")
    c.border = thin_border()
    if fmt:
        c.number_format = fmt

def linea_cab(ws, row, c1, c2):
    """Celda con solo línea inferior (para CHOFER, PLACA, FECHA ENVIO)"""
    if c1 < c2:
        ws.merge_cells(start_row=row, start_column=c1, end_row=row, end_column=c2)
    c = ws.cell(row, c1)
    c.fill = PatternFill("solid", fgColor=C_BLN)
    bottom = Side(style="thin", color="646464")
    thin  = Side(style="thin", color=C_BRD)
    c.border = Border(left=thin, right=thin, top=thin, bottom=bottom)

# ── Normalización ─────────────────────────────────────────────
def norm(txt):
    return str(txt).strip().upper() if txt else ""

def nombre_hoja(txt):
    s = str(txt)[:28]
    for ch in ["/", "\\", ":", "*", "?", "[", "]", "!"]:
        s = s.replace(ch, "-")
    return s

def serial_a_texto(serial):
    """Convierte serial Excel a DD/MM/YYYY — region-independent"""
    try:
        s = int(serial)
        if s < 1 or s > 200000:
            return "--"
        if s > 60:
            s -= 1
        base = datetime.date(1900, 1, 1)
        result = base + datetime.timedelta(days=s - 1)
        return result.strftime("%d/%m/%Y")
    except:
        return "--"

def fecha_de_celda(cell):
    """
    Lee fecha de celda openpyxl y devuelve (serial_int, texto_DD/MM/YYYY).
    Usa .value que en openpyxl ya devuelve datetime correcto (no afectado por regional).
    """
    v = cell.value
    if v is None:
        return 0, "--"
    if isinstance(v, (datetime.datetime, datetime.date)):
        base = datetime.datetime(1899, 12, 30)
        if isinstance(v, datetime.date) and not isinstance(v, datetime.datetime):
            v = datetime.datetime(v.year, v.month, v.day)
        serial = (v - base).days
        texto = v.strftime("%d/%m/%Y")
        return serial, texto
    if isinstance(v, (int, float)):
        return int(v), serial_a_texto(int(v))
    return 0, str(v)

# ── Leer maestros ─────────────────────────────────────────────
def cargar_centros(wb):
    """CENTROS: fila 1=titulo, 2=subtitulo, 3=cabecera, datos desde 4"""
    centros_nombre  = {}
    centros_empresa = {}
    if "CENTROS" not in wb.sheetnames:
        return centros_nombre, centros_empresa, "No se encontró la hoja CENTROS"
    ws = wb["CENTROS"]
    last = ws.max_row
    for i in range(4, last + 1):
        v = ws.cell(i, 1).value
        if v is None:
            continue
        try:
            key = str(int(float(str(v))))
        except:
            key = norm(str(v))
        if key and key not in centros_nombre:
            centros_nombre[key]  = str(ws.cell(i, 2).value or "").strip()
            centros_empresa[key] = str(ws.cell(i, 3).value or "").strip()
    return centros_nombre, centros_empresa, None

def cargar_clientes(wb):
    """CLIENTES: fila 1=titulo, 2=subtitulo, 3=cabecera, datos desde 4"""
    cli_vendedor  = {}
    cli_prioridad = {}
    if "CLIENTES" not in wb.sheetnames:
        return cli_vendedor, cli_prioridad, "No se encontró la hoja CLIENTES"
    ws = wb["CLIENTES"]
    last = ws.max_row
    if last < 4:
        return cli_vendedor, cli_prioridad, None
    arr = []
    for i in range(4, last + 1):
        row = [ws.cell(i, c).value for c in range(1, 5)]
        arr.append(row)
    for row in arr:
        key = norm(str(row[0])) if row[0] else ""
        if not key or key in cli_vendedor:
            continue
        try:
            prio = int(float(str(row[3])))
            if prio < 1 or prio > 998:
                prio = 999
        except:
            prio = 999
        cli_vendedor[key]  = str(row[2] or "").strip()
        cli_prioridad[key] = prio
    return cli_vendedor, cli_prioridad, None

def cargar_rutas(wb):
    """RUTAS: fila 1=titulo, 2=subtitulo, 3=cabecera, datos desde 4"""
    ruta_camion = {}
    ruta_dia    = {}
    ruta_fecha  = {}
    if "RUTAS" not in wb.sheetnames:
        return ruta_camion, ruta_dia, ruta_fecha, "No se encontró la hoja RUTAS"
    ws = wb["RUTAS"]
    last = ws.max_row
    for i in range(4, last + 1):
        ciudad = norm(str(ws.cell(i, 2).value or ""))
        if not ciudad or ciudad in ruta_camion:
            continue
        ruta_camion[ciudad] = str(ws.cell(i, 1).value or "").strip()
        ruta_dia[ciudad]    = str(ws.cell(i, 4).value or "").strip()
        v5 = ws.cell(i, 5).value
        if isinstance(v5, (datetime.datetime, datetime.date)):
            fdesp = v5.strftime("%d/%m/%Y")
        elif v5:
            fdesp = str(v5).strip()
        else:
            fdesp = "--"
        ruta_fecha[ciudad] = fdesp
    return ruta_camion, ruta_dia, ruta_fecha, None

# ── Detectar columnas DATA ────────────────────────────────────
ALIASES = {
    "fecha":   ["fecha", "fecha factura", "date"],
    "comp":    ["comprobante", "no. comprobante", "numero comprobante", "nro comprobante", "doc.", "documento"],
    "total":   ["total", "monto", "valor", "importe", "total factura"],
    "codcli":  ["codigo cliente", "cod. cliente", "cod cliente", "cliente cod", "id cliente"],
    "nomdest": ["nombre destinatario", "destinatario", "nombre cliente", "cliente"],
    "ciudad":  ["ciudad destinatario", "ciudad dest", "ciudad", "destino"],
    "ruta":    ["ruta", "cod ruta", "codigo ruta"],
    "bodega":  ["bodega", "centro sap", "cod bodega", "almacen sap"],
}

DEFAULTS = {"fecha":1,"comp":4,"total":5,"codcli":6,"nomdest":9,"ciudad":10,"ruta":11,"bodega":17}

def match_alias(col_norm, aliases):
    for a in aliases:
        if col_norm == a:
            return True
    for a in aliases:
        if " " in a and a in col_norm:
            return True
    return False

def detectar_columnas(ws):
    ci = dict(DEFAULTS)
    for c in range(1, 31):
        v = ws.cell(5, c).value
        if not v:
            continue
        cn = str(v).strip().lower()
        for field, aliases in ALIASES.items():
            if match_alias(cn, aliases):
                ci[field] = c
    return ci

# ── Procesar DATA ─────────────────────────────────────────────
def procesar_data(wb, centros_nombre, centros_empresa, cli_vendedor, cli_prioridad,
                  ruta_dia, filtro_activo, fact_desde, fact_hasta):
    ws = wb["DATA"]
    last_row = ws.max_row
    if last_row < 6:
        return None, "No hay datos en DATA desde fila 6."

    ci = detectar_columnas(ws)
    errores = []

    df       = defaultdict(list)
    df_monto = defaultdict(float)
    df_fecha = {}
    df_ruta  = {}
    df_bodega= {}
    df_emp   = defaultdict(set)
    df_vend  = defaultdict(set)

    for i in range(6, last_row + 1):
        ciudad = norm(str(ws.cell(i, ci["ciudad"]).value or ""))
        if not ciudad:
            continue

        comp_str = str(ws.cell(i, ci["comp"]).value or "").strip()
        if filtro_activo:
            try:
                comp_num = float(comp_str)
                if comp_num < fact_desde or comp_num > fact_hasta:
                    continue
            except:
                continue

        try:
            monto = float(ws.cell(i, ci["total"]).value or 0)
        except:
            monto = 0.0

        # Bodega como entero
        bod_v = ws.cell(i, ci["bodega"]).value
        try:
            bodega_num = str(int(float(str(bod_v))))
        except:
            bodega_num = str(bod_v or "").strip()

        emp_res   = centros_empresa.get(bodega_num, "")
        centro_nom = centros_nombre.get(bodega_num, "")

        cod_cli = norm(str(ws.cell(i, ci["codcli"]).value or ""))

        if cod_cli in cli_vendedor:
            vendedor  = cli_vendedor[cod_cli]
            prioridad = cli_prioridad[cod_cli]
        else:
            vendedor  = "VENDEDOR NO ASIGNADO"
            prioridad = 9999
            errores.append(f"Fila {i}: Cliente '{cod_cli}' no encontrado en CLIENTES")

        # Fecha: leer serial desde openpyxl (datetime nativo, no afectado por regional)
        serial_fecha, texto_fecha = fecha_de_celda(ws.cell(i, ci["fecha"]))

        if ciudad not in df_fecha:
            df_fecha[ciudad]  = texto_fecha
            df_ruta[ciudad]   = str(ws.cell(i, ci["ruta"]).value or "").strip()
            df_bodega[ciudad] = bodega_num

        if emp_res:
            df_emp[ciudad].add(emp_res)
        if vendedor != "VENDEDOR NO ASIGNADO":
            df_vend[ciudad].add(vendedor)

        df[ciudad].append({
            "comp":      comp_str,
            "nom_dest":  str(ws.cell(i, ci["nomdest"]).value or "").strip(),
            "monto":     monto,
            "fecha_serial": serial_fecha,
            "fecha_txt": texto_fecha,
            "cod_cli":   cod_cli,
            "bodega":    bodega_num,
            "ciudad":    ciudad,
            "vendedor":  vendedor,
            "prioridad": prioridad,
            "centro_nom":centro_nom,
        })
        df_monto[ciudad] += monto

    if not df:
        return None, "No se encontraron registros."

    return {
        "df": df, "df_monto": df_monto, "df_fecha": df_fecha,
        "df_ruta": df_ruta, "df_bodega": df_bodega,
        "df_emp": df_emp, "df_vend": df_vend,
        "errores": errores
    }, None

# ── Generar Excel de salida ───────────────────────────────────
def generar_excel(datos, centros_nombre, ruta_camion, ruta_dia, ruta_fecha):
    wb_out = openpyxl.Workbook()
    ws_res = wb_out.active
    ws_res.title = "RESUMEN"

    df       = datos["df"]
    df_monto = datos["df_monto"]
    df_fecha = datos["df_fecha"]
    df_ruta  = datos["df_ruta"]
    df_bodega= datos["df_bodega"]
    df_emp   = datos["df_emp"]
    df_vend  = datos["df_vend"]

    ciudades = sorted(df.keys())
    total_f  = 0
    total_m  = 0.0
    res_row  = 5

    for ciudad in ciudades:
        registros = df[ciudad]
        # Ordenar por prioridad ASC, luego comprobante ASC
        registros.sort(key=lambda r: (r["prioridad"],
                                      r["comp"].zfill(20) if r["comp"].isdigit() else r["comp"]))
        n_rows = len(registros)
        sh_name = nombre_hoja(ciudad)

        ws_r = wb_out.create_sheet(sh_name)
        fecha_str  = df_fecha.get(ciudad, "--")
        ruta_str   = df_ruta.get(ciudad, "--")
        bodega_num = df_bodega.get(ciudad, "")
        bodega_str = bodega_num
        if bodega_num in centros_nombre and centros_nombre[bodega_num]:
            bodega_str = bodega_num + " -- " + centros_nombre[bodega_num]

        empresas_ruta  = ", ".join(sorted(df_emp[ciudad]))  or "--"
        camion_str     = ruta_camion.get(ciudad, "NO ASIGNADO")
        dia_str        = ruta_dia.get(ciudad, "--")
        fecha_desp_str = ruta_fecha.get(ciudad, "--")

        # ── Anchos columnas exactos ──
        anchos = {"A":11.43,"B":7.43,"C":6.86,"D":11.86,"E":39.14,
                  "F":16.43,"G":11.0,"H":8.86,"I":6.0,"J":31.29}
        for col, w in anchos.items():
            ws_r.column_dimensions[col].width = w

        # ── F1: Titulo ──
        ws_r.merge_cells("A1:J1")
        c = ws_r["A1"]
        c.value = "RUTERO DE DESPACHO"
        c.font  = Font(name="Arial", bold=True, size=13, color=C_BLN)
        c.fill  = PatternFill("solid", fgColor=C_AZO)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = thin_border()
        ws_r.row_dimensions[1].height = 21.95

        # ── F2: EMPRESA ──
        lbl_cab(ws_r, 2, 1, 2, "EMPRESA:")
        val_cab(ws_r, 2, 3, 10, empresas_ruta, bold=True)

        # ── F3: CIUDAD | FECHA FACTURA ──
        lbl_cab(ws_r, 3, 1, 2, "CIUDAD DEST.:")
        val_cab(ws_r, 3, 3, 5, ciudad)
        lbl_cab(ws_r, 3, 6, 7, "FECHA FACTURA:")
        val_cab(ws_r, 3, 8, 10, fecha_str)

        # ── F4: BODEGA ──
        lbl_cab(ws_r, 4, 1, 2, "BODEGA / CENTRO:")
        val_cab(ws_r, 4, 3, 10, bodega_str)

        # ── F5: FECHA ENVIO | CHOFER | PLACA ──
        lbl_cab(ws_r, 5, 1, 1, "FECHA ENVIO:")
        linea_cab(ws_r, 5, 2, 3)
        lbl_cab(ws_r, 5, 4, 4, "CHOFER:")
        linea_cab(ws_r, 5, 5, 7)
        lbl_cab(ws_r, 5, 8, 8, "PLACA:")
        linea_cab(ws_r, 5, 9, 10)

        # ── F6: N FACTURAS | MONTO ──
        lbl_cab(ws_r, 6, 1, 2, "N\u00b0 FACTURAS:")
        val_cab(ws_r, 6, 3, 3, n_rows, bold=True)
        lbl_cab(ws_r, 6, 4, 5, "MONTO TOTAL:")
        val_cab(ws_r, 6, 6, 8, round(df_monto[ciudad], 2), bold=True, fmt="$#,##0.00")
        ws_r.merge_cells("I6:J6")
        ws_r["I6"].fill = PatternFill("solid", fgColor=C_BLN)
        ws_r["I6"].border = thin_border()

        # Alturas cabecera
        for rr in range(2, 7):
            ws_r.row_dimensions[rr].height = 15.0

        # Línea separadora azul en F6
        for col in range(1, 11):
            c = ws_r.cell(6, col)
            old_b = c.border
            c.border = Border(
                left=old_b.left, right=old_b.right, top=old_b.top,
                bottom=Side(style="medium", color=C_AZO)
            )

        # ── F7: Link volver ──
        ws_r.merge_cells("A7:J7")
        c7 = ws_r["A7"]
        c7.value = "<< Volver al Resumen"
        c7.font  = Font(name="Arial", size=8, italic=True, color=C_AZM)
        c7.fill  = PatternFill("solid", fgColor=C_GRS)
        c7.alignment = Alignment(horizontal="left", vertical="center")
        c7.border = thin_border()
        c7.hyperlink = "#RESUMEN!A1"
        ws_r.row_dimensions[7].height = 12.0

        # ── F8: Cabecera tabla ──
        hdrs = ["Seq.","Prior.","Bodega","Comprobante","Nombre Cliente",
                "Agente","Fecha Fact.","Total ($)","Bultos","Observaciones"]
        for fc, h in enumerate(hdrs, 1):
            c = ws_r.cell(8, fc)
            c.value = h
            c.font  = Font(name="Arial", bold=True, size=8, color=C_BLN)
            c.fill  = PatternFill("solid", fgColor=C_AZM)
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = thin_border()
        ws_r.row_dimensions[8].height = 15.95

        # ── Datos (fondo blanco) ──
        data_start = 9
        for idx, row in enumerate(registros):
            dr = data_start + idx
            for fc in range(1, 11):
                c = ws_r.cell(dr, fc)
                c.fill   = PatternFill("solid", fgColor=C_BLN)
                c.font   = Font(name="Arial", size=8, color="323232")
                c.border = thin_border()
                c.alignment = Alignment(
                    horizontal="left" if fc == 5 else "center",
                    vertical="center"
                )

            ws_r.cell(dr, 1).value = idx + 1
            ws_r.cell(dr, 2).value = row["prioridad"]
            ws_r.cell(dr, 3).value = row["bodega"]
            ws_r.cell(dr, 4).value = row["comp"]
            ws_r.cell(dr, 5).value = row["nom_dest"]

            ws_r.cell(dr, 6).value = row["vendedor"]
            if row["vendedor"] == "VENDEDOR NO ASIGNADO":
                ws_r.cell(dr, 6).font = Font(name="Arial", size=8,
                                              color="963232", italic=True)

            # Fecha: escribir serial con NumberFormat DD/MM/YYYY
            serial = row["fecha_serial"]
            if serial and serial > 1:
                ws_r.cell(dr, 7).value = serial
                ws_r.cell(dr, 7).number_format = "DD/MM/YYYY"
            else:
                ws_r.cell(dr, 7).value = row["fecha_txt"]

            ws_r.cell(dr, 8).value = round(row["monto"], 2)
            ws_r.cell(dr, 8).number_format = "#,##0.00"
            ws_r.cell(dr, 10).alignment = Alignment(horizontal="left", vertical="center")

            ws_r.row_dimensions[dr].height = 14.1

            # Color prioridad
            p = row["prioridad"]
            if p == 1:
                ws_r.cell(dr, 2).font = Font(name="Arial", size=8,
                                              color="8B0000", bold=True)
            elif p == 2:
                ws_r.cell(dr, 2).font = Font(name="Arial", size=8,
                                              color="823C00", bold=True)
            else:
                ws_r.cell(dr, 2).font = Font(name="Arial", size=8, color="646464")

        # ── Subtotal ──
        sub_r = data_start + n_rows
        ws_r.merge_cells(
            start_row=sub_r, start_column=1, end_row=sub_r, end_column=7)
        c = ws_r.cell(sub_r, 1)
        c.value = "SUBTOTAL"
        c.font  = Font(name="Arial", bold=True, size=9, color=C_BLN)
        c.fill  = PatternFill("solid", fgColor=C_AZO)
        c.alignment = Alignment(horizontal="right", vertical="center")
        c.border = thin_border()

        ws_r.cell(sub_r, 8).value  = f"=SUM(H{data_start}:H{data_start+n_rows-1})"
        ws_r.cell(sub_r, 8).number_format = "$#,##0.00"
        ws_r.cell(sub_r, 8).font  = Font(name="Arial", bold=True, size=9, color=C_BLN)
        ws_r.cell(sub_r, 8).fill  = PatternFill("solid", fgColor=C_AZO)
        ws_r.cell(sub_r, 8).alignment = Alignment(horizontal="center", vertical="center")
        ws_r.cell(sub_r, 8).border = thin_border()
        for fc in [9, 10]:
            ws_r.cell(sub_r, fc).fill  = PatternFill("solid", fgColor=C_AZO)
            ws_r.cell(sub_r, fc).border = thin_border()
        ws_r.row_dimensions[sub_r].height = 16

        # ── Firmas ──
        fir_r = sub_r + 5
        ws_r.merge_cells(
            start_row=fir_r, start_column=1, end_row=fir_r, end_column=3)
        c = ws_r.cell(fir_r, 1)
        c.value = "Firma Chofer: _______________________"
        c.font  = Font(name="Arial", size=9)
        c.alignment = Alignment(horizontal="left", vertical="center")

        ws_r.merge_cells(
            start_row=fir_r, start_column=8, end_row=fir_r, end_column=10)
        c = ws_r.cell(fir_r, 8)
        c.value = "Firma Receptor: _______________________"
        c.font  = Font(name="Arial", size=9)
        c.alignment = Alignment(horizontal="right", vertical="center")
        ws_r.row_dimensions[fir_r].height = 20

        # ── Impresión ──
        ws_r.page_setup.orientation  = "landscape"
        ws_r.page_setup.paperSize    = 9  # A4
        ws_r.page_setup.fitToWidth   = 1
        ws_r.page_setup.fitToHeight  = 0
        ws_r.sheet_properties.pageSetUpPr.fitToPage = True
        ws_r.page_margins.top    = 0.75
        ws_r.page_margins.bottom = 0.75
        ws_r.page_margins.left   = 0.25
        ws_r.page_margins.right  = 0.25
        ws_r.page_margins.header = 0.3
        ws_r.page_margins.footer = 0.3
        ws_r.print_title_rows    = "1:8"
        ws_r.print_area          = f"A1:J{fir_r}"
        ws_r.oddHeader.center.text = (
            f"&\"Arial,Bold\"&9 {ciudad}  |  DESPACHO: {fecha_desp_str}"
        )
        ws_r.oddFooter.center.text = "&\"Arial\"&8 Página &P de &N"

        total_f += n_rows
        total_m += df_monto[ciudad]

        # Fila en RESUMEN
        bg = C_AZC if res_row % 2 == 0 else C_BLN
        for fc, val in enumerate([ciudad, ruta_str, n_rows,
                                   round(df_monto[ciudad], 2),
                                   camion_str, fecha_desp_str,
                                   empresas_ruta], 1):
            c = ws_res.cell(res_row, fc)
            c.value = val
            c.font  = Font(name="Arial", size=9)
            c.fill  = PatternFill("solid", fgColor=bg)
            c.border = thin_border()
            c.alignment = Alignment(
                horizontal="left" if fc in [1, 7] else "center",
                vertical="center"
            )
            if fc == 4:
                c.number_format = "#,##0.00"

        # Link >> Ver
        link_c = ws_res.cell(res_row, 8)
        link_c.value = ">> Ver"
        link_c.font  = Font(name="Arial", size=9, bold=True, color=C_AZM)
        link_c.fill  = PatternFill("solid", fgColor=bg)
        link_c.border = thin_border()
        link_c.alignment = Alignment(horizontal="center", vertical="center")
        link_c.hyperlink = f"#{sh_name}!A1"
        ws_res.row_dimensions[res_row].height = 16
        res_row += 1

    # ── RESUMEN: cabecera ──
    ws_res.merge_cells("A1:H1")
    c = ws_res["A1"]
    c.value = "RESUMEN DE DESPACHOS"
    c.font  = Font(name="Arial", bold=True, size=13, color=C_BLN)
    c.fill  = PatternFill("solid", fgColor=C_AZO)
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.border = thin_border()
    ws_res.row_dimensions[1].height = 28

    ws_res.merge_cells("A2:H2")
    c = ws_res["A2"]
    c.value = (f"Generado: {datetime.datetime.now().strftime('%d/%m/%Y %H:%M')}"
               f"   |   Facturas: {total_f}"
               f"   |   Monto: ${total_m:,.2f}")
    c.font  = Font(name="Arial", size=9, italic=True)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws_res.row_dimensions[2].height = 16

    hdrs_r = ["CIUDAD","RUTA","N° FACT.","MONTO ($)","CAMION",
              "FECHA DESP.","EMPRESAS","HOJA"]
    widths_r = [24,10,10,14,14,13,34,10]
    for fc, (h, w) in enumerate(zip(hdrs_r, widths_r), 1):
        c = ws_res.cell(4, fc)
        c.value = h
        c.font  = Font(name="Arial", bold=True, size=10, color=C_BLN)
        c.fill  = PatternFill("solid", fgColor=C_AZM)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = thin_border()
        ws_res.column_dimensions[get_column_letter(fc)].width = w
    ws_res.row_dimensions[4].height = 22

    # Total general
    ws_res.merge_cells(
        start_row=res_row, start_column=1, end_row=res_row, end_column=2)
    for fc in range(1, 9):
        c = ws_res.cell(res_row, fc)
        c.fill  = PatternFill("solid", fgColor=C_AZO)
        c.font  = Font(name="Arial", bold=True, size=10, color=C_BLN)
        c.border = thin_border()
        c.alignment = Alignment(horizontal="center", vertical="center")
    ws_res.cell(res_row, 1).value = "TOTAL GENERAL"
    ws_res.cell(res_row, 3).value = total_f
    ws_res.cell(res_row, 4).value = round(total_m, 2)
    ws_res.cell(res_row, 4).number_format = "#,##0.00"
    ws_res.row_dimensions[res_row].height = 20

    return wb_out

# ── Interfaz Streamlit ────────────────────────────────────────
def main():
    st.set_page_config(
        page_title="Rutero de Despacho",
        page_icon="🚚",
        layout="centered"
    )

    # Header
    st.markdown("""
        <div style="background:#1F3864;padding:18px 24px;border-radius:6px;margin-bottom:20px">
            <h2 style="color:white;margin:0;font-family:Arial">🚚 Generador de Rutero de Despacho</h2>
            <p style="color:#D6E4F6;margin:4px 0 0 0;font-size:14px">
                Suba la plantilla Excel con DATA, CENTROS, CLIENTES y RUTAS
            </p>
        </div>
    """, unsafe_allow_html=True)

    # Upload
    uploaded = st.file_uploader(
        "Seleccione la plantilla (.xlsx o .xlsm)",
        type=["xlsx", "xlsm"],
        help="El archivo debe contener las hojas: DATA, CENTROS, CLIENTES, RUTAS"
    )

    if not uploaded:
        st.info("📂 Suba la plantilla para comenzar.")
        st.markdown("""
        **Estructura requerida de la plantilla:**
        | Hoja | Descripción |
        |---|---|
        | DATA | Exportación SAP. Cabecera en fila 5, datos desde fila 6 |
        | CENTROS | Código bodega → Nombre → Empresa. Datos desde fila 4 |
        | CLIENTES | Código cliente → Vendedor → Prioridad. Datos desde fila 4 |
        | RUTAS | Camión → Ciudad → Día despacho → Fecha. Datos desde fila 4 |
        """)
        return

    # Filtro opcional
    with st.expander("⚙️ Filtro por rango de comprobantes (opcional)"):
        col1, col2 = st.columns(2)
        with col1:
            fact_desde = st.number_input("Comprobante desde", min_value=0, value=0, step=1)
        with col2:
            fact_hasta = st.number_input("Comprobante hasta", min_value=0, value=0, step=1)
        filtro_activo = fact_desde > 0 and fact_hasta > 0
        if filtro_activo:
            if fact_desde > fact_hasta:
                st.error("'Desde' no puede ser mayor que 'Hasta'")
                return
            st.success(f"Filtro activo: {int(fact_desde):,} — {int(fact_hasta):,}")

    if st.button("🚀 Generar Rutero", type="primary", use_container_width=True):
        with st.spinner("Procesando..."):
            try:
                wb = openpyxl.load_workbook(uploaded, keep_vba=False, data_only=True)

                # Verificar hojas
                hojas_req = ["DATA", "CENTROS", "CLIENTES", "RUTAS"]
                faltantes = [h for h in hojas_req if h not in wb.sheetnames]
                if faltantes:
                    st.error(f"Faltan hojas: {', '.join(faltantes)}")
                    return

                # Cargar maestros
                centros_nombre, centros_empresa, err = cargar_centros(wb)
                if err:
                    st.error(err); return

                cli_vendedor, cli_prioridad, err = cargar_clientes(wb)
                if err:
                    st.error(err); return

                ruta_camion, ruta_dia, ruta_fecha, err = cargar_rutas(wb)
                if err:
                    st.error(err); return

                # Procesar DATA
                resultado, err = procesar_data(
                    wb, centros_nombre, centros_empresa,
                    cli_vendedor, cli_prioridad, ruta_dia,
                    filtro_activo, fact_desde, fact_hasta
                )
                if err:
                    st.error(err); return

                # Advertir errores de clientes
                errores = resultado.get("errores", [])
                if errores:
                    with st.expander(f"⚠️ {len(errores)} cliente(s) sin agente asignado"):
                        for e in errores[:20]:
                            st.text(e)
                        if len(errores) > 20:
                            st.text(f"... y {len(errores)-20} más")

                # Generar Excel
                wb_out = generar_excel(
                    resultado, centros_nombre,
                    ruta_camion, ruta_dia, ruta_fecha
                )

                # Guardar en buffer
                buf = io.BytesIO()
                wb_out.save(buf)
                buf.seek(0)

                n_ciudades = len(resultado["df"])
                n_facturas = sum(len(v) for v in resultado["df"].values())
                monto_total = sum(resultado["df_monto"].values())

                # Resumen
                st.success("✅ Rutero generado correctamente")
                col1, col2, col3 = st.columns(3)
                col1.metric("Rutas", n_ciudades)
                col2.metric("Facturas", f"{n_facturas:,}")
                col3.metric("Monto Total", f"${monto_total:,.2f}")

                # Botón descarga
                nombre_archivo = (
                    f"DESPACHO_{datetime.datetime.now().strftime('%Y-%m-%d_%H-%M')}.xlsx"
                )
                st.download_button(
                    label="📥 Descargar Rutero Excel",
                    data=buf,
                    file_name=nombre_archivo,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    type="primary"
                )

            except Exception as e:
                st.error(f"Error inesperado: {str(e)}")
                st.exception(e)

if __name__ == "__main__":
    main()
