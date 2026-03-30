#!/usr/bin/env python3
"""
app.py — Streamlit App
Automatizacion de Estados de Cuenta (GBM / Prestadero)
"""

import streamlit as st
import pdfplumber
import pandas as pd
import os
import re
import io
import shutil
import tempfile
from collections import defaultdict
from copy import copy
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ════════════════════════════════════════════════════════════
# CONFIGURACION DE PAGINA
# ════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="Estados de Cuenta - GBM",
    page_icon="📊",
    layout="wide",
)

st.title("📊 Automatizacion de Estados de Cuenta")
st.markdown("**Extrae datos de PDFs (GBM / Prestadero) y consolida con el maestro anterior.**")
st.divider()


# ════════════════════════════════════════════════════════════
# FUNCIONES DE EXTRACCION (del extractor_gbm.py)
# ════════════════════════════════════════════════════════════
def extraer_numeros(texto):
    nums = re.findall(r'[\d,]+\.\d+', texto)
    return [float(n.replace(',', '')) for n in nums]


def extraer_todos_numeros(texto):
    nums = re.findall(r'[\d,]+\.?\d*', texto)
    return [float(n.replace(',', '')) for n in nums if n]


def extraer_numero_despues_de(texto, clave):
    idx = texto.find(clave)
    if idx == -1:
        return None
    sub = texto[idx + len(clave):]
    nums = extraer_numeros(sub)
    return nums[0] if nums else None


def extraer_nombre_cliente(pdf, plataforma):
    texto = pdf.pages[0].extract_text() or ""
    lineas = texto.split('\n')
    if plataforma == "Prestadero":
        for linea in lineas:
            if "Periodo:" in linea and "Estado de Cuenta" not in linea:
                return re.split(r'\s+Periodo:', linea)[0].strip().upper()
    else:
        for linea in lineas:
            if "Contrato:" in linea:
                parte = re.split(r'\s+Contrato:', linea)[0].strip()
                parte = parte.replace("PUBLICO EN GENERAL - ", "")
                return parte.upper()
    return "DESCONOCIDO"


def detectar_plataforma(texto):
    return "Prestadero" if ("Prestadero" in texto or "PRESTADERO" in texto) else "GBM"


def es_smart_cash(texto):
    for linea in texto.split('\n'):
        if "RENTA VARIABLE" in linea and "VALORES EN CORTO" not in linea:
            nums = extraer_numeros(linea)
            if len(nums) >= 2 and nums[1] > 0:
                return False
    return True


def extraer_saldo_anterior(lineas):
    for l in lineas:
        if "VALOR DEL PORTAFOLIO" in l and "TOTAL" not in l:
            nums = extraer_numeros(l)
            if nums:
                return nums[0]
    return 0.0


def extraer_portafolio_gbm(pdf):
    portafolio = []
    en_desglose = en_acciones = False
    for pag in pdf.pages:
        texto = pag.extract_text() or ""
        for l in texto.split("\n"):
            ls, lu = l.strip(), l.strip().upper()
            if "DESGLOSE DEL PORTAFOLIO" in lu:
                en_desglose = True; continue
            if not en_desglose: continue
            if lu == "ACCIONES":
                en_acciones = True; continue
            if en_acciones and ("EMISORA" in lu or "MES ANTERIOR" in lu or "EN PR" in lu):
                continue
            if en_acciones and lu.startswith("TOTAL"):
                en_acciones = False; continue
            if lu in ("DESGLOSE DE MOVIMIENTOS", "RENDIMIENTO DEL PORTAFOLIO", "EFECTIVO"):
                en_desglose = en_acciones = False; continue
            if not en_acciones: continue
            m = re.match(r'^([A-Z]+(?:\s+\d+)?)\s+', ls)
            if m:
                try:
                    emisora = m.group(1).strip()
                    nums = extraer_todos_numeros(ls[m.end():])
                    if len(nums) >= 8:
                        portafolio.append({
                            "Emisora": emisora,
                            "Títulos Mes Anterior": int(nums[0]),
                            "Títulos Mes Actual": int(nums[1]),
                            "Costo Total": nums[4],
                            "Precio Mercado Mes Anterior": nums[5],
                            "Precio Mercado Mes Actual": nums[6],
                            "Valor a Mercado": nums[7],
                        })
                except (ValueError, IndexError):
                    continue
    return portafolio


def extraer_deuda_gbm(pdf):
    deuda = []
    en_desglose = en_deuda = False
    for pag in pdf.pages:
        texto = pag.extract_text() or ""
        for l in texto.split("\n"):
            ls, lu = l.strip(), l.strip().upper()
            if "DESGLOSE DEL PORTAFOLIO" in lu:
                en_desglose = True; continue
            if not en_desglose: continue
            if "DEUDA EN REPORTO" in lu and "TOTAL" not in lu:
                en_deuda = True; continue
            if en_deuda and ("EMISORA" in lu or "ANTERIOR" in lu):
                continue
            if en_deuda and lu.startswith("TOTAL"):
                en_deuda = False; continue
            if lu in ("RENTA VARIABLE", "EFECTIVO", "DESGLOSE DE MOVIMIENTOS", "RENDIMIENTO DEL PORTAFOLIO"):
                en_desglose = en_deuda = False; continue
            if not en_deuda: continue
            m = re.match(r'^([A-Z]+\s+\d+)\s+', ls)
            if m:
                try:
                    emisora = m.group(1).strip()
                    nums = extraer_todos_numeros(ls[m.end():])
                    if len(nums) >= 8:
                        deuda.append({
                            "Emisora": emisora,
                            "Títulos Mes Anterior": int(nums[0]),
                            "Títulos Mes Actual": int(nums[1]),
                            "Tasa": nums[2],
                            "Valor del Reporto": nums[7],
                            "% Cartera": nums[9] if len(nums) >= 10 else 0.0,
                        })
                except (ValueError, IndexError):
                    continue
    return deuda


def extraer_movimientos_acciones(pdf):
    movimientos = []
    en_mov = False
    for pag in pdf.pages:
        texto = pag.extract_text() or ""
        for l in texto.split("\n"):
            lu = l.strip().upper()
            if "DESGLOSE DE MOVIMIENTOS" in lu:
                en_mov = True; continue
            if en_mov and lu in ("RENDIMIENTO DEL PORTAFOLIO", "COMPOSICIÓN FISCAL INFORMATIVA"):
                en_mov = False; continue
            if not en_mov: continue
            if "Compra de Acciones" not in l and "Venta de Acciones" not in l:
                continue
            try:
                ls = l.strip()
                if "Compra de Acciones" in l:
                    op = "Compra"
                    resto = l[l.find("Compra de Acciones.") + len("Compra de Acciones."):].strip()
                else:
                    op = "Venta"
                    resto = l[l.find("Venta de Acciones.") + len("Venta de Acciones."):].strip()
                fm = re.match(r'(\d{2}/\d{2})', ls)
                fecha = fm.group(1) if fm else ""
                em = re.match(r'^([A-Z]+(?:\s+\d+)?)\s+', resto)
                if em:
                    emisora = em.group(1).strip()
                    nums = extraer_todos_numeros(resto[em.end():])
                    movimientos.append({
                        "Fecha": fecha, "Operación": op, "Emisora": emisora,
                        "Títulos": int(nums[0]) if nums else 0,
                        "Precio Unitario": nums[1] if len(nums) >= 2 else 0,
                        "Comisión": nums[2] if len(nums) >= 3 else 0,
                        "Neto": nums[5] if len(nums) >= 6 else 0,
                    })
            except Exception:
                continue
    return movimientos


def extraer_movimientos_efectivo_smart_cash(pdf):
    movimientos = []
    en_mov = False
    for pag in pdf.pages:
        texto = pag.extract_text() or ""
        for l in texto.split("\n"):
            lu = l.strip().upper()
            if "DESGLOSE DE MOVIMIENTOS" in lu:
                en_mov = True; continue
            if en_mov and lu in ("RENDIMIENTO DEL PORTAFOLIO", "COMPOSICIÓN FISCAL INFORMATIVA"):
                en_mov = False; continue
            if not en_mov: continue
            if "DEPOSITO" not in lu and "RETIRO" not in lu:
                continue
            try:
                ls = l.strip()
                fecha_match = re.match(r'(\d{2}/\d{2})', ls)
                fecha = fecha_match.group(1) if fecha_match else ""
                operacion = "Depósito" if "DEPOSITO" in lu else "Retiro"
                nums = extraer_numeros(ls)
                monto = nums[-2] if len(nums) >= 2 else (nums[0] if nums else 0.0)
                movimientos.append({"Fecha": fecha, "Operación": operacion, "Monto": monto})
            except Exception:
                continue
    return movimientos


# ════════════════════════════════════════════════════════════
# FORMATEO EXCEL (para el reporte de extraccion)
# ════════════════════════════════════════════════════════════
TITULO_FONT = Font(bold=True, size=12, color="FFFFFF")
TITULO_FILL = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
TOTAL_FONT = Font(bold=True, size=10)
TOTAL_FILL = PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid")


def escribir_seccion(writer, hoja, fila, titulo, df):
    ws = writer.sheets[hoja]
    celda = ws.cell(row=fila + 1, column=1, value=titulo)
    celda.font = TITULO_FONT
    celda.fill = TITULO_FILL
    celda.alignment = Alignment(horizontal="left")
    for c in range(2, len(df.columns) + 1):
        ws.cell(row=fila + 1, column=c).fill = TITULO_FILL
    df.to_excel(writer, sheet_name=hoja, startrow=fila + 1, index=False)
    return fila + len(df) + 3


def escribir_fila_total(writer, hoja, fila, label, valor):
    ws = writer.sheets[hoja]
    c1 = ws.cell(row=fila + 1, column=1, value=label)
    c1.font = TOTAL_FONT; c1.fill = TOTAL_FILL
    c2 = ws.cell(row=fila + 1, column=2, value=valor)
    c2.font = TOTAL_FONT; c2.fill = TOTAL_FILL; c2.number_format = '#,##0.00'
    return fila + 2


def ajustar_columnas(ws):
    for col in ws.columns:
        max_len = 0
        col_letter = col[0].column_letter
        for cell in col:
            if cell.value:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = min(max_len + 3, 35)


# ════════════════════════════════════════════════════════════
# PROCESAR PDFs SUBIDOS
# ════════════════════════════════════════════════════════════
def procesar_pdfs(archivos_pdf):
    """Procesa una lista de archivos PDF subidos y retorna datos por cliente."""
    clientes = defaultdict(lambda: {"gbm": None, "smart_cash": None, "prestadero": None})
    log = []

    for archivo in archivos_pdf:
        nombre_archivo = archivo.name
        log.append(f"📄 Procesando: {nombre_archivo}")

        try:
            with pdfplumber.open(archivo) as pdf:
                texto_p1 = pdf.pages[0].extract_text() or ""
                texto_completo = texto_p1
                for p in pdf.pages[1:]:
                    t = p.extract_text()
                    if t:
                        texto_completo += "\n" + t

                plataforma = detectar_plataforma(texto_completo)
                nombre = extraer_nombre_cliente(pdf, plataforma)
                lineas_p1 = texto_p1.split('\n')

                log.append(f"   → Cliente: {nombre} | {plataforma}")

                # Prestadero
                if plataforma == "Prestadero":
                    abonos = retiros = interes = valor = 0.0
                    for l in lineas_p1:
                        try:
                            if "Abonos:" in l and "Cuenta Abonos:" not in l:
                                v = extraer_numero_despues_de(l, "Abonos:")
                                if v is not None: abonos = v
                            if "Valor de la Cuenta:" in l:
                                v = extraer_numero_despues_de(l, "Valor de la Cuenta:")
                                if v is not None: valor = v
                            if "Interés Recibido" in l or "Interes Recibido" in l:
                                ns = extraer_numeros(l)
                                if ns: interes = ns[0]
                            if "Retiros:" in l and "Detalle" not in l:
                                v = extraer_numero_despues_de(l, "Retiros:")
                                if v is not None: retiros = v
                        except Exception:
                            continue
                    clientes[nombre]["prestadero"] = {
                        "resumen": pd.DataFrame([{
                            "Plataforma": "Prestadero", "Abonos": abonos,
                            "Retiros": retiros, "Interés Recibido": interes,
                            "Valor de la Cuenta": valor,
                        }]),
                        "abonos": abonos, "retiros": retiros,
                        "interes": interes, "valor": valor,
                    }

                # GBM
                else:
                    entradas = salidas = valor_total = saldo_ant = 0.0
                    saldo_ant = extraer_saldo_anterior(lineas_p1)
                    for l in lineas_p1:
                        try:
                            if "ENTRADAS DE EFECTIVO" in l:
                                ns = extraer_numeros(l)
                                if ns: entradas = ns[-1]
                            elif "SALIDAS DE EFECTIVO" in l:
                                ns = extraer_numeros(l)
                                if ns: salidas = ns[-1]
                            elif "VALOR DEL PORTAFOLIO" in l and "TOTAL" not in l:
                                ns = extraer_numeros(l)
                                if len(ns) >= 2: valor_total = ns[1]
                                elif ns: valor_total = ns[0]
                        except Exception:
                            continue

                    smart = es_smart_cash(texto_p1)
                    portafolio = deuda = movimientos = []
                    movimientos_efectivo = []
                    try: deuda = extraer_deuda_gbm(pdf)
                    except Exception: pass
                    if smart:
                        try: movimientos_efectivo = extraer_movimientos_efectivo_smart_cash(pdf)
                        except Exception: pass
                    else:
                        try: portafolio = extraer_portafolio_gbm(pdf)
                        except Exception: pass
                        try: movimientos = extraer_movimientos_acciones(pdf)
                        except Exception: pass

                    tipo = "smart_cash" if smart else "gbm"
                    log.append(f"   → Tipo: {'Smart Cash' if smart else 'GBM Regular'}")

                    clientes[nombre][tipo] = {
                        "resumen": pd.DataFrame([{
                            "Plataforma": "GBM Smart Cash" if smart else "GBM",
                            "Saldo Anterior": saldo_ant,
                            "Entradas de Efectivo": entradas,
                            "Salidas de Efectivo": salidas,
                            "Valor Total del Portafolio": valor_total,
                        }]),
                        "entradas": entradas, "salidas": salidas,
                        "valor_total": valor_total, "saldo_anterior": saldo_ant,
                        "portafolio": portafolio, "deuda": deuda,
                        "movimientos": movimientos,
                        "movimientos_efectivo": movimientos_efectivo,
                    }
        except Exception as e:
            log.append(f"   ❌ Error: {e}")

    return clientes, log


# ════════════════════════════════════════════════════════════
# GENERAR EXCEL DE EXTRACCION
# ════════════════════════════════════════════════════════════
def generar_excel_extraccion(clientes):
    """Genera el Reporte_Consolidado.xlsx en memoria."""
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for nombre_cliente, cuentas in sorted(clientes.items()):
            hoja = nombre_cliente[:31]
            pd.DataFrame().to_excel(writer, sheet_name=hoja, index=False)
            fila = 0

            # GBM Regular
            if cuentas["gbm"]:
                gbm = cuentas["gbm"]
                fila = escribir_seccion(writer, hoja, fila, "RESUMEN GBM", gbm["resumen"])
                if gbm.get("portafolio"):
                    df = pd.DataFrame(gbm["portafolio"])
                    fila = escribir_seccion(writer, hoja, fila, "PORTAFOLIO (ACCIONES / FIBRAS)", df)
                if gbm.get("deuda"):
                    df = pd.DataFrame(gbm["deuda"])
                    fila = escribir_seccion(writer, hoja, fila, "DEUDA", df)
                if gbm.get("movimientos"):
                    df_mov = pd.DataFrame(gbm["movimientos"])
                    fila = escribir_seccion(writer, hoja, fila, "COMPRA / VENTA DE ACCIONES", df_mov)
                    compras = df_mov[df_mov["Operación"] == "Compra"]
                    ventas = df_mov[df_mov["Operación"] == "Venta"]
                    fila = escribir_fila_total(writer, hoja, fila, "COMPRA TOTAL",
                                               compras["Neto"].sum() if len(compras) > 0 else 0.0)
                    fila = escribir_fila_total(writer, hoja, fila, "VENTA TOTAL",
                                               ventas["Neto"].sum() if len(ventas) > 0 else 0.0)
                    fila += 1

            # Smart Cash
            if cuentas["smart_cash"]:
                sc = cuentas["smart_cash"]
                fila = escribir_seccion(writer, hoja, fila, "RESUMEN SMART CASH", sc["resumen"])
                if sc.get("deuda"):
                    df = pd.DataFrame(sc["deuda"])
                    fila = escribir_seccion(writer, hoja, fila, "DEUDA SMART CASH", df)
                if sc.get("movimientos_efectivo"):
                    df_efec = pd.DataFrame(sc["movimientos_efectivo"])
                    fila = escribir_seccion(writer, hoja, fila, "MOVIMIENTOS EFECTIVO SMART CASH", df_efec)
                    depositos = df_efec[df_efec["Operación"] == "Depósito"]
                    retiros = df_efec[df_efec["Operación"] == "Retiro"]
                    fila = escribir_fila_total(writer, hoja, fila, "TOTAL DEPOSITOS",
                                               depositos["Monto"].sum() if len(depositos) > 0 else 0.0)
                    fila = escribir_fila_total(writer, hoja, fila, "TOTAL RETIROS",
                                               retiros["Monto"].sum() if len(retiros) > 0 else 0.0)
                    fila += 1

            # Prestadero
            if cuentas["prestadero"]:
                fila = escribir_seccion(writer, hoja, fila, "RESUMEN PRESTADERO",
                                        cuentas["prestadero"]["resumen"])

            ajustar_columnas(writer.sheets[hoja])

    output.seek(0)
    return output


# ════════════════════════════════════════════════════════════
# FUNCIONES DE CONSOLIDACION (del consolidador.py)
# ════════════════════════════════════════════════════════════
MESES = {
    "ENERO": 1, "FEBRERO": 2, "MARZO": 3, "ABRIL": 4,
    "MAYO": 5, "JUNIO": 6, "JULIO": 7, "AGOSTO": 8,
    "SEPTIEMBRE": 9, "OCTUBRE": 10, "NOVIEMBRE": 11, "DICIEMBRE": 12,
}
MESES_INV = {v: k for k, v in MESES.items()}

ALIASES = {
    "FIBRAPL 14": ["FIBRA PL 14", "FIBRA PL14", "FIBRAPL14"],
    "FIHO 12": ["FIHO12"], "FMTY 14": ["FMTY14"],
    "FUNO 11": ["FUNO11"], "FIBRAMQ 12": ["FIBRAMQ12"],
    "DAHANOS 13": ["DANHOS 13", "DANHOS13", "DAHANOS13"],
}
_ALIAS_MAP = {}
for canonical, alts in ALIASES.items():
    group = {canonical.upper().strip()} | {a.upper().strip() for a in alts}
    for name in group:
        _ALIAS_MAP[name] = group


def normalizar(nombre):
    if not nombre or nombre == "-": return ""
    n = str(nombre).upper().strip()
    n = re.sub(r"\n", " ", n)
    n = re.sub(r"\s+", " ", n)
    return n


def instrumentos_coinciden(np_, nm_):
    np_ = normalizar(np_)
    nm_ = normalizar(nm_)
    if np_ == nm_: return True
    grupo = _ALIAS_MAP.get(np_, {np_})
    return nm_ in grupo


def encontrar_fila(ws, texto, col=1, rango=(1, 50)):
    for r in range(rango[0], rango[1]):
        v = ws.cell(r, col).value
        if v and texto in str(v).upper():
            return r
    return None


def valor_numerico(v, default=0.0):
    return float(v) if isinstance(v, (int, float)) else default


def actualizar_celda(ws, row, col, value, forzar=False):
    target_row, target_col = row, col
    for rng in ws.merged_cells.ranges:
        if rng.min_row <= row <= rng.max_row and rng.min_col <= col <= rng.max_col:
            target_row, target_col = rng.min_row, rng.min_col
            break
    celda = ws.cell(target_row, target_col)
    if not forzar and isinstance(celda.value, str) and celda.value.startswith("="):
        return
    celda.value = value


def copiar_formato_fila(ws, fila_origen, fila_destino):
    for col in range(1, 16):
        src = ws.cell(fila_origen, col)
        dst = ws.cell(fila_destino, col)
        dst.font = copy(src.font); dst.fill = copy(src.fill)
        dst.border = copy(src.border); dst.number_format = src.number_format
        dst.alignment = copy(src.alignment)


def insertar_instrumento(ws, fila_totales, datos, fila_ref, periodo):
    ws.insert_rows(fila_totales)
    nueva_fila = fila_totales
    copiar_formato_fila(ws, fila_ref, nueva_fila)
    emisora = datos["emisora"]
    valor = datos["valor_a_mercado"]
    compra = datos["compra_neto"]
    venta = datos["venta_neto"]
    b = compra if compra > 0 else valor
    c = valor
    e = c - b
    f = (e / b) if b > 0 else 0.0
    g = c - 0 + venta - compra
    h = (g / b) if b > 0 else 0.0
    ws.cell(nueva_fila, 1).value = emisora
    ws.cell(nueva_fila, 2).value = round(b, 2)
    ws.cell(nueva_fila, 3).value = round(c, 2)
    if periodo:
        ws.cell(nueva_fila, 4).value = f"{periodo['mes_nombre'].lower()} {periodo['anio']}"
    ws.cell(nueva_fila, 5).value = round(e, 2)
    ws.cell(nueva_fila, 6).value = round(f, 10)
    ws.cell(nueva_fila, 7).value = round(g, 2)
    ws.cell(nueva_fila, 8).value = round(h, 10)
    ws.cell(nueva_fila, 9).value = "RENTA VARIABLE"
    ws.cell(nueva_fila, 10).value = round(venta, 2)
    ws.cell(nueva_fila, 11).value = round(compra, 2)
    ws.cell(nueva_fila, 14).value = round(c, 2)
    ws.cell(nueva_fila, 15).value = "GBM"
    return fila_totales + 1


def expandir_formulas_totales(ws, fila_totales):
    patron = re.compile(r"(SUM\([A-Z]+)(\d+)(:[A-Z]+)(\d+)(\))")
    nueva_fin = fila_totales - 1
    for col in range(2, 16):
        celda = ws.cell(fila_totales, col)
        val = celda.value
        if not isinstance(val, str) or not val.startswith("="):
            continue
        nueva = patron.sub(
            lambda m: f"{m.group(1)}{m.group(2)}{m.group(3)}{nueva_fin}{m.group(5)}", val)
        if nueva != val:
            celda.value = nueva


def leer_instrumentos_master(ws, fila_header, fila_totales):
    instrumentos = []
    r = fila_header + 1
    while r < fila_totales:
        nombre = ws.cell(r, 1).value
        if nombre and str(nombre).strip() and str(nombre).strip() != "-":
            fila_fin = r
            for rng in ws.merged_cells.ranges:
                if rng.min_row == r and rng.min_col == 1:
                    fila_fin = rng.max_row; break
            instrumentos.append({
                "fila": r, "fila_fin": fila_fin,
                "nombre": str(nombre).strip(),
                "B": ws.cell(r, 2).value, "C": ws.cell(r, 3).value,
                "D": ws.cell(r, 4).value, "E": ws.cell(r, 5).value,
                "F": ws.cell(r, 6).value, "G": ws.cell(r, 7).value,
                "H": ws.cell(r, 8).value, "I": ws.cell(r, 9).value,
                "J": ws.cell(r, 10).value, "K": ws.cell(r, 11).value,
                "L": ws.cell(r, 12).value, "M": ws.cell(r, 13).value,
                "N": ws.cell(r, 14).value, "O": ws.cell(r, 15).value,
            })
            r = fila_fin + 1
        else:
            r += 1
    return instrumentos


def extraer_periodo_pdf_from_uploaded(pdf, plataforma):
    texto = pdf.pages[0].extract_text() or ""
    if plataforma == "Prestadero":
        m = re.search(r"(\d{4}).(\d{2}).(\d{2})\s+al\s+(\d{4}).(\d{2}).(\d{2})", texto)
        if m:
            anio, mes = int(m.group(4)), int(m.group(5))
            dia_fin = int(m.group(6))
            nombre_mes = MESES_INV.get(mes, str(mes))
            return {"mes": mes, "anio": anio, "mes_nombre": nombre_mes,
                    "periodo": f"01-{dia_fin} {nombre_mes[:3]} DE {anio}"}
    else:
        m = re.search(r"DEL\s+(\d+)\s+AL\s+(\d+)\s+DE\s+(\w+)\s+DE\s+(\d{4})", texto, re.I)
        if m:
            dia_ini, dia_fin = int(m.group(1)), int(m.group(2))
            nombre_mes = m.group(3).upper()
            anio = int(m.group(4))
            mes = MESES.get(nombre_mes, 0)
            return {"mes": mes, "anio": anio, "mes_nombre": nombre_mes,
                    "periodo": f"{dia_ini:02d}-{dia_fin} {nombre_mes[:3]} DE {anio}"}
    return None


def _mejor_match_deuda(old_c, fuentes):
    if not fuentes: return None, None
    if old_c <= 0:
        k = min(fuentes, key=lambda k: fuentes[k]["valor"])
        return k, fuentes[k]
    mejor_key = min(fuentes, key=lambda k: abs(fuentes[k]["valor"] - old_c))
    return mejor_key, fuentes[mejor_key]


def buscar_hoja(wb, nombre_cliente):
    nc = nombre_cliente.upper().strip()
    for s in wb.sheetnames:
        if s.upper().strip() == nc: return s
    for s in wb.sheetnames:
        if nc in s.upper() or s.upper().strip() in nc: return s
    partes = nc.split()
    for s in wb.sheetnames:
        su = s.upper()
        if sum(1 for p in partes if p in su) >= 2: return s
    return None


def actualizar_hoja(ws, datos, nombre_hoja, log):
    gbm = datos.get("gbm")
    smart_cash = datos.get("smart_cash")
    prestadero = datos.get("prestadero")
    periodo = datos.get("periodo")

    fila_header = encontrar_fila(ws, "INSTRUMENTO") or 23
    fila_totales = encontrar_fila(ws, "TOTALES", rango=(fila_header, fila_header + 40))
    if not fila_totales:
        log.append(f"   ⚠️ No se encontro TOTALES en '{nombre_hoja}'")
        return

    instrumentos = leer_instrumentos_master(ws, fila_header, fila_totales)
    log.append(f"   📊 Instrumentos: {len(instrumentos)}")

    pdf_port = {}
    compras_map = defaultdict(float)
    ventas_map = defaultdict(float)

    if gbm:
        for item in gbm.get("portafolio", []):
            key = normalizar(item["Emisora"])
            pdf_port[key] = item["Valor a Mercado"]
        for mov in gbm.get("movimientos", []):
            key = normalizar(mov["Emisora"])
            if mov["Operación"] == "Compra":
                compras_map[key] += mov["Neto"]
            else:
                ventas_map[key] += mov["Neto"]

    deuda_gbm_total = sum(d["Valor del Reporto"] for d in gbm.get("deuda", [])) if gbm else 0.0
    deuda_sc_total = sum(d["Valor del Reporto"] for d in smart_cash.get("deuda", [])) if smart_cash else 0.0
    sc_entradas = smart_cash.get("entradas", 0) if smart_cash else 0
    sc_salidas = smart_cash.get("salidas", 0) if smart_cash else 0

    fuentes_deuda = {}
    if prestadero:
        fuentes_deuda["prestadero"] = {
            "valor": prestadero["valor"], "retiros": prestadero["retiros"],
            "depositos": prestadero["abonos"], "interes": prestadero["interes"],
            "tipo": "prestadero",
        }
    if smart_cash and deuda_sc_total > 0:
        fuentes_deuda["smart_cash"] = {
            "valor": deuda_sc_total, "retiros": sc_salidas,
            "depositos": sc_entradas, "interes": 0.0, "tipo": "smart_cash",
        }
    if gbm and deuda_gbm_total > 0:
        fuentes_deuda["gbm_deuda"] = {
            "valor": deuda_gbm_total, "retiros": 0.0,
            "depositos": 0.0, "interes": 0.0, "tipo": "gbm_deuda",
        }

    matched_pdf_keys = set()
    deuda_gbm_matched = 0.0
    efectivo_instr = None

    for instr in instrumentos:
        fila = instr["fila"]
        nom = instr["nombre"]
        nom_n = normalizar(nom)
        old_c = valor_numerico(instr["C"])
        old_b = valor_numerico(instr["B"])
        matched = False
        new_c = new_g = None
        new_j = new_k = 0.0

        if "EFECTIVO" in nom_n and "GBM" in nom_n:
            efectivo_instr = instr; continue

        elif "DEUDA" in nom_n and instr["O"]:
            prov = normalizar(str(instr["O"]))
            fuente_key = None
            if "PRESTADERO" in prov and "prestadero" in fuentes_deuda:
                fuente_key = "prestadero"
            elif "SMART" in prov and "smart_cash" in fuentes_deuda:
                fuente_key = "smart_cash"
            elif "GBM" in prov and "gbm_deuda" in fuentes_deuda:
                fuente_key = "gbm_deuda"
            if fuente_key:
                mf = fuentes_deuda[fuente_key]
                new_c = mf["valor"]; new_j = mf["retiros"]; new_k = mf["depositos"]
                new_g = mf["interes"] if mf["tipo"] == "prestadero" else new_c - old_c
                matched = True
                if mf["tipo"] == "gbm_deuda": deuda_gbm_matched += new_c
                log.append(f"   🔗 DEUDA (fila {fila}) <- {fuente_key.upper()} ${new_c:,.2f}")
                del fuentes_deuda[fuente_key]

        elif "DEUDA" in nom_n and not instr["O"] and fuentes_deuda:
            mk, mf = _mejor_match_deuda(old_c, fuentes_deuda)
            if mf:
                new_c = mf["valor"]; new_j = mf["retiros"]; new_k = mf["depositos"]
                new_g = mf["interes"] if mf["tipo"] == "prestadero" else new_c - old_c
                matched = True
                if mf["tipo"] == "gbm_deuda": deuda_gbm_matched += new_c
                del fuentes_deuda[mk]

        else:
            for pdf_key, pdf_valor in pdf_port.items():
                if instrumentos_coinciden(pdf_key, nom_n):
                    new_c = pdf_valor
                    compra_neto = compras_map.get(pdf_key, 0)
                    venta_neto = ventas_map.get(pdf_key, 0)
                    new_j = venta_neto; new_k = compra_neto
                    new_g = new_c - old_c + new_j - new_k
                    matched = True; matched_pdf_keys.add(pdf_key)
                    break

        if not matched: continue

        if "DEUDA" not in nom_n:
            new_b = old_b + new_k - new_j
            actualizar_celda(ws, fila, 2, round(new_b, 2))
            old_b = new_b

        actualizar_celda(ws, fila, 3, round(new_c, 2))
        new_e = new_c - old_b
        actualizar_celda(ws, fila, 5, round(new_e, 2))
        new_f = (new_e / old_b) if old_b > 0 else 0.0
        actualizar_celda(ws, fila, 6, round(new_f, 10))
        if new_g is not None:
            actualizar_celda(ws, fila, 7, round(new_g, 2))
        if new_g is not None and old_b > 0:
            actualizar_celda(ws, fila, 8, round(new_g / old_b, 10))
        actualizar_celda(ws, fila, 10, round(new_j, 2))
        actualizar_celda(ws, fila, 11, round(new_k, 2))
        actualizar_celda(ws, fila, 14, round(new_c, 2))
        log.append(f"   ✅ {nom:<20s} C=${new_c:>12,.2f}")

    # Efectivo GBM
    if efectivo_instr and gbm:
        fila = efectivo_instr["fila"]
        old_c = valor_numerico(efectivo_instr["C"])
        new_c = round(gbm["valor_total"] - sum(pdf_port.values()) - deuda_gbm_matched, 2)
        if new_c < 0: new_c = 0.0
        new_g = new_c - old_c
        actualizar_celda(ws, fila, 3, round(new_c, 2))
        actualizar_celda(ws, fila, 5, "-"); actualizar_celda(ws, fila, 6, "-")
        actualizar_celda(ws, fila, 7, round(new_g, 2)); actualizar_celda(ws, fila, 8, "-")
        actualizar_celda(ws, fila, 10, 0.0); actualizar_celda(ws, fila, 11, 0.0)
        actualizar_celda(ws, fila, 14, round(new_c, 2))
        log.append(f"   ✅ {'EFECTIVO GBM':<20s} C=${new_c:>12,.2f}")

    # Nuevos instrumentos
    nuevos = []
    for pk, pv in pdf_port.items():
        if pk not in matched_pdf_keys:
            c = compras_map.get(pk, 0); v = ventas_map.get(pk, 0)
            if pv == 0 and c == 0 and v == 0: continue
            nuevos.append({"emisora": pk, "valor_a_mercado": pv, "compra_neto": c, "venta_neto": v})
    if nuevos:
        fi = efectivo_instr["fila"] if efectivo_instr else fila_totales
        fr = instrumentos[-1]["fila"] if instrumentos else fila_header + 1
        for n in nuevos:
            insertar_instrumento(ws, fi, n, fr, periodo)
            fi += 1; fila_totales += 1
            log.append(f"   ➕ NUEVO: {n['emisora']}")
        expandir_formulas_totales(ws, fila_totales)

    # Limpiar debajo de TOTALES
    ultima = ws.max_row
    if ultima > fila_totales:
        ws.delete_rows(fila_totales + 1, ultima - fila_totales)

    if periodo:
        actualizar_celda(ws, 2, 9, f"CORTE MENSUAL {periodo['mes_nombre']}", forzar=True)
        actualizar_celda(ws, 3, 9, periodo["periodo"], forzar=True)


def consolidar_con_maestro(clientes, maestro_file, pdfs_files):
    """Consolida datos extraidos con el maestro. Retorna (bytes, nombre, log)."""
    log = []

    # Extraer periodos de los PDFs
    for archivo in pdfs_files:
        try:
            archivo.seek(0)
            with pdfplumber.open(archivo) as pdf:
                texto = pdf.pages[0].extract_text() or ""
                plataforma = detectar_plataforma(texto)
                nombre = extraer_nombre_cliente(pdf, plataforma)
                periodo = extraer_periodo_pdf_from_uploaded(pdf, plataforma)
                if nombre.upper() in {k.upper() for k in clientes}:
                    for k in clientes:
                        if k.upper() == nombre.upper():
                            if periodo and (clientes[k].get("periodo") is None
                                            or periodo["mes"] > clientes[k].get("periodo", {}).get("mes", 0)):
                                clientes[k]["periodo"] = periodo
        except Exception:
            continue

    # Determinar nombre salida
    any_periodo = None
    for c in clientes.values():
        if c.get("periodo"):
            any_periodo = c["periodo"]; break
    mes_nombre = any_periodo["mes_nombre"] if any_periodo else "ACTUALIZADO"
    anio = any_periodo["anio"] if any_periodo else ""
    salida_nombre = f"ESTADOS DE CUENTA {mes_nombre} {anio}.xlsx".strip()

    # Cargar maestro
    maestro_file.seek(0)
    wb = load_workbook(maestro_file)
    log.append(f"📂 Hojas en maestro: {wb.sheetnames}")

    clientes_ok = 0
    for nombre_cliente, datos in sorted(clientes.items()):
        hoja = buscar_hoja(wb, nombre_cliente)
        if hoja:
            log.append(f"\n🔄 {nombre_cliente} -> Hoja: '{hoja}'")
            try:
                actualizar_hoja(wb[hoja], datos, hoja, log)
                clientes_ok += 1
            except Exception as e:
                log.append(f"   ❌ Error: {e}")
        else:
            log.append(f"\n⚠️ {nombre_cliente} — Sin hoja en maestro")

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    log.append(f"\n✅ Clientes actualizados: {clientes_ok}")
    return output, salida_nombre, log


# ════════════════════════════════════════════════════════════
# INTERFAZ STREAMLIT
# ════════════════════════════════════════════════════════════
tab1, tab2 = st.tabs(["📄 Paso 1: Extraer datos de PDFs", "📊 Paso 2: Consolidar con Maestro"])

# ─── TAB 1: EXTRACCION ───
with tab1:
    st.header("Extraer datos de PDFs")
    st.markdown("Sube los estados de cuenta en PDF (GBM y/o Prestadero) para extraer los datos.")

    pdfs = st.file_uploader(
        "Selecciona los PDFs",
        type=["pdf"],
        accept_multiple_files=True,
        key="pdfs_extraccion",
    )

    if pdfs and st.button("🚀 Extraer datos", key="btn_extraer", type="primary"):
        with st.spinner("Procesando PDFs..."):
            clientes, log = procesar_pdfs(pdfs)
            st.session_state["clientes"] = clientes
            st.session_state["pdfs_files"] = pdfs

        # Mostrar log
        with st.expander("📋 Log de procesamiento", expanded=True):
            for linea in log:
                st.text(linea)

        # Mostrar resumen por cliente
        if clientes:
            st.success(f"Se procesaron {len(clientes)} cliente(s)")

            for nombre, cuentas in sorted(clientes.items()):
                with st.expander(f"👤 {nombre}"):
                    if cuentas["gbm"]:
                        st.subheader("GBM Regular")
                        st.dataframe(cuentas["gbm"]["resumen"], use_container_width=True)
                        if cuentas["gbm"].get("portafolio"):
                            st.markdown("**Portafolio:**")
                            st.dataframe(pd.DataFrame(cuentas["gbm"]["portafolio"]),
                                         use_container_width=True)
                    if cuentas["smart_cash"]:
                        st.subheader("Smart Cash")
                        st.dataframe(cuentas["smart_cash"]["resumen"], use_container_width=True)
                    if cuentas["prestadero"]:
                        st.subheader("Prestadero")
                        st.dataframe(cuentas["prestadero"]["resumen"], use_container_width=True)

            # Generar Excel descargable
            excel_data = generar_excel_extraccion(clientes)
            st.divider()
            st.download_button(
                label="⬇️ Descargar Reporte de Extraccion (.xlsx)",
                data=excel_data,
                file_name="Reporte_Consolidado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
            )
        else:
            st.warning("No se encontraron datos en los PDFs.")

# ─── TAB 2: CONSOLIDACION ───
with tab2:
    st.header("Consolidar con Maestro Anterior")
    st.markdown("Sube el archivo maestro del mes anterior y los PDFs para generar el nuevo estado consolidado.")

    col1, col2 = st.columns(2)
    with col1:
        maestro = st.file_uploader(
            "📁 Archivo Maestro (.xlsx)",
            type=["xlsx"],
            key="maestro_upload",
        )
    with col2:
        pdfs_cons = st.file_uploader(
            "📄 PDFs del mes",
            type=["pdf"],
            accept_multiple_files=True,
            key="pdfs_consolidar",
        )

    if maestro and pdfs_cons and st.button("🚀 Consolidar", key="btn_consolidar", type="primary"):
        with st.spinner("Extrayendo datos de PDFs..."):
            clientes, log_ext = procesar_pdfs(pdfs_cons)

        with st.expander("📋 Log de extraccion"):
            for linea in log_ext:
                st.text(linea)

        if not clientes:
            st.error("No se encontraron datos en los PDFs.")
        else:
            with st.spinner("Consolidando con maestro anterior..."):
                excel_cons, nombre_salida, log_cons = consolidar_con_maestro(
                    clientes, maestro, pdfs_cons)

            with st.expander("📋 Log de consolidacion", expanded=True):
                for linea in log_cons:
                    st.text(linea)

            st.success(f"Consolidacion completada: {nombre_salida}")

            # Tambien generar el reporte de extraccion
            excel_ext = generar_excel_extraccion(clientes)

            st.divider()
            c1, c2 = st.columns(2)
            with c1:
                st.download_button(
                    label="⬇️ Descargar Estado Consolidado",
                    data=excel_cons,
                    file_name=nombre_salida,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                )
            with c2:
                st.download_button(
                    label="⬇️ Descargar Reporte de Extraccion",
                    data=excel_ext,
                    file_name="Reporte_Consolidado.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

# Footer
st.divider()
st.caption("Automatizacion de Estados de Cuenta v1.0 — GBM / Prestadero")
