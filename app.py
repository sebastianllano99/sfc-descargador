#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
app.py — SFC Descargador · Streamlit
======================================
Descarga informes de Establecimientos de Crédito de la SFC
y genera el consolidado apalancamiento.xlsx.

Render + GitHub: https://render.com
"""

import os
import glob
import time
import shutil
import zipfile
import threading
import tempfile
import io
from pathlib import Path
from datetime import datetime

import streamlit as st

# ─────────────────────────────────────────────────────────────────────────────
# CONSTANTES
# ─────────────────────────────────────────────────────────────────────────────
URL_BASE = (
    "https://www.superfinanciera.gov.co/publicaciones/60765/"
    "informes-y-cifrasinformesinforme-actualidad-del-sistema-"
    "financiero-colombiano-60765/"
)
URL_CA = (
    "https://www.superfinanciera.gov.co/publicaciones/60950/"
    "informes-y-cifrascifrasestablecimientos-de-creditoinformacion-"
    "periodicamensualevolucion-cartera-de-creditos-60950/"
)
URL_PV = (
    "https://www.superfinanciera.gov.co/publicaciones/60949/"
    "informes-y-cifrascifrasestablecimientos-de-creditoinformacion-"
    "periodicamensualprincipales-variables-de-los-establecimientos-"
    "de-credito-60949/"
)

MESES_TEXTO = [
    "enero", "febrero", "marzo", "abril", "mayo", "junio",
    "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"
]

TIMEOUT_DESCARGA = 120

ENTIDADES_DEFAULT = [
    "BANCOLOMBIA", "BANCO POPULAR", "DAVIVIENDA",
    "BBVA", "BANCO DE BOGOTA", "OCCIDENTE",
]

CANONICA = {
    "POPULAR":           "BANCO POPULAR",
    "BANCO POPULAR":     "BANCO POPULAR",
    "BBVA":              "BANCO BBVA",
    "OCCIDENTE":         "BANCO DE OCCIDENTE",
    "BANCO DAVIVIENDA":  "DAVIVIENDA",
    "BANCO DE BOGOTA":   "BANCO DE BOGOTA",
    "BANCO DE BOGOTÁ":   "BANCO DE BOGOTA",
    "BANCO BOGOTA":      "BANCO DE BOGOTA",
    "BANCO BOGOTÁ":      "BANCO DE BOGOTA",
}

MESES_NUM = {
    "enero":1,"febrero":2,"marzo":3,"abril":4,
    "mayo":5,"junio":6,"julio":7,"agosto":8,
    "septiembre":9,"setiembre":9,"octubre":10,
    "noviembre":11,"diciembre":12
}

ESPECIALES = [
    {"anio":2021,"mes":"julio","mes_num":7,"tipo":"pv","texto_mes":"Julio 2021","url":URL_PV},
    {"anio":2021,"mes":"julio","mes_num":7,"tipo":"ca","texto_mes":"Julio 2021","url":URL_CA},
    {"anio":2023,"mes":"marzo","mes_num":3,"tipo":"pv","texto_mes":"Marzo 2023","url":URL_PV},
    {"anio":2023,"mes":"marzo","mes_num":3,"tipo":"ca","texto_mes":"Marzo 2023","url":URL_CA},
    {"anio":2023,"mes":"diciembre","mes_num":12,"tipo":"pv","texto_mes":"Diciembre 2023","url":URL_PV},
    {"anio":2023,"mes":"diciembre","mes_num":12,"tipo":"ca","texto_mes":"Diciembre 2023","url":URL_CA},
]

# ─────────────────────────────────────────────────────────────────────────────
# ESTADO DE SESIÓN
# ─────────────────────────────────────────────────────────────────────────────
def init_state():
    defaults = {
        "corriendo":    False,
        "log":          [],
        "archivos":     {},   # {anio: {nombre: bytes}}
        "ok":           0,
        "errores":      0,
        "sin_link":     0,
        "total":        0,
        "procesados":   0,
        "hilo":         None,
        "carpeta_tmp":  None,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

# ─────────────────────────────────────────────────────────────────────────────
# UTILIDADES SCRAPING
# ─────────────────────────────────────────────────────────────────────────────

def log(tipo, msg):
    ts = datetime.now().strftime("%H:%M:%S")
    st.session_state.log.append({"tipo": tipo, "msg": msg, "ts": ts})
    iconos = {"ok": "✓", "err": "✗", "warn": "!", "info": "→"}
    print(f"{iconos.get(tipo,'·')} [{ts}] {msg}")


def _es_xlsx_disfrazado(ruta):
    try:
        with zipfile.ZipFile(ruta, "r") as z:
            return "[Content_Types].xml" in z.namelist()
    except Exception:
        return False


def convertir_xls_a_xlsx(ruta_xls):
    if not os.path.exists(ruta_xls):
        return ruta_xls
    ruta_xlsx = ruta_xls[:-4] + ".xlsx"
    try:
        import xlrd, openpyxl
        wb_in  = xlrd.open_workbook(ruta_xls)
        wb_out = openpyxl.Workbook()
        wb_out.remove(wb_out.active)
        for idx in range(wb_in.nsheets):
            hi = wb_in.sheet_by_index(idx)
            ho = wb_out.create_sheet(title=hi.name)
            for r in range(hi.nrows):
                for c in range(hi.ncols):
                    ho.cell(row=r+1, column=c+1, value=hi.cell_value(r, c))
        wb_out.save(ruta_xlsx)
        os.remove(ruta_xls)
        return ruta_xlsx
    except Exception:
        pass
    try:
        import pandas as pd
        dfs = pd.read_excel(ruta_xls, sheet_name=None, header=None, engine="xlrd")
        with pd.ExcelWriter(ruta_xlsx, engine="openpyxl") as w:
            for n, df in dfs.items():
                df.to_excel(w, sheet_name=n, index=False, header=False)
        os.remove(ruta_xls)
        return ruta_xlsx
    except Exception:
        pass
    try:
        shutil.copy2(ruta_xls, ruta_xlsx)
        os.remove(ruta_xls)
        return ruta_xlsx
    except Exception:
        return ruta_xls


def esperar_descarga(carpeta_temp, timeout=TIMEOUT_DESCARGA):
    fin = time.time() + timeout
    while time.time() < fin:
        archivos = [
            f for f in glob.glob(os.path.join(carpeta_temp, "*"))
            if not f.endswith(".crdownload") and not f.endswith(".tmp")
            and os.path.getsize(f) > 1000
        ]
        if archivos:
            ruta = max(archivos, key=os.path.getctime)
            tam1 = os.path.getsize(ruta)
            time.sleep(2)
            if os.path.getsize(ruta) == tam1:
                return ruta
        time.sleep(1)
    return None


def iniciar_chrome(carpeta_dl):
    from selenium import webdriver
    from selenium.webdriver.chrome.service import Service
    from selenium.webdriver.chrome.options import Options
    import shutil as _sh

    op = Options()
    op.add_argument("--headless=new")
    op.add_argument("--no-sandbox")
    op.add_argument("--disable-dev-shm-usage")
    op.add_argument("--disable-gpu")
    op.add_argument("--window-size=1920,1080")
    op.add_argument("--log-level=3")
    op.add_argument("--disable-extensions")
    op.add_argument("--disable-setuid-sandbox")
    op.add_experimental_option("prefs", {
        "download.default_directory": os.path.abspath(carpeta_dl),
        "download.prompt_for_download": False,
        "directory_upgrade": True,
        "plugins.always_open_pdf_externally": True,
    })

    chrome_bin = (
        os.environ.get("CHROME_BIN")
        or _sh.which("google-chrome")
        or _sh.which("google-chrome-stable")
        or _sh.which("chromium")
        or _sh.which("chromium-browser")
    )
    chromedriver_bin = (
        os.environ.get("CHROMEDRIVER_PATH")
        or _sh.which("chromedriver")
    )

    if chrome_bin:
        op.binary_location = chrome_bin

    driver = None
    if chromedriver_bin:
        try:
            driver = webdriver.Chrome(service=Service(chromedriver_bin), options=op)
        except Exception as e:
            log("warn", f"chromedriver explícito falló: {e}")

    if driver is None:
        try:
            import concurrent.futures
            with concurrent.futures.ThreadPoolExecutor(max_workers=1) as ex:
                driver = ex.submit(webdriver.Chrome, options=op).result(timeout=60)
        except Exception as e:
            log("err", f"No se pudo iniciar Chrome: {e}")

    return driver


def click_js(driver, el):
    driver.execute_script(
        "arguments[0].scrollIntoView({behavior:'smooth',block:'center'});", el
    )
    time.sleep(0.6)
    driver.execute_script("arguments[0].click();", el)


def _buscar_link(driver, texto, anio_ref=None, mes_ref=None):
    from selenium.webdriver.common.by import By
    for variante in [texto, texto.lower(), texto.upper(), texto.title()]:
        try:
            el = driver.find_element(By.LINK_TEXT, variante)
            if el.is_displayed(): return el
        except Exception: pass
    try:
        el = driver.find_element(By.PARTIAL_LINK_TEXT, texto)
        if el.is_displayed(): return el
    except Exception: pass
    try:
        el = driver.find_element(By.XPATH, f"//a[contains(normalize-space(.), '{texto}')]")
        if el.is_displayed(): return el
    except Exception: pass
    if anio_ref and mes_ref:
        for lk in driver.find_elements(By.XPATH, "//a[@href]"):
            try:
                if not lk.is_displayed(): continue
                t = lk.text.strip().lower()
                if str(anio_ref) in t and mes_ref.lower() in t: return lk
            except Exception: pass
    return None


# ─────────────────────────────────────────────────────────────────────────────
# SCRAPING PRINCIPAL  (corre en hilo separado)
# ─────────────────────────────────────────────────────────────────────────────

def hilo_scraping(anos_deseados, delay, forzar, carpeta_base):
    """
    Descarga los archivos de la SFC en carpeta_base/AÑO/nombre.xlsx
    Actualiza st.session_state directamente (hilo daemon).
    """
    from selenium.webdriver.common.by import By

    log("info", f"Hilo iniciado. Años: {anos_deseados}")

    carpeta_temp = os.path.join(carpeta_base, "_temp_chrome")
    os.makedirs(carpeta_temp, exist_ok=True)

    def limpiar_temp():
        shutil.rmtree(carpeta_temp, ignore_errors=True)
        os.makedirs(carpeta_temp, exist_ok=True)

    driver = iniciar_chrome(carpeta_temp)
    if driver is None:
        log("err", "Chrome no disponible.")
        st.session_state.corriendo = False
        return

    try:
        log("info", f"Cargando página SFC...")
        driver.get(URL_BASE)
        time.sleep(3)

        anos_set  = set(anos_deseados)
        anio_ini  = min(anos_deseados)
        anio_fin  = max(anos_deseados)

        # Mapear años por posición Y
        mapa_anios = []
        for y in range(anio_fin, anio_ini - 1, -1):
            elems = driver.find_elements(By.XPATH, f"//*[contains(text(), '{y}')]")
            for el in elems:
                try:
                    if el.is_displayed() and len(el.text.strip()) < 12:
                        mapa_anios.append((y, el.location["y"]))
                        break
                except Exception:
                    continue
        mapa_anios.sort(key=lambda x: x[1])

        if not mapa_anios:
            log("err", "No se encontraron años en la página.")
            st.session_state.corriendo = False
            return

        log("info", f"Años detectados: {[a for a,_ in mapa_anios]}")
        mapa_filtrado = [(a, y) for a, y in mapa_anios if a in anos_set]

        if not mapa_filtrado:
            log("err", f"Ningún año encontrado en página: {anos_deseados}")
            st.session_state.corriendo = False
            return

        st.session_state.total = len(mapa_filtrado) * 12
        contador = 0

        for idx_anio, (anio, pos_y_actual) in enumerate(mapa_filtrado):
            pos_y_siguiente = 999999
            for a, y in mapa_anios:
                if y > pos_y_actual:
                    pos_y_siguiente = y
                    break

            log("info", f"=== AÑO {anio} ===")
            carpeta_anio = os.path.join(carpeta_base, str(anio))
            os.makedirs(carpeta_anio, exist_ok=True)

            for mes in MESES_TEXTO:
                contador += 1
                st.session_state.procesados = contador
                st.session_state.progreso   = round(contador / st.session_state.total * 100)

                nombre_destino = os.path.join(carpeta_anio, f"{anio}_{mes}_Establecimientos.xlsx")
                if os.path.exists(nombre_destino) and not forzar:
                    log("info", f"  {mes} {anio}: ya existe, omitiendo")
                    st.session_state.ok += 1
                    continue

                limpiar_temp()
                driver.get(URL_BASE)
                time.sleep(2)

                # Buscar link del mes en el rango Y del año
                xpath_mes = (
                    f"//a[@href and contains("
                    f"translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz')"
                    f", '{mes}')]"
                )
                candidatos  = driver.find_elements(By.XPATH, xpath_mes)
                link_mes    = None
                for cand in candidatos:
                    try:
                        if not cand.is_displayed(): continue
                        y_cand = cand.location["y"]
                        if not (pos_y_actual < y_cand < pos_y_siguiente): continue
                        try:
                            tabla = cand.find_element(By.XPATH, "./ancestor::table[1]")
                            if str(anio) not in tabla.text: continue
                        except Exception: pass
                        link_mes = cand
                        break
                    except Exception:
                        continue

                if not link_mes:
                    log("warn", f"  {mes} {anio}: mes no encontrado")
                    st.session_state.sin_link += 1
                    continue

                log("info", f"  Abriendo {mes} {anio}...")
                click_js(driver, link_mes)
                time.sleep(2)

                # Buscar enlace Establecimientos de Crédito
                enlace = None
                for _ in range(5):
                    try:
                        enlace = driver.find_element(By.LINK_TEXT, "Establecimientos de crédito")
                        break
                    except Exception: pass
                    try:
                        enlace = driver.find_element(By.XPATH,
                            "//a[contains(translate(.,'ÁÉÍÓÚÁÉÍÓÚ','AEIOUaeiou'),'Establecimientos') "
                            "and contains(translate(.,'ÁÉÍÓÚÁÉÍÓÚ','AEIOUaeiou'),'Credito')]")
                        break
                    except Exception: pass
                    time.sleep(1)

                if not enlace:
                    log("warn", f"  {mes} {anio}: no se encontró 'Establecimientos de Crédito'")
                    st.session_state.sin_link += 1
                    continue

                click_js(driver, enlace)
                time.sleep(3)

                archivo = esperar_descarga(carpeta_temp)
                if archivo:
                    _procesar_archivo_descargado(archivo, carpeta_anio, anio, mes)
                else:
                    log("err", f"  {mes} {anio}: timeout")
                    st.session_state.errores += 1

                if delay > 0:
                    time.sleep(delay)

        # ── Casos especiales ──────────────────────────────────────────────────
        especiales_a_procesar = [e for e in ESPECIALES if e["anio"] in anos_set]
        if especiales_a_procesar:
            log("info", "=== CASOS ESPECIALES ===")

        for esp in especiales_a_procesar:
            anio_e = esp["anio"]
            mes_e  = esp["mes"]
            tipo_e = esp["tipo"]
            prefijo = f"{anio_e}_{mes_e}_{tipo_e}{esp['mes_num']:02d}{anio_e}"
            carpeta_e = os.path.join(carpeta_base, str(anio_e))
            os.makedirs(carpeta_e, exist_ok=True)

            limpiar_temp()
            try:
                driver.get(esp["url"])
                time.sleep(4)
                link_mes = _buscar_link(driver, esp["texto_mes"], anio_e, mes_e)
                if link_mes is None:
                    log("warn", f"  [{anio_e}/{mes_e}/{tipo_e.upper()}] No se encontró el link")
                    st.session_state.sin_link += 1
                    continue
                click_js(driver, link_mes)
                time.sleep(3)
                archivo_e = esperar_descarga(carpeta_temp)
                if archivo_e:
                    _procesar_archivo_descargado(archivo_e, carpeta_e, anio_e, mes_e, prefijo=prefijo)
                else:
                    log("warn", f"  [{anio_e}/{mes_e}/{tipo_e.upper()}] Timeout")
                    st.session_state.sin_link += 1
            except Exception as ee:
                log("err", f"  [{anio_e}/{mes_e}/{tipo_e.upper()}] Error: {ee}")

        log("ok", f"Proceso completado — "
            f"{st.session_state.ok} descargados, "
            f"{st.session_state.sin_link} sin link, "
            f"{st.session_state.errores} errores")

    except Exception as e:
        import traceback
        log("err", f"Error inesperado: {e}")
        log("err", traceback.format_exc()[:400])
    finally:
        try:
            driver.quit()
        except Exception: pass
        shutil.rmtree(carpeta_temp, ignore_errors=True)
        st.session_state.corriendo = False


def _procesar_archivo_descargado(archivo, carpeta_destino, anio, mes, prefijo=None):
    """Mueve/extrae el archivo descargado a carpeta_destino y actualiza estado."""
    ext = os.path.splitext(archivo)[1].lower()
    os.makedirs(carpeta_destino, exist_ok=True)

    if (ext == ".zip" or zipfile.is_zipfile(archivo)) and not _es_xlsx_disfrazado(archivo):
        try:
            with zipfile.ZipFile(archivo, "r") as z:
                contenido = [n for n in z.namelist() if not n.endswith("/")]
                extraidos = 0
                for idx_z, nombre_zip in enumerate(contenido):
                    ext_inner = os.path.splitext(nombre_zip)[1].lower() or ".xlsx"
                    if prefijo:
                        nombre_final = f"{prefijo}{ext_inner}" if len(contenido)==1 else f"{prefijo}_{idx_z+1}{ext_inner}"
                    else:
                        nombre_final = f"{anio}_{mes}_Establecimientos{ext_inner}" if len(contenido)==1 else f"{anio}_{mes}_{os.path.basename(nombre_zip)}"
                    ruta_dest = os.path.join(carpeta_destino, nombre_final)
                    with open(ruta_dest, "wb") as f_out:
                        f_out.write(z.read(nombre_zip))
                    if ext_inner == ".xls":
                        ruta_dest = convertir_xls_a_xlsx(ruta_dest)
                        nombre_final = os.path.basename(ruta_dest)
                    kb = os.path.getsize(ruta_dest) / 1024
                    log("ok", f"  Extraído: {nombre_final} ({kb:.0f} KB)")
                    extraidos += 1
            os.remove(archivo)
            if extraidos > 0:
                st.session_state.ok += 1
            else:
                log("warn", f"  ZIP vacío: {mes} {anio}")
        except Exception as e:
            log("err", f"  Error ZIP {mes} {anio}: {e}")
            st.session_state.errores += 1
    else:
        ext_final = ".xlsx" if _es_xlsx_disfrazado(archivo) else (ext or ".xlsx")
        nombre_final = f"{prefijo or f'{anio}_{mes}_Establecimientos'}{ext_final}"
        destino = os.path.join(carpeta_destino, nombre_final)
        shutil.move(archivo, destino)
        if ext_final == ".xls":
            destino = convertir_xls_a_xlsx(destino)
            nombre_final = os.path.basename(destino)
        kb = os.path.getsize(destino) / 1024
        log("ok", f"  Guardado: {nombre_final} ({kb:.0f} KB)")
        st.session_state.ok += 1


# ─────────────────────────────────────────────────────────────────────────────
# EXTRACTOR / CONSOLIDADO
# ─────────────────────────────────────────────────────────────────────────────

def _norm(t):
    t = str(t)
    for a, b in [("Á","A"),("É","E"),("Í","I"),("Ó","O"),("Ú","U"),
                 ("á","a"),("é","e"),("í","i"),("ó","o"),("ú","u")]:
        t = t.replace(a, b)
    return t.upper().strip()


def _extraer_anio_mes(nombre):
    partes = nombre.lower().replace(".xlsx","").split("_")
    anio = mes = None
    for p in partes:
        if p.isdigit() and 2000 <= int(p) <= 2030:
            anio = int(p)
        if p in MESES_NUM:
            mes = MESES_NUM[p]
    return anio, mes


def _procesar_pv(path_archivo, anio, mes, entidades_sel):
    try:
        import pandas as pd
        xls = pd.ExcelFile(path_archivo)
        meses_abrev = {1:"ene",2:"feb",3:"mar",4:"abr",5:"may",6:"jun",
                       7:"jul",8:"ago",9:"sep",10:"oct",11:"nov",12:"dic"}
        abrev_mes = f"{meses_abrev.get(mes,'')}-{str(anio)[2:]}"
        hoja = None
        for h in xls.sheet_names:
            if h.strip().lower() == abrev_mes:
                hoja = h; break
        if hoja is None:
            for h in xls.sheet_names:
                if h.strip().upper() in ["3", "PV3"]:
                    hoja = h; break
        if hoja is None:
            hojas = [h for h in xls.sheet_names
                     if pd.read_excel(path_archivo, sheet_name=h, nrows=1).shape[1] > 0]
            hoja = hojas[-1] if hojas else xls.sheet_names[0]

        df = pd.read_excel(path_archivo, sheet_name=hoja, header=None)
        fila0_vals = [_norm(str(v)) for v in df.iloc[0, 1:10] if str(v).strip() not in ("","NAN")]
        col0_vals  = [_norm(str(v)) for v in df.iloc[2:8, 0]  if str(v).strip() not in ("","NAN")]
        es_formato_b = (
            any("TOTAL" in v or "BANCO" in v for v in fila0_vals) and
            any(v in ("ACTIVOS","PASIVOS","PATRIMONIO","RESULTADOS") for v in col0_vals)
        )
        return _fmt_b(df, anio, mes, entidades_sel) if es_formato_b else _fmt_a(df, anio, mes, entidades_sel)
    except Exception as e:
        return []


def _fmt_a(df, anio, mes, entidades_sel):
    import pandas as pd
    fila_ent = col_ent = None
    for i in range(min(15, len(df))):
        for j in range(min(10, len(df.columns))):
            v = _norm(str(df.iloc[i, j]))
            if v in ("ENTIDAD","NOMBRE ENTIDAD"):
                fila_ent = i; col_ent = j; break
        if fila_ent is not None: break
    if fila_ent is None: return []

    col_act = col_pas = col_pat = col_res_acum = col_res_nuevo = None
    for fc in range(max(0, fila_ent-1), fila_ent+2):
        if fc >= len(df): continue
        cab = df.iloc[fc]
        for j in range(len(cab)):
            v = _norm(str(cab.iloc[j]))
            if v == "ACTIVO" and col_act is None:       col_act = j
            elif v == "PASIVO" and col_pas is None:     col_pas = j
            elif v == "PATRIMONIO" and col_pat is None: col_pat = j
            elif "RESULTADO" in v and "ACUMULADO" in v and col_res_acum is None: col_res_acum = j
            elif v == "RESULTADOS DEL EJERCICIO":       col_res_nuevo = j
    if col_act is None: return []

    col_res  = col_res_nuevo if col_res_nuevo is not None else (
               col_res_acum + 2 if col_res_acum is not None else None)
    sub_val  = _norm(str(df.iloc[fila_ent+1, col_act])) if fila_ent+1 < len(df) else ""
    fila_dat = fila_ent + 2 if sub_val == "ACTIVO" else fila_ent + 1
    datos    = df.iloc[fila_dat:].reset_index(drop=True)

    resultados = []; vistos = set()
    for ent in entidades_sel:
        can = CANONICA.get(_norm(ent), _norm(ent))
        if can in vistos: continue
        mask = datos[col_ent].apply(lambda x: _norm(ent) in _norm(x))
        fila_df = datos[mask]
        if fila_df.empty: continue
        fila = fila_df.iloc[0]
        def safe(c):
            try: v=fila[c]; return None if pd.isna(v) else v
            except: return None
        resultados.append({"Año":anio,"Mes":mes,"Entidad":can,
            "Activo":safe(col_act),"Pasivo":safe(col_pas),
            "Patrimonio":safe(col_pat),"Resultado del ejercicio":safe(col_res)})
        vistos.add(can)
    return resultados


def _fmt_b(df, anio, mes, entidades_sel):
    import pandas as pd
    fila_act = fila_pas = fila_pat = fila_res = None
    for i in range(min(15, len(df))):
        v = _norm(str(df.iloc[i, 0]))
        if v in ("ACTIVOS","ACTIVO") and fila_act is None:     fila_act = i
        elif v in ("PASIVOS","PASIVO") and fila_pas is None:   fila_pas = i
        elif v == "PATRIMONIO" and fila_pat is None:           fila_pat = i
        elif v in ("RESULTADOS","RESULTADO","RESULTADO DEL EJERCICIO",
                   "RESULTADO DEL PERIODO") and fila_res is None: fila_res = i
    if fila_act is None: return []

    ent_cols = {}
    for j in range(1, len(df.columns)):
        nombre = _norm(str(df.iloc[0, j]))
        if nombre and nombre != "NAN":
            ent_cols[nombre] = j

    resultados = []; vistos = set()
    for ent in entidades_sel:
        can = CANONICA.get(_norm(ent), _norm(ent))
        if can in vistos: continue
        col = next((j for n, j in ent_cols.items() if _norm(ent) in n), None)
        if col is None: continue
        def safe_b(fi):
            try:
                if fi is None: return None
                v = df.iloc[fi, col]
                return None if __import__("pandas").isna(v) else v
            except: return None
        resultados.append({"Año":anio,"Mes":mes,"Entidad":can,
            "Activo":safe_b(fila_act),"Pasivo":safe_b(fila_pas),
            "Patrimonio":safe_b(fila_pat),"Resultado del ejercicio":safe_b(fila_res)})
        vistos.add(can)
    return resultados


def _detectar_hoja_icm(anio, mes):
    if   (anio, mes) <= (2008,  5): return ["10"],        9
    elif (anio, mes) <= (2011,  1): return ["11"],        9
    elif (anio, mes) <= (2014, 12): return ["C11", "11"], 9
    elif (anio, mes) <= (2016, 12): return ["C6"],        7
    else:                           return ["C6"],        10


def _extraer_icm(path_archivo, anio, mes, entidades_sel):
    try:
        import pandas as pd
        candidatos_hoja, col_val = _detectar_hoja_icm(anio, mes)
        xls = pd.ExcelFile(path_archivo)
        hoja = None
        for c in candidatos_hoja:
            for h in xls.sheet_names:
                if h.strip() == c:
                    hoja = h; break
            if hoja: break
        if hoja is None: return []

        df = pd.read_excel(path_archivo, sheet_name=hoja, header=None)

        COL_ENT = 1
        def es_grupo(n): return any(p in n for p in ("TOTAL","ESTABLEC","BANCOS","SECTOR"))
        def _coincide(ent, cel): return _norm(ent) in cel or cel in _norm(ent)

        resultados = []; vistos = set()
        for ent in entidades_sel:
            can = CANONICA.get(_norm(ent), _norm(ent))
            if can in vistos: continue
            fila_match = None
            for i in range(len(df)):
                celda = df.iloc[i, COL_ENT]
                if pd.isna(celda): continue
                nombre_cel = _norm(str(celda))
                if es_grupo(nombre_cel): continue
                if _coincide(ent, nombre_cel):
                    fila_match = i; break
            if fila_match is None: continue
            try:
                val = df.iloc[fila_match, col_val]
                icm = None if pd.isna(val) else float(val)
            except Exception:
                icm = None
            resultados.append({"Año":anio,"Mes":mes,"Entidad":can,"ICM":icm})
            vistos.add(can)
        return resultados
    except Exception:
        return []


def generar_consolidado(archivos_subidos, entidades_sel, config):
    """
    archivos_subidos: lista de UploadedFile de Streamlit
    Retorna bytes del Excel consolidado.
    """
    import pandas as pd

    registros_pv  = []
    registros_icm = []

    with tempfile.TemporaryDirectory() as tmp:
        for uploaded in archivos_subidos:
            nombre = uploaded.name.lower()
            anio, mes = _extraer_anio_mes(uploaded.name)
            if anio is None or mes is None:
                continue

            ruta = os.path.join(tmp, uploaded.name)
            with open(ruta, "wb") as f:
                f.write(uploaded.read())

            if "pv" in nombre:
                datos = _procesar_pv(ruta, anio, mes, entidades_sel)
                registros_pv.extend(datos)
            elif "_ca" in nombre or nombre.startswith("ca"):
                datos = _extraer_icm(ruta, anio, mes, entidades_sel)
                registros_icm.extend(datos)

    salida = io.BytesIO()
    with pd.ExcelWriter(salida, engine="openpyxl") as writer:
        if registros_pv:
            df_pv = pd.DataFrame(registros_pv)
            df_pv["Año"] = pd.to_numeric(df_pv["Año"], errors="coerce").astype("Int64")
            df_pv["Mes"] = pd.to_numeric(df_pv["Mes"], errors="coerce").astype("Int64")
            df_pv.sort_values(["Año","Mes","Entidad"], inplace=True)
            df_pv["Indicador de Apalancamiento"] = (
                df_pv["Pasivo"] / df_pv["Activo"]
            ).where(df_pv["Activo"].notna() & (df_pv["Activo"] != 0))
            df_pv.to_excel(writer, sheet_name="apalancamiento", index=False)
        else:
            pd.DataFrame(columns=["Año","Mes","Entidad","Activo","Pasivo",
                                   "Patrimonio","Resultado del ejercicio"]
                        ).to_excel(writer, sheet_name="apalancamiento", index=False)

        if registros_icm:
            df_icm = pd.DataFrame(registros_icm)
            df_icm["Año"] = pd.to_numeric(df_icm["Año"], errors="coerce").astype("Int64")
            df_icm["Mes"] = pd.to_numeric(df_icm["Mes"], errors="coerce").astype("Int64")
            df_icm.sort_values(["Año","Mes","Entidad"], inplace=True)
            df_icm.to_excel(writer, sheet_name="ICV", index=False)
        else:
            pd.DataFrame(columns=["Año","Mes","Entidad","ICM"]
                        ).to_excel(writer, sheet_name="ICV", index=False)

    return salida.getvalue(), len(registros_pv), len(registros_icm)


# ─────────────────────────────────────────────────────────────────────────────
# INTERFAZ STREAMLIT
# ─────────────────────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="SFC Descargador",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Syne:wght@400;700;800&family=DM+Mono:wght@400;500&display=swap');

html, body, [class*="css"] { font-family: 'Syne', sans-serif; }
.stApp { background: #060b14; color: #dde6f4; }

/* Ocultar elementos de Streamlit que no necesitamos */
#MainMenu, footer, header { visibility: hidden; }
.block-container { padding-top: 2rem; padding-bottom: 2rem; max-width: 900px; }

/* Títulos */
h1 { font-weight: 800 !important; letter-spacing: -1px; }
h2, h3 { font-weight: 700 !important; }

/* Cards */
.sfc-card {
    background: #0d1624;
    border: 1px solid #1a2d47;
    border-radius: 14px;
    padding: 24px;
    margin-bottom: 18px;
}
.sfc-card-teal  { border-color: rgba(0,212,184,.3); }
.sfc-card-amber { border-color: rgba(245,158,11,.3); }

/* Badge */
.sfc-badge {
    display: inline-flex; align-items: center; gap: 6px;
    background: rgba(26,111,255,.12);
    border: 1px solid rgba(26,111,255,.25);
    border-radius: 99px; padding: 5px 14px;
    font-family: 'DM Mono', monospace; font-size: 11px; color: #7aa8ff;
    margin-bottom: 12px;
}

/* Log box */
.log-entry { font-family: 'DM Mono', monospace; font-size: 12px; line-height: 1.8; }
.log-ok   { color: #22c55e; }
.log-err  { color: #f43f5e; }
.log-warn { color: #f59e0b; }
.log-info { color: #00d4b8; }
.log-ts   { color: #4d6480; margin-right: 8px; }

/* Archivo descargable */
.dl-item {
    display: flex; justify-content: space-between; align-items: center;
    background: #111e30; border: 1px solid #1a2d47;
    border-radius: 8px; padding: 10px 14px; margin-bottom: 6px;
    font-family: 'DM Mono', monospace; font-size: 12px;
}
</style>
""", unsafe_allow_html=True)

init_state()

# ── ENCABEZADO ────────────────────────────────────────────────────────────────
st.markdown('<div class="sfc-badge">SFC · Superfinanciera Colombia</div>', unsafe_allow_html=True)
st.markdown("# Descargador de *Informes*")
st.markdown("**Establecimientos de Crédito · 2005–2026 · Automatizado**")
st.divider()

tab1, tab2 = st.tabs(["📥  Descargar archivos SFC", "📊  Generar consolidado"])

# ═══════════════════════════════════════════════════════════════════════════════
# TAB 1 — DESCARGA
# ═══════════════════════════════════════════════════════════════════════════════
with tab1:
    st.markdown("### Años a descargar")

    col_sel, col_cfg = st.columns([2, 1])

    with col_sel:
        anos_disponibles = list(range(2026, 2004, -1))
        anos_seleccionados = st.multiselect(
            "Selecciona uno o más años",
            options=anos_disponibles,
            default=[],
            placeholder="Ej: 2025, 2024, 2023…",
        )

    with col_cfg:
        delay  = st.number_input("Delay entre páginas (seg)", min_value=1.0, max_value=10.0, value=2.5, step=0.5)
        forzar = st.checkbox("Re-descargar existentes", value=False)

    st.caption(
        "Los archivos se descargarán en el servidor y podrás descargarlos uno a uno "
        "desde esta interfaz. Se organizan en `Descargas_SFC/AÑO/archivo.xlsx`."
    )

    iniciar_btn = st.button(
        "▶  Iniciar descarga",
        type="primary",
        disabled=st.session_state.corriendo or not anos_seleccionados,
        use_container_width=True,
    )

    if iniciar_btn and anos_seleccionados:
        # Limpiar estado anterior
        st.session_state.log        = []
        st.session_state.ok         = 0
        st.session_state.errores    = 0
        st.session_state.sin_link   = 0
        st.session_state.total      = 0
        st.session_state.procesados = 0
        st.session_state.corriendo  = True

        # Carpeta temporal en el servidor
        carpeta_tmp = tempfile.mkdtemp(prefix="sfc_")
        st.session_state.carpeta_tmp = carpeta_tmp

        hilo = threading.Thread(
            target=hilo_scraping,
            args=(anos_seleccionados, delay, forzar, carpeta_tmp),
            daemon=True,
        )
        hilo.start()
        st.session_state.hilo = hilo
        st.rerun()

    # ── Progreso ─────────────────────────────────────────────────────────────
    if st.session_state.corriendo or st.session_state.ok > 0 or st.session_state.errores > 0:

        col_ok, col_skip, col_err = st.columns(3)
        col_ok.metric("✅ Descargados",  st.session_state.ok)
        col_skip.metric("⚠️ Sin link",   st.session_state.sin_link)
        col_err.metric("❌ Errores",     st.session_state.errores)

        total = st.session_state.total or 1
        pct   = st.session_state.procesados / total
        st.progress(pct, text=f"{st.session_state.procesados}/{total} procesados")

        # Log
        if st.session_state.log:
            log_html = ""
            for entrada in st.session_state.log[-60:]:
                cls = {"ok":"log-ok","err":"log-err","warn":"log-warn","info":"log-info"}.get(entrada["tipo"],"log-info")
                msg = entrada["msg"].replace("<","&lt;").replace(">","&gt;")
                log_html += f'<div class="log-entry"><span class="log-ts">{entrada["ts"]}</span><span class="{cls}">{msg}</span></div>'
            st.markdown(f'<div style="background:#060b14;border:1px solid #1a2d47;border-radius:8px;padding:12px 16px;max-height:280px;overflow-y:auto">{log_html}</div>', unsafe_allow_html=True)

        if st.session_state.corriendo:
            time.sleep(1.5)
            st.rerun()

    # ── Archivos disponibles para descarga ───────────────────────────────────
    carpeta_tmp = st.session_state.get("carpeta_tmp")
    if carpeta_tmp and os.path.exists(carpeta_tmp) and not st.session_state.corriendo:
        archivos_en_servidor = []
        for anio_dir in sorted(Path(carpeta_tmp).iterdir()):
            if anio_dir.is_dir() and anio_dir.name.isdigit():
                for f in sorted(anio_dir.iterdir()):
                    if f.suffix.lower() in (".xlsx", ".xls"):
                        archivos_en_servidor.append(f)

        if archivos_en_servidor:
            st.divider()
            st.markdown(f"### 📂 {len(archivos_en_servidor)} archivos listos para descargar")
            st.caption("Haz clic en cada botón para guardar el archivo en tu computador.")

            # Agrupar por año
            por_anio = {}
            for f in archivos_en_servidor:
                anio = f.parent.name
                por_anio.setdefault(anio, []).append(f)

            for anio in sorted(por_anio.keys()):
                with st.expander(f"📁 {anio}  ({len(por_anio[anio])} archivos)", expanded=True):
                    for f in por_anio[anio]:
                        col_nombre, col_btn = st.columns([3, 1])
                        col_nombre.markdown(
                            f'<div style="font-family:monospace;font-size:12px;padding:8px 0;color:#dde6f4">'
                            f'📄 {f.name}  <span style="color:#4d6480">({f.stat().st_size//1024} KB)</span></div>',
                            unsafe_allow_html=True
                        )
                        col_btn.download_button(
                            label="⬇ Descargar",
                            data=f.read_bytes(),
                            file_name=f.name,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key=f"dl_{f.parent.name}_{f.name}",
                            use_container_width=True,
                        )

            st.divider()
            # Opción de descargar todo como ZIP
            st.markdown("#### 📦 O descarga todo de una vez como ZIP")
            zip_buf = io.BytesIO()
            with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                for f in archivos_en_servidor:
                    zf.write(f, arcname=f"Descargas_SFC/{f.parent.name}/{f.name}")
            st.download_button(
                label=f"⬇ Descargar ZIP completo ({len(archivos_en_servidor)} archivos)",
                data=zip_buf.getvalue(),
                file_name="Descargas_SFC.zip",
                mime="application/zip",
                use_container_width=True,
                type="primary",
            )


# ═══════════════════════════════════════════════════════════════════════════════
# TAB 2 — CONSOLIDADO
# ═══════════════════════════════════════════════════════════════════════════════
with tab2:
    st.markdown("### Generar consolidado `apalancamiento.xlsx`")
    st.markdown(
        "Sube los archivos **PV** y **CA** de los meses que quieras consolidar. "
        "Puedes seleccionar múltiples archivos a la vez desde tu carpeta `Descargas_SFC`."
    )

    # Entidades
    st.markdown("#### Entidades a incluir")
    col_ents = st.columns(3)
    entidades_activas = []
    ents_default = ["BANCOLOMBIA","BANCO POPULAR","DAVIVIENDA","BBVA","BANCO DE BOGOTA","OCCIDENTE"]
    for i, ent in enumerate(ents_default):
        if col_ents[i % 3].checkbox(ent.title(), value=True, key=f"ent_{ent}"):
            entidades_activas.append(ent)

    ent_extra = st.text_input(
        "Agregar entidad adicional",
        placeholder="Ej: BANCO CAJA SOCIAL",
        help="Escribe el nombre en mayúsculas y presiona Enter"
    )
    if ent_extra.strip():
        entidades_activas.append(ent_extra.strip().upper())

    st.divider()

    # Subida de archivos
    st.markdown("#### Archivos a procesar")
    archivos_subidos = st.file_uploader(
        "Selecciona los archivos Excel (PV y CA) desde tu carpeta Descargas_SFC",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        help="Puedes seleccionar archivos de múltiples años. "
             "Busca en Descargas_SFC/2024/, Descargas_SFC/2023/, etc."
    )

    if archivos_subidos:
        pv_count = sum(1 for f in archivos_subidos if "pv" in f.name.lower())
        ca_count = sum(1 for f in archivos_subidos if "_ca" in f.name.lower() or f.name.lower().startswith("ca"))
        st.caption(f"**{len(archivos_subidos)}** archivos cargados — {pv_count} PV · {ca_count} CA")

        anos_detectados = set()
        for f in archivos_subidos:
            anio, _ = _extraer_anio_mes(f.name)
            if anio: anos_detectados.add(anio)
        if anos_detectados:
            st.caption(f"Años detectados: {', '.join(str(a) for a in sorted(anos_detectados))}")

    st.divider()

    generar_btn = st.button(
        "📊  Generar consolidado",
        type="primary",
        disabled=not archivos_subidos or not entidades_activas,
        use_container_width=True,
    )

    if generar_btn:
        if not entidades_activas:
            st.error("Selecciona al menos una entidad.")
        elif not archivos_subidos:
            st.error("Sube al menos un archivo Excel.")
        else:
            with st.spinner("Procesando archivos..."):
                try:
                    xlsx_bytes, n_pv, n_icm = generar_consolidado(
                        archivos_subidos,
                        entidades_activas,
                        {}
                    )
                    st.success(
                        f"✅ Consolidado generado — "
                        f"{n_pv} registros de apalancamiento · {n_icm} registros ICV"
                    )
                    st.download_button(
                        label="⬇  Descargar apalancamiento.xlsx",
                        data=xlsx_bytes,
                        file_name="apalancamiento.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        type="primary",
                    )
                except Exception as e:
                    st.error(f"Error generando consolidado: {e}")
