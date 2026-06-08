import json, math, os
from datetime import date, datetime
from typing import Optional

_DATA_DIR = os.path.join(os.path.dirname(__file__), "..", "data")
with open(os.path.join(_DATA_DIR, "detracciones.json"), encoding="utf-8") as f:
    DETRACCIONES = json.load(f)
with open(os.path.join(_DATA_DIR, "tipos_comprobante.json"), encoding="utf-8") as f:
    TIPOS_CP = json.load(f)

IGV_TASA = 0.18
TOLERANCIA = 0.05
LIM_BANC_PEN = 3500
LIM_BANC_USD = 1000
UIT = {2024: 5150, 2025: 5350, 2026: 5350}

TASAS_CONOCIDAS = {
    0.18: "18% (estandar)",
    0.10: "10% (Amazonia/turismo)",
    0.08: "8% (Amazonia)",
    0.04: "4% (Amazonia diferencial)",
    0.00: "0% (exonerado/exportacion)",
}

def _rit(n):
    return math.floor(float(n or 0) * 100 + 0.5) / 100

def get_uit(anio=2026):
    return UIT.get(anio, 5350)

def _tasa_igv(base, igv):
    if base <= 0:
        return 0.0, "N/A"
    t = igv / base
    for k, v in TASAS_CONOCIDAS.items():
        if abs(t - k) < 0.005:
            return k, v
    return t, f"{t:.2%} (tasa no estandar)"

# --- VALIDACIONES ---

def validar_calculo_igv(base, igv, total, moneda="PEN"):
    obs = []
    ok = True
    if base <= 0:
        return {"ok": True, "tasa_detectada": 0, "diferencia_igv": 0,
                "diferencia_total": 0, "observaciones": []}
    tasa, desc = _tasa_igv(base, igv)
    igv_calc = _rit(base * tasa)
    dif_total = abs(total - (base + igv))

    if tasa == IGV_TASA:
        dif_igv = abs(igv - _rit(base * IGV_TASA))
        if dif_igv > TOLERANCIA:
            ok = False
            obs.append(f"IGV incorrecto: {igv:.2f} vs {_rit(base*IGV_TASA):.2f} (base {base:.2f} x 18%). Dif: {dif_igv:.2f}")
    elif tasa in TASAS_CONOCIDAS:
        obs.append(f"Tasa IGV: {desc}. Verifique regimen especial del emisor.")
    else:
        ok = False
        obs.append(f"Tasa IGV no reconocida: {tasa:.2%} (IGV {igv:.2f} / base {base:.2f}).")

    if dif_total > TOLERANCIA:
        ok = False
        obs.append(f"Total incorrecto: {total:.2f} != {base:.2f}+{igv:.2f}={base+igv:.2f}. Dif: {dif_total:.2f}")

    return {"ok": ok, "tasa_detectada": tasa, "diferencia_igv": abs(igv - igv_calc),
            "diferencia_total": dif_total, "observaciones": obs}


def validar_tipo_comprobante_cf(tipo_cp, es_rus, base_grav, igv):
    obs = []
    ok = True
    otorga_cf = True
    desc = TIPOS_CP["descripcion"].get(tipo_cp, f"Tipo {tipo_cp}")

    if tipo_cp == "03":
        otorga_cf = False
        if es_rus is True:
            ok = False
            obs.append("Boleta de contribuyente RUS: NO otorga credito fiscal (Art. 19 LIGV).")
        else:
            obs.append("Boleta de venta: verifique regimen del emisor (generalmente no otorga CF).")
    elif tipo_cp == "02" and igv > 0:
        ok = False
        obs.append("Recibo por honorarios no genera IGV. Verifique importe registrado.")
    elif tipo_cp not in TIPOS_CP["descripcion"]:
        obs.append(f"Tipo de comprobante '{tipo_cp}' no reconocido en tabla SUNAT.")

    return {"ok": ok, "tipo_cp": tipo_cp, "descripcion": desc,
            "otorga_cf": otorga_cf, "observaciones": obs}


def validar_bancarizacion(importe, moneda, tipo_cp, tiene_medio_pago=None):
    obs = []
    ok = True
    moneda = (moneda or "PEN").upper()
    limite = LIM_BANC_PEN if moneda == "PEN" else LIM_BANC_USD

    if importe > limite:
        if tiene_medio_pago is False:
            ok = False
            obs.append(f"Sin bancarizar: {importe:.2f} {moneda} > {limite:,}. Pierde CF y gasto (Art. 8 Ley 28194).")
        elif tiene_medio_pago is None:
            obs.append(f"Monto {importe:.2f} {moneda} supera limite ({limite:,}). Confirme medio de pago bancarizado.")
        else:
            obs.append(f"Bancarizado OK. Monto {importe:.2f} {moneda}.")

    return {"ok": ok, "requiere": importe > limite, "limite": limite, "observaciones": obs}


def validar_detraccion(tipo_cp, descripcion, importe, moneda,
                       tiene_constancia=None, codigo=None):
    obs = []
    ok = True
    aplica = False
    tasa_dt = None
    desc_dt = ""
    moneda = (moneda or "PEN").upper()
    limite_dt = DETRACCIONES["monto_minimo_soles"] if moneda == "PEN" else DETRACCIONES["monto_minimo_dolares"]

    if tipo_cp not in ["01", "04"]:
        return {"ok": True, "aplica_detraccion": False, "tasa": None, "observaciones": []}

    if importe >= limite_dt:
        if codigo:
            for item in DETRACCIONES["bienes"] + DETRACCIONES["servicios"]:
                if item["codigo"] == str(codigo).zfill(3):
                    aplica = True
                    tasa_dt = item["tasa"]
                    desc_dt = item["descripcion"]
                    break
        if not aplica:
            obs.append(f"Monto {importe:.2f} {moneda} >= {limite_dt}. Verifique si aplica detraccion (SPOT).")

    if aplica:
        monto_dt = _rit(importe * tasa_dt / 100)
        constancia = (tiene_constancia or "").strip()
        if not constancia or constancia in ("0", "nan", "None", "-"):
            ok = False
            obs.append(f"DETRACCION: {desc_dt} ({tasa_dt}%). Monto S/{monto_dt:.2f}. "
                       f"Sin constancia: CF NO valido (Art. 16 D.Leg. 940).")
        else:
            obs.append(f"Detraccion {tasa_dt}% depositada. Constancia: {constancia}.")

    return {"ok": ok, "aplica_detraccion": aplica, "tasa": tasa_dt,
            "descripcion_dt": desc_dt, "observaciones": obs}


def validar_no_domiciliado(tipo_serv, base, igv_reg, ret_ir_reg=0, tiene_cdr=False):
    obs = []
    ok = True
    if base > 0:
        igv_calc = _rit(base * IGV_TASA)
        if abs(igv_reg - igv_calc) > TOLERANCIA:
            ok = False
            obs.append(f"IGV no domiciliado: {igv_reg:.2f} vs calculado {igv_calc:.2f}")
        ir_calc = _rit(base * 0.30)
        if ret_ir_reg == 0:
            obs.append(f"Verifique retencion IR no domiciliado (30% = {ir_calc:.2f}). Puede variar por CDI.")
        elif abs(ret_ir_reg - ir_calc) > TOLERANCIA:
            obs.append(f"Retencion IR no dom: {ret_ir_reg:.2f} vs 30%: {ir_calc:.2f}. Verifique CDI aplicable.")
    if not tiene_cdr:
        obs.append("Verifique presentacion de Formulario 617 (servicios no domiciliado).")
    return {"ok": ok, "observaciones": obs}


def validar_gasto_representacion(importe, acum_previo, ingresos_anual, anio=2026):
    uit = get_uit(anio)
    limite = min(ingresos_anual * 0.005, 40 * uit)
    nuevo_acum = acum_previo + importe
    exceso = max(0.0, nuevo_acum - limite)
    obs = []
    if exceso > 0:
        obs.append(f"Gtos. representacion acumulado {nuevo_acum:.2f} supera limite {limite:.2f}. Exceso: {exceso:.2f}.")
    return {"ok": exceso == 0, "limite": limite, "acumulado": nuevo_acum,
            "exceso": exceso, "observaciones": obs}


def validar_regalia(base, retencion_reg):
    calc = _rit(base * 0.30)
    dif = abs(retencion_reg - calc)
    ok = dif <= TOLERANCIA
    obs = [f"Retencion regalia {'correcta' if ok else 'incorrecta'}: {retencion_reg:.2f} vs 30%={calc:.2f}"]
    return {"ok": ok, "retencion_calculada": calc, "observaciones": obs}


def calcular_prorrateo(ventas_grav, ventas_no_grav, igv_comun):
    total = ventas_grav + ventas_no_grav
    if total == 0:
        return {"porcentaje": 0, "igv_aceptado": 0, "igv_no_aceptado": 0, "observaciones": []}
    pct = ventas_grav / total
    igv_ac = _rit(igv_comun * pct)
    igv_no = _rit(igv_comun - igv_ac)
    return {"porcentaje": pct, "igv_aceptado": igv_ac, "igv_no_aceptado": igv_no,
            "observaciones": [f"Prorrateo {pct:.2%}: CF aceptado {igv_ac:.2f} / no aceptado {igv_no:.2f}"]}


def validar_fila_compras(fila, contexto=None):
    if contexto is None:
        contexto = {}

    resultados = {}
    obs_total = []
    tiene_error = False

    tipo_cp  = str(fila.get("TIPO DOC") or "01").strip().zfill(2)
    base_grav = float(fila.get("ADQ GRA DES OPE GRA - BASE") or 0)
    igv_comp  = float(fila.get("ADQ GRA DES OPE GRA - IGV") or 0)
    val_no_grav = float(fila.get("VALOR NO GRABADO") or 0)
    imp_mn    = float(fila.get("IMPORTE TOTAL MN") or 0)
    imp_me    = float(fila.get("IMPORTE TOTAL ME") or 0)
    moneda    = str(fila.get("MONEDA") or "PEN").strip()
    tc_libro  = float(fila.get("TIPO DE CAMBIO") or 0)
    ruc_prov  = str(fila.get("NRO DOC PROVEEDOR") or "").strip()
    nro_const = str(fila.get("NRO CONSTANCIA DETRACCION") or
                    fila.get("NRO CONSTANCIA DETRACCION", "") or "").strip()
    nro_no_dom = str(fila.get("NRO DE COMPROB NO DOMICILIADO") or "").strip()

    importe = imp_mn if moneda == "PEN" else imp_me
    if importe == 0:
        importe = base_grav + igv_comp + val_no_grav

    # 1. IGV
    if base_grav > 0:
        r = validar_calculo_igv(base_grav, igv_comp, base_grav + igv_comp + val_no_grav)
        resultados["igv"] = r
        if not r["ok"]:
            tiene_error = True
        for o in r["observaciones"]:
            obs_total.append(f"[IGV] {o}")

    # 2. Tipo comprobante
    es_rus = contexto.get(f"rus_{ruc_prov}")
    r = validar_tipo_comprobante_cf(tipo_cp, es_rus, base_grav, igv_comp)
    resultados["tipo_cp"] = r
    if not r["ok"]:
        tiene_error = True
    for o in r["observaciones"]:
        obs_total.append(f"[TIPO] {o}")

    # 3. Bancarizacion
    r = validar_bancarizacion(importe, moneda, tipo_cp)
    resultados["bancarizacion"] = r
    if not r["ok"]:
        tiene_error = True
    for o in r["observaciones"]:
        obs_total.append(f"[BANC] {o}")

    # 4. Detraccion
    r = validar_detraccion(tipo_cp, str(fila.get("DESCRIPCION") or ""),
                           importe, moneda, nro_const)
    resultados["detraccion"] = r
    if not r["ok"]:
        tiene_error = True
    for o in r["observaciones"]:
        obs_total.append(f"[DT] {o}")

    # 5. No domiciliado
    es_no_dom = nro_no_dom and nro_no_dom not in ("0", "nan", "None", "-")
    if es_no_dom:
        r = validar_no_domiciliado(
            str(fila.get("DESCRIPCION") or ""),
            float(fila.get("ADQ GRA DES OPE - BASE") or 0),
            float(fila.get("ADQ GRA DES OPE - IGV") or 0),
            0,
        )
        resultados["no_domiciliado"] = r
        if not r["ok"]:
            tiene_error = True
        for o in r["observaciones"]:
            obs_total.append(f"[NODOM] {o}")

    # 6. Tipo de cambio
    tc_sunat = contexto.get("tc_sunat")
    if moneda != "PEN" and tc_libro > 0 and tc_sunat:
        if abs(tc_libro - tc_sunat) > 0.001:
            tiene_error = True
            obs_total.append(f"[TC] TC incorrecto: libro {tc_libro:.4f} vs SUNAT {tc_sunat:.4f}")
        else:
            obs_total.append(f"[TC] TC correcto ({tc_libro:.4f})")

    # Estado final
    tiene_alerta = any("Verifique" in o or "Confirme" in o or "Tasa IGV" in o for o in obs_total)
    if tiene_error:
        estado, sem = "ERROR", "ROJO"
    elif tiene_alerta:
        estado, sem = "ALERTA", "AMARILLO"
    else:
        estado, sem = "OK", "VERDE"

    return {"semaforo": sem, "estado": estado,
            "observaciones": obs_total, "detalle": resultados}
