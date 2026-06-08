"""
Módulo de integración con APIs de SUNAT.
Cubre: tipo de cambio, consulta RUC, validación de comprobantes electrónicos.
"""

import requests
import json
from datetime import date, datetime
from typing import Optional
import re
import time

# ─────────────────────────────────────────────
# TIPO DE CAMBIO
# ─────────────────────────────────────────────

def obtener_tipo_cambio(fecha: date) -> dict:
    """
    Obtiene el tipo de cambio compra/venta del USD para una fecha dada.
    Fuente: API pública de SUNAT (via apis.net.pe como proxy confiable).
    Retorna: {"compra": float, "venta": float, "fecha": str, "fuente": str}
    """
    fecha_str = fecha.strftime("%Y-%m-%d")
    result = {"compra": None, "venta": None, "fecha": fecha_str, "fuente": None, "error": None}

    # Fuente 1: APIs.net.pe (agrega datos de SBS/SUNAT, sin auth requerida)
    try:
        url = f"https://api.apis.net.pe/v1/tipo-cambio?fecha={fecha_str}"
        headers = {"Referer": "https://apis.net.pe"}
        r = requests.get(url, headers=headers, timeout=8)
        if r.status_code == 200:
            data = r.json()
            result["compra"] = float(data.get("compra", 0))
            result["venta"] = float(data.get("venta", 0))
            result["fuente"] = "SUNAT via apis.net.pe"
            return result
    except Exception:
        pass

    # Fuente 2: SBS directamente
    try:
        url_sbs = (
            f"https://www.sbs.gob.pe/app/pp/SISTIP_PORTAL/Paginas/Tipo-Cambio-Promedio/"
            f"TipoCambioPromedioBancario.aspx"
        )
        # Si falla la API, usamos un valor de referencia y marcamos como "no verificado"
        result["error"] = "No se pudo obtener TC de SUNAT/SBS. Verifique manualmente."
    except Exception as e:
        result["error"] = str(e)

    return result


def validar_tipo_cambio(tc_libro: float, fecha: date, tolerancia: float = 0.001) -> dict:
    """
    Compara el TC del libro con el TC oficial de SUNAT.
    Retorna: {"ok": bool, "tc_sunat": float, "tc_libro": float, "diferencia": float, "mensaje": str}
    """
    tc_info = obtener_tipo_cambio(fecha)

    if tc_info.get("error") or not tc_info.get("venta"):
        return {
            "ok": None,
            "tc_sunat": None,
            "tc_libro": tc_libro,
            "diferencia": None,
            "mensaje": f"⚠ No se pudo verificar TC SUNAT: {tc_info.get('error', 'Sin datos')}",
        }

    tc_sunat = tc_info["venta"]  # Para compras se usa TC venta
    diferencia = abs(tc_libro - tc_sunat)

    if diferencia <= tolerancia:
        return {
            "ok": True,
            "tc_sunat": tc_sunat,
            "tc_libro": tc_libro,
            "diferencia": diferencia,
            "mensaje": f"✔ TC correcto ({tc_libro} vs SUNAT {tc_sunat})",
        }
    else:
        return {
            "ok": False,
            "tc_sunat": tc_sunat,
            "tc_libro": tc_libro,
            "diferencia": diferencia,
            "mensaje": f"✘ TC incorrecto. Libro: {tc_libro} | SUNAT: {tc_sunat} | Dif: {diferencia:.4f}",
        }


# ─────────────────────────────────────────────
# CONSULTA RUC
# ─────────────────────────────────────────────

def consultar_ruc(ruc: str) -> dict:
    """
    Consulta el estado de un RUC en SUNAT.
    Retorna: {"activo": bool, "habido": bool, "razon_social": str, "condicion": str,
              "estado": str, "tipo_contribuyente": str, "es_agente_retencion": bool,
              "regimen": str, "error": str}
    """
    ruc = str(ruc).strip().zfill(11) if ruc else ""
    result = {
        "activo": None, "habido": None, "razon_social": "",
        "condicion": "", "estado": "", "tipo_contribuyente": "",
        "es_agente_retencion": False, "regimen": "", "error": None
    }

    if not ruc or len(ruc) != 11 or not ruc.isdigit():
        result["error"] = "RUC inválido (debe tener 11 dígitos)"
        return result

    # Fuente: apis.net.pe (datos de SUNAT, sin auth)
    try:
        url = f"https://api.apis.net.pe/v2/sunat/ruc?numero={ruc}"
        headers = {"Referer": "https://apis.net.pe", "Authorization": "Bearer apis-token-12345.fVEFiGOdaKv4"}
        r = requests.get(url, timeout=10)
        if r.status_code == 200:
            data = r.json()
            condicion = data.get("condicion", "").upper()
            estado = data.get("estado", "").upper()
            result["razon_social"] = data.get("razonSocial", "")
            result["condicion"] = condicion
            result["estado"] = estado
            result["activo"] = "ACTIVO" in estado
            result["habido"] = "HABIDO" in condicion
            result["tipo_contribuyente"] = data.get("tipoContribuyente", "")
            result["regimen"] = data.get("sistemaEmision", "")
            result["es_agente_retencion"] = data.get("esAgenteRetencion", False)
            return result
        elif r.status_code == 422:
            result["error"] = "RUC no encontrado en SUNAT"
            return result
    except Exception as e:
        result["error"] = f"Error al consultar RUC: {str(e)}"

    return result


def es_rus(ruc: str) -> Optional[bool]:
    """Determina si el contribuyente está en el Régimen Único Simplificado (RUS)."""
    datos = consultar_ruc(ruc)
    if datos.get("error"):
        return None
    regimen = datos.get("regimen", "").upper()
    tipo = datos.get("tipo_contribuyente", "").upper()
    return "RUS" in regimen or "SIMPLIFICADO" in regimen


# ─────────────────────────────────────────────
# VALIDACIÓN COMPROBANTE ELECTRÓNICO (SOL)
# ─────────────────────────────────────────────

def validar_comprobante_sunat(
    ruc_emisor: str,
    tipo_comprobante: str,
    serie: str,
    numero: str,
    importe_total: float,
    clave_sol_ruc: str,
    clave_sol_usuario: str,
    clave_sol_clave: str,
) -> dict:
    """
    Valida un comprobante electrónico en SUNAT usando Clave SOL.
    Llama al servicio REST de consulta de validez de comprobantes.

    Retorna: {"valido": bool, "estado": str, "mensaje": str}
    """
    result = {"valido": None, "estado": "", "mensaje": ""}

    # Endpoint SUNAT para consulta de comprobantes
    url = "https://e-factura.sunat.gob.pe/v1/contribuyente/gre/v1.0/contribuyente/gem/comprobantes/validarcomprobante"

    # Primero obtenemos token OAuth de SUNAT
    try:
        token_url = "https://api-seguridad.sunat.gob.pe/v1/clientessol/{ruc}/oauth2/token/".format(
            ruc=clave_sol_ruc
        )
        token_payload = {
            "grant_type": "password",
            "scope": "https://api.sunat.gob.pe/v1/contribuyente/contribuyentes",
            "client_id": "test-fiscalizacion",
            "username": f"{clave_sol_ruc}{clave_sol_usuario}",
            "password": clave_sol_clave,
        }
        token_r = requests.post(token_url, data=token_payload, timeout=15)

        if token_r.status_code != 200:
            result["mensaje"] = "⚠ No se pudo autenticar con SUNAT SOL. Verifique credenciales."
            return result

        token = token_r.json().get("access_token")

        # Consulta de validez del comprobante
        consulta_url = (
            f"https://api.sunat.gob.pe/v1/contribuyente/contribuyentes/{clave_sol_ruc}"
            f"/validarcomprobante"
        )
        payload = {
            "numRuc": ruc_emisor,
            "codComp": tipo_comprobante,
            "numeroSerie": serie,
            "numero": numero,
            "fechaEmision": "",  # Opcional
            "monto": str(importe_total),
        }
        headers = {"Authorization": f"Bearer {token}"}
        r = requests.get(consulta_url, params=payload, headers=headers, timeout=15)

        if r.status_code == 200:
            data = r.json()
            estado_cp = data.get("data", {}).get("estadoCp", "")
            observacion = data.get("data", {}).get("observaciones", "")

            if estado_cp == "1":
                result["valido"] = True
                result["estado"] = "ACEPTADO"
                result["mensaje"] = "✔ Comprobante aceptado en SUNAT"
            elif estado_cp == "2":
                result["valido"] = False
                result["estado"] = "ANULADO"
                result["mensaje"] = f"✘ Comprobante ANULADO en SUNAT. {observacion}"
            elif estado_cp == "3":
                result["valido"] = False
                result["estado"] = "RECHAZADO"
                result["mensaje"] = f"✘ Comprobante RECHAZADO. {observacion}"
            else:
                result["valido"] = None
                result["estado"] = "NO ENCONTRADO"
                result["mensaje"] = f"⚠ Comprobante no encontrado en SUNAT. Estado: {estado_cp}"
        else:
            result["mensaje"] = f"⚠ SUNAT respondió con código {r.status_code}"

    except requests.exceptions.Timeout:
        result["mensaje"] = "⚠ Tiempo de espera agotado al consultar SUNAT"
    except Exception as e:
        result["mensaje"] = f"⚠ Error al validar comprobante: {str(e)}"

    return result


def validar_comprobante_batch(
    comprobantes: list,
    clave_sol_ruc: str,
    clave_sol_usuario: str,
    clave_sol_clave: str,
    delay_segundos: float = 0.3,
) -> list:
    """
    Valida una lista de comprobantes con pausa entre llamadas para no saturar SUNAT.
    comprobantes: lista de dicts con keys: ruc_emisor, tipo_comprobante, serie, numero, importe_total
    """
    resultados = []
    for cp in comprobantes:
        res = validar_comprobante_sunat(
            ruc_emisor=cp.get("ruc_emisor", ""),
            tipo_comprobante=cp.get("tipo_comprobante", ""),
            serie=cp.get("serie", ""),
            numero=cp.get("numero", ""),
            importe_total=cp.get("importe_total", 0),
            clave_sol_ruc=clave_sol_ruc,
            clave_sol_usuario=clave_sol_usuario,
            clave_sol_clave=clave_sol_clave,
        )
        resultados.append(res)
        time.sleep(delay_segundos)
    return resultados
