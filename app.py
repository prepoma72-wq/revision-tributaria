"""
Sistema de Revisión Tributaria
Supervisor: Pedro | Contadores: 12 usuarios
Módulo Fase 1: Registro de Compras + IGV
"""

import streamlit as st
import pandas as pd
import json
import os
import io
from datetime import datetime, date

# Módulos propios
import sys
sys.path.insert(0, os.path.dirname(__file__))
from validators.igv import validar_fila_compras
from utils.excel_reader import leer_registro_compras, obtener_hojas_disponibles
from utils.excel_writer import generar_excel_evidencia_compras
from utils.sunat_api import consultar_ruc, validar_tipo_cambio, validar_comprobante_sunat

# ─── CONFIGURACIÓN ────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="Revisión Tributaria",
    page_icon="🧾",
    layout="wide",
    initial_sidebar_state="expanded",
)

DATA_DIR = os.path.join(os.path.dirname(__file__), "data")

# ─── AUTENTICACIÓN SIMPLE ─────────────────────────────────────────────────────

def cargar_usuarios() -> dict:
    path = os.path.join(DATA_DIR, "usuarios.json")
    if os.path.exists(path):
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    return {"usuarios": [], "empresas": []}

def verificar_login(username: str, password: str) -> dict | None:
    """Verifica credenciales. En producción usar bcrypt o Streamlit secrets."""
    config = cargar_usuarios()
    for u in config.get("usuarios", []):
        # En producción: bcrypt.checkpw(password.encode(), u["password_hash"].encode())
        # Para demo usamos contraseña directa guardada en secrets o usuario==contraseña
        stored = u.get("password_plain", u.get("username"))
        if u["username"] == username and password == stored:
            return u
    return None

def login_form():
    st.markdown("## 🧾 Sistema de Revisión Tributaria")
    st.markdown("---")
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown("### Iniciar sesión")
        username = st.text_input("Usuario", placeholder="Ingrese su usuario")
        password = st.text_input("Contraseña", type="password", placeholder="Ingrese su contraseña")
        if st.button("Ingresar", use_container_width=True, type="primary"):
            user = verificar_login(username, password)
            if user:
                st.session_state["usuario"] = user
                st.rerun()
            else:
                st.error("Usuario o contraseña incorrectos.")
        st.markdown("---")
        st.caption("Sistema desarrollado para revisión tributaria mensual · Perú")

# ─── ESTADO DE SESIÓN ─────────────────────────────────────────────────────────

if "usuario" not in st.session_state:
    login_form()
    st.stop()

usuario = st.session_state["usuario"]
es_supervisor = usuario.get("rol") == "supervisor"

# ─── SIDEBAR ──────────────────────────────────────────────────────────────────

with st.sidebar:
    st.markdown(f"**👤 {usuario.get('nombre', usuario.get('username'))}**")
    st.markdown(f"Rol: `{'Supervisor' if es_supervisor else 'Contador'}`")
    st.markdown("---")

    modulo = st.radio(
        "Módulo",
        ["📋 Registro de Compras", "📊 Registro de Ventas", "🏢 Dashboard Supervisor"],
        disabled=not es_supervisor if True else False,
    )

    st.markdown("---")
    if st.button("Cerrar sesión"):
        del st.session_state["usuario"]
        st.rerun()

# ─── MÓDULO: REGISTRO DE COMPRAS ──────────────────────────────────────────────

if modulo == "📋 Registro de Compras":

    st.title("📋 Revisión — Registro de Compras")
    st.markdown("Carga el Excel exportado de tu sistema contable y ejecuta todas las validaciones automáticamente.")

    # ── Datos de la revisión ──
    with st.expander("📌 Datos de la revisión", expanded=True):
        col1, col2, col3 = st.columns(3)
        with col1:
            empresa_ruc = st.text_input("RUC empresa", placeholder="20XXXXXXXXX")
            empresa_nombre = st.text_input("Razón social", placeholder="EMPRESA SAC")
        with col2:
            periodo_mes = st.selectbox("Mes", range(1, 13), format_func=lambda x: [
                "Enero","Febrero","Marzo","Abril","Mayo","Junio",
                "Julio","Agosto","Setiembre","Octubre","Noviembre","Diciembre"][x-1])
            periodo_anio = st.number_input("Año", min_value=2020, max_value=2030, value=datetime.now().year)
        with col3:
            contador_nombre = st.text_input("Contador revisor", value=usuario.get("nombre", ""))
            usar_sol = st.checkbox("Validar comprobantes con Clave SOL", value=False)

        periodo = f"{['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Setiembre','Octubre','Noviembre','Diciembre'][periodo_mes-1]} {periodo_anio}"

    # Clave SOL (si activado)
    sol_ruc = sol_usuario = sol_clave = ""
    if usar_sol:
        with st.expander("🔑 Credenciales Clave SOL", expanded=True):
            st.warning("Las credenciales se usan solo para la consulta en esta sesión y no se almacenan.")
            col1, col2, col3 = st.columns(3)
            with col1:
                sol_ruc = st.text_input("RUC SOL", placeholder="RUC del representante")
            with col2:
                sol_usuario = st.text_input("Usuario SOL")
            with col3:
                sol_clave = st.text_input("Clave SOL", type="password")

    # ── Carga del archivo ──
    st.markdown("---")
    archivo = st.file_uploader(
        "📂 Carga el Registro de Compras (Excel .xlsx o .xlsm)",
        type=["xlsx", "xlsm", "xls"],
    )

    if archivo:
        try:
            # Detectar hojas
            hojas = obtener_hojas_disponibles(archivo)
            hoja_sugerida = next(
                (h for h in hojas if "COMPRA" in h.upper() or "II-1" in h.upper()), hojas[0]
            )
            col1, col2 = st.columns([3, 1])
            with col1:
                hoja_sel = st.selectbox("Hoja del Registro de Compras:", hojas,
                                        index=hojas.index(hoja_sugerida) if hoja_sugerida in hojas else 0)
            with col2:
                fila_inicio = st.number_input("Fila de encabezado (opcional)", min_value=1, value=10, step=1)

            if st.button("🔍 Ejecutar validación completa", type="primary", use_container_width=True):
                with st.spinner("Cargando y procesando el registro..."):
                    archivo.seek(0)
                    df, _ = leer_registro_compras(archivo, nombre_hoja=hoja_sel)

                st.success(f"✔ {len(df)} comprobantes cargados.")

                # ── Consultar RUCs únicos ──
                rucs_unicos = df["NRO DOC PROVEEDOR"].dropna().unique().tolist() if "NRO DOC PROVEEDOR" in df.columns else []
                contexto_rucs = {}

                if rucs_unicos:
                    with st.spinner(f"Consultando {len(rucs_unicos)} RUCs en SUNAT..."):
                        barra_ruc = st.progress(0)
                        for i, ruc in enumerate(rucs_unicos):
                            datos_ruc = consultar_ruc(str(ruc))
                            contexto_rucs[ruc] = datos_ruc
                            if datos_ruc.get("regimen"):
                                es_rus = "RUS" in datos_ruc["regimen"].upper()
                                contexto_rucs[f"rus_{ruc}"] = es_rus
                            barra_ruc.progress((i + 1) / len(rucs_unicos))
                        barra_ruc.empty()

                # ── Obtener tipo de cambio por fechas únicas ──
                fechas_tc = {}
                if "FECHA DE EMISIÓN" in df.columns:
                    fechas_unicas = df["FECHA DE EMISIÓN"].dropna().unique()
                    with st.spinner(f"Obteniendo tipos de cambio SUNAT ({len(fechas_unicas)} fechas)..."):
                        for fecha in fechas_unicas:
                            if isinstance(fecha, date):
                                tc_info = validar_tipo_cambio(0, fecha)  # Solo para obtener el TC
                                fechas_tc[str(fecha)] = tc_info.get("tc_sunat")

                # ── Ejecutar validaciones ──
                with st.spinner("Ejecutando validaciones tributarias..."):
                    resultados = []
                    barra_val = st.progress(0)
                    total = len(df)

                    for i, (_, fila) in enumerate(df.iterrows()):
                        # Agregar contexto RUC
                        ruc_prov = str(fila.get("NRO DOC PROVEEDOR", ""))
                        ctx = {**contexto_rucs}

                        # Agregar TC SUNAT para la fecha
                        fecha_em = fila.get("FECHA DE EMISIÓN")
                        tc_sunat = fechas_tc.get(str(fecha_em)) if fecha_em else None
                        ctx["tc_sunat"] = tc_sunat

                        resultado = validar_fila_compras(fila.to_dict(), ctx)

                        # Agregar info de RUC al resultado
                        datos_ruc = contexto_rucs.get(ruc_prov, {})
                        resultado["ruc"] = {
                            "activo": datos_ruc.get("activo"),
                            "habido": datos_ruc.get("habido"),
                            "condicion": f"{'ACTIVO' if datos_ruc.get('activo') else 'INACTIVO'} / {'HABIDO' if datos_ruc.get('habido') else 'NO HABIDO'}" if datos_ruc else "N/V",
                        }
                        if datos_ruc.get("activo") is False or datos_ruc.get("habido") is False:
                            resultado["estado"] = "ERROR"
                            resultado["semaforo"] = "🔴"
                            msg = f"✘ RUC {ruc_prov} {datos_ruc.get('condicion', '')}: crédito fiscal NO válido."
                            resultado["observaciones"].insert(0, f"[RUC] {msg}")

                        resultados.append(resultado)
                        barra_val.progress((i + 1) / total)

                    barra_val.empty()

                # Guardar en sesión
                st.session_state["df_compras"] = df
                st.session_state["resultados_compras"] = resultados
                st.session_state["empresa_nombre"] = empresa_nombre
                st.session_state["empresa_ruc"] = empresa_ruc
                st.session_state["periodo"] = periodo
                st.session_state["contador_nombre"] = contador_nombre

        except Exception as e:
            st.error(f"Error al procesar el archivo: {e}")
            st.exception(e)

    # ── Mostrar resultados ──
    if "df_compras" in st.session_state and "resultados_compras" in st.session_state:
        df = st.session_state["df_compras"]
        resultados = st.session_state["resultados_compras"]

        total = len(resultados)
        errores = sum(1 for r in resultados if r.get("estado") == "ERROR")
        alertas = sum(1 for r in resultados if r.get("estado") == "ALERTA")
        ok_count = sum(1 for r in resultados if r.get("estado") == "OK")

        # KPIs
        st.markdown("---")
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Total revisados", total)
        col2.metric("✅ Conformes", ok_count, f"{ok_count/total*100:.0f}%" if total else "")
        col3.metric("⚠️ Alertas", alertas)
        col4.metric("🔴 Errores", errores)

        # Filtros
        st.markdown("#### Filtrar resultados")
        col1, col2 = st.columns(2)
        with col1:
            filtro_estado = st.multiselect(
                "Estado", ["OK", "ALERTA", "ERROR"],
                default=["ALERTA", "ERROR"] if errores + alertas > 0 else ["OK", "ALERTA", "ERROR"]
            )
        with col2:
            buscar_razon = st.text_input("Buscar razón social o RUC", placeholder="Escriba para filtrar...")

        # Construir tabla de resultados
        filas_tabla = []
        cols_mostrar_tabla = [
            "NUM CORRELATIVO", "FECHA DE EMISIÓN", "TIPO DOC", "SERIE",
            "NRO COMPROBANTE", "NRO DOC PROVEEDOR", "RAZÓN SOCIAL",
            "ADQ GRA DES OPE GRA - BASE", "ADQ GRA DES OPE GRA - IGV", "IMPORTE TOTAL MN",
        ]
        cols_tabla = [c for c in cols_mostrar_tabla if c in df.columns]

        for i, (_, fila) in enumerate(df.iterrows()):
            r = resultados[i] if i < len(resultados) else {}
            estado = r.get("estado", "")
            semaforo = r.get("semaforo", "")

            if filtro_estado and estado not in filtro_estado:
                continue

            razon = str(fila.get("RAZÓN SOCIAL", ""))
            ruc = str(fila.get("NRO DOC PROVEEDOR", ""))
            if buscar_razon and buscar_razon.lower() not in razon.lower() and buscar_razon not in ruc:
                continue

            fila_dict = {c: fila.get(c, "") for c in cols_tabla}
            fila_dict["🚦"] = semaforo
            fila_dict["Estado"] = estado
            fila_dict["Observaciones"] = " | ".join(r.get("observaciones", []))[:200]
            filas_tabla.append(fila_dict)

        if filas_tabla:
            df_tabla = pd.DataFrame(filas_tabla)
            # Mostrar semáforo primero
            cols_orden = ["🚦", "Estado"] + cols_tabla + ["Observaciones"]
            cols_orden = [c for c in cols_orden if c in df_tabla.columns]
            st.dataframe(
                df_tabla[cols_orden],
                use_container_width=True,
                height=500,
                column_config={
                    "🚦": st.column_config.TextColumn(width="small"),
                    "Estado": st.column_config.TextColumn(width="small"),
                    "Observaciones": st.column_config.TextColumn(width="large"),
                }
            )
        else:
            st.info("No hay registros con los filtros seleccionados.")

        # ── Exportar evidencia ──
        st.markdown("---")
        st.markdown("### 📥 Exportar evidencia")
        col1, col2 = st.columns(2)

        with col1:
            if st.button("📊 Generar Excel de evidencia", type="primary", use_container_width=True):
                with st.spinner("Generando Excel de evidencia..."):
                    excel_bytes = generar_excel_evidencia_compras(
                        df_original=df,
                        resultados_validacion=resultados,
                        empresa_nombre=st.session_state.get("empresa_nombre", ""),
                        empresa_ruc=st.session_state.get("empresa_ruc", ""),
                        periodo=st.session_state.get("periodo", ""),
                        contador_nombre=st.session_state.get("contador_nombre", ""),
                        supervisor_nombre="Pedro - Supervisor" if es_supervisor else "",
                        fecha_revision=datetime.now(),
                    )
                    nombre_archivo = (
                        f"REV_COMPRAS_{st.session_state.get('empresa_ruc','')}_{periodo.replace(' ','_')}.xlsx"
                    )
                    st.download_button(
                        label="⬇️ Descargar Excel",
                        data=excel_bytes,
                        file_name=nombre_archivo,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )


# ─── MÓDULO: REGISTRO DE VENTAS ───────────────────────────────────────────────

elif modulo == "📊 Registro de Ventas":
    st.title("📊 Revisión — Registro de Ventas")
    st.info("🚧 Módulo en construcción — Fase 2. Disponible próximamente.")


# ─── MÓDULO: DASHBOARD SUPERVISOR ─────────────────────────────────────────────

elif modulo == "🏢 Dashboard Supervisor":
    if not es_supervisor:
        st.error("Acceso restringido al supervisor.")
        st.stop()

    st.title("🏢 Dashboard Supervisor")
    st.markdown("Vista general de todas las revisiones del período.")
    st.info("📊 El dashboard mostrará el estado de revisión de las 20 empresas y los 12 contadores en tiempo real. Disponible en Fase 2 con integración de base de datos centralizada.")

    st.markdown("#### Estado actual")
    config = cargar_usuarios()
    empresas = config.get("empresas", [])
    if empresas:
        df_emp = pd.DataFrame(empresas)
        st.dataframe(df_emp, use_container_width=True)
    else:
        st.markdown("No hay empresas registradas aún. Agrega empresas en `data/usuarios.json`.")
