# Sistema de Revisión Tributaria — Instrucciones de Instalación

## Opción A: Streamlit Cloud (recomendado — acceso web para todos)

### Paso 1: Crear cuenta GitHub (si no tienes)
1. Ve a https://github.com y crea una cuenta gratuita.

### Paso 2: Subir el proyecto
1. En GitHub, crea un nuevo repositorio: "revision-tributaria"
2. Sube todos los archivos de esta carpeta (`revision_tributaria/`) al repositorio.

### Paso 3: Publicar en Streamlit Cloud
1. Ve a https://share.streamlit.io
2. Inicia sesión con tu cuenta GitHub.
3. Haz clic en "New app".
4. Selecciona el repositorio "revision-tributaria".
5. En "Main file path" escribe: `app.py`
6. Haz clic en "Deploy".
7. En ~2 minutos tendrás una URL como: `https://revision-tributaria-pedro.streamlit.app`

### Paso 4: Configurar usuarios
Edita el archivo `data/usuarios.json` con los usuarios de tus 12 contadores:

```json
{
  "usuarios": [
    {
      "username": "pedro",
      "nombre": "Pedro García - Supervisor",
      "rol": "supervisor",
      "password_plain": "TU_CLAVE_AQUI"
    },
    {
      "username": "contador1",
      "nombre": "María López",
      "rol": "contador",
      "password_plain": "clave_contador1"
    }
  ],
  "empresas": [
    {"ruc": "20467225155", "razon_social": "VISUAL IMPACT S.A.C.", "contador_asignado": "contador1"}
  ]
}
```

### Paso 5: Compartir con tu equipo
Comparte la URL con tus 12 contadores. Cada uno accede con su usuario y clave.

---

## Opción B: Ejecutar localmente (sin internet)

### Requisitos
- Python 3.10 o superior instalado
- pip actualizado

### Instalación
```bash
cd revision_tributaria
pip install -r requirements.txt
streamlit run app.py
```

Esto abre el sistema en tu navegador en http://localhost:8501

---

## Flujo de trabajo para los contadores

1. **Entrar al sistema** con usuario y contraseña.
2. **Ingresar datos** de la empresa y período.
3. **Cargar el Excel** exportado del sistema contable.
4. **Seleccionar la hoja** "Registro de Compras".
5. **Ejecutar la validación** — el sistema:
   - Verifica cálculos de IGV (Base × 18%)
   - Consulta RUC del proveedor en SUNAT (activo/habido)
   - Valida tipo de cambio contra SUNAT
   - Revisa bancarización (>S/3,500 o $1,000)
   - Valida detracciones por tipo de bien/servicio
   - Verifica comprobante electrónico (con Clave SOL)
   - Revisa régimen del emisor (RUS, etc.)
6. **Revisar los resultados**: 🟢 OK | 🟡 Alerta | 🔴 Error
7. **Generar Excel de evidencia** con todo formateado y coloreado.
8. **Enviar al supervisor** para aprobación.

---

## Validaciones implementadas (Fase 1)

| Validación | Fuente | Estado |
|---|---|---|
| IGV = Base × 18% | Cálculo interno | ✅ Activo |
| Total = Base + IGV | Cálculo interno | ✅ Activo |
| RUC activo y habido | SUNAT (API pública) | ✅ Activo |
| Tipo de cambio correcto | SUNAT/SBS (API pública) | ✅ Activo |
| Tipo de comprobante válido para CF | Tabla SUNAT | ✅ Activo |
| Emisor en régimen RUS | SUNAT (API pública) | ✅ Activo |
| Bancarización >S/3,500 | Ley 28194 | ✅ Activo |
| Detracción aplicable | Tabla RS 183-2004 | ✅ Activo |
| Estado comprobante electrónico | SUNAT (Clave SOL) | ✅ Activo |
| Gastos de representación (límite) | Art. 37 LIR | ✅ Activo |
| No domiciliados (IGV + IR) | Art. 76 LIGV/LIR | ✅ Activo |
| Regalías (retención 30%) | Art. 56 LIR | ✅ Activo |

## Próximas fases

- **Fase 2**: Registro de Ventas, Impuesto a la Renta mensual, PLAME
- **Fase 3**: Dashboard supervisor en tiempo real, base de datos centralizada, notificaciones
- **Fase 4**: ITAN, no domiciliados completo, activos fijos, vinculadas
