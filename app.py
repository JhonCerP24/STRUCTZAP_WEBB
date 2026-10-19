import os
import io
import tempfile
import pandas as pd
import streamlit as st
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.drawing.image import Image as OpenpyxlImage


# IMPORTACIÓN MODULAR DESDE TU LIBRERÍA (codigo_base.py)

from codigo_base import (
    calcular_zapata,
    verificar_armadura,
    cargar_datos_archivo,
    generar_tabla_alternativas,
    generar_grafico_planta_2d,
    generar_grafico_modelo_3d,
    generar_grafico_esfuerzos_suelo,
    generar_grafico_verificacion_estructural,
    TABLA_VARILLAS
)


# CONFIGURACIÓN DE LA PÁGINA WEB Y ESTILOS VISUALES (CSS PERSONALIZADO)

st.set_page_config(
    page_title="STRUCTZAP WEB PRO | Diseño de Zapatas Aisladas",
    page_icon="🏗️",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    /* Fondo general estilo ingeniería CAD oscuro */
    .stApp {
        background-color: #070c18;
        color: #f8fafc;
    }
    /* Encabezado principal */
    .main-header {
        background: linear-gradient(90deg, #0f1b33 0%, #1e3a5f 100%);
        padding: 20px 26px;
        border-radius: 12px;
        border-left: 6px solid #00e5ff;
        margin-bottom: 20px;
        box-shadow: 0 4px 15px rgba(0, 0, 0, 0.4);
    }
    .main-header h1 {
        color: #ffffff;
        font-size: 1.85rem;
        margin: 0;
        font-weight: 800;
        letter-spacing: 0.5px;
    }
    .main-header p {
        color: #94a3b8;
        margin: 6px 0 0 0;
        font-size: 0.95rem;
    }
    /* Tarjetas Ejecutivas KPI */
    .kpi-card {
        background-color: #0f1b33;
        border: 1px solid #1e3a5f;
        border-top: 3px solid #38bdf8;
        border-radius: 10px;
        padding: 14px 16px;
        text-align: center;
        box-shadow: 0 2px 8px rgba(0,0,0,0.35);
        height: 100%;
    }
    .kpi-title {
        color: #94a3b8;
        font-size: 0.78rem;
        text-transform: uppercase;
        font-weight: 700;
        margin-bottom: 6px;
        letter-spacing: 0.5px;
    }
    .kpi-value {
        color: #f8fafc;
        font-size: 1.15rem;
        font-weight: 800;
        margin-bottom: 4px;
    }
    .kpi-sub {
        font-size: 0.80rem;
        font-weight: 600;
    }
    /* Badges de cumplimiento */
    .badge-ok {
        background-color: #065f46;
        color: #6ee7b7;
        padding: 3px 10px;
        border-radius: 12px;
        font-size: 0.78rem;
        font-weight: bold;
    }
    .badge-err {
        background-color: #7f1d1d;
        color: #fca5a5;
        padding: 3px 10px;
        border-radius: 12px;
        font-size: 0.78rem;
        font-weight: bold;
    }
    /* Contenedores de gráficos */
    .chart-box {
        background-color: #0b1325;
        border: 1px solid #1e3a5f;
        border-radius: 10px;
        padding: 10px;
        margin-bottom: 12px;
    }
</style>
""", unsafe_allow_html=True)


# ESTADO DE SESIÓN (DATOS POR DEFECTO DEL PROYECTO)

CARGAS_DEFECTO = pd.DataFrame(
    [
        [32.5810, 0.1533, 1.0600],  # Carga Muerta (CM)
        [9.8700,  0.0154, 0.0118],  # Carga Viva (CV)
        [8.8031,  2.7582, 10.9464], # Sismo en X (Sx)
        [15.9962, 9.2800, 0.5030]   # Sismo en Y (Sy)
    ],
    columns=["P", "Mx", "My"],
    index=["Carga Muerta (CM)", "Carga Viva (CV)", "Sismo en X (Sx)", "Sismo en Y (Sy)"]
)

def inicializar_estado():
    defaults = {
        "df_cargas": CARGAS_DEFECTO.copy(),
        "q_adm": 1.80,
        "fc": 210.0,
        "fy": 4200.0,
        "h_inicial": 0.60,
        "l_c": 0.30,
        "b_c": 0.60,
        "varilla": '5/8"',
        "espaciamiento": 0.150,
        "origen_datos_txt": "Datos predefinidos del proyecto (Editable manualmente)"
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

def restaurar_valores():
    st.session_state["df_cargas"] = CARGAS_DEFECTO.copy()
    st.session_state["q_adm"] = 1.80
    st.session_state["fc"] = 210.0
    st.session_state["fy"] = 4200.0
    st.session_state["h_inicial"] = 0.60
    st.session_state["l_c"] = 0.30
    st.session_state["b_c"] = 0.60
    st.session_state["varilla"] = '5/8"'
    st.session_state["espaciamiento"] = 0.150
    st.session_state["origen_datos_txt"] = "Datos predefinidos del proyecto (Restaurados)"

inicializar_estado()


# FUNCIÓN GENERADORA DEL REPORTE EXCEL EN MEMORIA (PARA DESCARGA WEB)

def generar_excel_bytes(r):
    arm = r["armadura"]
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Memoria de Cálculo"

    font_titulo = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
    font_seccion = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    font_header = Font(name="Calibri", size=10, bold=True, color="0F213E")
    font_bold = Font(name="Calibri", size=10, bold=True, color="1E293B")
    font_normal = Font(name="Calibri", size=10, color="1E293B")

    fill_titulo = PatternFill("solid", fgColor="0F213E")
    fill_seccion = PatternFill("solid", fgColor="1E3A5F")
    fill_subhdr = PatternFill("solid", fgColor="D6EAF8")
    fill_ok = PatternFill("solid", fgColor="D1FAE5")
    fill_err = PatternFill("solid", fgColor="FEE2E2")

    borde_fino = Border(
        left=Side(style='thin', color='CBD5E1'),
        right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'),
        bottom=Side(style='thin', color='CBD5E1')
    )

    ws.merge_cells("A1:D1")
    ws["A1"] = "REPORTE DE DISEÑO Y VERIFICACIÓN ESTRUCTURAL DE ZAPATA AISLADA"
    ws["A1"].font = font_titulo
    ws["A1"].fill = fill_titulo
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    # 1. Geometría y Parámetros
    ws.merge_cells("A3:D3")
    ws["A3"] = "1. PARÁMETROS GEOTÉCNICOS Y DIMENSIONES FINALES"
    ws["A3"].font = font_seccion
    ws["A3"].fill = fill_seccion

    filas_geo = [
        ("Capacidad Admisible (q_adm)", f"{r['qadm']:.2f} kg/cm² ({r['Qatn']:.1f} Tn/m²)", "Resistencia Concreto (f'c)", f"{r['fc']:.0f} kg/cm²"),
        ("Fluencia del Acero (fy)", f"{r['fy']:.0f} kg/cm²", "Sección de Columna (Lc × Bc)", f"{r['L_C']:.2f} × {r['B_C']:.2f} m"),
        ("Base Final de Zapata (B)", f"{r['Base_final']:.2f} m", "Largo Final de Zapata (L)", f"{r['Largo_final']:.2f} m"),
        ("Peralte Total (H)", f"{r['Peralte_H']:.2f} m", "Peralte Efectivo (d)", f"{r['d_efectivo']:.2f} m"),
        ("Volado 1 (Dirección Y)", f"{r['Volado_1']:.2f} m", "Volado 2 (Dirección X)", f"{r['Volado_2']:.2f} m"),
    ]
    for idx, (k1, v1, k2, v2) in enumerate(filas_geo, start=4):
        ws.cell(row=idx, column=1, value=k1).font = font_bold
        ws.cell(row=idx, column=2, value=v1).font = font_normal
        ws.cell(row=idx, column=3, value=k2).font = font_bold
        ws.cell(row=idx, column=4, value=v2).font = font_normal
        for col in range(1, 5):
            ws.cell(row=idx, column=col).border = borde_fino

    # 2. Verificación de Presiones en el Suelo
    ws.merge_cells("A10:D10")
    ws["A10"] = "2. VERIFICACIÓN DE ESFUERZOS TRANSMITIDOS AL SUELO"
    ws["A10"].font = font_seccion
    ws["A10"].fill = fill_seccion

    headers_suelo = ["Estado de Carga", "σ máx (Tn/m²)", "σ mín (Tn/m²)", "Límite Admisible (Tn/m²)"]
    for col, h_txt in enumerate(headers_suelo, start=1):
        c = ws.cell(row=11, column=col, value=h_txt)
        c.font = font_header
        c.fill = fill_subhdr
        c.border = borde_fino

    datos_suelo = [
        ("Carga de Servicio (CM + CV)", r['esfuerzo_maximo_CS'], r['esfuerzo_minimo_CS'], r['Qatn']),
        ("Sismo en Dirección X (CS + 0.8 Sx)", r['esfuerzo_maximo_Sx'], r['esfuerzo_minimo_Sx'], r['Qatn_Sx']),
        ("Sismo en Dirección Y (CS + 0.8 Sy)", r['esfuerzo_maximo_Sy'], r['esfuerzo_minimo_Sy'], r['Qatn_Sx']),
    ]
    for idx, (nom, smax, smin, qlim) in enumerate(datos_suelo, start=12):
        ws.cell(row=idx, column=1, value=nom).font = font_bold
        ws.cell(row=idx, column=2, value=smax).font = font_normal
        ws.cell(row=idx, column=3, value=smin).font = font_normal
        ws.cell(row=idx, column=4, value=qlim).font = font_normal
        for col in range(1, 5):
            ws.cell(row=idx, column=col).border = borde_fino

    # 3. Corte y Punzonamiento
    ws.merge_cells("A16:D16")
    ws["A16"] = "3. VERIFICACIÓN POR CORTE UNIDIRECCIONAL Y PUNZONAMIENTO BIDIRECCIONAL"
    ws["A16"].font = font_seccion
    ws["A16"].fill = fill_seccion

    headers_corte = ["Sección Crítica", "Demanda Vu (Tn)", "Capacidad φVc (Tn)", "Estado"]
    for col, h_txt in enumerate(headers_corte, start=1):
        c = ws.cell(row=17, column=col, value=h_txt)
        c.font = font_header
        c.fill = fill_subhdr
        c.border = borde_fino

    datos_corte = [
        ("Corte Unidireccional Volado 1 (Y)", r['V_u1'], r['φVc1'], "CUMPLE" if r['cumple_corte_1'] else "NO CUMPLE"),
        ("Corte Unidireccional Volado 2 (X)", r['V_u2'], r['φVc2'], "CUMPLE" if r['cumple_corte_2'] else "NO CUMPLE"),
        (f"Punzonamiento Bidireccional (b₀={r['Pm']}m)", r['Vu_Punzonamiento'], r['φVc_Punzonamiento'], "CUMPLE" if r['cumple_punzonamiento'] else "NO CUMPLE"),
    ]
    for idx, (sec, vu, pvc, est) in enumerate(datos_corte, start=18):
        ws.cell(row=idx, column=1, value=sec).font = font_bold
        ws.cell(row=idx, column=2, value=vu).font = font_normal
        ws.cell(row=idx, column=3, value=pvc).font = font_normal
        c_est = ws.cell(row=idx, column=4, value=est)
        c_est.font = font_bold
        c_est.fill = fill_ok if est == "CUMPLE" else fill_err
        for col in range(1, 5):
            ws.cell(row=idx, column=col).border = borde_fino

    # 4. Armadura por Flexión
    ws.merge_cells("A22:D22")
    ws["A22"] = "4. DISEÑO DE ARMADURA PRINCIPAL POR FLEXIÓN"
    ws["A22"].font = font_seccion
    ws["A22"].fill = fill_seccion

    filas_acero = [
        ("Distribución Elegida", arm['descripcion'], "Momento Último (Mu)", f"{arm['MU_Tnm']:.2f} Tn·m"),
        ("Acero Colocado (As)", f"{arm['As_colocado']:.2f} cm²/m", "Capacidad Resistente (φMn)", f"{arm['φMn_Tnm']:.2f} Tn·m"),
        ("Acero Mínimo (As mín)", f"{arm['As_minimo']:.2f} cm²/m", "Estado de Verificación", arm['estado']),
    ]
    for idx, (k1, v1, k2, v2) in enumerate(filas_acero, start=23):
        ws.cell(row=idx, column=1, value=k1).font = font_bold
        ws.cell(row=idx, column=2, value=v1).font = font_normal
        ws.cell(row=idx, column=3, value=k2).font = font_bold
        c_v2 = ws.cell(row=idx, column=4, value=v2)
        c_v2.font = font_bold if idx == 25 else font_normal
        if idx == 25:
            c_v2.fill = fill_ok if arm['es_favorable'] else fill_err
        for col in range(1, 5):
            ws.cell(row=idx, column=col).border = borde_fino

    for col_let, w_val in {"A": 34, "B": 24, "C": 28, "D": 28}.items():
        ws.column_dimensions[col_let].width = w_val

    # Incrustar los 4 gráficos en el archivo Excel
    rutas_graficos = [
        (r.get("ruta_img_2d"), "F2"),
        (r.get("ruta_img_3d"), "M2"),
        (r.get("ruta_img_esfuerzos"), "F22"),
        (r.get("ruta_img_estructural"), "M22"),
    ]
    for ruta_g, celda_ancla in rutas_graficos:
        if ruta_g and os.path.exists(ruta_g):
            img_xl = OpenpyxlImage(ruta_g)
            img_xl.width = int(img_xl.width * 0.46)
            img_xl.height = int(img_xl.height * 0.46)
            img_xl.anchor = celda_ancla
            ws.add_image(img_xl)

    # Hoja 2: Matriz de Alternativas de Acero
    ws_alt = wb.create_sheet(title="Alternativas de Acero")
    df_alt = r["tabla_alternativas"]
    cols_export = ["Varilla", "Espaciamiento_m", "As_cm2_m", "As_min_cm2_m", "φMn_Tnm", "Mu_Tnm", "Ratio", "Estado"]
    for c_idx, nombre_col in enumerate(cols_export, start=1):
        celda = ws_alt.cell(row=1, column=c_idx, value=nombre_col)
        celda.font = font_seccion
        celda.fill = fill_seccion
        ws_alt.column_dimensions[openpyxl.utils.get_column_letter(c_idx)].width = 20

    for r_idx, (_, fila_df) in enumerate(df_alt.iterrows(), start=2):
        es_fav = bool(fila_df["Favorable"])
        for c_idx, nombre_col in enumerate(cols_export, start=1):
            c = ws_alt.cell(row=r_idx, column=c_idx, value=fila_df[nombre_col])
            c.border = borde_fino
            if c_idx == len(cols_export):
                c.fill = fill_ok if es_fav else fill_err

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


# BARRA LATERAL (SIDEBAR): CARGA DE ARCHIVOS + 8 INPUTS MANUALES

with st.sidebar:
    st.markdown("## ⚙️ Panel de Control e Inputs")
    st.caption("Configura los parámetros geotécnicos, geométricos y de armado o carga un archivo externo.")

    # 1. CARGA DESDE ARCHIVO EXTERNO (.CSV, .XLSX, .TXT)
    with st.expander("📂 1. Cargar Datos desde Archivo", expanded=True):
        archivo_subido = st.file_uploader(
            "Subir archivo (.csv, .xlsx, .txt)",
            type=["csv", "xlsx", "xls", "txt"],
            help="El archivo debe contener las 4 filas de cargas (CM, CV, Sx, Sy) con columnas P, Mx, My."
        )
        if archivo_subido is not None:
            if st.button("📥 Procesar Archivo Subido", use_container_width=True, type="primary"):
                try:
                    sufijo = os.path.splitext(archivo_subido.name)[1]
                    with tempfile.NamedTemporaryFile(delete=False, suffix=sufijo) as tmp:
                        tmp.write(archivo_subido.getbuffer())
                        ruta_tmp = tmp.name

                    info = cargar_datos_archivo(ruta_tmp)
                    df_cargado = info["df_cargas"].copy()
                    df_cargado.index = ["Carga Muerta (CM)", "Carga Viva (CV)", "Sismo en X (Sx)", "Sismo en Y (Sy)"]
                    st.session_state["df_cargas"] = df_cargado

                    if info["q_adm"] is not None:
                        st.session_state["q_adm"] = float(info["q_adm"])
                    if info["L_C"] is not None:
                        st.session_state["l_c"] = float(info["L_C"])
                    if info["B_C"] is not None:
                        st.session_state["b_c"] = float(info["B_C"])

                    st.session_state["origen_datos_txt"] = f"Archivo cargado: {archivo_subido.name}"
                    os.remove(ruta_tmp)
                    st.success(f"¡Datos cargados desde {archivo_subido.name}!")
                except Exception as e:
                    st.error(f"Error al leer el archivo: {e}")

    # 2. PARÁMETROS GEOTÉCNICOS Y MATERIALES (INPUTS MANUALES 1, 2 y 3)
    with st.expander("🧱 2. Suelo y Materiales (3 Inputs)", expanded=True):
        q_adm = st.number_input(
            "Capacidad Portante q_adm (kg/cm²)",
            min_value=0.50, max_value=10.00, step=0.10, format="%.2f",
            key="q_adm"
        )
        fc = st.number_input(
            "Resistencia Concreto f'c (kg/cm²)",
            min_value=175.0, max_value=420.0, step=5.0, format="%.0f",
            key="fc"
        )
        fy = st.number_input(
            "Fluencia del Acero fy (kg/cm²)",
            min_value=2800.0, max_value=5000.0, step=100.0, format="%.0f",
            key="fy"
        )

    # 3. GEOMETRÍA INICIAL Y COLUMNA (INPUTS MANUALES 4, 5 y 6)
    with st.expander("📐 3. Geometría de Zapata y Columna (3 Inputs)", expanded=True):
        h_inicial = st.number_input(
            "Peralte Inicial Tentativo H (m)",
            min_value=0.30, max_value=2.50, step=0.05, format="%.2f",
            key="h_inicial"
        )
        col_c1, col_c2 = st.columns(2)
        with col_c1:
            l_c = st.number_input(
                "Columna Lc (m)",
                min_value=0.20, max_value=2.00, step=0.05, format="%.2f",
                key="l_c"
            )
        with col_c2:
            b_c = st.number_input(
                "Columna Bc (m)",
                min_value=0.20, max_value=2.00, step=0.05, format="%.2f",
                key="b_c"
            )

    # 4. ARMADURA DE FLEXIÓN (INPUTS MANUALES 7 y 8)
    with st.expander("🔩 4. Diseño de Acero Inferior (2 Inputs)", expanded=True):
        lista_varillas = list(TABLA_VARILLAS.keys())
        varilla = st.selectbox(
            "Diámetro de Varilla (Ø)",
            options=lista_varillas,
            key="varilla"
        )
        espaciamiento = st.slider(
            "Espaciamiento s (m)",
            min_value=0.100, max_value=0.300, step=0.025, format="%.3f",
            key="espaciamiento"
        )

    st.divider()
    col_btn1, col_btn2 = st.columns(2)
    with col_btn1:
        btn_calcular = st.button("⚡ Recalcular", type="primary", use_container_width=True)
    with col_btn2:
        st.button("🔄 Restaurar", on_click=restaurar_valores, use_container_width=True)


# ENCABEZADO PRINCIPAL DE LA PÁGINA WEB

st.markdown("""
<div class="main-header">
    <h1>🏗️ STRUCTZAP WEB PRO v2.0 — Diseño y Verificación de Zapatas Aisladas</h1>
    <p>Automatización Geotécnica y Estructural según <b>NTP E.050 (Suelos y Cimentaciones)</b> y <b>NTP E.060 (Concreto Armado)</b></p>
</div>
""", unsafe_allow_html=True)


# SECCIÓN DESPLEGABLE: TABLA EDITABLE DE CARGAS ACTUANTES (MANUAL O ARCHIVO)

with st.expander("📊 Matriz de Cargas Actuantes en la Base de la Columna (Editable Manualmente)", expanded=False):
    st.markdown(f"**📄 Origen actual de datos:** `{st.session_state['origen_datos_txt']}`")
    df_editado = st.data_editor(
        st.session_state["df_cargas"],
        column_config={
            "P": st.column_config.NumberColumn("Carga Axial P (Tn)", format="%.4f"),
            "Mx": st.column_config.NumberColumn("Momento Mx (Tn·m)", format="%.4f"),
            "My": st.column_config.NumberColumn("Momento My (Tn·m)", format="%.4f"),
        },
        use_container_width=True,
        num_rows="fixed"
    )
    st.session_state["df_cargas"] = df_editado


# EJECUCIÓN DEL MOTOR DE CÁLCULO (LLAMADA A codigo_base.py)

df_para_calculo = st.session_state["df_cargas"].reset_index(drop=True)

resultados = calcular_zapata(
    origen_datos=df_para_calculo,
    q_adm=st.session_state["q_adm"],
    fc=st.session_state["fc"],
    H_inicial=st.session_state["h_inicial"],
    L_C=st.session_state["l_c"],
    B_C=st.session_state["b_c"],
    varilla_elegida=st.session_state["varilla"],
    espaciamiento_m=st.session_state["espaciamiento"],
    fy=st.session_state["fy"]
)

arm = resultados["armadura"]


# FILA DE 5 TARJETAS EJECUTIVAS (KPIs DE INGENIERÍA)

k1, k2, k3, k4, k5 = st.columns(5)

with k1:
    st.markdown(f"""
    <div class="kpi-card">
        <div class="kpi-title">1. Geometría Final Zapata</div>
        <div class="kpi-value" style="color:#38bdf8;">{resultados['Base_final']:.2f} × {resultados['Largo_final']:.2f} × {resultados['Peralte_H']:.2f} m</div>
        <div class="kpi-sub" style="color:#94a3b8;">Área = {resultados['Area_final']:.2f} m² | d = {resultados['d_efectivo']:.2f} m</div>
    </div>
    """, unsafe_allow_html=True)

with k2:
    ok_suelo = resultados["cumple_suelo_CS"] and resultados["cumple_suelo_Sismo"]
    badge_s = '<span class="badge-ok">✓ CUMPLE</span>' if ok_suelo else '<span class="badge-err">✗ REVISAR</span>'
    st.markdown(f"""
    <div class="kpi-card">
        <div class="kpi-title">2. Presión en Suelo (CS)</div>
        <div class="kpi-value">{resultados['esfuerzo_maximo_CS']:.2f} / {resultados['Qatn']:.1f} Tn/m²</div>
        <div class="kpi-sub">{badge_s} (σ mín = {resultados['esfuerzo_minimo_CS']:.2f})</div>
    </div>
    """, unsafe_allow_html=True)

with k3:
    vu_max = max(resultados["V_u1"], resultados["V_u2"])
    ok_corte = resultados["cumple_corte_1"] and resultados["cumple_corte_2"]
    badge_c = '<span class="badge-ok">✓ CUMPLE</span>' if ok_corte else '<span class="badge-err">✗ FALLA</span>'
    st.markdown(f"""
    <div class="kpi-card">
        <div class="kpi-title">3. Corte Unidireccional</div>
        <div class="kpi-value">Vu = {vu_max:.2f} ≤ {resultados['φVc1']:.2f} Tn</div>
        <div class="kpi-sub">{badge_c} (Ratio: {vu_max/resultados['φVc1']:.2f})</div>
    </div>
    """, unsafe_allow_html=True)

with k4:
    ok_punz = resultados["cumple_punzonamiento"]
    badge_p = '<span class="badge-ok">✓ CUMPLE</span>' if ok_punz else '<span class="badge-err">✗ FALLA</span>'
    st.markdown(f"""
    <div class="kpi-card">
        <div class="kpi-title">4. Punzonamiento (b₀={resultados['Pm']:.2f}m)</div>
        <div class="kpi-value">{resultados['Vu_Punzonamiento']:.1f} ≤ {resultados['φVc_Punzonamiento']:.1f} Tn</div>
        <div class="kpi-sub">{badge_p} (Ratio: {resultados['Vu_Punzonamiento']/resultados['φVc_Punzonamiento']:.2f})</div>
    </div>
    """, unsafe_allow_html=True)

with k5:
    ok_acero = arm["es_favorable"]
    badge_a = '<span class="badge-ok">✓ FAVORABLE</span>' if ok_acero else '<span class="badge-err">✗ DESFAVORABLE</span>'
    st.markdown(f"""
    <div class="kpi-card">
        <div class="kpi-title">5. Armadura de Flexión</div>
        <div class="kpi-value" style="color:#34d399;">{arm['descripcion']}</div>
        <div class="kpi-sub">{badge_a} (φMn={arm['φMn_Tnm']:.2f} ≥ Mu={arm['MU_Tnm']:.2f})</div>
    </div>
    """, unsafe_allow_html=True)

st.write("")

# Barra de progreso de utilización a flexión + Botón de Exportación directa
col_prog, col_exp = st.columns([3.2, 1.2])
with col_prog:
    pct_uso = min(1.0, float(arm["ratio_demanda_capacidad"]))
    st.progress(
        pct_uso,
        text=f"🔧 Ratio Demanda/Capacidad a Flexión (Mu / φMn): {arm['ratio_demanda_capacidad']:.2f} ({arm['porcentaje_uso']:.1f}%) — As colocado = {arm['As_colocado']:.2f} cm²/m (As mín = {arm['As_minimo']:.2f} cm²/m)"
    )
with col_exp:
    excel_bytes = generar_excel_bytes(resultados)
    st.download_button(
        label="📥 Exportar Reporte Completo (.XLSX)",
        data=excel_bytes,
        file_name=f"Reporte_Zapata_{resultados['Base_final']:.2f}x{resultados['Largo_final']:.2f}m.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True
    )


# ORGANIZACIÓN EN PESTAÑAS (TABS): 4 GRÁFICOS, MATRIZ DE ACERO Y MEMORIA TÉCNICA

tab1, tab2, tab3, tab4 = st.tabs([
    "📊 Panel de 4 Gráficos de Ingeniería",
    "🔩 Matriz Comercial de Alternativas de Acero",
    "📋 Memoria de Cálculo Detallada",
    "🖼️ Vista Individual Ampliada de Gráficos"
])

# --- PESTAÑA 1: CUADRÍCULA 2x2 CON LOS 4 GRÁFICOS DE CÁLCULOS Y RESULTADOS ---
with tab1:
    st.markdown("#### Visualización Técnica en Tiempo Real (Geometría, 3D, Suelo y Estructuras)")
    g_col1, g_col2 = st.columns(2)

    with g_col1:
        st.markdown("**1. Planta Estructural 2D y Secciones Críticas**")
        if os.path.exists(resultados["ruta_img_2d"]):
            st.image(resultados["ruta_img_2d"], use_container_width=True)

        st.markdown("**3. Verificación de Presiones Actuantes vs. Admisibles en el Suelo**")
        if os.path.exists(resultados["ruta_img_esfuerzos"]):
            st.image(resultados["ruta_img_esfuerzos"], use_container_width=True)

    with g_col2:
        st.markdown("**2. Modelo Volumétrico 3D y Parrilla de Acero**")
        if os.path.exists(resultados["ruta_img_3d"]):
            st.image(resultados["ruta_img_3d"], use_container_width=True)

        st.markdown("**4. Verificación de Corte, Punzonamiento y Capacidad a Flexión**")
        if os.path.exists(resultados["ruta_img_estructural"]):
            st.image(resultados["ruta_img_estructural"], use_container_width=True)

# --- PESTAÑA 2: TABLA INTERACTIVA DE ALTERNATIVAS COMERCIALES DE ACERO ---
with tab2:
    st.markdown("#### Explorador de Combinaciones Comerciales de Acero (ASTM A615 / NTP E.060)")
    col_f1, col_f2 = st.columns([1, 2])
    with col_f1:
        solo_fav = st.checkbox("✅ Mostrar únicamente alternativas FAVORABLES", value=True)
    with col_f2:
        st.info(f"**Diseño Actual Seleccionado:** `{arm['descripcion']}` → **As = {arm['As_colocado']:.2f} cm²/m** | **φMn = {arm['φMn_Tnm']:.2f} Tn·m** vs **Mu = {arm['MU_Tnm']:.2f} Tn·m**")

    df_alt_mostrar = resultados["tabla_alternativas"].copy()
    if solo_fav:
        df_alt_mostrar = df_alt_mostrar[df_alt_mostrar["Favorable"] == True].reset_index(drop=True)

    df_alt_vista = df_alt_mostrar[[
        "Varilla", "Area_barra_cm2", "Espaciamiento_m",
        "As_cm2_m", "As_min_cm2_m", "φMn_Tnm", "Mu_Tnm", "Ratio", "Estado"
    ]].rename(columns={
        "Varilla": "Diámetro (Ø)",
        "Area_barra_cm2": "Área Barra (cm²)",
        "Espaciamiento_m": "Espaciamiento s (m)",
        "As_cm2_m": "As Colocado (cm²/m)",
        "As_min_cm2_m": "As Mínimo (cm²/m)",
        "φMn_Tnm": "Capacidad φMn (Tn·m)",
        "Mu_Tnm": "Demanda Mu (Tn·m)",
        "Ratio": "Ratio (Mu/φMn)",
        "Estado": "Verificación Normativa"
    })

    st.dataframe(df_alt_vista, use_container_width=True, hide_index=True)

# --- PESTAÑA 3: MEMORIA DE CÁLCULO DETALLADA ---
with tab3:
    st.markdown("### 📋 Memoria de Cálculo Estructural Paso a Paso")
    m_col1, m_col2 = st.columns(2)

    with m_col1:
        st.markdown("##### 1. Dimensionamiento Geométrico y Propiedades")
        df_geo = pd.DataFrame([
            ["Base Final de Zapata (B)", f"{resultados['Base_final']:.2f} m", "Mínimo normativo ≥ 1.50 m"],
            ["Largo Final de Zapata (L)", f"{resultados['Largo_final']:.2f} m", f"Área Final A = {resultados['Area_final']:.2f} m²"],
            ["Peralte Total (H) / Efectivo (d)", f"{resultados['Peralte_H']:.2f} m / {resultados['d_efectivo']:.2f} m", "Recubrimiento libre + barra = 0.10 m"],
            ["Dimensiones de Columna (Lc × Bc)", f"{resultados['L_C']:.2f} m × {resultados['B_C']:.2f} m", f"Ix = {resultados['Ix']:.3f} m⁴ | Iy = {resultados['Iy']:.3f} m⁴"],
            ["Volados Libres (Dy / Dx)", f"{resultados['Volado_1']:.2f} m / {resultados['Volado_2']:.2f} m", f"Factor peso propio = {resultados['factor_peso']:.2f}"],
        ], columns=["Parámetro", "Valor Calculado", "Observación Técnica"])
        st.table(df_geo)

        st.markdown("##### 2. Verificación Geotécnica de Presiones en el Terreno")
        df_suelo = pd.DataFrame([
            ["Carga de Servicio (CM + CV)", f"{resultados['esfuerzo_maximo_CS']:.3f}", f"{resultados['esfuerzo_minimo_CS']:.3f}", f"{resultados['Qatn']:.2f}", "✓ CUMPLE" if resultados['cumple_suelo_CS'] else "✗ REVISAR"],
            ["Sismo X (CS + 0.80 Sx)", f"{resultados['esfuerzo_maximo_Sx']:.3f}", f"{resultados['esfuerzo_minimo_Sx']:.3f}", f"{resultados['Qatn_Sx']:.2f}", "✓ CUMPLE" if (resultados['esfuerzo_maximo_Sx'] <= resultados['Qatn_Sx'] and resultados['esfuerzo_minimo_Sx'] > 0) else "✗ REVISAR"],
            ["Sismo Y (CS + 0.80 Sy)", f"{resultados['esfuerzo_maximo_Sy']:.3f}", f"{resultados['esfuerzo_minimo_Sy']:.3f}", f"{resultados['Qatn_Sx']:.2f}", "✓ CUMPLE" if (resultados['esfuerzo_maximo_Sy'] <= resultados['Qatn_Sx'] and resultados['esfuerzo_minimo_Sy'] > 0) else "✗ REVISAR"],
        ], columns=["Combinación", "σ máx (Tn/m²)", "σ mín (Tn/m²)", "q adm (Tn/m²)", "Estado"])
        st.table(df_suelo)

    with m_col2:
        st.markdown(f"##### 3. Verificación por Corte y Punzonamiento (Wu = {resultados['Wu']:.2f} Tn/m²)")
        df_corte = pd.DataFrame([
            ["Corte Unidireccional Volado 1 (Y)", f"{resultados['V_u1']:.2f} Tn", f"{resultados['φVc1']:.2f} Tn", f"{resultados['V_u1']/resultados['φVc1']:.2f}", "✓ CUMPLE" if resultados['cumple_corte_1'] else "✗ NO CUMPLE"],
            ["Corte Unidireccional Volado 2 (X)", f"{resultados['V_u2']:.2f} Tn", f"{resultados['φVc2']:.2f} Tn", f"{resultados['V_u2']/resultados['φVc2']:.2f}", "✓ CUMPLE" if resultados['cumple_corte_2'] else "✗ NO CUMPLE"],
            [f"Punzonamiento Bidireccional (b₀={resultados['Pm']:.2f}m)", f"{resultados['Vu_Punzonamiento']:.2f} Tn", f"{resultados['φVc_Punzonamiento']:.2f} Tn", f"{resultados['Vu_Punzonamiento']/resultados['φVc_Punzonamiento']:.2f}", "✓ CUMPLE" if resultados['cumple_punzonamiento'] else "✗ NO CUMPLE"],
        ], columns=["Sección Crítica", "Demanda Vu", "Capacidad φVc", "Ratio", "Estado"])
        st.table(df_corte)

        st.markdown("##### 4. Diseño de Refuerzo Longitudinal por Flexión")
        df_flex = pd.DataFrame([
            ["Armadura Seleccionada", arm["descripcion"]],
            ["Área de Acero Colocado (As)", f"{arm['As_colocado']:.2f} cm²/m (Mínimo = {arm['As_minimo']:.2f} cm²/m)"],
            ["Momento Último Actuante (Mu)", f"{arm['MU_Tnm']:.2f} Tn·m"],
            ["Momento Resistente de Diseño (φMn)", f"{arm['φMn_Tnm']:.2f} Tn·m (Bloque a = {arm['a_cm']:.2f} cm)"],
            ["Estado Normativo NTP E.060", arm["estado"]],
        ], columns=["Parámetro de Diseño", "Resultado"])
        st.table(df_flex)

# --- PESTAÑA 4: VISOR AMPLIADO DE CADA GRÁFICO ---
with tab4:
    opcion_grafico = st.radio(
        "Seleccione el gráfico para inspeccionar en alta resolución:",
        [
            "1. Planta Estructural 2D y Secciones Críticas",
            "2. Modelo Volumétrico 3D y Armadura",
            "3. Verificación de Presiones en el Suelo",
            "4. Verificación de Corte, Punzonamiento y Flexión"
        ],
        horizontal=True
    )
    mapa_rutas = {
        "1. Planta Estructural 2D y Secciones Críticas": resultados["ruta_img_2d"],
        "2. Modelo Volumétrico 3D y Armadura": resultados["ruta_img_3d"],
        "3. Verificación de Presiones en el Suelo": resultados["ruta_img_esfuerzos"],
        "4. Verificación de Corte, Punzonamiento y Flexión": resultados["ruta_img_estructural"],
    }
    ruta_sel = mapa_rutas[opcion_grafico]
    if os.path.exists(ruta_sel):
        st.image(ruta_sel, use_container_width=True)