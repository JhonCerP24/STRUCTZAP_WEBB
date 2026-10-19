import os
import pandas as pd
import numpy as np


import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.patches import Rectangle
from mpl_toolkits.mplot3d.art3d import Poly3DCollection


# 1. TABLA Y FUNCIONES DE ACERO MODULARES 


# Tabla de diámetros nominales y áreas por varilla (ASTM A615 / NTP E.060)
TABLA_VARILLAS = {
    '3/8"': 0.71,   # cm² (Ø 9.5 mm)
    '1/2"': 1.29,   # cm² (Ø 12.7 mm)
    '5/8"': 2.00,   # cm² (Ø 15.9 mm)
    '3/4"': 2.84,   # cm² (Ø 19.1 mm)
    '1"':   5.10    # cm² (Ø 25.4 mm)
}


def cargar_datos_archivo(ruta_archivo):
    """
    Lee un archivo externo (.csv, .txt o .xlsx) con las cargas actuantes (P, Mx, My)
    para CM, CV, Sismo X y Sismo Y, además de parámetros opcionales (q_adm, L_C, B_C).
    """
    if not os.path.exists(ruta_archivo):
        raise FileNotFoundError(f"No se encontró el archivo: {ruta_archivo}")

    ext = os.path.splitext(ruta_archivo)[1].lower()
    if ext in ['.xlsx', '.xls']:
        df_raw = pd.read_excel(ruta_archivo)
    else:
        try:
            df_raw = pd.read_csv(ruta_archivo, sep=r'\s+', engine='python')
            if df_raw.shape[1] < 3:
                df_raw = pd.read_csv(ruta_archivo, sep=None, engine='python')
        except Exception:
            df_raw = pd.read_csv(ruta_archivo, sep=None, engine='python')

    # Limpiar filas completamente vacías
    df_clean = df_raw.dropna(how='all').reset_index(drop=True)

    # Extraer las primeras 4 filas numéricas correspondientes a CM, CV, Sx, Sy
    cargas_filas = []
    for idx in range(min(4, len(df_clean))):
        fila_vals = pd.to_numeric(df_clean.iloc[idx, :3], errors='coerce').fillna(0.0).tolist()
        while len(fila_vals) < 3:
            fila_vals.append(0.0)
        cargas_filas.append(fila_vals[:3])

    while len(cargas_filas) < 4:
        cargas_filas.append([0.0, 0.0, 0.0])

    df_cargas = pd.DataFrame(cargas_filas, columns=['P', 'Mx', 'My'])

    # Extraer parámetros adicionales si el archivo los incluye en las filas 4 y 5
    q_adm_extra = None
    l_c_extra = None
    b_c_extra = None

    if len(df_clean) > 4:
        val_q = pd.to_numeric(df_clean.iloc[4, 1], errors='coerce')
        if pd.notna(val_q) and float(val_q) > 0:
            q_adm_extra = float(val_q)

    if len(df_clean) > 5:
        val_lc = pd.to_numeric(df_clean.iloc[5, 0], errors='coerce')
        val_bc = pd.to_numeric(df_clean.iloc[5, 1], errors='coerce')
        if pd.notna(val_lc) and float(val_lc) > 0:
            l_c_extra = float(val_lc)
        if pd.notna(val_bc) and float(val_bc) > 0:
            b_c_extra = float(val_bc)

    return {
        "df_cargas": df_cargas,
        "q_adm": q_adm_extra,
        "L_C": l_c_extra,
        "B_C": b_c_extra,
        "ruta_origen": ruta_archivo
    }


def verificar_armadura(
    MU, H, d, As_min, fc=210, fy=4200,
    varilla_elegida='5/8"', espaciamiento_m=0.15,
    As_propuesto=None, b_cm=100
):
   
    d_cm = round(float(d) * 100, 2)
    MU = float(MU)
    As_min = float(As_min)
    fc = float(fc)
    fy = float(fy)

    if As_propuesto is not None and As_propuesto > 0:
        As_colocado = round(float(As_propuesto), 2)
        descripcion = f"As = {As_colocado:.2f} cm²/m"
    else:
        area_var = TABLA_VARILLAS.get(varilla_elegida, 2.00)
        As_colocado = round(area_var / espaciamiento_m, 2) if espaciamiento_m > 0 else 0.0
        descripcion = f'Ø {varilla_elegida} @ {espaciamiento_m:.3f} m'

    # Área de acero requerida por flexión (bloque de Whitney)
    phi = 0.90
    termino_raiz = 1.0 - (2.0 * (MU * 1e5)) / (phi * 0.85 * fc * b_cm * (d_cm ** 2))
    if termino_raiz > 0:
        cuantia_req = (0.85 * fc / fy) * (1.0 - np.sqrt(termino_raiz))
        As_req_flexion = round(cuantia_req * b_cm * d_cm, 2)
    else:
        As_req_flexion = round(As_min * 1.5, 2)

    As_gobernante = round(max(As_min, As_req_flexion), 2)

    # Profundidad del bloque de compresión de Whitney (a)
    a = round((As_colocado * fy) / (0.85 * fc * b_cm), 2)

    # Momento nominal resistente de diseño φMn (Tn.m), con φ = 0.90 para flexión
    phi_Mn = round(phi * As_colocado * fy * (d_cm - a / 2.0) / 10**5, 2)

    # Verificaciones normativas (NTP E.060)
    cumple_resistencia = bool(phi_Mn >= MU)
    cumple_as_min = bool(As_colocado >= As_min)
    es_favorable = bool(cumple_resistencia and cumple_as_min)
    ratio = round(MU / phi_Mn, 2) if phi_Mn > 0 else 999.0

    if es_favorable:
        estado = "✓ FAVORABLE (Mu ≤ φMn y As ≥ As mín)"
        color_gui = "#10b981"
    elif not cumple_resistencia:
        estado = "✗ DESFAVORABLE (Demanda Mu > φMn)"
        color_gui = "#ef4444"
    else:
        estado = "✗ DESFAVORABLE (As < As mínimo)"
        color_gui = "#f59e0b"

    return {
        "descripcion": descripcion,
        "varilla": varilla_elegida,
        "area_varilla_cm2": TABLA_VARILLAS.get(varilla_elegida, 2.00),
        "espaciamiento_m": round(float(espaciamiento_m), 3),
        "As_colocado": As_colocado,
        "As_minimo": round(As_min, 2),
        "As_req_flexion": As_req_flexion,
        "As_gobernante": As_gobernante,
        "a_cm": a,
        "d_cm": d_cm,
        "MU_Tnm": round(MU, 2),
        "Momento_Ultimo": round(MU, 2),
        "φMn_Tnm": phi_Mn,
        "Phi_Mn_Tnm": phi_Mn,
        "φ_Mn_Tnm": phi_Mn,
        "ratio_demanda_capacidad": ratio,
        "porcentaje_uso": round(min(ratio * 100.0, 999.0), 1),
        "cumple_resistencia": cumple_resistencia,
        "cumple_as_min": cumple_as_min,
        "es_favorable": es_favorable,
        "estado": estado,
        "color_gui": color_gui
    }


def generar_tabla_alternativas(MU, H, d, As_min, fc=210, fy=4200, b_cm=100):
   

    espaciamientos = [0.10, 0.125, 0.15, 0.175, 0.20, 0.25, 0.30]
    filas = []

    for var, area_var in TABLA_VARILLAS.items():
        for s in espaciamientos:
            res = verificar_armadura(
                MU, H, d, As_min,
                fc=fc, fy=fy,
                varilla_elegida=var,
                espaciamiento_m=s,
                b_cm=b_cm
            )
            filas.append({
                "Varilla": var,
                "Area_barra_cm2": area_var,
                "Espaciamiento_m": s,
                "As_cm2_m": res["As_colocado"],
                "As_min_cm2_m": res["As_minimo"],
                "φMn_Tnm": res["φMn_Tnm"],
                "Phi_Mn_Tnm": res["φMn_Tnm"],
                "Mu_Tnm": res["MU_Tnm"],
                "Ratio": res["ratio_demanda_capacidad"],
                "Estado": res["estado"],
                "Favorable": res["es_favorable"]
            })

    return pd.DataFrame(filas)



# 2. GENERADORES MODULARES DE LOS 4 GRÁFICOS DE INGENIERÍA 


def _aplicar_estilo_cad_oscuro(ax, fig, titulo=""):
    
    fig.patch.set_facecolor('#0b1325')
    ax.set_facecolor('#0f1b33')
    if titulo:
        ax.set_title(titulo, fontsize=11, fontweight='bold', color='#f8fafc', pad=12)
    ax.tick_params(colors='#cbd5e1', labelsize=8.5)
    for spine in ax.spines.values():
        spine.set_color('#1e3a5f')
        spine.set_linewidth(1.1)


def generar_grafico_planta_2d(Nv_Base, Largo, L_C, B_C, d, varilla, espaciamiento, ruta_img_2d):
    
    fig, ax = plt.subplots(figsize=(6.8, 6.2))
    _aplicar_estilo_cad_oscuro(
        ax, fig,
        f"1. PLANTA ESTRUCTURAL 2D Y SECCIONES CRÍTICAS\nZapata: {Nv_Base:.2f}×{Largo:.2f} m | Columna: {L_C:.2f}×{B_C:.2f} m"
    )

    D_x = round((Nv_Base - L_C) / 2.0, 2)
    D_y = round((Largo - B_C) / 2.0, 2)

    # Sombra sutil de la zapata
    sombra = Rectangle((0.04, -0.04), Nv_Base, Largo, facecolor='#050914', edgecolor='none', zorder=1)
    ax.add_patch(sombra)

    # Zapata principal
    zapata = Rectangle(
        (0, 0), Nv_Base, Largo,
        facecolor='#162a4a', edgecolor='#38bdf8', linewidth=2.2, zorder=2
    )
    ax.add_patch(zapata)

    # Parrilla esquemática de acero inferior (líneas tenues)
    paso_visual = max(0.18, float(espaciamiento))
    for x_bar in np.arange(0.12, Nv_Base - 0.08, paso_visual):
        ax.plot([x_bar, x_bar], [0.08, Largo - 0.08], color='#334155', lw=0.7, ls='-', alpha=0.55, zorder=2.5)
    for y_bar in np.arange(0.12, Largo - 0.08, paso_visual):
        ax.plot([0.08, Nv_Base - 0.08], [y_bar, y_bar], color='#334155', lw=0.7, ls='-', alpha=0.55, zorder=2.5)

    # Perímetro crítico de punzonamiento a d/2 alrededor de la columna
    px = max(0.02, D_x - d / 2.0)
    py = max(0.02, D_y - d / 2.0)
    pw = min(Nv_Base - 2 * px, L_C + d)
    ph = min(Largo - 2 * py, B_C + d)
    perimetro_punz = Rectangle(
        (px, py), pw, ph,
        facecolor='#f59e0b', alpha=0.14, edgecolor='#f59e0b',
        linewidth=1.6, linestyle='--', zorder=3
    )
    ax.add_patch(perimetro_punz)
    ax.text(
        px + pw / 2.0, py + ph + 0.04,
        f"Perímetro Punz. b₀ (d/2 = {d/2:.2f}m)",
        ha='center', va='bottom', color='#fbbf24', fontsize=8.0, fontweight='bold', zorder=5
    )

    # Sección crítica de corte unidireccional a distancia d (derecho)
    x_corte_d = D_x + L_C + d
    if x_corte_d < Nv_Base:
        ax.plot(
            [x_corte_d, x_corte_d], [0, Largo],
            color='#f43f5e', linestyle='-.', linewidth=1.5, zorder=3.5
        )
        ax.text(
            x_corte_d + 0.03, 0.12, "Corte Unidireccional (a dist. d)",
            rotation=90, ha='left', va='bottom', color='#fda4af', fontsize=7.8, fontweight='bold', zorder=5
        )

    # Columna central
    columna = Rectangle(
        (D_x, D_y), L_C, B_C,
        facecolor='#0d9488', edgecolor='#5eead4', linewidth=2.0, zorder=4
    )
    ax.add_patch(columna)
    ax.text(
        D_x + L_C / 2.0, D_y + B_C / 2.0,
        f"COLUMNA\n{L_C:.2f}×{B_C:.2f}m",
        ha='center', va='center', color='#ffffff', fontweight='bold', fontsize=7.8, zorder=6
    )

    # 1. COTAS EXTERIORES DE LA ZAPATA (B y L)
    offset_zap = max(0.20, Nv_Base * 0.11)
    ax.annotate(
        '', xy=(0, -offset_zap), xytext=(Nv_Base, -offset_zap),
        arrowprops=dict(arrowstyle='<->', color='#38bdf8', lw=1.6)
    )
    ax.text(
        Nv_Base / 2.0, -offset_zap - 0.06, f"Base B = {Nv_Base:.2f} m",
        ha='center', va='top', fontweight='bold', color='#38bdf8', fontsize=9.5
    )

    ax.annotate(
        '', xy=(-offset_zap, 0), xytext=(-offset_zap, Largo),
        arrowprops=dict(arrowstyle='<->', color='#38bdf8', lw=1.6)
    )
    ax.text(
        -offset_zap - 0.06, Largo / 2.0, f"Largo L = {Largo:.2f} m",
        ha='right', va='center', fontweight='bold', color='#38bdf8', fontsize=9.5, rotation=90
    )

    # 2. COTAS DE VOLADOS (Dx y Dy)
    y_dx = D_y + B_C / 2.0
    ax.annotate(
        '', xy=(0, y_dx), xytext=(D_x, y_dx),
        arrowprops=dict(arrowstyle='<->', color='#34d399', lw=1.4)
    )
    ax.text(
        D_x / 2.0, y_dx + 0.04, f"Volado X\nDx = {D_x:.2f} m",
        ha='center', va='bottom', fontweight='bold', color='#34d399', fontsize=8.2
    )

    x_dy = D_x + L_C / 2.0
    ax.annotate(
        '', xy=(x_dy, 0), xytext=(x_dy, D_y),
        arrowprops=dict(arrowstyle='<->', color='#34d399', lw=1.4)
    )
    ax.text(
        x_dy - 0.04, D_y / 2.0, f"Volado Y\nDy = {D_y:.2f} m",
        ha='right', va='center', fontweight='bold', color='#34d399', fontsize=8.2
    )

    # Etiqueta de armadura en la parte superior
    ax.text(
        Nv_Base / 2.0, Largo + 0.10,
        f"Malla Inferior: Ø {varilla} @ {espaciamiento:.3f} m (en ambas direcciones)",
        ha='center', va='bottom', color='#a78bfa', fontweight='bold', fontsize=8.8
    )

    plt.axis('off')
    ax.set_aspect('equal')
    margen_x = max(0.52, Nv_Base * 0.25)
    margen_y = max(0.45, Largo * 0.22)
    plt.xlim(-margen_x, Nv_Base + margen_x * 0.75)
    plt.ylim(-margen_y, Largo + margen_y * 0.75)

    fig.savefig(ruta_img_2d, dpi=140, bbox_inches='tight', facecolor=fig.get_facecolor())
    plt.close(fig)


def _crear_prisma_3d(x0, y0, z0, dx, dy, dz):
   
    v = np.array([
        [x0, y0, z0],
        [x0 + dx, y0, z0],
        [x0 + dx, y0 + dy, z0],
        [x0, y0 + dy, z0],
        [x0, y0, z0 + dz],
        [x0 + dx, y0, z0 + dz],
        [x0 + dx, y0 + dy, z0 + dz],
        [x0, y0 + dy, z0 + dz]
    ])
    return [
        [v[0], v[1], v[2], v[3]],
        [v[4], v[5], v[6], v[7]],
        [v[0], v[1], v[5], v[4]],
        [v[2], v[3], v[7], v[6]],
        [v[1], v[2], v[6], v[5]],
        [v[4], v[7], v[3], v[0]]
    ]


def generar_grafico_modelo_3d(Nv_Base, Largo, H, L_C, B_C, varilla, espaciamiento, ruta_img_3d):
    
    fig = plt.figure(figsize=(6.8, 6.2))
    fig.patch.set_facecolor('#0b1325')
    ax = fig.add_subplot(111, projection='3d')
    ax.set_facecolor('#0b1325')

    # Estética de paneles 3D oscuros
    for pane in [ax.xaxis.pane, ax.yaxis.pane, ax.zaxis.pane]:
        pane.set_facecolor((0.06, 0.11, 0.20, 0.9))
        pane.set_edgecolor('#1e3a5f')

    ax.tick_params(colors='#94a3b8', labelsize=8)
    ax.xaxis.label.set_color('#38bdf8')
    ax.yaxis.label.set_color('#38bdf8')
    ax.zaxis.label.set_color('#f59e0b')

    # 1. Bloque de la Zapata (translúcido para ver la parrilla de acero interna)
    caras_zapata = _crear_prisma_3d(0, 0, 0, Nv_Base, Largo, H)
    poly_zapata = Poly3DCollection(
        caras_zapata, facecolors='#1e3a8a', edgecolors='#38bdf8',
        linewidths=1.4, alpha=0.38
    )
    ax.add_collection3d(poly_zapata)

    # 2. Parrilla de Acero Inferior (z = recubrimiento 0.075 m)
    z_acero = min(0.08, H * 0.15)
    paso = max(0.15, float(espaciamiento))
    xs = np.arange(0.08, Nv_Base - 0.05, paso)
    ys = np.arange(0.08, Largo - 0.05, paso)

    for xb in xs:
        ax.plot([xb, xb], [0.07, Largo - 0.07], [z_acero, z_acero], color='#10b981', lw=1.5, alpha=0.9)
    for yb in ys:
        ax.plot([0.07, Nv_Base - 0.07], [yb, yb], [z_acero + 0.015, z_acero + 0.015], color='#34d399', lw=1.5, alpha=0.9)

    # 3. Pedestal / Columna Central sobre la Zapata
    D_x = (Nv_Base - L_C) / 2.0
    D_y = (Largo - B_C) / 2.0
    h_col = max(0.55, H * 0.95)
    caras_col = _crear_prisma_3d(D_x, D_y, H, L_C, B_C, h_col)
    poly_col = Poly3DCollection(
        caras_col, facecolors='#0d9488', edgecolors='#5eead4',
        linewidths=1.5, alpha=0.88
    )
    ax.add_collection3d(poly_col)

    # 4. Aceros longitudinales de columna (4 esquinas)
    for cx in [D_x + 0.04, D_x + L_C - 0.04]:
        for cy in [D_y + 0.04, D_y + B_C - 0.04]:
            ax.plot([cx, cx], [cy, cy], [z_acero, H + h_col + 0.12], color='#fbbf24', lw=1.8, alpha=0.95)

    # Etiquetas 3D limpias sin colisión
    ax.text(
        Nv_Base / 2.0, Largo / 2.0, H + h_col + 0.18,
        f"Columna {L_C:.2f}×{B_C:.2f} m\nMalla Ø {varilla} @ {espaciamiento:.2f} m",
        color='#5eead4', fontweight='bold', fontsize=8.8, ha='center',
        bbox=dict(boxstyle='round,pad=0.25', facecolor='#0f1b33', edgecolor='#0d9488', alpha=0.85)
    )
    ax.text(
        Nv_Base, 0.0, H + 0.06,
        f"Peralte H = {H:.2f} m",
        color='#fbbf24', fontweight='bold', fontsize=8.5, ha='right'
    )

    ax.set_xlim(0, Nv_Base)
    ax.set_ylim(0, Largo)
    ax.set_zlim(0, H + h_col + 0.30)
    ax.set_xlabel(f"Base B = {Nv_Base:.2f} m (X)", fontsize=8.8, fontweight='bold', labelpad=6)
    ax.set_ylabel(f"Largo L = {Largo:.2f} m (Y)", fontsize=8.8, fontweight='bold', labelpad=6)
    ax.set_zlabel(f"Altura Z (H={H:.2f}m)", fontsize=8.8, fontweight='bold', labelpad=4)

    try:
        ax.set_box_aspect((Nv_Base, Largo, (H + h_col) * 1.15))
    except Exception:
        pass

    ax.view_init(elev=25, azim=-46)
    ax.set_title(
        f"2. MODELO VOLUMÉTRICO 3D Y ARMADURA\nZapata {Nv_Base:.2f}×{Largo:.2f}×{H:.2f} m | Parrilla Ø {varilla} @ {espaciamiento:.2f} m",
        fontsize=11, fontweight='bold', color='#f8fafc', pad=10
    )

    fig.savefig(ruta_img_3d, dpi=140, bbox_inches='tight', facecolor=fig.get_facecolor())
    plt.close(fig)


def generar_grafico_esfuerzos_suelo(
    esf_max_CS, esf_min_CS, esf_max_Sx, esf_min_Sx,
    esf_max_Sy, esf_min_Sy, Qatn, Qatn_Sx, ruta_img_esfuerzos
):
   
    fig, ax = plt.subplots(figsize=(6.8, 5.8))
    _aplicar_estilo_cad_oscuro(
        ax, fig,
        "3. VERIFICACIÓN DE PRESIONES EN EL SUELO (Tn/m²)\nDemanda Actuante (σ máx / σ mín) vs. Capacidad Portante Admisible"
    )

    estados = ['Carga Servicio\n(CM + CV)', 'Sismo Dirección X\n(CS + 0.8 Sx)', 'Sismo Dirección Y\n(CS + 0.8 Sy)']
    x = np.arange(len(estados))
    ancho = 0.25

    maximos = [esf_max_CS, esf_max_Sx, esf_max_Sy]
    minimos = [esf_min_CS, esf_min_Sx, esf_min_Sy]
    admisibles = [Qatn, Qatn_Sx, Qatn_Sx]

    b_max = ax.bar(x - ancho, maximos, width=ancho, label='σ máx actuante (Tn/m²)', color='#38bdf8', edgecolor='#bae6fd', linewidth=1.1, zorder=3)
    b_min = ax.bar(x, minimos, width=ancho, label='σ mín actuante (Tn/m²)', color='#10b981', edgecolor='#a7f3d0', linewidth=1.1, zorder=3)
    b_adm = ax.bar(x + ancho, admisibles, width=ancho, label='q admisible límite (Tn/m²)', color='#6366f1', edgecolor='#c7d2fe', linewidth=1.1, alpha=0.85, zorder=3)

    # Etiquetas numéricas sobre cada barra
    for bars, color_txt in [(b_max, '#7dd3fc'), (b_min, '#6ee7b7'), (b_adm, '#c7d2fe')]:
        for bar in bars:
            h = bar.get_height()
            ax.annotate(
                f"{h:.2f}",
                xy=(bar.get_x() + bar.get_width() / 2.0, h),
                xytext=(0, 3), textcoords="offset points",
                ha='center', va='bottom', color=color_txt, fontsize=8.2, fontweight='bold'
            )

    # Indicador de Ratio de Uso colocado justo encima de la barra más alta de cada grupo
    y_max_global = max(max(maximos), max(admisibles))
    y_top = y_max_global * 1.34
    for idx in range(3):
        ratio_pct = (maximos[idx] / admisibles[idx]) * 100.0 if admisibles[idx] > 0 else 0.0
        ok = (maximos[idx] <= admisibles[idx]) and (minimos[idx] > 0)
        badge_txt = f"{'✓ OK' if ok else '✗ REVISAR'} ({ratio_pct:.1f}%)"
        badge_col = '#059669' if ok else '#dc2626'
        y_grupo = max(maximos[idx], admisibles[idx]) + y_max_global * 0.11
        ax.text(
            x[idx], y_grupo, badge_txt,
            ha='center', va='bottom', color='#ffffff', fontweight='bold', fontsize=8.0,
            bbox=dict(boxstyle='round,pad=0.28', facecolor=badge_col, edgecolor='#a7f3d0' if ok else '#fecaca', lw=0.8, alpha=0.92)
        )

    ax.axhline(0, color='#ef4444', linestyle='--', linewidth=1.0, alpha=0.7)
    ax.set_xticks(x)
    ax.set_xticklabels(estados, fontsize=8.8, fontweight='bold', color='#e2e8f0')
    ax.set_ylabel("Esfuerzo en el Terreno (Tn/m²)", color='#cbd5e1', fontweight='bold', fontsize=9.2)
    ax.set_ylim(min(0.0, min(minimos) * 1.2), y_top)
    ax.grid(axis='y', linestyle='--', color='#1e3a5f', alpha=0.6, zorder=0)

    leg = ax.legend(
        loc='upper center', bbox_to_anchor=(0.5, 0.99), ncol=3,
        frameon=True, facecolor='#0b1325', edgecolor='#1e3a5f', fontsize=7.8
    )
    for text in leg.get_texts():
        text.set_color('#e2e8f0')

    fig.savefig(ruta_img_esfuerzos, dpi=140, bbox_inches='tight', facecolor=fig.get_facecolor())
    plt.close(fig)


def generar_grafico_verificacion_estructural(
    V_u1, phi_Vc1, V_u2, phi_Vc2, Vu_punz, phi_Vc_punz,
    MU, H, d, As_min, fc, fy, varilla_elegida, espaciamiento_m, ruta_img_estructural
):
   
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(7.4, 5.8))
    fig.patch.set_facecolor('#0b1325')
    fig.suptitle(
        "4. VERIFICACIÓN DE CORTE, PUNZONAMIENTO Y CAPACIDAD A FLEXIÓN",
        fontsize=11, fontweight='bold', color='#f8fafc', y=0.98
    )

    # --- SUBPLOT 1: CORTE Y PUNZONAMIENTO (% DE UTILIZACIÓN Y VALORES) ---
    _aplicar_estilo_cad_oscuro(ax1, fig, "Corte y Punzonamiento (Tn)")
    categorias = ['Corte\nVolado 1 (Y)', 'Corte\nVolado 2 (X)', 'Punzonamiento\nBidireccional (b₀)']
    demandas = [V_u1, V_u2, Vu_punz]
    capacidades = [phi_Vc1, phi_Vc2, phi_Vc_punz]

    y_pos = np.arange(len(categorias))
    h_bar = 0.34

    b1 = ax1.barh(y_pos + h_bar / 2.0, capacidades, height=h_bar, label='Capacidad φVc (Tn)', color='#10b981', edgecolor='#a7f3d0')
    b2 = ax1.barh(y_pos - h_bar / 2.0, demandas, height=h_bar, label='Demanda Vu (Tn)', color='#f59e0b', edgecolor='#fde68a')

    max_v = max(max(capacidades), max(demandas), 1.0)
    ax1.set_xlim(0, max_v * 1.34)
    ax1.set_yticks(y_pos)
    ax1.set_yticklabels(categorias, fontsize=8.2, fontweight='bold', color='#e2e8f0')
    ax1.set_xlabel("Fuerza Cortante (Tn)", color='#cbd5e1', fontsize=8.5, fontweight='bold')
    ax1.grid(axis='x', linestyle='--', color='#1e3a5f', alpha=0.6)

    for bar, col_t in [(b1, '#6ee7b7'), (b2, '#fcd34d')]:
        for rect in bar:
            w = rect.get_width()
            ax1.text(
                w + max_v * 0.02, rect.get_y() + rect.get_height() / 2.0,
                f"{w:.1f} Tn", va='center', ha='left', color=col_t, fontsize=7.8, fontweight='bold'
            )

    leg1 = ax1.legend(loc='lower right', frameon=True, facecolor='#0b1325', edgecolor='#1e3a5f', fontsize=7.5)
    for t in leg1.get_texts():
        t.set_color('#e2e8f0')

    # --- SUBPLOT 2: CAPACIDAD A FLEXIÓN φMn POR VARILLA (@ espaciamiento actual) ---
    _aplicar_estilo_cad_oscuro(ax2, fig, f"Flexión φMn vs Mu (@ s={espaciamiento_m:.2f}m)")
    varillas = list(TABLA_VARILLAS.keys())
    phi_mns = []
    colores_bar = []

    for v in varillas:
        r_v = verificar_armadura(MU, H, d, As_min, fc=fc, fy=fy, varilla_elegida=v, espaciamiento_m=espaciamiento_m)
        phi_mns.append(r_v["φMn_Tnm"])
        if v == varilla_elegida:
            colores_bar.append('#00e5ff' if r_v["es_favorable"] else '#ef4444')
        else:
            colores_bar.append('#3b82f6' if r_v["es_favorable"] else '#64748b')

    x_v = np.arange(len(varillas))
    bars_mn = ax2.bar(x_v, phi_mns, width=0.55, color=colores_bar, edgecolor='#e2e8f0', linewidth=1.0, zorder=3)
    ax2.axhline(MU, color='#f43f5e', linestyle='--', linewidth=1.8, label=f'Demanda Mu = {MU:.2f} Tn·m', zorder=4)

    max_m = max(max(phi_mns), MU, 1.0)
    ax2.set_ylim(0, max_m * 1.25)
    ax2.set_xticks(x_v)
    ax2.set_xticklabels(varillas, fontsize=8.5, fontweight='bold', color='#e2e8f0')
    ax2.set_ylabel("Momento φMn (Tn·m)", color='#cbd5e1', fontsize=8.5, fontweight='bold')
    ax2.grid(axis='y', linestyle='--', color='#1e3a5f', alpha=0.6, zorder=0)

    for idx, rect in enumerate(bars_mn):
        h_val = rect.get_height()
        es_actual = (varillas[idx] == varilla_elegida)
        etiqueta = f"★ {h_val:.1f}" if es_actual else f"{h_val:.1f}"
        offset_y = max_m * 0.05 if abs(h_val - MU) < (max_m * 0.06) else max_m * 0.02
        ax2.text(
            rect.get_x() + rect.get_width() / 2.0, h_val + offset_y,
            etiqueta, ha='center', va='bottom',
            color='#00e5ff' if es_actual else '#cbd5e1',
            fontsize=8.0 if es_actual else 7.5, fontweight='bold'
        )

    leg2 = ax2.legend(loc='upper left', frameon=True, facecolor='#0b1325', edgecolor='#1e3a5f', fontsize=7.5)
    for t in leg2.get_texts():
        t.set_color('#e2e8f0')

    plt.tight_layout(rect=[0, 0, 1, 0.94])
    fig.savefig(ruta_img_estructural, dpi=140, bbox_inches='tight', facecolor=fig.get_facecolor())
    plt.close(fig)



# 3. FUNCIÓN PRINCIPAL DE CÁLCULO DE ZAPATA Y GENERACIÓN DE LOS 4 GRÁFICOS


def calcular_zapata(
    origen_datos=r"DATOS DE ZAPATA.csv",
    q_adm=1.8,
    fc=210,
    H_inicial=0.60,
    L_C=0.30,
    B_C=0.60,
    varilla_elegida='5/8"',
    espaciamiento_m=0.15,
    As_propuesto=None,
    fy=4200
):
    
    # 1. Carga o recepción de datos
    if isinstance(origen_datos, pd.DataFrame):
        df = origen_datos.copy().reset_index(drop=True)
    else:
        ruta_busqueda = str(origen_datos)
        if not os.path.exists(ruta_busqueda):
            dir_script = os.path.dirname(os.path.abspath(__file__))
            candidatos = [
                os.path.join(dir_script, "DATOS DE ZAPATA.csv"),
                os.path.join(dir_script, "DATOS", "DATOS DE ZAPATA.csv"),
                os.path.join(dir_script, "RECURSOS", "DATOS", "DATOS DE ZAPATA.csv"),
                os.path.join(os.path.dirname(dir_script), ruta_busqueda)
            ]
            for cand in candidatos:
                if os.path.exists(cand):
                    ruta_busqueda = cand
                    break
        info_archivo = cargar_datos_archivo(ruta_busqueda)
        df = info_archivo["df_cargas"].copy()
        if q_adm is None and info_archivo["q_adm"] is not None:
            q_adm = info_archivo["q_adm"]
        if L_C is None and info_archivo["L_C"] is not None:
            L_C = info_archivo["L_C"]
        if B_C is None and info_archivo["B_C"] is not None:
            B_C = info_archivo["B_C"]

    Qadm = float(q_adm) if q_adm is not None else 1.8
    Qatn = round(Qadm * 10.0, 2)
    L_C = float(L_C) if L_C is not None else 0.30
    B_C = float(B_C) if B_C is not None else 0.60
    fc = float(fc) if fc is not None else 210.0
    fy = float(fy) if fy is not None else 4200.0

    # Guardar copia de las 4 cargas de entrada originales (CM, CV, Sx, Sy)
    df_entrada = df.iloc[0:4, 0:3].copy()
    df_entrada.index = ["Carga Muerta (CM)", "Carga Viva (CV)", "Sismo en X (Sx)", "Sismo en Y (Sy)"]

    # 2. Procesamiento de cargas y combinaciones normativas
    # Fila 0: CM + CV (Carga de servicio)
    df.iloc[0:1, 0:3] = df.iloc[0:2, 0:3].sum()
    # Fila 1: Sismo en X reducido al 80%
    df.iloc[1:2, 0:3] = df.iloc[2:3, 0:3] * 0.8
    # Fila 2: Sismo en Y reducido al 80%
    df.iloc[2:3, 0:3] = df.iloc[3:4, 0:3] * 0.8
    # Eliminamos fila 3 y reseteamos índice
    df = df.drop(3).reset_index(drop=True)

    P_CS, Mx_CS, My_CS = float(df.iloc[0, 0]), float(df.iloc[0, 1]), float(df.iloc[0, 2])
    P_Sx, Mx_Sx, My_Sx = float(df.iloc[1, 0]), float(df.iloc[1, 1]), float(df.iloc[1, 2])
    P_Sy, Mx_Sy, My_Sy = float(df.iloc[2, 0]), float(df.iloc[2, 1]), float(df.iloc[2, 2])

    # 3. Factor de peso propio y relleno según Qadm
    if Qadm < 1.5:
        factor = 1.10
    elif 1.5 <= Qadm < 2.5:
        factor = 1.07
    else:
        factor = 1.05

    # Área tentativa de la zapata
    denominador_q = max(1.0, (Qadm * 10.0) - 3.0)
    A_tentativa = round((P_CS * factor) / denominador_q, 3)

    # Dimensiones preliminares de la zapata (mínimo 1.50 m)
    Base = round(np.sqrt(A_tentativa), 1)
    if Base < 1.5:
        Base = 1.5

    Largo = round(A_tentativa * 1.2 / Base, 1)
    if Largo < 1.5:
        Largo = 1.5

    A_D = round(Base * Largo, 2)
    Ix = round((Base * Largo**3) / 12.0, 3)
    Iy = round((Largo * Base**3) / 12.0, 3)

    # 4. Verificación bajo Carga de Servicio (CS)
    esfuerzo_maximo_CS = round(((P_CS * factor / A_D) + (Mx_CS * (Largo / 2.0) / Ix) + (My_CS * (Base / 2.0) / Iy)), 3)
    esfuerzo_minimo_CS = round(((P_CS * factor / A_D) - (Mx_CS * (Largo / 2.0) / Ix) - (My_CS * (Base / 2.0) / Iy)), 3)

    while esfuerzo_maximo_CS >= Qatn:
        Base = round(Base + 0.1, 2)
        A_D = round(Base * Largo, 2)
        Ix = round((Base * Largo**3) / 12.0, 3)
        Iy = round((Largo * Base**3) / 12.0, 3)
        esfuerzo_maximo_CS = round(((P_CS * factor / A_D) + (Mx_CS * (Largo / 2.0) / Ix) + (My_CS * (Base / 2.0) / Iy)), 3)
        esfuerzo_minimo_CS = round(((P_CS * factor / A_D) - (Mx_CS * (Largo / 2.0) / Ix) - (My_CS * (Base / 2.0) / Iy)), 3)

    # 5. Verificación bajo Sismo (X e Y)
    Nv_Base = Base
    Nv_A_D = A_D
    Qatn_Sx = round(Qatn * 1.3, 2)

    esfuerzo_maximo_Sx = round((((P_CS + P_Sx) * factor / Nv_A_D) + ((Mx_CS + Mx_Sx) * (Largo / 2.0) / Ix) + ((My_CS + My_Sx) * (Nv_Base / 2.0) / Iy)), 3)
    esfuerzo_minimo_Sx = round((((P_CS + P_Sx) * factor / Nv_A_D) - ((Mx_CS + Mx_Sx) * (Largo / 2.0) / Ix) - ((My_CS + My_Sx) * (Nv_Base / 2.0) / Iy)), 3)
    esfuerzo_maximo_Sy = round((((P_CS + P_Sy) * factor / Nv_A_D) + ((Mx_CS + Mx_Sy) * (Largo / 2.0) / Ix) + ((My_CS + My_Sy) * (Nv_Base / 2.0) / Iy)), 3)
    esfuerzo_minimo_Sy = round((((P_CS + P_Sy) * factor / Nv_A_D) - ((Mx_CS + Mx_Sy) * (Largo / 2.0) / Ix) - ((My_CS + My_Sy) * (Nv_Base / 2.0) / Iy)), 3)

    while esfuerzo_maximo_Sx >= Qatn_Sx or esfuerzo_maximo_Sy >= Qatn_Sx:
        Nv_Base = round(Nv_Base + 0.2, 2)
        Nv_A_D = round(Nv_Base * Largo, 2)
        Ix = round((Nv_Base * Largo**3) / 12.0, 3)
        Iy = round((Largo * Nv_Base**3) / 12.0, 3)

        esfuerzo_maximo_CS = round(((P_CS * factor / Nv_A_D) + (Mx_CS * (Largo / 2.0) / Ix) + (My_CS * (Nv_Base / 2.0) / Iy)), 3)
        esfuerzo_minimo_CS = round(((P_CS * factor / Nv_A_D) - (Mx_CS * (Largo / 2.0) / Ix) - (My_CS * (Nv_Base / 2.0) / Iy)), 3)

        esfuerzo_maximo_Sx = round((((P_CS + P_Sx) * factor / Nv_A_D) + ((Mx_CS + Mx_Sx) * (Largo / 2.0) / Ix) + ((My_CS + My_Sx) * (Nv_Base / 2.0) / Iy)), 3)
        esfuerzo_minimo_Sx = round((((P_CS + P_Sx) * factor / Nv_A_D) - ((Mx_CS + Mx_Sx) * (Largo / 2.0) / Ix) - ((My_CS + My_Sx) * (Nv_Base / 2.0) / Iy)), 3)

        esfuerzo_maximo_Sy = round((((P_CS + P_Sy) * factor / Nv_A_D) + ((Mx_CS + Mx_Sy) * (Largo / 2.0) / Ix) + ((My_CS + My_Sy) * (Nv_Base / 2.0) / Iy)), 3)
        esfuerzo_minimo_Sy = round((((P_CS + P_Sy) * factor / Nv_A_D) - ((Mx_CS + Mx_Sy) * (Largo / 2.0) / Ix) - ((My_CS + My_Sy) * (Nv_Base / 2.0) / Iy)), 3)

    # 6. Diseño por Corte Unidireccional (Volados 1 y 2) y Punzonamiento Bidireccional
    qmax_CS = round(esfuerzo_maximo_CS * 1.25, 4)
    qmax_S = round(max(esfuerzo_maximo_Sx, esfuerzo_maximo_Sy) * 1.55, 4)

    Volado_1 = round(((Largo / 2.0) - (B_C / 2.0)), 2)
    Volado_2 = round(((Nv_Base / 2.0) - (L_C / 2.0)), 2)

    Wu = round(max(qmax_CS, qmax_S), 4)
    Vu_1 = round(Wu * Volado_1, 2)
    Vu_2 = round(Wu * Volado_2, 2)

    H = float(H_inicial)
    d = round(H - 0.10, 2)
    b = 0.10

    # Verificación en Volado 1
    V_u1 = round(Wu * max(0.0, Volado_1 - d), 2)
    M_u1 = round(Wu * (Volado_1**2) / 2.0, 2)
    φVc1 = round(0.85 * 0.53 * np.sqrt(fc) * b * d * 10**2, 2)
    while φVc1 < V_u1:
        H = round(H + 0.10, 2)
        d = round(H - 0.10, 2)
        V_u1 = round(Wu * max(0.0, Volado_1 - d), 2)
        φVc1 = round(0.85 * 0.53 * np.sqrt(fc) * b * d * 10**2, 2)

    # Verificación en Volado 2
    V_u2 = round(Wu * max(0.0, Volado_2 - d), 2)
    M_u2 = round(Wu * (Volado_2**2) / 2.0, 2)
    φVc2 = round(0.85 * 0.53 * np.sqrt(fc) * b * d * 10**2, 2)
    while φVc2 < V_u2:
        H = round(H + 0.10, 2)
        d = round(H - 0.10, 2)
        V_u2 = round(Wu * max(0.0, Volado_2 - d), 2)
        φVc2 = round(0.85 * 0.53 * np.sqrt(fc) * b * d * 10**2, 2)

    # Verificación por Punzonamiento Bidireccional (Sección crítica a d/2)
    Am = round((B_C + d) * (L_C + d), 2)
    Pm = round(2.0 * (B_C + d) + 2.0 * (L_C + d), 2)
    Vu_Punzonamiento = round(Wu * max(0.0, Nv_A_D - Am), 2)
    φVc_Punzonamiento = round(0.85 * 1.06 * np.sqrt(fc) * Pm * d * 10.0, 2)

    while φVc_Punzonamiento < Vu_Punzonamiento:
        H = round(H + 0.10, 2)
        d = round(H - 0.10, 2)
        Am = round((B_C + d) * (L_C + d), 2)
        Pm = round(2.0 * (B_C + d) + 2.0 * (L_C + d), 2)
        Vu_Punzonamiento = round(Wu * max(0.0, Nv_A_D - Am), 2)
        φVc_Punzonamiento = round(0.85 * 1.06 * np.sqrt(fc) * Pm * d * 10.0, 2)
        V_u1 = round(Wu * max(0.0, Volado_1 - d), 2)
        V_u2 = round(Wu * max(0.0, Volado_2 - d), 2)
        φVc1 = round(0.85 * 0.53 * np.sqrt(fc) * b * d * 10**2, 2)
        φVc2 = round(0.85 * 0.53 * np.sqrt(fc) * b * d * 10**2, 2)

    # 7. Momento Último y Acero Mínimo
    As_min = round(0.0018 * H * 10**4, 2)
    MU = round(max(M_u1, M_u2), 2)

    # 8. Evaluación de Acero Modular
    res_acero = verificar_armadura(
        MU=MU,
        H=H,
        d=d,
        As_min=As_min,
        fc=fc,
        fy=fy,
        varilla_elegida=varilla_elegida,
        espaciamiento_m=espaciamiento_m,
        As_propuesto=As_propuesto
    )

    tabla_alternativas = generar_tabla_alternativas(MU, H, d, As_min, fc=fc, fy=fy)

    # 9. Generación de los 4 Gráficos de Cálculos y Resultados en RECURSOS/IMAGENES
    directorio_actual = os.path.dirname(os.path.abspath(__file__))
    dir_imagenes = os.path.join(directorio_actual, "RECURSOS", "IMAGENES")
    os.makedirs(dir_imagenes, exist_ok=True)

    ruta_img_2d = os.path.join(dir_imagenes, "planta_2d.png")
    ruta_img_3d = os.path.join(dir_imagenes, "modelo_3d.png")
    ruta_img_esfuerzos = os.path.join(dir_imagenes, "esfuerzos_suelo.png")
    ruta_img_estructural = os.path.join(dir_imagenes, "verificacion_estructural.png")

    generar_grafico_planta_2d(Nv_Base, Largo, L_C, B_C, d, varilla_elegida, espaciamiento_m, ruta_img_2d)
    generar_grafico_modelo_3d(Nv_Base, Largo, H, L_C, B_C, varilla_elegida, espaciamiento_m, ruta_img_3d)
    generar_grafico_esfuerzos_suelo(
        esfuerzo_maximo_CS, esfuerzo_minimo_CS,
        esfuerzo_maximo_Sx, esfuerzo_minimo_Sx,
        esfuerzo_maximo_Sy, esfuerzo_minimo_Sy,
        Qatn, Qatn_Sx, ruta_img_esfuerzos
    )
    generar_grafico_verificacion_estructural(
        V_u1, φVc1, V_u2, φVc2, Vu_Punzonamiento, φVc_Punzonamiento,
        MU, H, d, As_min, fc, fy, varilla_elegida, espaciamiento_m, ruta_img_estructural
    )

    return {
        "Base_final": Nv_Base,
        "Largo_final": Largo,
        "Area_final": Nv_A_D,
        "Area_tentativa": A_tentativa,
        "factor_peso": factor,
        "Ix": Ix,
        "Iy": Iy,
        "Peralte_H": round(H, 2),
        "d_efectivo": round(d, 2),
        "L_C": round(L_C, 2),
        "B_C": round(B_C, 2),
        "Volado_1": Volado_1,
        "Volado_2": Volado_2,
        "qadm": Qadm,
        "Qatn": Qatn,
        "Qatn_Sx": Qatn_Sx,
        "fc": fc,
        "fy": fy,
        "df_entrada": df_entrada,
        "df_combinaciones": df,
        "esfuerzo_maximo_CS": esfuerzo_maximo_CS,
        "esfuerzo_minimo_CS": esfuerzo_minimo_CS,
        "esfuerzo_maximo_Sx": esfuerzo_maximo_Sx,
        "esfuerzo_minimo_Sx": esfuerzo_minimo_Sx,
        "esfuerzo_maximo_Sy": esfuerzo_maximo_Sy,
        "esfuerzo_minimo_Sy": esfuerzo_minimo_Sy,
        "cumple_suelo_CS": bool(esfuerzo_maximo_CS <= Qatn and esfuerzo_minimo_CS > 0),
        "cumple_suelo_Sismo": bool(esfuerzo_maximo_Sx <= Qatn_Sx and esfuerzo_maximo_Sy <= Qatn_Sx and esfuerzo_minimo_Sx > 0 and esfuerzo_minimo_Sy > 0),
        "qmax_CS": round(qmax_CS, 2),
        "qmax_S": round(qmax_S, 2),
        "Wu": round(Wu, 2),
        "Vu_1_cara": Vu_1,
        "Vu_2_cara": Vu_2,
        "V_u1": V_u1,
        "φVc1": φVc1,
        "phi_Vc1": φVc1,
        "M_u1": M_u1,
        "V_u2": V_u2,
        "φVc2": φVc2,
        "phi_Vc2": φVc2,
        "M_u2": M_u2,
        "cumple_corte_1": bool(φVc1 >= V_u1),
        "cumple_corte_2": bool(φVc2 >= V_u2),
        "Am": Am,
        "Pm": Pm,
        "Vu_Punzonamiento": Vu_Punzonamiento,
        "φVc_Punzonamiento": φVc_Punzonamiento,
        "phi_Vc_Punzonamiento": φVc_Punzonamiento,
        "cumple_punzonamiento": bool(φVc_Punzonamiento >= Vu_Punzonamiento),
        "Momento_Ultimo": MU,
        "Momento_Ultimo_Mu": MU,
        "Acero_minimo": As_min,
        "As_min": As_min,
        "armadura": res_acero,
        "tabla_alternativas": tabla_alternativas,
        "ruta_img_2d": ruta_img_2d,
        "ruta_img_3d": ruta_img_3d,
        "ruta_img_esfuerzos": ruta_img_esfuerzos,
        "ruta_img_estructural": ruta_img_estructural
    }


if __name__ == "__main__":
    datos_defecto = pd.DataFrame([
        [32.581, 0.1533, 1.06],
        [9.87, 0.0154, 0.0118],
        [8.8031, 2.7582, 10.9464],
        [15.9962, 9.28, 0.503]
    ], columns=['P', 'Mx', 'My'])

    res = calcular_zapata(origen_datos=datos_defecto)
    print("\n=== RESUMEN DE DISEÑO ESTRUCTURAL DE ZAPATA ===")
    print(f"Dimensiones Finales : B = {res['Base_final']} m | L = {res['Largo_final']} m | H = {res['Peralte_H']} m (d = {res['d_efectivo']} m)")
    print(f"Esfuerzos Suelo (CS): σ_máx = {res['esfuerzo_maximo_CS']} Tn/m² | σ_mín = {res['esfuerzo_minimo_CS']} Tn/m² (Límite = {res['Qatn']} Tn/m²)")
    print(f"Corte Unidireccional: Vu1 = {res['V_u1']} Tn, Vu2 = {res['V_u2']} Tn vs φVc = {res['φVc1']} Tn")
    print(f"Punzonamiento (b₀)  : Vup = {res['Vu_Punzonamiento']} Tn vs φVcp = {res['φVc_Punzonamiento']} Tn")
    print(f"Diseño por Flexión  : {res['armadura']['descripcion']} -> φMn = {res['armadura']['φMn_Tnm']} Tn·m vs Mu = {res['Momento_Ultimo_Mu']} Tn·m ({res['armadura']['estado']})")