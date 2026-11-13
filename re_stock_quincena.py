import streamlit as st
import pandas as pd
import io
import os
import xlwt
import time
from contextlib import contextmanager

# =====================================================================
# GENERADOR DE PLANTILLAS NAVASOFT (.xls) - misma estructura que la original
# Hoja1 | Codi, Codigo, Descripcion, Cantidad, Costo, Codalm
# =====================================================================
CABECERAS_NAVASOFT = ['Codi', 'Codigo', 'Descripcion', 'Cantidad', 'Costo', 'Codalm']
COD_ALMACEN = {'dansey': '08', 'san luis': '01'}


def _texto(v):
    if v is None or (not isinstance(v, str) and pd.isna(v)):
        return ''
    return str(v).strip()


def armar_filas_navasoft(df, cod_almacen):
    filas = []
    for _, r in df.iterrows():
        codigo = _texto(r.get('codigo_comercial'))
        codi = _texto(r.get('codi_interno')) or codigo
        desc = _texto(r.get('descripcion'))
        cant = pd.to_numeric(r.get('pedido'), errors='coerce')
        costo = pd.to_numeric(r.get('precio_unitario'), errors='coerce')
        cant = int(round(cant)) if pd.notna(cant) else 0
        costo = round(float(costo), 2) if pd.notna(costo) else 0.0
        if cant <= 0:
            continue
        filas.append([codi, codigo, desc, cant, costo, cod_almacen])
    return filas


def generar_xls_navasoft(filas):
    """Crea el .xls (BIFF8) igual a PLANTILLA NAVASOFT: hoja 'Hoja1', cabecera en fila 1,
    datos desde la fila 2 y sin filas vacías debajo. Devuelve bytes."""
    wb = xlwt.Workbook(encoding='utf-8')
    ws = wb.add_sheet('Hoja1')

    st_head = xlwt.easyxf('font: name Calibri, height 220;', num_format_str='@')
    st_txt = xlwt.easyxf('font: name Calibri, height 180;', num_format_str='@')
    st_cant = xlwt.easyxf('font: name Calibri, height 180;', num_format_str='#,##0')
    st_costo = xlwt.easyxf('font: name Calibri, height 180;', num_format_str='0.00')

    for c, h in enumerate(CABECERAS_NAVASOFT):
        ws.write(0, c, h, st_head)

    for i, (codi, codigo, desc, cant, costo, alm) in enumerate(filas, start=1):
        ws.write(i, 0, codi, st_txt)
        ws.write(i, 1, codigo, st_txt)
        ws.write(i, 2, desc[:255], st_txt)
        ws.write(i, 3, cant, st_cant)
        ws.write(i, 4, costo, st_costo)
        ws.write(i, 5, alm, st_txt)

    for c, w in enumerate([4000, 6000, 14000, 2500, 2500, 2500]):
        ws.col(c).width = w

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def procesar_maestro_navasoft(df_term):
    req = ['codi_interno', 'codigo_comercial', 'descripcion', 'pedido', 'precio_unitario', 'sucursal']
    faltan = [c for c in req if c not in df_term.columns]
    if faltan:
        raise ValueError(f"Faltan columnas en el archivo: {', '.join(faltan)}")

    df = df_term.copy()
    df['pedido'] = pd.to_numeric(df['pedido'], errors='coerce')
    df = df[df['pedido'] > 0]
    suc = df['sucursal'].astype(str).str.strip().str.lower()

    out = {}
    for nombre, cod in COD_ALMACEN.items():
        filas = armar_filas_navasoft(df[suc == nombre], cod)
        out[nombre] = (generar_xls_navasoft(filas), len(filas)) if filas else (None, 0)
    return out


# =====================================================================
# LOADER ANIMADO (icono llamativo mientras se procesa)
# =====================================================================
@contextmanager
def loader(texto, emoji="⚙️", min_seg=2.5):
    """Muestra el loader y lo mantiene al menos `min_seg` segundos (aunque el proceso sea más rápido)."""
    inicio = time.time()
    ph = st.empty()
    ph.markdown(
        f"""
        <div class="loader-box">
            <div class="loader-emoji">{emoji}</div>
            <div class="loader-text">{texto}<span class="dots"><span>.</span><span>.</span><span>.</span></span></div>
            <div class="loader-bar"><div class="loader-bar-fill"></div></div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    try:
        yield
    finally:
        restante = min_seg - (time.time() - inicio)
        if restante > 0:
            time.sleep(restante)
        ph.empty()


# Configuración de la página web
st.set_page_config(page_title="RodaRestock", page_icon="📦", layout="centered")

# Estilo visual limpio
st.markdown("""
    <style>
    .main { background-color: #0f172a; color: #f1f5f9; }
    .stButton>button { width: 100%; background-color: #059669; color: white; font-weight: bold; border-radius: 10px; padding: 14px; font-size: 16px; }
    .stButton>button:hover { background-color: #047857; color: white; }
    .stTabs [data-baseweb="tab-list"] { gap: 10px; }
    .stTabs [data-baseweb="tab"] { background-color: #1e293b; color: #f1f5f9; border-radius: 8px; padding: 10px 20px; font-weight: bold; }
    .stTabs [aria-selected="true"] { background-color: #059669; color: white; }
    .loader-box { text-align: center; padding: 22px; margin: 10px 0; background: #1e293b; border: 1px solid #334155; border-radius: 14px; }
    .loader-emoji { font-size: 54px; display: inline-block; animation: girar 1.2s linear infinite, latir 1.2s ease-in-out infinite; }
    .loader-text { color: #f1f5f9; font-weight: bold; font-size: 17px; margin-top: 8px; }
    .loader-bar { height: 6px; background: #0f172a; border-radius: 6px; margin-top: 14px; overflow: hidden; }
    .loader-bar-fill { height: 100%; width: 40%; background: linear-gradient(90deg, #059669, #60a5fa); border-radius: 6px; animation: correr 1.4s ease-in-out infinite; }
    .dots span { animation: parpadeo 1.4s infinite both; }
    .dots span:nth-child(2) { animation-delay: .2s; }
    .dots span:nth-child(3) { animation-delay: .4s; }
    @keyframes girar { from { transform: rotate(0deg); } to { transform: rotate(360deg); } }
    @keyframes latir { 0%, 100% { filter: drop-shadow(0 0 4px #059669); } 50% { filter: drop-shadow(0 0 14px #60a5fa); } }
    @keyframes correr { 0% { margin-left: -40%; } 100% { margin-left: 100%; } }
    @keyframes parpadeo { 0%, 80%, 100% { opacity: 0; } 40% { opacity: 1; } }
    </style>
""", unsafe_allow_html=True)

# =====================================================================
# ENCABEZADO: LOGO Y TÍTULO MÁS GRANDES ARRIBA (EDITABLES)
# =====================================================================

TITULO_APP = "RodaRestock"
SUBTITULO_APP = "Sistema inteligente de Cálculo de Reposición"

archivo_logo = "logo.jpg"

if os.path.exists(archivo_logo):
    import base64
    with open(archivo_logo, "rb") as img_file:
        encoded_img = base64.b64encode(img_file.read()).decode()
    
    st.markdown(f"""
        <div style="display: flex; flex-direction: column; align-items: center; justify-content: center; margin-top: 20px; margin-bottom: 10px;">
            <div style="display: flex; align-items: center; gap: 20px;">
                <img src="data:image/jpeg;base64,{encoded_img}" style="width: 110px; height: auto;">
                <h1 style='color: #60a5fa; margin: 0; font-size: 80px; font-weight: 800;'>{TITULO_APP}</h1>
            </div>
            <p style='color: #94a3b8; margin: 14px 0 0 0; font-size: 19px; text-align: center; font-weight: 500;'>{SUBTITULO_APP}</p>
        </div>
    """, unsafe_allow_html=True)
else:
    st.markdown(f"""
        <div style="text-align: center; margin-top: 20px; margin-bottom: 10px;">
            <h1 style='color: #60a5fa; margin: 0; font-size: 62px; font-weight: 800;'>📦 {TITULO_APP}</h1>
            <p style='color: #94a3b8; margin: 14px 0 0 0; font-size: 19px; text-align: center; font-weight: 500;'>{SUBTITULO_APP}</p>
        </div>
    """, unsafe_allow_html=True)

st.divider()

# Pestañas principales definidas correctamente
tab_calculo, tab_navasoft = st.tabs(["📊 1. Cálculo y Maestro de Transferencias", "📦 2. Generar Plantillas Navasoft (Dansey / San Luis)"])

with tab_calculo:
    almacen_destino = st.selectbox(
        "Selecciona la Sucursal de Destino",
        ["San Luis", "Dansey", "Chiclayo", "Arequipa", "Trujillo", "Piura"]
    )

    es_provincia = almacen_destino not in ["San Luis", "Dansey"]

    st.divider()
    st.subheader(f"📂 Archivos requeridos para Destino: {almacen_destino}")

    f_sisco = st.file_uploader(f"1. Reporte Sisco {almacen_destino} (Incluye su stock actual)", type=["xlsx", "xls"])
    f_consumo = st.file_uploader(f"2. OLAP de Consumo ({almacen_destino})", type=["xlsx", "xls"])

    if almacen_destino == "San Luis":
        f_stk_ref = st.file_uploader("3. Stock Actual (Dansey)", type=["xlsx", "xls"])
        f_stk_dansey = f_stk_ref
        f_stk_sanluis = None
    elif almacen_destino == "Dansey":
        f_stk_ref = st.file_uploader("3. Stock Actual (San Luis)", type=["xlsx", "xls"])
        f_stk_sanluis = f_stk_ref
        f_stk_dansey = None
    else:
        f_stk_dansey = st.file_uploader("3. Stock Actual (Dansey)", type=["xlsx", "xls"])
        f_stk_sanluis = st.file_uploader("4. Stock Actual (San Luis)", type=["xlsx", "xls"])

    st.divider()

    if st.button("🚀 Ejecutar Procesamiento Maestro"):
        if not es_provincia:
            archivos_ok = all([f_sisco, f_consumo, f_stk_ref])
        else:
            archivos_ok = all([f_sisco, f_consumo, f_stk_dansey, f_stk_sanluis])

        if not archivos_ok:
            st.error("Por favor, sube todos los archivos requeridos antes de continuar.")
        else:
            with loader(f"Procesando datos para {almacen_destino}", "⚡"):
                try:
                    ruta_precios = "archivo_codigos_preciou.xlsx"
                    ruta_cajas = "numero_cajas.xlsx"

                    if not os.path.exists(ruta_precios) or not os.path.exists(ruta_cajas):
                        st.error("No se encuentran los archivos maestros locales 'archivo_codigos_preciou.xlsx' o 'numero_cajas.xlsx' en la carpeta.")
                    else:
                        def limpiar_codigo(cod):
                            if pd.isna(cod): return ""
                            s = str(cod).strip().upper()
                            if s.startswith('/HO/'): s = s[4:]
                            return s.replace('*', '').strip()

                        def limpiar_marca(m):
                            if pd.isna(m): return ""
                            return str(m).strip().upper()

                        df_precios = pd.read_excel(ruta_precios, engine='openpyxl')
                        col_desc = next((c for c in df_precios.columns if 'desc' in str(c).lower()), df_precios.columns[2] if len(df_precios.columns) > 2 else None)
                        col_marca_precios = next((c for c in df_precios.columns if 'marc' in str(c).lower() or 'marca' in str(c).lower()), None)

                        mapa_maestro = {}
                        for _, r_p in df_precios.iterrows():
                            c_codi = limpiar_codigo(r_p.get('codi', ''))
                            c_codf = limpiar_codigo(r_p.get('codf', ''))
                            c_marca = limpiar_marca(r_p.get(col_marca_precios, '')) if col_marca_precios else ""

                            precio_val = pd.to_numeric(r_p.get('precio', 0), errors='coerce')
                            precio_val = float(precio_val) if pd.notna(precio_val) else 0.0
                            desc_val = str(r_p.get(col_desc, '')).strip() if col_desc else ""

                            mapa_maestro[(c_codf, c_marca)] = {'codi': c_codi, 'precio': precio_val, 'descripcion': desc_val}
                            mapa_maestro[c_codi] = {'codi': c_codi, 'precio': precio_val, 'descripcion': desc_val}

                        df_cajas = pd.read_excel(ruta_cajas, engine='openpyxl')
                        col_cod_caja = df_cajas.columns[0]
                        col_cant_caja = df_cajas.columns[1] if len(df_cajas.columns) > 1 else df_cajas.columns[0]

                        mapa_cajas = {}
                        for _, r_c in df_cajas.iterrows():
                            c_clean = limpiar_codigo(r_c[col_cod_caja])
                            cant_val = pd.to_numeric(r_c[col_cant_caja], errors='coerce')
                            mapa_cajas[c_clean] = float(cant_val) if pd.notna(cant_val) else 1.0

                        df_crit = pd.read_excel(f_sisco, skiprows=6, engine='openpyxl')
                        col_codigo_crit = next((c for c in df_crit.columns if 'código' in str(c).lower() or 'codigo' in str(c).lower()), df_crit.columns[4])
                        col_marca_crit = next((c for c in df_crit.columns if 'marca' in str(c).lower()), df_crit.columns[5] if len(df_crit.columns) > 5 else None)
                        col_stk_dest = next((c for c in df_crit.columns if 'stock' in str(c).lower() or 'stoc' in str(c).lower()), df_crit.columns[-1] if len(df_crit.columns) > 0 else None)

                        nombre_col_dest = almacen_destino.upper()
                        lista_base = []

                        for idx, row in df_crit.iterrows():
                            val_cod = row[col_codigo_crit]
                            if pd.isna(val_cod) or str(val_cod).strip() in ['', 'nan']: continue

                            codigo_leido = str(val_cod).strip().upper()
                            codigo_limpio = limpiar_codigo(codigo_leido)
                            marca = limpiar_marca(row[col_marca_crit]) if col_marca_crit and pd.notna(row[col_marca_crit]) else ""

                            info_maestro = mapa_maestro.get((codigo_limpio, marca), None)
                            if not info_maestro:
                                info_maestro = mapa_maestro.get(codigo_limpio, {'codi': codigo_limpio, 'precio': 0.0, 'descripcion': ''})

                            codi_interno = info_maestro['codi'] if info_maestro['codi'] else codigo_limpio
                            codf_comercial = codigo_limpio
                            precio = info_maestro['precio']
                            descripcion = info_maestro['descripcion']

                            cant_caja = mapa_cajas.get(codi_interno, mapa_cajas.get(codf_comercial, 1.0))

                            stk_dest_val = 0.0
                            if col_stk_dest:
                                try:
                                    val_s = pd.to_numeric(row[col_stk_dest], errors='coerce')
                                    stk_dest_val = float(val_s) if pd.notna(val_s) else 0.0
                                except:
                                    stk_dest_val = 0.0

                            lista_base.append({
                                'codigo_comercial': codf_comercial,
                                'codi_interno': codi_interno,
                                'descripcion': descripcion,
                                'marca': marca,
                                'cant_caja': cant_caja,
                                'precio_unitario': precio,
                                'prioridad_sisco': idx,
                                nombre_col_dest: stk_dest_val
                            })

                        df_maestro_base = pd.DataFrame(lista_base)

                        df_raw_olap = pd.read_excel(f_consumo, sheet_name='Resultado', header=None)
                        fila_inicio = 8
                        for r in range(len(df_raw_olap)):
                            val = df_raw_olap.iloc[r, 1]
                            if pd.notna(val) and str(val).strip().lower() == 'codigo':
                                fila_inicio = r + 1
                                break

                        fila_meses = fila_inicio - 1
                        fila_anios = fila_meses - 1

                        indices_2025 = []
                        indices_2026 = []
                        anio_actual_en_curso = ""

                        for c in range(3, df_raw_olap.shape[1]):
                            val_anio = df_raw_olap.iloc[fila_anios, c] if fila_anios >= 0 else None
                            val_mes = df_raw_olap.iloc[fila_meses, c] if fila_meses >= 0 else None

                            txt_anio = str(val_anio).strip() if pd.notna(val_anio) else ""
                            txt_mes = str(val_mes).strip().lower() if pd.notna(val_mes) else ""

                            if '2025' in txt_anio: anio_actual_en_curso = "2025"
                            elif '2026' in txt_anio: anio_actual_en_curso = "2026"

                            if 'total' in txt_mes or txt_mes in ['nan', 'none', '']: continue

                            if anio_actual_en_curso == "2025": indices_2025.append(c)
                            elif anio_actual_en_curso == "2026": indices_2026.append(c)

                        dict_olap = {}
                        ultimo_codigo = ""
                        i = fila_inicio
                        while i < len(df_raw_olap):
                            codigo_val = df_raw_olap.iloc[i, 1]
                            if pd.notna(codigo_val):
                                cod_str = str(codigo_val).strip()
                                if cod_str.lower() not in ['', 'nan', 'total general']:
                                    ultimo_codigo = limpiar_codigo(cod_str)

                            if not ultimo_codigo:
                                i += 1
                                continue

                            marca_olap = limpiar_marca(df_raw_olap.iloc[i, 2]) if pd.notna(df_raw_olap.iloc[i, 2]) else ""

                            todos_indices = indices_2025 + indices_2026
                            todos_vals = []
                            for c in todos_indices:
                                if c < df_raw_olap.shape[1]:
                                    v = pd.to_numeric(df_raw_olap.iloc[i, c], errors='coerce')
                                    if pd.notna(v): todos_vals.append(v)
                            total_gral = sum(todos_vals)

                            vals_2025 = []
                            for c in indices_2025:
                                if c < df_raw_olap.shape[1]:
                                    v = pd.to_numeric(df_raw_olap.iloc[i, c], errors='coerce')
                                    if pd.notna(v): vals_2025.append(v)
                            mov_2025 = [v for v in vals_2025 if v > 0]
                            prom_2025 = (sum(mov_2025) / len(mov_2025)) if len(mov_2025) > 0 else 0

                            vals_2026 = []
                            for c in indices_2026:
                                if c < df_raw_olap.shape[1]:
                                    v = pd.to_numeric(df_raw_olap.iloc[i, c], errors='coerce')
                                    if pd.notna(v): vals_2026.append(v)
                            mov_2026 = [v for v in vals_2026 if v > 0]
                            prom_2026 = (sum(mov_2026) / len(mov_2026)) if len(mov_2026) > 0 else 0

                            mayor_prom = round(max(prom_2025, prom_2026))

                            dict_olap[(ultimo_codigo, marca_olap)] = {
                                'total_general': total_gral,
                                'mayor_promedio_mensual': mayor_prom
                            }
                            i += 1

                        def extraer_olap(row):
                            codf = row['codigo_comercial']
                            m_row = row['marca']
                            if (codf, m_row) in dict_olap:
                                datos = dict_olap[(codf, m_row)]
                                return pd.Series([datos['total_general'], datos['mayor_promedio_mensual']])
                            for (k_cod, k_marca), datos in dict_olap.items():
                                if codf == k_cod and (m_row in k_marca or k_marca in m_row or not m_row or not k_marca):
                                    return pd.Series([datos['total_general'], datos['mayor_promedio_mensual']])
                            return pd.Series([0, 0])

                        df_maestro_base[['total_general', 'mayor_promedio_mensual']] = df_maestro_base.apply(extraer_olap, axis=1)

                        def procesar_stock_sisco_con_marca(file_obj, nombre_col):
                            df_s = pd.read_excel(file_obj, skiprows=6, engine='openpyxl')
                            col_cod = next((c for c in df_s.columns if 'código' in str(c).lower() or 'codigo' in str(c).lower()), df_s.columns[4])
                            col_marca = next((c for c in df_s.columns if 'marca' in str(c).lower()), df_s.columns[5] if len(df_s.columns) > 5 else None)
                            col_stk = next((c for c in df_s.columns if 'stoc' in str(c).lower() or 'físico' in str(c).lower() or 'fisico' in str(c).lower()), df_s.columns[10])

                            stk_data = []
                            for _, r_s in df_s.iterrows():
                                c_val = r_s[col_cod]
                                if pd.isna(c_val): continue
                                c_clean = limpiar_codigo(c_val)
                                m_clean = limpiar_marca(r_s[col_marca]) if col_marca and pd.notna(r_s[col_marca]) else ""
                                s_val = pd.to_numeric(r_s[col_stk], errors='coerce')
                                s_val = float(s_val) if pd.notna(s_val) else 0.0

                                info_m = mapa_maestro.get((c_clean, m_clean), {'codi': c_clean})
                                c_codi_real = info_m['codi'] if info_m['codi'] else c_clean
                                stk_data.append({'codigo': c_codi_real, 'codigo_comercial': c_clean, 'marca': m_clean, nombre_col: s_val})

                            df_stk = pd.DataFrame(stk_data)
                            return df_stk.groupby(['codigo_comercial', 'marca'], as_index=False)[[nombre_col]].sum()

                        df_final = df_maestro_base.copy()

                        if not es_provincia:
                            nombre_col_ref = 'DANSEY' if almacen_destino == 'San Luis' else 'SAN LUIS'
                            stk_ref = procesar_stock_sisco_con_marca(f_stk_ref, nombre_col_ref)

                            df_final = pd.merge(df_final, stk_ref, left_on=['codigo_comercial', 'marca'], right_on=['codigo_comercial', 'marca'], how='left')
                            df_final[nombre_col_ref] = df_final[nombre_col_ref].fillna(0)

                            def calcular_reposicion(row):
                                ventas = row['total_general']
                                if ventas < 50: return 0, '⚪ Baja Rotación'
                                stk_actual_dest = row[nombre_col_dest]
                                prom_alto = row['mayor_promedio_mensual']
                                balance = int(round(stk_actual_dest - prom_alto))
                                if balance >= 0: return balance, '🟢 Stock Suficiente / Excedente'
                                else: return balance, '🔴 Déficit (Faltante)'

                            resultados = df_final.apply(calcular_reposicion, axis=1)
                            df_final['cantidad_a_transferir'] = [r[0] for r in resultados]
                            df_final['estado_reposicion'] = [r[1] for r in resultados]
                            df_final['pedido'] = df_final['cantidad_a_transferir']
                            df_final['sucursal'] = 'Dansey' if almacen_destino == 'San Luis' else 'San Luis'

                            columnas_ordenadas = [
                                'codigo_comercial', 'codi_interno', 'descripcion', 'marca', 'cant_caja',
                                'total_general', 'mayor_promedio_mensual', nombre_col_dest, nombre_col_ref,
                                'precio_unitario', 'cantidad_a_transferir', 'pedido', 'sucursal', 'estado_reposicion'
                            ]
                        else:
                            stk_dan = procesar_stock_sisco_con_marca(f_stk_dansey, 'DANSEY')
                            stk_sl = procesar_stock_sisco_con_marca(f_stk_sanluis, 'SAN LUIS')

                            df_final = pd.merge(df_final, stk_dan, left_on=['codigo_comercial', 'marca'], right_on=['codigo_comercial', 'marca'], how='left')
                            df_final = pd.merge(df_final, stk_sl, left_on=['codigo_comercial', 'marca'], right_on=['codigo_comercial', 'marca'], how='left', suffixes=('_dan', '_sl'))

                            df_final[nombre_col_dest] = df_final[nombre_col_dest].fillna(0)
                            if 'DANSEY' in df_final.columns: df_final['DANSEY'] = df_final['DANSEY'].fillna(0)
                            if 'SAN LUIS' in df_final.columns: df_final['SAN LUIS'] = df_final['SAN LUIS'].fillna(0)

                            def calcular_reposicion_prov(row):
                                ventas = row['total_general']
                                if ventas < 50: return 0, '⚪ Baja Rotación'
                                stk_actual_dest = row[nombre_col_dest]
                                prom_alto = row['mayor_promedio_mensual']
                                balance = int(round(stk_actual_dest - prom_alto))
                                if balance >= 0: return balance, '🟢 Stock Suficiente / Excedente'
                                else: return balance, '🔴 Déficit (Faltante)'

                            resultados = df_final.apply(calcular_reposicion_prov, axis=1)
                            df_final['cantidad_a_transferir'] = [r[0] for r in resultados]
                            df_final['estado_reposicion'] = [r[1] for r in resultados]

                            df_final['pedido'] = df_final['cantidad_a_transferir'].apply(lambda x: abs(x) if x < 0 else 0)
                            df_final['sucursal'] = ""

                            columnas_ordenadas = [
                                'codigo_comercial', 'codi_interno', 'descripcion', 'marca', 'cant_caja',
                                'total_general', 'mayor_promedio_mensual', nombre_col_dest,
                                'DANSEY', 'SAN LUIS',
                                'precio_unitario', 'cantidad_a_transferir', 'pedido', 'sucursal', 'estado_reposicion'
                            ]

                        df_final_ordenado = df_final.sort_values(by=['prioridad_sisco'], ascending=True)[columnas_ordenadas]

                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            df_final_ordenado.to_excel(writer, index=False)
                            if es_provincia:
                                worksheet = writer.sheets['Sheet1']
                                for row_idx in range(2, len(df_final_ordenado) + 2):
                                    worksheet[f'N{row_idx}'] = f'=IF(M{row_idx}<=0, "-", IF(M{row_idx}<=I{row_idx}, "Dansey", "San Luis"))'

                        output.seek(0)
                        st.success(f"¡Cálculo maestro para **{almacen_destino}** completado con éxito!")
                        st.download_button(
                            label=f"📥 Descargar Reporte de {almacen_destino}",
                            data=output,
                            file_name=f'Maestro_Transferencias_{almacen_destino}.xlsx',
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                except Exception as e:
                    st.error(f"Error al procesar los archivos: {str(e)}")

with tab_navasoft:
    st.subheader("📦 Generar Plantillas Navasoft (Dansey y San Luis)")

    f_terminado = st.file_uploader("Sube tu Excel de Transferencia Terminado", type=["xlsx"], key="terminado_uploader")

    if f_terminado and st.button("🚀 Procesar Plantillas Navasoft"):
        st.session_state.pop('nav', None)  # limpia descargas anteriores al empezar
        try:
            with loader("Leyendo y generando plantillas Navasoft", "📦"):
                df_term = pd.read_excel(f_terminado, engine='openpyxl')
                res = procesar_maestro_navasoft(df_term)
            st.session_state['nav'] = res
            if res['dansey'][1] == 0 and res['san luis'][1] == 0:
                st.warning("No se encontraron ítems con pedido mayor a 0 para Dansey o San Luis.")
            else:
                st.success("¡Plantillas generadas! Ya puedes descargarlas abajo.")
        except Exception as e:
            st.error(f"Error al procesar las plantillas: {e}")

    if 'nav' in st.session_state:
        st.divider()
        col_d, col_s = st.columns(2)
        for col, clave, titulo, archivo in [
            (col_d, 'dansey', "🏢 Almacén Dansey (08)", "PLANTILLA_NAVASOFT_DANSEY_08.xls"),
            (col_s, 'san luis', "🏢 Almacén San Luis (01)", "PLANTILLA_NAVASOFT_SANLUIS_01.xls"),
        ]:
            datos, n = st.session_state['nav'][clave]
            with col:
                st.markdown(f"### {titulo}")
                st.metric("Ítems a pedir", n)
                if datos:
                    st.download_button(
                        label="📥 Descargar",
                        data=datos,
                        file_name=archivo,
                        mime="application/vnd.ms-excel",
                        key=f"dl_{clave}"
                    )
                else:
                    st.info("No hay ítems para este almacén.")
