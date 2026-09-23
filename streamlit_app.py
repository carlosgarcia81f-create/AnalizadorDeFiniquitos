import streamlit as st
import pandas as pd
import numpy as np
import io  # Importante para manejar la exportación a Excel en memoria
from st_aggrid import AgGrid, GridOptionsBuilder
#-------------F U N C I O N E S  D E  A P O Y O ------------------------------------------------------------------------#

# ------------ C O N F I G U R A C I Ó N  D E  P Á G I N A -------------------------------------------------------------#
# 1. Título
st.title("Analizador de Finiquitos")
# 2. Controles laterales (lo que eran tus variables de @param)
filas_a_saltar = st.sidebar.number_input("Filas a saltar", value=11)
nombre_hoja = st.sidebar.text_input("Nombre de la hoja", value="11")
umbral_exceso = st.sidebar.number_input("Umbral de exceso respecto a contrato en %", value=30)
# NUEVO: Filtro de materialidad económica
umbral_monetario = st.sidebar.number_input("Monto mínimo de exceso ($)", value=20000.0)
porcentaje_pareto = st.sidebar.number_input("% Pareto", value=80)

uploaded_file = st.file_uploader("Sube tu archivo (.xlsm)", type=["xlsm"])

# Definir display como un alias de st.write para que no marque error
display = st.write

# 3. Configuración para anchos de columnas y tablas
st.set_page_config(layout="wide")

# M Ó D U L O 1
#---------------- P R O C E S A M I E N T O,  L I M P I E Z A  Y  A N Á L I S I S --------------------------------------#
# 1. Lectura de archivo
if uploaded_file is not None:
    df_finiquito = pd.read_excel(uploaded_file, sheet_name=nombre_hoja, skiprows=int(filas_a_saltar), engine='openpyxl')
       
    ## Controles para ver que funcione una vez cargado el archivo bien ## 
    # DIAGNÓSTICO:
    #st.write(f"Leyendo hoja: {nombre_hoja} saltando {filas_a_saltar} filas")
    #st.write("Primeras 5 filas detectadas:")
    #st.header("Vista previa de los datos leídos:")
    #st.dataframe(df_finiquito.head()) # Esto te mostrará si los títulos están en su lugar
    #st.write("Columnas detectadas:", df_finiquito.columns.tolist())

    # Esto elimina espacios, saltos de línea y tabuladores en los títulos de las columnas
    df_finiquito.columns = [str(c).strip() for c in df_finiquito.columns]
    # Forzar la columna Clave a Texto
    if 'Clave' in df_finiquito.columns:
        df_finiquito['Clave'] = df_finiquito['Clave'].astype(str).replace('nan', '')
    
    #---------------------- 2. Renombramos columnas----------------------------------------------------------------------------------------------------
    df_finiquito = df_finiquito.rename(columns={
        'Precio Unitario/Costo': 'PU',
        'Importe Contratado': 'Monto_Contratado',
        'Importe total estimado': 'Monto_Ejecutado',
        'Cantidad total estimada' : 'Cantidad_Ejecutada'
    })
    
    #---------------------- 3. Identificamos Nombres de Partidas y Subpartidas basado en la columna NIVEL-----------------------------------------------
    # Si la columna 'NIVEL' tiene un 1, es Partida Principal
    df_finiquito.loc[df_finiquito['NIVEL'] == 1, 'Partida_Principal'] = df_finiquito['Concepto']
    
    # Si la columna 'NIVEL' tiene un 2, es Subpartida
    df_finiquito.loc[df_finiquito['NIVEL'] == 2, 'Subpartida'] = df_finiquito['Concepto']
    
    #---------------------- 4. Aplicamos el relleno (Forward Fill)--------------------------------------------------------------------------------------
    # 1. Rellenamos primero las Partidas Principales (Nivel 1)
    df_finiquito['Partida_Principal'] = df_finiquito['Partida_Principal'].ffill()
    
    # 2. Si una fila es una nueva Partida Principal (Nivel 1),
    # forzamos que la Subpartida sea un texto vacío o "Sin Subpartida" para que no arrastre la anterior.
    df_finiquito.loc[df_finiquito['NIVEL'] == 1, 'Subpartida'] = "N/A"
    # Al haber puesto "N/A" en el inicio de cada partida, el ffill solo arrastrará dentro de su propia sección.
    df_finiquito['Subpartida'] = df_finiquito['Subpartida'].ffill()
    
    #------------------- 5. Limpieza -------------------------------------------------------------------------------------------------------------------
    #1.Nos quedamos solo con los renglones que SON conceptos (los que no tienen 1 ni 2)
    #Generalmente los conceptos no tienen marca en esa columna o tienen un 0
    df_finiquito_auditoria = df_finiquito[df_finiquito['PU'].notna()].copy()
    
    # 2. Borramos las filas que digan TOTAL, SUMA, SUBTOTAL en la columna Precio Unitario
    # El parámetro 'case=False' ignora si está en mayúsculas o minúsculas
    # Se añade una comprobación para asegurar que 'PU' es de tipo string antes de usar .str
    if not df_finiquito_auditoria['PU'].empty:
        df_finiquito_auditoria = df_finiquito_auditoria[
            ~df_finiquito_auditoria['PU'].astype(str).str.contains('TOTAL|SUMA|SUBTOTAL|RESUMEN', case=False, na=False)
        ]
    
    # NUEVA LIMPIEZA: Borramos filas donde 'Concepto' es nulo o 'N/A'
    df_finiquito_auditoria = df_finiquito_auditoria[
        df_finiquito_auditoria['Concepto'].notna() &
        (df_finiquito_auditoria['Concepto'] != 'N/A')
    ].copy()
    
    #---------------------- 6. Visualización del resultado-----------------------------------------------------------------------------------------------
    #Solo para efectos de diagnóstico
    #df_finiquito_auditoria[['Partida_Principal', 'Subpartida', 'Clave', 'Concepto', 'Monto_Contratado','Monto_Ejecutado']]
    
    #----------------------7. Calculamos la variación porcentual e impacto económico-------------------------
    porcentajeRespectoContrato = umbral_exceso / 100
    
    # NUEVO: Limpiar símbolos de moneda/comas y forzar formato numérico
    for col in ['Monto_Contratado', 'Monto_Ejecutado']:
        if df_finiquito_auditoria[col].dtype == 'object':
            df_finiquito_auditoria[col] = df_finiquito_auditoria[col].astype(str).str.replace(r'[\$,\s]', '', regex=True)
        df_finiquito_auditoria[col] = pd.to_numeric(df_finiquito_auditoria[col], errors='coerce').fillna(0)
    
    # Para evitar división entre cero si algún monto contratado viene vacío o en cero
    df_finiquito_auditoria['Monto_Contratado'] = df_finiquito_auditoria['Monto_Contratado'].replace(0, np.nan)
    
    # 1. Calcular variación porcentual y diferencia absoluta en pesos (Línea 99 actual)
    df_finiquito_auditoria['Variacion_Pct'] = (df_finiquito_auditoria['Monto_Ejecutado'] - df_finiquito_auditoria['Monto_Contratado']) / df_finiquito_auditoria['Monto_Contratado']
    df_finiquito_auditoria['Diferencia_Absoluta'] = df_finiquito_auditoria['Monto_Ejecutado'] - df_finiquito_auditoria['Monto_Contratado']
    
    # Revertir los NaN a 0 para que no causen problemas en la visualización posterior
    df_finiquito_auditoria['Monto_Contratado'] = df_finiquito_auditoria['Monto_Contratado'].fillna(0)
    df_finiquito_auditoria['Variacion_Pct'] = df_finiquito_auditoria['Variacion_Pct'].fillna(0)
    
    # 2. Formato para visualización
    df_finiquito_auditoria['Variacion_Pct_%'] = df_finiquito_auditoria['Variacion_Pct'].apply(lambda x: f'{x:.2%}')
    
    # 3. FILTRO DOBLE: Supera el porcentaje Y supera el monto mínimo económico
    excesos = df_finiquito_auditoria[
        (df_finiquito_auditoria['Variacion_Pct'] > porcentajeRespectoContrato) & 
        (df_finiquito_auditoria['Diferencia_Absoluta'] > umbral_monetario)
    ]
    
    # 4. Ordenar de mayor a menor impacto económico
    excesos = excesos.sort_values(by='Diferencia_Absoluta', ascending=False)
    
    st.write(f"Se encontraron {len(excesos)} conceptos atípicos (>{umbral_exceso}% y >${umbral_monetario:,.2f} de exceso)")
    
    # Mostrar la tabla incluyendo la Diferencia Absoluta para dar contexto
    display(excesos[['Clave', 'Partida_Principal', 'Subpartida', 'Concepto', 'Unidad', 'Monto_Contratado', 'Monto_Ejecutado', 'Diferencia_Absoluta', 'Variacion_Pct_%']].style.format({
        'Monto_Contratado': '${:,.2f}', 
        'Monto_Ejecutado': '${:,.2f}',
        'Diferencia_Absoluta': '${:,.2f}'
    }).background_gradient(subset=['Diferencia_Absoluta'], cmap='Reds'))
    
    #---------------------- 8. RESUMEN EJECUTIVO (CORREGIDO) ---------------------------------------------------------------------------------------------
    resumen_ejecutivo = df_finiquito_auditoria.groupby(['Partida_Principal', 'Subpartida']).agg({
        'Monto_Contratado': 'sum',
        'Monto_Ejecutado': 'sum'
    }).reset_index()
    
    # 1. Recalculate Diferencia_Absoluta
    resumen_ejecutivo['Diferencia_Absoluta'] = resumen_ejecutivo['Monto_Ejecutado'] - resumen_ejecutivo['Monto_Contratado']
    
    # 2. Initialize %_Variacion_Global with a default value (0)
    # Ensure the column is explicitly float type from the start
    resumen_ejecutivo['%_Variacion_Global'] = 0.0
    
    # 3. Handle cases where Monto_Contratado is 0 and Diferencia_Absoluta is also 0 (already 0 by initialization)
    
    # 4. Identify rows where Monto_Contratado is 0 but Diferencia_Absoluta is not 0
    condition_contracted_zero_diff_not_zero = (
        (resumen_ejecutivo['Monto_Contratado'] == 0) &
        (resumen_ejecutivo['Diferencia_Absoluta'] != 0)
    )
    resumen_ejecutivo.loc[condition_contracted_zero_diff_not_zero, '%_Variacion_Global'] = 100.0
    
    # 5. For all other rows (where Monto_Contratado is not 0), calculate normally
    condition_monto_contratado_not_zero = (resumen_ejecutivo['Monto_Contratado'] != 0)
    resumen_ejecutivo.loc[condition_monto_contratado_not_zero, '%_Variacion_Global'] = (
        (resumen_ejecutivo.loc[condition_monto_contratado_not_zero, 'Diferencia_Absoluta'] /
        resumen_ejecutivo.loc[condition_monto_contratado_not_zero, 'Monto_Contratado']) * 100
    ).astype(float) # Explicitly cast to float to prevent FutureWarning
    
    # Ordenamos
    resumen_ejecutivo = resumen_ejecutivo.sort_values(by='Diferencia_Absoluta', ascending=False)
    
    # Visualización con formato
    st.write("--- RESUMEN DE AUDITORÍA POR SUBPARTIDAS ---")
    display(resumen_ejecutivo.style.format({
        'Monto_Contratado': '${:,.2f}',
        'Monto_Ejecutado': '${:,.2f}',
        'Diferencia_Absoluta': '${:,.2f}',
        '%_Variacion_Global': '{:.2f}%'
    }).background_gradient(subset=['%_Variacion_Global'], cmap='YlOrRd'))
   #---------------------- 9. PLANEACIÓN DE INSPECCIÓN FÍSICA (PARETO POR CATEGORÍAS) ------------------------------------------

    # 1. Normalizar la columna Unidad (convierte a mayúsculas y quita espacios en los extremos)
    if 'Unidad' in df_finiquito_auditoria.columns:
        df_finiquito_auditoria['Unidad_Norm'] = df_finiquito_auditoria['Unidad'].astype(str).str.strip().str.upper()
    else:
        df_finiquito_auditoria['Unidad_Norm'] = 'N/A'
    
    # 2. Clasificar los conceptos con validación flexible
    # Lista exhaustiva para todas las variantes posibles de piezas
    variantes_piezas = ['PZA', 'PZA.', 'PZAS', 'PZAS.', 'PIEZA', 'PIEZAS', 'PZ', 'PZ.', 'P.']
    
    # Condiciones usando búsqueda exacta (para piezas) y Regex (para acarreos)
    condiciones_unidad = [
        df_finiquito_auditoria['Unidad_Norm'].isin(variantes_piezas),
        # El regex r'M3\s*[-\/]\s*KM' encuentra "M3-KM", "M3/KM", e incluso variaciones con espacios como "M3 / KM"
        df_finiquito_auditoria['Unidad_Norm'].str.contains(r'M3\s*[-\/]\s*KM', regex=True, na=False)
    ]
    
    elecciones_categoria = ['REVISIÓN GABINETE/CONTEO (Piezas)', 'REVISIÓN VOLUMÉTRICA (Acarreos)']
    df_finiquito_auditoria['Categoria_Analisis'] = np.select(condiciones_unidad, elecciones_categoria, default='INSPECCIÓN FÍSICA CAMPO (Visibles)')
    
    # 3. Aplicar Pareto separado por cada categoría
    dfs_pareto = []
    threshold_alta = porcentaje_pareto
    threshold_media = threshold_alta + 5
    
    for categoria in df_finiquito_auditoria['Categoria_Analisis'].unique():
        df_cat = df_finiquito_auditoria[df_finiquito_auditoria['Categoria_Analisis'] == categoria].copy()
        
        if not df_cat.empty:
            # Ordenamos de mayor a menor importe dentro de la categoría
            df_cat = df_cat.sort_values(by='Monto_Ejecutado', ascending=False)
            
            # El peso y el acumulado se calculan respecto al total de su propia categoría
            total_cat = df_cat['Monto_Ejecutado'].sum()
            if total_cat > 0:
                df_cat['%_Peso'] = (df_cat['Monto_Ejecutado'] / total_cat) * 100
            else:
                df_cat['%_Peso'] = 0.0
                
            df_cat['%_Acumulado'] = df_cat['%_Peso'].cumsum()
        
            # NUEVA LÓGICA: Calculamos el acumulado de la fila anterior. 
            # Esto asegura que el concepto que "rompe" la barrera del 80% sí se incluya en ALTA.
            acumulado_anterior = df_cat['%_Acumulado'] - df_cat['%_Peso']
            
            # Asignación de prioridades corregida
            condiciones_pareto = [
                (acumulado_anterior < threshold_alta),
                (acumulado_anterior < threshold_media)
            ]
            elecciones_pareto = ['ALTA', 'MEDIA']
            df_cat['Prioridad'] = np.select(condiciones_pareto, elecciones_pareto, default='BAJA')
            
            dfs_pareto.append(df_cat)
    
    # 4. Unir y filtrar los resultados finales
    df_plan_inspeccion = pd.concat(dfs_pareto)
    
    # Nos quedamos con la prioridad ALTA de las tres categorías
    df_plan_inspeccion_filtrado = df_plan_inspeccion[
        (df_plan_inspeccion['Prioridad'] == 'ALTA') & 
        (df_plan_inspeccion['Monto_Ejecutado'] > 0)
    ].copy()
    
    st.write(f"\n--- ESTRATEGIA DE INSPECCIÓN FÍSICA SEPARADA (PARETO {threshold_alta}%) ---")

    # Cuadro explicativo de la metodología para el usuario
    st.info(
        "💡 **Nota Metodológica:** Para evitar que los acarreos o suministros sesguen la muestra, "
        "el análisis de Pareto se calculó de forma **independiente** (cada grupo suma su propio 100%):\n"
        "* **Visibles:** Conceptos revisables físicamente en campo (pisos, banquetas, etc.).\n"
        "* **Acarreos (m3-km):** Revisión basada en volumetría y generadores.\n"
        "* **Piezas (pza):** Revisión de gabinete o conteo de inventario."
    )
    
    # --- NUEVOS CÁLCULOS DE REPRESENTATIVIDAD ---
    monto_total_obra = df_finiquito_auditoria['Monto_Ejecutado'].sum()
    monto_revisar_total = df_plan_inspeccion_filtrado['Monto_Ejecutado'].sum()
    pct_revisar_total = (monto_revisar_total / monto_total_obra) * 100 if monto_total_obra > 0 else 0
    
    st.write(f"Monto Total Ejecutado de la Obra: **${monto_total_obra:,.2f}**")
    st.write(f"Monto Total a Revisar (Suma de Prioridades ALTA): **${monto_revisar_total:,.2f} ({pct_revisar_total:.2f}% de la obra)**")
    st.write("-" * 50)
    
    # Crear pestañas en Streamlit
    tab1, tab2, tab3 = st.tabs(["🏗️ Visibles en Campo", "🚚 Acarreos (Volumetría)", "📦 Piezas (Gabinete)"])
    
    # Función auxiliar para dar formato a las sub-tablas
    def formato_tabla_pareto(df_sub):
        return df_sub[['Partida_Principal', 'Clave', 'Concepto', 'Cantidad_Ejecutada', 'Monto_Ejecutado', '%_Peso', '%_Acumulado']].style.format({
            '%_Peso': '{:.2f}%',
            'Monto_Ejecutado': '${:,.2f}',
            'Cantidad_Ejecutada': '{:,.2f}',
            '%_Acumulado': '{:.2f}%'
        }).background_gradient(subset=['%_Acumulado'], cmap='Blues')
    
    with tab1:
        df_visibles = df_plan_inspeccion_filtrado[df_plan_inspeccion_filtrado['Categoria_Analisis'] == 'INSPECCIÓN FÍSICA CAMPO (Visibles)']
        monto_visibles = df_visibles['Monto_Ejecutado'].sum()
        pct_visibles = (monto_visibles / monto_total_obra) * 100 if monto_total_obra > 0 else 0
        
        st.write(f"**{len(df_visibles)} conceptos** prioritarios para revisión física.")
        st.success(f"💰 Representan **${monto_visibles:,.2f}** (El **{pct_visibles:.2f}%** del total de la obra)")
        display(formato_tabla_pareto(df_visibles))
    
    with tab2:
        df_acarreos = df_plan_inspeccion_filtrado[df_plan_inspeccion_filtrado['Categoria_Analisis'] == 'REVISIÓN VOLUMÉTRICA (Acarreos)']
        monto_acarreos = df_acarreos['Monto_Ejecutado'].sum()
        pct_acarreos = (monto_acarreos / monto_total_obra) * 100 if monto_total_obra > 0 else 0
        
        st.write(f"**{len(df_acarreos)} conceptos** prioritarios para revisión de generadores/topografía.")
        st.success(f"💰 Representan **${monto_acarreos:,.2f}** (El **{pct_acarreos:.2f}%** del total de la obra)")
        display(formato_tabla_pareto(df_acarreos))
    
    with tab3:
        df_piezas = df_plan_inspeccion_filtrado[df_plan_inspeccion_filtrado['Categoria_Analisis'] == 'REVISIÓN GABINETE/CONTEO (Piezas)']
        monto_piezas = df_piezas['Monto_Ejecutado'].sum()
        pct_piezas = (monto_piezas / monto_total_obra) * 100 if monto_total_obra > 0 else 0
        
        st.write(f"**{len(df_piezas)} conceptos** prioritarios para conteo físico o revisión de facturas.")
        st.success(f"💰 Representan **${monto_piezas:,.2f}** (El **{pct_piezas:.2f}%** del total de la obra)")
        display(formato_tabla_pareto(df_piezas))
    ####------DEPURACIÓN FINAL PARA INSPECCIÓN FÍSICA (INTERACTIVA POR EL USUARIO)-----#####
    st.write("---")
    st.subheader("🛠️ Depuración Final para Inspección Física")
    
    # 1. Hacemos una copia de todos los conceptos de prioridad ALTA
    df_prioritarios = df_plan_inspeccion_filtrado.copy()
    
    # 2. Agregamos la columna de control (True por defecto)
    df_prioritarios['Inspeccion_Fisica'] = True
    
    # 3. Mostrar la tabla como un editor interactivo
    st.write("Desmarca los conceptos cuyo cálculo dependa de otro (ej. rellenos, cimbras). Estos se conservarán etiquetados para revisión de gabinete:")
    df_editado = st.data_editor(
        df_prioritarios[['Inspeccion_Fisica', 'Clave', 'Concepto', 'Unidad', 'Categoria_Analisis', 'Monto_Ejecutado']],
        column_config={
            "Inspeccion_Fisica": st.column_config.CheckboxColumn(
                "¿Inspección Física?",
                help="Desmarca si el volumen se deduce de otro concepto y solo requiere cálculo en gabinete",
                default=True,
            )
        },
        disabled=["Clave", "Concepto", "Unidad", "Categoria_Analisis", "Monto_Ejecutado"], 
        hide_index=True,
        use_container_width=True
    )
    
    # 4. Clasificar el tipo de revisión en el dataframe final que se va a descargar
    df_final_descarga = df_prioritarios.copy()
    # Actualizamos la columna con los cambios que hizo el usuario en pantalla
    df_final_descarga['Inspeccion_Fisica'] = df_editado['Inspeccion_Fisica'] 
    # Creamos la etiqueta clara para Excel
    df_final_descarga['Estrategia_Revision'] = df_final_descarga['Inspeccion_Fisica'].apply(
        lambda x: 'Campo (Medición Directa)' if x else 'Gabinete (Cálculo Dependiente)'
    )
    #---------------------------------- M O D U L O 2 ----------------------------------------------------------#
    #-------------D E S C A R G A  D E  A R C H I V O  A  E X C E L---------------------------------------------#
    
    # --- 1. LISTADO ORIGINAL COMPLETO (Formato intacto con títulos) ---
    # Usamos df_finiquito, que contiene TODAS las filas originales (incluyendo títulos de partidas sin PU)
    df_original_marcado = df_finiquito.copy()
    
    # Mapeamos la Prioridad y la Categoría de todos los conceptos analizados
    df_original_marcado['Prioridad'] = df_plan_inspeccion['Prioridad']
    df_original_marcado['Categoria_Analisis'] = df_plan_inspeccion['Categoria_Analisis'] # <--- AGREGA ESTA LÍNEA
    
    # Mapeamos el Peso y Acumulado ÚNICAMENTE de la tabla filtrada final (para que coincidan exacto con la app)
    df_original_marcado['%_Peso'] = df_plan_inspeccion_filtrado['%_Peso']
    df_original_marcado['%_Acumulado'] = df_plan_inspeccion_filtrado['%_Acumulado']
    
    # Definimos las columnas a exportar
    columnas_disponibles_raw = df_original_marcado.columns.tolist()
    cols_interes_original = ['Clave', 'Concepto','PU', 'Cantidad_Ejecutada','Monto_Ejecutado', 'Partida_Principal', 'Subpartida', '%_Peso', '%_Acumulado', 'Prioridad','Categoria_Analisis']
    
    if 'Unidad' in columnas_disponibles_raw:
        cols_interes_original.insert(2, 'Unidad')
    
    df_listado_completo = df_original_marcado[cols_interes_original].copy()
    
        # --- 2. LISTADO FILTRADO PARETO (Prioridad, % Peso y % Acumulado) ---
    columnas_disponibles_filtrado = df_plan_inspeccion_filtrado.columns.tolist()
    cols_interes_resumen_prioridades = ['Clave', 'Concepto','PU', 'Cantidad_Ejecutada','Monto_Ejecutado', 'Partida_Principal', 'Subpartida', '%_Peso', '%_Acumulado','Prioridad','Categoria_Analisis']
    
    if 'Unidad' in columnas_disponibles_filtrado:
        cols_interes_resumen_prioridades.insert(2, 'Unidad')
    
    df_resumen_final = df_plan_inspeccion_filtrado[cols_interes_resumen_prioridades].copy()
    
    
    # --- 3. TABLA DE EXCESOS ---
    cols_interes_excesos = ['Clave', 'Concepto','PU', 'Monto_Contratado', 'Cantidad_Ejecutada','Monto_Ejecutado', 'Partida_Principal', 'Subpartida', 'Variacion_Pct']
    
    if 'Unidad' in columnas_disponibles_filtrado:
        cols_interes_excesos.insert(2, 'Unidad')
    
    df_excesos = df_plan_inspeccion_filtrado[cols_interes_excesos].copy()
    
    
    # --- ESCRITURA EN EXCEL ---
    # 1. Crear el objeto en memoria
    buffer_excel = io.BytesIO()
    
    # 2. Iniciar el Writer
    with pd.ExcelWriter(buffer_excel, engine='xlsxwriter') as writer:
        
        # Agregar todas las tablas al diccionario de hojas
        hojas = {
            'Conceptos_sobre_Umbral': df_excesos,
            'Var_Por_Partidas': resumen_ejecutivo,
            'Resumen_Prioridades': df_resumen_final,
            'Listado_Completo': df_listado_completo  # Nueva pestaña agregada
        }
    
        workbook = writer.book
        
        # --- DEFINICIÓN DE FORMATOS ---
        header_format = workbook.add_format({
            'bold': True, 'text_wrap': True, 'font_color': 'white',
            'valign': 'vcenter', 'align': 'center', 'bg_color': '#FF5E12', 'border': 1
        })
    
        # Formato para MONTO (Moneda: $ #,##0.00)
        money_format = workbook.add_format({
            'num_format': '"$"#,##0.00',
            'text_wrap': True,
            'valign': 'center'
        })
    
        # Formato para CANTIDADES (Miles con decimales: #,##0.00)
        number_format = workbook.add_format({
            'num_format': '#,##0.00',
            'text_wrap': True,
            'valign': 'center'
        })
    
        # Formato para Porcentajes (0.00%)
        percent_format = workbook.add_format({
            'num_format': '0.00%',
            'text_wrap': True,
            'valign': 'center'
        })
        
        # Formato base para TEXTO (Conceptos)
        body_format = workbook.add_format({
            'text_wrap': True,
            'valign': 'top'
        })
    
        # --- PROCESO POR HOJA ---
        for nombre_hoja, df in hojas.items():
            if not df.empty:
                df.to_excel(writer, sheet_name=nombre_hoja, index=False, startrow=1, header=False)
                worksheet = writer.sheets[nombre_hoja]
    
                # Aplicar encabezados
                for col_num, value in enumerate(df.columns.values):
                    worksheet.write(0, col_num, value, header_format)
    
                # --- APLICAR ANCHOS Y FORMATOS DE CELDA ---
                for i, col in enumerate(df.columns):
                    # Determinar el ancho
                    if col in ['Concepto', 'Partida_Principal']:
                        ancho = 45
                        formato_celda = body_format
                    elif col == 'Clave':
                        ancho = 12
                        formato_celda = body_format
                    elif col == 'Variacion_Pct':
                        ancho = 12
                        formato_celda = percent_format    
                    # Si la columna es de DINERO (Monto, Importe, PU, Diferencia)
                    elif any(x in col for x in ['Monto', 'Importe', 'PU', 'Diferencia']):
                        ancho = 14
                        formato_celda = money_format
                    # Si la columna es de CANTIDAD (Cantidad, Volumen)
                    elif any(x in col for x in ['Cantidad', 'Volumen']):
                        ancho = 9
                        formato_celda = number_format
                    else:
                        ancho = 14
                        formato_celda = body_format
                    
                    # Aplicar a toda la columna (desde la fila 1 hasta la 1048576)
                    worksheet.set_column(i, i, ancho, formato_celda)
            else:
                pd.DataFrame(["No hay datos."]).to_excel(writer, sheet_name=nombre_hoja, index=False, header=False)
    
    # 3. Botón de descarga
    st.download_button(
        label="📥 Descargar Reporte de Auditoría Final",
        data=buffer_excel.getvalue(),
        file_name="Reporte_Final_Auditoria.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
