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
    
    #----------------------7. Calculamos la variación porcentual de cada concepto respecto de lo contratado----------------------------------------------
    #Esto nos ayuda a ver cuales conceptos rebasaron más el importe contratado
    porcentajeRespectoContrato = umbral_exceso/100 # Usamos la variable de configuración
    df_finiquito_auditoria['Variacion_Pct'] = (df_finiquito_auditoria['Monto_Ejecutado'] - df_finiquito_auditoria['Monto_Contratado']) / df_finiquito_auditoria['Monto_Contratado']
    # Create a new column for the formatted percentage for display purposes
    df_finiquito_auditoria['Variacion_Pct_%'] = df_finiquito_auditoria['Variacion_Pct'].apply(lambda x: f'{x:.2%}')
    # Filtramos los que superan el porcentaje señalado (using the numeric Variacion_Pct)
    excesos = df_finiquito_auditoria[df_finiquito_auditoria['Variacion_Pct'] > porcentajeRespectoContrato]
    
    st.write(f"Se encontraron {len(excesos)} conceptos con un porcentaje de {porcentajeRespectoContrato*100}% superior respecto del porcentaje contratado")
    display(excesos[['Clave', 'Partida_Principal', 'Subpartida', 'Concepto', 'Unidad', 'Monto_Contratado', 'Monto_Ejecutado', 'Variacion_Pct_%']])
    
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
        
        # NUEVO: Agregamos una columna booleana para los checkboxes
        df_plan_inspeccion_filtrado.insert(0, 'Seleccionado', True)
        
        # --- CÁLCULOS BASE ---
        monto_total_obra = df_finiquito_auditoria['Monto_Ejecutado'].sum()
        
        st.write(f"\n--- ESTRATEGIA DE INSPECCIÓN FÍSICA SEPARADA (PARETO {threshold_alta}%) ---")
        st.info(
            "💡 **Nota Metodológica:** El análisis se calculó de forma independiente. "
            "Usa las tablas inferiores para desmarcar los conceptos que por cuestiones técnicas o físicas no revisarás. "
            "El dashboard se actualizará automáticamente."
        )
        
        # --- CREACIÓN DEL DASHBOARD INTERACTIVO ---
        st.markdown("### 📊 Dashboard de Selección de Auditoría")
        # Usamos un contenedor vacío para poder imprimir los resultados DESPUÉS de que el usuario interactúe con las tablas
        dashboard_placeholder = st.empty() 
        
        # Crear pestañas en Streamlit
        tab1, tab2, tab3 = st.tabs(["🏗️ Visibles en Campo", "🚚 Acarreos (Volumetría)", "📦 Piezas (Gabinete)"])
        
        # Función auxiliar para crear tablas editables (Checkboxes)
        def mostrar_editor_interactivo(df_sub, key):
            if df_sub.empty:
                st.warning("No hay conceptos en esta categoría que cumplan el criterio.")
                return df_sub
                
            # Bloqueamos todas las columnas excepto el Checkbox para que no modifiquen montos por error
            columnas_bloqueadas = df_sub.columns.drop('Seleccionado').tolist()
            
            edited_df = st.data_editor(
                df_sub,
                column_config={
                    "Seleccionado": st.column_config.CheckboxColumn("Revisar ✅", default=True),
                    "Monto_Ejecutado": st.column_config.NumberColumn("Monto", format="$ %.2f"),
                    "Cantidad_Ejecutada": st.column_config.NumberColumn("Cantidad", format="%.2f"),
                    "%_Peso": st.column_config.NumberColumn("% Peso", format="%.2f %%"),
                    "%_Acumulado": st.column_config.NumberColumn("% Acumulado", format="%.2f %%"),
                    "Categoria_Analisis": None, # Ocultamos la categoría porque ya está en la pestaña
                    "Prioridad": None # Ocultamos para ahorrar espacio
                },
                disabled=columnas_bloqueadas,
                hide_index=True,
                use_container_width=True,
                key=key
            )
            return edited_df
        
        # Desplegar tablas interactivas
        with tab1:
            df_visibles = df_plan_inspeccion_filtrado[df_plan_inspeccion_filtrado['Categoria_Analisis'] == 'INSPECCIÓN FÍSICA CAMPO (Visibles)']
            edited_visibles = mostrar_editor_interactivo(df_visibles, "editor_visibles")
        
        with tab2:
            df_acarreos = df_plan_inspeccion_filtrado[df_plan_inspeccion_filtrado['Categoria_Analisis'] == 'REVISIÓN VOLUMÉTRICA (Acarreos)']
            edited_acarreos = mostrar_editor_interactivo(df_acarreos, "editor_acarreos")
        
        with tab3:
            df_piezas = df_plan_inspeccion_filtrado[df_plan_inspeccion_filtrado['Categoria_Analisis'] == 'REVISIÓN GABINETE/CONTEO (Piezas)']
            edited_piezas = mostrar_editor_interactivo(df_piezas, "editor_piezas")
        
        # --- ACTUALIZACIÓN DE VARIABLES POST-SELECCIÓN ---
        # Filtramos para quedarnos solo con lo que el usuario DEJÓ marcado
        sel_visibles = edited_visibles[edited_visibles['Seleccionado']]
        sel_acarreos = edited_acarreos[edited_acarreos['Seleccionado']]
        sel_piezas = edited_piezas[edited_piezas['Seleccionado']]
        
        monto_sel_visibles = sel_visibles['Monto_Ejecutado'].sum()
        monto_sel_acarreos = sel_acarreos['Monto_Ejecutado'].sum()
        monto_sel_piezas = sel_piezas['Monto_Ejecutado'].sum()
        
        monto_revisar_total = monto_sel_visibles + monto_sel_acarreos + monto_sel_piezas
        pct_revisar_total = (monto_revisar_total / monto_total_obra) * 100 if monto_total_obra > 0 else 0
        
        # Llenar el Dashboard en la parte superior con la info actualizada
        with dashboard_placeholder.container():
            cols = st.columns(4)
            # Mostramos los indicadores. delta_color="off" pone el subtítulo en gris (neutral).
            cols[0].metric("💰 Total de la Obra", f"${monto_total_obra:,.2f}")
            cols[1].metric("🎯 Total a Revisar", f"${monto_revisar_total:,.2f}", f"{pct_revisar_total:.2f}% de la obra", delta_color="off")
            cols[2].metric("🏗️ Visibles (Campo)", f"${monto_sel_visibles:,.2f}", f"{(monto_sel_visibles/monto_total_obra)*100:.2f}%", delta_color="off")
            cols[3].metric("📦 Acarreos + Piezas", f"${(monto_sel_acarreos + monto_sel_piezas):,.2f}", f"{((monto_sel_acarreos + monto_sel_piezas)/monto_total_obra)*100:.2f}%", delta_color="off")
            
            # Barra de progreso visual
            progreso_normalizado = min(pct_revisar_total / 100.0, 1.0) # Evita que pase del 100% y lance error
            st.progress(progreso_normalizado, text=f"Porcentaje de auditoría cubierto en el plan de inspección: {pct_revisar_total:.2f}%")
            st.write("-" * 50)
        
        # 5. SOBRESCRIBIR EL DATAFRAME FINAL
        # Al juntar solo las selecciones, garantizamos que el Módulo 2 de Excel exporte ÚNICAMENTE lo que elegiste.
        df_plan_inspeccion_filtrado = pd.concat([sel_visibles, sel_acarreos, sel_piezas])

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
