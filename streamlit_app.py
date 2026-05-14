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
nombre_hoja = st.sidebar.text_input("Nombre de la hoja", value="12")
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
    display(excesos[['Clave', 'Partida_Principal', 'Subpartida', 'Concepto', 'Monto_Contratado','Monto_Ejecutado','Variacion_Pct_%']])
    
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
    #---------------------- 9. PLANEACIÓN DE INSPECCIÓN FÍSICA (ANÁLISIS DE PARETO CON CONFIGURACIÓN) ------------------------------------------
    
    # 1. Ordenamos de mayor a menor importancia económica
    df_plan_inspeccion = df_finiquito_auditoria.sort_values(by='Monto_Ejecutado', ascending=False).copy()
    
    # 2. Calculamos el peso de cada concepto y su acumulado
    total_finiquito = df_plan_inspeccion['Monto_Ejecutado'].sum()
    df_plan_inspeccion['%_Peso'] = (df_plan_inspeccion['Monto_Ejecutado'] / total_finiquito) * 100
    df_plan_inspeccion['%_Acumulado'] = df_plan_inspeccion['%_Peso'].cumsum()
    
    # 3. Definimos quiénes son del Grupo A (Prioridad Alta) usando el porcentaje configurable
    #    El grupo B será un 5% adicional, y el resto el Grupo C.
    threshold_alta = porcentaje_pareto # Usamos la variable de configuración (ej. 90%)
    threshold_media = threshold_alta + 5 # Por ejemplo, 95%
    
    df_plan_inspeccion['Prioridad'] = df_plan_inspeccion['%_Acumulado'].apply(
        lambda x: f'ALTA (Grupo A - cubre el {threshold_alta}%)' if x <= threshold_alta
        else (f'MEDIA (Grupo B - cubre hasta el {threshold_media}%)' if x <= threshold_media else 'BAJA (Grupo C)')
    )
    
    # 4. Filtramos la lista para tu bitácora de campo (solo los de alta prioridad)
    # Se añade una comprobación para evitar el AttributeError si el DataFrame está vacío
    if not df_plan_inspeccion.empty:
        df_plan_inspeccion['Prioridad'] = df_plan_inspeccion['Prioridad'].astype(str)
        lista_campo = df_plan_inspeccion[df_plan_inspeccion['Prioridad'].str.startswith('ALTA')]
    else:
        st.write("Advertencia: df_plan_inspeccion está vacío. No se pueden generar elementos para la lista de campo.")
        lista_campo = pd.DataFrame(columns=['Partida_Principal', 'Subpartida', 'Clave', 'Concepto', 'Monto_Ejecutado', '%_Peso', '%_Acumulado', 'Prioridad']) # Crear un DataFrame vacío con las columnas esperadas
    
    st.write(f"\n--- ESTRATEGIA DE INSPECCIÓN FÍSICA (ANÁLISIS DE PARETO {threshold_alta}/{100-threshold_alta}) ---")
    st.write(f"Total de conceptos en la obra: {len(df_plan_inspeccion)}")
    st.write(f"Conceptos críticos a revisar para cubrir el {threshold_alta}% del monto: {len(lista_campo)}")
    st.write("-" * 50)
    
    # Filtramos los conceptos con Monto_Ejecutado > 0 (y por lo tanto %_Peso > 0)
    df_plan_inspeccion_filtrado = df_plan_inspeccion[df_plan_inspeccion['Monto_Ejecutado'] > 0].copy()
    display(df_plan_inspeccion_filtrado[['Prioridad','Partida_Principal', 'Subpartida', 'Clave', 'Concepto','Cantidad_Ejecutada', 'Monto_Ejecutado', '%_Peso', '%_Acumulado']].style.format({
        '%_Peso': '{:.2f}%',
        'Monto_Ejecutado': '${:,.2f}',
        'Cantidad_Ejecutada': '{:,.0f}',
        '%_Acumulado': '{:.2f}%'
    }).background_gradient(subset=['%_Acumulado'], cmap='Blues'))

#---------------------------------- M O D U L O 2 ----------------------------------------------------------#
#-------------D E S C A R G A  D E  A R C H I V O  A  E X C E L---------------------------------------------#
# --- LIMPIEZA DE COLUMNAS PARA EL REPORTE ---
    # Selecciona solo lo que el auditor necesita ver en campo
    cols_interes_resumen_prioridades = ['Clave', 'Concepto', 'Unidad', 'PU', 'Cantidad_Ejecutada','Monto_Ejecutado', 'Partida_Principal', 'Subpartida', '%_Peso', '%_Acumulado','Prioridad']
    df_resumen_final = df_plan_inspeccion_filtrado[cols_interes_resumen_prioridades].copy()
    
    cols_interes_excesos = ['Clave', 'Concepto', 'Unidad', 'PU', 'Monto_Contratado', 'Cantidad_Ejecutada','Monto_Ejecutado', 'Partida_Principal', 'Subpartida', 'Variacion_Pct',]
    df_excesos = df_plan_inspeccion_filtrado[cols_interes_excesos].copy()


    # 1. Crear el objeto en memoria
    buffer_excel = io.BytesIO()
    
    # 2. Iniciar el Writer
    with pd.ExcelWriter(buffer_excel, engine='xlsxwriter') as writer:
        
        hojas = {
            'Conceptos_sobre_Umbral': df_excesos,
            'Var_Por_Partidas': resumen_ejecutivo,
            'Resumen_Prioridades': df_resumen_final #Solo mostrará las columnas de interés
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
