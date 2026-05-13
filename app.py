import streamlit as st
import pandas as pd
from st_aggrid import AgGrid, GridOptionsBuilder

# ... (Toda tu lógica de limpieza y niveles de AnalizadorDeFiniquito.ipynb) ...

st.title("Asistente de Auditoría de Obra")

# 1. Carga de datos
uploaded_file = st.file_uploader("Sube tu archivo de finiquito", type=["xlsm"])

if uploaded_file:
    # 2. Mostrar propuesta de Pareto
    st.subheader("Propuesta de Inspección Física")
    
    # Configurar tabla interactiva para reordenar
    gb = GridOptionsBuilder.from_dataframe(df_plan_inspeccion)
    gb.configure_row_drag(True) # Activa el arrastre de filas
    gridOptions = gb.build()
    
    response = AgGrid(df_plan_inspeccion, gridOptions=gridOptions)
    
    # 3. Recalcular según el nuevo orden
    df_usuario = pd.DataFrame(response['data'])
    nuevo_acumulado = df_usuario['Monto_Ejecutado'].cumsum() / total_finiquito
    
    # 4. Botón de exportación
    st.download_button("Exportar Propuesta a Excel", data=buffer_excel, file_name="Propuesta_Auditoria.xlsx")
