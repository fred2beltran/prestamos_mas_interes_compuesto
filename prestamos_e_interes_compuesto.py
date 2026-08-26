import streamlit as st
import pandas as pd
import numpy_financial as npf
import io
from openpyxl.worksheet.table import Table, TableStyleInfo 
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# ==========================================
# CONFIGURACIÓN DE LA PÁGINA
# ==========================================
st.set_page_config(page_title="Simulador Financiero 360", layout="centered", page_icon="🏦")

# ==========================================
# 1. LÓGICA DE NEGOCIO: PRÉSTAMOS
# ==========================================
def calcular_amortizacion(monto, tasa_anual, anios, sistema="Francés"):
    meses = int(anios * 12)
    tasa_mensual = tasa_anual / 12
    saldo_restante = monto
    tabla = []
    
    if sistema == "Francés":
        cuota_mensual = npf.pmt(tasa_mensual, meses, -monto)
        for mes in range(1, meses + 1):
            pago_interes = saldo_restante * tasa_mensual
            pago_capital = cuota_mensual - pago_interes
            saldo_restante -= pago_capital
            tabla.append({
                "Mes": mes, "Cuota Mensual": round(cuota_mensual, 2),
                "Pago Interés": round(pago_interes, 2), "Pago Capital": round(pago_capital, 2),
                "Saldo Restante": round(max(0, saldo_restante), 2)
            })
            
    elif sistema == "Alemán":
        pago_capital_fijo = monto / meses
        for mes in range(1, meses + 1):
            pago_interes = saldo_restante * tasa_mensual
            cuota_mensual = pago_capital_fijo + pago_interes
            saldo_restante -= pago_capital_fijo
            tabla.append({
                "Mes": mes, "Cuota Mensual": round(cuota_mensual, 2),
                "Pago Interés": round(pago_interes, 2), "Pago Capital": round(pago_capital_fijo, 2),
                "Saldo Restante": round(max(0, saldo_restante), 2)
            })

    return pd.DataFrame(tabla)

# ==========================================
# 2. LÓGICA DE NEGOCIO: AHORRO COMPUESTO
# ==========================================
def calcular_ahorro_compuesto(capital_inicial, aporte_mensual, tasa_anual, anios):
    meses = int(anios * 12)
    tasa_mensual = tasa_anual / 12
    saldo = capital_inicial
    total_aportado = capital_inicial
    tabla = []

    for mes in range(1, meses + 1):
        interes_ganado = saldo * tasa_mensual
        saldo += interes_ganado + aporte_mensual
        total_aportado += aporte_mensual
        
        tabla.append({
            "Mes": mes,
            "Aporte Mensual": round(aporte_mensual, 2),
            "Interés del Mes": round(interes_ganado, 2),
            "Total Aportado": round(total_aportado, 2),
            "Saldo Final": round(saldo, 2)
        })

    return pd.DataFrame(tabla)

# ==========================================
# 3. GENERADORES DE EXCEL (ESTILO DARK)
# ==========================================
def aplicar_estilos_base_excel(writer, df, titulo_resumen):
    """Aplica el diseño oscuro, minimalista y retorna el worksheet configurado."""
    df.to_excel(writer, sheet_name="Reporte", index=False, startrow=7)
    worksheet = writer.sheets["Reporte"]
    
    fondo_principal = PatternFill(start_color="1A1A1A", end_color="1A1A1A", fill_type="solid")
    fuente_titulo = Font(name="Segoe UI", color="FFFFFF", bold=True, size=16)
    alineacion_centro = Alignment(horizontal="center", vertical="center")
    
    worksheet.sheet_view.showGridLines = False
    for row in worksheet.iter_rows(min_row=1, max_row=7, min_col=1, max_col=6):
        for cell in row:
            cell.fill = fondo_principal

    # Título Principal
    worksheet.merge_cells("A1:E1")
    worksheet["A1"] = titulo_resumen
    worksheet["A1"].font = fuente_titulo 
    worksheet["A1"].alignment = alineacion_centro
    
    # Créditos impresos en el archivo Excel
    worksheet.merge_cells("A2:E2")
    worksheet["A2"] = "Generado por Simulador Financiero 360 - Desarrollado por [TU NOMBRE AQUÍ]"
    worksheet["A2"].font = Font(name="Segoe UI", color="A6A6A6", italic=True, size=10)
    worksheet["A2"].alignment = Alignment(horizontal="right", vertical="center")
    
    return worksheet

def finalizar_excel(worksheet, ultima_fila):
    """Aplica tabla dinámica, anchos y estilos finales."""
    rango_tabla = f"A8:E{ultima_fila}" 
    tabla_excel = Table(displayName="TablaDatos", ref=rango_tabla)
    tabla_excel.tableStyleInfo = TableStyleInfo(name="TableStyleDark1", showRowStripes=True)
    worksheet.add_table(tabla_excel)
    
    formato_moneda = '$#,##0.00'
    for fila in worksheet.iter_rows(min_row=9, max_row=ultima_fila, min_col=1, max_col=5):
        fila[0].alignment = Alignment(horizontal="center")
        for celda in fila[1:]:
            celda.number_format = formato_moneda
            celda.alignment = Alignment(horizontal="right")
            
    for col in worksheet.columns:
        max_len = max(len(str(celda.value or '')) for celda in col)
        worksheet.column_dimensions[get_column_letter(col[0].column)].width = max(max_len + 4, 18)

def pintar_cajas_resumen(worksheet):
    """Colorea las celdas de las etiquetas de resumen en el Excel."""
    fondo_secundario = PatternFill(start_color="252525", end_color="252525", fill_type="solid")
    fuente_etiquetas = Font(name="Segoe UI", color="A6A6A6", size=11) 
    fuente_valores = Font(name="Segoe UI", color="4DA8DA", bold=True, size=12) 
    
    for fila_res in [4, 5, 6]:
        for col, alineacion, fuente in [("A", "right", fuente_etiquetas), ("D", "right", fuente_etiquetas), 
                                        ("B", "left", fuente_valores), ("E", "left", fuente_valores)]:
            celda = worksheet[f"{col}{fila_res}"]
            celda.font = fuente
            celda.alignment = Alignment(horizontal=alineacion, vertical="center")
            celda.fill = fondo_secundario

# --- Excel para Préstamos ---
def generar_excel_prestamo(df, titulo_reporte, sistema):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        worksheet = aplicar_estilos_base_excel(writer, df, f"RESUMEN: {titulo_reporte.upper()} ({sistema.upper()})")
        ultima_fila = len(df) + 8
        
        worksheet["A4"] = "Monto del Préstamo:"
        worksheet["A5"] = "Primera Cuota:" if sistema == "Alemán" else "Cuota Fija:"
        worksheet["A6"] = "Plazo:"
        worksheet["D4"] = "Total Intereses:"
        worksheet["D5"] = "Total a Pagar:"
        worksheet["D6"] = "Costo del Crédito:"
        
        worksheet["B4"] = f"=SUM(D9:D{ultima_fila})"
        worksheet["B5"] = f"=B9" 
        worksheet["B6"] = f'={len(df)} & " meses"' 
        worksheet["E4"] = f"=SUM(C9:C{ultima_fila})"   
        worksheet["E5"] = f"=SUM(B9:B{ultima_fila})"   
        worksheet["E6"] = f"=(E5/B4)-1"                
        
        pintar_cajas_resumen(worksheet)
        for celda in ["B4", "B5", "E4", "E5"]: worksheet[celda].number_format = '$#,##0.00'
        worksheet["E6"].number_format = '0.0%'  
        
        finalizar_excel(worksheet, ultima_fila)
    return buffer.getvalue()

# --- Excel para Ahorros ---
def generar_excel_ahorro(df, capital_inicial, aporte_mensual):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        worksheet = aplicar_estilos_base_excel(writer, df, "RESUMEN: PLAN DE AHORRO COMPUESTO")
        ultima_fila = len(df) + 8
        
        worksheet["A4"] = "Capital Inicial:"
        worksheet["A5"] = "Aporte Mensual:"
        worksheet["A6"] = "Plazo:"
        worksheet["D4"] = "Total Aportado:"
        worksheet["D5"] = "Total Intereses Ganados:"
        worksheet["D6"] = "Saldo Final:"
        
        worksheet["B4"] = capital_inicial
        worksheet["B5"] = aporte_mensual
        worksheet["B6"] = f'={len(df)} & " meses"' 
        worksheet["E4"] = f"=D{ultima_fila}"           
        worksheet["E5"] = f"=SUM(C9:C{ultima_fila})"   
        worksheet["E6"] = f"=E{ultima_fila}"           
        
        pintar_cajas_resumen(worksheet)
        for celda in ["B4", "B5", "E4", "E5", "E6"]: worksheet[celda].number_format = '$#,##0.00'
        
        finalizar_excel(worksheet, ultima_fila)
    return buffer.getvalue()


# ==========================================
# 4. INTERFAZ WEB CON STREAMLIT
# ==========================================

st.sidebar.title("⚙️ Menú Principal")
st.sidebar.markdown("Elige el tipo de simulador:")
opcion_menu = st.sidebar.radio(
    "",
    ["🏠 Préstamo Hipotecario", "🏢 Préstamo Comercial", "💰 Ahorro / Inversión"]
)

# --- CRÉDITOS DEL DESARROLLADOR ---
st.sidebar.markdown("---") 
st.sidebar.markdown("### 👨‍💻 Desarrollado por")
st.sidebar.info(
    """
    **[TU NOMBRE AQUÍ]**  
    Analista Financiero / Desarrollador  
    
    [![LinkedIn](https://img.shields.io/badge/LinkedIn-0077B5?style=for-the-badge&logo=linkedin&logoColor=white)](https://tu-link-de-linkedin.com)
    """
)

tipo_calculo = None
df_resultado = None

# --- VISTA: HIPOTECARIO ---
if opcion_menu == "🏠 Préstamo Hipotecario":
    st.title("🏠 Simulador Hipotecario")
    st.markdown("Calcula tu crédito de vivienda (Utiliza Sistema Francés por defecto).")
    col1, col2, col3 = st.columns(3)
    with col1: monto = st.number_input("Monto ($)", min_value=1000.0, value=50000.0, step=1000.0)
    with col2: tasa = st.number_input("Tasa Anual (%)", min_value=0.1, value=9.0, step=0.5) / 100
    with col3: plazo = st.number_input("Años", min_value=1, value=20, step=1)
    
    tipo_calculo = "Hipotecario"
    sistema_elegido = "Francés"

# --- VISTA: COMERCIAL ---
elif opcion_menu == "🏢 Préstamo Comercial":
    st.title("🏢 Simulador Comercial")
    st.markdown("Ideal para negocios y pymes. Personaliza tu sistema de amortización.")
    col1, col2, col3 = st.columns(3)
    with col1: monto = st.number_input("Monto ($)", min_value=1000.0, value=15000.0, step=1000.0)
    with col2: tasa = st.number_input("Tasa Anual (%)", min_value=0.1, value=12.0, step=0.5) / 100
    with col3: plazo = st.number_input("Años", min_value=1, value=5, step=1)
        
    sistema_elegido = st.radio("Sistema de Amortización:", ["Francés", "Alemán"], horizontal=True)
    tipo_calculo = "Comercial"

# --- VISTA: AHORRO / INVERSIÓN ---
elif opcion_menu == "💰 Ahorro / Inversión":
    st.title("💰 Simulador de Ahorro e Inversión")
    st.markdown("Descubre cómo crece tu dinero usando el **interés compuesto**.")
    
    col1, col2 = st.columns(2)
    with col1:
        capital_ini = st.number_input("Capital Inicial ($)", min_value=0.0, value=1000.0, step=500.0)
        tasa_ahorro = st.number_input("Tasa de Rendimiento Anual (%)", min_value=0.1, value=8.0, step=0.5) / 100
    with col2:
        aporte_mes = st.number_input("Aporte Adicional Mensual ($)", min_value=0.0, value=200.0, step=50.0)
        plazo_ahorro = st.number_input("Plazo de Inversión (Años)", min_value=1, value=10, step=1)
        
    tipo_calculo = "Ahorro"

# ==========================================
# 5. BOTÓN DE CÁLCULO Y RESULTADOS
# ==========================================
st.markdown("---")
if st.button("📊 Calcular y Generar Reporte"):
    
    if tipo_calculo in ["Hipotecario", "Comercial"]:
        df_resultado = calcular_amortizacion(monto, tasa, plazo, sistema=sistema_elegido)
        primera_cuota = df_resultado.loc[0, "Cuota Mensual"]
        
        st.success(f"**{'Cuota Fija Mensual' if sistema_elegido == 'Francés' else 'Primera Cuota'}:** ${primera_cuota:,.2f}")
        with st.expander("Ver tabla previa (Haz clic para expandir)"): st.dataframe(df_resultado, use_container_width=True)
        
        excel_binario = generar_excel_prestamo(df_resultado, tipo_calculo, sistema_elegido)
        st.download_button("📥 Descargar Excel", data=excel_binario, file_name=f"Prestamo_{tipo_calculo}.xlsx")

    elif tipo_calculo == "Ahorro":
        df_resultado = calcular_ahorro_compuesto(capital_ini, aporte_mes, tasa_ahorro, plazo_ahorro)
        saldo_final = df_resultado.iloc[-1]["Saldo Final"]
        total_interes = df_resultado["Interés del Mes"].sum()
        
        col_res1, col_res2 = st.columns(2)
        col_res1.success(f"**Saldo Final Proyectado:** ${saldo_final:,.2f}")
        col_res2.info(f"**Total Intereses Ganados:** ${total_interes:,.2f}")
        
        with st.expander("Ver tabla previa (Haz clic para expandir)"): st.dataframe(df_resultado, use_container_width=True)
        
        excel_binario = generar_excel_ahorro(df_resultado, capital_ini, aporte_mes)
        st.download_button("📥 Descargar Excel de Ahorro", data=excel_binario, file_name="Plan_Ahorro_Compuesto.xlsx")
