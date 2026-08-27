import io
from dataclasses import dataclass
from typing import Literal

import pandas as pd
import streamlit as st
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

Sistema = Literal["Francés", "Alemán"]
FORMATO_MONEDA = "$#,##0.00"
FILA_ENCABEZADO = 8


@dataclass(frozen=True)
class ConfiguracionPrestamo:
    nombre: str
    icono: str
    descripcion: str
    monto: float
    tasa: float
    plazo: int
    permite_elegir_sistema: bool = False


PRESTAMOS = {
    "🏠 Préstamo Hipotecario": ConfiguracionPrestamo(
        "Hipotecario", "🏠", "Calcula tu crédito de vivienda con el sistema francés.",
        50_000.0, 9.0, 20,
    ),
    "🏢 Préstamo Comercial": ConfiguracionPrestamo(
        "Comercial", "🏢", "Ideal para negocios y pymes.",
        15_000.0, 12.0, 5, True,
    ),
}


def calcular_amortizacion(
    monto: float, tasa_anual: float, anios: int, sistema: Sistema = "Francés"
) -> pd.DataFrame:
    """Genera una tabla de amortización francesa o alemana."""
    meses = anios * 12
    tasa_mensual = tasa_anual / 12
    saldo = monto
    filas = []

    if sistema == "Francés":
        cuota_fija = (
            monto / meses
            if tasa_mensual == 0
            else monto * tasa_mensual / (1 - (1 + tasa_mensual) ** -meses)
        )
        capital_fijo = None
    elif sistema == "Alemán":
        cuota_fija = 0.0
        capital_fijo = monto / meses
    else:
        raise ValueError(f"Sistema no válido: {sistema}")

    for mes in range(1, meses + 1):
        interes = saldo * tasa_mensual
        capital = cuota_fija - interes if sistema == "Francés" else capital_fijo
        cuota = cuota_fija if sistema == "Francés" else capital + interes

        # La última cuota absorbe pequeñas diferencias de punto flotante.
        if mes == meses:
            capital = saldo
            cuota = capital + interes

        saldo = max(0.0, saldo - capital)
        filas.append({
            "Mes": mes,
            "Cuota Mensual": round(cuota, 2),
            "Pago Interés": round(interes, 2),
            "Pago Capital": round(capital, 2),
            "Saldo Restante": round(saldo, 2),
        })

    return pd.DataFrame(filas)


def calcular_ahorro_compuesto(
    capital_inicial: float,
    aporte_mensual: float,
    tasa_anual: float,
    anios: int,
) -> pd.DataFrame:
    """Proyecta aportes realizados al final de cada mes."""
    tasa_mensual = tasa_anual / 12
    saldo = capital_inicial
    total_aportado = capital_inicial
    filas = []

    for mes in range(1, anios * 12 + 1):
        interes = saldo * tasa_mensual
        saldo += interes + aporte_mensual
        total_aportado += aporte_mensual
        filas.append({
            "Mes": mes,
            "Aporte Mensual": round(aporte_mensual, 2),
            "Interés del Mes": round(interes, 2),
            "Total Aportado": round(total_aportado, 2),
            "Saldo Final": round(saldo, 2),
        })

    return pd.DataFrame(filas)


def crear_reporte_excel(
    datos: pd.DataFrame,
    titulo: str,
    etiquetas: dict[str, str],
    valores: dict[str, object],
    celdas_porcentaje: tuple[str, ...] = (),
) -> bytes:
    """Crea el Excel común de préstamos y ahorros."""
    buffer = io.BytesIO()
    ultima_fila = FILA_ENCABEZADO + len(datos)

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        datos.to_excel(writer, sheet_name="Reporte", index=False, startrow=7)
        hoja = writer.sheets["Reporte"]
        fondo_oscuro = PatternFill("solid", fgColor="1A1A1A")
        fondo_resumen = PatternFill("solid", fgColor="252525")
        hoja.sheet_view.showGridLines = False

        for fila in hoja.iter_rows(min_row=1, max_row=7, min_col=1, max_col=5):
            for celda in fila:
                celda.fill = fondo_oscuro

        hoja.merge_cells("A1:E1")
        hoja["A1"] = titulo
        hoja["A1"].font = Font(name="Segoe UI", color="FFFFFF", bold=True, size=16)
        hoja["A1"].alignment = Alignment(horizontal="center", vertical="center")

        hoja.merge_cells("A2:E2")
        hoja["A2"] = "Generado por ZoraEC · Desarrollado por Freddy Beltrán A. (2026)"
        hoja["A2"].font = Font(name="Segoe UI", color="A6A6A6", italic=True, size=10)
        hoja["A2"].alignment = Alignment(horizontal="right", vertical="center")

        for referencia, texto in etiquetas.items():
            celda = hoja[referencia]
            celda.value = texto
            celda.font = Font(name="Segoe UI", color="A6A6A6", size=11)
            celda.alignment = Alignment(horizontal="right", vertical="center")
            celda.fill = fondo_resumen

        for referencia, valor in valores.items():
            celda = hoja[referencia]
            celda.value = valor
            celda.font = Font(name="Segoe UI", color="4DA8DA", bold=True, size=12)
            celda.alignment = Alignment(horizontal="left", vertical="center")
            celda.fill = fondo_resumen
            if referencia not in {"B6"}:
                celda.number_format = (
                    "0.0%" if referencia in celdas_porcentaje else FORMATO_MONEDA
                )

        tabla = Table(displayName="TablaDatos", ref=f"A8:E{ultima_fila}")
        tabla.tableStyleInfo = TableStyleInfo(
            name="TableStyleDark1", showRowStripes=True
        )
        hoja.add_table(tabla)

        for fila in hoja.iter_rows(min_row=9, max_row=ultima_fila):
            fila[0].alignment = Alignment(horizontal="center")
            for celda in fila[1:]:
                celda.number_format = FORMATO_MONEDA
                celda.alignment = Alignment(horizontal="right")

        for columna in hoja.columns:
            ancho = max(len(str(celda.value or "")) for celda in columna)
            letra = get_column_letter(columna[0].column)
            hoja.column_dimensions[letra].width = max(ancho + 4, 18)

    return buffer.getvalue()


def generar_excel_prestamo(
    datos: pd.DataFrame, nombre: str, sistema: Sistema
) -> bytes:
    fin = FILA_ENCABEZADO + len(datos)
    return crear_reporte_excel(
        datos,
        f"RESUMEN: {nombre.upper()} ({sistema.upper()})",
        {
            "A4": "Monto del Préstamo:",
            "A5": "Primera Cuota:" if sistema == "Alemán" else "Cuota Fija:",
            "A6": "Plazo:",
            "D4": "Total Intereses:",
            "D5": "Total a Pagar:",
            "D6": "Costo del Crédito:",
        },
        {
            "B4": f"=SUM(D9:D{fin})",
            "B5": "=B9",
            "B6": f'={len(datos)} & " meses"',
            "E4": f"=SUM(C9:C{fin})",
            "E5": f"=SUM(B9:B{fin})",
            "E6": "=(E5/B4)-1",
        },
        ("E6",),
    )


def generar_excel_ahorro(
    datos: pd.DataFrame, capital_inicial: float, aporte_mensual: float
) -> bytes:
    fin = FILA_ENCABEZADO + len(datos)
    return crear_reporte_excel(
        datos,
        "RESUMEN: PLAN DE AHORRO COMPUESTO",
        {
            "A4": "Capital Inicial:",
            "A5": "Aporte Mensual:",
            "A6": "Plazo:",
            "D4": "Total Aportado:",
            "D5": "Total Intereses Ganados:",
            "D6": "Saldo Final:",
        },
        {
            "B4": capital_inicial,
            "B5": aporte_mensual,
            "B6": f'={len(datos)} & " meses"',
            "E4": f"=D{fin}",
            "E5": f"=SUM(C9:C{fin})",
            "E6": f"=E{fin}",
        },
    )


def mostrar_tabla(datos: pd.DataFrame) -> None:
    with st.expander("Ver tabla previa"):
        st.dataframe(datos, use_container_width=True, hide_index=True)


def mostrar_prestamo(config: ConfiguracionPrestamo) -> None:
    st.title(f"{config.icono} Simulador {config.nombre}")
    st.markdown(config.descripcion)

    with st.form(f"formulario_{config.nombre.lower()}"):
        columna_1, columna_2, columna_3 = st.columns(3)
        monto = columna_1.number_input(
            "Monto ($)", min_value=1_000.0, value=config.monto, step=1_000.0
        )
        tasa = columna_2.number_input(
            "Tasa Anual (%)", min_value=0.0, value=config.tasa, step=0.5
        ) / 100
        plazo = columna_3.number_input(
            "Años", min_value=1, value=config.plazo, step=1
        )
        sistema: Sistema = "Francés"
        if config.permite_elegir_sistema:
            sistema = st.radio(
                "Sistema de Amortización:",
                ["Francés", "Alemán"],
                horizontal=True,
            )
        calcular = st.form_submit_button("📊 Calcular y Generar Reporte")

    if not calcular:
        return

    datos = calcular_amortizacion(monto, tasa, plazo, sistema)
    etiqueta = "Cuota Fija Mensual" if sistema == "Francés" else "Primera Cuota"
    st.success(f"**{etiqueta}:** ${datos.iloc[0]['Cuota Mensual']:,.2f}")
    mostrar_tabla(datos)
    st.download_button(
        "📥 Descargar Excel",
        data=generar_excel_prestamo(datos, config.nombre, sistema),
        file_name=f"Prestamo_{config.nombre}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


def mostrar_ahorro() -> None:
    st.title("💰 Simulador de Ahorro e Inversión")
    st.markdown("Descubre cómo crece tu dinero usando el **interés compuesto**.")

    with st.form("formulario_ahorro"):
        columna_1, columna_2 = st.columns(2)
        capital = columna_1.number_input(
            "Capital Inicial ($)", min_value=0.0, value=1_000.0, step=500.0
        )
        tasa = columna_1.number_input(
            "Tasa de Rendimiento Anual (%)", min_value=0.0, value=8.0, step=0.5
        ) / 100
        aporte = columna_2.number_input(
            "Aporte Adicional Mensual ($)", min_value=0.0, value=200.0, step=50.0
        )
        plazo = columna_2.number_input(
            "Plazo de Inversión (Años)", min_value=1, value=10, step=1
        )
        calcular = st.form_submit_button("📊 Calcular y Generar Reporte")

    if not calcular:
        return

    datos = calcular_ahorro_compuesto(capital, aporte, tasa, plazo)
    columna_1, columna_2 = st.columns(2)
    columna_1.success(
        f"**Saldo Final Proyectado:** ${datos.iloc[-1]['Saldo Final']:,.2f}"
    )
    columna_2.info(
        f"**Total Intereses Ganados:** ${datos['Interés del Mes'].sum():,.2f}"
    )
    mostrar_tabla(datos)
    st.download_button(
        "📥 Descargar Excel de Ahorro",
        data=generar_excel_ahorro(datos, capital, aporte),
        file_name="Plan_Ahorro_Compuesto.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


def main() -> None:
    st.set_page_config(
        page_title="Simulador Financiero 360",
        layout="centered",
        page_icon="🏦",
    )
    st.sidebar.title("⚙️ Menú Principal")
    opcion = st.sidebar.radio(
        "Elige el tipo de simulador:", [*PRESTAMOS, "💰 Ahorro / Inversión"]
    )
    st.sidebar.divider()
    st.sidebar.markdown("### 👨‍💻 Creado por ZoraEC")
    st.sidebar.info("**Freddy Beltrán A.**  \nDesarrollador")

    if opcion in PRESTAMOS:
        mostrar_prestamo(PRESTAMOS[opcion])
    else:
        mostrar_ahorro()


if __name__ == "__main__":
    main()
