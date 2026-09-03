
import streamlit as st
from logic import load_stock, load_salidas, load_oc, calculate, excel_bytes

st.set_page_config(page_title="ABASTECIMIENTO", layout="wide")
st.title("ABASTECIMIENTO – Análisis profesional de inventario")

st.caption(
    "Conexiones fijas: STOCK → Sistema; SALIDAS → Material + F.Almac + Unidades; "
    "OC → Material + F. Docum. + Fecha Guia + Unidades + Estado."
)

c1, c2, c3 = st.columns(3)
with c1:
    f_stock = st.file_uploader("Stock", type=["xlsx", "xls"], key="upload_stock")
with c2:
    f_sal = st.file_uploader("Salidas", type=["xlsx", "xls"], key="upload_salidas")
with c3:
    f_oc = st.file_uploader("Órdenes de compra", type=["xlsx", "xls"], key="upload_oc")

c4, c5 = st.columns(2)
with c4:
    horizonte = st.number_input(
        "Horizonte para cantidad a abastecer (meses)",
        min_value=1, max_value=12, value=3
    )
with c5:
    z = st.number_input(
        "Factor de seguridad Z",
        min_value=0.0, max_value=3.0, value=1.65, step=0.05
    )

if f_stock and f_sal and f_oc:
    try:
        stock, map_stock = load_stock(f_stock)
        sal, map_sal = load_salidas(f_sal)
        oc, map_oc = load_oc(f_oc)

        # Validación de conexiones antes del análisis.
        stock_keys = set(stock["key"])
        sal_keys = set(sal["key"])
        oc_keys = set(oc["key"])

        sal_match = len(stock_keys & sal_keys)
        oc_match = len(stock_keys & oc_keys)

        st.success(
            f"Archivos leídos correctamente. "
            f"Stock: {len(stock):,} materiales | "
            f"Salidas: {len(sal):,} movimientos | "
            f"OC: {len(oc):,} registros."
        )

        with st.expander("Validación de columnas utilizadas"):
            st.write("STOCK", map_stock)
            st.write("SALIDAS", map_sal)
            st.write("ÓRDENES DE COMPRA", map_oc)
            st.write(f"Materiales Stock ↔ Salidas: {sal_match:,}")
            st.write(f"Materiales Stock ↔ OC: {oc_match:,}")

        df, periods = calculate(
            stock, sal, oc,
            horizon_months=horizonte,
            z=z
        )

        st.subheader("Resultado")
        st.dataframe(df, use_container_width=True, height=700)

        st.download_button(
            "⬇️ Descargar análisis completo en Excel",
            data=excel_bytes(df),
            file_name="Analisis_Abastecimiento.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_excel"
        )

    except Exception as e:
        st.error("No se pudo procesar el análisis.")
        st.exception(e)
else:
    st.info("Carga los tres archivos para ejecutar el análisis.")
