import datetime
import io
import os
import re
import pandas as pd
import pypdf
import streamlit as st

# Configuración adaptada para celulares y pantallas táctiles
st.set_page_config(
    page_title="Sistema Bodega",
    page_icon="📦",
    layout="centered",
    initial_sidebar_state="collapsed",
)

DB_FECHAS_LOCAL = "Digitacion_Fechas_Local.xlsx"


# ==========================================
# FUNCIONES PARA EL MÓDULO DE FECHAS (OFFLINE)
# ==========================================
def gestionar_base_fechas():
    """Carga los registros locales y elimina automáticamente los mayores a 3 días."""
    if os.path.exists(DB_FECHAS_LOCAL):
        df = pd.read_excel(DB_FECHAS_LOCAL)
        df["Fecha_Ingreso"] = pd.to_datetime(df["Fecha_Ingreso"])

        # Borrado automático a los 3 días
        limite = datetime.datetime.now() - datetime.timedelta(days=3)
        df_limpio = df[df["Fecha_Ingreso"] >= limite].copy()

        if len(df_limpio) < len(df):
            df_limpio.to_excel(DB_FECHAS_LOCAL, index=False)

        return df_limpio
    else:
        return pd.DataFrame(
            columns=[
                "SKU",
                "Descripción",
                "Fecha_Vencimiento",
                "Fecha_Ingreso",
            ]
        )


def guardar_fecha_local(sku, descripcion, fecha_venc):
    """Guarda una fecha de vencimiento ingresada."""
    df = gestionar_base_fechas()
    nuevo = pd.DataFrame([{
        "SKU": str(sku).strip(),
        "Descripción": descripcion.strip(),
        "Fecha_Vencimiento": fecha_venc.strftime("%d/%m/%Y"),
        "Fecha_Ingreso": datetime.datetime.now(),
    }])
    df = pd.concat([df, nuevo], ignore_index=True)
    df.to_excel(DB_FECHAS_LOCAL, index=False)
    return df


# ==========================================
# INTERFAZ EN PESTAÑAS (TABS)
# ==========================================
st.title("📦 Sistema de Bodega")

tab1, tab2 = st.tabs([
    "📄 Extraer PDF a Excel",
    "📅 Digitación de Fechas",
])

# ------------------------------------------
# PESTAÑA 1: EXTRAER PLANILLA PDF
# ------------------------------------------
with tab1:
    st.header("📄 Convertir PDF de Planilla")
    pdf_file = st.file_uploader(
        "Sube la planilla en PDF desde tu celular", type=["pdf"]
    )

    if pdf_file is not None:
        reader = pypdf.PdfReader(pdf_file)
        st.success(
            f"Planilla cargada con éxito ({len(reader.pages)} páginas)."
        )

        items = []
        current_ruta = None
        current_fecha = None

        for page in reader.pages:
            p_text = page.extract_text()
            ruta_match = re.search(
                r"Ruta\s*/\s*No\.de Carga:\s*([A-Z0-9/]+)", p_text
            )
            fecha_match = re.search(r"Fecha de Entrega:\s*([\d\.]+)", p_text)

            if ruta_match:
                current_ruta = ruta_match.group(1)
            if fecha_match:
                current_fecha = fecha_match.group(1)

            lines = p_text.split("\n")
            for line in lines:
                line_str = line.strip()
                if re.match(r"^\d{5,6}\s+", line_str):
                    m = re.match(
                        r"^(\d{5,6})\s+(.*?)\s+([\d]*\s*/\s*[\d]+)", line_str
                    )
                    if m:
                        sku = m.group(1)
                        desc = m.group(2).strip()
                        qty_str = m.group(3).replace(" ", "")
                        parts = qty_str.split("/")
                        cajas = int(parts[0]) if parts[0] != "" else 0
                        unidades = int(parts[1]) if parts[1] != "" else 0

                        items.append({
                            "Fecha Entrega": current_fecha,
                            "Ruta": current_ruta,
                            "SKU": sku,
                            "Descripción": desc,
                            "Cajas": cajas,
                            "Unidades": unidades,
                        })

        if items:
            df_all = pd.DataFrame(items)
            st.subheader("Rutas Generadas para Descargar:")

            for ruta, group in df_all.groupby("Ruta"):
                clean_name = ruta.replace("/", "_")
                output = io.BytesIO()

                df_export = group[[
                    "Fecha Entrega",
                    "Ruta",
                    "SKU",
                    "Descripción",
                    "Cajas",
                    "Unidades",
                ]].copy()
                df_export.columns = [
                    "Fecha de Entrega",
                    "Ruta / Carga",
                    "SKU (Material)",
                    "Descripción del Producto",
                    "Cantidad (Cajas)",
                    "Cantidad (Unidades)",
                ]

                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    df_export.to_excel(writer, index=False, sheet_name="Carga")
                    ws = writer.sheets["Carga"]
                    r_tot = len(df_export) + 2
                    ws.cell(row=r_tot, column=4, value="TOTAL")
                    ws.cell(row=r_tot, column=5, value=group["Cajas"].sum())
                    ws.cell(row=r_tot, column=6, value=group["Unidades"].sum())

                st.download_button(
                    label=f"📥 Descargar Excel Ruta {clean_name}",
                    data=output.getvalue(),
                    file_name=f"Ruta_{clean_name}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

# ------------------------------------------
# PESTAÑA 2: DIGITACIÓN DE FECHAS (CALENDARIO)
# ------------------------------------------
with tab2:
    st.header("📅 Digitación de Fechas")

    with st.form("form_fechas", clear_on_submit=True):
        sku_in = st.text_input("Código SKU / Material:")
        desc_in = st.text_input("Descripción del Producto:")
        fecha_in = st.date_input(
            "Selecciona Fecha de Vencimiento:", datetime.date.today()
        )

        btn_save = st.form_submit_button("💾 Guardar Registro")

    if btn_save and sku_in:
        guardar_fecha_local(sku_in, desc_in, fecha_in)
        st.success("¡Guardado correctamente!")

    st.markdown("---")
    st.subheader("📋 Registros (Últimos 3 Días)")

    df_act = gestionar_base_fechas()
    if not df_act.empty:
        st.dataframe(
            df_act[["SKU", "Descripción", "Fecha_Vencimiento"]],
            use_container_width=True,
        )

        with open(DB_FECHAS_LOCAL, "rb") as f:
            st.download_button(
                label="📥 Descargar Base de Fechas en Excel",
                data=f,
                file_name=f"Fechas_Vencimiento_{datetime.date.today().strftime('%d_%m_%Y')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
    else:
        st.info("No hay registros guardados en los últimos 3 días.")
