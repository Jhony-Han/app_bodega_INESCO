import datetime
import io
import os
import re
import pandas as pd
import pypdf
import streamlit as st
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# Configuración de página
st.set_page_config(
    page_title="Sistema Bodega",
    page_icon="📦",
    layout="centered",
    initial_sidebar_state="collapsed",
)

DB_FECHAS_LOCAL = "Digitacion_Fechas_Local.xlsx"
DB_PRODUCTOS_LOCAL = "Catalogo_Productos.xlsx"


# ==========================================
# GESTIÓN DE BASES DE DATOS (EXCEL LOCAL)
# ==========================================
def cargar_catalogo_productos():
    """Carga el catálogo de productos SKU-Descripción."""
    if os.path.exists(DB_PRODUCTOS_LOCAL):
        df = pd.read_excel(DB_PRODUCTOS_LOCAL, dtype={"SKU": str})
        df["SKU"] = df["SKU"].str.strip()
        df["Descripción"] = df["Descripción"].astype(str).str.strip()
        return df
    else:
        # Base inicial por defecto si no existe
        df_init = pd.DataFrame(
            columns=["SKU", "Descripción"],
            data=[
                ["160053", "COCA-COLA"],
                ["135718", "COCA-COLA ZERO"],
                ["135763", "COCA-COLA SABOR ORIGINAL"],
                ["56704", "QUATRO TORONJA"],
                ["56521", "SPRITE ZERO"],
                ["56624", "PREMIO ROJO"],
                ["160048", "AGUA BRISA CON GAS"],
                ["120131", "FRUTAL MANZANA"],
            ],
        )
        df_init.to_excel(DB_PRODUCTOS_LOCAL, index=False)
        return df_init


def guardar_catalogo_productos(df):
    """Guarda el catálogo de productos actualizado."""
    df.to_excel(DB_PRODUCTOS_LOCAL, index=False)


def gestionar_base_fechas():
    """Carga los registros locales y elimina automáticamente los mayores a 3 días."""
    if os.path.exists(DB_FECHAS_LOCAL):
        df = pd.read_excel(DB_FECHAS_LOCAL, dtype={"SKU": str})
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
                "ID",
                "SKU",
                "Descripción",
                "Fecha_Vencimiento",
                "Fecha_Ingreso",
            ]
        )


def guardar_base_fechas(df):
    """Guarda la base de registros de fechas."""
    df.to_excel(DB_FECHAS_LOCAL, index=False)


# ==========================================
# INTERFAZ PRINCIPAL
# ==========================================
st.title("📦 Sistema de Bodega")

tab1, tab2, tab3 = st.tabs([
    "📄 Extraer PDF a Excel",
    "📅 Digitación de Fechas",
    "⚙️ Catálogo de Productos",
])

# ------------------------------------------
# PESTAÑA 1: EXTRAER PLANILLA PDF A EXCEL ESTILIZADO
# ------------------------------------------
with tab1:
    st.header("📄 Convertir PDF de Planilla")
    pdf_file = st.file_uploader(
        "Sube la planilla en PDF desde tu celular",
        type=["pdf"],
        key="pdf_uploader_key",
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

            for idx, (ruta, group) in enumerate(df_all.groupby("Ruta")):
                clean_name = ruta.replace("/", "_")
                fecha_hdr = group["Fecha Entrega"].iloc[0] or "N/A"

                output = io.BytesIO()

                # Crear Excel con openpyxl estilizado
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    # Crear hoja vacía
                    wb = writer.book
                    ws = wb.create_sheet(title="Carga")

                    # Definir Estilos
                    fill_header_info = PatternFill(
                        start_color="1F4E78",
                        end_color="1F4E78",
                        fill_type="solid",
                    )
                    fill_header_tbl = PatternFill(
                        start_color="2F5597",
                        end_color="2F5597",
                        fill_type="solid",
                    )
                    fill_total = PatternFill(
                        start_color="D9E1F2",
                        end_color="D9E1F2",
                        fill_type="solid",
                    )

                    font_title = Font(name="Calibri", size=11, bold=True)
                    font_header = Font(
                        name="Calibri", size=11, bold=True, color="FFFFFF"
                    )
                    font_data = Font(name="Calibri", size=11)
                    font_total = Font(name="Calibri", size=11, bold=True)

                    thin_border = Border(
                        left=Side(style="thin", color="BFBFBF"),
                        right=Side(style="thin", color="BFBFBF"),
                        top=Side(style="thin", color="BFBFBF"),
                        bottom=Side(style="thin", color="BFBFBF"),
                    )

                    align_center = Alignment(
                        horizontal="center", vertical="center"
                    )
                    align_left = Alignment(horizontal="left", vertical="center")
                    align_right = Alignment(
                        horizontal="right", vertical="center"
                    )

                    # 1. Encabezado de Información
                    ws["A1"] = "Fecha de Entrega:"
                    ws["B1"] = fecha_hdr
                    ws["A2"] = "Ruta / Carga:"
                    ws["B2"] = ruta

                    ws["A1"].font = font_title
                    ws["A2"].font = font_title
                    ws["B1"].font = font_data
                    ws["B2"].font = font_data

                    # 2. Encabezados de Tabla (Fila 4)
                    headers = [
                        "SKU (Material)",
                        "Descripción del Producto",
                        "Cantidad (Cajas)",
                        "Cantidad (Unidades)",
                    ]
                    for col_num, h_text in enumerate(headers, 1):
                        cell = ws.cell(row=4, column=col_num)
                        cell.value = h_text
                        cell.fill = fill_header_tbl
                        cell.font = font_header
                        cell.alignment = align_center
                        cell.border = thin_border

                    # 3. Filas de Datos
                    row_idx = 5
                    for _, r_data in group.iterrows():
                        ws.cell(
                            row=row_idx, column=1, value=str(r_data["SKU"])
                        ).alignment = align_center
                        ws.cell(
                            row=row_idx,
                            column=2,
                            value=str(r_data["Descripción"]),
                        ).alignment = align_left
                        ws.cell(
                            row=row_idx, column=3, value=int(r_data["Cajas"])
                        ).alignment = align_right
                        ws.cell(
                            row=row_idx, column=4, value=int(r_data["Unidades"])
                        ).alignment = align_right

                        for col_num in range(1, 5):
                            c = ws.cell(row=row_idx, column=col_num)
                            c.font = font_data
                            c.border = thin_border
                        row_idx += 1

                    # 4. Fila de Totales
                    ws.cell(
                        row=row_idx, column=2, value="TOTAL"
                    ).alignment = align_right
                    ws.cell(
                        row=row_idx, column=3, value=group["Cajas"].sum()
                    ).alignment = align_right
                    ws.cell(
                        row=row_idx, column=4, value=group["Unidades"].sum()
                    ).alignment = align_right

                    for col_num in range(1, 5):
                        c = ws.cell(row=row_idx, column=col_num)
                        c.font = font_total
                        c.fill = fill_total
                        c.border = thin_border

                    # Ajuste de Ancho de Columnas
                    for col in ws.columns:
                        max_len = max(
                            len(str(cell.value or "")) for cell in col
                        )
                        col_letter = get_column_letter(col[0].column)
                        ws.column_dimensions[col_letter].width = max(
                            max_len + 4, 15
                        )

                st.download_button(
                    label=f"📥 Descargar Excel Ruta {clean_name}",
                    data=output.getvalue(),
                    file_name=f"Ruta_{clean_name}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"btn_dl_ruta_{idx}",
                )

# ------------------------------------------
# PESTAÑA 2: DIGITACIÓN, EDICIÓN Y REGISTRO DE FECHAS
# ------------------------------------------
with tab2:
    st.header("📅 Digitación y Control de Fechas")

    df_cat = cargar_catalogo_productos()
    dict_prod = dict(zip(df_cat["SKU"], df_cat["Descripción"]))

    st.subheader("➕ Registrar Nueva Fecha")

    sku_sel = st.selectbox(
        "Selecciona SKU / Código:",
        options=[""] + list(dict_prod.keys()),
        format_func=lambda x: f"{x} - {dict_prod.get(x, '')}"
        if x in dict_prod
        else "Escribe un SKU nuevo...",
        key="sb_sku_key",
    )

    sku_manual = st.text_input("O ingresa SKU manualmente:", key="txt_sku_key")
    sku_final = sku_sel if sku_sel else sku_manual.strip()

    desc_autocompletada = dict_prod.get(sku_final, "")
    desc_final = st.text_input(
        "Descripción del Producto:",
        value=desc_autocompletada,
        key="txt_desc_key",
    )

    fecha_venc = st.date_input(
        "Selecciona Fecha de Vencimiento:",
        datetime.date.today(),
        key="dt_venc_key",
    )

    if st.button("💾 Guardar Registro", key="btn_save_fecha"):
        if sku_final and desc_final:
            df_fechas = gestionar_base_fechas()
            nuevo_id = (
                int(df_fechas["ID"].max()) + 1 if not df_fechas.empty else 1
            )

            nuevo_reg = pd.DataFrame([{
                "ID": nuevo_id,
                "SKU": str(sku_final),
                "Descripción": desc_final,
                "Fecha_Vencimiento": fecha_venc.strftime("%d/%m/%Y"),
                "Fecha_Ingreso": datetime.datetime.now(),
            }])

            df_fechas = pd.concat([df_fechas, nuevo_reg], ignore_index=True)
            guardar_base_fechas(df_fechas)

            st.success("¡Registro guardado exitosamente!")
            st.rerun()
        else:
            st.warning("Por favor completa el SKU y la Descripción.")

    st.markdown("---")
    st.subheader("📋 Registros (Últimos 3 Días)")

    df_act = gestionar_base_fechas()

    if not df_act.empty:
        st.dataframe(
            df_act[["ID", "SKU", "Descripción", "Fecha_Vencimiento"]],
            use_container_width=True,
        )

        # Modificación o eliminación de registro
        with st.expander("🛠️ Modificar / Eliminar un Registro"):
            id_sel = st.number_input(
                "Ingresa el ID del registro a modificar/eliminar:",
                min_value=1,
                step=1,
                key="num_id_mod",
            )

            reg_mod = df_act[df_act["ID"] == id_sel]

            if not reg_mod.empty:
                st.info(
                    f"Registro seleccionado: SKU {reg_mod.iloc[0]['SKU']} - {reg_mod.iloc[0]['Descripción']}"
                )

                col_mod1, col_mod2 = st.columns(2)
                with col_mod1:
                    nueva_f_venc = st.date_input(
                        "Nueva Fecha Vencimiento:", key="dt_mod_key"
                    )
                    if st.button(
                        "✏️ Actualizar Fecha", key="btn_update_fecha"
                    ):
                        df_act.loc[
                            df_act["ID"] == id_sel, "Fecha_Vencimiento"
                        ] = nueva_f_venc.strftime("%d/%m/%Y")
                        guardar_base_fechas(df_act)
                        st.success("¡Registro actualizado!")
                        st.rerun()

                with col_mod2:
                    if st.button("🗑️ Eliminar Registro", key="btn_del_fecha"):
                        df_act = df_act[df_act["ID"] != id_sel]
                        guardar_base_fechas(df_act)
                        st.success("¡Registro eliminado!")
                        st.rerun()

        # Descarga de Excel de Fechas
        output_fechas = io.BytesIO()
        with pd.ExcelWriter(output_fechas, engine="openpyxl") as writer:
            df_act[["SKU", "Descripción", "Fecha_Vencimiento"]].to_excel(
                writer, index=False, sheet_name="Fechas"
            )

        st.download_button(
            label="📥 Descargar Base de Fechas en Excel",
            data=output_fechas.getvalue(),
            file_name=f"Fechas_Vencimiento_{datetime.date.today().strftime('%d_%m_%Y')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="btn_dl_fechas_key",
        )
    else:
        st.info("No hay registros guardados en los últimos 3 días.")

# ------------------------------------------
# PESTAÑA 3: GESTIÓN DEL CATÁLOGO DE PRODUCTOS
# ------------------------------------------
with tab3:
    st.header("⚙️ Administración de Productos")

    df_cat_actual = cargar_catalogo_productos()

    st.subheader("➕ Agregar / Actualizar Producto en Catálogo")

    with st.form("form_cat_prod", clear_on_submit=True):
        new_sku = st.text_input("Código SKU / Material:", key="cat_sku_in")
        new_desc = st.text_input(
            "Descripción del Producto:", key="cat_desc_in"
        )
        btn_cat_save = st.form_submit_button(
            "💾 Guardar / Actualizar Producto", key="btn_cat_save_key"
        )

    if btn_cat_save:
        if new_sku and new_desc:
            sku_clean = new_sku.strip()
            desc_clean = new_desc.strip()

            # Si ya existe, actualiza; si no, agrega
            if sku_clean in df_cat_actual["SKU"].values:
                df_cat_actual.loc[
                    df_cat_actual["SKU"] == sku_clean, "Descripción"
                ] = desc_clean
                st.success(
                    f"Producto {sku_clean} actualizado a '{desc_clean}'."
                )
            else:
                nuevo_prod = pd.DataFrame(
                    [{"SKU": sku_clean, "Descripción": desc_clean}]
                )
                df_cat_actual = pd.concat(
                    [df_cat_actual, nuevo_prod], ignore_index=True
                )
                st.success(f"Producto {sku_clean} agregado al catálogo.")

            guardar_catalogo_productos(df_cat_actual)
            st.rerun()
        else:
            st.warning("Completa el SKU y la Descripción.")

    st.markdown("---")
    st.subheader("📋 Catálogo Actual de Productos")
    st.dataframe(df_cat_actual, use_container_width=True)

    # Eliminar producto del catálogo
    with st.expander("🗑️ Eliminar Producto del Catálogo"):
        sku_del = st.selectbox(
            "Selecciona el SKU a eliminar:",
            options=[""] + list(df_cat_actual["SKU"]),
            key="sb_sku_del",
        )
        if st.button("Eliminar del Catálogo", key="btn_del_cat_item"):
            if sku_del:
                df_cat_actual = df_cat_actual[
                    df_cat_actual["SKU"] != sku_del
                ].reset_index(drop=True)
                guardar_catalogo_productos(df_cat_actual)
                st.success(f"Producto {sku_del} eliminado del catálogo.")
                st.rerun()
