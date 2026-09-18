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
    page_title="Distribuciones INESCO - Bodega",
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
    cols_esperadas = ["SKU", "Descripción"]
    if os.path.exists(DB_PRODUCTOS_LOCAL):
        try:
            df = pd.read_excel(DB_PRODUCTOS_LOCAL, dtype={"SKU": str})
            df["SKU"] = df["SKU"].astype(str).str.strip()
            df["Descripción"] = df["Descripción"].astype(str).str.strip()
            for col in cols_esperadas:
                if col not in df.columns:
                    raise ValueError("Columna faltante")
            return df
        except Exception:
            pass

    # Base por defecto si falla o no existe
    df_init = pd.DataFrame(
        columns=cols_esperadas,
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
    """Guarda el catálogo de productos."""
    df.to_excel(DB_PRODUCTOS_LOCAL, index=False)


def gestionar_base_fechas():
    """Carga los registros locales de fechas y limpia viejos."""
    cols_esperadas = [
        "ID",
        "SKU",
        "Descripción",
        "Fecha_Vencimiento",
        "Fecha_Ingreso",
    ]
    if os.path.exists(DB_FECHAS_LOCAL):
        try:
            df = pd.read_excel(DB_FECHAS_LOCAL, dtype={"SKU": str})
            # Verificar columnas
            if not all(col in df.columns for col in cols_esperadas):
                raise ValueError("Estructura antigua detectada")

            df["Fecha_Ingreso"] = pd.to_datetime(df["Fecha_Ingreso"])
            limite = datetime.datetime.now() - datetime.timedelta(days=3)
            df_limpio = df[df["Fecha_Ingreso"] >= limite].copy()

            if len(df_limpio) < len(df):
                df_limpio.to_excel(DB_FECHAS_LOCAL, index=False)

            return df_limpio
        except Exception:
            pass

    # Si no existe o tiene error de estructura, crea uno nuevo
    df_empty = pd.DataFrame(columns=cols_esperadas)
    df_empty.to_excel(DB_FECHAS_LOCAL, index=False)
    return df_empty


def guardar_base_fechas(df):
    """Guarda los registros de fechas."""
    df.to_excel(DB_FECHAS_LOCAL, index=False)


# ==========================================
# INTERFAZ PRINCIPAL
# ==========================================
st.title("🏭 Distribuciones INESCO")

tab1, tab2, tab3 = st.tabs([
    "📅 Digitación de Fechas",
    "📄 Extraer PDF a Excel",
    "⚙️ Catálogo de Productos",
])

# ------------------------------------------
# PESTAÑA 1: DIGITACIÓN DE FECHAS
# ------------------------------------------
with tab1:
    st.header("📅 Digitación de Fechas de Vencimiento")

    df_cat = cargar_catalogo_productos()
    dict_prod = dict(zip(df_cat["SKU"], df_cat["Descripción"]))

    st.subheader("➕ Registrar Producto")

    # Selector de SKU
    opciones_sku = [""] + list(dict_prod.keys())
    sku_seleccionado = st.selectbox(
        "Selecciona el Código SKU:",
        options=opciones_sku,
        format_func=lambda x: f"{x} — {dict_prod.get(x, '')}"
        if x != ""
        else "--- Seleccionar SKU ---",
        key="sb_sku_fechas",
    )

    # Autocompletado automático de la descripción
    desc_auto = dict_prod.get(sku_seleccionado, "")

    st.text_input(
        "Descripción del Producto (Autocompletado):",
        value=desc_auto,
        disabled=True,
        key="txt_desc_disabled",
    )

    fecha_vencimiento = st.date_input(
        "Selecciona Fecha de Vencimiento:",
        datetime.date.today(),
        key="dt_venc_input",
    )

    if st.button("💾 Guardar Registro", key="btn_guardar_registro"):
        if sku_seleccionado != "":
            df_fechas = gestionar_base_fechas()
            nuevo_id = (
                int(df_fechas["ID"].max()) + 1 if not df_fechas.empty else 1
            )

            nuevo_reg = pd.DataFrame([{
                "ID": nuevo_id,
                "SKU": str(sku_seleccionado),
                "Descripción": desc_auto,
                "Fecha_Vencimiento": fecha_vencimiento.strftime("%d/%m/%Y"),
                "Fecha_Ingreso": datetime.datetime.now(),
            }])

            df_fechas = pd.concat([df_fechas, nuevo_reg], ignore_index=True)
            guardar_base_fechas(df_fechas)

            st.success(
                f"¡Guardado! SKU: {sku_seleccionado} - Fecha: {fecha_vencimiento.strftime('%d/%m/%Y')}"
            )
            st.rerun()
        else:
            st.warning("Por favor selecciona un Código SKU.")

    st.markdown("---")
    st.subheader("📋 Registros Guardados (Últimos 3 Días)")

    df_registros = gestionar_base_fechas()

    if not df_registros.empty:
        st.dataframe(
            df_registros[["SKU", "Descripción", "Fecha_Vencimiento"]],
            use_container_width=True,
        )

        # Edición y Eliminación de registros erróneos
        with st.expander("🛠️ Modificar o Eliminar un Registro"):
            id_mod = st.number_input(
                "Ingresa el ID o número de fila a gestionar (1, 2, 3...):",
                min_value=1,
                max_value=len(df_registros),
                step=1,
                key="num_id_mod_idx",
            )
            idx_real = id_mod - 1

            reg_actual = df_registros.iloc[idx_real]
            st.info(
                f"Seleccionado: **{reg_actual['SKU']}** - {reg_actual['Descripción']} (Fecha actual: {reg_actual['Fecha_Vencimiento']})"
            )

            col_a, col_b = st.columns(2)
            with col_a:
                nueva_f = st.date_input("Nueva Fecha:", key="dt_edit_val")
                if st.button("✏️ Actualizar Fecha", key="btn_update_reg"):
                    df_registros.iloc[idx_real, df_registros.columns.get_loc("Fecha_Vencimiento")] = nueva_f.strftime("%d/%m/%Y")
                    guardar_base_fechas(df_registros)
                    st.success("¡Fecha modificada correctamente!")
                    st.rerun()

            with col_b:
                if st.button("🗑️ Borrar este Registro", key="btn_delete_reg"):
                    df_registros = df_registros.drop(df_registros.index[idx_real]).reset_index(drop=True)
                    guardar_base_fechas(df_registros)
                    st.success("¡Registro eliminado!")
                    st.rerun()

        # EXPORTACIÓN BONITA EN EXCEL PARA COMPARTIR
        output_excel = io.BytesIO()
        with pd.ExcelWriter(output_excel, engine="openpyxl") as writer:
            wb = writer.book
            ws = wb.create_sheet(title="Fechas_Vencimiento")

            # Estilos
            fill_header = PatternFill(
                start_color="1F4E78", end_color="1F4E78", fill_type="solid"
            )
            font_title = Font(
                name="Calibri", size=14, bold=True, color="1F4E78"
            )
            font_header = Font(
                name="Calibri", size=11, bold=True, color="FFFFFF"
            )
            font_data = Font(name="Calibri", size=11)
            thin_border = Border(
                left=Side(style="thin", color="BFBFBF"),
                right=Side(style="thin", color="BFBFBF"),
                top=Side(style="thin", color="BFBFBF"),
                bottom=Side(style="thin", color="BFBFBF"),
            )
            align_center = Alignment(horizontal="center", vertical="center")
            align_left = Alignment(horizontal="left", vertical="center")

            # Encabezado Empresa
            ws.merge_cells("A1:C1")
            ws["A1"] = "DISTRIBUCIONES INESCO"
            ws["A1"].font = font_title
            ws["A1"].alignment = align_center

            ws["A2"] = (
                f"Reporte de Fechas de Vencimiento — Generado: {datetime.date.today().strftime('%d/%m/%Y')}"
            )
            ws["A2"].font = Font(name="Calibri", size=10, italic=True)

            # Encabezados de Tabla (Fila 4)
            headers = ["SKU", "Descripción del Producto", "Fecha Vencimiento"]
            for col_idx, text in enumerate(headers, 1):
                cell = ws.cell(row=4, column=col_idx, value=text)
                cell.fill = fill_header
                cell.font = font_header
                cell.alignment = align_center
                cell.border = thin_border

            # Datos (Fila 5 en adelante)
            r_idx = 5
            for _, r in df_registros.iterrows():
                c1 = ws.cell(row=r_idx, column=1, value=str(r["SKU"]))
                c2 = ws.cell(row=r_idx, column=2, value=str(r["Descripción"]))
                c3 = ws.cell(
                    row=r_idx, column=3, value=str(r["Fecha_Vencimiento"])
                )

                c1.alignment = align_center
                c2.alignment = align_left
                c3.alignment = align_center

                for c in (c1, c2, c3):
                    c.font = font_data
                    c.border = thin_border
                r_idx += 1

            # Ancho automático de columnas
            for col in ws.columns:
                max_len = max(
                    len(str(cell.value or "")) for cell in col
                )
                col_letter = get_column_letter(col[0].column)
                ws.column_dimensions[col_letter].width = max(max_len + 4, 16)

        st.download_button(
            label="📥 Exportar Excel Distribuciones INESCO",
            data=output_excel.getvalue(),
            file_name=f"Distribuciones_INESCO_Fechas_{datetime.date.today().strftime('%d_%m_%Y')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="btn_download_fechas_inesco",
        )
    else:
        st.info("No hay registros guardados en los últimos 3 días.")

# ------------------------------------------
# PESTAÑA 2: EXTRAER PDF A EXCEL
# ------------------------------------------
with tab2:
    st.header("📄 Convertir PDF de Planilla a Excel")
    pdf_file = st.file_uploader(
        "Sube la planilla en PDF desde tu celular",
        type=["pdf"],
        key="pdf_uploader_key",
    )

    if pdf_file is not None:
        reader = pypdf.PdfReader(pdf_file)
        st.success(f"Planilla cargada ({len(reader.pages)} páginas).")

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
            st.subheader("Archivos por Ruta para Descargar:")

            for idx, (ruta, group) in enumerate(df_all.groupby("Ruta")):
                clean_name = ruta.replace("/", "_")
                fecha_hdr = group["Fecha Entrega"].iloc[0] or "N/A"

                out_pdf_excel = io.BytesIO()
                with pd.ExcelWriter(
                    out_pdf_excel, engine="openpyxl"
                ) as writer:
                    wb = writer.book
                    ws = wb.create_sheet(title="Planilla")

                    fill_header = PatternFill(
                        start_color="1F4E78", end_color="1F4E78", fill_type="solid"
                    )
                    font_title = Font(
                        name="Calibri", size=13, bold=True, color="1F4E78"
                    )
                    font_header = Font(
                        name="Calibri", size=11, bold=True, color="FFFFFF"
                    )
                    font_data = Font(name="Calibri", size=11)
                    thin_border = Border(
                        left=Side(style="thin", color="BFBFBF"),
                        right=Side(style="thin", color="BFBFBF"),
                        top=Side(style="thin", color="BFBFBF"),
                        bottom=Side(style="thin", color="BFBFBF"),
                    )

                    align_center = Alignment(horizontal="center", vertical="center")
                    align_left = Alignment(horizontal="left", vertical="center")
                    align_right = Alignment(horizontal="right", vertical="center")

                    # Empresa y Encabezado
                    ws.merge_cells("A1:D1")
                    ws["A1"] = "DISTRIBUCIONES INESCO"
                    ws["A1"].font = font_title
                    ws["A1"].alignment = align_center

                    ws["A2"] = f"Fecha de Entrega: {fecha_hdr}"
                    ws["C2"] = f"Ruta / Carga: {ruta}"
                    ws["A2"].font = Font(name="Calibri", size=11, bold=True)
                    ws["C2"].font = Font(name="Calibri", size=11, bold=True)

                    # Tabla Headers
                    headers = [
                        "SKU (Material)",
                        "Descripción del Producto",
                        "Cantidad (Cajas)",
                        "Cantidad (Unidades)",
                    ]
                    for col_n, h_text in enumerate(headers, 1):
                        cell = ws.cell(row=4, column=col_n, value=h_text)
                        cell.fill = fill_header
                        cell.font = font_header
                        cell.alignment = align_center
                        cell.border = thin_border

                    # Filas
                    row_i = 5
                    for _, r_data in group.iterrows():
                        ws.cell(
                            row=row_i, column=1, value=str(r_data["SKU"])
                        ).alignment = align_center
                        ws.cell(
                            row=row_i,
                            column=2,
                            value=str(r_data["Descripción"]),
                        ).alignment = align_left
                        ws.cell(
                            row=row_i, column=3, value=int(r_data["Cajas"])
                        ).alignment = align_right
                        ws.cell(
                            row=row_i, column=4, value=int(r_data["Unidades"])
                        ).alignment = align_right

                        for c_n in range(1, 5):
                            c = ws.cell(row=row_i, column=c_n)
                            c.font = font_data
                            c.border = thin_border
                        row_i += 1

                    # Anchos
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
                    data=out_pdf_excel.getvalue(),
                    file_name=f"Ruta_{clean_name}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"btn_dl_ruta_pdf_{idx}",
                )

# ------------------------------------------
# PESTAÑA 3: CATÁLOGO DE PRODUCTOS (INDEPENDIENTE)
# ------------------------------------------
with tab3:
    st.header("⚙️ Administración del Catálogo de Productos")
    st.write(
        "Aquí puedes agregar, actualizar o eliminar productos para que aparezcan automáticamente en la digitación."
    )

    df_cat_actual = cargar_catalogo_productos()

    st.subheader("➕ Agregar / Editar Producto")
    with st.form("form_cat_productos", clear_on_submit=True):
        nuevo_sku = st.text_input("Código SKU / Material:", key="in_cat_sku")
        nueva_desc = st.text_input(
            "Descripción del Producto:", key="in_cat_desc"
        )
        btn_cat_guardar = st.form_submit_button("💾 Guardar en Catálogo")

    if btn_cat_guardar:
        if nuevo_sku and nueva_desc:
            sku_clean = nuevo_sku.strip()
            desc_clean = nueva_desc.strip()

            if sku_clean in df_cat_actual["SKU"].values:
                df_cat_actual.loc[
                    df_cat_actual["SKU"] == sku_clean, "Descripción"
                ] = desc_clean
                st.success(
                    f"Producto SKU {sku_clean} actualizado a '{desc_clean}'."
                )
            else:
                nuevo_p = pd.DataFrame(
                    [{"SKU": sku_clean, "Descripción": desc_clean}]
                )
                df_cat_actual = pd.concat(
                    [df_cat_actual, nuevo_p], ignore_index=True
                )
                st.success(f"Producto SKU {sku_clean} agregado al catálogo.")

            guardar_catalogo_productos(df_cat_actual)
            st.rerun()
        else:
            st.warning("Completa el SKU y la Descripción.")

    st.markdown("---")
    st.subheader("📋 Productos Actualmente en Sistema")
    st.dataframe(df_cat_actual, use_container_width=True)

    with st.expander("🗑️ Eliminar Producto del Catálogo"):
        sku_borrar = st.selectbox(
            "Selecciona el SKU que deseas borrar:",
            options=[""] + list(df_cat_actual["SKU"]),
            key="sb_borrar_cat",
        )
        if st.button("Eliminar del Catálogo", key="btn_borrar_cat"):
            if sku_borrar != "":
                df_cat_actual = df_cat_actual[
                    df_cat_actual["SKU"] != sku_borrar
                ].reset_index(drop=True)
                guardar_catalogo_productos(df_cat_actual)
                st.success(f"SKU {sku_borrar} eliminado del catálogo.")
                st.rerun()
