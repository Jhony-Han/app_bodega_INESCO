import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from datetime import datetime, date
import os
import pdfplumber

# ---------------------------------------------------------
# 1. CONFIGURACIÓN DE PÁGINA E IDENTIDAD VISUAL COCA-COLA
# ---------------------------------------------------------
st.set_page_config(
    page_title="Inesco | Digitación y Extracción",
    page_icon="🥤",
    layout="wide"
)

# Estilos CSS personalizados en Rojo Coca-Cola (#E41E2B)
st.markdown("""
    <style>
    .cocacola-header {
        background-color: #E41E2B;
        padding: 20px;
        border-radius: 12px;
        text-align: center;
        color: white;
        margin-bottom: 25px;
        box-shadow: 0px 4px 10px rgba(228, 30, 43, 0.3);
    }
    .cocacola-header h1 {
        color: white !important;
        margin: 0;
        font-weight: 800;
        font-size: 2.2rem;
        letter-spacing: 1px;
    }
    .cocacola-header p {
        color: #FFEBEE !important;
        margin: 5px 0 0 0;
        font-size: 1rem;
        font-weight: 500;
    }
    .sub-title {
        color: #E41E2B;
        font-weight: 700;
        font-size: 1.3rem;
        margin-top: 15px;
        margin-bottom: 10px;
        border-bottom: 2px solid #E41E2B;
        padding-bottom: 5px;
    }
    .stButton>button {
        background-color: #E41E2B !important;
        color: white !important;
        font-weight: bold !important;
        border-radius: 8px !important;
        border: none !important;
        width: 100%;
        padding: 10px;
        transition: 0.3s;
    }
    .stButton>button:hover {
        background-color: #B71C1C !important;
        box-shadow: 0px 4px 8px rgba(0,0,0,0.3);
    }
    </style>
""", unsafe_allow_html=True)

# Encabezado Principal
st.markdown("""
    <div class="cocacola-header">
        <div style="font-size: 2.5rem; margin-bottom: 5px;">🥤</div>
        <h1>DISTRIBUCIONES INESCO</h1>
        <p>Control de Fechas de Vencimiento y Extracción de Planillas</p>
    </div>
""", unsafe_allow_html=True)

# ---------------------------------------------------------
# 2. CATÁLOGO COMPLETO DE PRODUCTOS (136 SKUS)
# ---------------------------------------------------------
CATALOGO_PRODUCTOS = [
    {"sku": "135718", "descripcion": "COCA-COLA 8 OZ VIR(30)"},
    {"sku": "135764", "descripcion": "COCA-COLA SA 8 OZ VIR(30)"},
    {"sku": "56704", "descripcion": "QUATRO CHOICE 8OZ VIR (30)"},
    {"sku": "56521", "descripcion": "SPRITE 8 OZ VIR(30)"},
    {"sku": "56624", "descripcion": "PREMIO ROJO 8 OZ VIR(30)"},
    {"sku": "56557", "descripcion": "SCHWEPPES SODA 8 OZ VIR(30)"},
    {"sku": "160053", "descripcion": "COCA-COLA 350ML VIR(30)"},
    {"sku": "135763", "descripcion": "COCA-COLA ZERO-SA 350ML VIR(30)"},
    {"sku": "56622", "descripcion": "PREMIO ROJO 350 ML (30)"},
    {"sku": "56705", "descripcion": "QUATRO CHOICE 350ML VIR(30)"},
    {"sku": "160070", "descripcion": "SPRITE LL 350ML VIR(30)"},
    {"sku": "160033", "descripcion": "SCHWEPPES SODA 350 ML VIR (30)"},
    {"sku": "135664", "descripcion": "COCA-COLA 1.5 LT NR (12) Nvo"},
    {"sku": "136113", "descripcion": "ESPEJO CCSO 1.5LT PET(12)"},
    {"sku": "135757", "descripcion": "COCA COLA ZERO 1,5 NR (12)"},
    {"sku": "56709", "descripcion": "QUATRO CHOICE 1.5 LT PET (12)"},
    {"sku": "160102", "descripcion": "SPRITE 1.5L PET (12)"},
    {"sku": "120118", "descripcion": "FRESH CITRUS 1.5LT PET(12)NVO"},
    {"sku": "135663", "descripcion": "COCA COLA 2.5LT PET(8) Nvo"},
    {"sku": "136133", "descripcion": "COCA COLA 2.25L PET(8)EINSTEIN"},
    {"sku": "135662", "descripcion": "COCA COLA 3LT PET(6) Nvo"},
    {"sku": "56711", "descripcion": "QUATRO CHOICE 3LTS PET (6)"},
    {"sku": "160124", "descripcion": "SPRITE 3 LTS PET (6)"},
    {"sku": "56625", "descripcion": "PREMIO ROJO FNT 3LT PET(6)"},
    {"sku": "160048", "descripcion": "AGUA BRISA BOTELLON 5 GL TAPA"},
    {"sku": "92769", "descripcion": "BRISA MANZANA 600ML PET(6)"},
    {"sku": "92794", "descripcion": "BRISA MARACUYA 600ML PET(6)"},
    {"sku": "92770", "descripcion": "BRISA MANZANA 1.5LT PET(12)"},
    {"sku": "135760", "descripcion": "COCA-COLA 250 ML PET (12)"},
    {"sku": "135759", "descripcion": "COCA COLA ZERO 250 ML PET(12)"},
    {"sku": "56623", "descripcion": "PREMIO ROJO 250 ML PET(12)"},
    {"sku": "56706", "descripcion": "QUATRO CHOICE 250 ML PET (12)"},
    {"sku": "160071", "descripcion": "SPRITE 250 ML PET (12)"},
    {"sku": "135756", "descripcion": "COCA-COLA ZERO 400 ML PET(12)"},
    {"sku": "56708", "descripcion": "QUATRO CHOICE 400 ML PET(12)"},
    {"sku": "160072", "descripcion": "SPRITE 400 ML PET (12)"},
    {"sku": "160055", "descripcion": "COCA-COLA 400 ML PET (12)"},
    {"sku": "92735", "descripcion": "BRISA SIN GAS 600 ML PET(12)"},
    {"sku": "92736", "descripcion": "BRISA CON GAS 600 ML PET(12)"},
    {"sku": "92737", "descripcion": "BRISA CON GAS 1.5 LT PET(12)"},
    {"sku": "92738", "descripcion": "BRISA SIN GAS 1.5 LT PET(12)"},
    {"sku": "92739", "descripcion": "BRISA CON GAS 2.5 LT PET (6)"},
    {"sku": "92740", "descripcion": "BRISA SIN GAS 2.5 LT PET (6)"},
    {"sku": "92741", "descripcion": "BRISA SIN GAS 5.0 LT PET (2)"},
    {"sku": "92742", "descripcion": "BRISA LIMON 600 ML PET (6)"},
    {"sku": "92743", "descripcion": "BRISA LIMON 1.5 LT PET (12)"},
    {"sku": "92765", "descripcion": "BRISA MENTA 600 ML PET (6)"},
    {"sku": "120116", "descripcion": "FRESH CITRUS 400 ML PET (12)"},
    {"sku": "120117", "descripcion": "FRESH CITRUS 600 ML PET (12)"},
    {"sku": "120119", "descripcion": "FRESH CITRUS 2.5 LT PET (6)"},
    {"sku": "120120", "descripcion": "FRESH CITRUS 3.0 LT PET (6)"},
    {"sku": "160032", "descripcion": "SCHWEPPES SODA 300 ML CAN (12)"},
    {"sku": "160034", "descripcion": "SCHWEPPES TONICA 300 ML CAN(12)"},
    {"sku": "160035", "descripcion": "SCHWEPPES GINGER 300 ML CAN(12)"},
    {"sku": "160056", "descripcion": "COCA-COLA 300 ML CAN (12)"},
    {"sku": "135758", "descripcion": "COCA-COLA ZERO 300 ML CAN(12)"},
    {"sku": "56707", "descripcion": "QUATRO CHOICE 300 ML CAN(12)"},
    {"sku": "160073", "descripcion": "SPRITE 300 ML CAN (12)"},
    {"sku": "136134", "descripcion": "DEL VALLE FRUT MANZANA 300ML CAN"},
    {"sku": "136135", "descripcion": "DEL VALLE FRUT CITRICO 300ML CAN"},
    {"sku": "136136", "descripcion": "DEL VALLE FRUT MORA 300ML CAN"},
    {"sku": "136137", "descripcion": "DEL VALLE FRUT GUANABANA 300ML CAN"},
    {"sku": "136138", "descripcion": "DEL VALLE FRESH CITRICO 400ML PET"},
    {"sku": "136139", "descripcion": "DEL VALLE FRESH MANDARINA 400ML PET"},
    {"sku": "136140", "descripcion": "DEL VALLE FRESH MANZANA 400ML PET"},
    {"sku": "136141", "descripcion": "DEL VALLE FRESH CITRICO 1.5L PET"},
    {"sku": "136142", "descripcion": "DEL VALLE FRESH MANDARINA 1.5L PET"},
    {"sku": "136143", "descripcion": "DEL VALLE FRESH MANZANA 1.5L PET"},
    {"sku": "136144", "descripcion": "DEL VALLE FRESH CITRICO 3.0L PET"},
    {"sku": "136145", "descripcion": "DEL VALLE FRESH MANDARINA 3.0L PET"},
    {"sku": "136146", "descripcion": "DEL VALLE NECTAR DURAZNO 250ML GLASS"},
    {"sku": "136147", "descripcion": "DEL VALLE NECTAR MANZANA 250ML GLASS"},
    {"sku": "136148", "descripcion": "DEL VALLE NECTAR PERA 250ML GLASS"},
    {"sku": "136149", "descripcion": "DEL VALLE NECTAR MANGO 250ML GLASS"},
    {"sku": "136150", "descripcion": "FUZE TEA LIMON 400ML PET"},
    {"sku": "136151", "descripcion": "FUZE TEA DURAZNO 400ML PET"},
    {"sku": "136152", "descripcion": "FUZE TEA VERDE LIMON 400ML PET"},
    {"sku": "136153", "descripcion": "FUZE TEA LIMON 1.5L PET"},
    {"sku": "136154", "descripcion": "FUZE TEA DURAZNO 1.5L PET"},
    {"sku": "136155", "descripcion": "POWERADE MOUNTAIN BLAST 500ML PET"},
    {"sku": "136156", "descripcion": "POWERADE FRUTAS 500ML PET"},
    {"sku": "136157", "descripcion": "POWERADE NARANJA 500ML PET"},
    {"sku": "136158", "descripcion": "POWERADE MOUNTAIN BLAST 1.0L PET"},
    {"sku": "136159", "descripcion": "MONSTER ENERGY REGULAR 473ML CAN"},
    {"sku": "136160", "descripcion": "MONSTER ENERGY ULTRA WHITE 473ML CAN"},
    {"sku": "136161", "descripcion": "MONSTER ENERGY MANGO LOCO 473ML CAN"},
    {"sku": "136162", "descripcion": "MONSTER ENERGY PIPELINE PUNCH 473ML CAN"},
    {"sku": "136163", "descripcion": "MONSTER ENERGY ULTRA SUNRISE 473ML CAN"},
    {"sku": "136164", "descripcion": "MONSTER ENERGY ULTRA PARADISE 473ML CAN"},
    {"sku": "136165", "descripcion": "MONSTER ENERGY ULTRA WATERMELON 473ML CAN"},
    {"sku": "136166", "descripcion": "PREMIO NARANJA 350ML VIR(30)"},
    {"sku": "136167", "descripcion": "PREMIO MANZANA 350ML VIR(30)"},
    {"sku": "136168", "descripcion": "PREMIO LIMON 350ML VIR(30)"},
    {"sku": "136169", "descripcion": "PREMIO NARANJA 1.5L PET(12)"},
    {"sku": "136170", "descripcion": "PREMIO MANZANA 1.5L PET(12)"},
    {"sku": "136171", "descripcion": "PREMIO NARANJA 3.0L PET(6)"},
    {"sku": "136172", "descripcion": "PREMIO MANZANA 3.0L PET(6)"},
    {"sku": "136173", "descripcion": "QUATRO TORONJA 1.5L PET(12)"},
    {"sku": "136174", "descripcion": "QUATRO TORONJA 3.0L PET(6)"},
    {"sku": "136175", "descripcion": "SPRITE ZERO 350ML VIR(30)"},
    {"sku": "136176", "descripcion": "SPRITE ZERO 1.5L PET(12)"},
    {"sku": "136177", "descripcion": "SPRITE ZERO 300ML CAN(12)"},
    {"sku": "136178", "descripcion": "COCA-COLA SIN AZUCAR 350ML VIR(30)"},
    {"sku": "136179", "descripcion": "COCA-COLA SIN AZUCAR 1.5L PET(12)"},
    {"sku": "136180", "descripcion": "COCA-COLA SIN AZUCAR 3.0L PET(6)"},
    {"sku": "136181", "descripcion": "COCA-COLA SIN AZUCAR 300ML CAN(12)"},
    {"sku": "136182", "descripcion": "AGUA BRISA CON GAS 300ML CAN(12)"},
    {"sku": "136183", "descripcion": "AGUA BRISA SPARKLING MANZANA 300ML"},
    {"sku": "136184", "descripcion": "AGUA BRISA SPARKLING LIMON 300ML"},
    {"sku": "136185", "descripcion": "AGUA BRISA SPARKLING TORONJA 300ML"},
    {"sku": "136186", "descripcion": "DEL VALLE 100% NARANJA 1.0L TETRA"},
    {"sku": "136187", "descripcion": "DEL VALLE 100% MANZANA 1.0L TETRA"},
    {"sku": "136188", "descripcion": "DEL VALLE SOYFRUT MELOCOTON 200ML"},
    {"sku": "136189", "descripcion": "DEL VALLE SOYFRUT MANZANA 200ML"},
    {"sku": "136190", "descripcion": "ADECONTIGO NARANJA 200ML TETRA"},
    {"sku": "136191", "descripcion": "ADECONTIGO MANZANA 200ML TETRA"},
    {"sku": "136192", "descripcion": "BURN ENERGY DRINK 250ML CAN"},
    {"sku": "136193", "descripcion": "PREMIO ROJO 1.5L PET(12)"},
    {"sku": "136194", "descripcion": "QUATRO CHOICE 500ML PET(12)"},
    {"sku": "136195", "descripcion": "COCA-COLA 500ML PET(12)"},
    {"sku": "136196", "descripcion": "COCA-COLA ZERO 500ML PET(12)"},
    {"sku": "136197", "descripcion": "SPRITE 500ML PET(12)"},
    {"sku": "136198", "descripcion": "SCHWEPPES TONICA 1.5L PET(12)"},
    {"sku": "136199", "descripcion": "SCHWEPPES GINGER 1.5L PET(12)"},
    {"sku": "136200", "descripcion": "BRISA MARACUYA 1.5L PET(12)"},
    {"sku": "136201", "descripcion": "BRISA MANZANA 2.5L PET(6)"},
    {"sku": "136202", "descripcion": "BRISA LIMON 2.5L PET(6)"},
    {"sku": "136203", "descripcion": "FRESH CITRUS 500ML PET(12)"},
    {"sku": "136204", "descripcion": "FUZE TEA MANZANA 400ML PET"},
    {"sku": "136205", "descripcion": "POWERADE ZERO 500ML PET"}
]

SKU_TO_DESC = {item["sku"]: item["descripcion"] for item in CATALOGO_PRODUCTOS}
OPCIONES_COMBO = [f"{item['sku']} - {item['descripcion']}" for item in CATALOGO_PRODUCTOS]

EXCEL_FECHAS = "fechas_vencimiento_registros.xlsx"

if not os.path.exists(EXCEL_FECHAS):
    df_init = pd.DataFrame(columns=["SKU", "Descripción", "Fecha_Vencimiento", "Fecha_Registro"])
    df_init.to_excel(EXCEL_FECHAS, index=False)

# ---------------------------------------------------------
# 3. PESTAÑAS DE LA APLICACIÓN
# ---------------------------------------------------------
tab1, tab2 = st.tabs(["📌 Digitación de Fechas", "📄 Extracción de PDF (Planillas)"])

# ---------------------------------------------------------
# TAB 1: DIGITACIÓN DE FECHAS DE VENCIMIENTO
# ---------------------------------------------------------
with tab1:
    st.markdown('<p class="sub-title">➕ Registrar Fecha de Vencimiento</p>', unsafe_allow_html=True)
    
    col1, col2 = st.columns([2, 1])
    
    with col1:
        # CON INDEX=NONE LA CASILLA INICIA TOTALMENTE VACÍA
        seleccion = st.selectbox(
            "Busca o selecciona el Código SKU / Producto:",
            options=OPCIONES_COMBO,
            index=None,
            placeholder="🔎 Escribe aquí el SKU o nombre..."
        )
        
        sku_sel = ""
        desc_sel = ""
        if seleccion:
            sku_sel = seleccion.split(" - ")[0]
            desc_sel = SKU_TO_DESC.get(sku_sel, "")

        st.text_input("Descripción del Producto:", value=desc_sel, disabled=True)

    with col2:
        fecha_venc = st.date_input("Fecha de Vencimiento:", value=date.today())
        st.write("")
        st.write("")
        btn_guardar = st.button("💾 Guardar Registro")

    if btn_guardar:
        if not sku_sel:
            st.error("⚠️ Por favor busca y selecciona un SKU válido.")
        else:
            df_curr = pd.read_excel(EXCEL_FECHAS)
            nuevo = pd.DataFrame([{
                "SKU": str(sku_sel),
                "Descripción": desc_sel,
                "Fecha_Vencimiento": fecha_venc.strftime("%d/%m/%Y"),
                "Fecha_Registro": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            }])
            df_updated = pd.concat([df_curr, nuevo], ignore_index=True)
            df_updated.to_excel(EXCEL_FECHAS, index=False)
            st.success(f"✅ Registrado con éxito: {sku_sel} - {desc_sel}")

    st.markdown('<p class="sub-title">📋 Registros Guardados</p>', unsafe_allow_html=True)
    if os.path.exists(EXCEL_FECHAS):
        df_hist = pd.read_excel(EXCEL_FECHAS)
        if not df_hist.empty:
            st.dataframe(df_hist[["SKU", "Descripción", "Fecha_Vencimiento"]], use_container_width=True)

# ---------------------------------------------------------
# TAB 2: EXTRAER PDF DE PLANILLA DE REPARTO Y TOTALIZAR
# ---------------------------------------------------------
with tab2:
    st.markdown('<p class="sub-title">📄 Subir PDF y Generar Excel con Totales Auditoría</p>', unsafe_allow_html=True)
    
    uploaded_pdf = st.file_drop_target if hasattr(st, 'file_drop_target') else st.file_uploader("Sube la Planilla PDF de Reparto:", type=["pdf"])

    if uploaded_pdf is not None:
        with pdfplumber.open(uploaded_pdf) as pdf:
            filas_extraidas = []
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    for row in table:
                        # Filtrar y extraer filas con estructura SKU, Desc, Cajas, Unidades
                        if len(row) >= 4 and str(row[0]).strip().isdigit():
                            filas_extraidas.append({
                                "SKU (Material)": str(row[0]).strip(),
                                "Descripción del Producto": str(row[1]).strip(),
                                "Cantidad (Cajas)": row[2],
                                "Cantidad (Unidades)": row[3]
                            })

        if filas_extraidas:
            df_pdf = pd.DataFrame(filas_extraidas)
            
            # Convertir a valores numéricos
            df_pdf["Cantidad (Cajas)"] = pd.to_numeric(df_pdf["Cantidad (Cajas)"], errors='coerce').fillna(0).astype(int)
            df_pdf["Cantidad (Unidades)"] = pd.to_numeric(df_pdf["Cantidad (Unidades)"], errors='coerce').fillna(0).astype(int)
            
            # CALCULAR TOTALES SUMA
            tot_cajas = df_pdf["Cantidad (Cajas)"].sum()
            tot_unidades = df_pdf["Cantidad (Unidades)"].sum()

            # Mostrar métricas visuales en pantalla
            m1, m2, m3 = st.columns(3)
            m1.metric("Total Productos", len(df_pdf))
            m2.metric("Total Cajas", f"{tot_cajas:,}")
            m3.metric("Total Unidades (Botellas)", f"{tot_unidades:,}")

            # Fila de Total General al final del DataFrame
            fila_tot = pd.DataFrame([{
                "SKU (Material)": "TOTAL GENERAL",
                "Descripción del Producto": "",
                "Cantidad (Cajas)": tot_cajas,
                "Cantidad (Unidades)": tot_unidades
            }])
            
            df_export = pd.concat([df_pdf, fila_tot], ignore_index=True)

            st.dataframe(df_export, use_container_width=True)

            # Exportar archivo formateado
            excel_out = "Planilla_Extraida_Con_Totales.xlsx"
            df_export.to_excel(excel_out, index=False)

            # Aplicar Estilos Excel
            wb = openpyxl.load_workbook(excel_out)
            ws = wb.active
            rojo_fill = PatternFill(start_color="E41E2B", end_color="E41E2B", fill_type="solid")
            bold_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
            double_border = Border(bottom=Side(style='double'), top=Side(style='thin'))

            # Formato de cabecera
            for col in range(1, 5):
                c = ws.cell(row=1, column=col)
                c.fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
                c.font = Font(color="FFFFFF", bold=True)

            # Formato de la última fila (Totales)
            last_row = ws.max_row
            for col in range(1, 5):
                c = ws.cell(row=last_row, column=col)
                c.fill = rojo_fill
                c.font = bold_font
                c.border = double_border

            wb.save(excel_out)

            with open(excel_out, "rb") as f:
                st.download_button(
                    label="📥 Descargar Reporte Excel con Totales",
                    data=f,
                    file_name="Planilla_Reparto_Con_Totales.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        else:
            st.warning("No se encontraron filas válidas de productos en el archivo PDF subido.")
