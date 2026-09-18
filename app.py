import streamlit as st
import pandas as pd
import pdfplumber
import io
import re
from datetime import datetime, timedelta, date
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import dataframe_to_rows

# ---------------------------------------------------------
# 1. CONFIGURACIÓN DE PÁGINA Y ESTILOS
# ---------------------------------------------------------
st.set_page_config(
    page_title="Inesco | Gestión y Extracción",
    page_icon="🥤",
    layout="wide"
)

st.markdown("""
    <style>
    .cocacola-header {
        background-color: #E41E2B;
        padding: 18px;
        border-radius: 12px;
        text-align: center;
        color: white;
        margin-bottom: 20px;
        box-shadow: 0px 4px 10px rgba(228, 30, 43, 0.3);
    }
    .cocacola-header h1 { color: white !important; margin: 0; font-weight: 800; font-size: 1.8rem; }
    .cocacola-header p { color: #FFEBEE !important; margin: 3px 0 0 0; font-size: 0.9rem; }
    .sub-title {
        color: #E41E2B; font-weight: 700; font-size: 1.2rem;
        margin-top: 15px; margin-bottom: 10px;
        border-bottom: 2px solid #E41E2B; padding-bottom: 5px;
    }
    .stButton>button {
        background-color: #E41E2B !important; color: white !important;
        font-weight: bold !important; border-radius: 8px !important; width: 100%;
    }
    </style>
""", unsafe_allow_html=True)

st.markdown("""
    <div class="cocacola-header">
        <h1>DISTRIBUCIONES INESCO</h1>
        <p>Gestión de Inventario, Vencimientos y Extracción por Rutas</p>
    </div>
""", unsafe_allow_html=True)

# ---------------------------------------------------------
# 2. CATÁLOGO COMPLETO EXTRAÍDO DEL EXCEL (136 SKUS)
# ---------------------------------------------------------
CATALOGO_INICIAL = [
    {"sku": "135718", "descripcion": "COCA-COLA 8 OZ VIR(30)"},
    {"sku": "135764", "descripcion": "COCA-COLA SA 8 OZ VIR(30)"},
    {"sku": "56704", "descripcion": "QUATRO CHOICE 8OZ VIR (30)"},
    {"sku": "56521", "descripcion": "SPRITE  8 OZ VIR(30)"},
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
    {"sku": "160099", "descripcion": "KOLA ROMAN 1.5L"},
    {"sku": "56709", "descripcion": "QUATRO 1.5 L"},
    {"sku": "160102", "descripcion": "SPRITE 1.5L (12)"},
    {"sku": "56651", "descripcion": "PREMIO ROJO 1.5 L"},
    {"sku": "160101", "descripcion": "SCHWEPPES SODA 1.5 L (12)"},
    {"sku": "135662", "descripcion": "COCA-COLA 3 LITROS PET(6) Nvo"},
    {"sku": "160124", "descripcion": "SPRITE 3 LTS PET (6)"},
    {"sku": "56711", "descripcion": "QUATRO CHOICE 3LTS PET (6)"},
    {"sku": "56625", "descripcion": "PREMIO ROJO 3 LTS  PET (6)"},
    {"sku": "135755", "descripcion": "COCA COLA ZERO 2.5 LTS (8)"},
    {"sku": "160318", "descripcion": "COCA-COLA 400ML PET  (12)"},
    {"sku": "135760", "descripcion": "COCA-COLA SA 400ml PET(12)"},
    {"sku": "56452", "descripcion": "SCHWEPPES SODA 400ML PET (12)"},
    {"sku": "160479", "descripcion": "KOLA ROMAN 400 ML NR (12)"},
    {"sku": "56627", "descripcion": "PREMIO 400PET"},
    {"sku": "160297", "descripcion": "SPRITE 400PET"},
    {"sku": "56706", "descripcion": "QUATRO CHOICE 400ML NR (12)"},
    {"sku": "135761", "descripcion": "CC ZERO 250ML (12) CON PRECIO"},
    {"sku": "160200", "descripcion": "COCA COLA 250 ML (12)"},
    {"sku": "120205", "descripcion": "MONSTER GREEN ZERO LAT 473ML(12"},
    {"sku": "120180", "descripcion": "MONSTER VIOLET 473ML LAT(12) MQ"},
    {"sku": "120175", "descripcion": "MONSTER PIPELINE PUNCH 473ML LA"},
    {"sku": "120182", "descripcion": "MONSTER PARADISE 473ML LAT(12) MQ"},
    {"sku": "120183", "descripcion": "MONSTER ULTRA 473ML LAT(12) MQ"},
    {"sku": "120215", "descripcion": "FURY LAT 473ML(12)"},
    {"sku": "120156", "descripcion": "MONSTER GREEN 473ML LAT(12) MQ"},
    {"sku": "120181", "descripcion": "MONSTER MGLOCO 473ML LAT(12) MQ"},
    {"sku": "160186", "descripcion": "FUZE TEA DURAZNO PET 400ML  (6)"},
    {"sku": "160187", "descripcion": "FUZE TEA LIMON PET 400 (6)"},
    {"sku": "120117", "descripcion": "FRESH CITRUS 400ML PET (12)NVO"},
    {"sku": "120158", "descripcion": "FRESH MANDARINA 400ML(12)NVO"},
    {"sku": "120184", "descripcion": "FRESH PONCHE FRUTAS 400ML PET(12)"},
    {"sku": "120159", "descripcion": "FRESH MANDARINA 1.5LT PET(12)NVO"},
    {"sku": "92795", "descripcion": "BRISA MARACUYA 1.5LT PET(12)"},
    {"sku": "120118", "descripcion": "DEL VALLE CITRUS 1.5 LTS PET (12)"},
    {"sku": "92794", "descripcion": "BRISA MARACUYA 600ML PET(6)"},
    {"sku": "120116", "descripcion": "FRESH CITRUS 188ML TRP(24)NVO"},
    {"sku": "120134", "descripcion": "FRUTAL MORA 188ML TRP(24)NVO"},
    {"sku": "120136", "descripcion": "FRUTAL MANGO 188ML TRP(24)"},
    {"sku": "120135", "descripcion": "FRUTAL SALPICON 188ML TRP(24)NV"},
    {"sku": "56498", "descripcion": "SODA 10OZ VNR(12)NvaB VIDR"},
    {"sku": "56497", "descripcion": "GINGER 10oz VNR Nva VIDRIO"},
    {"sku": "160192", "descripcion": "COCACOLA 10OZ VNR (12) VIDRIO"},
    {"sku": "135766", "descripcion": "CC SA 10onz"},
    {"sku": "120208", "descripcion": "POWERADE CONTRATAQUE 500ML PET(6)"},
    {"sku": "160121", "descripcion": "POWERADE ROJO FT 500ML PET(6)"},
    {"sku": "160122", "descripcion": "POWERADE AZUL MB 500ML PET(6)"},
    {"sku": "160241", "descripcion": "MANANTIAL 600 ML (12)"},
    {"sku": "160240", "descripcion": "MANANTIAL GAS 600 ML (12)"},
    {"sku": "160168", "descripcion": "AGUA BRISA PET 280 ML (24)"},
    {"sku": "92893", "descripcion": "AGUA BRISA GAS LIMA LIMON PET 280 ML (24)"},
    {"sku": "160047", "descripcion": "AGUA BRISA PET 600ML (24)"},
    {"sku": "160046", "descripcion": "AGUA BRISA GAS PET 600ML (24)"},
    {"sku": "92892", "descripcion": "BRISA LIMA LIMON 600 ML (6)"},
    {"sku": "160008", "descripcion": "BRISA ECOFLEX 1LT"},
    {"sku": "160048", "descripcion": "AGUA BRISA BOTELLON 5 GL TV"},
    {"sku": "92793", "descripcion": "BRISA MARACUYA 280ML PET(24)"},
    {"sku": "120120", "descripcion": "2PK FRESH CITRUS 2.5L PET NVO"},
    {"sku": "136247", "descripcion": "2 PK (1CCSO 3LT + 1 QT3LT) Nvo"},
    {"sku": "160201", "descripcion": "6PACK CCR 250ML MINI PET"},
    {"sku": "120131", "descripcion": "FRUTAL MANGO 500ML PET(6)NVO"},
    {"sku": "120129", "descripcion": "FRUTAL MORA 500ML PET(6)NVO"},
    {"sku": "120128", "descripcion": "FRUTAL MANGOFRS 500ML PET(6)NVO"},
    {"sku": "136401", "descripcion": "CCZR CAFE 269ML LAT(12)"},
    {"sku": "120190", "descripcion": "FRUTAL MANGOFRS 1.2LT PET(6)"},
    {"sku": "120188", "descripcion": "FRUTAL MANGO 1.2L PET (6)"},
    {"sku": "120191", "descripcion": "FRUTAL MORA 1.2LT PET(6)"},
    {"sku": "160029", "descripcion": "COCA-COLA 330ML LATA (12)"},
    {"sku": "120194", "descripcion": "6PK(2MG+2MR+2MF)188ML"},
    {"sku": "92769", "descripcion": "BRISA MANZANA 600ML"},
    {"sku": "56703", "descripcion": "QUATRO CHOICE 250ML PET (12)"},
    {"sku": "92768", "descripcion": "BRISA MANZANA 288ML"},
    {"sku": "120130", "descripcion": "FRUTAL SALPICON 500ML PET(6)NVO"},
    {"sku": "136159", "descripcion": "COCA COLA SO 1L PET (12)"},
    {"sku": "136133", "descripcion": "COCA COLA 2.25L PET(8)EINSTEIN"},
    {"sku": "160060", "descripcion": "MANANTIAL 1.5 PET(9)"},
    {"sku": "135675", "descripcion": "COMBO AHORRO CCSO 3LT"},
    {"sku": "92911", "descripcion": "BRISA LIMA LIMON 1.5L PET(12)NV"},
    {"sku": "92770", "descripcion": "BRISA MANZANA 1.5 L"},
    {"sku": "135663", "descripcion": "CC PET 2.5LTNR 8 PK  Nvo"},
    {"sku": "120119", "descripcion": "FRESH CITRUS 2,5LT"},
    {"sku": "56820", "descripcion": "QUATRO CHOICE 2.5L PET (8)"},
    {"sku": "136403", "descripcion": "6PK LAT CCZR KZ 269ML"},
    {"sku": "136665", "descripcion": "6PK CCSO 269ML LAT"},
    {"sku": "135795", "descripcion": "6PK CCZR 250 ML PET KZ"},
    {"sku": "56842", "descripcion": "FLASHLYTE COCO LIM 625ML PET(6)"},
    {"sku": "56841", "descripcion": "FLASHLYTE UVA 625ML PET(6)"},
    {"sku": "136691", "descripcion": "EXH TRUL 72UX70G 6.7K KIT1(1)"},
    {"sku": "136468", "descripcion": "HALLS 100S EXT STRONG COL 310GR KIT30(1)"},
    {"sku": "136487", "descripcion": "LOKINO CARAMELO 400GR KIT24(1)"},
    {"sku": "136747", "descripcion": "CHOCOLORES SURT 540G KIT12(1)"},
    {"sku": "136475", "descripcion": "BARRILETE 440GR PAIS KIT18(1)"},
    {"sku": "136701", "descripcion": "BIANCHI CHOCOBLA 600G KIT18(1)"},
    {"sku": "136700", "descripcion": "BIANCHI CHOCOLA 600G KIT18(1)"},
    {"sku": "136479", "descripcion": "CHAO FRESA 350GR KIT24(1)"},
    {"sku": "136478", "descripcion": "CHAO MENTA 350GR KIT24(1)"},
    {"sku": "136466", "descripcion": "OREO ROLLO 109GR KIT30(1)"},
    {"sku": "136712", "descripcion": "TRULU 70U MAS VENDID 1.6K KIT6(1)"},
    {"sku": "136467", "descripcion": "OREO REGULAR COL 324GR KIT18(1)"},
    {"sku": "136671", "descripcion": "TRUL MAXI+VEND 12U 1.2K KIT6(1)"},
    {"sku": "136630", "descripcion": "TRULULU MINIPACK 24U 960GR KIT1(1)"},
    {"sku": "136720", "descripcion": "TRUL MINIPACK 24U 840GR KIT12(1)"},
    {"sku": "136632", "descripcion": "TRUL MAXIPACK 12U 1.2K KIT1(1)"},
    {"sku": "136675", "descripcion": "TRUL SABOR SURT 12U 360G KIT12(1)"},
    {"sku": "136670", "descripcion": "TRUL MXPK 12U 1.3K TOPEX KIT6(1)"},
    {"sku": "136020", "descripcion": "BOCAD ATUN ACEIT 140GR KIT48(1)"},
    {"sku": "136024", "descripcion": "ATUN RALL ACEITE 150GR KIT48(1)"},
    {"sku": "136027", "descripcion": "BOCADO ATUN AGUA 140GR KIT18(1)"},
    {"sku": "136474", "descripcion": "ATUN ISABEL ACEI 142G KIT48(1)"},
    {"sku": "136399", "descripcion": "COCA COLA ORIGINAL 269ML LAT(12)"},
    {"sku": "136400", "descripcion": "COCA COLA ZERO 269ML LAT(12)KZ"},
    {"sku": "92899", "descripcion": "REFAJO KOLA ROMAN 330M LAT(6)"},
    {"sku": "136461", "descripcion": "ZUCARITAS MEGA 115G KIT54(1)"},
    {"sku": "136465", "descripcion": "FROOT LOOPS MEGA 90GR KIT54(1)"},
    {"sku": "136462", "descripcion": "CHOKRISPIS MEGA 115G KIT54(1)"},
    {"sku": "136560", "descripcion": "CHOCO KRISPIS 240GR KIT30(1)"},
    {"sku": "136559", "descripcion": "ZUCARITAS 240GR KIT30(1)"},
    {"sku": "136561", "descripcion": "FROOT LOOPS 200GR KIT30(1)"},
    {"sku": "136705", "descripcion": "TRIDENT 5SX18U MENTA 153G KIT30(1)"},
    {"sku": "136707", "descripcion": "TRIDENT 1SX60U MENTA 90G KIT40(1)"},
    {"sku": "136710", "descripcion": "CLUB SOCIAL 9U ORIG 216G KIT24(1)"}
]

# ---------------------------------------------------------
# 3. GESTIÓN DE ESTADO Y AUTO-ELIMINACIÓN A LOS 3 DÍAS
# ---------------------------------------------------------
if "catalogo" not in st.session_state:
    st.session_state.catalogo = CATALOGO_INICIAL

if "vencimientos" not in st.session_state:
    st.session_state.vencimientos = []

# Limpieza automática: elimina registros mayores a 3 días (72 horas)
ahora = datetime.now()
st.session_state.vencimientos = [
    reg for reg in st.session_state.vencimientos
    if (ahora - datetime.strptime(reg["created_at"], "%Y-%m-%d %H:%M:%S")).total_seconds() < 3 * 86400
]

# ---------------------------------------------------------
# 4. FUNCIÓN PARA GENERAR EXCEL FORMATO COCA-COLA (ROJO)
# ---------------------------------------------------------
def exportar_excel_cocacola(df, titulo_hoja="Hoja1"):
    wb = Workbook()
    ws = wb.active
    ws.title = titulo_hoja[:30].replace(":", "").replace("/", "")
    
    red_fill = PatternFill(start_color="E41E2B", end_color="E41E2B", fill_type="solid")
    white_bold_font = Font(color="FFFFFF", bold=True, name="Calibri", size=11)
    thin_border = Border(
        left=Side(style='thin', color='CCCCCC'),
        right=Side(style='thin', color='CCCCCC'),
        top=Side(style='thin', color='CCCCCC'),
        bottom=Side(style='thin', color='CCCCCC')
    )
    
    for r in dataframe_to_rows(df, index=False, header=True):
        ws.append(r)
        
    for cell in ws[1]:
        cell.fill = red_fill
        cell.font = white_bold_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            cell.border = thin_border
            cell.alignment = Alignment(vertical="center")
            
    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = col[0].column_letter
        ws.column_dimensions[col_letter].width = max(max_len + 3, 12)
        
    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

# ---------------------------------------------------------
# 5. PESTAÑAS Y NAVEGACIÓN
# ---------------------------------------------------------
tab1, tab2, tab3 = st.tabs(["📅 Fechas de Vencimiento", "📦 Administrar SKUs", "📄 Extracción PDF por Ruta"])

# --- TAB 1: FECHAS DE VENCIMIENTO ---
with tab1:
    st.markdown('<p class="sub-title">➕ Agregar Registro de Vencimiento</p>', unsafe_allow_html=True)
    
    skus_opt = [f"{item['sku']} - {item['descripcion']}" for item in st.session_state.catalogo]
    
    c1, c2 = st.columns([2, 1])
    with c1:
        sel_sku = st.selectbox("Seleccionar Producto:", options=skus_opt, index=None, placeholder="🔎 Escribe aquí SKU o Nombre...")
    with c2:
        f_venc = st.date_input("Fecha de Vencimiento:", value=date.today())
        
    if st.button("💾 Guardar Fecha de Vencimiento"):
        if sel_sku:
            s_code, s_desc = sel_sku.split(" - ", 1)
            st.session_state.vencimientos.append({
                "id": len(st.session_state.vencimientos) + 1,
                "SKU": s_code,
                "Descripción": s_desc,
                "Fecha Vencimiento": f_venc.strftime("%d/%m/%Y"),
                "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            })
            st.success("✅ Registro guardado con éxito.")
            st.rerun()
        else:
            st.warning("⚠️ Debes seleccionar un SKU primero.")

    st.markdown('<p class="sub-title">📋 Registros Guardados (Permanecen 3 días activos)</p>', unsafe_allow_html=True)
    
    if st.session_state.vencimientos:
        for idx, row in enumerate(st.session_state.vencimientos):
            col_a, col_b, col_c, col_d = st.columns([2, 4, 3, 2])
            col_a.write(f"**{row['SKU']}**")
            col_b.write(row['Descripción'])
            
            # Opción para Modificar la fecha
            fecha_actual = datetime.strptime(row['Fecha Vencimiento'], "%d/%m/%Y").date()
            nueva_f = col_c.date_input("Fecha", value=fecha_actual, key=f"date_{row['id']}")
            st.session_state.vencimientos[idx]['Fecha Vencimiento'] = nueva_f.strftime("%d/%m/%Y")
            
            # Opción para Borrar
            if col_d.button("❌ Borrar", key=f"del_{row['id']}"):
                st.session_state.vencimientos.pop(idx)
                st.rerun()

        st.divider()
        df_venc_out = pd.DataFrame(st.session_state.vencimientos)[["SKU", "Descripción", "Fecha Vencimiento"]]
        excel_bytes = exportar_excel_cocacola(df_venc_out, "Vencimientos")
        
        col_dl, col_sh = st.columns([1, 1])
        with col_dl:
            st.download_button(
                label="📥 Descargar Reporte en Excel (Rojo)",
                data=excel_bytes,
                file_name=f"Vencimientos_Inesco_{date.today()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        with col_sh:
            mensaje_wa = f"Reporte de Vencimientos Inesco - {date.today()}"
            st.markdown(f'[📲 Compartir por WhatsApp](https://api.whatsapp.com/send?text={mensaje_wa})')
    else:
        st.info("No hay registros activos actualmente.")

# --- TAB 2: ADMINISTRAR SKUS ---
with tab2:
    st.markdown('<p class="sub-title">⚙️ Agregar o Eliminar SKUs del Catálogo</p>', unsafe_allow_html=True)
    
    col_add1, col_add2 = st.columns([1, 2])
    with col_add1:
        nuevo_sku = st.text_input("Nuevo Código SKU:")
    with col_add2:
        nueva_desc = st.text_input("Descripción del Producto:")
        
    if st.button("➕ Agregar Nuevo SKU"):
        if nuevo_sku and nueva_desc:
            st.session_state.catalogo.append({"sku": nuevo_sku.strip(), "descripcion": nueva_desc.strip()})
            st.success(f"SKU {nuevo_sku} agregado al catálogo.")
            st.rerun()
        else:
            st.error("Por favor completa el SKU y la Descripción.")

    st.divider()
    st.write(f"**Catálogo Actual ({len(st.session_state.catalogo)} SKUs):**")
    
    for idx, item in enumerate(st.session_state.catalogo):
        c_k, c_d, c_b = st.columns([2, 5, 2])
        c_k.write(f"**{item['sku']}**")
        c_d.write(item['descripcion'])
        if c_b.button("🗑️ Eliminar", key=f"cat_del_{idx}"):
            st.session_state.catalogo.pop(idx)
            st.rerun()

# --- TAB 3: EXTRACTION PDF POR RUTAS SEPARADAS ---
with tab3:
    st.markdown('<p class="sub-title">📄 Extraer PDF en Archivos Excel Independientes por Ruta</p>', unsafe_allow_html=True)
    
    uploaded_file = st.file_uploader("Cargar documento PDF de planillas:", type=["pdf"])
    
    if uploaded_file:
        with pdfplumber.open(uploaded_file) as pdf:
            texto_completo = ""
            for page in pdf.pages:
                texto_completo += (page.extract_text() or "") + "\n"
                
        # Separación de rutas por patrón
        bloques = re.split(r'(RUTA\s*:\s*\d+)', texto_completo, flags=re.IGNORECASE)
        
        rutas_data = {}
        if len(bloques) > 1:
            for i in range(1, len(bloques), 2):
                nombre_ruta = bloques[i].strip().upper()
                contenido = bloques[i+1]
                
                lineas = contenido.split('\n')
                items = []
                for line in lineas:
                    match = re.search(r'(\d{5,6})\s+(.+?)\s+(\d+)\s*$', line)
                    if match:
                        items.append({
                            "SKU": match.group(1),
                            "Descripción": match.group(2).strip(),
                            "Cantidad": int(match.group(3))
                        })
                if items:
                    rutas_data[nombre_ruta] = pd.DataFrame(items)
        
        if rutas_data:
            st.success(f"Se detectaron {len(rutas_data)} rutas independientes en el PDF.")
            
            for ruta, df_ruta in rutas_data.items():
                st.write(f"### {ruta}")
                st.dataframe(df_ruta, use_container_width=True)
                
                excel_ruta = exportar_excel_cocacola(df_ruta, ruta)
                
                st.download_button(
                    label=f"📥 Descargar Excel para {ruta}",
                    data=excel_ruta,
                    file_name=f"Planilla_{ruta.replace(' ', '_').replace(':', '')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"dl_{ruta}"
                )
                st.divider()
        else:
            st.warning("No se detectaron marcas de RUTA en el PDF. Verifique el formato.")
