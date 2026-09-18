import streamlit as st
import pandas as pd
import pdfplumber
import io
import re
import json
from datetime import datetime, date
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------
# 1. CONFIGURACIÓN DE PÁGINA Y ESTILOS
# ---------------------------------------------------------
st.set_page_config(
    page_title="Inesco | Gestión y Extracción",
    page_icon="🚛",
    layout="wide",
    initial_sidebar_state="collapsed"
)

st.markdown("""
    <style>
    .cocacola-header {
        background: linear-gradient(135deg, #E41E2B 0%, #B3000C 100%);
        padding: 20px;
        border-radius: 15px;
        text-align: center;
        color: white;
        margin-bottom: 20px;
        box-shadow: 0px 6px 15px rgba(228, 30, 43, 0.4);
    }
    .cocacola-header .logo-truck {
        font-size: 2.8rem;
        margin-bottom: 5px;
        display: inline-block;
    }
    .cocacola-header h1 { 
        color: white !important; 
        margin: 0; 
        font-weight: 800; 
        font-size: 1.8rem;
        letter-spacing: 1px;
    }
    .cocacola-header p { 
        color: #FFEBEE !important; 
        margin: 5px 0 0 0; 
        font-size: 0.95rem; 
    }
    .sub-title {
        color: #E41E2B; 
        font-weight: 700; 
        font-size: 1.2rem;
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
        width: 100%;
        padding: 10px 0px;
    }
    @media (max-width: 640px) {
        .cocacola-header h1 { font-size: 1.4rem; }
        .cocacola-header .logo-truck { font-size: 2rem; }
        .cocacola-header p { font-size: 0.8rem; }
    }
    </style>
""", unsafe_allow_html=True)

st.markdown("""
    <div class="cocacola-header">
        <div class="logo-truck">🚛 🥤</div>
        <h1>DISTRIBUCIONES INESCO</h1>
        <p>Gestión de Inventario, Vencimientos y Extracción por Rutas</p>
    </div>
""", unsafe_allow_html=True)

# ---------------------------------------------------------
# 2. GESTIÓN DE DATOS EN SESIÓN CON RESPALDO SEGURO
# ---------------------------------------------------------
if "vencimientos" not in st.session_state:
    st.session_state.vencimientos = []

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

if "catalogo" not in st.session_state:
    st.session_state.catalogo = CATALOGO_INICIAL

# ---------------------------------------------------------
# 3. EXPORTADOR MULTI-PESTAÑA SEGURO
# ---------------------------------------------------------
def exportar_excel_multiruta(rutas_dict, fecha_str):
    wb = Workbook()
    default_sheet = wb.active
    
    blue_title_font = Font(color="003366", bold=True, size=14, name="Calibri")
    sub_font = Font(italic=True, size=10, name="Calibri", color="333333")
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True, size=11, name="Calibri")
    
    total_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    total_font = Font(color="002060", bold=True, size=11, name="Calibri")
    
    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )
    
    first_sheet = True
    
    for nombre_ruta, df_r in rutas_dict.items():
        safe_title = re.sub(r'[\\/*?:[\]]', '_', nombre_ruta)
        
        if first_sheet:
            ws = default_sheet
            ws.title = safe_title[:30]
            first_sheet = False
        else:
            ws = wb.create_sheet(title=safe_title[:30])
            
        num_cols = len(df_r.columns)
        
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max(num_cols, 4))
        cell_t = ws.cell(row=1, column=1, value="DISTRIBUCIONES INESCO")
        cell_t.font = blue_title_font
        cell_t.alignment = Alignment(horizontal="center", vertical="center")
        
        subtitulo = f"Fecha de Entrega: {fecha_str} | Ruta: {nombre_ruta}"
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=max(num_cols, 4))
        cell_s = ws.cell(row=2, column=1, value=subtitulo)
        cell_s.font = sub_font
        cell_s.alignment = Alignment(horizontal="left", vertical="center")
        
        ws.row_dimensions[1].height = 25
        ws.row_dimensions[2].height = 18
        ws.row_dimensions[3].height = 10
        
        for col_idx, col_name in enumerate(df_r.columns, start=1):
            c = ws.cell(row=4, column=col_idx, value=col_name)
            c.fill = header_fill
            c.font = header_font
            c.alignment = Alignment(horizontal="center" if col_idx != 2 else "left", vertical="center")
        ws.row_dimensions[4].height = 22
        
        for row_idx, row_data in enumerate(df_r.values, start=5):
            val_first = str(row_data[0]).upper()
            is_total_row = ("TOTAL" in val_first or "SUBTOTAL" in val_first)
            
            for col_idx, val in enumerate(row_data, start=1):
                c = ws.cell(row=row_idx, column=col_idx, value=val)
                c.border = thin_border
                
                if is_total_row:
                    c.fill = total_fill
                    c.font = total_font
                
                if col_idx in [1, 3, 4]:
                    c.alignment = Alignment(horizontal="center", vertical="center")
                else:
                    c.alignment = Alignment(horizontal="left", vertical="center")
                    
        for col_idx in range(1, num_cols + 1):
            col_letter = get_column_letter(col_idx)
            max_len = 0
            for row_idx in range(4, 5 + len(df_r)):
                cell_val = ws.cell(row=row_idx, column=col_idx).value
                if cell_val:
                    max_len = max(max_len, len(str(cell_val)))
            ws.column_dimensions[col_letter].width = max(max_len + 5, 12)
            
    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

# ---------------------------------------------------------
# 4. PESTAÑAS Y NAVEGACIÓN
# ---------------------------------------------------------
tab1, tab2, tab3 = st.tabs(["📅 Fechas de Vencimiento", "📦 Administrar SKUs", "📄 Extracción PDF (Rutas)"])

# --- TAB 1: FECHAS DE VENCIMIENTO CON RESPALDO NUBE ---
with tab1:
    st.markdown('<p class="sub-title">➕ Agregar Registro de Vencimiento</p>', unsafe_allow_html=True)
    
    # Herramientas de respaldo rápido para la bodega
    with st.expander("📂 Opciones de Respaldo (Guardar / Cargar en Celular)"):
        if st.session_state.vencimientos:
            json_str = json.dumps(st.session_state.vencimientos, ensure_ascii=False, indent=4)
            st.download_button(
                label="📥 Descargar Archivo Respaldo (.json)",
                data=json_str,
                file_name=f"respaldo_vencimientos_{date.today()}.json",
                mime="application/json"
            )
        
        uploaded_backup = st.file_uploader("📤 Subir Respaldo Previo (.json)", type=["json"])
        if uploaded_backup is not None:
            try:
                data_recuperada = json.load(uploaded_backup)
                st.session_state.vencimientos = data_recuperada
                st.success("✅ ¡Datos restaurados exitosamente desde tu respaldo!")
                st.rerun()
            except Exception as e:
                st.error("⚠️ El archivo no es válido.")

    skus_opt = [f"{item['sku']} - {item['descripcion']}" for item in st.session_state.catalogo]
    
    c1, c2 = st.columns([2, 1])
    with c1:
        sel_sku = st.selectbox("Seleccionar Producto:", options=skus_opt, index=None, placeholder="🔎 Buscar SKU o Nombre...")
    with c2:
        f_venc = st.date_input("Fecha de Vencimiento:", value=date.today())
        
    if st.button("💾 Guardar Fecha de Vencimiento"):
        if sel_sku:
            s_code, s_desc = sel_sku.split(" - ", 1)
            
            existe_idx = next((i for i, r in enumerate(st.session_state.vencimientos) if r["SKU"] == s_code), None)
            
            if existe_idx is not None:
                st.session_state.vencimientos[existe_idx]["Fecha Vencimiento"] = f_venc.strftime("%d/%m/%Y")
                st.session_state.vencimientos[existe_idx]["created_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                st.warning(f"⚠️ El SKU {s_code} ya estaba registrado. Se actualizó su fecha a {f_venc.strftime('%d/%m/%Y')}.")
            else:
                st.session_state.vencimientos.append({
                    "id": len(st.session_state.vencimientos) + 1,
                    "SKU": s_code,
                    "Descripción del Producto": s_desc,
                    "Fecha Vencimiento": f_venc.strftime("%d/%m/%Y"),
                    "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                })
                st.success("✅ Registro guardado con éxito.")
            st.rerun()
        else:
            st.warning("⚠️ Debes seleccionar un SKU primero.")

    st.markdown('<p class="sub-title">📋 Registros Guardados</p>', unsafe_allow_html=True)
    
    if st.session_state.vencimientos:
        for idx, row in enumerate(st.session_state.vencimientos):
            col_a, col_b, col_c, col_d = st.columns([2, 4, 3, 2])
            col_a.write(f"**{row['SKU']}**")
            col_b.write(row['Descripción del Producto'])
            
            fecha_actual = datetime.strptime(row['Fecha Vencimiento'], "%d/%m/%Y").date()
            nueva_f = col_c.date_input("Fecha", value=fecha_actual, key=f"date_{row['id']}")
            st.session_state.vencimientos[idx]['Fecha Vencimiento'] = nueva_f.strftime("%d/%m/%Y")
            
            if col_d.button("❌ Borrar", key=f"del_{row['id']}"):
                st.session_state.vencimientos.pop(idx)
                st.rerun()

        st.divider()
        df_venc_out = pd.DataFrame(st.session_state.vencimientos)[["SKU", "Descripción del Producto", "Fecha Vencimiento"]]
        
        dict_venc = {"Vencimientos": df_venc_out}
        excel_bytes = exportar_excel_multiruta(dict_venc, fecha_str=date.today().strftime('%d/%m/%Y'))
        
        col_dl, col_sh = st.columns([1, 1])
        with col_dl:
            st.download_button(
                label="📥 Descargar Reporte en Excel",
                data=excel_bytes,
                file_name=f"Vencimientos_Inesco_{date.today()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        with col_sh:
            mensaje_wa = f"Reporte de Vencimientos Inesco - {date.today()}"
            st.markdown(f'[📲 Compartir por WhatsApp](https://api.whatsapp.com/send?text={mensaje_wa})')
    else:
        st.info("No hay registros guardados todavía.")

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

# --- TAB 3: EXTRACCIÓN PDF (RUTAS) ---
with tab3:
    st.markdown('<p class="sub-title">📄 Extracción de Rutas (ML3E51, ML3E52, ML3E53)</p>', unsafe_allow_html=True)
    
    uploaded_file = st.file_uploader("Cargar documento PDF con las planillas:", type=["pdf"])
    
    if uploaded_file:
        rutas_crudas = {"ML3E51": [], "ML3E52": [], "ML3E53": []}
        ruta_actual = "ML3E51"
        fecha_detectada = date.today().strftime('%d.%m.%Y')
        
        palabras_ignorar = ['REPARTIDOR', 'USUARIO', 'ESTATUS', 'TRANSPORTE', 'CAMION', 'CAMIÓN', 'ENTREGAS', 'SUBTOTAL', 'TOTAL']
        
        with pdfplumber.open(uploaded_file) as pdf:
            for page in pdf.pages:
                texto = page.extract_text() or ""
                lineas = texto.split('\n')
                
                for line in lineas:
                    line_clean = line.strip()
                    line_upper = line_clean.upper()
                    
                    if not line_clean:
                        continue
                        
                    if "FECHA" in line_upper:
                        match_f = re.search(r'(\d{2}[/.-]\d{2}[/.-]\d{4}|\d{4}[/.-]\d{2}[/.-]\d{2})', line_clean)
                        if match_f:
                            fecha_detectada = match_f.group(1)
                            
                    if "ML3E51" in line_upper or " 51" in line_upper and "RUTA" in line_upper:
                        ruta_actual = "ML3E51"
                        continue
                    elif "ML3E52" in line_upper or " 52" in line_upper and "RUTA" in line_upper:
                        ruta_actual = "ML3E52"
                        continue
                    elif "ML3E53" in line_upper or " 53" in line_upper and "RUTA" in line_upper:
                        ruta_actual = "ML3E53"
                        continue
                        
                    if any(p in line_upper for p in palabras_ignorar):
                        continue
                        
                    match_prod = re.search(r'^(\d{5,6})\s+(.+?)\s+(\d+)(?:\s*/\s*(\d+)|\s+(\d+))?\s*$', line_clean)
                    if match_prod:
                        sku = match_prod.group(1)
                        desc = match_prod.group(2).strip()
                        cajas = int(match_prod.group(3))
                        
                        if match_prod.group(4):
                            unidades = int(match_prod.group(4))
                        elif match_prod.group(5):
                            unidades = int(match_prod.group(5))
                        else:
                            unidades = 0
                            
                        if not any(d['SKU (Material)'] == sku for d in rutas_crudas[ruta_actual]):
                            rutas_crudas[ruta_actual].append({
                                "SKU (Material)": sku,
                                "Descripción del Producto": desc,
                                "Cajas": cajas,
                                "Unidades": unidades
                            })

        rutas_crudas = {k: v for k, v in rutas_crudas.items() if len(v) > 0}

        if rutas_crudas:
            st.success(f"✅ Extracción completada para {len(rutas_crudas)} ruta(s).")
            
            excel_dict = {}
            
            for nombre_ruta in ["ML3E51", "ML3E52", "ML3E53"]:
                if nombre_ruta in rutas_crudas:
                    items = rutas_crudas[nombre_ruta]
                    df_temp = pd.DataFrame(items)
                    
                    sum_cajas = int(df_temp["Cajas"].sum())
                    sum_unidades = int(df_temp["Unidades"].sum())
                    
                    df_final = pd.DataFrame()
                    df_final["SKU (Material)"] = df_temp["SKU (Material)"]
                    df_final["Descripción del Producto"] = df_temp["Descripción del Producto"]
                    df_final["Cantidad (Cajas)"] = df_temp["Cajas"]
                    df_final["Cantidad (Unidades)"] = df_temp["Unidades"]
                    
                    df_final.loc[len(df_final)] = {
                        "SKU (Material)": "TOTALES",
                        "Descripción del Producto": "SUMATORIA TOTAL CALCULADA",
                        "Cantidad (Cajas)": sum_cajas,
                        "Cantidad (Unidades)": sum_unidades
                    }
                    
                    excel_dict[nombre_ruta] = df_final
                    
                    st.markdown(f"### 🚚 Ruta: {nombre_ruta} ({len(items)} productos encontrados)")
                    
                    m1, m2 = st.columns(2)
                    m1.metric("📦 Total Cajas Calculadas", f"{sum_cajas:,}")
                    m2.metric("🍾 Total Unidades (Botellas) Calculadas", f"{sum_unidades:,}")
                    
                    st.dataframe(df_final, use_container_width=True)
                    st.divider()

            if excel_dict:
                excel_bytes_multiruta = exportar_excel_multiruta(excel_dict, fecha_str=fecha_detectada)
                
                st.download_button(
                    label="📥 Descargar Excel Unificado (Pestañas ML3E51, ML3E52, ML3E53)",
                    data=excel_bytes_multiruta,
                    file_name=f"Planillas_Rutas_Inesco_{date.today()}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        else:
            st.warning("⚠️ No se detectaron SKUs válidos. Comprueba que el PDF contenga texto seleccionable.")
