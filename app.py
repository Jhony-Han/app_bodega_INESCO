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
import urllib.parse

# ---------------------------------------------------------
# 1. CONFIGURACIÓN DE PÁGINA Y ESTILOS (DISEÑO CLÁSICO INESCO)
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
    .success-box {
        background-color: #D4EDDA;
        color: #155724;
        padding: 12px;
        border-radius: 8px;
        border-left: 5px solid #28A745;
        font-weight: bold;
        margin-bottom: 15px;
    }
    </style>
""", unsafe_allow_html=True)

# Encabezado clásico
col_h1, col_h2, col_h3 = st.columns([1, 6, 1])
with col_h1:
    st.markdown("<h1 style='text-align: center; font-size: 2.5rem; margin-top: 20px;'>🥤</h1>", unsafe_allow_html=True)
with col_h2:
    st.markdown("""
        <div class="cocacola-header">
            <h1>DISTRIBUCIONES INESCO</h1>
            <p>Gestión de Inventario, Vencimientos y Extracción por Rutas</p>
        </div>
    """, unsafe_allow_html=True)
with col_h3:
    st.markdown("<h1 style='text-align: center; font-size: 2.5rem; margin-top: 20px;'>🚚</h1>", unsafe_allow_html=True)

# ---------------------------------------------------------
# 2. GESTIÓN DE DATOS EN SESIÓN Y CATÁLOGO PERSISTENTE
# ---------------------------------------------------------
if "vencimientos" not in st.session_state:
    st.session_state.vencimientos = []

if "modo_captura" not in st.session_state:
    st.session_state.modo_captura = "Tomar datos manual"

if "voz_temp_input" not in st.session_state:
    st.session_state.voz_temp_input = ""

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
# 3. EXPORTADOR EXCEL PROFESIONAL (CON SUMATORIAS Y TÍTULO DE RUTAS)
# ---------------------------------------------------------
def exportar_excel_multiruta(rutas_dict, fecha_str, es_reporte_rutas=False):
    wb = Workbook()
    default_sheet = wb.active
    
    red_title_font = Font(color="FFFFFF", bold=True, size=13, name="Calibri")
    title_fill = PatternFill(start_color="C00000", end_color="C00000", fill_type="solid")
    
    sub_font = Font(color="333333", italic=True, size=10, name="Calibri")
    sub_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    
    header_fill = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True, size=11, name="Calibri")
    
    zebra_fill = PatternFill(start_color="F9FBFD", end_color="F9FBFD", fill_type="solid")
    white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    total_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    total_font = Font(bold=True, size=11, name="Calibri", color="000000")
    
    thin_border = Border(
        left=Side(style='thin', color='BFBFBF'), right=Side(style='thin', color='BFBFBF'),
        top=Side(style='thin', color='BFBFBF'), bottom=Side(style='thin', color='BFBFBF')
    )
    
    thick_bottom = Border(
        left=Side(style='thin', color='BFBFBF'), right=Side(style='thin', color='BFBFBF'),
        top=Side(style='thin', color='BFBFBF'), bottom=Side(style='double', color='000000')
    )
    
    first_sheet = True
    for nombre_ruta, df_r in rutas_dict.items():
        if not df_r.empty and "SKU" in df_r.columns:
            df_r = df_r.sort_values(by="SKU", key=lambda col: col.astype(str).str.zfill(10)).reset_index(drop=True)

        safe_title = re.sub(r'[\\/*?:[\]]', '_', nombre_ruta)
        ws = default_sheet if first_sheet else wb.create_sheet(title=safe_title[:30])
        first_sheet = False
            
        num_cols = len(df_r.columns)
        
        titulo_reporte = "DISTRIBUCIONES INESCO - REPORTE DE EXTRACCIÓN DE RUTAS" if es_reporte_rutas else "DISTRIBUCIONES INESCO - REPORTE DE VENCIMIENTOS"
        
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max(num_cols, 3))
        cell_t = ws.cell(row=1, column=1, value=titulo_reporte)
        cell_t.font = red_title_font
        cell_t.fill = title_fill
        cell_t.alignment = Alignment(horizontal="center", vertical="center")
        for col in range(1, max(num_cols, 3) + 1):
            ws.cell(row=1, column=col).border = thin_border
            ws.cell(row=1, column=col).fill = title_fill
        
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=max(num_cols, 3))
        cell_s = ws.cell(row=2, column=1, value=f"Fecha de Exportación: {fecha_str}    |    Sección: {nombre_ruta}")
        cell_s.font = sub_font
        cell_s.fill = sub_fill
        cell_s.alignment = Alignment(horizontal="center", vertical="center")
        for col in range(1, max(num_cols, 3) + 1):
            ws.cell(row=2, column=col).border = thin_border
            ws.cell(row=2, column=col).fill = sub_fill
        
        ws.row_dimensions[1].height = 28
        ws.row_dimensions[2].height = 22
        ws.row_dimensions[3].height = 10
        
        ws.row_dimensions[4].height = 24
        for col_idx, col_name in enumerate(df_r.columns, start=1):
            c = ws.cell(row=4, column=col_idx, value=col_name)
            c.fill = header_fill
            c.font = header_font
            c.border = thin_border
            c.alignment = Alignment(horizontal="center", vertical="center")
            
        for row_idx, row_data in enumerate(df_r.values, start=5):
            ws.row_dimensions[row_idx].height = 20
            is_even = (row_idx % 2 == 0)
            row_fill = zebra_fill if is_even else white_fill
            
            for col_idx, val in enumerate(row_data, start=1):
                c = ws.cell(row=row_idx, column=col_idx, value=val)
                c.border = thin_border
                c.fill = row_fill
                c.font = Font(name="Calibri", size=11)
                
                col_header_name = str(df_r.columns[col_idx-1]).lower()
                if "sku" in col_header_name or "fecha" in col_header_name or "cajas" in col_header_name or "unidades" in col_header_name:
                    c.alignment = Alignment(horizontal="center", vertical="center")
                    if "cajas" in col_header_name or "unidades" in col_header_name:
                        c.number_format = '#,##0'
                else:
                    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
                    
        # Fila de Totales si es reporte de rutas
        if es_reporte_rutas and not df_r.empty:
            last_row = 4 + len(df_r)
            total_row_idx = last_row + 1
            ws.row_dimensions[total_row_idx].height = 22
            
            for col_idx in range(1, num_cols + 1):
                c = ws.cell(row=total_row_idx, column=col_idx)
                c.fill = total_fill
                c.font = total_font
                c.border = thick_bottom
                
                col_header_name = str(df_r.columns[col_idx-1]).lower()
                if col_idx == 1:
                    c.value = "TOTALES"
                    c.alignment = Alignment(horizontal="center", vertical="center")
                elif "cajas" in col_header_name or "unidades" in col_header_name:
                    col_letter = get_column_letter(col_idx)
                    c.value = f"=SUM({col_letter}5:{col_letter}{last_row})"
                    c.alignment = Alignment(horizontal="center", vertical="center")
                    c.number_format = '#,##0'
                else:
                    c.value = ""

        for col_idx in range(1, num_cols + 1):
            col_letter = get_column_letter(col_idx)
            max_len = 0
            for row in range(4, ws.max_row + 1):
                cell_val = ws.cell(row=row, column=col_idx).value
                if cell_val:
                    max_len = max(max_len, len(str(cell_val)))
            ws.column_dimensions[col_letter].width = max(max_len + 8, 25)
            
    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

# ---------------------------------------------------------
# 4. PARSER INTELIGENTE MULTILÍNEA PARA EL PDF DE RUTAS
# ---------------------------------------------------------
def procesar_pdf_rutas(pdf_file):
    rutas_encontradas = {}
    ruta_actual = "General"
    datos_actuales = []
    
    sku_actual = None
    desc_partes = []
    
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            texto = page.extract_text()
            if not texto:
                continue
                
            lineas = texto.split('\n')
            for linea in lineas:
                linea_str = linea.strip()
                
                # Detectar cambio de ruta
                if "Ruta / No.de Carga:" in linea_str:
                    # Si había un producto pendiente por cerrar antes de cambiar de ruta
                    if sku_actual and desc_partes:
                        datos_actuales.append({
                            "SKU": sku_actual,
                            "Descripción del Producto": " ".join(desc_partes).replace('|', '').strip(),
                            "Cajas": 0,
                            "Unidades": 0
                        })
                        sku_actual = None
                        desc_partes = []

                    if datos_actuales and ruta_actual:
                        rutas_encontradas[ruta_actual] = pd.DataFrame(datos_actuales)
                    
                    match_ruta = re.search(r'ML3E5[1-3]', linea_str)
                    if match_ruta:
                        ruta_actual = match_ruta.group(0)
                    else:
                        match_gen = re.search(r'Ruta / No.de Carga:\s*([^\s]+)', linea_str)
                        ruta_actual = match_gen.group(1) if match_gen else "Ruta_Desconocida"
                    
                    datos_actuales = []
                    continue
                
                # Cortar lectura si llegamos a secciones de sumarios o materiales adicionales
                if "Materiales Adicionales" in linea_str or "Sub-Total Familia" in linea_str:
                    if sku_actual and desc_partes:
                        datos_actuales.append({
                            "SKU": sku_actual,
                            "Descripción del Producto": " ".join(desc_partes).replace('|', '').strip(),
                            "Cajas": 0,
                            "Unidades": 0
                        })
                        sku_actual = None
                        desc_partes = []
                    continue

                # 1. Caso línea en formato compacto: SKU | DESCRIPCIÓN | CANTIDAD (ej: 160187 | FUZE NEGRO... | $5/0$)
                match_compacto = re.search(r'^(\d{5,6})\s*\|\s*(.*?)\s*\|\s*\$?(\d+)/(\d+)', linea_str)
                if match_compacto:
                    if sku_actual and desc_partes: # Guardar pendiente anterior si lo hubiera
                        datos_actuales.append({
                            "SKU": sku_actual,
                            "Descripción del Producto": " ".join(desc_partes).replace('|', '').strip(),
                            "Cajas": 0,
                            "Unidades": 0
                        })
                    sku_actual = match_compacto.group(1)
                    desc = match_compacto.group(2).replace('|', '').strip()
                    cajas = int(match_compacto.group(3))
                    unidades = int(match_compacto.group(4))
                    
                    datos_actuales.append({
                        "SKU": sku_actual,
                        "Descripción del Producto": desc,
                        "Cajas": cajas,
                        "Unidades": unidades
                    })
                    sku_actual = None
                    desc_partes = []
                    continue

                # 2. Caso formato multilínea (PDF extrae SKU solo en una línea)
                if re.fullmatch(r'\d{5,6}', linea_str):
                    if sku_actual and desc_partes:
                        datos_actuales.append({
                            "SKU": sku_actual,
                            "Descripción del Producto": " ".join(desc_partes).replace('|', '').strip(),
                            "Cajas": 0,
                            "Unidades": 0
                        })
                    sku_actual = linea_str
                    desc_partes = []
                    continue

                # 3. Detectar línea de cantidad (ej: $48/0$ o 48/0) cuando ya tenemos un SKU activo
                match_qty = re.search(r'\$?(\d+)/(\d+)', linea_str)
                if match_qty and sku_actual:
                    cajas = int(match_qty.group(1))
                    unidades = int(match_qty.group(2))
                    desc_limpia = " ".join(desc_partes).replace('|', '').strip()
                    
                    datos_actuales.append({
                        "SKU": sku_actual,
                        "Descripción del Producto": desc_limpia,
                        "Cajas": cajas,
                        "Unidades": unidades
                    })
                    sku_actual = None
                    desc_partes = []
                    continue

                # 4. Si tenemos un SKU activo, las líneas intermedias corresponden a la descripción del producto
                if sku_actual:
                    # Limpiamos caracteres sobrantes de tuberías o espacios
                    texto_limpio = linea_str.replace('|', '').strip()
                    if texto_limpio:
                        desc_partes.append(texto_limpio)
                        
        if datos_actuales and ruta_actual:
            rutas_encontradas[ruta_actual] = pd.DataFrame(datos_actuales)
            
    return rutas_encontradas

# ---------------------------------------------------------
# 5. PESTAÑAS Y NAVEGACIÓN
# ---------------------------------------------------------
tab1, tab2, tab3 = st.tabs(["📅 Fechas de Vencimiento", "📦 Administrar SKUs", "📄 Extracción PDF (Rutas)"])

# --- TAB 1: FECHAS DE VENCIMIENTO ---
with tab1:
    st.markdown('<p class="sub-title">➕ Agregar Registro de Vencimiento</p>', unsafe_allow_html=True)
    
    with st.expander("📂 Opciones de Respaldo y Sincronización (JSON)"):
        datos_respaldo = {
            "vencimientos": st.session_state.vencimientos,
            "catalogo": st.session_state.catalogo
        }
        json_str = json.dumps(datos_respaldo, ensure_ascii=False, indent=4)
        st.download_button(
            label="📥 Descargar Archivo Respaldo Completo (.json)",
            data=json_str,
            file_name=f"respaldo_inesco_{date.today()}.json",
            mime="application/json"
        )
        
        uploaded_backup = st.file_uploader("📤 Subir Respaldo Previo (.json)", type=["json"], key="uploader_backup")
        if uploaded_backup is not None:
            if "last_uploaded_file" not in st.session_state or st.session_state.last_uploaded_file != uploaded_backup.name:
                try:
                    stringio = io.StringIO(uploaded_backup.getvalue().decode("utf-8"))
                    data_recuperada = json.load(stringio)
                    if isinstance(data_recuperada, dict) and "vencimientos" in data_recuperada:
                        st.session_state.vencimientos = data_recuperada["vencimientos"]
                        if "catalogo" in data_recuperada:
                            st.session_state.catalogo = data_recuperada["catalogo"]
                        st.session_state.last_uploaded_file = uploaded_backup.name
                        st.success("✅ ¡Datos y catálogo restaurados exitosamente!")
                        st.rerun()
                    elif isinstance(data_recuperada, list):
                        st.session_state.vencimientos = data_recuperada
                        st.session_state.last_uploaded_file = uploaded_backup.name
                        st.success("✅ ¡Vencimientos restaurados con éxito!")
                        st.rerun()
                except Exception as e:
                    st.error(f"⚠️ Error al leer el respaldo: {e}")

    st.session_state.modo_captura = st.selectbox(
        "🎛️ Selecciona el método de entrada de datos:",
        ["Tomar datos manual", "Tomar datos con voz"],
        index=0,
        key="select_modo_captura"
    )

    skus_opt = [f"{item['sku']} - {item['descripcion']}" for item in st.session_state.catalogo]

    if st.session_state.modo_captura == "Tomar datos con voz":
        st.markdown("""
            <div style="background-color: #F8D7DA; padding: 10px; border-radius: 8px; border-left: 5px solid #E41E2B; margin-bottom: 15px; color: #721C24;">
                <strong>🎙️ Modo Dictado / Búsqueda Rápida:</strong> Usa el icono de micrófono del teclado de tu celular o PC, o escribe el nombre del producto para filtrar y autocompletar el SKU.
            </div>
        """, unsafe_allow_html=True)
        
        filtro_voz = st.text_input("🎤 Dicta o escribe el producto:", value=st.session_state.voz_temp_input, placeholder="Ej: Coca-Cola 350, Brisa, Quatro...", key="input_voz_busqueda")
        st.session_state.voz_temp_input = filtro_voz

        if filtro_voz:
            skus_filtrados = [s for s in skus_opt if filtro_voz.lower() in s.lower()]
        else:
            skus_filtrados = skus_opt
    else:
        st.markdown("""
            <div style="background-color: #D1ECF1; padding: 10px; border-radius: 8px; border-left: 5px solid #0C5460; margin-bottom: 15px; color: #0C5460;">
                <strong>⌨️ Modo Manual Activo:</strong> Selecciona el producto directamente de la lista desplegable.
            </div>
        """, unsafe_allow_html=True)
        skus_filtrados = skus_opt

    c1, c2 = st.columns([2, 1])
    with c1:
        sel_sku = st.selectbox("Seleccionar Producto:", options=skus_filtrados, index=0 if len(skus_filtrados) == 1 else None, placeholder="🔎 Buscar SKU o Nombre...", key="select_sku_venc")
    with c2:
        fecha_seleccionada = st.date_input("Fecha Vencimiento:", value=date.today(), key="input_calendar_venc")
        
    if st.button("💾 Guardar Fecha de Vencimiento", key="btn_save_venc"):
        if sel_sku:
            s_code, s_desc = sel_sku.split(" - ", 1)
            fecha_formateada = fecha_seleccionada.strftime("%d/%m/%Y")

            existe_idx = next((i for i, r in enumerate(st.session_state.vencimientos) if r["SKU"] == s_code), None)
            
            if existe_idx is not None:
                st.session_state.vencimientos[existe_idx]["Fecha Vencimiento"] = fecha_formateada
                st.warning(f"⚠️ El SKU {s_code} ya estaba registrado. Se actualizó su fecha.")
            else:
                st.session_state.vencimientos.insert(0, {
                    "id": len(st.session_state.vencimientos) + 1,
                    "SKU": s_code,
                    "Descripción del Producto": s_desc,
                    "Fecha Vencimiento": fecha_formateada
                })
            
            st.toast("¡Guardado correctamente!", icon="✅")
            st.markdown('<div class="success-box">✅ ¡Guardado con éxito! El nuevo registro aparece de primero en la lista.</div>', unsafe_allow_html=True)
            st.rerun()
        else:
            st.warning("⚠️ Debes seleccionar un SKU primero.")

    st.markdown('<p class="sub-title">📋 Registros Guardados (El más reciente aparece de primero)</p>', unsafe_allow_html=True)
    
    if st.session_state.vencimientos:
        for idx, row in enumerate(st.session_state.vencimientos):
            col_a, col_b, col_c, col_d = st.columns([2, 4, 3, 2])
            col_a.write(f"**{row['SKU']}**")
            col_b.write(row['Descripción del Producto'])
            
            try:
                fecha_default_row = datetime.strptime(row['Fecha Vencimiento'], "%d/%m/%Y").date()
            except ValueError:
                fecha_default_row = date.today()

            nueva_f_date = col_c.date_input("Fecha", value=fecha_default_row, key=f"date_row_{row['SKU']}_{idx}")
            nueva_f_texto = nueva_f_date.strftime("%d/%m/%Y")
            
            if nueva_f_texto != row['Fecha Vencimiento']:
                st.session_state.vencimientos[idx]['Fecha Vencimiento'] = nueva_f_texto
            
            if col_d.button("❌ Borrar", key=f"del_row_{row['SKU']}_{idx}"):
                st.session_state.vencimientos.pop(idx)
                st.rerun()

        st.divider()
        df_venc_out = pd.DataFrame(st.session_state.vencimientos)[["SKU", "Descripción del Producto", "Fecha Vencimiento"]]
        excel_bytes = exportar_excel_multiruta({"Vencimientos": df_venc_out}, fecha_str=date.today().strftime('%d/%m/%Y'), es_reporte_rutas=False)
        
        col_dl1, col_dl2 = st.columns([1, 1])
        
        with col_dl1:
            st.download_button(
                label="📥 Descargar Reporte en Excel",
                data=excel_bytes,
                file_name=f"Vencimientos_Inesco_{date.today()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_venc_excel"
            )
            
        with col_dl2:
            texto_wa = f"Hola, comparto el reporte de vencimientos de Distribuciones Inesco del {date.today().strftime('%d/%m/%Y')}."
            texto_encoded = urllib.parse.quote(texto_wa)
            url_whatsapp = f"https://api.whatsapp.com/send?text={texto_encoded}"
            
            st.markdown(f"""
                <a href="{url_whatsapp}" target="_blank" style="text-decoration: none;">
                    <div style="background-color: #25D366; color: white; padding: 10px 15px; border-radius: 8px; text-align: center; font-weight: bold; font-size: 0.95rem; box-shadow: 0px 4px 10px rgba(37, 211, 102, 0.3);">
                        💬 Abrir WhatsApp con Aviso
                    </div>
                </a>
            """, unsafe_allow_html=True)
    else:
        st.info("No hay registros guardados todavía. Sube tu archivo de respaldo o agrega un producto.")

# --- TAB 2: ADMINISTRAR SKUS (PERSISTENTE) ---
with tab2:
    st.markdown('<p class="sub-title">⚙️ Agregar o Eliminar SKUs del Catálogo (Permanente)</p>', unsafe_allow_html=True)
    
    col_add1, col_add2 = st.columns([1, 2])
    with col_add1:
        nuevo_sku = st.text_input("Nuevo Código SKU:", key="input_new_sku")
    with col_add2:
        nueva_desc = st.text_input("Descripción del Producto:", key="input_new_desc")
        
    if st.button("➕ Agregar Nuevo SKU al Catálogo", key="btn_add_sku"):
        if nuevo_sku and nueva_desc:
            sku_limpio = nuevo_sku.strip()
            desc_limpia = nueva_desc.strip()
            existe = any(item['sku'] == sku_limpio for item in st.session_state.catalogo)
            if not existe:
                st.session_state.catalogo.append({"sku": sku_limpio, "descripcion": desc_limpia})
                st.success(f"✅ SKU {sku_limpio} agregado permanentemente. Recuerda descargar tu respaldo JSON para conservarlo.")
                st.rerun()
            else:
                st.warning("⚠️ Este código SKU ya se encuentra registrado en el catálogo.")
        else:
            st.error("Por favor completa tanto el SKU como la Descripción.")

    st.divider()
    st.write(f"**Catálogo Actual ({len(st.session_state.catalogo)} SKUs):**")
    
    for idx, item in enumerate(st.session_state.catalogo):
        c_k, c_d, c_b = st.columns([2, 5, 2])
        c_k.write(f"**{item['sku']}**")
        c_d.write(item['descripcion'])
        if c_b.button("🗑️ Eliminar", key=f"cat_del_{item['sku']}_{idx}"):
            st.session_state.catalogo.pop(idx)
            st.rerun()

# --- TAB 3: EXTRACCIÓN PDF (RUTAS) ---
with tab3:
    st.markdown('<p class="sub-title">📄 Extracción de Rutas (ML3E51, ML3E52, ML3E53)</p>', unsafe_allow_html=True)
    st.info("ℹ️ Sube el archivo PDF de cargue de FEMSA para extraer y separar automáticamente las Cajas y Unidades por cada ruta.")
    
    uploaded_pdf = st.file_uploader("📂 Seleccionar archivo PDF de rutas", type=["pdf"], key="uploader_pdf_rutas")
    
    if uploaded_pdf is not None:
        with st.spinner("Procesando y extrayendo datos de las rutas..."):
            try:
                dict_rutas = procesar_pdf_rutas(uploaded_pdf)
                
                if dict_rutas:
                    st.success(f"✅ ¡Se procesaron exitosamente {len(dict_rutas)} rutas del documento!")
                    
                    for r_nombre, r_df in dict_rutas.items():
                        with st.expander(f"Ruta: {r_nombre} ({len(r_df)} productos encontrados)"):
                            st.dataframe(r_df, use_container_width=True)
                            
                    # Generar Excel multiruta profesional con sumatorias de Cajas y Unidades
                    excel_rutas_bytes = exportar_excel_multiruta(dict_rutas, fecha_str=date.today().strftime('%d/%m/%Y'), es_reporte_rutas=True)
                    
                    st.divider()
                    col_pdf1, col_pdf2 = st.columns([1, 1])
                    
                    with col_pdf1:
                        st.download_button(
                            label="📥 Descargar Excel Consolidado de Rutas",
                            data=excel_rutas_bytes,
                            file_name=f"Rutas_Inesco_{date.today()}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="dl_rutas_excel"
                        )
                        
                    with col_pdf2:
                        texto_wa_rutas = f"Hola, comparto el reporte consolidado de extracción de rutas de Distribuciones Inesco del {date.today().strftime('%d/%m/%Y')}."
                        texto_encoded_rutas = urllib.parse.quote(texto_wa_rutas)
                        url_whatsapp_rutas = f"https://api.whatsapp.com/send?text={texto_encoded_rutas}"
                        
                        st.markdown(f"""
                            <a href="{url_whatsapp_rutas}" target="_blank" style="text-decoration: none;">
                                <div style="background-color: #25D366; color: white; padding: 10px 15px; border-radius: 8px; text-align: center; font-weight: bold; font-size: 0.95rem; box-shadow: 0px 4px 10px rgba(37, 211, 102, 0.3);">
                                    💬 Abrir WhatsApp con Aviso de Rutas
                                </div>
                            </a>
                        """, unsafe_allow_html=True)
                else:
                    st.warning("⚠️ No se detectaron estructuras de rutas válidas en el PDF.")
            except Exception as e:
                st.error(f"⚠️ Ocurrió un error al procesar el PDF: {e}")
