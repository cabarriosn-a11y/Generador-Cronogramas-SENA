import streamlit as st
from fpdf import FPDF
import pandas as pd
from datetime import date, timedelta
import math
import os
import io

# --- 1. MOTOR DE ASIGNACIÓN DE INSTRUCTORES ---
def asignar_instructor(aa, nombre_tecnico):
    mapa = {
        "240201528": "Matemáticas", "220201501": "Física", "220601501": "SST y Medio Ambiente",
        "240201533": "Emprendimiento", "240201524": "Comunicación", "240201526": "Ética",
        "210201501": "Derechos del Trabajo", "230101507": "Actividad Física", "220501046": "TIC",
        "240202501": "Bilingüismo"
    }
    texto_aa = str(aa)
    for cod, area in mapa.items():
        if cod in texto_aa:
            return f"Instructor de {area}"
    return nombre_tecnico

# --- 2. UTILIDADES DE FORMATO ---
def limpiar(t):
    if not isinstance(t, str): return str(t)
    d = {"á":"a","é":"e","í":"i","ó":"o","ú":"u","ñ":"n","Á":"A","É":"E","Í":"I","Ó":"O","Ú":"U","Ñ":"N",
         "“":"\"","”":"\"","–":"-","—":"-","\r":"","\n":" ","\t":" "}
    for k, v in d.items(): t = t.replace(k, v)
    return t

def format_f(d):
    return d.strftime('%d/%m/%y')

def extraer_tipo(evidencia):
    ev = evidencia.lower()
    if "conocimiento" in ev: return "Conocimiento"
    if "desempeño" in ev or "desempeno" in ev: return "Desempeño"
    if "producto" in ev: return "Producto"
    return "Evidencia"

# --- 3. MOTOR DE FECHAS ---
def proximo_valido(f, festivs, vaca_activa):
    cursor = f
    while True:
        if cursor.weekday() == 6 or cursor in festivs:
            cursor += timedelta(days=1); continue
        if vaca_activa and ((cursor.month == 12 and cursor.day >= 15) or (cursor.month == 1) or (cursor.month == 2 and cursor.day <= 2)):
            cursor = date(cursor.year + (1 if cursor.month == 12 else 0), 2, 3); continue
        break
    return cursor

# --- 4. EXCEL GENERAL (CON LOGO Y CELDAS UNIDAS) ---
def generar_excel_general_pro(datos, programa, horas_dict):
    output = io.BytesIO()
    df = pd.DataFrame(datos)
    f_info = df.groupby('Fase').agg({'Inicio': 'min', 'Fin': 'max'}).to_dict('index')
    
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_export = df[['Fase', 'Actividad_Proyecto', 'Actividad_Aprendizaje', 'RAP', 'Evidencia', 'Instructor']].copy()
        df_export['Tipo'] = df_export['Evidencia'].apply(extraer_tipo)
        columnas = ['Fase', 'Actividad_Proyecto', 'Actividad_Aprendizaje', 'RAP', 'Tipo', 'Evidencia', 'Instructor']
        df_export = df_export[columnas]
        
        df_export.to_excel(writer, sheet_name='Cronograma General', index=False, startrow=4)
        workbook  = writer.book
        worksheet = writer.sheets['Cronograma General']
        
        # Formatos
        fmt_h = workbook.add_format({'bold': True, 'bg_color': '#39A900', 'font_color': 'white', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
        fmt_m = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'text_wrap': True, 'font_size': 9})

        # Encabezado e Imagen
        if os.path.exists("logo_sena.png"):
            worksheet.insert_image('A1', 'logo_sena.png', {'x_scale': 0.15, 'y_scale': 0.15, 'x_offset': 10, 'y_offset': 5})
        
        worksheet.merge_range('C1:I1', 'CRONOGRAMA GENERAL DE ACTIVIDADES', workbook.add_format({'bold': True, 'size': 14, 'align': 'center'}))
        worksheet.merge_range('C2:I2', f'PROGRAMA: {programa}', workbook.add_format({'bold': True, 'size': 11, 'align': 'center'}))
        worksheet.merge_range('C3:I3', 'CENTRO INDUSTRIAL Y DE ENERGIAS ALTERNATIVAS - REGIONAL GUAJIRA', workbook.add_format({'size': 10, 'align': 'center'}))

        # Columnas Adicionales
        worksheet.write('H5', 'Duración (Hrs)', fmt_h)
        worksheet.write('I5', 'Fechas Fase', fmt_h)
        for col, val in enumerate(df_export.columns): worksheet.write(4, col, val, fmt_h)

        # Merge de Fases
        row = 5
        for fase in df['Fase'].unique():
            n = len(df[df['Fase'] == fase])
            fin = row + n - 1
            worksheet.merge_range(row, 0, fin, 0, fase, fmt_m)
            worksheet.merge_range(row, 7, fin, 7, f"{horas_dict.get(fase, 0)} Horas", fmt_m)
            txt_f = f"{format_f(f_info[fase]['Inicio'])} al {format_f(f_info[fase]['Fin'])}"
            worksheet.merge_range(row, 8, fin, 8, txt_f, fmt_m)
            row += n

        worksheet.set_column('A:I', 22)
    return output.getvalue()

# --- 5. PDF OPERATIVO (CON LOGO) ---
class PDF_SENA_FASE(FPDF):
    def __init__(self, programa, fase, inicio, fin, **kwargs):
        super().__init__(**kwargs)
        self.prog = programa; self.fase_n = fase; self.f_i = inicio; self.f_f = fin

    def header(self):
        if os.path.exists("logo_sena.png"):
            self.image("logo_sena.png", 10, 8, 22)
        self.set_font('Arial', 'B', 11)
        self.cell(0, 5, f'CRONOGRAMA DE ACTIVIDADES - FASE {limpiar(self.fase_n)}', 0, 1, 'C')
        self.set_font('Arial', '', 9)
        self.cell(0, 5, f'PROGRAMA: {limpiar(self.prog)}', 0, 1, 'C')
        self.cell(0, 5, 'CENTRO INDUSTRIAL Y DE ENERGIAS ALTERNATIVAS - REGIONAL GUAJIRA', 0, 1, 'C')
        self.ln(2); self.set_font('Arial', 'B', 8)
        self.cell(0, 5, f'PERIODO: {self.f_i} AL {self.f_f}', 0, 1, 'C')
        self.ln(10)

def generar_pdf_fase(datos, programa, fase_nombre):
    f_i, f_f = format_f(datos[0]['Inicio']), format_f(datos[-1]['Fin'])
    pdf = PDF_SENA_FASE(programa=programa, fase=fase_nombre, inicio=f_i, fin=f_f, orientation='L')
    pdf.add_page()
    w = [20, 25, 38, 38, 55, 45, 27, 27] 
    h_labels = ["Fase", "AP", "AA", "RAP", "Evidencia", "Responsable", "Inicio", "Fin"]
    pdf.set_fill_color(57, 169, 0); pdf.set_text_color(255, 255, 255); pdf.set_font('Arial', 'B', 7)
    for i, l in enumerate(h_labels): pdf.cell(w[i], 8, l, 1, 0, 'C', True)
    pdf.ln(); pdf.set_text_color(0, 0, 0); pdf.set_font('Arial', '', 6)
    for r in datos:
        txts = [limpiar(r['Fase']), limpiar(r['Actividad_Proyecto']), limpiar(r['Actividad_Aprendizaje']), 
                limpiar(r['RAP']), limpiar(r['Evidencia']), limpiar(r['Instructor']), format_f(r['Inicio']), format_f(r['Fin'])]
        h_row = max(max([math.ceil(pdf.get_string_width(t)/(w[i]-2.5)) * 4 for i, t in enumerate(txts)]), 8)
        if pdf.get_y() + h_row > 185: pdf.add_page()
        y_a = pdf.get_y()
        for i, t in enumerate(txts):
            pdf.set_xy(10 + sum(w[:i]), y_a); pdf.multi_cell(w[i], 3.8, t, 1, 'L')
            pdf.set_xy(10 + sum(w[:i]), y_a); pdf.cell(w[i], h_row, '', 1)
        pdf.set_y(y_a + h_row)
    return pdf.output(dest='S').encode('latin-1')

# --- 6. INTERFAZ ---
st.set_page_config(page_title="SENA Pro 2026", layout="wide")
st.title("🛡️ Estación de Mando: Cronogramas con Identidad SENA")

with st.sidebar:
    archivo = st.file_uploader("Suba su Excel (5 Columnas)", type=["xlsx"])
    config_semanas = {}; config_horas = {}
    if archivo:
        df_t = pd.read_excel(archivo)
        for f in df_t['Fase'].unique():
            st.write(f"**Fase: {f}**")
            c1, c2 = st.columns(2)
            config_semanas[f] = c1.number_input(f"Semanas:", 1, 50, 4, key=f"s_{f}")
            config_horas[f] = c2.number_input(f"Horas:", 1, 2000, 40, key=f"h_{f}")
    f_ini_in = st.date_input("Inicio de Lectiva", date.today())
    inst_nombre = st.text_input("Instructor Técnico", "Carlos Barrios")
    vaca = st.toggle("Omitir Vacaciones", value=True)
    festivs = st.multiselect("Festivos:", pd.date_range(start=date.today(), periods=365).date.tolist())

if archivo:
    df_raw = pd.read_excel(archivo)
    if st.button("🚀 PROCESAR TODO"):
        res = []; cursor = f_ini_in
        for f in df_raw['Fase'].unique():
            items = df_raw[df_raw['Fase'] == f].to_dict('records')
            sem = config_semanas[f]; avg = len(items) / sem; idx = 0
            for s in range(sem):
                count = math.ceil((s + 1) * avg) - math.ceil(s * avg)
                lote = items[idx:idx+count]; idx += count
                cursor = proximo_valido(cursor, festivs, vaca)
                f_i_s = cursor; f_f_s = proximo_valido(f_i_s + timedelta(days=6), festivs, vaca)
                for item in lote:
                    item.update({"Inicio": f_i_s, "Fin": f_f_s, "Instructor": asignar_instructor(item['Actividad_Aprendizaje'], inst_nombre)})
                    res.append(item)
                cursor = f_f_s + timedelta(days=1)
        st.session_state.crono = res; st.session_state.horas = config_horas
        st.session_state.prog = archivo.name.split('.')[0].upper()

if 'crono' in st.session_state:
    st.dataframe(pd.DataFrame(st.session_state.crono), use_container_width=True)
    st.divider()
    
    st.subheader("📥 Centro de Descargas")
    col1, col2 = st.columns(2)
    with col1:
        st.success("✅ **Excel General (Estratégico)**")
        ex_gen = generar_excel_general_pro(st.session_state.crono, st.session_state.prog, st.session_state.horas)
        st.download_button("📥 DESCARGAR GENERAL (.XLSX)", data=ex_gen, file_name=f"General_{st.session_state.prog}.xlsx")
    with col2:
        st.write("📂 **Cronogramas por Fase (PDF)**")
        df_r = pd.DataFrame(st.session_state.crono)
        fases = df_r['Fase'].unique()
        btns = st.columns(len(fases))
        for i, f in enumerate(fases):
            df_f = [x for x in st.session_state.crono if x['Fase'] == f]
            pdf_f = generar_pdf_fase(df_f, st.session_state.prog, f)
            btns[i].download_button(f"📄 Fase {f}", data=pdf_f, file_name=f"Fase_{f}.pdf")