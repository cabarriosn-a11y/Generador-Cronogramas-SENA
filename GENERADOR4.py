import streamlit as st
from fpdf import FPDF
import pandas as pd
from datetime import date, timedelta
import math
import os
import io

# ==========================================
# 1. CONFIGURACIÓN GLOBAL Y ESTILOS
# ==========================================
st.set_page_config(page_title="SENA Pro CIEA 2026", layout="wide")

VERDE_SENA = "#39A900"
LOGO_SENA_URL = "https://oficinavirtualderadicacion.sena.edu.co/oficinavirtual/Resources/logoSenaNaranja.png"

# Estilos para Botón 3D Verde y botones de acción
st.markdown(f"""
    <style>
    /* Botón de Descarga 3D Verde */
    div.stDownloadButton > button {{
        background: linear-gradient(145deg, #2ecc71, #27ae60) !important;
        color: white !important;
        height: 4em !important;
        width: 100% !important;
        border-radius: 12px !important;
        font-weight: bold !important;
        font-size: 20px !important;
        border: none !important;
        box-shadow: 0px 6px 0px #1e8449, 0px 8px 15px rgba(0,0,0,0.2) !important;
        transition: all 0.1s ease !important;
    }}
    div.stDownloadButton > button:hover {{ transform: translateY(2px); }}

    /* Estilo para los títulos */
    .title-ciea {{ color: {VERDE_SENA}; text-align: center; font-weight: bold; }}
    </style>
    """, unsafe_allow_html=True)

# ==========================================
# 2. FUNCIONES DE APOYO (MOTORES)
# ==========================================
def asignar_instructor(aa, nombre_tecnico):
    mapa = {
        "240201528": "Matemáticas", "220201501": "Física", "220601501": "SST y Medio Ambiente",
        "240201533": "Emprendimiento", "240201524": "Comunicación", "240201526": "Ética",
        "210201501": "Derechos del Trabajo", "230101507": "Actividad Física", "220501046": "TIC",
        "240202501": "Bilingüismo"
    }
    texto_aa = str(aa)
    for cod, area in mapa.items():
        if cod in texto_aa: return f"Instructor de {area}"
    return nombre_tecnico

def limpiar(t):
    if not isinstance(t, str): return str(t)
    d = {"á":"a","é":"e","í":"i","ó":"o","ú":"u","ñ":"n","Á":"A","É":"E","Í":"I","Ó":"O","Ú":"U","Ñ":"N",
         "“":"\"","”":"\"","–":"-","—":"-","\r":"","\n":" ","\t":" "}
    for k, v in d.items(): t = t.replace(k, v)
    return t

def format_f(d): return d.strftime('%d/%m/%y')

# ==========================================
# 3. MENÚ DE NAVEGACIÓN LATERAL
# ==========================================
with st.sidebar:
    st.image(LOGO_SENA_URL, width=150)
    st.markdown("---")
    st.markdown("### 🛰️ MENÚ PRINCIPAL")
    opcion = st.radio("Seleccione herramienta:", ["📢 Anuncios para Zajuna", "📅 Cronogramas Técnicos"])
    st.markdown("---")
    st.info("C.I.E.A. - Regional Guajira")

# ==========================================
# 4. APLICATIVO: ANUNCIOS
# ==========================================
if opcion == "📢 Anuncios para Zajuna":
    st.markdown(f"<h1 class='title-ciea'>🌿 Generador de Anuncios Institucionales</h1>", unsafe_allow_html=True)
    
    with st.sidebar:
        st.header("⚙️ Ajustes del Anuncio")
        tipo = st.selectbox("Categoría:", ["1. Bienvenida y Presentación", "2. Información General / Cambios", "3. Evidencias y Actividades", "4. Sondeos, Evaluaciones y Recursos", "5. Recordatorio de Fechas / Cronograma", "6. Orientaciones y Aclaraciones", "7. Sesión en Línea (Sincrónica)", "8. Cuadro de Honor (Felicitaciones)"])
        prog_an = st.text_input("Programa", "Coordinación de Procesos Logísticos", key="an_prog")
        fich_an = st.text_input("Ficha", "2874563", key="an_fich")
        inst_an = st.text_input("Instructor", "Carlos Barrios", key="an_inst")

    col_in, col_pre = st.columns([1, 1.3])
    with col_in:
        st.subheader("🖋️ Contenido")
        if "1." in tipo: titulo, icono, contenido = "BIENVENIDA AL PROGRAMA", "👋", st.text_area("Mensaje:", "Es un placer saludarlos e iniciar este proceso de formación...")
        elif "7." in tipo:
            titulo, icono = "SESIÓN EN LÍNEA (SINCRÓNICA)", "💻"
            tema_s = st.text_input("Tema de la Sesión", "Explicación de la Guía 2")
            fecha_s, hora_s = st.date_input("Fecha"), st.time_input("Hora")
            link_s = st.text_input("URL de Teams/Meet", "https://teams.microsoft.com/...")
            contenido = f"Tema: {tema_s}\nFecha: {fecha_s.strftime('%d/%m/%Y')}\nHora: {hora_s.strftime('%H:%M')}"
        elif "8." in tipo:
            titulo, icono = "CUADRO DE HONOR", "🏆"
            nombres_h = st.text_area("Aprendices:", "Juan Pérez\nMaría López")
            motivo_h = st.text_input("Motivo", "Excelente participación")
        else:
            titulo, icono = "INFORMACIÓN IMPORTANTE", "ℹ️"
            contenido = st.text_area("Novedad:", "Se informa el siguiente ajuste...")
            if "3." in tipo or "5." in tipo: fecha_an = st.date_input("Fecha Límite")

    # Lógica HTML
    if "7." in tipo:
        cuerpo_inner = f"""<div style='background:#f0f9f0; padding:15px; border-radius:10px; border:1px solid {VERDE_SENA}; text-align:center;'>
        <p><strong>{tema_s}</strong></p><p>📅 {fecha_s.strftime('%d/%m/%Y')} | ⏰ {hora_s.strftime('%H:%M')}</p>
        <br><a href='{link_s}' target='_blank' style='background:{VERDE_SENA}; color:white; padding:12px 25px; text-decoration:none; border-radius:5px; font-weight:bold; display:inline-block;'>INGRESAR A LA SESIÓN</a></div>"""
    elif "8." in tipo:
        lista_n = "".join([f"• {n.strip().upper()}<br>" for n in nombres_h.split('\n') if n.strip()])
        cuerpo_inner = f"""<div style='text-align:center; border:3px double {VERDE_SENA}; padding:20px; background:#f9fff9;'><h2 style='color:{VERDE_SENA};'>🏆 ¡FELICITACIONES! 🏆</h2><div style='font-size:1.3em; font-weight:bold; margin:15px 0;'>{lista_n}</div><p>Motivo: <strong>{motivo_h}</strong></p></div>"""
    else:
        lista_items = "".join([f"<li>✅ {i.strip()}</li>" for i in contenido.split('\n') if i.strip()])
        cuerpo_inner = f"<ul style='line-height:1.7;'>{lista_items}</ul>"
        if "3." in tipo or "5." in tipo: cuerpo_inner += f"<p style='background:#e8f5e9; padding:10px; border-radius:5px;'><strong>📅 Límite:</strong> {fecha_an.strftime('%d/%m/%Y')}</p>"

    html_final = f"""<div style="font-family: Arial, sans-serif; max-width: 750px; margin: auto; border: 1px solid #ddd; border-radius: 15px; background: white; overflow: hidden; box-shadow: 0 8px 16px rgba(0,0,0,0.1);">
    <div style="background: {VERDE_SENA}; padding: 25px; text-align: center;"><img src="{LOGO_SENA_URL}" width="110" style="background: white; border-radius: 8px; padding: 5px;"><h2 style="color: white; margin: 15px 0 0 0; text-transform: uppercase;">{icono} {titulo}</h2></div>
    <div style="padding: 30px; color: #333;"><p style="font-size: 1.1em;">Estimados aprendices de la ficha <strong>{fich_an}</strong> ({prog_an}):</p><div style="margin: 20px 0;">{cuerpo_inner}</div><p style="font-size: 0.9em; color: #666; font-style: italic;">Por favor, estar atentos a la plataforma Zajuna.</p></div>
    <div style="background: #fdfdfd; padding: 20px; border-top: 1px solid #eee; text-align: right;"><strong style="color: {VERDE_SENA};">{inst_an}</strong><br><small>Instructor SENA | Regional Guajira</small></div></div>"""

    with col_pre:
        st.subheader("👁️ Vista Previa Real")
        st.markdown(html_final, unsafe_allow_html=True)
        st.divider()
        st.subheader("📋 Código para Zajuna")
        st.code(html_final, language="html")
        if "8." in tipo: st.balloons()

# ==========================================
# 5. APLICATIVO: CRONOGRAMAS
# ==========================================
elif opcion == "📅 Cronogramas Técnicos":
    st.markdown(f"<h1 class='title-ciea'>📅 Cronogramas con Identidad CIEA</h1>", unsafe_allow_html=True)
    
    # --- FUNCIONES CRONOGRAMA ---
    def proximo_valido(f, festivs, vaca_activa):
        cursor = f
        while True:
            if cursor.weekday() == 6 or cursor in festivs: cursor += timedelta(days=1); continue
            if vaca_activa and ((cursor.month == 12 and cursor.day >= 15) or (cursor.month == 1) or (cursor.month == 2 and cursor.day <= 2)):
                cursor = date(cursor.year + (1 if cursor.month == 12 else 0), 2, 3); continue
            break
        return cursor

    def generar_excel_general_pro(datos, programa, horas_dict):
        output = io.BytesIO()
        df = pd.DataFrame(datos)
        f_info = df.groupby('Fase').agg({'Inicio': 'min', 'Fin': 'max'}).to_dict('index')
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_export = df[['Fase', 'Actividad_Proyecto', 'Actividad_Aprendizaje', 'RAP', 'Evidencia', 'Instructor']].copy()
            df_export['Tipo'] = df_export['Evidencia'].apply(lambda x: "Conocimiento" if "conocimiento" in str(x).lower() else ("Desempeño" if "desempeño" in str(x).lower() else "Producto"))
            df_export.to_excel(writer, sheet_name='Cronograma General', index=False, startrow=4)
            wb = writer.book; ws = writer.sheets['Cronograma General']
            fmt_h = wb.add_format({'bold': True, 'bg_color': '#39A900', 'font_color': 'white', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
            fmt_m = wb.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'text_wrap': True, 'font_size': 9})
            if os.path.exists("logo_sena.png"): ws.insert_image('A1', 'logo_sena.png', {'x_scale': 0.15, 'y_scale': 0.15})
            ws.merge_range('C1:I1', 'CENTRO INDUSTRIAL Y DE ENERGIAS ALTERNATIVAS - REGIONAL GUAJIRA', wb.add_format({'bold': True, 'size': 12, 'align': 'center'}))
            ws.merge_range('C2:I2', f'PROGRAMA: {programa}', wb.add_format({'bold': True, 'align': 'center'}))      
            ws.merge_range('C3:I3', 'CRONOGRAMA GENERAL', wb.add_format({'bold': True, 'size': 11,'align': 'center'}))
            for col, val in enumerate(df_export.columns): ws.write(4, col, val, fmt_h)
            ws.write('H5', 'Duración (Hrs)', fmt_h); ws.write('I5', 'Fechas Fase', fmt_h)
            row = 5
            for fase in df['Fase'].unique():
                n = len(df[df['Fase'] == fase])
                ws.merge_range(row, 0, row+n-1, 0, fase, fmt_m)
                ws.merge_range(row, 7, row+n-1, 7, f"{horas_dict.get(fase, 0)} Horas", fmt_m)
                ws.merge_range(row, 8, row+n-1, 8, f"{format_f(f_info[fase]['Inicio'])} al {format_f(f_info[fase]['Fin'])}", fmt_m)
                row += n
            ws.set_column('A:I', 25)
        return output.getvalue()

    class PDF_CIEA_FASE(FPDF):
        def __init__(self, programa, fase, inicio, fin, **kwargs):
            super().__init__(**kwargs)
            self.prog = programa; self.fase_n = fase; self.f_i = inicio; self.f_f = fin
        def header(self):
            if os.path.exists("logo_sena.png"): self.image("logo_sena.png", 10, 8, 22)
            self.set_font('Arial', 'B', 11)
            self.cell(0, 5, f'CENTRO INDUSTRIAL Y DE ENERGIAS ALTERNATIVAS - REGIONAL GUAJIRA', 0, 1, 'C')
            self.set_font('Arial', '', 9)
            self.cell(0, 5, f'CRONOGRAMA DE ACTIVIDADES - FASE {limpiar(self.fase_n)}', 0, 1, 'C')
            self.cell(0, 5, f'PROGRAMA: {limpiar(self.prog)}', 0, 1, 'C')
            self.set_font('Arial', 'B', 8); self.cell(0, 5, f'PERIODO: {self.f_i} AL {self.f_f}', 0, 1, 'C')
            self.ln(10)

    def generar_pdf_fase_ciea(datos, programa, fase_nombre):
        f_i, f_f = format_f(datos[0]['Inicio']), format_f(datos[-1]['Fin'])
        pdf = PDF_CIEA_FASE(programa, fase_nombre, f_i, f_f, orientation='L')
        pdf.add_page()
        w = [15, 25, 45, 50, 65, 40, 18, 18]
        pdf.set_fill_color(57, 169, 0); pdf.set_text_color(255); pdf.set_font('Arial', 'B', 8)
        for i, l in enumerate(["Fase", "AP", "AA", "RAP", "Evidencia", "Responsable", "Inicio", "Fin"]): pdf.cell(w[i], 8, l, 1, 0, 'C', True)
        pdf.ln(); pdf.set_text_color(0); pdf.set_font('Arial', '', 7)
        for r in datos:
            txts = [limpiar(r['Fase']), limpiar(r['Actividad_Proyecto']), limpiar(r['Actividad_Aprendizaje']), limpiar(r['RAP']), limpiar(r['Evidencia']), limpiar(r['Instructor']), format_f(r['Inicio']), format_f(r['Fin'])]
            # Cálculo de altura de fila para que NO se monten las celdas
            h_row = max(max([pdf.get_string_width(t) / (w[i]-2) for i, t in enumerate(txts)]), 1) * 4
            h_row = max(h_row, 8) 
            if pdf.get_y() + h_row > 180: pdf.add_page()
            y_a = pdf.get_y()
            for i, t in enumerate(txts):
                pdf.set_xy(10 + sum(w[:i]), y_a); pdf.multi_cell(w[i], 2, t, 0, 'L')
                pdf.set_xy(10 + sum(w[:i]), y_a); pdf.cell(w[i], h_row, "", 1)
            pdf.set_y(y_a + h_row)
        return pdf.output(dest='S').encode('latin-1')

    with st.sidebar:
        archivo = st.file_uploader("Suba su Excel (Fase, AP, AA, RAP, Evidencia)", type=["xlsx"], key="cr_file")
        c_sem, c_hrs = {}, {}
        if archivo:
            df_t = pd.read_excel(archivo)
            for f in df_t['Fase'].unique():
                st.write(f"**Fase: {f}**")
                c1, c2 = st.columns(2)
                c_sem[f] = c1.number_input(f"Semanas:", 1, 50, 4, key=f"s_{f}")
                c_hrs[f] = c2.number_input(f"Horas:", 1, 2000, 40, key=f"h_{f}")
        f_ini_in = st.date_input("Inicio de Lectiva", date.today(), key="cr_date")
        inst_nom = st.text_input("Instructor Técnico", "Carlos Barrios", key="cr_inst")
        vaca_tog = st.toggle("Omitir Vacaciones", value=True)
        festivs_list = st.multiselect("Festivos:", pd.date_range(start=date.today(), periods=365).date.tolist())

    if archivo:
        df_raw = pd.read_excel(archivo)
        if st.button("🚀 PROCESAR CRONOGRAMA"):
            res = []; cursor = f_ini_in
            for f in df_raw['Fase'].unique():
                items = df_raw[df_raw['Fase'] == f].to_dict('records')
                sem = c_sem[f]; avg = len(items)/sem; idx = 0
                for s in range(sem):
                    count = math.ceil((s+1)*avg) - math.ceil(s*avg)
                    lote = items[idx:idx+count]; idx += count
                    cursor = proximo_valido(cursor, festivs_list, vaca_tog)
                    f_i_s = cursor; f_f_s = proximo_valido(f_i_s + timedelta(days=6), festivs_list, vaca_tog)
                    for item in lote:
                        item.update({"Inicio": f_i_s, "Fin": f_f_s, "Instructor": asignar_instructor(item['Actividad_Aprendizaje'], inst_nom)})
                        res.append(item)
                    cursor = f_f_s + timedelta(days=1)
            st.session_state.crono = res; st.session_state.horas = c_hrs; st.session_state.prog = archivo.name.split('.')[0].upper()

    if 'crono' in st.session_state:
        st.dataframe(pd.DataFrame(st.session_state.crono), use_container_width=True)
        st.subheader("📥 Centro de Descargas")
        col1, col2 = st.columns(2)
        with col1:
            ex_gen = generar_excel_general_pro(st.session_state.crono, st.session_state.prog, st.session_state.horas)
            st.download_button("📥 EXCEL GENERAL (.XLSX)", data=ex_gen, file_name=f"Cronograma_{st.session_state.prog}.xlsx")
        with col2:
            st.write("📄 **PDFs por Fase:**")
            df_r = pd.DataFrame(st.session_state.crono)
            fases = df_r['Fase'].unique()
            for f in fases:
                df_f = [x for x in st.session_state.crono if x['Fase'] == f]
                pdf_f = generar_pdf_fase_ciea(df_f, st.session_state.prog, f)
                st.download_button(f"Fase {f}", data=pdf_f, file_name=f"Fase_{f}.pdf")
