# app.py
from flask import Flask, render_template, request, send_file, abort
from collections import defaultdict
from fpdf import FPDF
from datetime import datetime, timedelta
import pandas as pd
import os

app = Flask(__name__)

# = = = DEFINICION DE SUB PROCESOS = = =
def limpiar_texto(texto):
    if not isinstance(texto, str):
        return texto
    return (
        texto.replace("–", "-")
             .replace("—", "-")
             .replace("“", '"')
             .replace("”", '"')
             .replace("‘", "'")
             .replace("’", "'")
             .replace("…", "...")
             .replace("á", "a").replace("é", "e").replace("í", "i").replace("ó", "o").replace("ú", "u")
             .replace("Á", "A").replace("É", "E").replace("Í", "I").replace("Ó", "O").replace("Ú", "U")
             .replace("ñ", "n").replace("Ñ", "N")
    )

# Helper: nombres de meses en español (minuscula)
MESES_ES = {
    1: "enero", 2: "febrero", 3: "marzo", 4: "abril",
    5: "mayo", 6: "junio", 7: "julio", 8: "agosto",
    9: "setiembre", 10: "octubre", 11: "noviembre", 12: "diciembre"
}

# RUTA DEL FORMULARIO Y PROCESO
@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        # === INPUTS desde el formulario ===
        try:
            ubicacion_input = int(request.form.get("ubicacion"))
        except Exception:
            return "Número de instalación inválido", 400

        ingeniero = request.form.get("ingeniero", "CT")
        fecha_inspeccion_str = request.form.get("fecha_inspeccion")  # esperado 'YYYY-MM-DD' desde <input type="date">

        # validar fecha_inspeccion
        try:
            fecha_dt = datetime.strptime(fecha_inspeccion_str, "%Y-%m-%d")
        except Exception:
            return "Fecha de inspección inválida. Use formato de calendario.", 400

        # Calcular fecha_emision = fecha_inspeccion + 7 dias
        fecha_emision_dt = fecha_dt + timedelta(days=7)
        # Formato textual igual al ejemplo: "Lima, 19 de julio de 2022"
        mes_text = MESES_ES.get(fecha_emision_dt.month, fecha_emision_dt.strftime("%B").lower())
        fecha_emision = f"Lima, {fecha_emision_dt.day} de {mes_text} de {fecha_emision_dt.year}"

        # considerar_nuevo_formato: 0 si año de emisión == 2025, else 1
        considerar_nuevo_formato = 0 if fecha_emision_dt.year == 2025 else 1

        # === LEER EXCEL igual que tu script original ===
        excel_path = "data_base.xlsm"
        if not os.path.exists(excel_path):
            return "No se encontró el archivo data_base.xlsm en el directorio.", 500

        try:
            df = pd.read_excel(excel_path, sheet_name="DATA")
        except Exception as e:
            return f"Error leyendo data_base.xlsm: {e}", 500

        # === FILTRAR DATOS DE LA UBICACIÓN ===
        df_filtrado = df[df["Ubicacion"] == ubicacion_input]
        if df_filtrado.empty:
            return f"No se ha podido obtener el dato para la ubicación técnica {ubicacion_input}.", 404

        cliente = df_filtrado.iloc[0][("Nombre Titular")]
        direccion = df_filtrado.iloc[0]["Direccion"]
        texto_adicional = """(1) En aras de mantener el estricto cumplimiento del marco normativo vigente, y en caso requieran hacer alguna modificación en el área que se encuentra alrededor de la zona de almacenamiento de tanques de GLP como: instalar equipos eléctricos, cámaras de seguridad, iluminación, tomacorrientes, edificaciones contiguas, equipos de aire acondicionado, cableado eléctrico, ductos, sumideros, almacenar materiales, construcción de muros perimetrales, etc. tienen la obligación de comunicar previamente al equipo técnico de Solgas lo pertinente, a fin de que puedan recibir la asesoría técnica y validación correspondiente, de tal manera que, se evite incurrir en algunos incumplimientos normativos que puedan ser materia de suspensión de la Ficha de Registro y sanciones pecuniarias por parte del ente fiscalizador, así como evitar cualquier riesgo innecesario en la instalación; cabe precisar que en caso de no informar acerca de las modificaciones que realicen en la instalación, usted será el único y exclusivo responsable por las consecuencias que se deriven de su accionar."""

        # === CREAR LISTA DE TANQUES ===
        tanques = []
        for _, row in df_filtrado.iterrows():
            capacidad = row.get("Capacidad")
            serie = str(row.get("Serie")).strip()
            tipo = str(row.get("Tipo")).strip().upper()

            if pd.notna(capacidad) and serie and tipo:
                tanques.append({
                    "capacidad": capacidad,
                    "serie": serie,
                    "tipo": tipo
                })

        # === EXTRAER AÑO DE INSPECCIÓN ===
        fecha_inspeccion_formato = fecha_dt.strftime("%d/%m/%Y")
        ano = fecha_dt.year

        # === AGRUPAR TANQUES POR (TIPO, CAPACIDAD) ===
        grupo_tanques = defaultdict(lambda: defaultdict(list))
        for t in tanques:
            grupo_tanques[t["tipo"]][t["capacidad"]].append(t["serie"])

        # === FORMATEAR TEXTO DE TANQUES ===
        partes = []
        for tipo in sorted(grupo_tanques.keys()):
            # sorted(...items(), reverse=True) tal como tu script
            for cap, series in sorted(grupo_tanques[tipo].items(), reverse=True):
                n = len(series)
                if n > 1:
                    # Aquí replicamos exactamente tu lógica de unir series
                    serie_str = " y ".join([", ".join(series[:-1]), series[-1]])
                else:
                    serie_str = series[0]
                tipo_lower = tipo.lower()
                plural = "" if n == 1 else "s"
                texto = f"{n} tanque{plural} {tipo_lower}{plural} de {cap} galones de GLP con número{'s' if n > 1 else ''} de serie {serie_str}"
                partes.append(texto)

        if partes:
            texto_tanques = "Se inspeccionaron " + ", ".join(partes) + ". Se comprobó que los tanques no cuentan con abolladuras, hendiduras o áreas en estado avanzado de abrasión, erosión o corrosión. Asimismo, se sometió a inspección los accesorios del tanque comprobando su correcto funcionamiento y hermeticidad."
        else:
            texto_tanques = "No se registraron tanques asociados en la base de datos para esta ubicación."

        # = = = ADECUACION DE TEXTOS = = =
        cliente = limpiar_texto(cliente)
        direccion = limpiar_texto(direccion)
        texto_tanques = limpiar_texto(texto_tanques)
        fecha_emision = limpiar_texto(fecha_emision)

        # === CREAR PDF ===
        pdf = FPDF()
        pdf.add_page()
        pdf.set_auto_page_break(auto=True, margin=40)

        # --- ENCABEZADO (manteniendo exactamente parámetros del ejemplo)
        pdf.set_font("Helvetica", size=10)

        logo_path = os.path.join("static", "logo_solgaspro.png")
        if os.path.exists(logo_path):
            pdf.image(logo_path, x=10, y=10, w=55)
        # si no existe logo, seguimos sin error

        pdf.set_xy(140, 30)
        pdf.multi_cell(60, 1, fecha_emision, align='R')

        pdf.ln(3)
        pdf.cell(0, 1, "Señor(a):")
        pdf.ln(6)
        pdf.set_font("Helvetica", "B", 10)
        pdf.cell(0, 1, cliente)
        pdf.ln(6)
        pdf.set_font("Helvetica", size=10)
        pdf.cell(0, 1, "Presente:")
        pdf.ln(3)

        # --- REFERENCIA
        pdf.set_font("Helvetica", "B", 10)
        pdf.cell(0, 1, "Referencia: Certificado de Operatividad Instalación GLP", align="R")
        pdf.ln(2)

        # --- CUERPO (texto principal, manteniendo el texto original)
        pdf.set_font("Helvetica", size=11)
        pdf.multi_cell(0, 4, f"""Estimado(a),

Sirva la presente para saludarlo(a) cordialmente e informarle que SOLGAS S.A. con fecha {fecha_inspeccion_formato} ha realizado los trabajos de Mantenimiento Preventivo Anual en la instalación de la zona del tanque de GLP y las redes de media presión en la dirección {direccion}, en cumplimiento de la Norma Técnica Peruana NTP 321.123 (REVISADA 2025) de instalaciones de consumidores directos (Capítulo 5.1.16.1) y de acuerdo con los estándares de seguridad y calidad de la empresa.""")

        pdf.ln(3)
        pdf.multi_cell(0, 4, texto_tanques)

        pdf.ln(3)
        pdf.multi_cell(0, 4, """Asimismo, se realizó la inspección, revisión y mantenimiento de los reguladores de 1era etapa, verificando que se encuentran en condiciones seguras de operación.""")

        pdf.ln(3)
        pdf.multi_cell(0, 4, """Los tanques son recipientes a presión fabricados de acuerdo con el Código ASME Sección VIII, API 510 con altos estándares internacionales de seguridad y calidad.""")

        pdf.ln(3)
        pdf.multi_cell(0, 4, """Es importante indicar que el tanque instalado por Solgas S.A. se encuentra cubierto por una póliza de Responsabilidad Civil Extracontractual de hasta 733 UIT.""")

        pdf.ln(3)
        pdf.multi_cell(0, 4, "El presente Certificado de Operatividad tiene un periodo de vigencia de un año.")

        if considerar_nuevo_formato == 0 :
            pdf.ln(3)
            pdf.multi_cell(0,4,texto_adicional)

        pdf.ln(1)
        pdf.multi_cell(0,4,"Sin otro particular, \nAtentamente")

        # --- FIRMA (mismos nombres de archivos que en tu script)
        firma_map = {
            "CC": ("CC-FIRMA.png", 85, 40),
            "ML": ("ML-FIRMA.png", 75, 60),
            "CT": ("CT-FIRMA.png", 85, 40),
            "EP": ("EP-FIRMA.png", 80, 40),
            "AR": ("AR-FIRMA.png", 85, 40),
        }
        firma_file = firma_map.get(ingeniero, ("CT-FIRMA.png", 85, 40))[0]
        firma_x = firma_map.get(ingeniero, ("CT-FIRMA.png", 85, 40))[1]
        firma_w = firma_map.get(ingeniero, ("CT-FIRMA.png", 85, 40))[2]

        firma_path = os.path.join("static", firma_file)
        if os.path.exists(firma_path):
            # intentamos colocar la imagen de la firma, sin forzar y permitiendo que FPDF la ubique verticalmente
            try:
                pdf.image(firma_path, x=firma_x, w=firma_w)
            except Exception:
                # si por alguna razon falla la insercion, la ignoramos y seguimos
                pass

        # --- PIE DE PÁGINA
        pdf.set_text_color(150, 150, 150)
        pdf.ln(3)
        pdf.set_font("Helvetica", size=8)
        pdf.cell(0, 1, "Jr. Vittore Scarpazza Carpaccio N° 250 Piso 7, San Borja", align="C")
        pdf.ln(4)
        pdf.cell(0, 1, "Telf: (+511) 613-3330", align="C")
        pdf.ln(4)
        pdf.cell(0, 1, "www.solgaspro.com.pe", align="C")

        # === GUARDAR PDF CON NOMBRE PERSONALIZADO ===
        cliente_limpio = str(cliente).replace("/", "-").replace("\\", "-").replace(":", "").replace("*", "").replace("?", "").replace('"', "").replace("<", "").replace(">", "").replace("|", "")
        nombre_archivo = f"CO_{ubicacion_input}_{cliente_limpio}_{ano}.pdf"

        # guardar en el directorio actual
        pdf.output(nombre_archivo)

        # enviar el archivo resultante para descarga
        return send_file(nombre_archivo, as_attachment=True)

    # GET -> mostrar formulario
    return render_template("form.html")


if __name__ == "__main__":
    app.run(debug=True)
