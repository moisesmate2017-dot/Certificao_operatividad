# app.py
from flask import Flask, render_template, request, send_file, redirect, url_for, flash
from collections import defaultdict
from fpdf import FPDF
from datetime import datetime, timedelta
import pandas as pd
import os

app = Flask(__name__)
app.secret_key = "cambia_esto_por_una_clave_muy_segura"  # cambia en producción

# = = = utilidades = = =
def limpiar_texto(texto):
    if not isinstance(texto, str):
        return texto
    return (texto.replace("–", "-").replace("—", "-").replace("“", '"').replace("”", '"')
            .replace("‘", "'").replace("’", "'").replace("…", "...")
            .replace("á", "a").replace("é", "e").replace("í", "i").replace("ó", "o").replace("ú", "u")
            .replace("Á", "A").replace("É", "E").replace("Í", "I").replace("Ó", "O").replace("Ú", "U")
            .replace("ñ", "n").replace("Ñ", "N"))

MESES_ES = {
    1: "enero", 2: "febrero", 3: "marzo", 4: "abril",
    5: "mayo", 6: "junio", 7: "julio", 8: "agosto",
    9: "setiembre", 10: "octubre", 11: "noviembre", 12: "diciembre"
}

DATA_FILE = "data_base.xlsm"
UPDATED_FILE_PREFIX = "data_base_updated"

# = = = Ruta principal: formulario = = =
@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        ubicacion = request.form.get("ubicacion", "").strip()
        fecha_inspeccion = request.form.get("fecha_inspeccion", "").strip()
        ingeniero = request.form.get("ingeniero", "CT").strip()

        if not ubicacion or not fecha_inspeccion:
            flash("Complete número de instalación y fecha de inspección.", "danger")
            return redirect(url_for("index"))

        # Directamente cargamos los datos y renderizamos la vista previa (no hay loader separado)
        # Leer Excel y extraer cliente/direccion/tanques
        cliente = ""
        direccion = ""
        tanques = []

        if os.path.exists(DATA_FILE):
            try:
                df = pd.read_excel(DATA_FILE, sheet_name="DATA", dtype=str)
            except Exception as e:
                flash(f"Error leyendo {DATA_FILE}: {e}", "danger")
                df = None
            if df is not None:
                # robusto: si columna Ubicacion existe, comparar como string
                if "Ubicacion" in df.columns:
                    df["Ubicacion_str"] = df["Ubicacion"].astype(str)
                    df_filtrado = df[df["Ubicacion_str"] == str(ubicacion)]
                else:
                    df_filtrado = pd.DataFrame()

                if not df_filtrado.empty:
                    cliente = df_filtrado.iloc[0].get("Nombre Titular", "") or ""
                    direccion = df_filtrado.iloc[0].get("Direccion", "") or ""
                    # extraer todos los tanques asociados (cada fila representa un tanque)
                    for _, row in df_filtrado.iterrows():
                        capacidad = row.get("Capacidad")
                        serie = row.get("Serie")
                        tipo = row.get("Tipo")
                        # validar valores no nulos
                        if pd.notna(capacidad) and capacidad != "nan" and serie and tipo:
                            tanques.append({
                                "tipo": str(tipo).strip(),
                                "capacidad": str(capacidad).strip(),
                                "serie": str(serie).strip()
                            })

        # calcular fecha_emision = inspeccion + 7 dias
        try:
            fecha_dt = datetime.strptime(fecha_inspeccion, "%Y-%m-%d")
        except Exception:
            flash("Formato de fecha inválido; use el calendario.", "danger")
            return redirect(url_for("index"))
        fecha_emision_dt = fecha_dt + timedelta(days=7)
        if fecha_emision_dt > datetime.today() :
            fecha_emision_dt = datetime.today()
            mes_text = MESES_ES.get(fecha_emision_dt.month, fecha_emision_dt.strftime("%B").lower())
            fecha_emision = f"Lima, {fecha_emision_dt.day} de {mes_text} de {fecha_emision_dt.year}"
        else:
            mes_text = MESES_ES.get(fecha_emision_dt.month, fecha_emision_dt.strftime("%B").lower())
            fecha_emision = f"Lima, {fecha_emision_dt.day} de {mes_text} de {fecha_emision_dt.year}"
        

        # render preview template (está llamado preview_loader.html por compatibilidad con tu estructura)
        return render_template("preview_loader.html",
                               ubicacion=ubicacion,
                               cliente=cliente,
                               direccion=direccion,
                               tanques=tanques,
                               fecha_inspeccion=fecha_inspeccion,
                               fecha_emision=fecha_emision,
                               ingeniero=ingeniero)

    return render_template("form.html")

# = = = Ruta generar PDF = = =
@app.route("/generar_pdf", methods=["POST"])
def generar_pdf():
    # recolectar campos
    ubicacion = request.form.get("ubicacion", "").strip()
    cliente = request.form.get("cliente", "").strip()
    direccion = request.form.get("direccion", "").strip()
    fecha_inspeccion = request.form.get("fecha_inspeccion", "").strip()
    fecha_emision = request.form.get("fecha_emision", "").strip()
    ingeniero = request.form.get("ingeniero", "CT").strip()
    guardar_en_base = request.form.get("guardar_en_base", "off") == "on"

    # listas de tanques
    tipos = request.form.getlist("tipo[]")
    capacidades = request.form.getlist("capacidad[]")
    series = request.form.getlist("serie[]")

    # construir lista de tanques (ignorar vacías)
    tanques = []
    for t, c, s in zip(tipos, capacidades, series):
        t_ = (t or "").strip()
        c_ = (c or "").strip()
        s_ = (s or "").strip()
        if not (t_ == "" and c_ == "" and s_ == ""):
            tanques.append({"tipo": t_.upper(), "capacidad": c_, "serie": s_})

    # parsear fecha_inspeccion
    try:
        fecha_dt = datetime.strptime(fecha_inspeccion, "%Y-%m-%d")
    except Exception:
        try:
            fecha_dt = datetime.strptime(fecha_inspeccion, "%d/%m/%Y")
        except Exception:
            return "Fecha de inspección inválida", 400
    ano = fecha_dt.year

    # guardar en base si se solicitó (crea archivo xlsx con timestamp)
    if guardar_en_base:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        updated_name = f"{UPDATED_FILE_PREFIX}_{timestamp}.xlsx"
        try:
            if os.path.exists(DATA_FILE):
                df_orig = pd.read_excel(DATA_FILE, sheet_name="DATA", dtype=str)
            else:
                df_orig = pd.DataFrame(columns=["Ubicacion", "Nombre Titular", "Direccion", "Tipo", "Capacidad", "Serie"])

            # eliminar filas de la ubicacion actual
            if "Ubicacion" in df_orig.columns:
                df_orig["Ubicacion_str"] = df_orig["Ubicacion"].astype(str)
                df_sin = df_orig[df_orig["Ubicacion_str"] != str(ubicacion)].drop(columns=["Ubicacion_str"])
            else:
                df_sin = df_orig

            nuevas = []
            if tanques:
                for t in tanques:
                    nuevas.append({
                        "Ubicacion": ubicacion,
                        "Nombre Titular": cliente,
                        "Direccion": direccion,
                        "Tipo": t.get("tipo", ""),
                        "Capacidad": t.get("capacidad", ""),
                        "Serie": t.get("serie", "")
                    })
            else:
                nuevas.append({
                    "Ubicacion": ubicacion,
                    "Nombre Titular": cliente,
                    "Direccion": direccion,
                    "Tipo": "",
                    "Capacidad": "",
                    "Serie": ""
                })
            df_nuevas = pd.DataFrame(nuevas)
            df_result = pd.concat([df_sin, df_nuevas], ignore_index=True, sort=False)
            df_result.to_excel(updated_name, sheet_name="DATA", index=False)
            flash(f"Guardado en {updated_name}", "success")
        except Exception as e:
            flash(f"Error guardando base actualizada: {e}", "danger")

    # agrupar tanques y formar texto (idéntico al script original)
    grupo_tanques = defaultdict(lambda: defaultdict(list))
    for t in tanques:
        tipo = t.get("tipo", "").upper()
        cap = t.get("capacidad", "")
        serie = t.get("serie", "")
        if tipo and cap and serie:
            grupo_tanques[tipo][cap].append(serie)

    partes = []
    for tipo in sorted(grupo_tanques.keys()):
        for cap, series_list in sorted(grupo_tanques[tipo].items(), reverse=True):
            n = len(series_list)
            if n > 1:
                serie_str = " y ".join([", ".join(series_list[:-1]), series_list[-1]])
            else:
                serie_str = series_list[0]
            tipo_lower = tipo.lower()
            plural = "" if n == 1 else "s"
            texto = f"{n} tanque{plural} {tipo_lower}{plural} de {cap} galones de GLP con número{'s' if n > 1 else ''} de serie {serie_str}"
            partes.append(texto)

    if partes:
        texto_tanques = "Se inspeccionaron " + ", ".join(partes) + ". Se comprobó que los tanques no cuentan con abolladuras, hendiduras o áreas en estado avanzado de abrasión, erosión o corrosión. Asimismo, se sometió a inspección los accesorios del tanque comprobando su correcto funcionamiento y hermeticidad."
    else:
        texto_tanques = "No se registraron tanques asociados en la base de datos para esta ubicación."

    cliente = limpiar_texto(cliente)
    direccion = limpiar_texto(direccion)
    texto_tanques = limpiar_texto(texto_tanques)
    fecha_emision = limpiar_texto(fecha_emision)

    # considerar nuevo formato si año emisión == 2025
    try:
        partes_fecha = fecha_emision.split()
        ano_emision = int(partes_fecha[-1]) if partes_fecha[-1].isdigit() else fecha_dt.year
    except Exception:
        ano_emision = fecha_dt.year
    considerar_nuevo_formato = 0 if ano_emision >= 2025 else 1

    # === generar PDF (manteniendo tu layout original) ===
    pdf = FPDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=40)

    pdf.set_font("Helvetica", size=10)
    logo_path = os.path.join("static", "logo_solgaspro.png")
    if os.path.exists(logo_path):
        try:
            pdf.image(logo_path, x=10, y=10, w=55)
        except Exception:
            pass

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

    pdf.set_font("Helvetica", "B", 10)
    pdf.cell(0, 1, "Referencia: Certificado de Operatividad Instalación GLP", align="R")
    pdf.ln(2)

    pdf.set_font("Helvetica", size=11)
    fecha_inspeccion_formato = fecha_dt.strftime("%d/%m/%Y")
    pdf.multi_cell(0, 4, f"""Estimado(a),

Sirva la presente para saludarlo(a) cordialmente e informarle que SOLGAS S.A. con fecha {fecha_inspeccion_formato} ha realizado los trabajos de Mantenimiento Preventivo Anual en la instalación de la zona del tanque de GLP y las redes de media presión en la dirección {direccion}, en cumplimiento de la Norma Técnica Peruana NTP 321.123 (REVISADA 2025) de instalaciones de consumidores directos (Capítulo 5.2.2) y de acuerdo con los estándares de seguridad y calidad de la empresa.""")
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

    texto_adicional = """(1) En aras de mantener el estricto cumplimiento del marco normativo vigente, y en caso requieran hacer alguna modificación en el área que se encuentra alrededor de la zona de almacenamiento de tanques de GLP como: instalar equipos eléctricos, cámaras de seguridad, iluminación, tomacorrientes, edificaciones contiguas, equipos de aire acondicionado, cableado eléctrico, ductos, sumideros, almacenar materiales, construcción de muros perimetrales, etc. tienen la obligación de comunicar previamente al equipo técnico de Solgas lo pertinente, a fin de que puedan recibir la asesoría técnica y validación correspondiente, de tal manera que, se evite incurrir en algunos incumplimientos normativos que puedan ser materia de suspensión de la Ficha de Registro y sanciones pecuniarias por parte del ente fiscalizador, así como evitar cualquier riesgo innecesario en la instalación; cabe precisar que en caso de no informar acerca de las modificaciones que realicen en la instalación, usted será el único y exclusivo responsable por las consecuencias que se deriven de su accionar."""

    if considerar_nuevo_formato == 0:
        pdf.ln(3)
        pdf.multi_cell(0,4,texto_adicional)

    pdf.ln(1)
    pdf.multi_cell(0,4,"Sin otro particular, \nAtentamente")

    # firmas
    firma_map = {
        "CC": ("CC-FIRMA.png", 85, 40),
        "ML": ("ML-FIRMA.png", 75, 60),
        "CT": ("CT-FIRMA.png", 85, 40),
        "LM": ("LM-FIRMA.png", 80, 40),
        "AR": ("AR-FIRMA.png", 85, 40),
    }
    firma_file, firma_x, firma_w = firma_map.get(ingeniero, ("CT-FIRMA.png", 85, 40))
    firma_path = os.path.join("static", firma_file)
    if os.path.exists(firma_path):
        try:
            pdf.image(firma_path, x=firma_x, w=firma_w)
        except Exception:
            pass

    # pie de pagina
    pdf.set_text_color(150, 150, 150)
    pdf.ln(3)
    pdf.set_font("Helvetica", size=8)
    pdf.cell(0, 1, "Jr. Vittore Scarpazza Carpaccio N° 250 Piso 7, San Borja", align="C")
    pdf.ln(4)
    pdf.cell(0, 1, "Telf: (+511) 613-3330", align="C")
    pdf.ln(4)
    pdf.cell(0, 1, "www.solgaspro.com.pe", align="C")

    # guardar pdf con nombre unico
    cliente_limpio = str(cliente).replace("/", "-").replace("\\", "-").replace(":", "").replace("*", "").replace("?", "").replace('"', "").replace("<", "").replace(">", "").replace("|", "")
    nombre_archivo = f"CO_{ubicacion}_{cliente_limpio}_{ano}.pdf"
    pdf.output(nombre_archivo)

    return send_file(nombre_archivo, as_attachment=True)

# tiny preview_loader (reenvía POST)
@app.route("/preview_loader", methods=["POST"])
def preview_loader():
    ubicacion = request.form.get("ubicacion", "")
    fecha_inspeccion = request.form.get("fecha_inspeccion", "")
    ingeniero = request.form.get("ingeniero", "CT")
    return render_template("preview_loader.html",
                           ubicacion=ubicacion,
                           fecha_inspeccion=fecha_inspeccion,
                           ingeniero=ingeniero)

if __name__ == "__main__":
    # En Render: usar gunicorn app:app
    app.run(debug=True)
