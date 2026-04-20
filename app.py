import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import unicodedata
import io
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT

st.set_page_config(page_title="Analizador de Deportes", layout="wide")

def limpiar_texto(texto):
    """Limpia tildes, mayúsculas, espacios extra y caracteres invisibles."""
    if pd.isna(texto):
        return ""
    texto = str(texto)
    texto = texto.strip()
    texto = " ".join(texto.split())
    texto = texto.lower()
    texto_normalizado = unicodedata.normalize('NFKD', texto)
    return "".join([c for c in texto_normalizado if not unicodedata.combining(c)])


def generar_pdf(fig, palabras_input, columna_objetivo, total_con, total_filas, df_filtrado):
    """Genera un PDF con el gráfico, resumen y tabla de coincidencias."""
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=2*cm,
        leftMargin=2*cm,
        topMargin=2*cm,
        bottomMargin=2*cm
    )

    styles = getSampleStyleSheet()
    style_titulo = ParagraphStyle(
        'Titulo',
        parent=styles['Title'],
        fontSize=18,
        textColor=colors.HexColor('#1a1a2e'),
        spaceAfter=6,
        alignment=TA_CENTER
    )
    style_subtitulo = ParagraphStyle(
        'Subtitulo',
        parent=styles['Heading2'],
        fontSize=13,
        textColor=colors.HexColor('#16213e'),
        spaceBefore=14,
        spaceAfter=6,
        alignment=TA_LEFT
    )
    style_normal = ParagraphStyle(
        'Normal2',
        parent=styles['Normal'],
        fontSize=10,
        textColor=colors.HexColor('#333333'),
        spaceAfter=4
    )
    style_centro = ParagraphStyle(
        'Centro',
        parent=styles['Normal'],
        fontSize=10,
        textColor=colors.HexColor('#555555'),
        alignment=TA_CENTER,
        spaceAfter=10
    )

    porcentaje = f"{(total_con / total_filas) * 100:.2f}%" if total_filas > 0 else "N/A"
    elementos = []

    # --- Título ---
    elementos.append(Paragraph("Informe de Estadisticas de Deportes", style_titulo))
    elementos.append(Spacer(1, 0.3*cm))
    elementos.append(Paragraph(f"Columna analizada: <b>{columna_objetivo}</b>", style_centro))
    elementos.append(Paragraph(f"Palabra(s) clave: <b>{palabras_input}</b>", style_centro))
    elementos.append(Spacer(1, 0.4*cm))

    # --- Resumen numérico ---
    elementos.append(Paragraph("Resumen", style_subtitulo))

    datos_resumen = [
        ["Concepto", "Valor"],
        ["Total de inscripciones (filas)", str(total_filas)],
        ["Coincidencias encontradas", str(total_con)],
        ["Porcentaje sobre el total", porcentaje],
    ]
    tabla_resumen = Table(datos_resumen, colWidths=[9*cm, 6*cm])
    tabla_resumen.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1a1a2e')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 11),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.HexColor('#f0f4f8'), colors.white]),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#cccccc')),
        ('FONTSIZE', (0, 1), (-1, -1), 10),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
    ]))
    elementos.append(tabla_resumen)
    elementos.append(Spacer(1, 0.5*cm))

    # --- Gráfico ---
    elementos.append(Paragraph("Distribucion Grafica", style_subtitulo))
    img_buffer = io.BytesIO()
    fig.savefig(img_buffer, format='PNG', dpi=150, bbox_inches='tight')
    img_buffer.seek(0)
    img = Image(img_buffer, width=10*cm, height=10*cm)
    img.hAlign = 'CENTER'
    elementos.append(img)
    elementos.append(Spacer(1, 0.5*cm))

    # --- Tabla de coincidencias ---
    elementos.append(Paragraph(f"Vista previa de coincidencias ({total_con} encontradas)", style_subtitulo))

    df_tabla = df_filtrado.head(100).fillna("").astype(str)
    columnas = list(df_tabla.columns)

    ancho_pagina = 17*cm
    n_cols = len(columnas)
    ancho_col = ancho_pagina / n_cols

    filas_tabla = [columnas] + df_tabla.values.tolist()
    tabla_datos = Table(filas_tabla, colWidths=[ancho_col] * n_cols, repeatRows=1)
    tabla_datos.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#4CAF50')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 7),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.HexColor('#f9f9f9'), colors.white]),
        ('GRID', (0, 0), (-1, -1), 0.4, colors.HexColor('#dddddd')),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ('WORDWRAP', (0, 0), (-1, -1), True),
    ]))
    elementos.append(tabla_datos)

    if len(df_filtrado) > 100:
        elementos.append(Spacer(1, 0.3*cm))
        elementos.append(Paragraph(
            f"Se muestran las primeras 100 filas de {len(df_filtrado)} coincidencias totales.",
            style_normal
        ))

    doc.build(elementos)
    buffer.seek(0)
    return buffer


# ─────────────────────────────────────────────
#  APP PRINCIPAL
# ─────────────────────────────────────────────

st.title("📊 Estadísticas de Deportes en Excel")
st.markdown("Carga tu archivo para analizar la frecuencia de palabras clave en tus columnas.")

archivo = st.file_uploader("Sube tu archivo Excel", type=["xlsx", "xls"])

if archivo:
    df = pd.read_excel(archivo)
    st.success(f"Archivo cargado: {len(df)} filas encontradas")

    col1, col2 = st.columns(2)

    with col1:
        columna_objetivo = st.selectbox("Selecciona la columna a analizar", df.columns)
        palabras_input = st.text_input("Palabra(s) clave — separadas por coma (ej: futbol, tenis)", "")

        desglosar = False
        if "," in palabras_input:
            desglosar = st.checkbox("🔍 Desglosar porcentaje por cada palabra clave")

    if palabras_input:
        palabras_lista_original = [p.strip() for p in palabras_input.split(",") if p.strip()]
        palabras_limpias = [limpiar_texto(p) for p in palabras_lista_original]

        df['_texto_limpio'] = df[columna_objetivo].apply(limpiar_texto)

        coincidencias_mask = df['_texto_limpio'].apply(
            lambda celda: any(p in celda for p in palabras_limpias)
        )

        total_con = int(coincidencias_mask.sum())
        total_filas = len(df)
        total_sin = total_filas - total_con

        with col2:
            fig, ax = plt.subplots(figsize=(6, 6))
            plt.subplots_adjust(left=0.1, right=0.9, top=0.9, bottom=0.1)

            if desglosar and len(palabras_limpias) > 1:
                conteos_individuales = []
                for p in palabras_limpias:
                    count = df['_texto_limpio'].str.contains(p, na=False).sum()
                    conteos_individuales.append(count)

                labels = palabras_lista_original + ['Otros / Sin datos']
                sizes = conteos_individuales + [total_sin]
                cmap = plt.get_cmap('tab10')
                colors_list = [cmap(i) for i in range(len(palabras_limpias))]
                colors_list.append('#E0E0E0')
                ax.pie(sizes, labels=labels, autopct='%1.1f%%', colors=colors_list,
                       startangle=140, radius=1, pctdistance=0.85)
                ax.set_title("Desglose Detallado")
            else:
                etiqueta_general = ", ".join(palabras_lista_original)
                labels = [f'Con "{etiqueta_general}"', 'Otros / Sin datos']
                sizes = [total_con, total_sin]
                colors_list = ['#4CAF50', '#E0E0E0']
                if total_con > 0:
                    ax.pie(sizes, labels=labels, autopct='%1.1f%%', colors=colors_list,
                           startangle=140, radius=1)
                else:
                    ax.text(0.5, 0.5, "Sin coincidencias", ha='center', va='center')
                ax.set_title("Distribucion General")

            ax.axis('equal')
            st.pyplot(fig)

        st.subheader(f"Vista previa de coincidencias — {total_con} encontradas de {total_filas} filas")
        df_filtrado = df[coincidencias_mask].drop(columns=['_texto_limpio'], errors='ignore')
        st.dataframe(df_filtrado, use_container_width=True)

        # --- Reporte PDF ---
        st.divider()
        st.subheader("📥 Descargar Informe en PDF")

        pdf_buffer = generar_pdf(
            fig=fig,
            palabras_input=palabras_input,
            columna_objetivo=columna_objetivo,
            total_con=total_con,
            total_filas=total_filas,
            df_filtrado=df_filtrado
        )

        nombre_archivo = palabras_input.replace(" ", "_").replace(",", "-")
        st.download_button(
            label="⬇️ Descargar reporte en PDF",
            data=pdf_buffer,
            file_name=f"reporte_{nombre_archivo}.pdf",
            mime="application/pdf"
        )

        df.drop(columns=['_texto_limpio'], inplace=True, errors='ignore')
