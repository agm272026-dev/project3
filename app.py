import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import unicodedata
import io

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
        
        # Nueva lógica para el checkbox de desglose
        desglosar = False
        if "," in palabras_input:
            desglosar = st.checkbox("🔍 Desglosar porcentaje por cada palabra clave")

    if palabras_input:
        # Extraemos las palabras originales para las etiquetas y las limpias para la búsqueda
        palabras_lista_original = [p.strip() for p in palabras_input.split(",") if p.strip()]
        palabras_limpias = [limpiar_texto(p) for p in palabras_lista_original]

        df['_texto_limpio'] = df[columna_objetivo].apply(limpiar_texto)

        # Máscara global (Cualquiera de las palabras)
        coincidencias_mask = df['_texto_limpio'].apply(
            lambda celda: any(p in celda for p in palabras_limpias)
        )

        total_con = int(coincidencias_mask.sum())
        total_filas = len(df)
        total_sin = total_filas - total_con

        with col2:
            # Creamos la figura con un tamaño fijo
            fig, ax = plt.subplots(figsize=(6, 6))
            
            # Ajustamos los márgenes para que el contenido no se mueva
            plt.subplots_adjust(left=0.1, right=0.9, top=0.9, bottom=0.1)

            if desglosar and len(palabras_limpias) > 1:
                conteos_individuales = []
                for p in palabras_limpias:
                    count = df['_texto_limpio'].str.contains(p, na=False).sum()
                    conteos_individuales.append(count)
                
                labels = palabras_lista_original + ['Otros']
                sizes = conteos_individuales + [total_sin]
                
                # Paleta de colores + Gris al final
                cmap = plt.get_cmap('tab10')
                colors = [cmap(i) for i in range(len(palabras_limpias))]
                colors.append('#E0E0E0')
                
                # El parámetro radius=1 mantiene el círculo siempre del mismo tamaño
                ax.pie(sizes, labels=labels, autopct='%1.1f%%', colors=colors, 
                       startangle=140, radius=1, pctdistance=0.85)
                ax.set_title("Desglose Detallado")
            else:
                etiqueta_general = ", ".join(palabras_lista_original)
                labels = [f'Con "{etiqueta_general}"', 'Otros / Sin datos']
                sizes = [total_con, total_sin]
                colors = ['#4CAF50', '#E0E0E0']
                
                if total_con > 0:
                    ax.pie(sizes, labels=labels, autopct='%1.1f%%', colors=colors, 
                           startangle=140, radius=1)
                else:
                    ax.text(0.5, 0.5, "Sin coincidencias", ha='center', va='center')
                ax.set_title("Distribución General")

            # Forzamos que sea un círculo y dibujamos
            ax.axis('equal') 
            st.pyplot(fig)

        st.subheader(f"Vista previa de coincidencias — {total_con} encontradas de {total_filas} filas")

        df_filtrado = df[coincidencias_mask].drop(columns=['_texto_limpio'], errors='ignore')
        st.dataframe(df_filtrado, use_container_width=True)

        # Limpieza y preparación de reporte
        st.divider()
        st.subheader("📥 Descargar Informe")

        porcentaje = f"{(total_con / total_filas) * 100:.2f}%" if total_filas > 0 else "N/A"
        resumen_data = {
            "Concepto": ["Palabra(s) clave", "Columna analizada", "Coincidencias", "Total filas", "Porcentaje"],
            "Valor": [palabras_input, columna_objetivo, total_con, total_filas, porcentaje]
        }
        df_resumen = pd.DataFrame(resumen_data)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_resumen.to_excel(writer, sheet_name='Resumen', index=False)
            df_filtrado.to_excel(writer, sheet_name='Datos_Filtrados', index=False)

        nombre_archivo = palabras_input.replace(" ", "_").replace(",", "-")
        st.download_button(
            label="⬇️ Descargar reporte en Excel",
            data=output.getvalue(),
            file_name=f"reporte_{nombre_archivo}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        # Eliminar columna temporal
        df.drop(columns=['_texto_limpio'], inplace=True, errors='ignore')
