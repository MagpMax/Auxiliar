import pandas as pd
from datetime import datetime, timedelta
from docx import Document
from configuracion import Configuracion


class Funciones:
    
    @staticmethod
    def formatear_fecha(fecha):
        if pd.isnull(fecha):
            return ""
        return fecha.strftime("%d-%m-%Y")

    @staticmethod
    def obtener_semanas(fi_str, ff_str):
        fi = datetime.strptime(fi_str, "%d-%m-%Y")
        ff = datetime.strptime(ff_str, "%d-%m-%Y")

        def semana_mes(fecha):
            primer_dia = fecha.replace(day=1)

            # weekday(): lunes=0 ... domingo=6
            ajuste = primer_dia.weekday()

            return ((fecha.day + ajuste - 1) // 7) + 1

        return semana_mes(fi), semana_mes(ff)

    @staticmethod
    def texto_semana(fi, ff):
        s1, s2 = Funciones.obtener_semanas(fi, ff)
        mes = datetime.strptime(fi, "%d-%m-%Y").strftime("%B")
        return f"Parte de la {s1}° y {s2}° semana de {mes}\n({fi} al {ff})"
    
    @staticmethod
    def filtrar_por_fecha(df):
        fi = pd.to_datetime(Configuracion.FECHA_INICIO, dayfirst=True)
        ff = pd.to_datetime(Configuracion.FECHA_FIN, dayfirst=True)

        return df[
            (df["FECHA SOLICITUD"] >= fi) &
            (df["FECHA SOLICITUD"] <= ff)
        ]


    @staticmethod
    def normalizar_tiempo(valor):

        if pd.isnull(valor):
            return "---"

        txt = str(valor).lower().strip()

        if txt in ["", "nan", "none"]:
            return "---"

        # minutos
        if "min" in txt:
            num = int(''.join(filter(str.isdigit, txt)) or 0)
            horas = round(num / 60)
            return f"{horas} h" if horas > 0 else "---"

        # formato hh:mm
        if ":" in txt:
            try:
                h, m = txt.split(":")
                h = int(h)
                m = int(m)
                if m >= 30:
                    h += 1
                return f"{h} h"
            except:
                return "---"

        # horas directas
        if "h" in txt or "hora" in txt:
            num = int(''.join(filter(str.isdigit, txt)) or 0)
            return f"{num} h" if num > 0 else "---"

        # solo número
        if txt.isdigit():
            return f"{int(txt)} h"

        return "---"

    # =========================
    # WORD
    # =========================

    @staticmethod
    def actualizar_cabecera(doc, equipo):
        tabla = doc.tables[0]

        for row in tabla.rows:
            key = row.cells[0].text.lower()

            if "servicio" in key:
                row.cells[1].text = equipo

            if "semana" in key:
                row.cells[1].text = Funciones.texto_semana(Configuracion.FECHA_INICIO, Configuracion.FECHA_FIN)

    @staticmethod
    def preparar_datos(df):
        df.columns = df.columns.str.strip().str.upper()

        df["FECHA SOLICITUD"] = pd.to_datetime(
            df["FECHA SOLICITUD"], dayfirst=True, errors="coerce"
        )

        col_estado = next((c for c in df.columns if "ESTADO" in c), None)
        col_tiempo = next((c for c in df.columns if "TIEMPO" in c), None)

        df["ESTADO"] = df[col_estado].astype(str).str.lower() if col_estado else ""
        df["TIEMPO ESTIMADO"] = df[col_tiempo] if col_tiempo else ""

        return df

    @staticmethod
    def leer_todos_los_profesionales(path):
        dataframes = []

        for profesional in Configuracion.PROFESIONALES:
            try:
                df = pd.read_excel(path, sheet_name=profesional)
                df.columns = df.columns.str.strip().str.upper()

                df["PROFESIONAL"] = profesional
                df["EQUIPO"] = Configuracion.MAPEO_EQUIPOS[profesional]

                dataframes.append(df)

            except Exception as e:
                print(f"⚠️ Error leyendo {profesional}: {e}")

        return pd.concat(dataframes, ignore_index=True)
    
    
    @staticmethod
    def obtener_columna(nombre, df_prof):
        candidatas = []

        for c in df_prof.columns:
            c_norm = str(c).upper().replace("\n", " ").strip()

            if nombre in c_norm:
                candidatas.append(c)

        # devolver la primera columna que realmente tenga datos
        for c in candidatas:
            serie = (
                df_prof[c]
                .astype(str)
                .str.strip()
                .replace(["", "nan", "None", "---"], pd.NA)
            )

            if serie.notna().any():
                return c

        return candidatas[0] if candidatas else None
    
    @staticmethod
    def normalizar_columnas(df):
        df.columns = [
            col.strip().replace("\n", " ").replace("  ", " ").upper()
            for col in df.columns
        ]
        return df

    @staticmethod
    def obtener_columna_tiempo(df):
        for col in df.columns:
            if "TIEMPO" in col and "ACTIVIDAD" in col:
                return col
        return None

    @staticmethod
    def obtener_valor(row, posibles_columnas):
        for col in posibles_columnas:
            if col in row and pd.notnull(row[col]):
                val = str(row[col]).strip()
                if val and val.lower() != "nan":
                    return val
        return "---"
    
    @staticmethod
    def normalizar_tiempo(valor):

        if valor is None:
            return "---"

        if isinstance(valor, float) or isinstance(valor, int):
            return str(valor)

        valor = str(valor).strip().replace("\xa0", "")

        if valor == "" or valor.lower() == "nan":
            return "---"

        return valor
            