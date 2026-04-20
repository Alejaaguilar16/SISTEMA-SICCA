from flask import Flask, render_template, request, send_file
import pandas as pd
import re
import unicodedata
import os
import tempfile
from difflib import SequenceMatcher

app = Flask(__name__)

#CONVERSIÓN INTERNA ARCHIVO D BLACKBOARD

def convertir_blackboard_si_es_xls(archivo_subido):

    nombre = archivo_subido.filename.lower()

    temp_dir = tempfile.mkdtemp()

    ruta_original = os.path.join(
        temp_dir,
        archivo_subido.filename
    )

    archivo_subido.save(ruta_original)

    if nombre.endswith(".xlsx"):
        return ruta_original

    if nombre.endswith(".xls"):

        import win32com.client as win32

        excel = win32.Dispatch(
            "Excel.Application"
        )

        excel.Visible = False
        excel.DisplayAlerts = False

        try:

            wb = excel.Workbooks.Open(
                os.path.abspath(
                    ruta_original
                ),
                CorruptLoad=1
            )

            ruta_convertida = (
                ruta_original + "x"
            )

            wb.SaveAs(
                os.path.abspath(
                    ruta_convertida
                ),
                FileFormat=51
            )

            wb.Close()

            return ruta_convertida

        finally:
            excel.Quit()

    raise Exception(
        "Formato Blackboard no soportado"
    )

# LIMPIEZA

def limpiar_columnas(df):
    df.columns = [str(c).strip().lower() for c in df.columns]
    return df


def limpiar_texto(x):
    if isinstance(x, str):
        x = re.sub(r'_x0000_', '', x)
        return x.strip().lower()
    return x


# Cambiado solo por compatibilidad pandas
def limpiar_df(df):
    df = limpiar_columnas(df)
    return df.map(limpiar_texto)

# NORMALIZACIÓN DE NOTA

def normalizar_nota(valor):

    if pd.isna(valor):
        return None

    valor = str(valor).strip()

    valor = valor.replace(
        ",",
        "."
    )

    return valor

# NORMALIZACIÓN DE NOMBRE

def normalizar_nombre(nombre):

    nombre = str(nombre).lower().strip()

    nombre = ''.join(
        c for c in unicodedata.normalize(
            'NFD',
            nombre
        )
        if unicodedata.category(c) != 'Mn'
    )

    nombre = re.sub(
        r'[^a-z\s]',
        '',
        nombre
    )

    nombre = re.sub(
        r'\s+',
        ' ',
        nombre
    )

    palabras = nombre.split()
    palabras.sort()

    return " ".join(
        palabras
    )

#DETECTAR COLUMNAS
def detectar_columna(df, posibles):

    for col in df.columns:
        for p in posibles:
            if p in col:
                return col

    return None

#SIMILITUD

def similitud(a,b):

    a_set = set(a.split())
    b_set = set(b.split())

    inter = a_set.intersection(
        b_set
    )

    if not a_set or not b_set:
        return 0

    score = len(inter) / max(
        len(a_set),
        len(b_set)
    )

    if (
        a_set.issubset(b_set)
        or
        b_set.issubset(a_set)
    ):
        score += 0.3

    return min(score,1)

#ELEGIR CORREO

def elegir_correo(grupo):

    for u in grupo["usuario"]:

        if "@alumnoeseit.edu.co" in str(u):
            return u

    return grupo[
        "usuario"
    ].iloc[0]

#RESOLVER NOTA BLACKBOARD

def resolver_nota_bb(grupo):

    notas=[]

    for n in grupo["nota_bb"]:

        n = normalizar_nota(n)

        if n is not None:

            try:
                notas.append(
                    float(n)
                )

            except:
                continue

    if not notas:
        return None,False,[]

    notas_unicas = sorted(
        list(
            set(notas)
        )
    )

    if len(notas_unicas)>1:

        notas_validas=[
            n for n in notas_unicas
            if n>0
        ]

        if notas_validas:

            return (
                str(max(notas_validas)),
                True,
                notas_unicas
            )

        else:

            return (
                "0.0",
                True,
                notas_unicas
            )

    return (
        str(notas_unicas[0]),
        False,
        notas_unicas
    )
#MATCH ORIGINAL (completo)
def match_estudiantes(bb, at):

    resultados=[]
    usados_at=set()

    for _, row_bb in bb.iterrows():

        clave_bb=row_bb["clave"]
        nota_bb=row_bb["nota_bb"]

        encontrado=None
        idx_encontrado=None

        # match exacto
        for idx,row_at in at.iterrows():

            if clave_bb==row_at["clave"]:

                encontrado=row_at
                idx_encontrado=idx
                break

        # similitud
        if encontrado is None:

            mejor=0

            for idx,row_at in at.iterrows():

                s=similitud(
                    clave_bb,
                    row_at["clave"]
                )

                if s>mejor:

                    mejor=s
                    encontrado=row_at
                    idx_encontrado=idx

            if mejor<0.6:
                encontrado=None


        if encontrado is not None:

            usados_at.add(
                idx_encontrado
            )

            nota_at=encontrado["nota_at"]

            if row_bb.get(
                "conflicto_bb",
                False
            ):

                notas=row_bb.get(
                    "notas_detectadas",
                    []
                )

                estado=f"Duplicado en Blackboard: notas detectadas {notas}"

            elif nota_bb==nota_at:

                estado="Correcto"

            else:

                estado="Nota no coincide"


            resultados.append({

                "estudiante":
                row_bb["nombre_original"],

                "correo":
                row_bb["correo"],

                "usuario":
                row_bb["usuario"],

                "nota_bb":
                nota_bb,

                "nota_at":
                nota_at,

                "estado":
                estado

            })

        else:

            resultados.append({

                "estudiante":
                row_bb["nombre_original"],

                "correo":
                row_bb["correo"],

                "usuario":
                row_bb["usuario"],

                "nota_bb":
                nota_bb,

                "nota_at":
                None,

                "estado":
                "No está en Atenea"

            })

    for idx,row_at in at.iterrows():

        if idx not in usados_at:

            resultados.append({

                "estudiante":
                row_at["clave"],

                "correo":"",
                "usuario":"",

                "nota_bb":
                None,

                "nota_at":
                row_at["nota_at"],

                "estado":
                "No está en Blackboard"

            })


    return pd.DataFrame(
        resultados
    )

#RUTA PRINCIPAL
@app.route(
"/",
methods=["GET","POST"]
)

def index():

    if request.method=="POST":

        try:

            # ÚNICO CAMBIO REAL:
            ruta_bb = convertir_blackboard_si_es_xls(
                request.files[
                    "blackboard"
                ]
            )

            df_bb = pd.read_excel(
                ruta_bb
            )

            # Atenea intacto
            df_at = pd.read_excel(
                request.files[
                    "atenea"
                ]
            )

            df_bb=limpiar_df(df_bb)
            df_at=limpiar_df(df_at)


            # ATENEA
            col_student=detectar_columna(
                df_at,
                ["student"]
            )

            col_score=detectar_columna(
                df_at,
                ["score"]
            )

            col_final=detectar_columna(
                df_at,
                ["final"]
            )

            df_at=df_at[
                df_at[col_final]=="yes"
            ]

            df_at["clave"]=df_at[
                col_student
            ].apply(
                normalizar_nombre
            )

            df_at["nota_at"]=df_at[
                col_score
            ].apply(
                normalizar_nota
            )

            df_at=df_at.groupby(
                "clave",
                as_index=False
            ).agg({
                "nota_at":"first"
            })


            # BLACKBOARD
            col_apellidos=detectar_columna(
                df_bb,
                ["apellido"]
            )

            col_nombre=detectar_columna(
                df_bb,
                ["nombre"]
            )

            col_nota=detectar_columna(
                df_bb,
                ["total"]
            )

            col_usuario=detectar_columna(
                df_bb,
                ["usuario"]
            )

            df_bb["usuario"]=df_bb[
                col_usuario
            ] if col_usuario else ""

            df_bb["nombre"]=(
                df_bb[col_apellidos]
                +" "+
                df_bb[col_nombre]
            )

            df_bb["clave"]=df_bb[
                "nombre"
            ].apply(
                normalizar_nombre
            )

            df_bb["nota_bb"]=df_bb[
                col_nota
            ].apply(
                normalizar_nota
            )

            df_bb=df_bb[
                df_bb["nota_bb"].notna()
            ]


            df_bb=df_bb.groupby(
                "clave"
            ).apply(
                lambda g: pd.Series({

                    "nota_bb":
                    resolver_nota_bb(g)[0],

                    "conflicto_bb":
                    resolver_nota_bb(g)[1],

                    "notas_detectadas":
                    resolver_nota_bb(g)[2],

                    "usuario":
                    g["usuario"].iloc[0],

                    "correo":
                    elegir_correo(g),

                    "nombre_original":
                    g["nombre"].iloc[0]

                })
            ).reset_index() 
            df = match_estudiantes(
                df_bb,
                df_at
            )

            conciliados = df[
                df["estado"] == "Correcto"
            ]

            errores = df[
                df["estado"] != "Correcto"
            ]

            conciliados_final = conciliados[
                ["estudiante","correo","usuario","nota_bb"]
            ].copy()

            conciliados_final.columns = [
                "Estudiante",
                "Correo",
                "Usuario",
                "Nota Final"
            ]


            errores_final = errores[
                ["estudiante","correo","nota_bb","nota_at","estado"]
            ].copy()

            errores_final = errores_final.fillna(
                "No Aplica"
            )

            errores_final.columns = [
                "Estudiante",
                "Correo",
                "Nota Blackboard",
                "Nota Atenea",
                "Tipo de inconsistencia"
            ]


            with pd.ExcelWriter(
                "reporte_sicca.xlsx"
            ) as writer:

                conciliados_final.to_excel(
                    writer,
                    sheet_name="Notas_Finales",
                    index=False
                )

                errores_final.to_excel(
                    writer,
                    sheet_name="Inconsistencias",
                    index=False
                )


            return render_template(
                "index.html",
                conciliados=conciliados_final.to_html(
                    classes="table table-success",
                    index=False
                ),
                errores=errores_final.to_html(
                    classes="table table-danger",
                    index=False
                )
            )

            return render_template(
                "index.html",

                conciliados=
                conciliados_final.to_html(
                    classes="table table-success",
                    index=False
                ),

                errores=
            errores_final.to_html(
                    classes="table table-danger",
                    index=False
                )
            )

        except Exception as e:

            return f"<h3>Error:</h3><p>{str(e)}</p>"


    return render_template(
        "index.html",
        conciliados=None,
        errores=None
    )


@app.route("/descargar")
def descargar():

    return send_file(
        "reporte_sicca.xlsx",
        as_attachment=True
    )


if __name__=="__main__":
    app.run(debug=True) 