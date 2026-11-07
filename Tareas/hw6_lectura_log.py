"""Funciones de la HW6: leer las cuatro tablas, interpretarlas y llenar el Excel."""
import math
from pathlib import Path

import openpyxl
import pandas as pd
from scipy import stats

ALFA = 0.05
N_OBS = 400

# Cambio en x que se interpreta. Las variables en log se mueven en porcentaje. Con
# y en log, 1,000 km o 1 hp son muy poco para quien vende, así que se leen 10 unidades.
CAMBIO_PCT = 10
PASOS_LOG_Y = {"km_mil": 10, "potencia_hp": 10}

COLUMNAS = ["variable", "coef", "std_err", "t", "p", "ic_bajo", "ic_alto"]
FILAS_COEF = range(6, 13)
DESFASE_FILA_RESPUESTA = 11  # el coeficiente de la fila 7 se responde en la 18, etc.

NOMBRES_X = {"km_mil": "el kilometraje", "potencia_hp": "la potencia"}
DUMMIES = {
    "importado": ("un carro importado", "uno nacional", "importado vs nacional"),
    "automatica": ("un carro automático", "uno manual", "automático vs manual"),
}


def es_log(termino):
    return termino.startswith("np.log(")


def nombre_base(termino):
    return termino.removeprefix("np.log(").removesuffix(")")


def leer_modelos(ruta):
    """Regresa {hoja: {"formula": str, "coefs": DataFrame}}, indexado por fila de Excel."""
    libro = openpyxl.load_workbook(ruta, data_only=True)
    modelos = {}
    for hoja in libro.worksheets[1:]:
        filas = [[hoja.cell(r, c).value for c in range(1, 8)] for r in FILAS_COEF]
        coefs = pd.DataFrame(filas, columns=COLUMNAS, index=FILAS_COEF)
        modelos[hoja.title] = {"formula": hoja["B2"].value, "coefs": coefs}
    return modelos


def resumen_modelos(modelos):
    """Dónde lleva el log cada modelo, leído de la fórmula de su hoja."""
    filas = []
    for nombre, modelo in modelos.items():
        terminos = modelo["coefs"]["variable"]
        filas.append({
            "modelo": nombre,
            "y en log": "Sí" if es_log(modelo["formula"]) else "No",
            "x en log": ", ".join(t for t in terminos if es_log(t)) or "ninguna",
        })
    return pd.DataFrame(filas).style.hide()


def revisar_inferencia(coefs):
    """Recalcula p y compara la decisión por p con la del intervalo de 95%."""
    gl = N_OBS - len(coefs)
    t = (coefs["coef"] / coefs["std_err"]).abs()
    return pd.DataFrame({
        "p_reportado": coefs["p"],
        "p_recalculado": 2 * stats.t.sf(t, gl),
        "sig_por_p": coefs["p"] < ALFA,
        "sig_por_ic": (coefs["ic_bajo"] > 0) | (coefs["ic_alto"] < 0),
    })


def tabla_significancia(modelos):
    """p-value y decisión de cada variable (sin intercepto) en cada modelo."""
    columnas = {}
    for nombre, modelo in modelos.items():
        coefs = modelo["coefs"].iloc[1:]
        decision = coefs["p"].map(lambda p: f"{p:.3f} ({'Sí' if p < ALFA else 'No'})")
        columnas[nombre] = decision.set_axis(coefs["variable"].map(nombre_base))
    return pd.DataFrame(columnas).rename_axis("variable")


def calculo(celda, coef, y_log, x_log, paso=1):
    """Fórmula de Excel de un coeficiente y su valor. Sin fórmula si β se lee directo."""
    crecimiento = 1 + CAMBIO_PCT / 100
    if y_log and x_log:
        return f"={crecimiento}^{celda}-1", crecimiento**coef - 1
    if y_log:
        exponente = celda if paso == 1 else f"{celda}*{paso}"
        return f"=EXP({exponente})-1", math.exp(coef * paso) - 1
    if x_log:
        return f"={celda}*LN({crecimiento})", coef * math.log(crecimiento)
    return None, coef


def etiqueta_cambio(base, x_log, paso):
    """Cambio en x, en corto, para las tablas."""
    if x_log:
        return f"+{CAMBIO_PCT}%"
    return {
        "km_mil": f"+{paso * 1000:,} km",
        "antiguedad_anios": f"+{paso} año",
        "potencia_hp": f"+{paso} hp",
    }[base]


def frase_cambio(base, x_log, paso):
    """Cambio en x, como lo dice la oración."""
    if x_log:
        return f"{NOMBRES_X[base]} sube {CAMBIO_PCT}%"
    return {
        "km_mil": f"el carro trae {paso * 1000:,} km más",
        "antiguedad_anios": f"el carro tiene {paso} año más",
        "potencia_hp": f"el carro trae {paso} hp más",
    }[base]


def oracion(base, valor, y_log, x_log, paso):
    """Una oración de interpretación, con unidades y 'con todo lo demás igual'."""
    if y_log:
        tamano = f"como {abs(valor):.2%}"
    else:
        tamano = f"como {abs(valor) * 1000:,.0f} pesos ({abs(valor):.3f} miles de pesos)"
    if base in DUMMIES:
        sujeto, referencia, _ = DUMMIES[base]
        direccion = "más" if valor > 0 else "menos"
        return f"Por {sujeto} piden {tamano} {direccion} que por {referencia}, con todo lo demás igual."
    verbo = "sube" if valor > 0 else "baja"
    cambio = frase_cambio(base, x_log, paso)
    return f"Si {cambio}, el precio que piden {verbo} {tamano}, con todo lo demás igual."


def resolver_modelo(modelo):
    """Responde una hoja: una fila por variable, sin el intercepto."""
    y_log = es_log(modelo["formula"])
    filas = []
    for fila, coef in modelo["coefs"].iloc[1:].iterrows():
        termino, x_log = coef["variable"], es_log(coef["variable"])
        base = nombre_base(termino)
        respuesta = {
            "variable": termino, "base": base, "fila_respuesta": fila + DESFASE_FILA_RESPUESTA,
            "coef": coef["coef"], "p": coef["p"], "significativa": "No", "paso": 1,
            "formula": None, "valor": None, "cambio": None, "efecto": None,
            "interpretacion": "Not significant",
        }
        if coef["p"] < ALFA:
            paso = PASOS_LOG_Y.get(base, 1) if y_log and not x_log else 1
            formula, valor = calculo(f"B{fila}", coef["coef"], y_log, x_log, paso)
            es_dummy = base in DUMMIES
            respuesta |= {
                "significativa": "Yes", "paso": paso, "formula": formula, "valor": valor,
                "cambio": DUMMIES[base][2] if es_dummy else etiqueta_cambio(base, x_log, paso),
                "efecto": f"{valor:+.2%}" if y_log else f"{valor:+.3f} miles de pesos",
                "interpretacion": oracion(base, valor, y_log, x_log, paso),
            }
        filas.append(respuesta)
    return pd.DataFrame(filas)


def mostrar(tabla):
    """Respuestas de un modelo, tal como van en la hoja."""
    columnas = ["variable", "p", "significativa", "formula", "efecto", "interpretacion"]
    return (
        tabla[columnas]
        .rename(columns={"formula": "fórmula", "interpretacion": "interpretación"})
        .style.hide()
        .format({"p": "{:.3f}"}, na_rep="")
        .set_properties(subset=["interpretación"], **{"text-align": "left"})
    )


def tabla_lecturas(respuestas):
    """Efecto de cada variable por modelo. Las unidades cambian de columna a columna."""
    columnas = {}
    for nombre, tabla in respuestas.items():
        celdas = tabla["cambio"].fillna("Not significant") + ": " + tabla["efecto"].fillna("")
        columnas[nombre] = celdas.str.removesuffix(": ").set_axis(tabla["base"])
    return pd.DataFrame(columnas).rename_axis("variable")


def exacto_vs_aprox(tabla):
    """Atajo contra efecto exacto, para un modelo con y en log (en %)."""
    t = tabla[tabla["significativa"] == "Yes"]
    elasticidad = t["variable"].map(es_log)
    aprox = 100 * t["coef"] * t["paso"].where(~elasticidad, CAMBIO_PCT / 100)
    exacto = 100 * t["valor"]
    return pd.DataFrame({
        "variable": t["variable"],
        "cambio": t["cambio"],
        "coef": t["coef"],
        "aprox_%": aprox.round(3),
        "exacto_%": exacto.round(3),
        "brecha_pts": (exacto - aprox).round(3),
    }).set_index("variable")


def escribir_respuestas(plantilla, salida, respuestas):
    """Guarda una copia de la plantilla con solo las celdas amarillas llenas."""
    libro = openpyxl.load_workbook(plantilla)
    for nombre, tabla in respuestas.items():
        hoja = libro[nombre]
        for r in tabla.itertuples():
            assert hoja.cell(r.fila_respuesta, 1).value == r.variable, (nombre, r.fila_respuesta)
            hoja.cell(r.fila_respuesta, 2).value = r.significativa
            hoja.cell(r.fila_respuesta, 3).value = None if pd.isna(r.formula) else r.formula
            hoja.cell(r.fila_respuesta, 5).value = r.interpretacion
    libro.save(salida)


def recalcular_con_excel(ruta):
    """Deja que Excel calcule y guarde los resultados, para que las vistas previas no salgan vacías."""
    from win32com.client import DispatchEx  # solo en Windows con Excel

    excel = DispatchEx("Excel.Application")  # instancia propia: no toca ventanas abiertas
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        libro = excel.Workbooks.Open(str(Path(ruta).resolve()))
        excel.CalculateFull()
        libro.Save()
        libro.Close(False)
    finally:
        excel.Quit()


def comparar_con_excel(ruta, respuestas):
    """Pone cada resultado calculado por Excel junto al calculado en Python."""
    libro = openpyxl.load_workbook(ruta, data_only=True)
    filas = []
    for nombre, tabla in respuestas.items():
        for r in tabla[tabla["formula"].notna()].itertuples():
            filas.append({
                "hoja": nombre,
                "celda": f"C{r.fila_respuesta}",
                "fórmula": r.formula,
                "python": r.valor,
                "excel": libro[nombre].cell(r.fila_respuesta, 3).value,
            })
    resultado = pd.DataFrame(filas)
    resultado["brecha"] = (resultado["python"] - resultado["excel"]).abs()
    return resultado
