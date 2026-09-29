"""
Genera el libro MetodosNumericos_Excel.xlsx automatizando Microsoft Excel (COM).

Todo el cálculo del libro vive en fórmulas de Excel: este script solo construye
las hojas, los nombres definidos, las funciones LAMBDA, los formatos, los
gráficos, la tabla estructurada, las tablas dinámicas y las segmentaciones.

Requisitos: Windows + Microsoft 365 (Excel 16) + `pip install pywin32 openpyxl`
Uso:        py scripts/generar_libro.py
"""
import os
import re
import time
import openpyxl
from openpyxl.workbook.defined_name import DefinedName
import win32com.client as win32

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ARCHIVO = "MetodosNumericos_Excel.xlsx"
SALIDA = os.path.join(BASE, ARCHIVO)
IMG = os.path.join(BASE, "docs", "img")

# ----------------------------------------------------------------- constantes
XL_EXPRESSION, XL_WHOLE, XL_DECIMAL, XL_BETWEEN = 2, 1, 2, 1
XL_ROW, XL_COL, XL_PAGE = 1, 2, 3
XL_MAX, XL_MIN, XL_AVG = -4136, -4139, -4106
XL_LOG = -4133
XL_SCATTER, XL_SCATTER_LINES, XL_SCATTER_SMOOTH = -4169, 74, 73
XL_LINE_MARKERS, XL_COLUMN, XL_AREA = 65, 51, 1
XL_CENTER = -4108


def rgb(r, g, b):
    return r + g * 256 + b * 65536


AZUL = rgb(31, 78, 121)
BLANCO = rgb(255, 255, 255)
ENTRADA = rgb(255, 242, 204)
SALIDA_C = rgb(226, 239, 218)
SUAVE = rgb(221, 235, 247)
GRIS = rgb(118, 118, 118)
VERDE_OK = rgb(198, 239, 206)
BORDE = rgb(191, 191, 191)

HOJAS = ["Inicio", "Resumen", "Errores", "SEL", "Raíces", "Interpolación",
         "Ajuste", "Integración", "EDO", "Datos", "Análisis"]
COLOR_HOJA = {"Inicio": AZUL, "Resumen": AZUL, "Errores": rgb(91, 155, 213),
              "SEL": rgb(91, 155, 213), "Raíces": rgb(112, 173, 71),
              "Interpolación": rgb(112, 173, 71), "Ajuste": rgb(237, 125, 49),
              "Integración": rgb(237, 125, 49), "EDO": rgb(237, 125, 49),
              "Datos": rgb(165, 165, 165), "Análisis": rgb(255, 192, 0)}

# Funciones con nombre (Administrador de nombres). Cambiarlas aquí o en Excel
# actualiza todas las hojas que las usan. (nombre, fórmula, descripción, parámetros)
LAMBDAS = [
    ("f_raiz", "=LAMBDA(x,EXP(-x)-x)", "Función cuya raíz se busca: f(x) = e^(-x) - x", ["x"]),
    ("df_raiz", "=LAMBDA(x,-EXP(-x)-1)", "Derivada de f_raiz (Newton-Raphson)", ["x"]),
    ("g_raiz", "=LAMBDA(x,EXP(-x))", "Despeje x = g(x) para aproximaciones sucesivas", ["x"]),
    ("f_int", "=LAMBDA(x,x^3-2*x^2+x+2)", "Integrando f(x) = x^3 - 2x^2 + x + 2", ["x"]),
    ("f_edo", "=LAMBDA(x,y,y-x^2+1)", "EDO de primer orden y' = f(x, y)", ["x", "y"]),
    ("y_exacta", "=LAMBDA(x,(x+1)^2-0.5*EXP(x))", "Solución analítica de la EDO", ["x"]),
    ("ERROR_REL", '=LAMBDA(nuevo,anterior,IFERROR(ABS((nuevo-anterior)/nuevo)*100,""))',
     "Error relativo porcentual aproximado entre dos iteraciones", ["nuevo", "anterior"]),
    ("INTERP_LAGRANGE",
     "=LAMBDA(xs,ys,t,LET(n,ROWS(xs),k,SEQUENCE(n),SUM(MAP(k,LAMBDA(i,"
     "INDEX(ys,i)*PRODUCT(IF(k=i,1,(t-xs)/(INDEX(xs,i)-xs))))))))",
     "Polinomio de Lagrange evaluado en t a partir de los puntos (xs, ys)",
     ["xs", "ys", "t", "n", "k", "i"]),
]
FUNCIONES_NUEVAS = {"LAMBDA", "LET", "SEQUENCE", "MAP"}


def formula_a_xml(fx, params):
    """Convierte una fórmula LAMBDA a la sintaxis de almacenamiento de OOXML
    (_xlfn.LAMBDA, _xlpm.parametro)."""
    s = fx[1:]
    s = re.sub(r"\b([A-Z]+)\(", lambda m: ("_xlfn." if m.group(1) in FUNCIONES_NUEVAS else "")
               + m.group(1) + "(", s)
    return re.sub(r"(?<![\w.])(" + "|".join(params) + r")\b(?!\()", r"_xlpm.\1", s)


def libro_base(ruta):
    """El modelo de objetos de Excel (Names.Add) no admite LAMBDA/LET, así que los
    nombres se escriben directamente en el archivo y luego Excel lo abre."""
    base = openpyxl.Workbook()
    base.properties.creator = "Ernesto Jurado"
    base.properties.title = "Métodos Numéricos para Ingenieros"
    base.active.title = HOJAS[0]
    for h in HOJAS[1:]:
        base.create_sheet(h)
    for nm, fx, _, params in LAMBDAS:
        base.defined_names[nm] = DefinedName(nm, attr_text=formula_a_xml(fx, params))
    base.save(ruta)

xl = wb = None



# ------------------------------------------------------------------- helpers
def esperar_excel(limite=120):
    """Espera a que Excel termine de recalcular o actualizar (evita llamadas rechazadas)."""
    fin = time.time() + limite
    while time.time() < fin:
        try:
            if xl.Ready and xl.CalculationState == 0:
                return
        except Exception:
            pass
        time.sleep(0.5)


def nombre(nm, ref, comentario=None):
    # Names.Add interpreta RefersTo con el separador de listas de la configuración regional
    n = wb.Names.Add(Name=nm, RefersTo=ref.replace(",", xl.International[4]))
    if comentario:
        n.Comment = comentario
    return n


def f(ws, celda, formula):
    ws.Range(celda).Formula2 = formula


def v(ws, celda, valor):
    ws.Range(celda).Value = valor


def titulo(ws, texto, sub):
    v(ws, "A1", texto)
    ws.Range("A1").Font.Size = 18
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Color = AZUL
    v(ws, "A2", sub)
    ws.Range("A2").Font.Italic = True
    ws.Range("A2").Font.Color = GRIS


def seccion(ws, celda, texto):
    v(ws, celda, texto)
    c = ws.Range(celda)
    c.Font.Bold = True
    c.Font.Size = 12
    c.Font.Color = AZUL


def encabezado(ws, rango, textos=None):
    r = ws.Range(rango)
    if textos:
        for i, t in enumerate(textos):
            r.Cells(1, i + 1).Value = t
    r.Interior.Color = AZUL
    r.Font.Color = BLANCO
    r.Font.Bold = True
    r.HorizontalAlignment = XL_CENTER
    r.WrapText = True


def borde(ws, rango):
    b = ws.Range(rango).Borders
    for i in (7, 8, 9, 10, 11, 12):
        b(i).LineStyle = 1
        b(i).Color = BORDE


def entrada(ws, rango):
    ws.Range(rango).Interior.Color = ENTRADA
    borde(ws, rango)


def resultado(ws, rango):
    ws.Range(rango).Interior.Color = SALIDA_C
    ws.Range(rango).Font.Bold = True
    borde(ws, rango)


def fmt(ws, rango, formato):
    ws.Range(rango).NumberFormat = formato


def resaltar(ws, rango, formula, color=VERDE_OK):
    fc = ws.Range(rango).FormatConditions.Add(Type=XL_EXPRESSION, Formula1=formula)
    fc.Interior.Color = color
    return fc


def validar_entero(ws, celda, minimo, maximo, msg):
    val = ws.Range(celda).Validation
    val.Delete()
    val.Add(Type=XL_WHOLE, AlertStyle=1, Operator=XL_BETWEEN,
            Formula1=str(minimo), Formula2=str(maximo))
    val.InputTitle = "Dato de entrada"
    val.InputMessage = msg
    val.ErrorMessage = f"Ingrese un entero entre {minimo} y {maximo}."


def validar_decimal(ws, celda, minimo, maximo, msg):
    val = ws.Range(celda).Validation
    val.Delete()
    val.Add(Type=XL_DECIMAL, AlertStyle=1, Operator=XL_BETWEEN,
            Formula1=minimo, Formula2=maximo)
    val.InputTitle = "Dato de entrada"
    val.InputMessage = msg
    val.ErrorMessage = "Valor fuera del rango permitido."


def grafico(ws, ancla, ancho, alto, tit, series, xt, yt, tipo=XL_SCATTER_LINES,
            log_y=False):
    """series: lista de (nombre, xvalues, values, tipo_serie|None)"""
    c = ws.Range(ancla)
    ch = ws.ChartObjects().Add(c.Left, c.Top, ancho, alto).Chart
    ch.ChartType = tipo
    for nom, xs, ys, t in series:
        s = ch.SeriesCollection().NewSeries()
        s.Name = nom
        s.XValues = xs
        s.Values = ys
        if t is not None:
            s.ChartType = t
    ch.HasTitle = True
    ch.ChartTitle.Text = tit
    ch.ChartTitle.Font.Size = 12
    ch.HasLegend = True
    ch.Legend.Position = -4107  # abajo
    for eje, texto in ((1, xt), (2, yt)):
        if texto:
            ch.Axes(eje).HasTitle = True
            ch.Axes(eje).AxisTitle.Text = texto
    if log_y:
        ch.Axes(2).ScaleType = XL_LOG
        ch.Axes(2).Crosses = 4      # el eje X cruza en el mínimo: etiquetas abajo
    return ch


def rango_serie(hoja, rng):
    return f"='{hoja}'!{rng}"


# ============================================================== hoja Errores
def hoja_errores():
    ws = wb.Worksheets("Errores")
    titulo(ws, "Módulo I · Representación de números y teoría del error",
           "Sistemas posicionales, redondeo vs. truncamiento, error de truncamiento en series y error propagado.")

    seccion(ws, "A4", "1. Representación interna de enteros")
    v(ws, "B5", "Número decimal (entero)")
    v(ws, "C5", 156)
    entrada(ws, "C5")
    validar_entero(ws, "C5", 0, 511, "Entero entre 0 y 511 (límite de DEC2BIN).")
    nombre("n_dec", "='Errores'!$C$5")
    for fila, (etq, fx) in enumerate([
        ("Binario", "=DEC2BIN(n_dec)"), ("Octal", "=DEC2OCT(n_dec)"),
        ("Hexadecimal", "=DEC2HEX(n_dec)"), ("Verificación BIN → DEC", "=BIN2DEC(C6)")], start=6):
        v(ws, f"B{fila}", etq)
        f(ws, f"C{fila}", fx)
    resultado(ws, "C6:C9")
    ws.Range("C6:C9").HorizontalAlignment = -4152

    encabezado(ws, "E4:H4", ["Paso", "Dividendo", "Cociente (÷2)", "Residuo"])
    for i in range(10):
        r = 5 + i
        v(ws, f"E{r}", i + 1)
        if i == 0:
            f(ws, f"F{r}", "=n_dec")
        else:
            f(ws, f"F{r}", f'=IF(OR(F{r-1}="",G{r-1}=0),"",G{r-1})')
        f(ws, f"G{r}", f'=IF(F{r}="","",QUOTIENT(F{r},2))')
        f(ws, f"H{r}", f'=IF(F{r}="","",MOD(F{r},2))')
    borde(ws, "E5:H14")
    v(ws, "E15", "Binario por divisiones sucesivas (residuos de abajo hacia arriba):")
    f(ws, "H16", '=LET(r,FILTER(H5:H14,H5:H14<>""),CONCAT(INDEX(r,SEQUENCE(ROWS(r),1,ROWS(r),-1))))')
    resultado(ws, "H16")
    ws.Range("H16").HorizontalAlignment = -4152

    seccion(ws, "A18", "2. Redondeo vs. truncamiento")
    v(ws, "B19", "Valor verdadero")
    f(ws, "C19", "=PI()")
    entrada(ws, "C19")
    fmt(ws, "C19", "0.000000000000")
    nombre("x_real", "='Errores'!$C$19")
    encabezado(ws, "B21:H21", ["Decimales", "Redondeado", "Truncado", "Error abs. redondeo",
                               "Error abs. truncamiento", "Er % redondeo", "Er % truncamiento"])
    for i in range(8):
        r = 22 + i
        v(ws, f"B{r}", i + 1)
        f(ws, f"C{r}", f"=ROUND(x_real,B{r})")
        f(ws, f"D{r}", f"=TRUNC(x_real,B{r})")
        f(ws, f"E{r}", f"=ABS(x_real-C{r})")
        f(ws, f"F{r}", f"=ABS(x_real-D{r})")
        f(ws, f"G{r}", f"=E{r}/ABS(x_real)*100")
        f(ws, f"H{r}", f"=F{r}/ABS(x_real)*100")
    fmt(ws, "C22:D29", "0.00000000")
    fmt(ws, "E22:H29", "0.00E+00")
    borde(ws, "B22:H29")
    ws.Range("E22:F29").FormatConditions.AddDatabar()
    grafico(ws, "J18", 430, 250, "Error relativo: redondeo vs. truncamiento",
            [("Redondeo", ws.Range("B22:B29"), ws.Range("G22:G29"), None),
             ("Truncamiento", ws.Range("B22:B29"), ws.Range("H22:H29"), None)],
            "Decimales conservados", "Er % (escala log)", log_y=True)

    seccion(ws, "A32", "3. Error de truncamiento: serie de Maclaurin de eˣ")
    for fila, etq, val in [(33, "x", 0.5), (34, "Cifras significativas deseadas (n)", 3)]:
        v(ws, f"B{fila}", etq)
        v(ws, f"C{fila}", val)
    entrada(ws, "C33:C34")
    validar_entero(ws, "C34", 1, 10, "Número de cifras significativas (criterio de Scarborough).")
    nombre("x_serie", "='Errores'!$C$33")
    nombre("cs_serie", "='Errores'!$C$34")
    v(ws, "B35", "Tolerancia Es = 0.5 × 10^(2 − n) %")
    f(ws, "C35", "=0.5*10^(2-cs_serie)")
    nombre("es_serie", "='Errores'!$C$35")
    v(ws, "B36", "Valor verdadero eˣ")
    f(ws, "C36", "=EXP(x_serie)")
    nombre("ex_real", "='Errores'!$C$36")
    resultado(ws, "C35:C36")
    fmt(ws, "C36", "0.0000000000")
    encabezado(ws, "B38:G38", ["n", "Término xⁿ/n!", "Suma acumulada", "Et % (verdadero)",
                               "Ea % (aproximado)", "¿Ea < Es?"])
    for i in range(12):
        r = 39 + i
        v(ws, f"B{r}", i)
        f(ws, f"C{r}", f"=x_serie^B{r}/FACT(B{r})")
        f(ws, f"D{r}", f"=SUM(C$39:C{r})")
        f(ws, f"E{r}", f"=ABS(ex_real-D{r})/ex_real*100")
        f(ws, f"F{r}", '=""' if i == 0 else f"=ERROR_REL(D{r},D{r-1})")
        f(ws, f"G{r}", f'=IF(F{r}="","",IF(F{r}<es_serie,"Sí","No"))')
    fmt(ws, "C39:D50", "0.0000000000")
    fmt(ws, "E39:F50", "0.00E+00")
    borde(ws, "B39:G50")
    resaltar(ws, "B39:G50", "=$F39<es_serie")
    ws.Range("G39:G50").HorizontalAlignment = XL_CENTER

    seccion(ws, "A53", "4. Error propagado  (f(x) = x³)")
    filas = [("x̃ (valor aproximado)", 2.5), ("Δx (incertidumbre)", 0.01)]
    for i, (etq, val) in enumerate(filas):
        v(ws, f"B{54+i}", etq)
        v(ws, f"C{54+i}", val)
    entrada(ws, "C54:C55")
    for fila, etq, fx in [(56, "f(x̃) = x̃³", "=C54^3"),
                          (57, "Δf ≈ |f'(x̃)|·Δx = |3x̃²|·Δx", "=ABS(3*C54^2)*C55"),
                          (58, "Intervalo de f(x)", '=TEXT(C56-C57,"0.0000")&"  ≤ f(x) ≤  "&TEXT(C56+C57,"0.0000")'),
                          (59, "Error relativo propagado (%)", "=C57/ABS(C56)*100")]:
        v(ws, f"B{fila}", etq)
        f(ws, f"C{fila}", fx)
    resultado(ws, "C56:C59")
    fmt(ws, "C56:C57", "0.000000")
    fmt(ws, "C59", "0.0000")
    ws.Columns("A").ColumnWidth = 3
    ws.Columns("B").ColumnWidth = 34
    ws.Columns("C:H").ColumnWidth = 16


# ================================================================== hoja SEL
A_SEL = [[3, -0.1, -0.2, 7.85], [0.1, 7, -0.3, -19.3], [0.3, -0.2, 10, 71.4]]


def hoja_sel():
    ws = wb.Worksheets("SEL")
    titulo(ws, "Módulo I · Sistemas de ecuaciones lineales",
           "Solución directa con funciones matriciales, Gauss-Jordan paso a paso y métodos iterativos de Jacobi y Gauss-Seidel.")
    seccion(ws, "A3", "1. Sistema A·x = c")
    encabezado(ws, "B4:E4", ["x1", "x2", "x3", "c"])
    for i, fila in enumerate(A_SEL):
        v(ws, f"A{5+i}", f"Ec. {i+1}")
        for j, val in enumerate(fila):
            ws.Cells(5 + i, 2 + j).Value = val
    entrada(ws, "B5:E7")
    nombre("A_sel", "='SEL'!$B$5:$D$7")
    nombre("c_sel", "='SEL'!$E$5:$E$7")

    encabezado(ws, "G4", ["¿Diagonal dominante?"])
    for i in range(3):
        f(ws, f"G{5+i}", f'=IF(ABS(INDEX(A_sel,{i+1},{i+1}))>SUM(ABS(INDEX(A_sel,{i+1},0)))-ABS(INDEX(A_sel,{i+1},{i+1})),"Sí","No")')
    borde(ws, "G5:G7")
    ws.Range("G5:G7").HorizontalAlignment = XL_CENTER

    encabezado(ws, "I4:J4", ["Variable", "x = A⁻¹·c"])
    for i in range(3):
        v(ws, f"I{5+i}", f"x{i+1}")
    f(ws, "J5", "=MMULT(MINVERSE(A_sel),c_sel)")
    resultado(ws, "J5:J7")
    fmt(ws, "J5:J7", "0.000000")
    v(ws, "I9", "det(A)")
    f(ws, "J9", "=MDETERM(A_sel)")
    f(ws, "J10", '=IF(ABS(J9)>1E-12,"Solución única","Sistema singular")')
    resultado(ws, "J9:J10")
    fmt(ws, "J9", "0.0000")

    v(ws, "L4", "Tolerancia Ea (%)")
    v(ws, "M4", 0.001)
    entrada(ws, "M4")
    nombre("tol_sel", "='SEL'!$M$4")
    fmt(ws, "M4", "0.0000")

    seccion(ws, "A11", "2. Gauss-Jordan (paso a paso sobre la matriz aumentada)")
    # bloque k en filas 13+5k .. 15+5k, encabezado en 12+5k
    for k in range(4):
        h = 12 + 5 * k
        encabezado(ws, f"B{h}:E{h}", ["x1", "x2", "x3", "c"])
        v(ws, f"A{h}", "Matriz aumentada" if k == 0 else f"Pivote fila {k}")
        ws.Range(f"A{h}").Font.Bold = True
        for i in range(3):
            r = h + 1 + i
            for j in range(4):
                col = "BCDE"[j]
                if k == 0:
                    ws.Range(f"{col}{r}").Formula = f"={col}{5+i}"
                    continue
                p = h - 5 + 1           # primera fila del bloque anterior
                piv_r = p + (k - 1)     # fila pivote en el bloque anterior
                piv_c = "BCD"[k - 1]
                if i == k - 1:
                    fx = f"={col}{p+i}/{piv_c}{piv_r}"
                else:
                    fx = f"={col}{p+i}-{piv_c}{p+i}*{col}{piv_r}/{piv_c}{piv_r}"
                ws.Range(f"{col}{r}").Formula = fx
        fmt(ws, f"B{h+1}:E{h+3}", "0.0000")
        borde(ws, f"B{h+1}:E{h+3}")
    resaltar(ws, "B28:D30", "=B28=1", rgb(255, 230, 153))
    encabezado(ws, "G27:H27", ["Variable", "Gauss-Jordan"])
    for i in range(3):
        v(ws, f"G{28+i}", f"x{i+1}")
        f(ws, f"H{28+i}", f"=E{28+i}")
    resultado(ws, "H28:H30")
    fmt(ws, "H28:H30", "0.000000")

    seccion(ws, "A32", "3. Métodos iterativos")
    v(ws, "A33", "Jacobi (usa los valores de la iteración anterior)")
    v(ws, "G33", "Gauss-Seidel (usa los valores recién calculados)")
    encabezado(ws, "A34:E34", ["k", "x1", "x2", "x3", "Ea % (máx.)"])
    encabezado(ws, "G34:K34", ["k", "x1", "x2", "x3", "Ea % (máx.)"])
    ini = [("A", "B", "C", "D", "E"), ("G", "H", "I", "J", "K")]
    for k in range(21):
        r = 35 + k
        for n, (ck, c1, c2, c3, ce) in enumerate(ini):
            v(ws, f"{ck}{r}", k)
            if k == 0:
                for c in (c1, c2, c3):
                    v(ws, f"{c}{r}", 0)
                v(ws, f"{ce}{r}", "")
                continue
            p = r - 1
            if n == 0:   # Jacobi
                x1, x2, x3 = f"{c1}{p}", f"{c2}{p}", f"{c3}{p}"
                f(ws, f"{c1}{r}", f"=(INDEX(c_sel,1)-INDEX(A_sel,1,2)*{x2}-INDEX(A_sel,1,3)*{x3})/INDEX(A_sel,1,1)")
                f(ws, f"{c2}{r}", f"=(INDEX(c_sel,2)-INDEX(A_sel,2,1)*{x1}-INDEX(A_sel,2,3)*{x3})/INDEX(A_sel,2,2)")
                f(ws, f"{c3}{r}", f"=(INDEX(c_sel,3)-INDEX(A_sel,3,1)*{x1}-INDEX(A_sel,3,2)*{x2})/INDEX(A_sel,3,3)")
            else:        # Gauss-Seidel
                f(ws, f"{c1}{r}", f"=(INDEX(c_sel,1)-INDEX(A_sel,1,2)*{c2}{p}-INDEX(A_sel,1,3)*{c3}{p})/INDEX(A_sel,1,1)")
                f(ws, f"{c2}{r}", f"=(INDEX(c_sel,2)-INDEX(A_sel,2,1)*{c1}{r}-INDEX(A_sel,2,3)*{c3}{p})/INDEX(A_sel,2,2)")
                f(ws, f"{c3}{r}", f"=(INDEX(c_sel,3)-INDEX(A_sel,3,1)*{c1}{r}-INDEX(A_sel,3,2)*{c2}{r})/INDEX(A_sel,3,3)")
            f(ws, f"{ce}{r}", f'=IFERROR(MAX(ABS(({c1}{r}:{c3}{r}-{c1}{p}:{c3}{p})/{c1}{r}:{c3}{r}))*100,"")')
    for ck, c1, c2, c3, ce in ini:
        fmt(ws, f"{c1}35:{c3}55", "0.000000")
        fmt(ws, f"{ce}35:{ce}55", "0.00E+00")
        borde(ws, f"{ck}35:{ce}55")
        resaltar(ws, f"{ck}36:{ce}55", f"=${ce}36<tol_sel")

    encabezado(ws, "M34:Q34", ["Método", "Iteraciones", "x1", "x2", "x3"])
    for fila, met, col_e, cols in [(35, "Jacobi", "E", "BCD"), (36, "Gauss-Seidel", "K", "HIJ")]:
        v(ws, f"M{fila}", met)
        f(ws, f"N{fila}", f'=IFERROR(XMATCH(1,ISNUMBER({col_e}36:{col_e}55)*({col_e}36:{col_e}55<tol_sel)),"No converge")')
        for j, c in enumerate(cols):
            f(ws, f"{'OPQ'[j]}{fila}", f'=IFERROR(INDEX({c}36:{c}55,$N{fila}),"—")')
    v(ws, "M37", "Gauss-Jordan")
    v(ws, "N37", "Directo")
    v(ws, "M38", "MINVERSE · MMULT")
    v(ws, "N38", "Directo")
    for j in range(3):
        f(ws, f"{'OPQ'[j]}37", f"=H{28+j}")
        f(ws, f"{'OPQ'[j]}38", f"=INDEX(J5:J7,{j+1})")
    resultado(ws, "M35:Q38")
    fmt(ws, "O35:Q38", "0.000000")
    ws.Range("N35:N38").HorizontalAlignment = XL_CENTER

    grafico(ws, "M40", 430, 260, "Convergencia: Jacobi vs. Gauss-Seidel",
            [("Jacobi", ws.Range("A36:A55"), ws.Range("E36:E55"), None),
             ("Gauss-Seidel", ws.Range("G36:G55"), ws.Range("K36:K55"), None)],
            "Iteración k", "Ea % (escala log)", log_y=True)
    ws.Columns("A").ColumnWidth = 16
    ws.Columns("B:K").ColumnWidth = 12
    ws.Columns("M").ColumnWidth = 18
    ws.Columns("N:Q").ColumnWidth = 12
    ws.Columns("F").ColumnWidth = 3
    ws.Columns("L").ColumnWidth = 16


# =============================================================== hoja Raíces
N_IT = 30


def hoja_raices():
    ws = wb.Worksheets("Raíces")
    titulo(ws, "Módulo II · Raíces de ecuaciones",
           "Intervalo medio (bisección), Regla falsa, Aproximaciones sucesivas y Newton-Raphson sobre la misma función.")
    seccion(ws, "A3", "Parámetros")
    v(ws, "B4", "f(x)")
    v(ws, "C4", "e^(−x) − x   ·  LAMBDA: f_raiz, df_raiz, g_raiz")
    ws.Range("C4").Font.Italic = True
    for fila, etq, val, nm in [(5, "a (límite inferior)", 0, "a_raiz"), (6, "b (límite superior)", 1, "b_raiz"),
                               (7, "x0 (valor inicial)", 0, "x0_raiz"), (8, "Tolerancia Ea (%)", 0.0001, "tol_raiz")]:
        v(ws, f"B{fila}", etq)
        v(ws, f"C{fila}", val)
        nombre(nm, f"='Raíces'!$C${fila}")
    entrada(ws, "C5:C8")
    fmt(ws, "C8", "0.000000")
    v(ws, "B9", "¿f(a)·f(b) < 0?")
    f(ws, "C9", '=IF(f_raiz(a_raiz)*f_raiz(b_raiz)<0,"Sí: hay raíz en [a, b]","No: cambie el intervalo")')
    resultado(ws, "C9")

    encabezado(ws, "E4:I4", ["Método", "Raíz aproximada", "Iteraciones", "f(raíz)", "Ea % final"])
    last = 12 + N_IT
    metodos = [("Intervalo medio", "D", "G"), ("Regla falsa", "N", "P"),
               ("Aproximaciones sucesivas", "T", "U"), ("Newton-Raphson", "AA", "AB")]
    for i, (met, cx, ce) in enumerate(metodos):
        r = 5 + i
        v(ws, f"E{r}", met)
        f(ws, f"G{r}", f'=IFERROR(XMATCH(1,ISNUMBER({ce}13:{ce}{last})*({ce}13:{ce}{last}<tol_raiz)),"No converge")')
        f(ws, f"F{r}", f'=IFERROR(INDEX({cx}13:{cx}{last},G{r}),"—")')
        f(ws, f"H{r}", f'=IFERROR(f_raiz(F{r}),"—")')
        f(ws, f"I{r}", f'=IFERROR(INDEX({ce}13:{ce}{last},G{r}),"—")')
    resultado(ws, "E5:I8")
    fmt(ws, "F5:F8", "0.0000000000")
    fmt(ws, "H5:I8", "0.00E+00")
    ws.Range("G5:G8").HorizontalAlignment = XL_CENTER

    # --- tablas de iteración (filas 13..42)
    v(ws, "A11", "Intervalo medio (bisección)")
    v(ws, "I11", "Regla falsa (Regula-Falsi)")
    v(ws, "R11", "Aproximaciones sucesivas  x = g(x)")
    v(ws, "W11", "Newton-Raphson")
    for c in ("A11", "I11", "R11", "W11"):
        ws.Range(c).Font.Bold = True
    encabezado(ws, "A12:G12", ["i", "a", "b", "xm", "f(a)", "f(xm)", "Ea %"])
    encabezado(ws, "I12:P12", ["i", "a", "b", "f(a)", "f(b)", "xr", "f(xr)", "Ea %"])
    encabezado(ws, "R12:U12", ["i", "xi", "g(xi)", "Ea %"])
    encabezado(ws, "W12:AB12", ["i", "xi", "f(xi)", "f'(xi)", "xi+1", "Ea %"])
    for i in range(N_IT):
        r, p = 13 + i, 12 + i
        for c in ("A", "I", "R", "W"):
            v(ws, f"{c}{r}", i + 1)
        # bisección
        if i == 0:
            f(ws, f"B{r}", "=a_raiz")
            f(ws, f"C{r}", "=b_raiz")
            f(ws, f"G{r}", '=""')
        else:
            f(ws, f"B{r}", f"=IF(E{p}*F{p}<0,B{p},D{p})")
            f(ws, f"C{r}", f"=IF(E{p}*F{p}<0,D{p},C{p})")
            f(ws, f"G{r}", f"=ERROR_REL(D{r},D{p})")
        f(ws, f"D{r}", f"=(B{r}+C{r})/2")
        f(ws, f"E{r}", f"=f_raiz(B{r})")
        f(ws, f"F{r}", f"=f_raiz(D{r})")
        # regla falsa
        if i == 0:
            f(ws, f"J{r}", "=a_raiz")
            f(ws, f"K{r}", "=b_raiz")
            f(ws, f"P{r}", '=""')
        else:
            f(ws, f"J{r}", f"=IF(L{p}*O{p}<0,J{p},N{p})")
            f(ws, f"K{r}", f"=IF(L{p}*O{p}<0,N{p},K{p})")
            f(ws, f"P{r}", f"=ERROR_REL(N{r},N{p})")
        f(ws, f"L{r}", f"=f_raiz(J{r})")
        f(ws, f"M{r}", f"=f_raiz(K{r})")
        f(ws, f"N{r}", f"=K{r}-M{r}*(J{r}-K{r})/(L{r}-M{r})")
        f(ws, f"O{r}", f"=f_raiz(N{r})")
        # aproximaciones sucesivas
        f(ws, f"S{r}", "=x0_raiz" if i == 0 else f"=T{p}")
        f(ws, f"T{r}", f"=g_raiz(S{r})")
        f(ws, f"U{r}", f"=ERROR_REL(T{r},S{r})")
        # Newton-Raphson
        f(ws, f"X{r}", "=x0_raiz" if i == 0 else f"=AA{p}")
        f(ws, f"Y{r}", f"=f_raiz(X{r})")
        f(ws, f"Z{r}", f"=df_raiz(X{r})")
        f(ws, f"AA{r}", f"=X{r}-Y{r}/Z{r}")
        f(ws, f"AB{r}", f"=ERROR_REL(AA{r},X{r})")
    for bloque, num, err in [("A13:G", "B13:F", "G"), ("I13:P", "J13:O", "P"),
                             ("R13:U", "S13:T", "U"), ("W13:AB", "X13:AA", "AB")]:
        borde(ws, f"{bloque}{last}")
        fmt(ws, f"{num}{last}", "0.0000000")
        fmt(ws, f"{err}13:{err}{last}", "0.00E+00")
        resaltar(ws, f"{bloque}{last}", f"=${err}13<tol_raiz")

    # --- datos para gráficos
    encabezado(ws, "AD12:AE12", ["x", "f(x)"])
    for k in range(41):
        r = 13 + k
        v(ws, f"AD{r}", -0.5 + 0.0625 * k)
        f(ws, f"AE{r}", f"=f_raiz(AD{r})")
    encabezado(ws, "AG12:AK12", ["i", "Intervalo medio", "Regla falsa", "Aprox. sucesivas", "Newton-Raphson"])
    for i in range(N_IT):
        r = 13 + i
        f(ws, f"AG{r}", f"=A{r}")
        for c, src in zip(("AH", "AI", "AJ", "AK"), ("G", "P", "U", "AB")):
            f(ws, f"{c}{r}", f"=IF(AND(ISNUMBER({src}{r}),N({src}{r})>0),{src}{r},NA())")
    ws.Range(f"AD13:AK{last+11}").Font.Color = GRIS
    fmt(ws, f"AH13:AK{last}", "0.00E+00")

    grafico(ws, "E46", 430, 260, "f(x) = e^(−x) − x",
            [("f(x)", ws.Range("AD13:AD53"), ws.Range("AE13:AE53"), None)],
            "x", "f(x)", tipo=XL_SCATTER_SMOOTH).HasLegend = False
    grafico(ws, "N46", 520, 260, "Velocidad de convergencia por método",
            [(n, ws.Range(f"AG13:AG{last}"), ws.Range(f"{c}13:{c}{last}"), None)
             for n, c in zip(["Intervalo medio", "Regla falsa", "Aprox. sucesivas", "Newton-Raphson"],
                             ["AH", "AI", "AJ", "AK"])],
            "Iteración", "Ea % (escala log)", log_y=True)
    ws.Columns("A:AK").ColumnWidth = 11
    ws.Columns("B").ColumnWidth = 20
    ws.Columns("C").ColumnWidth = 14
    ws.Columns("E").ColumnWidth = 24
    ws.Columns("F").ColumnWidth = 15
    for c in ("H", "Q", "V", "AC", "AF"):
        ws.Columns(c).ColumnWidth = 3


# ======================================================== hoja Interpolación
def hoja_interpolacion():
    ws = wb.Worksheets("Interpolación")
    titulo(ws, "Módulo II · Interpolación numérica",
           "Estimar ln(2) a partir de cuatro puntos con los polinomios de Lagrange y de Newton (diferencias divididas).")
    seccion(ws, "A3", "Datos")
    encabezado(ws, "B4:C4", ["x", "y = ln(x)"])
    for i, (x, y) in enumerate([(1, 0), (4, 1.386294), (6, 1.791759), (5, 1.609438)]):
        v(ws, f"B{5+i}", x)
        v(ws, f"C{5+i}", y)
    entrada(ws, "B5:C8")
    nombre("xs_int", "='Interpolación'!$B$5:$B$8")
    nombre("ys_int", "='Interpolación'!$C$5:$C$8")
    v(ws, "B10", "xp (punto a interpolar)")
    v(ws, "C10", 2)
    entrada(ws, "C10")
    nombre("xp_int", "='Interpolación'!$C$10")
    v(ws, "B11", "Valor real ln(xp)")
    f(ws, "C11", "=LN(xp_int)")
    resultado(ws, "C11")
    fmt(ws, "C11", "0.000000")

    seccion(ws, "E3", "1. Lagrange")
    encabezado(ws, "E4:I4", ["i", "xi", "yi", "Li(xp)", "yi · Li(xp)"])
    for i in range(4):
        r = 5 + i
        v(ws, f"E{r}", i)
        f(ws, f"F{r}", f"=B{r}")
        f(ws, f"G{r}", f"=C{r}")
        f(ws, f"H{r}", f"=LET(xi,F{r},d,xs_int,PRODUCT(IF(d=xi,1,(xp_int-d)/(xi-d))))")
        f(ws, f"I{r}", f"=G{r}*H{r}")
    borde(ws, "E5:I8")
    fmt(ws, "G5:I8", "0.000000")
    for fila, etq, fx in [(9, "P(xp) = Σ yi·Li", "=SUM(I5:I8)"),
                          (10, "Con la función INTERP_LAGRANGE", "=INTERP_LAGRANGE(xs_int,ys_int,xp_int)"),
                          (11, "Error verdadero %", "=ABS((C11-I9)/C11)*100")]:
        v(ws, f"E{fila}", etq)
        f(ws, f"I{fila}", fx)
    resultado(ws, "I9:I11")
    fmt(ws, "I9:I11", "0.000000")

    seccion(ws, "A14", "2. Newton – diferencias divididas")
    encabezado(ws, "B15:F15", ["xi", "f[xi]", "1er orden", "2do orden", "3er orden"])
    for i in range(4):
        f(ws, f"B{16+i}", f"=B{5+i}")
        f(ws, f"C{16+i}", f"=C{5+i}")
    cols = "CDEF"
    for k in range(1, 4):
        for i in range(4 - k):
            r = 16 + i
            prev = cols[k - 1]
            f(ws, f"{cols[k]}{r}", f"=({prev}{r+1}-{prev}{r})/(B{r+k}-B{r})")
    borde(ws, "B16:F19")
    fmt(ws, "C16:F19", "0.000000")
    ws.Range("C16:F16").Interior.Color = rgb(255, 230, 153)   # coeficientes bk
    encabezado(ws, "H15:K15", ["k", "bk", "Π (xp − xj)", "Término"])
    for k in range(4):
        r = 16 + k
        v(ws, f"H{r}", k)
        f(ws, f"I{r}", f"={cols[k]}16")
        f(ws, f"J{r}", "=1" if k == 0 else f"=PRODUCT(xp_int-$B$16:B{15+k})")
        f(ws, f"K{r}", f"=I{r}*J{r}")
    borde(ws, "H16:K19")
    fmt(ws, "I16:K19", "0.000000")
    v(ws, "H20", "P(xp)")
    f(ws, "K20", "=SUM(K16:K19)")
    v(ws, "H21", "Error verdadero %")
    f(ws, "K21", "=ABS((C11-K20)/C11)*100")
    resultado(ws, "K20:K21")
    fmt(ws, "K20:K21", "0.000000")

    encabezado(ws, "M4:O4", ["x", "P(x)", "ln(x)"])
    for k in range(27):
        r = 5 + k
        v(ws, f"M{r}", 0.5 + 0.25 * k)
        f(ws, f"N{r}", f"=INTERP_LAGRANGE(xs_int,ys_int,M{r})")
        f(ws, f"O{r}", f"=LN(M{r})")
    fmt(ws, "N5:O31", "0.0000")
    ws.Range("M5:O31").Font.Color = GRIS
    ch = grafico(ws, "B24", 500, 280, "Polinomio interpolante vs. ln(x)",
                 [("Polinomio de grado 3", ws.Range("M5:M31"), ws.Range("N5:N31"), XL_SCATTER_SMOOTH),
                  ("ln(x)", ws.Range("M5:M31"), ws.Range("O5:O31"), XL_SCATTER_SMOOTH),
                  ("Puntos dados", ws.Range("B5:B8"), ws.Range("C5:C8"), XL_SCATTER),
                  ("P(xp)", ws.Range("C10"), ws.Range("I9"), XL_SCATTER)],
                 "x", "y")
    ch.SeriesCollection(2).Format.Line.DashStyle = 4
    ch.SeriesCollection(3).MarkerSize = 8
    ch.SeriesCollection(4).MarkerSize = 10
    ch.SeriesCollection(4).MarkerStyle = 2
    ws.Columns("A").ColumnWidth = 3
    ws.Columns("B").ColumnWidth = 22
    ws.Columns("C:K").ColumnWidth = 13
    ws.Columns("E").ColumnWidth = 30
    ws.Columns("L").ColumnWidth = 3


# =============================================================== hoja Ajuste
def hoja_ajuste():
    ws = wb.Worksheets("Ajuste")
    titulo(ws, "Módulo III · Aproximación funcional (mínimos cuadrados)",
           "Ecuaciones normales resueltas con MMULT/MINVERSE, verificación con LINEST y comparación de modelos por R².")
    seccion(ws, "A3", "Datos y sumatorias")
    encabezado(ws, "B4:J4", ["x", "y", "x²", "x³", "x⁴", "x·y", "x²·y", "ln y", "x·ln y"])
    for i, y in enumerate([2.1, 7.7, 13.6, 27.2, 40.9, 61.1]):
        r = 5 + i
        v(ws, f"B{r}", i)
        v(ws, f"C{r}", y)
        for c, fx in zip("DEFGHIJ", [f"=B{r}^2", f"=B{r}^3", f"=B{r}^4", f"=B{r}*C{r}",
                                     f"=B{r}^2*C{r}", f"=LN(C{r})", f"=B{r}*LN(C{r})"]):
            f(ws, f"{c}{r}", fx)
    entrada(ws, "B5:C10")
    borde(ws, "D5:J10")
    v(ws, "A11", "Σ")
    for c in "BCDEFGHIJ":
        f(ws, f"{c}11", f"=SUM({c}5:{c}10)")
    resultado(ws, "B11:J11")
    fmt(ws, "I5:J11", "0.0000")
    v(ws, "A12", "n")
    f(ws, "B12", "=COUNT(B5:B10)")
    resultado(ws, "B12")

    seccion(ws, "L3", "Ecuaciones normales")
    # recta
    v(ws, "L4", "Recta  y = a0 + a1·x")
    ws.Range("L4").Font.Bold = True
    for (r, c), fx in {(5, "L"): "=B12", (5, "M"): "=B11", (6, "L"): "=B11", (6, "M"): "=D11",
                       (5, "O"): "=C11", (6, "O"): "=G11"}.items():
        f(ws, f"{c}{r}", fx)
    f(ws, "Q5", "=MMULT(MINVERSE(L5:M6),O5:O6)")
    # parábola
    v(ws, "L8", "Parábola  y = a0 + a1·x + a2·x²")
    ws.Range("L8").Font.Bold = True
    mat = [["B12", "B11", "D11"], ["B11", "D11", "E11"], ["D11", "E11", "F11"]]
    for i in range(3):
        for j in range(3):
            f(ws, f"{'LMN'[j]}{9+i}", f"={mat[i][j]}")
        f(ws, f"O{9+i}", f"={['C11', 'G11', 'H11'][i]}")
    f(ws, "Q9", "=MMULT(MINVERSE(L9:N11),O9:O11)")
    # exponencial
    v(ws, "L13", "Exponencial  y = a·e^(b·x)  →  ln y = ln a + b·x")
    ws.Range("L13").Font.Bold = True
    for (r, c), fx in {(14, "L"): "=B12", (14, "M"): "=B11", (15, "L"): "=B11", (15, "M"): "=D11",
                       (14, "O"): "=I11", (15, "O"): "=J11"}.items():
        f(ws, f"{c}{r}", fx)
    f(ws, "Q14", "=MMULT(MINVERSE(L14:M15),O14:O15)")
    for fila, et in [(5, "a0"), (6, "a1"), (9, "a0"), (10, "a1"), (11, "a2"), (14, "ln a"), (15, "b")]:
        v(ws, f"P{fila}", et)
    v(ws, "P16", "a")
    f(ws, "Q16", "=EXP(Q14)")
    for rg in ("L5:M6", "L9:N11", "L14:M15", "O5:O6", "O9:O11", "O14:O15"):
        borde(ws, rg)
    resultado(ws, "Q5:Q6")
    resultado(ws, "Q9:Q11")
    resultado(ws, "Q14:Q16")
    fmt(ws, "Q5:Q16", "0.000000")
    ws.Range("P5:P16").HorizontalAlignment = -4152

    v(ws, "L18", "Verificación con LINEST (ESTIMACION.LINEAL)")
    ws.Range("L18").Font.Bold = True
    encabezado(ws, "L19:N19", ["a2", "a1", "a0"])
    f(ws, "L20", "=LINEST(C5:C10,B5:B10^{1,2})")
    ws.Range("L21").Formula2 = ("=IF(MAX(ABS(L20:N20-CHOOSE({1,2,3},Q11,Q10,Q9)))<1E-9,"
                                "\"✔ Coincide con las ecuaciones normales\",\"Revisar\")")
    resultado(ws, "L20:N20")
    fmt(ws, "L20:N20", "0.000000")

    seccion(ws, "A15", "Valores ajustados")
    encabezado(ws, "B16:F16", ["x", "y", "ŷ recta", "ŷ parábola", "ŷ exponencial"])
    for i in range(6):
        r = 17 + i
        f(ws, f"B{r}", f"=B{5+i}")
        f(ws, f"C{r}", f"=C{5+i}")
        f(ws, f"D{r}", f"=$Q$5+$Q$6*B{r}")
        f(ws, f"E{r}", f"=$Q$9+$Q$10*B{r}+$Q$11*B{r}^2")
        f(ws, f"F{r}", f"=$Q$16*EXP($Q$15*B{r})")
    borde(ws, "B17:F22")
    fmt(ws, "D17:F22", "0.0000")
    v(ws, "A23", "R²")
    for c in "DEF":
        f(ws, f"{c}23", f"=1-SUMXMY2($C$17:$C$22,{c}17:{c}22)/DEVSQ($C$17:$C$22)")
    resultado(ws, "D23:F23")
    fmt(ws, "D23:F23", "0.0000")

    seccion(ws, "A25", "Comparación de modelos")
    encabezado(ws, "B26:G26", ["Modelo", "Ecuación ajustada", "", "", "", "R²"])
    modelos = [("Recta", '="y = "&TEXT(Q5,"0.0000")&" + "&TEXT(Q6,"0.0000")&"x"', "=D23"),
               ("Parábola", '="y = "&TEXT(Q9,"0.0000")&" + "&TEXT(Q10,"0.0000")&"x + "&TEXT(Q11,"0.0000")&"x²"', "=E23"),
               ("Exponencial", '="y = "&TEXT(Q16,"0.0000")&"·e^("&TEXT(Q15,"0.0000")&"x)"', "=F23")]
    for r in range(26, 30):
        ws.Range(f"C{r}:F{r}").Merge()
    for i, (m, eq, r2) in enumerate(modelos):
        v(ws, f"B{27+i}", m)
        f(ws, f"C{27+i}", eq)
        f(ws, f"G{27+i}", r2)
    borde(ws, "B27:G29")
    fmt(ws, "G27:G29", "0.0000")
    ws.Range("G27:G29").FormatConditions.AddColorScale(3)
    v(ws, "B30", "Mejor modelo")
    f(ws, "C30", "=XLOOKUP(MAX(G27:G29),G27:G29,B27:B29)")
    resultado(ws, "C30")

    encabezado(ws, "W4:Z4", ["x", "Recta", "Parábola", "Exponencial"])
    for k in range(21):
        r = 5 + k
        v(ws, f"W{r}", 0.25 * k)
        f(ws, f"X{r}", f"=$Q$5+$Q$6*W{r}")
        f(ws, f"Y{r}", f"=$Q$9+$Q$10*W{r}+$Q$11*W{r}^2")
        f(ws, f"Z{r}", f"=$Q$16*EXP($Q$15*W{r})")
    ws.Range("W5:Z25").Font.Color = GRIS
    fmt(ws, "X5:Z25", "0.00")
    ch = grafico(ws, "I25", 480, 280, "Datos y modelos ajustados",
                 [("Datos", ws.Range("B5:B10"), ws.Range("C5:C10"), XL_SCATTER),
                  ("Recta", ws.Range("W5:W25"), ws.Range("X5:X25"), XL_SCATTER_SMOOTH),
                  ("Exponencial", ws.Range("W5:W25"), ws.Range("Z5:Z25"), XL_SCATTER_SMOOTH)],
                 "x", "y")
    s = ch.SeriesCollection(1)
    s.MarkerSize = 8
    tl = s.Trendlines().Add(Type=3, Order=2)   # polinómica de grado 2
    tl.DisplayEquation = True
    tl.DisplayRSquared = True
    tl.Name = "Tendencia polinómica (grado 2)"
    ws.Columns("A").ColumnWidth = 4
    ws.Columns("B:J").ColumnWidth = 10
    ws.Columns("C").ColumnWidth = 12
    ws.Columns("K").ColumnWidth = 3
    ws.Columns("L:Q").ColumnWidth = 11
    ws.Columns("P").ColumnWidth = 6


# ========================================================== hoja Integración
N_MAX = 60


def hoja_integracion():
    ws = wb.Worksheets("Integración")
    titulo(ws, "Módulo III · Integración numérica",
           "Trapecios, Simpson 1/3 y Simpson 3/8 con pesos dinámicos y análisis de sensibilidad con Tabla de datos (Análisis Y si).")
    seccion(ws, "A3", "Parámetros")
    v(ws, "B4", "f(x)")
    v(ws, "C4", "x³ − 2x² + x + 2  ·  LAMBDA f_int")
    ws.Range("C4").Font.Italic = True
    for fila, etq, val, nm in [(5, "a", 0, "a_int"), (6, "b", 3, "b_int"), (7, "n (subintervalos)", 12, "n_int")]:
        v(ws, f"B{fila}", etq)
        v(ws, f"C{fila}", val)
        nombre(nm, f"='Integración'!$C${fila}")
    entrada(ws, "C5:C7")
    validar_entero(ws, "C7", 1, N_MAX, "Número de subintervalos (1 a 60). Simpson 1/3 requiere n par; Simpson 3/8, múltiplo de 3.")
    v(ws, "B8", "h = (b − a) / n")
    f(ws, "C8", "=(b_int-a_int)/n_int")
    nombre("h_int", "='Integración'!$C$8")
    v(ws, "B9", "Referencia (Simpson, n = 1000)")
    f(ws, "C9", "=LET(N,1000,h,(b_int-a_int)/N,i,SEQUENCE(N+1,1,0),x,a_int+i*h,"
                "w,IF((i=0)+(i=N),1,IF(MOD(i,2)=1,4,2)),h/3*SUM(w*f_int(x)))")
    nombre("ref_int", "='Integración'!$C$9")
    v(ws, "B10", "Valor analítico (función original)")
    v(ws, "C10", 12.75)
    resultado(ws, "C8:C10")
    fmt(ws, "C8:C10", "0.000000")

    encabezado(ws, "E4:H4", ["Método", "Resultado", "Error absoluto", "Error %"])
    fil = "C14:C74"
    for r, met, fx in [
        (5, "Trapecios", f"=h_int/2*SUMPRODUCT({fil},D14:D74)"),
        (6, "Simpson 1/3", f'=IF(MOD(n_int,2)=0,h_int/3*SUMPRODUCT({fil},E14:E74),"Requiere n par")'),
        (7, "Simpson 3/8", f'=IF(MOD(n_int,3)=0,3*h_int/8*SUMPRODUCT({fil},F14:F74),"Requiere n múltiplo de 3")')]:
        v(ws, f"E{r}", met)
        f(ws, f"F{r}", fx)
        f(ws, f"G{r}", f'=IFERROR(ABS(F{r}-ref_int),"—")')
        f(ws, f"H{r}", f'=IFERROR(G{r}/ABS(ref_int)*100,"—")')
    resultado(ws, "E5:H7")
    fmt(ws, "F5:F7", "0.000000")
    fmt(ws, "G5:H7", "0.00E+00")

    seccion(ws, "A12", "Puntos muestrales y pesos")
    encabezado(ws, "A13:F13", ["i", "xi", "f(xi)", "w Trapecio", "w Simpson 1/3", "w Simpson 3/8"])
    for i in range(N_MAX + 1):
        r = 14 + i
        v(ws, f"A{r}", i)
        f(ws, f"B{r}", f'=IF(A{r}<=n_int,a_int+A{r}*h_int,"")')
        f(ws, f"C{r}", f'=IF(B{r}="","",f_int(B{r}))')
        f(ws, f"D{r}", f'=IF(B{r}="","",IF(OR(A{r}=0,A{r}=n_int),1,2))')
        f(ws, f"E{r}", f'=IF(B{r}="","",IF(OR(A{r}=0,A{r}=n_int),1,IF(MOD(A{r},2)=1,4,2)))')
        f(ws, f"F{r}", f'=IF(B{r}="","",IF(OR(A{r}=0,A{r}=n_int),1,IF(MOD(A{r},3)=0,2,3)))')
    fmt(ws, "B14:C74", "0.0000")
    borde(ws, "A14:F74")
    resaltar(ws, "A14:F74", '=$B14=""', BLANCO).Font.Color = rgb(200, 200, 200)

    seccion(ws, "J11", "Sensibilidad al número de subintervalos (Tabla de datos)")
    encabezado(ws, "J12:Q12", ["n", "Trapecios", "Simpson 1/3", "Simpson 3/8", "",
                               "Error % Trap.", "Error % S 1/3", "Error % S 3/8"])
    ws.Range("N12").Interior.ColorIndex = -4142
    v(ws, "J13", "actual →")
    f(ws, "K13", "=F5")
    f(ws, "L13", "=F6")
    f(ws, "M13", "=F7")
    ws.Range("J13:M13").Font.Color = GRIS
    for i, n in enumerate([6, 12, 18, 24, 30, 36, 48, 60]):
        r = 14 + i
        v(ws, f"J{r}", n)
        for c_err, c_val in zip("OPQ", "KLM"):
            f(ws, f"{c_err}{r}", f"=ABS({c_val}{r}-ref_int)/ABS(ref_int)*100")
    ws.Range("J13:M21").Table(ColumnInput=ws.Range("C7"))
    fmt(ws, "K13:M21", "0.000000")
    fmt(ws, "O14:Q21", "0.00E+00")
    borde(ws, "J14:M21")
    borde(ws, "O14:Q21")

    # columnas auxiliares: #N/A no se grafica, así el gráfico se ajusta a n
    encabezado(ws, "S13:T13", ["xi (gráfico)", "f(xi) (gráfico)"])
    for i in range(N_MAX + 1):
        r = 14 + i
        f(ws, f"S{r}", f'=IF(B{r}="",NA(),B{r})')
        f(ws, f"T{r}", f'=IF(C{r}="",NA(),C{r})')
    ws.Range("S14:T74").Font.Color = GRIS
    fmt(ws, "S14:T74", "0.0000")
    ch = grafico(ws, "H24", 430, 250, "f(x) y puntos muestrales (n subintervalos)",
                 [("f(xi)", ws.Range("S14:S74"), ws.Range("T14:T74"), None)],
                 "x", "f(x)")
    ch.HasLegend = False
    grafico(ws, "O24", 430, 250, "Error % vs. n",
            [(n, ws.Range("J14:J21"), ws.Range(f"{c}14:{c}21"), None)
             for n, c in [("Trapecios", "O"), ("Simpson 1/3", "P"), ("Simpson 3/8", "Q")]],
            "n", "Error % (escala log)", log_y=True)
    ws.Columns("A").ColumnWidth = 6
    ws.Columns("B").ColumnWidth = 28
    ws.Columns("C:H").ColumnWidth = 14
    ws.Columns("I").ColumnWidth = 3
    ws.Columns("J:Q").ColumnWidth = 13
    ws.Columns("N").ColumnWidth = 3


# ================================================================== hoja EDO
N_EDO = 50


def hoja_edo():
    ws = wb.Worksheets("EDO")
    titulo(ws, "Módulo III · Solución numérica de ecuaciones diferenciales",
           "Euler, Euler mejorado (Heun) y Runge-Kutta de 4.º orden comparados con la solución exacta.")
    seccion(ws, "A3", "Problema de valor inicial")
    v(ws, "B4", "y' = f(x, y)")
    v(ws, "C4", "y − x² + 1   ·  LAMBDA f_edo")
    ws.Range("C4").Font.Italic = True
    for fila, etq, val, nm in [(5, "x0", 0, "x0_edo"), (6, "y0 = y(x0)", 0.5, "y0_edo"),
                               (7, "x final", 2, "xf_edo"), (8, "h (paso)", 0.2, "h_edo")]:
        v(ws, f"B{fila}", etq)
        v(ws, f"C{fila}", val)
        nombre(nm, f"='EDO'!$C${fila}")
    entrada(ws, "C5:C8")
    validar_decimal(ws, "C8", "=1/100", "=5", "Tamaño de paso h > 0. Se admiten hasta 50 pasos.")
    v(ws, "B9", "n = (xf − x0) / h")
    f(ws, "C9", f"=MIN(ROUND((xf_edo-x0_edo)/h_edo,0),{N_EDO})")
    nombre("n_edo", "='EDO'!$C$9")
    v(ws, "B10", "Solución exacta")
    v(ws, "C10", "(x+1)² − 0.5·eˣ  ·  LAMBDA y_exacta")
    ws.Range("C10").Font.Italic = True
    resultado(ws, "C9")

    encabezado(ws, "E4:H4", ["Método", "y(x final)", "Error % en x final", "Error % máximo"])
    fin = 13 + N_EDO
    for r, met, cy, ce in [(5, "Euler", "C", "L"), (6, "Euler mejorado (Heun)", "E", "M"), (7, "Runge-Kutta 4", "J", "N")]:
        v(ws, f"E{r}", met)
        f(ws, f"F{r}", f"=INDEX({cy}13:{cy}{fin},n_edo+1)")
        f(ws, f"G{r}", f"=INDEX({ce}13:{ce}{fin},n_edo+1)")
        f(ws, f"H{r}", f"=MAX({ce}14:{ce}{fin})")
    v(ws, "E8", "Solución exacta")
    f(ws, "F8", "=y_exacta(x0_edo+n_edo*h_edo)")
    resultado(ws, "E5:H8")
    fmt(ws, "F5:F8", "0.000000")
    fmt(ws, "G5:H7", "0.00E+00")
    ws.Range("G5:G7").FormatConditions.AddIconSetCondition()

    encabezado(ws, "A12:N12", ["i", "xi", "Euler", "Heun y*", "Heun", "k1", "k2", "k3", "k4",
                               "RK4", "Exacta", "Er % Euler", "Er % Heun", "Er % RK4"])
    for i in range(N_EDO + 1):
        r, p = 13 + i, 12 + i
        v(ws, f"A{r}", i)
        if i == 0:
            f(ws, f"B{r}", "=x0_edo")
            for c in "CEJ":
                f(ws, f"{c}{r}", "=y0_edo")
            for c in "DFGHI":
                v(ws, f"{c}{r}", "")
        else:
            f(ws, f"B{r}", f'=IF(A{r}<=n_edo,x0_edo+A{r}*h_edo,"")')
            g = f'IF($B{r}="","",'
            f(ws, f"C{r}", f"={g}C{p}+h_edo*f_edo(B{p},C{p}))")
            f(ws, f"D{r}", f"={g}E{p}+h_edo*f_edo(B{p},E{p}))")
            f(ws, f"E{r}", f"={g}E{p}+h_edo/2*(f_edo(B{p},E{p})+f_edo(B{r},D{r})))")
            f(ws, f"F{r}", f"={g}f_edo(B{p},J{p}))")
            f(ws, f"G{r}", f"={g}f_edo(B{p}+h_edo/2,J{p}+h_edo/2*F{r}))")
            f(ws, f"H{r}", f"={g}f_edo(B{p}+h_edo/2,J{p}+h_edo/2*G{r}))")
            f(ws, f"I{r}", f"={g}f_edo(B{p}+h_edo,J{p}+h_edo*H{r}))")
            f(ws, f"J{r}", f"={g}J{p}+h_edo/6*(F{r}+2*G{r}+2*H{r}+I{r}))")
        f(ws, f"K{r}", f'=IF($B{r}="","",y_exacta(B{r}))')
        for ce, cy in zip("LMN", "CEJ"):
            f(ws, f"{ce}{r}", f'=IF($B{r}="","",ABS(({cy}{r}-K{r})/K{r})*100)')
    fmt(ws, f"B13:B{fin}", "0.00")
    fmt(ws, f"C13:K{fin}", "0.000000")
    fmt(ws, f"L13:N{fin}", "0.00E+00")
    borde(ws, f"A13:N{fin}")
    resaltar(ws, f"A13:N{fin}", '=$B13=""', BLANCO).Font.Color = rgb(200, 200, 200)
    ws.Range(f"C13:C{fin}").Interior.Color = rgb(252, 228, 214)
    ws.Range(f"E13:E{fin}").Interior.Color = rgb(255, 242, 204)
    ws.Range(f"J13:J{fin}").Interior.Color = rgb(226, 239, 218)

    # columnas auxiliares para el gráfico (#N/A fuera del rango de n)
    encabezado(ws, "AA12:AE12", ["x", "Euler", "Heun", "RK4", "Exacta"])
    for i in range(N_EDO + 1):
        r = 13 + i
        for c, src in zip(("AA", "AB", "AC", "AD", "AE"), "BCEJK"):
            f(ws, f"{c}{r}", f'=IF($B{r}="",NA(),{src}{r})')
    ws.Range(f"AA13:AE{fin}").Font.Color = GRIS
    fmt(ws, f"AA13:AE{fin}", "0.0000")
    X = ws.Range(f"AA13:AA{fin}")
    ch = grafico(ws, "P4", 480, 290, "Aproximaciones vs. solución exacta",
                 [(n, X, ws.Range(f"{c}13:{c}{fin}"), None) for n, c in
                  [("Euler", "AB"), ("Heun", "AC"), ("RK4", "AD"), ("Exacta", "AE")]],
                 "x", "y")
    ch.SeriesCollection(4).Format.Line.DashStyle = 4
    ws.Columns("A").ColumnWidth = 5
    ws.Columns("B").ColumnWidth = 16
    ws.Columns("C:N").ColumnWidth = 11
    ws.Columns("E").ColumnWidth = 22
    ws.Columns("O").ColumnWidth = 3


# ======================================================== hoja Datos (tabla)
def hoja_datos():
    ws = wb.Worksheets("Datos")
    cab = ["Tipo", "Módulo", "Tema", "Método", "Iteración", "Aproximación", "Error_pct", "Conv", "Estado"]
    filas = []

    def agrega(tipo, mod, tema, met, it, aprox, err):
        filas.append([tipo, mod, tema, met, it, aprox, err, "", ""])

    for met, cx, ce in [("Jacobi", "B", "E"), ("Gauss-Seidel", "H", "K")]:
        for k in range(1, 21):
            agrega("Iterativo", "Módulo I", "Sistemas de ecuaciones", met, k,
                   f"='SEL'!{cx}{35+k}", f"='SEL'!{ce}{35+k}")
    for met, cx, ce in [("Intervalo medio", "D", "G"), ("Regla falsa", "N", "P"),
                        ("Aproximaciones sucesivas", "T", "U"), ("Newton-Raphson", "AA", "AB")]:
        for i in range(1, N_IT + 1):
            agrega("Iterativo", "Módulo II", "Raíces de ecuaciones", met, i,
                   f"='Raíces'!{cx}{12+i}", f"='Raíces'!{ce}{12+i}")
    for met, cx, ce in [("Euler", "C", "L"), ("Euler mejorado (Heun)", "E", "M"), ("Runge-Kutta 4", "J", "N")]:
        for i in range(1, 21):
            agrega("Discretización", "Módulo III", "Ecuaciones diferenciales", met,
                   f"=IF('EDO'!B{13+i}=\"\",\"\",{i})", f"='EDO'!{cx}{13+i}", f"='EDO'!{ce}{13+i}")
    for met, cx, ce in [("Trapecios", "K", "O"), ("Simpson 1/3", "L", "P"), ("Simpson 3/8", "M", "Q")]:
        for i in range(8):
            agrega("Discretización", "Módulo III", "Integración numérica", met,
                   f"='Integración'!J{14+i}", f"='Integración'!{cx}{14+i}", f"='Integración'!{ce}{14+i}")

    n = len(filas)
    ws.Range(ws.Cells(1, 1), ws.Cells(1, len(cab))).Value = cab
    ws.Range(ws.Cells(2, 1), ws.Cells(n + 1, len(cab))).Formula = filas
    lo = ws.ListObjects.Add(SourceType=1, Source=ws.Range(ws.Cells(1, 1), ws.Cells(n + 1, len(cab))),
                            XlListObjectHasHeaders=1)
    lo.Name = "tblConvergencia"
    lo.TableStyle = "TableStyleMedium2"
    lo.ListColumns("Conv").DataBodyRange.Formula2 = (
        '=IF([@Tipo]="Iterativo",IF(AND(ISNUMBER([@[Error_pct]]),N([@[Error_pct]])<'
        'IF([@Tema]="Sistemas de ecuaciones",tol_sel,tol_raiz)),1,0),0)')
    lo.ListColumns("Estado").DataBodyRange.Formula2 = (
        '=IF([@Tipo]<>"Iterativo","N/A",LET(prev,COUNTIFS(INDEX([Método],1):[@Método],[@Método],'
        'INDEX([Conv],1):[@Conv],1),IF(prev=0,"Iterando",IF(AND([@Conv]=1,prev=1),"Convergió","Post-convergencia"))))')
    fmt(ws, f"F2:F{n+1}", "0.000000")
    fmt(ws, f"G2:G{n+1}", "0.00E+00")
    ws.Columns("A:I").ColumnWidth = 16
    ws.Columns("D").ColumnWidth = 26
    ws.Columns("C").ColumnWidth = 26
    ws.Range("K1").Value = "Tabla estructurada tblConvergencia: origen de las tablas dinámicas de la hoja Análisis."
    ws.Range("K2").Value = "Se alimenta con fórmulas de las hojas de cada método: al cambiar un parámetro, use Datos > Actualizar todo."
    ws.Range("K1:K2").Font.Color = GRIS
    return lo


# ============================================ hoja Análisis (tablas dinámicas)
def hoja_analisis():
    ws = wb.Worksheets("Análisis")
    titulo(ws, "Análisis de convergencia con tablas dinámicas",
           "Origen: tabla tblConvergencia (hoja Datos). Use las segmentaciones para filtrar por tema o método.")
    cache = wb.PivotCaches().Create(SourceType=1, SourceData="tblConvergencia", Version=5)

    # --- TD 1: iteraciones y error al converger
    pt1 = cache.CreatePivotTable(TableDestination="'Análisis'!R7C1", TableName="tdIteraciones", DefaultVersion=5)
    pt1.PivotFields("Tipo").Orientation = XL_PAGE
    pt1.PivotFields("Tipo").CurrentPage = "Iterativo"
    est = pt1.PivotFields("Estado")
    est.Orientation = XL_PAGE
    est.EnableMultiplePageItems = True
    for it in ("Post-convergencia", "N/A"):
        try:
            est.PivotItems(it).Visible = False
        except Exception:
            pass
    pt1.PivotFields("Tema").Orientation = XL_ROW
    pt1.PivotFields("Método").Orientation = XL_ROW
    d = pt1.AddDataField(pt1.PivotFields("Iteración"), "Iteraciones hasta converger", XL_MAX)
    d.NumberFormat = "0"
    d = pt1.AddDataField(pt1.PivotFields("Error_pct"), "Ea % al converger", XL_MIN)
    d.NumberFormat = "0.00E+00"
    d = pt1.AddDataField(pt1.PivotFields("Error_pct"), "Ea % promedio", XL_AVG)
    d.NumberFormat = "0.00E+00"
    pt1.RowAxisLayout(1)
    pt1.TableStyle2 = "PivotStyleMedium2"

    # --- TD 2: Ea % por iteración y método (para gráfico dinámico)
    pt2 = cache.CreatePivotTable(TableDestination="'Análisis'!R7C7", TableName="tdCurvas", DefaultVersion=5)
    pt2.PivotFields("Tipo").Orientation = XL_PAGE
    pt2.PivotFields("Tipo").CurrentPage = "Iterativo"
    est = pt2.PivotFields("Estado")
    est.Orientation = XL_PAGE
    est.EnableMultiplePageItems = True
    for it in ("Post-convergencia", "N/A"):
        try:
            est.PivotItems(it).Visible = False
        except Exception:
            pass
    pt2.PivotFields("Iteración").Orientation = XL_ROW
    pt2.PivotFields("Método").Orientation = XL_COL
    d = pt2.AddDataField(pt2.PivotFields("Error_pct"), "Ea % ", XL_AVG)
    d.NumberFormat = "0.00E+00"
    pt2.ColumnGrand = False
    pt2.DisplayErrorString = True   # la iteración 1 de los métodos de intervalo no tiene Ea
    pt2.ErrorString = ""
    pt2.RowGrand = False
    pt2.TableStyle2 = "PivotStyleLight16"

    # --- TD 3: métodos de discretización
    pt3 = cache.CreatePivotTable(TableDestination="'Análisis'!R26C1", TableName="tdDiscretizacion", DefaultVersion=5)
    pt3.PivotFields("Tipo").Orientation = XL_PAGE
    pt3.PivotFields("Tipo").CurrentPage = "Discretización"
    pt3.PivotFields("Tema").Orientation = XL_ROW
    pt3.PivotFields("Método").Orientation = XL_ROW
    d = pt3.AddDataField(pt3.PivotFields("Error_pct"), "Error % máximo", XL_MAX)
    d.NumberFormat = "0.00E+00"
    d = pt3.AddDataField(pt3.PivotFields("Error_pct"), "Error % promedio", XL_AVG)
    d.NumberFormat = "0.00E+00"
    pt3.RowAxisLayout(1)
    pt3.TableStyle2 = "PivotStyleMedium4"

    # --- gráficos dinámicos
    ancla = ws.Range("R4")
    sh = ws.Shapes.AddChart2(-1, XL_LINE_MARKERS, ancla.Left, ancla.Top, 520, 300)
    ch = sh.Chart
    ch.SetSourceData(pt2.TableRange1)
    ch.HasTitle = True
    ch.ChartTitle.Text = "Ea % por iteración (gráfico dinámico)"
    ch.Axes(2).ScaleType = XL_LOG
    ch.Axes(2).Crosses = 4
    ch.Axes(2).HasTitle = True
    ch.Axes(2).AxisTitle.Text = "Ea % (escala log)"
    ch.Axes(1).HasTitle = True
    ch.Axes(1).AxisTitle.Text = "Iteración"
    ch.HasLegend = True
    ch.Legend.Position = -4107
    ch.ShowAllFieldButtons = False
    ancla = ws.Range("R26")
    sh = ws.Shapes.AddChart2(-1, XL_COLUMN, ancla.Left, ancla.Top, 520, 300)
    ch2 = sh.Chart
    ch2.SetSourceData(pt3.TableRange1)
    ch2.HasTitle = True
    ch2.ChartTitle.Text = "Error % máximo y promedio por método"
    ch2.Axes(2).ScaleType = XL_LOG
    ch2.Axes(2).Crosses = 4
    ch2.Axes(2).HasTitle = True
    ch2.Axes(2).AxisTitle.Text = "Error % (escala log)"
    ch2.HasLegend = True
    ch2.Legend.Position = -4107
    ch2.ShowAllFieldButtons = False

    # --- segmentaciones
    sc_tema = wb.SlicerCaches.Add2(pt2, "Tema")
    sc_tema.PivotTables.AddPivotTable(pt1)
    c = ws.Range("AD4")
    sc_tema.Slicers.Add(SlicerDestination=ws, Name="segTema", Caption="Tema",
                        Top=c.Top, Left=c.Left, Width=190, Height=110)
    sc_met = wb.SlicerCaches.Add2(pt2, "Método")
    sc_met.PivotTables.AddPivotTable(pt1)
    c = ws.Range("AD12")
    sc_met.Slicers.Add(SlicerDestination=ws, Name="segMetodo", Caption="Método",
                       Top=c.Top, Left=c.Left, Width=190, Height=210)
    ws.Columns("A").ColumnWidth = 24
    ws.Columns("B").ColumnWidth = 26
    ws.Columns("C:E").ColumnWidth = 14
    cache.RefreshOnFileOpen = True
    return ch, ch2


# ============================================================= hoja Resumen
def hoja_resumen():
    ws = wb.Worksheets("Resumen")
    titulo(ws, "Resumen de resultados",
           "Cada fila se calcula en vivo desde la hoja del método correspondiente.")
    encabezado(ws, "B9:H9", ["Módulo", "Tema", "Método", "Resultado", "Referencia", "Error %", "Iteraciones / pasos"])
    ref_raiz = f"'Raíces'!AA{12+N_IT}"
    vec = lambda rng: f'="("&TEXTJOIN("; ",,TEXT({rng},"0.0000"))&")"'
    filas = [
        ("Módulo I", "Sistemas de ecuaciones", "Gauss-Jordan", vec("'SEL'!H28:H30"), vec("'SEL'!J5:J7"),
         "=MAX(ABS(('SEL'!H28:H30-'SEL'!J5:J7)/'SEL'!J5:J7))*100", "Directo"),
        ("Módulo I", "Sistemas de ecuaciones", "Jacobi", vec("'SEL'!O35:Q35"), vec("'SEL'!J5:J7"),
         "=MAX(ABS(('SEL'!O35:Q35-TRANSPOSE('SEL'!J5:J7))/TRANSPOSE('SEL'!J5:J7)))*100", "='SEL'!N35"),
        ("Módulo I", "Sistemas de ecuaciones", "Gauss-Seidel", vec("'SEL'!O36:Q36"), vec("'SEL'!J5:J7"),
         "=MAX(ABS(('SEL'!O36:Q36-TRANSPOSE('SEL'!J5:J7))/TRANSPOSE('SEL'!J5:J7)))*100", "='SEL'!N36"),
    ]
    for i, met in enumerate(["Intervalo medio", "Regla falsa", "Aproximaciones sucesivas", "Newton-Raphson"]):
        filas.append(("Módulo II", "Raíces de ecuaciones", met, f"='Raíces'!F{5+i}", f"={ref_raiz}",
                      f"=ABS(('Raíces'!F{5+i}-{ref_raiz})/{ref_raiz})*100", f"='Raíces'!G{5+i}"))
    filas += [
        ("Módulo II", "Interpolación", "Lagrange", "='Interpolación'!I9", "='Interpolación'!C11",
         "='Interpolación'!I11", "4 puntos"),
        ("Módulo II", "Interpolación", "Newton (dif. divididas)", "='Interpolación'!K20", "='Interpolación'!C11",
         "='Interpolación'!K21", "4 puntos"),
    ]
    for i, met in enumerate(["Recta", "Parábola", "Exponencial"]):
        filas.append(("Módulo III", "Mínimos cuadrados", met, f"='Ajuste'!G{27+i}", 1,
                      f"=(1-'Ajuste'!G{27+i})*100", "R² (Resultado)"))
    for i, met in enumerate(["Trapecios", "Simpson 1/3", "Simpson 3/8"]):
        filas.append(("Módulo III", "Integración numérica", met, f"='Integración'!F{5+i}", "=ref_int",
                      f"='Integración'!H{5+i}", "=\"n = \"&n_int"))
    for i, met in enumerate(["Euler", "Euler mejorado (Heun)", "Runge-Kutta 4"]):
        filas.append(("Módulo III", "Ecuaciones diferenciales", met, f"='EDO'!F{5+i}", "='EDO'!F8",
                      f"='EDO'!G{5+i}", "=n_edo&\" pasos\""))
    fin = 9 + len(filas)
    for i, fila in enumerate(filas):
        r = 10 + i
        for j, val in enumerate(fila):
            c = ws.Cells(r, 2 + j)
            if isinstance(val, str) and val.startswith("="):
                c.Formula2 = val
            else:
                c.Value = val
    borde(ws, f"B10:H{fin}")
    fmt(ws, f"E10:F{fin}", "0.000000")
    fmt(ws, f"G10:G{fin}", "0.00E+00")
    ws.Range(f"H10:H{fin}").HorizontalAlignment = XL_CENTER
    ws.Range(f"E10:F{fin}").HorizontalAlignment = -4152
    # semáforo: verde < 0.01 %, amarillo < 1 %, rojo >= 1 %
    for regla, color in [(f"=$G10<1/100", VERDE_OK), (f"=$G10<1", rgb(255, 235, 156)),
                         (f"=$G10>=1", rgb(255, 199, 206))]:
        resaltar(ws, f"B10:H{fin}", regla, color).StopIfTrue = True
    for r in range(10, fin + 1):
        if ws.Cells(r, 2).Value != ws.Cells(r - 1, 2).Value and r > 10:
            ws.Range(f"B{r}:H{r}").Borders(8).Weight = -4138

    # KPI
    kpis = [
        ("B4", "Métodos implementados", f"=ROWS(D10:D{fin})", "0"),
        ("D4", "Método de raíces más rápido",
         f"=LET(m,D13:D16,it,H13:H16,XLOOKUP(MIN(it),it,m))", "@"),
        ("F4", "Menor error en EDO", f"=LET(m,D{fin-2}:D{fin},e,G{fin-2}:G{fin},XLOOKUP(MIN(e),e,m))", "@"),
        ("H4", "Mejor ajuste (R²)", "='Ajuste'!C30", "@"),
    ]
    for celda, etq, fx, nf in kpis:
        col = celda[0]
        sig = chr(ord(col) + 1)
        c = ws.Range(celda)
        c.Value = etq
        c.Font.Color = GRIS
        c.Font.Size = 10
        k = ws.Range(f"{col}5")
        k.Formula2 = fx
        k.NumberFormat = nf
        k.Font.Size = 16
        k.Font.Bold = True
        k.Font.Color = AZUL
        k.HorizontalAlignment = -4131
        blk = ws.Range(f"{col}4:{sig}6")
        blk.Interior.Color = SUAVE
    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = 12
    ws.Columns("C").ColumnWidth = 25
    ws.Columns("D").ColumnWidth = 25
    ws.Columns("E:F").ColumnWidth = 30
    ws.Columns("G").ColumnWidth = 12
    ws.Columns("H").ColumnWidth = 20


# =============================================================== hoja Inicio
def hoja_inicio():
    ws = wb.Worksheets("Inicio")
    titulo(ws, "Métodos Numéricos para Ingenieros — Libro interactivo en Excel",
           "Proyecto integrador de la asignatura 0638 · Universidad Tecnológica de Panamá · Autor: Ernesto Jurado")
    seccion(ws, "B4", "Cómo usar el libro")
    pasos = [
        "1. Modifique solo las celdas amarillas (datos de entrada); todo lo demás se recalcula con fórmulas.",
        "2. Las funciones matemáticas están definidas como LAMBDA en Fórmulas > Administrador de nombres (p. ej. f_raiz, f_int, f_edo).",
        "3. Las celdas verdes son resultados; las filas en verde claro de las tablas iterativas indican que se alcanzó la tolerancia.",
        "4. Tras cambiar parámetros, actualice las tablas dinámicas con Datos > Actualizar todo (Ctrl + Alt + F5).",
    ]
    for i, t in enumerate(pasos):
        v(ws, f"B{5+i}", t)
    v(ws, "B10", "Entrada")
    entrada(ws, "B10")
    v(ws, "C10", "Resultado")
    resultado(ws, "C10")

    seccion(ws, "B12", "Contenido")
    encabezado(ws, "B13:D13", ["Hoja", "Tema del curso", "Herramientas de Excel aplicadas"])
    indice = [
        ("Resumen", "Resultados de todos los métodos", "KPI con LET/XLOOKUP, TEXTJOIN, semáforo con formato condicional"),
        ("Errores", "Módulo I: representación de números y teoría del error", "DEC2BIN/HEX, QUOTIENT, FILTER + SEQUENCE, ROUND vs TRUNC, barras de datos, validación"),
        ("SEL", "Módulo I: Gauss-Jordan, Jacobi, Gauss-Seidel", "MMULT, MINVERSE, MDETERM, XMATCH con matrices, formato condicional por fórmula"),
        ("Raíces", "Módulo II: bisección, regla falsa, aprox. sucesivas, Newton-Raphson", "Funciones LAMBDA con nombre, ERROR_REL, gráfico logarítmico"),
        ("Interpolación", "Módulo II: Lagrange y Newton", "LAMBDA + MAP + SEQUENCE (INTERP_LAGRANGE), fórmulas de matriz dinámica"),
        ("Ajuste", "Módulo III: mínimos cuadrados (recta, parábola, exponencial)", "Ecuaciones normales matriciales, LINEST, línea de tendencia, escala de color"),
        ("Integración", "Módulo III: trapecios, Simpson 1/3 y 3/8", "SUMPRODUCT con pesos, Tabla de datos (Análisis Y si), gráficos que se ajustan a n con NA()"),
        ("EDO", "Módulo III: Euler, Heun, Runge-Kutta 4", "LAMBDA de dos variables, íconos de semáforo, gráfico dinámico con NA()"),
        ("Datos", "Registro de todas las iteraciones", "Tabla estructurada, referencias estructuradas, COUNTIFS acumulado"),
        ("Análisis", "Comparación de convergencia", "Tablas dinámicas, gráficos dinámicos y segmentaciones"),
    ]
    for i, (h, tema, herr) in enumerate(indice):
        r = 14 + i
        ws.Hyperlinks.Add(Anchor=ws.Range(f"B{r}"), Address="", SubAddress=f"'{h}'!A1", TextToDisplay=h)
        v(ws, f"C{r}", tema)
        v(ws, f"D{r}", herr)
    borde(ws, f"B14:D{13+len(indice)}")

    seccion(ws, "B26", "Funciones LAMBDA definidas en el libro")
    encabezado(ws, "B27:D27", ["Nombre", "Descripción", "Definición"])
    for i, (nm, fx, desc, _) in enumerate(LAMBDAS):
        r = 28 + i
        v(ws, f"B{r}", nm)
        v(ws, f"C{r}", desc)
        v(ws, f"D{r}", "'" + fx)
    borde(ws, f"B28:D{27+len(LAMBDAS)}")
    ws.Range(f"D28:D{27+len(LAMBDAS)}").Font.Name = "Consolas"
    ws.Range(f"D28:D{27+len(LAMBDAS)}").Font.Size = 9
    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = 20
    ws.Columns("C").ColumnWidth = 58
    ws.Columns("D").ColumnWidth = 95
    ws.Range("B14:D40").WrapText = True
    ws.Range("B14:D40").VerticalAlignment = -4160


# ================================================================ imágenes
def exportar_imagenes(graficos):
    os.makedirs(IMG, exist_ok=True)
    for nombre_png, ch in graficos:
        ruta = os.path.join(IMG, nombre_png)
        ch.Parent.Parent.Activate()
        for _ in range(5):
            ch.Refresh()
            ch.Export(ruta)
            if os.path.getsize(ruta) > 0:
                break
            time.sleep(1)
        if os.path.getsize(ruta) == 0:
            # algunas exportaciones salen vacías: se copia el gráfico a una hoja temporal
            tmp = wb.Worksheets.Add()
            ch.Parent.Copy()
            tmp.Paste()
            tmp.ChartObjects(1).Chart.Export(ruta)
            tmp.Delete()


def captura_rango(ws, rango, nombre_png):
    """Exporta un rango como imagen (copia como imagen y la pega en un gráfico vacío)."""
    ws.Activate()
    r = ws.Range(rango)
    r.CopyPicture(1, 2)
    co = ws.ChartObjects().Add(0, 0, r.Width + 4, r.Height + 4)
    co.Activate()
    co.Chart.Paste()
    co.Chart.Export(os.path.join(IMG, nombre_png))
    co.Delete()


def main():
    global xl, wb
    # instancia aislada de Excel con enlace temprano (makepy), necesario para
    # pasar argumentos con nombre a la API
    xl = win32.gencache.EnsureDispatch(win32.DispatchEx("Excel.Application"))
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.ScreenUpdating = False
    try:
        if os.path.exists(SALIDA):
            os.remove(SALIDA)
        libro_base(SALIDA)
        wb = xl.Workbooks.Open(SALIDA)
        xl.Calculation = -4135      # manual mientras se construye (evita llamadas rechazadas)
        for h in HOJAS:
            wb.Worksheets(h).Tab.Color = COLOR_HOJA[h]
        for nm, _, desc, _ in LAMBDAS:
            wb.Names(nm).Comment = desc

        hoja_errores()
        hoja_sel()
        hoja_raices()
        hoja_interpolacion()
        hoja_ajuste()
        hoja_integracion()
        hoja_edo()
        xl.CalculateFull()
        hoja_datos()
        xl.CalculateFull()
        ch_piv, ch_piv2 = hoja_analisis()
        hoja_resumen()
        hoja_inicio()
        xl.Calculation = -4105      # automático
        xl.CalculateFull()
        wb.RefreshAll()
        esperar_excel()
        xl.CalculateFull()
        esperar_excel()

        xl.ScreenUpdating = True
        for ws in wb.Worksheets:
            ws.Activate()
            xl.ActiveWindow.DisplayGridlines = False
            xl.ActiveWindow.Zoom = 90
            ws.Range("A1").Select()

        os.makedirs(IMG, exist_ok=True)
        graficos = [
            ("raices_convergencia.png", wb.Worksheets("Raíces").ChartObjects(2).Chart),
            ("sel_convergencia.png", wb.Worksheets("SEL").ChartObjects(1).Chart),
            ("interpolacion.png", wb.Worksheets("Interpolación").ChartObjects(1).Chart),
            ("ajuste_modelos.png", wb.Worksheets("Ajuste").ChartObjects(1).Chart),
            ("integracion_error_n.png", wb.Worksheets("Integración").ChartObjects(2).Chart),
            ("edo_comparacion.png", wb.Worksheets("EDO").ChartObjects(1).Chart),
            ("analisis_grafico_dinamico.png", ch_piv),
            ("analisis_discretizacion.png", ch_piv2),
        ]
        exportar_imagenes(graficos)
        for hoja, rango, png in [("Resumen", "B2:H33", "resumen.png"),
                                 ("Raíces", "A11:G34", "raices_tabla.png"),
                                 ("Análisis", "A4:E40", "analisis_tablas.png")]:
            try:
                captura_rango(wb.Worksheets(hoja), rango, png)
            except Exception as e:  # la captura depende del portapapeles
                print("Aviso: no se pudo capturar", png, e)

        wb.Worksheets("Inicio").Activate()
        wb.Save()
        print("Libro generado:", SALIDA)
    finally:
        if wb is not None:
            wb.Close(False)
        xl.Quit()


if __name__ == "__main__":
    main()
