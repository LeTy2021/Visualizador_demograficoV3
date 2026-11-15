"""
Configuración del dashboard demográfico.

Todo lo que puede cambiar con una nueva revisión de las estimaciones y proyecciones, con un
cambio de nombre de los cuadros del INE o con un ajuste de diseño se toca SOLO en este archivo:
rutas, años, códigos de los cuadros, catálogo de indicadores, textos y colores.
"""
from pathlib import Path

# ----------------------------------------------------------------------------
# Rutas (relativas a la carpeta del proyecto)
# ----------------------------------------------------------------------------
CARPETA_PROYECTO = Path(__file__).resolve().parent
DIR_DATOS = CARPETA_PROYECTO / "datos"                       # Excel originales del INE
DIR_PROCESADOS = CARPETA_PROYECTO / "datos_procesados"       # CSV que lee el dashboard
DIR_DOCS = CARPETA_PROYECTO / "docs"                         # PDF técnicos (opcional)
LOGO = CARPETA_PROYECTO / "logo_ine.png"

# Archivos que genera procesamiento.py y lee app.py
ARCHIVOS_PROCESADOS = {
    "ind": "indicadores.csv",
    "ps": "poblacion_edad_simple.csv",
    "pq": "poblacion_quinquenal.csv",
    "cat": "catalogo.csv",
    "tef": "tef.csv",
    "tv": "tablas_mortalidad.csv",
}

# ----------------------------------------------------------------------------
# Revisión y períodos
# ----------------------------------------------------------------------------
REVISION = "2025"
ANIO_INI_EST = 1996          # primer año estimado
ANIO_PROY = 2024             # primer año proyectado
ANIO_FIN_PAIS = 2070         # último año proyectado, total país
ANIO_FIN_DEPTOS = 2045       # último año proyectado, departamentos
ANIO_REF_DEFECTO = 2025      # año que se muestra al abrir el dashboard
EDAD_ABIERTA = 90            # grupo abierto armonizado (estimaciones 90+, proyecciones 100+)
SEXOS = ("Ambos sexos", "Hombres", "Mujeres")
# Tasa neta de migración donde el INE no la publica: si es True se calcula como TCT − TCN.
# Se deja en False para mostrar sólo cifras publicadas.
CALCULAR_TNM = False
NIVEL_REEMPLAZO = 2.1

# ----------------------------------------------------------------------------
# Cuadros del INE: se buscan por el código al inicio del nombre del archivo
# ----------------------------------------------------------------------------
CUADROS_POB_SIMPLE = ("A.1.1_", "B.1.1_")        # población por edad simple (total país)
CUADROS_POB_QUINQ = ("A.1.2_", "B.1.2_")         # población quinquenal por departamento
CUADROS_INDICADORES = ("A.2.1_", "A.2.2_", "A.2.3_", "A.2.4_", "B.2.1.1_", "B.2.1.2_")
CUADRO_TEF, HOJA_TEF = "A.2.4_", "TEF"
CUADRO_TABLAS_VIDA = "Tablas_mortalidad"
HOJAS_TABLAS_VIDA = {"Tabla hombres": "Hombres", "Tabla mujeres": "Mujeres"}
PREFIJO_HOJA_NACIONAL = "Uruguay_"               # hojas nacionales de proyección (llegan a 2070)
MASC_NACIMIENTO_DEFECTO = 1.05                   # si falta la población de 0 años

# ----------------------------------------------------------------------------
# Catálogo de indicadores: código -> (nombre, tema, unidad)
# ----------------------------------------------------------------------------
TEMAS = ["Población", "Crecimiento", "Estructura", "Fecundidad", "Mortalidad"]

CATALOGO = {
    "POB":    ("Población total", "Población", "personas"),
    "TBN":    ("Tasa bruta de natalidad", "Crecimiento", "por mil habitantes"),
    "TBM":    ("Tasa bruta de mortalidad", "Crecimiento", "por mil habitantes"),
    "TCN":    ("Tasa de crecimiento natural", "Crecimiento", "por mil habitantes"),
    "TNM":    ("Tasa neta de migración", "Crecimiento", "por mil habitantes"),
    "TCT":    ("Tasa de crecimiento total", "Crecimiento", "por mil habitantes"),
    "MASC":   ("Relación de masculinidad", "Estructura", "hombres por cada 100 mujeres"),
    "MASC65": ("Relación de masculinidad 65 y más", "Estructura", "hombres por cada 100 mujeres (65+)"),
    "P014":   ("Población de 0 a 14 años", "Estructura", "% de la población"),
    "P1564":  ("Población de 15 a 64 años", "Estructura", "% de la población"),
    "P65":    ("Población de 65 años y más", "Estructura", "% de la población"),
    "IE":     ("Índice de envejecimiento", "Estructura", "personas 65+ por cada 100 de 0-14"),
    "ISE":    ("Índice de sobreenvejecimiento", "Estructura", "personas 85+ por cada 100 de 65+"),
    "RDT":    ("Relación de dependencia total", "Estructura", "por cada 100 personas de 15-64"),
    "RDI":    ("Relación de dependencia de la infancia", "Estructura", "por cada 100 personas de 15-64"),
    "RDV":    ("Relación de dependencia de la vejez", "Estructura", "por cada 100 personas de 15-64"),
    "IRPEA":  ("Índice de reemplazo de la población en edad activa", "Estructura", "personas 60-64 por cada 100 de 15-19"),
    "EMED":   ("Edad mediana de la población", "Estructura", "años"),
    "TFG":    ("Tasa de fecundidad general", "Fecundidad", "nacimientos por mil mujeres de 15-49"),
    "TGF":    ("Tasa global de fecundidad", "Fecundidad", "hijos por mujer"),
    "EMM":    ("Edad media a la maternidad", "Fecundidad", "años"),
    "TEM":    ("Tasa estandarizada de mortalidad", "Mortalidad", "por mil habitantes"),
    "E0":     ("Esperanza de vida al nacer", "Mortalidad", "años"),
    "E65":    ("Esperanza de vida a los 65 años", "Mortalidad", "años"),
}

# Definiciones breves (se muestran en el dashboard)
DEFINICIONES = {
    "POB": "Población residente estimada o proyectada al 30 de junio de cada año.",
    "TBN": "Nacimientos del año por cada mil habitantes.",
    "TBM": "Defunciones del año por cada mil habitantes.",
    "TCN": "Diferencia entre nacimientos y defunciones por cada mil habitantes (TBN − TBM).",
    "TNM": "Saldo entre inmigrantes y emigrantes por cada mil habitantes." +
           (f" Para el período de estimación ({ANIO_INI_EST}-{ANIO_PROY - 1}) se calcula como tasa de "
            "crecimiento total menos tasa de crecimiento natural." if CALCULAR_TNM else ""),
    "TCT": "Variación anual de la población por cada mil habitantes (crecimiento natural + migración neta).",
    "MASC": "Cantidad de hombres por cada 100 mujeres.",
    "MASC65": "Cantidad de hombres de 65 años y más por cada 100 mujeres de esas edades.",
    "P014": "Porcentaje de la población que tiene entre 0 y 14 años.",
    "P1564": "Porcentaje de la población que tiene entre 15 y 64 años.",
    "P65": "Porcentaje de la población que tiene 65 años y más.",
    "IE": "Personas de 65 años y más por cada 100 personas de 0 a 14 años.",
    "ISE": "Personas de 85 años y más por cada 100 personas de 65 años y más.",
    "RDT": "Personas de 0-14 y de 65+ años por cada 100 personas de 15 a 64 años.",
    "RDI": "Personas de 0 a 14 años por cada 100 personas de 15 a 64 años.",
    "RDV": "Personas de 65 años y más por cada 100 personas de 15 a 64 años.",
    "IRPEA": "Personas de 60 a 64 años por cada 100 personas de 15 a 19 años "
             "(los que se retiran de la edad activa frente a los que ingresan).",
    "EMED": "Edad que divide a la población en dos mitades: la mitad de las personas tiene menos "
            "de esa edad y la otra mitad, más.",
    "TFG": "Nacimientos por cada mil mujeres de 15 a 49 años.",
    "TGF": "Número medio de hijos que tendría una mujer al final de su vida fértil según las tasas "
           "de fecundidad por edad del año.",
    "EMM": "Edad media de las madres al tener a sus hijos, ponderada por las tasas de fecundidad por edad.",
    "TEM": "Tasa de mortalidad que se observaría con una estructura por edad estándar; permite "
           "comparar departamentos y años sin el efecto de la estructura por edad.",
    "E0": "Años que en promedio viviría un recién nacido si se mantuvieran las tasas de mortalidad "
          f"por edad del año. Desde {ANIO_PROY} surge de las tablas de mortalidad proyectadas.",
    "E65": "Años que en promedio le quedarían por vivir a una persona de 65 años con las tasas de "
           f"mortalidad por edad del año. Desde {ANIO_PROY} surge de las tablas de mortalidad proyectadas.",
}

# Nombre de hoja (cuadros por departamento A.2.x y B.2.x) -> código
HOJAS = {
    "TBN": "TBN", "TBM": "TBM", "TCN": "TCN", "TNM": "TNM", "TCT": "TCT",
    "Masc total": "MASC", "Masc 65+": "MASC65", "% 0-14": "P014", "% 15-64": "P1564",
    "% 65+": "P65", "Índice envej": "IE", "Índice sobreenvej": "ISE",
    "Depen total": "RDT", "Depen infancia": "RDI", "Depen vejez": "RDV",
    "IRPEA": "IRPEA", "Edad mediana": "EMED",
    "TFG": "TFG", "TGF": "TGF", "EMM": "EMM",
    "estmort": "TEM", "e0": "E0", "e65": "E65",
}

# Columnas de las hojas nacionales de proyección (Uruguay_CRE. y Uruguay_EST) -> (código, sexo)
COLS_URUGUAY = {
    "tasa bruta de natalidad": ("TBN", "Ambos sexos"),
    "tasa bruta de mortalidad": ("TBM", "Ambos sexos"),
    "tasa de crecimiento natural": ("TCN", "Ambos sexos"),
    "tasa neta de migración": ("TNM", "Ambos sexos"),
    "tasa de crecimiento total": ("TCT", "Ambos sexos"),
    "relación masculinidad total": ("MASC", "Ambos sexos"),
    "relación masculinidad 65+ años": ("MASC65", "Ambos sexos"),
    "población 0-14 %": ("P014", "Ambos sexos"),
    "población 15-64 %": ("P1564", "Ambos sexos"),
    "población 65+ %": ("P65", "Ambos sexos"),
    "índice de envejecimiento": ("IE", "Ambos sexos"),
    "índice de sobreenvejecimiento": ("ISE", "Ambos sexos"),
    "relación dependencia total": ("RDT", "Ambos sexos"),
    "relación dependencia infancia": ("RDI", "Ambos sexos"),
    "relación dependencia vejez": ("RDV", "Ambos sexos"),
    "índice reemplazo población edad activa": ("IRPEA", "Ambos sexos"),
    "edad mediana hombres": ("EMED", "Hombres"),
    "edad mediana mujeres": ("EMED", "Mujeres"),
    "edad mediana total": ("EMED", "Ambos sexos"),
}

# ----------------------------------------------------------------------------
# Selecciones por defecto del dashboard
# ----------------------------------------------------------------------------
INDICADORES_CLAVE = [("TGF", "Ambos sexos"), ("E0", None), ("P65", "Ambos sexos"),
                     ("IE", "Ambos sexos"), ("RDT", "Ambos sexos"), ("EMED", "Ambos sexos")]
ANIOS_TEF_DEFECTO = (1996, 2005, 2015)       # más el último disponible
ANIOS_TV_DEFECTO = (2045,)                   # más el primero y el último disponibles
RANGO_DATOS_DEFECTO = (2010, 2030)

# ----------------------------------------------------------------------------
# Tema INE: colores tomados del logo institucional
# ----------------------------------------------------------------------------
AZUL_OSC = "#00396B"     # azul INE (títulos, encabezado, Uruguay)
AZUL = "#0A94D6"         # celeste INE (series principales)
AMARILLO = "#FCD116"     # amarillo INE (líneas decorativas)
NARANJA = "#E0A800"      # dorado (amarillo INE oscurecido para que se lea sobre blanco)
CELESTE = "#7FC4E8"      # celeste claro
AZUL_CLARO = "#E6F3FB"   # fondos suaves
GRIS = "#5F6B7A"
AZUL_GRIS = "#5B7FA6"
TEXTO = "#263238"
FUENTE = "Open Sans"
COLOR_SEXO = {"Hombres": AZUL, "Mujeres": NARANJA, "Ambos sexos": AZUL_OSC}
PALETA = [AZUL_OSC, AZUL, NARANJA, AZUL_GRIS, "#2E8B57", "#8E5EA2", "#C0392B", "#7F8C8D",
          "#16A085", "#8C6D00"]
# Componentes del crecimiento (pestaña Perfil): líneas de natalidad, mortalidad y TCT
COMPONENTES = {"TBN": ("Tasa bruta de natalidad", AZUL),
               "TBM": ("Tasa bruta de mortalidad", NARANJA)}
LINEA_TOTAL = ("TCT", "Tasa de crecimiento total", AZUL_OSC)

# ----------------------------------------------------------------------------
# Textos institucionales y documentación
# ----------------------------------------------------------------------------
TITULO = "Panorama demográfico del Uruguay"
FUENTE_TXT = f"INE, Estimaciones y proyecciones de población (revisión {REVISION})."
PAGINA_INE = "https://www.gub.uy/instituto-nacional-estadistica/estimaciones_proyecciones"
CONTACTO = "difusion@ine.gub.uy"
# Enlaces a cada documento en el sitio del INE (se abren en una pestaña nueva)
URL_DOC_TECNICO = ("https://www5.ine.gub.uy/documents/Demograf%C3%ADayEESS/PDF/Estimaciones%20y%20proyecciones/"
                   "Estimaciones%20y%20proyecciones%20de%20la%20poblaci%C3%B3n%20de%20Uruguay%202012-2070.pdf")

URL_RESUMEN = ("https://www5.ine.gub.uy/documents/Demograf%C3%ADayEESS/PDF/Estimaciones%20y%20proyecciones/"
               "Estimaciones%20y%20proyecciones%20de%20la%20poblaci%C3%B3n%20de%20Uruguay%202012-2070_resumen%20ejecutivo.pdf")
URL_DEFINICIONES = ("https://www5.ine.gub.uy/documents/Demograf%C3%ADayEESS/PDF/Estimaciones%20y%20proyecciones/"
                    "Definiciones%20de%20indicadores%20demogr%C3%A1ficos%202025.pdf")
DOCUMENTOS = [
    # (título, descripción, enlace, palabra clave para buscar el PDF en la carpeta docs/)
    ("Documento técnico", f"Estimaciones y proyecciones de población, revisión {REVISION}: fuentes, "
     "supuestos y metodología.", URL_DOC_TECNICO, "tecnico"),
    ("Resumen ejecutivo", f"Principales resultados de la revisión {REVISION}.", URL_RESUMEN, "resumen"),
    ("Definiciones", "Conceptos e indicadores utilizados.", URL_DEFINICIONES, "definic"),
]
NOTA_PROY_MIGRACION = f"Proyecciones del total país con migración internacional nula desde {ANIO_PROY}."
NOTA_TNM = (f"Migración neta {ANIO_INI_EST}-{ANIO_PROY - 1}: calculada como tasa de crecimiento total "
            "menos tasa de crecimiento natural." if CALCULAR_TNM else "")

# ----------------------------------------------------------------------------
# Servidor
# ----------------------------------------------------------------------------
PUERTO = 8501
