# -*- coding: utf-8 -*-
"""Reporte mensual de inventario, mermas y compras.

Pensado para correr tal cual en Google Colab (monta Drive y descarga el PDF
al terminar) o en local/CI apuntando BASE_DIR a cualquier carpeta con el
Excel de origen. Este repo incluye datos sinteticos en datos_demo/ para que
se pueda ejecutar de principio a fin sin depender de informacion real de
ningun negocio.
"""

# Coneccion con Google Drive (solo aplica si se corre dentro de Colab)
try:
    from google.colab import drive
    drive.mount('/content/drive')
    EN_COLAB = True
except Exception:
    EN_COLAB = False

import importlib.util
import subprocess
import sys

PAQUETES = {
    'reportlab': 'reportlab',
    'openpyxl': 'openpyxl',
    'xlsxwriter': 'xlsxwriter',
}
faltantes = [paquete for modulo, paquete in PAQUETES.items() if importlib.util.find_spec(modulo) is None]
if faltantes:
    subprocess.check_call([sys.executable, '-m', 'pip', '-q', 'install', *faltantes])

print('Entorno listo:', 'Google Colab' if EN_COLAB else 'Python local')

# Librerias
import re
import unicodedata
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd
import matplotlib
if not EN_COLAB:
    matplotlib.use('Agg')  # sin ventanas emergentes al correr como script local
import matplotlib.pyplot as plt
import matplotlib.ticker as mtick

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.lib.utils import ImageReader
from reportlab.platypus import (
    Image,
    KeepTogether,
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

#Configuracion del reporte
# ESTABLECIMIENTO, ARCHIVO_EXCEL_NOMBRE y BASE_DIR se ajustan por negocio.
# Los valores de abajo apuntan al Excel sintetico de datos_demo/ para que el
# script corra de inmediato; en produccion se cambian por la carpeta y el
# archivo reales (por ejemplo, la carpeta de auditorias en Google Drive).
ESTABLECIMIENTO = 'DEMO'  # @param {type:"string"}
ANIO = 2026  # @param {type:"integer"}
MES = 7  # @param {type:"integer", min:1, max:12}
ARCHIVO_EXCEL_NOMBRE = 'inventario_demo.xlsx'  # @param {type:"string"}
DESCARGA_AUTOMATICA = True  # @param {type:"boolean"}

if EN_COLAB:
    BASE_DIR = Path(f'/content/drive/MyDrive/AUDITORIAS/{ESTABLECIMIENTO}')
else:
    BASE_DIR = Path(__file__).resolve().parent.parent / 'datos_demo'
ARCHIVO_EXCEL = BASE_DIR / ARCHIVO_EXCEL_NOMBRE

# Dia y duracion de la semana operativa del negocio (0=lunes ... 6=domingo).
# Ejemplo: un negocio que opera viernes-sabado (y a veces jueves) usa 3/2.
DIA_INICIO_SEMANA = 3
DIAS_DURACION_SEMANA = 2
COBERTURA_OBJETIVO_DIAS = 2
# Parametros de priorizacion y abastecimiento
ABC_UMBRAL_A = 0.80
ABC_UMBRAL_B = 0.95
NIVEL_SERVICIO_Z = 1.65
LEAD_TIME_DIAS_DEFAULT = COBERTURA_OBJETIVO_DIAS
LEAD_TIME_DIAS_POR_PRODUCTO = {
    # 'NOMBRE DEL PRODUCTO': 3,
}

# Parametros de auditoria
UMBRAL_ALERTA_UNIDADES = 0.25
UMBRAL_ALERTA_COSTO = 100
MARGEN_COSTO_RIESGO = 15
UMBRAL_SEVERIDAD_MEDIA = 250
UMBRAL_SEVERIDAD_ALTA = 1000

# visuales
COLOR_PRIMARIO = '#111827'
COLOR_SECUNDARIO = '#C28B36'
COLOR_ACENTO = '#2563EB'
COLOR_EXITO = '#15803D'
COLOR_ALERTA = '#D97706'
COLOR_RIESGO = '#B91C1C'
LOGO_PATH = BASE_DIR / 'logo.png'

MESES_NOMBRE = {
    1: 'ENERO', 2: 'FEBRERO', 3: 'MARZO', 4: 'ABRIL',
    5: 'MAYO', 6: 'JUNIO', 7: 'JULIO', 8: 'AGOSTO',
    9: 'SEPTIEMBRE', 10: 'OCTUBRE', 11: 'NOVIEMBRE', 12: 'DICIEMBRE',
}
if MES not in MESES_NOMBRE:
    raise ValueError('MES debe estar entre 1 y 12.')
if not ARCHIVO_EXCEL.exists():
    raise FileNotFoundError(f'No se encontro el archivo: {ARCHIVO_EXCEL}')

NOMBRE_MES = MESES_NOMBRE[MES]
NOMBRE_PERIODO = f"{ESTABLECIMIENTO}_{NOMBRE_MES}_{ANIO}".replace(' ', '_')
NOMBRE_PERIODO_ARCHIVO = NOMBRE_PERIODO.lower()
TITULO_REPORTE = f"{ESTABLECIMIENTO} - {NOMBRE_MES} {ANIO}"

OUT_BASE_DIR = BASE_DIR if EN_COLAB else Path(__file__).resolve().parent.parent / 'ejemplo_salida'
PROCESADO_DIR = OUT_BASE_DIR / 'procesado'
REPORTES_DIR = OUT_BASE_DIR / 'reportes'
OUT_DIR = PROCESADO_DIR / NOMBRE_PERIODO
GRAFICAS_DIR = OUT_DIR / 'graficas'
for carpeta in [PROCESADO_DIR, REPORTES_DIR, OUT_DIR, GRAFICAS_DIR]:
    carpeta.mkdir(parents=True, exist_ok=True)

RUTA_EXCEL_SALIDA = OUT_DIR / f"reporte_base_{NOMBRE_PERIODO_ARCHIVO}_v3.xlsx"
RUTA_PDF = REPORTES_DIR / f"{ESTABLECIMIENTO.title()} Reporte {NOMBRE_MES.title()}.pdf"

print('Negocio:', ESTABLECIMIENTO)
print('Periodo:', TITULO_REPORTE)
print('Fuente:', ARCHIVO_EXCEL)
print('PDF:', RUTA_PDF)
print('Excel:', RUTA_EXCEL_SALIDA)

#diccionario y limpieza
MESES = {
    'ENERO': 1,
    'FEBRERO': 2,
    'MARZO': 3,
    'ABRIL': 4,
    'MAYO': 5,
    'JUNIO': 6,
    'JULIO': 7,
    'AGOSTO': 8,
    'SEPTIEMBRE': 9,
    'SETIEMBRE': 9,
    'OCTUBRE': 10,
    'NOVIEMBRE': 11,
    'DICIEMBRE': 12,
}

DIAS = {
    'LUNES',
    'MARTES',
    'MIERCOLES',
    'MIERCOLES',
    'JUEVES',
    'VIERNES',
    'SABADO',
    'SABADO',
    'DOMINGO',
}


def quitar_acentos(texto):
    texto = str(texto)
    texto = unicodedata.normalize('NFKD', texto)
    return ''.join(c for c in texto if not unicodedata.combining(c))


def normalizar_texto(texto):
    texto = quitar_acentos(str(texto).upper().strip())
    texto = re.sub(r'\s+', ' ', texto)
    return texto


def limpiar_numero(valor):
    if pd.isna(valor):
        return np.nan
    if isinstance(valor, (int, float, np.integer, np.floating)):
        return float(valor)
    texto = str(valor).strip()
    if texto == '':
        return np.nan
    errores_excel = {'#REF!', '#NAME?', '#DIV/0!', '#VALUE!', '#N/A', '#NULL!'}
    if texto.upper() in errores_excel:
        return np.nan
    texto = texto.replace('$', '').replace(' ', '')
    if re.match(r'^-?\d{1,3}(\.\d{3})+,\d+$', texto):
        texto = texto.replace('.', '').replace(',', '.')
    elif re.match(r'^-?\d+,\d+$', texto):
        texto = texto.replace(',', '.')
    else:
        texto = texto.replace(',', '')
    return pd.to_numeric(texto, errors='coerce')


def formato_pesos(valor):
    try:
        return f"${float(valor):,.2f}"
    except Exception:
        return str(valor)


def formato_num(valor):
    try:
        return f"{float(valor):,.2f}"
    except Exception:
        return str(valor)


def formato_pct(valor):
    try:
        return f"{float(valor) * 100:,.1f}%"
    except Exception:
        return str(valor)


def formato_fecha(valor):
    try:
        return pd.to_datetime(valor).strftime('%d/%m/%Y')
    except Exception:
        return str(valor)

#lectura de excel
def leer_metadata_hoja(nombre_hoja, anio=2026):
    limpio = normalizar_texto(nombre_hoja)
    partes = limpio.split()

    if limpio in {'SEMANA', 'RESUMEN', 'CATALOGO'}:
        return None

    indice_mes = None
    for i, parte in enumerate(partes):
        if parte in MESES:
            indice_mes = i
            break
    if indice_mes is None:
        return None

    establecimiento = ' '.join(partes[:indice_mes]).strip()
    mes_nombre = partes[indice_mes]
    mes_numero = MESES[mes_nombre]
    resto = partes[indice_mes + 1:]

    dia_semana = None
    for parte in resto:
        if parte in DIAS:
            dia_semana = parte
            break

    numeros = [int(parte) for parte in resto if parte.isdigit()]
    if len(numeros) == 0:
        return None

    dia_numero = numeros[-1]
    try:
        fecha = pd.Timestamp(year=anio, month=mes_numero, day=dia_numero)
    except Exception:
        return None

    return {
        'establecimiento': establecimiento,
        'mes_nombre': mes_nombre,
        'mes': mes_numero,
        'dia_semana': dia_semana,
        'dia_numero': dia_numero,
        'fecha': fecha,
    }


COLUMNAS = [
    'producto',
    'precio_compra',
    'presentacion_unidades',
    'precio_unitario_excel',
    'inventario_inicial_botellas',
    'inventario_inicial_botellas_uni',
    'inventario_inicial_suelto',
    'inventario_inicial_total_excel',
    'compras_caja_bot',
    'compras_unidades_excel',
    'disponible_total_excel',
    'ventas_botellas',
    'ventas_botellas_unidades_excel',
    'ventas_tragos',
    'cortesias_merma',
    'salidas_total_excel',
    'inventario_final_botellas',
    'inventario_final_botellas_uni',
    'tara_peso_actual_oz',
    'peso_botella_vacia_oz',
    'inventario_final_suelto',
    'inventario_final_total_excel',
    'inventario_teorico_excel',
    'diferencia_unidades_excel',
    'diferencia_costo_excel',
    'costo_almacen_final_excel',
    'costo_salidas_total_excel',
    'costo_inventario_inicial_excel',
    'costo_inicial_mas_compras_excel',
    'costo_compras_excel',
]


def leer_hoja_diaria(ruta_excel, nombre_hoja, anio=2026):
    metadata = leer_metadata_hoja(nombre_hoja, anio=anio)
    if metadata is None:
        return None

    df_raw = pd.read_excel(
        ruta_excel,
        sheet_name=nombre_hoja,
        header=None,
        engine='openpyxl',
    )
    df_raw = df_raw.reindex(columns=range(30))
    df_raw = df_raw.iloc[:, :30].copy()
    df_raw.columns = COLUMNAS

    df = df_raw.iloc[2:].copy()
    df['fila_excel'] = df.index + 1
    df = df[df['producto'].notna()].copy()
    df['producto'] = df['producto'].astype(str).str.strip()
    df = df[df['producto'] != '']
    df = df[~df['producto'].str.upper().eq('PRODUCTO')]

    df['archivo_origen'] = Path(ruta_excel).name
    df['hoja_origen'] = nombre_hoja
    df['establecimiento'] = metadata['establecimiento']
    df['fecha'] = metadata['fecha']
    df['mes'] = metadata['mes']
    df['mes_nombre'] = metadata['mes_nombre']
    df['dia_semana'] = metadata['dia_semana']
    df['dia_numero'] = metadata['dia_numero']

    columnas_numericas = [col for col in COLUMNAS if col != 'producto']
    for col in columnas_numericas:
        df[col] = df[col].apply(limpiar_numero)
    df[columnas_numericas] = df[columnas_numericas].fillna(0)

    return df


def cargar_excel_completo(ruta_excel, establecimiento=None, mes=None, anio=2026):
    archivo = pd.ExcelFile(ruta_excel, engine='openpyxl')
    lista_hojas = []
    hojas_leidas = []

    for hoja in archivo.sheet_names:
        df_hoja = leer_hoja_diaria(ruta_excel, hoja, anio=anio)
        if df_hoja is None:
            continue
        if establecimiento is not None:
            df_hoja = df_hoja[df_hoja['establecimiento'].str.upper() == establecimiento.upper()]
        if mes is not None:
            df_hoja = df_hoja[df_hoja['mes'] == mes]
        if len(df_hoja) > 0:
            lista_hojas.append(df_hoja)
            hojas_leidas.append(hoja)

    if len(lista_hojas) == 0:
        raise ValueError('No se encontraron hojas validas con los filtros indicados.')

    df = pd.concat(lista_hojas, ignore_index=True)
    print('Hojas procesadas:')
    for hoja in hojas_leidas:
        print('-', hoja)
    return df

#cargar y primeros calculos
inventario = cargar_excel_completo(
    ruta_excel=ARCHIVO_EXCEL,
    establecimiento=ESTABLECIMIENTO,
    mes=MES,
    anio=ANIO,
)

print('Filas cargadas:', len(inventario))
print('Fecha inicial:', inventario['fecha'].min())
print('Fecha final:', inventario['fecha'].max())

df = inventario.copy()

# Precio unitario confiable
df['precio_unitario_calc'] = np.where(
    df['presentacion_unidades'] > 0,
    df['precio_compra'] / df['presentacion_unidades'],
    0,
)
df['precio_unitario'] = np.where(
    df['precio_unitario_excel'] > 0,
    df['precio_unitario_excel'],
    df['precio_unitario_calc'],
)

# Inventario inicial y final
df['inventario_inicial_unidades'] = np.where(
    df['inventario_inicial_total_excel'] > 0,
    df['inventario_inicial_total_excel'],
    df['inventario_inicial_botellas_uni'] + df['inventario_inicial_suelto'],
)

df['compras_unidades'] = df['compras_unidades_excel']
condicion_compras_vacias = df['compras_unidades'] == 0
df.loc[condicion_compras_vacias, 'compras_unidades'] = (
    df.loc[condicion_compras_vacias, 'compras_caja_bot']
    * df.loc[condicion_compras_vacias, 'presentacion_unidades']
)

df['disponible_unidades'] = df['inventario_inicial_unidades'] + df['compras_unidades']
df['ventas_botellas_unidades'] = df['ventas_botellas'] * df['presentacion_unidades']
df['ventas_unidades'] = df['ventas_botellas_unidades'] + df['ventas_tragos']
df['cortesias_merma_unidades'] = df['cortesias_merma']
df['salidas_totales_unidades_calc'] = df['ventas_unidades'] + df['cortesias_merma_unidades']
df['salidas_totales_unidades'] = np.where(
    df['salidas_total_excel'] > 0,
    df['salidas_total_excel'],
    df['salidas_totales_unidades_calc'],
)

df['inventario_final_unidades'] = np.where(
    df['inventario_final_total_excel'] > 0,
    df['inventario_final_total_excel'],
    df['inventario_final_botellas_uni'] + df['inventario_final_suelto'],
)

df['inventario_teorico_unidades'] = df['disponible_unidades'] - df['salidas_totales_unidades']
df['diferencia_unidades_calculada'] = df['inventario_final_unidades'] - df['inventario_teorico_unidades']

# La columna X del Excel es la referencia oficial de diferencia en unidades.
# Si X es cero, una diferencia monetaria aislada en Y no genera faltante ni sobrante.
df['diferencia_unidades'] = df['diferencia_unidades_excel']
df['diferencia_costo'] = np.where(
    df['diferencia_unidades'].abs() > 0,
    np.where(
        df['diferencia_costo_excel'].abs() > 0,
        df['diferencia_costo_excel'],
        df['diferencia_unidades'] * df['precio_unitario'],
    ),
    0,
)

df['costo_inventario_inicial'] = df['inventario_inicial_unidades'] * df['precio_unitario']
df['costo_compras'] = df['compras_unidades'] * df['precio_unitario']
df['costo_ventas'] = df['ventas_unidades'] * df['precio_unitario']
df['costo_cortesias_merma'] = df['cortesias_merma_unidades'] * df['precio_unitario']
df['costo_salidas_total'] = df['salidas_totales_unidades'] * df['precio_unitario']
df['costo_inventario_final'] = df['inventario_final_unidades'] * df['precio_unitario']

#clasificacion
def clasificar_producto(row):
    producto = str(row['producto']).upper()
    tiene_datos_botella = (
        row['inventario_inicial_botellas'] > 0
        or row['inventario_final_botellas'] > 0
        or row['ventas_botellas'] > 0
        or row['tara_peso_actual_oz'] > 0
    )
    if tiene_datos_botella:
        return 'BOTELLA'
    palabras_cerveza = [
        'TECATE', 'XX', 'DOS EQUIS', 'INDIO', 'HEINEKEN', 'AMSTEL',
        'ULTRA', 'MODELO', 'CORONA', 'VICTORIA', 'CARTA BLANCA', 'LAGUNITAS',
    ]
    if any(palabra in producto for palabra in palabras_cerveza):
        return 'CERVEZA'
    palabras_refresco = [
        'COCA', 'SPRITE', 'FRESCA', 'AGUA', 'RED BULL', 'TONIC', 'GINGER',
        'MINERAL', 'JUGO', 'SQUIRT', 'PENAFIEL', 'TOPO CHICO',
    ]
    if any(palabra in producto for palabra in palabras_refresco):
        return 'REFRESCO / MEZCLADOR'
    return 'MISELANEA / INSUMO'


def obtener_semana_operativa(fechas, dia_inicio=2):
    return fechas - pd.to_timedelta((fechas.dt.weekday - dia_inicio) % 7, unit='D')


df['tipo_producto'] = df.apply(clasificar_producto, axis=1)
df['semana_inicio'] = obtener_semana_operativa(df['fecha'], dia_inicio=DIA_INICIO_SEMANA)
df['semana_fin'] = df['semana_inicio'] + pd.Timedelta(days=DIAS_DURACION_SEMANA)
df['periodo_mes'] = df['fecha'].dt.to_period('M').astype(str)

#resumen operativo dias
resumen_dia = (
    df.groupby(['establecimiento', 'fecha', 'dia_semana'], as_index=False)
    .agg(
        productos=('producto', 'nunique'),
        inventario_inicial=('costo_inventario_inicial', 'sum'),
        compras=('costo_compras', 'sum'),
        costo_ventas=('costo_ventas', 'sum'),
        cortesias_merma=('costo_cortesias_merma', 'sum'),
        salidas_totales=('costo_salidas_total', 'sum'),
        inventario_final=('costo_inventario_final', 'sum'),
        diferencia_costo=('diferencia_costo', 'sum'),
        diferencia_unidades_abs=('diferencia_unidades', lambda x: x.abs().sum()),
    )
    .sort_values('fecha')
)

fecha_inicial_mes = resumen_dia['fecha'].min()
fecha_final_mes = resumen_dia['fecha'].max()
inventario_inicial_mes = resumen_dia.loc[resumen_dia['fecha'] == fecha_inicial_mes, 'inventario_inicial'].sum()
inventario_final_mes = resumen_dia.loc[resumen_dia['fecha'] == fecha_final_mes, 'inventario_final'].sum()
compras_mes = resumen_dia['compras'].sum()
costo_ventas_mes = resumen_dia['costo_ventas'].sum()
cortesias_merma_mes = resumen_dia['cortesias_merma'].sum()
salidas_totales_mes = resumen_dia['salidas_totales'].sum()
# Para riesgo se excluyen cambios normales de miscelanea, refrescos y mezcladores.
CATEGORIAS_CAMBIO_NORMAL = {'MISELANEA / INSUMO', 'REFRESCO / MEZCLADOR'}
df_diferencias_control = df[
    (~df['tipo_producto'].isin(CATEGORIAS_CAMBIO_NORMAL))
    & (df['diferencia_unidades'].abs() > UMBRAL_ALERTA_UNIDADES)
    & (df['diferencia_costo'].abs() > MARGEN_COSTO_RIESGO)
].copy()

diferencia_costo_mes = df_diferencias_control['diferencia_costo'].sum()
diferencia_abs_mes = df_diferencias_control['diferencia_costo'].abs().sum()
faltante_costo_mes = abs(
    df_diferencias_control.loc[df_diferencias_control['diferencia_costo'] < 0, 'diferencia_costo'].sum()
)
sobrante_costo_mes = df_diferencias_control.loc[
    df_diferencias_control['diferencia_costo'] > 0, 'diferencia_costo'
].sum()

resumen_mes_correcto = pd.DataFrame({
    'concepto': [
        'Fecha inicial', 'Fecha final', 'Inventario inicial mensual', 'Compras del mes',
        'Costo ventas del mes', 'Consumo y cortesias del mes', 'Salidas totales del mes',
        'Inventario final mensual', 'Diferencia costo acumulada', 'Diferencia absoluta revisada',
        'Faltante estimado', 'Sobrante estimado',
    ],
    'valor': [
        fecha_inicial_mes.strftime('%Y-%m-%d'), fecha_final_mes.strftime('%Y-%m-%d'),
        inventario_inicial_mes, compras_mes, costo_ventas_mes, cortesias_merma_mes,
        salidas_totales_mes, inventario_final_mes, diferencia_costo_mes, diferencia_abs_mes,
        faltante_costo_mes, sobrante_costo_mes,
    ]
})

# Semana operativa
semanas_por_fecha = df[['fecha', 'semana_inicio', 'semana_fin']].drop_duplicates().sort_values('fecha')
resumen_dia_semana = resumen_dia.merge(semanas_por_fecha, on='fecha', how='left')


def resumir_semana(grupo):
    grupo = grupo.sort_values('fecha')
    fecha_inicial = grupo['fecha'].min()
    fecha_final = grupo['fecha'].max()
    inventario_inicial = grupo.loc[grupo['fecha'] == fecha_inicial, 'inventario_inicial'].sum()
    inventario_final = grupo.loc[grupo['fecha'] == fecha_final, 'inventario_final'].sum()
    return pd.Series({
        'fecha_inicial': fecha_inicial,
        'fecha_final': fecha_final,
        'dias_operados': grupo['fecha'].nunique(),
        'inventario_inicial': inventario_inicial,
        'compras': grupo['compras'].sum(),
        'costo_ventas': grupo['costo_ventas'].sum(),
        'cortesias_merma': grupo['cortesias_merma'].sum(),
        'salidas_totales': grupo['salidas_totales'].sum(),
        'inventario_final': inventario_final,
        'diferencia_costo': grupo['diferencia_costo'].sum(),
        'diferencia_unidades_abs': grupo['diferencia_unidades_abs'].sum(),
    })

resumen_semana = (
    resumen_dia_semana.groupby(['establecimiento', 'semana_inicio', 'semana_fin'])
    .apply(resumir_semana)
    .reset_index()
    .sort_values('semana_inicio')
)

#alertas
TOLERANCIA_UNIDADES = UMBRAL_ALERTA_UNIDADES
TOLERANCIA_COSTO = UMBRAL_ALERTA_COSTO

alertas = df[
    (df['diferencia_unidades'].abs() > TOLERANCIA_UNIDADES)
    | (df['diferencia_costo'].abs() > TOLERANCIA_COSTO)
    | (df['precio_unitario'] <= 0)
].copy()

alertas['severidad'] = np.select(
    [
        alertas['diferencia_costo'].abs() >= UMBRAL_SEVERIDAD_ALTA,
        alertas['diferencia_costo'].abs() >= UMBRAL_SEVERIDAD_MEDIA,
    ],
    ['ALTA', 'MEDIA'],
    default='BAJA',
)
alertas['tipo_alerta'] = np.select(
    [
        alertas['precio_unitario'] <= 0,
        alertas['diferencia_unidades'] < 0,
        alertas['diferencia_unidades'] > 0,
    ],
    ['PRECIO UNITARIO EN CERO', 'FALTANTE', 'SOBRANTE'],
    default='REVISAR',
)
# Miscelanea, refrescos y mezcladores pueden compensarse por sustituciones normales.
es_cambio_normal = (
    alertas['tipo_producto'].isin(CATEGORIAS_CAMBIO_NORMAL)
    & alertas['tipo_alerta'].isin(['FALTANTE', 'SOBRANTE'])
)
es_variacion_menor = (
    alertas['tipo_alerta'].isin(['FALTANTE', 'SOBRANTE'])
    & (alertas['diferencia_costo'].abs() <= MARGEN_COSTO_RIESGO)
)
alertas = alertas[~es_cambio_normal & ~es_variacion_menor].copy()
alertas['severidad'] = pd.Categorical(
    alertas['severidad'], categories=['ALTA', 'MEDIA', 'BAJA'], ordered=True
)

alertas = alertas[[
    'severidad', 'tipo_alerta', 'establecimiento', 'fecha', 'dia_semana', 'producto',
    'tipo_producto', 'precio_unitario', 'inventario_inicial_unidades', 'compras_unidades',
    'ventas_unidades', 'cortesias_merma_unidades', 'inventario_final_unidades',
    'inventario_teorico_unidades', 'diferencia_unidades', 'diferencia_costo',
    'hoja_origen', 'fila_excel',
]].sort_values(by=['severidad', 'diferencia_costo'], ascending=[True, True])

if len(alertas) > 0:
    resumen_alertas = (
        alertas.groupby(['severidad', 'tipo_alerta'], as_index=False, observed=True)
        .agg(
            productos_afectados=('producto', 'nunique'),
            registros=('producto', 'count'),
            diferencia_costo_total=('diferencia_costo', 'sum'),
            diferencia_costo_abs=('diferencia_costo', lambda x: x.abs().sum()),
        )
        .sort_values(['severidad', 'diferencia_costo_abs'], ascending=[True, False])
    )
else:
    resumen_alertas = pd.DataFrame(columns=['severidad', 'tipo_alerta', 'productos_afectados', 'registros', 'diferencia_costo_total', 'diferencia_costo_abs'])

resumen_producto = (
    df.groupby(['establecimiento', 'producto', 'tipo_producto'], as_index=False)
    .agg(
        dias_contados=('fecha', 'nunique'),
        compras=('costo_compras', 'sum'),
        costo_ventas=('costo_ventas', 'sum'),
        cortesias_merma=('costo_cortesias_merma', 'sum'),
        salidas_totales=('costo_salidas_total', 'sum'),
        inventario_final=('costo_inventario_final', 'last'),
        inventario_final_unidades=('inventario_final_unidades', 'last'),
        ventas_unidades=('ventas_unidades', 'sum'),
        diferencia_unidades=('diferencia_unidades', 'sum'),
        diferencia_costo=('diferencia_costo', 'sum'),
        precio_unitario=('precio_unitario', 'last'),
    )
)

resumen_producto['diferencia_costo_abs'] = resumen_producto['diferencia_costo'].abs()
resumen_producto['rotacion_costo_vs_inventario'] = np.where(
    resumen_producto['inventario_final'] > 0,
    resumen_producto['costo_ventas'] / resumen_producto['inventario_final'],
    np.nan,
)
resumen_producto['cobertura_dias_estimada'] = np.where(
    resumen_producto['ventas_unidades'] > 0,
    resumen_producto['inventario_final_unidades'] / (resumen_producto['ventas_unidades'] / max(resumen_dia['fecha'].nunique(), 1)),
    np.nan,
)

top_ventas = resumen_producto.sort_values('costo_ventas', ascending=False).head(20)
top_cortesias = resumen_producto.sort_values('cortesias_merma', ascending=False).head(20)

resumen_producto_control = (
    df_diferencias_control.groupby(
        ['establecimiento', 'producto', 'tipo_producto'],
        as_index=False,
    )
    .agg(
        inventario_final=('costo_inventario_final', 'last'),
        diferencia_unidades=('diferencia_unidades', 'sum'),
        diferencia_costo=('diferencia_costo', 'sum'),
    )
)
resumen_producto_control['diferencia_costo_abs'] = resumen_producto_control['diferencia_costo'].abs()
resumen_producto_control = resumen_producto_control[
    resumen_producto_control['diferencia_costo_abs'] > MARGEN_COSTO_RIESGO
].copy()
top_diferencias = resumen_producto_control.sort_values('diferencia_costo_abs', ascending=False).head(20)
top_faltantes = resumen_producto_control[
    resumen_producto_control['diferencia_unidades'] < 0
].sort_values('diferencia_costo').head(15)
top_sobrantes = resumen_producto_control[
    resumen_producto_control['diferencia_unidades'] > 0
].sort_values('diferencia_costo', ascending=False).head(15)
productos_lentos = resumen_producto[
    (resumen_producto['inventario_final'] > 0) &
    (resumen_producto['costo_ventas'] <= resumen_producto['costo_ventas'].quantile(0.35))
].sort_values('inventario_final', ascending=False).head(15)

# Priorizacion ABC por valor en inventario y movimiento, seguida del calculo de cobertura.
def clasificacion_abc(tabla, columna):
    orden = tabla[columna].fillna(0).clip(lower=0).sort_values(ascending=False)
    clases = pd.Series('C', index=tabla.index, dtype='object')
    total = float(orden.sum())
    if total <= 0:
        return clases
    participacion = orden / total
    acumulado_anterior = participacion.cumsum() - participacion
    clases.loc[orden.index[acumulado_anterior < ABC_UMBRAL_A]] = 'A'
    clases.loc[orden.index[(acumulado_anterior >= ABC_UMBRAL_A) & (acumulado_anterior < ABC_UMBRAL_B)]] = 'B'
    return clases


demanda_diaria_producto = (
    df.groupby(['establecimiento', 'producto', 'tipo_producto', 'fecha'], as_index=False)
    .agg(
        demanda_diaria=('ventas_unidades', 'sum'),
        compra_diaria=('compras_unidades', 'sum'),
        presentacion_unidades=('presentacion_unidades', 'last'),
    )
)

historico_demanda = (
    demanda_diaria_producto.groupby(['establecimiento', 'producto', 'tipo_producto'], as_index=False)
    .agg(
        demanda_promedio_diaria=('demanda_diaria', 'mean'),
        desviacion_demanda_diaria=('demanda_diaria', lambda x: x.std(ddof=1)),
        lote_compra_habitual=('compra_diaria', lambda x: x[x > 0].median() if (x > 0).any() else np.nan),
        presentacion_unidades=('presentacion_unidades', 'last'),
    )
)
historico_demanda['desviacion_demanda_diaria'] = historico_demanda['desviacion_demanda_diaria'].fillna(0)
historico_demanda['lote_compra_habitual'] = (
    historico_demanda['lote_compra_habitual']
    .fillna(historico_demanda['presentacion_unidades'])
    .fillna(1)
    .clip(lower=1)
)

plan_compra = resumen_producto.merge(
    historico_demanda.drop(columns=['presentacion_unidades']),
    on=['establecimiento', 'producto', 'tipo_producto'],
    how='left',
)
plan_compra['valor_total_inventario'] = (
    plan_compra['inventario_final_unidades'].fillna(0).clip(lower=0)
    * plan_compra['precio_unitario'].fillna(0).clip(lower=0)
)
plan_compra['clasificacion_valor'] = clasificacion_abc(plan_compra, 'valor_total_inventario')
plan_compra['clasificacion_movimiento'] = clasificacion_abc(plan_compra, 'ventas_unidades')
plan_compra['clasificacion_abc'] = np.select(
    [
        (plan_compra['clasificacion_valor'] == 'A') | (plan_compra['clasificacion_movimiento'] == 'A'),
        (plan_compra['clasificacion_valor'] == 'B') | (plan_compra['clasificacion_movimiento'] == 'B'),
    ],
    ['A', 'B'],
    default='C',
)
plan_compra['unidades_vendidas_periodo'] = plan_compra['ventas_unidades'].fillna(0)
plan_compra['lead_time_dias'] = (
    plan_compra['producto'].map(LEAD_TIME_DIAS_POR_PRODUCTO)
    .fillna(LEAD_TIME_DIAS_DEFAULT)
    .astype(float)
    .clip(lower=0)
)
plan_compra['demanda_promedio_diaria'] = plan_compra['demanda_promedio_diaria'].fillna(0).clip(lower=0)
plan_compra['desviacion_demanda_diaria'] = plan_compra['desviacion_demanda_diaria'].fillna(0).clip(lower=0)
plan_compra['stock_seguridad'] = (
    NIVEL_SERVICIO_Z
    * plan_compra['desviacion_demanda_diaria']
    * np.sqrt(plan_compra['lead_time_dias'])
)
plan_compra['punto_reorden'] = (
    plan_compra['demanda_promedio_diaria'] * plan_compra['lead_time_dias']
    + plan_compra['stock_seguridad']
)
plan_compra['stock_maximo'] = plan_compra['punto_reorden'] + plan_compra['lote_compra_habitual']
plan_compra['stock_actual'] = plan_compra['inventario_final_unidades'].fillna(0).clip(lower=0)
plan_compra['comprar'] = np.where(plan_compra['stock_actual'] <= plan_compra['punto_reorden'], 'SI', 'NO')
plan_compra['cantidad_sugerida'] = np.where(
    plan_compra['comprar'] == 'SI',
    np.ceil(np.maximum(0, plan_compra['stock_maximo'] - plan_compra['stock_actual'])),
    0,
)

# El puntaje coloca primero los productos A en ambos criterios y despues pondera valor y rotacion.
puntos_abc = {'A': 3, 'B': 2, 'C': 1}
max_valor = max(float(plan_compra['valor_total_inventario'].max()), 1.0)
max_movimiento = max(float(plan_compra['ventas_unidades'].max()), 1.0)
plan_compra['puntaje_riesgo'] = (
    plan_compra['clasificacion_valor'].map(puntos_abc)
    + plan_compra['clasificacion_movimiento'].map(puntos_abc)
    + plan_compra['valor_total_inventario'] / max_valor
    + plan_compra['ventas_unidades'].fillna(0) / max_movimiento
)
compras_sugeridas = (
    plan_compra[
        (plan_compra['clasificacion_valor'] == 'A')
        | (plan_compra['clasificacion_movimiento'] == 'A')
    ]
    .sort_values(['puntaje_riesgo', 'valor_total_inventario', 'unidades_vendidas_periodo'], ascending=False)
    [[
        'producto', 'valor_total_inventario', 'unidades_vendidas_periodo', 'clasificacion_abc',
        'demanda_promedio_diaria', 'desviacion_demanda_diaria', 'stock_seguridad',
        'punto_reorden', 'stock_actual', 'comprar', 'cantidad_sugerida',
    ]]
    .reset_index(drop=True)
)

# KPIs
promedio_inventario = (inventario_inicial_mes + inventario_final_mes) / 2 if (inventario_inicial_mes + inventario_final_mes) != 0 else np.nan
variacion_inventario = inventario_final_mes - inventario_inicial_mes
variacion_inventario_pct = variacion_inventario / inventario_inicial_mes if inventario_inicial_mes else np.nan
cortesias_pct_salidas = cortesias_merma_mes / salidas_totales_mes if salidas_totales_mes else 0
compras_pct_salidas = compras_mes / salidas_totales_mes if salidas_totales_mes else 0
diferencia_abs_pct_salidas = diferencia_abs_mes / salidas_totales_mes if salidas_totales_mes else 0
rotacion_inventario = costo_ventas_mes / promedio_inventario if promedio_inventario else np.nan
alertas_altas = int((alertas['severidad'] == 'ALTA').sum()) if len(alertas) else 0
alertas_medias = int((alertas['severidad'] == 'MEDIA').sum()) if len(alertas) else 0

riesgo_score = min(100, round(
    (diferencia_abs_pct_salidas * 40) +
    (cortesias_pct_salidas * 35) +
    (alertas_altas * 3) +
    (alertas_medias * 1),
    1,
))
if riesgo_score >= 70:
    nivel_riesgo = 'ALTO'
elif riesgo_score >= 35:
    nivel_riesgo = 'MEDIO'
else:
    nivel_riesgo = 'CONTROLADO'

kpis_ejecutivos = pd.DataFrame({
    'KPI': [
        'Inventario inicial', 'Inventario final', 'Variacion inventario', 'Compras',
        'Costo ventas', 'Consumo y cortesias', 'Diferencia absoluta', 'Faltante estimado',
        'Sobrante estimado', 'Rotacion inventario', 'Alertas altas', 'Nivel de riesgo',
    ],
    'Valor': [
        inventario_inicial_mes, inventario_final_mes, variacion_inventario, compras_mes,
        costo_ventas_mes, cortesias_merma_mes, diferencia_abs_mes, faltante_costo_mes,
        sobrante_costo_mes, rotacion_inventario, alertas_altas, nivel_riesgo,
    ],
    'Lectura ejecutiva': [
        'Capital inicial en inventario al primer dia operativo del mes.',
        'Capital final en inventario al ultimo dia operativo del mes.',
        'Cambio neto del capital inmovilizado en inventario.',
        'Compra acumulada del periodo.',
        'Costo asociado al producto vendido.',
        'Consumo de personal y cortesias registradas en el periodo.',
        'Monto total sujeto a revision, sin compensar faltantes con sobrantes.',
        'Estimacion de producto faltante.',
        'Estimacion de producto sobrante.',
        'Veces que el costo de venta cubre el inventario promedio.',
        'Registros que requieren revision prioritaria.',
        'Lectura integral de riesgo por diferencias, cortesias y alertas.',
    ]
})


def construir_diagnostico():
    textos = []
    if variacion_inventario > 0:
        textos.append(
            f"El inventario final aumento {formato_pesos(variacion_inventario)} ({formato_pct(variacion_inventario_pct)}) frente al inventario inicial."
        )
    elif variacion_inventario < 0:
        textos.append(
            f"El inventario final disminuyo {formato_pesos(abs(variacion_inventario))} ({formato_pct(abs(variacion_inventario_pct))}) frente al inventario inicial."
        )
    else:
        textos.append('El inventario final se mantuvo sin variacion relevante frente al inventario inicial.')

    if cortesias_pct_salidas > 0.10:
        textos.append(
            f"Consumo y cortesias representan {formato_pct(cortesias_pct_salidas)} de las salidas y requieren conciliacion documental."
        )
    else:
        textos.append(
            f"Consumo y cortesias representan {formato_pct(cortesias_pct_salidas)} de las salidas."
        )

    productos_con_diferencia = len(resumen_producto_control)
    textos.append(
        f"Se identificaron {productos_con_diferencia} productos auditables con diferencias relevantes, excluyendo miscelanea, refrescos y mezcladores."
    )
    return textos


def construir_recomendaciones():
    productos_revision = pd.concat(
        [top_faltantes[['producto']], top_sobrantes[['producto']]],
        ignore_index=True,
    ).drop_duplicates()
    if len(productos_revision) > 0:
        nombres = ', '.join(productos_revision['producto'].head(3).astype(str).tolist())
        accion_conteo = (
            f"Recontar fisicamente y cotejar contra la hoja de auditoria los productos con diferencia prioritaria: {nombres}."
        )
    else:
        accion_conteo = (
            'Confirmar con un segundo conteo que los productos auditados no presenten diferencias en unidades.'
        )

    return [
        accion_conteo,
        'Conciliar consumo de personal y cortesias contra comandas, autorizaciones o bitacoras del periodo.',
        'Documentar cada diferencia confirmada con evidencia, responsable de seguimiento y correccion solicitada en el archivo fuente.',
    ]

diagnostico_ejecutivo = construir_diagnostico()
recomendaciones_ejecutivas = construir_recomendaciones()

# Exportacion con el analisis completo
TABLAS_SALIDA = {
    'inventario_diario_limpio.csv': df,
    'resumen_dia.csv': resumen_dia,
    'resumen_semana.csv': resumen_semana,
    'resumen_mes.csv': resumen_mes_correcto,
    'kpis_ejecutivos.csv': kpis_ejecutivos,
    'resumen_producto.csv': resumen_producto,
    'alertas.csv': alertas,
    'resumen_alertas.csv': resumen_alertas,
    'compras_sugeridas.csv': compras_sugeridas,
    'productos_lentos.csv': productos_lentos,
}
for nombre, tabla in TABLAS_SALIDA.items():
    tabla.to_csv(OUT_DIR / nombre, index=False, encoding='utf-8-sig')

ETIQUETAS_COLUMNAS = {
    'concepto': 'Concepto', 'valor': 'Valor', 'KPI': 'Indicador', 'Valor': 'Valor',
    'Lectura ejecutiva': 'Lectura ejecutiva', 'fecha': 'Fecha',
    'fecha_inicial': 'Inicio', 'fecha_final': 'Cierre', 'semana_inicio': 'Inicio de semana',
    'semana_fin': 'Fin de semana', 'dias_operados': 'Dias operados',
    'producto': 'Producto', 'tipo_producto': 'Categoria', 'severidad': 'Severidad',
    'tipo_alerta': 'Tipo de alerta', 'productos_afectados': 'Productos afectados',
    'registros': 'Registros', 'inventario_inicial': 'Inventario inicial',
    'inventario_final': 'Inventario final', 'compras': 'Compras',
    'costo_ventas': 'Costo de ventas', 'cortesias_merma': 'Consumo y cortesias',
    'cortesias_merma_unidades': 'Consumo y cortesias (unidades)',
    'costo_cortesias_merma': 'Consumo y cortesias (costo)',
    'salidas_totales': 'Salidas totales', 'diferencia_costo': 'Costo por unidad',
    'diferencia_costo_total': 'Costo por unidad neto', 'diferencia_costo_abs': 'Costo por unidad absoluto',
    'diferencia_unidades': 'Unidades',
    'diferencia_unidades_excel': 'Unidades (archivo)',
    'diferencia_costo_excel': 'Costo por unidad (archivo)',
    'inventario_final_unidades': 'Inventario final (unidades)',
    'ventas_unidades': 'Ventas (unidades)', 'compra_sugerida_unidades': 'Compra sugerida (unidades)',
    'compra_sugerida_costo': 'Compra sugerida (costo)',
    'valor_total_inventario': 'Valor en inventario', 'clasificacion_abc': 'Clasificación (A/B/C)',
    'unidades_vendidas_periodo': 'Unidades vendidas (periodo)',
    'clasificacion_valor': 'Clase por valor', 'clasificacion_movimiento': 'Clase por movimiento',
    'demanda_promedio_diaria': 'Demanda prom. diaria', 'desviacion_demanda_diaria': 'Desviacion diaria',
    'stock_seguridad': 'Stock de seguridad', 'punto_reorden': 'Punto de reorden',
    'stock_actual': 'Stock actual', 'comprar': '¿Comprar?', 'cantidad_sugerida': 'Cantidad sugerida',
    'lead_time_dias': 'Lead time (dias)', 'lote_compra_habitual': 'Lote habitual',
    'stock_maximo': 'Stock maximo', 'puntaje_riesgo': 'Puntaje de riesgo',
    'rotacion_costo_vs_inventario': 'Rotacion', 'cobertura_dias_estimada': 'Cobertura estimada (dias)',
    'precio_unitario': 'Precio unitario', 'dia_semana': 'Dia', 'establecimiento': 'Negocio',
    'dias_contados': 'Dias auditados', 'hoja_origen': 'Hoja origen', 'fila_excel': 'Fila de Excel',
}

def etiqueta_columna(nombre):
    return ETIQUETAS_COLUMNAS.get(nombre, str(nombre).replace('_', ' ').strip().title())

def clase_columna(nombre):
    n = str(nombre).lower()
    if any(x in n for x in ['fecha', 'semana_inicio', 'semana_fin']):
        return 'fecha'
    if n == 'dias_operados':
        return 'entero'
    if n in ['demanda_promedio_diaria', 'desviacion_demanda_diaria', 'stock_seguridad', 'punto_reorden', 'stock_actual', 'cantidad_sugerida', 'lote_compra_habitual', 'stock_maximo']:
        return 'numero'
    if any(x in n for x in ['unidades', 'dias_', 'dias ', 'registros', 'productos_afectados', 'fila_excel']):
        return 'numero'
    if any(x in n for x in ['porcentaje', '_pct', 'tasa']):
        return 'porcentaje'
    if any(x in n for x in ['inventario', 'compras', 'costo', 'cortesias', 'merma', 'salidas', 'faltante', 'sobrante', 'precio']):
        return 'moneda'
    if 'rotacion' in n or 'cobertura' in n:
        return 'decimal'
    return 'texto'

HOJAS = {
    'KPIs': kpis_ejecutivos,
    'Resumen mes': resumen_mes_correcto,
    'Resumen semana': resumen_semana,
    'Resumen dia': resumen_dia,
    'Productos': resumen_producto,
    'Resumen alertas': resumen_alertas,
    'Top ventas': top_ventas,
    'Top cortesias': top_cortesias,
    'Top diferencias': top_diferencias,
    'Top faltantes': top_faltantes,
    'Top sobrantes': top_sobrantes,
    'Productos lentos': productos_lentos,
    'Compra sugerida': compras_sugeridas,
    'Alertas detalle': alertas,
    'Base limpia': df,
}

with pd.ExcelWriter(RUTA_EXCEL_SALIDA, engine='xlsxwriter', datetime_format='dd/mm/yyyy') as writer:
    workbook = writer.book
    inicio = workbook.add_worksheet('Inicio')
    writer.sheets['Inicio'] = inicio
    inicio.hide_gridlines(2)
    inicio.set_tab_color(COLOR_SECUNDARIO)
    inicio.set_column('A:A', 3)
    inicio.set_column('B:B', 28)
    inicio.set_column('C:C', 22)
    inicio.set_column('D:D', 3)
    inicio.set_column('E:G', 22)

    fmt_titulo = workbook.add_format({
        'bold': True, 'font_size': 24, 'font_color': '#FFFFFF',
        'bg_color': COLOR_PRIMARIO, 'align': 'left', 'valign': 'vcenter',
    })
    fmt_subtitulo = workbook.add_format({'font_size': 12, 'font_color': '#475569'})
    fmt_seccion = workbook.add_format({
        'bold': True, 'font_size': 12, 'font_color': '#FFFFFF',
        'bg_color': COLOR_PRIMARIO, 'align': 'left',
    })
    fmt_kpi_nombre = workbook.add_format({
        'bold': True, 'font_color': '#FFFFFF', 'bg_color': COLOR_PRIMARIO,
        'align': 'center', 'border': 1, 'border_color': '#FFFFFF',
    })
    fmt_kpi_valor = workbook.add_format({
        'bold': True, 'font_size': 15, 'num_format': '$#,##0.00;[Red]-$#,##0.00',
        'bg_color': '#F1F5F9', 'align': 'center', 'border': 1, 'border_color': '#FFFFFF',
    })
    fmt_texto = workbook.add_format({'text_wrap': True, 'valign': 'top', 'font_color': '#334155'})
    fmt_riesgo = workbook.add_format({
        'bold': True, 'font_size': 15, 'font_color': '#FFFFFF',
        'bg_color': COLOR_EXITO if nivel_riesgo == 'CONTROLADO' else COLOR_ALERTA if nivel_riesgo == 'MEDIO' else COLOR_RIESGO,
        'align': 'center', 'valign': 'vcenter',
    })

    inicio.merge_range('B2:G3', 'REPORTE MENSUAL DE CONTROL', fmt_titulo)
    inicio.write('B4', TITULO_REPORTE, fmt_subtitulo)
    inicio.write('G4', f"Generado: {datetime.now():%d/%m/%Y}", fmt_subtitulo)
    inicio.merge_range('B6:G6', 'INDICADORES PRINCIPALES', fmt_seccion)

    kpis_inicio = [
        ('Inventario inicial', inventario_inicial_mes),
        ('Inventario final', inventario_final_mes),
        ('Consumo y cortesias', cortesias_merma_mes),
    ]
    for i, (nombre, valor) in enumerate(kpis_inicio):
        col = 1 + (i % 3) * 2
        row = 7 + (i // 3) * 3
        inicio.merge_range(row, col, row, col + 1, nombre, fmt_kpi_nombre)
        inicio.merge_range(row + 1, col, row + 2, col + 1, float(valor), fmt_kpi_valor)

    inicio.merge_range('B12:G12', 'NIVEL DE RIESGO', fmt_seccion)
    inicio.merge_range('B13:G14', nivel_riesgo, fmt_riesgo)
    inicio.merge_range('B16:G16', 'ACCIONES PRIORITARIAS', fmt_seccion)
    for fila, recomendacion in enumerate(recomendaciones_ejecutivas[:3], start=17):
        inicio.write(fila - 1, 1, f"{fila - 16}.", fmt_texto)
        inicio.merge_range(fila - 1, 2, fila - 1, 6, recomendacion, fmt_texto)
        inicio.set_row(fila - 1, 32)
    inicio.merge_range('B22:G22', 'Las hojas de detalle contienen filtros, encabezados legibles y formatos por tipo de dato.', fmt_subtitulo)
    inicio.freeze_panes(5, 1)

    formatos = {
        'moneda': workbook.add_format({'num_format': '$#,##0.00;[Red]-$#,##0.00', 'align': 'center', 'valign': 'vcenter'}),
        'numero': workbook.add_format({'num_format': '#,##0.00;[Red]-#,##0.00', 'align': 'center', 'valign': 'vcenter'}),
        'entero': workbook.add_format({'num_format': '#,##0', 'align': 'center', 'valign': 'vcenter'}),
        'porcentaje': workbook.add_format({'num_format': '0.0%', 'align': 'center', 'valign': 'vcenter'}),
        'decimal': workbook.add_format({'num_format': '0.00', 'align': 'center', 'valign': 'vcenter'}),
        'fecha': workbook.add_format({'num_format': 'dd/mm/yyyy', 'align': 'center', 'valign': 'vcenter'}),
        'texto': workbook.add_format({'text_wrap': False, 'valign': 'vcenter'}),
    }
    formato_header = workbook.add_format({
        'bold': True, 'bg_color': COLOR_PRIMARIO, 'font_color': '#FFFFFF',
        'border': 0, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True,
    })

    for indice, (nombre_hoja, tabla_original) in enumerate(HOJAS.items(), start=1):
        tabla = tabla_original.copy()
        columnas_originales = list(tabla.columns)
        tabla.columns = [etiqueta_columna(c) for c in columnas_originales]
        tabla.to_excel(writer, sheet_name=nombre_hoja, index=False)
        ws = writer.sheets[nombre_hoja]
        ws.hide_gridlines(2)
        ws.freeze_panes(1, 0)
        ws.set_zoom(90)
        ws.set_tab_color(COLOR_SECUNDARIO if indice <= 6 else COLOR_ACENTO)
        ws.set_row(0, 34, formato_header)
        ws.set_default_row(22)

        for col_idx, (original, visible) in enumerate(zip(columnas_originales, tabla.columns)):
            clase = clase_columna(original)
            ancho = 30 if original in ['producto', 'Lectura ejecutiva'] else 22 if clase == 'texto' else 16
            ws.set_column(col_idx, col_idx, ancho, formatos[clase])

        if len(tabla.columns) > 0:
            if len(tabla) > 0:
                nombre_tabla = f"Tabla_{indice:02d}"
                ws.add_table(0, 0, len(tabla), len(tabla.columns) - 1, {
                    'name': nombre_tabla,
                    'style': 'Table Style Medium 2',
                    'columns': [{'header': c} for c in tabla.columns],
                })
            else:
                ws.autofilter(0, 0, 0, len(tabla.columns) - 1)

        if 'Severidad' in tabla.columns and len(tabla) > 0:
            c = tabla.columns.get_loc('Severidad')
            ws.conditional_format(1, c, len(tabla), c, {
                'type': 'text', 'criteria': 'containing', 'value': 'ALTA',
                'format': workbook.add_format({'bg_color': '#FEE2E2', 'font_color': '#991B1B', 'bold': True}),
            })
            ws.conditional_format(1, c, len(tabla), c, {
                'type': 'text', 'criteria': 'containing', 'value': 'MEDIA',
                'format': workbook.add_format({'bg_color': '#FEF3C7', 'font_color': '#92400E', 'bold': True}),
            })

print('Excel V3 generado:', RUTA_EXCEL_SALIDA)

# Graficas
plt.rcParams.update({
    'font.family': 'DejaVu Sans',
    'axes.titlesize': 13,
    'axes.titleweight': 'bold',
    'axes.labelsize': 9,
    'xtick.labelsize': 8,
    'ytick.labelsize': 8,
    'figure.facecolor': 'white',
    'axes.facecolor': 'white',
})

def formato_pesos_eje_y(ax):
    ax.yaxis.set_major_formatter(mtick.StrMethodFormatter('${x:,.0f}'))

def formato_pesos_eje_x(ax):
    ax.xaxis.set_major_formatter(mtick.StrMethodFormatter('${x:,.0f}'))

def estilizar_ax(ax, eje='y'):
    ax.spines[['top', 'right']].set_visible(False)
    ax.spines[['left', 'bottom']].set_color('#CBD5E1')
    ax.grid(axis=eje, color='#E2E8F0', linewidth=0.8, alpha=0.8)
    ax.set_axisbelow(True)

def guardar_grafica(fig, nombre):
    ruta = GRAFICAS_DIR / nombre
    fig.tight_layout()
    fig.savefig(ruta, dpi=200, bbox_inches='tight', facecolor='white')
    plt.show()
    plt.close(fig)
    print('Grafica guardada:', ruta)
    return ruta

_g = resumen_semana.sort_values('semana_inicio').copy()
_g['semana'] = _g['fecha_inicial'].dt.strftime('%d-%b') + ' - ' + _g['fecha_final'].dt.strftime('%d-%b')

fig, ax = plt.subplots(figsize=(11, 5))
ax.plot(_g['semana'], _g['inventario_inicial'], marker='o', linewidth=2.4, color=COLOR_ACENTO, label='Inventario inicial')
ax.plot(_g['semana'], _g['inventario_final'], marker='o', linewidth=2.4, color=COLOR_SECUNDARIO, label='Inventario final')
ax.fill_between(range(len(_g)), _g['inventario_inicial'], _g['inventario_final'], color=COLOR_ACENTO, alpha=0.08)
ax.set_title('Evolucion del inventario por semana', loc='left')
ax.set_ylabel('Importe')
ax.legend(frameon=False, ncol=2)
formato_pesos_eje_y(ax)
estilizar_ax(ax)
ax.tick_params(axis='x', rotation=20)
grafica_inventario = guardar_grafica(fig, '01_inventario_inicial_vs_final.png')

fig, ax = plt.subplots(figsize=(11, 5))
ax.plot(_g['semana'], _g['compras'], marker='o', linewidth=2.2, color=COLOR_ACENTO, label='Compras')
ax.plot(_g['semana'], _g['salidas_totales'], marker='o', linewidth=2.2, color=COLOR_SECUNDARIO, label='Salidas totales')
ax.set_title('Compras y salidas por semana', loc='left')
ax.set_ylabel('Importe')
ax.legend(frameon=False, ncol=2)
formato_pesos_eje_y(ax)
estilizar_ax(ax)
ax.tick_params(axis='x', rotation=20)
grafica_movimiento = guardar_grafica(fig, '02_compras_ventas_salidas.png')

fig, ax = plt.subplots(figsize=(11, 5))
ax.bar(_g['semana'], _g['cortesias_merma'], color=COLOR_ALERTA, width=0.62)
ax.set_title('Consumo y cortesias por semana', loc='left')
ax.set_ylabel('Importe')
formato_pesos_eje_y(ax)
estilizar_ax(ax)
ax.tick_params(axis='x', rotation=20)
grafica_cortesias = guardar_grafica(fig, '03_cortesias_merma_semana.png')

_gv = top_ventas.sort_values('costo_ventas', ascending=True).tail(12)
fig, ax = plt.subplots(figsize=(10, 6.5))
ax.barh(_gv['producto'], _gv['costo_ventas'], color=COLOR_ACENTO)
ax.set_title('Productos con mayor costo de venta', loc='left')
ax.set_xlabel('Costo de venta')
formato_pesos_eje_x(ax)
estilizar_ax(ax, eje='x')
grafica_top_ventas = guardar_grafica(fig, '04_top_productos_costo_venta.png')

_gd = top_diferencias.sort_values('diferencia_costo_abs', ascending=True).tail(12)
colores_dif = [COLOR_RIESGO if v < 0 else COLOR_EXITO for v in _gd['diferencia_costo']]
fig, ax = plt.subplots(figsize=(10, 6.5))
ax.barh(_gd['producto'], _gd['diferencia_costo'], color=colores_dif)
ax.axvline(0, color='#64748B', linewidth=0.8)
ax.set_title('Diferencias relevantes por producto', loc='left')
ax.set_xlabel('Costo por unidad')
formato_pesos_eje_x(ax)
estilizar_ax(ax, eje='x')
grafica_top_diferencias = guardar_grafica(fig, '05_top_diferencias_costo.png')

faltantes_sobrantes_df = pd.DataFrame({
    'concepto': ['Faltante estimado', 'Sobrante estimado'],
    'importe': [faltante_costo_mes, sobrante_costo_mes],
})
fig, ax = plt.subplots(figsize=(7, 4.5))
ax.bar(faltantes_sobrantes_df['concepto'], faltantes_sobrantes_df['importe'], color=[COLOR_RIESGO, COLOR_EXITO], width=0.58)
ax.set_title('Faltantes y sobrantes estimados', loc='left')
ax.set_ylabel('Importe')
formato_pesos_eje_y(ax)
estilizar_ax(ax)
grafica_falt_sobr = guardar_grafica(fig, '06_faltantes_vs_sobrantes.png')

graficas_pdf = [
    grafica_inventario, grafica_movimiento, grafica_cortesias,
    grafica_top_ventas, grafica_top_diferencias, grafica_falt_sobr,
]

# Componentes reutilizables para el PDF
ETIQUETAS_PDF = {
    'concepto': 'Concepto', 'valor': 'Valor', 'fecha_inicial': 'Inicio', 'fecha_final': 'Cierre',
    'dias_operados': 'Dias', 'inventario_inicial': 'Inventario inicial', 'inventario_final': 'Inventario final',
    'compras': 'Compras', 'costo_ventas': 'Costo de ventas', 'cortesias_merma': 'Consumo y cortesias',
    'diferencia_costo': 'Costo', 'diferencia_unidades': 'Unidades',
    'tipo_producto': 'Categoria', 'producto': 'Producto', 'severidad': 'Severidad',
    'tipo_alerta': 'Alerta', 'productos_afectados': 'Productos', 'registros': 'Registros',
    'diferencia_costo_total': 'Costo por unidad neto', 'diferencia_costo_abs': 'Costo por unidad absoluto',
    'inventario_final_unidades': 'Inv. final (unid.)', 'ventas_unidades': 'Ventas (unid.)',
    'compra_sugerida_unidades': 'Compra (unid.)', 'compra_sugerida_costo': 'Compra (costo)',
    'rotacion_costo_vs_inventario': 'Rotacion', 'costo_referencia': 'Valor',
    'valor_total_inventario': 'Valor en inventario', 'clasificacion_abc': 'Clasificación (A/B/C)',
    'unidades_vendidas_periodo': 'Unidades vendidas (periodo)',
    'demanda_promedio_diaria': 'Demanda prom. diaria', 'desviacion_demanda_diaria': 'sigma d',
    'stock_seguridad': 'Stock seguridad', 'punto_reorden': 'Punto reorden',
    'stock_actual': 'Stock actual', 'comprar': '¿Comprar?', 'cantidad_sugerida': 'Cantidad sugerida',
}

def paragraph(text, style):
    return Paragraph(str(text).replace('\n', '<br/>'), style)

def etiqueta_pdf(nombre):
    return ETIQUETAS_PDF.get(nombre, str(nombre).replace('_', ' ').title())

def tipo_dato_columna(nombre):
    n = str(nombre).lower()
    if 'fecha' in n or 'semana' in n:
        return 'fecha'
    if n == 'dias_operados':
        return 'entero'
    if n in ['demanda_promedio_diaria', 'desviacion_demanda_diaria', 'stock_seguridad', 'punto_reorden', 'stock_actual', 'cantidad_sugerida']:
        return 'numero'
    if any(x in n for x in ['unidades', 'dias_operados', 'registros', 'productos_afectados']):
        return 'numero'
    if 'rotacion' in n or 'cobertura' in n:
        return 'decimal'
    if any(x in n for x in ['inventario', 'compras', 'ventas', 'cortesias', 'merma', 'salidas', 'diferencia', 'costo', 'faltante', 'sobrante', 'precio']):
        return 'moneda'
    return 'texto'

def formatear_celda(valor, tipo):
    if pd.isna(valor):
        return '-'
    if tipo == 'fecha':
        return formato_fecha(valor)
    if tipo == 'moneda':
        return formato_pesos(valor)
    if tipo == 'entero':
        return f"{int(round(float(valor)))}"
    if tipo == 'numero':
        return formato_num(valor)
    if tipo == 'decimal':
        return formato_num(valor)
    return str(valor)

def tabla_pdf(df_tabla, columnas=None, max_filas=None, col_widths=None, font_size=7):
    tabla = df_tabla.copy()
    if columnas is not None:
        tabla = tabla[columnas].copy()
    if max_filas is not None:
        tabla = tabla.head(max_filas).copy()

    originales = list(tabla.columns)
    encabezados = [Paragraph(etiqueta_pdf(c), ParagraphStyle(
        f'h_{i}', fontName='Helvetica-Bold', fontSize=font_size, leading=font_size + 1,
        textColor=colors.white, alignment=TA_CENTER,
    )) for i, c in enumerate(originales)]

    filas = []
    for _, row in tabla.iterrows():
        fila = []
        for col in originales:
            tipo = tipo_dato_columna(col)
            texto = formatear_celda(row[col], tipo)
            estilo = ParagraphStyle(
                f'c_{col}', fontName='Helvetica', fontSize=font_size, leading=font_size + 1.5,
                textColor=colors.HexColor('#1E293B'),
                alignment=(
                    TA_RIGHT if tipo in ['moneda', 'numero', 'decimal']
                    else TA_CENTER if col in ['fecha_inicial', 'fecha_final', 'dias_operados', 'tipo_producto', 'clasificacion_abc', 'comprar', 'severidad']
                    else TA_LEFT
                ),
            )
            fila.append(Paragraph(texto, estilo))
        filas.append(fila)

    data = [encabezados] + filas
    if col_widths is None:
        col_widths = [7.35 * inch / max(len(originales), 1)] * len(originales)
    t = Table(data, repeatRows=1, colWidths=col_widths, hAlign='CENTER')
    estilos = [
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor(COLOR_PRIMARIO)),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('GRID', (0, 0), (-1, -1), 0.25, colors.HexColor('#CBD5E1')),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#F8FAFC')]),
        ('TOPPADDING', (0, 0), (-1, -1), 5.5),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 5.5),
        ('LEFTPADDING', (0, 0), (-1, -1), 4),
        ('RIGHTPADDING', (0, 0), (-1, -1), 4),
    ]
    if 'severidad' in originales:
        col_sev = originales.index('severidad')
        for fila_idx, valor in enumerate(tabla['severidad'].astype(str), start=1):
            if valor == 'ALTA':
                estilos.append(('BACKGROUND', (col_sev, fila_idx), (col_sev, fila_idx), colors.HexColor('#FEE2E2')))
                estilos.append(('TEXTCOLOR', (col_sev, fila_idx), (col_sev, fila_idx), colors.HexColor('#991B1B')))
            elif valor == 'MEDIA':
                estilos.append(('BACKGROUND', (col_sev, fila_idx), (col_sev, fila_idx), colors.HexColor('#FEF3C7')))
    if 'comprar' in originales:
        col_comprar = originales.index('comprar')
        for fila_idx, valor in enumerate(tabla['comprar'].astype(str), start=1):
            if valor == 'SI':
                estilos.append(('BACKGROUND', (col_comprar, fila_idx), (col_comprar, fila_idx), colors.HexColor('#DCFCE7')))
                estilos.append(('TEXTCOLOR', (col_comprar, fila_idx), (col_comprar, fila_idx), colors.HexColor('#166534')))
                estilos.append(('FONTNAME', (col_comprar, fila_idx), (col_comprar, fila_idx), 'Helvetica-Bold'))
    t.setStyle(TableStyle(estilos))
    t.spaceBefore = 7
    t.spaceAfter = 14
    return t

def imagen_ajustada(ruta, max_width=7.1 * inch, max_height=3.2 * inch):
    ancho, alto = ImageReader(str(ruta)).getSize()
    escala = min(max_width / ancho, max_height / alto)
    imagen = Image(str(ruta), width=ancho * escala, height=alto * escala)
    imagen.hAlign = 'CENTER'
    return imagen

def kpi_cards_pdf():
    cards = [
        ('Inventario inicial', formato_pesos(inventario_inicial_mes)),
        ('Inventario final', formato_pesos(inventario_final_mes)),
        ('Consumo y cortesias', formato_pesos(cortesias_merma_mes)),
    ]
    data = [
        [card[0] for card in cards],
        [card[1] for card in cards],
    ]
    t = Table(data, colWidths=[2.42 * inch] * 3, rowHeights=[0.34 * inch, 0.54 * inch], hAlign='CENTER')
    t.setStyle(TableStyle([
        ('GRID', (0, 0), (-1, -1), 1, colors.white),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor(COLOR_PRIMARIO)),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 7.5),
        ('BACKGROUND', (0, 1), (-1, 1), colors.HexColor('#F1F5F9')),
        ('FONTNAME', (0, 1), (-1, 1), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 1), (-1, 1), 13),
    ]))
    return t

def riesgo_badge_pdf():
    color = COLOR_EXITO if nivel_riesgo == 'CONTROLADO' else COLOR_ALERTA if nivel_riesgo == 'MEDIO' else COLOR_RIESGO
    t = Table([['NIVEL DE RIESGO', nivel_riesgo]], colWidths=[2.2 * inch, 2.2 * inch], rowHeights=[0.42 * inch], hAlign='CENTER')
    t.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (0, 0), colors.HexColor(COLOR_PRIMARIO)),
        ('BACKGROUND', (1, 0), (1, 0), colors.HexColor(color)),
        ('TEXTCOLOR', (0, 0), (-1, -1), colors.white),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ]))
    return t

def encabezado_pie(canvas, doc):
    canvas.saveState()
    canvas.setFillColor(colors.white)
    canvas.rect(0, 0, letter[0], letter[1], stroke=0, fill=1)
    ancho, alto = letter
    canvas.setStrokeColor(colors.HexColor(COLOR_SECUNDARIO))
    canvas.setLineWidth(1.2)
    canvas.line(doc.leftMargin, alto - 30, ancho - doc.rightMargin, alto - 30)
    canvas.setFont('Helvetica-Bold', 8)
    canvas.setFillColor(colors.HexColor(COLOR_PRIMARIO))
    canvas.drawString(doc.leftMargin, alto - 23, ESTABLECIMIENTO)
    canvas.setFont('Helvetica', 7.5)
    canvas.setFillColor(colors.HexColor('#64748B'))
    canvas.drawRightString(ancho - doc.rightMargin, alto - 23, TITULO_REPORTE)
    canvas.drawString(doc.leftMargin, 20, 'Documento de control interno')
    canvas.drawRightString(ancho - doc.rightMargin, 20, f'Pagina {doc.page}')
    canvas.restoreState()

# PDF
doc = SimpleDocTemplate(
    str(RUTA_PDF),
    pagesize=letter,
    rightMargin=32,
    leftMargin=32,
    topMargin=46,
    bottomMargin=34,
    title=f"Reporte mensual {TITULO_REPORTE}",
    author=ESTABLECIMIENTO,
)
styles = getSampleStyleSheet()
titulo_style = ParagraphStyle(
    'TituloV3', parent=styles['Title'], fontName='Helvetica-Bold',
    fontSize=22, leading=25, alignment=TA_LEFT, textColor=colors.HexColor(COLOR_PRIMARIO),
    spaceAfter=6,
)
subtitulo_style = ParagraphStyle(
    'SubtituloV3', parent=styles['Normal'], fontSize=11, leading=14,
    textColor=colors.HexColor('#64748B'), alignment=TA_LEFT, spaceAfter=14,
)
seccion_style = ParagraphStyle(
    'SeccionV3', parent=styles['Heading2'], fontName='Helvetica-Bold',
    fontSize=14, leading=17, alignment=TA_LEFT, textColor=colors.HexColor(COLOR_PRIMARIO),
    spaceBefore=4, spaceAfter=8,
)
normal_style = ParagraphStyle(
    'NormalV3', parent=styles['Normal'], fontSize=8.5, leading=11.5,
    textColor=colors.HexColor('#1E293B'), alignment=TA_LEFT,
)
bullet_style = ParagraphStyle(
    'BulletV3', parent=normal_style, leftIndent=14, firstLineIndent=-8,
    bulletIndent=4, spaceAfter=5,
)
nota_style = ParagraphStyle(
    'NotaV3', parent=styles['Normal'], fontSize=7.2, leading=9.2,
    textColor=colors.HexColor('#64748B'), alignment=TA_CENTER,
)
mini_style = ParagraphStyle(
    'MiniV3', parent=styles['Normal'], fontSize=7.5, leading=9.5,
    textColor=colors.HexColor('#475569'), alignment=TA_CENTER,
)

resumen_referencia_pdf = pd.DataFrame({
    'concepto': ['Compras del periodo', 'Costo de ventas', 'Diferencia absoluta revisada'],
    'costo_referencia': [compras_mes, costo_ventas_mes, diferencia_abs_mes],
})

story = []

# 1. Portada
if LOGO_PATH.exists():
    story.append(imagen_ajustada(LOGO_PATH, max_width=1.6 * inch, max_height=0.8 * inch))
    story.append(Spacer(1, 0.10 * inch))
story.append(paragraph('REPORTE MENSUAL DE CONTROL', titulo_style))
story.append(paragraph(TITULO_REPORTE, subtitulo_style))
story.append(riesgo_badge_pdf())
story.append(Spacer(1, 0.18 * inch))
story.append(kpi_cards_pdf())
story.append(Spacer(1, 0.18 * inch))
story.append(paragraph('Lectura ejecutiva', seccion_style))
for texto in diagnostico_ejecutivo:
    story.append(Paragraph('- ' + texto, bullet_style))
story.append(Spacer(1, 0.08 * inch))
story.append(paragraph('Acciones prioritarias', seccion_style))
for numero, recomendacion in enumerate(recomendaciones_ejecutivas[:3], start=1):
    story.append(Paragraph(f"{numero}. {recomendacion}", bullet_style))
story.append(Spacer(1, 0.06 * inch))
story.append(paragraph('Movimientos de referencia', seccion_style))
story.append(Paragraph(
    'Resume las compras, el costo de ventas y las diferencias del periodo como referencia general.',
    mini_style,
))
story.append(tabla_pdf(
    resumen_referencia_pdf,
    columnas=['concepto', 'costo_referencia'],
    col_widths=[3.6 * inch, 2.2 * inch],
    font_size=7.0,
))
story.append(PageBreak())

# 2 Tendencia operativa
story.append(paragraph('1. Tendencia operativa', seccion_style))
story.append(paragraph(
    'Las graficas muestran como cambiaron el inventario, las compras y las salidas durante el mes.',
    normal_style,
))
story.append(Spacer(1, 0.08 * inch))
if Path(grafica_inventario).exists():
    story.append(imagen_ajustada(grafica_inventario, max_height=2.55 * inch))
story.append(Spacer(1, 0.08 * inch))
if Path(grafica_movimiento).exists():
    story.append(imagen_ajustada(grafica_movimiento, max_height=2.55 * inch))
story.append(Spacer(1, 0.10 * inch))
story.append(paragraph('Resumen semanal', seccion_style))
story.append(Paragraph(
    'Resume por semana las fechas operadas, el inventario inicial y final, y el consumo y cortesias.',
    mini_style,
))
story.append(tabla_pdf(
    resumen_semana,
    columnas=['fecha_inicial', 'fecha_final', 'dias_operados', 'inventario_inicial', 'inventario_final', 'cortesias_merma'],
    max_filas=6,
    col_widths=[0.90 * inch, 0.90 * inch, 0.52 * inch, 1.40 * inch, 1.40 * inch, 1.55 * inch],
    font_size=6.5,
))
story.append(PageBreak())

# 3 Auditoria y riesgos
story.append(paragraph('2. Auditoria y riesgos', seccion_style))
if Path(grafica_top_diferencias).exists():
    story.append(imagen_ajustada(grafica_top_diferencias, max_height=2.75 * inch))
story.append(Spacer(1, 0.28 * inch))
story.append(paragraph('Productos que requieren revision', seccion_style))
story.append(Paragraph(
    'Muestra los productos que conviene revisar primero. Las unidades indican cuanto falta o sobra y el costo estima su efecto en dinero.',
    mini_style,
))
story.append(tabla_pdf(
    top_diferencias,
    columnas=['producto', 'tipo_producto', 'diferencia_unidades', 'diferencia_costo', 'inventario_final'],
    max_filas=8,
    col_widths=[2.45 * inch, 1.30 * inch, 1.12 * inch, 1.15 * inch, 1.15 * inch],
    font_size=6.3,
))
story.append(PageBreak())

# 4 Inventario, consumo y cortesias
story.append(paragraph('3. Inventario, consumo y cortesias', seccion_style))
graficas_decision = []
if Path(grafica_cortesias).exists():
    graficas_decision.append(imagen_ajustada(grafica_cortesias, max_width=3.55 * inch, max_height=2.45 * inch))
else:
    graficas_decision.append(Paragraph('Sin grafica disponible', mini_style))
if Path(grafica_falt_sobr).exists():
    graficas_decision.append(imagen_ajustada(grafica_falt_sobr, max_width=3.55 * inch, max_height=2.45 * inch))
else:
    graficas_decision.append(Paragraph('Sin grafica disponible', mini_style))
panel_graficas = Table([graficas_decision], colWidths=[3.65 * inch, 3.65 * inch], hAlign='CENTER')
panel_graficas.setStyle(TableStyle([('VALIGN', (0, 0), (-1, -1), 'TOP'), ('LEFTPADDING', (0, 0), (-1, -1), 0), ('RIGHTPADDING', (0, 0), (-1, -1), 4)]))
story.append(panel_graficas)
story.append(Spacer(1, 0.10 * inch))
story.append(paragraph('Consumo y cortesias con mayor impacto', seccion_style))
story.append(Paragraph(
    'Ordena los productos por el importe registrado como consumo de personal y cortesias durante el mes.',
    mini_style,
))
story.append(tabla_pdf(
    top_cortesias,
    columnas=['producto', 'tipo_producto', 'cortesias_merma'],
    max_filas=7,
    col_widths=[3.60 * inch, 1.65 * inch, 1.70 * inch],
    font_size=6.5,
))
story.append(Spacer(1, 0.10 * inch))
story.append(Spacer(1, 0.12 * inch))
story.append(paragraph('Capital inmovilizado: productos con baja rotacion', seccion_style))
story.append(Paragraph(
    'Muestra productos con inventario disponible y poco movimiento. Ayuda a ubicar dinero detenido en existencias.',
    mini_style,
))
story.append(tabla_pdf(
    productos_lentos,
    columnas=['producto', 'tipo_producto', 'inventario_final', 'costo_ventas', 'rotacion_costo_vs_inventario'],
    max_filas=7,
    col_widths=[2.45 * inch, 1.35 * inch, 1.20 * inch, 1.20 * inch, 1.20 * inch],
    font_size=6.3,
))
story.append(PageBreak())

# 5 Compra sugerida por riesgo y cobertura
story.append(paragraph('4. Compra sugerida por riesgo y cobertura', seccion_style))
story.append(Paragraph(
    'Reune los productos mas importantes por su valor o movimiento. Compara el stock disponible con el consumo observado '
    'y muestra si conviene comprar, junto con una cantidad aproximada.',
    mini_style,
))
story.append(Spacer(1, 0.20 * inch))
story.append(tabla_pdf(
    compras_sugeridas,
    columnas=[
        'producto', 'valor_total_inventario', 'unidades_vendidas_periodo', 'clasificacion_abc',
        'demanda_promedio_diaria', 'desviacion_demanda_diaria', 'stock_seguridad',
        'punto_reorden', 'stock_actual', 'comprar', 'cantidad_sugerida',
    ],
    max_filas=12,
    col_widths=[
        1.30 * inch, 0.73 * inch, 0.66 * inch, 0.48 * inch, 0.63 * inch,
        0.46 * inch, 0.64 * inch, 0.64 * inch, 0.57 * inch, 0.50 * inch, 0.64 * inch,
    ],
    font_size=4.8,
))
story.append(Spacer(1, 0.42 * inch))
story.append(paragraph('Criterio metodologico', seccion_style))
story.append(Paragraph(
    'Esta tabla ayuda a decidir que productos revisar primero y cuando conviene comprar. '
    'El punto de reorden estima el nivel en el que seria necesario hacer un pedido. '
    'La cantidad sugerida es una referencia de cuanto comprar para recuperar una cobertura razonable.',
    nota_style,
))
story.append(Spacer(1, 0.15 * inch))
story.append(Paragraph(
    f"Fuente: {ARCHIVO_EXCEL.name} - Elaborado: {datetime.now():%d/%m/%Y %H:%M}",
    nota_style,
))

doc.build(story, onFirstPage=encabezado_pie, onLaterPages=encabezado_pie)
print('PDF ejecutivo V3 generado:', RUTA_PDF)

# Validacion final y descarga
print('\nVALIDACION FINAL')
print('-' * 60)
artefactos = {
    'Excel analitico V3': RUTA_EXCEL_SALIDA,
    'PDF ejecutivo V3': RUTA_PDF,
}
for nombre, ruta in artefactos.items():
    existe = ruta.exists()
    tamano = ruta.stat().st_size if existe else 0
    print(f"{nombre}: {'OK' if existe else 'FALTA'} - {tamano / 1024:,.1f} KB - {ruta}")

print('\nIndicadores de control')
print('Filas procesadas:', len(df))
print('Periodo:', formato_fecha(fecha_inicial_mes), 'a', formato_fecha(fecha_final_mes))
print('Alertas altas:', alertas_altas)
print('Nivel de riesgo:', nivel_riesgo)

if DESCARGA_AUTOMATICA and EN_COLAB:
    from google.colab import files
    for ruta in [RUTA_PDF, RUTA_EXCEL_SALIDA]:
        if ruta.exists():
            files.download(str(ruta))