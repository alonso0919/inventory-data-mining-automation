# -*- coding: utf-8 -*-
"""
Genera un Excel de demostracion con la MISMA estructura que usa el pipeline
real (una hoja por dia operado, 30 columnas fijas, nombre de hoja
"ESTABLECIMIENTO MES DIA_SEMANA DIA") pero con productos, precios e
inventarios totalmente ficticios.

Sirve para poder clonar este repo y correr `src/reporte_mensual.py` de
principio a fin sin depender de ningun dato real de un negocio.
"""
import random
from calendar import monthrange
from datetime import date
from pathlib import Path

import numpy as np
import pandas as pd

random.seed(7)
np.random.seed(7)

ESTABLECIMIENTO = "DEMO"  # nombre de ejemplo; Excel limita el nombre de hoja a 31 caracteres
ANIO = 2026
MES = 7
DIAS_OPERACION = {3, 4, 5}  # 0=lunes ... 3=jueves, 4=viernes, 5=sabado
ARCHIVO_SALIDA = Path(__file__).parent / "inventario_demo.xlsx"

MESES_NOMBRE = {
    1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL", 5: "MAYO", 6: "JUNIO",
    7: "JULIO", 8: "AGOSTO", 9: "SEPTIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE",
}
DIAS_SEMANA_NOMBRE = {
    0: "LUNES", 1: "MARTES", 2: "MIERCOLES", 3: "JUEVES",
    4: "VIERNES", 5: "SABADO", 6: "DOMINGO",
}

# catalogo ficticio: (nombre, tipo, precio_compra, presentacion_unidades)
# tipo define como se simula la venta: 'BOTELLA' se vende por trago, el resto por pieza/unidad.
CATALOGO = [
    ("RON EJEMPLO AÑEJO", "BOTELLA", 420.0, 25),
    ("VODKA PRUEBA PREMIUM", "BOTELLA", 380.0, 25),
    ("WHISKY DEMO 8 ANOS", "BOTELLA", 650.0, 25),
    ("TEQUILA MUESTRA BLANCO", "BOTELLA", 410.0, 25),
    ("GINEBRA TEST LONDON DRY", "BOTELLA", 470.0, 25),
    ("CORONA EXTRA", "CERVEZA", 21.0, 1),
    ("HEINEKEN LAGER", "CERVEZA", 24.0, 1),
    ("MODELO ESPECIAL", "CERVEZA", 22.0, 1),
    ("COCA COLA", "REFRESCO / MEZCLADOR", 14.0, 1),
    ("AGUA MINERAL TONICA", "REFRESCO / MEZCLADOR", 16.0, 1),
    ("RED BULL ENERGETICA", "REFRESCO / MEZCLADOR", 32.0, 1),
    ("LIMON PARA COCTELERIA", "MISELANEA / INSUMO", 3.5, 1),
    ("SAL DE GRANO", "MISELANEA / INSUMO", 1.2, 1),
    ("SERVILLETAS BARRA", "MISELANEA / INSUMO", 0.8, 1),
]

COLUMNAS = [
    "producto", "precio_compra", "presentacion_unidades", "precio_unitario_excel",
    "inventario_inicial_botellas", "inventario_inicial_botellas_uni", "inventario_inicial_suelto",
    "inventario_inicial_total_excel", "compras_caja_bot", "compras_unidades_excel",
    "disponible_total_excel", "ventas_botellas", "ventas_botellas_unidades_excel",
    "ventas_tragos", "cortesias_merma", "salidas_total_excel",
    "inventario_final_botellas", "inventario_final_botellas_uni", "tara_peso_actual_oz",
    "peso_botella_vacia_oz", "inventario_final_suelto", "inventario_final_total_excel",
    "inventario_teorico_excel", "diferencia_unidades_excel", "diferencia_costo_excel",
    "costo_almacen_final_excel", "costo_salidas_total_excel", "costo_inventario_inicial_excel",
    "costo_inicial_mas_compras_excel", "costo_compras_excel",
]

ETIQUETAS = {
    "producto": "PRODUCTO", "precio_compra": "PRECIO COMPRA",
    "presentacion_unidades": "PRESENTACION (UNIDADES)", "precio_unitario_excel": "PRECIO UNITARIO",
    "inventario_inicial_botellas": "INV. INICIAL (BOTELLAS)", "inventario_inicial_botellas_uni": "INV. INICIAL (UNIDADES)",
    "inventario_inicial_suelto": "INV. INICIAL (SUELTO)", "inventario_inicial_total_excel": "INV. INICIAL TOTAL",
    "compras_caja_bot": "COMPRAS (CAJA/BOT)", "compras_unidades_excel": "COMPRAS (UNIDADES)",
    "disponible_total_excel": "DISPONIBLE TOTAL", "ventas_botellas": "VENTAS (BOTELLAS/PIEZAS)",
    "ventas_botellas_unidades_excel": "VENTAS BOTELLAS (UNIDADES)", "ventas_tragos": "VENTAS (TRAGOS)",
    "cortesias_merma": "CORTESIAS / MERMA", "salidas_total_excel": "SALIDAS TOTAL",
    "inventario_final_botellas": "INV. FINAL (BOTELLAS)", "inventario_final_botellas_uni": "INV. FINAL (UNIDADES)",
    "tara_peso_actual_oz": "TARA PESO ACTUAL (OZ)", "peso_botella_vacia_oz": "PESO BOTELLA VACIA (OZ)",
    "inventario_final_suelto": "INV. FINAL (SUELTO)", "inventario_final_total_excel": "INV. FINAL TOTAL",
    "inventario_teorico_excel": "INV. TEORICO", "diferencia_unidades_excel": "DIFERENCIA (UNIDADES)",
    "diferencia_costo_excel": "DIFERENCIA (COSTO)", "costo_almacen_final_excel": "COSTO ALMACEN FINAL",
    "costo_salidas_total_excel": "COSTO SALIDAS TOTAL", "costo_inventario_inicial_excel": "COSTO INV. INICIAL",
    "costo_inicial_mas_compras_excel": "COSTO INICIAL + COMPRAS", "costo_compras_excel": "COSTO COMPRAS",
}


def dias_operados(anio, mes, dias_operacion):
    total_dias = monthrange(anio, mes)[1]
    fechas = [date(anio, mes, d) for d in range(1, total_dias + 1)]
    return [f for f in fechas if f.weekday() in dias_operacion]


def simular_inventario(fechas):
    """Genera, producto por producto, la cadena inicial->compra->venta->final
    encadenando el final de un dia operado como inicial del siguiente
    (igual que en el flujo real: el jefe de barra reporta su final del dia
    y ese es el inicial del proximo dia de operacion)."""
    filas_por_fecha = {f: [] for f in fechas}
    stock_actual = {}
    nivel_par = {}

    for producto, tipo, precio_compra, presentacion in CATALOGO:
        nivel_par[producto] = random.uniform(25, 60) * presentacion
        stock_actual[producto] = nivel_par[producto]

    ultimas_fechas_sin_reposicion = set(fechas[-2:])

    for fecha in fechas:
        for producto, tipo, precio_compra, presentacion in CATALOGO:
            precio_unitario = precio_compra / presentacion
            inicial = stock_actual[producto]
            par = nivel_par[producto]

            # se repone cuando el stock cae por debajo de la mitad del nivel
            # objetivo (punto de reorden simple), salvo en los ultimos dias
            # operados del mes: igual que en la operacion real, la compra
            # para reponer eso se decide y se recibe hasta el periodo
            # siguiente, asi que el cierre de mes puede quedar con stock bajo.
            compra_bool = inicial < par * 0.55 and fecha not in ultimas_fechas_sin_reposicion
            if compra_bool:
                compras_unidades = max(0, (par - inicial)) * random.uniform(0.85, 1.15)
                compras_caja_bot = max(1, round(compras_unidades / presentacion))
                compras_unidades = compras_caja_bot * presentacion
            else:
                compras_caja_bot = 0
                compras_unidades = 0

            disponible = inicial + compras_unidades

            demanda_base = disponible * random.uniform(0.12, 0.30)
            ventas_unidades_totales = max(0, np.random.normal(demanda_base, demanda_base * 0.2))
            cortesias = ventas_unidades_totales * random.uniform(0.01, 0.06)
            salidas = ventas_unidades_totales + cortesias
            salidas = min(salidas, disponible)

            # ventas_botellas queda en 0 para todo lo que no es BOTELLA: el
            # clasificador de producto del pipeline real (clasificar_producto)
            # prioriza cualquier dato de botella/tara sobre las palabras clave,
            # asi que solo los espirituosos deben tocar esas columnas.
            ventas_botellas = 0
            ventas_botellas_unidades = 0
            ventas_tragos = ventas_unidades_totales

            final_teorico = disponible - salidas

            # unos cuantos productos, en unos cuantos dias, quedan con una
            # pequeña diferencia real para que el motor de alertas del
            # reporte tenga algo que mostrar (igual que en una auditoria real).
            diferencia_unidades = 0.0
            if random.random() < 0.06:
                diferencia_unidades = round(random.uniform(-3, 3) * (1 if tipo == "BOTELLA" else presentacion), 2)

            final_real = max(0, final_teorico + diferencia_unidades)
            diferencia_costo = diferencia_unidades * precio_unitario

            if tipo == "BOTELLA":
                inv_inicial_botellas = round(inicial // presentacion)
                inv_final_botellas = round(final_real // presentacion)
                tara_actual = round(random.uniform(2, 20), 1)
                peso_vacia = 24.0
            else:
                inv_inicial_botellas = 0
                inv_final_botellas = 0
                tara_actual = 0.0
                peso_vacia = 0.0

            fila = {
                "producto": producto,
                "precio_compra": round(precio_compra, 2),
                "presentacion_unidades": presentacion,
                "precio_unitario_excel": round(precio_unitario, 2),
                "inventario_inicial_botellas": inv_inicial_botellas,
                "inventario_inicial_botellas_uni": inv_inicial_botellas * presentacion,
                "inventario_inicial_suelto": round(inicial - inv_inicial_botellas * presentacion, 2),
                "inventario_inicial_total_excel": round(inicial, 2),
                "compras_caja_bot": compras_caja_bot,
                "compras_unidades_excel": round(compras_unidades, 2),
                "disponible_total_excel": round(disponible, 2),
                "ventas_botellas": ventas_botellas,
                "ventas_botellas_unidades_excel": round(ventas_botellas_unidades, 2),
                "ventas_tragos": round(ventas_tragos, 2),
                "cortesias_merma": round(cortesias, 2),
                "salidas_total_excel": round(salidas, 2),
                "inventario_final_botellas": inv_final_botellas,
                "inventario_final_botellas_uni": inv_final_botellas * presentacion,
                "tara_peso_actual_oz": tara_actual,
                "peso_botella_vacia_oz": peso_vacia,
                "inventario_final_suelto": round(final_real - inv_final_botellas * presentacion, 2),
                "inventario_final_total_excel": round(final_real, 2),
                "inventario_teorico_excel": round(final_teorico, 2),
                "diferencia_unidades_excel": diferencia_unidades,
                "diferencia_costo_excel": round(diferencia_costo, 2),
                "costo_almacen_final_excel": round(final_real * precio_unitario, 2),
                "costo_salidas_total_excel": round(salidas * precio_unitario, 2),
                "costo_inventario_inicial_excel": round(inicial * precio_unitario, 2),
                "costo_inicial_mas_compras_excel": round(disponible * precio_unitario, 2),
                "costo_compras_excel": round(compras_unidades * precio_unitario, 2),
            }
            filas_por_fecha[fecha].append(fila)
            stock_actual[producto] = final_real

    return filas_por_fecha


def nombre_hoja(fecha):
    mes_nombre = MESES_NOMBRE[fecha.month]
    dia_semana = DIAS_SEMANA_NOMBRE[fecha.weekday()]
    return f"{ESTABLECIMIENTO} {mes_nombre} {dia_semana} {fecha.day:02d}"


def construir_excel():
    fechas = dias_operados(ANIO, MES, DIAS_OPERACION)
    filas_por_fecha = simular_inventario(fechas)

    with pd.ExcelWriter(ARCHIVO_SALIDA, engine="xlsxwriter") as writer:
        workbook = writer.book
        fmt_titulo = workbook.add_format({"bold": True, "font_size": 12})
        fmt_header = workbook.add_format({"bold": True, "bg_color": "#DDEBF7", "border": 1})

        for fecha in fechas:
            hoja = nombre_hoja(fecha)
            filas = filas_por_fecha[fecha]
            df = pd.DataFrame(filas, columns=COLUMNAS)

            ws = workbook.add_worksheet(hoja[:31])
            writer.sheets[hoja[:31]] = ws
            ws.write(0, 0, f"Inventario diario - {hoja} (DATOS SINTETICOS)", fmt_titulo)
            for col_idx, col in enumerate(COLUMNAS):
                ws.write(1, col_idx, ETIQUETAS[col], fmt_header)
            for row_idx, fila in enumerate(filas, start=2):
                for col_idx, col in enumerate(COLUMNAS):
                    ws.write(row_idx, col_idx, fila[col])
            ws.set_column(0, 0, 26)
            ws.set_column(1, len(COLUMNAS) - 1, 16)

    print(f"Archivo demo generado: {ARCHIVO_SALIDA}")
    print(f"Hojas creadas: {len(fechas)}")


if __name__ == "__main__":
    construir_excel()
