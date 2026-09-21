"""
Generador de histórico mensual de nómina y costos de personal
----------------------------------------------------------------
Toma la base maestra de empleados (con cargo, IBC vigente 2026 y nivel de
riesgo ARL ya asignados) y construye una tabla de hechos mensual (formato
largo, ideal para Power BI / Tableau) para el periodo SEP-2024 a AGO-2026
(24 meses).

Reglas de negocio aplicadas (normativa laboral colombiana):
- El salario histórico de cada empleado se escala en la misma proporción
  en que varió el SMLV año a año. Así, quienes ganan el salario mínimo
  coinciden exactamente con el SMLV vigente cada año, y los demás cargos
  suben en la misma proporción (práctica común de incremento salarial).
- El auxilio de transporte solo se paga si el IBC del mes es <= 2 SMLV
  vigentes ese año.
- Un empleado solo se considera ACTIVO en los meses comprendidos entre su
  FECHA INGRESO y su FECHA RETIRO (o hasta el mes actual si sigue activo).
- Los aportes a seguridad social y las provisiones de prestaciones sociales
  se calculan con las fórmulas y porcentajes vigentes por ley (ver
  hoja "Parametros" del archivo de salida).

Fuente valores históricos SMLV / Auxilio de transporte:
  Decretos del Ministerio del Trabajo de Colombia 2022-2025.
"""

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows
from dateutil.relativedelta import relativedelta
from datetime import datetime

# ------------------------------------------------------------------
# 1. Parámetros normativos históricos (fuente: decretos MinTrabajo)
# ------------------------------------------------------------------
SMLV = {2023: 1_160_000, 2024: 1_300_000, 2025: 1_423_500, 2026: 1_750_905}
AUX_TRANSPORTE = {2023: 140_606, 2024: 162_000, 2025: 200_000, 2026: 249_095}

PCT_SALUD_TRABAJADOR = 0.04
PCT_PENSION_TRABAJADOR = 0.04
PCT_SALUD_EMPLEADOR = 0.085
PCT_PENSION_EMPLEADOR = 0.12
PCT_CCF_EMPLEADOR = 0.04
PCT_CESANTIAS_MES = 1 / 12
PCT_INTERESES_CESANTIAS = 0.12
PCT_PRIMA_MES = 1 / 12
PCT_VACACIONES = 0.0417

PERIODO_INICIO = datetime(2024, 9, 1)
PERIODO_FIN = datetime(2026, 8, 1)  # 24 meses


def meses_periodo(inicio, fin):
    meses = []
    m = inicio
    while m <= fin:
        meses.append(m)
        m = m + relativedelta(months=1)
    return meses


def cargar_empleados(path_maestro):
    wb = load_workbook(path_maestro, data_only=True)
    ws = wb['Nomina']
    headers = [c.value for c in ws[1]]
    rows = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        rows.append(dict(zip(headers, r)))
    df = pd.DataFrame(rows)
    return df[['ID EMPLEADO', 'PRIMER APELLIDO', 'PRIMER NOMBRE', 'SEXO',
               'DEPARTAMENTO', 'MUNICIPIO', 'OCUPACION', 'FECHA INGRESO',
               'FECHA RETIRO', 'EPS', 'AFP', 'ARL', 'CCF', 'NIVEL RIESGO', 'IBC']]


def generar_historico(df_emp, meses):
    registros = []
    for _, emp in df_emp.iterrows():
        fecha_ingreso = pd.to_datetime(emp['FECHA INGRESO'])
        fecha_retiro = pd.to_datetime(emp['FECHA RETIRO']) if pd.notna(emp['FECHA RETIRO']) else None
        ibc_2026 = float(emp['IBC'])
        riesgo = float(emp['NIVEL RIESGO'])

        for mes in meses:
            fin_mes = mes + relativedelta(months=1) - relativedelta(days=1)
            activo = fecha_ingreso <= fin_mes and (fecha_retiro is None or fecha_retiro >= mes)
            if not activo:
                continue

            anio = mes.year if mes.year in SMLV else 2026
            smlv_anio = SMLV[anio]
            aux_anio = AUX_TRANSPORTE[anio]

            ibc_mes = round(ibc_2026 * (smlv_anio / SMLV[2026]))
            aux_transp = aux_anio if ibc_mes <= 2 * smlv_anio else 0

            desc_eps = ibc_mes * PCT_SALUD_TRABAJADOR
            desc_pension = ibc_mes * PCT_PENSION_TRABAJADOR
            base_prestacional = ibc_mes + aux_transp
            cesantias = base_prestacional * PCT_CESANTIAS_MES
            intereses_cesantias = cesantias * PCT_INTERESES_CESANTIAS
            provision_prima = base_prestacional * PCT_PRIMA_MES
            provision_vacaciones = ibc_mes * PCT_VACACIONES
            aporte_salud_empl = ibc_mes * PCT_SALUD_EMPLEADOR
            aporte_pension_empl = ibc_mes * PCT_PENSION_EMPLEADOR
            aporte_ccf_empl = ibc_mes * PCT_CCF_EMPLEADOR
            aporte_arl_empl = ibc_mes * riesgo

            costo_total_empleador = (ibc_mes + aux_transp + aporte_salud_empl +
                                      aporte_pension_empl + aporte_ccf_empl +
                                      aporte_arl_empl + cesantias +
                                      intereses_cesantias + provision_prima +
                                      provision_vacaciones)

            registros.append({
                'PERIODO': mes.strftime('%Y-%m'),
                'ANIO': mes.year,
                'MES': mes.month,
                'ID EMPLEADO': emp['ID EMPLEADO'],
                'PRIMER APELLIDO': emp['PRIMER APELLIDO'],
                'PRIMER NOMBRE': emp['PRIMER NOMBRE'],
                'SEXO': emp['SEXO'],
                'DEPARTAMENTO': emp['DEPARTAMENTO'],
                'MUNICIPIO': emp['MUNICIPIO'],
                'OCUPACION': emp['OCUPACION'],
                'EPS': emp['EPS'], 'AFP': emp['AFP'], 'ARL': emp['ARL'], 'CCF': emp['CCF'],
                'NIVEL RIESGO': riesgo,
                'SMLV VIGENTE': smlv_anio,
                'IBC': ibc_mes,
                'AUX DE TRANSPORTE': aux_transp,
                'DESCUENTO TRABAJADOR EPS': round(desc_eps, 1),
                'DESCUENTO TRABAJADOR PENSION': round(desc_pension, 1),
                'CESANTIAS': round(cesantias, 1),
                'INTERESES DE CESANTIAS': round(intereses_cesantias, 1),
                'PROVISION PRIMA': round(provision_prima, 1),
                'PROVISION VACACIONES': round(provision_vacaciones, 1),
                'APORTE SALUD EMPLEADOR': round(aporte_salud_empl, 1),
                'APORTE PENSION EMPLEADOR': round(aporte_pension_empl, 1),
                'APORTE CCF EMPLEADOR': round(aporte_ccf_empl, 1),
                'APORTE ARL EMPLEADOR': round(aporte_arl_empl, 1),
                'COSTO TOTAL EMPLEADOR': round(costo_total_empleador, 1),
            })
    return pd.DataFrame(registros)


def escribir_excel(df_hist, df_emp, path_salida):
    wb = Workbook()

    # ---- Parametros ----
    ws_p = wb.active
    ws_p.title = 'Parametros'
    ws_p.append(['Año', 'SMLV', 'Auxilio de Transporte'])
    for anio in sorted(SMLV):
        ws_p.append([anio, SMLV[anio], AUX_TRANSPORTE[anio]])
    ws_p.append([])
    ws_p.append(['Parametro', 'Valor', 'Fuente / Norma'])
    otros = [
        ('% Aporte Salud Trabajador', PCT_SALUD_TRABAJADOR, 'Ley 100 de 1993'),
        ('% Aporte Pension Trabajador', PCT_PENSION_TRABAJADOR, 'Ley 100 de 1993'),
        ('% Aporte Salud Empleador', PCT_SALUD_EMPLEADOR, 'Ley 100 de 1993'),
        ('% Aporte Pension Empleador', PCT_PENSION_EMPLEADOR, 'Ley 100 de 1993'),
        ('% Aporte CCF Empleador', PCT_CCF_EMPLEADOR, 'Ley 21 de 1982'),
        ('% Provision Cesantias mensual', PCT_CESANTIAS_MES, 'Ley 50 de 1990'),
        ('% Intereses sobre Cesantias', PCT_INTERESES_CESANTIAS, 'Ley 52 de 1975'),
        ('% Provision Prima mensual', PCT_PRIMA_MES, 'CST Art. 306'),
        ('% Provision Vacaciones', PCT_VACACIONES, 'CST Art. 186'),
    ]
    for row in otros:
        ws_p.append(list(row))
    for cell in ws_p[1]:
        cell.font = Font(name='Arial', bold=True, color='FFFFFF')
        cell.fill = PatternFill('solid', fgColor='1F4E78')
    ws_p.column_dimensions['A'].width = 30
    ws_p.column_dimensions['B'].width = 14
    ws_p.column_dimensions['C'].width = 40

    # ---- Empleados (maestro) ----
    ws_e = wb.create_sheet('Empleados')
    for r in dataframe_to_rows(df_emp, index=False, header=True):
        ws_e.append(r)
    for cell in ws_e[1]:
        cell.font = Font(name='Arial', bold=True, color='FFFFFF')
        cell.fill = PatternFill('solid', fgColor='1F4E78')

    # ---- Historico_Mensual ----
    ws_h = wb.create_sheet('Historico_Mensual')
    for r in dataframe_to_rows(df_hist, index=False, header=True):
        ws_h.append(r)
    for cell in ws_h[1]:
        cell.font = Font(name='Arial', bold=True, color='FFFFFF', size=9)
        cell.fill = PatternFill('solid', fgColor='1F4E78')
        cell.alignment = Alignment(horizontal='center', wrap_text=True)
    ws_h.freeze_panes = 'A2'
    for i in range(1, df_hist.shape[1] + 1):
        ws_h.column_dimensions[get_column_letter(i)].width = 15

    wb.save(path_salida)


if __name__ == '__main__':
    meses = meses_periodo(PERIODO_INICIO, PERIODO_FIN)
    df_emp = cargar_empleados('/mnt/user-data/outputs/nomina_transporte_logistica.xlsx')
    df_hist = generar_historico(df_emp, meses)
    escribir_excel(df_hist, df_emp, '/home/claude/proyecto_nomina/nomina_historica_24m.xlsx')
    print('Filas generadas:', len(df_hist))
    print('Periodos:', df_hist['PERIODO'].nunique())
    print('Empleados unicos:', df_hist['ID EMPLEADO'].nunique())
    print(df_hist.groupby('PERIODO').agg(
        headcount=('ID EMPLEADO', 'count'),
        costo_total=('COSTO TOTAL EMPLEADOR', 'sum')
    ))
