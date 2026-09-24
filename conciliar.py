"""
Automatización (RPA) para la Conciliación de Movimientos de Mercado Pago.
Lee el archivo descargado de Mercado Pago y reconstruye el importe total bruto
de cada operación a partir de los reportes de retenciones y gastos.
"""
from __future__ import annotations

import argparse
import unicodedata
from collections import Counter, defaultdict
from copy import copy
from datetime import date, datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from math import ceil
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent
CENT = Decimal('0.01')
ZERO = Decimal('0')

HEADERS = (
    'RELEASE_DATE', 'TRANSACTION_TYPE', 'REFERENCE_ID', 'TRANSACTION',
    'IDC COMPRAS', 'SIRCUPA', 'SIRTAC', 'GASTOS', 'IDC', 'MOV TOTAL'
)
HEADERS_PENDIENTES = (
    'HOJA ORIGEN', 'FILA ORIGEN', 'ID OPERACION', 'FECHA RETENCION',
    'FECHA OPERACION', 'BASE IMPONIBLE', 'IMPORTE CARGO', 'MOTIVO'
)
NUMBER_FORMAT = '#,##0.00;[Red]-#,##0.00'


def parse_money(value) -> Decimal:
    """Convierte importes de distintos formatos a Decimal con 2 decimales."""
    if value is None or value == '':
        return ZERO
    if isinstance(value, str):
        value = value.strip().replace('$', '').replace(' ', '')
        if ',' in value:
            value = value.replace('.', '').replace(',', '.')
    try:
        return Decimal(str(value)).quantize(CENT, rounding=ROUND_HALF_UP)
    except InvalidOperation:
        raise ValueError(f"Importe monetario inválido: {value!r}")


def parse_id(value) -> str:
    """Conserva identificadores numéricos y alfanuméricos tal como vienen."""
    if value is None:
        return ''
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def norm(text) -> str:
    """Normaliza texto quitando tildes y pasando a minúsculas."""
    return ''.join(
        c for c in unicodedata.normalize('NFKD', str(text or ''))
        if not unicodedata.combining(c)
    ).lower().strip()


def parse_date(value) -> date:
    """Convierte fechas a objeto date."""
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    for fmt in ('%d-%m-%Y', '%d/%m/%Y', '%Y-%m-%d'):
        try:
            return datetime.strptime(str(value), fmt).date()
        except ValueError:
            pass
    raise ValueError(f"Fecha inválida: {value!r}")


def cargar_datos(excel_path: Path, periodo: str, hoja_tp: str = 'sheet0'):
    """Lee el Excel fuente y extrae los movimientos principales y deducciones del período."""
    wb = openpyxl.load_workbook(excel_path, read_only=True, data_only=True)
    try:
        # Resolver hoja principal si no existe el nombre especificado
        if hoja_tp not in wb.sheetnames:
            for alt in ('sheet0', 'TP', 'Detallado'):
                if alt in wb.sheetnames:
                    hoja_tp = alt
                    break

        # 1. Movimientos principales
        source_rows = list(wb[hoja_tp].values)
        movimientos = []
        for rownum, row in enumerate(source_rows, 1):
            if rownum <= 4:
                continue
            if not row or not row[0]:
                break  # Fin de la tabla principal de movimientos
            if not row[1] or row[2] is None:
                continue
            try:
                d = parse_date(row[0])
            except ValueError:
                continue
            if d.strftime('%Y-%m') != periodo:
                continue

            movimientos.append({
                'row': rownum,
                'date': d,
                'type': str(row[1]),
                'id': parse_id(row[2]),
                'net': parse_money(row[3]),
                'charges': []
            })

        # 2. Cargos de deducciones (SIRCUPA, SIRTAC, IDC, Gastos)
        cargos = []
        for sheet_name in ('SIRCUPA', 'SIRTAC', 'IDC', 'Gastos'):
            if sheet_name not in wb.sheetnames:
                continue

            started = False
            for rownum, row in enumerate(wb[sheet_name].values, 1):
                if not row:
                    continue
                primera_col = norm(row[0])
                if primera_col in ('numero de operacion', 'operacion relacionada'):
                    started = True
                    continue
                if not started:
                    continue
                if not row[0]:
                    break  # Fin de la tabla oficial de Mercado Pago; se ignoran notas o controles posteriores

                # Fechas e importes según la hoja correspondiente
                try:
                    d = parse_date(row[2] if sheet_name == 'Gastos' else row[4])
                except ValueError:
                    continue

                # Filtrar fuera del período seleccionado
                if d.strftime('%Y-%m') != periodo:
                    continue

                # En Gastos, excluir los marcados con 'No' en cobrado en operación
                if sheet_name == 'Gastos' and norm(row[6]) == 'no':
                    continue

                try:
                    op_date = parse_date(row[2] if sheet_name != 'SIRCUPA' else row[1]) if sheet_name != 'Gastos' else d
                except ValueError:
                    op_date = d

                monto = parse_money(row[9] if sheet_name == 'Gastos' else row[7])
                base = parse_money(row[16] if sheet_name == 'Gastos' else row[6])

                cargos.append({
                    'sheet': sheet_name,
                    'row': rownum,
                    'id': parse_id(row[0]),
                    'date': d,
                    'op_date': op_date,
                    'amount': monto,
                    'base': base,
                    'target': None
                })

        return movimientos, cargos
    finally:
        wb.close()


def es_impuesto_extraccion(m) -> bool:
    return norm(m['type']) == 'impuesto por extraccion'


def es_anulacion(m) -> bool:
    t = norm(m['type'])
    return 'cancelada' in t or 'devolucion' in t


def total_cargos(m, sheet_name: str | None = None) -> Decimal:
    return sum((c['amount'] for c in m['charges'] if sheet_name is None or c['sheet'] == sheet_name), ZERO)


def asociar_cargos(movimientos: list[dict], cargos: list[dict], tolerancia: Decimal = CENT):
    """Asocia retenciones y comisiones a sus operaciones correspondientes en TP."""
    por_id = defaultdict(list)
    for m in movimientos:
        if not es_impuesto_extraccion(m):
            por_id[m['id']].append(m)

    # Nivel 1: Asociación directa por ID de operación
    for c in cargos:
        candidatos = [m for m in por_id[c['id']] if es_anulacion(m) == (c['amount'] < 0)]
        if len(candidatos) > 1:
            candidatos = [m for m in candidatos if m['date'] in (c['date'], c['op_date'])]
        if len(candidatos) == 1:
            c['target'] = candidatos[0]['row']
            candidatos[0]['charges'].append(c)

    # Nivel 2: IDC con identificadores alfanuméricos internos
    propuestas = []
    for c in cargos:
        if c['target'] is not None or c['sheet'] != 'IDC':
            continue

        candidatos = []
        for m in movimientos:
            if es_impuesto_extraccion(m) or any(x['sheet'] == 'IDC' for x in m['charges']):
                continue
            if m['date'] not in (c['date'], c['op_date']):
                continue
            if es_anulacion(m) != (c['amount'] < 0):
                continue

            esperado = c['base'] if m['net'] >= 0 or es_anulacion(m) else -c['base']
            if abs(m['net'] + total_cargos(m) + c['amount'] - esperado) <= tolerancia:
                if m['net'] > 0 and not any(abs(x['base'] - c['base']) <= tolerancia for x in m['charges']):
                    continue
                candidatos.append(m)

        if len(candidatos) == 1:
            propuestas.append((c, candidatos[0]))

    # Asignación de IDC únicos (evita colisiones o asignaciones ambiguas)
    conteo_movs = Counter(m['row'] for _, m in propuestas)
    for c, m in propuestas:
        if conteo_movs[m['row']] == 1:
            c['target'] = m['row']
            m['charges'].append(c)


def calcular_totales(movimientos: list[dict]):
    """Calcula el total reconstruido para cada fila de TP."""
    separados = defaultdict(list)
    for m in movimientos:
        if es_impuesto_extraccion(m):
            separados[(m['id'], m['date'])].append(m)

    filas_resultado = []
    for m in movimientos:
        sircupa = total_cargos(m, 'SIRCUPA')
        sirtac = total_cargos(m, 'SIRTAC')
        idc = total_cargos(m, 'IDC')
        gastos = total_cargos(m, 'Gastos')

        ajuste = sircupa + sirtac + idc + gastos
        filas_sep = separados[(m['id'], m['date'])] if not es_impuesto_extraccion(m) else []
        if filas_sep:
            ajuste -= idc

        reconstruido = m['net'] + ajuste

        if es_impuesto_extraccion(m):
            reconstruido = None

        filas_resultado.append({
            'fila_origen': m['row'],
            'fecha': m['date'],
            'tipo': m['type'],
            'reference_id': m['id'],
            'neto': m['net'],
            'sircupa': sircupa,
            'sirtac': sirtac,
            'idc': idc,
            'gastos': gastos,
            'total': reconstruido
        })
    return filas_resultado


def generar_excel(
    archivo_origen: Path,
    archivo_destino: Path,
    resultados: list[dict],
    hoja_tp: str = 'sheet0',
    cargos_pendientes: list[dict] | None = None,
    movimientos: list[dict] | None = None
):
    """Genera el Excel final con formato contable profesional, las 10 columnas oficiales y hoja de pendientes."""
    template = openpyxl.load_workbook(archivo_origen, data_only=True)
    wb_salida = openpyxl.Workbook()
    ws = wb_salida.active
    ws.title = hoja_tp

    try:
        if hoja_tp not in template.sheetnames:
            for alt in ('sheet0', 'TP', 'Detallado'):
                if alt in template.sheetnames:
                    hoja_tp = alt
                    break

        ws_original = template[hoja_tp]
        ws.append(HEADERS)

        # Mapeo de columnas originales para estilos:
        # RELEASE_DATE, TRANSACTION_TYPE, REFERENCE_ID, TRANSACTION, IDC COMPRAS
        cols_origen = (1, 2, 3, 4, 8, 9, 10, 11, 12, 13)

        for row_idx, res in enumerate(resultados, start=2):
            src_row = res['fila_origen']
            total = res['total']

            valores = (
                ws_original.cell(src_row, 1).value,
                res['tipo'],
                res['reference_id'],
                res['neto'],
                ws_original.cell(src_row, 8).value,  # IDC COMPRAS
                -res['sircupa'],
                -res['sirtac'],
                -res['gastos'],
                -res['idc'],
                res['neto'] if total is None else total
            )

            for col_idx, val in enumerate(valores, start=1):
                cell = ws.cell(row_idx, col_idx, float(val) if isinstance(val, Decimal) else val)
                src_cell = ws_original.cell(src_row, cols_origen[col_idx - 1])

                for attr in ('font', 'fill', 'border'):
                    setattr(cell, attr, copy(getattr(src_cell, attr)))

                cell.alignment = Alignment(
                    vertical='center',
                    horizontal='right' if col_idx >= 4 else 'left',
                    wrap_text=(col_idx == 2)
                )
                cell.number_format = NUMBER_FORMAT if col_idx >= 4 else ('@' if col_idx == 3 else src_cell.number_format)

            ws.row_dimensions[row_idx].height = max(20, 16 * ceil(len(res['tipo']) / 65))

        # Estilo del encabezado y anchos de columnas en hoja principal
        anchos = (18, 74, 21, 20, 18, 18, 18, 18, 18, 21)
        for col_idx, width in enumerate(anchos, start=1):
            ws.column_dimensions[get_column_letter(col_idx)].width = width
            h_cell = ws.cell(1, col_idx)
            h_cell.font = Font(name='Calibri', size=11, bold=True)
            h_cell.fill = PatternFill('solid', fgColor='E7E6E6')
            h_cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

        ws.row_dimensions[1].height = 32
        ws.freeze_panes = 'D2'
        ws.auto_filter.ref = f"A1:J{ws.max_row}"
        ws.sheet_view.zoomScale = 85

        # -------------------------------------------------------------
        # 2. Hoja de Cargos No Conciliados / Pendientes
        # -------------------------------------------------------------
        cargos_pend = cargos_pendientes or []
        ws_pend = wb_salida.create_sheet(title='No Conciliados')
        ws_pend.sheet_properties.tabColor = "FFC000" if cargos_pend else "70AD47"
        ws_pend.append(HEADERS_PENDIENTES)

        ids_movs = {m['id'] for m in (movimientos or [])}

        if cargos_pend:
            for row_idx, c in enumerate(cargos_pend, start=2):
                if c['id'] not in ids_movs:
                    if c['sheet'] == 'IDC' and any(ch.isalpha() for ch in c['id']):
                        motivo = 'Código interno sin operación vinculable en fecha o base imponible'
                    else:
                        motivo = 'ID de operación no encontrado en los movimientos del período'
                else:
                    motivo = 'Candidato no unívoco o discrepancia en importe/anulación'

                f_ret = c['date'].strftime('%d/%m/%Y') if hasattr(c['date'], 'strftime') else str(c['date'])
                f_op = c['op_date'].strftime('%d/%m/%Y') if hasattr(c['op_date'], 'strftime') else str(c['op_date'])

                valores_pend = (
                    c['sheet'],
                    c['row'],
                    c['id'],
                    f_ret,
                    f_op,
                    float(c['base']),
                    float(c['amount']),
                    motivo
                )
                for col_idx, val in enumerate(valores_pend, start=1):
                    cell = ws_pend.cell(row_idx, col_idx, val)
                    cell.font = Font(name='Calibri', size=10)
                    cell.alignment = Alignment(
                        vertical='center',
                        horizontal='right' if col_idx in (6, 7) else ('left' if col_idx == 8 else 'center')
                    )
                    if col_idx in (6, 7):
                        cell.number_format = NUMBER_FORMAT
                    elif col_idx in (2, 3):
                        cell.number_format = '@'

                ws_pend.row_dimensions[row_idx].height = 20

            # Fila totalizadora de cargos no asociados
            total_row_idx = len(cargos_pend) + 2
            total_amt = float(sum((c['amount'] for c in cargos_pend), ZERO))
            ws_pend.cell(total_row_idx, 1, 'TOTAL PENDIENTE')
            ws_pend.cell(total_row_idx, 7, total_amt)
            ws_pend.cell(total_row_idx, 8, f"{len(cargos_pend)} cargos sin asociar")

            for col_idx in range(1, 9):
                cell = ws_pend.cell(total_row_idx, col_idx)
                cell.font = Font(name='Calibri', size=11, bold=True)
                if col_idx == 7:
                    cell.number_format = NUMBER_FORMAT
                    cell.alignment = Alignment(horizontal='right', vertical='center')
                cell.fill = PatternFill('solid', fgColor='F2DCDB')
            ws_pend.row_dimensions[total_row_idx].height = 22
            ws_pend.auto_filter.ref = f"A1:H{len(cargos_pend) + 1}"
        else:
            ws_pend.cell(2, 1, "Sin cargos pendientes: el 100% de las retenciones y deducciones fue conciliado con éxito.")
            ws_pend.cell(2, 1).font = Font(name='Calibri', size=11, italic=True, color='375623')
            ws_pend.row_dimensions[2].height = 25

        # Estilo del encabezado en hoja No Conciliados
        anchos_pend = (16, 14, 24, 18, 18, 18, 18, 60)
        for col_idx, width in enumerate(anchos_pend, start=1):
            ws_pend.column_dimensions[get_column_letter(col_idx)].width = width
            h_cell = ws_pend.cell(1, col_idx)
            h_cell.font = Font(name='Calibri', size=11, bold=True)
            h_cell.fill = PatternFill('solid', fgColor='FFF2CC' if cargos_pend else 'E2EFDA')
            h_cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

        ws_pend.row_dimensions[1].height = 28
        ws_pend.freeze_panes = 'A2'
        ws_pend.sheet_view.zoomScale = 85

        archivo_destino.parent.mkdir(parents=True, exist_ok=True)
        destino_final = archivo_destino
        try:
            wb_salida.save(archivo_destino)
        except PermissionError:
            import time
            ts = int(time.time())
            destino_final = archivo_destino.with_name(f"{archivo_destino.stem}_{ts}.xlsx")
            wb_salida.save(destino_final)
            print(f"-> AVISO: El archivo '{archivo_destino.name}' está abierto en Excel.")
            print(f"-> Se guardó como copia alternativa en: {destino_final.name}")

        return destino_final
    finally:
        template.close()
        wb_salida.close()


def main():
    import re

    parser = argparse.ArgumentParser(description="Conciliación automatizada de Mercado Pago")
    parser.add_argument('--archivo', type=Path, default=ROOT / 'docs' / 'MP 2025-12 (M).xlsx')
    parser.add_argument('--salida', type=Path, default=ROOT / 'outputs')
    parser.add_argument('--periodo', default=None, help="Período YYYY-MM (se auto-detecta si no se especifica)")
    parser.add_argument('--hoja', default='sheet0')
    args = parser.parse_args()

    periodo = args.periodo
    if not periodo:
        match = re.search(r'20\d{2}-\d{2}', args.archivo.name)
        periodo = match.group(0) if match else '2025-12'

    print("=" * 60)
    print(f" Iniciando Conciliación de Mercado Pago ({periodo})")
    print("=" * 60)

    movimientos, cargos = cargar_datos(args.archivo, periodo, args.hoja)
    print(f"-> Movimientos leídos de '{args.hoja}': {len(movimientos)}")
    print(f"-> Retenciones y cargos leídos: {len(cargos)}")

    asociar_cargos(movimientos, cargos)
    cargos_asociados = sum(1 for c in cargos if c['target'] is not None)
    cargos_pendientes = [c for c in cargos if c['target'] is None]

    print(f"-> Cargos asociados con éxito: {cargos_asociados}")
    if cargos_pendientes:
        print(f"-> Cargos NO conciliados: {len(cargos_pendientes)} (detallados en pestaña 'No Conciliados')")
    else:
        print(f"-> Cargos NO conciliados: 0 (100% conciliado con éxito)")

    resultados = calcular_totales(movimientos)

    excel_salida = args.salida / f"conciliacion_{periodo}.xlsx"
    print(f"-> Generando Excel con totales en: {excel_salida.name}")
    guardado_en = generar_excel(args.archivo, excel_salida, resultados, args.hoja, cargos_pendientes, movimientos)

    print("=" * 60)
    print(" Proceso finalizado correctamente!")
    print(f" Archivo listo en: {guardado_en.resolve()}")
    print("=" * 60)


if __name__ == '__main__':
    main()
