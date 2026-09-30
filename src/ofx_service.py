from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import datetime
from io import BytesIO
from typing import Any

import pandas as pd
from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation


@dataclass(frozen=True)
class OfxStatement:
    metadata: dict[str, str]
    transactions: pd.DataFrame


def decode_ofx(data: bytes) -> str:
    """Decode OFX files that frequently declare a different encoding than they use."""
    if data.startswith(b"\xef\xbb\xbf"):
        return data.decode("utf-8-sig")
    for encoding in ("utf-8", "cp1252", "latin-1"):
        try:
            return data.decode(encoding)
        except UnicodeDecodeError:
            continue
    return data.decode("latin-1", errors="replace")


def _tag(block: str, name: str, default: str = "") -> str:
    match = re.search(rf"<{name}>\s*([^<\r\n]+)", block, flags=re.IGNORECASE)
    return match.group(1).strip() if match else default


def _ofx_date(value: str) -> pd.Timestamp | pd.NaT:
    digits = re.sub(r"\D", "", value)[:8]
    if len(digits) != 8:
        return pd.NaT
    return pd.to_datetime(digits, format="%Y%m%d", errors="coerce")


def parse_ofx(data: bytes) -> OfxStatement:
    text = decode_ofx(data)
    blocks = re.findall(r"<STMTTRN>(.*?)(?:</STMTTRN>|(?=<STMTTRN>|</BANKTRANLIST>))", text, re.I | re.S)
    if not blocks:
        raise ValueError("Nenhum lançamento <STMTTRN> foi encontrado no arquivo OFX.")

    rows: list[dict[str, Any]] = []
    for block in blocks:
        amount_text = _tag(block, "TRNAMT", "0").replace(",", ".")
        try:
            amount = float(amount_text)
        except ValueError:
            amount = 0.0
        raw_type = _tag(block, "TRNTYPE", "OTHER").upper()
        friendly_type = "Crédito" if amount >= 0 else "Débito"
        rows.append(
            {
                "Conferido": "Não",
                "Data": _ofx_date(_tag(block, "DTPOSTED")),
                "Tipo": friendly_type,
                "Valor": amount,
                "Documento": _tag(block, "CHECKNUM"),
                "Descrição": _tag(block, "NAME"),
                "Histórico": _tag(block, "MEMO"),
                "ID OFX": _tag(block, "FITID"),
                "Tipo OFX": raw_type,
            }
        )

    frame = pd.DataFrame(rows)
    metadata = {
        "Banco": _tag(text, "BANKID"),
        "Agência": _tag(text, "BRANCHID"),
        "Conta": _tag(text, "ACCTID"),
        "Moeda": _tag(text, "CURDEF", "BRL"),
        "Saldo": _tag(text, "BALAMT"),
        "Início": _tag(text, "DTSTART"),
        "Fim": _tag(text, "DTEND"),
    }
    return OfxStatement(metadata=metadata, transactions=frame)


def build_excel(statement: OfxStatement, source_name: str = "extrato.ofx") -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Lançamentos"
    summary = wb.create_sheet("Resumo")
    instructions = wb.create_sheet("Como conferir")

    navy = "17324D"
    blue = "1F6E8C"
    teal = "159A9C"
    pale = "EAF3F5"
    green = "DFF3E4"
    red = "FCE2E2"
    white = "FFFFFF"
    gray = "667085"
    thin = Side(style="thin", color="D0D5DD")

    headers = ["Conferido", "Data", "Tipo", "Valor (R$)", "Documento", "Descrição", "Histórico", "ID OFX", "Tipo OFX"]
    start_row = 8
    end_row = start_row + len(statement.transactions) - 1

    ws.merge_cells("A1:I2")
    ws["A1"] = "CONFERÊNCIA DE LANÇAMENTOS"
    ws["A1"].font = Font(size=18, bold=True, color=white)
    ws["A1"].fill = PatternFill("solid", fgColor=navy)
    ws["A1"].alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 25
    ws["A3"] = f"Origem: {source_name}"
    ws["A3"].font = Font(color=gray, italic=True)

    metrics = [
        ("A4", "Entradas", f'=SUMIF(C{start_row}:C{end_row},"Crédito",D{start_row}:D{end_row})'),
        ("C4", "Saídas", f'=SUMIF(C{start_row}:C{end_row},"Débito",D{start_row}:D{end_row})'),
        ("E4", "Resultado", f"=SUM(D{start_row}:D{end_row})"),
        ("G4", "Total conferido", f'=SUMIF(A{start_row}:A{end_row},"Sim",D{start_row}:D{end_row})'),
    ]
    for anchor, label, formula in metrics:
        col = ws[anchor].column
        ws.cell(4, col, label)
        ws.cell(4, col).font = Font(bold=True, color=gray, size=10)
        ws.merge_cells(start_row=4, start_column=col, end_row=4, end_column=col + 1)
        ws.cell(5, col, formula)
        ws.cell(5, col).font = Font(bold=True, color=navy, size=14)
        ws.cell(5, col).number_format = 'R$ #,##0.00;[Red]-R$ #,##0.00'
        ws.merge_cells(start_row=5, start_column=col, end_row=5, end_column=col + 1)

    for col, header in enumerate(headers, 1):
        cell = ws.cell(7, col, header)
        cell.font = Font(bold=True, color=white)
        cell.fill = PatternFill("solid", fgColor=blue)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(bottom=thin)
    ws.row_dimensions[7].height = 28

    display_columns = ["Conferido", "Data", "Tipo", "Valor", "Documento", "Descrição", "Histórico", "ID OFX", "Tipo OFX"]
    for row_idx, row in enumerate(statement.transactions[display_columns].itertuples(index=False, name=None), start_row):
        for col_idx, value in enumerate(row, 1):
            cell = ws.cell(row_idx, col_idx, value)
            cell.border = Border(bottom=thin)
            cell.alignment = Alignment(vertical="top", wrap_text=col_idx in (6, 7))
            if row_idx % 2 == 0:
                cell.fill = PatternFill("solid", fgColor="F8FAFC")
        ws.cell(row_idx, 2).number_format = "dd/mm/yyyy"
        ws.cell(row_idx, 4).number_format = 'R$ #,##0.00;[Red]-R$ #,##0.00'
        ws.cell(row_idx, 3).font = Font(color="19725A" if row[3] >= 0 else "B42318")

    if len(statement.transactions):
        validation = DataValidation(type="list", formula1='"Não,Sim"', allow_blank=False)
        validation.promptTitle = "Conferência"
        validation.prompt = "Selecione Sim após validar o lançamento."
        ws.add_data_validation(validation)
        validation.add(f"A{start_row}:A{end_row}")
        ws.conditional_formatting.add(
            f"A{start_row}:I{end_row}",
            FormulaRule(formula=[f'$A{start_row}="Sim"'], fill=PatternFill("solid", fgColor=green)),
        )
        ws.conditional_formatting.add(
            f"D{start_row}:D{end_row}",
            CellIsRule(operator="lessThan", formula=["0"], font=Font(color="B42318")),
        )

        total_row = end_row + 2
        ws.cell(total_row, 3, "TOTAL VISÍVEL")
        ws.cell(total_row, 3).font = Font(bold=True, color=white)
        ws.cell(total_row, 3).fill = PatternFill("solid", fgColor=navy)
        ws.cell(total_row, 4, f"=SUBTOTAL(109,D{start_row}:D{end_row})")
        ws.cell(total_row, 4).font = Font(bold=True, color=white)
        ws.cell(total_row, 4).fill = PatternFill("solid", fgColor=navy)
        ws.cell(total_row, 4).number_format = 'R$ #,##0.00;[Red]-R$ #,##0.00'

    ws.auto_filter.ref = f"A7:I{max(end_row, 7)}"
    ws.freeze_panes = "A8"
    ws.sheet_view.showGridLines = False
    widths = [15, 13, 12, 16, 16, 34, 30, 22, 14]
    for idx, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(idx)].width = width

    summary.merge_cells("A1:F2")
    summary["A1"] = "RESUMO DO EXTRATO"
    summary["A1"].font = Font(size=18, bold=True, color=white)
    summary["A1"].fill = PatternFill("solid", fgColor=navy)
    summary["A1"].alignment = Alignment(vertical="center")
    summary.sheet_view.showGridLines = False
    meta_rows = [("Arquivo", source_name), ("Banco", statement.metadata.get("Banco", "")), ("Agência", statement.metadata.get("Agência", "")), ("Conta", statement.metadata.get("Conta", "")), ("Moeda", statement.metadata.get("Moeda", "BRL"))]
    for idx, (label, value) in enumerate(meta_rows, 4):
        summary.cell(idx, 1, label).font = Font(bold=True, color=gray)
        summary.cell(idx, 2, value)
    summary["A11"] = "Indicador"
    summary["B11"] = "Valor"
    for cell in summary[11]:
        if cell.column <= 2:
            cell.fill = PatternFill("solid", fgColor=blue)
            cell.font = Font(bold=True, color=white)
    report_rows = [
        ("Total de lançamentos", f"=COUNTA('Lançamentos'!B{start_row}:B{end_row})", "0"),
        ("Créditos", f'=SUMIF(\'Lançamentos\'!C{start_row}:C{end_row},"Crédito",\'Lançamentos\'!D{start_row}:D{end_row})', 'R$ #,##0.00'),
        ("Débitos", f'=SUMIF(\'Lançamentos\'!C{start_row}:C{end_row},"Débito",\'Lançamentos\'!D{start_row}:D{end_row})', 'R$ #,##0.00;[Red]-R$ #,##0.00'),
        ("Resultado líquido", f"=SUM('Lançamentos'!D{start_row}:D{end_row})", 'R$ #,##0.00;[Red]-R$ #,##0.00'),
        ("Itens conferidos", f'=COUNTIF(\'Lançamentos\'!A{start_row}:A{end_row},"Sim")', "0"),
        ("Itens pendentes", f'=COUNTIF(\'Lançamentos\'!A{start_row}:A{end_row},"Não")', "0"),
        ("Valor conferido", f'=SUMIF(\'Lançamentos\'!A{start_row}:A{end_row},"Sim",\'Lançamentos\'!D{start_row}:D{end_row})', 'R$ #,##0.00;[Red]-R$ #,##0.00'),
    ]
    for idx, (label, formula, number_format) in enumerate(report_rows, 12):
        summary.cell(idx, 1, label)
        summary.cell(idx, 2, formula)
        summary.cell(idx, 2).number_format = number_format
        if idx % 2 == 0:
            summary.cell(idx, 1).fill = PatternFill("solid", fgColor=pale)
            summary.cell(idx, 2).fill = PatternFill("solid", fgColor=pale)
    summary.column_dimensions["A"].width = 28
    summary.column_dimensions["B"].width = 24

    instructions.sheet_view.showGridLines = False
    instructions.merge_cells("A1:F2")
    instructions["A1"] = "COMO FAZER A CONFERÊNCIA"
    instructions["A1"].font = Font(size=18, bold=True, color=white)
    instructions["A1"].fill = PatternFill("solid", fgColor=navy)
    steps = [
        "1. Abra a aba Lançamentos e compare cada item com seus comprovantes.",
        "2. Na coluna Conferido, troque “Não” por “Sim”. A linha ficará verde.",
        "3. Use os filtros para analisar datas, tipos e descrições.",
        "4. O total visível acompanha os filtros aplicados.",
        "5. A aba Resumo atualiza os totais automaticamente no Excel.",
    ]
    for idx, value in enumerate(steps, 4):
        instructions.cell(idx, 1, value)
        instructions.cell(idx, 1).alignment = Alignment(wrap_text=True, vertical="top")
    instructions.column_dimensions["A"].width = 92
    instructions["A11"] = "Atenção: confira os dados antes de usá-los para fins contábeis ou bancários."
    instructions["A11"].font = Font(bold=True, color="B54708")

    try:
        from openpyxl.workbook.properties import CalcProperties

        wb.calculation = CalcProperties(calcMode="auto", fullCalcOnLoad=True, forceFullCalc=True)
    except ImportError:
        pass

    output = BytesIO()
    wb.save(output)
    return output.getvalue()

