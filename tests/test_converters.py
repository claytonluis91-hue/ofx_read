from pathlib import Path

import pytest
from openpyxl import load_workbook

from src.ofx_service import build_excel, parse_ofx
from src.pdf_service import build_ofx, parse_pdf


EXAMPLES = Path(__file__).parents[1] / "Exemplos"
OFX_EXAMPLES = list(EXAMPLES.glob("*.ofx"))
PDF_EXAMPLES = list(EXAMPLES.glob("*.pdf"))


@pytest.mark.skipif(not OFX_EXAMPLES, reason="Exemplo OFX privado não está disponível neste ambiente")
def test_ofx_to_excel_from_example():
    path = OFX_EXAMPLES[0]
    statement = parse_ofx(path.read_bytes())
    assert len(statement.transactions) > 100
    assert {"Crédito", "Débito"}.issubset(set(statement.transactions["Tipo"]))
    payload = build_excel(statement, path.name)
    workbook = load_workbook(filename=__import__("io").BytesIO(payload), data_only=False)
    assert workbook.sheetnames == ["Lançamentos", "Resumo", "Como conferir"]
    assert workbook["Lançamentos"]["A8"].value == "Não"


@pytest.mark.skipif(not PDF_EXAMPLES, reason="Exemplo PDF privado não está disponível neste ambiente")
def test_pdf_to_ofx_from_example():
    path = PDF_EXAMPLES[0]
    statement = parse_pdf(path.read_bytes())
    assert len(statement.transactions) > 100
    assert statement.metadata["Banco"] == "237"
    assert statement.metadata["Agência"] == "480"
    assert statement.metadata["Conta"] == "151385-0"
    assert "BIG SUPERMERCADOS" in " ".join(statement.transactions["Descrição"])
    payload = build_ofx(statement.transactions, statement.metadata)
    reparsed = parse_ofx(payload)
    assert len(reparsed.transactions) == len(statement.transactions)
    assert reparsed.metadata["Saldo"] == "176.70"
