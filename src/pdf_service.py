from __future__ import annotations

import hashlib
import html
import re
from dataclasses import dataclass
from datetime import datetime
from io import BytesIO
from typing import Any

import pymupdf as fitz
import pandas as pd


DATE_RE = re.compile(r"^\d{2}/\d{2}/\d{4}$")
MONEY_RE = re.compile(r"^-?(?:\d{1,3}(?:\.\d{3})*|\d+),\d{2}$")


@dataclass(frozen=True)
class PdfStatement:
    metadata: dict[str, str]
    transactions: pd.DataFrame
    warnings: list[str]


def _br_money(value: str) -> float:
    return float(value.replace(".", "").replace(",", "."))


def _group_lines(words: list[tuple[Any, ...]], tolerance: float = 1.6) -> list[dict[str, Any]]:
    grouped: list[dict[str, Any]] = []
    for word in sorted(words, key=lambda item: (item[1], item[0])):
        y = float(word[1])
        if not grouped or abs(grouped[-1]["y"] - y) > tolerance:
            grouped.append({"y": y, "words": [word]})
        else:
            grouped[-1]["words"].append(word)
            grouped[-1]["y"] = sum(float(w[1]) for w in grouped[-1]["words"]) / len(grouped[-1]["words"])
    for line in grouped:
        line["words"].sort(key=lambda item: item[0])
    return grouped


def _column_text(line: dict[str, Any], left: float, right: float) -> str:
    return " ".join(str(w[4]) for w in line["words"] if left <= float(w[0]) < right).strip()


def _core_transaction(line: dict[str, Any]) -> dict[str, Any] | None:
    date = next((str(w[4]) for w in line["words"] if float(w[0]) < 120 and DATE_RE.match(str(w[4]))), None)
    type_word = next((str(w[4]) for w in line["words"] if 385 <= float(w[0]) < 435 and str(w[4]).lower() in {"crédito", "credito", "débito", "debito"}), None)
    money_words = [(float(w[0]), str(w[4])) for w in line["words"] if MONEY_RE.match(str(w[4]))]
    if not date or not type_word or len(money_words) < 2:
        return None
    money_words.sort()
    amount = _br_money(money_words[-2][1])
    balance = _br_money(money_words[-1][1])
    if type_word.lower() in {"débito", "debito"} and amount > 0:
        amount *= -1
    document = _column_text(line, 335, 385)
    return {
        "Data": pd.to_datetime(date, format="%d/%m/%Y"),
        "Tipo": "Crédito" if amount >= 0 else "Débito",
        "Valor": amount,
        "Saldo": balance,
        "Documento": document,
        "Descrição": "",
    }


def _description(line: dict[str, Any]) -> str:
    return _column_text(line, 120, 335)


def _extract_metadata(text: str) -> dict[str, str]:
    company = re.search(r"Empresa\s*\|\s*CNPJ:\s*(.*?)\s*\|\s*([\d./-]+)", text, re.I)
    account = re.search(r"Agência\s*\|\s*Conta:\s*([\d]+)\s*\|\s*([\d-]+)", text, re.I)
    period = re.search(r"Período:\s*(\d{2}/\d{2}/\d{4})\s*a\s*(\d{2}/\d{2}/\d{4})", text, re.I)
    return {
        "Banco": "237",
        "Instituição": "Bradesco",
        "Titular": company.group(1).strip() if company else "",
        "CNPJ": company.group(2).strip() if company else "",
        "Agência": account.group(1).strip() if account else "",
        "Conta": account.group(2).strip() if account else "",
        "Início": period.group(1) if period else "",
        "Fim": period.group(2) if period else "",
        "Moeda": "BRL",
    }


def parse_pdf(data: bytes) -> PdfStatement:
    try:
        document = fitz.open(stream=data, filetype="pdf")
    except Exception as exc:
        raise ValueError("Não foi possível abrir o PDF. Verifique se ele não está protegido ou corrompido.") from exc

    full_text = "\n".join(page.get_text() for page in document)
    if len(full_text.strip()) < 80:
        raise ValueError("O PDF parece ser escaneado e não contém texto selecionável. Esta versão não executa OCR.")
    if "Saldo (R$)" not in full_text or not re.search(r"Crédito|Débito", full_text, re.I):
        raise ValueError("O layout deste PDF ainda não é compatível. Use um extrato Bradesco Empresas no mesmo padrão do exemplo.")

    transactions: list[dict[str, Any]] = []
    pending_continuation: list[str] = []
    for page_index, page in enumerate(document):
        lines = _group_lines(page.get_text("words"))
        cores = [(idx, _core_transaction(line)) for idx, line in enumerate(lines)]
        cores = [(idx, core) for idx, core in cores if core]
        if not cores:
            continue

        first_core_idx = cores[0][0]
        first_core_y = lines[first_core_idx]["y"]
        if page_index > 0 and transactions:
            for line in lines[:first_core_idx]:
                desc = _description(line)
                if desc and line["y"] < first_core_y - 7 and line["y"] < 48:
                    pending_continuation.append(desc)
            if pending_continuation:
                transactions[-1]["Descrição"] = " ".join([transactions[-1]["Descrição"], *pending_continuation]).strip()
                pending_continuation.clear()

        for core_idx, core in cores:
            core_y = lines[core_idx]["y"]
            description_parts: list[str] = []
            for nearby in lines:
                if abs(nearby["y"] - core_y) <= 6.6:
                    desc = _description(nearby)
                    if desc and desc not in description_parts:
                        description_parts.append(desc)
            core["Descrição"] = " ".join(description_parts).strip()
            transactions.append(core)

    document.close()
    if not transactions:
        raise ValueError("Nenhum lançamento bancário foi identificado no PDF.")

    frame = pd.DataFrame(transactions)
    frame["Data"] = pd.to_datetime(frame["Data"])
    warnings: list[str] = []
    empty_descriptions = int((frame["Descrição"].str.strip() == "").sum())
    if empty_descriptions:
        warnings.append(f"{empty_descriptions} lançamento(s) ficaram sem descrição e devem ser revisados.")
    return PdfStatement(metadata=_extract_metadata(full_text), transactions=frame, warnings=warnings)


def _clean(value: Any, limit: int = 255) -> str:
    text = " ".join(str(value or "").split())[:limit]
    return html.escape(text, quote=False)


def _fitid(row: pd.Series) -> str:
    raw = "|".join(
        [
            pd.Timestamp(row["Data"]).strftime("%Y%m%d"),
            str(row.get("Documento", "")),
            f"{float(row['Valor']):.2f}",
            str(row.get("Descrição", "")),
        ]
    )
    return hashlib.sha1(raw.encode("utf-8")).hexdigest()[:24].upper()


def build_ofx(transactions: pd.DataFrame, metadata: dict[str, str]) -> bytes:
    if transactions.empty:
        raise ValueError("É necessário ter pelo menos um lançamento para gerar o OFX.")
    frame = transactions.copy()
    frame["Data"] = pd.to_datetime(frame["Data"], errors="coerce")
    frame["Valor"] = pd.to_numeric(frame["Valor"], errors="coerce")
    if frame["Data"].isna().any() or frame["Valor"].isna().any():
        raise ValueError("Existem datas ou valores inválidos na tabela editada.")

    start = frame["Data"].min().strftime("%Y%m%d000000")
    end = frame["Data"].max().strftime("%Y%m%d235959")
    now = datetime.now().strftime("%Y%m%d%H%M%S")
    latest_date = frame["Data"].max()
    # O extrato vem com datas decrescentes, mas mantém a sequência dos
    # lançamentos dentro de cada dia. O último item do dia mais recente traz
    # o saldo de fechamento correto.
    latest = frame.loc[frame["Data"] == latest_date].iloc[-1]
    ledger = float(latest.get("Saldo", 0) or 0)

    lines = [
        "OFXHEADER:100",
        "DATA:OFXSGML",
        "VERSION:102",
        "SECURITY:NONE",
        "ENCODING:USASCII",
        "CHARSET:1252",
        "COMPRESSION:NONE",
        "OLDFILEUID:NONE",
        "NEWFILEUID:NONE",
        "",
        "<OFX>",
        "  <SIGNONMSGSRSV1><SONRS><STATUS><CODE>0</CODE><SEVERITY>INFO</SEVERITY></STATUS>",
        f"  <DTSERVER>{now}</DTSERVER><LANGUAGE>POR</LANGUAGE></SONRS></SIGNONMSGSRSV1>",
        "  <BANKMSGSRSV1><STMTTRNRS><TRNUID>0</TRNUID><STATUS><CODE>0</CODE><SEVERITY>INFO</SEVERITY></STATUS><STMTRS>",
        f"  <CURDEF>{_clean(metadata.get('Moeda', 'BRL'))}</CURDEF>",
        "  <BANKACCTFROM>",
        f"    <BANKID>{_clean(metadata.get('Banco', '237'))}</BANKID>",
        f"    <BRANCHID>{_clean(metadata.get('Agência', ''))}</BRANCHID>",
        f"    <ACCTID>{_clean(metadata.get('Conta', ''))}</ACCTID>",
        "    <ACCTTYPE>CHECKING</ACCTTYPE>",
        "  </BANKACCTFROM>",
        f"  <BANKTRANLIST><DTSTART>{start}</DTSTART><DTEND>{end}</DTEND>",
    ]
    for _, row in frame.sort_values("Data", ascending=False, kind="stable").iterrows():
        amount = float(row["Valor"])
        description = _clean(row.get("Descrição", ""))
        lines.extend(
            [
                "    <STMTTRN>",
                f"      <TRNTYPE>{'CREDIT' if amount >= 0 else 'DEBIT'}</TRNTYPE>",
                f"      <DTPOSTED>{pd.Timestamp(row['Data']).strftime('%Y%m%d')}120000</DTPOSTED>",
                f"      <TRNAMT>{amount:.2f}</TRNAMT>",
                f"      <FITID>{_fitid(row)}</FITID>",
                f"      <CHECKNUM>{_clean(row.get('Documento', ''))}</CHECKNUM>",
                f"      <NAME>{description[:32]}</NAME>",
                f"      <MEMO>{description}</MEMO>",
                "    </STMTTRN>",
            ]
        )
    lines.extend(
        [
            "  </BANKTRANLIST>",
            f"  <LEDGERBAL><BALAMT>{ledger:.2f}</BALAMT><DTASOF>{end}</DTASOF></LEDGERBAL>",
            "  </STMTRS></STMTTRNRS></BANKMSGSRSV1>",
            "</OFX>",
        ]
    )
    return "\r\n".join(lines).encode("cp1252", errors="replace")
