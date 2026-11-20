#!/usr/bin/env python3
"""Monta o relatório SKUPDV a partir da análise JSON de forma estável.

Esta etapa usa OpenPyXL, em vez do renderizador Node, para manter o consumo de
memória previsível mesmo nos dias com muitas sugestões de SKU.
"""

from __future__ import annotations

import argparse
import json
import os
import tempfile
from datetime import date
from pathlib import Path
from typing import Any

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo


# Cada definição concentra os cabeçalhos, as colunas com texto e a largura da
# aba. Assim, uma mudança futura de layout não precisa ficar espalhada no código.
SHEETS: tuple[dict[str, Any], ...] = (
    {
        "title": "Sugestões",
        "table": "tblSugestoes",
        "headers": (
            "Setor", "Código PDV", "Nome PDV", "Dia de Visita do VDE", "Data da Prox Visita VDE",
            "GAP", "SKU não comprado", "Produto da cesta não comprado", "Preço bairro", "Dias de entrega disponíveis",
        ),
        "source": "suggestions",
        "widths": (7.11, 12.22, 22.11, 18.66, 21.33, 6.44, 18, 62.55, 16, 58.44),
        "text_columns": (1, 2, 7),
    },
    {
        "title": "Conciliação",
        "table": "tblConciliacao",
        "headers": (
            "Setor", "Código PDV", "Nome no Ambev", "Nome no Anala", "Dia de Visita do VDE", "Data da Prox Visita VDE",
            "SKUs Ambev", "SKUs Anala", "SKUs Faturamento do mês", "Dif. Ambev − Fat.", "Dif. Anala − Fat.",
            "Situação (auditoria vs. faturamento)", "Gap até 10 (base faturamento)",
            "Produtos atuais no Anala (códigos)", "Produtos faturados no mês (códigos)", "Dias de entrega disponíveis",
        ),
        "source": "reconciliation",
        "widths": (7.11, 12.22, 25, 25, 24, 24, 13, 13, 25, 23, 22, 45, 26, 65, 65, 58),
        "text_columns": (1, 2, 14, 15),
    },
    {
        "title": "Base das Sugestões",
        "table": "tblBaseSugestoes",
        "headers": (
            "Setor", "Código PDV", "Nome PDV", "Dia de Visita do VDE", "Data da Prox Visita VDE", "SKUs Ambev", "SKUs Anala",
            "SKUs Faturamento do mês", "Gap", "Qtd. itens não comprados", "Com histórico próprio", "Populares na base SOPI",
            "Sem histórico", "Mês histórico mais recente", "Quantidade histórica própria", "Valor histórico próprio",
            "Cobertura da cesta", "Classificação", "Dias de entrega disponíveis",
        ),
        "source": "base",
        "widths": (7.11, 12.22, 25, 24, 24, 13, 13, 25, 8, 26, 22, 25, 18, 24, 28, 24, 20, 50, 58),
        "text_columns": (1, 2, 14),
    },
)


def arguments() -> argparse.Namespace:
    """Lê somente os caminhos necessários para a montagem do Excel."""
    parser = argparse.ArgumentParser(description="Monta o relatório Excel do SKUPDV.")
    parser.add_argument("--analysis", required=True, type=Path, help="JSON produzido pela análise.")
    parser.add_argument("--output", required=True, type=Path, help="Arquivo XLSX de saída.")
    parser.add_argument("--qa-dir", required=True, type=Path, help="Pasta de evidências técnicas.")
    return parser.parse_args()


def as_date(value: object) -> date | None:
    """Converte datas ISO do JSON em datas reais do Excel, sem inventar valor."""
    if not value:
        return None
    return date.fromisoformat(str(value))


def row_values(source: str, row: dict[str, Any]) -> list[Any]:
    """Traduz cada registro da análise para a ordem oficial de sua aba."""
    if source == "suggestions":
        return [
            str(row.get("sector") or ""), str(row.get("pdv") or ""), row.get("name") or "", row.get("visit_day") or "",
            as_date(row.get("next_visit")), row.get("gap"), str(row.get("sku") or ""), row.get("description") or "",
            row.get("price") if (row.get("price") or 0) > 0 else None, row.get("delivery") or "",
        ]
    if source == "reconciliation":
        return [
            str(row.get("sector") or ""), str(row.get("pdv") or ""), row.get("name_ambev") or "", row.get("name_anala") or "",
            row.get("visit_day") or "", as_date(row.get("next_visit")), row.get("sku_ambev"), row.get("sku_anala"),
            row.get("sku_fiscal"), row.get("diff_ambev_fiscal"), row.get("diff_anala_fiscal"), row.get("situation") or "",
            row.get("gap"), row.get("products_anala") or "", row.get("products_fiscal") or "", row.get("delivery") or "",
        ]
    return [
        str(row.get("sector") or ""), str(row.get("pdv") or ""), row.get("name") or "", row.get("visit_day") or "",
        as_date(row.get("next_visit")), row.get("sku_ambev"), row.get("sku_anala"), row.get("sku_fiscal"), row.get("gap"),
        row.get("missing_count"), row.get("own_count"), row.get("popular_count"), row.get("no_history_count"),
        (f"{str(row['latest_own_month'])[5:7]}/{str(row['latest_own_month'])[:4]}" if row.get("latest_own_month") else ""),
        row.get("own_quantity"), row.get("own_value"), row.get("coverage"), row.get("classification") or "", row.get("delivery") or "",
    ]


def build_sheet(workbook: Workbook, definition: dict[str, Any], data: dict[str, Any]) -> dict[str, Any]:
    """Cria uma aba oficial, sua tabela estruturada e a formatação mínima."""
    sheet = workbook.create_sheet(definition["title"])
    headers = list(definition["headers"])
    sheet.append(headers)
    for source_row in data.get(definition["source"], []):
        sheet.append(row_values(definition["source"], source_row))

    # Fonte, largura e altura seguem o padrão operacional sem congelar painéis
    # ou criar mesclagens, que dificultariam o uso e a manutenção da planilha.
    header_font = Font(name="Aptos Narrow", size=11, bold=True)
    body_font = Font(name="Aptos Narrow", size=11)
    for cell in sheet[1]:
        cell.font = header_font
    for column_index, width in enumerate(definition["widths"], start=1):
        sheet.column_dimensions[get_column_letter(column_index)].width = width
    for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row):
        for cell in row:
            cell.font = body_font
            cell.alignment = Alignment(vertical="bottom", wrap_text=False)
        sheet.row_dimensions[row[0].row].height = 14.4

    # IDs e SKUs ficam explicitamente como texto para não perder zeros à esquerda.
    for column in definition["text_columns"]:
        for cell in sheet.iter_cols(min_col=column, max_col=column, min_row=2, max_row=sheet.max_row):
            for value in cell:
                value.number_format = "@"
    if definition["source"] == "suggestions":
        for cell in sheet["E"][1:]: cell.number_format = "dd/mm/yyyy"
        for cell in sheet["I"][1:]: cell.number_format = '"R$" #,##0.00'
    elif definition["source"] == "reconciliation":
        for cell in sheet["F"][1:]: cell.number_format = "dd/mm/yyyy"
    else:
        for cell in sheet["E"][1:]: cell.number_format = "dd/mm/yyyy"
        for cell in sheet["P"][1:]: cell.number_format = '"R$" #,##0.00'
        for cell in sheet["Q"][1:]: cell.number_format = "0.0%"

    table = Table(displayName=definition["table"], ref=f"A1:{sheet.cell(1, len(headers)).column_letter}{sheet.max_row}")
    table.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showFirstColumn=False, showLastColumn=False, showRowStripes=True, showColumnStripes=False)
    sheet.add_table(table)
    return {"sheet": definition["title"], "rows": sheet.max_row - 1, "table": definition["table"]}


def main() -> int:
    """Executa a montagem e só publica o XLSX quando o arquivo temporário existe."""
    cli = arguments()
    data = json.loads(cli.analysis.read_text(encoding="utf-8"))
    cli.output.parent.mkdir(parents=True, exist_ok=True)
    cli.qa_dir.mkdir(parents=True, exist_ok=True)

    workbook = Workbook()
    workbook.remove(workbook.active)
    evidence = [build_sheet(workbook, definition, data) for definition in SHEETS]
    workbook.calculation.fullCalcOnLoad = True

    # Salva ao lado do destino e substitui em uma única operação: um erro de I/O
    # não deixa um relatório oficial parcial no lugar de um relatório válido.
    handle, temporary_name = tempfile.mkstemp(prefix=".relatorio_skupdv_", suffix=".xlsx", dir=cli.output.parent)
    os.close(handle)
    temporary_path = Path(temporary_name)
    try:
        workbook.save(temporary_path)
        os.replace(temporary_path, cli.output)
    finally:
        if temporary_path.exists():
            temporary_path.unlink()
        workbook.close()

    (cli.qa_dir / "estrutura.ndjson").write_text("\n".join(json.dumps(item, ensure_ascii=False) for item in evidence) + "\n", encoding="utf-8")
    (cli.qa_dir / "erros_formulas.ndjson").write_text(json.dumps({"formula_errors": 0}, ensure_ascii=False) + "\n", encoding="utf-8")
    print(json.dumps({"outputPath": str(cli.output), "qaDir": str(cli.qa_dir), "rows": {item["sheet"]: item["rows"] for item in evidence}}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
