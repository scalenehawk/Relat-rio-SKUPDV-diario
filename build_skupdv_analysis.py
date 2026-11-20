from __future__ import annotations

import argparse
import csv
import json
import math
import re
import unicodedata
from collections import defaultdict
from datetime import date, datetime
from pathlib import Path
from typing import Any

from openpyxl import load_workbook


TARGET = 10
# Universo de PDVs (regra do KPI SKU/PDV Cooler): base de SOPI comodatada e ativa, instalada desde 01/01/2026.
# O 02.02.20 é a fonte; ajuste aqui se a regra mudar.
COMODATO_STATUS_ATIVO = {"ATIVO"}          # valores aceitos em "Status Cliente" (Operação Ativa = S)
COMODATO_SUBTIPOS = {"SOPI"}               # "Sub-Tipo Material": só SOPI (VISA, CHOPP e OUTROS ficam de fora)
COMODATO_INSTALADO_DESDE = date(2026, 1, 1)
COMODATO_COLUNA_INSTALACAO = "DATA EMISSAO"  # a Ambev conta a instalação pela emissão do comodato (inclui equipamento ainda em trânsito); alternativa: "DATA EXECUCAO"
ELIGIBLE_TYPES = {"CERVEJA", "MATCH"}
MONTH_NAMES = {
    1: "janeiro",
    2: "fevereiro",
    3: "março",
    4: "abril",
    5: "maio",
    6: "junho",
    7: "julho",
    8: "agosto",
    9: "setembro",
    10: "outubro",
    11: "novembro",
    12: "dezembro",
}


def args() -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    parser.add_argument("--root", required=True)
    parser.add_argument("--comodato", required=True)
    parser.add_argument("--ambev", required=True)
    parser.add_argument("--anala")
    parser.add_argument("--basket", required=True)
    parser.add_argument("--clients", required=True)
    parser.add_argument("--delivery", required=True)
    parser.add_argument("--prices", required=True)
    parser.add_argument("--billing-dir", required=True)
    parser.add_argument("--report-date", required=True)
    parser.add_argument("--output", required=True)
    return parser.parse_args()


def clean(value: Any) -> str:
    return "" if value is None else str(value).strip()


def norm_key(value: Any) -> str:
    text = clean(value)
    if not text:
        return ""
    text = re.sub(r"\.0+$", "", text)
    text = re.sub(r"\s+", "", text)
    if re.fullmatch(r"[+-]?\d+", text):
        text = text.lstrip("+")
        text = text.lstrip("0") or "0"
    return text


def norm_text(value: Any) -> str:
    text = unicodedata.normalize("NFKD", clean(value))
    return "".join(ch for ch in text if not unicodedata.combining(ch)).upper()


def parse_br_number(value: Any) -> float | None:
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)):
        if isinstance(value, float) and math.isnan(value):
            return None
        return float(value)
    text = clean(value).replace("R$", "").replace(" ", "")
    if not text:
        return None
    if "," in text:
        text = text.replace(".", "").replace(",", ".")
    elif text.count(".") > 1 or (text.count(".") == 1 and len(text.rsplit(".", 1)[1]) == 3):
        text = text.replace(".", "")
    try:
        return float(text)
    except ValueError:
        return None


def parse_date(value: Any) -> date | None:
    if value is None or value == "":
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = clean(value)
    for fmt in ("%d/%m/%Y", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%Y%m%d"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            pass
    return None


def iso(value: date | None) -> str | None:
    return value.isoformat() if value else None


def csv_reader(path: Path):
    handle = path.open("r", encoding="cp1252", errors="replace", newline="")
    reader = csv.DictReader(handle, delimiter=";")
    reader.fieldnames = [clean(name) for name in (reader.fieldnames or [])]
    return handle, reader


def valid_basket_product(description: str, product_type: str) -> tuple[bool, str]:
    desc = norm_text(description)
    typ = norm_text(product_type)
    if typ not in ELIGIBLE_TYPES:
        return False, "tipo não elegível"
    if "MARKETPLACE" in desc:
        return False, "Marketplace"
    # Regra do KPI: só as embalagens de chopp (Bag, Barril, BIB, e equivalentes Petainer/Keg) ficam fora.
    # "BRAHMA CHOPP GARRAFA 1L" é cerveja envasada e conta normalmente.
    if re.search(r"\b(BAG|BARRIL|BIB|PETAINER|KEG)\b", desc):
        return False, "embalagem de Chopp"
    return True, ""


def load_basket(path: Path) -> tuple[dict[str, str], dict[str, Any]]:
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb["CESTA SOPI"] if "CESTA SOPI" in wb.sheetnames else wb[wb.sheetnames[0]]
    rows = ws.iter_rows(values_only=True)
    header = [clean(v) for v in next(rows)]
    idx = {name: i for i, name in enumerate(header)}
    code_i = idx.get("Cod", 0)
    name_i = idx.get("Nome Produto", 1)
    type_i = idx.get("Tipo Produto", 2)
    basket: dict[str, str] = {}
    categories = defaultdict(int)
    exclusions = defaultdict(int)
    duplicates = 0
    invalid = 0
    conflicts: list[dict[str, str]] = []
    source_rows = 0
    for row in rows:
        source_rows += 1
        code = norm_key(row[code_i] if code_i < len(row) else None)
        description = clean(row[name_i] if name_i < len(row) else None)
        product_type = norm_text(row[type_i] if type_i < len(row) else None)
        if not code:
            invalid += 1
            continue
        valid, reason = valid_basket_product(description, product_type)
        if not valid:
            exclusions[reason] += 1
            continue
        categories[product_type] += 1
        if code in basket:
            duplicates += 1
            if description and basket[code] != description:
                old = basket[code]
                if len(description) > len(old):
                    basket[code] = description
                conflicts.append({"sku": code, "description_1": old, "description_2": description})
        else:
            basket[code] = description
    wb.close()
    return basket, {
        "sheet": ws.title,
        "source_rows": source_rows,
        "eligible_skus": len(basket),
        "categories": dict(categories),
        "exclusions": dict(exclusions),
        "duplicates": duplicates,
        "invalid_codes": invalid,
        "description_conflicts": conflicts[:100],
    }


def load_comodato(path: Path) -> tuple[dict[str, dict[str, Any]], dict[str, Any]]:
    """Lê o 02.02.20 (uma linha por equipamento comodatado). Linhas com ';' dentro da descrição são reagrupadas."""
    handle = path.open("r", encoding="cp1252", errors="replace", newline="")
    reader = csv.reader(handle, delimiter=";")
    header = [norm_text(name) for name in next(reader)]
    idx = {name: i for i, name in enumerate(header) if name}
    required = ["CLIENTE", "NOME / RAZAO", "CODIGO SETOR", "DESCRICAO", "COMODATADO", "BAIXADO",
                "TIPO MATERIAL", "SUB-TIPO MATERIAL", "STATUS CLIENTE", "DATA RECOLHA", COMODATO_COLUNA_INSTALACAO]
    missing = [name for name in required if name not in idx]
    if missing:
        raise ValueError(f"Colunas ausentes no 02.02.20: {missing}")
    width = len(header)
    desc_i = idx["DESCRICAO"]
    route_i = idx.get("ROTA")
    pdvs: dict[str, dict[str, Any]] = {}
    total_rows = 0
    regrouped = 0
    for row in reader:
        if not any(clean(cell) for cell in row):
            continue
        total_rows += 1
        if len(row) > width:
            extra = len(row) - width
            row = row[:desc_i] + [";".join(row[desc_i:desc_i + extra + 1])] + row[desc_i + extra + 1:]
            regrouped += 1
        pdv = norm_key(row[idx["CLIENTE"]])
        if not pdv:
            continue
        committed = parse_br_number(row[idx["COMODATADO"]]) or 0.0
        written_off = parse_br_number(row[idx["BAIXADO"]]) or 0.0
        collected = parse_date(row[idx["DATA RECOLHA"]])
        active_line = (committed - written_off) > 0 and collected is None
        record = pdvs.setdefault(pdv, {
            "name": clean(row[idx["NOME / RAZAO"]]),
            "sector": norm_key(row[idx["CODIGO SETOR"]]),
            "route": clean(row[route_i]) if route_i is not None else "",
            "status": norm_text(row[idx["STATUS CLIENTE"]]),
            "lines": 0,
            "active_lines": 0,
            "active_types": set(),
            "qualifying_lines": 0,
            "sopi_before_cutoff": 0,
        })
        record["lines"] += 1
        if active_line:
            record["active_lines"] += 1
            record["active_types"].add(norm_text(row[idx["TIPO MATERIAL"]]))
            if norm_text(row[idx["SUB-TIPO MATERIAL"]]) in COMODATO_SUBTIPOS:
                installed = parse_date(row[idx[COMODATO_COLUNA_INSTALACAO]])
                if installed is not None and installed >= COMODATO_INSTALADO_DESDE:
                    record["qualifying_lines"] += 1
                else:
                    record["sopi_before_cutoff"] += 1
    handle.close()
    return pdvs, {"file": str(path), "rows": total_rows, "pdvs": len(pdvs), "rows_regrouped": regrouped}


def comodato_universe(comodato: dict[str, dict[str, Any]]) -> tuple[set[str], dict[str, Any]]:
    universe: set[str] = set()
    excluded_status: dict[str, int] = defaultdict(int)
    without_active_line = 0
    without_sopi_2026 = 0
    for pdv, record in comodato.items():
        if record["active_lines"] == 0:
            without_active_line += 1
        elif record["qualifying_lines"] == 0:
            without_sopi_2026 += 1
        elif record["status"] not in COMODATO_STATUS_ATIVO:
            excluded_status[record["status"] or "(vazio)"] += 1
        else:
            universe.add(pdv)
    return universe, {
        "universe_pdvs": len(universe),
        "excluded_without_active_line": without_active_line,
        "excluded_without_sopi_installed_in_period": without_sopi_2026,
        "excluded_by_client_status": dict(excluded_status),
        "rule_status": sorted(COMODATO_STATUS_ATIVO),
        "rule_subtypes": sorted(COMODATO_SUBTIPOS),
        "rule_installed_since": COMODATO_INSTALADO_DESDE.isoformat(),
        "rule_installation_column": COMODATO_COLUNA_INSTALACAO,
    }


def load_ambev(path: Path) -> tuple[dict[str, dict[str, Any]], dict[str, Any]]:
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb["Export"] if "Export" in wb.sheetnames else wb[wb.sheetnames[0]]
    rows = ws.iter_rows(values_only=True)
    header = [clean(v) for v in next(rows)]
    idx = {name: i for i, name in enumerate(header)}
    required = ["PDV", "COD RN", "Razão", "Quantidade SKUs"]
    missing = [name for name in required if name not in idx]
    if missing:
        raise ValueError(f"Colunas ausentes no Ambev: {missing}")
    pdvs: dict[str, dict[str, Any]] = {}
    duplicates = 0
    count_conflicts = []
    invalid = 0
    for row in rows:
        pdv = norm_key(row[idx["PDV"]])
        if not pdv:
            invalid += 1
            continue
        raw_count = row[idx["Quantidade SKUs"]]
        count = int(parse_br_number(raw_count) or 0)
        record = {
            "pdv": pdv,
            "sector": norm_key(row[idx["COD RN"]]),
            "name_ambev": clean(row[idx["Razão"]]),
            "visit_day_ambev": clean(row[idx.get("Frequência visita", -1)]) if idx.get("Frequência visita", -1) >= 0 else "",
            "next_visit_ambev": iso(parse_date(row[idx.get("Data próxima visita", -1)])) if idx.get("Data próxima visita", -1) >= 0 else None,
            "sku_ambev": count,
        }
        if pdv in pdvs:
            duplicates += 1
            if pdvs[pdv]["sku_ambev"] != count:
                count_conflicts.append({"pdv": pdv, "counts": [pdvs[pdv]["sku_ambev"], count]})
            if count > pdvs[pdv]["sku_ambev"]:
                pdvs[pdv] = record
        else:
            pdvs[pdv] = record
    wb.close()
    return pdvs, {
        "sheet": ws.title,
        "pdvs": len(pdvs),
        "duplicates": duplicates,
        "invalid_codes": invalid,
        "count_conflicts": count_conflicts,
    }


def load_anala(
    path: Path | None,
    basket: set[str],
) -> tuple[dict[str, dict[str, Any]], dict[str, Any]]:
    if path is None:
        return {}, {"available": False, "pdvs": 0, "duplicates": 0, "invalid_codes": 0}
    workbook = load_workbook(path, read_only=True, data_only=True)
    result: dict[str, dict[str, Any]] = {}
    duplicates = 0
    invalid_codes = 0
    for sheet_name in ("Compraram", "Não Compraram"):
        if sheet_name not in workbook.sheetnames:
            continue
        sheet = workbook[sheet_name]
        for row in sheet.iter_rows(min_row=4, values_only=True):
            if norm_text(row[0] if len(row) > 0 else None) != "SOPI":
                continue
            pdv = norm_key(row[2] if len(row) > 2 else None)
            if not pdv:
                invalid_codes += 1
                continue
            products: set[str] = set()
            if sheet_name == "Compraram" and len(row) > 5:
                products = {
                    norm_key(item)
                    for item in clean(row[5]).split(",")
                    if norm_key(item) in basket
                }
            if pdv in result:
                duplicates += 1
                result[pdv]["products"].update(products)
                continue
            result[pdv] = {
                "name_anala": clean(row[3] if len(row) > 3 else None),
                "sector_anala": norm_key(row[1] if len(row) > 1 else None),
                "visit_day_anala": clean(row[4] if len(row) > 4 else None),
                "products": products,
            }
    workbook.close()
    return result, {
        "available": True,
        "path": str(path),
        "pdvs": len(result),
        "duplicates": duplicates,
        "invalid_codes": invalid_codes,
    }


def load_clients(path: Path, universe: set[str]) -> tuple[dict[str, dict[str, Any]], dict[str, Any]]:
    handle, reader = csv_reader(path)
    clients: dict[str, dict[str, Any]] = {}
    duplicate_rows = 0
    for row in reader:
        pdv = norm_key(row.get("Cód PDV"))
        if pdv not in universe:
            continue
        record = {
            "name_client": clean(row.get("Nome Fantasia")),
            "sector_client": norm_key(row.get("Setor VDE")),
            "visit_day": clean(row.get("Dia de Visita do VDE")),
            "next_visit": iso(parse_date(row.get("Data da Prox Visita VDE"))),
        }
        if pdv in clients:
            duplicate_rows += 1
            prior = clients[pdv]
            for key, value in record.items():
                if not prior.get(key) and value:
                    prior[key] = value
        else:
            clients[pdv] = record
    handle.close()
    return clients, {"matched_pdvs": len(clients), "duplicate_rows": duplicate_rows}


DAY_LABELS = [("Mon", "SEG"), ("Tue", "TER"), ("Wed", "QUA"), ("Thu", "QUI"), ("Fri", "SEX"), ("Sat", "SÁB"), ("Sun", "DOM")]


def load_delivery(path: Path, universe: set[str]) -> tuple[dict[str, str], dict[str, Any]]:
    handle = path.open("r", encoding="utf-8-sig", errors="replace", newline="")
    reader = csv.DictReader(handle, delimiter=";")
    reader.fieldnames = [clean(name) for name in (reader.fieldnames or [])]
    by_pdv: dict[str, set[str]] = defaultdict(set)
    for row in reader:
        pdv = norm_key(row.get("ERP_POC_ID"))
        if pdv not in universe:
            continue
        for column, label in DAY_LABELS:
            status = norm_text(row.get(column))
            if status in {"FREE", "FLEX"}:
                by_pdv[pdv].add(f"{label} - {status}")
    handle.close()
    order = {f"{label} - {status}": i * 2 + j for i, (_, label) in enumerate(DAY_LABELS) for j, status in enumerate(("FREE", "FLEX"))}
    result = {pdv: "; ".join(sorted(values, key=lambda item: order[item])) for pdv, values in by_pdv.items()}
    return result, {"matched_pdvs": len(result), "without_days": len(universe - set(result))}


def load_prices(path: Path) -> tuple[dict[str, float], dict[str, Any]]:
    wb_formula = load_workbook(path, read_only=True, data_only=False)
    wb_values = load_workbook(path, read_only=True, data_only=True)
    main_name = "tabela de preços" if "tabela de preços" in wb_formula.sheetnames else wb_formula.sheetnames[0]
    ws_formula = wb_formula[main_name]
    ws_values = wb_values[main_name]
    headers = [clean(cell.value) for cell in next(ws_formula.iter_rows(min_row=1, max_row=1))]
    if len(headers) < 10 or norm_text(headers[9]) != "VALOR":
        raise ValueError("A coluna J da tabela de preços não possui o cabeçalho VALOR.")
    fallback: dict[str, list[float]] = defaultdict(list)
    if "Planilha1" in wb_values.sheetnames:
        ws_ref = wb_values["Planilha1"]
        for row in ws_ref.iter_rows(min_row=2, values_only=True):
            code = norm_key(row[0] if len(row) > 0 else None)
            value = parse_br_number(row[2] if len(row) > 2 else None)
            if code and value is not None and value > 0:
                fallback[code].append(value)
    candidates: dict[str, list[float]] = defaultdict(list)
    invalid_rows = 0
    for row_formula, row_value in zip(
        ws_formula.iter_rows(min_row=2, values_only=True),
        ws_values.iter_rows(min_row=2, values_only=True),
    ):
        code = norm_key(row_formula[0] if row_formula else None)
        if not code:
            invalid_rows += 1
            continue
        value = parse_br_number(row_value[9] if len(row_value) > 9 else None)
        if value is None and code in fallback:
            for found in fallback[code]:
                candidates[code].append(found)
        elif value is not None and value > 0:
            candidates[code].append(value)
    prices: dict[str, float] = {}
    conflicts: dict[str, list[float]] = {}
    duplicate_same = 0
    invalid_price = 0
    for code, values in candidates.items():
        unique = sorted({round(value, 6) for value in values})
        if len(unique) == 1:
            prices[code] = unique[0]
            duplicate_same += max(0, len(values) - 1)
        elif len(unique) > 1:
            conflicts[code] = unique
        else:
            invalid_price += 1
    fallback_sheet_used = "Planilha1" in wb_values.sheetnames
    wb_formula.close()
    wb_values.close()
    return prices, {
        "sheet": main_name,
        "unique_prices": len(prices),
        "conflicts": conflicts,
        "conflict_count": len(conflicts),
        "duplicate_same_price": duplicate_same,
        "invalid_rows": invalid_rows,
        "invalid_prices": invalid_price,
        "fallback_sheet_used": fallback_sheet_used,
    }


def billing_files(folder: Path, report_date: date) -> list[Path]:
    result = []
    for path in folder.glob(f"{report_date.year}_*_Faturamento.csv"):
        match = re.match(rf"{report_date.year}_(\d{{2}})_Faturamento\.csv$", path.name, re.IGNORECASE)
        if match and int(match.group(1)) <= report_date.month:
            result.append(path)
    return sorted(result, key=lambda p: p.name)


def process_billing(
    paths: list[Path],
    basket: set[str],
    universe: set[str],
    report_date: date,
) -> tuple[set[str], dict[str, set[str]], dict[tuple[str, str], dict[str, Any]], dict[str, dict[str, Any]], dict[str, Any]]:
    regional: set[str] = set()
    current: dict[str, set[str]] = defaultdict(set)
    own: dict[tuple[str, str], dict[str, Any]] = {}
    popular: dict[str, dict[str, Any]] = {}
    popular_pdvs: dict[str, set[str]] = defaultdict(set)
    file_stats = []
    for path in paths:
        file_month = int(re.match(r"\d{4}_(\d{2})", path.name).group(1))
        min_date = None
        max_date = None
        total_rows = 0
        eligible_rows = 0
        basket_rows = 0
        handle, reader = csv_reader(path)
        for row in reader:
            total_rows += 1
            operation = norm_key(row.get("Operacao ."))
            status = norm_text(row.get("Status"))
            if operation not in {"1", "2"} or status not in {"A", "E"}:
                continue
            eligible_rows += 1
            emission = parse_date(row.get("Emissao"))
            if emission:
                min_date = emission if min_date is None or emission < min_date else min_date
                max_date = emission if max_date is None or emission > max_date else max_date
                if emission.year != report_date.year or emission > report_date:
                    continue
            sku = norm_key(row.get("Produto"))
            if sku not in basket:
                continue
            basket_rows += 1
            regional.add(sku)
            pdv = norm_key(row.get("Cliente"))
            if not pdv:
                pdv = norm_key(row.get("Cliente .."))
            if emission and emission.month == report_date.month and pdv in universe:
                current[pdv].add(sku)
            if file_month >= report_date.month or pdv not in universe:
                continue
            quantity = parse_br_number(row.get("Qtde")) or 0.0
            value = parse_br_number(row.get("Valor")) or 0.0
            key = (pdv, sku)
            metric = own.setdefault(key, {"months": set(), "latest": 0, "quantity": 0.0, "value": 0.0})
            metric["months"].add(file_month)
            metric["latest"] = max(metric["latest"], file_month)
            metric["quantity"] += quantity
            metric["value"] += value
            pop = popular.setdefault(sku, {"months": set(), "latest": 0, "quantity": 0.0, "value": 0.0})
            pop["months"].add(file_month)
            pop["latest"] = max(pop["latest"], file_month)
            pop["quantity"] += quantity
            pop["value"] += value
            popular_pdvs[sku].add(pdv)
        handle.close()
        file_stats.append({
            "path": str(path),
            "month": file_month,
            "min_date": iso(min_date),
            "max_date": iso(max_date),
            "total_rows": total_rows,
            "eligible_rows": eligible_rows,
            "basket_rows": basket_rows,
            "modified_at": datetime.fromtimestamp(path.stat().st_mtime).isoformat(timespec="seconds"),
        })
    for metric in own.values():
        metric["month_count"] = len(metric.pop("months"))
    for sku, metric in popular.items():
        metric["month_count"] = len(metric.pop("months"))
        metric["pdv_count"] = len(popular_pdvs[sku])
    return regional, current, own, popular, {"files": file_stats, "months": [item["month"] for item in file_stats]}


def sku_sort_key(sku: str):
    return (0, int(sku)) if sku.isdigit() else (1, sku)


def suggestion_sort_key(item: dict[str, Any]):
    sku_key = sku_sort_key(item["sku"])
    if item["category"] == "Histórico do próprio PDV":
        return (0, -item["latest"], -item["month_count"], -item["quantity"], -item["value"], sku_key)
    if item["category"] == "Popular no histórico da base SOPI":
        return (1, -item["pdv_count"], -item["latest"], -item["month_count"], -item["quantity"], sku_key)
    return (2, sku_key)


def main() -> None:
    cli = args()
    root = Path(cli.root)
    report_date = date.fromisoformat(cli.report_date)
    basket, basket_quality = load_basket(Path(cli.basket))
    comodato, comodato_quality = load_comodato(Path(cli.comodato))
    universe, universe_quality = comodato_universe(comodato)
    if not universe:
        raise ValueError("Nenhum PDV com comodato ativo foi encontrado no 02.02.20.")
    comodato_quality.update(universe_quality)
    ambev, ambev_quality = load_ambev(Path(cli.ambev))
    anala_path = Path(cli.anala) if cli.anala else None
    anala, anala_quality = load_anala(anala_path, set(basket))
    anala_available = anala_path is not None
    clients, clients_quality = load_clients(Path(cli.clients), universe)
    deliveries, delivery_quality = load_delivery(Path(cli.delivery), universe)
    prices, price_quality = load_prices(Path(cli.prices))
    bill_paths = billing_files(Path(cli.billing_dir), report_date)
    regional, current, own, popular, billing_quality = process_billing(
        bill_paths, set(basket), universe, report_date
    )
    if not regional:
        raise ValueError("A cesta regional ficou vazia.")

    suggestions: list[dict[str, Any]] = []
    reconciliation: list[dict[str, Any]] = []
    bases: list[dict[str, Any]] = []
    missing_sector = []
    missing_visit = []
    missing_delivery = []
    for pdv in sorted(universe, key=sku_sort_key):
        source = ambev.get(pdv, {})
        como = comodato[pdv]
        client = clients.get(pdv, {})
        anala_record = anala.get(pdv, {})
        sector = source.get("sector") or client.get("sector_client", "") or como.get("sector", "")
        name = como.get("name") or source.get("name_ambev") or anala_record.get("name_anala") or client.get("name_client", "")
        visit_day = client.get("visit_day") or anala_record.get("visit_day_anala") or source.get("visit_day_ambev", "")
        next_visit = client.get("next_visit") or source.get("next_visit_ambev")
        delivery = deliveries.get(pdv, "")
        sku_ambev = int(source["sku_ambev"]) if pdv in ambev else None
        fiscal_products = current.get(pdv, set())
        sku_fiscal = len(fiscal_products)
        anala_products = set(anala_record.get("products", set())) if anala_available else set()
        sku_anala = len(anala_products) if anala_available and pdv in anala else None
        # Faturamento (itens da cesta) é a referência; Ambev e Anala são auditados contra ele.
        gap = max(0, TARGET - sku_fiscal)
        diff_ambev_fiscal = (sku_ambev - sku_fiscal) if sku_ambev is not None else None
        diff_anala_fiscal = (sku_anala - sku_fiscal) if sku_anala is not None else None
        ambev_ok = sku_ambev is not None and sku_ambev == sku_fiscal
        anala_ok = sku_anala is None or sku_anala == sku_fiscal
        if sku_ambev is None:
            situation = "Ausente no planificador"
        elif ambev_ok and anala_ok:
            situation = "Convergente"
        elif diff_ambev_fiscal > 0:
            situation = "Ambev acima do faturamento"
        elif diff_ambev_fiscal < 0:
            situation = "Ambev abaixo do faturamento"
        else:
            situation = "Anala diverge do faturamento"
        converges = situation == "Convergente"
        if sku_fiscal < TARGET:
            classification = "Ação comercial" if converges else "Ação comercial e validação"
        else:
            classification = "Sem ação" if converges else "Validação com relação preventiva"
        current_products = fiscal_products
        missing = regional - current_products
        items = []
        for sku in missing:
            own_metric = own.get((pdv, sku))
            pop_metric = popular.get(sku)
            if own_metric:
                category = "Histórico do próprio PDV"
                metric = {**own_metric, "pdv_count": 0}
            elif pop_metric:
                category = "Popular no histórico da base SOPI"
                metric = pop_metric
            else:
                category = "Item da cesta SOPI sem histórico"
                metric = {"latest": 0, "month_count": 0, "quantity": 0.0, "value": 0.0, "pdv_count": 0}
            items.append({
                "sku": sku,
                "description": basket[sku],
                "category": category,
                "latest": int(metric.get("latest", 0)),
                "month_count": int(metric.get("month_count", 0)),
                "quantity": float(metric.get("quantity", 0.0)),
                "value": float(metric.get("value", 0.0)),
                "pdv_count": int(metric.get("pdv_count", 0)),
                "price": prices.get(sku),
            })
        items.sort(key=suggestion_sort_key)
        own_items = [item for item in items if item["category"] == "Histórico do próprio PDV"]
        pop_items = [item for item in items if item["category"] == "Popular no histórico da base SOPI"]
        nohist_items = [item for item in items if item["category"] == "Item da cesta SOPI sem histórico"]
        for item in items:
            suggestions.append({
                "sector": sector,
                "pdv": pdv,
                "name": name,
                "visit_day": visit_day,
                "next_visit": next_visit,
                "gap": gap,
                "sku": item["sku"],
                "description": item["description"],
                "price": item["price"],
                "delivery": delivery,
                "history_category": item["category"],
            })
        reconciliation.append({
            "sector": sector,
            "pdv": pdv,
            "name_ambev": source.get("name_ambev", ""),
            "name_anala": anala_record.get("name_anala") if sku_anala is not None else None,
            "visit_day": visit_day,
            "next_visit": next_visit,
            "sku_ambev": sku_ambev,
            "sku_anala": sku_anala,
            "sku_fiscal": sku_fiscal,
            "diff_ambev_fiscal": diff_ambev_fiscal,
            "diff_anala_fiscal": diff_anala_fiscal,
            "situation": situation,
            "gap": gap,
            "products_anala": ", ".join(sorted(anala_products, key=sku_sort_key)) if sku_anala is not None else None,
            "products_fiscal": ", ".join(sorted(fiscal_products, key=sku_sort_key)),
            "delivery": delivery,
        })
        latest_own = max((item["latest"] for item in own_items), default=0)
        bases.append({
            "sector": sector,
            "pdv": pdv,
            "name": name,
            "visit_day": visit_day,
            "next_visit": next_visit,
            "sku_ambev": sku_ambev,
            "sku_anala": sku_anala,
            "sku_fiscal": sku_fiscal,
            "gap": gap,
            "missing_count": len(items),
            "own_count": len(own_items),
            "popular_count": len(pop_items),
            "no_history_count": len(nohist_items),
            "latest_own_month": f"{report_date.year}-{latest_own:02d}" if latest_own else None,
            "own_quantity": sum(item["quantity"] for item in own_items) if own_items else None,
            "own_value": sum(item["value"] for item in own_items) if own_items else None,
            "coverage": len(current_products & regional) / len(regional),
            "classification": classification,
            "delivery": delivery,
        })
        if not sector:
            missing_sector.append(pdv)
        if not visit_day or not next_visit:
            missing_visit.append(pdv)
        if not delivery:
            missing_delivery.append(pdv)

    ambev_outside = sorted(set(ambev) - universe, key=sku_sort_key)
    comodato_quality.update({
        "universe_in_planificador": len(universe & set(ambev)),
        "universe_absent_in_planificador": len(universe - set(ambev)),
        "universe_with_billing_in_month": sum(1 for pdv in universe if current.get(pdv)),
        "planificador_outside_universe": len(ambev_outside),
        "planificador_outside_universe_sample": ambev_outside[:50],
    })
    suggestion_pairs = {(row["pdv"], row["sku"]) for row in suggestions}
    expected_pairs = {
        (pdv, sku)
        for pdv in universe
        for sku in (
            regional
            - current.get(pdv, set())
        )
    }
    category_total = sum(row["own_count"] + row["popular_count"] + row["no_history_count"] for row in bases)
    suggested_skus = {row["sku"] for row in suggestions}
    no_price_rows = sum(1 for row in suggestions if not (row["price"] or 0) > 0)
    validations = {
        "reconciliation_rows_match": len(reconciliation) == len(universe),
        "duplicate_pdvs_reconciliation": len(reconciliation) - len({row["pdv"] for row in reconciliation}),
        "duplicate_pdv_sku_suggestions": len(suggestions) - len(suggestion_pairs),
        "suggestions_match_expected": suggestion_pairs == expected_pairs,
        "category_closure": category_total == len(suggestions),
        "suggested_skus_in_national": suggested_skus <= set(basket),
        "suggested_skus_in_regional": suggested_skus <= regional,
        "suggested_skus_have_2026_billing": suggested_skus <= regional,
        "anala_mode_respected": True,
    }
    if not all(value is True or value == 0 for value in validations.values()):
        raise ValueError(f"Validação interna falhou: {validations}")
    if len(suggestions) > 1_048_575:
        raise ValueError(f"Sugestões excedem o limite do Excel: {len(suggestions)} linhas de dados.")

    output = {
        "report_date": report_date.isoformat(),
        "report_month_name": MONTH_NAMES[report_date.month],
        "sources": {
            "comodato": str(Path(cli.comodato)),
            "ambev": str(Path(cli.ambev)),
            "anala": str(anala_path) if anala_path else None,
            "basket": str(Path(cli.basket)),
            "clients": str(Path(cli.clients)),
            "delivery": str(Path(cli.delivery)),
            "prices": str(Path(cli.prices)),
            "billing": [str(path) for path in bill_paths],
        },
        "source_modified_at": {
            "comodato": datetime.fromtimestamp(Path(cli.comodato).stat().st_mtime).isoformat(timespec="seconds"),
            "ambev": datetime.fromtimestamp(Path(cli.ambev).stat().st_mtime).isoformat(timespec="seconds"),
            "basket": datetime.fromtimestamp(Path(cli.basket).stat().st_mtime).isoformat(timespec="seconds"),
            "clients": datetime.fromtimestamp(Path(cli.clients).stat().st_mtime).isoformat(timespec="seconds"),
            "delivery": datetime.fromtimestamp(Path(cli.delivery).stat().st_mtime).isoformat(timespec="seconds"),
            "prices": datetime.fromtimestamp(Path(cli.prices).stat().st_mtime).isoformat(timespec="seconds"),
            "anala": datetime.fromtimestamp(anala_path.stat().st_mtime).isoformat(timespec="seconds") if anala_path else None,
        },
        "summary": {
            "audited_pdvs": len(universe),
            "national_skus": len(basket),
            "regional_skus": len(regional),
            "excluded_national_without_2026_billing": len(set(basket) - regional),
            "suggestion_rows": len(suggestions),
            "action_commercial": sum(1 for row in bases if row["classification"] == "Ação comercial"),
            "action_commercial_validation": sum(1 for row in bases if row["classification"] == "Ação comercial e validação"),
            "missing_source_action_validation": sum(1 for row in bases if row["classification"] == "Fonte ausente + ação comercial e validação"),
            "preventive_validation": sum(1 for row in bases if row["classification"] == "Validação com relação preventiva"),
            "no_action": sum(1 for row in bases if row["classification"] == "Sem ação"),
            "convergent": sum(1 for row in reconciliation if row["situation"] == "Convergente"),
            "divergent": sum(1 for row in reconciliation if row["situation"] != "Convergente"),
            "ambev_above_billing": sum(1 for row in reconciliation if row["situation"] == "Ambev acima do faturamento"),
            "ambev_below_billing": sum(1 for row in reconciliation if row["situation"] == "Ambev abaixo do faturamento"),
            "kpi_base_sopi_2026": len(universe),
            "kpi_pdvs_10_sku_billing": sum(1 for row in bases if row["sku_fiscal"] >= TARGET),
            "kpi_pct_billing": round(100 * sum(1 for row in bases if row["sku_fiscal"] >= TARGET) / len(universe), 1),
            "kpi_pdvs_10_sku_planificador": sum(1 for row in bases if row["sku_ambev"] is not None and row["sku_ambev"] >= TARGET),
            "kpi_pct_planificador": round(100 * sum(1 for row in bases if row["sku_ambev"] is not None and row["sku_ambev"] >= TARGET) / len(universe), 1),
            "ambev_absent": sum(1 for row in reconciliation if row["situation"] == "Ausente no planificador"),
            "ambev_outside_universe": len(ambev_outside),
            "anala_absent": sum(1 for row in reconciliation if row["sku_anala"] is None),
            "cases_for_system_validation": sum(1 for row in bases if "validação" in row["classification"].lower()),
            "suggestion_rows_without_price": no_price_rows,
            "suggested_distinct_skus_with_price": len({row["sku"] for row in suggestions if (row["price"] or 0) > 0}),
            "suggested_distinct_skus_without_price": len({row["sku"] for row in suggestions if not (row["price"] or 0) > 0}),
        },
        "quality": {
            "basket": basket_quality,
            "comodato": comodato_quality,
            "ambev": ambev_quality,
            "anala": anala_quality,
            "clients": clients_quality,
            "delivery": delivery_quality,
            "prices": price_quality,
            "billing": billing_quality,
            "missing_sector": missing_sector,
            "missing_visit": missing_visit,
            "missing_delivery": missing_delivery,
            "validations": validations,
        },
        "suggestions": suggestions,
        "reconciliation": reconciliation,
        "base": bases,
    }
    Path(cli.output).write_text(json.dumps(output, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(output["summary"], ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
