from __future__ import annotations

import argparse
from contextlib import contextmanager
import hashlib
import json
import os
import shutil
import subprocess
import sys
from datetime import date, datetime
from pathlib import Path


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
EXPECTED_SHEETS = ["Sugestões", "Conciliação", "Base das Sugestões"]
DEFAULT_ROOT = Path(
    r"C:\Users\matheus.manoel\OneDrive - Vyas\Matheus - NS - Sinop\SPO - Gestão A Vista\SKU-PDV"
)

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")


def arguments() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Gera o relatório mensal SKU/PDV SOPI a partir das fontes oficiais."
    )
    parser.add_argument("--data", help="Data da análise em DD-MM-AAAA ou AAAA-MM-DD.")
    parser.add_argument("--raiz", type=Path, default=DEFAULT_ROOT, help="Pasta raiz do projeto SKU-PDV.")
    parser.add_argument("--saida", type=Path, help="Caminho completo do XLSX de saída.")
    parser.add_argument("--sobrescrever", action="store_true", help="Substitui a saída se ela já existir.")
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument("--sem-anala", action="store_true", help="Desconsidera o Anala nesta análise.")
    mode.add_argument("--usar-anala", action="store_true", help="Usa o Anala mais recente disponível até a data.")
    parser.add_argument("--comodato", type=Path, help="Arquivo 02.02.20 específico (padrão: o mais recente da pasta de comodato).")
    parser.add_argument("--anala", type=Path, help="Arquivo Anala específico. Implica --usar-anala.")
    parser.add_argument("--nao-interativo", action="store_true", help="Falha em vez de fazer perguntas.")
    return parser.parse_args()


def parse_report_date(value: str) -> date:
    for pattern in ("%d-%m-%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(value.strip(), pattern).date()
        except ValueError:
            pass
    raise ValueError("Data inválida. Use DD-MM-AAAA ou AAAA-MM-DD.")


def prompt_date() -> date:
    while True:
        try:
            return parse_report_date(input("Informe a data da análise em DD-MM-AAAA: "))
        except ValueError as exc:
            print(f"ERRO: {exc}")


def prompt_yes_no(message: str, default: bool = False) -> bool:
    suffix = "S/n" if default else "s/N"
    answer = input(f"{message} ({suffix}): ").strip().casefold()
    if not answer:
        return default
    return answer in {"s", "sim", "y", "yes"}


def dated_directories(folder: Path, report_date: date) -> list[tuple[date, Path]]:
    found: list[tuple[date, Path]] = []
    if not folder.exists():
        return found
    for child in folder.iterdir():
        if not child.is_dir():
            continue
        try:
            folder_date = datetime.strptime(child.name, "%d-%m-%Y").date()
        except ValueError:
            continue
        if folder_date <= report_date:
            found.append((folder_date, child))
    return sorted(found, key=lambda item: item[0], reverse=True)


def latest_dated_file(folder: Path, report_date: date, suffixes: tuple[str, ...]) -> Path:
    for _, dated in dated_directories(folder, report_date):
        files = [
            path
            for path in dated.rglob("*")
            if path.is_file() and path.suffix.casefold() in suffixes and not path.name.startswith("~$")
        ]
        if files:
            return max(files, key=lambda path: path.stat().st_mtime)
    raise FileNotFoundError(f"Nenhuma fonte válida encontrada em {folder} até {report_date:%d-%m-%Y}.")


def latest_fixed_file(folder: Path, suffixes: tuple[str, ...], preferred: str | None = None) -> Path:
    if preferred:
        candidate = folder / preferred
        if candidate.exists():
            return candidate
    files = [
        path
        for path in folder.rglob("*")
        if path.is_file() and path.suffix.casefold() in suffixes and not path.name.startswith("~$")
    ]
    if not files:
        raise FileNotFoundError(f"Nenhum arquivo válido encontrado em {folder}.")
    return max(files, key=lambda path: path.stat().st_mtime)


def choose_output(root: Path, report_date: date, requested: Path | None, overwrite: bool) -> Path:
    if requested:
        base = requested.resolve()
    else:
        day = report_date.strftime("%d-%m-%Y")
        base = root / f"SKUPDV - {MONTH_NAMES[report_date.month]}" / day / f"Relatorio_KPI_SKU_PDV_SOPI_{report_date:%Y-%m-%d}.xlsx"
    if overwrite or not base.exists():
        return base
    index = 2
    while True:
        candidate = base.with_name(f"{base.stem}_v{index}{base.suffix}")
        if not candidate.exists():
            return candidate
        index += 1


def locate_node() -> Path:
    candidates = [
        Path.home() / ".cache" / "codex-runtimes" / "codex-primary-runtime" / "dependencies" / "node" / "bin" / "node.exe",
        Path.home() / ".cache" / "codex-runtimes" / "codex-primary-runtime" / "dependencies" / "node" / "bin" / "node",
    ]
    for candidate in candidates:
        if candidate.exists():
            return candidate
    executable = shutil.which("node")
    if executable:
        return Path(executable)
    raise FileNotFoundError("Node.js não foi localizado. Instale o Node.js ou execute pelo Codex Desktop.")


def run(command: list[str], cwd: Path) -> None:
    print("\nExecutando:", " ".join(f'\"{part}\"' if " " in part else part for part in command[:2]), "...")
    completed = subprocess.run(command, cwd=cwd, check=False)
    if completed.returncode != 0:
        raise RuntimeError(f"A etapa terminou com o código {completed.returncode}.")


@contextmanager
def execution_directory(root: Path):
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    folder = (root / f"execucao_{stamp}_{os.getpid()}").resolve()
    if folder.parent != root.resolve():
        raise ValueError("Diretório temporário inválido.")
    folder.mkdir(parents=True, exist_ok=False)
    yield str(folder)


def validate_report(path: Path) -> dict[str, object]:
    from openpyxl import load_workbook

    workbook = load_workbook(path, read_only=False, data_only=False)
    if workbook.sheetnames != EXPECTED_SHEETS:
        raise ValueError(f"Abas inesperadas: {workbook.sheetnames}")
    expected_tables = {
        "Sugestões": "tblSugestoes",
        "Conciliação": "tblConciliacao",
        "Base das Sugestões": "tblBaseSugestoes",
    }
    formula_errors = 0
    for sheet in workbook.worksheets:
        if list(sheet.tables) != [expected_tables[sheet.title]]:
            raise ValueError(f"Tabela estruturada ausente ou inesperada em {sheet.title}.")
        if sheet.freeze_panes is not None:
            raise ValueError(f"Painel congelado inesperado em {sheet.title}.")
        if sheet.merged_cells.ranges:
            raise ValueError(f"Células mescladas inesperadas em {sheet.title}.")
        # Percorre apenas células materializadas. ``iter_rows`` visita todo o
        # retângulo teórico da planilha e tornava a validação muito lenta em
        # relatórios com dezenas de milhares de sugestões.
        for cell in sheet._cells.values():
            if isinstance(cell.value, str) and cell.value.startswith(("#REF!", "#DIV/0!", "#VALUE!", "#NAME?", "#N/A")):
                formula_errors += 1
    if formula_errors:
        raise ValueError(f"Foram encontrados {formula_errors} erros de fórmula.")
    suggestion = workbook["Sugestões"]
    for coordinate in ("A2", "B2", "G2"):
        if suggestion[coordinate].data_type != "s":
            raise ValueError(f"O identificador {coordinate} não foi gravado como texto.")
    if suggestion["E2"].data_type != "d":
        raise ValueError("A próxima visita não foi gravada como data real.")
    result = {
        "sheets": workbook.sheetnames,
        "rows": {sheet.title: sheet.max_row - 1 for sheet in workbook.worksheets},
        "formula_errors": formula_errors,
        "size_bytes": path.stat().st_size,
        "sha256": hashlib.sha256(path.read_bytes()).hexdigest(),
    }
    workbook.close()
    return result


def remove_inspection_sidecar(report_path: Path) -> bool:
    sidecar = Path(f"{report_path}.inspect.ndjson")
    if not sidecar.exists():
        return False
    sidecar.unlink()
    return True


def main() -> int:
    cli = arguments()
    script_dir = Path(__file__).resolve().parent
    analysis_script = script_dir / "build_skupdv_analysis.py"
    builder_script = script_dir / "montar_relatorio_skupdv.py"
    if not analysis_script.exists() or not builder_script.exists():
        raise FileNotFoundError("Os componentes de análise e montagem devem permanecer ao lado deste programa.")
    report_date = parse_report_date(cli.data) if cli.data else (
        (_ for _ in ()).throw(ValueError("Informe --data no modo não interativo."))
        if cli.nao_interativo
        else prompt_date()
    )
    root = cli.raiz.resolve()
    if not root.exists():
        raise FileNotFoundError(f"A pasta raiz não existe: {root}")
    use_anala = cli.usar_anala or cli.anala is not None
    if not cli.sem_anala and not use_anala:
        if cli.nao_interativo:
            raise ValueError("Informe --usar-anala ou --sem-anala no modo não interativo.")
        use_anala = prompt_yes_no("Usar o relatório Anala mais recente disponível?", default=True)

    comodato_dir = root.parent / "Giro de Comodato" / "CHOPEIRA" / "02.02.20"
    sources = {
        "comodato": cli.comodato.resolve() if cli.comodato else latest_fixed_file(comodato_dir, (".csv",)),
        "ambev": latest_dated_file(root / "PLANIFICADOR SKU-PDV", report_date, (".xlsx", ".xlsm")),
        "basket": latest_fixed_file(root / "CESTA SKUPDV", (".xlsx", ".xlsm"), "Nova Cesta Produtos.xlsx"),
        "clients": latest_dated_file(root / "CLIENTES +", report_date, (".csv",)),
        "delivery": latest_dated_file(root / "DIAS DE ENTREGA", report_date, (".csv",)),
        "prices": latest_fixed_file(root / "TABELA DE PREÇOS", (".xlsx", ".xlsm"), "Tabela_de_preços.xlsx"),
    }
    anala = cli.anala.resolve() if cli.anala else None
    if use_anala and anala is None:
        anala = latest_dated_file(root / "PLANIFICADOR ANALA", report_date, (".xlsx", ".xlsm"))
    if anala and not anala.exists():
        raise FileNotFoundError(f"O arquivo Anala não existe: {anala}")
    billing_dir = root / "FATURAMENTO"
    output = choose_output(root, report_date, cli.saida, cli.sobrescrever)

    print("\nFONTES SELECIONADAS")
    print(f"Data da análise: {report_date:%d/%m/%Y}")
    print(f"Universo (02.02.20 comodato): {sources['comodato']}")
    print(f"Planificador Ambev (confronto): {sources['ambev']}")
    print(f"Anala: {anala if anala else 'DESCONSIDERADO'}")
    print(f"CLIENTES+: {sources['clients']}")
    print(f"Dias de entrega: {sources['delivery']}")
    print(f"Cesta: {sources['basket']}")
    print(f"Tabela de preços: {sources['prices']}")
    print(f"Saída: {output}")

    temp_root = script_dir / ".temp_relatorio_skupdv"
    temp_root.mkdir(exist_ok=True)
    with execution_directory(temp_root) as temp_name:
        temp = Path(temp_name)
        analysis_json = temp / "analise.json"
        qa_dir = temp / "qa"
        ambev_copy = temp / sources["ambev"].name
        shutil.copy2(sources["ambev"], ambev_copy)
        analysis_command = [
            sys.executable,
            str(analysis_script),
            "--root", str(root),
            "--comodato", str(sources["comodato"]),
            "--ambev", str(ambev_copy),
            "--basket", str(sources["basket"]),
            "--clients", str(sources["clients"]),
            "--delivery", str(sources["delivery"]),
            "--prices", str(sources["prices"]),
            "--billing-dir", str(billing_dir),
            "--report-date", report_date.isoformat(),
            "--output", str(analysis_json),
        ]
        if anala:
            analysis_command.extend(["--anala", str(anala)])
        run(analysis_command, script_dir)
        output.parent.mkdir(parents=True, exist_ok=True)
        run([
            sys.executable,
            str(builder_script),
            "--analysis", str(analysis_json),
            "--output", str(output),
            "--qa-dir", str(qa_dir),
        ], script_dir)
        summary = json.loads(analysis_json.read_text(encoding="utf-8"))["summary"]

    validation = validate_report(output)
    sidecar_removed = remove_inspection_sidecar(output)
    print("\nPROCESSO CONCLUÍDO COM SUCESSO")
    print(f"PDVs auditados: {summary['audited_pdvs']}")
    print(f"Sugestões: {summary['suggestion_rows']}")
    print(f"%SKU/PDV pelo faturamento: {summary['kpi_pct_billing']}% ({summary['kpi_pdvs_10_sku_billing']} de {summary['kpi_base_sopi_2026']} PDVs)")
    print(f"%SKU/PDV pelo planificador Ambev: {summary['kpi_pct_planificador']}% ({summary['kpi_pdvs_10_sku_planificador']} de {summary['kpi_base_sopi_2026']} PDVs)")
    print(f"Relatório: {output}")
    print(f"SHA-256: {validation['sha256']}")
    if sidecar_removed:
        print("Arquivo técnico de inspeção removido automaticamente.")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except KeyboardInterrupt:
        print("\nOperação cancelada pelo usuário.")
        raise SystemExit(130)
    except Exception as exc:
        print(f"\nERRO: {exc}", file=sys.stderr)
        raise SystemExit(1)
