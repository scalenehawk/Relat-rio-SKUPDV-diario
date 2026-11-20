import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const workDir = path.dirname(fileURLToPath(import.meta.url));
function argument(name, fallback) {
  const index = process.argv.indexOf(name);
  return index >= 0 && process.argv[index + 1] ? process.argv[index + 1] : fallback;
}
const analysisPath = path.resolve(argument("--analysis", path.join(workDir, "analysis_skupdv_2026-10-19.json")));
const outputPath = path.resolve(argument("--output", path.join(workDir, "Relatorio_KPI_SKU_PDV_SOPI_2026-10-19.xlsx")));
const qaDir = path.resolve(argument("--qa-dir", path.join(workDir, "qa_skupdv_2026-10-19")));
const data = JSON.parse(await fs.readFile(analysisPath, "utf8"));
await fs.mkdir(qaDir, { recursive: true });
await fs.mkdir(path.dirname(outputPath), { recursive: true });

const workbook = Workbook.create();
const sheets = {
  suggestions: workbook.worksheets.add("Sugestões"),
  reconciliation: workbook.worksheets.add("Conciliação"),
  base: workbook.worksheets.add("Base das Sugestões"),
};

const font = { name: "Aptos Narrow", size: 11, color: "#000000" };
const headerFont = { ...font, bold: true };

function excelDate(value) {
  return value ? new Date(`${value}T12:00:00Z`) : null;
}

function columnName(index) {
  let n = index + 1;
  let result = "";
  while (n > 0) {
    n -= 1;
    result = String.fromCharCode(65 + (n % 26)) + result;
    n = Math.floor(n / 26);
  }
  return result;
}

function configureSheet(sheet, rowCount, columnCount, widths) {
  sheet.showGridLines = true;
  sheet.freezePanes.unfreeze();
  const used = sheet.getRangeByIndexes(0, 0, rowCount, columnCount);
  used.format.font = font;
  used.format.wrapText = false;
  used.format.verticalAlignment = "bottom";
  used.format.rowHeight = 14.4;
  sheet.getRangeByIndexes(0, 0, 1, columnCount).format.font = headerFont;
  widths.forEach((width, columnIndex) => {
    sheet.getRangeByIndexes(0, columnIndex, rowCount, 1).format.columnWidth = width;
  });
}

function addTable(sheet, rowCount, columnCount, name) {
  const table = sheet.tables.add(`A1:${columnName(columnCount - 1)}${rowCount}`, true, name);
  table.style = "TableStyleLight1";
  table.showHeaders = true;
  table.showTotals = false;
  table.showBandedRows = true;
  table.showBandedColumns = false;
  table.showFilterButton = true;
}

const suggestionHeaders = [
  "Setor", "Código PDV", "Nome PDV", "Dia de Visita do VDE", "Data da Prox Visita VDE",
  "GAP", "SKU não comprado", "Produto da cesta não comprado", "Preço bairro", "Dias de entrega disponíveis",
];
const suggestionRows = data.suggestions.map((row) => [
  String(row.sector ?? ""), String(row.pdv ?? ""), row.name ?? "", row.visit_day ?? "", excelDate(row.next_visit),
  row.gap, String(row.sku ?? ""), row.description ?? "", row.price > 0 ? row.price : "", row.delivery ?? "",
]);
const suggestionValues = [suggestionHeaders, ...suggestionRows];
sheets.suggestions.getRangeByIndexes(0, 0, suggestionValues.length, suggestionHeaders.length).values = suggestionValues;
data.suggestions.forEach((row, index) => {
  if (!(row.price > 0)) sheets.suggestions.getRange(`I${index + 2}`).clear({ applyTo: "contents" });
});
configureSheet(sheets.suggestions, suggestionValues.length, suggestionHeaders.length, [7.11, 12.22, 22.11, 18.66, 21.33, 6.44, 18, 62.55, 16, 58.44]);
sheets.suggestions.getRange(`A2:B${suggestionValues.length}`).format.numberFormat = "@";
sheets.suggestions.getRange(`E2:E${suggestionValues.length}`).format.numberFormat = "dd/mm/yyyy";
sheets.suggestions.getRange(`F2:F${suggestionValues.length}`).format.numberFormat = "0";
sheets.suggestions.getRange(`G2:G${suggestionValues.length}`).format.numberFormat = "@";
sheets.suggestions.getRange(`I2:I${suggestionValues.length}`).format.numberFormat = '"R$" #,##0.00';
addTable(sheets.suggestions, suggestionValues.length, suggestionHeaders.length, "tblSugestoes");

const reconciliationHeaders = [
  "Setor", "Código PDV", "Nome no Ambev", "Nome no Anala", "Dia de Visita do VDE", "Data da Prox Visita VDE",
  "SKUs Ambev", "SKUs Anala", "SKUs Faturamento do mês", "Dif. Ambev − Fat.", "Dif. Anala − Fat.", "Situação (auditoria vs. faturamento)",
  "Gap até 10 (base faturamento)", "Produtos atuais no Anala (códigos)", "Produtos faturados no mês (códigos)", "Dias de entrega disponíveis",
];
const reconciliationRows = data.reconciliation.map((row) => [
  String(row.sector ?? ""), String(row.pdv ?? ""), row.name_ambev ?? "", row.name_anala ?? "", row.visit_day ?? "", excelDate(row.next_visit),
  row.sku_ambev, row.sku_anala, row.sku_fiscal, row.diff_ambev_fiscal, row.diff_anala_fiscal, row.situation,
  row.gap, row.products_anala ?? "", row.products_fiscal ?? "", row.delivery ?? "",
]);
const reconciliationValues = [reconciliationHeaders, ...reconciliationRows];
sheets.reconciliation.getRangeByIndexes(0, 0, reconciliationValues.length, reconciliationHeaders.length).values = reconciliationValues;
configureSheet(sheets.reconciliation, reconciliationValues.length, reconciliationHeaders.length, [7.11, 12.22, 25, 25, 24, 24, 13, 13, 25, 23, 22, 45, 26, 65, 65, 58]);
sheets.reconciliation.getRange(`A2:B${reconciliationValues.length}`).format.numberFormat = "@";
sheets.reconciliation.getRange(`F2:F${reconciliationValues.length}`).format.numberFormat = "dd/mm/yyyy";
sheets.reconciliation.getRange(`G2:K${reconciliationValues.length}`).format.numberFormat = "0";
sheets.reconciliation.getRange(`M2:M${reconciliationValues.length}`).format.numberFormat = "0";
sheets.reconciliation.getRange(`N2:O${reconciliationValues.length}`).format.numberFormat = "@";
addTable(sheets.reconciliation, reconciliationValues.length, reconciliationHeaders.length, "tblConciliacao");

const baseHeaders = [
  "Setor", "Código PDV", "Nome PDV", "Dia de Visita do VDE", "Data da Prox Visita VDE", "SKUs Ambev", "SKUs Anala",
  "SKUs Faturamento do mês", "Gap", "Qtd. itens não comprados", "Com histórico próprio", "Populares na base SOPI", "Sem histórico",
  "Mês histórico mais recente", "Quantidade histórica própria", "Valor histórico próprio", "Cobertura da cesta", "Classificação",
  "Dias de entrega disponíveis",
];
const baseRows = data.base.map((row) => [
  String(row.sector ?? ""), String(row.pdv ?? ""), row.name ?? "", row.visit_day ?? "", excelDate(row.next_visit), row.sku_ambev,
  row.sku_anala, row.sku_fiscal, row.gap, row.missing_count, row.own_count, row.popular_count, row.no_history_count,
  row.latest_own_month ? `${row.latest_own_month.slice(5, 7)}/${row.latest_own_month.slice(0, 4)}` : "", row.own_quantity,
  row.own_value, row.coverage, row.classification, row.delivery ?? "",
]);
const baseValues = [baseHeaders, ...baseRows];
sheets.base.getRangeByIndexes(0, 0, baseValues.length, baseHeaders.length).values = baseValues;
configureSheet(sheets.base, baseValues.length, baseHeaders.length, [7.11, 12.22, 25, 24, 24, 13, 13, 25, 8, 26, 22, 25, 18, 24, 28, 24, 20, 50, 58]);
sheets.base.getRange(`A2:B${baseValues.length}`).format.numberFormat = "@";
sheets.base.getRange(`E2:E${baseValues.length}`).format.numberFormat = "dd/mm/yyyy";
sheets.base.getRange(`F2:M${baseValues.length}`).format.numberFormat = "0";
sheets.base.getRange(`N2:N${baseValues.length}`).format.numberFormat = "@";
sheets.base.getRange(`O2:O${baseValues.length}`).format.numberFormat = "#,##0.0";
sheets.base.getRange(`P2:P${baseValues.length}`).format.numberFormat = '"R$" #,##0.00';
sheets.base.getRange(`Q2:Q${baseValues.length}`).format.numberFormat = "0.0%";
addTable(sheets.base, baseValues.length, baseHeaders.length, "tblBaseSugestoes");

workbook.recalculate();

const structural = await workbook.inspect({ kind: "sheet,table", maxChars: 12000, tableMaxRows: 4, tableMaxCols: 20 });
await fs.writeFile(path.join(qaDir, "estrutura.ndjson"), structural.ndjson, "utf8");
const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!",
  options: { useRegex: true, maxResults: 300 },
  summary: "final formula error scan",
});
await fs.writeFile(path.join(qaDir, "erros_formulas.ndjson"), errors.ndjson, "utf8");

const renders = [
  ["Sugestões", "A1:J28", "sugestoes_inicio.png"],
  ["Sugestões", `A${Math.max(1, suggestionValues.length - 25)}:J${suggestionValues.length}`, "sugestoes_final.png"],
  ["Conciliação", "A1:P28", "conciliacao.png"],
  ["Base das Sugestões", "A1:S28", "base_sugestoes.png"],
];
for (const [sheetName, range, fileName] of renders) {
  const preview = await workbook.render({ sheetName, range, scale: 1.25, format: "png" });
  await fs.writeFile(path.join(qaDir, fileName), new Uint8Array(await preview.arrayBuffer()));
}

const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);
console.log(JSON.stringify({ outputPath, qaDir, rows: { suggestions: suggestionRows.length, reconciliation: reconciliationRows.length, base: baseRows.length } }, null, 2));
