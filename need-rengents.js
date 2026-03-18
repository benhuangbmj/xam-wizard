const fs = require("fs");
const path = require("path");
const dotenv = require("dotenv");
const XLSX = require("xlsx");
const ExcelJS = require("exceljs");

dotenv.config({ path: path.join(__dirname, ".env") });

function normalizeName(name) {
  if (typeof name !== "string") return "";
  return name
    .trim()
    .replace(/\s+/g, " ")
    .replace(/\s*,\s*/g, ", ")
    .toLowerCase();
}

function getFirstColumnValues(rows) {
  if (rows.length === 0) return [];
  const firstKey = Object.keys(rows[0])[0];
  return rows
    .map((row) => row[firstKey])
    .filter((v) => v != null && String(v).trim() !== "");
}

async function main() {
  const folderPath = process.env.NEED_REGENTS;
  if (!folderPath) {
    throw new Error(
      "NEED_REGENTS is not set. Define it in .env or your shell environment.",
    );
  }
  if (!fs.existsSync(folderPath)) {
    throw new Error(`Folder does not exist: ${folderPath}`);
  }

  const entries = fs.readdirSync(folderPath);

  const SOURCE_FILE = "need-regents.xlsx";
  const csvFiles = entries.filter((name) =>
    name.toLowerCase().endsWith(".csv"),
  );

  if (!fs.existsSync(path.join(folderPath, SOURCE_FILE))) {
    throw new Error(`Source file "${SOURCE_FILE}" not found in ${folderPath}`);
  }

  // --- Load xlsx names ---
  const xlsxFileName = SOURCE_FILE;
  console.log(`\nReading xlsx: "${xlsxFileName}"`);
  const xlsxWorkbook = XLSX.readFile(path.join(folderPath, xlsxFileName));
  const xlsxSheet = xlsxWorkbook.Sheets[xlsxWorkbook.SheetNames[0]];
  const xlsxRows = XLSX.utils.sheet_to_json(xlsxSheet, { defval: null });

  const xlsxRawNames = getFirstColumnValues(xlsxRows);
  // Map normalized name -> original raw name (first occurrence wins)
  const xlsxNameMap = new Map();
  for (const raw of xlsxRawNames) {
    const normalized = normalizeName(String(raw));
    if (normalized && !xlsxNameMap.has(normalized)) {
      xlsxNameMap.set(normalized, String(raw));
    }
  }
  console.log(`  Loaded ${xlsxNameMap.size} unique names from xlsx.`);

  const matchedXlsxNames = new Set();

  // --- Process each csv, collecting matched names per column ---
  const columns = []; // { header: string, names: string[] }

  for (const csvFileName of csvFiles) {
    const csvPath = path.join(folderPath, csvFileName);
    const csvWorkbook = XLSX.readFile(csvPath, { raw: true });
    const csvSheet = csvWorkbook.Sheets[csvWorkbook.SheetNames[0]];
    const csvRows = XLSX.utils.sheet_to_json(csvSheet, { defval: null });

    if (csvRows.length === 0) {
      console.log(`\nSkipped "${csvFileName}": empty.`);
      continue;
    }

    const firstKey = Object.keys(csvRows[0])[0];
    console.log(`\nProcessing "${csvFileName}" (name column: "${firstKey}")`);

    const matchedNames = [];
    const unmatchedCsvNames = [];
    for (const row of csvRows) {
      const raw = row[firstKey];
      if (raw == null || String(raw).trim() === "") continue;
      const normalized = normalizeName(String(raw));
      if (xlsxNameMap.has(normalized)) {
        matchedNames.push(xlsxNameMap.get(normalized));
        matchedXlsxNames.add(normalized);
      } else {
        unmatchedCsvNames.push(String(raw));
      }
    }

    console.log(
      `  Matched ${matchedNames.length} of ${csvRows.length} students.`,
    );

    columns.push({
      header: path.parse(csvFileName).name,
      names: matchedNames,
      unmatched: unmatchedCsvNames,
    });
  }

  // --- Build single output xlsx with one column per csv using ExcelJS ---
  const maxRows = Math.max(0, ...columns.map((c) => c.names.length));

  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Matches");

  const borderStyle = {
    top: { style: "thin" },
    left: { style: "thin" },
    bottom: { style: "thin" },
    right: { style: "thin" },
  };

  columns.forEach((col, colIdx) => {
    // Header row (row 1)
    const headerCell = sheet.getCell(1, colIdx + 1);
    headerCell.value = col.header;
    headerCell.border = borderStyle;

    // Data rows
    col.names.forEach((name, rowIdx) => {
      const cell = sheet.getCell(rowIdx + 2, colIdx + 1);
      cell.value = name;
      cell.border = borderStyle;
    });

    // Fill empty cells in shorter columns so all rows have borders
    for (let r = col.names.length + 2; r <= maxRows + 1; r++) {
      sheet.getCell(r, colIdx + 1).border = borderStyle;
    }
  });

  const outputPath = path.join(folderPath, "matches.xlsx");
  await workbook.xlsx.writeFile(outputPath);
  console.log(`\nSaved all matches to "matches.xlsx".`);

  // --- Build unmatched-csv xlsx (names in CSVs not found in xlsx) ---
  const unmatchedMaxRows = Math.max(
    0,
    ...columns.map((c) => c.unmatched.length),
  );

  const unmatchedWorkbook = new ExcelJS.Workbook();
  const unmatchedSheet = unmatchedWorkbook.addWorksheet("Unmatched");

  columns.forEach((col, colIdx) => {
    const headerCell = unmatchedSheet.getCell(1, colIdx + 1);
    headerCell.value = col.header;
    headerCell.border = borderStyle;

    col.unmatched.forEach((name, rowIdx) => {
      const cell = unmatchedSheet.getCell(rowIdx + 2, colIdx + 1);
      cell.value = name;
      cell.border = borderStyle;
    });

    for (let r = col.unmatched.length + 2; r <= unmatchedMaxRows + 1; r++) {
      unmatchedSheet.getCell(r, colIdx + 1).border = borderStyle;
    }
  });

  const unmatchedOutputPath = path.join(folderPath, "unmatched.xlsx");
  await unmatchedWorkbook.xlsx.writeFile(unmatchedOutputPath);
  console.log(`Saved unmatched CSV names to "unmatched.xlsx".`);

  // --- Report xlsx names never matched in any CSV ---
  const neverMatchedXlsxNames = [...xlsxNameMap.keys()].filter(
    (name) => !matchedXlsxNames.has(name),
  );

  console.log(`\nProcessed ${csvFiles.length} csv file(s).`);

  if (neverMatchedXlsxNames.length > 0) {
    console.log(
      `\n${neverMatchedXlsxNames.length} xlsx name(s) never matched in any CSV:`,
    );
    for (const name of neverMatchedXlsxNames) {
      console.log(`  "${xlsxNameMap.get(name)}"`);
    }
  } else {
    console.log("\nAll xlsx names were matched in at least one CSV.");
  }
}

main().catch((err) => {
  console.error(err.message);
  process.exit(1);
});
