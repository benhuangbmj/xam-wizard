const fs = require("fs");
const path = require("path");
const dotenv = require("dotenv");
const XLSX = require("xlsx");

dotenv.config({ path: path.join(__dirname, ".env") });

function loadGradebookPath() {
  const gradebookPath = process.env.GRADEBOOK_PATH;

  if (!gradebookPath) {
    throw new Error(
      "GRADEBOOK_PATH is not set. Define it in .env or your shell environment.",
    );
  }

  return gradebookPath;
}

function normalizeName(name) {
  if (typeof name !== "string") return "";

  return name
    .trim()
    .replace(/\s+/g, " ")
    .replace(/\s*,\s*/g, ", ")
    .toLowerCase();
}

function toNumber(value) {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  if (typeof value !== "string") return null;

  const cleaned = value.replace(/[%,$\s]/g, "");
  if (!cleaned) return null;

  const parsed = Number(cleaned);
  return Number.isFinite(parsed) ? parsed : null;
}

function findNameColumns(row) {
  const keys = Object.keys(row || {});
  const keyLower = new Map(keys.map((k) => [k.toLowerCase(), k]));

  for (const key of keys) {
    const lowerKey = String(key).trim().toLowerCase();
    if (
      lowerKey === "last name, first name" ||
      (lowerKey.includes("last name") && lowerKey.includes("first name"))
    ) {
      return { combined: key };
    }
  }

  const simpleNameKey = Array.from(keyLower.entries()).find(
    ([k]) => k === "name" || k === "student name",
  )?.[1];
  if (simpleNameKey) {
    return { combined: simpleNameKey };
  }

  const firstKey = Array.from(keyLower.entries()).find(
    ([k]) => k === "first" || k === "first name",
  )?.[1];
  const lastKey = Array.from(keyLower.entries()).find(
    ([k]) => k === "last" || k === "last name",
  )?.[1];

  if (firstKey && lastKey) {
    return { first: firstKey, last: lastKey };
  }

  return null;
}

function getStudentName(row, nameColumns) {
  if (!nameColumns) return null;

  if (nameColumns.combined) {
    return row[nameColumns.combined];
  }

  if (nameColumns.first && nameColumns.last) {
    const first = row[nameColumns.first];
    const last = row[nameColumns.last];
    if (first || last) {
      return `${last || ""}, ${first || ""}`.trim();
    }
  }

  return null;
}

function buildStudentGradesMap(folderPath) {
  const xlsxFiles = fs
    .readdirSync(folderPath)
    .filter(
      (name) => name.toLowerCase().endsWith(".xlsx") && !name.startsWith("~$"),
    );

  const gradesByStudent = new Map();

  for (const fileName of xlsxFiles) {
    console.log(`\nReading xlsx: "${fileName}"`);
    const workbook = XLSX.readFile(path.join(folderPath, fileName));

    for (const sheetName of workbook.SheetNames) {
      const worksheet = workbook.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(worksheet, { defval: null });

      if (rows.length === 0) {
        console.log(`  Sheet "${sheetName}": empty`);
        continue;
      }

      const nameColumns = findNameColumns(rows[0]);
      if (!nameColumns) {
        console.log(
          `  Sheet "${sheetName}": no name columns found. Columns: [${Object.keys(rows[0]).join(", ")}]`,
        );
        continue;
      }

      console.log(`  Sheet "${sheetName}": ${rows.length} rows`);

      for (const row of rows) {
        const rawName = getStudentName(row, nameColumns);
        const normalized = normalizeName(rawName);

        if (!normalized) {
          continue;
        }

        const values = Object.entries(row)
          .filter(([column]) => {
            return column === "Grade";
          })
          .map(([, cell]) => toNumber(cell))
          .filter((num) => num !== null);

        if (values.length === 0) {
          continue;
        }

        if (!gradesByStudent.has(normalized)) {
          gradesByStudent.set(normalized, []);
        }

        gradesByStudent.get(normalized).push(...values);
      }
    }
  }

  console.log(`\nBuilt grades map with ${gradesByStudent.size} students:`);
  for (const [name, grades] of gradesByStudent.entries()) {
    const avg =
      Math.round(
        (grades.reduce((a, b) => a + b, 0) / grades.length / 100) * 4 * 2,
      ) / 2;
    console.log(`  "${name}": ${grades.length} grades, avg=${avg}`);
  }

  return { xlsxFiles, gradesByStudent };
}

function processCsvFiles(folderPath, gradesByStudent) {
  const csvFiles = fs
    .readdirSync(folderPath)
    .filter((name) => name.toLowerCase().endsWith(".csv"))
    .filter((name) => !name.toLowerCase().endsWith("_with_averages.csv"));

  const outputFiles = [];
  const matchedXlsxNames = new Set();

  for (const csvFileName of csvFiles) {
    const csvPath = path.join(folderPath, csvFileName);
    const workbook = XLSX.readFile(csvPath, { raw: true });
    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(firstSheet, { defval: null });

    if (rows.length === 0) continue;

    const nameColumns = findNameColumns(rows[0]);
    if (!nameColumns) {
      console.warn(`Skipped ${csvFileName}: no recognizable name column.`);
      continue;
    }

    console.log(`\nProcessing "${csvFileName}"`);

    const header = Object.keys(rows[0]);
    const averageColumn = "Score";

    let matchCount = 0;
    const outputRows = rows.map((row) => {
      const rawName = getStudentName(row, nameColumns);
      const normalized = normalizeName(rawName);
      const grades = gradesByStudent.get(normalized) || [];

      if (grades.length > 0) {
        matchCount++;
        matchedXlsxNames.add(normalized);
      }

      let average = "M";
      if (grades.length > 0) {
        const sum = grades.reduce((acc, val) => acc + val, 0);
        average = Math.round(Number((sum / grades.length / 100) * 4) * 2) / 2;
      }

      return {
        ...row,
        [averageColumn]: average,
      };
    });

    console.log(`  Matched ${matchCount} of ${outputRows.length} students`);

    const worksheet = XLSX.utils.json_to_sheet(outputRows, {
      header: [...header, averageColumn],
    });
    const outputCsv = XLSX.utils.sheet_to_csv(worksheet);

    const outputName = `${path.parse(csvFileName).name}_with_averages.csv`;
    const outputPath = path.join(folderPath, outputName);

    fs.writeFileSync(outputPath, outputCsv, "utf8");
    outputFiles.push(outputPath);
  }

  return { csvFiles, outputFiles, matchedXlsxNames };
}

function main() {
  const gradebookPath = loadGradebookPath();

  if (!fs.existsSync(gradebookPath)) {
    throw new Error(`Folder does not exist: ${gradebookPath}`);
  }

  const { xlsxFiles, gradesByStudent } = buildStudentGradesMap(gradebookPath);
  const { csvFiles, outputFiles, matchedXlsxNames } = processCsvFiles(
    gradebookPath,
    gradesByStudent,
  );

  console.log(`Loaded ${xlsxFiles.length} xlsx file(s).`);
  console.log(`Scanned ${csvFiles.length} csv file(s).`);
  console.log(`Created ${outputFiles.length} output csv file(s).`);

  for (const outputFile of outputFiles) {
    console.log(`- ${outputFile}`);
  }

  const unmatchedXlsxNames = [...gradesByStudent.keys()].filter(
    (name) => !matchedXlsxNames.has(name),
  );
  if (unmatchedXlsxNames.length > 0) {
    console.log(
      `\n${unmatchedXlsxNames.length} xlsx name(s) never matched in any CSV:`,
    );
    for (const name of unmatchedXlsxNames) {
      console.log(`  "${name}"`);
    }
  } else {
    console.log("\nAll xlsx names were matched in at least one CSV.");
  }
}

main();
