import ExcelJS from "exceljs";
import fs from "node:fs/promises";
import path from "node:path";
import { logger } from "./utils/logger";

const INPUT_DIR = path.resolve(process.cwd(), "input");
const DEFAULT_OUTPUT_DIR = path.resolve(process.cwd(), "output");
const DEFAULT_MAX_CENTRES = Number(process.env.SPLIT_MAX_CENTRES ?? 1000);

interface SplitCliOptions {
  file?: string;
  outputDir?: string;
  maxCentres?: number;
  maxFormations?: number;
  prefix?: string;
  help?: boolean;
}

interface CapturedRow {
  id: string;
  values: ExcelJS.CellValue[];
}

const normalizeText = (value?: string | null): string | undefined => {
  if (!value) return undefined;
  return value
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .trim();
};

const getCellText = (
  value: ExcelJS.CellValue | undefined
): string | undefined => {
  if (value == null) return undefined;
  if (typeof value === "string") return value.trim() || undefined;
  if (typeof value === "number" || typeof value === "boolean") {
    return String(value);
  }

  if (value instanceof Date) return value.toISOString();

  if (typeof value === "object") {
    if ("text" in value && typeof value.text === "string") {
      return value.text.trim() || undefined;
    }

    if ("hyperlink" in value && typeof value.hyperlink === "string") {
      return value.hyperlink.trim() || undefined;
    }

    if ("richText" in value && Array.isArray(value.richText)) {
      return (
        value.richText
          .map((part) => part.text)
          .join("")
          .trim() || undefined
      );
    }
  }

  return undefined;
};

const parsePositiveInteger = (value?: string): number | undefined => {
  if (!value) return undefined;
  const parsed = Number(value);
  if (!Number.isFinite(parsed) || parsed <= 0) {
    throw new Error(`Valeur numérique invalide: ${value}`);
  }
  return Math.floor(parsed);
};

const parseCli = (): SplitCliOptions => {
  const args = process.argv.slice(2);
  const options: SplitCliOptions = {};

  for (let index = 0; index < args.length; index += 1) {
    const arg = args[index];

    if (arg === "--help" || arg === "-h") {
      options.help = true;
      break;
    }

    if (arg.startsWith("--file=")) {
      options.file = arg.split("=", 2)[1];
      continue;
    }

    if (arg === "--file" || arg === "-f") {
      const value = args[index + 1];
      if (value && !value.startsWith("-")) {
        options.file = value;
        index += 1;
      }
      continue;
    }

    if (arg.startsWith("--output-dir=")) {
      options.outputDir = arg.split("=", 2)[1];
      continue;
    }

    if (arg === "--output-dir" || arg === "-o") {
      const value = args[index + 1];
      if (value && !value.startsWith("-")) {
        options.outputDir = value;
        index += 1;
      }
      continue;
    }

    if (arg.startsWith("--prefix=")) {
      options.prefix = arg.split("=", 2)[1];
      continue;
    }

    if (arg === "--prefix" || arg === "-p") {
      const value = args[index + 1];
      if (value && !value.startsWith("-")) {
        options.prefix = value;
        index += 1;
      }
      continue;
    }

    if (arg.startsWith("--max-centres=") || arg.startsWith("--max-centers=")) {
      const [, raw] = arg.split("=", 2);
      options.maxCentres = parsePositiveInteger(raw);
      continue;
    }

    if (
      arg === "--max-centres" ||
      arg === "--max-centers" ||
      arg === "-c"
    ) {
      const value = args[index + 1];
      if (value && !value.startsWith("-")) {
        options.maxCentres = parsePositiveInteger(value);
        index += 1;
      }
      continue;
    }

    if (arg.startsWith("--max-formations=")) {
      const [, raw] = arg.split("=", 2);
      options.maxFormations = parsePositiveInteger(raw);
      continue;
    }

    if (arg === "--max-formations" || arg === "-m") {
      const value = args[index + 1];
      if (value && !value.startsWith("-")) {
        options.maxFormations = parsePositiveInteger(value);
        index += 1;
      }
    }
  }

  return options;
};

const showHelp = () => {
  // eslint-disable-next-line no-console
  console.log(`Usage: npm run split -- [options]

Options:
  -f, --file <p>             Chemin vers le fichier Excel à découper (par défaut: dernier .xlsx de input/).
  -o, --output-dir <p>       Dossier de sortie (défaut: ./output).
  -p, --prefix <texte>       Préfixe des fichiers générés (défaut: nom du fichier source).
  -c, --max-centres <n>      Nombre max. de centres par fichier (défaut: ${DEFAULT_MAX_CENTRES}).
  -m, --max-formations <n>   Nombre max. de formations par fichier (optionnel).
  -h, --help                 Affiche cette aide.
`);
};

const resolveInputFile = async (explicit?: string): Promise<string> => {
  if (explicit) {
    const candidate = path.resolve(process.cwd(), explicit);
    await fs.access(candidate);
    return candidate;
  }

  const entries = await fs.readdir(INPUT_DIR, { withFileTypes: true });
  const excelFiles = await Promise.all(
    entries
      .filter(
        (entry) => entry.isFile() && entry.name.toLowerCase().endsWith(".xlsx")
      )
      .map(async (entry) => {
        const fullPath = path.join(INPUT_DIR, entry.name);
        const stats = await fs.stat(fullPath);
        return { path: fullPath, mtime: stats.mtimeMs };
      })
  );

  if (excelFiles.length === 0) {
    throw new Error(
      "Aucun fichier .xlsx trouvé dans input/. Utilisez --file pour préciser un chemin."
    );
  }

  excelFiles.sort((a, b) => b.mtime - a.mtime);
  return excelFiles[0].path;
};

const findHeaderRow = (
  sheet: ExcelJS.Worksheet,
  targets: string[]
): { row: ExcelJS.Row; rowNumber: number } => {
  for (let index = 1; index <= sheet.rowCount; index += 1) {
    const row = sheet.getRow(index);
    let matches = false;
    row.eachCell((cell) => {
      const normalized = normalizeText(getCellText(cell.value));
      if (normalized && targets.includes(normalized)) {
        matches = true;
      }
    });
    if (matches) return { row, rowNumber: index };
  }
  throw new Error(
    `Impossible de trouver une ligne d'entête contenant: ${targets.join(", ")}`
  );
};

const findColumnIndex = (
  headerRow: ExcelJS.Row,
  targets: string[]
): number | undefined => {
  let found: number | undefined;
  headerRow.eachCell((cell, colNumber) => {
    const normalized = normalizeText(getCellText(cell.value));
    if (normalized && targets.includes(normalized)) {
      found = colNumber;
    }
  });
  return found;
};

const captureRowValues = (
  row: ExcelJS.Row,
  columnCount: number
): ExcelJS.CellValue[] => {
  const values: ExcelJS.CellValue[] = [];
  for (let column = 1; column <= columnCount; column += 1) {
    const cell = row.getCell(column);
    values.push((cell?.value ?? null) as ExcelJS.CellValue);
  }
  return values;
};

const rowHasValues = (row: ExcelJS.Row): boolean => {
  let hasValue = false;
  row.eachCell({ includeEmpty: false }, (cell) => {
    if (hasValue) return;
    const value = cell.value;
    if (value === null || value === undefined) return;
    if (typeof value === "string" && value.trim() === "") return;
    hasValue = true;
  });
  return hasValue;
};

const detectDataStartRow = (
  sheet: ExcelJS.Worksheet,
  headerRowNumber: number,
  idColumn: number
): number => {
  let lastHeaderRow = headerRowNumber;
  for (let rowNumber = headerRowNumber + 1; rowNumber <= sheet.rowCount; rowNumber += 1) {
    const row = sheet.getRow(rowNumber);
    if (!rowHasValues(row)) {
      lastHeaderRow = rowNumber;
      continue;
    }

    const cellText = normalizeText(getCellText(row.getCell(idColumn).value));
    if (!cellText || cellText === "id") {
      lastHeaderRow = rowNumber;
      continue;
    }

    return rowNumber;
  }
  return lastHeaderRow + 1;
};

const extractIdentifier = (
  value: ExcelJS.CellValue | undefined
): string | undefined => {
  if (value == null) return undefined;
  if (typeof value === "number") return String(value);
  if (typeof value === "string") return value.trim() || undefined;
  if (value instanceof Date) return String(value.getTime());

  const text = getCellText(value);
  return text?.trim() || undefined;
};

const applyColumnWidths = (
  source: ExcelJS.Worksheet,
  target: ExcelJS.Worksheet
): void => {
  source.columns?.forEach((column, index) => {
    if (column && typeof column.width === "number") {
      target.getColumn(index + 1).width = column.width;
    }
  });
};

const main = async () => {
  const options = parseCli();
  if (options.help) {
    showHelp();
    return;
  }

  const inputPath = await resolveInputFile(options.file);
  const outputDir = path.resolve(
    process.cwd(),
    options.outputDir ?? DEFAULT_OUTPUT_DIR
  );
  const maxCentres = options.maxCentres ?? DEFAULT_MAX_CENTRES;
  const maxFormations = options.maxFormations;

  if (maxCentres <= 0) {
    throw new Error("Le nombre maximum de centres doit être supérieur à 0.");
  }

  if (maxFormations !== undefined && maxFormations <= 0) {
    throw new Error(
      "Le nombre maximum de formations doit être supérieur à 0."
    );
  }

  logger.info(
    "Découpage du fichier %s (max %d centres, max %s formations)...",
    path.basename(inputPath),
    maxCentres,
    maxFormations ?? "∞"
  );

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(inputPath);

  const centresSheet = workbook.getWorksheet("Centres");
  const formationsSheet = workbook.getWorksheet("Formations");

  if (!centresSheet || !formationsSheet) {
    throw new Error(
      "Le fichier doit contenir les feuilles « Centres » et « Formations »."
    );
  }

  const centresHeader = findHeaderRow(centresSheet, ["id"]);
  const centresIdColumn = findColumnIndex(centresHeader.row, ["id"]);
  if (!centresIdColumn) {
    throw new Error(
      "Impossible d'identifier la colonne ID dans la feuille Centres."
    );
  }

  const formationsHeader = findHeaderRow(formationsSheet, [
    "centre id",
    "id centre",
  ]);
  const formationCentreIdColumn = findColumnIndex(formationsHeader.row, [
    "centre id",
    "id centre",
  ]);
  if (!formationCentreIdColumn) {
    throw new Error(
      "Impossible d'identifier la colonne « Centre ID » dans la feuille Formations."
    );
  }

  const centresDataStart = detectDataStartRow(
    centresSheet,
    centresHeader.rowNumber,
    centresIdColumn
  );
  const formationsDataStart = detectDataStartRow(
    formationsSheet,
    formationsHeader.rowNumber,
    formationCentreIdColumn
  );

  const centreColumnCount = centresSheet.columnCount;
  const formationColumnCount = formationsSheet.columnCount;

  const centreHeaders: ExcelJS.CellValue[][] = [];
  for (let rowNumber = 1; rowNumber < centresDataStart; rowNumber += 1) {
    const row = centresSheet.getRow(rowNumber);
    if (rowHasValues(row)) {
      centreHeaders.push(captureRowValues(row, centreColumnCount));
    }
  }

  const formationHeaders: ExcelJS.CellValue[][] = [];
  for (let rowNumber = 1; rowNumber < formationsDataStart; rowNumber += 1) {
    const row = formationsSheet.getRow(rowNumber);
    if (rowHasValues(row)) {
      formationHeaders.push(captureRowValues(row, formationColumnCount));
    }
  }

  const centres: CapturedRow[] = [];
  for (let rowNumber = centresDataStart; rowNumber <= centresSheet.rowCount; rowNumber += 1) {
    const row = centresSheet.getRow(rowNumber);
    if (!rowHasValues(row)) continue;
    const identifier = extractIdentifier(row.getCell(centresIdColumn).value);
    if (!identifier) {
      logger.warn(
        "Ligne Centres #%d ignorée: impossible de lire l'ID.",
        rowNumber
      );
      continue;
    }
    centres.push({
      id: identifier,
      values: captureRowValues(row, centreColumnCount),
    });
  }

  const formationByCentre = new Map<string, ExcelJS.CellValue[][]>();
  for (
    let rowNumber = formationsDataStart;
    rowNumber <= formationsSheet.rowCount;
    rowNumber += 1
  ) {
    const row = formationsSheet.getRow(rowNumber);
    if (!rowHasValues(row)) continue;
    const identifier = extractIdentifier(
      row.getCell(formationCentreIdColumn).value
    );
    if (!identifier) {
      logger.warn(
        "Ligne Formations #%d ignorée: aucun Centre ID.",
        rowNumber
      );
      continue;
    }
    const existing = formationByCentre.get(identifier) ?? [];
    existing.push(captureRowValues(row, formationColumnCount));
    formationByCentre.set(identifier, existing);
  }

  const chunks: CapturedRow[][] = [];
  let currentChunk: CapturedRow[] = [];
  let currentFormationCount = 0;

  for (const centre of centres) {
    const formations = formationByCentre.get(centre.id) ?? [];
    if (maxFormations && formations.length > maxFormations) {
      throw new Error(
        `Le centre ${centre.id} possède ${formations.length} formations (> ${maxFormations}).`
      );
    }

    const wouldExceedCentres =
      maxCentres && currentChunk.length >= maxCentres;
    const wouldExceedFormations =
      maxFormations &&
      currentChunk.length > 0 &&
      currentFormationCount + formations.length > maxFormations;

    if (
      currentChunk.length > 0 &&
      (wouldExceedCentres || wouldExceedFormations)
    ) {
      chunks.push(currentChunk);
      currentChunk = [];
      currentFormationCount = 0;
    }

    currentChunk.push(centre);
    currentFormationCount += formations.length;
  }

  if (currentChunk.length > 0) {
    chunks.push(currentChunk);
  }

  if (chunks.length === 0) {
    logger.warn("Aucune donnée à écrire.");
    return;
  }

  await fs.mkdir(outputDir, { recursive: true });
  const baseName =
    options.prefix ??
    path.basename(inputPath, path.extname(inputPath)).replace(/\s+/g, "_");
  const digits = Math.max(String(chunks.length).length, 2);

  const summary: Array<{ file: string; centres: number; formations: number }> =
    [];

  for (const [index, chunk] of chunks.entries()) {
    const workbookPart = new ExcelJS.Workbook();
    const centresOut = workbookPart.addWorksheet("Centres");
    const formationsOut = workbookPart.addWorksheet("Formations");

    applyColumnWidths(centresSheet, centresOut);
    applyColumnWidths(formationsSheet, formationsOut);

    centreHeaders.forEach((row) => {
      centresOut.addRow(row);
    });
    chunk.forEach((row) => {
      centresOut.addRow(row.values);
    });

    formationHeaders.forEach((row) => {
      formationsOut.addRow(row);
    });
    let chunkFormationsCount = 0;
    chunk.forEach((centre) => {
      const rows = formationByCentre.get(centre.id) ?? [];
      chunkFormationsCount += rows.length;
      rows.forEach((values) => formationsOut.addRow(values));
    });

    const resume = workbookPart.addWorksheet("Résumé");
    resume.addRow(["Fichier source", path.basename(inputPath)]);
    resume.addRow(["Partie", `${index + 1}/${chunks.length}`]);
    resume.addRow(["Centres", chunk.length]);
    resume.addRow(["Formations", chunkFormationsCount]);
    resume.addRow([
      "Paramètres",
      `max centres=${maxCentres}${
        maxFormations ? `, max formations=${maxFormations}` : ""
      }`,
    ]);

    const fileName = `${baseName}_part-${String(index + 1).padStart(
      digits,
      "0"
    )}.xlsx`;
    const outputPath = path.join(outputDir, fileName);
    await workbookPart.xlsx.writeFile(outputPath);
    summary.push({ file: fileName, centres: chunk.length, formations: chunkFormationsCount });
    logger.info(
      "Fichier %s généré (%d centres, %d formations).",
      fileName,
      chunk.length,
      chunkFormationsCount
    );
  }

  logger.info(
    "Découpage terminé: %d fichiers créés dans %s.",
    summary.length,
    outputDir
  );
};

main().catch((error) => {
  logger.error(error.message, { stack: error.stack });
  process.exitCode = 1;
});
