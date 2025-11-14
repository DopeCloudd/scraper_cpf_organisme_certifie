import ExcelJS from "exceljs";
import fs from "node:fs/promises";
import path from "node:path";
import puppeteer, { Browser, Page } from "puppeteer";
import { logger } from "./utils/logger";
import { randomUserAgent } from "./utils/userAgent";

const INPUT_DIR = path.resolve(process.cwd(), "input");
const DEFAULT_SHEET_NAME = "Formations";
const CPF_BASE_URL =
  "https://www.moncompteformation.gouv.fr/espace-prive/html/";
const NAVIGATION_TIMEOUT_MS = Number(
  process.env.NAVIGATION_TIMEOUT_MS ?? 45000
);
const CONTENT_LOAD_DELAY_MS = Number(process.env.CONTENT_LOAD_DELAY_MS ?? 2000);
const BETWEEN_DELAY_MS = Number(process.env.BETWEEN_DELAY_MS ?? 500);
const HEADLESS =
  (process.env.PUPPETEER_HEADLESS ?? "false").toLowerCase() === "true";

interface CliOptions {
  file?: string;
  help?: boolean;
}

const normalizeText = (value?: string | null): string | undefined => {
  if (!value) return undefined;
  return value
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .trim();
};

const wait = (ms: number) =>
  new Promise((resolve) => {
    setTimeout(resolve, ms);
  });

const parseCli = (): CliOptions => {
  const args = process.argv.slice(2);
  const options: CliOptions = {};

  for (let index = 0; index < args.length; index += 1) {
    const arg = args[index];

    if (arg === "--help" || arg === "-h") {
      options.help = true;
      break;
    }

    if (arg.startsWith("--file=")) {
      options.file = arg.split("=")[1];
      continue;
    }

    if (arg === "--file" || arg === "-f") {
      const value = args[index + 1];
      if (value && !value.startsWith("-")) {
        options.file = value;
        index += 1;
      }
    }
  }

  return options;
};

const showHelp = () => {
  // eslint-disable-next-line no-console
  console.log(`Usage: npm run extract [-- --file=<chemin>]

Options:
  -f, --file <p>   Chemin vers le fichier Excel (par défaut: dernier .xlsx dans input/).
  -h, --help       Affiche cette aide.
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

const getCellText = (
  value: ExcelJS.CellValue | undefined
): string | undefined => {
  if (value == null) return undefined;
  if (typeof value === "string") return value.trim() || undefined;
  if (typeof value === "number" || typeof value === "boolean") {
    return String(value);
  }

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

const findHeaderRow = (
  sheet: ExcelJS.Worksheet
): { row: ExcelJS.Row; rowNumber: number } => {
  for (let index = 1; index <= sheet.rowCount; index += 1) {
    const row = sheet.getRow(index);
    let matches = false;
    row.eachCell((cell) => {
      const normalized = normalizeText(getCellText(cell.value));
      if (normalized === "url detail") {
        matches = true;
      }
    });
    if (matches) return { row, rowNumber: index };
  }
  throw new Error('Impossible de trouver la colonne "URL détail".');
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

const ensureContentColumn = (
  sheet: ExcelJS.Worksheet,
  headerRow: ExcelJS.Row
): number => {
  const existing = findColumnIndex(headerRow, ["contenu"]);
  if (existing) return existing;

  const newIndex = headerRow.actualCellCount + 1;
  headerRow.getCell(newIndex).value = "Contenu";

  const nextRow = sheet.getRow(headerRow.number + 1);
  const duplicate =
    normalizeText(getCellText(headerRow.getCell(1).value)) ===
    normalizeText(getCellText(nextRow.getCell(1).value));
  if (duplicate) {
    nextRow.getCell(newIndex).value = "Contenu";
  }

  return newIndex;
};

const buildDetailUrl = (raw: string): string | undefined => {
  const trimmed = raw.trim();
  if (!trimmed) return undefined;
  if (/^https?:\/\//i.test(trimmed)) return trimmed;

  let normalized = trimmed.replace(/^#/, "").replace(/^\/+/, "");
  if (!normalized.startsWith("formation/")) {
    normalized = `formation/recherche/${normalized}`;
  }

  return `${CPF_BASE_URL}#/${normalized}`;
};

type FormationRow = {
  rowNumber: number;
  url: string;
};

const collectRows = (
  sheet: ExcelJS.Worksheet
): {
  rows: FormationRow[];
  contentColumn: number;
} => {
  const { row: headerRow, rowNumber } = findHeaderRow(sheet);
  const urlColumn = findColumnIndex(headerRow, ["url detail"]);
  if (!urlColumn) throw new Error('Colonne "URL détail" introuvable.');

  const contentColumn = ensureContentColumn(sheet, headerRow);
  const rows: FormationRow[] = [];

  sheet.eachRow((row, index) => {
    if (index <= rowNumber) return;
    const urlValue = getCellText(row.getCell(urlColumn).value);
    if (!urlValue) return;
    const detailUrl = buildDetailUrl(urlValue);
    if (!detailUrl) return;

    const contentValue = getCellText(row.getCell(contentColumn).value);
    if (contentValue && contentValue.trim().length > 0) {
      return;
    }

    rows.push({
      rowNumber: index,
      url: detailUrl,
    });
  });

  return { rows, contentColumn };
};

const DEFAULT_BROWSER_ARGS = [
  "--no-sandbox",
  "--disable-setuid-sandbox",
  "--disable-dev-shm-usage",
  "--disable-blink-features=AutomationControlled",
];

const createBrowser = async (): Promise<Browser> => {
  return puppeteer.launch({
    headless: HEADLESS ? "new" : false,
    executablePath:
      "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
    defaultViewport: { width: 1366, height: 768 },
    args: DEFAULT_BROWSER_ARGS,
  });
};

const expandContentSection = async (page: Page): Promise<void> => {
  await page
    .evaluate(() => {
      const normalize = (text?: string | null) =>
        text
          ?.normalize("NFD")
          .replace(/[\u0300-\u036f]/g, "")
          .toLowerCase()
          .trim();
      const heading = Array.from(
        document.querySelectorAll<HTMLHeadingElement>("h3")
      ).find((node) => {
        const normalized = normalize(node.textContent);
        return normalized ? normalized.includes("contenu") : false;
      });
      if (!heading) return;

      const scopes = [heading.parentElement, heading.nextElementSibling].filter(
        Boolean
      ) as Element[];
      const btn = scopes
        .flatMap((scope) =>
          Array.from(scope.querySelectorAll<HTMLButtonElement>("button"))
        )
        .find((button) => {
          const controls = (
            button.getAttribute("aria-controls") ?? ""
          ).toLowerCase();
          const described = (
            button.getAttribute("aria-describedby") ?? ""
          ).toLowerCase();
          return controls.includes("contenu") || described.includes("contenu");
        });
      if (btn && btn.getAttribute("aria-expanded") === "false") {
        btn.click();
      }
    })
    .catch(() => undefined);

  await page.waitForTimeout(250);
};

const extractContentFromPage = async (
  page: Page
): Promise<string | undefined> => {
  return page
    .evaluate(() => {
      const normalize = (text?: string | null) =>
        text
          ?.normalize("NFD")
          .replace(/[\u0300-\u036f]/g, "")
          .toLowerCase()
          .trim();

      const heading = Array.from(
        document.querySelectorAll<HTMLHeadingElement>("h3")
      ).find((node) => {
        const normalized = normalize(node.textContent);
        return normalized ? normalized.includes("contenu") : false;
      });
      if (!heading) return undefined;

      const candidates: Element[] = [];
      if (heading.parentElement) candidates.push(heading.parentElement);
      if (heading.nextElementSibling)
        candidates.push(heading.nextElementSibling);

      for (const candidate of [...candidates]) {
        const collapse = candidate.querySelector(
          "[data-fr-js-collapse], .fr-collapse"
        );
        if (collapse) candidates.push(collapse);
      }

      const container =
        candidates.find(
          (node) => node.textContent && node.textContent.trim()
        ) ?? heading.parentElement;
      if (!container) return undefined;

      const htmlPieces = Array.from(
        container.querySelectorAll("p, ul, ol, li")
      )
        .map((node) => node.outerHTML?.trim())
        .filter((snippet): snippet is string => Boolean(snippet));

      if (htmlPieces.length > 0) {
        return htmlPieces.join("\n");
      }

      const fallback = container.innerHTML?.trim();
      return fallback && fallback.length > 0 ? fallback : undefined;
    })
    .catch(() => undefined);
};

const scrapeContent = async (page: Page, url: string) => {
  await page.setUserAgent(randomUserAgent());
  await page.goto(url, {
    waitUntil: "networkidle2",
    timeout: NAVIGATION_TIMEOUT_MS,
  });
  await page.waitForTimeout(CONTENT_LOAD_DELAY_MS);
  await expandContentSection(page);
  return extractContentFromPage(page);
};

const runWorkbookUpdate = async (filePath: string) => {
  logger.info("Chargement du fichier: %s", filePath);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);

  const sheet =
    workbook.getWorksheet(DEFAULT_SHEET_NAME) ?? workbook.worksheets[0];
  if (!sheet) throw new Error("Aucune feuille dans le fichier Excel.");

  const { rows, contentColumn } = collectRows(sheet);
  if (rows.length === 0) {
    logger.warn("Aucun lien détecté.");
    return;
  }

  const browser = await createBrowser();
  const page = await browser.newPage();

  let mutated = false;
  let processed = 0;
  let filled = 0;
  let failed = 0;

  try {
    for (const rowInfo of rows) {
      try {
        logger.info(
          "Extraction contenu ligne %d (%s)",
          rowInfo.rowNumber,
          rowInfo.url
        );
        const content = await scrapeContent(page, rowInfo.url);
        if (!content) {
          failed += 1;
          logger.warn(
            "Bloc Contenu introuvable pour la ligne %d",
            rowInfo.rowNumber
          );
        } else {
          sheet.getRow(rowInfo.rowNumber).getCell(contentColumn).value =
            content;
          mutated = true;
          filled += 1;
          await workbook.xlsx.writeFile(filePath);
        }
      } catch (error) {
        failed += 1;
        logger.error(
          "Erreur ligne %d (%s): %s",
          rowInfo.rowNumber,
          rowInfo.url,
          error instanceof Error ? error.message : String(error)
        );
      } finally {
        processed += 1;
        if (BETWEEN_DELAY_MS > 0) {
          await wait(BETWEEN_DELAY_MS);
        }
      }
    }
  } finally {
    await page.close().catch(() => undefined);
    await browser.close().catch(() => undefined);
  }

  if (mutated) {
    await workbook.xlsx.writeFile(filePath);
    logger.info("Fichier mis à jour: %s", filePath);
  } else {
    logger.info("Aucun changement à écrire.");
  }

  logger.info(
    "Bilan: total=%d, remplis=%d, erreurs=%d",
    processed,
    filled,
    failed
  );
};

const main = async () => {
  const options = parseCli();
  if (options.help) {
    showHelp();
    return;
  }

  const filePath = await resolveInputFile(options.file);
  await runWorkbookUpdate(filePath);
};

main().catch((error) => {
  logger.error("Exécution échouée: %s", (error as Error).message, {
    stack: (error as Error).stack,
  });
  process.exitCode = 1;
});
