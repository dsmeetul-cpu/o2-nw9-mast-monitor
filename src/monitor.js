import { chromium } from "playwright";
import fs from "fs";
import path from "path";
import XLSX from "xlsx";

const POSTCODE = "NW9 0RY";
const COMPLAINT_REFERENCE = "C-1308267357";
const O2_URL = "https://status.o2.co.uk/";

const root = process.cwd();

const dataDir = path.join(root, "data");
const evidenceDir = path.join(root, "evidence");

fs.mkdirSync(dataDir, { recursive: true });
fs.mkdirSync(evidenceDir, { recursive: true });

const csvFile = path.join(dataDir, "o2-nw9-0ry-monitor.csv");
const xlsxFile = path.join(dataDir, "O2-NW9-0RY-Evidence.xlsx");

function londonTimestamp() {
  const now = new Date();

  return new Intl.DateTimeFormat("en-GB", {
    timeZone: "Europe/London",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit",
    hour12: false
  }).format(now);
}

function safeFileTimestamp() {
  const now = new Date();

  const parts = new Intl.DateTimeFormat("en-GB", {
    timeZone: "Europe/London",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit",
    hour12: false
  }).formatToParts(now);

  const get = name => parts.find(p => p.type === name)?.value;

  return `${get("year")}-${get("month")}-${get("day")}_${get("hour")}-${get("minute")}-${get("second")}`;
}

function csvEscape(value) {
  if (value === null || value === undefined) return "";

  const text = String(value).replace(/\r?\n/g, " ");

  if (/[",]/.test(text)) {
    return `"${text.replace(/"/g, '""')}"`;
  }

  return text;
}

function readExistingCsv() {
  if (!fs.existsSync(csvFile)) return [];

  const text = fs.readFileSync(csvFile, "utf8").trim();

  if (!text) return [];

  const lines = text.split(/\r?\n/);

  const headers = lines[0]
    .split(",")
    .map(x => x.replace(/^"|"$/g, ""));

  return lines.slice(1).map(line => {
    const values = [];
    let current = "";
    let quoted = false;

    for (let i = 0; i < line.length; i++) {
      const char = line[i];

      if (char === '"') {
        if (quoted && line[i + 1] === '"') {
          current += '"';
          i++;
        } else {
          quoted = !quoted;
        }
      } else if (char === "," && !quoted) {
        values.push(current);
        current = "";
      } else {
        current += char;
      }
    }

    values.push(current);

    const row = {};

    headers.forEach((header, index) => {
      row[header] = values[index] ?? "";
    });

    return row;
  });
}

async function findPostcodeInput(page) {
  const candidates = [
    'input[placeholder*="postcode" i]',
    'input[aria-label*="postcode" i]',
    'input[name*="postcode" i]',
    'input[id*="postcode" i]',
    'input[type="text"]'
  ];

  for (const selector of candidates) {
    const locator = page.locator(selector).first();

    if (await locator.count()) {
      try {
        if (await locator.isVisible()) {
          return locator;
        }
      } catch {
        // Continue searching.
      }
    }
  }

  return null;
}

async function findSubmitButton(page) {
  const candidates = [
    'button:has-text("Check")',
    'button:has-text("Search")',
    'button:has-text("Submit")',
    'input[type="submit"]',
    'button'
  ];

  for (const selector of candidates) {
    const locator = page.locator(selector).first();

    if (await locator.count()) {
      try {
        if (await locator.isVisible()) {
          return locator;
        }
      } catch {
        // Continue.
      }
    }
  }

  return null;
}

function classifyStatus(text) {
  const lower = text.toLowerCase();

  if (
    lower.includes("issue") ||
    lower.includes("problem") ||
    lower.includes("outage") ||
    lower.includes("not working") ||
    lower.includes("fault") ||
    lower.includes("maintenance") ||
    lower.includes("engineer")
  ) {
    return "ISSUE";
  }

  if (
    lower.includes("working normally") ||
    lower.includes("no known") ||
    lower.includes("no issues") ||
    lower.includes("no issue")
  ) {
    return "OK";
  }

  return "UNKNOWN";
}

async function main() {
  const browser = await chromium.launch({
    headless: true
  });

  const context = await browser.newContext({
    locale: "en-GB",
    timezoneId: "Europe/London",
    viewport: {
      width: 1440,
      height: 1200
    }
  });

  const page = await context.newPage();

  const timestamp = londonTimestamp();
  const fileTimestamp = safeFileTimestamp();

  let status = "UNKNOWN";
  let message = "";
  let expectedResolution = "";
  let pageText = "";
  let checkResult = "FAILED";

  try {
    await page.goto(O2_URL, {
      waitUntil: "domcontentloaded",
      timeout: 60000
    });

    await page.waitForTimeout(3000);

    const postcodeInput = await findPostcodeInput(page);

    if (!postcodeInput) {
      throw new Error("Could not find O2 postcode input.");
    }

    await postcodeInput.fill(POSTCODE);

    const submitButton = await findSubmitButton(page);

    if (submitButton) {
      await submitButton.click();
    } else {
      await postcodeInput.press("Enter");
    }

    await page.waitForTimeout(7000);

    pageText = await page.locator("body").innerText();

    status = classifyStatus(pageText);

    message = pageText
      .replace(/\s+/g, " ")
      .trim()
      .slice(0, 5000);

    const resolutionMatch = pageText.match(
      /(?:expected|estimated|should|aim).*?(?:resolved|fixed).*?(\d{1,2}[:.]\d{2}|\d{1,2}\s?(?:am|pm)|\d{1,2}\/\d{1,2}\/\d{2,4})/i
    );

    if (resolutionMatch) {
      expectedResolution = resolutionMatch[0];
    }

    checkResult = "SUCCESS";
  } catch (error) {
    message = `MONITORING ERROR: ${error.message}`;
    checkResult = "FAILED";
  }

  const screenshotName = `${fileTimestamp}_NW9-0RY.png`;
  const screenshotPath = path.join(evidenceDir, screenshotName);

  await page.screenshot({
    path: screenshotPath,
    fullPage: true
  });

  const existing = readExistingCsv();

  const previous = existing.length
    ? existing[existing.length - 1]
    : null;

  const statusChanged =
    !previous || previous["O2 Status"] !== status
      ? "YES"
      : "NO";

  const record = {
    "Date": timestamp.split(",")[0],
    "Time": timestamp.split(",")[1]?.trim() || timestamp,
    "Date & Time": timestamp,
    "Postcode": POSTCODE,
    "O2 Status": status,
    "O2 Message": message,
    "Expected Resolution": expectedResolution,
    "Complaint Reference": COMPLAINT_REFERENCE,
    "Source URL": O2_URL,
    "Screenshot": `evidence/${screenshotName}`,
    "Check Result": checkResult,
    "Status Changed?": statusChanged,
    "Previous Status": previous?.["O2 Status"] || "",
    "Notes": ""
  };

  const headers = Object.keys(record);

  if (!fs.existsSync(csvFile)) {
    fs.writeFileSync(
      csvFile,
      headers.map(csvEscape).join(",") + "\n",
      "utf8"
    );
  }

  fs.appendFileSync(
    csvFile,
    headers.map(h => csvEscape(record[h])).join(",") + "\n",
    "utf8"
  );

  const allRecords = [...existing, record];

  const worksheet = XLSX.utils.json_to_sheet(allRecords);

  worksheet["!cols"] = [
    { wch: 14 },
    { wch: 12 },
    { wch: 24 },
    { wch: 14 },
    { wch: 18 },
    { wch: 80 },
    { wch: 30 },
    { wch: 20 },
    { wch: 22 },
    { wch: 55 },
    { wch: 18 },
    { wch: 18 },
    { wch: 18 },
    { wch: 30 }
  ];

  const workbook = XLSX.utils.book_new();

  XLSX.utils.book_append_sheet(
    workbook,
    worksheet,
    "O2 Mast Log"
  );

  const summary = XLSX.utils.aoa_to_sheet([
    ["O2 NW9 0RY NETWORK EVIDENCE"],
    [],
    ["Postcode", POSTCODE],
    ["Complaint Reference", COMPLAINT_REFERENCE],
    ["Monitoring Source", O2_URL],
    ["Total Checks", allRecords.length],
    [
      "Issue Checks",
      allRecords.filter(x => x["O2 Status"] === "ISSUE").length
    ],
    [
      "OK Checks",
      allRecords.filter(x => x["O2 Status"] === "OK").length
    ],
    [
      "Unknown Checks",
      allRecords.filter(x => x["O2 Status"] === "UNKNOWN").length
    ],
    [
      "Failed Checks",
      allRecords.filter(x => x["Check Result"] === "FAILED").length
    ],
    [],
    [
      "Important",
      "Screenshots are stored in the evidence folder and provide visual evidence of each check."
    ]
  ]);

  summary["!cols"] = [
    { wch: 30 },
    { wch: 100 }
  ];

  XLSX.utils.book_append_sheet(
    workbook,
    summary,
    "Summary"
  );

  XLSX.writeFile(workbook, xlsxFile);

  await browser.close();

  console.log(JSON.stringify({
    timestamp,
    postcode: POSTCODE,
    status,
    screenshot: screenshotName,
    checkResult
  }, null, 2));
}

main().catch(error => {
  console.error(error);
  process.exit(1);
});
