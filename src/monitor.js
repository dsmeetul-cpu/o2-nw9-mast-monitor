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

function getLondonTime() {
  return new Intl.DateTimeFormat("en-GB", {
    timeZone: "Europe/London",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit",
    hour12: false
  }).format(new Date());
}

function getFileTimestamp() {
  const parts = new Intl.DateTimeFormat("en-GB", {
    timeZone: "Europe/London",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit",
    hour12: false
  }).formatToParts(new Date());

  const value = type =>
    parts.find(part => part.type === type)?.value || "";

  return [
    value("year"),
    value("month"),
    value("day")
  ].join("-") +
    "_" +
    [
      value("hour"),
      value("minute"),
      value("second")
    ].join("-");
}

function csvEscape(value) {
  if (value === null || value === undefined) {
    return "";
  }

  const text = String(value).replace(/\r?\n/g, " ");

  if (text.includes(",") || text.includes('"')) {
    return `"${text.replace(/"/g, '""')}"`;
  }

  return text;
}

function parseCsvLine(line) {
  const result = [];
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
      result.push(current);
      current = "";
    } else {
      current += char;
    }
  }

  result.push(current);

  return result;
}

function readExistingCsv() {
  if (!fs.existsSync(csvFile)) {
    return [];
  }

  const content = fs.readFileSync(csvFile, "utf8").trim();

  if (!content) {
    return [];
  }

  const lines = content.split(/\r?\n/);
  const headers = parseCsvLine(lines[0]);

  return lines.slice(1).map(line => {
    const values = parseCsvLine(line);
    const record = {};

    headers.forEach((header, index) => {
      record[header] = values[index] || "";
    });

    return record;
  });
}

async function findPostcodeInput(page) {
  const selectors = [
    'input[placeholder*="postcode" i]',
    'input[aria-label*="postcode" i]',
    'input[name*="postcode" i]',
    'input[id*="postcode" i]',
    'input[placeholder*="post code" i]',
    'input[aria-label*="post code" i]',
    'input[name*="post code" i]',
    'input[id*="post code" i]'
  ];

  for (const selector of selectors) {
    const locator = page.locator(selector);
    const count = await locator.count();

    for (let i = 0; i < count; i++) {
      const candidate = locator.nth(i);

      try {
        if (await candidate.isVisible()) {
          return candidate;
        }
      } catch {
        // Continue.
      }
    }
  }

  return null;
}

async function findCheckButton(page) {
  const selectors = [
    'button:has-text("Check")',
    'button:has-text("Search")',
    'button:has-text("Submit")',
    'input[type="submit"]'
  ];

  for (const selector of selectors) {
    const locator = page.locator(selector);
    const count = await locator.count();

    for (let i = 0; i < count; i++) {
      const candidate = locator.nth(i);

      try {
        if (await candidate.isVisible()) {
          return candidate;
        }
      } catch {
        // Continue.
      }
    }
  }

  return null;
}

function classifyStatus(text) {
  const value = text.toLowerCase();

  const issueWords = [
    "nearby phone mast",
    "phone mast",
    "mast isn't working",
    "mast isn’t working",
    "not working as it should",
    "service might come and go",
    "engineers will be on the case",
    "known issue",
    "known problem",
    "network issue",
    "network problem",
    "outage",
    "maintenance",
    "fault"
  ];

  for (const word of issueWords) {
    if (value.includes(word)) {
      return "ISSUE";
    }
  }

  const okWords = [
    "no known issues",
    "no known issue",
    "working normally",
    "everything is working",
    "service is working"
  ];

  for (const word of okWords) {
    if (value.includes(word)) {
      return "OK";
    }
  }

  return "UNKNOWN";
}

async function saveErrorRecord(errorMessage, screenshotName = "") {
  const existing = readExistingCsv();

  const previous =
    existing.length > 0
      ? existing[existing.length - 1]
      : null;

  const now = getLondonTime();

  const record = {
    "Date": now.split(",")[0],
    "Time": now.split(",")[1]?.trim() || "",
    "Date & Time": now,
    "Postcode": POSTCODE,
    "O2 Status": "UNKNOWN",
    "O2 Message": `MONITORING ERROR: ${errorMessage}`,
    "Expected Resolution": "",
    "Complaint Reference": COMPLAINT_REFERENCE,
    "Source URL": O2_URL,
    "Screenshot": screenshotName
      ? `evidence/${screenshotName}`
      : "",
    "Check Result": "FAILED",
    "Status Changed?":
      !previous || previous["O2 Status"] !== "UNKNOWN"
        ? "YES"
        : "NO",
    "Previous Status":
      previous?.["O2 Status"] || "",
    "Notes":
      "Automation failed. This check is not evidence of an O2 outage."
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
    headers.map(header => csvEscape(record[header])).join(",") + "\n",
    "utf8"
  );
}

async function main() {
  const timestamp = getLondonTime();
  const fileTimestamp = getFileTimestamp();

  let browser = null;

  try {
    console.log("========================================");
    console.log("O2 NW9 0RY NETWORK MONITOR");
    console.log("========================================");

    console.log(`Opening: ${O2_URL}`);

    browser = await chromium.launch({
      headless: true
    });

    const context = await browser.newContext({
      locale: "en-GB",
      timezoneId: "Europe/London",
      viewport: {
        width: 1440,
        height: 1400
      }
    });

    const page = await context.newPage();

    await page.goto(O2_URL, {
      waitUntil: "domcontentloaded",
      timeout: 60000
    });

    console.log(`Page title: ${await page.title()}`);
    console.log(`URL: ${page.url()}`);

    await page.waitForTimeout(5000);

    console.log("Looking for postcode input...");

    const postcodeInput =
      await findPostcodeInput(page);

    if (!postcodeInput) {
      const errorScreenshot =
        `${fileTimestamp}_ERROR_NW9-0RY.png`;

      await page.screenshot({
        path: path.join(
          evidenceDir,
          errorScreenshot
        ),
        fullPage: true
      });

      await saveErrorRecord(
        "Could not find O2 postcode input.",
        errorScreenshot
      );

      console.error(
        "MONITORING ERROR: Could not find O2 postcode input."
      );

      await browser.close();
      return;
    }

    console.log("Postcode input FOUND.");

    /*
     * IMPORTANT:
     * Enter the postcode exactly as:
     *
     * NW9 0RY
     *
     * There is a SPACE.
     * There is NO hyphen.
     */

    await postcodeInput.click();

    await postcodeInput.fill("NW9 0RY");

    await postcodeInput.dispatchEvent("input");
    await postcodeInput.dispatchEvent("change");

    await page.waitForTimeout(1000);

    const actualValue =
      await postcodeInput.inputValue();

    console.log(
      `Postcode entered into O2: "${actualValue}"`
    );

    if (actualValue !== POSTCODE) {
      throw new Error(
        `O2 input contains "${actualValue}" instead of "${POSTCODE}".`
      );
    }

    console.log("Postcode confirmed.");

    const checkButton =
      await findCheckButton(page);

    if (checkButton) {
      console.log("Check button FOUND.");
      console.log("Clicking O2 Check button...");

      await checkButton.click();
    } else {
      console.log(
        "Check button not found. Pressing ENTER..."
      );

      await postcodeInput.press("Enter");
    }

    console.log(
      "Waiting for O2 result..."
    );

    await page.waitForTimeout(10000);

    try {
      await page.waitForLoadState(
        "networkidle",
        { timeout: 15000 }
      );
    } catch {
      console.log(
        "Network did not become completely idle; continuing."
      );
    }

    await page.waitForTimeout(3000);

    const resultText =
      await page.locator("body").innerText();

    const finalUrl = page.url();

    console.log("========================================");
    console.log("O2 RESULT");
    console.log("========================================");
    console.log(`Final URL: ${finalUrl}`);
    console.log(resultText.slice(0, 8000));
    console.log("========================================");

    /*
     * Check whether O2 actually returned the postcode.
     *
     * Comparison ignores spaces/hyphens ONLY for verification.
     * The actual submitted value remains "NW9 0RY".
     */
    const normalisedPage =
      resultText
        .toUpperCase()
        .replace(/[\s-]/g, "");

    const normalisedPostcode =
      POSTCODE
        .toUpperCase()
        .replace(/[\s-]/g, "");

    if (!normalisedPage.includes(normalisedPostcode)) {
      const errorScreenshot =
        `${fileTimestamp}_ERROR_NW9-0RY.png`;

      await page.screenshot({
        path: path.join(
          evidenceDir,
          errorScreenshot
        ),
        fullPage: true
      });

      await saveErrorRecord(
        "Postcode was entered but the O2 result page did not contain NW9 0RY.",
        errorScreenshot
      );

      console.error(
        "MONITORING ERROR: O2 result did not contain postcode."
      );

      await browser.close();
      return;
    }

    /*
     * Determine O2's actual reported status.
     */
    const status =
      classifyStatus(resultText);

    const message =
      resultText
        .replace(/\s+/g, " ")
        .trim()
        .slice(0, 5000);

    /*
     * PRIMARY EVIDENCE SCREENSHOT.
     *
     * This is taken AFTER:
     *
     * 1. O2 page loaded
     * 2. NW9 0RY entered
     * 3. Check clicked
     * 4. O2 result loaded
     */
    const screenshotName =
      `${fileTimestamp}_NW9-0RY.png`;

    const screenshotPath =
      path.join(
        evidenceDir,
        screenshotName
      );

    await page.screenshot({
      path: screenshotPath,
      fullPage: true
    });

    console.log(
      `PRIMARY EVIDENCE SCREENSHOT: ${screenshotName}`
    );

    const existing =
      readExistingCsv();

    const previous =
      existing.length > 0
        ? existing[existing.length - 1]
        : null;

    const record = {
      "Date": timestamp.split(",")[0],
      "Time": timestamp.split(",")[1]?.trim() || "",
      "Date & Time": timestamp,
      "Postcode": POSTCODE,
      "O2 Status": status,
      "O2 Message": message,
      "Expected Resolution": "",
      "Complaint Reference":
        COMPLAINT_REFERENCE,
      "Source URL": O2_URL,
      "Screenshot":
        `evidence/${screenshotName}`,
      "Check Result": "SUCCESS",
      "Status Changed?":
        !previous ||
        previous["O2 Status"] !== status
          ? "YES"
          : "NO",
      "Previous Status":
        previous?.["O2 Status"] || "",
      "Notes":
        "NW9 0RY entered and submitted to O2. Screenshot captured after result loaded."
    };

    const headers =
      Object.keys(record);

    if (!fs.existsSync(csvFile)) {
      fs.writeFileSync(
        csvFile,
        headers.map(csvEscape).join(",") + "\n",
        "utf8"
      );
    }

    fs.appendFileSync(
      csvFile,
      headers
        .map(header =>
          csvEscape(record[header])
        )
        .join(",") + "\n",
      "utf8"
    );

    /*
     * Rebuild Excel workbook.
     */
    const allRecords = [
      ...existing,
      record
    ];

    const worksheet =
      XLSX.utils.json_to_sheet(
        allRecords
      );

    worksheet["!cols"] = [
      { wch: 14 },
      { wch: 12 },
      { wch: 24 },
      { wch: 14 },
      { wch: 18 },
      { wch: 100 },
      { wch: 30 },
      { wch: 22 },
      { wch: 30 },
      { wch: 65 },
      { wch: 18 },
      { wch: 18 },
      { wch: 18 },
      { wch: 60 }
    ];

    const workbook =
      XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(
      workbook,
      worksheet,
      "O2 Mast Log"
    );

    const summary =
      XLSX.utils.aoa_to_sheet([
        ["O2 NW9 0RY NETWORK EVIDENCE"],
        [],
        ["Postcode", POSTCODE],
        [
          "Complaint Reference",
          COMPLAINT_REFERENCE
        ],
        [
          "Monitoring Source",
          O2_URL
        ],
        [
          "Total Checks",
          allRecords.length
        ],
        [
          "Successful Checks",
          allRecords.filter(
            x => x["Check Result"] === "SUCCESS"
          ).length
        ],
        [
          "Failed Checks",
          allRecords.filter(
            x => x["Check Result"] === "FAILED"
          ).length
        ],
        [
          "Issue Checks",
          allRecords.filter(
            x => x["O2 Status"] === "ISSUE" &&
                 x["Check Result"] === "SUCCESS"
          ).length
        ],
        [
          "OK Checks",
          allRecords.filter(
            x => x["O2 Status"] === "OK" &&
                 x["Check Result"] === "SUCCESS"
          ).length
        ],
        [
          "Unknown Checks",
          allRecords.filter(
            x => x["O2 Status"] === "UNKNOWN"
          ).length
        ],
        [],
        [
          "Evidence Rule",
          "Only SUCCESS checks are considered valid O2 status checks. FAILED checks represent automation failures and are not treated as outage evidence."
        ],
        [],
        [
          "Primary Evidence",
          "The timestamped screenshot is captured after NW9 0RY is entered and the O2 result is displayed."
        ]
      ]);

    summary["!cols"] = [
      { wch: 30 },
      { wch: 110 }
    ];

    XLSX.utils.book_append_sheet(
      workbook,
      summary,
      "Summary"
    );

    XLSX.writeFile(
      workbook,
      xlsxFile
    );

    console.log("========================================");
    console.log("SUCCESS");
    console.log(`Postcode: ${POSTCODE}`);
    console.log(`O2 Status: ${status}`);
    console.log("Check Result: SUCCESS");
    console.log(
      `Screenshot: ${screenshotName}`
    );
    console.log("========================================");

    await browser.close();

  } catch (error) {
    console.error(
      "MONITORING ERROR:",
      error.message
    );

    try {
      if (browser) {
        const contexts =
          browser.contexts();

        if (contexts.length > 0) {
          const pages =
            contexts[0].pages();

          if (pages.length > 0) {
            const errorScreenshot =
              `${fileTimestamp}_ERROR_NW9-0RY.png`;

            await pages[0].screenshot({
              path: path.join(
                evidenceDir,
                errorScreenshot
              ),
              fullPage: true
            });

            await saveErrorRecord(
              error.message,
              errorScreenshot
            );
          }
        }
      }
    } catch (secondaryError) {
      console.error(
        "Could not save error evidence:",
        secondaryError.message
      );

      await saveErrorRecord(
        error.message
      );
    }

    if (browser) {
      try {
        await browser.close();
      } catch {
        // Ignore.
      }
    }
  }
}

main();
