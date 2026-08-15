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

function safeFileTimestamp() {
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

  const get = name =>
    parts.find(part => part.type === name)?.value || "00";

  return `${get("year")}-${get("month")}-${get("day")}_${get("hour")}-${get("minute")}-${get("second")}`;
}

function csvEscape(value) {
  if (value === null || value === undefined) {
    return "";
  }

  const text = String(value).replace(/\r?\n/g, " ");

  if (/[",]/.test(text)) {
    return `"${text.replace(/"/g, '""')}"`;
  }

  return text;
}

function parseCsvLine(line) {
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

  return values;
}

function readExistingCsv() {
  if (!fs.existsSync(csvFile)) {
    return [];
  }

  const text = fs.readFileSync(csvFile, "utf8").trim();

  if (!text) {
    return [];
  }

  const lines = text.split(/\r?\n/);
  const headers = parseCsvLine(lines[0]);

  return lines.slice(1).map(line => {
    const values = parseCsvLine(line);
    const record = {};

    headers.forEach((header, index) => {
      record[header] = values[index] ?? "";
    });

    return record;
  });
}

function classifyStatus(text) {
  const lower = text.toLowerCase();

  if (
    lower.includes("nearby phone mast isn't working") ||
    lower.includes("nearby phone mast isn't working as it should") ||
    lower.includes("mast isn’t working") ||
    lower.includes("mast is not working") ||
    lower.includes("phone mast") ||
    lower.includes("engineers will be on the case") ||
    lower.includes("service might come and go") ||
    lower.includes("outage") ||
    lower.includes("fault") ||
    lower.includes("maintenance")
  ) {
    return "ISSUE";
  }

  if (
    lower.includes("working normally") ||
    lower.includes("no known issues") ||
    lower.includes("no issues")
  ) {
    return "OK";
  }

  return "UNKNOWN";
}

async function dumpDiagnostics(page) {
  console.log("");
  console.log("============================================================");
  console.log("O2 DIAGNOSTIC INFORMATION");
  console.log("============================================================");

  console.log("");
  console.log("MAIN PAGE URL:");
  console.log(page.url());

  console.log("");
  console.log("============================================================");
  console.log("ALL FRAMES");
  console.log("============================================================");

  const frames = page.frames();

  frames.forEach((frame, index) => {
    console.log("");
    console.log(`FRAME ${index}`);
    console.log(`URL: ${frame.url()}`);
  });

  for (let frameIndex = 0; frameIndex < frames.length; frameIndex++) {
    const frame = frames[frameIndex];

    console.log("");
    console.log("============================================================");
    console.log(`VISIBLE INPUTS - FRAME ${frameIndex}`);
    console.log(`FRAME URL: ${frame.url()}`);
    console.log("============================================================");

    try {
      const inputs = await frame.locator("input").evaluateAll(elements =>
        elements
          .filter(element => {
            const style = window.getComputedStyle(element);
            const rect = element.getBoundingClientRect();

            return (
              style.display !== "none" &&
              style.visibility !== "hidden" &&
              rect.width > 0 &&
              rect.height > 0
            );
          })
          .map((element, index) => ({
            index,
            type: element.type || "",
            name: element.name || "",
            id: element.id || "",
            placeholder: element.placeholder || "",
            ariaLabel: element.getAttribute("aria-label") || "",
            value: element.value || "",
            outerHTML: element.outerHTML.slice(0, 1000)
          }))
      );

      if (!inputs.length) {
        console.log("NO VISIBLE INPUTS");
      } else {
        inputs.forEach(input => {
          console.log("");
          console.log(`INPUT ${input.index}`);
          console.log(`type: ${input.type}`);
          console.log(`name: ${input.name}`);
          console.log(`id: ${input.id}`);
          console.log(`placeholder: ${input.placeholder}`);
          console.log(`aria-label: ${input.ariaLabel}`);
          console.log(`value: ${input.value}`);
          console.log(`HTML: ${input.outerHTML}`);
        });
      }
    } catch (error) {
      console.log(`Unable to inspect inputs: ${error.message}`);
    }

    console.log("");
    console.log("============================================================");
    console.log(`VISIBLE BUTTONS - FRAME ${frameIndex}`);
    console.log(`FRAME URL: ${frame.url()}`);
    console.log("============================================================");

    try {
      const buttons = await frame.locator(
        'button, input[type="submit"], input[type="button"], [role="button"]'
      ).evaluateAll(elements =>
        elements
          .filter(element => {
            const style = window.getComputedStyle(element);
            const rect = element.getBoundingClientRect();

            return (
              style.display !== "none" &&
              style.visibility !== "hidden" &&
              rect.width > 0 &&
              rect.height > 0
            );
          })
          .map((element, index) => ({
            index,
            tag: element.tagName,
            type: element.getAttribute("type") || "",
            text: (element.innerText || element.value || "")
              .replace(/\s+/g, " ")
              .trim(),
            name: element.getAttribute("name") || "",
            id: element.id || "",
            ariaLabel: element.getAttribute("aria-label") || "",
            outerHTML: element.outerHTML.slice(0, 1000)
          }))
      );

      if (!buttons.length) {
        console.log("NO VISIBLE BUTTONS");
      } else {
        buttons.forEach(button => {
          console.log("");
          console.log(`BUTTON ${button.index}`);
          console.log(`tag: ${button.tag}`);
          console.log(`type: ${button.type}`);
          console.log(`text: ${button.text}`);
          console.log(`name: ${button.name}`);
          console.log(`id: ${button.id}`);
          console.log(`aria-label: ${button.ariaLabel}`);
          console.log(`HTML: ${button.outerHTML}`);
        });
      }
    } catch (error) {
      console.log(`Unable to inspect buttons: ${error.message}`);
    }
  }

  console.log("");
  console.log("============================================================");
  console.log("END DIAGNOSTIC INFORMATION");
  console.log("============================================================");
  console.log("");
}

async function findPostcodeInputInAnyFrame(page) {
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

  for (const frame of page.frames()) {
    for (const selector of selectors) {
      try {
        const locator = frame.locator(selector).first();

        if (await locator.count() && await locator.isVisible()) {
          return {
            frame,
            locator,
            selector
          };
        }
      } catch {
        // Continue.
      }
    }
  }

  return null;
}

async function findButtonInFrame(frame) {
  const selectors = [
    'button:has-text("Check")',
    'button:has-text("Search")',
    'button:has-text("Submit")',
    'input[type="submit"]',
    '[role="button"]:has-text("Check")',
    '[role="button"]:has-text("Search")',
    "button"
  ];

  for (const selector of selectors) {
    try {
      const locator = frame.locator(selector).first();

      if (await locator.count() && await locator.isVisible()) {
        return locator;
      }
    } catch {
      // Continue.
    }
  }

  return null;
}

async function main() {
  console.log("");
  console.log("============================================================");
  console.log("O2 NW9 0RY MONITOR");
  console.log("============================================================");
  console.log(`Postcode: ${POSTCODE}`);
  console.log(`Complaint: ${COMPLAINT_REFERENCE}`);
  console.log(`Source: ${O2_URL}`);
  console.log(`Time: ${londonTimestamp()}`);

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
  let checkResult = "FAILED";
  let notes = "";

  try {
    console.log("");
    console.log("Opening O2 status page...");

    await page.goto(O2_URL, {
      waitUntil: "domcontentloaded",
      timeout: 60000
    });

    await page.waitForTimeout(5000);

    console.log(`Loaded URL: ${page.url()}`);

    await dumpDiagnostics(page);

    console.log("");
    console.log("Searching all frames for postcode input...");

    const postcodeResult =
      await findPostcodeInputInAnyFrame(page);

    if (!postcodeResult) {
      throw new Error(
        "Could not find O2 postcode input in any frame."
      );
    }

    console.log("");
    console.log("POSTCODE INPUT FOUND");
    console.log(`Frame URL: ${postcodeResult.frame.url()}`);
    console.log(`Selector: ${postcodeResult.selector}`);

    await postcodeResult.locator.scrollIntoViewIfNeeded();

    /*
     * IMPORTANT:
     * O2 expects the postcode exactly as entered by the user.
     * We deliberately use a SPACE:
     *
     * NW9 0RY
     */
    await postcodeResult.locator.fill(POSTCODE);

    console.log(`Entered postcode: "${POSTCODE}"`);

    console.log("");
    console.log("Looking for submit button...");

    const submitButton =
      await findButtonInFrame(postcodeResult.frame);

    if (submitButton) {
      console.log(
        `Submit button found: "${await submitButton.innerText().catch(() => "")}"`
      );

      await submitButton.scrollIntoViewIfNeeded();

      await submitButton.click();
    } else {
      console.log("No submit button found.");
      console.log("Pressing Enter on postcode input.");

      await postcodeResult.locator.press("Enter");
    }

    console.log("");
    console.log("Postcode submitted.");

    await page.waitForTimeout(8000);

    console.log("");
    console.log("Checking resulting page...");

    /*
     * Dump diagnostics again AFTER the postcode submission.
     * This is important because O2's result may be inside
     * a dynamically created iframe.
     */
    await dumpDiagnostics(page);

    const allFrameTexts = [];

    for (const frame of page.frames()) {
      try {
        const text = await frame.locator("body").innerText();

        if (text && text.trim()) {
          allFrameTexts.push(
            `FRAME URL: ${frame.url()}\n${text}`
          );
        }
      } catch {
        // Continue.
      }
    }

    const combinedText = allFrameTexts.join("\n\n");

    console.log("");
    console.log("============================================================");
    console.log("O2 RESULT TEXT");
    console.log("============================================================");
    console.log(combinedText.slice(0, 10000));
    console.log("============================================================");

    status = classifyStatus(combinedText);

    message = combinedText
      .replace(/\s+/g, " ")
      .trim()
      .slice(0, 5000);

    const resolutionMatch = combinedText.match(
      /(?:updated|expected|estimated|should|aim).*?(?:resolved|fixed).*?(\d{1,2}[:.]\d{2}|\d{1,2}\s?(?:am|pm)|\d{1,2}\/\d{1,2}\/\d{2,4})/i
    );

    if (resolutionMatch) {
      expectedResolution = resolutionMatch[0];
    }

    checkResult = "SUCCESS";

    notes =
      "Successfully located the O2 postcode input, entered NW9 0RY with a space, submitted the check and captured the resulting page.";
  } catch (error) {
    status = "UNKNOWN";
    checkResult = "FAILED";

    message = `MONITORING ERROR: ${error.message}`;

    notes =
      "Automation failed. This check is not evidence of an O2 outage.";

    console.error("");
    console.error("============================================================");
    console.error("MONITOR ERROR");
    console.error("============================================================");
    console.error(error);
    console.error("============================================================");
  }

  const screenshotName =
    `${fileTimestamp}_NW9-0RY.png`;

  const screenshotPath =
    path.join(evidenceDir, screenshotName);

  console.log("");
  console.log(`Saving screenshot: ${screenshotPath}`);

  await page.screenshot({
    path: screenshotPath,
    fullPage: true
  });

  const existing = readExistingCsv();

  const previous =
    existing.length
      ? existing[existing.length - 1]
      : null;

  const statusChanged =
    !previous ||
    previous["O2 Status"] !== status
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
    "Notes": notes
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

  const allRecords = [
    ...existing,
    record
  ];

  const worksheet =
    XLSX.utils.json_to_sheet(allRecords);

  worksheet["!cols"] = [
    { wch: 14 },
    { wch: 12 },
    { wch: 24 },
    { wch: 14 },
    { wch: 18 },
    { wch: 80 },
    { wch: 30 },
    { wch: 20 },
    { wch: 40 },
    { wch: 55 },
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
      ["Complaint Reference", COMPLAINT_REFERENCE],
      ["Monitoring Source", O2_URL],
      ["Total Checks", allRecords.length],
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
          x => x["O2 Status"] === "ISSUE"
        ).length
      ],
      [
        "OK Checks",
        allRecords.filter(
          x => x["O2 Status"] === "OK"
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
        "Important",
        "SUCCESS means the monitor located the O2 postcode checker, entered NW9 0RY, submitted it and captured the resulting O2 page."
      ],
      [
        "Evidence",
        "Timestamped screenshots are stored in the evidence folder."
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

  console.log("");
  console.log("============================================================");
  console.log("FINAL RESULT");
  console.log("============================================================");
  console.log(`Timestamp: ${timestamp}`);
  console.log(`Postcode: ${POSTCODE}`);
  console.log(`Status: ${status}`);
  console.log(`Check Result: ${checkResult}`);
  console.log(`Screenshot: ${screenshotName}`);
  console.log("============================================================");

  await browser.close();
}

main().catch(error => {
  console.error("");
  console.error("FATAL MONITOR ERROR");
  console.error(error);
  process.exit(1);
});
