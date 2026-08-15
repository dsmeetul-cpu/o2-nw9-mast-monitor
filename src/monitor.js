import { chromium } from "playwright";
import fs from "fs";
import path from "path";

const POSTCODE = "NW9 0RY";
const O2_URL = "https://status.o2.co.uk/";

const root = process.cwd();
const evidenceDir = path.join(root, "evidence");

fs.mkdirSync(evidenceDir, { recursive: true });

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
    parts.find(p => p.type === name)?.value || "";

  return `${get("year")}-${get("month")}-${get("day")}_${get("hour")}-${get("minute")}-${get("second")}`;
}

async function main() {

  console.log("========================================");
  console.log("O2 NW9 DIAGNOSTIC MONITOR");
  console.log("========================================");
  console.log(`Time: ${londonTimestamp()}`);
  console.log(`Postcode: ${POSTCODE}`);
  console.log(`Starting URL: ${O2_URL}`);
  console.log("");

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

  /*
   * Log navigation.
   */
  page.on("framenavigated", frame => {
    if (frame === page.mainFrame()) {
      console.log(
        `[NAVIGATION] ${frame.url()}`
      );
    }
  });

  /*
   * Log console messages from the O2 page.
   */
  page.on("console", msg => {
    console.log(
      `[BROWSER CONSOLE] ${msg.type()}: ${msg.text()}`
    );
  });

  /*
   * Log failed network requests.
   */
  page.on("requestfailed", request => {
    console.log(
      `[REQUEST FAILED] ${request.method()} ${request.url()}`
    );

    console.log(
      `  Error: ${request.failure()?.errorText || "unknown"}`
    );
  });

  /*
   * Log HTTP errors.
   */
  page.on("response", response => {
    if (response.status() >= 400) {
      console.log(
        `[HTTP ${response.status()}] ${response.url()}`
      );
    }
  });

  try {

    console.log("");
    console.log("STEP 1: Opening O2 status page...");
    console.log("");

    await page.goto(O2_URL, {
      waitUntil: "domcontentloaded",
      timeout: 60000
    });

    await page.waitForTimeout(5000);

    console.log("");
    console.log("CURRENT PAGE");
    console.log("----------------------------------------");
    console.log(`URL: ${page.url()}`);
    console.log(`Title: ${await page.title()}`);
    console.log("");

    /*
     * PAGE TEXT
     */
    console.log("");
    console.log("VISIBLE PAGE TEXT");
    console.log("----------------------------------------");

    const bodyText =
      await page.locator("body").innerText();

    console.log(
      bodyText.slice(0, 15000)
    );

    /*
     * LINKS
     */
    console.log("");
    console.log("ALL VISIBLE LINKS");
    console.log("----------------------------------------");

    const links =
      await page.locator("a").evaluateAll(
        elements =>
          elements
            .filter(el => {
              const style =
                window.getComputedStyle(el);

              return (
                style.display !== "none" &&
                style.visibility !== "hidden"
              );
            })
            .map((el, index) => ({
              index,
              text:
                el.innerText
                  ?.replace(/\s+/g, " ")
                  .trim(),
              href:
                el.href || "",
              aria:
                el.getAttribute("aria-label") || "",
              title:
                el.getAttribute("title") || ""
            }))
      );

    links.forEach(link => {

      console.log(
        `[LINK ${link.index}]`
      );

      console.log(
        `  Text: ${link.text}`
      );

      console.log(
        `  Href: ${link.href}`
      );

      if (link.aria) {
        console.log(
          `  ARIA: ${link.aria}`
        );
      }

      if (link.title) {
        console.log(
          `  Title: ${link.title}`
        );
      }

      console.log("");
    });

    /*
     * INPUTS
     */
    console.log("");
    console.log("ALL INPUTS");
    console.log("----------------------------------------");

    const inputs =
      await page.locator("input").evaluateAll(
        elements =>
          elements.map((el, index) => ({
            index,
            type:
              el.getAttribute("type") || "",
            name:
              el.getAttribute("name") || "",
            id:
              el.getAttribute("id") || "",
            placeholder:
              el.getAttribute("placeholder") || "",
            aria:
              el.getAttribute("aria-label") || "",
            value:
              el.getAttribute("value") || "",
            autocomplete:
              el.getAttribute("autocomplete") || "",
            visible:
              !!(
                el.offsetWidth ||
                el.offsetHeight ||
                el.getClientRects().length
              )
          }))
      );

    if (inputs.length === 0) {
      console.log("NO INPUT ELEMENTS FOUND");
    }

    inputs.forEach(input => {
      console.log(
        `[INPUT ${input.index}]`
      );

      console.log(
        `  Type: ${input.type}`
      );

      console.log(
        `  Name: ${input.name}`
      );

      console.log(
        `  ID: ${input.id}`
      );

      console.log(
        `  Placeholder: ${input.placeholder}`
      );

      console.log(
        `  ARIA: ${input.aria}`
      );

      console.log(
        `  Autocomplete: ${input.autocomplete}`
      );

      console.log(
        `  Visible: ${input.visible}`
      );

      console.log("");
    });

    /*
     * BUTTONS
     */
    console.log("");
    console.log("ALL BUTTONS");
    console.log("----------------------------------------");

    const buttons =
      await page.locator("button").evaluateAll(
        elements =>
          elements.map((el, index) => ({
            index,
            text:
              el.innerText
                ?.replace(/\s+/g, " ")
                .trim(),
            type:
              el.getAttribute("type") || "",
            id:
              el.getAttribute("id") || "",
            name:
              el.getAttribute("name") || "",
            aria:
              el.getAttribute("aria-label") || "",
            disabled:
              el.disabled,
            visible:
              !!(
                el.offsetWidth ||
                el.offsetHeight ||
                el.getClientRects().length
              )
          }))
      );

    if (buttons.length === 0) {
      console.log("NO BUTTON ELEMENTS FOUND");
    }

    buttons.forEach(button => {

      console.log(
        `[BUTTON ${button.index}]`
      );

      console.log(
        `  Text: ${button.text}`
      );

      console.log(
        `  Type: ${button.type}`
      );

      console.log(
        `  ID: ${button.id}`
      );

      console.log(
        `  Name: ${button.name}`
      );

      console.log(
        `  ARIA: ${button.aria}`
      );

      console.log(
        `  Disabled: ${button.disabled}`
      );

      console.log(
        `  Visible: ${button.visible}`
      );

      console.log("");
    });

    /*
     * FORMS
     */
    console.log("");
    console.log("FORMS");
    console.log("----------------------------------------");

    const forms =
      await page.locator("form").evaluateAll(
        elements =>
          elements.map((el, index) => ({
            index,
            action:
              el.getAttribute("action") || "",
            method:
              el.getAttribute("method") || "",
            text:
              el.innerText
                ?.replace(/\s+/g, " ")
                .trim()
                .slice(0, 1000)
          }))
      );

    if (forms.length === 0) {
      console.log("NO FORMS FOUND");
    }

    forms.forEach(form => {

      console.log(
        `[FORM ${form.index}]`
      );

      console.log(
        `  Action: ${form.action}`
      );

      console.log(
        `  Method: ${form.method}`
      );

      console.log(
        `  Text: ${form.text}`
      );

      console.log("");
    });

    /*
     * IFRAMES
     */
    console.log("");
    console.log("IFRAMES");
    console.log("----------------------------------------");

    const frames =
      page.frames();

    frames.forEach((frame, index) => {

      console.log(
        `[FRAME ${index}]`
      );

      console.log(
        `  URL: ${frame.url()}`
      );

      console.log("");
    });

    /*
     * SAVE HTML
     */
    console.log("");
    console.log("SAVING HTML...");
    console.log("----------------------------------------");

    const html =
      await page.content();

    const htmlFile =
      path.join(
        evidenceDir,
        "diagnostic-o2-page.html"
      );

    fs.writeFileSync(
      htmlFile,
      html,
      "utf8"
    );

    console.log(
      `HTML saved: ${htmlFile}`
    );

    /*
     * SCREENSHOT
     */
    console.log("");
    console.log("SAVING SCREENSHOT...");
    console.log("----------------------------------------");

    const screenshotFile =
      path.join(
        evidenceDir,
        `diagnostic-${safeFileTimestamp()}.png`
      );

    await page.screenshot({
      path: screenshotFile,
      fullPage: true
    });

    console.log(
      `Screenshot saved: ${screenshotFile}`
    );

    /*
     * IMPORTANT:
     * We deliberately DO NOT submit the postcode.
     * This is diagnostic only.
     */
    console.log("");
    console.log("========================================");
    console.log("DIAGNOSTIC COMPLETE");
    console.log("========================================");
    console.log("");
    console.log(
      "No postcode was submitted."
    );
    console.log(
      "No O2 status was classified."
    );
    console.log(
      "No CSV record was created."
    );
    console.log(
      "No historical evidence was modified."
    );
    console.log("");

  } catch (error) {

    console.error("");
    console.error(
      "DIAGNOSTIC ERROR:"
    );

    console.error(
      error
    );

    /*
     * Capture whatever page state remains.
     */
    try {

      const emergencyScreenshot =
        path.join(
          evidenceDir,
          `diagnostic-error-${safeFileTimestamp()}.png`
        );

      await page.screenshot({
        path: emergencyScreenshot,
        fullPage: true
      });

      console.log(
        `Emergency screenshot: ${emergencyScreenshot}`
      );

    } catch {
      // Nothing else to do.
    }

    process.exitCode = 1;

  } finally {

    await browser.close();

  }
}

main();function parseCsvLine(line) {
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
  if (!fs.existsSync(csvFile)) return [];

  const text = fs.readFileSync(csvFile, "utf8").trim();

  if (!text) return [];

  const lines = text.split(/\r?\n/);

  const headers = parseCsvLine(lines[0]);

  return lines.slice(1).map(line => {
    const values = parseCsvLine(line);
    const row = {};

    headers.forEach((header, index) => {
      row[header] = values[index] ?? "";
    });

    return row;
  });
}

async function findNetworkCheckerLink(page) {
  const candidates = [
    'a:has-text("network status checker")',
    'a:has-text("network status")',
    'a:has-text("Network Issues")',
    'a:has-text("network checker")',
    'a[href*="network"]'
  ];

  for (const selector of candidates) {
    const locator = page.locator(selector);

    const count = await locator.count();

    for (let i = 0; i < count; i++) {
      const link = locator.nth(i);

      try {
        if (await link.isVisible()) {
          const href = await link.getAttribute("href");

          if (href) {
            return href;
          }
        }
      } catch {
        // Continue searching.
      }
    }
  }

  return null;
}

async function findPostcodeInput(page) {
  const candidates = [
    'input[placeholder*="postcode" i]',
    'input[aria-label*="postcode" i]',
    'input[name*="postcode" i]',
    'input[id*="postcode" i]',
    'input[placeholder*="post code" i]',
    'input[aria-label*="post code" i]',
    'input[name*="post code" i]',
    'input[id*="post code" i]',
    'input[type="text"]'
  ];

  for (const selector of candidates) {
    const locator = page.locator(selector);

    const count = await locator.count();

    for (let i = 0; i < count; i++) {
      const input = locator.nth(i);

      try {
        if (await input.isVisible()) {
          return input;
        }
      } catch {
        // Continue.
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
    'button:has-text("Find")',
    'button:has-text("Continue")',
    'input[type="submit"]',
    'button'
  ];

  for (const selector of candidates) {
    const locator = page.locator(selector);

    const count = await locator.count();

    for (let i = 0; i < count; i++) {
      const button = locator.nth(i);

      try {
        if (await button.isVisible()) {
          return button;
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

  const issuePatterns = [
    "issue",
    "problem",
    "outage",
    "not working",
    "fault",
    "maintenance",
    "engineer",
    "service disruption",
    "network disruption",
    "affected"
  ];

  const okPatterns = [
    "working normally",
    "no known",
    "no issues",
    "no issue",
    "no problems",
    "everything is working"
  ];

  if (issuePatterns.some(x => lower.includes(x))) {
    return "ISSUE";
  }

  if (okPatterns.some(x => lower.includes(x))) {
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
  let checkResult = "FAILED";

  try {
    console.log("Opening O2 status page...");

    await page.goto(O2_URL, {
      waitUntil: "domcontentloaded",
      timeout: 60000
    });

    await page.waitForTimeout(3000);

    console.log("Looking for O2 network status checker...");

    const checkerHref =
      await findNetworkCheckerLink(page);

    if (!checkerHref) {
      throw new Error(
        "Could not find O2 network status checker link."
      );
    }

    const checkerUrl =
      new URL(
        checkerHref,
        O2_URL
      ).href;

    console.log(
      `Network checker found: ${checkerUrl}`
    );

    await page.goto(checkerUrl, {
      waitUntil: "domcontentloaded",
      timeout: 60000
    });

    await page.waitForTimeout(5000);

    console.log(
      "Looking for postcode input..."
    );

    let postcodeInput =
      await findPostcodeInput(page);

    if (!postcodeInput) {
      /*
       * Some modern O2 journeys may load
       * their content dynamically.
       */

      await page.waitForTimeout(5000);

      postcodeInput =
        await findPostcodeInput(page);
    }

    if (!postcodeInput) {
      throw new Error(
        "Could not find postcode input on O2 network status checker."
      );
    }

    console.log(
      "Postcode input found."
    );

    await postcodeInput.fill(POSTCODE);

    const submitButton =
      await findSubmitButton(page);

    if (submitButton) {
      console.log(
        "Submitting postcode..."
      );

      await submitButton.click();
    } else {
      console.log(
        "No submit button found; pressing Enter..."
      );

      await postcodeInput.press("Enter");
    }

    await page.waitForTimeout(8000);

    const pageText =
      await page.locator("body").innerText();

    const cleanText =
      pageText
        .replace(/\s+/g, " ")
        .trim();

    status =
      classifyStatus(cleanText);

    message =
      cleanText.slice(0, 5000);

    const resolutionMatch =
      pageText.match(
        /(?:expected|estimated|should|aim).*?(?:resolved|fixed).*?(\d{1,2}[:.]\d{2}|\d{1,2}\s?(?:am|pm)|\d{1,2}\/\d{1,2}\/\d{2,4})/i
      );

    if (resolutionMatch) {
      expectedResolution =
        resolutionMatch[0];
    }

    checkResult = "SUCCESS";

    console.log(
      `O2 status classified as: ${status}`
    );

  } catch (error) {
    status = "UNKNOWN";

    message =
      `MONITORING ERROR: ${error.message}`;

    checkResult = "FAILED";

    console.error(message);
  }

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

  const existing =
    readExistingCsv();

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
    "Date":
      timestamp.split(",")[0],

    "Time":
      timestamp.split(",")[1]?.trim()
      || timestamp,

    "Date & Time":
      timestamp,

    "Postcode":
      POSTCODE,

    "O2 Status":
      status,

    "O2 Message":
      message,

    "Expected Resolution":
      expectedResolution,

    "Complaint Reference":
      COMPLAINT_REFERENCE,

    "Source URL":
      O2_URL,

    "Screenshot":
      `evidence/${screenshotName}`,

    "Check Result":
      checkResult,

    "Status Changed?":
      statusChanged,

    "Previous Status":
      previous?.["O2 Status"] || "",

    "Notes":
      ""
  };

  const headers =
    Object.keys(record);

  if (!fs.existsSync(csvFile)) {
    fs.writeFileSync(
      csvFile,
      headers
        .map(csvEscape)
        .join(",") +
        "\n",
      "utf8"
    );
  }

  fs.appendFileSync(
    csvFile,
    headers
      .map(h => csvEscape(record[h]))
      .join(",") +
      "\n",
    "utf8"
  );

  const allRecords =
    [...existing, record];

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
      [
        "Failed Checks",
        allRecords.filter(
          x => x["Check Result"] === "FAILED"
        ).length
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

  XLSX.writeFile(
    workbook,
    xlsxFile
  );

  await browser.close();

  console.log(
    JSON.stringify(
      {
        timestamp,
        postcode: POSTCODE,
        status,
        screenshot: screenshotName,
        checkResult
      },
      null,
      2
    )
  );
}

main().catch(error => {
  console.error(error);
  process.exit(1);
});  const now = new Date();

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
