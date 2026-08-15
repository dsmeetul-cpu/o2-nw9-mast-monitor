import { chromium } from "playwright";
import fs from "fs";
import path from "path";
import XLSX from "xlsx";

const POSTCODE = "NW9 0RY";
const COMPLAINT_REFERENCE = "C-1308267357";

const O2_STATUS_URL = "https://status.o2.co.uk/";

const root = process.cwd();

const dataDir = path.join(root, "data");
const evidenceDir = path.join(root, "evidence");

fs.mkdirSync(dataDir, { recursive: true });
fs.mkdirSync(evidenceDir, { recursive: true });

const csvFile =
  path.join(dataDir, "o2-nw9-0ry-monitor.csv");

const xlsxFile =
  path.join(dataDir, "O2-NW9-0RY-Evidence.xlsx");

/*
===========================================================
TIME
===========================================================
*/

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

  const parts =
    new Intl.DateTimeFormat("en-GB", {
      timeZone: "Europe/London",
      year: "numeric",
      month: "2-digit",
      day: "2-digit",
      hour: "2-digit",
      minute: "2-digit",
      second: "2-digit",
      hour12: false
    }).formatToParts(now);

  const get = name =>
    parts.find(
      part => part.type === name
    )?.value || "";

  return (
    `${get("year")}-${get("month")}-${get("day")}` +
    `_${get("hour")}-${get("minute")}-${get("second")}`
  );

}

/*
===========================================================
CSV
===========================================================
*/

function csvEscape(value) {

  if (
    value === null ||
    value === undefined
  ) {
    return "";
  }

  const text =
    String(value)
      .replace(/\r?\n/g, " ");

  if (/[",]/.test(text)) {

    return (
      '"' +
      text.replace(/"/g, '""') +
      '"'
    );

  }

  return text;

}

function parseCSVLine(line) {

  const result = [];

  let current = "";

  let quoted = false;

  for (
    let i = 0;
    i < line.length;
    i++
  ) {

    const char = line[i];

    if (char === '"') {

      if (
        quoted &&
        line[i + 1] === '"'
      ) {

        current += '"';

        i++;

      } else {

        quoted = !quoted;

      }

    } else if (
      char === "," &&
      !quoted
    ) {

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

  const text =
    fs.readFileSync(
      csvFile,
      "utf8"
    ).trim();

  if (!text) {
    return [];
  }

  const lines =
    text.split(/\r?\n/);

  const headers =
    parseCSVLine(lines[0]);

  return lines
    .slice(1)
    .map(line => {

      const values =
        parseCSVLine(line);

      const row = {};

      headers.forEach(
        (header, index) => {

          row[header] =
            values[index] ?? "";

        }
      );

      return row;

    });

}

/*
===========================================================
HELPERS
===========================================================
*/

async function visibleLocator(
  page,
  selectors
) {

  for (const selector of selectors) {

    try {

      const locator =
        page.locator(selector);

      const count =
        await locator.count();

      for (
        let i = 0;
        i < count;
        i++
      ) {

        const candidate =
          locator.nth(i);

        if (
          await candidate.isVisible()
        ) {

          return candidate;

        }

      }

    } catch {
      // Continue.
    }

  }

  return null;

}

/*
===========================================================
FIND POSTCODE FIELD
===========================================================
*/

async function findPostcodeInput(page) {

  const selectors = [

    'input[autocomplete="postal-code"]',

    'input[autocomplete="postcode"]',

    'input[name*="postcode" i]',

    'input[id*="postcode" i]',

    'input[name*="post-code" i]',

    'input[id*="post-code" i]',

    'input[placeholder*="postcode" i]',

    'input[placeholder*="post code" i]',

    'input[aria-label*="postcode" i]',

    'input[aria-label*="post code" i]',

    'input[type="text"]',

    'input:not([type])'

  ];

  return visibleLocator(
    page,
    selectors
  );

}

/*
===========================================================
FIND O2 CHECKER LINK
===========================================================
*/

async function findNetworkCheckerLink(page) {

  const links =
    page.locator("a");

  const count =
    await links.count();

  for (
    let i = 0;
    i < count;
    i++
  ) {

    const link =
      links.nth(i);

    try {

      if (
        !(await link.isVisible())
      ) {
        continue;
      }

      const text =
        (
          await link.innerText()
        )
          .replace(/\s+/g, " ")
          .trim()
          .toLowerCase();

      const href =
        await link.getAttribute("href");

      if (
        text.includes(
          "network status checker"
        )
      ) {

        return {
          link,
          href
        };

      }

      if (
        text.includes(
          "check our network"
        )
      ) {

        return {
          link,
          href
        };

      }

      if (
        text.includes(
          "check network"
        )
      ) {

        return {
          link,
          href
        };

      }

    } catch {
      // Continue.
    }

  }

  return null;

}

/*
===========================================================
FIND SUBMIT BUTTON
===========================================================
*/

async function findSubmitButton(page) {

  const selectors = [

    'button:has-text("Check")',

    'button:has-text("Check now")',

    'button:has-text("Check status")',

    'button:has-text("Search")',

    'button:has-text("Submit")',

    'input[type="submit"]',

    'button'

  ];

  return visibleLocator(
    page,
    selectors
  );

}

/*
===========================================================
ENTER POSTCODE
===========================================================
*/

async function enterPostcode(
  page,
  input
) {

  /*
   * IMPORTANT:
   *
   * O2 postcode must be:
   *
   * NW9 0RY
   *
   * There is a REAL SPACE between NW9 and 0RY.
   */

  const expected =
    "NW9 0RY";

  await input.scrollIntoViewIfNeeded();

  await input.click();

  await input.fill("");

  /*
   * Use fill first.
   */

  await input.fill(expected);

  let actual =
    await input.inputValue();

  console.log(
    "O2 postcode field after fill:",
    JSON.stringify(actual)
  );

  /*
   * If O2 has transformed it,
   * explicitly replace the value.
   */

  if (actual !== expected) {

    await input.evaluate(
      (element, value) => {

        const setter =
          Object.getOwnPropertyDescriptor(
            HTMLInputElement.prototype,
            "value"
          ).set;

        setter.call(
          element,
          value
        );

        element.dispatchEvent(
          new Event(
            "input",
            {
              bubbles: true
            }
          )
        );

        element.dispatchEvent(
          new Event(
            "change",
            {
              bubbles: true
            }
          )
        );

      },
      expected
    );

  }

  actual =
    await input.inputValue();

  console.log(
    "O2 postcode field final value:",
    JSON.stringify(actual)
  );

  /*
   * Absolute verification.
   */

  if (actual !== expected) {

    throw new Error(
      `Postcode entry failed. Expected "${expected}" but O2 field contains "${actual}".`
    );

  }

  /*
   * Screenshot proving the postcode
   * has actually been entered.
   */

  return actual;

}

/*
===========================================================
STATUS CLASSIFICATION
===========================================================
*/

function classifyStatus(text) {

  const lower =
    text.toLowerCase();

  /*
   * Strong O2 mast/outage wording.
   */

  const issuePatterns = [

    "nearby phone mast",

    "mast isn’t working",

    "mast isn't working",

    "mast is not working",

    "mast not working",

    "network issue",

    "network issues",

    "network problem",

    "network problems",

    "service is affected",

    "service may be affected",

    "service might come and go",

    "engineers will be on the case",

    "engineers are on the case",

    "known issue",

    "known problem",

    "outage",

    "fault"

  ];

  for (
    const phrase of issuePatterns
  ) {

    if (
      lower.includes(phrase)
    ) {

      return "ISSUE";

    }

  }

  /*
   * Normal service wording.
   */

  const okPatterns = [

    "no known issues",

    "no known issue",

    "working normally",

    "there are no issues",

    "there is no issue",

    "no network issues"

  ];

  for (
    const phrase of okPatterns
  ) {

    if (
      lower.includes(phrase)
    ) {

      return "OK";

    }

  }

  return "UNKNOWN";

}

/*
===========================================================
EXTRACT O2 MESSAGE
===========================================================
*/

function extractO2Message(text) {

  const cleaned =
    text
      .replace(/\s+/g, " ")
      .trim();

  const patterns = [

    /Looks like a nearby phone mast.*?(?=(?:Updated|Just so you know|$))/i,

    /Our engineers.*?(?=(?:Updated|Just so you know|$))/i,

    /There (?:is|are) .*?(?:issue|problem|fault).*?(?=(?:Updated|$))/i

  ];

  for (
    const pattern of patterns
  ) {

    const match =
      cleaned.match(pattern);

    if (match) {

      return match[0].trim();

    }

  }

  return cleaned.slice(
    0,
    5000
  );

}

/*
===========================================================
DIAGNOSTIC PAGE INFORMATION
===========================================================
*/

async function getPageDiagnostics(page) {

  const inputs =
    await page.locator("input")
      .evaluateAll(elements =>
        elements.map(
          element => ({
            type:
              element.getAttribute(
                "type"
              ),
            name:
              element.getAttribute(
                "name"
              ),
            id:
              element.getAttribute(
                "id"
              ),
            placeholder:
              element.getAttribute(
                "placeholder"
              ),
            aria:
              element.getAttribute(
                "aria-label"
              ),
            autocomplete:
              element.getAttribute(
                "autocomplete"
              ),
            value:
              element.value
          })
        )
      );

  const buttons =
    await page.locator(
      "button"
    )
      .evaluateAll(elements =>
        elements.map(
          element => ({
            text:
              (
                element.innerText ||
                ""
              )
                .replace(/\s+/g, " ")
                .trim(),
            type:
              element.getAttribute(
                "type"
              ),
            aria:
              element.getAttribute(
                "aria-label"
              )
          })
        )
      );

  return {
    url:
      page.url(),

    title:
      await page.title(),

    inputs,

    buttons

  };

}

/*
===========================================================
MAIN
===========================================================
*/

async function main() {

  const browser =
    await chromium.launch({
      headless: true
    });

  const context =
    await browser.newContext({
      locale: "en-GB",
      timezoneId:
        "Europe/London",
      viewport: {
        width: 1440,
        height: 1200
      }
    });

  const page =
    await context.newPage();

  /*
   * Avoid browser dialogs stopping automation.
   */

  page.on(
    "dialog",
    async dialog => {

      console.log(
        "Browser dialog:",
        dialog.message()
      );

      await dialog.dismiss();

    }
  );

  const timestamp =
    londonTimestamp();

  const fileTimestamp =
    safeFileTimestamp();

  let status =
    "UNKNOWN";

  let message =
    "";

  let expectedResolution =
    "";

  let checkResult =
    "FAILED";

  let pageText =
    "";

  let postcodeEntered =
    false;

  let checkerSubmitted =
    false;

  let diagnostics =
    "";

  const preEntryScreenshot =
    `${fileTimestamp}_NW9-0RY-PRE-ENTRY.png`;

  const finalScreenshot =
    `${fileTimestamp}_NW9-0RY.png`;

  const preEntryPath =
    path.join(
      evidenceDir,
      preEntryScreenshot
    );

  const finalScreenshotPath =
    path.join(
      evidenceDir,
      finalScreenshot
    );

  try {

    /*
     * =====================================================
     * STEP 1
     * Open O2 service status.
     * =====================================================
     */

    console.log(
      "Opening O2 service status:",
      O2_STATUS_URL
    );

    await page.goto(
      O2_STATUS_URL,
      {
        waitUntil:
          "domcontentloaded",
        timeout:
          60000
      }
    );

    await page.waitForTimeout(
      5000
    );

    console.log(
      "Initial O2 URL:",
      page.url()
    );

    console.log(
      "Initial O2 title:",
      await page.title()
    );

    /*
     * =====================================================
     * STEP 2
     * Look for postcode field on current page.
     * =====================================================
     */

    let postcodeInput =
      await findPostcodeInput(page);

    /*
     * =====================================================
     * STEP 3
     * If not found, follow O2's network checker link.
     * =====================================================
     */

    if (!postcodeInput) {

      console.log(
        "Postcode field not found on initial O2 page."
      );

      const checkerLink =
        await findNetworkCheckerLink(
          page
        );

      if (!checkerLink) {

        diagnostics =
          JSON.stringify(
            await getPageDiagnostics(
              page
            ),
            null,
            2
          );

        throw new Error(
          "Could not find O2 postcode checker or network status link."
        );

      }

      console.log(
        "Found O2 network checker link:",
        checkerLink.href
      );

      /*
       * Prefer clicking the actual link.
       */

      await checkerLink.link.click();

      await page.waitForLoadState(
        "domcontentloaded"
      )
        .catch(() => {});

      await page.waitForTimeout(
        5000
      );

      console.log(
        "Network checker URL:",
        page.url()
      );

      postcodeInput =
        await findPostcodeInput(
          page
        );

    }

    /*
     * =====================================================
     * STEP 4
     * Some modern O2 pages render the form late.
     * Wait and retry.
     * =====================================================
     */

    if (!postcodeInput) {

      console.log(
        "Waiting for O2 postcode field..."
      );

      try {

        await page.waitForFunction(
          () => {

            const inputs =
              Array.from(
                document.querySelectorAll(
                  "input"
                )
              );

            return inputs.some(
              input => {

                const value =
                  (
                    input.getAttribute(
                      "name"
                    ) ||
                    ""
                  ).toLowerCase();

                const id =
                  (
                    input.getAttribute(
                      "id"
                    ) ||
                    ""
                  ).toLowerCase();

                const placeholder =
                  (
                    input.getAttribute(
                      "placeholder"
                    ) ||
                    ""
                  ).toLowerCase();

                return (
                  value.includes(
                    "postcode"
                  ) ||
                  id.includes(
                    "postcode"
                  ) ||
                  placeholder.includes(
                    "postcode"
                  ) ||
                  value.includes(
                    "post-code"
                  ) ||
                  id.includes(
                    "post-code"
                  )
                );

              }
            );

          },
          {
            timeout:
              20000
          }
        );

      } catch {
        // Retry with normal selectors.
      }

      postcodeInput =
        await findPostcodeInput(
          page
        );

    }

    /*
     * =====================================================
     * STEP 5
     * If still unavailable, inspect every input.
     * =====================================================
     */

    if (!postcodeInput) {

      const diagnosticData =
        await getPageDiagnostics(
          page
        );

      diagnostics =
        JSON.stringify(
          diagnosticData,
          null,
          2
        );

      /*
       * Save diagnostics.
       */

      const diagnosticFile =
        path.join(
          evidenceDir,
          `${fileTimestamp}_diagnostics.json`
        );

      fs.writeFileSync(
        diagnosticFile,
        diagnostics,
        "utf8"
      );

      throw new Error(
        "Could not find O2 postcode input. Diagnostic information saved."
      );

    }

    /*
     * =====================================================
     * STEP 6
     * ENTER NW9 0RY EXACTLY.
     * =====================================================
     */

    console.log(
      "Entering postcode:",
      JSON.stringify(
        POSTCODE
      )
    );

    await enterPostcode(
      page,
      postcodeInput
    );

    postcodeEntered =
      true;

    /*
     * =====================================================
     * STEP 7
     * SCREENSHOT AFTER POSTCODE ENTRY.
     *
     * This is important evidence.
     * =====================================================
     */

    await page.screenshot({
      path:
        preEntryPath,
      fullPage:
        true
    });

    console.log(
      "Pre-submission screenshot:",
      preEntryScreenshot
    );

    /*
     * =====================================================
     * STEP 8
     * Find O2 submit button.
     * =====================================================
     */

    const submitButton =
      await findSubmitButton(
        page
      );

    if (!submitButton) {

      throw new Error(
        "Postcode was entered successfully, but the O2 check button could not be found."
      );

    }

    console.log(
      "Submitting O2 postcode check."
    );

    await submitButton.scrollIntoViewIfNeeded();

    await submitButton.click();

    checkerSubmitted =
      true;

    /*
     * =====================================================
     * STEP 9
     * Wait for O2 result.
     * =====================================================
     */

    await page.waitForTimeout(
      8000
    );

    /*
     * Wait for page/network activity to settle.
     */

    await page.waitForLoadState(
      "networkidle",
      {
        timeout:
          15000
      }
    )
      .catch(() => {});

    await page.waitForTimeout(
      3000
    );

    /*
     * =====================================================
     * STEP 10
     * Read resulting O2 page.
     * =====================================================
     */

    pageText =
      await page.locator(
        "body"
      ).innerText();

    const cleanedText =
      pageText
        .replace(/\s+/g, " ")
        .trim();

    console.log(
      "O2 result URL:",
      page.url()
    );

    console.log(
      "O2 result text:",
      cleanedText.slice(
        0,
        3000
      )
    );

    /*
     * =====================================================
     * STEP 11
     * CLASSIFY RESULT.
     * =====================================================
     */

    status =
      classifyStatus(
        cleanedText
      );

    message =
      extractO2Message(
        cleanedText
      );

    /*
     * =====================================================
     * Expected resolution.
     * =====================================================
     */

    const resolutionPatterns = [

      /Updated\s+([0-9]{1,2}:[0-9]{2})/i,

      /updated\s+([0-9]{1,2}:[0-9]{2})/i,

      /expected.*?([0-9]{1,2}:[0-9]{2})/i,

      /estimated.*?([0-9]{1,2}:[0-9]{2})/i

    ];

    for (
      const pattern of resolutionPatterns
    ) {

      const match =
        cleanedText.match(
          pattern
        );

      if (match) {

        expectedResolution =
          match[0];

        break;

      }

    }

    /*
     * =====================================================
     * SUCCESS ONLY IF:
     *
     * 1. Postcode actually entered
     * 2. Submit button clicked
     * 3. O2 returned a page
     *
     * UNKNOWN can still be SUCCESS because it means
     * the O2 check completed but classification needs
     * review.
     * =====================================================
     */

    if (
      postcodeEntered &&
      checkerSubmitted &&
      cleanedText.length > 0
    ) {

      checkResult =
        "SUCCESS";

    }

    /*
     * =====================================================
     * STEP 12
     * FINAL RESULT SCREENSHOT.
     *
     * THIS IS THE PRIMARY EVIDENCE.
     * =====================================================
     */

    await page.screenshot({
      path:
        finalScreenshotPath,
      fullPage:
        true
    });

    console.log(
      "Final evidence screenshot:",
      finalScreenshot
    );

  } catch (error) {

    message =
      `MONITORING ERROR: ${error.message}`;

    checkResult =
      "FAILED";

    console.error(
      message
    );

    /*
     * Always capture the current page
     * even if automation failed.
     */

    try {

      await page.screenshot({
        path:
          finalScreenshotPath,
        fullPage:
          true
      });

    } catch {
      // Nothing else to do.
    }

  }

  /*
   * =======================================================
   * EXISTING RECORDS
   * =======================================================
   */

  const existing =
    readExistingCsv();

  const previous =
    existing.length
      ? existing[
          existing.length - 1
        ]
      : null;

  const statusChanged =
    !previous ||
    previous["O2 Status"] !== status
      ? "YES"
      : "NO";

  /*
   * =======================================================
   * RECORD
   * =======================================================
   */

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
      page.url(),

    "Screenshot":
      `evidence/${finalScreenshot}`,

    "Check Result":
      checkResult,

    "Status Changed?":
      statusChanged,

    "Previous Status":
      previous?.["O2 Status"]
      || "",

    "Notes":
      postcodeEntered
        ? (
            checkerSubmitted
              ? "Postcode entered as NW9 0RY and O2 checker submitted."
              : "Postcode entered as NW9 0RY but O2 checker was not submitted."
          )
        : "Postcode was not entered."

  };

  /*
   * =======================================================
   * WRITE CSV
   * =======================================================
   */

  const headers =
    Object.keys(record);

  if (
    !fs.existsSync(csvFile)
  ) {

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
      .map(
        header =>
          csvEscape(
            record[header]
          )
      )
      .join(",") +
      "\n",
    "utf8"
  );

  /*
   * =======================================================
   * XLSX
   * =======================================================
   */

  const allRecords =
    [
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
    { wch: 20 },
    { wch: 45 },
    { wch: 65 },
    { wch: 18 },
    { wch: 18 },
    { wch: 18 },
    { wch: 70 }

  ];

  const workbook =
    XLSX.utils.book_new();

  XLSX.utils.book_append_sheet(
    workbook,
    worksheet,
    "O2 Mast Log"
  );

  /*
   * =======================================================
   * SUMMARY SHEET
   * =======================================================
   */

  const issueCount =
    allRecords.filter(
      row =>
        row["O2 Status"] ===
        "ISSUE" &&
        row["Check Result"] ===
        "SUCCESS"
    ).length;

  const okCount =
    allRecords.filter(
      row =>
        row["O2 Status"] ===
        "OK" &&
        row["Check Result"] ===
        "SUCCESS"
    ).length;

  const unknownCount =
    allRecords.filter(
      row =>
        row["O2 Status"] ===
        "UNKNOWN"
    ).length;

  const failedCount =
    allRecords.filter(
      row =>
        row["Check Result"] ===
        "FAILED"
    ).length;

  const summary =
    XLSX.utils.aoa_to_sheet([

      [
        "O2 NW9 0RY NETWORK EVIDENCE"
      ],

      [],

      [
        "Postcode",
        POSTCODE
      ],

      [
        "Complaint Reference",
        COMPLAINT_REFERENCE
      ],

      [
        "Monitoring Source",
        O2_STATUS_URL
      ],

      [
        "Total Checks",
        allRecords.length
      ],

      [
        "Confirmed Issue Checks",
        issueCount
      ],

      [
        "Confirmed OK Checks",
        okCount
      ],

      [
        "Unknown Checks",
        unknownCount
      ],

      [
        "Failed Checks",
        failedCount
      ],

      [],

      [
        "Important",
        "SUCCESS means the postcode was entered into O2 and the checker was submitted."
      ],

      [
        "Evidence",
        "The final timestamped O2 screenshot is the primary evidence."
      ],

      [
        "Postcode",
        "The monitor explicitly verifies the exact value NW9 0RY, including the space."
      ]

    ]);

  summary["!cols"] = [
    {
      wch: 30
    },
    {
      wch: 110
    }
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

  /*
   * =======================================================
   * CLOSE
   * =======================================================
   */

  await browser.close();

  /*
   * =======================================================
   * FINAL LOG
   * =======================================================
   */

  console.log(
    JSON.stringify(
      {
        timestamp,
        postcode:
          POSTCODE,
        postcodeEntered,
        checkerSubmitted,
        status,
        checkResult,
        screenshot:
          finalScreenshot,
        preEntryScreenshot
      },
      null,
      2
    )
  );

}

/*
===========================================================
START
===========================================================
*/

main().catch(
  error => {

    console.error(
      "Fatal monitor error:",
      error
    );

    process.exit(1);

  }
);  }

  return text;
}

function parseCSVLine(line) {
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

  const headers = parseCSVLine(lines[0]);

  return lines.slice(1).map(line => {
    const values = parseCSVLine(line);
    const row = {};

    headers.forEach((header, index) => {
      row[header] = values[index] ?? "";
    });

    return row;
  });
}

/*
 * O2's current public status page is dynamic.
 *
 * Instead of assuming one fixed input selector, inspect every
 * visible input and textarea and identify one that behaves like
 * a postcode field.
 */
async function findPostcodeInput(page) {
  const selectors = [
    'input',
    'textarea'
  ];

  for (const selector of selectors) {
    const locators = page.locator(selector);

    const count = await locators.count();

    for (let i = 0; i < count; i++) {
      const input = locators.nth(i);

      try {
        if (!(await input.isVisible())) {
          continue;
        }

        const type = (await input.getAttribute("type") || "").toLowerCase();
        const name = (await input.getAttribute("name") || "").toLowerCase();
        const id = (await input.getAttribute("id") || "").toLowerCase();
        const placeholder =
          (await input.getAttribute("placeholder") || "").toLowerCase();
        const aria =
          (await input.getAttribute("aria-label") || "").toLowerCase();

        const description =
          `${type} ${name} ${id} ${placeholder} ${aria}`;

        if (
          description.includes("postcode") ||
          description.includes("post code") ||
          description.includes("postal")
        ) {
          return input;
        }
      } catch {
        // Continue searching.
      }
    }
  }

  /*
   * If O2 has changed the accessibility attributes again,
   * use visible text around the form as a final fallback.
   */
  const text = (await page.locator("body").innerText()).toLowerCase();

  if (
    text.includes("postcode") ||
    text.includes("post code")
  ) {
    const inputs = page.locator("input:visible");

    const count = await inputs.count();

    for (let i = 0; i < count; i++) {
      const input = inputs.nth(i);

      const type = (
        await input.getAttribute("type") || "text"
      ).toLowerCase();

      if (
        type === "text" ||
        type === "search" ||
        type === ""
      ) {
        return input;
      }
    }
  }

  return null;
}

async function findCheckButton(page) {
  const selectors = [
    'button:has-text("Check")',
    'button:has-text("check")',
    'button:has-text("Search")',
    'button:has-text("search")',
    'button:has-text("Submit")',
    'input[type="submit"]',
    'button'
  ];

  for (const selector of selectors) {
    const buttons = page.locator(selector);

    const count = await buttons.count();

    for (let i = 0; i < count; i++) {
      const button = buttons.nth(i);

      try {
        if (await button.isVisible()) {
          const text = (
            await button.innerText().catch(() => "")
          ).trim().toLowerCase();

          const aria = (
            await button.getAttribute("aria-label") || ""
          ).toLowerCase();

          const combined = `${text} ${aria}`;

          if (
            combined.includes("check") ||
            combined.includes("search") ||
            combined.includes("submit") ||
            selector === "button"
          ) {
            return button;
          }
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

  /*
   * These are intentionally checked first because this is the
   * exact type of O2 outage wording we want to capture.
   */
  if (
    lower.includes("nearby phone mast") ||
    lower.includes("mast isn't working") ||
    lower.includes("mast isn't working as it should") ||
    lower.includes("service might come and go") ||
    lower.includes("engineers will be on the case") ||
    lower.includes("known issue") ||
    lower.includes("network issue") ||
    lower.includes("network fault") ||
    lower.includes("outage") ||
    lower.includes("maintenance")
  ) {
    return "ISSUE";
  }

  if (
    lower.includes("working normally") ||
    lower.includes("no known issues") ||
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
      height: 1600
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
    console.log(`Opening O2 status page: ${O2_STATUS_URL}`);

    await page.goto(O2_STATUS_URL, {
      waitUntil: "domcontentloaded",
      timeout: 60000
    });

    await page.waitForTimeout(5000);

    /*
     * Log the URL because O2 may redirect the checker.
     */
    console.log(`Current URL: ${page.url()}`);

    /*
     * First try to locate the postcode field.
     */
    let postcodeInput = await findPostcodeInput(page);

    /*
     * O2 can expose the network checker through a link.
     * If the field isn't on the landing page, click the
     * network checker link and try again.
     */
    if (!postcodeInput) {
      const links = page.locator("a");

      const count = await links.count();

      for (let i = 0; i < count; i++) {
        const link = links.nth(i);

        try {
          if (!(await link.isVisible())) {
            continue;
          }

          const text = (
            await link.innerText().catch(() => "")
          ).trim().toLowerCase();

          const href =
            (await link.getAttribute("href") || "").toLowerCase();

          if (
            text.includes("network status") ||
            text.includes("network checker") ||
            text.includes("check our network") ||
            href.includes("network")
          ) {
            console.log(
              `Opening network checker link: ${await link.getAttribute("href")}`
            );

            await link.click();

            await page.waitForLoadState("domcontentloaded")
              .catch(() => {});

            await page.waitForTimeout(5000);

            postcodeInput = await findPostcodeInput(page);

            if (postcodeInput) {
              break;
            }
          }
        } catch {
          // Continue.
        }
      }
    }

    /*
     * Last attempt: inspect iframes.
     */
    if (!postcodeInput) {
      for (const frame of page.frames()) {
        try {
          const inputs = frame.locator("input:visible");

          const count = await inputs.count();

          for (let i = 0; i < count; i++) {
            const input = inputs.nth(i);

            const attrs = [
              await input.getAttribute("name"),
              await input.getAttribute("id"),
              await input.getAttribute("placeholder"),
              await input.getAttribute("aria-label")
            ]
              .filter(Boolean)
              .join(" ")
              .toLowerCase();

            if (
              attrs.includes("postcode") ||
              attrs.includes("post code") ||
              attrs.includes("postal")
            ) {
              postcodeInput = input;
              break;
            }
          }

          if (postcodeInput) {
            break;
          }
        } catch {
          // Continue.
        }
      }
    }

    if (!postcodeInput) {
      /*
       * IMPORTANT:
       * Capture the actual O2 page even when the selector fails.
       * This means the evidence shows exactly what O2 presented.
       */
      throw new Error(
        `Could not find O2 postcode input. Current URL: ${page.url()}`
      );
    }

    console.log("Postcode input found.");

    await postcodeInput.scrollIntoViewIfNeeded();

    await postcodeInput.click();

    await postcodeInput.fill("");

    await postcodeInput.fill(POSTCODE);

    /*
     * Verify that Playwright actually entered the postcode.
     */
    const enteredValue = await postcodeInput.inputValue();

    console.log(`Postcode entered: ${enteredValue}`);

    if (
      enteredValue.replace(/\s+/g, "").toUpperCase() !==
      POSTCODE.replace(/\s+/g, "").toUpperCase()
    ) {
      throw new Error(
        `Postcode was not entered correctly. Value returned: "${enteredValue}"`
      );
    }

    const checkButton = await findCheckButton(page);

    if (checkButton) {
      console.log("Check button found.");

      await checkButton.scrollIntoViewIfNeeded();

      await checkButton.click();

      console.log("Check button clicked.");
    } else {
      console.log("No check button found. Pressing Enter.");

      await postcodeInput.press("Enter");
    }

    /*
     * Allow O2's result to load.
     */
    await page.waitForTimeout(10000);

    /*
     * Capture the complete page after the postcode check.
     */
    pageText = await page.locator("body").innerText();

    console.log("===== O2 RESULT =====");
    console.log(pageText);
    console.log("=====================");

    status = classifyStatus(pageText);

    message = pageText
      .replace(/\s+/g, " ")
      .trim()
      .slice(0, 5000);

    const resolutionMatch = pageText.match(
      /(?:expected|estimated|should|aim|updated).*?(?:resolved|fixed|back|working).*?(\d{1,2}[:.]\d{2}|\d{1,2}\s?(?:am|pm)|\d{1,2}\/\d{1,2}\/\d{2,4})/i
    );

    if (resolutionMatch) {
      expectedResolution = resolutionMatch[0];
    }

    /*
     * A successful screenshot is only considered successful
     * when we actually got a result after entering the postcode.
     */
    checkResult = "SUCCESS";
  } catch (error) {
    message = `MONITORING ERROR: ${error.message}`;

    console.error(message);
  }

  /*
   * Screenshot is ALWAYS taken.
   *
   * This is important for your complaint evidence:
   * even if O2 changes its page structure, we retain what
   * O2 actually showed at the time of the check.
   */
  const screenshotName =
    `${fileTimestamp}_NW9-0RY.png`;

  const screenshotPath =
    path.join(evidenceDir, screenshotName);

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
    "Source URL": O2_STATUS_URL,
    "Screenshot": `evidence/${screenshotName}`,
    "Check Result": checkResult,
    "Status Changed?": statusChanged,
    "Previous Status": previous?.["O2 Status"] || "",
    "Notes":
      status === "ISSUE"
        ? "O2 network issue detected after postcode check."
        : ""
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
    { wch: 100 },
    { wch: 30 },
    { wch: 20 },
    { wch: 30 },
    { wch: 60 },
    { wch: 18 },
    { wch: 18 },
    { wch: 18 },
    { wch: 50 }
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
      ["Monitoring Source", O2_STATUS_URL],
      ["Total Checks", allRecords.length],
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
        "Every automated check produces a timestamped screenshot in the evidence folder."
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
});
  const text = String(value).replace(/\r?\n/g, " ");

  return /[",]/.test(text)
    ? `"${text.replace(/"/g, '""')}"`
    : text;
}

function parseCSV(line) {
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

  const text = fs.readFileSync(csvFile, "utf8").trim();

  if (!text) {
    return [];
  }

  const lines = text.split(/\r?\n/);
  const headers = parseCSV(lines[0]);

  return lines.slice(1).map(line => {
    const values = parseCSV(line);
    const row = {};

    headers.forEach((header, index) => {
      row[header] = values[index] ?? "";
    });

    return row;
  });
}

function classifyStatus(text) {
  const lower = text.toLowerCase();

  const issuePatterns = [
    "nearby phone mast",
    "mast isn’t working",
    "mast isn't working",
    "mast is not working",
    "working as it should",
    "network issue",
    "network issues",
    "network problem",
    "known issue",
    "known issues",
    "service issue",
    "service problem",
    "outage",
    "fault",
    "engineers will be on the case",
    "engineer will be on the case",
    "service might come and go",
    "maintenance"
  ];

  for (const pattern of issuePatterns) {
    if (lower.includes(pattern)) {
      return "ISSUE";
    }
  }

  const okPatterns = [
    "no known issues",
    "no issues in your area",
    "working normally",
    "everything is working",
    "there are no known issues"
  ];

  for (const pattern of okPatterns) {
    if (lower.includes(pattern)) {
      return "OK";
    }
  }

  return "UNKNOWN";
}

/*
 * Find an input anywhere in the page,
 * including frames.
 */
async function findInputInPage(page) {
  const selectors = [
    'input[placeholder*="postcode" i]',
    'input[placeholder*="post code" i]',
    'input[aria-label*="postcode" i]',
    'input[aria-label*="post code" i]',
    'input[name*="postcode" i]',
    'input[name*="post code" i]',
    'input[id*="postcode" i]',
    'input[id*="post-code" i]',
    'input[autocomplete="postal-code"]',
    'input[data-testid*="postcode" i]',
    'input[type="text"]"
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
      } catch {}
    }
  }

  return null;
}

/*
 * Search every iframe for the postcode field.
 */
async function findPostcodeInputEverywhere(page) {
  let input = await findInputInPage(page);

  if (input) {
    return input;
  }

  for (const frame of page.frames()) {
    try {
      const selectors = [
        'input[placeholder*="postcode" i]',
        'input[placeholder*="post code" i]',
        'input[aria-label*="postcode" i]',
        'input[name*="postcode" i]',
        'input[id*="postcode" i]',
        'input[autocomplete="postal-code"]',
        'input[type="text"]'
      ];

      for (const selector of selectors) {
        const locator = frame.locator(selector);
        const count = await locator.count();

        for (let i = 0; i < count; i++) {
          const candidate = locator.nth(i);

          try {
            if (await candidate.isVisible()) {
              return candidate;
            }
          } catch {}
        }
      }
    } catch {}
  }

  return null;
}

/*
 * Look for anything that can open
 * the network checker.
 */
async function clickNetworkChecker(page) {
  const texts = [
    "O2 Network Issues",
    "Network Issues",
    "network issues",
    "Check network",
    "Check your network",
    "Check coverage",
    "Network status",
    "network status",
    "Check status"
  ];

  for (const text of texts) {
    try {
      const locator = page.getByText(text, {
        exact: false
      });

      const count = await locator.count();

      for (let i = 0; i < count; i++) {
        const element = locator.nth(i);

        try {
          if (await element.isVisible()) {
            console.log(`Clicking O2 checker link: ${text}`);

            await element.click({
              timeout: 10000
            });

            await page.waitForTimeout(5000);

            return true;
          }
        } catch {}
      }
    } catch {}
  }

  /*
   * Search links directly.
   */
  const links = await page.locator("a").evaluateAll(links =>
    links.map(link => ({
      text: (link.innerText || "").trim(),
      href: link.href
    }))
  );

  for (const link of links) {
    if (
      /network|status|coverage|issue|fault/i.test(
        `${link.text} ${link.href}`
      )
    ) {
      if (link.href) {
        console.log(
          `Navigating directly to possible O2 checker: ${link.href}`
        );

        await page.goto(link.href, {
          waitUntil: "domcontentloaded",
          timeout: 60000
        });

        await page.waitForTimeout(5000);

        return true;
      }
    }
  }

  return false;
}

/*
 * Diagnostic information.
 */
async function saveDiagnostics(page, timestamp) {
  const diagnosticFile = path.join(
    evidenceDir,
    `${timestamp}_DIAGNOSTIC.png`
  );

  await page.screenshot({
    path: diagnosticFile,
    fullPage: true
  });

  const htmlFile = path.join(
    evidenceDir,
    `${timestamp}_DIAGNOSTIC.html`
  );

  fs.writeFileSync(
    htmlFile,
    await page.content(),
    "utf8"
  );

  console.log(`Diagnostic screenshot: ${diagnosticFile}`);
  console.log(`Diagnostic HTML: ${htmlFile}`);
}

/*
 * Try to dismiss cookies / privacy banners.
 */
async function dismissPopups(page) {
  const buttons = [
    "Accept all",
    "Accept All",
    "Accept",
    "Allow all",
    "Allow All",
    "Continue",
    "Got it",
    "I agree"
  ];

  for (const text of buttons) {
    try {
      const locator = page.getByRole("button", {
        name: text,
        exact: true
      });

      if (await locator.count()) {
        if (await locator.first().isVisible()) {
          await locator.first().click({
            timeout: 5000
          });

          await page.waitForTimeout(1000);
        }
      }
    } catch {}
  }
}

/*
 * Submit the postcode.
 */
async function submitPostcode(page, postcodeInput) {
  const buttons = [
    "Check",
    "Search",
    "Submit",
    "Continue",
    "View",
    "Check status",
    "Check network"
  ];

  for (const text of buttons) {
    try {
      const button = page.getByRole("button", {
        name: text,
        exact: false
      });

      const count = await button.count();

      for (let i = 0; i < count; i++) {
        const candidate = button.nth(i);

        try {
          if (await candidate.isVisible()) {
            console.log(
              `Submitting using button: ${text}`
            );

            await candidate.click({
              timeout: 10000
            });

            return;
          }
        } catch {}
      }
    } catch {}
  }

  console.log("No suitable button found. Pressing Enter.");

  await postcodeInput.press("Enter");
}

/*
 * Wait for O2's actual result.
 */
async function waitForResult(page) {
  const resultPatterns = [
    "nearby phone mast",
    "mast isn’t working",
    "mast isn't working",
    "mast is not working",
    "engineers will be on the case",
    "service might come and go",
    "no known issues",
    "no issues in your area",
    "working normally",
    "network issue",
    "network issues"
  ];

  for (let i = 0; i < 30; i++) {
    const text = (
      await page.locator("body").innerText()
    ).replace(/\s+/g, " ");

    for (const pattern of resultPatterns) {
      if (
        text.toLowerCase().includes(pattern)
      ) {
        return text;
      }
    }

    await page.waitForTimeout(1000);
  }

  return (
    await page.locator("body").innerText()
  ).replace(/\s+/g, " ");
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
      height: 1400
    }
  });

  const page = await context.newPage();

  const timestamp = londonTimestamp();
  const fileTimestamp = safeFileTimestamp();

  let status = "UNKNOWN";
  let message = "";
  let expectedResolution = "";
  let checkResult = "FAILED";
  let screenshotName = "";

  try {
    console.log("===== O2 MONITOR START =====");

    console.log(`Opening ${O2_URL}`);

    await page.goto(O2_URL, {
      waitUntil: "domcontentloaded",
      timeout: 60000
    });

    await page.waitForTimeout(5000);

    await dismissPopups(page);

    console.log(`URL: ${page.url()}`);
    console.log(`TITLE: ${await page.title()}`);

    /*
     * First attempt: postcode already available.
     */
    let postcodeInput =
      await findPostcodeInputEverywhere(page);

    /*
     * If not available, open the network checker.
     */
    if (!postcodeInput) {
      console.log(
        "Postcode field not on initial page."
      );

      await clickNetworkChecker(page);

      await page.waitForTimeout(5000);

      await dismissPopups(page);

      postcodeInput =
        await findPostcodeInputEverywhere(page);
    }

    /*
     * Final attempt after another wait.
     */
    if (!postcodeInput) {
      await page.waitForTimeout(5000);

      postcodeInput =
        await findPostcodeInputEverywhere(page);
    }

    /*
     * We will NOT claim a valid check unless
     * the postcode field was actually found.
     */
    if (!postcodeInput) {
      await saveDiagnostics(
        page,
        fileTimestamp
      );

      throw new Error(
        "Could not find O2 postcode input."
      );
    }

    /*
     * ENTER POSTCODE.
     */
    console.log(
      `Entering postcode: ${POSTCODE}`
    );

    await postcodeInput.click();

    await postcodeInput.fill("");

    await postcodeInput.fill(
      POSTCODE
    );

    /*
     * Verify that the field really contains
     * the postcode before taking evidence.
     */
    const enteredValue =
      await postcodeInput.inputValue();

    if (
      enteredValue
        .replace(/\s+/g, "")
        .toUpperCase() !==
      POSTCODE
        .replace(/\s+/g, "")
        .toUpperCase()
    ) {
      throw new Error(
        `Postcode entry verification failed. Field contains: "${enteredValue}"`
      );
    }

    /*
     * SCREENSHOT 1:
     * This proves NW9 0RY was actually entered.
     */
    const enteredScreenshot =
      `${fileTimestamp}_01_POSTCODE_ENTERED.png`;

    await page.screenshot({
      path: path.join(
        evidenceDir,
        enteredScreenshot
      ),
      fullPage: true
    });

    console.log(
      `POSTCODE ENTERED SCREENSHOT: ${enteredScreenshot}`
    );

    /*
     * Submit.
     */
    await submitPostcode(
      page,
      postcodeInput
    );

    /*
     * Wait for actual result.
     */
    const resultText =
      await waitForResult(page);

    message =
      resultText
        .replace(/\s+/g, " ")
        .trim()
        .slice(0, 5000);

    status =
      classifyStatus(resultText);

    /*
     * Extract resolution information.
     */
    const resolutionMatch =
      resultText.match(
        /(?:updated|expected|estimated|should|aim).*?(\d{1,2}[:.]\d{2}\s?(?:am|pm)?|\d{1,2}\s?(?:am|pm))/i
      );

    if (resolutionMatch) {
      expectedResolution =
        resolutionMatch[0];
    }

    /*
     * SCREENSHOT 2:
     * This is the main evidence screenshot.
     */
    screenshotName =
      `${fileTimestamp}_02_O2_RESULT.png`;

    await page.screenshot({
      path: path.join(
        evidenceDir,
        screenshotName
      ),
      fullPage: true
    });

    console.log(
      `O2 RESULT SCREENSHOT: ${screenshotName}`
    );

    /*
     * The check is successful because:
     *
     * 1. Postcode field was found
     * 2. NW9 0RY was entered
     * 3. Entry was verified
     * 4. O2 result page was captured
     */
    checkResult = "SUCCESS";

    console.log(
      `O2 STATUS: ${status}`
    );

  } catch (error) {
    message =
      `MONITORING ERROR: ${error.message}`;

    checkResult = "FAILED";

    console.error(message);

    /*
     * Always preserve the failure page.
     */
    try {
      if (!screenshotName) {
        screenshotName =
          `${fileTimestamp}_FAILED.png`;
      }

      await page.screenshot({
        path: path.join(
          evidenceDir,
          screenshotName
        ),
        fullPage: true
      });
    } catch {}
  }

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
      page.url(),

    "Screenshot":
      screenshotName
        ? `evidence/${screenshotName}`
        : "",

    "Check Result":
      checkResult,

    "Status Changed?":
      statusChanged,

    "Previous Status":
      previous?.["O2 Status"] || "",

    "Notes":
      checkResult === "SUCCESS"
        ? "Postcode entered, verified and O2 result screenshot captured."
        : "Automation failed. This record must not be treated as evidence of O2 service status."
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
      .map(h => csvEscape(record[h]))
      .join(",") + "\n",
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
    { wch: 100 },
    { wch: 30 },
    { wch: 20 },
    { wch: 60 },
    { wch: 60 },
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

  const issueCount =
    allRecords.filter(
      x => x["O2 Status"] === "ISSUE"
    ).length;

  const okCount =
    allRecords.filter(
      x => x["O2 Status"] === "OK"
    ).length;

  const unknownCount =
    allRecords.filter(
      x => x["O2 Status"] === "UNKNOWN"
    ).length;

  const failedCount =
    allRecords.filter(
      x => x["Check Result"] === "FAILED"
    ).length;

  const successfulCount =
    allRecords.filter(
      x => x["Check Result"] === "SUCCESS"
    ).length;

  const summary =
    XLSX.utils.aoa_to_sheet([
      [
        "O2 NW9 0RY NETWORK EVIDENCE"
      ],
      [],
      [
        "Postcode",
        POSTCODE
      ],
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
        successfulCount
      ],
      [
        "Failed Checks",
        failedCount
      ],
      [
        "Issue Checks",
        issueCount
      ],
      [
        "OK Checks",
        okCount
      ],
      [
        "Unknown Checks",
        unknownCount
      ],
      [],
      [
        "Evidence Rule",
        "A check is only successful when the postcode was entered and the O2 result page was captured."
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

  await browser.close();

  console.log(
    "===== O2 MONITOR COMPLETE ====="
  );

  console.log(
    JSON.stringify(
      {
        timestamp,
        postcode: POSTCODE,
        status,
        checkResult,
        screenshot: screenshotName
      },
      null,
      2
    )
  );

  if (checkResult === "FAILED") {
    process.exit(1);
  }
}

main().catch(error => {
  console.error(error);
  process.exit(1);
});
