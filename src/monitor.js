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

const csvFile = path.join(
  dataDir,
  "o2-nw9-0ry-monitor.csv"
);

const xlsxFile = path.join(
  dataDir,
  "O2-NW9-0RY-Evidence.xlsx"
);

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

  const get = type =>
    parts.find(p => p.type === type)?.value || "";

  return `${get("year")}-${get("month")}-${get("day")}_${get("hour")}-${get("minute")}-${get("second")}`;
}

function csvEscape(value) {
  if (value === null || value === undefined) {
    return "";
  }

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
