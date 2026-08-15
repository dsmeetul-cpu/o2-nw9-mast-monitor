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
      p => p.type === name
    )?.value;

  return `${get("year")}-${get("month")}-${get("day")}_${get("hour")}-${get("minute")}-${get("second")}`;

}


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

    return `"${text.replace(
      /"/g,
      '""'
    )}"`;

  }

  return text;

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
    parseCSV(lines[0]);

  return lines
    .slice(1)
    .map(line => {

      const values =
        parseCSV(line);

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


function parseCSV(line) {

  const result = [];

  let current = "";

  let quoted = false;

  for (
    let i = 0;
    i < line.length;
    i++
  ) {

    const char =
      line[i];

    if (char === '"') {

      if (
        quoted &&
        line[i + 1] === '"'
      ) {

        current += '"';

        i++;

      } else {

        quoted =
          !quoted;

      }

    }

    else if (
      char === "," &&
      !quoted
    ) {

      result.push(current);

      current = "";

    }

    else {

      current += char;

    }

  }

  result.push(current);

  return result;

}


/*
 * Find a postcode input.
 *
 * O2 may change the HTML, so we deliberately
 * try several different selectors.
 */
async function findPostcodeInput(page) {

  const selectors = [

    'input[placeholder*="postcode" i]',

    'input[aria-label*="postcode" i]',

    'input[name*="postcode" i]',

    'input[id*="postcode" i]',

    'input[autocomplete="postal-code"]',

    'input[inputmode="text"]',

    'input[type="text"]'

  ];


  for (const selector of selectors) {

    const locator =
      page.locator(selector);


    const count =
      await locator.count();


    console.log(
      `Selector ${selector}: ${count} match(es)`
    );


    for (
      let i = 0;
      i < count;
      i++
    ) {

      const candidate =
        locator.nth(i);

      try {

        if (
          await candidate.isVisible()
        ) {

          console.log(
            `Using visible postcode candidate: ${selector} [${i}]`
          );

          return candidate;

        }

      } catch {

        // Continue.

      }

    }

  }


  return null;

}


/*
 * Find the most likely submit button.
 */
async function findSubmitButton(page) {

  const selectors = [

    'button:has-text("Check")',

    'button:has-text("Search")',

    'button:has-text("Submit")',

    'button:has-text("Continue")',

    'button:has-text("View")',

    'input[type="submit"]',

    'button'

  ];


  for (const selector of selectors) {

    const locator =
      page.locator(selector);


    const count =
      await locator.count();


    console.log(
      `Button selector ${selector}: ${count} match(es)`
    );


    for (
      let i = 0;
      i < count;
      i++
    ) {

      const candidate =
        locator.nth(i);


      try {

        if (
          await candidate.isVisible()
        ) {

          const text =
            await candidate.innerText()
              .catch(() => "");


          console.log(
            `Using button: ${selector} [${i}] "${text}"`
          );

          return candidate;

        }

      } catch {

        // Continue.

      }

    }

  }


  return null;

}


/*
 * Print every input on the page.
 *
 * This is extremely useful if O2 changes
 * the HTML again.
 */
async function diagnoseInputs(page) {

  console.log(
    "===== INPUT DIAGNOSTICS ====="
  );

  const inputs =
    await page.locator("input").evaluateAll(
      elements =>
        elements.map(
          (input, index) => ({

            index,

            type:
              input.getAttribute("type"),

            name:
              input.getAttribute("name"),

            id:
              input.getAttribute("id"),

            placeholder:
              input.getAttribute(
                "placeholder"
              ),

            ariaLabel:
              input.getAttribute(
                "aria-label"
              ),

            autocomplete:
              input.getAttribute(
                "autocomplete"
              ),

            value:
              input.value,

            visible:
              !!(
                input.offsetWidth ||
                input.offsetHeight ||
                input.getClientRects().length
              )

          })
        )
    );


  console.log(
    JSON.stringify(
      inputs,
      null,
      2
    )
  );

  console.log(
    "===== END INPUT DIAGNOSTICS ====="
  );

}


/*
 * Find the O2 Network Issues link on
 * the landing page.
 */
async function openNetworkIssuesChecker(page) {

  console.log(
    "Looking for O2 Network Issues checker..."
  );


  const links =
    page.getByText(
      "O2 Network Issues",
      {
        exact: true
      }
    );


  if (
    await links.count()
  ) {

    const link =
      links.first();


    if (
      await link.isVisible()
    ) {

      console.log(
        "Found O2 Network Issues link."
      );


      const href =
        await link.getAttribute(
          "href"
        );


      console.log(
        `O2 Network Issues href: ${href}`
      );


      await link.click();


      await page.waitForLoadState(
        "domcontentloaded",
        {
          timeout: 30000
        }
      ).catch(() => {});


      await page.waitForTimeout(
        3000
      );


      console.log(
        `After navigation URL: ${page.url()}`
      );


      return true;

    }

  }


  /*
   * Fallback: search all links for
   * network/status wording.
   */
  const possibleLinks =
    await page.locator("a").evaluateAll(
      links =>
        links.map(
          link => ({

            text:
              (link.innerText || "")
                .trim(),

            href:
              link.href

          })
        )
        .filter(
          x =>
            /network|status|issue/i
              .test(x.text)
        )
    );


  console.log(
    "Possible network links:"
  );

  console.log(
    JSON.stringify(
      possibleLinks,
      null,
      2
    )
  );


  if (possibleLinks.length) {

    const href =
      possibleLinks[0].href;


    if (href) {

      console.log(
        `Navigating to: ${href}`
      );


      await page.goto(
        href,
        {
          waitUntil:
            "domcontentloaded",
          timeout: 60000
        }
      );


      await page.waitForTimeout(
        3000
      );


      return true;

    }

  }


  return false;

}


/*
 * Classify the result.
 */
function classifyStatus(text) {

  const lower =
    text.toLowerCase();


  /*
   * Strong outage wording.
   */
  const issuePatterns = [

    "nearby phone mast",

    "mast isn't working",

    "mast isn’t working",

    "mast is not working",

    "isn't working as it should",

    "isn’t working as it should",

    "network issue",

    "network problem",

    "known issue",

    "service issue",

    "service problem",

    "outage",

    "fault",

    "engineers will be on the case",

    "engineer is on the case",

    "service might come and go",

    "maintenance"

  ];


  for (
    const pattern of issuePatterns
  ) {

    if (
      lower.includes(pattern)
    ) {

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


  for (
    const pattern of okPatterns
  ) {

    if (
      lower.includes(pattern)
    ) {

      return "OK";

    }

  }


  return "UNKNOWN";

}


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

        height: 1400

      }

    });


  const page =
    await context.newPage();


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


  let pageText =
    "";


  let checkResult =
    "FAILED";


  let screenshotName =
    "";


  try {

    console.log(
      "===== O2 MONITOR START ====="
    );


    console.log(
      `Opening: ${O2_URL}`
    );


    await page.goto(
      O2_URL,
      {
        waitUntil:
          "domcontentloaded",

        timeout:
          60000
      }
    );


    await page.waitForTimeout(
      3000
    );


    console.log(
      `Initial URL: ${page.url()}`
    );


    console.log(
      `Initial title: ${await page.title()}`
    );


    /*
     * First check whether the landing page
     * itself contains a postcode field.
     */
    let postcodeInput =
      await findPostcodeInput(page);


    /*
     * If it doesn't, follow the actual
     * Network Issues checker.
     */
    if (!postcodeInput) {

      console.log(
        "No postcode input on landing page."
      );


      const navigated =
        await openNetworkIssuesChecker(
          page
        );


      if (!navigated) {

        await diagnoseInputs(
          page
        );


        throw new Error(
          "Could not locate the O2 Network Issues checker."
        );

      }


      postcodeInput =
        await findPostcodeInput(
          page
        );

    }


    /*
     * If still missing, produce diagnostics
     * before failing.
     */
    if (!postcodeInput) {

      await diagnoseInputs(
        page
      );


      await page.screenshot({

        path:
          path.join(
            evidenceDir,
            `${fileTimestamp}_DIAGNOSTIC.png`
          ),

        fullPage:
          true

      });


      throw new Error(
        "Could not find O2 postcode input after opening the Network Issues checker."
      );

    }


    /*
     * ENTER POSTCODE.
     */
    console.log(
      `Entering postcode: ${POSTCODE}`
    );


    await postcodeInput.fill(
      POSTCODE
    );


    /*
     * IMPORTANT:
     *
     * This screenshot proves that the
     * monitor actually entered the postcode.
     */
    const enteredScreenshot =
      `${fileTimestamp}_01_POSTCODE_ENTERED.png`;


    await page.screenshot({

      path:
        path.join(
          evidenceDir,
          enteredScreenshot
        ),

      fullPage:
        true

    });


    console.log(
      `Postcode-entered screenshot: ${enteredScreenshot}`
    );


    /*
     * Find submit button.
     */
    const submitButton =
      await findSubmitButton(
        page
      );


    if (submitButton) {

      console.log(
        "Submitting postcode..."
      );


      await submitButton.click();

    } else {

      console.log(
        "No submit button found. Pressing Enter."
      );


      await postcodeInput.press(
        "Enter"
      );

    }


    /*
     * Give O2 enough time to display
     * the result.
     */
    await page.waitForTimeout(
      8000
    );


    console.log(
      `Result URL: ${page.url()}`
    );


    console.log(
      `Result title: ${await page.title()}`
    );


    /*
     * Capture the actual O2 result.
     */
    pageText =
      await page.locator(
        "body"
      ).innerText();


    /*
     * Determine status.
     */
    status =
      classifyStatus(
        pageText
      );


    /*
     * Clean message for CSV.
     */
    message =
      pageText
        .replace(/\s+/g, " ")
        .trim()
        .slice(
          0,
          5000
        );


    /*
     * Try to extract expected resolution.
     */
    const resolutionMatch =
      pageText.match(
        /(?:expected|estimated|should|aim).*?(?:resolved|fixed).*?(\d{1,2}[:.]\d{2}|\d{1,2}\s?(?:am|pm)|\d{1,2}\/\d{1,2}\/\d{2,4})/i
      );


    if (resolutionMatch) {

      expectedResolution =
        resolutionMatch[0];

    }


    /*
     * Final evidence screenshot.
     *
     * This is the important screenshot:
     * postcode has been submitted and the
     * O2 result is visible.
     */
    screenshotName =
      `${fileTimestamp}_02_O2_RESULT.png`;


    await page.screenshot({

      path:
        path.join(
          evidenceDir,
          screenshotName
        ),

      fullPage:
        true

    });


    console.log(
      `Result screenshot: ${screenshotName}`
    );


    /*
     * A successful monitor run means we
     * successfully interacted with O2.
     *
     * UNKNOWN can therefore still be a
     * successful check.
     */
    checkResult =
      "SUCCESS";


    console.log(
      `O2 Status: ${status}`
    );


  } catch (error) {

    message =
      `MONITORING ERROR: ${error.message}`;


    checkResult =
      "FAILED";


    /*
     * Always preserve a failure screenshot.
     */
    try {

      screenshotName =
        screenshotName ||
        `${fileTimestamp}_FAILED.png`;


      await page.screenshot({

        path:
          path.join(
            evidenceDir,
            screenshotName
          ),

        fullPage:
          true

      });

    } catch {

      // Ignore screenshot failure.

    }


    console.error(
      message
    );

  }


  /*
   * Read previous records.
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
      previous?.["O2 Status"]
      || "",

    "Notes":
      checkResult === "SUCCESS"
        ? "Postcode entered and O2 result captured."
        : "Monitor failed before a valid O2 result was captured."

  };


  /*
   * Write CSV.
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
        h =>
          csvEscape(
            record[h]
          )
      )
      .join(",") +
      "\n",

    "utf8"

  );


  /*
   * Update Excel evidence file.
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
    { wch: 80 },
    { wch: 30 },
    { wch: 20 },
    { wch: 55 },
    { wch: 55 },
    { wch: 18 },
    { wch: 18 },
    { wch: 18 },
    { wch: 45 }

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
      x =>
        x["O2 Status"] ===
        "ISSUE"
    ).length;


  const okCount =
    allRecords.filter(
      x =>
        x["O2 Status"] ===
        "OK"
    ).length;


  const unknownCount =
    allRecords.filter(
      x =>
        x["O2 Status"] ===
        "UNKNOWN"
    ).length;


  const failedCount =
    allRecords.filter(
      x =>
        x["Check Result"] ===
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
        O2_URL
      ],

      [
        "Total Checks",
        allRecords.length
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

      [
        "Failed Checks",
        failedCount
      ],

      [],

      [
        "Important",
        "A successful check means the postcode was entered into the O2 checker and the resulting page was captured."
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
    "===== O2 MONITOR COMPLETE ====="
  );


  console.log(
    JSON.stringify(
      {
        timestamp,
        postcode:
          POSTCODE,
        status,
        checkResult,
        screenshot:
          screenshotName
      },
      null,
      2
    )
  );


  /*
   * IMPORTANT:
   *
   * Don't make GitHub Actions fail simply
   * because O2 returned UNKNOWN.
   *
   * But do fail the workflow if the
   * automation itself failed.
   */
  if (
    checkResult === "FAILED"
  ) {

    process.exit(1);

  }

}


main().catch(error => {

  console.error(
    error
  );

  process.exit(1);

});  }

  const text =
    String(value).replace(/\r?\n/g, " ");

  if (/[",]/.test(text)) {
    return `"${text.replace(/"/g, '""')}"`;
  }

  return text;
}

function readExistingCsv() {

  if (!fs.existsSync(csvFile)) {
    return [];
  }

  const text =
    fs.readFileSync(csvFile, "utf8").trim();

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

      headers.forEach((header, index) => {
        row[header] =
          values[index] ?? "";
      });

      return row;
    });
}

function parseCSVLine(line) {

  const values = [];

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

      values.push(current);
      current = "";

    } else {

      current += char;
    }
  }

  values.push(current);

  return values;
}


/*
 * Find postcode input.
 *
 * We deliberately search very broadly because
 * O2's page is dynamically rendered.
 */
async function findPostcodeInput(page) {

  console.log("");
  console.log("Searching for postcode field...");

  const selectors = [

    'input[placeholder*="postcode" i]',

    'input[aria-label*="postcode" i]',

    'input[name*="postcode" i]',

    'input[id*="postcode" i]',

    'input[data-testid*="postcode" i]',

    'input[autocomplete="postal-code"]',

    'input[autocomplete="postcode"]',

    'input[type="text"]',

    'input:not([type])'
  ];

  for (const selector of selectors) {

    const locator =
      page.locator(selector);

    const count =
      await locator.count();

    console.log(
      `Selector ${selector}: ${count}`
    );

    for (
      let i = 0;
      i < count;
      i++
    ) {

      const candidate =
        locator.nth(i);

      try {

        if (
          await candidate.isVisible()
        ) {

          console.log(
            `Visible candidate found: ${selector} [${i}]`
          );

          return candidate;
        }

      } catch {
        // Continue.
      }
    }
  }

  /*
   * Search labels.
   */
  const labels =
    await page.locator("label").all();

  for (const label of labels) {

    try {

      const text =
        (
          await label.innerText()
        )
          .replace(/\s+/g, " ")
          .trim();

      console.log(
        `Label found: ${text}`
      );

      if (
        text.toLowerCase()
          .includes("postcode")
      ) {

        const forId =
          await label.getAttribute("for");

        if (forId) {

          const candidate =
            page.locator(`#${forId}`);

          if (
            await candidate.count() &&
            await candidate.isVisible()
          ) {

            console.log(
              "Postcode field found through label."
            );

            return candidate;
          }
        }
      }

    } catch {
      // Continue.
    }
  }

  return null;
}


/*
 * Find the appropriate submit/check button.
 */
async function findSubmitButton(page) {

  console.log("");
  console.log(
    "Searching for postcode submit button..."
  );

  const candidates = [

    'button:has-text("Check")',

    'button:has-text("Check status")',

    'button:has-text("Search")',

    'button:has-text("Submit")',

    'button:has-text("Continue")',

    'input[type="submit"]',

    'button'
  ];

  for (const selector of candidates) {

    const locator =
      page.locator(selector);

    const count =
      await locator.count();

    console.log(
      `${selector}: ${count}`
    );

    for (
      let i = 0;
      i < count;
      i++
    ) {

      const button =
        locator.nth(i);

      try {

        if (
          await button.isVisible() &&
          await button.isEnabled()
        ) {

          const text =
            (
              await button.innerText()
            )
              .replace(/\s+/g, " ")
              .trim();

          console.log(
            `Button candidate: "${text}"`
          );

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

  const lower =
    text.toLowerCase();

  /*
   * This is deliberately based on
   * O2's actual wording.
   */

  if (
    lower.includes(
      "nearby phone mast"
    ) ||

    lower.includes(
      "mast isn’t working"
    ) ||

    lower.includes(
      "mast isn't working"
    ) ||

    lower.includes(
      "isn’t working as it should"
    ) ||

    lower.includes(
      "isn't working as it should"
    ) ||

    lower.includes(
      "service might come and go"
    ) ||

    lower.includes(
      "engineers will be on the case"
    ) ||

    lower.includes("outage") ||

    lower.includes("network issue") ||

    lower.includes("network fault")
  ) {

    return "ISSUE";
  }

  if (
    lower.includes(
      "working normally"
    ) ||

    lower.includes(
      "no known issues"
    ) ||

    lower.includes(
      "no issues"
    )
  ) {

    return "OK";
  }

  return "UNKNOWN";
}


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

  const timestamp =
    londonTimestamp();

  const fileTimestamp =
    safeFileTimestamp();

  let status = "UNKNOWN";
  let message = "";
  let expectedResolution = "";
  let checkResult = "FAILED";

  let postcodeEntered = false;
  let postcodeSubmitted = false;

  try {

    console.log("");
    console.log(
      "========================================"
    );

    console.log(
      "O2 NW9 0RY NETWORK CHECK"
    );

    console.log(
      "========================================"
    );

    console.log(
      `Starting URL: ${O2_URL}`
    );

    console.log(
      `Postcode: ${POSTCODE}`
    );

    console.log("");

    /*
     * OPEN O2
     */

    await page.goto(
      O2_URL,
      {
        waitUntil:
          "domcontentloaded",

        timeout: 60000
      }
    );

    await page.waitForTimeout(5000);

    console.log(
      `Loaded URL: ${page.url()}`
    );

    /*
     * IMPORTANT:
     * Look through the page and any iframe.
     */

    let postcodeInput =
      await findPostcodeInput(page);

    /*
     * If not found, inspect frames.
     */

    if (!postcodeInput) {

      console.log(
        "No postcode field in main page."
      );

      for (
        const frame of page.frames()
      ) {

        if (
          frame === page.mainFrame()
        ) {
          continue;
        }

        console.log(
          `Checking iframe: ${frame.url()}`
        );

        try {

          const candidate =
            await findPostcodeInput(frame);

          if (candidate) {

            postcodeInput =
              candidate;

            console.log(
              "Postcode field found in iframe."
            );

            break;
          }

        } catch {
          // Continue.
        }
      }
    }

    if (!postcodeInput) {

      throw new Error(
        "Could not locate O2 postcode field."
      );
    }

    /*
     * ENTER POSTCODE
     */

    console.log("");
    console.log(
      "ENTERING POSTCODE..."
    );

    await postcodeInput.scrollIntoViewIfNeeded();

    await postcodeInput.click();

    await postcodeInput.fill("");

    await postcodeInput.fill(
      POSTCODE
    );

    await page.waitForTimeout(1000);

    /*
     * VERIFY VALUE
     */

    const enteredValue =
      await postcodeInput.inputValue();

    console.log(
      `Postcode field now contains: "${enteredValue}"`
    );

    if (
      enteredValue
        .replace(/\s+/g, "")
        .toUpperCase() !==
      POSTCODE
        .replace(/\s+/g, "")
        .toUpperCase()
    ) {

      throw new Error(
        `Postcode verification failed. Field contains "${enteredValue}".`
      );
    }

    postcodeEntered = true;

    /*
     * TAKE PRE-SUBMISSION SCREENSHOT.
     *
     * This proves the postcode was actually entered.
     */

    const preSubmitScreenshot =
      path.join(
        evidenceDir,
        `${fileTimestamp}_NW9-0RY_POSTCODE-ENTERED.png`
      );

    await page.screenshot({
      path:
        preSubmitScreenshot,
      fullPage: true
    });

    console.log("");
    console.log(
      "POSTCODE ENTERED SUCCESSFULLY."
    );

    console.log(
      `Pre-submit screenshot: ${preSubmitScreenshot}`
    );

    /*
     * FIND SUBMIT BUTTON
     */

    const submitButton =
      await findSubmitButton(page);

    /*
     * SUBMIT
     */

    console.log("");
    console.log(
      "SUBMITTING POSTCODE..."
    );

    if (submitButton) {

      await submitButton.scrollIntoViewIfNeeded();

      await submitButton.click();

    } else {

      console.log(
        "No button found - pressing Enter."
      );

      await postcodeInput.press(
        "Enter"
      );
    }

    postcodeSubmitted = true;

    console.log(
      "Postcode submitted."
    );

    /*
     * WAIT FOR O2 RESULT
     */

    console.log("");
    console.log(
      "WAITING FOR O2 RESULT..."
    );

    await page.waitForTimeout(3000);

    /*
     * Give O2 up to 20 seconds to
     * render its result.
     */

    for (
      let i = 0;
      i < 10;
      i++
    ) {

      const currentText =
        await page
          .locator("body")
          .innerText();

      console.log(
        `Result check ${i + 1}/10`
      );

      console.log(
        currentText
          .replace(/\s+/g, " ")
          .trim()
          .slice(0, 1000)
      );

      /*
       * Stop early if O2 has clearly
       * returned a network result.
       */

      const lower =
        currentText.toLowerCase();

      if (
        lower.includes("nearby phone mast") ||
        lower.includes("mast isn't working") ||
        lower.includes("mast isn’t working") ||
        lower.includes("service might come and go") ||
        lower.includes("network issue") ||
        lower.includes("network fault") ||
        lower.includes("working normally") ||
        lower.includes("no known issues")
      ) {

        console.log(
          "O2 network result detected."
        );

        break;
      }

      await page.waitForTimeout(2000);
    }

    /*
     * CAPTURE FINAL RESULT
     */

    const finalText =
      await page
        .locator("body")
        .innerText();

    const cleanedText =
      finalText
        .replace(/\s+/g, " ")
        .trim();

    console.log("");
    console.log(
      "========================================"
    );

    console.log(
      "FINAL O2 PAGE"
    );

    console.log(
      "========================================"
    );

    console.log(
      `URL: ${page.url()}`
    );

    console.log("");

    console.log(
      cleanedText.slice(0, 10000)
    );

    /*
     * CLASSIFY
     */

    status =
      classifyStatus(cleanedText);

    message =
      cleanedText.slice(0, 5000);

    /*
     * Look for O2 update time.
     */

    const resolutionMatch =
      cleanedText.match(
        /(?:updated|update|expected|estimated|resolved|fixed).*?(\d{1,2}[:.]\d{2}\s?(?:am|pm)?)/i
      );

    if (resolutionMatch) {

      expectedResolution =
        resolutionMatch[0];
    }

    /*
     * FINAL SCREENSHOT
     *
     * THIS is the important evidence screenshot.
     */

    const screenshotName =
      `${fileTimestamp}_NW9-0RY.png`;

    const screenshotPath =
      path.join(
        evidenceDir,
        screenshotName
      );

    await page.screenshot({
      path:
        screenshotPath,
      fullPage: true
    });

    console.log("");
    console.log(
      "========================================"
    );

    console.log(
      "FINAL EVIDENCE SCREENSHOT"
    );

    console.log(
      "========================================"
    );

    console.log(
      `Saved: ${screenshotPath}`
    );

    /*
     * Only a successful check can
     * produce a SUCCESS result.
     */

    if (
      postcodeEntered &&
      postcodeSubmitted
    ) {

      checkResult = "SUCCESS";

    } else {

      checkResult = "FAILED";
    }

  } catch (error) {

    status = "UNKNOWN";

    message =
      `MONITORING ERROR: ${error.message}`;

    checkResult = "FAILED";

    console.error("");
    console.error(
      "MONITORING ERROR:"
    );

    console.error(
      error.message
    );

    /*
     * Diagnostic screenshot.
     *
     * This is NOT treated as outage evidence.
     */

    const diagnosticName =
      `${fileTimestamp}_NW9-0RY_DIAGNOSTIC.png`;

    const diagnosticPath =
      path.join(
        evidenceDir,
        diagnosticName
      );

    try {

      await page.screenshot({
        path:
          diagnosticPath,
        fullPage: true
      });

      console.log(
        `Diagnostic screenshot: ${diagnosticPath}`
      );

    } catch {
      // Ignore screenshot failure.
    }
  }

  /*
   * RECORD RESULT
   */

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

  const screenshotName =
    `${fileTimestamp}_NW9-0RY.png`;

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
      postcodeEntered
        ? "Postcode verified before submission."
        : "Postcode was not successfully entered."
  };

  const headers =
    Object.keys(record);

  if (!fs.existsSync(csvFile)) {

    fs.writeFileSync(
      csvFile,
      headers
        .map(csvEscape)
        .join(",") + "\n",
      "utf8"
    );
  }

  fs.appendFileSync(
    csvFile,
    headers
      .map(h =>
        csvEscape(record[h])
      )
      .join(",") + "\n",
    "utf8"
  );

  /*
   * UPDATE EXCEL
   */

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
    { wch: 45 }
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
        "Issue Checks",
        allRecords.filter(
          x =>
            x["O2 Status"] === "ISSUE"
        ).length
      ],
      [
        "OK Checks",
        allRecords.filter(
          x =>
            x["O2 Status"] === "OK"
        ).length
      ],
      [
        "Unknown Checks",
        allRecords.filter(
          x =>
            x["O2 Status"] === "UNKNOWN"
        ).length
      ],
      [
        "Failed Checks",
        allRecords.filter(
          x =>
            x["Check Result"] === "FAILED"
        ).length
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

  console.log("");
  console.log(
    "========================================"
  );

  console.log(
    "MONITOR COMPLETE"
  );

  console.log(
    "========================================"
  );

  console.log(
    `Timestamp: ${timestamp}`
  );

  console.log(
    `Postcode entered: ${postcodeEntered}`
  );

  console.log(
    `Postcode submitted: ${postcodeSubmitted}`
  );

  console.log(
    `O2 Status: ${status}`
  );

  console.log(
    `Check Result: ${checkResult}`
  );

  console.log("");
}

main().catch(error => {

  console.error(error);

  process.exit(1);
});      width: 1440,
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
