import { chromium } from "playwright";
import fs from "fs";
import path from "path";

const O2_URL = "https://status.o2.co.uk/";
const POSTCODE = "NW9 0RY";

const root = process.cwd();

const dataDir = path.join(root, "data");
const evidenceDir = path.join(root, "evidence");

fs.mkdirSync(dataDir, { recursive: true });
fs.mkdirSync(evidenceDir, { recursive: true });

const timestamp = new Date()
  .toISOString()
  .replace(/[:.]/g, "-");

const diagnosticFile = path.join(
  dataDir,
  "o2-diagnostic.txt"
);

const screenshotFile = path.join(
  evidenceDir,
  `${timestamp}_DIAGNOSTIC_NW9-0RY.png`
);

const output = [];

function log(message = "") {
  console.log(message);
  output.push(message);
}

function separator() {
  log("");
  log("============================================================");
  log("");
}

async function safeAttribute(locator, attribute) {
  try {
    return (await locator.getAttribute(attribute)) || "";
  } catch {
    return "";
  }
}

async function safeText(locator) {
  try {
    return ((await locator.innerText()) || "").trim();
  } catch {
    return "";
  }
}

async function isVisible(locator) {
  try {
    return await locator.isVisible();
  } catch {
    return false;
  }
}

async function inspectFrame(frame, frameNumber) {
  separator();

  log(`FRAME ${frameNumber}`);
  log(`URL: ${frame.url()}`);

  if (!frame.url()) {
    log("URL: [EMPTY]");
  }

  // ----------------------------------------------------------
  // INPUTS
  // ----------------------------------------------------------

  log("");
  log(`--- INPUTS IN FRAME ${frameNumber} ---`);

  const inputs = frame.locator("input");

  let inputCount = 0;

  try {
    inputCount = await inputs.count();
  } catch {
    inputCount = 0;
  }

  log(`Total input elements found: ${inputCount}`);

  for (let i = 0; i < inputCount; i++) {
    const input = inputs.nth(i);

    if (!(await isVisible(input))) {
      continue;
    }

    const type = await safeAttribute(input, "type");
    const placeholder = await safeAttribute(input, "placeholder");
    const name = await safeAttribute(input, "name");
    const id = await safeAttribute(input, "id");
    const ariaLabel = await safeAttribute(input, "aria-label");
    const value = await safeAttribute(input, "value");
    const autocomplete = await safeAttribute(input, "autocomplete");
    const className = await safeAttribute(input, "class");

    log("");
    log(`VISIBLE INPUT #${i + 1}`);
    log(`  type:         ${type}`);
    log(`  placeholder:  ${placeholder}`);
    log(`  name:         ${name}`);
    log(`  id:           ${id}`);
    log(`  aria-label:   ${ariaLabel}`);
    log(`  value:        ${value}`);
    log(`  autocomplete: ${autocomplete}`);
    log(`  class:        ${className}`);
  }

  // ----------------------------------------------------------
  // TEXTAREAS
  // ----------------------------------------------------------

  log("");
  log(`--- TEXTAREAS IN FRAME ${frameNumber} ---`);

  const textareas = frame.locator("textarea");

  let textareaCount = 0;

  try {
    textareaCount = await textareas.count();
  } catch {
    textareaCount = 0;
  }

  log(`Total textarea elements found: ${textareaCount}`);

  for (let i = 0; i < textareaCount; i++) {
    const textarea = textareas.nth(i);

    if (!(await isVisible(textarea))) {
      continue;
    }

    log("");
    log(`VISIBLE TEXTAREA #${i + 1}`);
    log(`  placeholder: ${await safeAttribute(textarea, "placeholder")}`);
    log(`  name:        ${await safeAttribute(textarea, "name")}`);
    log(`  id:          ${await safeAttribute(textarea, "id")}`);
    log(`  aria-label:  ${await safeAttribute(textarea, "aria-label")}`);
  }

  // ----------------------------------------------------------
  // BUTTONS
  // ----------------------------------------------------------

  log("");
  log(`--- BUTTONS IN FRAME ${frameNumber} ---`);

  const buttons = frame.locator("button");

  let buttonCount = 0;

  try {
    buttonCount = await buttons.count();
  } catch {
    buttonCount = 0;
  }

  log(`Total button elements found: ${buttonCount}`);

  for (let i = 0; i < buttonCount; i++) {
    const button = buttons.nth(i);

    if (!(await isVisible(button))) {
      continue;
    }

    const text = await safeText(button);
    const type = await safeAttribute(button, "type");
    const ariaLabel = await safeAttribute(button, "aria-label");
    const id = await safeAttribute(button, "id");
    const name = await safeAttribute(button, "name");
    const title = await safeAttribute(button, "title");
    const className = await safeAttribute(button, "class");

    log("");
    log(`VISIBLE BUTTON #${i + 1}`);
    log(`  text:        ${text}`);
    log(`  type:        ${type}`);
    log(`  aria-label:  ${ariaLabel}`);
    log(`  id:          ${id}`);
    log(`  name:        ${name}`);
    log(`  title:       ${title}`);
    log(`  class:       ${className}`);
  }

  // ----------------------------------------------------------
  // INPUT-LIKE ELEMENTS
  // ----------------------------------------------------------

  log("");
  log(`--- INPUT-LIKE ELEMENTS IN FRAME ${frameNumber} ---`);

  const inputLike = frame.locator(
    '[contenteditable="true"], [role="textbox"], [role="combobox"]'
  );

  let inputLikeCount = 0;

  try {
    inputLikeCount = await inputLike.count();
  } catch {
    inputLikeCount = 0;
  }

  log(`Total input-like elements found: ${inputLikeCount}`);

  for (let i = 0; i < inputLikeCount; i++) {
    const element = inputLike.nth(i);

    if (!(await isVisible(element))) {
      continue;
    }

    log("");
    log(`VISIBLE INPUT-LIKE ELEMENT #${i + 1}`);
    log(`  tag:         ${await element.evaluate(el => el.tagName)}`);
    log(`  role:        ${await safeAttribute(element, "role")}`);
    log(`  contenteditable: ${await safeAttribute(element, "contenteditable")}`);
    log(`  placeholder: ${await safeAttribute(element, "placeholder")}`);
    log(`  name:        ${await safeAttribute(element, "name")}`);
    log(`  id:          ${await safeAttribute(element, "id")}`);
    log(`  aria-label:  ${await safeAttribute(element, "aria-label")}`);
    log(`  aria-labelledby: ${await safeAttribute(element, "aria-labelledby")}`);
    log(`  class:       ${await safeAttribute(element, "class")}`);
  }

  // ----------------------------------------------------------
  // LINKS
  // ----------------------------------------------------------

  log("");
  log(`--- VISIBLE LINKS IN FRAME ${frameNumber} ---`);

  const links = frame.locator("a");

  let linkCount = 0;

  try {
    linkCount = await links.count();
  } catch {
    linkCount = 0;
  }

  let visibleLinks = 0;

  for (let i = 0; i < linkCount; i++) {
    const link = links.nth(i);

    if (!(await isVisible(link))) {
      continue;
    }

    visibleLinks++;

    const text = await safeText(link);
    const href = await safeAttribute(link, "href");
    const ariaLabel = await safeAttribute(link, "aria-label");

    log(
      `LINK #${i + 1}: text="${text}" href="${href}" aria-label="${ariaLabel}"`
    );
  }

  log(`Visible links: ${visibleLinks}`);

  // ----------------------------------------------------------
  // FORMS
  // ----------------------------------------------------------

  log("");
  log(`--- FORMS IN FRAME ${frameNumber} ---`);

  const forms = frame.locator("form");

  let formCount = 0;

  try {
    formCount = await forms.count();
  } catch {
    formCount = 0;
  }

  log(`Forms found: ${formCount}`);

  for (let i = 0; i < formCount; i++) {
    const form = forms.nth(i);

    if (!(await isVisible(form))) {
      continue;
    }

    log("");
    log(`VISIBLE FORM #${i + 1}`);
    log(`  action: ${await safeAttribute(form, "action")}`);
    log(`  method: ${await safeAttribute(form, "method")}`);
    log(`  id:     ${await safeAttribute(form, "id")}`);
    log(`  name:   ${await safeAttribute(form, "name")}`);
  }

  // ----------------------------------------------------------
  // BODY TEXT
  // ----------------------------------------------------------

  log("");
  log(`--- VISIBLE TEXT IN FRAME ${frameNumber} ---`);

  try {
    const bodyText = await frame.locator("body").innerText();

    const cleaned = bodyText
      .replace(/\n{3,}/g, "\n\n")
      .trim();

    log(cleaned.slice(0, 12000));
  } catch (error) {
    log(`Could not read body text: ${error.message}`);
  }
}

async function main() {
  log("============================================================");
  log("O2 NW9 0RY DIAGNOSTIC MONITOR");
  log("============================================================");
  log(`Target URL: ${O2_URL}`);
  log(`Postcode intended for test: "${POSTCODE}"`);
  log(`Timestamp: ${new Date().toISOString()}`);
  log("");

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

  // Capture console messages from the O2 page.
  page.on("console", message => {
    log(
      `[PAGE CONSOLE] ${message.type()}: ${message.text()}`
    );
  });

  // Capture failed network requests.
  page.on("requestfailed", request => {
    log("");
    log(
      `[REQUEST FAILED] ${request.method()} ${request.url()}`
    );

    const failure = request.failure();

    if (failure) {
      log(`  Error: ${failure.errorText}`);
    }
  });

  // Capture responses that return errors.
  page.on("response", response => {
    if (response.status() >= 400) {
      log(
        `[HTTP ${response.status()}] ${response.request().method()} ${response.url()}`
      );
    }
  });

  try {
    log("");
    log("Opening O2 status page...");

    await page.goto(O2_URL, {
      waitUntil: "domcontentloaded",
      timeout: 60000
    });

    log(`Initial page URL: ${page.url()}`);

    log("");
    log("Waiting 5 seconds for JavaScript/iframes to load...");

    await page.waitForTimeout(5000);

    // Scroll through the page to trigger lazy-loaded content.
    try {
      await page.evaluate(async () => {
        window.scrollTo(0, document.body.scrollHeight);
        await new Promise(resolve => setTimeout(resolve, 1500));
        window.scrollTo(0, 0);
      });
    } catch {
      // Ignore scrolling errors.
    }

    log("");
    log("Waiting another 3 seconds for dynamic content...");

    await page.waitForTimeout(3000);

    // --------------------------------------------------------
    // MAIN PAGE
    // --------------------------------------------------------

    await inspectFrame(page.mainFrame(), 0);

    // --------------------------------------------------------
    // ALL IFRAMES
    // --------------------------------------------------------

    separator();

    log("============================================================");
    log("IFRAME INVENTORY");
    log("============================================================");

    const frames = page.frames();

    log(`Total frames detected: ${frames.length}`);

    for (let i = 0; i < frames.length; i++) {
      const frame = frames[i];

      log("");
      log(`FRAME ${i}`);
      log(`URL: ${frame.url()}`);

      if (frame.parentFrame()) {
        log(
          `Parent frame URL: ${frame.parentFrame().url()}`
        );
      } else {
        log("Parent: MAIN PAGE");
      }
    }

    // --------------------------------------------------------
    // INSPECT EVERY FRAME
    // --------------------------------------------------------

    separator();

    log("============================================================");
    log("DETAILED FRAME INSPECTION");
    log("============================================================");

    for (let i = 0; i < frames.length; i++) {
      try {
        await inspectFrame(frames[i], i);
      } catch (error) {
        log("");
        log(
          `ERROR INSPECTING FRAME ${i}: ${error.message}`
        );
      }
    }

    // --------------------------------------------------------
    // ALL FRAME ELEMENTS FROM MAIN PAGE
    // --------------------------------------------------------

    separator();

    log("============================================================");
    log("IFRAME ELEMENTS AS SEEN FROM MAIN PAGE");
    log("============================================================");

    const iframeElements = page.locator("iframe");

    const iframeCount = await iframeElements.count();

    log(`iframe elements found: ${iframeCount}`);

    for (let i = 0; i < iframeCount; i++) {
      const iframe = iframeElements.nth(i);

      log("");
      log(`IFRAME ELEMENT #${i + 1}`);
      log(`  src:         ${await safeAttribute(iframe, "src")}`);
      log(`  id:          ${await safeAttribute(iframe, "id")}`);
      log(`  name:        ${await safeAttribute(iframe, "name")}`);
      log(`  title:       ${await safeAttribute(iframe, "title")}`);
      log(`  class:       ${await safeAttribute(iframe, "class")}`);
      log(`  aria-label:  ${await safeAttribute(iframe, "aria-label")}`);

      try {
        log(`  visible:     ${await iframe.isVisible()}`);
      } catch {
        log("  visible:     unknown");
      }
    }

    // --------------------------------------------------------
    // SEARCH FOR POSTCODE WORDING EVERYWHERE
    // --------------------------------------------------------

    separator();

    log("============================================================");
    log("POSTCODE SEARCH");
    log("============================================================");

    for (let i = 0; i < frames.length; i++) {
      const frame = frames[i];

      try {
        const bodyText = (
          await frame.locator("body").innerText()
        ).toLowerCase();

        const postcodePositions = [
          "postcode",
          "post code",
          "enter your postcode",
          "enter postcode",
          "location"
        ];

        const found = postcodePositions.filter(
          term => bodyText.includes(term)
        );

        log("");
        log(`FRAME ${i}: ${frame.url()}`);
        log(
          found.length
            ? `FOUND WORDS: ${found.join(", ")}`
            : "No postcode/location wording detected."
        );
      } catch (error) {
        log(
          `FRAME ${i}: Unable to inspect text - ${error.message}`
        );
      }
    }

    // --------------------------------------------------------
    // SCREENSHOT
    // --------------------------------------------------------

    separator();

    log("============================================================");
    log("SCREENSHOT");
    log("============================================================");

    await page.screenshot({
      path: screenshotFile,
      fullPage: true
    });

    log(`Diagnostic screenshot: ${screenshotFile}`);

    // --------------------------------------------------------
    // SAVE DIAGNOSTIC FILE
    // --------------------------------------------------------

    separator();

    log("Saving diagnostic report...");

    fs.writeFileSync(
      diagnosticFile,
      output.join("\n"),
      "utf8"
    );

    log(`Diagnostic report: ${diagnosticFile}`);

    separator();

    log("============================================================");
    log("DIAGNOSTIC COMPLETE");
    log("============================================================");
    log("");
    log("IMPORTANT:");
    log("This diagnostic intentionally does NOT attempt to submit");
    log("the postcode.");
    log("");
    log("Use the GitHub Actions log to identify:");
    log("1. The O2 iframe containing the postcode checker.");
    log("2. The exact input selector.");
    log("3. The exact submit button.");
    log("4. Any dynamically generated iframe URL.");
    log("5. Any postcode-related elements.");
    log("");

  } catch (error) {
    log("");
    log("============================================================");
    log("DIAGNOSTIC ERROR");
    log("============================================================");
    log(error.stack || error.message);
  } finally {
    await browser.close();
  }

  // Always exit successfully so we can inspect the diagnostics.
  process.exit(0);
}

main();
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
    const record = {};

    headers.forEach((header, index) => {
      record[header] = values[index] ?? "";
    });

    return record;
  });
}

function classifyStatus(text) {
  const lower = text.toLowerCase();

  /*
   * O2 outage wording.
   */
  const issuePatterns = [
    "nearby phone mast",
    "mast isn’t working",
    "mast isn't working",
    "phone mast isn't working",
    "phone mast isn’t working",
    "isn’t working as it should",
    "isn't working as it should",
    "service might come and go",
    "service may come and go",
    "outage",
    "fault",
    "engineers will be on the case",
    "engineer is working",
    "engineers are working",
    "not working",
    "problem with your service",
    "problem in your area",
    "issue in your area",
    "known issue",
    "maintenance"
  ];

  for (const pattern of issuePatterns) {
    if (lower.includes(pattern)) {
      return "ISSUE";
    }
  }

  const okPatterns = [
    "working normally",
    "no known issues",
    "no issues",
    "no issue",
    "there are no known problems",
    "everything is working"
  ];

  for (const pattern of okPatterns) {
    if (lower.includes(pattern)) {
      return "OK";
    }
  }

  return "UNKNOWN";
}

async function getAllFrames(page) {
  return page.frames();
}

async function findPostcodeInputInFrame(frame) {
  const selectors = [
    'input[placeholder*="postcode" i]',
    'input[aria-label*="postcode" i]',
    'input[name*="postcode" i]',
    'input[id*="postcode" i]',
    'input[data-testid*="postcode" i]',
    'input[placeholder*="post code" i]',
    'input[aria-label*="post code" i]',
    'input[type="text"]'
  ];

  for (const selector of selectors) {
    try {
      const locator = frame.locator(selector);

      const count = await locator.count();

      for (let i = 0; i < count; i++) {
        const candidate = locator.nth(i);

        if (await candidate.isVisible()) {
          return candidate;
        }
      }
    } catch {
      // Continue.
    }
  }

  return null;
}

async function findButtonInFrame(frame) {
  const selectors = [
    'button:has-text("Check")',
    'button:has-text("Search")',
    'button:has-text("Check status")',
    'button:has-text("Check Status")',
    'button:has-text("Submit")',
    'input[type="submit"]',
    '[role="button"]:has-text("Check")',
    '[role="button"]:has-text("Search")',
    'button'
  ];

  for (const selector of selectors) {
    try {
      const locator = frame.locator(selector);

      const count = await locator.count();

      for (let i = 0; i < count; i++) {
        const candidate = locator.nth(i);

        if (await candidate.isVisible()) {
          return candidate;
        }
      }
    } catch {
      // Continue.
    }
  }

  return null;
}

async function locateO2Postcode(page) {
  /*
   * First inspect the main page.
   */
  let input = await findPostcodeInputInFrame(page);

  if (input) {
    return {
      frame: page,
      input
    };
  }

  /*
   * Then inspect every iframe.
   *
   * This is the important part because O2's checker is
   * rendered inside an embedded application.
   */
  const frames = await getAllFrames(page);

  console.log(`Found ${frames.length} browser frame(s).`);

  for (const frame of frames) {
    try {
      console.log(`Inspecting frame: ${frame.url()}`);

      input = await findPostcodeInputInFrame(frame);

      if (input) {
        console.log("Postcode input found.");
        return {
          frame,
          input
        };
      }
    } catch {
      // Continue.
    }
  }

  return null;
}

async function waitForO2Result(frame, page) {
  /*
   * Give React time to render the result.
   */
  await page.waitForTimeout(3000);

  /*
   * Look for known O2 result text.
   */
  const resultPatterns = [
    "nearby phone mast",
    "mast isn’t working",
    "mast isn't working",
    "service might come and go",
    "service may come and go",
    "engineers will be on the case",
    "working normally",
    "no known issues",
    "no issues",
    "outage",
    "fault",
    "maintenance"
  ];

  for (let attempt = 0; attempt < 10; attempt++) {
    let text = "";

    try {
      text = await frame.locator("body").innerText();
    } catch {
      try {
        text = await page.locator("body").innerText();
      } catch {
        text = "";
      }
    }

    const lower = text.toLowerCase();

    if (resultPatterns.some(pattern => lower.includes(pattern))) {
      console.log("O2 result detected.");
      return text;
    }

    await page.waitForTimeout(1000);
  }

  try {
    return await frame.locator("body").innerText();
  } catch {
    return await page.locator("body").innerText();
  }
}

function extractResolution(text) {
  const patterns = [
    /updated\s+([0-9]{1,2}:[0-9]{2})/i,
    /updated\s+([0-9]{1,2}:[0-9]{2}\s*(?:am|pm))/i,
    /expected.*?([0-9]{1,2}:[0-9]{2})/i,
    /estimated.*?([0-9]{1,2}:[0-9]{2})/i,
    /resolved.*?([0-9]{1,2}:[0-9]{2})/i,
    /fixed.*?([0-9]{1,2}:[0-9]{2})/i
  ];

  for (const pattern of patterns) {
    const match = text.match(pattern);

    if (match) {
      return match[0];
    }
  }

  return "";
}

async function captureScreenshot(page, screenshotPath) {
  await page.screenshot({
    path: screenshotPath,
    fullPage: true
  });

  console.log(`Screenshot saved: ${screenshotPath}`);
}

function writeCsv(record, existing) {
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

  return [...existing, record];
}

function writeExcel(records) {
  const worksheet = XLSX.utils.json_to_sheet(records);

  worksheet["!cols"] = [
    { wch: 14 },
    { wch: 12 },
    { wch: 24 },
    { wch: 14 },
    { wch: 18 },
    { wch: 90 },
    { wch: 30 },
    { wch: 22 },
    { wch: 35 },
    { wch: 65 },
    { wch: 18 },
    { wch: 18 },
    { wch: 18 },
    { wch: 40 }
  ];

  const workbook = XLSX.utils.book_new();

  XLSX.utils.book_append_sheet(
    workbook,
    worksheet,
    "O2 Mast Log"
  );

  const issueCount = records.filter(
    record => record["O2 Status"] === "ISSUE"
  ).length;

  const okCount = records.filter(
    record => record["O2 Status"] === "OK"
  ).length;

  const unknownCount = records.filter(
    record => record["O2 Status"] === "UNKNOWN"
  ).length;

  const failedCount = records.filter(
    record => record["Check Result"] === "FAILED"
  ).length;

  const successfulCount = records.filter(
    record => record["Check Result"] === "SUCCESS"
  ).length;

  const summary = XLSX.utils.aoa_to_sheet([
    ["O2 NW9 0RY NETWORK EVIDENCE"],
    [],
    ["Postcode", POSTCODE],
    ["Complaint Reference", COMPLAINT_REFERENCE],
    ["Monitoring Source", O2_URL],
    [],
    ["Total Checks", records.length],
    ["Successful Checks", successfulCount],
    ["Failed Checks", failedCount],
    ["Issue Checks", issueCount],
    ["OK Checks", okCount],
    ["Unknown Checks", unknownCount],
    [],
    [
      "IMPORTANT",
      "Only SUCCESS checks are considered valid automated O2 checks. FAILED checks are not evidence of an O2 outage."
    ],
    [
      "PRIMARY EVIDENCE",
      "Timestamped screenshots stored in the evidence folder."
    ],
    [
      "POSTCODE",
      POSTCODE
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

  XLSX.writeFile(workbook, xlsxFile);

  console.log(`Workbook saved: ${xlsxFile}`);
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
  let previousStatus = "";

  const screenshotName =
    `${fileTimestamp}_NW9-0RY.png`;

  const screenshotPath =
    path.join(evidenceDir, screenshotName);

  const existing = readExistingCsv();

  if (existing.length) {
    previousStatus =
      existing[existing.length - 1]["O2 Status"] || "";
  }

  try {
    console.log("Opening O2 status page...");

    await page.goto(O2_URL, {
      waitUntil: "domcontentloaded",
      timeout: 60000
    });

    await page.waitForTimeout(5000);

    console.log(`Main page URL: ${page.url()}`);

    console.log("Searching for O2 postcode checker...");

    const located = await locateO2Postcode(page);

    if (!located) {
      /*
       * IMPORTANT:
       * Even on failure we capture the actual O2 page.
       * This means the screenshot shows what the monitor
       * actually saw at that exact time.
       */
      message =
        "MONITORING ERROR: Could not find O2 postcode input in the main page or embedded O2 iframe.";

      console.log(message);

      await captureScreenshot(
        page,
        screenshotPath
      );

      throw new Error(message);
    }

    const { frame, input } = located;

    console.log("Entering postcode:");

    /*
     * Explicitly use a SPACE.
     *
     * NW9 0RY
     */
    await input.fill("");

    await input.fill(POSTCODE);

    /*
     * Verify what was actually entered.
     */
    const enteredValue =
      await input.inputValue();

    console.log(
      `Postcode field contains: "${enteredValue}"`
    );

    if (enteredValue !== POSTCODE) {
      throw new Error(
        `Postcode verification failed. Expected "${POSTCODE}" but field contains "${enteredValue}".`
      );
    }

    console.log("Postcode verified.");

    const submitButton =
      await findButtonInFrame(frame);

    if (submitButton) {
      console.log("O2 check button found.");

      await submitButton.scrollIntoViewIfNeeded();

      await submitButton.click();

      console.log("O2 check button clicked.");
    } else {
      console.log(
        "No dedicated button found. Pressing Enter."
      );

      await input.press("Enter");
    }

    console.log(
      "Waiting for O2 result..."
    );

    const resultText =
      await waitForO2Result(
        frame,
        page
      );

    message = resultText
      .replace(/\s+/g, " ")
      .trim()
      .slice(0, 10000);

    status =
      classifyStatus(resultText);

    expectedResolution =
      extractResolution(resultText);

    /*
     * The check is only SUCCESS if we found the field,
     * entered the exact postcode, submitted it and
     * obtained the O2 result page.
     */
    checkResult = "SUCCESS";

    console.log(
      `O2 status classified as: ${status}`
    );

    console.log(
      "Capturing post-check screenshot..."
    );

    await captureScreenshot(
      page,
      screenshotPath
    );

  } catch (error) {
    if (!message) {
      message =
        `MONITORING ERROR: ${error.message}`;
    }

    status =
      status === "UNKNOWN"
        ? "UNKNOWN"
        : status;

    checkResult = "FAILED";

    /*
     * Always capture a screenshot on failure too.
     */
    if (!fs.existsSync(screenshotPath)) {
      try {
        await captureScreenshot(
          page,
          screenshotPath
        );
      } catch {
        console.error(
          "Could not capture failure screenshot."
        );
      }
    }

    console.error(message);
  }

  const statusChanged =
    !previousStatus ||
    previousStatus !== status
      ? "YES"
      : "NO";

  const record = {
    "Date":
      timestamp.split(",")[0],

    "Time":
      timestamp.split(",")[1]?.trim() ||
      timestamp,

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
      previousStatus,

    "Notes":
      checkResult === "SUCCESS"
        ? "Automated O2 postcode check completed. Screenshot captured after result."
        : "Automated check failed. This record is not considered evidence of an O2 outage."
  };

  const allRecords =
    writeCsv(record, existing);

  writeExcel(allRecords);

  console.log("");
  console.log("======================================");
  console.log("O2 NW9 0RY MONITOR RESULT");
  console.log("======================================");
  console.log(`Timestamp: ${timestamp}`);
  console.log(`Postcode: ${POSTCODE}`);
  console.log(`Status: ${status}`);
  console.log(`Check Result: ${checkResult}`);
  console.log(`Screenshot: ${screenshotName}`);
  console.log("======================================");

  await browser.close();
}

main().catch(error => {
  console.error(
    "FATAL MONITOR ERROR:",
    error
  );

  process.exit(1);
});
