const { Builder, By, Key, until } = require("selenium-webdriver");
const assert = require("assert");
const fs = require("fs");
const path = require("path");
const xlsx = require("xlsx");
const { parse } = require("json2csv"); // CSV conversion

// Helper function to write logs to JSON and CSV files
function writeLogToJsonAndCsv(logType, logData) {
  const jsonFilePath = `logs/json/${logType}_log.json`;
  const csvFilePath = `logs/csv/${logType}_log.csv`;

  // Ensure the logs directory and subdirectories exist
  if (!fs.existsSync("logs")) {
    fs.mkdirSync("logs");
  }
  if (!fs.existsSync("logs/json")) {
    fs.mkdirSync("logs/json");
  }
  if (!fs.existsSync("logs/csv")) {
    fs.mkdirSync("logs/csv");
  }

  // Write to JSON log file (append new data)
  try {
    let jsonLog = [];
    if (fs.existsSync(jsonFilePath)) {
      const fileContent = fs.readFileSync(jsonFilePath, "utf-8");
      if (fileContent.trim()) {
        jsonLog = JSON.parse(fileContent);
      }
    }
    jsonLog.push(logData);
    fs.writeFileSync(jsonFilePath, JSON.stringify(jsonLog, null, 2));
  } catch (error) {
    console.error("Error writing JSON log:", error);
  }

  // Write to CSV log file (append new data)
  try {
    const row =
      [
        logData.NIK || "",
        logData.TGLLHR || "",
        logData.status || "",
        logData.timestamp || "",
        logData.error_message || "",
        logData.retry_count || "",
        logData.step || "",
      ].join(",") + "\n";

    fs.appendFileSync(csvFilePath, row);
  } catch (error) {
    console.error("Error writing CSV log:", error);
  }
}

// Function to process NIK and TGLLHR (Tanggal Lahir)
async function processNIKandTGLLHR(driver, row) {
  try {
    const tanggalLahir = excelSerialToDate(row["TGLLHR"]);
    if (!tanggalLahir) {
      console.error("Failed to process Tanggal Lahir for NIK:", row["NIK"]);
      writeLogToJsonAndCsv("failure", {
        NIK: row["NIK"],
        TGLLHR: row["TGLLHR"],
        status: "failure",
        timestamp: new Date().toISOString(),
        error_message: "Invalid NIK or TGLLHR",
        retry_count: 1,
        step: "NIK and TGLLHR Validation",
      });
      return null; // Return null if date is invalid
    }

    // Process NIK and Tanggal Lahir (Birth Date)
    await driver.findElement(By.id("nik_txt")).sendKeys(row["NIK"]);
    await driver.findElement(By.id("TglLahir_src")).sendKeys(tanggalLahir);
    await driver.findElement(By.id("TglLahir_src")).sendKeys(Key.ESCAPE);

    console.log(`Processed NIK: ${row["NIK"]}, Tanggal Lahir: ${tanggalLahir}`);

    return tanggalLahir; // Return Tanggal Lahir if successful
  } catch (error) {
    console.error("Error processing NIK and TGLLHR:", error);
    writeLogToJsonAndCsv("failure", {
      NIK: row["NIK"],
      TGLLHR: row["TGLLHR"],
      status: "failure",
      timestamp: new Date().toISOString(),
      error_message: error.message,
      retry_count: 1,
      step: "NIK and TGLLHR Processing",
    });
    return null; // Return null if an error occurs
  }
}

// Convert Excel serial date to JavaScript Date
function excelSerialToDate(serialDate) {
  if (serialDate === null || serialDate === undefined) {
    return null;
  }
  const date = new Date(Math.round((serialDate - 25569) * 86400 * 1000));
  return date.toLocaleDateString("en-GB", {
    day: "2-digit",
    month: "2-digit",
    year: "numeric",
  });
}

// Retry logic for retryable operations
async function retryOperation(driver, operation, retries = 1) {
  let attempt = 0;
  while (attempt < retries) {
    try {
      return await operation();
    } catch (error) {
      attempt++;
      if (attempt >= retries) {
        throw error; // Max retries reached
      }
      await new Promise((resolve) => setTimeout(resolve, 1000));
    }
  }
}

// Function to handle CAPTCHA
async function handleCaptcha(driver) {
  console.log("Waiting for captcha");
  await driver.executeScript("window.scrollTo(0, document.body.scrollHeight);");
  await driver.findElement(By.id("captchaCode_txt")).click();

  let inputCaptcha = await driver
    .findElement(By.id("captchaCode_txt"))
    .getAttribute("value");
  let startTime = Date.now();
  let timeElapsed = 0;
  while (inputCaptcha.length < 5 && timeElapsed < 7000) {
    await driver.sleep(200);
    inputCaptcha = await driver
      .findElement(By.id("captchaCode_txt"))
      .getAttribute("value");
  }
}

// Function to handle login and success check
async function handleLogin(driver, row) {
  try {
    console.log("Click login");
    await driver.findElement(By.id("btnCariPetugas")).click();

    // Wait for any success buttons to appear with timeout
    try {
      await driver.wait(
        async () => {
          const successButtons = await driver.findElements(
            By.xpath("//button[contains(text(),'Setuju')]")
          );
          console.log("Found success buttons:", successButtons.length);

          if (successButtons.length > 0) {
            await driver.wait(until.elementIsVisible(successButtons[0]), 2000);
            await successButtons[0].click();
            console.log("Click success button");
            return true;
          }
          return false;
        },
        5000,
        "Timeout waiting for success button"
      );

      // If success button was found and clicked
      logCompletion(row, "Filled before");
      return false;
    } catch (successError) {
      // If no success button found, try back button
      await driver.executeScript(
        "window.scrollTo(0, document.body.scrollHeight);"
      );
      const backButton = await driver.wait(
        until.elementLocated(By.id("btnBacktoHome1")),
        5000,
        "Timeout waiting for back button"
      );
      await driver.sleep(1000);
      await backButton.click();
      console.log("Click back button");

      logCompletion(row, "Filled before");
      console.log("Screening successful, found back button");
      return false;
    }
  } catch (error) {
    console.log("No success/back button found, proceeding with login", error);
    return true;
  }
}

// Function to fill out form
async function fillOutForm(driver) {
  console.log("Fill out form");
  await driver.executeScript("window.scrollTo(0, document.body.scrollHeight);");

  await driver.findElement(By.id("beratBadan_txt")).sendKeys("60");
  console.log("BB filled");
  await driver.findElement(By.id("tinggiBadan_txt")).sendKeys("160");
  console.log("TB filled");
  await driver.findElement(By.id("nextGenBtn")).click();
  console.log("button clicked");
}

// Function to handle answering questions
async function handleQuestionSet(driver) {
  await driver.wait(
    until.elementsLocated(By.css(".answer-item:nth-child(2) > .answertext")),
    5000
  );
  const elements = await driver.findElements(
    By.css(".answer-item:nth-child(2) > .answertext")
  );
  for (const element of elements) {
    await driver.executeScript("arguments[0].click();", element);
    console.log("Answered a question");
  }
  await driver.findElement(By.id("nextGenBtn")).click();
}

// Function to submit form
async function submitForm(driver) {
  let attempts = 0;
  while (true) {
    let saveButton;
    try {
      saveButton = await driver.findElement(
        By.xpath("//button[contains(text(),'Setuju')]")
      );
      await driver.executeScript("arguments[0].click();", saveButton);
      // await saveButton.click();
      console.log("Clicked save button, form completed");
      break;
    } catch (error) {
      await handleQuestionSet(driver);
      attempts++;
    }
  }
}

// Function to log completion
function logCompletion(row, status = "Success") {
  writeLogToJsonAndCsv("success", {
    NIK: row["NIK"],
    TGLLHR: row["TGLLHR"],
    status: status,
    timestamp: new Date().toISOString(),
    step: "Login Success",
    error_message: "",
    retry_count: "",
  });
}

// Function to log failure
function logFailure(row, errorMessage, step = "Unknown") {
  writeLogToJsonAndCsv("failure", {
    NIK: row["NIK"],
    TGLLHR: row["TGLLHR"],
    status: "failure",
    timestamp: new Date().toISOString(),
    error_message: errorMessage,
    retry_count: 1,
    step: step,
  });
}

describe("4 Sessions Layout Test", function () {
  this.timeout(0); // Set Mocha's timeout for test cases

  let drivers = [];
  const sessionState = 4; // Set to 1 for a single session, 4 for multiple sessions

  beforeEach(async function () {
    drivers = [];

    const scaleFactor = 1.5;
    const positions = [
      { x: 0, y: 0 },
      { x: 1280 / scaleFactor, y: 0 },
      { x: 0, y: 800 / scaleFactor },
      { x: 1280 / scaleFactor, y: 800 / scaleFactor },
    ];

    if (sessionState === 1) {
      const driver = await new Builder().forBrowser("chrome").build();
      drivers.push(driver);
    } else if (sessionState === 4) {
      const size = { width: 1280 / scaleFactor, height: 800 / scaleFactor };
      for (let i = 0; i < 4; i++) {
        const driver = await new Builder().forBrowser("chrome").build();
        await driver.manage().window().setRect({
          width: size.width,
          height: size.height,
          x: positions[i].x,
          y: positions[i].y,
        });
        drivers.push(driver);
      }
    } else if (sessionState === 2) {
      const size = { width: 1280 / scaleFactor, height: 800 / scaleFactor };
      for (let i = 0; i < 2; i++) {
        const driver = await new Builder().forBrowser("chrome").build();
        await driver.manage().window().setRect({
          width: size.width,
          height: size.height,
          x: positions[i].x,
          y: positions[i].y,
        });
        drivers.push(driver);
      }
    }
  });

  afterEach(async function () {
    for (const driver of drivers) {
      await driver.quit();
    }
  });

  it("should fill the form", async function () {
    const workbook = xlsx.readFile("memeysel.xlsx");
    const worksheet = workbook.Sheets[workbook.SheetNames[0]];
    const data = xlsx.utils.sheet_to_json(worksheet);

    // Read success log to get processed NIKs
    let processedNIKs = new Set();
    try {
      const successLog = JSON.parse(
        fs.readFileSync("./logs/json/success_log.json", "utf8")
      );
      processedNIKs = new Set(successLog.map((entry) => entry.NIK));
    } catch (error) {
      console.log(
        "No existing success log found or error reading it:",
        error.message
      );
    }

    let currentIndex = 0;
    const totalSessions = sessionState === 2 ? 2 : sessionState === 4 ? 4 : 1;
    const promises = [];

    while (currentIndex < data.length) {
      const sessionPromises = drivers.map(async (driver, sessionIndex) => {
        const dataIndex = currentIndex + sessionIndex;
        if (dataIndex >= data.length) return; // Skip if we've processed all data

        const row = data[dataIndex];

        // Skip if NIK is already processed
        if (processedNIKs.has(row.NIK)) {
          console.log(`Skipping already processed NIK: ${row.NIK}`);
          return;
        }

        try {
          await retryOperation(driver, async () => {
            const url = "https://webskrining.bpjs-kesehatan.go.id/skrining";
            await driver.get(url);

            const tanggalLahir = await processNIKandTGLLHR(driver, row);
            if (!tanggalLahir) {
              throw new Error("Failed to process NIK and TGLLHR");
            }

            await handleCaptcha(driver);
            const loginSuccess = await handleLogin(driver, row);
            if (loginSuccess) return true;

            await fillOutForm(driver);
            await submitForm(driver);

            logCompletion(row);
            return true;
          });
        } catch (error) {
          const step = error.message.includes("Step:")
            ? error.message.split("Step:")[1].trim().replace(")", "")
            : "Unknown";
          await logFailure(row, error.message, step);
        }
      });

      promises.push(...sessionPromises);
      currentIndex += totalSessions; // Increment by number of sessions
      await Promise.all(sessionPromises); // Wait for current batch to complete
    }

    await Promise.all(promises);
  });
});
