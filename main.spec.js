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
async function retryOperation(
  driver,
  operation,
  step = "Unknown",
  retries = 3
) {
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
    console.log(inputCaptcha, Date.now() - startTime);
  }
}

// Function to handle login and success check
async function handleLogin(driver, row) {
  console.log("Click login");
  await driver.findElement(By.id("btnCariPetugas")).click();

  try {
    // First try to find and click success button
    await driver.sleep(1000); // Wait for 1 second
    const successButtons = await driver.findElements(
      By.xpath("//button[contains(text(),'Setuju')]")
    );

    if (successButtons.length > 0) {
      await driver.wait(until.elementIsVisible(successButtons[0]), 5000);
      await successButtons[0].click();
      return false;
    }

    // If no success button, try back button
    await driver.executeScript(
      "window.scrollTo(0, document.body.scrollHeight);"
    );
    console.log("Click back button");

    const backButton = await driver.findElement(By.id("btnBacktoHome1"));

    console.log("Found back button:", backButton);

    await driver.wait(until.elementIsVisible(backButton), 5000);
    await driver.wait(until.elementIsEnabled(backButton), 5000);

    logCompletion(row, "Filled before");
    console.log("Screening successful, found back button");
    return true;
  } catch (error) {
    console.log("Error in handling result:", error.message);
    return false;
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
      // await driver.executeScript("arguments[0].click();", saveButton);
      await saveButton.click();
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
    }
  });

  afterEach(async function () {
    // Quit all 4 drivers
    for (let driver of drivers) {
      await driver.quit();
    }
  });

  it("Should Process Each NIK", async function () {
    const workbook = xlsx.readFile("memeysel.xlsx");
    const worksheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = xlsx.utils.sheet_to_json(worksheet);

    const url = "https://webskrining.bpjs-kesehatan.go.id/skrining";

    const operations = drivers.map(async (driver, driverIndex) => {
      for (const row of rows) {
        try {
          // Skip if NIK already processed
          const successLogPath = "logs/json/success_log.json";
          if (fs.existsSync(successLogPath)) {
            const fileContent = fs.readFileSync(successLogPath, "utf-8");
            if (fileContent.trim()) {
              const successLog = JSON.parse(fileContent);
              if (successLog.some((item) => item.NIK === row["NIK"])) {
                console.log(
                  `Skipping NIK ${row["NIK"]} (Driver ${driverIndex + 1})`
                );
                continue;
              }
            }
          }

          const result = await retryOperation(
            driver,
            async () => {
              const step = "";
              await driver.get(url);
              const tanggalLahir = await processNIKandTGLLHR(driver, row);
              if (!tanggalLahir) {
                throw new Error("Failed to process NIK and TGLLHR");
              }

              try {
                await handleCaptcha(driver);
              } catch (error) {
                throw new Error(`Captcha handling failed: ${error}`);
              }

              try {
                const loginSuccess = await handleLogin(driver, row);
                if (loginSuccess) return true;
              } catch (error) {
                throw new Error(`Login process failed: ${error}`);
              }

              try {
                await fillOutForm(driver);
              } catch (error) {
                throw new Error(`Form filling failed: ${error}`);
              }

              try {
                await submitForm(driver);
              } catch (error) {
                throw new Error(`Form submission failed: ${error}`);
              }

              logCompletion(row);
              return true;
            },
            "Process Flow"
          );

          if (result) continue;
        } catch (error) {
          const step = error.message.includes("Step:")
            ? error.message.split("Step:")[1].trim().replace(")", "")
            : "Unknown";
          logFailure(row, error.message.split("(Step:")[0].trim(), step);
        }
      }
    });

    await Promise.all(operations); // Wait for all drivers to process their tasks concurrently
  });
});
