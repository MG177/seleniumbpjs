const { Builder, By, Key, until } = require("selenium-webdriver");
const assert = require("assert");
const fs = require("fs");
const path = require("path");
const xlsx = require("xlsx");
const { parse } = require("json2csv"); // CSV conversion

// Helper function to write logs to JSON and CSV files
function writeLogToJsonAndCsv(logType, logData) {
  const jsonFilePath = `logs/${logType}_log.json`;
  const csvFilePath = `logs/${logType}_log.csv`;

  // Ensure the logs directory exists
  if (!fs.existsSync("logs")) {
    fs.mkdirSync("logs");
  }

  // Write to JSON log file (append new data)
  try {
    let jsonLog = [];
    if (fs.existsSync(jsonFilePath)) {
      jsonLog = JSON.parse(fs.readFileSync(jsonFilePath, "utf-8"));
    }
    jsonLog.push(logData);
    fs.writeFileSync(jsonFilePath, JSON.stringify(jsonLog, null, 2));
  } catch (error) {
    console.error("Error writing JSON log:", error);
  }

  // Write to CSV log file (append new data)
  try {
    let csvLog = [];
    if (fs.existsSync(csvFilePath)) {
      const csvContent = fs.readFileSync(csvFilePath, "utf-8");
      csvLog = csvContent.split("\n").map((line) => line.split(","));
    }

    // Add new row to CSV log
    const row = [
      logData.NIK,
      logData.TGLLHR,
      logData.status,
      logData.timestamp,
      logData.error_message || "",
      logData.retry_count || "",
    ];
    csvLog.push(row);

    // Convert array to CSV and write to file
    const csvString = parse(csvLog);
    fs.writeFileSync(csvFilePath, csvString);
  } catch (error) {
    console.error("Error writing CSV log:", error);
  }
}

// Function to retry an operation
async function retryOperation(driver, operation, retries = 3) {
  let attempt = 0;
  while (attempt < retries) {
    try {
      await operation();
      return; // Success, exit the retry loop
    } catch (error) {
      attempt++;
      if (attempt >= retries) {
        throw error; // Max retries reached
      }
      console.log(`Retrying operation for NIK, attempt ${attempt}...`);
      await driver.sleep(1000); // Wait before retrying
    }
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

describe("4 Sessions Layout Test", function () {
  this.timeout(0); // Set Mocha's timeout for test cases

  let drivers = [];
  let vars;

  // Launch 4 browser windows with 2x2 layout
  beforeEach(async function () {
    vars = {};

    // Scale factors
    const scaleFactor = 1.5;

    // Positions (logical resolution for 150% scale, as per your screen size)
    const positions = [
      { x: 0, y: 0 }, // Top-left
      { x: 1280 / scaleFactor, y: 0 }, // Top-right
      { x: 0, y: 800 / scaleFactor }, // Bottom-left
      { x: 1280 / scaleFactor, y: 800 / scaleFactor }, // Bottom-right
    ];

    // Logical size for each window (scaled size)
    const size = { width: 1280 / scaleFactor, height: 800 / scaleFactor };

    // Launch 4 browser windows with 2x2 layout, applying scaling
    for (let i = 0; i < 4; i++) {
      const driver = await new Builder().forBrowser("chrome").build();
      // Set both position and size using setRect() while considering scale factor
      await driver.manage().window().setRect({
        width: size.width,
        height: size.height,
        x: positions[i].x,
        y: positions[i].y,
      });
      drivers.push(driver);
    }
  });

  afterEach(async function () {
    // Quit all 4 drivers
    for (let driver of drivers) {
      await driver.quit();
    }
  });

  it("should process data in 4 sessions", async function () {
    // Load Excel file
    const filePath = "C:\\Users\\Lenovo\\Downloads\\memeysel.xlsx"; // Update with your Excel file path
    const workbook = xlsx.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const data = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]);

    // Distribute the data into 4 groups for parallel processing
    const chunks = [[], [], [], []];
    for (let i = 0; i < data.length; i++) {
      chunks[i % 4].push(data[i]);
    }

    // Start 4 separate sessions
    await Promise.all(
      drivers.map(async (driver, index) => {
        for (const row of chunks[index]) {
          try {
            const timestamp = new Date().toISOString();
            writeLogToJsonAndCsv("main", {
              NIK: row["NIK"],
              TGLLHR: row["TGLLHR"],
              status: "started",
              timestamp: timestamp,
            });

            // Retry operation in case of failure
            await retryOperation(driver, async () => {
              await driver.get(
                "https://webskrining.bpjs-kesehatan.go.id/skrining"
              );
              await driver.manage().window().setRect({ width: 866 });
            });

            // Process NIK and TGLLHR
            const tanggalLahir = excelSerialToDate(row["TGLLHR"]);
            if (!tanggalLahir) {
              console.error(
                "Failed to process Tanggal Lahir for NIK:",
                row["NIK"]
              );
              writeLogToJsonAndCsv("failure", {
                NIK: row["NIK"],
                TGLLHR: row["TGLLHR"],
                status: "failure",
                timestamp: timestamp,
                error_message: "Invalid Tanggal Lahir",
                retry_count: 1,
              });
              continue; // Skip to next row if date is invalid
            }

            await driver.findElement(By.id("nik_txt")).sendKeys(row["NIK"]);
            await driver
              .findElement(By.id("TglLahir_src"))
              .sendKeys(tanggalLahir);
            await driver
              .findElement(By.id("TglLahir_src"))
              .sendKeys(Key.ESCAPE);

            await driver.executeScript(
              "window.scrollTo(0, document.body.scrollHeight);"
            );

            // Handle captcha and login process
            console.log("Waiting for captcha");
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

            console.log("Click login");
            await driver.findElement(By.id("btnCariPetugas")).click();

            // Check for success buttons and results
            const successButtons = await driver.wait(
              until.elementsLocated(
                By.css(
                  "body > div.bootbox.modal.fade.bootbox-confirm.in > div > div > div.modal-footer > button.btn.btn-success"
                )
              ),
              10000
            );

            const hasilElements = await driver.findElements(
              By.id("hasilSkrJudul_Top")
            );

            if (hasilElements.length > 0 && successButtons.length == 0) {
              const hasilText = await hasilElements[0].getText();
              writeLogToJsonAndCsv("success", {
                NIK: row["NIK"],
                TGLLHR: row["TGLLHR"],
                status: "success",
                timestamp: new Date().toISOString(),
              });
              console.log("Hasil skrining ditemukan:", hasilText);
              continue;
            }

            if (successButtons.length > 0) {
              await driver.executeScript(
                "arguments[0].click();",
                successButtons[0]
              );
              console.log("Clicked .btn-success, continuing with form");
            }

            // Fill out form
            await driver.findElement(By.id("beratBadan_txt")).click();
            await driver.findElement(By.id("beratBadan_txt")).sendKeys("60");
            await driver.findElement(By.id("tinggiBadan_txt")).sendKeys("160");
            await driver.findElement(By.id("nextGenBtn")).click();

            // Answer questions
            async function handleQuestionSet(driver) {
              await driver.wait(
                until.elementsLocated(
                  By.css(".answer-item:nth-child(2) > .answertext")
                ),
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

            let attempts = 0;
            while (true) {
              let saveButton;
              try {
                saveButton = await driver.findElement(
                  By.xpath("/html/body/div[6]/div/div/div[3]/button[2]")
                );
                await driver.executeScript("arguments[0].click();", saveButton);
                console.log("Clicked save button, form completed");
                break;
              } catch (error) {
                await handleQuestionSet(driver);
                attempts++;
              }
            }

            writeLogToJsonAndCsv("success", {
              NIK: row["NIK"],
              TGLLHR: row["TGLLHR"],
              status: "completed",
              timestamp: new Date().toISOString(),
            });
          } catch (error) {
            console.error("Error processing data for NIK:", row["NIK"], error);
            writeLogToJsonAndCsv("failure", {
              NIK: row["NIK"],
              TGLLHR: row["TGLLHR"],
              status: "failure",
              timestamp: new Date().toISOString(),
              error_message: error.message,
              retry_count: 1,
            });
          }
        }
      })
    );
  });
});
