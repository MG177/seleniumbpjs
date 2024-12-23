const { Builder, By, Key, until } = require("selenium-webdriver");
const assert = require("assert");
const fs = require("fs");
const path = require("path");
const xlsx = require("xlsx");
const { parse } = require("json2csv");

// Import helper functions from main.spec.js
const {
  writeLogToJsonAndCsv,
  processNIKandTGLLHR,
  excelSerialToDate,
  retryOperation,
  handleCaptcha,
  handleLogin,
  fillOutForm,
  handleQuestionSet,
  submitForm,
  logCompletion,
  logFailure,
} = require("./main.spec.js");

describe("Session 2 Test", function () {
  this.timeout(0);

  let driver;
  const maxRetries = 3;
  const batchSize = 20;

  beforeEach(async function () {
    driver = await new Builder().forBrowser("chrome").build();
    // Position window in top-right quadrant
    await driver
      .manage()
      .window()
      .setRect({
        width: 1280 / 1.5,
        height: 800 / 1.5,
        x: 1280 / 1.5,
        y: 0,
      });
  });

  afterEach(async function () {
    if (driver) {
      try {
        await driver.quit();
      } catch (error) {
        console.error("Error closing session:", error);
      }
    }
  });

  it("should fill the form", async function () {
    const workbook = xlsx.readFile("memeysel.xlsx");
    const worksheet = workbook.Sheets[workbook.SheetNames[2]];
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

    // Process data in batches
    for (
      let batchStart = 0;
      batchStart < data.length;
      batchStart += batchSize
    ) {
      const batchEnd = Math.min(batchStart + batchSize, data.length);
      const currentBatch = data.slice(batchStart, batchEnd);
      console.log(
        `Processing batch ${batchStart / batchSize + 1}, entries ${
          batchStart + 1
        } to ${batchEnd}`
      );

      for (const row of currentBatch) {
        if (processedNIKs.has(row.NIK)) {
          console.log(`Skipping already processed NIK: ${row.NIK}`);
          continue;
        }

        let retryCount = 0;
        let success = await retryOperation(
          driver,
          async () => {
            try {
              const url = "https://webskrining.bpjs-kesehatan.go.id/skrining";
              await driver.get(url);

              const tanggalLahir = await processNIKandTGLLHR(driver, row);
              if (!tanggalLahir)
                throw new Error("Failed to process NIK and TGLLHR");

              await handleCaptcha(driver);
              const loginSuccess = await handleLogin(driver, row);
              if (!loginSuccess) {
                console.log(`Login failed for NIK: ${row.NIK}`);
                return false; // This will be returned by retryOperation
              }

              await fillOutForm(driver);
              await submitForm(driver);
              await logCompletion(row);
              return true; // Success case
            } catch (error) {
              throw error;
            }
          },
          maxRetries
        );

        if (!success) {
          console.log(
            `Skipping NIK ${row.NIK} due to login failure or data already filled before`
          );
          continue; // Continue to next row in the main loop
        }

        processedNIKs.add(row.NIK);
      }

      // Small delay between batches to prevent overload
      await new Promise((resolve) => setTimeout(resolve, 2000));
    }
  });
});
