const puppeteer = require("puppeteer");
const xlsx = require("xlsx");
const fs = require("fs");

// Path ke file Excel
const excelFilePath = "C:\\Users\\Lenovo\\Downloads\\memeysel.xlsx";

// Baca data dari Excel
function readExcel(filePath) {
  const workbook = xlsx.readFile(filePath);
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  return xlsx.utils.sheet_to_json(sheet);
}

// Fungsi untuk menulis log
function writeLog(message) {
  const logFile = "log_pengisian.log";
  fs.appendFileSync(logFile, `${new Date().toISOString()} - ${message}\n`);
}

function excelSerialToDate(serial) {
  if (!serial || isNaN(serial)) {
    console.error("Invalid excel serial date:", serial);
    return "";
  }

  // Excel's starting date is January 1, 1900
  const excelStartDate = new Date(1900, 0, 1);
  // Adjust for the Excel leap year bug (there is no February 29, 1900)
  const adjustedSerial = serial - 2;
  // Add the serial number as days to the starting date
  const resultDate = new Date(
    excelStartDate.getTime() + adjustedSerial * 24 * 60 * 60 * 1000
  );

  // Validate result date
  if (isNaN(resultDate.getTime())) {
    console.error("Invalid date result for serial:", serial);
    return "";
  }

  // Format the date as DD-MM-YYYY
  const day = String(resultDate.getDate()).padStart(2, "0");
  const month = String(resultDate.getMonth() + 1).padStart(2, "0"); // Months are 0-based
  const year = resultDate.getFullYear();

  return `${day}-${month}-${year}`;
}

// Fungsi untuk delay
const delay = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

// Fungsi untuk retry
async function retryOperation(operation, maxAttempts = 3, delay = 5000) {
  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    try {
      return await operation();
    } catch (error) {
      if (attempt === maxAttempts) throw error;
      writeLog(`Attempt ${attempt} failed, retrying after ${delay}ms...`);
      await new Promise((resolve) => setTimeout(resolve, delay));
    }
  }
}

(async () => {
  const data = readExcel(excelFilePath);
  const browser = await puppeteer.launch({
    headless: false,
    defaultViewport: null,
    timeout: 60000,
    args: [
      "--start-maximized",
      "--disable-extensions",
      "--no-sandbox",
      "--disable-setuid-sandbox",
    ],
  });

  try {
    const page = await browser.newPage();
    // Set default navigation timeout
    page.setDefaultNavigationTimeout(60000);
    // Set default timeout
    page.setDefaultTimeout(60000);

    for (const row of data) {
      try {
        writeLog(`Proses data untuk NIK: ${row["NIK"]}`);

        // Retry navigation if it fails
        await retryOperation(async () => {
          await page.goto(
            "https://webskrining.bpjs-kesehatan.go.id/skrining/index.html",
            {
              waitUntil: "networkidle2",
              timeout: 60000,
            }
          );
        });

        // Tunggu elemen NIK muncul
        await retryOperation(async () => {
          await page.waitForSelector("#noKartu_txt", { timeout: 30000 });
        });

        console.log(row["TGLLHR"]);
        console.log(row["NIK"]);

        console.log(data);
        console.log(excelSerialToDate(row["TGLLHR"]));
        // Isi NIK
        await page.evaluate(() => {
          document.querySelector("#noKartu_txt").value = "";
        });
        await page.type("#noKartu_txt", row["NIK"]); // Tambahkan ini

        // Isi tanggal lahir
        await page.evaluate(() => {
          document.querySelector("#tglLhr_txt").value = "";
        });
        const tanggalLahir = excelSerialToDate(row["TGLLHR"]);
        if (!tanggalLahir) {
          console.error("Gagal memproses tanggal lahir untuk NIK:", row["NIK"]);
          continue; // Skip ke data berikutnya jika tanggal tidak valid
        }
        await page.type("#tglLhr_txt", tanggalLahir);

        // Loop tak terbatas sampai captcha terisi
        while (true) {
          writeLog("Menunggu 7 detik untuk pengisian captcha secara manual...");
          await delay(7000);

          // Cek apakah captcha sudah diisi
          const captchaValue = await page.evaluate(() => {
            const element = document.querySelector("#captchaCode_txt");
            return element ? element.value : "";
          });

          if (captchaValue && captchaValue.trim() !== "") {
            break;
          }
        }

        // Klik tombol "Cari Peserta"
        await page.click('[ng-click="pilihPeserta()"]');
        await delay(5000);

        // Tunggu dan klik tombol konfirmasi
        try {
          await page.waitForSelector(
            "body > div.bootbox.modal.fade.in > div > div > div.modal-footer > button.btn.btn-success",
            { timeout: 5000 }
          );
          await page.click(
            "body > div.bootbox.modal.fade.in > div > div > div.modal-footer > button.btn.btn-success"
          );
        } catch (error) {
          writeLog(
            `Warning: Tombol konfirmasi tidak ditemukan untuk NIK: ${row["NIK"]}`
          );
        }

        // Tunggu form untuk berat dan tinggi badan dengan retry
        await retryOperation(async () => {
          await page.waitForSelector("#beratBadan_txt", { timeout: 30000 });
        });

        // Isi Berat Badan
        await page.type("#beratBadan_txt", row["berat Badan"].toString());

        // Isi Tinggi Badan
        await page.type("#tinggiBadan_txt", row["tinggi Badan"].toString());

        // Isi Nama jika ada
        if (row["Nama"]) {
          await page.type("#nama_txt", row["NAMA"]);
        }

        // Isi Alamat jika ada
        if (row["Alamat"]) {
          await page.type("#alamatKlg_txt", row["ALAMAT"]);
        }

        await delay(2000);

        writeLog(`BERHASIL: Data NIK: ${row["NIK"]}, Nama: ${row["Nama"]}`);
      } catch (innerError) {
        writeLog(
          `GAGAL: Data NIK: ${row["NIK"]}, Error: ${innerError.message}`
        );
      }
    }
  } catch (error) {
    writeLog(`Terjadi error fatal: ${error.message}`);
  } finally {
    await browser.close();
  }
})();
