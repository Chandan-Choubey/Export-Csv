const express = require("express");
const multer = require("multer");
const fs = require("fs");
const ExcelJS = require("exceljs");
const archiver = require("archiver");
const path = require("path");
const cors = require("cors");

const app = express();
const upload = multer({ dest: "uploads/" });
app.use(cors());
app.use(express.json());

app.get("/", (req, res) => {
  res.send("Hello");
});

app.post("/convert", upload.single("file"), async (req, res) => {
  const Uploadpath = "./uploads";

  if (!fs.existsSync(Uploadpath)) {
    fs.mkdirSync(Uploadpath);
  }
  const input = req.body;
  console.log(input);
  const workbook = new ExcelJS.Workbook();
  const csvFiles = [];

  for (const [sheetName, rawSheet] of Object.entries(input)) {
    let sheetData;
    let style;

    try {
      const parsed = JSON.parse(rawSheet);
      sheetData = parsed.data;
      style = parsed.style || {};
    } catch (err) {
      continue;
    }

    if (!Array.isArray(sheetData)) continue;

    const sheet = workbook.addWorksheet(sheetName);

    // Header row with styling
    const headerRow = sheet.addRow(sheetData[0]);
    const headerFontColor = style.fontColor || "FFFFFFFF";
    const headerBgColor = style.backgroundColor || "FF4472C4";

    headerRow.eachCell((cell) => {
      cell.font = { bold: true, color: { argb: headerFontColor } };
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: headerBgColor },
      };
    });

    // Data rows
    for (let i = 1; i < sheetData.length; i++) {
      const row = sheet.addRow();
      const rowData = sheetData[i];

      rowData.forEach((cellData, colIndex) => {
        const cell = row.getCell(colIndex + 1);

        // Handle rich text objects like { text: "Everest", bold: true }
        if (
          typeof cellData === "object" &&
          !Array.isArray(cellData) &&
          cellData.text
        ) {
          cell.value = {
            richText: [
              {
                text: cellData.text,
                font: cellData.bold ? { bold: true } : undefined,
              },
            ],
          };
        }

        // Handle hyperlinks
        else if (typeof cellData === "string" && cellData.startsWith("http")) {
          cell.value = { text: cellData, hyperlink: cellData };
          cell.font = { color: { argb: "FF0000FF" }, underline: true };
        }

        // Handle numbers
        else if (typeof cellData === "number") {
          cell.value = cellData;
          cell.numFmt = "#,##0";

          // Example: Apply conditional red background if < 800 in "Visitors" column
          if (colIndex === 2 && cellData < 800) {
            // colIndex 2 = 3rd column
            cell.fill = {
              type: "pattern",
              pattern: "solid",
              fgColor: { argb: "FFFF0000" },
            };
          }
        }

        // Default: plain text
        else {
          cell.value = cellData;
        }
      });
    }

    // Auto-fit columns
    sheet.columns.forEach((column) => {
      let maxLength = 10;
      column.eachCell({ includeEmpty: true }, (cell) => {
        const len = cell.value ? cell.value.toString().length : 0;
        if (len > maxLength) maxLength = len;
      });
      column.width = maxLength + 2;
    });

    // Add formula row for "Total Visitors" for Sheet1
    if (Array.isArray(sheetData) && sheetData[0].length >= 3) {
      const startRow = 2;
      const endRow = sheet.lastRow.number;
      const labelCell = sheet.getCell(`C${endRow + 1}`);
      const formulaCell = sheet.getCell(`D${endRow + 1}`);

      labelCell.value = "Total Visitors";
      labelCell.font = { bold: true };
      formulaCell.value = { formula: `SUM(C${startRow}:C${endRow})` };
    }

    // Check if an uploaded image matches the sheet name
    const uploadedImagePath = req.file ? req.file.path : null;
    const imageExtension = req.file
      ? path.extname(req.file.originalname).substring(1)
      : null;

    if (
      uploadedImagePath &&
      path.basename(
        req.file.originalname,
        path.extname(req.file.originalname)
      ) === sheetName &&
      ["png", "jpg", "jpeg"].includes(imageExtension)
    ) {
      const imageId = workbook.addImage({
        filename: uploadedImagePath,
        extension: imageExtension,
      });

      const imageWidthPx = 500; // Set your desired image width
      const imageHeightPx = 500; // Set your desired image height

      const targetCol = 1; // Excel column A = 1
      const targetRow = sheet.lastRow.number + 1;

      // Set the column width (1 unit ≈ 7.5px)
      sheet.getColumn(targetCol).width = imageWidthPx / 7.5;

      // Set the row height (1 unit ≈ 0.75pt ≈ 1.33px, ExcelJS uses points)
      sheet.getRow(targetRow).height = imageHeightPx / 1.33;

      // Add image to that exact cell
      sheet.addImage(imageId, {
        tl: { col: targetCol - 1, row: targetRow - 1 },
        ext: { width: imageWidthPx, height: imageHeightPx },
        editAs: "oneCell",
      });
    }

    // Save CSV
    const csvPath = `${sheetName}.csv`;
    const csvContent = [sheetData[0].join(",")]
      .concat(sheetData.slice(1).map((row) => row.join(",")))
      .join("\n");
    fs.writeFileSync(csvPath, csvContent);
    csvFiles.push(csvPath);
  }

  // Save Excel file
  const xlsxPath = "output.xlsx";
  await workbook.xlsx.writeFile(xlsxPath);

  // Create ZIP
  const zipPath = "output.zip";
  const archive = archiver("zip", { zlib: { level: 9 } });
  const output = fs.createWriteStream(zipPath);

  output.on("close", () => {
    res.download(zipPath, "output.zip", () => {
      fs.unlinkSync(xlsxPath);
      fs.unlinkSync(zipPath);
      csvFiles.forEach((f) => fs.unlinkSync(f));
    });
  });

  archive.pipe(output);
  archive.file(xlsxPath, { name: "output.xlsx" });
  csvFiles.forEach((f) => archive.file(f, { name: f }));
  archive.finalize();
});

app.listen(3000, () => {
  console.log("Server running at http://localhost:3000");
});
