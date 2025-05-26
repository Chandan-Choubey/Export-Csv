const express = require("express");
const multer = require("multer");
const fs = require("fs");
const ExcelJS = require("exceljs");
const archiver = require("archiver");
const path = require("path");
const cors = require("cors");

const app = express();
const UPLOAD_DIR = path.join(__dirname, "uploads");

app.use(cors());
app.use(express.json());

const upload = multer({
  dest: UPLOAD_DIR,
  fileFilter: (req, file, cb) => {
    const allowedTypes = ["image/jpeg", "image/png", "image/gif"];
    if (allowedTypes.includes(file.mimetype)) {
      cb(null, true);
    } else {
      cb(new Error("Only JPG, PNG, and GIF files are supported."));
    }
  },
});

if (!fs.existsSync(UPLOAD_DIR)) {
  fs.mkdirSync(UPLOAD_DIR);
}

const deleteAllUploads = () => {
  fs.readdirSync(UPLOAD_DIR).forEach((file) => {
    fs.unlinkSync(path.join(UPLOAD_DIR, file));
  });
};

app.get("/", (req, res) => {
  res.send("Server is running.");
});

app.post("/convert", upload.single("file"), async (req, res) => {
  try {
    const input = req.body;
    const workbook = new ExcelJS.Workbook();
    const csvFiles = [];
    // console.log(input);
    // console.log(Object.entries(input));
    for (const [sheetName, rawSheet] of Object.entries(input)) {
      let sheetData, style, config;
      try {
        const parsed = JSON.parse(rawSheet);
        sheetData = parsed.data;
        style = parsed.style || {};
        config = parsed.config || {};
      } catch {
        continue;
      }

      if (!Array.isArray(sheetData)) continue;

      const sheet = workbook.addWorksheet(sheetName);

      const headerRow = sheet.addRow(sheetData[0]);
      headerRow.eachCell((cell) => {
        cell.font = {
          bold: true,
          color: { argb: style.fontColor || "FFFFFFFF" },
        };
        cell.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: style.backgroundColor || "FF1F4E78" },
        };
      });

      for (let i = 1; i < sheetData.length; i++) {
        const row = sheet.addRow();
        const rowData = sheetData[i];

        rowData.forEach((cellData, colIndex) => {
          const cell = row.getCell(colIndex + 1);

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
          } else if (
            typeof cellData === "string" &&
            cellData.startsWith("http")
          ) {
            cell.value = { text: cellData, hyperlink: cellData };
            cell.font = { color: { argb: "FF0000FF" }, underline: true };
          } else if (typeof cellData === "number") {
            cell.value = cellData;
            cell.numFmt = "#,##0";
            if (colIndex === 2 && cellData < 800) {
              cell.fill = {
                type: "pattern",
                pattern: "solid",
                fgColor: { argb: "FFFF0000" },
              };
            }
          } else {
            cell.value = cellData;
          }
        });
      }

      sheet.columns.forEach((col) => {
        let maxLen = 10;
        col.eachCell({ includeEmpty: true }, (cell) => {
          const len = cell.value ? cell.value.toString().length : 0;
          if (len > maxLen) maxLen = len;
        });
        col.width = maxLen + 2;
      });

      if (config.sumColumn && config.startRow && config.operation) {
        const supportedOps = ["SUM", "AVERAGE", "MIN", "MAX", "COUNT"];
        if (!supportedOps.includes(config.operation)) {
          return res.status(400).json({
            error: `Invalid or missing operation. Supported operations are: ${supportedOps.join(
              ", "
            )}.`,
          });
        }
        const endRow = sheet.lastRow.number;

        const columnLetter = (index) => String.fromCharCode(64 + index);

        const labelColLetter = columnLetter(
          config.labelColumn || config.sumColumn
        );
        const formulaColLetter = columnLetter(
          config.formulaColumn || config.sumColumn + 1
        );
        const targetColumnLetter = columnLetter(config.sumColumn);

        const labelCell = sheet.getCell(`${labelColLetter}${endRow + 1}`);
        const formulaCell = sheet.getCell(`${formulaColLetter}${endRow + 1}`);

        labelCell.value = config.label || config.operation + " Result";
        labelCell.font = { bold: true };

        formulaCell.value = {
          formula: `${config.operation}(${targetColumnLetter}${config.startRow}:${targetColumnLetter}${endRow})`,
        };
      }

      if (req.file) {
        try {
          const imageName = path.parse(req.file.originalname).name;
          const sheetMatchesFile = imageName === sheetName;
          const ext = path
            .extname(req.file.originalname)
            .toLowerCase()
            .replace(".", "");
          const supported = ["jpeg", "jpg", "png", "gif"];

          if (sheetMatchesFile && supported.includes(ext)) {
            const imageId = workbook.addImage({
              filename: req.file.path,
              extension: ext === "jpg" ? "jpeg" : ext,
            });

            const imageWidthPx = 500;
            const imageHeightPx = 500;
            const targetCol = 1;
            const targetRow = sheet.lastRow.number + 1;

            sheet.getColumn(targetCol).width = imageWidthPx / 7.0017;
            sheet.getRow(targetRow).height = imageHeightPx / 1.33;

            sheet.addImage(imageId, {
              tl: { col: targetCol - 1, row: targetRow - 1 },
              ext: { width: imageWidthPx, height: imageHeightPx },
              editAs: "oneCell",
            });
          }
        } catch (err) {
          console.error("Image embedding failed:", err.message);
        }
      }

      const csvPath = `${sheetName}.csv`;
      const csvContent = [sheetData[0].join(",")]
        .concat(sheetData.slice(1).map((row) => row.join(",")))
        .join("\n");
      fs.writeFileSync(csvPath, csvContent);
      csvFiles.push(csvPath);
    }

    const xlsxPath = "output.xlsx";
    await workbook.xlsx.writeFile(xlsxPath);

    const zipPath = "output.zip";
    const archive = archiver("zip", { zlib: { level: 9 } });
    const output = fs.createWriteStream(zipPath);

    output.on("close", () => {
      res.download(zipPath, "output.zip", () => {
        [xlsxPath, zipPath, ...csvFiles].forEach((file) => fs.unlinkSync(file));
        deleteAllUploads();
      });
    });

    archive.pipe(output);
    archive.file(xlsxPath, { name: "output.xlsx" });
    csvFiles.forEach((f) => archive.file(f, { name: f }));
    archive.finalize();
  } catch (err) {
    console.error("Error during conversion:", err);
    res.status(500).json({ error: err.message });
    deleteAllUploads();
  }
});

app.listen(3000, () => {
  console.log("Server running at http://localhost:3000");
});
