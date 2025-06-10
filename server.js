const express = require("express");
const multer = require("multer");
const fs = require("fs");
const pdf = require("pdf-parse");
const ExcelJS = require("exceljs");
const app = express();
const upload = multer({ dest: "uploads/" });

app.use(express.static("public"));

app.post("/upload", upload.single("pdf"), async (req, res) => {
  const dataBuffer = fs.readFileSync(req.file.path);
  const data = await pdf(dataBuffer);

  const lines = data.text
    .split(/\r?\n/)
    .map((l) => l.trim())
    .filter((l) => l);
  const itemRegex =
    /(\d+0)\s+(RED|Equal|Tee|Elbow|Purge|Red)([\s\S]*?)(?=\n\d+0\s+(?:RED|Equal|Tee|Elbow|Purge|Red)|$)/gi;
  const rows = [];
  const text = lines.join("\n");
  let match;
  while ((match = itemRegex.exec(text)) !== null) {
    rows.push({
      Item: match[1],
      "Short text": (match[2] + match[3]).replace(/\s+/g, " ").trim(),
    });
  }

  if (!fs.existsSync("./public")) fs.mkdirSync("./public");

  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("Estratti");

  worksheet.columns = [
    { header: "Item", key: "Item" },
    { header: "Short text", key: "Short text" },
  ];

  worksheet.addRows(rows);

  const outputPath = "./public/output.xlsx";
  await workbook.xlsx.writeFile(outputPath);
  console.log("✅ File Excel creato con", rows.length, "righe.");
  res.download(outputPath);
});

app.listen(3000, () => {
  console.log("✅ Server avviato su http://localhost:3000");
});
