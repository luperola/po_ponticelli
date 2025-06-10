
const express = require('express');
const multer = require('multer');
const fs = require('fs');
const pdf = require('pdf-parse');
const ExcelJS = require('exceljs');
const app = express();
const upload = multer({ dest: 'uploads/' });

app.use(express.static('public'));

app.post('/upload', upload.single('pdf'), async (req, res) => {
    const dataBuffer = fs.readFileSync(req.file.path);
    const data = await pdf(dataBuffer);

    const lines = data.text.split(/\r?\n/).map(l => l.trim()).filter(l => l);
    const rows = [];
    let i = 0;

    while (i < lines.length) {
        const line = lines[i];
        const posMatch = line.match(/^(\d{1,2}\.\d)Art\.$/);
        if (posMatch) {
            const row = { "N. riga": posMatch[1] };
            const desc = [];

            i++;
            if (i < lines.length) desc.push(lines[i++]);
            if (i < lines.length) desc.push(lines[i++]);
            row["Cod. articolo/Descrizione"] = desc.join(" ");

            // Cerca una riga NRxxxx,yyy,zzz
            while (i < lines.length && !lines[i].startsWith("NR")) i++;
            if (i < lines.length && lines[i].startsWith("NR")) {
                const parts = lines[i].replace("NR", "").split(",");
                if (parts.length === 3) {
                    const digitsBefore = parts[0];
                    const digitsAfter = parts[1];
                    const final = parts[2];

                    if (digitsBefore.length === 4) {
                        row["Q.ty"] = digitsBefore.slice(0, 2);
                        row["P.U."] = digitsBefore.slice(2) + "," + digitsAfter.slice(0, 3);
                        row["Importo"] = digitsAfter.slice(3) + "," + final;
                    } else if (digitsBefore.length === 3) {
                        if (digitsBefore[1] === '0') {
                            row["Q.ty"] = digitsBefore.slice(0, 2);
                            row["P.U."] = digitsBefore[2] + "," + digitsAfter.slice(0, 3);
                            row["Importo"] = digitsAfter.slice(3) + "," + final;
                        } else {
                            row["Q.ty"] = digitsBefore[0];
                            row["P.U."] = digitsBefore.slice(1) + "," + digitsAfter.slice(0, 3);
                            row["Importo"] = digitsAfter.slice(3) + "," + final;
                        }
                    }
                }
                i++;
            }

            // Cerca data consegna
            if (i < lines.length && /\d{2}-[A-Z]{3}-\d{2}/.test(lines[i])) {
                row["Data consegna"] = lines[i];
                i++;
            }

            rows.push(row);
        } else {
            i++;
        }
    }

    if (!fs.existsSync("./public")) fs.mkdirSync("./public");

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Estratti");

    worksheet.columns = [
        { header: "N. riga", key: "N. riga" },
        { header: "Cod. articolo/Descrizione", key: "Cod. articolo/Descrizione" },
        { header: "Q.ty", key: "Q.ty" },
        { header: "P.U.", key: "P.U." },
        { header: "Importo", key: "Importo" },
        { header: "Data consegna", key: "Data consegna" }
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
