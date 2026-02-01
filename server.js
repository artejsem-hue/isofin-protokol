const express = require("express");
const fs = require("fs");
const path = require("path");
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");

const app = express();
app.use(express.json());

// CESTA K WORD ŠABLONĚ (JE V KOŘENI REPA)
const templatePath = path.join(__dirname, "template.docx");

app.post("/generate", (req, res) => {
  try {
    // 1️⃣ kontrola existence šablony
    if (!fs.existsSync(templatePath)) {
      return res.status(500).send("Template not found");
    }

    // 2️⃣ načtení šablony
    const content = fs.readFileSync(templatePath);
    const zip = new PizZip(content);

    // 3️⃣ inicializace docxtemplateru
    // 👉 DŮLEŽITÉ: vlastní delimitery [[ ]]
    const doc = new Docxtemplater(zip, {
      paragraphLoop: true,
      linebreaks: true,
      delimiters: {
        start: "[[",
        end: "]]"
      }
    });

    // 4️⃣ naplnění dat (musí odpovídat [[PLACEHOLDER]] ve Wordu)
    doc.setData({
      CERT_NUMBER: req.body.cert_number || "",
      REPORT_NUMBER: req.body.report_number || "",
      OP_CODE: req.body.op_code || "",
      OBJECT_NAME: req.body.object_name || "",
      CUSTOMER_NAME: req.body.customer_name || "",
      CUSTOMER_ICO: req.body.customer_ico || "",
      CUSTOMER_REPRESENTATIVE: req.body.customer_representative || "",
      ISSUE_DATE: req.body.issue_date || "",
      NEXT_INSPECTION_DATE: req.body.next_inspection_date || ""
    });

    // 5️⃣ render dokumentu
    doc.render();

    // 6️⃣ generování Word bufferu
    const buffer = doc.getZip().generate({
      type: "nodebuffer"
    });

    // 7️⃣ HTTP hlavičky pro Word
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    );
    res.setHeader(
      "Content-Disposition",
      'attachment; filename="protokol.docx"'
    );

    // 8️⃣ odeslání souboru
    res.send(buffer);

  } catch (err) {
    console.error("GENERATE ERROR:", err);
    res.status(500).send("GENERATE ERROR: " + err.message);
  }
});

// PORT PRO RAILWAY
const PORT = process.env.PORT || 8080;
app.listen(PORT, () => {
  console.log("Server běží na portu", PORT);
});
