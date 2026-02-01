const express = require("express");
const fs = require("fs");
const path = require("path");
const cors = require("cors");
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");

const app = express();

/* ===== MIDDLEWARE ===== */
app.use(cors());               // ← nutné pro Netlify → Railway
app.use(express.json());       // ← čtení JSON z wizardu

/* ===== CESTA K WORD ŠABLONĚ ===== */
const templatePath = path.join(__dirname, "template.docx");

/* ===== API ENDPOINT (SEDÍ S FRONTENDEM) ===== */
app.post("/api/protocol", (req, res) => {
  try {
    // 1️⃣ kontrola šablony
    if (!fs.existsSync(templatePath)) {
      return res.status(500).send("Template not found");
    }

    // 2️⃣ načtení Word šablony
    const content = fs.readFileSync(templatePath, "binary");
    const zip = new PizZip(content);

    // 3️⃣ inicializace docxtemplateru + [[ ]]
    const doc = new Docxtemplater(zip, {
      paragraphLoop: true,
      linebreaks: true,
      delimiters: {
        start: "[[",
        end: "]]"
      }
    });

    // 4️⃣ DATA Z WIZARDU – 1:1 SHODA S PLACEHOLDERY
    doc.setData({
      CERT_NUMBER: req.body.CERT_NUMBER || "",
      REPORT_NUMBER: req.body.REPORT_NUMBER || "",
      OP_CODE: req.body.OP_CODE || "",
      OBJECT_NAME: req.body.OBJECT_NAME || "",
      CUSTOMER_NAME: req.body.CUSTOMER_NAME || "",
      CUSTOMER_ICO: req.body.CUSTOMER_ICO || "",
      CUSTOMER_REPRESENTATIVE: req.body.CUSTOMER_REPRESENTATIVE || "",
      ISSUE_DATE: req.body.ISSUE_DATE || "",
      NEXT_INSPECTION_DATE: req.body.NEXT_INSPECTION_DATE || ""
    });

    // 5️⃣ render dokumentu
    doc.render();

    // 6️⃣ generování Word souboru
    const buffer = doc.getZip().generate({
      type: "nodebuffer"
    });

    // 7️⃣ HTTP hlavičky
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    );
    res.setHeader(
      "Content-Disposition",
      'attachment; filename="protokol.docx"'
    );

    // 8️⃣ odeslání Wordu
    res.send(buffer);

  } catch (err) {
    console.error("GENERATE ERROR:", err);
    res.status(500).send("GENERATE ERROR: " + err.message);
  }
});

/* ===== PORT (RAILWAY) ===== */
const PORT = process.env.PORT || 8080;
app.listen(PORT, () => {
  console.log("Server běží na portu", PORT);
});
