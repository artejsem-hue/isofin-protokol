const express = require("express");
const fs = require("fs");
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");

const app = express();
app.use(express.json());

app.post("/generate", (req, res) => {
  const data = req.body;

  const content = fs.readFileSync("template.docx", "binary");
  const zip = new PizZip(content);
  const doc = new Docxtemplater(zip);

  doc.setData(data);

  try {
    doc.render();
  } catch (error) {
    return res.status(500).json({ error: "Chyba generování Wordu" });
  }

  const buf = doc.getZip().generate({ type: "nodebuffer" });
  fs.writeFileSync("protokol.docx", buf);

  res.download("protokol.docx");
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log("Server běží na portu", PORT);
});
