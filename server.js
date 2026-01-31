const express = require("express");
const fs = require("fs");
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const path = require("path");

const templatePath = path.join(__dirname, "template.docx");


const app = express();
app.use(express.json());

app.post("/generate", async (req, res) => {
  try {
    console.log("STEP 1: handler start");

    const templatePath = path.join(__dirname, "template.docx");
    console.log("STEP 2: templatePath", templatePath);

    const content = fs.readFileSync(templatePath);
    console.log("STEP 3: template loaded, size", content.length);

    const zip = new PizZip(content);
    console.log("STEP 4: PizZip OK");

    const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });
    console.log("STEP 5: Docxtemplater OK");

    doc.setData({
      CERT_NUMBER: req.body.cert_number,
      REPORT_NUMBER: req.body.report_number,
      OP_CODE: req.body.op_code,
      OBJECT_NAME: req.body.object_name,
      CUSTOMER_NAME: req.body.customer_name,
      CUSTOMER_ICO: req.body.customer_ico,
      CUSTOMER_REPRESENTATIVE: req.body.customer_representative,
      ISSUE_DATE: req.body.issue_date,
      NEXT_INSPECTION_DATE: req.body.next_inspection_date
    });

    console.log("STEP 6: data set");

    doc.render();
    console.log("STEP 7: render OK");

    const buf = doc.getZip().generate({ type: "nodebuffer" });
    console.log("STEP 8: buffer generated, size", buf.length);

    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    );
    res.setHeader(
      "Content-Disposition",
      'attachment; filename="protokol.docx"'
    );

    res.send(buf);
    console.log("STEP 9: response sent");

  } catch (err) {
    console.error("🔥 GENERATE ERROR:", err);
    res.status(500).send("GENERATE ERROR: " + err.message);
  }
});


  console.log("CWD:", process.cwd());
console.log("DIRNAME:", __dirname);
console.log("Template path:", templatePath);
console.log("Exists:", fs.existsSync(templatePath));
if (fs.existsSync(templatePath)) {
  console.log("Template size:", fs.statSync(templatePath).size);
}
  const content = fs.readFileSync(templatePath, "binary");
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
