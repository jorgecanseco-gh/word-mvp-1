import express from "express";
import multer from "multer";
import mammoth from "mammoth";
import { Document, Packer, Paragraph } from "docx";

const app = express();
const upload = multer(); // memory storage

// --------------------------------------------------
// 1️⃣ HEALTH CHECK (optional but recommended)
// --------------------------------------------------
app.get("/", (req, res) => {
  res.send("Word MVP API is running");
});

// --------------------------------------------------
// 2️⃣ UPLOAD DOCX → EXTRACT TEXT (existing, KEEP THIS)
// --------------------------------------------------
app.post("/upload", upload.single("file"), async (req, res) => {
  if (!req.file) {
    return res.status(400).json({
      ok: false,
      error: "No file received",
    });
  }

  try {
      const result = await mammoth.convertToHtml({
    buffer: req.file.buffer,
  });


     res.json({
    ok: true,
    html: result.value,
  });

  } catch (error) {
    console.error("Mammoth error:", error);
    res.status(500).json({
      ok: false,
      error: "Failed to extract text",
    });
  }
});

// --------------------------------------------------
// 3️⃣ TEXT → DOCX (NEW ENDPOINT)
// --------------------------------------------------
app.post(
  "/generate-docx",
  express.json({ limit: "5mb" }),
  async (req, res) => {
    const { text } = req.body;

    if (!text) {
      return res.status(400).json({
        ok: false,
        error: "No text provided",
      });
    }

    try {
      const doc = new Document({
        sections: [
          {
            children: text
              .split("\n")
              .map((line) => new Paragraph(line)),
          },
        ],
      });

      const buffer = await Packer.toBuffer(doc);

      res.setHeader(
        "Content-Type",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
      );
      res.setHeader(
        "Content-Disposition",
        "attachment; filename=edited.docx"
      );

      res.send(buffer);
    } catch (error) {
      console.error("DOCX generation error:", error);
      res.status(500).json({
        ok: false,
        error: "Failed to generate DOCX",
      });
    }
  }
);

// --------------------------------------------------
// 4️⃣ START SERVER (Railway compatible)
// --------------------------------------------------
const PORT = process.env.PORT || 3000;

app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});





