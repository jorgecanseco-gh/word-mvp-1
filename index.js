import express from "express";
import multer from "multer";
import mammoth from "mammoth";
import { parseDocument } from "htmlparser2";
import { Document, Packer, Paragraph, TextRun } from "docx";

const app = express();
const upload = multer(); // memory storage

// --------------------------------------------------
// 1️⃣ HEALTH CHECK
// --------------------------------------------------
app.get("/", (req, res) => {
  res.send("Word MVP API is running");
});

// --------------------------------------------------
// 2️⃣ UPLOAD DOCX → EXTRACT HTML (preserve structure)
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
      error: "Failed to extract document",
    });
  }
});

// --------------------------------------------------
// 3️⃣ HTML → DOCX (preserve paragraphs, lists, bold)
// --------------------------------------------------
app.post(
  "/generate-docx",
  express.json({ limit: "5mb" }),
  async (req, res) => {
    const { html } = req.body;

    if (!html) {
      return res.status(400).json({
        ok: false,
        error: "No HTML provided",
      });
    }

    try {
      const dom = parseDocument(html);
      const paragraphs = [];

      function walk(nodes) {
  for (const node of nodes) {

    // ✅ HEADINGS (ADD THIS)
    if (node.name === "h1" || node.name === "h2" || node.name === "h3") {
      const level =
        node.name === "h1" ? 1 :
        node.name === "h2" ? 2 : 3;

      paragraphs.push(
        new Paragraph({
          text: node.children?.[0]?.data || "",
          heading: `Heading${level}`,
        })
      );
    }

    // Paragraphs (ALREADY EXISTING)
    if (node.name === "p") {
      const runs = [];

      node.children?.forEach((child) => {
        if (child.type === "text") {
          runs.push(new TextRun(child.data));
        }

        if (child.name === "strong") {
          runs.push(
            new TextRun({
              text: child.children?.[0]?.data || "",
              bold: true,
            })
          );
        }

        if (child.name === "em") {
          runs.push(
            new TextRun({
              text: child.children?.[0]?.data || "",
              italics: true,
            })
          );
        }
      });

      paragraphs.push(new Paragraph({ children: runs }));
    }

    // Bullet lists (ALREADY EXISTING)
    if (node.name === "ul") {
      node.children?.forEach((li) => {
        if (li.name === "li") {
          paragraphs.push(
            new Paragraph({
              text: li.children?.[0]?.data || "",
              bullet: { level: 0 },
            })
          );
        }
      });
    }
  }
}

      walk(dom.children);

      const document = new Document({
        sections: [{ children: paragraphs }],
      });

      const buffer = await Packer.toBuffer(document);

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
        error: "Failed to generate DOCX from HTML",
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








