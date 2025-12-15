import express from "express";
import multer from "multer";
import mammoth from "mammoth";
import { parseDocument } from "htmlparser2";
import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  HeadingLevel,
} from "docx";

const app = express();
const upload = multer();

// --------------------------------------------------
// 1️⃣ HEALTH CHECK
// --------------------------------------------------
app.get("/", (req, res) => {
  res.send("Word MVP API is running");
});

// --------------------------------------------------
// 2️⃣ UPLOAD DOCX → HTML
// --------------------------------------------------
app.post("/upload", upload.single("file"), async (req, res) => {
  if (!req.file) {
    return res.status(400).json({ ok: false, error: "No file received" });
  }

  try {
    const result = await mammoth.convertToHtml({
      buffer: req.file.buffer,
    });

    res.json({ ok: true, html: result.value });
  } catch (error) {
    console.error(error);
    res.status(500).json({ ok: false, error: "Failed to extract document" });
  }
});

// --------------------------------------------------
// 3️⃣ HTML → DOCX (COLOR FIXED)
// --------------------------------------------------
app.post(
  "/generate-docx",
  express.json({ limit: "5mb" }),
  async (req, res) => {
    const { html } = req.body;

    if (!html) {
      return res.status(400).json({ ok: false, error: "No HTML provided" });
    }

    try {
      const dom = parseDocument(html);
      const paragraphs = [];

      // rgb(5, 99, 193) → 0563C1
      function rgbToHex(rgb) {
        const match = rgb.match(
          /rgb\(\s*(\d+),\s*(\d+),\s*(\d+)\s*\)/
        );
        if (!match) return undefined;

        return match
          .slice(1)
          .map((n) => Number(n).toString(16).padStart(2, "0"))
          .join("")
          .toUpperCase();
      }

      function extractColor(style = "") {
        if (style.includes("rgb")) {
          return rgbToHex(style);
        }

        const hex = style.match(/color:\s*(#[0-9a-fA-F]{6})/);
        return hex ? hex[1].replace("#", "").toUpperCase() : undefined;
      }

      function buildTextRuns(nodes, inheritedStyle = {}) {
        const runs = [];

        for (const node of nodes) {
          if (node.type === "text") {
            runs.push(
              new TextRun({
                text: node.data,
                ...inheritedStyle,
              })
            );
          }

          if (node.name === "strong") {
            runs.push(
              ...buildTextRuns(node.children || [], {
                ...inheritedStyle,
                bold: true,
              })
            );
          }

          if (node.name === "em") {
            runs.push(
              ...buildTextRuns(node.children || [], {
                ...inheritedStyle,
                italics: true,
              })
            );
          }

          if (node.name === "span") {
            const color = extractColor(node.attribs?.style || "");
            runs.push(
              ...buildTextRuns(node.children || [], {
                ...inheritedStyle,
                ...(color ? { color } : {}),
              })
            );
          }
        }

        return runs;
      }

      function walk(nodes) {
        for (const node of nodes) {
          // HEADINGS
          if (["h1", "h2", "h3"].includes(node.name)) {
            const levelMap = {
              h1: HeadingLevel.HEADING_1,
              h2: HeadingLevel.HEADING_2,
              h3: HeadingLevel.HEADING_3,
            };

            paragraphs.push(
              new Paragraph({
                children: buildTextRuns(node.children || []),
                heading: levelMap[node.name],
              })
            );
          }

          // PARAGRAPHS
          if (node.name === "p") {
            paragraphs.push(
              new Paragraph({
                children: buildTextRuns(node.children || []),
              })
            );
          }

          // BULLET LISTS
          if (node.name === "ul") {
            node.children?.forEach((li) => {
              if (li.name === "li") {
                paragraphs.push(
                  new Paragraph({
                    children: buildTextRuns(li.children || []),
                    bullet: { level: 0 },
                  })
                );
              }
            });
          }

          if (node.children) {
            walk(node.children);
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
      console.error(error);
      res.status(500).json({
        ok: false,
        error: "Failed to generate DOCX from HTML",
      });
    }
  }
);

// --------------------------------------------------
// 4️⃣ START SERVER
// --------------------------------------------------
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});













