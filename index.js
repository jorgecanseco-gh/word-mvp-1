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
// 3️⃣ HTML → DOCX (with color support)
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

      function extractColor(style = "") {
        const match = style.match(/color:\s*(#[0-9a-fA-F]{6})/);
        return match ? match[1].replace("#", "") : undefined;
      }

      function buildTextRuns(nodes) {
        const runs = [];

        for (const node of nodes) {
          if (node.type === "text") {
            runs.push(new TextRun(node.data));
          }

          if (node.name === "strong") {
            runs.push(
              new TextRun({
                text: node.children?.[0]?.data || "",
                bold: true,
              })
            );
          }

          if (node.name === "em") {
            runs.push(
              new TextRun({
                text: node.children?.[0]?.data || "",
                italics: true,
              })
            );
          }

          if (node.name === "span") {
            const color = extractColor(node.attribs?.style || "");
            runs.push(
              new TextRun({
                text: node.children?.[0]?.data || "",
                color,
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











