import express from "express";
import multer from "multer";
import mammoth from "mammoth";


const app = express();
const upload = multer(); // memory storage

app.post("/upload", upload.single("file"), async (req, res) => {
  if (!req.file) {
    return res.status(400).json({
      ok: false,
      error: "No file received"
    });
  }

  try {
    const result = await mammoth.extractRawText({
      buffer: req.file.buffer
    });

    res.json({
      ok: true,
      text: result.value
    });
  } catch (error) {
    console.error("Mammoth error:", error);

    res.status(500).json({
      ok: false,
      error: "Failed to extract text"
    });
  }
});

