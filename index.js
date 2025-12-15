import express from "express";
import multer from "multer";

const app = express();
const upload = multer(); // memory storage

app.post("/upload", upload.single("file"), (req, res) => {
  if (!req.file) {
    return res.status(400).json({
      ok: false,
      error: "No file received"
    });
  }

  console.log("File received:", {
    originalName: req.file.originalname,
    mimeType: req.file.mimetype,
    size: req.file.size
  });

  res.json({
    ok: true,
    message: "File received",
    fileName: req.file.originalname,
    fileSize: req.file.size
  });
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});
