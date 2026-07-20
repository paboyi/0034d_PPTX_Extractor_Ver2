// Minimal web server that serves the upload UI and exposes a single
// endpoint: POST /api/extract  (multipart/form-data, field name "file").


import express, { Request, Response } from "express";
import multer from "multer";
import * as path from "path";
import { extractPptxFromBuffer, saveImagesToFolder} from "./extractor";
import cors from "cors";
import * as os from "os";

const PPTX_MIME =
  "application/vnd.openxmlformats-officedocument.presentationml.presentation";

// Keep uploads in memory; we never need them on disk.
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 100 * 1024 * 1024 }, // 100 MB cap 
  //note Lambda itself will reject anything over ~6MB before this even matters

  fileFilter: (_req, file, cb) => {
    const hasPptxExt = file.originalname.toLowerCase().endsWith(".pptx");
      // Some browsers send octet-stream; allow it only alongside a .pptx name.
  const mimeOk =
      file.mimetype === PPTX_MIME ||
      file.mimetype === "application/octet-stream";
    if (hasPptxExt && mimeOk) {
      cb(null, true);
    } else {
      cb(new Error("Only .pptx files are accepted."));
    }
  },
});

const app = express();

// Allow your Vercel frontend (a different domain) to call this API.
app.use(cors());

// Serve the frontend from /public.
app.use(express.static(path.join(__dirname, "..", "public")));

app.post("/api/extract", (req: Request, res: Response) => {
  upload.single("file")(req, res, async (err: unknown) => {
        // Multer errors (wrong type, too large, etc.)
    if (err instanceof multer.MulterError) {
      const msg =
        err.code === "LIMIT_FILE_SIZE"
          ? "File is too large (max 100 MB)."
          : err.message;
      return res.status(400).json({ error: msg });
    }
    if (err instanceof Error) {
      return res.status(400).json({ error: err.message });
    }

    if (!req.file) {
      return res
        .status(400)
        .json({ error: "No file received. Please choose a .pptx file." });
    }

    try {
      const result = await extractPptxFromBuffer(req.file.buffer);

        // also write the real image files to a folder (one per upload so different files don't overwrite each other)
      /* AWS Lambda - Debug 
      
      Debug copy of extracted images — write to /tmp, the ONLY writable
      // path inside a Lambda execution environment. This disappears whenever
      // the Lambda container is recycled, so don't rely on it for anything real. */
      const base = (req.file.originalname || "upload").replace(/\.pptx$/i, "");
      const outDir = path.join(os.tmpdir(), "output_images", `${base}-${Date.now()}`);
      const saved = saveImagesToFolder(result, outDir);
      console.log(`Saved ${saved.length} images to ${outDir}`);

      return res.json(result);
    } catch (e) {
      return res.status(422).json({
        error:
          "That file could not be read as a PowerPoint (.pptx). It may be corrupt or not a real .pptx.",
        detail: e instanceof Error ? e.message : String(e),
      });
    }
  });
});

export { app };

// Only bind a real port when running locally on your machine.
// AWS sets AWS_LAMBDA_FUNCTION_NAME automatically inside Lambda
// we use that to detect Lambda and skip .listen(), since Lambda
// never calls it anyway (it calls the handler in lambda.ts instead).
if (!process.env.AWS_LAMBDA_FUNCTION_NAME) {
  const PORT = Number(process.env.PORT) || 3000;
  app.listen(PORT, () => {
    console.log(`PPTX extractor running at http://localhost:${PORT}`);
  });
}