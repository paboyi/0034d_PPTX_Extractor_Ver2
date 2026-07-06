// Refactored from src/extrtact (1st ver).ts to be importable and to work from an in-memory buffer (so the web server never has to touch the disk).
// Difference: images are returned inline as base64 data URIs instead of being written to an output folder, so a single JSON response is self-contained.
import * as path from "path";
import * as unzipper from "unzipper";
import { XMLParser } from "fast-xml-parser";
import * as fs from "fs";




// PPTX measures in EMUs. Convert to pixels (96 DPI => 9525 EMU per px).
function emuToPx(emu: number): number {
  return Math.round(emu / 9525);
}

// Turn an image buffer into a data URI so it can travel inside the JSON.
function toDataUri(buffer: Buffer, filename: string): string {
  const ext = path.extname(filename).slice(1).toLowerCase() || "png";
  const mime =
    ext === "jpg" || ext === "jpeg"
      ? "image/jpeg"
      : ext === "svg"
      ? "image/svg+xml"
      : `image/${ext}`;
  return `data:${mime};base64,${buffer.toString("base64")}`;
}

// Extract text + geometry from a shape or group element. If there.
function extractTextFromElement(element: any): {
  texts: string[];
  position: { x: number; y: number };
  size: { width: number; height: number };
} {
  const texts: string[] = [];
  let position = { x: 0, y: 0 };
  let size = { width: 0, height: 0 };

  const xfrm = element["p:spPr"]?.["a:xfrm"];
  if (xfrm) {
    position = {
      x: emuToPx(Number(xfrm?.["a:off"]?.x || 0)),
      y: emuToPx(Number(xfrm?.["a:off"]?.y || 0)),
    };
    size = {
      width: emuToPx(Number(xfrm?.["a:ext"]?.cx || 0)),
      height: emuToPx(Number(xfrm?.["a:ext"]?.cy || 0)),
    };
  }

  if (element["p:txBody"]?.["a:p"]) {
    const paras = element["p:txBody"]["a:p"];
    const paraArr = Array.isArray(paras) ? paras : [paras];

    for (const p of paraArr) {
      if (!p) continue;
      const runs = p["a:r"];
      if (runs) {
        const runArr = Array.isArray(runs) ? runs : [runs];
        for (const r of runArr) {
          if (r?.["a:t"]) texts.push(r["a:t"]);
        }
      }
    }
  }

  return { texts, position, size };
}

// Custom-geometry shapes that carry an image fill (blipFill).
async function processCustomShapes(
  shapes: any[],
  elements: any[],
  slideIndex: number,
  directory: any,
  relArr: any[]
) {
  for (const shape of shapes) {
    // Check if this shape has an image fill (blipFill) //e.g. slide 3 images
    const blipFill = shape["p:spPr"]?.["a:blipFill"];
    if (!blipFill) continue;

    const xfrm = shape["p:spPr"]?.["a:xfrm"];
    const position = {
      x: emuToPx(Number(xfrm?.["a:off"]?.x || 0)),
      y: emuToPx(Number(xfrm?.["a:off"]?.y || 0)),
    };
    const size = {
      width: emuToPx(Number(xfrm?.["a:ext"]?.cx || 0)),
      height: emuToPx(Number(xfrm?.["a:ext"]?.cy || 0)),
    };


    // Get the embed ID
    const embedId = blipFill["a:blip"]?.["r:embed"];
    if (!embedId) continue;

    const rel = relArr.find((r: any) => r && r.Id === embedId);
    if (!rel) continue;

    const target = rel.Target;
    const imagePath = target.startsWith("../")
      ? `ppt/${target.substring(3)}`
      : `ppt/${target}`;

    const imageFile = directory.files.find((f: any) => f.path === imagePath);
    if (!imageFile) continue;

    try {
      const imageBuffer = await imageFile.buffer();
      const imageName = `slide${slideIndex + 1}_${path.basename(imagePath)}`;
      elements.push({
        type: "image",
        name: imageName,
        src: toDataUri(imageBuffer, imagePath),
        position,
        size,
        imageType: "customShape", // Marking as a custom shape for refernce
        hasCustomGeometry: true,
      });
    } catch (error) {
      console.warn(`Failed to process custom shape image ${imagePath}:`, error);
    }
  }
}




// Regular pictures, including rotated/cropped ones.
async function processImages(
  pics: any[],
  elements: any[],
  slideIndex: number,
  directory: any,
  relArr: any[]
) {
  for (const pic of pics) {
    const xfrm = pic["p:spPr"]?.["a:xfrm"];
    const position = {
      x: emuToPx(Number(xfrm?.["a:off"]?.x || 0)),
      y: emuToPx(Number(xfrm?.["a:off"]?.y || 0)),
    };
    const size = {
      width: emuToPx(Number(xfrm?.["a:ext"]?.cx || 0)),
      height: emuToPx(Number(xfrm?.["a:ext"]?.cy || 0)),
    };

    let embedId = pic["p:blipFill"]?.["a:blip"]?.["r:embed"];
    if (!embedId) embedId = pic["p:blipFill"]?.["a:blip"]?.["r:link"];

    const srcRect = pic["p:blipFill"]?.["a:srcRect"];
    const hasCropping =
      srcRect && (srcRect.l || srcRect.t || srcRect.r || srcRect.b);

    if (!embedId) continue;

    const rel = relArr.find((r) => r && r.Id === embedId);
    if (!rel) continue;

    const target = rel.Target;
    const imagePath = `ppt/${target.replace("..", "").replace(/^\//, "")}`;

    const imageFile = directory.files.find((f: any) => f.path === imagePath);
    if (!imageFile) continue;

    const imageBuffer = await imageFile.buffer();
    const imageName = `slide${slideIndex + 1}_${path.basename(imagePath)}`;
    elements.push({
      type: "image",
      name: imageName,
      src: toDataUri(imageBuffer, imagePath),
      position,
      size,
      cropped: hasCropping,
      cropInfo: srcRect || null,
    });
  }
}





// Walk a shape tree, handling shapes, groups (incl. nested), and images.
async function processShapes(
  shapeElement: any,
  elements: any[],
  slideIndex: number,
  directory: any,
  relArr: any[]
) {
  if (!shapeElement) return;

  // Group shapes.
  if (shapeElement["p:grpSp"]) {
    const groups = Array.isArray(shapeElement["p:grpSp"])
      ? shapeElement["p:grpSp"]
      : [shapeElement["p:grpSp"]];

    for (const group of groups) {
      if (group["p:sp"]) {
        const shapesInGroup = Array.isArray(group["p:sp"])
          ? group["p:sp"]
          : [group["p:sp"]];
        for (const shape of shapesInGroup) {
          const { texts, position, size } = extractTextFromElement(shape);
          if (texts.length > 0) {
            elements.push({
              type: "text",
              content: texts.join(" "),
              position,
              size,
              source: "group",
            });
          }
        }
      }

      if (group["p:pic"]) {
        const picsInGroup = Array.isArray(group["p:pic"])
          ? group["p:pic"]
          : [group["p:pic"]];
        await processImages(picsInGroup, elements, slideIndex, directory, relArr);
      }

      if (group["p:grpSp"]) {
        await processShapes(group, elements, slideIndex, directory, relArr);
      }
    }
  }

  // Individual shapes (text).
  if (shapeElement["p:sp"]) {
    const shapes = Array.isArray(shapeElement["p:sp"])
      ? shapeElement["p:sp"]
      : [shapeElement["p:sp"]];

    for (const shape of shapes) {
      const { texts, position, size } = extractTextFromElement(shape);
      if (texts.length > 0) {
        elements.push({
          type: "text",
          content: texts.join(" "),
          position,
          size,
          source: "individual",
        });
      }
    }
  }

  // Individual pictures.
  if (shapeElement["p:pic"]) {
    const pics = Array.isArray(shapeElement["p:pic"])
      ? shapeElement["p:pic"]
      : [shapeElement["p:pic"]];
    await processImages(pics, elements, slideIndex, directory, relArr);
  }

  // Custom-geometry image fills on individual shapes.
  if (shapeElement["p:sp"]) {
    const shapes = Array.isArray(shapeElement["p:sp"])
      ? shapeElement["p:sp"]
      : [shapeElement["p:sp"]];
    await processCustomShapes(shapes, elements, slideIndex, directory, relArr);
  }

  // Custom-geometry image fills inside groups.
  if (shapeElement["p:grpSp"]) {
    const groups = Array.isArray(shapeElement["p:grpSp"])
      ? shapeElement["p:grpSp"]
      : [shapeElement["p:grpSp"]];
    for (const group of groups) {
      if (group["p:sp"]) {
        const shapesInGroup = Array.isArray(group["p:sp"])
          ? group["p:sp"]
          : [group["p:sp"]];
        await processCustomShapes(
          shapesInGroup,
          elements,
          slideIndex,
          directory,
          relArr
        );
      }
    }
  }
}




// Decodes the base64 images already in the result and writes them to `outputDir`
// same as in V1
// as real files. The JSON is untouched. Images stay as data-URI strings too.
export function saveImagesToFolder(result: ExtractResult, outputDir: string): string[] {
  if (!fs.existsSync(outputDir)) fs.mkdirSync(outputDir, { recursive: true });
  const written: string[] = [];
  for (const slide of Object.values(result.slides)) {
    for (const el of slide.elements) {
      if (el.type !== "image" || typeof el.src !== "string") continue;
      const m = el.src.match(/^data:.*?;base64,(.*)$/);
      if (!m || !m[1]) continue;
      const filename = el.name || "image.png";
      fs.writeFileSync(path.join(outputDir, filename), Buffer.from(m[1], "base64"));
      written.push(filename);
    }
  }
  return written;
}





export interface ExtractResult {
  head: {
    slides: number;
    size: { width: number; height: number };
  };
  slides: Record<string, { title: string; elements: any[] }>;
}

// Extract a PPTX given its raw bytes. 
// Throws if the buffer is not a valid
// PPTX (e.g. missing ppt/presentation.xml).
export async function extractPptxFromBuffer(
  buffer: Buffer
): Promise<ExtractResult> {
  const directory = await unzipper.Open.buffer(buffer);
  const parser = new XMLParser({
    ignoreAttributes: false,
    attributeNamePrefix: "",
  });

  // Presentation metadata.
  const presFile = directory.files.find(
    (f) => f.path === "ppt/presentation.xml"
  );
  if (!presFile) {
    throw new Error(
      "Not a valid .pptx: ppt/presentation.xml was not found in the archive."
    );
  }
  const presXml = await presFile.buffer();
  const presJson = parser.parse(presXml.toString());

  const slideSize = presJson["p:presentation"]["p:sldSz"];
  const sldIdRaw = presJson["p:presentation"]["p:sldIdLst"]?.["p:sldId"];
  const slideCount = Array.isArray(sldIdRaw) ? sldIdRaw.length : sldIdRaw ? 1 : 0;

  const presentationHead = {
    slides: slideCount,
    size: {
      width: emuToPx(slideSize.cx),
      height: emuToPx(slideSize.cy),
    },
  };

  const slides: Record<string, { title: string; elements: any[] }> = {};

  // Sort slide files numerically (slide2 before slide10).
  const slideFiles = directory.files
    .filter((f) => /^ppt\/slides\/slide\d+\.xml$/.test(f.path))
    .sort((a, b) => {
      const na = Number(a.path.match(/slide(\d+)\.xml$/)?.[1] || 0);
      const nb = Number(b.path.match(/slide(\d+)\.xml$/)?.[1] || 0);
      return na - nb;
    });

  for (let i = 0; i < slideFiles.length; i++) {
    const slideFile = slideFiles[i];
    if (!slideFile) continue;

    const slideXml = await slideFile.buffer();
    const slideJson = parser.parse(slideXml.toString());

    const elements: any[] = [];

    const slideNum = Number(slideFile.path.match(/slide(\d+)\.xml$/)?.[1] || i + 1);
    const relFile = directory.files.find(
      (f) => f.path === `ppt/slides/_rels/slide${slideNum}.xml.rels`
    );
    let relJson: any = {};
    if (relFile) {
      const relXml = await relFile.buffer();
      relJson = parser.parse(relXml.toString());
    }

    const rels = relJson?.["Relationships"]?.["Relationship"];
    const relArr = rels ? (Array.isArray(rels) ? rels : [rels]) : [];

    const spTree = slideJson["p:sld"]?.["p:cSld"]?.["p:spTree"];
    if (spTree) {
      await processShapes(spTree, elements, slideNum - 1, directory, relArr);
    }

    slides[`slide${slideNum}`] = {
      title: `Slide ${slideNum}`,
      elements,
    };
  }

  return { head: presentationHead, slides };
}
