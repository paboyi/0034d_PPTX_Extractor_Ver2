import * as fs from "fs";
import * as path from "path";
import * as unzipper from "unzipper";
import {XMLParser} from "fast-xml-parser";



// measurement: EMUs --> Pixels
//EMU ia a pptx measuring unit
function emuToPx(emu: number): number {
    return Math.round(emu/9525); 
}




// Extracting text from shape or group.  slide 10 has groups with text & images
function extractTextFromElement(element: any): { texts: string[], position: any, size: any } {
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

    // Extract text if this element has text body
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





async function processCustomShapes(shapes: any[], elements: any[], slideIndex: number, directory: any, relArr: any[], outputDir: string) {

    for (const shape of shapes) {

        // Check if this shape has an image fill (blipFill) //e.g. slide 3 images
        const blipFill = shape["p:spPr"]?.["a:blipFill"];
        if (blipFill) {
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
            
            if (embedId) {
                const rel = relArr.find((r: any) => r && r.Id === embedId);
                if (rel) {
                    const target = rel.Target;
                    const imagePath = target.startsWith("../") 
                        ? `ppt/${target.substring(3)}` 
                        : `ppt/${target}`;

                    const imageFile = directory.files.find((f: any) => f.path === imagePath);
                    if (imageFile) {
                        try {
                            const imageBuffer = await imageFile.buffer();
                            const imageName = `slide${slideIndex + 1}_${path.basename(imagePath)}`;
                            fs.writeFileSync(path.join(outputDir, imageName), imageBuffer);

                            elements.push({
                                type: "image",
                                src: imageName,
                                position,
                                size,
                                imageType: "customShape", // Marking as a custom shape for refernce
                                hasCustomGeometry: true
                            });
                        } catch (error) {
                            console.warn(`Failed to process custom shape image ${imagePath}:`, error);
                        }
                    }
                }
            }
        }
    }
}






// Process shapes (handles both individual shapes and groups)
async function processShapes(shapeElement: any, elements: any[], slideIndex: number, directory: any, relArr: any[], outputDir: string) {
    if (!shapeElement) return;



    // Handling GROUP shapes
    if (shapeElement["p:grpSp"]) {
        const groups = Array.isArray(shapeElement["p:grpSp"]) ? shapeElement["p:grpSp"] : [shapeElement["p:grpSp"]];
        
        for (const group of groups) {
            // Process shapes within the group
            if (group["p:sp"]) {
                const shapesInGroup = Array.isArray(group["p:sp"]) ? group["p:sp"] : [group["p:sp"]];
                for (const shape of shapesInGroup) {
                    const { texts, position, size } = extractTextFromElement(shape);
                    if (texts.length > 0) {
                        elements.push({
                            type: "text",
                            content: texts.join(" "),
                            position,
                            size,
                            source: "group"
                        });
                    }
                }
            }
            
            // Process images within the group
            if (group["p:pic"]) {
                const picsInGroup = Array.isArray(group["p:pic"]) ? group["p:pic"] : [group["p:pic"]];
                processImages(picsInGroup, elements, slideIndex, directory, relArr, outputDir);
            }

            // Recursively process nested groups -> kinda [[[]]]
            if (group["p:grpSp"]) {
                processShapes(group, elements, slideIndex, directory, relArr, outputDir);
            }
        }
    }



    // Handling INDIVIDUAL shapes
    if (shapeElement["p:sp"]) {
        const shapes = Array.isArray(shapeElement["p:sp"]) ? shapeElement["p:sp"] : [shapeElement["p:sp"]];

        for (const shape of shapes) {
            const { texts, position, size } = extractTextFromElement(shape);
            if (texts.length > 0) {
                elements.push({
                    type: "text",
                    content: texts.join(" "),
                    position,
                    size,
                    source: "individual"
                });
            }
        }
    }

    // Handling individual images
    if (shapeElement["p:pic"]) {
        const pics = Array.isArray(shapeElement["p:pic"]) ? shapeElement["p:pic"] : [shapeElement["p:pic"]];
        processImages(pics, elements, slideIndex, directory, relArr, outputDir);
    }



    // Handling custom geometry shapes with image fills (like slide 3)
    if (shapeElement["p:sp"]) {
        const shapes = Array.isArray(shapeElement["p:sp"]) ? shapeElement["p:sp"] : [shapeElement["p:sp"]];
        await processCustomShapes(shapes, elements, slideIndex, directory, relArr, outputDir);
    }


    // Also handle custom shapes within groups
    if (shapeElement["p:grpSp"]) {
    const groups = Array.isArray(shapeElement["p:grpSp"]) ? shapeElement["p:grpSp"] : [shapeElement["p:grpSp"]];
    
        for (const group of groups) {
            if (group["p:sp"]) {
                const shapesInGroup = Array.isArray(group["p:sp"]) ? group["p:sp"] : [group["p:sp"]];
                await processCustomShapes(shapesInGroup, elements, slideIndex, directory, relArr, outputDir);
            }
        }
    }

}





// support for transformed images e.g slide 8 of example pptx
// by transformed. I mean rotated, cropped etc
async function processImages(pics: any[], elements: any[], slideIndex: number, directory: any, relArr: any[], outputDir: string) {

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



        // Checking for different types of image references
        let embedId = pic["p:blipFill"]?.["a:blip"]?.["r:embed"];
        

        // --> Sometimes images use "r:link" instead of "r:embed"
        if (!embedId) {
            embedId = pic["p:blipFill"]?.["a:blip"]?.["r:link"];
        }

        // Checking  for 'cropping' information 
        const srcRect = pic["p:blipFill"]?.["a:srcRect"];
        const hasCropping = srcRect && (srcRect.l || srcRect.t || srcRect.r || srcRect.b);

        if (embedId) {
            const rel = relArr.find(r => r && r.Id === embedId);
            if (rel) {
                const target = rel.Target;
                const imagePath = `ppt/${target.replace("..", "").replace(/^\//, "")}`;

                const imageFile = directory.files.find((f: any) => f.path === imagePath);
                if (imageFile) {
                    const imageBuffer = await imageFile.buffer();
                    const imageName = `slide${slideIndex + 1}_${path.basename(imagePath)}`;
                    fs.writeFileSync(path.join(outputDir, imageName), imageBuffer);

                    elements.push({
                        type: "image",
                        src: imageName,
                        position,
                        size,
                        cropped: hasCropping,
                        cropInfo: srcRect || null
                    });
                }
            }
        }
    }
}


















// Main function?

async function extractPptx(filePath: string, outputDir: string) {
    const directory = await unzipper.Open.file(filePath);

    // 🪼🐙 REMOVE (choice) 🪼🐙
    fs.writeFileSync("directory.json", JSON.stringify(directory, null, 2));
    // 🪼🐙 REMOVE 🪼🐙

    const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: ""});

    if (!fs.existsSync(outputDir)) {
        fs.mkdirSync(outputDir, { recursive: true});
    }







    // ----------------------
    // Presentation metadata
    // ----------------------
    const presFile = directory.files.find(f => f.path === "ppt/presentation.xml");
    const presXml = await presFile!.buffer();
    const presJson = parser.parse(presXml.toString());

    const slideSize = presJson["p:presentation"]["p:sldSz"];
    const presentationHead = {
        slides: presJson["p:presentation"]["p:sldIdLst"]["p:sldId"].length,
        size: {
            width: emuToPx(slideSize.cx),
            height: emuToPx(slideSize.cy),
        }
    };









    const slides: any = {};

    // ----------------------
    // Process slides
    // ----------------------
    const slideFiles = directory.files.filter(f => f.path.startsWith("ppt/slides/slide"));
    
    for (let i = 0; i < slideFiles.length; i++) {
        const slideFile = slideFiles[i];
        if (!slideFile) {
            console.warn(`Slide file at index ${i} not found.`);
            continue;
        }
        
        const slideXml = await slideFile.buffer();
        const slideJson = parser.parse(slideXml.toString());

        const elements: any[] = [];



        // Get relationship data for this slide
        const relFile = directory.files.find(f => f.path === `ppt/slides/_rels/slide${i+1}.xml.rels`);
        let relJson: any = {};
        if (relFile) {
            const relXml = await relFile.buffer();
            relJson = parser.parse(relXml.toString());
        }

        const rels = relJson?.["Relationships"]?.["Relationship"];
        const relArr = rels ? (Array.isArray(rels) ? rels : [rels]) : [];



        // Processing all elements in the slide (shapes, groups, images)
        const spTree = slideJson["p:sld"]["p:cSld"]["p:spTree"];
        if (spTree) {
            processShapes(spTree, elements, i, directory, relArr, outputDir);
        }

        slides[`slide${i+1}`] = {
            title: `Slide ${i+1}`,
            elements
        };
    }


    // create object, return object 
    return { head: presentationHead, slides};
}







// ----------------------
// Run the extractor
// PS: name your pptx 'example.pptx' (or change the 'argument name')
// ----------------------
(async () => {
    const result = await extractPptx("example.pptx", "output_images"); 
    fs.writeFileSync("output.json", JSON.stringify(result, null, 2));
    console.log("Extraction complete. Check output.json and output_images/");
})();

// run with 'npx ts-node src/extract.ts'