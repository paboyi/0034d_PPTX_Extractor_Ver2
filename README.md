still working on the ReadME. It will be uploaded soon.

# 0034d_PPTX_Extractor_Ver2



## ABOUT
// It accepts .pptx ONLY, enforced in three places:
//   1. file extension check
//   2. MIME type check
//   3. actual archive validation inside the extractor (must contain
//      ppt/presentation.xml)  (this is the check that can't be spoofed.)


## Run

npm install
npm start
# open http://localhost:3000

for a clean re-install:
rm -rf node_modules package-lock.json
npm install
npm start



## Strick Project Requirement
On the "pptx only" requirement, the gate is enforced in three layers so it can't be trivially bypassed:

Browser — accept=".pptx" plus a name/type check before upload.
Server — multer rejects anything not ending in .pptx → 400 "Only .pptx files are accepted."
Content — if someone renames a random file to .pptx, the extractor fails to find ppt/presentation.xml and returns 422 "couldn't be read as PowerPoint."


Why?
Because this extracvtor only know how to work on PPTX file for the time being.
See road Map


## Road Map
Current version V2
- [] Save images to extrernal folder like V1
- [] Accept other file types e.g epub, pdf, docx etc
- [] Decide if it's wise to have frontend & Backend separated for later versions
right now it runs as a web service on Render
<!-- only variable server.ts reads is PORT -->

## It's plain documentation for a developer
describing how to call your backend directly without the web page (for example from curl, Postman, or your own code.) 

Reads literally:
* POST /api/extract: send an HTTP POST request to the /api/extract URL (that's the endpoint in server.ts). <br />
* field file: put the PowerPoint in a form field named file (this is the name multer reads in upload.single("file")). <br />
* multipart/form-data: send it as a file upload, not JSON.