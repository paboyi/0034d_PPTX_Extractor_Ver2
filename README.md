0034d_PPTX_Extractor_ ~~v.2~~ v3



## About
This is a PPTX extraction tool.

It accepts .pptx ONLY, enforced in three places:
1. file extension check
2. MIME type check
3. actual archive validation inside the extractor (must contain ppt/presentation.xml)  (this is the check that can't be spoofed.)


## Run
    npm install
    npm start
should start on port `http://localhost:{PORT}`

for a clean re-install: run `rm -rf node_modules package-lock.json`
then:
    npm install
    npm start


## Strick Project Requirement
On the "*pptx only*" requirement, the gate is enforced in three layers so it can't be trivially bypassed:

* Browser → accept = ".pptx" plus a name/type check before upload.
* Server → multer rejects anything not ending in .pptx → 400 "Only .pptx files are accepted."
* Content → if someone renames a random file to .pptx, the extractor fails to find ppt/presentation.xml and returns 422 "couldn't be read as PowerPoint."


<b>Why?</b> </br>
Because this extractor only know how to work on PPTX file for the time being.
_See road Map for possible future extension_


## Road Map
Current version V2
- [ ] Save images to extrernal folder like in V1
    - [ ] downloadable with JSON result
- [ ] Accept other file types e.g epub, pdf, docx etc
- [ ] Decide if it's wise to have frontend & Backend separated for later versions (currently on ~~Render~~ AWS Lambda)
<!-- only variable server.ts reads is PORT -->


## It's plain documentation for a developer
This describes how to call (access) the backend directly without the frontend web page 
(for example from curl, Postman, or your own code.) 

Literal endpoints:
* <b>POST /api/extract</b>: 
    send an HTTP POST request to the /api/extract URL (that's the endpoint in server.ts). <br />

* <b>field file</b>: 
    put the PowerPoint in a form field named file (this is the name multer reads in upload.single("file")). <br />

* <b>multipart/form-data</b>: 
    send it as a file upload, not JSON.