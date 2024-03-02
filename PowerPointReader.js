const fs = require('fs').promises;
const JSZip = require('jszip');
const xml2js = require('xml2js').parseStringPromise;

class PowerPointReader {
    constructor() {
        this.slides = [];
    }

    async load(pptxPath) {
        try {
            const data = await fs.readFile(pptxPath);
            const zip = await JSZip.loadAsync(data);
            const slideFiles = Object.keys(zip.files).filter(fileName => fileName.match(/ppt\/slides\/slide[0-9]+.xml/));

            for (let i = 0; i < slideFiles.length; i++) {
                const slideFile = slideFiles[i];
                const content = await zip.files[slideFile].async("string");
                const result = await xml2js(content);
                const title = this.extractTitle(result);
                
                this.slides.push({
                    getTitle: () => title || "No title"
                });
            }

            return this; // Allow chaining
        } catch (err) {
            console.error('Error loading PowerPoint file:', err);
            throw err; // Re-throw to allow caller to handle it
        }
    }

    extractTitle(slideXml) {
        // Simplified extraction logic for demonstration purposes
        try {
            const shapes = slideXml['p:sld']['p:cSld'][0]['p:spTree'][0]['p:sp'] || [];
            for (const shape of shapes) {
                if (shape['p:nvSpPr'][0]['p:nvPr'][0]['p:ph']) {
                    const placeholder = shape['p:nvSpPr'][0]['p:nvPr'][0]['p:ph'][0]['$'];
                    if (placeholder.type === 'ctrTitle' || placeholder.type === 'title') {
                        return shape['p:txBody'][0]['a:p'][0]['a:r'][0]['a:t'][0];
                    }
                }
            }
        } catch (error) {
            return "No title";
        }
        return "No title";
    }
}

module.exports = PowerPointReader;
