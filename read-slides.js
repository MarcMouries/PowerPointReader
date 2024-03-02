import PowerPointReader from './PowerPointReader.js';
import { exit } from 'process';

function printUsage() {
    console.error("Usage:");
    console.error("  node readPPTX.js <pptx file path>");
}
if (process.argv.length !== 3) {
    printUsage();
    exit(1);
}

const filePath = process.argv[2];
const reader = new PowerPointReader();

try {
    const pres = await reader.load(filePath);
    const slideCount = pres.slides.length;
    console.log(`Total slides: ${slideCount}\n`);

    for (let i = 0; i < slideCount; i++) {
        const slide = pres.slides[i];
        const title = slide.getTitle();
        console.log(`${i + 1}: ${title}`);
    }
} catch (error) {
    console.error("Error reading PowerPoint file:", error);
    exit(1);
}