const mammoth = require('mammoth');
const TurndownService = require('turndown');
const fs = require('fs');
const path = require('path');

(async () => {
  try {
    const dir = __dirname;
    const files = await fs.promises.readdir(dir);
    const docxFiles = files.filter(f => f.toLowerCase().endsWith('.docx'));
    if (docxFiles.length === 0) {
      console.log('No .docx files found in', dir);
      return;
    }

    const turndownService = new TurndownService({ headingStyle: 'atx' });

    for (const file of docxFiles) {
      const inputPath = path.join(dir, file);
      const outputName = file.replace(/\.docx$/i, '.md');
      const outputPath = path.join(dir, outputName);
      try {
        const result = await mammoth.convertToHtml({ path: inputPath });
        const html = result.value; // The generated HTML
        const md = turndownService.turndown(html);
        await fs.promises.writeFile(outputPath, md, 'utf8');
        console.log(`Converted: ${file} -> ${outputName}`);
      } catch (err) {
        console.error(`Failed to convert ${file}:`, err.message || err);
      }
    }
  } catch (err) {
    console.error('Error:', err.message || err);
    process.exit(1);
  }
})();
