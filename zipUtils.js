const fs = require('fs-extra');
const JSZip = require('jszip');
const path = require('path');

async function unzipDocx(docxPath, tempFolder) {
  if (process.env.DEBUG) {
    console.log(`Starting to unzip DOCX file: ${docxPath}`);
  }
  const data = await fs.readFile(docxPath);
  const zip = await JSZip.loadAsync(data);
  if (process.env.DEBUG) {
    console.log(`Loaded DOCX file into JSZip.`);
  }
  const fileWrites = [];
  for (const relativePath in zip.files) {
    const file = zip.files[relativePath];
    const destPath = path.join(tempFolder, relativePath);
    if (file.dir) {
      if (process.env.DEBUG) {
        console.log(`Ensuring directory exists: ${destPath}`);
      }
      await fs.ensureDir(destPath);
    } else {
      // Ensure the directory exists before trying to write the file
      const dirName = path.dirname(destPath);
      await fs.ensureDir(dirName);
      
      if (process.env.DEBUG) {
        console.log(`Writing file: ${destPath}`);
      }
      const fileWrite = file.async('nodebuffer').then(content => {
        return fs.writeFile(destPath, content);
      }).catch(error => {
        if (process.env.DEBUG) {
          console.error(`Error writing file ${destPath}:`, error);
        }
        throw error; // Re-throw the error to handle it in the outer scope if necessary
      });
      fileWrites.push(fileWrite);
    }
  }
  await Promise.all(fileWrites);
  if (process.env.DEBUG) {
    console.log(`Finished unzipping DOCX file.`);
  }
}

async function createDocx(docxPath, tempFolder) {
  const zip = new JSZip();
  const allFiles = (dir, fileList = []) => {
    fs.readdirSync(dir).forEach(file => {
      const filePath = path.join(dir, file);
      if (fs.lstatSync(filePath).isDirectory()) {
        allFiles(filePath, fileList);
      } else {
        fileList.push(filePath);
      }
    });
    return fileList;
  };

  allFiles(tempFolder).forEach(filePath => {
    const relativePath = path.relative(tempFolder, filePath);
    const content = fs.readFileSync(filePath);
    zip.file(relativePath, content);
  });
  const content = await zip.generateAsync({ type: 'nodebuffer' });
  fs.writeFileSync(docxPath, content);
}

module.exports = { unzipDocx, createDocx };
