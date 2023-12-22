#!/usr/bin/env node

const path = require('path');
const readline = require('readline');
const { generateTempFolderName, deleteFolder } = require('./fileUtils');
const { unzipDocx, createDocx } = require('./zipUtils');
const { removeProtection } = require('./xmlUtils');

const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout
});

const mainUnprotect = async function mainUnprotect(protected, unprotected) {
  const tempFolder = generateTempFolderName();
  try {
    if (process.env.DEBUG) {
      console.log('Debugging is on.');
    }
    if (process.env.DEBUG) {
      console.log(`Temporary folder created: ${tempFolder}`);
    }

    await unzipDocx(protected, tempFolder);
    if (process.env.DEBUG) {
      console.log(`Unzipped DOCX file to temporary folder.`);
    }
    await removeProtection(tempFolder);
    if (process.env.DEBUG) {
      console.log(`Removed protection from DOCX file.`);
    }
    await createDocx(unprotected, tempFolder);
    if (process.env.DEBUG) {
      console.log(`Created unprotected DOCX file.`);
    }
  } finally {
    deleteFolder(tempFolder);
    if (process.env.DEBUG) {
      console.log(`Deleted temporary folder.`);
    }
    process.exit(0);
  }
}

const args = process.argv.slice(2);
let protectedDocx = args[0];
let unprotectedDocx = args[1];

if (!protectedDocx) {
  rl.question('Enter the path to the protected DOCX file: ', (inputProtectedDocx) => {
    protectedDocx = inputProtectedDocx;
    if (!unprotectedDocx) {
      const defaultUnprotectedName = path.join(path.dirname(protectedDocx), path.basename(protectedDocx, '.docx') + '_unprotected.docx');
      rl.question(`Enter the path to save the unprotected DOCX file [${defaultUnprotectedName}]: `, (inputUnprotectedDocx) => {
        unprotectedDocx = inputUnprotectedDocx || defaultUnprotectedName;
        rl.close();
        mainUnprotect(protectedDocx, unprotectedDocx);
      });
    } else {
      rl.close();
      mainUnprotect(protectedDocx, unprotectedDocx);
    }
  });
} else {
  if (!unprotectedDocx) {
    unprotectedDocx = path.join(path.dirname(protectedDocx), path.basename(protectedDocx, '.docx') + '_unprotected.docx');
  }
  mainUnprotect(protectedDocx, unprotectedDocx);
}

module.exports = {mainUnprotect}