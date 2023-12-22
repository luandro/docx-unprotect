const fs = require('fs-extra');
const os = require('os');
const crypto = require('crypto');
const path = require('path');

function deleteFolder(folderPath) {
  fs.removeSync(folderPath);
}

function generateTempFolderName() {
  const baseString = "docx_unprotect_";
  const randomString = crypto.randomBytes(8).toString('hex');
  const tempFolderName = path.join(os.tmpdir(), `${baseString}${randomString}`);
  return tempFolderName;
}

module.exports = { deleteFolder, generateTempFolderName };
