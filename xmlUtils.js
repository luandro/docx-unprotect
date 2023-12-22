const fs = require('fs-extra');
const { parseString, Builder } = require('xml2js');
const path = require('path');

async function removeProtection(tempFolder) {
  const settingsPath = path.join(tempFolder, 'word', 'settings.xml');
  if (process.env.DEBUG) {
    console.log(`Reading settings.xml from: ${settingsPath}`);
  }
  const xml = await fs.readFile(settingsPath, 'utf8');
  parseString(xml, (err, result) => {
    if (err) {
      if (process.env.DEBUG) {
        console.error(`Error parsing XML: ${err}`);
      }
      throw err;
    }
    if (process.env.DEBUG) {
      console.log(`Parsed settings.xml successfully.`);
    }
    if (result['w:settings'] && result['w:settings']['w:documentProtection']) {
      if (process.env.DEBUG) {
        console.log(`Document protection found, removing...`);
      }
      delete result['w:settings']['w:documentProtection'];
    } else if (process.env.DEBUG) {
      console.log(`No document protection found.`);
    }
    const builder = new Builder();
    const newXml = builder.buildObject(result);
    fs.writeFileSync(settingsPath, newXml);
    if (process.env.DEBUG) {
      console.log(`Written new settings.xml without protection.`);
    }
  });
}

module.exports = { removeProtection };
