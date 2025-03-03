const fs = require('fs');
const path = './.clasp.json';
const masterScriptId = '1eh1_6mAXZa0X0OFHLOUvEVgI3Sat4Y4oXyLD8exDors4wi85mqvA8Yj2';
const devScriptId = '1_QhTZARA2aYoHU3vmc_AG50WQLPwYRI3RYBc_Fmbq65z4EGUyE3kGKaF';
const arg = process.argv[2]; 
const scriptId = arg === 'master' ? masterScriptId : devScriptId;

// Read the current JSON file
fs.readFile(path, 'utf8', (err, data) => {
  if (err) {
    console.error('Error reading file:', err);
    return;
  }

  // Parse the JSON data
  const json = JSON.parse(data);

  // Modify the value field based on the passed argument
  json.scriptId = scriptId;

  // Write the updated JSON back to the file
  fs.writeFile(path, JSON.stringify(json, null, 2), (err) => {
    if (err) {
      console.error('Error writing file:', err);
    } else {
      console.log(`File updated with scriptId: ${scriptId}`);
    }
  });
});
