const fs = require('fs');
const text = fs.readFileSync('d:\\Project\\Timesheet\\TIME-SHEET-APNS-FORM\\index.html', 'utf8');

// Find the merged-content colspan td and what's after it
const mergedIdx = text.indexOf('colspan="8" class="merged-content"');
console.log('MERGED CONTENT + ACTION TD:');
console.log(JSON.stringify(text.substring(mergedIdx - 20, mergedIdx + 700)));
