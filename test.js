const fs = require('fs');
const html = fs.readFileSync('/Users/macbook/Documents/ODC/Index.html', 'utf8');
const script = html.match(/<script>([\s\S]*)<\/script>/)[1];
fs.writeFileSync('test_script.js', script);
