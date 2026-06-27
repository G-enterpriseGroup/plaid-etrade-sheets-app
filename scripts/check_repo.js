const fs = require('fs');
const path = require('path');

const required = [
  'apps-script/Code.js',
  'apps-script/Sidebar.html',
  'apps-script/appsscript.json',
  '.gitignore',
  '.claspignore',
  'README.md',
  'SECURITY.md',
  'SETUP.md'
];

let failed = false;
for (const file of required) {
  if (!fs.existsSync(path.join(process.cwd(), file))) {
    console.error(`Missing required file: ${file}`);
    failed = true;
  }
}

const forbiddenPatterns = [
  /access-[a-z]+-[a-z0-9-]{20,}/i,
  /PLAID_SECRET\s*=\s*[^\n\r]+/i,
  /SECRET=\w{20,}/i
];

function walk(dir) {
  let out = [];
  for (const item of fs.readdirSync(dir)) {
    if (['.git', 'node_modules'].includes(item)) continue;
    const full = path.join(dir, item);
    const stat = fs.statSync(full);
    if (stat.isDirectory()) out = out.concat(walk(full));
    else out.push(full);
  }
  return out;
}

for (const file of walk(process.cwd())) {
  const rel = path.relative(process.cwd(), file);
  if (rel.endsWith('.png') || rel.endsWith('.zip')) continue;
  const text = fs.readFileSync(file, 'utf8');
  for (const re of forbiddenPatterns) {
    if (re.test(text)) {
      console.error(`Possible secret found in ${rel}`);
      failed = true;
    }
  }
}

const code = fs.readFileSync(path.join(process.cwd(), 'apps-script/Code.js'), 'utf8');
new Function(code);

if (failed) process.exit(1);
console.log('Repo checks passed.');
