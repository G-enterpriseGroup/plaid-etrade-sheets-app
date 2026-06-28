const fs = require('fs');
const path = require('path');

const root = process.cwd();

const required = [
  'apps-script/Code.gs',
  'apps-script/Sidebar.html',
  'apps-script/PlaidLinkModal.html',
  'apps-script/PlaidModalBackend.gs',
  'apps-script/MarketAnalysisReport.gs',
  '.gitignore',
  '.claspignore',
  'README.md',
  'SECURITY.md',
  'SETUP.md'
];

let failed = false;

for (const file of required) {
  if (!fs.existsSync(path.join(root, file))) {
    console.error(`Missing required file: ${file}`);
    failed = true;
  }
}

const skipDirs = new Set(['.git', 'node_modules', '.venv', 'venv', '__pycache__']);
const skipExtensions = new Set(['.png', '.jpg', '.jpeg', '.gif', '.webp', '.zip', '.pdf', '.ico']);
const skipSecretScanFiles = new Set([
  'SETUP_SECRETS.md',
  'docs/GITHUB_ACTIONS_MARKET_ANALYSIS_SETUP.md',
  'backend/README.md',
  'backend/.env.example'
]);

const forbiddenPatterns = [
  /access-[a-z]+-[a-z0-9-]{20,}/i,
  /PLAID_SECRET\s*[:=]\s*(?!your\b|replace\b|<|\$\{|process\.env)(['"]?)[A-Za-z0-9_\-]{20,}\1/i,
  /PLAID_ACCESS_TOKEN\s*[:=]\s*(?!your\b|replace\b|<|\$\{|process\.env)(['"]?)[A-Za-z0-9_\-]{20,}\1/i,
  /SECRET\s*[:=]\s*(?!your\b|replace\b|<|\$\{|process\.env)(['"]?)[A-Za-z0-9_\-]{20,}\1/i
];

function walk(dir) {
  let out = [];
  for (const item of fs.readdirSync(dir)) {
    if (skipDirs.has(item)) continue;
    const full = path.join(dir, item);
    const stat = fs.statSync(full);
    if (stat.isDirectory()) out = out.concat(walk(full));
    else out.push(full);
  }
  return out;
}

for (const file of walk(root)) {
  const rel = path.relative(root, file).replace(/\\/g, '/');
  if (skipExtensions.has(path.extname(file).toLowerCase())) continue;
  if (skipSecretScanFiles.has(rel)) continue;
  if (rel.startsWith('runtime/market-inputs/') || rel.startsWith('runtime/market-outputs/')) continue;

  const text = fs.readFileSync(file, 'utf8');
  for (const re of forbiddenPatterns) {
    if (re.test(text)) {
      console.error(`Possible secret found in ${rel}`);
      failed = true;
    }
  }
}

const appsScriptFiles = [
  'apps-script/Code.gs',
  'apps-script/PlaidModalBackend.gs',
  'apps-script/MarketAnalysisReport.gs'
];

for (const file of appsScriptFiles) {
  const full = path.join(root, file);
  if (!fs.existsSync(full)) continue;
  const code = fs.readFileSync(full, 'utf8');
  try {
    new Function(code);
  } catch (err) {
    console.error(`Syntax check failed in ${file}: ${err.message}`);
    failed = true;
  }
}

if (failed) process.exit(1);
console.log('Repo checks passed.');
