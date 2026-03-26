const fetch = (...args) => import('node-fetch').then(({default: f}) => f(...args));
const XLSX = require('xlsx');
const fs = require('fs');

const TENANT_ID     = process.env.TENANT_ID;
const CLIENT_ID     = process.env.CLIENT_ID;
const CLIENT_SECRET = process.env.CLIENT_SECRET;

async function main() {
  // ── 1. Auth ──────────────────────────────────────────────────────────────
  console.log('Fetching access token...');
  const tokenResp = await fetch(
    `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`,
    {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        grant_type:    'client_credentials',
        client_id:     CLIENT_ID,
        client_secret: CLIENT_SECRET,
        scope:         'https://graph.microsoft.com/.default',
      }).toString(),
    }
  );
  const tokenData = await tokenResp.json();
  if (!tokenData.access_token) {
    console.error('Token error:', JSON.stringify(tokenData));
    process.exit(1);
  }
  const { access_token } = tokenData;
  console.log('✓ Token obtained');

  // ── 2. Site ───────────────────────────────────────────────────────────────
  console.log('Getting site...');
  const siteResp = await fetch(
    'https://graph.microsoft.com/v1.0/sites/seveninsurancebrokers.sharepoint.com:/sites/SIBIntranet',
    { headers: { Authorization: `Bearer ${access_token}` } }
  );
  const siteData = await siteResp.json();
  if (!siteData.id) {
    console.error('Site error:', JSON.stringify(siteData));
    process.exit(1);
  }
  const { id: siteId } = siteData;
  console.log('✓ Site ID:', siteId);

  // ── 3. Drive ──────────────────────────────────────────────────────────────
  console.log('Listing drives...');
  const drivesResp = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${siteId}/drives`,
    { headers: { Authorization: `Bearer ${access_token}` } }
  );
  const drivesData = await drivesResp.json();
  console.log('Available drives:', drivesData.value?.map(d => d.name));

  const docsDrive = drivesData.value?.find(d =>
    d.name === 'Documents' || d.name === 'Shared Documents' ||
    (d.webUrl || '').includes('Shared%20Documents') ||
    (d.webUrl || '').includes('Shared Documents')
  );
  if (!docsDrive) {
    console.error('Could not find Documents drive. Drives:', JSON.stringify(drivesData.value?.map(d => ({ name: d.name, id: d.id }))));
    process.exit(1);
  }
  console.log('✓ Using drive:', docsDrive.name, docsDrive.id);

  // ── 4. File ───────────────────────────────────────────────────────────────
  console.log('Getting file...');
  const fileResp = await fetch(
    `https://graph.microsoft.com/v1.0/drives/${docsDrive.id}/root:/GIM Weekly Report - 2026.xlsx`,
    { headers: { Authorization: `Bearer ${access_token}` } }
  );
  const fileData = await fileResp.json();
  if (!fileData['@microsoft.graph.downloadUrl']) {
    console.error('File error:', JSON.stringify(fileData));
    process.exit(1);
  }
  const downloadUrl = fileData['@microsoft.graph.downloadUrl'];
  console.log('✓ Download URL obtained');

  // ── 5. Download ───────────────────────────────────────────────────────────
  console.log('Downloading file...');
  const dlResp = await fetch(downloadUrl);
  const buf = await dlResp.arrayBuffer();
  console.log('✓ File size:', buf.byteLength, 'bytes');

  // ── 6. Parse ──────────────────────────────────────────────────────────────
  console.log('Parsing Excel...');
  const wb = XLSX.read(Buffer.from(buf), { type: 'buffer', cellDates: true });
  console.log('Sheet names:', wb.SheetNames);

  const sheetName =
    wb.SheetNames.find(n => /written business.*2026/i.test(n)) ||
    wb.SheetNames.find(n => /written business/i.test(n)) ||
    wb.SheetNames[0];
  console.log('✓ Using sheet:', sheetName);

  const ws = wb.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });
  console.log('Total rows:', rows.length);

  if (rows.length === 0) {
    console.error('ERROR: No rows found in sheet. Check the sheet name and format.');
    process.exit(1);
  }

  // Log ALL column names so you can spot mismatches easily
  const keys = Object.keys(rows[0] || {});
  console.log('Column names found:', JSON.stringify(keys));

  const exactCol = (name) =>
    keys.find(k => k.trim().toLowerCase() === name.trim().toLowerCase()) || '';

  const C = {
    week:     exactCol('WEEK'),
    div:      exactCol('DIV'),
    broker:   exactCol('Broker'),
    month:    exactCol('Month'),
    policy:   keys.find(k => k.trim() === 'Policy Type') || exactCol('Policy Type'),
    product:  exactCol('Product Type') || exactCol('Product name'),
    provider: exactCol('Insurance Co. / Provider') || exactCol('Insurance Co'),
    tier:     exactCol('Category'),
    commUSD:  exactCol('Commission Value (USD)'),
  };
  console.log('Column mapping:', JSON.stringify(C));

  // Warn about unmapped critical columns
  if (!C.commUSD) console.warn('⚠️  Could not map "Commission Value (USD)" — check column name in Excel');
  if (!C.week)    console.warn('⚠️  Could not map "WEEK" — check column name in Excel');

  // ── 7. Build records ──────────────────────────────────────────────────────
  const records = [];
  let skipped = 0;

  for (const row of rows) {
    const u = parseFloat(row[C.commUSD]) || 0;
    const w = parseFloat(row[C.week])    || 0;
    if (!u || !w) { skipped++; continue; }

    const div  = String(row[C.div]  || 'Unknown').trim();
    let   prod = String(row[C.product] || '').trim();

    if      (/group medical/i.test(prod))          prod = 'Group Medical';
    else if (/group life/i.test(prod))             prod = 'Group Life';
    else if (/individual.*basic/i.test(prod))      prod = 'Med Individual (Basic)';
    else if (/individual/i.test(prod))             prod = 'Med Individual';
    else if (/retail/i.test(prod))                 prod = 'Retail';
    else if (/general/i.test(prod))                prod = 'General';

    const tier     = String(row[C.tier]   || 'D').trim();
    const policyVal = String(row[C.policy] || '');

    records.push({
      w,
      b:    String(row[C.broker] || div).trim(),
      div,
      m:    String(row[C.month]  || '').trim(),
      pt:   /new business/i.test(policyVal) ? 'New' : 'Renewal',
      prod: prod || 'Other',
      prov: String(row[C.provider] || '').trim(),
      t:    ['A','B','C','D'].includes(tier) ? tier : 'D',
      u,
    });
  }

  console.log(`✓ Parsed records: ${records.length}  (skipped ${skipped} rows with no commission/week)`);

  if (records.length === 0) {
    console.error('ERROR: 0 records after parsing. Common causes:');
    console.error('  1. Column names in Excel differ from what the script expects (see mapping above)');
    console.error('  2. Commission or week values are not numeric in the sheet');
    process.exit(1);
  }

  // ── 8. Write — root data.json (matches what the workflow commits) ─────────
  const output = { updated: new Date().toISOString(), records };
  fs.writeFileSync('data.json', JSON.stringify(output));
  console.log('✓ Saved data.json with', records.length, 'records');
}

main().catch(err => {
  console.error('ERROR:', err.message);
  process.exit(1);
});
