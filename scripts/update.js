// Descarga los Excel de la carpeta compartida de OneDrive y genera data.json.
// Lo ejecuta GitHub Actions cada pocos minutos. El vínculo está en el secreto ONEDRIVE_SHARE_URL.
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');
const { parseAll, norm } = require('./parse.js');

const SHARE = (process.env.ONEDRIVE_SHARE_URL || '').trim();
const OUT = path.join(__dirname, '..', 'data.json');

async function main() {
  if (!SHARE) throw new Error('Falta el secreto ONEDRIVE_SHARE_URL');
  const tok = await fetch('https://api-badgerp.svc.ms/v1.0/token', {
    method: 'POST', headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ appId: '5cbed6ac-a083-4e14-b191-b4ba07653de2' }),
  }).then(r => r.json());
  const enc = 'u!' + Buffer.from(SHARE).toString('base64').replace(/=+$/, '').replace(/\//g, '_').replace(/\+/g, '-');
  const res = await fetch(`https://my.microsoftpersonalcontent.com/_api/v2.0/shares/${enc}/driveitem?$expand=children`, {
    headers: { Authorization: 'Badger ' + tok.token, Prefer: 'autoredeem' },
  });
  if (!res.ok) throw new Error('OneDrive respondió ' + res.status + ': ' + (await res.text()).slice(0, 300));
  const folder = await res.json();
  const kids = folder.children || [];
  const pick = re => kids.find(k => re.test(norm(k.name)) && /\.xlsx?$/i.test(k.name));
  const want = { proyecto: pick(/^proyecto porcino/), contable: pick(/sistema.?contable/) };
  const books = {}, files = [];
  for (const [key, item] of Object.entries(want)) {
    if (!item) continue;
    const url = item['@content.downloadUrl'];
    const buf = Buffer.from(await (await fetch(url)).arrayBuffer());
    books[key] = XLSX.read(buf, { type: 'buffer', cellDates: true });
    files.push({ nombre: item.name, modificado: item.lastModifiedDateTime, por: item.lastModifiedBy && item.lastModifiedBy.user ? item.lastModifiedBy.user.displayName : null });
  }
  const data = parseAll(XLSX, books);
  const payload = { fuente: 'OneDrive · Proyecto_Porcino', archivos: files, ...data };
  let prev = null;
  try { prev = JSON.parse(fs.readFileSync(OUT, 'utf8')); } catch (e) { /* primera vez */ }
  if (prev) { const { generado, ...rest } = prev; if (JSON.stringify(rest) === JSON.stringify(payload)) { console.log('Sin cambios.'); return; } }
  fs.writeFileSync(OUT, JSON.stringify({ generado: new Date().toISOString(), ...payload }, null, 1));
  console.log('data.json actualizado:', data.libro.length, 'movimientos,', data.partos.length, 'partos,', data.aportes.length, 'aportes');
  if (data.avisos.length) console.log('Avisos:', data.avisos.join(' | '));
}
main().catch(e => { console.error(e); process.exit(1); });
