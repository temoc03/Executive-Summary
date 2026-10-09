/**
 * sync.js
 *
 * Revisa una carpeta de Google Drive (y todas sus subcarpetas, una por
 * proyecto) y carga solas al dashboard las facturas EVX nuevas, igual que si
 * se hubieran subido desde el panel ⚙ "Subir facturas":
 *
 *   - Factura nueva y limpia -> dashboardContent/upload_<Proyecto>_<número>
 *     (kind 'invoiceUpload', source 'drive') con sus renglones tal como vienen
 *     en el archivo. El dashboard la aplica en el siguiente snapshot y le pone
 *     a cada renglón su fila de la pestaña con las mismas reglas que el panel.
 *   - Algo raro (nota de crédito, sin número o fecha, proyecto que no se
 *     reconoce, de otro año, total que no cuadra con los renglones) ->
 *     pendingInvoices/<fileId> con el motivo; el dashboard lo muestra en el
 *     aviso amarillo para que se revise a mano.
 *   - Factura que ya está en el dashboard (lista base o ya subida) -> nada.
 *
 * La primera vez (sin estado guardado) solo carga las facturas EVX-INV que
 * falten en el dashboard; todo lo demás que ya estaba en la carpeta se da por
 * visto, para no llenar el aviso con notas de crédito viejas.
 *
 * Autenticación: Application Default Credentials. En GitHub Actions las pone
 * el paso "google-github-actions/auth" con Workload Identity Federation (sin
 * llaves: la organización bloquea las llaves JSON), entrando como la cuenta de
 * servicio drive-sync@evolve-dashboard-dcd20.iam.gserviceaccount.com. Necesita:
 *   - la carpeta de Drive compartida con su correo (Lector), y
 *   - el rol "Cloud Datastore User" en el proyecto de Firebase.
 *
 * Variables de entorno:
 *   DRIVE_ROOT_FOLDER_ID   -- ID de la carpeta raíz de Drive a vigilar (requerida)
 *   FIRESTORE_PROJECT_ID   -- por default 'evolve-dashboard-dcd20'
 *   DASHBOARD_YEAR         -- por default 2026 (el mismo window.DASHBOARD_YEAR del dashboard)
 *   DRY_RUN                -- '1' = solo muestra lo que haría, no escribe nada
 */

const { google } = require('googleapis');
const { Firestore, FieldValue } = require('@google-cloud/firestore');
const JSZip = require('jszip');

const ROOT_FOLDER_ID = process.env.DRIVE_ROOT_FOLDER_ID;
const PROJECT_ID = process.env.FIRESTORE_PROJECT_ID || 'evolve-dashboard-dcd20';
const YEAR = parseInt(process.env.DASHBOARD_YEAR || '2026', 10);
const DRY_RUN = process.env.DRY_RUN === '1' || process.env.DRY_RUN === 'true';

// Same projects as window.PROJECT_REGISTRY in the dashboard; the ones added
// from "Proyectos y anticipos" are read from Firestore (project_* docs).
const BASE_PROJECTS = {
  Candelaria: 'Candelaria', CandelariaRC: 'Candelaria RC', Rochester: 'Rochester',
  Kougarok: 'Kougarok', Cuprite: 'Cuprite', Mojave: 'Mojave', GoldenGate: 'Golden Gate', Apex: 'Apex',
  MajubaHill: 'Majuba Hill', Palmarejo: 'Palmarejo',
};

function round2(n) { return Math.round((n + Number.EPSILON) * 100) / 100; }
function norm(s) { return String(s || '').toLowerCase().replace(/&[a-z]+;/g, ' ').replace(/[^a-z0-9]/g, ''); }
function num(s) {
  if (s == null) return null;
  const t = String(s).trim(); if (!/\d/.test(t)) return null;
  const v = parseFloat(t.replace(/[^0-9.\-]/g, ''));
  return isNaN(v) ? null : v;
}
function money(n) { return (n < 0 ? '-$' : '$') + Math.abs(n).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }); }

// ---------------- xlsx reading (same cells as the dashboard's parser) ----------------
function unxml(s) { return s.replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'"); }
async function readXlsxCells(buffer) {
  const zip = await JSZip.loadAsync(buffer);
  const shared = [];
  const sstFile = zip.file('xl/sharedStrings.xml');
  if (sstFile) {
    const sst = await sstFile.async('string');
    sst.replace(/<si>([\s\S]*?)<\/si>/g, (_, inner) => { let s = ''; inner.replace(/<t[^>]*>([\s\S]*?)<\/t>/g, (__, t) => { s += t; }); shared.push(unxml(s)); });
  }
  let sheetFile = zip.file('xl/worksheets/sheet1.xml');
  if (!sheetFile) sheetFile = zip.file(/^xl\/worksheets\/sheet\d+\.xml$/)[0];
  if (!sheetFile) throw new Error('El archivo no tiene hoja de datos.');
  const sheet = await sheetFile.async('string');
  const cells = {};
  sheet.replace(/<c r="([A-Z]+\d+)"([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g, (_, ref, attrs, inner) => {
    if (!inner) return;
    const t = (attrs.match(/t="([^"]+)"/) || [])[1];
    let v;
    if (t === 'inlineStr') v = (inner.match(/<t[^>]*>([\s\S]*?)<\/t>/) || [])[1] || '';
    else { const m = inner.match(/<v>([\s\S]*?)<\/v>/); if (!m) return; v = t === 's' ? (shared[parseInt(m[1], 10)] || '') : m[1]; }
    cells[ref] = unxml(String(v));
  });
  return cells;
}
function isoDate(text) {
  const s = String(text || '').trim();
  let m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return m[0];
  if (/^\d+(\.\d+)?$/.test(s)) {
    const serial = parseFloat(s);
    if (serial > 40000 && serial < 60000) return new Date(Date.UTC(1899, 11, 30) + Math.floor(serial) * 86400000).toISOString().slice(0, 10);
    return null;
  }
  const d = new Date(s);
  if (!isNaN(d.getTime())) return d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2) + '-' + ('0' + d.getDate()).slice(-2);
  return null;
}
function parseInvoice(cells) {
  const docNo = (cells.I4 || '').trim();
  const numMatch = docNo.match(/(\d+)\s*$/);
  const inv = {
    docNo, n: numMatch ? '#' + parseInt(numMatch[1], 10) : null,
    isCredit: /-CN-/i.test(docNo), date: isoDate(cells.I5), projectRaw: (cells.C8 || '').trim(),
    lines: [], subtotal: null, tax: 0, total: null,
  };
  for (let r = 12; r <= 90; r++) {
    const a = cells['A' + r], j = num(cells['J' + r]);
    if (a && j != null && Math.abs(j) >= 0.005 && !/^description$/i.test(a.trim())) {
      inv.lines.push({ desc: a.replace(/\s+/g, ' ').trim(), qty: (cells['F' + r] || '').trim(), rate: num(cells['H' + r]), amount: round2(j) });
    }
    const g = (cells['G' + r] || '').trim();
    if (/^subtotal/i.test(g)) inv.subtotal = num(cells['I' + r]);
    else if (/^tax/i.test(g)) inv.tax = round2(num(cells['I' + r]) || 0);
    else if (/^total/i.test(g)) inv.total = num(cells['I' + r]);
  }
  return inv;
}
// Same split as the dashboard's invoiceTotals(): the applied advance doesn't
// add to what's billed, an invoiced advance is informative only.
function lineKind(desc) {
  if (/advance payment applied/i.test(desc)) return 'advance';
  if (/^advance payment\b|deposit for future services/i.test(desc)) return 'advanceInvoice';
  return 'concept';
}
function invoiceTotals(inv) {
  let gross = inv.tax || 0, adv = 0, advInv = 0;
  inv.lines.forEach((l) => {
    const k = lineKind(l.desc);
    if (k === 'advance') adv += -l.amount;
    else if (k === 'advanceInvoice') advInv += l.amount;
    else gross += l.amount;
  });
  return { gross: round2(gross), advance: round2(adv), advanceInvoice: round2(advInv), net: round2(gross - adv) };
}
function conceptOf(inv) {
  const drill = inv.lines.some((l) => /(Meters|Feet|Footage)\s+[\d.,]+\s*-\s*[\d.,]+/i.test(l.desc));
  if (drill) return 'Drilled Mts & Consumables';
  const first = inv.lines.filter((l) => lineKind(l.desc) === 'concept')[0];
  return ((first && first.desc) || 'Invoice').slice(0, 80);
}

// ---------------- Drive ----------------
const GSHEET = 'application/vnd.google-apps.spreadsheet';
const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
async function listFilesRecursive(drive, rootId) {
  const out = [], seen = { folders: 0, other: 0 };
  const queue = [{ id: rootId, path: [] }];
  while (queue.length) {
    const { id: fid, path } = queue.shift();
    let pageToken;
    do {
      const resp = await drive.files.list({
        q: `'${fid}' in parents and trashed = false`,
        fields: 'nextPageToken, files(id, name, mimeType)',
        pageSize: 1000, pageToken, supportsAllDrives: true, includeItemsFromAllDrives: true,
      });
      for (const f of resp.data.files || []) {
        if (f.mimeType === 'application/vnd.google-apps.folder') { seen.folders++; queue.push({ id: f.id, path: path.concat(f.name) }); }
        else if (/\.xlsx$/i.test(f.name) && !/^~\$/.test(f.name)) out.push({ id: f.id, name: f.name, path });
        // An .xlsx converted to Google Sheets on upload: read it exported back to .xlsx.
        else if (f.mimeType === GSHEET && /^EVX-/i.test(f.name)) out.push({ id: f.id, name: f.name, path, gsheet: true });
        else seen.other++;
      }
      pageToken = resp.data.nextPageToken;
    } while (pageToken);
  }
  console.log(`${seen.folders} subcarpeta(s), ${out.length} factura(s)/.xlsx, ${seen.other} archivo(s) de otro tipo.`);
  return out;
}
async function download(drive, file) {
  const resp = file.gsheet
    ? await drive.files.export({ fileId: file.id, mimeType: XLSX_MIME }, { responseType: 'arraybuffer' })
    : await drive.files.get({ fileId: file.id, alt: 'media', supportsAllDrives: true }, { responseType: 'arraybuffer' });
  return Buffer.from(resp.data);
}

// ---------------- main ----------------
async function main() {
  if (!ROOT_FOLDER_ID) throw new Error('Falta la variable de entorno DRIVE_ROOT_FOLDER_ID.');
  const auth = new google.auth.GoogleAuth({ scopes: ['https://www.googleapis.com/auth/drive.readonly'] });
  const drive = google.drive({ version: 'v3', auth });
  const db = new Firestore({ projectId: PROJECT_ID, ignoreUndefinedProperties: true });
  const content = db.collection('dashboardContent');

  // Projects (base + added from the panel) and what the dashboard already has.
  const projects = Object.assign({}, BASE_PROJECTS);
  const baseNumbers = {}, baseByDateAmount = {}, uploaded = new Set();
  const all = await content.get();
  all.forEach((doc) => {
    const d = doc.data() || {};
    if (/^project_/.test(doc.id) && d.isNew && d.name) projects[doc.id.slice(8)] = d.name;
    if (/^upload_/.test(doc.id)) uploaded.add(doc.id);
    if (doc.id === 'invoiceBreakdown' && d.json) {
      const IB = JSON.parse(d.json);
      Object.keys(IB).forEach((k) => {
        const m = /^port-(.+)-invoiced$/.exec(k); if (!m) return;
        baseNumbers[m[1]] = new Set((IB[k] || []).map((r) => String(r.n).replace(/\s*\(.*\)\s*$/, '').trim()));
        // date|amount of each base invoice: catches the same invoice under another number.
        const sums = {};
        (IB[k] || []).forEach((r) => { const n = String(r.n); if (/\(/.test(n)) return; const s = sums[n] || (sums[n] = { d: r.d, a: 0 }); s.a += Number(r.a) || 0; });
        baseByDateAmount[m[1]] = {};
        Object.keys(sums).forEach((n) => { baseByDateAmount[m[1]][sums[n].d + '|' + round2(sums[n].a)] = n; });
      });
    }
  });
  const byFolder = {}; Object.keys(projects).forEach((k) => { byFolder[norm(projects[k])] = k; });
  function projectOf(file, inv) {
    for (let i = file.path.length - 1; i >= 0; i--) { const k = byFolder[norm(file.path[i])]; if (k) return k; }
    // Not in a project folder: by the client / project text of the invoice (C8), longest name wins.
    const t = String(inv.projectRaw || '').toLowerCase(); let best = null;
    Object.keys(projects).forEach((k) => { if (t.indexOf(projects[k].toLowerCase()) !== -1 && (!best || projects[k].length > projects[best].length)) best = k; });
    return best;
  }

  // processedFileIds (v2). The old detect-only version kept knownFileIds, but it
  // never loaded anything, so v2 starts over: its first run loads what's missing.
  const stateRef = db.collection('driveSync').doc('state');
  const stateSnap = await stateRef.get();
  const knownList = (stateSnap.exists && stateSnap.data().processedFileIds) || [];
  const known = new Set(knownList);
  const bootstrap = knownList.length === 0;
  // The old version's notices ("detected, not captured", no reason) are replaced by this one.
  const oldNotices = [];
  if (bootstrap) {
    const pend = await db.collection('pendingInvoices').where('processed', '==', false).get();
    pend.forEach((doc) => { if (!doc.data().reason) oldNotices.push(doc.ref); });
  }

  console.log(`Revisando carpeta de Drive${DRY_RUN ? ' (DRY_RUN: no se escribe nada)' : ''}${bootstrap ? ' (primera vez)' : ''}...`);
  // The root folder first: a clear message if it doesn't exist or isn't shared with this account.
  try {
    const root = await drive.files.get({ fileId: ROOT_FOLDER_ID, fields: 'id, name, mimeType', supportsAllDrives: true });
    console.log(`Carpeta raíz: "${root.data.name}"`);
  } catch (err) {
    const who = (await auth.getCredentials().catch(() => ({}))).client_email || 'la cuenta de servicio';
    throw new Error(`No se puede abrir la carpeta del secret DRIVE_ROOT_FOLDER_ID (${err.code || err.message}). ` +
      `Revisa que el ID sea el de la carpeta Projects y que esté compartida con ${who} como Lector.`);
  }
  const files = await listFilesRecursive(drive, ROOT_FOLDER_ID);
  const newFiles = files.filter((f) => !known.has(f.id));
  console.log(`${files.length} archivo(s) .xlsx en total, ${newFiles.length} sin revisar.`);

  const writes = [], loaded = [], pending = [];
  for (const f of newFiles) {
    const folder = f.path.join(' / ');
    const markPending = (reason) => {
      if (bootstrap) return; // first run: what was already in the folder counts as seen
      pending.push(`${folder} / ${f.name}: ${reason}`);
      writes.push(['set', db.collection('pendingInvoices').doc(f.id), {
        fileId: f.id, fileName: f.name, projectFolder: f.path[f.path.length - 1] || null,
        driveLink: `https://drive.google.com/file/d/${f.id}/view`, detectedAt: new Date().toISOString(), processed: false, reason,
      }]);
    };
    known.add(f.id);
    if (!/^EVX-/i.test(f.name)) continue; // trackers, consolidated files...: not invoices
    let inv;
    try { inv = parseInvoice(await readXlsxCells(await download(drive, f))); }
    catch (err) { markPending('no se pudo leer: ' + err.message); continue; }
    if (inv.isCredit) { markPending('nota de crédito: aplícala en ⚙ Editar facturas'); continue; }
    if (!inv.n) { markPending('no se pudo leer el número de factura (I4)'); continue; }
    if (inv.lines.length && inv.lines.every((l) => lineKind(l.desc) === 'advanceInvoice')) {
      markPending('factura de anticipo: regístrala en ⚙ Proyectos y anticipos'); continue;
    }
    const p = projectOf(f, inv);
    if (!p) { markPending('no se reconoció el proyecto (carpeta "' + (f.path[f.path.length - 1] || '') + '")'); continue; }
    const docId = 'upload_' + p + '_' + String(inv.n).replace(/[^0-9A-Za-z]/g, '');
    if ((baseNumbers[p] && baseNumbers[p].has(inv.n)) || uploaded.has(docId)) continue; // already in the dashboard
    if (!inv.date) { markPending('no se pudo leer la fecha (I5)'); continue; }
    if (parseInt(inv.date.slice(0, 4), 10) !== YEAR) { markPending('es de ' + inv.date.slice(0, 4) + ' y el dashboard es de ' + YEAR); continue; }
    const t = invoiceTotals(inv);
    const twin = baseByDateAmount[p] && baseByDateAmount[p][inv.date + '|' + t.gross];
    if (twin) { markPending('parece la misma factura que la ' + twin + ' del dashboard (misma fecha y monto)'); continue; }
    if (inv.total != null && Math.abs(t.net - inv.total) > 0.05) {
      markPending('los renglones suman ' + money(t.net) + ' y la factura dice ' + money(inv.total));
      continue;
    }
    uploaded.add(docId);
    loaded.push(`${projects[p]} ${inv.n} (${inv.date}) ${money(t.net)}`);
    writes.push(['set', content.doc(docId), {
      kind: 'invoiceUpload', source: 'drive', project: p, n: inv.n, date: inv.date, concept: conceptOf(inv),
      gross: t.gross, advance: t.advance, advanceInvoice: t.advanceInvoice, tax: inv.tax || 0, net: t.net, printedTotal: inv.total,
      lines: inv.lines.map((l) => ({ desc: l.desc, qty: l.qty, rate: l.rate, amount: l.amount })),
      fileName: f.name, driveId: f.id, uploadedBy: 'drive-sync', uploadedAt: FieldValue.serverTimestamp(),
    }]);
  }

  loaded.forEach((s) => console.log('  cargada: ' + s));
  pending.forEach((s) => console.log('  para revisar: ' + s));
  console.log(`${loaded.length} factura(s) cargada(s), ${pending.length} para revisar.`);
  if (DRY_RUN) { console.log('DRY_RUN: no se escribió nada.'); return; }

  if (oldNotices.length) console.log(`${oldNotices.length} aviso(s) de la versión anterior se dan por atendidos.`);
  oldNotices.forEach((ref) => writes.push(['set', ref, { processed: true }, { merge: true }]));
  writes.push(['set', stateRef, {
    processedFileIds: Array.from(known), lastRunAt: new Date().toISOString(), lastRunFileCount: files.length,
    lastLoaded: loaded.length, lastPending: pending.length,
  }, { merge: true }]);
  // Firestore batches take up to 500 writes.
  for (let i = 0; i < writes.length; i += 400) {
    const batch = db.batch();
    writes.slice(i, i + 400).forEach(([, ref, data, opts]) => (opts ? batch.set(ref, data, opts) : batch.set(ref, data)));
    await batch.commit();
  }
  console.log('Listo.');
}

main().catch((err) => {
  console.error('Error en sync.js:', err);
  process.exit(1);
});
