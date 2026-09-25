/**
 * sync.js
 *
 * Revisa una carpeta de Google Drive (y todas sus subcarpetas) buscando
 * archivos .xlsx nuevos que no se han visto antes, y los agrega a la
 * colección "pendingInvoices" de Firestore para que alguien los revise y
 * capture manualmente -- este script NUNCA modifica los totales del
 * dashboard por sí solo, solo detecta y avisa.
 *
 * Autenticación: usa Application Default Credentials (ADC). En GitHub
 * Actions, la acción "google-github-actions/auth" genera esas credenciales
 * automáticamente vía Workload Identity Federation (sin llaves ni
 * contraseñas guardadas en el repo). Localmente, puedes correr
 * `gcloud auth application-default login` para probarlo con tu propia cuenta.
 *
 * Variables de entorno requeridas:
 *   DRIVE_ROOT_FOLDER_ID   -- ID de la carpeta raíz de Drive a vigilar
 *   FIRESTORE_PROJECT_ID   -- (opcional) por default 'evolve-dashboard-dcd20'
 */

const { google } = require('googleapis');
const { Firestore } = require('@google-cloud/firestore');

const ROOT_FOLDER_ID = process.env.DRIVE_ROOT_FOLDER_ID;
const PROJECT_ID = process.env.FIRESTORE_PROJECT_ID || 'evolve-dashboard-dcd20';

// Filtra archivos que no son facturas reales: archivos de bloqueo temporal
// de Excel ("~$..."), consolidados/resúmenes armados a mano, y trackers de
// facturación -- ninguno de esos es una factura individual que alguien deba
// "capturar".
function isLikelyInvoiceFile(name) {
  if (/^~\$/.test(name)) return false;
  if (/consolidad[oa]/i.test(name)) return false;
  if (/tracker[_ ]?facturaci[oó]n/i.test(name)) return false;
  return true;
}

async function listFilesRecursive(drive, rootId) {
  const out = [];
  const queue = [{ id: rootId, name: null }];
  while (queue.length) {
    const { id: fid, name: parentName } = queue.shift();
    let pageToken;
    do {
      const resp = await drive.files.list({
        q: `'${fid}' in parents and trashed = false`,
        fields: 'nextPageToken, files(id, name, mimeType)',
        pageSize: 1000,
        pageToken,
      });
      for (const f of resp.data.files || []) {
        if (f.mimeType === 'application/vnd.google-apps.folder') {
          queue.push({ id: f.id, name: f.name });
        } else if (/\.xlsx$/i.test(f.name) && isLikelyInvoiceFile(f.name)) {
          out.push({ id: f.id, name: f.name, parentFolderName: parentName });
        }
      }
      pageToken = resp.data.nextPageToken;
    } while (pageToken);
  }
  return out;
}

async function main() {
  if (!ROOT_FOLDER_ID) {
    throw new Error('Falta la variable de entorno DRIVE_ROOT_FOLDER_ID.');
  }

  const auth = new google.auth.GoogleAuth({
    scopes: ['https://www.googleapis.com/auth/drive.readonly'],
  });
  const drive = google.drive({ version: 'v3', auth });
  const firestore = new Firestore({ projectId: PROJECT_ID });

  const stateRef = firestore.collection('driveSync').doc('state');
  const stateSnap = await stateRef.get();
  const known = new Set((stateSnap.exists && stateSnap.data().knownFileIds) || []);

  console.log(`Revisando carpeta de Drive ${ROOT_FOLDER_ID}...`);
  const files = await listFilesRecursive(drive, ROOT_FOLDER_ID);
  const newFiles = files.filter((f) => !known.has(f.id));

  console.log(`Encontrados ${files.length} archivo(s) .xlsx en total, ${newFiles.length} nuevo(s).`);

  if (newFiles.length === 0) {
    await stateRef.set(
      { lastRunAt: new Date().toISOString(), lastRunFileCount: files.length },
      { merge: true }
    );
    console.log('Nada nuevo. Listo.');
    return;
  }

  const batch = firestore.batch();
  for (const f of newFiles) {
    const ref = firestore.collection('pendingInvoices').doc(f.id);
    batch.set(ref, {
      fileId: f.id,
      fileName: f.name,
      projectFolder: f.parentFolderName || null,
      driveLink: `https://drive.google.com/file/d/${f.id}/view`,
      detectedAt: new Date().toISOString(),
      processed: false,
    });
  }

  const allKnown = Array.from(new Set([...known, ...newFiles.map((f) => f.id)]));
  batch.set(
    stateRef,
    {
      knownFileIds: allKnown,
      lastRunAt: new Date().toISOString(),
      lastRunFileCount: files.length,
      lastNewCount: newFiles.length,
    },
    { merge: true }
  );

  await batch.commit();

  console.log(`Se agregaron ${newFiles.length} factura(s) nueva(s) a "pendingInvoices":`);
  newFiles.forEach((f) => console.log(`  - ${f.parentFolderName || '(raíz)'} / ${f.name}`));
}

main().catch((err) => {
  console.error('Error en sync.js:', err);
  process.exit(1);
});
