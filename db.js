// db.js - IndexedDB helper para almacenar último contrato (.docx blob) y settings como Ejecutivo

const DB_NAME = 'ContratoDB';
const STORE_NAME = 'Contratos';

function openDB() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, 1);
    req.onupgradeneeded = (ev) => {
      const db = ev.target.result;
      if (!db.objectStoreNames.contains(STORE_NAME)) db.createObjectStore(STORE_NAME);
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

async function saveContrato(blob) {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE_NAME, 'readwrite');
    const store = tx.objectStore(STORE_NAME);
    const req = store.put(blob, 'ultimoContrato');
    req.onsuccess = () => resolve(true);
    req.onerror = () => reject(req.error);
  });
}

async function getContrato() {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE_NAME, 'readonly');
    const store = tx.objectStore(STORE_NAME);
    const req = store.get('ultimoContrato');
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

// --- Nuevas utilidades para Ejecutivo (persistencia y borrado) ---
async function saveEjecutivo(name) {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE_NAME, 'readwrite');
    const store = tx.objectStore(STORE_NAME);
    const req = store.put(String(name || ''), 'ejecutivo');
    req.onsuccess = () => resolve(true);
    req.onerror = () => reject(req.error);
  });
}

async function getEjecutivo() {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE_NAME, 'readonly');
    const store = tx.objectStore(STORE_NAME);
    const req = store.get('ejecutivo');
    req.onsuccess = () => resolve(req.result || '');
    req.onerror = () => reject(req.error);
  });
}

async function deleteEjecutivo() {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE_NAME, 'readwrite');
    const store = tx.objectStore(STORE_NAME);
    const req = store.delete('ejecutivo');
    req.onsuccess = () => resolve(true);
    req.onerror = () => reject(req.error);
  });
}