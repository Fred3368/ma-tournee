import React, { useState, useEffect } from 'react';
import * as XLSX from 'xlsx';

// Clés de stockage local
const KEY_TOUR = 'maTournee_current';
const KEY_HISTORY = 'maTournee_history';

function saveTour(tour) { localStorage.setItem(KEY_TOUR, JSON.stringify(tour)); }
function loadTour() { try { return JSON.parse(localStorage.getItem(KEY_TOUR)) || []; } catch (e) { return []; } }
function saveHistory(h) { localStorage.setItem(KEY_HISTORY, JSON.stringify(h)); }
function loadHistory() { try { return JSON.parse(localStorage.getItem(KEY_HISTORY)) || []; } catch (e) { return []; } }

// Reconnaissance vocale
function createRecognizer(onResult, onError) {
  const S = window.SpeechRecognition || window.webkitSpeechRecognition;
  if (!S) return null;
  const rec = new S();
  rec.lang = 'fr-FR';
  rec.interimResults = false;
  rec.maxAlternatives = 1;
  rec.onresult = (e) => { const t = e.results[0][0].transcript; onResult(t); };
  rec.onerror = (e) => onError && onError(e);
  return rec;
}

function normalize(s) {
  if (!s) return '';
  return s.toLowerCase().normalize('NFD').replace(/\p{Diacritic}/gu, '').replace(/[.,]/g, '').trim();
}

function parseAddressFromText(text) {
  const numMatch = text.match(/\b(\d{1,4})\b/);
  const no = numMatch ? numMatch[1] : '';
  const cleaned = text.replace(/colis|pour|a|à|chez/gi, '');
  const r = cleaned.match(/(?:rue|avenue|av\.|place|boulevard|bd|chemin|impasse)?\s*(.*)/i);
  const rue = r ? r[1].trim() : cleaned.trim();
  return { no, rue };
}

export default function App() {
  const [tour, setTour] = useState(loadTour());
  const [history, setHistory] = useState(loadHistory());
  const [toast, setToast] = useState('');

  useEffect(() => { saveTour(tour); }, [tour]);
  useEffect(() => { saveHistory(history); }, [history]);

  // Fonction d'import Excel
  function handleImportExcel(e) {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      const data = new Uint8Array(evt.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet);

      const newTour = rows.map((row, i) => ({
        id: Date.now() + i,
        no: row['N°'] || row['No'] || row['NUM'] || '',
        rue: row['RUE'] || row['Adresse'] || '',
        village: row['VILLAGE'] || row['Commune'] || '',
        count: 0,
        done: false,
      }));

      setTour(newTour);
      setToast(`✅ ${newTour.length} adresses importées`);
      setTimeout(() => setToast(''), 2000);
    };
    reader.readAsArrayBuffer(file);
  }

  function handleAddMatch(addr, isNew = false) {
    setTour(prev => {
      if (isNew) return [{ ...addr, count: 1, done: false, id: addr.id || Date.now() }, ...prev];
      return prev.map(a => a.id === addr.id ? { ...a, count: (a.count || 0) + 1 } : a);
    });
    setToast('📦 Colis ajouté');
    setTimeout(() => setToast(''), 1200);
  }

  function handleDeliver(id) {
    const updated = tour.map(a => a.id === id ? { ...a, done: true } : a);
    setTour(updated);
    const delivered = updated.find(a => a.id === id);
    if (delivered) {
      setHistory(h => [{ ...delivered, when: Date.now() }, ...h]);
    }
  }

  function startDictation() {
    const rec = createRecognizer((text) => {
      const parsed = parseAddressFromText(text);
      const norm = normalize((parsed.rue || '') + ' ' + (parsed.no || ''));
      let matched = null;
      for (const a of tour) {
        const cand = normalize(a.rue + ' ' + String(a.no));
        if (cand.includes(norm) || norm.includes(cand) || (parsed.no && String(a.no) === parsed.no)) {
          matched = a; break;
        }
      }
      if (matched) handleAddMatch(matched);
      else if (window.confirm(`Adresse non trouvée: "${text}"\nVoulez-vous l’ajouter ?`)) {
        const newAddr = { id: Date.now(), no: parsed.no || '', rue: parsed.rue || text, village: '', count: 1, done: false };
        handleAddMatch(newAddr, true);
      }
    }, (e) => alert('Erreur reconnaissance : ' + (e.error || e.message || '')));
    if (!rec) return alert('Reconnaissance vocale non disponible sur cet appareil.');
    rec.start();
  }

  return (
    <div style={{ maxWidth: 420, margin: '0 auto', padding: 18, fontFamily: 'Inter, sans-serif' }}>
      <h2>🚚 Ma Tournée</h2>
      <p>Importe ton fichier Excel (colonnes : N°, RUE, VILLAGE) puis dicte les colis.</p>

      <div style={{ display: 'flex', gap: 10, marginBottom: 12 }}>
        <input type="file" accept=".xlsx,.xls" onChange={handleImportExcel} />
        <button onClick={startDictation}>🎙️ Dicter</button>
      </div>

      <div style={{ border: '1px solid #ddd', borderRadius: 10, padding: 10, marginBottom: 12 }}>
        <h3>📋 Liste</h3>
        {tour.length === 0 && <div>Aucune adresse importée</div>}
        {tour.map(a => (
          <div key={a.id} style={{ marginBottom: 6, padding: 8, background: '#f8f8f8', borderRadius: 8, opacity: a.done ? 0.5 : 1 }}>
            <div><b>{a.no}</b> - {a.rue} {a.village ? '– ' + a.village : ''}</div>
            <div>📦 {a.count} colis</div>
            {!a.done && <button onClick={() => handleDeliver(a.id)}>✅ Livré</button>}
          </div>
        ))}
      </div>

      <div style={{ border: '1px solid #ddd', borderRadius: 10, padding: 10 }}>
        <h3>📜 Historique</h3>
        {history.length === 0 && <div>Aucun colis livré</div>}
        {history.map((h, i) => (
          <div key={i} style={{ marginBottom: 6, padding: 8, background: '#e0e0e0', borderRadius: 8 }}>
            ✔ {h.no} - {h.rue} {h.village ? '– ' + h.village : ''} • {h.count} colis • {new Date(h.when).toLocaleTimeString()}
          </div>
        ))}
      </div>

      {toast && (
        <div style={{ position: 'fixed', bottom: 20, left: '50%', transform: 'translateX(-50%)', background: '#333', color: '#fff', padding: '8px 12px', borderRadius: 8 }}>
          {toast}
        </div>
      )}
    </div>
  );
}