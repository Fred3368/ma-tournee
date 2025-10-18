import React, { useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";
import "./App.css";

function App() {
  const [allAddresses, setAllAddresses] = useState([]); // toutes les adresses Excel
  const [stops, setStops] = useState([]); // uniquement les arrêts avec colis
  const [listening, setListening] = useState(false);
  const recognitionRef = useRef(null);

  // --------- Charger les arrêts sauvegardés au démarrage ----------
  useEffect(() => {
    const savedStops = localStorage.getItem("stops");
    if (savedStops) setStops(JSON.parse(savedStops));
  }, []);

  // --------- Sauvegarde automatique ----------
  useEffect(() => {
    localStorage.setItem("stops", JSON.stringify(stops));
  }, [stops]);

  // --------- Import fichier Excel ----------
  const handleFileUpload = (event) => {
    const file = event.target.files[0];
    const reader = new FileReader();
    reader.onload = (e) => {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: "array" });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const json = XLSX.utils.sheet_to_json(sheet);

      const addresses = json.map((row, i) => ({
        no: row["N°"] || i + 1,
        rue: row["RUE"] || "",
        village: row["VILLAGE"] || "",
      }));

      setAllAddresses(addresses);
      setStops([]); // vide la tournée
    };
    reader.readAsArrayBuffer(file);
  };

  // --------- Ajouter un colis à une adresse ----------
  const addColis = (rueParlee) => {
    const match = allAddresses.find((a) =>
      a.rue.toLowerCase().includes(rueParlee.toLowerCase())
    );
    if (!match) {
      console.log("Adresse non trouvée pour :", rueParlee);
      return;
    }

    setStops((prevStops) => {
      const existing = prevStops.find((s) => s.no === match.no);
      if (existing) {
        return prevStops.map((s) =>
          s.no === match.no ? { ...s, count: s.count + 1 } : s
        );
      } else {
        return [...prevStops, { ...match, count: 1, done: false }];
      }
    });
  };

  // --------- Reconnaissance vocale ----------
  useEffect(() => {
    const SpeechRecognition =
      window.SpeechRecognition || window.webkitSpeechRecognition;
    if (!SpeechRecognition) {
      alert("La reconnaissance vocale n'est pas supportée sur ce navigateur.");
      return;
    }

    const recognition = new SpeechRecognition();
    recognition.lang = "fr-FR";
    recognition.continuous = true;
    recognition.interimResults = false;

    recognition.onresult = (event) => {
      const transcript =
        event.results[event.results.length - 1][0].transcript.trim();
      console.log("Vous avez dit :", transcript);
      addColis(transcript);
    };

    recognition.onend = () => {
      if (listening) recognition.start(); // relance automatique
    };

    recognitionRef.current = recognition;
  }, [allAddresses, listening]);

  const toggleListening = () => {
    if (!recognitionRef.current) return;

    if (!listening) {
      recognitionRef.current.start();
      setListening(true);
    } else {
      recognitionRef.current.stop();
      setListening(false);
    }
  };

  // --------- Fonctions utilitaires ----------
  const markDone = (no) => {
    setStops((prevStops) =>
      prevStops.map((s) => (s.no === no ? { ...s, done: !s.done } : s))
    );
  };

  const clearStops = () => {
    if (window.confirm("Effacer toute la tournée sauvegardée ?")) {
      setStops([]);
      localStorage.removeItem("stops");
    }
  };

  // --------- Ouvrir la carte Google Maps ----------
  const openMap = () => {
    if (stops.length === 0) return alert("Aucun arrêt à afficher.");
    const query = stops
      .map((s) => `${s.rue} ${s.village}`)
      .join(" -> ");
    const url = `https://www.google.com/maps/dir/${encodeURIComponent(query)}`;
    window.open(url, "_blank");
  };

  // --------- Rendu visuel ----------
  return (
    <div className="App">
      <header>
        <h1>🚚 Ma Tournée (déploiement automatique OK)</h1>
        <p>Importe ta tournée, dicte tes arrêts et reprends là où tu t’étais arrêté.</p>
      </header>

      <div className="import">
        <input type="file" accept=".xlsx" onChange={handleFileUpload} />
        <button onClick={clearStops}>🧹 Réinitialiser</button>
        <button onClick={toggleListening}>
          {listening ? "🛑 Arrêter le micro" : "🎙️ Démarrer le micro"}
        </button>
        <button onClick={openMap}>📍 Voir sur la carte</button>
      </div>

      <h2>🧾 Points de livraison</h2>
      {stops.length === 0 ? (
        <p>Aucun arrêt pour le moment.</p>
      ) : (
        <ul>
          {stops
            .sort((a, b) => a.no - b.no)
            .map((s) => (
              <li
                key={s.no}
                onClick={() => markDone(s.no)}
                style={{
                  textDecoration: s.done ? "line-through" : "none",
                  opacity: s.done ? 0.6 : 1,
                }}
              >
                {s.no} - {s.rue}, {s.village} • 📦 {s.count} colis
              </li>
            ))}
        </ul>
      )}
    </div>
  );
}

export default App;