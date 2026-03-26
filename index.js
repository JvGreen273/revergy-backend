const express = require('express');
const path = require('path');
const XLSX = require('xlsx');
const app = express();
const port = process.env.PORT || 3000;

// CORS mejorado para GitHub Pages
app.use((req, res, next) => {
  res.header('Access-Control-Allow-Origin', '*');
  res.header('Access-Control-Allow-Headers', 'Content-Type, Authorization');
  res.header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  if (req.method === 'OPTIONS') return res.sendStatus(200);
  next();
});

app.use(express.json());
app.use(express.static(__dirname));

// Caché
let cacheResumen = null;

function getResumenData() {
  if (cacheResumen) {
    console.log("✅ Datos desde caché");
    return cacheResumen;
  }
  console.log("📂 Leyendo Excel...");
  const workbook = XLSX.readFile(path.join(__dirname, 'Resumen.xlsx'), {
    cellFormula: true,
    cellNF: true,
    cellDates: true,
    sheetStubs: true
  });
  const sheet = workbook.Sheets['Resumen'];
  const data = XLSX.utils.sheet_to_json(sheet, { raw: true, defval: null, range: 2 });

  const resumen = {};
  data.forEach(row => {
    resumen[row.Indicador] = row.Valor;
  });

  cacheResumen = resumen;
  return cacheResumen;
}

// Rutas
app.get('/reload', (req, res) => {
  cacheResumen = null;
  res.json({ ok: true, mensaje: "Caché limpiada" });
});

app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'index.html'));
});

// === RUTA DEL CHAT (la que usa el frontend) ===
app.post('/chat', (req, res) => {
  const { message, project } = req.body || {};

  if (!message || !project) {
    return res.status(400).json({ reply: "Faltan datos en la solicitud." });
  }

  let reply = "Hmm, no encontré información específica para esa pregunta 🤔";

  if (project === "solar_sur" || project === "eolico_este" || project === "pv_melgar") {
    return res.json({ reply: "Este proyecto aún no tiene datos cargados en el copiloto. ¡Pronto estará disponible! 🚀" });
  }

  if (project === "puerta_de_oro") {
    let r = {};
    try {
      r = getResumenData();
    } catch (e) {
      console.error("Error Excel:", e.message);
      return res.json({ reply: "Tuve un problema al leer los datos del proyecto 😕 Intenta más tarde." });
    }

    // ... (todo tu lógica de respuestas se mantiene igual)
    const msg = message.toLowerCase();

    if (msg.includes("avance")) { /* tu código */ }
    else if (msg.includes("hinca")) { /* tu código */ }
    // ... (todas las demás condiciones se mantienen exactamente igual)

    // Si ninguna coincidió:
    reply = "📌 Entendido. ¿Puedes ser más específico? Puedo ayudarte con: avance, hincas, trackers, módulos, facturación, mano de obra, generación, dossier, rendimientos, etc.";
  }

  res.json({ reply });
});

app.listen(port, () => {
  console.log(`🚀 Servidor corriendo en puerto ${port}`);
  console.log(`Backend: https://revergy-backend.onrender.com`);
});
