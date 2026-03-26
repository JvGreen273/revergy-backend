const express = require('express');
const path = require('path');
const XLSX = require('xlsx');

const app = express();
const port = process.env.PORT || 3000;

// CORS para GitHub Pages (mejorado)
app.use((req, res, next) => {
  res.header('Access-Control-Allow-Origin', '*');
  res.header('Access-Control-Allow-Headers', 'Content-Type');
  res.header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  if (req.method === 'OPTIONS') return res.sendStatus(200);
  next();
});

app.use(express.json());

// Caché en memoria
let cacheResumen = null;

function getResumenData() {
  if (cacheResumen) {
    console.log("✅ Datos desde caché");
    return cacheResumen;
  }
  console.log("📂 Leyendo Excel...");
  try {
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
      if (row.Indicador) resumen[row.Indicador] = row.Valor;
    });

    cacheResumen = resumen;
    return resumen;
  } catch (e) {
    console.error("Error leyendo Excel:", e.message);
    return {};
  }
}

// Rutas básicas
app.get('/reload', (req, res) => {
  cacheResumen = null;
  res.json({ ok: true, mensaje: "Caché limpiada" });
});

app.get('/', (req, res) => {
  res.send("Backend Revergy funcionando ✅");
});

// === RUTA DEL CHAT ===
app.post('/chat', (req, res) => {
  const { message, project } = req.body || {};

  if (!message || !project) {
    return res.json({ reply: "Faltan datos en la solicitud." });
  }

  // Proyectos demo sin datos
  if (project === "solar_sur" || project === "eolico_este" || project === "pv_melgar") {
    return res.json({ reply: "Este proyecto aún no tiene datos cargados en el copiloto. ¡Pronto estará disponible! 🚀" });
  }

  let reply = "Hmm, no encontré información específica para esa pregunta 🤔 Puedes preguntarme por avance, hincas, trackers, módulos, facturación, mano de obra, generación, dossier o rendimientos.";

  if (project === "puerta_de_oro") {
    const r = getResumenData();
    const msg = message.toLowerCase().trim();

    // AVANCE
    if (msg.includes("avance")) {
      const real = r["Avance Real"] ? r["Avance Real"] + "%" : "N/D";
      const plan = r["Avance Plan"] ? r["Avance Plan"] + "%" : "N/D";
      reply = `📊 Avance general:\n• Real: ${real}\n• Planificado: ${plan}`;
    }
    // HINCAS
    else if (msg.includes("hinca") || msg.includes("hincas")) {
      const inst = r["Hincas_Instaladas"] || "N/D";
      const total = r["Hincas_Totales"] || "N/D";
      reply = `🔩 Hincas:\n• Instaladas: ${inst} de ${total}`;
    }
    // TRACKERS
    else if (msg.includes("tracker") || msg.includes("trackers")) {
      const inst = r["Trackers_Instalados"] || "N/D";
      const total = r["Trackers_Totales"] || "N/D";
      reply = `☀️ Trackers:\n• Instalados: ${inst} de ${total}`;
    }
    // MÓDULOS
    else if (msg.includes("modulo") || msg.includes("módulo") || msg.includes("panel")) {
      const inst = r["Modulos_Instalados"] || "N/D";
      const total = r["Modulos_Totales"] || "N/D";
      reply = `🔋 Módulos:\n• Instalados: ${inst} de ${total}`;
    }
    // Agrega aquí más condiciones si quieres (facturación, mano de obra, etc.)

    else {
      reply = "📌 Entendido. ¿Puedes ser más específico? Puedo ayudarte con: avance, hincas, trackers, módulos, facturación, mano de obra, generación, dossier, rendimientos, etc.";
    }
  }

  res.json({ reply });
});

app.listen(port, () => {
  console.log(`🚀 Backend corriendo en puerto ${port}`);
});
