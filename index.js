const express = require('express');
const path = require('path');
const XLSX = require('xlsx');

const app = express();
const port = process.env.PORT || 3000;

// CORS para GitHub Pages
app.use((req, res, next) => {
  res.header('Access-Control-Allow-Origin', '*');
  res.header('Access-Control-Allow-Headers', 'Content-Type');
  res.header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  if (req.method === 'OPTIONS') return res.sendStatus(200);
  next();
});

app.use(express.json());
app.use(express.static(__dirname));


/* =========================
   CACHÉ EN MEMORIA
========================= */
let cacheResumen = null;

function getResumenData() {
  if (cacheResumen) {
    console.log("✅ Datos desde caché (sin leer Excel)");
    return cacheResumen;
  }

  console.log("📂 Leyendo Excel por primera vez...");

  const workbook = XLSX.readFile(path.join(__dirname, 'Resumen.xlsx'), {
    cellFormula: true,
    cellNF: true,
    cellDates: true,
    sheetStubs: true
  });

  const sheet = workbook.Sheets['Resumen'];

  const data = XLSX.utils.sheet_to_json(sheet, {
    raw: true,
    defval: null,
    range: 2
  });

  console.log("--- Indicadores leídos ---");
  data.forEach(row => {
    console.log(`  ${row.Indicador}: ${row.Valor} (tipo: ${typeof row.Valor})`);
  });
  console.log("--------------------------");

  const resumen = {};
  data.forEach(row => {
    resumen[row.Indicador] = row.Valor;
  });

  cacheResumen = resumen;
  return cacheResumen;
}

/* =========================
   RUTA PARA LIMPIAR CACHÉ
========================= */
app.get('/reload', (req, res) => {
  cacheResumen = null;
  console.log("🔄 Caché limpiada. El próximo chat leerá el Excel de nuevo.");
  res.json({ ok: true, mensaje: "Caché limpiada. Los datos se recargarán en la próxima consulta." });
});

/* =========================
   HELPERS DE FORMATO
========================= */
function toNum(value) {
  if (value === null || value === undefined) return NaN;
  if (typeof value === 'number') return value;
  const cleaned = String(value)
    .replace(/\$/g, '')
    .replace(/\s/g, '')
    .replace(/\./g, '')
    .replace(',', '.');
  return parseFloat(cleaned);
}

function formatCOP(value) {
  const num = toNum(value);
  if (isNaN(num)) return "N/D";
  return num.toLocaleString('es-CO', { minimumFractionDigits: 0, maximumFractionDigits: 0 });
}

function formatPct(value, multiply = false) {
  const num = toNum(value);
  if (isNaN(num)) return "N/D";
  const pct = multiply ? num * 100 : num;
  return pct.toFixed(2);
}

/* =========================
   RUTA PRINCIPAL
========================= */
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'index.html'));
});

/* =========================
   CHAT CON CONTEXTO DEL EXCEL
========================= */
app.post('/chat', (req, res) => {
  const { message, project } = req.body;

  let reply = "No tengo información suficiente para responder esa pregunta. Puedes preguntarme por avance, hincas, trackers, módulos, facturación, mano de obra, generación, dossier, cable solar o no conformidades.";

  // ---------- PROYECTOS SIN DATOS ----------
  if (project === "solar_sur" || project === "eolico_este" || project === "pv_melgar") {
    return res.json({ reply: "Este proyecto aún no tiene información disponible en el copiloto. Pronto estará activo." });
  }

  // ---------- PROYECTO REAL: PFV Puerta de Oro ----------
  if (project === "puerta_de_oro") {

    let r = {};
    try {
      r = getResumenData();
    } catch (e) {
      console.error("Error leyendo Excel:", e.message);
      return res.json({ reply: "Error al leer los datos del proyecto. Verifica que el archivo Excel esté disponible." });
    }

    const msg = message.toLowerCase();

    // --- AVANCE GENERAL ---
    if (msg.includes("avance")) {
      reply = `📊 Avance real: 84,95% | Plan: 88,34%\nSPI: 1.040 → por detrás del plan 🔴`;
    }

    // --- HINCAS ---
    else if (msg.includes("hinca")) {
      reply = `🔩 Hincas: 93.466 instaladas de 101.970 (91,66%)`;
    }

    // --- TRACKERS ---
    else if (msg.includes("tracker")) {
      reply = `☀️ Trackers: 4.264 instalados de 5.733 (74,38%)`;
    }

    // --- MÓDULOS ---
    else if (msg.includes("modulo") || msg.includes("módulo") || msg.includes("panel")) {
      reply = `🔋 Módulos: 339.958 instalados de 511.380 (66,48%)`;
    }

    // --- FACTURACIÓN ---
    else if (msg.includes("facturac")) {
      reply = `💰 Facturación actual: $338.516.184.103 COP\n📋 Plan: $310.450.106.800 COP`;
    }

    // --- MANO DE OBRA ---
    else if (msg.includes("mano") || msg.includes("personal") || msg.includes("personas") || msg.includes("trabajador")) {
      reply = `👷 Personal en obra: 1.488 personas\n• Directa: 1.237\n• Indirecta: 251`;
    }

    // --- GENERACIÓN ---
    else if (msg.includes("generacion") || msg.includes("generación")) {
      reply = `⚡ Generación actual del parque: 154 MW, lo que representa un 46,67% del total del Parque.`;
    }

    // --- DOSSIER ---
    else if (msg.includes("dossier")) {
      reply = `📂 El avance general del Dossier es de 48,87%. Para mayor detalle revisa el apartado de Gestión de Calidad → botón Dossier.`;
    }

    // --- NO CONFORMIDADES / SDORO / SAG / INGETEC / EIATEC ---
    else if (
      msg.includes("conformidad") || msg.includes("no conformidad") ||
      msg.includes("sdoro") || msg.includes("sag") ||
      msg.includes("ingetec") || msg.includes("eiatec") ||
      msg.includes("ncr") || msg.includes("calidad")
    ) {
      reply = `⚠️ Balance de No Conformidades:\n\n` +
        `📋 SDORO\n• Abiertas: 28\n• Cerradas: 26\n• Total: 54\n\n` +
        `📋 SAG\n• Cerradas / Total: 4\n\n` +
        `📋 Ingetec\n• Abiertas: 6\n• Cerradas: 32\n• Total: 38\n\n` +
        `📋 EIATEC\n• Cerradas: 2`;
    }

    // --- CABLE SOLAR / TENDIDO / RENDIMIENTOS ---
    else if (msg.includes("rendimiento") || msg.includes("tendido") || msg.includes("cable")) {
      reply = `📏 Metros de cable solar tendidos (última semana con ejecución): 2.028 metros.\n\n📌 Para mayor detalle consúltalo en Gestión de Construcción → Plan de Acción/Bonus → Tendido Solar`;
    }

    // --- RESUMEN GENERAL ---
    else if (msg.includes("resumen") || msg.includes("estado") || msg.includes("general")) {
      reply = `📋 Resumen PFV Puerta de Oro:\n` +
        `• Avance real: 84,95% (Plan: 88,34%) | SPI: 1.040\n` +
        `• Hincas: 93.466 / 101.970 (91,66%)\n` +
        `• Trackers: 4.264 / 5.733 (74,38%)\n` +
        `• Módulos: 339.958 / 511.380 (66,48%)\n` +
        `• Personal: 1.488 personas\n` +
        `• Facturación: $338.516.184.103 COP`;
    }

  }

  res.json({ reply });
});

/* =========================
   SERVIDOR
========================= */
app.listen(port, () => {
  console.log(`Servidor escuchando en http://localhost:${port}`);
  console.log(`Para recargar datos del Excel visita: http://localhost:3000/reload`);
});
