const express = require('express');
const path = require('path');
const XLSX = require('xlsx');

const app = express();
const port = 3000;

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

  // ✅ Nombre del archivo actualizado
  const workbook = XLSX.readFile(path.join(__dirname, 'Resumen.xlsx'), {
    cellFormula: true,
    cellNF: true,
    cellDates: true,
    sheetStubs: true
  });

  const sheet = workbook.Sheets['Resumen'];

  // ✅ range: 2 le indica que el encabezado está en la fila 3 (índice 2)
  const data = XLSX.utils.sheet_to_json(sheet, {
    raw: true,
    defval: null,
    range: 2
  });

  // DEBUG: muestra indicadores y valores (solo la primera vez)
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
   Visita http://localhost:3000/reload
   cuando actualices el Excel
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

  let reply = "No tengo información suficiente para responder esa pregunta. Puedes preguntarme por avance, hincas, trackers, módulos, facturación o mano de obra.";

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
      const real   = formatPct(r["Avance Real"], false);  // Ya está en %
      const plan   = formatPct(r["Avance Plan"], false);  // Ya está en %
      const spi    = toNum(r["SPI"]);
      const spiStr = isNaN(spi) ? "N/D" : spi.toFixed(3);
      const estado = !isNaN(spi) && spi >= 1 ? "adelantado respecto al plan 🟢" : "por detrás del plan 🔴";
      reply = `📊 Avance real: ${real}% | Plan: ${plan}%\nSPI: ${spiStr} → ${estado}`;
    }

    // --- HINCAS ---
    else if (msg.includes("hinca")) {
      const inst  = toNum(r["Hincas_Instaladas"]);
      const total = toNum(r["Hincas_Totales"]);
      const pct   = (!isNaN(inst) && !isNaN(total)) ? (inst / total * 100).toFixed(2) : "N/D";
      reply = `🔩 Hincas: ${inst.toLocaleString('es-CO')} instaladas de ${total.toLocaleString('es-CO')} (${pct}%)`;
    }

    // --- TRACKERS ---
    else if (msg.includes("tracker")) {
      const inst  = toNum(r["Trackers_Instalados"]);
      const total = toNum(r["Trackers_Totales"]);
      const pct   = (!isNaN(inst) && !isNaN(total)) ? (inst / total * 100).toFixed(2) : "N/D";
      reply = `☀️ Trackers: ${inst.toLocaleString('es-CO')} instalados de ${total.toLocaleString('es-CO')} (${pct}%)`;
    }

    // --- MÓDULOS ---
    else if (msg.includes("modulo") || msg.includes("módulo") || msg.includes("panel")) {
      const inst  = toNum(r["Modulos_Instalados"]);
      const total = toNum(r["Modulos_Totales"]);
      const pct   = (!isNaN(inst) && !isNaN(total)) ? (inst / total * 100).toFixed(2) : "N/D";
      reply = `🔋 Módulos: ${inst.toLocaleString('es-CO')} instalados de ${total.toLocaleString('es-CO')} (${pct}%)`;
    }

    // --- FACTURACIÓN ---
    else if (msg.includes("facturac")) {
      const actual = formatCOP(r["Facturación Actual"]);
      const plan   = formatCOP(r["Facturación_Plan"]);
      const pct    = formatPct(r["Facturacion_Porcentaje"], false);
      reply = `💰 Facturación actual: $${actual} COP\n📋 Plan: $${plan} COP\n📈 Avance: ${pct}%`;
    }

    // --- MANO DE OBRA ---
    else if (msg.includes("mano") || msg.includes("personal") || msg.includes("personas") || msg.includes("trabajador")) {
      const total     = toNum(r["Mano_de_Obra_Total"]);
      const directa   = toNum(r["Mano_de_Obra_Directa"]);
      const indirecta = toNum(r["Mano_de_Obra_Indirecta"]);
      reply = `👷 Personal en obra: ${total} personas\n• Directa: ${directa}\n• Indirecta: ${indirecta}`;
    }

    // --- GENERACIÓN ---
    else if (msg.includes("generacion") || msg.includes("generación")) {
      const generacion = toNum(r["Generación"]);
      const pct        = toNum(r["Porcentaje de Generación"]);
      const generacionStr = !isNaN(generacion) ? generacion.toFixed(1) : "N/D";
      const pctStr        = !isNaN(pct) ? pct.toFixed(2) : "N/D";
      reply = `⚡ La generación actual del parque es ${generacionStr} MW, lo que representa un ${pctStr}% del total del Parque.`;
    }

    // --- DOSSIER ---
    else if (msg.includes("dossier")) {
      const avanceDossier = toNum(r["Dossier"]);
      const pctStr = !isNaN(avanceDossier) ? avanceDossier.toFixed(2) : "N/D";
      reply = `📂 El avance general del Dossier es de ${pctStr}%. Para mayor detalle revisa el apartado de Gestión de Calidad, botón Dossier.`;
    }

    // --- NO CONFORMIDADES ---
    else if (msg.includes("conformidad") || msg.includes("no conformidad")) {
      const noConformidades = r["Balance de No Conformidades"];
      reply = noConformidades ? `⚠️ ${noConformidades}` : "No se encontró información sobre No Conformidades.";
    }

    // --- RESUMEN GENERAL ---
    else if (msg.includes("resumen") || msg.includes("estado") || msg.includes("general")) {
      const real   = formatPct(r["Avance Real"], false);  // Ya está en %
      const plan   = formatPct(r["Avance Plan"], false);  // Ya está en %
      const spi    = toNum(r["SPI"]);
      const spiStr = isNaN(spi) ? "N/D" : spi.toFixed(3);
      const total  = toNum(r["Mano_de_Obra_Total"]);
      const actual = formatCOP(r["Facturación Actual"]);
      reply = `📋 Resumen PFV Puerta de Oro:\n• Avance real: ${real}% (Plan: ${plan}%)\n• SPI: ${spiStr}\n• Personal: ${total} personas\n• Facturación: $${actual} COP`;
    }

    // --- RENDIMIENTOS TENDIDO CABLE SOLAR ---
    else if (msg.includes("rendimiento") || msg.includes("tendido") || msg.includes("cable")) {
      const rendimiento = toNum(r["Rendimientos Tendido cable solar"]);
      const rendStr = isNaN(rendimiento) ? "N/D" : rendimiento.toFixed(0);
      reply = `📏 Los rendimientos presentados de la última semana que muestra ejecución para tendido de cable solar son de ${rendStr} metros en la semana.\n\n📌 Para mayor detalle consúltalo en el apartado de Gestión de Construcción → Plan de Acción/Bonus → Tendido Solar`;
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