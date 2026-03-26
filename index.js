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

  // range: 2 le indica que el encabezado está en la fila 3 (índice 2)
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

// Formatea porcentaje con exactamente los decimales del Excel (hasta 2 decimales, sin ceros innecesarios)
function formatPct(value, multiply = false) {
  const num = toNum(value);
  if (isNaN(num)) return "N/D";
  const pct = multiply ? num * 100 : num;
  // Muestra hasta 2 decimales pero elimina ceros finales
  return parseFloat(pct.toFixed(2)).toString();
}

// Formatea número entero con separador de miles colombiano
function formatInt(value) {
  const num = toNum(value);
  if (isNaN(num)) return "N/D";
  return Math.round(num).toLocaleString('es-CO');
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

  let reply = "Hmm, no encontré información específica para esa pregunta 🤔 Puedes preguntarme por avance, hincas, trackers, módulos, facturación, mano de obra, generación, dossier o rendimientos de tendido.";

  // ---------- PROYECTOS SIN DATOS ----------
  if (project === "solar_sur" || project === "eolico_este" || project === "pv_melgar") {
    return res.json({ reply: "Este proyecto aún no tiene datos cargados en el copiloto. ¡Pronto estará disponible! 🚀 Mientras tanto, puedes consultar el proyecto PFV Puerta de Oro." });
  }

  // ---------- PROYECTO REAL: PFV Puerta de Oro ----------
  if (project === "puerta_de_oro") {

    let r = {};
    try {
      r = getResumenData();
    } catch (e) {
      console.error("Error leyendo Excel:", e.message);
      return res.json({ reply: "Tuve un problema al leer los datos del proyecto 😕 Por favor verifica que el archivo Resumen.xlsx esté disponible en el servidor." });
    }

    const msg = message.toLowerCase();

    // --- AVANCE GENERAL ---
    if (msg.includes("avance")) {
      const real   = formatPct(r["Avance Real"], false);
      const plan   = formatPct(r["Avance Plan"], false);
      const spi    = toNum(r["SPI"]);
      const spiStr = isNaN(spi) ? "N/D" : parseFloat(spi.toFixed(3)).toString();
      const estado = !isNaN(spi) && spi >= 1
        ? "el proyecto va adelantado respecto al plan ✅"
        : "el proyecto está por detrás del plan ⚠️, hay que revisar el cronograma";
      reply = `📊 Aquí el estado de avance:\n• Avance real: ${real}%\n• Avance planificado: ${plan}%\n• SPI: ${spiStr} → ${estado}`;
    }

    // --- HINCAS ---
    else if (msg.includes("hinca")) {
      const inst  = toNum(r["Hincas_Instaladas"]);
      const total = toNum(r["Hincas_Totales"]);
      const pct   = (!isNaN(inst) && !isNaN(total) && total > 0)
        ? parseFloat((inst / total * 100).toFixed(2)).toString()
        : "N/D";
      reply = `🔩 Estado de hincas:\n• Instaladas: ${formatInt(inst)} de ${formatInt(total)}\n• Progreso: ${pct}%`;
    }

    // --- TRACKERS ---
    else if (msg.includes("tracker")) {
      const inst  = toNum(r["Trackers_Instalados"]);
      const total = toNum(r["Trackers_Totales"]);
      const pct   = (!isNaN(inst) && !isNaN(total) && total > 0)
        ? parseFloat((inst / total * 100).toFixed(2)).toString()
        : "N/D";
      reply = `☀️ Estado de trackers:\n• Instalados: ${formatInt(inst)} de ${formatInt(total)}\n• Progreso: ${pct}%`;
    }

    // --- MÓDULOS ---
    else if (msg.includes("modulo") || msg.includes("módulo") || msg.includes("panel")) {
      const inst  = toNum(r["Modulos_Instalados"]);
      const total = toNum(r["Modulos_Totales"]);
      const pct   = (!isNaN(inst) && !isNaN(total) && total > 0)
        ? parseFloat((inst / total * 100).toFixed(2)).toString()
        : "N/D";
      reply = `🔋 Estado de módulos:\n• Instalados: ${formatInt(inst)} de ${formatInt(total)}\n• Progreso: ${pct}%`;
    }

    // --- FACTURACIÓN ---
    else if (msg.includes("facturac")) {
      const actual = formatCOP(r["Facturación Actual"]);
      const plan   = formatCOP(r["Facturación_Plan"]);
      const pct    = formatPct(r["Facturacion_Porcentaje"], false);
      reply = `💰 Estado de facturación:\n• Facturado a la fecha: $${actual} COP\n• Meta planificada: $${plan} COP\n• Avance: ${pct}%`;
    }

    // --- MANO DE OBRA ---
    else if (msg.includes("mano") || msg.includes("personal") || msg.includes("personas") || msg.includes("trabajador")) {
      const total     = toNum(r["Mano_de_Obra_Total"]);
      const directa   = toNum(r["Mano_de_Obra_Directa"]);
      const indirecta = toNum(r["Mano_de_Obra_Indirecta"]);
      reply = `👷 Personal en obra:\n• Total: ${formatInt(total)} personas\n• Mano de obra directa: ${formatInt(directa)}\n• Mano de obra indirecta: ${formatInt(indirecta)}`;
    }

    // --- GENERACIÓN ---
    else if (msg.includes("generacion") || msg.includes("generación")) {
      const generacion = toNum(r["Generación"]);
      const pct        = toNum(r["Porcentaje de Generación"]);
      const generacionStr = !isNaN(generacion) ? parseFloat(generacion.toFixed(2)).toString() : "N/D";
      const pctStr        = !isNaN(pct) ? parseFloat(pct.toFixed(2)).toString() : "N/D";
      reply = `⚡ Generación del parque:\n• Generación actual: ${generacionStr} MW\n• Representa el ${pctStr}% del total del parque`;
    }

    // --- DOSSIER ---
    else if (msg.includes("dossier")) {
      const avanceDossier = toNum(r["Dossier"]);
      const pctStr = !isNaN(avanceDossier) ? parseFloat(avanceDossier.toFixed(2)).toString() : "N/D";
      reply = `📂 Avance del Dossier: ${pctStr}%\n\nPara mayor detalle, revisa el apartado de Gestión de Calidad → botón Dossier en el dashboard 👆`;
    }

    // --- NO CONFORMIDADES ---
    else if (msg.includes("conformidad") || msg.includes("no conformidad")) {
      const noConformidades = r["Balance de No Conformidades"];
      reply = noConformidades
        ? `⚠️ Balance de No Conformidades:\n${noConformidades}`
        : "No encontré información sobre No Conformidades en este momento. Verifica en el dashboard de calidad.";
    }

    // --- RESUMEN GENERAL ---
    else if (msg.includes("resumen") || msg.includes("estado") || msg.includes("general")) {
      const real   = formatPct(r["Avance Real"], false);
      const plan   = formatPct(r["Avance Plan"], false);
      const spi    = toNum(r["SPI"]);
      const spiStr = isNaN(spi) ? "N/D" : parseFloat(spi.toFixed(3)).toString();
      const total  = toNum(r["Mano_de_Obra_Total"]);
      const actual = formatCOP(r["Facturación Actual"]);
      reply = `📋 Resumen PFV Puerta de Oro:\n• Avance real: ${real}% (Plan: ${plan}%)\n• SPI: ${spiStr}\n• Personal en obra: ${formatInt(total)} personas\n• Facturación: $${actual} COP\n\n¿Quieres profundizar en algún tema? 😊`;
    }

    // --- RENDIMIENTOS TENDIDO CABLE SOLAR ---
    else if (msg.includes("rendimiento") || msg.includes("tendido") || msg.includes("cable")) {
      const rendimiento = toNum(r["Rendimientos Tendido cable solar"]);
      const rendStr = isNaN(rendimiento) ? "N/D" : formatInt(rendimiento);
      reply = `📏 Rendimiento de tendido cable solar:\n• Última semana registrada: ${rendStr} metros\n\n📌 Para detalle completo: Gestión de Construcción → Plan de Acción/Bonus → Tendido Solar`;
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
