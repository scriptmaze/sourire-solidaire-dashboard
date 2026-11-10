/* global XLSX */
(function () {
  "use strict";

  const MONTHS = ["Mai", "Juin", "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre", "Janvier", "Février", "Mars", "Avril"];
  const REGIONS = ["Laval", "Montréal", "Laurentides", "Lanaudière", "Autre"];
  const REFERRERS = ["Dentiste", "CISSS / CIUSSS", "Organisme communautaire", "Famille", "Pédiatre / médecin", "Autre"];
  const state = { history: emptyHistory(), latest: null, warnings: [], activeChart: "region" };

  const el = {
    historyFile: document.querySelector("#history-file"),
    dentitekFiles: document.querySelector("#dentitek-files"),
    fileSummary: document.querySelector("#file-summary"),
    process: document.querySelector("#process-files"),
    clear: document.querySelector("#clear-session"),
    messages: document.querySelector("#messages"),
    results: document.querySelector("#results"),
    latestMonth: document.querySelector("#latest-month"),
    kpis: document.querySelector("#kpi-grid"),
    regionTable: document.querySelector("#region-table"),
    referrerTable: document.querySelector("#referrer-table"),
    chart: document.querySelector("#trend-chart"),
    download: document.querySelector("#download-history"),
    tabs: document.querySelectorAll(".tab"),
  };

  function emptyHistory() {
    return Object.fromEntries(MONTHS.map((month) => [month, { regions: zeros(REGIONS), referrers: zeros(REFERRERS) }]));
  }

  function zeros(labels) {
    return Object.fromEntries(labels.map((label) => [label, null]));
  }

  function normalise(value) {
    return String(value ?? "")
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .toLowerCase()
      .trim()
      .replace(/\s+/g, " ");
  }

  function cleanLabel(value) {
    return String(value ?? "").trim();
  }

  function monthFromName(name) {
    const found = String(name).match(/(?:20)?26[-_ ]?(0[5-9]|1[0-2]|0[1-4])/i);
    if (!found) return null;
    const index = Number(found[1]) - 1;
    return [4, 5, 6, 7, 8, 9, 10, 11, 0, 1, 2, 3].includes(index) ? MONTHS[(index + 8) % 12] : null;
  }

  function fileArray(input) {
    return Array.from(input.files || []);
  }

  function updateFileSummary() {
    const history = fileArray(el.historyFile)[0];
    const exports = fileArray(el.dentitekFiles);
    const parts = [];
    if (history) parts.push(`Historique : ${history.name}`);
    if (exports.length) parts.push(`${exports.length} export${exports.length > 1 ? "s" : ""} Dentitek : ${exports.map((file) => file.name).join(", ")}`);
    el.fileSummary.textContent = parts.length ? parts.join(" · ") : "Aucun fichier chargé.";
  }

  async function readWorkbook(file) {
    const buffer = await file.arrayBuffer();
    return XLSX.read(buffer, { type: "array", cellDates: true });
  }

  function getRows(workbook, sheetName) {
    return XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: null, raw: true });
  }

  function findSheetWithHeaders(workbook, requiredHeaders) {
    for (const name of workbook.SheetNames) {
      const rows = getRows(workbook, name);
      const header = (rows[0] || []).map(normalise);
      if (requiredHeaders.every((required) => header.includes(normalise(required)))) return { name, rows, header };
    }
    return null;
  }

  function toNumber(value) {
    if (typeof value === "number") return value;
    const parsed = Number(String(value ?? "").replace(",", "."));
    return Number.isFinite(parsed) ? parsed : 0;
  }

  function classifyRegion(value) {
    const text = normalise(value);
    if (text.startsWith("13") || text.includes("laval")) return "Laval";
    if (text.startsWith("6") || text.includes("montreal")) return "Montréal";
    if (text.startsWith("15") || text.includes("laurentides")) return "Laurentides";
    if (text.startsWith("14") || text.includes("lanaudiere")) return "Lanaudière";
    return "Autre";
  }

  function classifyReferrer(value) {
    const text = normalise(value);
    if (text === "dentiste") return "Dentiste";
    if (/^\d/.test(text) || text.includes("cisss") || text.includes("ciuss")) return "CISSS / CIUSSS";
    if (text.includes("organisme")) return "Organisme communautaire";
    if (text.includes("fratrie") || text.includes("famille") || text === "patient") return "Famille";
    if (text.includes("pediatr") || text.includes("medecin") || text.includes("docteur")) return "Pédiatre / médecin";
    return "Autre";
  }

  function parseDentitek(workbook, filename) {
    const source = findSheetWithHeaders(workbook, ["nombre_refere", "ref_suffix", "nomprov"]);
    if (!source) throw new Error(`${filename} ne contient pas la feuille Dentitek attendue (nombre_refere, ref_suffix, nomprov).`);
    const indexes = {
      count: source.header.indexOf("nombre_refere"),
      suffix: source.header.indexOf("ref_suffix"),
      province: source.header.indexOf("nomprov"),
    };
    const result = { regions: zeros(REGIONS), referrers: zeros(REFERRERS), unclassified: new Map(), sourceSheet: source.name };
    Object.keys(result.regions).forEach((key) => { result.regions[key] = 0; });
    Object.keys(result.referrers).forEach((key) => { result.referrers[key] = 0; });
    source.rows.slice(1).forEach((row) => {
      const count = toNumber(row[indexes.count]);
      if (!count) return;
      result.regions[classifyRegion(row[indexes.province])] += count;
      const category = classifyReferrer(row[indexes.suffix]);
      result.referrers[category] += count;
      if (category === "Autre") {
        const sourceValue = cleanLabel(row[indexes.suffix]) || "(vide)";
        result.unclassified.set(sourceValue, (result.unclassified.get(sourceValue) || 0) + count);
      }
    });
    return result;
  }

  function parseHistory(workbook) {
    const name = workbook.SheetNames.find((sheet) => normalise(sheet) === "donnees mensuelles");
    if (!name) throw new Error("Le fichier historique ne contient pas l’onglet « Données mensuelles ».");
    const rows = getRows(workbook, name);
    const history = emptyHistory();
    const headers = rows[0] || [];
    const columns = headers.map((value) => MONTHS.indexOf(cleanLabel(value)));
    const targets = [
      ["1.1.1 laval", "regions", "Laval"],
      ["1.1.2 montreal", "regions", "Montréal"],
      ["1.1.3 laurentides", "regions", "Laurentides"],
      ["1.1.4 lanaudiere", "regions", "Lanaudière"],
      ["1.1.5 autre", "regions", "Autre"],
      ["1.2.1 dentiste", "referrers", "Dentiste"],
      ["1.2.2 cisss", "referrers", "CISSS / CIUSSS"],
      ["1.2.3 organisme", "referrers", "Organisme communautaire"],
      ["1.2.4 patient", "referrers", "Famille"],
      ["1.2.4 famille", "referrers", "Famille"],
      ["1.2.5 pediatre", "referrers", "Pédiatre / médecin"],
      ["1.2.6 autre", "referrers", "Autre"],
    ];
    rows.slice(1).forEach((row) => {
      const label = normalise(row[0]);
      const target = targets.find(([prefix]) => label.startsWith(prefix));
      if (!target) return;
      columns.forEach((monthIndex, col) => {
        if (monthIndex < 0 || row[col] === null || row[col] === "") return;
        const value = toNumber(row[col]);
        if (Number.isFinite(value)) history[MONTHS[monthIndex]][target[1]][target[2]] = value;
      });
    });
    return history;
  }

  function showMessages(messages) {
    el.messages.replaceChildren();
    messages.forEach(({ type, text }) => {
      const box = document.createElement("div");
      box.className = `message ${type}`;
      box.textContent = text;
      el.messages.append(box);
    });
  }

  function total(group) {
    return Object.values(group).reduce((sum, value) => sum + (Number(value) || 0), 0);
  }

  function htmlEscape(value) {
    return String(value).replace(/[&<>'"]/g, (char) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", "'": "&#39;", '"': "&quot;" }[char]));
  }

  function buildTable(group) {
    const totalValue = total(group);
    return `<table><thead><tr><th>Catégorie</th><th>Nombre</th><th>Part</th></tr></thead><tbody>${Object.entries(group).map(([label, value]) => {
      const share = totalValue ? `${Math.round((value / totalValue) * 100)} %` : "—";
      return `<tr><td>${htmlEscape(label)}</td><td>${value.toLocaleString("fr-CA")}</td><td>${share}</td></tr>`;
    }).join("")}</tbody><tfoot><tr><th>Total</th><th>${totalValue.toLocaleString("fr-CA")}</th><th>100 %</th></tr></tfoot></table>`;
  }

  function renderResults() {
    const data = state.history[state.latest];
    el.results.hidden = false;
    el.latestMonth.textContent = state.latest;
    const kpis = [
      ["Références reçues", total(data.regions)],
      ["Références de dentistes", data.referrers.Dentiste],
      ["Références CISSS / CIUSSS", data.referrers["CISSS / CIUSSS"]],
      ["Références de familles", data.referrers.Famille],
    ];
    el.kpis.innerHTML = kpis.map(([label, value]) => `<div class="kpi"><span>${htmlEscape(label)}</span><strong>${Number(value).toLocaleString("fr-CA")}</strong></div>`).join("");
    el.regionTable.innerHTML = buildTable(data.regions);
    el.referrerTable.innerHTML = buildTable(data.referrers);
    renderChart();
  }

  function renderChart() {
    const groupName = state.activeChart === "region" ? "regions" : "referrers";
    const labels = state.activeChart === "region" ? REGIONS : REFERRERS;
    const records = MONTHS.filter((month) => total(state.history[month][groupName]) > 0).map((month) => ({ month, data: state.history[month][groupName] }));
    if (!records.length) { el.chart.innerHTML = "<p>Aucune donnée historique disponible.</p>"; return; }
    const colors = ["#176b5a", "#e09848", "#5884bc", "#b45259", "#8768a8", "#6d7f4d"];
    const max = Math.max(1, ...records.flatMap(({ data }) => labels.map((label) => data[label] || 0)));
    const width = 800, height = 300, left = 46, right = 20, top = 20, bottom = 46;
    const x = (index) => left + (index * (width - left - right)) / Math.max(records.length - 1, 1);
    const y = (value) => top + (height - top - bottom) * (1 - value / max);
    const lines = labels.map((label, seriesIndex) => {
      const points = records.map(({ data }, index) => `${x(index)},${y(data[label] || 0)}`).join(" ");
      return `<polyline fill="none" stroke="${colors[seriesIndex]}" stroke-width="3" points="${points}" />`;
    }).join("");
    const dots = labels.map((label, seriesIndex) => records.map(({ data }, index) => `<circle cx="${x(index)}" cy="${y(data[label] || 0)}" r="3.5" fill="${colors[seriesIndex]}"><title>${htmlEscape(`${label} — ${records[index].month}: ${data[label] || 0}`)}</title></circle>`).join("")).join("");
    const xLabels = records.map(({ month }, index) => `<text x="${x(index)}" y="${height - 18}" text-anchor="middle">${month}</text>`).join("");
    const yLabels = [0, Math.round(max / 2), max].map((value) => `<text x="${left - 9}" y="${y(value) + 4}" text-anchor="end">${value}</text><line x1="${left}" x2="${width - right}" y1="${y(value)}" y2="${y(value)}" />`).join("");
    const legend = labels.map((label, index) => `<span><i style="background:${colors[index]}"></i>${htmlEscape(label)}</span>`).join("");
    el.chart.innerHTML = `<svg viewBox="0 0 ${width} ${height}" role="img" aria-label="Évolution mensuelle"><g class="grid">${yLabels}</g>${lines}${dots}<g class="axis-labels">${xLabels}</g></svg><div class="legend">${legend}</div>`;
  }

  function buildExportWorkbook() {
    const rows = [["Données mensuelles", ...MONTHS]];
    const addSection = (title, labels, key) => {
      rows.push([title]);
      labels.forEach((label) => rows.push([label, ...MONTHS.map((month) => state.history[month][key][label] ?? "")]));
    };
    addSection("1. Répartition des références par provenance", REGIONS, "regions");
    addSection("1.2 Répartition par type de référent", REFERRERS, "referrers");
    const wb = XLSX.utils.book_new();
    const dataSheet = XLSX.utils.aoa_to_sheet(rows);
    dataSheet["!cols"] = [{ wch: 42 }, ...MONTHS.map(() => ({ wch: 13 }))];
    XLSX.utils.book_append_sheet(wb, dataSheet, "Données mensuelles");
    const latest = state.history[state.latest];
    const dashboard = [["Tableau de bord — Sourire Solidaire"], ["Dernier mois", state.latest], [], ["Répartition par provenance", "Nombre"], ...REGIONS.map((label) => [label, latest.regions[label]]), ["Total", total(latest.regions)], [], ["Répartition par type de référent", "Nombre"], ...REFERRERS.map((label) => [label, latest.referrers[label]]), ["Total", total(latest.referrers)]];
    const dashboardSheet = XLSX.utils.aoa_to_sheet(dashboard);
    dashboardSheet["!cols"] = [{ wch: 42 }, { wch: 16 }];
    XLSX.utils.book_append_sheet(wb, dashboardSheet, "Tableau de bord");
    const notes = [["Confidentialité"], ["Ce fichier contient seulement des agrégats mensuels."], ["Les exports Dentitek bruts ne sont pas inclus."], [], ["Règles de classement"], ["CISSS / CIUSSS", "Codes numériques Dentitek ou libellés CISSS / CIUSSS"], ["Famille", "Fratrie Famille ou Patient"], ["Autre", "Toute valeur restante, y compris une valeur vide"]];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(notes), "Notes");
    return wb;
  }

  async function process() {
    const messages = [];
    const exports = fileArray(el.dentitekFiles);
    if (!exports.length) { showMessages([{ type: "error", text: "Ajoutez au moins un export Dentitek avant de lancer le calcul." }]); return; }
    state.history = emptyHistory();
    state.warnings = [];
    try {
      const historyFile = fileArray(el.historyFile)[0];
      if (historyFile) state.history = parseHistory(await readWorkbook(historyFile));
      for (const file of exports) {
        const month = monthFromName(file.name);
        if (!month) throw new Error(`${file.name} : le mois n’a pas pu être déterminé à partir du nom du fichier. Utilisez un nom contenant AAAA-MM.`);
        const result = parseDentitek(await readWorkbook(file), file.name);
        state.history[month] = { regions: result.regions, referrers: result.referrers };
        const unclassified = [...result.unclassified.entries()].filter(([label]) => label !== "(vide)");
        if (unclassified.length) messages.push({ type: "warning", text: `${month} : ${unclassified.length} valeur(s) classée(s) dans « Autre ». Vérifiez-les avant la diffusion du rapport.` });
      }
      const available = MONTHS.filter((month) => total(state.history[month].regions) > 0);
      state.latest = available[available.length - 1];
      if (!state.latest) throw new Error("Aucune référence exploitable n’a été trouvée dans les exports.");
      messages.unshift({ type: "success", text: `${exports.length} export(s) traité(s). Les résultats sont calculés dans ce navigateur.` });
      showMessages(messages);
      renderResults();
    } catch (error) {
      el.results.hidden = true;
      showMessages([{ type: "error", text: error.message || "Le traitement n’a pas pu être terminé." }]);
    }
  }

  function reset() {
    state.history = emptyHistory(); state.latest = null; state.warnings = [];
    el.historyFile.value = ""; el.dentitekFiles.value = ""; el.results.hidden = true;
    showMessages([{ type: "success", text: "La session a été effacée. Les données importées ne sont plus conservées par cette page." }]);
    updateFileSummary();
  }

  el.historyFile.addEventListener("change", updateFileSummary);
  el.dentitekFiles.addEventListener("change", updateFileSummary);
  el.process.addEventListener("click", process);
  el.clear.addEventListener("click", reset);
  el.download.addEventListener("click", () => {
    const wb = buildExportWorkbook();
    XLSX.writeFile(wb, `Historique_Sourire_Solidaire_${state.latest}_2026.xlsx`, { compression: true });
  });
  el.tabs.forEach((tab) => tab.addEventListener("click", () => {
    state.activeChart = tab.dataset.chart;
    el.tabs.forEach((button) => { const selected = button === tab; button.classList.toggle("active", selected); button.setAttribute("aria-selected", selected); });
    renderChart();
  }));
})();
