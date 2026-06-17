const DATA_URL = `data.json?v=${Date.now()}`;
const CACHE_KEY = "calottino-site-data";

let siteData;
let memberRows = [];

const euro = new Intl.NumberFormat("it-IT", {
  style: "currency",
  currency: "EUR",
  maximumFractionDigits: 2,
});

const integer = new Intl.NumberFormat("it-IT", {
  maximumFractionDigits: 0,
});

document.addEventListener("DOMContentLoaded", init);

async function init() {
  document.querySelector("#reloadDataButton")?.addEventListener("click", () => window.location.reload());

  try {
    const response = await fetch(DATA_URL, { cache: "no-store" });
    if (!response.ok) throw new Error(`HTTP ${response.status}`);
    siteData = await response.json();
    saveCachedData(siteData);
    memberRows = siteData.excel.people.rows;
    hideDataBanner();
    renderAll();
    bindEvents();
  } catch (error) {
    loadCachedData(error);
  }
}

function loadCachedData(error) {
  const cached = readCachedData();
  if (!cached) {
    showDataBanner(
      "Impossibile caricare i dati. Controlla la connessione e ricarica la pagina quando torna internet.",
    );
    return;
  }

  try {
    siteData = JSON.parse(cached);
    memberRows = siteData.excel.people.rows;
    renderAll();
    bindEvents();
    showDataBanner(
      `Dati online non raggiungibili (${error.message}). Sto mostrando l'ultimo aggiornamento salvato su questo dispositivo.`,
      "warning",
    );
  } catch {
    clearCachedData();
    showDataBanner(
      "Impossibile caricare i dati. Controlla la connessione e ricarica la pagina quando torna internet.",
    );
  }
}

function saveCachedData(data) {
  try {
    localStorage.setItem(CACHE_KEY, JSON.stringify(data));
  } catch {
    // Il sito resta pienamente leggibile anche se il browser blocca la cache locale.
  }
}

function readCachedData() {
  try {
    return localStorage.getItem(CACHE_KEY);
  } catch {
    return null;
  }
}

function clearCachedData() {
  try {
    localStorage.removeItem(CACHE_KEY);
  } catch {
    // Nessuna azione necessaria.
  }
}

function showDataBanner(message, tone = "error") {
  const banner = document.querySelector("#dataBanner");
  banner.hidden = false;
  banner.className = `data-banner data-banner-${tone}`;
  banner.querySelector("span").textContent = message;
}

function hideDataBanner() {
  const banner = document.querySelector("#dataBanner");
  banner.hidden = true;
}

function bindEvents() {
  const search = document.querySelector("#memberSearch");
  search.addEventListener("input", () => {
    const query = search.value.trim().toLowerCase();
    const filtered = memberRows.filter((row) => {
      return Object.values(row).some((value) => String(value ?? "").toLowerCase().includes(query));
    });
    renderPeople(filtered);
  });
}

function renderAll() {
  const excel = siteData.excel;
  const summary = excel.summary;

  document.querySelector("#updatedAt").textContent = formatDateTime(siteData.sources.excel.modifiedAt);
  document.querySelector("#workbookSubtitle").textContent = excel.workbookSubtitle;
  document.querySelector("#cashHero").textContent = euro.format(summary.cashAvailable);
  document.querySelector("#sourceInfo").textContent = `Fonte: ${siteData.sources.excel.file}`;

  renderKpis(summary);
  renderDues(summary);
  renderFlowList(excel.flows.rows);
  renderInventoryList(excel.inventory.rows);
  renderPeople(memberRows);
  renderDataTable("#salesTable", excel.sales.rows, ["Data", "Acquirente", "Descrizione", "Quantità", "Prezzo unitario", "Ricavo", "Utile lordo", "Incassato?", "Metodo", "Note"], "Nessuna vendita registrata.");
  renderDataTable("#purchasesTable", excel.purchases.rows, ["Data", "Fornitore", "Lotto / Descrizione", "Quantità", "Costo unitario", "Costo totale", "Stato pagamento"], "Nessun acquisto registrato.");
  renderDataTable("#expensesTable", excel.expenses.rows, ["Data", "Categoria", "Descrizione", "Importo", "Pagato da", "Metodo", "Approvata?", "Note"], "Nessuna spesa registrata.");
  renderDataTable("#parametersTable", excel.parameters.rows, ["Voce", "Valore", "Unità", "Note"], "Nessun parametro disponibile.");
  renderStatute(siteData.statute);
}

function renderKpis(summary) {
  const cards = [
    ["Cassa disponibile", euro.format(summary.cashAvailable), "Saldo stimato da entrate e uscite"],
    ["Quote incassate", euro.format(summary.duesCollected), `${summary.membersPaid}/${summary.membersTotal} persone in regola`],
    ["Ricavi patch", euro.format(summary.patchRevenue), `${euro.format(summary.patchRevenueExternal || 0)} esterni, ${euro.format(summary.patchRevenueInternal || 0)} interni`],
    ["Patch disponibili", integer.format(summary.patchAvailable), `${integer.format(summary.patchPurchased)} acquistate, ${integer.format(summary.patchSold)} vendute`],
    ["Margini unitari", `${euro.format(summary.patchUnitMarginExternal || summary.patchUnitMargin || 0)} / ${euro.format(summary.patchUnitMarginInternal || 0)}`, "Esterni / interni"],
  ];

  document.querySelector("#kpiGrid").innerHTML = cards
    .map(([label, value, note]) => `
      <article class="kpi-card">
        <span>${escapeHtml(label)}</span>
        <strong>${escapeHtml(value)}</strong>
        <small>${escapeHtml(note)}</small>
      </article>
    `)
    .join("");
}

function renderDues(summary) {
  const rate = Math.max(0, Math.min(100, Number(summary.duesCompletionRate) || 0));
  document.querySelector("#duesRate").textContent = `${rate.toLocaleString("it-IT")}%`;
  document.querySelector("#duesRing").style.setProperty("--progress", `${rate * 3.6}deg`);
  document.querySelector("#duesBar").style.width = `${rate}%`;
  document.querySelector("#duesText").textContent =
    `${euro.format(summary.duesCollected)} incassati su ${euro.format(summary.duesExpected)} previsti. ` +
    `${euro.format(summary.duesRemaining)} ancora da incassare.`;
}

function renderFlowList(rows) {
  document.querySelector("#flowList").innerHTML = rows
    .map((row) => `
      <div class="flow-item">
        <span>${escapeHtml(row.Voce)}</span>
        <strong>${formatByUnit(row.Importo, row.Segno)}</strong>
      </div>
    `)
    .join("");
}

function renderInventoryList(rows) {
  document.querySelector("#inventoryList").innerHTML = rows
    .map((row) => `
      <div class="mini-item">
        <span>${escapeHtml(row.Voce)}</span>
        <strong>${formatInventory(row.Valore, row.Unità)}</strong>
      </div>
    `)
    .join("");
}

function renderPeople(rows) {
  renderDataTable(
    "#peopleTable",
    rows,
    ["ID", "Nome / Cognome", "Quota dovuta", "Quota pagata", "Saldo", "Stato", "Data pagamento", "Metodo", "Note"],
    "Nessun componente trovato.",
  );
}

function renderDataTable(target, rows, columns, emptyText) {
  const element = document.querySelector(target);
  if (!rows.length) {
    element.innerHTML = `<div class="empty-state">${escapeHtml(emptyText)}</div>`;
    return;
  }

  element.innerHTML = `
    <table>
      <thead>
        <tr>${columns.map((column) => `<th>${escapeHtml(column)}</th>`).join("")}</tr>
      </thead>
      <tbody>
        ${rows.map((row) => `
          <tr>
            ${columns.map((column) => `<td>${formatCell(row[column], column, row)}</td>`).join("")}
          </tr>
        `).join("")}
      </tbody>
    </table>
  `;
}

function renderStatute(statute) {
  const sections = statute?.sections || [];
  document.querySelector("#statuteNotice").textContent = statute?.notice || "";

  if (!sections.length) {
    document.querySelector("#statuteIndex").innerHTML = "";
    document.querySelector("#statuteContent").innerHTML =
      `<div class="empty-state">Statuto non disponibile.</div>`;
    return;
  }

  document.querySelector("#statuteIndex").innerHTML = sections
    .map((section, index) => `<a href="#statute-${index}">${escapeHtml(section.title)}</a>`)
    .join("");

  document.querySelector("#statuteContent").innerHTML = sections
    .map((section, index) => `
      <details class="statute-card" id="statute-${index}" ${index < 3 ? "open" : ""}>
        <summary>${escapeHtml(section.title)}</summary>
        <div class="statute-body">
          ${section.blocks.map(renderStatuteBlock).join("")}
        </div>
      </details>
    `)
    .join("");
}

function renderStatuteBlock(block) {
  if (block.type === "paragraph") {
    return `<p>${escapeHtml(block.text)}</p>`;
  }

  if (block.rows.length === 1 && block.rows[0].length === 1) {
    return `<p>${escapeHtml(block.rows[0][0])}</p>`;
  }

  const [head, ...body] = block.rows;
  return `
    <div class="statute-table">
      <table>
        <thead><tr>${head.map((cell) => `<th>${escapeHtml(cell)}</th>`).join("")}</tr></thead>
        <tbody>
          ${body.map((row) => `<tr>${row.map((cell) => `<td>${escapeHtml(cell)}</td>`).join("")}</tr>`).join("")}
        </tbody>
      </table>
    </div>
  `;
}

function formatCell(value, column, row = {}) {
  if (value === null || value === undefined || value === "") return "";
  if (column === "Valore") {
    return row.Unità === "€" ? euro.format(Number(value) || 0) : escapeHtml(String(value));
  }
  if (["Quota dovuta", "Quota pagata", "Saldo", "Prezzo unitario", "Ricavo", "Costo unitario", "Costo totale", "Importo", "Utile lordo"].includes(column)) {
    return euro.format(Number(value) || 0);
  }
  if (column === "Stato" || column === "Stato pagamento" || column === "Stato consegna" || column === "Approvata?" || column === "Incassato?") {
    return `<span class="status ${statusClass(value)}">${escapeHtml(String(value))}</span>`;
  }
  return escapeHtml(String(value));
}

function statusClass(value) {
  const normalized = String(value).toLowerCase();
  if (normalized.includes("pagata") || normalized.includes("pagato") || normalized === "sì") return "status-paid";
  if (normalized.includes("parziale")) return "status-partial";
  return "status-open";
}

function formatByUnit(value, sign) {
  const numeric = Number(value) || 0;
  if (sign === "Entrata") return `+${euro.format(numeric)}`;
  if (sign === "Uscita") return `-${euro.format(numeric)}`;
  return euro.format(numeric);
}

function formatInventory(value, unit) {
  const numeric = Number(value) || 0;
  if (unit === "€") return euro.format(numeric);
  return `${integer.format(numeric)} ${escapeHtml(unit || "")}`.trim();
}

function formatDateTime(value) {
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return value;
  return date.toLocaleString("it-IT", {
    day: "2-digit",
    month: "2-digit",
    year: "numeric",
    hour: "2-digit",
    minute: "2-digit",
  });
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}
