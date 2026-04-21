let rawData = [];
let filteredProperties = [];
let filteredAvailableProperties = [];
let selectedPreparations = new Set();
let selectedProperties = new Set();

const DEFAULT_WORKBOOK = "suitbrowser_rules.xlsx";
const APPLICATION_COLUMN = "Applications";
const PREPARATION_COLUMN = "Preparation (to choose in the browser)";
const NAME_COLUMN = "Name in the SUITbrowser";
const TECHNICAL_NAME_COLUMN = "SUIT technical name";
const MITOPEDIA_COLUMN = "MitoPedia page";
const PROPERTY_START_AFTER = "Applications 2";
const SCORE_LABELS = {
  0: "Not applicable",
  1: "Not recommended",
  2: "Kind of suitable",
  3: "Very suitable",
};

const appSelect = document.getElementById("applicationSelect");
const appLayout = document.getElementById("appLayout");
const preparationOptions = document.getElementById("preparationOptions");
const propertySearch = document.getElementById("propertySearch");
const propertyOptions = document.getElementById("propertyOptions");
const resultsSummary = document.getElementById("resultsSummary");
const tableWrap = document.getElementById("tableWrap");
const openMatrixModalBtn = document.getElementById("openMatrixModalBtn");
const matrixModal = document.getElementById("matrixModal");
const matrixModalBackdrop = document.getElementById("matrixModalBackdrop");
const closeMatrixModalBtn = document.getElementById("closeMatrixModalBtn");
const matrixModalWrap = document.getElementById("matrixModalWrap");
const matrixModalSummary = document.getElementById("matrixModalSummary");
const fileStatus = document.getElementById("fileStatus");
const reloadDefaultBtn = document.getElementById("reloadDefaultBtn");
const clearPreparationsBtn = document.getElementById("clearPreparationsBtn");
const selectAllVisiblePropsBtn = document.getElementById("selectAllVisiblePropsBtn");
const clearPropertiesBtn = document.getElementById("clearPropertiesBtn");
const resetFiltersBtn = document.getElementById("resetFiltersBtn");
const toggleFiltersBtn = document.getElementById("toggleFiltersBtn");
const FILTERS_COLLAPSED_KEY = "suitbrowser.filtersCollapsed";

function uniqueValues(rows, key) {
  return [...new Set(rows.map((row) => row[key]).filter(Boolean))];
}

function getPropertyColumns(rows = rawData) {
  if (!rows.length) {
    return [];
  }

  const keys = Object.keys(rows[0]);
  const startIndex = keys.indexOf(PROPERTY_START_AFTER);
  return startIndex >= 0 ? keys.slice(startIndex + 1) : [];
}

function getAvailablePropertyColumns(rows) {
  return getPropertyColumns(rows).filter((property) =>
    rows.some((row) => String(row[property] ?? "").trim() !== "")
  );
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function scoreClass(value) {
  const normalized = String(value ?? "").trim();
  if (/^[0-3]$/.test(normalized)) {
    return `score-${normalized}`;
  }
  return "";
}

function hasPositiveScore(row, property) {
  const normalized = String(row[property] ?? "").trim();
  return normalized !== "" && normalized !== "0";
}

function parseNumericScore(value) {
  const normalized = String(value ?? "").trim();
  if (!/^[0-3]$/.test(normalized)) {
    return null;
  }
  return Number(normalized);
}

function formatScoreDisplay(value) {
  const normalized = String(value ?? "").trim();
  if (normalized === "") {
    return "";
  }

  return normalized
    .split("/")
    .map((part) => part.trim())
    .map((part) => {
      const numericScore = parseNumericScore(part);
      return numericScore === null ? part : SCORE_LABELS[numericScore];
    })
    .join(" / ");
}

function sanitizeUrl(value) {
  const url = String(value ?? "").trim();
  if (!/^https?:\/\//i.test(url)) {
    return "";
  }
  return url;
}

function setLoadingState(message) {
  if (fileStatus) {
    fileStatus.innerHTML = message;
  }
  appSelect.disabled = true;
  propertySearch.disabled = true;
  resetFiltersBtn.disabled = true;
  preparationOptions.textContent = "Loading workbook...";
  propertyOptions.textContent = "Loading workbook...";
  resultsSummary.textContent = "Loading workbook...";
}

function getRowsForPreparationOptions() {
  const selectedApplication = appSelect.value;

  return rawData.filter((row) => {
    return !selectedApplication || row[APPLICATION_COLUMN] === selectedApplication;
  });
}

function getRowsForPropertyOptions() {
  return getFilteredRows();
}

function renderPreparations() {
  const preparationRows = getRowsForPreparationOptions();
  const allPreparations = uniqueValues(preparationRows, PREPARATION_COLUMN);
  selectedPreparations = new Set(
    [...selectedPreparations].filter((preparation) =>
      allPreparations.includes(preparation)
    )
  );
  preparationOptions.innerHTML = "";

  if (!allPreparations.length) {
    preparationOptions.textContent = "No preparations available.";
    return;
  }

  allPreparations.forEach((preparation) => {
    const label = document.createElement("label");
    label.className = "chip-option";

    const input = document.createElement("input");
    input.type = "checkbox";
    input.value = preparation;
    input.checked = selectedPreparations.has(preparation);
    input.addEventListener("change", () => {
      if (input.checked) {
        selectedPreparations.add(preparation);
      } else {
        selectedPreparations.delete(preparation);
      }
      syncUiAndRender();
    });

    const span = document.createElement("span");
    span.textContent = preparation;

    label.append(input, span);
    preparationOptions.appendChild(label);
  });
}

function renderPropertyOptions() {
  const searchValue = propertySearch.value.trim().toLowerCase();
  const propertyRows = getRowsForPropertyOptions();
  const propertyColumns = getPropertyColumns();
  const availableProperties = new Set(getAvailablePropertyColumns(propertyRows));
  filteredProperties = propertyColumns.filter((property) =>
    property.toLowerCase().includes(searchValue)
  );
  filteredAvailableProperties = filteredProperties.filter((property) =>
    availableProperties.has(property)
  );

  propertyOptions.innerHTML = "";

  if (!filteredProperties.length) {
    propertyOptions.textContent = propertyColumns.length
      ? "No properties match this search."
      : "No properties available.";
    return;
  }

  filteredProperties.forEach((property) => {
    const label = document.createElement("label");
    label.className = "property-option";
    const isAvailable = availableProperties.has(property);

    if (!isAvailable) {
      label.classList.add("unavailable");
    }

    const input = document.createElement("input");
    input.type = "checkbox";
    input.value = property;
    input.checked = selectedProperties.has(property);
    input.addEventListener("change", () => {
      if (input.checked) {
        selectedProperties.add(property);
      } else {
        selectedProperties.delete(property);
      }
      renderTable();
    });

    const span = document.createElement("span");
    span.textContent = property;
    span.title = isAvailable ? property : `${property} is unavailable for the current filters`;

    label.append(input, span);
    propertyOptions.appendChild(label);
  });
}

function renderEmptyTable(message, detail) {
  tableWrap.classList.add("empty-wrap");
  tableWrap.innerHTML = `
    <div class="empty-state-box">
      <h3>${escapeHtml(message)}</h3>
      <p>${escapeHtml(detail)}</p>
    </div>
  `;
}

function renderEmptyModalTable(message, detail) {
  if (!matrixModalWrap) {
    return;
  }

  matrixModalWrap.classList.add("empty-wrap");
  matrixModalWrap.innerHTML = `
    <div class="empty-state-box">
      <h3>${escapeHtml(message)}</h3>
      <p>${escapeHtml(detail)}</p>
    </div>
  `;
}

function setModalSummary(text) {
  if (matrixModalSummary) {
    matrixModalSummary.textContent = text;
  }
}

function getFilteredRows() {
  const selectedApplication = appSelect.value;

  return rawData.filter((row) => {
    const matchesApplication =
      !selectedApplication || row[APPLICATION_COLUMN] === selectedApplication;
    const matchesPreparation =
      !selectedPreparations.size || selectedPreparations.has(row[PREPARATION_COLUMN]);
    return matchesApplication && matchesPreparation;
  });
}

function syncUiAndRender() {
  renderPreparations();
  renderPropertyOptions();
  renderTable();
}

function setFiltersCollapsed(collapsed) {
  if (!appLayout || !toggleFiltersBtn) {
    return;
  }

  appLayout.classList.toggle("filters-collapsed", collapsed);
  toggleFiltersBtn.setAttribute("aria-expanded", String(!collapsed));
  toggleFiltersBtn.textContent = collapsed ? "Show filters" : "Hide filters";

  try {
    localStorage.setItem(FILTERS_COLLAPSED_KEY, collapsed ? "true" : "false");
  } catch (_error) {
    // Ignore storage failures and keep the UI responsive.
  }
}

function restoreFiltersCollapsedState() {
  if (!appLayout || !toggleFiltersBtn) {
    return;
  }

  let collapsed = false;

  try {
    collapsed = localStorage.getItem(FILTERS_COLLAPSED_KEY) === "true";
  } catch (_error) {
    collapsed = false;
  }

  setFiltersCollapsed(collapsed);
}

function getGroupedRows(rows, selectedPropertyList) {
  const groups = new Map();

  rows.forEach((row) => {
    const name = row[NAME_COLUMN] || "Unnamed protocol";

    if (!groups.has(name)) {
      groups.set(name, {
        name,
        preparations: new Set(),
        technicalNames: new Set(),
        mitoPediaLinks: new Set(),
        propertyValues: new Map(),
        propertyScores: new Map(),
      });
    }

    const group = groups.get(name);
    const preparation = row[PREPARATION_COLUMN];

    if (preparation) {
      group.preparations.add(preparation);
    }

    const technicalName = String(row[TECHNICAL_NAME_COLUMN] ?? "").trim();
    if (technicalName) {
      group.technicalNames.add(technicalName);
    }

    const mitoPediaLink = sanitizeUrl(row[MITOPEDIA_COLUMN]);
    if (mitoPediaLink) {
      group.mitoPediaLinks.add(mitoPediaLink);
    }

    selectedPropertyList.forEach((property) => {
      const value = String(row[property] ?? "").trim();
      if (!value) {
        return;
      }

      if (!group.propertyValues.has(property)) {
        group.propertyValues.set(property, new Set());
      }

      group.propertyValues.get(property).add(value);

      const numericScore = parseNumericScore(value);
      if (numericScore !== null) {
        const currentScore = group.propertyScores.get(property);
        if (currentScore === undefined || numericScore > currentScore) {
          group.propertyScores.set(property, numericScore);
        }
      }
    });
  });

  return [...groups.values()].map((group) => {
    const numericScores = selectedPropertyList
      .map((property) => group.propertyScores.get(property))
      .filter((score) => score !== undefined);

    const averageScore = numericScores.length
      ? numericScores.reduce((sum, score) => sum + score, 0) / numericScores.length
      : 0;

    return {
      ...group,
      averageScore,
    };
  });
}

function buildTableMarkup(rankedRows, selectedPropertyList) {
  const headerCells = selectedPropertyList
    .map(
      (property) =>
        `<th class="diag-header"><span>${escapeHtml(property)}</span></th>`
    )
    .join("");

  const bodyRows = rankedRows
    .map((group) => {
      const scoreCells = selectedPropertyList
        .map((property) => {
          const values = [...(group.propertyValues.get(property) || [])];
          const value = values.join(" / ");
          const extraClass = scoreClass(value);
          const classes = ["score-cell"];

          if (extraClass) {
            classes.push(extraClass);
          } else if (value === "") {
            classes.push("empty-score");
          }

          return `<td class="${classes.join(" ")}" title="${escapeHtml(value)}">${escapeHtml(formatScoreDisplay(value))}</td>`;
        })
        .join("");

      const preparation = [...group.preparations].sort().join(", ");
      const technicalName = [...group.technicalNames].sort().join(" / ");
      const mitoPediaLink = [...group.mitoPediaLinks][0] || "";

      return `
        <tr>
          <th scope="row" class="row-header">
            <details class="row-protocol-details">
              <summary>${escapeHtml(group.name)}</summary>
              <div class="row-protocol-meta">
                ${technicalName ? `<small class="row-tech-name">${escapeHtml(technicalName)}</small>` : ""}
                ${mitoPediaLink ? `<a class="row-meta-link" href="${escapeHtml(mitoPediaLink)}" target="_blank" rel="noreferrer">MitoPedia page</a>` : ""}
              </div>
            </details>
            <small>Average score: ${group.averageScore.toFixed(2)}</small>
            ${preparation ? `<small>${escapeHtml(preparation)}</small>` : ""}
          </th>
          ${scoreCells}
        </tr>
      `;
    })
    .join("");

  return `
    <table class="matrix-table">
      <thead>
        <tr>
          <th class="corner-header">Protocol</th>
          ${headerCells}
        </tr>
      </thead>
      <tbody>
        ${bodyRows}
      </tbody>
    </table>
  `;
}

function setMatrixMarkup(markup) {
  tableWrap.classList.remove("empty-wrap");
  tableWrap.innerHTML = markup;

  if (matrixModalWrap) {
    matrixModalWrap.classList.remove("empty-wrap");
    matrixModalWrap.innerHTML = markup;
  }
}

function setMatrixEmptyState(message, detail) {
  renderEmptyTable(message, detail);
  renderEmptyModalTable(message, detail);
  if (openMatrixModalBtn) {
    openMatrixModalBtn.disabled = true;
  }
}

function openMatrixModal() {
  if (!matrixModal || !openMatrixModalBtn || openMatrixModalBtn.disabled) {
    return;
  }

  matrixModal.hidden = false;
  document.body.classList.add("modal-open");
}

function closeMatrixModal() {
  if (!matrixModal) {
    return;
  }

  matrixModal.hidden = true;
  document.body.classList.remove("modal-open");
}

function renderTable() {
  if (!rawData.length) {
    setMatrixEmptyState("No data yet", "The workbook has not been loaded.");
    setModalSummary("Loading workbook...");
    return;
  }

  const selectedPropertyList = getPropertyColumns().filter((property) =>
    selectedProperties.has(property)
  );
  const visibleRows = getFilteredRows().filter((row) => {
    if (!selectedPropertyList.length) {
      return true;
    }

    return selectedPropertyList.some((property) => hasPositiveScore(row, property));
  });
  const groupedRows = getGroupedRows(visibleRows, selectedPropertyList);
  const rankedRows = groupedRows
    .filter((group) => group.averageScore > 0)
    .sort((a, b) => {
      if (b.averageScore !== a.averageScore) {
        return b.averageScore - a.averageScore;
      }
      return a.name.localeCompare(b.name);
    });

  resetFiltersBtn.disabled = false;

  if (!selectedPropertyList.length) {
    resultsSummary.textContent = `${groupedRows.length} protocol${groupedRows.length === 1 ? "" : "s"} match the current filters. Select one or more properties to compare.`;
    setModalSummary(resultsSummary.textContent);
    setMatrixEmptyState(
      "No properties selected",
      "Choose one or more properties to render the protocol matrix."
    );
    return;
  }

  if (!rankedRows.length) {
    resultsSummary.textContent = "No protocols match the current filters.";
    setModalSummary(resultsSummary.textContent);
    setMatrixEmptyState(
      "No matching protocols",
      "Try a different application or preparation selection."
    );
    return;
  }

  resultsSummary.textContent = `${rankedRows.length} protocol${rankedRows.length === 1 ? "" : "s"} shown, sorted by average suitability across ${selectedPropertyList.length} propert${selectedPropertyList.length === 1 ? "y" : "ies"}.`;
  setModalSummary(resultsSummary.textContent);

  setMatrixMarkup(buildTableMarkup(rankedRows, selectedPropertyList));
  if (openMatrixModalBtn) {
    openMatrixModalBtn.disabled = false;
  }
}

function setupUi() {
  const applications = uniqueValues(rawData, APPLICATION_COLUMN);

  appSelect.innerHTML = '<option value="">All applications</option>';
  applications.forEach((application) => {
    const option = document.createElement("option");
    option.value = application;
    option.textContent = application;
    appSelect.appendChild(option);
  });

  appSelect.disabled = false;
  propertySearch.disabled = false;
  if (fileStatus) {
    fileStatus.innerHTML = `Loaded <strong>${DEFAULT_WORKBOOK}</strong>.`;
  }

  renderPreparations();
  renderPropertyOptions();
  renderTable();
}

function loadWorkbook(arrayBuffer) {
  const workbook = XLSX.read(arrayBuffer, { type: "array" });
  const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
  rawData = XLSX.utils.sheet_to_json(firstSheet, { defval: "" });
  setupUi();
}

async function fetchDefaultWorkbook() {
  setLoadingState(`Loading <strong>${DEFAULT_WORKBOOK}</strong> from the local folder...`);

  try {
    const response = await fetch(DEFAULT_WORKBOOK);
    if (!response.ok) {
      throw new Error(`HTTP ${response.status}`);
    }

    const data = await response.arrayBuffer();
    loadWorkbook(data);
  } catch (error) {
    console.error("Workbook auto-load failed:", error);
    const protocolHint =
      location.protocol === "file:"
        ? " Open this page through a local web server so the browser can request the workbook."
        : "";
    if (fileStatus) {
      fileStatus.innerHTML = `Could not load <strong>${DEFAULT_WORKBOOK}</strong>.${protocolHint}`;
    }
    resultsSummary.textContent = "Workbook failed to load.";
    setModalSummary(resultsSummary.textContent);
    setMatrixEmptyState(
      "Workbook not loaded",
      `The page could not fetch ${DEFAULT_WORKBOOK}.${protocolHint.trim()}`
    );
  }
}

appSelect.addEventListener("change", syncUiAndRender);
propertySearch.addEventListener("input", () => {
  renderPropertyOptions();
  renderTable();
});

if (reloadDefaultBtn) {
  reloadDefaultBtn.addEventListener("click", fetchDefaultWorkbook);
}

if (toggleFiltersBtn) {
  toggleFiltersBtn.addEventListener("click", () => {
    const collapsed = !appLayout.classList.contains("filters-collapsed");
    setFiltersCollapsed(collapsed);
  });
}

if (openMatrixModalBtn) {
  openMatrixModalBtn.addEventListener("click", openMatrixModal);
}

if (closeMatrixModalBtn) {
  closeMatrixModalBtn.addEventListener("click", closeMatrixModal);
}

if (matrixModalBackdrop) {
  matrixModalBackdrop.addEventListener("click", closeMatrixModal);
}

document.addEventListener("keydown", (event) => {
  if (event.key === "Escape" && matrixModal && !matrixModal.hidden) {
    closeMatrixModal();
  }
});

clearPreparationsBtn.addEventListener("click", () => {
  selectedPreparations = new Set();
  syncUiAndRender();
});

selectAllVisiblePropsBtn.addEventListener("click", () => {
  filteredAvailableProperties.forEach((property) => selectedProperties.add(property));
  renderPropertyOptions();
  renderTable();
});

clearPropertiesBtn.addEventListener("click", () => {
  selectedProperties = new Set();
  renderPropertyOptions();
  renderTable();
});

resetFiltersBtn.addEventListener("click", () => {
  appSelect.value = "";
  propertySearch.value = "";
  selectedPreparations = new Set();
  selectedProperties = new Set();
  syncUiAndRender();
});

fetchDefaultWorkbook();
restoreFiltersCollapsedState();
