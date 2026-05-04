let rawData = [];
let filteredProperties = [];
let filteredAvailableProperties = [];
let selectedPreparations = new Set();
let selectedProperties = new Set();
let selectedSortProperty = "";
let displayHeaderTextByColumn = new Map();
let displayValueTextByColumn = new Map();
let applicationPropertyAvailabilityByName = new Map();
let preparationPropertyAvailabilityByName = new Map();
let propertyLinkByName = new Map();
let preparationLinkByName = new Map();

const DEFAULT_WORKBOOK = "suitbrowser_rules_new.xlsx";
const APPLICATION_COLUMN = "Applications";
const PREPARATION_COLUMN = "Preparation (to choose in the browser)";
const PROTOCOL_PREPARATION_COLUMN = "Preparation";
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
const selectAllVisiblePreparationsBtn = document.getElementById("selectAllVisiblePreparationsBtn");
const selectAllVisiblePropsBtn = document.getElementById("selectAllVisiblePropsBtn");
const clearPropertiesBtn = document.getElementById("clearPropertiesBtn");
const clearSortBtn = document.getElementById("clearSortBtn");
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

const SUBSCRIPT_CHAR_MAP = {
  0: "₀",
  1: "₁",
  2: "₂",
  3: "₃",
  4: "₄",
  5: "₅",
  6: "₆",
  7: "₇",
  8: "₈",
  9: "₉",
  "+": "₊",
  "-": "₋",
  "=": "₌",
  "(": "₍",
  ")": "₎",
};

function toSubscriptText(value) {
  return String(value).replace(/[0-9+\-=()]/g, (char) => SUBSCRIPT_CHAR_MAP[char] || char);
}

function normalizeSearchText(value) {
  return String(value)
    .toLowerCase()
    .replace(/[₀₁₂₃₄₅₆₇₈₉₊₋₌₍₎]/g, (char) => {
      const entry = Object.entries(SUBSCRIPT_CHAR_MAP).find(([, subscriptChar]) => subscriptChar === char);
      return entry ? entry[0] : char;
    });
}

function richTextHtmlToPlainText(html) {
  if (!html) {
    return "";
  }

  const template = document.createElement("template");
  template.innerHTML = html;

  const renderNode = (node, inSubscript = false) => {
    if (node.nodeType === Node.TEXT_NODE) {
      return inSubscript ? toSubscriptText(node.nodeValue ?? "") : (node.nodeValue ?? "");
    }

    if (node.nodeType !== Node.ELEMENT_NODE) {
      return "";
    }

    if (node.tagName === "BR") {
      return "\n";
    }

    const style = String(node.getAttribute("style") || "");
    const nextInSubscript =
      inSubscript ||
      node.tagName === "SUB" ||
      /vertical-align\s*:\s*(sub|subscript)/i.test(style);

    return [...node.childNodes].map((child) => renderNode(child, nextInSubscript)).join("");
  };

  return [...template.content.childNodes].map((child) => renderNode(child)).join("");
}

function cellDisplayText(cell) {
  if (cell && typeof cell.h === "string" && cell.h) {
    return richTextHtmlToPlainText(cell.h);
  }

  return String(cell?.v ?? "");
}

function getDisplayTextForColumn(column, value) {
  const normalizedValue = String(value ?? "");
  const text = displayValueTextByColumn.get(column)?.get(normalizedValue);

  return text ?? normalizedValue;
}

function getDisplayTextForHeader(column) {
  return displayHeaderTextByColumn.get(column) ?? column;
}

function renderDisplayList(values, column, separator = ", ") {
  return values.map((value) => getDisplayTextForColumn(column, value)).join(separator);
}

function normalizeRuleKey(value) {
  return String(value ?? "")
    .replace(/\u00a0/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function getPropertyRuleKeys(property) {
  return [
    property,
    normalizeRuleKey(property),
    getDisplayTextForHeader(property),
    normalizeRuleKey(getDisplayTextForHeader(property)),
  ].filter((value, index, values) => value && values.indexOf(value) === index);
}

function getPropertyAvailabilityValue(availability, property) {
  for (const key of getPropertyRuleKeys(property)) {
    if (availability.has(key)) {
      return availability.get(key);
    }
  }

  return undefined;
}

function buildDisplayMaps(sheet, rows) {
  displayHeaderTextByColumn = new Map();
  displayValueTextByColumn = new Map();

  if (!rows.length || !sheet || !sheet["!ref"]) {
    return;
  }

  const range = XLSX.utils.decode_range(sheet["!ref"]);
  const headers = Object.keys(rows[0]);

  headers.forEach((header, colIndex) => {
    const headerCell = sheet[XLSX.utils.encode_cell({ r: range.s.r, c: colIndex })];
    displayHeaderTextByColumn.set(header, cellDisplayText(headerCell));
    displayValueTextByColumn.set(header, new Map());
  });

  rows.forEach((row, rowIndex) => {
    const sheetRow = range.s.r + 1 + rowIndex;

    headers.forEach((header, colIndex) => {
      const cell = sheet[XLSX.utils.encode_cell({ r: sheetRow, c: colIndex })];
      const normalizedValue = String(row[header] ?? "");
      const text = cellDisplayText(cell);

      if (text !== normalizedValue) {
        displayValueTextByColumn.get(header).set(normalizedValue, text);
      }
    });
  });
}

function buildApplicationPropertyAvailability(sheet) {
  applicationPropertyAvailabilityByName = new Map();

  if (!sheet || !sheet["!ref"]) {
    return;
  }

  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
  if (rows.length < 2) {
    return;
  }

  const headers = rows[0].slice(1).map((value) => normalizeRuleKey(value));

  rows.slice(1).forEach((row) => {
    const application = normalizeRuleKey(row[0]);
    if (!application) {
      return;
    }

    const propertyAvailability = new Map();
    headers.forEach((property, index) => {
      if (!property) {
        return;
      }

      propertyAvailability.set(property, String(row[index + 1] ?? "").trim() === "1");
    });

    applicationPropertyAvailabilityByName.set(application, propertyAvailability);
  });
}

function buildPreparationPropertyAvailability(sheet) {
  preparationPropertyAvailabilityByName = new Map();

  if (!sheet || !sheet["!ref"]) {
    return;
  }

  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
  if (rows.length < 2) {
    return;
  }

  const headers = rows[0].slice(1).map((value) => normalizeRuleKey(value));

  rows.slice(1).forEach((row) => {
    const preparation = normalizeRuleKey(row[0]);
    if (!preparation) {
      return;
    }

    const propertyAvailability = new Map();
    headers.forEach((property, index) => {
      if (!property) {
        return;
      }

      propertyAvailability.set(property, String(row[index + 1] ?? "").trim() === "1");
    });

    preparationPropertyAvailabilityByName.set(preparation, propertyAvailability);
  });
}

function buildPropertyLinkMap(sheet) {
  propertyLinkByName = new Map();

  if (!sheet || !sheet["!ref"]) {
    return;
  }

  const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
  rows.forEach((row) => {
    const property = String(row.Properties ?? row["Properties"] ?? "").trim();
    const link = sanitizeUrl(row["MitoPedia link"] ?? row["MitoPedia Link"]);

    if (property && link) {
      propertyLinkByName.set(property, link);
    }
  });
}

function buildPreparationLinkMap(sheet) {
  preparationLinkByName = new Map();

  if (!sheet || !sheet["!ref"]) {
    return;
  }

  const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
  rows.forEach((row) => {
    const preparation = String(row.Preparation ?? row["Preparation"] ?? "").trim();
    const link = sanitizeUrl(row["MitoPedia Link"] ?? row["MitoPedia link"]);

    if (preparation && link) {
      preparationLinkByName.set(preparation, link);
    }
  });
}

function getAvailablePropertiesForApplication(application, propertyColumns) {
  const availability = applicationPropertyAvailabilityByName.get(normalizeRuleKey(application));
  if (!availability) {
    return null;
  }

  return new Set(
    propertyColumns.filter((property) => getPropertyAvailabilityValue(availability, property) === true)
  );
}

function getAvailablePropertiesForPreparations(preparations, propertyColumns) {
  const selectedPreparationList = [...preparations];
  if (!selectedPreparationList.length) {
    return null;
  }

  const availabilityRules = selectedPreparationList
    .map(
      (preparation) =>
        preparationPropertyAvailabilityByName.get(normalizeRuleKey(preparation))
    )
    .filter(Boolean);

  if (!availabilityRules.length) {
    return null;
  }

  return new Set(
    propertyColumns.filter((property) =>
      availabilityRules.some(
        (availability) => getPropertyAvailabilityValue(availability, property) === true
      )
    )
  );
}

function intersectPropertyAvailability(propertyColumns, availabilitySets) {
  const activeSets = availabilitySets.filter(Boolean);
  if (!activeSets.length) {
    return new Set(propertyColumns);
  }

  return new Set(
    propertyColumns.filter((property) =>
      activeSets.every((availabilitySet) => availabilitySet.has(property))
    )
  );
}

function getPropertyLink(property) {
  return propertyLinkByName.get(property) || propertyLinkByName.get(String(property).trim()) || "";
}

function getPreparationLink(preparation) {
  return (
    preparationLinkByName.get(preparation) ||
    preparationLinkByName.get(String(preparation).trim()) ||
    ""
  );
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
  if (selectAllVisiblePreparationsBtn) {
    selectAllVisiblePreparationsBtn.disabled = true;
  }
  if (clearPreparationsBtn) {
    clearPreparationsBtn.disabled = true;
  }
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
    if (selectAllVisiblePreparationsBtn) {
      selectAllVisiblePreparationsBtn.disabled = true;
    }
    if (clearPreparationsBtn) {
      clearPreparationsBtn.disabled = true;
    }
    preparationOptions.textContent = "No preparations available.";
    return;
  }

  if (selectAllVisiblePreparationsBtn) {
    selectAllVisiblePreparationsBtn.disabled = selectedPreparations.size === allPreparations.length;
  }
  if (clearPreparationsBtn) {
    clearPreparationsBtn.disabled = !selectedPreparations.size;
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

    const prepText = document.createElement("span");
    prepText.className = "preparation-chip-text";
    prepText.textContent = getDisplayTextForColumn(PREPARATION_COLUMN, preparation);

    const preparationLink = getPreparationLink(preparation);
    span.appendChild(prepText);

    if (preparationLink) {
      const infoBadge = document.createElement("a");
      infoBadge.className = "preparation-info-badge";
      infoBadge.textContent = "i";
      infoBadge.title = `Open MitoPedia page for ${getDisplayTextForColumn(PREPARATION_COLUMN, preparation)}`;
      infoBadge.href = preparationLink;
      infoBadge.target = "_blank";
      infoBadge.rel = "noreferrer";
      infoBadge.setAttribute(
        "aria-label",
        `Open MitoPedia page for ${getDisplayTextForColumn(PREPARATION_COLUMN, preparation)}`
      );
      span.appendChild(infoBadge);
    }

    label.append(input, span);
    preparationOptions.appendChild(label);
  });
}

function renderPropertyOptions() {
  const searchValue = normalizeSearchText(propertySearch.value.trim());
  const selectedApplication = appSelect.value;
  const propertyRows = getRowsForPropertyOptions();
  const propertyColumns = getPropertyColumns();
  const baseAvailableProperties =
    getAvailablePropertiesForApplication(selectedApplication, propertyColumns) ||
    new Set(getAvailablePropertyColumns(propertyRows));
  const preparationAvailableProperties = getAvailablePropertiesForPreparations(
    selectedPreparations,
    propertyColumns
  );
  const availableProperties = intersectPropertyAvailability(propertyColumns, [
    baseAvailableProperties,
    preparationAvailableProperties,
  ]);
  filteredProperties = propertyColumns.filter((property) =>
    normalizeSearchText(property).includes(searchValue)
  );
  filteredAvailableProperties = filteredProperties.filter((property) =>
    availableProperties.has(property)
  );
  selectedProperties = new Set(
    [...selectedProperties].filter((property) => availableProperties.has(property))
  );

  propertyOptions.innerHTML = "";

  if (!filteredProperties.length) {
    propertyOptions.textContent = propertyColumns.length
      ? "No research questions match this search."
      : "No research questions available.";
    return;
  }

  filteredProperties.forEach((property, index) => {
    const option = document.createElement("div");
    option.className = "property-option";
    const isAvailable = availableProperties.has(property);
    const isUnavailableForPreparation =
      preparationAvailableProperties && !preparationAvailableProperties.has(property);
    const isUnavailableForApplication = !baseAvailableProperties.has(property);

    if (!isAvailable) {
      option.classList.add("unavailable");
    }

    const propertySlug = property
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, "-")
      .replace(/^-+|-+$/g, "") || "property";
    const optionId = `property-option-${propertySlug}-${index}`;
    const input = document.createElement("input");
    input.type = "checkbox";
    input.id = optionId;
    input.value = property;
    input.checked = selectedProperties.has(property);
    input.disabled = !isAvailable;
    input.addEventListener("change", () => {
      if (!isAvailable) {
        input.checked = false;
        selectedProperties.delete(property);
        return;
      }

      if (input.checked) {
        selectedProperties.add(property);
      } else {
        selectedProperties.delete(property);
      }
      renderTable();
    });

    const shell = document.createElement("div");
    shell.className = "property-card-shell";

    const label = document.createElement("label");
    label.className = "property-card";
    label.htmlFor = optionId;
    label.title = isAvailable
      ? property
      : isUnavailableForPreparation
        ? `${property} is unavailable for the selected preparation`
        : selectedApplication && isUnavailableForApplication
        ? `${property} is unavailable for the selected application`
        : `${property} is unavailable for the current filters`;

    const labelText = document.createElement("span");
    labelText.className = "property-card-label";
    labelText.textContent = getDisplayTextForHeader(property);

    const propertyLink = getPropertyLink(property);

    if (propertyLink) {
      shell.classList.add("has-info");
      const infoBadge = document.createElement("a");
      infoBadge.className = "property-info-badge";
      infoBadge.textContent = "i";
      infoBadge.title = `Open MitoPedia page for ${property}`;
      infoBadge.href = propertyLink;
      infoBadge.target = "_blank";
      infoBadge.rel = "noreferrer";
      infoBadge.setAttribute("aria-label", `Open MitoPedia page for ${property}`);
      shell.append(label, infoBadge);
    } else {
      shell.appendChild(label);
    }

    label.appendChild(labelText);
    option.append(input, shell);
    propertyOptions.appendChild(option);
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
  toggleFiltersBtn.setAttribute("aria-label", collapsed ? "Show filters" : "Hide filters");
  toggleFiltersBtn.setAttribute("title", collapsed ? "Show filters" : "Hide filters");

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
        protocolPreparations: new Set(),
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

    const protocolPreparation = String(row[PROTOCOL_PREPARATION_COLUMN] ?? "").trim();
    if (protocolPreparation) {
      group.protocolPreparations.add(protocolPreparation);
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

    const totalScore = numericScores.reduce((sum, score) => sum + score, 0);
    const averageScore = numericScores.length ? totalScore / numericScores.length : 0;

    return {
      ...group,
      totalScore,
      averageScore,
    };
  });
}

function getSortScore(group, property) {
  if (!property) {
    return group.totalScore;
  }

  const score = group.propertyScores.get(property);
  return score === undefined ? -1 : score;
}

function updateSortResetControl(selectedPropertyList) {
  if (selectedSortProperty && !selectedPropertyList.includes(selectedSortProperty)) {
    selectedSortProperty = "";
  }

  if (clearSortBtn) {
    clearSortBtn.disabled = !selectedSortProperty;
  }
}

function buildTableMarkup(rankedRows, selectedPropertyList) {
  const headerCells = selectedPropertyList
    .map((property) => {
      const isActiveSort = selectedSortProperty === property;
      return `
        <th class="diag-header${isActiveSort ? " active-sort" : ""}" aria-sort="${isActiveSort ? "descending" : "none"}">
          <button class="property-sort-btn" type="button" data-sort-property="${escapeHtml(property)}" title="Sort protocols by ${escapeHtml(property)}">
            <span>${escapeHtml(getDisplayTextForHeader(property))}</span>
            <small>${isActiveSort ? "Sorted" : "Sort"}</small>
          </button>
        </th>`;
    })
    .join("");

  const bodyRows = rankedRows
    .map((group) => {
      const technicalNames = [...group.technicalNames].sort();
      const browserPreparations = [...group.preparations].sort();
      const protocolPreparations = [...group.protocolPreparations].sort();
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

      const mitoPediaLink = [...group.mitoPediaLinks][0] || "";
      const preparationValues = [];

      if (protocolPreparations.length) {
        preparationValues.push({
          column: PROTOCOL_PREPARATION_COLUMN,
          values: protocolPreparations,
        });
      }

      if (
        browserPreparations.length &&
        browserPreparations.join("\u0000") !== protocolPreparations.join("\u0000")
      ) {
        preparationValues.push({
          column: PREPARATION_COLUMN,
          values: browserPreparations,
        });
      }

      const preparationLines = preparationValues
        .map(({ column, values }) => `<small>${escapeHtml(renderDisplayList(values, column))}</small>`)
        .join("");

      return `
        <tr>
          <th scope="row" class="row-header">
            <details class="row-protocol-details">
              <summary>${escapeHtml(getDisplayTextForColumn(NAME_COLUMN, group.name))}</summary>
              <div class="row-protocol-meta">
                ${technicalNames.length ? `<small class="row-tech-name">${escapeHtml(technicalNames.map((value) => getDisplayTextForColumn(TECHNICAL_NAME_COLUMN, value)).join(" / "))}</small>` : ""}
                ${mitoPediaLink ? `<a class="row-meta-link" href="${escapeHtml(mitoPediaLink)}" target="_blank" rel="noreferrer">MitoPedia page</a>` : ""}
              </div>
            </details>
            <small>Total score: ${group.totalScore}</small>
            <small>Average score: ${group.averageScore.toFixed(2)}</small>
            ${preparationLines}
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
  updateSortResetControl(selectedPropertyList);
  const visibleRows = getFilteredRows().filter((row) => {
    if (!selectedPropertyList.length) {
      return true;
    }

    return selectedPropertyList.some((property) => hasPositiveScore(row, property));
  });
  const groupedRows = getGroupedRows(visibleRows, selectedPropertyList);
  const rankedRows = groupedRows
    .filter((group) => group.totalScore > 0)
    .sort((a, b) => {
      const sortScoreA = getSortScore(a, selectedSortProperty);
      const sortScoreB = getSortScore(b, selectedSortProperty);

      if (sortScoreB !== sortScoreA) {
        return sortScoreB - sortScoreA;
      }

      if (b.totalScore !== a.totalScore) {
        return b.totalScore - a.totalScore;
      }

      return a.name.localeCompare(b.name);
    });

  if (!selectedPropertyList.length) {
    selectedSortProperty = "";
    updateSortResetControl(selectedPropertyList);
    resultsSummary.textContent = `${groupedRows.length} protocol${groupedRows.length === 1 ? "" : "s"} match the current filters. Select one or more research questions to compare.`;
    setModalSummary(resultsSummary.textContent);
    setMatrixEmptyState(
      "No research questions selected",
      "Choose one or more research questions to render the protocol matrix."
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

  const sortDescription = selectedSortProperty
    ? `sorted by ${selectedSortProperty}, then total applicability score`
    : `sorted by total applicability score across ${selectedPropertyList.length} research question${selectedPropertyList.length === 1 ? "" : "s"}`;
  resultsSummary.textContent = `${rankedRows.length} protocol${rankedRows.length === 1 ? "" : "s"} shown, ${sortDescription}.`;
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
    option.textContent = getDisplayTextForColumn(APPLICATION_COLUMN, application);
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
  const workbook = XLSX.read(arrayBuffer, { type: "array", cellHTML: true });
  const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
  const applicationRulesSheet = workbook.Sheets["Rules properties & applications"];
  const propertyLinksSheet = workbook.Sheets["Links properties"];
  const preparationLinksSheet = workbook.Sheets["Links preparations"];
  const preparationRulesSheet = workbook.Sheets["Rules properties & preparation"];
  rawData = XLSX.utils.sheet_to_json(firstSheet, { defval: "" });
  buildDisplayMaps(firstSheet, rawData);
  buildApplicationPropertyAvailability(applicationRulesSheet);
  buildPreparationPropertyAvailability(preparationRulesSheet);
  buildPropertyLinkMap(propertyLinksSheet);
  buildPreparationLinkMap(preparationLinksSheet);
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

function handleTableSortClick(event) {
  const sortButton = event.target.closest("[data-sort-property]");
  if (!sortButton) {
    return;
  }

  selectedSortProperty = sortButton.dataset.sortProperty || "";
  renderTable();
}

if (tableWrap) {
  tableWrap.addEventListener("click", handleTableSortClick);
}

if (matrixModalWrap) {
  matrixModalWrap.addEventListener("click", handleTableSortClick);
}

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

if (clearSortBtn) {
  clearSortBtn.addEventListener("click", () => {
    selectedSortProperty = "";
    renderTable();
  });
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

if (selectAllVisiblePreparationsBtn) {
  selectAllVisiblePreparationsBtn.addEventListener("click", () => {
    uniqueValues(getRowsForPreparationOptions(), PREPARATION_COLUMN).forEach((preparation) =>
      selectedPreparations.add(preparation)
    );
    syncUiAndRender();
  });
}

if (clearPreparationsBtn) {
  clearPreparationsBtn.addEventListener("click", () => {
    selectedPreparations = new Set();
    syncUiAndRender();
  });
}

selectAllVisiblePropsBtn.addEventListener("click", () => {
  filteredAvailableProperties.forEach((property) => selectedProperties.add(property));
  renderPropertyOptions();
  renderTable();
});

clearPropertiesBtn.addEventListener("click", () => {
  selectedProperties = new Set();
  selectedSortProperty = "";
  renderPropertyOptions();
  renderTable();
});

fetchDefaultWorkbook();
restoreFiltersCollapsedState();
