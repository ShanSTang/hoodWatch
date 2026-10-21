/* =========================================================
   HoodWatch — app.js
   Browser-only laboratory hood ACH analytics
   ========================================================= */

"use strict";

/* ---------------------------------------------------------
   State
   --------------------------------------------------------- */

let allRecords = [];
let filteredRecords = [];
let summary = [];
let selectedHoodKey = "";


/* ---------------------------------------------------------
   DOM
   --------------------------------------------------------- */

const fileInput = document.getElementById("fileInput");
const dropZone = document.getElementById("dropZone");
const fileStatus = document.getElementById("fileStatus");
const recordInfo = document.getElementById("recordInfo");
const clearBtn = document.getElementById("clearBtn");
const errorMessage = document.getElementById("errorMessage");

const buildingFilter = document.getElementById("buildingFilter");
const labFilter = document.getElementById("labFilter");
const startDate = document.getElementById("startDate");
const endDate = document.getElementById("endDate");
const resetFiltersBtn = document.getElementById("resetFiltersBtn");
const hoodFilter = document.getElementById("hoodFilter");

const exportBtn = document.getElementById("exportBtn");

const kpiHoods = document.getElementById("kpiHoods");
const kpiFlagged = document.getElementById("kpiFlagged");
const kpiBuildings = document.getElementById("kpiBuildings");
const kpiLabs = document.getElementById("kpiLabs");
const kpiVAVFailing = document.getElementById("kpiVAVFailing");

const flagCount = document.getElementById("flagCount");
const hoodList = document.getElementById("hoodList");

const trendChart = document.getElementById("trendChart");
const rankingChart = document.getElementById("rankingChart");
const heatmapChart = document.getElementById("heatmapChart");

const resultsBody = document.getElementById("resultsBody");
const tableCount = document.getElementById("tableCount");

const heatmapMetric = document.getElementById("heatmapMetric");
const heatmapTitle = document.getElementById("heatmapTitle");

/* ---------------------------------------------------------
   Required columns
   --------------------------------------------------------- */

const REQUIRED_COLUMNS = [
  "Building ID",
  "Lab ID",
  "Hood ID",
  "Date",
  "ACH",
   "Face velocity"
];


/* ---------------------------------------------------------
   Utility functions
   --------------------------------------------------------- */

function normalizeValue(value, fallback = "") {
  if (value === null || value === undefined || value === "") {
    return fallback;
  }

  return String(value).trim();
}


function parseACH(value) {
  if (value === null || value === undefined || value === "") {
    return NaN;
  }

  if (typeof value === "number") {
    return Number.isFinite(value) ? value : NaN;
  }

  const cleaned = String(value)
    .replace(/,/g, "")
    .replace(/[^\d.+\-eE]/g, "");

  const number = Number(cleaned);

  return Number.isFinite(number) ? number : NaN;
}


/*
 * Convert dates to a consistent YYYY-MM-DD representation
 * for filtering and sorting.
 */
function parseDate(value) {
  if (value instanceof Date && !isNaN(value)) {
    return new Date(
      value.getFullYear(),
      value.getMonth(),
      value.getDate()
    );
  }

  if (typeof value === "number") {
    // Excel serial date
    const parsed = XLSX.SSF.parse_date_code(value);

    if (parsed) {
      return new Date(
        parsed.y,
        parsed.m - 1,
        parsed.d
      );
    }
  }

  const text = normalizeValue(value);

  if (!text) {
    return null;
  }

  // ISO-ish date
  const iso = new Date(text);

  if (!isNaN(iso.getTime())) {
    return new Date(
      iso.getFullYear(),
      iso.getMonth(),
      iso.getDate()
    );
  }

  return null;
}


function dateKey(date) {
  if (!date || isNaN(date.getTime())) {
    return "";
  }

  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");

  return `${year}-${month}-${day}`;
}


function formatDate(date) {
  if (!date || isNaN(date.getTime())) {
    return "—";
  }

  return date.toLocaleDateString(undefined, {
    year: "numeric",
    month: "short",
    day: "numeric"
  });
}


function formatNumber(value, decimals = 1) {
  if (value === null || value === undefined || !Number.isFinite(value)) {
    return "—";
  }

  return Number(value).toLocaleString(undefined, {
    minimumFractionDigits: decimals,
    maximumFractionDigits: decimals
  });
}


function formatPercent(value) {
  if (!Number.isFinite(value)) {
    return "—";
  }

  return `${formatNumber(value, 1)}%`;
}


function escapeHTML(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}


/* ---------------------------------------------------------
   File handling
   --------------------------------------------------------- */

fileInput.addEventListener("change", event => {
  const file = event.target.files[0];

  if (file) {
    processFile(file);
  }
});


dropZone.addEventListener("dragover", event => {
  event.preventDefault();
  dropZone.classList.add("dragging");
});


dropZone.addEventListener("dragleave", () => {
  dropZone.classList.remove("dragging");
});


dropZone.addEventListener("drop", event => {
  event.preventDefault();
  dropZone.classList.remove("dragging");

  const file = event.dataTransfer.files[0];

  if (file) {
    processFile(file);
  }
});


async function processFile(file) {
  clearError();

  try {
    const arrayBuffer = await file.arrayBuffer();

    const workbook = XLSX.read(arrayBuffer, {
      type: "array",
      cellDates: true
    });

    if (!workbook.SheetNames.length) {
      throw new Error("The workbook contains no worksheets.");
    }

    const worksheet = workbook.Sheets[workbook.SheetNames[0]];

    const rows = XLSX.utils.sheet_to_json(worksheet, {
      defval: null,
      raw: true
    });

    if (!rows.length) {
      throw new Error("The selected file contains no data rows.");
    }

    validateColumns(rows);

    const parsed = parseRows(rows);

    if (!parsed.length) {
      throw new Error(
        "No valid inspection records were found. Check the Date and ACH columns."
      );
    }

    allRecords = parsed;

    initializeFilters();
    applyFilters();

    fileStatus.textContent = file.name;
    recordInfo.textContent =
      `${allRecords.length.toLocaleString()} inspection records loaded`;

    clearBtn.disabled = false;
    exportBtn.disabled = false;
    resetFiltersBtn.disabled = false;

    enableControls(true);

  } catch (error) {
    console.error(error);
    showError(error.message || "Unable to read the file.");
  }
}


/* ---------------------------------------------------------
   Validation / parsing
   --------------------------------------------------------- */

function validateColumns(rows) {
  const columns = Object.keys(rows[0]);

  const missing = REQUIRED_COLUMNS.filter(
    column => !columns.includes(column)
  );

  if (missing.length) {
    throw new Error(
      `Missing required column${missing.length > 1 ? "s" : ""}: ` +
      missing.join(", ")
    );
  }
}


function parseRows(rows) {
  const parsed = [];

  rows.forEach((row, index) => {
    const building = normalizeValue(
      row["Building ID"],
      "Unknown Building"
    );

    const lab = normalizeValue(
      row["Lab ID"],
      "Unknown Lab"
    );

    const hood = normalizeValue(
      row["Hood ID"],
      "Unknown Hood"
    );

    const date = parseDate(row["Date"]);
    const ach = parseACH(row["ACH"]);
    const faceVelocity = parseACH(row["Face velocity"]); 

    if (!date || !Number.isFinite(ach)) {
      return;
    }

    /*
     * Equivalent to:
     *
     * df["Hood Key"] =
     * Building ID + " | " + Lab ID + " | " + Hood ID
     */
    const hoodKey =
      `${building} | ${lab} | ${hood}`;

    parsed.push({
      rowNumber: index + 2,
      building,
      lab,
      hood,
      date,
      dateString: dateKey(date),
      ach,
      faceVelocity,
      hoodKey
    });
  });

  return parsed;
}


/* ---------------------------------------------------------
   Filtering
   --------------------------------------------------------- */

function initializeFilters() {
  const buildings = unique(
    allRecords.map(record => record.building)
  ).sort(naturalSort);

  buildingFilter.innerHTML =
    `<option value="">All buildings</option>` +
    buildings
      .map(
        building =>
          `<option value="${escapeHTML(building)}">
             ${escapeHTML(building)}
           </option>`
      )
      .join("");

  labFilter.innerHTML =
    `<option value="">All labs</option>`;

  updateLabOptions();

  const dates = allRecords
  .map(record => record.date)
  .filter(date =>
    date instanceof Date &&
    Number.isFinite(date.getTime())
  )
  .sort((a, b) => a - b);

if (dates.length) {
  const minDate = dateKey(dates[0]);
  const maxDate = dateKey(dates[dates.length - 1]);

  startDate.min = minDate;
  startDate.max = maxDate;
  endDate.min = minDate;
  endDate.max = maxDate;

  startDate.value = minDate;
  endDate.value = maxDate;
} else {
  startDate.value = "";
  endDate.value = "";
  startDate.removeAttribute("min");
  startDate.removeAttribute("max");
  endDate.removeAttribute("min");
  endDate.removeAttribute("max");
}
}


function updateLabOptions() {
  const selectedBuilding = buildingFilter.value;

  let records = allRecords;

  if (selectedBuilding) {
    records = records.filter(
      record => record.building === selectedBuilding
    );
  }

  const labs = unique(
    records.map(record => record.lab)
  ).sort(naturalSort);

  const previous = labFilter.value;

  labFilter.innerHTML =
    `<option value="">All labs</option>` +
    labs
      .map(
        lab =>
          `<option value="${escapeHTML(lab)}">
             ${escapeHTML(lab)}
           </option>`
      )
      .join("");

  if (labs.includes(previous)) {
    labFilter.value = previous;
  }
}


function applyFilters() {
  const building = buildingFilter.value || "";
  const lab = labFilter.value || "";

  const start = startDate.value
    ? new Date(`${startDate.value}T00:00:00`)
    : null;

  const end = endDate.value
    ? new Date(`${endDate.value}T23:59:59.999`)
    : null;

  // No data? Keep the dashboard in a safe empty state.
  if (!allRecords.length) {
    filteredRecords = [];
    summary = [];
    updateAllVisuals();
    return;
  }

  filteredRecords = allRecords.filter(record => {
    if (building && record.building !== building) {
      return false;
    }

    if (lab && record.lab !== lab) {
      return false;
    }

    if (start && record.date < start) {
      return false;
    }

    if (end && record.date > end) {
      return false;
    }

    return true;
  });

  summary = calculateSummary(filteredRecords);

  // Clear a hood selection if it no longer exists in the filtered data.
  if (!summary.some(row => row.hoodKey === selectedHoodKey)) {
    selectedHoodKey = "";
    hoodFilter.value = "";
  }

  updateAllVisuals();
}


/* ---------------------------------------------------------
   Core pandas-equivalent analysis
   --------------------------------------------------------- */

function calculateSummary(records) {

  /*
   * Equivalent to:
   *
   * df.sort_values(["Hood Key", "Date"])
   */
  const sorted = [...records].sort((a, b) => {
    const hoodCompare = naturalSort(a.hoodKey, b.hoodKey);

    if (hoodCompare !== 0) {
      return hoodCompare;
    }

    return a.date - b.date;
  });


  /*
   * Equivalent to:
   *
   * df.groupby("Hood Key")["ACH"].shift(1)
   * df.groupby("Hood Key")["ACH"].transform("first")
   */
  const groups = new Map();

  sorted.forEach(record => {
    if (!groups.has(record.hoodKey)) {
      groups.set(record.hoodKey, []);
    }

    groups.get(record.hoodKey).push(record);
  });


  const enriched = [];

  for (const [hoodKey, hoodRecords] of groups) {

    const firstACH = hoodRecords[0].ach;

    hoodRecords.forEach((record, index) => {

      const previousACH =
        index > 0
          ? hoodRecords[index - 1].ach
          : null;

      /*
       * Positive numbers mean ACH worsened,
       * matching the Python code.
       */
      const dropPrevious =
        previousACH !== null
          ? record.ach - previousACH
          : null;

      const dropBaseline =
        record.ach - firstACH;


      const dropPercentPrevious =
        previousACH !== null && previousACH > 0
          ? 100 * (record.ach - previousACH) / previousACH
          : null;
      
      const dropPercentBaseline =
        firstACH > 0
          ? 100 * (record.ach - firstACH) / firstACH
          : null;

      enriched.push({
        ...record,

        previousACH,
        firstACH,

        achDropPrevious: dropPrevious,
        achDropBaseline: dropBaseline,

        achDropPercentPrevious: dropPercentPrevious,
        achDropPercentBaseline: dropPercentBaseline
      });
    });
  }


  /*
   * Latest record for each hood.
   *
   * Equivalent to:
   *
   * groupby("Hood Key").tail(1)
   */
  const latestRecords = [];

  for (const hoodRecords of groups.values()) {
    latestRecords.push(
      hoodRecords[hoodRecords.length - 1]
    );
  }


  const results = latestRecords.map(latest => {

    const hoodRecords = groups.get(latest.hoodKey);

    const first = hoodRecords[0];

    const enrichedLatest = enriched.find(
      record =>
        record.hoodKey === latest.hoodKey &&
        record.rowNumber === latest.rowNumber
    );

    const inspectionCount = hoodRecords.length;

    return {
      hoodKey: latest.hoodKey,

      hood: latest.hood,
      building: latest.building,
      lab: latest.lab,

      firstInspection: first.date,
      latestInspection: latest.date,

      baselineACH: first.ach,
      latestACH: latest.ach,

      previousACH:
        enrichedLatest?.previousACH ?? null,

      achDropPrevious:
        enrichedLatest?.achDropPrevious ?? 0,

      achDropPercentPrevious:
        enrichedLatest?.achDropPercentPrevious ?? null,

      achDropBaseline:
        enrichedLatest?.achDropBaseline ?? 0,

      achDropPercentBaseline:
        enrichedLatest?.achDropPercentBaseline ?? null,

      inspectionCount,

      /*
       * Equivalent to:
       *
       * summary["Trend Available"] =
       * summary["Inspection Count"] >= 2
       */
      trendAvailable: inspectionCount >= 2
    };
  });


  /*
   * Worst deterioration first.
   *
   * Equivalent to:
   *
   * sort_values(
   *   "ACH Drop vs Baseline",
   *   ascending=False
   * )
   */
  return results.sort(
    (a, b) => b.latestACH - a.latestACH
  );
}


/* ---------------------------------------------------------
   Deterioration rules
   --------------------------------------------------------- */

function getDeteriorationStatus(row) {

  if (!row.trendAvailable) {
    return {
      flagged: false,
      reasons: ["Insufficient inspection history"]
    };
  }

  const reasons = [];

  /*
   * Rule 01:
   * ACH increased >= 100 ACH
   */
  if (
    row.achDropBaseline >= 100
  ) {
    reasons.push(
      `ACH increased ≥ 100 ACH`
    );
  }


  /*
   * Rule 02:
   *
   * Baseline is 200–400 inclusive AND
   * latest is outside that range.
   */
  if (
    row.baselineACH >= 200 &&
    row.baselineACH <= 400 &&
    (
      row.latestACH < 200 ||
      row.latestACH > 400
    )
  ) {
    reasons.push(
      `Crossed the 200–400 ACH range`
    );
  }


  return {
    flagged: reasons.length > 0,
    reasons
  };
}
/*---------------------------------------------------------
   Get VAV failure Counts
   ---------------------------------------------------------
*/

function getVAVFailureStatus(row) {
  if (!Number.isFinite(row.latestACH)) {
    return {
      failing: false,
      reason: "No valid latest ACH measurement"
    };
  }

  const failing = row.latestACH < 200 || row.latestACH > 400;

  return {
    failing,
    reason: failing
      ? "Latest ACH outside the 200–400 range"
      : "Latest ACH within the acceptable range"
  };
}

/* ---------------------------------------------------------
   Update dashboard
   --------------------------------------------------------- */

function updateAllVisuals() {
  updateKPIs();
  updateFlaggedHoods();
  updateHoodSelector();
  updateTable();
  updateTrendChart();
  updateFaceVelocityChart();
  updateHeatmapChart();
}

function updateHeatmap() {
  updateHeatmapChart();
}


/* ---------------------------------------------------------
   KPIs
   --------------------------------------------------------- */

function updateKPIs() {

  const flagged = summary.filter(
    row => getDeteriorationStatus(row).flagged
  );

  const failingVAVSystems = new Set(
     summary
        .filter(row=> getVAVFailureStatus(row).failing)
        .map(row => row.hoodKey)
  );
  const buildings = new Set(
    summary.map(row => row.building)
  );

  const labs = new Set(
    summary.map(row => row.lab)
  );

  kpiHoods.textContent =
    summary.length.toLocaleString();

  kpiFlagged.textContent =
    flagged.length.toLocaleString();
   kpiVAVFailing.textContent=
      failingVAVSystems.size.toLocaleString();

  kpiBuildings.textContent =
    buildings.size.toLocaleString();

  kpiLabs.textContent =
    labs.size.toLocaleString();

  flagCount.textContent =
    `${flagged.length} flagged`;
}


/* ---------------------------------------------------------
   Flagged hood list
   --------------------------------------------------------- */

function updateFlaggedHoods() {

  const flagged = summary.filter(
    row => getDeteriorationStatus(row).flagged
  );

  if (!flagged.length) {
    hoodList.innerHTML = `
      <div class="empty-state">
        No hoods currently meet the deterioration rules.
      </div>
    `;
    return;
  }


  hoodList.innerHTML = flagged.map(row => {

    const status = getDeteriorationStatus(row);

    return `
      <button
        class="hood-item"
        data-hood-key="${escapeHTML(row.hoodKey)}"
        type="button"
      >
        <div class="hood-item-main">
          <strong>${escapeHTML(row.hood)}</strong>
          <span>
            ${escapeHTML(row.building)} ·
            ${escapeHTML(row.lab)}
          </span>
        </div>

        <div class="hood-item-metric">
          <strong>
            ${formatNumber(row.achDropBaseline, 1)} ACH
          </strong>
          <span>
            ${formatPercent(row.achDropPercentBaseline)}
          </span>
        </div>

        <div class="hood-item-reason">
          ${status.reasons
            .map(reason => escapeHTML(reason))
            .join(" · ")}
        </div>
      </button>
    `;

  }).join("");


  hoodList.querySelectorAll(".hood-item")
    .forEach(button => {

      button.addEventListener("click", () => {

        selectedHoodKey =
          button.dataset.hoodKey;

        hoodFilter.value =
          selectedHoodKey;

        updateTrendChart();
        updateFaceVelocityChart();

        trendChart.scrollIntoView({
          behavior: "smooth",
          block: "center"
        });
      });
    });
}


/* ---------------------------------------------------------
   Hood selector
   --------------------------------------------------------- */

function updateHoodSelector() {

  const previous = selectedHoodKey;

  const sorted = [...summary].sort(
    (a, b) => naturalSort(a.hoodKey, b.hoodKey)
  );

  hoodFilter.innerHTML =
    `<option value="">Select a hood</option>` +
    sorted.map(row => `
      <option value="${escapeHTML(row.hoodKey)}">
        ${escapeHTML(row.hoodKey)}
      </option>
    `).join("");


  if (
    previous &&
    sorted.some(row => row.hoodKey === previous)
  ) {
    hoodFilter.value = previous;
  } else {
    selectedHoodKey = "";
  }
}


/* ---------------------------------------------------------
   Inspection table
   --------------------------------------------------------- */

function updateTable() {

  tableCount.textContent =
    `${summary.length} hoods`;

  if (!summary.length) {
    resultsBody.innerHTML = `
      <tr>
        <td colspan="10" class="empty-cell">
          No inspection records match the current filters.
        </td>
      </tr>
    `;

    return;
  }


  resultsBody.innerHTML = [...summary]
  .sort((a, b) => b.latestACH - a.latestACH)
  .map(row => {

    const status =
      getDeteriorationStatus(row);

    let statusHTML;

    if (!row.trendAvailable) {
      statusHTML =
        `<span class="status-neutral">
           Insufficient history
         </span>`;
    } else if (status.flagged) {
      statusHTML =
        `<span class="status-danger">
           Flagged
         </span>
         <small>
           ${status.reasons
             .map(reason => escapeHTML(reason))
             .join("<br>")}
         </small>`;
    } else {
      statusHTML =
        `<span class="status-good">
           No flag
         </span>`;
    }


    const changeClass =
      row.achDropBaseline > 0
        ? "danger-text"
        : row.achDropBaseline < 0
          ? "positive-text"
          : "";


    return `
      <tr>
        <td>${escapeHTML(row.building)}</td>
        <td>${escapeHTML(row.lab)}</td>
        <td>
          <button
            class="table-hood-button"
            data-hood-key="${escapeHTML(row.hoodKey)}"
            type="button"
          >
            ${escapeHTML(row.hood)}
          </button>
        </td>

        <td>${formatDate(row.firstInspection)}</td>
        <td>${formatDate(row.latestInspection)}</td>

        <td>${formatNumber(row.baselineACH, 1)}</td>
        <td>${formatNumber(row.latestACH, 1)}</td>

        <td class="${changeClass}">
          ${formatNumber(row.achDropBaseline, 1)}
        </td>

        <td class="${changeClass}">
          ${formatPercent(row.achDropPercentBaseline)}
        </td>

        <td>${statusHTML}</td>
      </tr>
    `;

  }).join("");


  resultsBody
    .querySelectorAll(".table-hood-button")
    .forEach(button => {

      button.addEventListener("click", () => {

        selectedHoodKey =
          button.dataset.hoodKey;

        hoodFilter.value =
          selectedHoodKey;

        updateTrendChart();
         
        trendChart.scrollIntoView({
          behavior: "smooth",
          block: "center"
        });
      });
    });
}


/* ---------------------------------------------------------
   Trend chart
   --------------------------------------------------------- */

function updateTrendChart() {

  if (!selectedHoodKey) {

    Plotly.react(
      trendChart,
      [],
      {
        ...baseChartLayout("Select a hood to view ACH history"),
        xaxis: { visible: false },
        yaxis: { visible: false }
      },
      plotConfig()
    );

    return;
  }


  const records = filteredRecords
    .filter(
      record => record.hoodKey === selectedHoodKey
    )
    .sort((a, b) => a.date - b.date);


  if (!records.length) {
    return;
  }


  const selectedSummary =
    summary.find(
      row => row.hoodKey === selectedHoodKey
    );


  const trace = {
    x: records.map(record => record.date),
    y: records.map(record => record.ach),
    type: "scatter",
    mode: "lines+markers",
    name: "ACH",
    hovertemplate:
      "<b>%{x|%b %d, %Y}</b><br>" +
      "ACH: %{y:.1f}<extra></extra>"
  };


  const shapes = [];

  if (selectedSummary) {

    shapes.push({
      type: "line",
      x0: records[0].date,
      x1: records[records.length - 1].date,
      y0: selectedSummary.baselineACH,
      y1: selectedSummary.baselineACH,
      line: {
        dash: "dash",
        width: 1
      }
    });
  }


  const layout = {
    ...baseChartLayout(
      `${selectedSummary?.hood || ""} — ACH history`
    ),

    xaxis: {
      title: "Inspection date",
      type: "date"
    },

    yaxis: {
      title: "ACH",
      rangemode: "tozero"
    },

    shapes
  };


  Plotly.react(
    trendChart,
    [trace],
    layout,
    plotConfig()
  );
}


/* ---------------------------------------------------------
   Face Velocity History chart
   Replaces the Deterioration Ranking chart
   --------------------------------------------------------- */

function updateFaceVelocityChart() {
  if (!selectedHoodKey) {
    Plotly.react(
      rankingChart,
      [],
      {
        ...baseChartLayout(
          "Select a hood to view Face Velocity History"
        ),
        xaxis: { visible: false },
        yaxis: { visible: false }
      },
      plotConfig()
    );
    return;
  }

  const records = filteredRecords
    .filter(record =>
      record.hoodKey === selectedHoodKey &&
      Number.isFinite(record.faceVelocity)
    )
    .sort((a, b) => a.date - b.date);

  if (!records.length) {
    Plotly.react(
      rankingChart,
      [],
      baseChartLayout(
        "No face velocity data available for this hood"
      ),
      plotConfig()
    );
    return;
  }

  const trace = {
    x: records.map(record => record.date),
    y: records.map(record => record.faceVelocity),
    type: "scatter",
    mode: "lines+markers",
    name: "Face velocity",
    hovertemplate:
      "<b>%{x|%b %d, %Y}</b><br>" +
      "Face velocity: %{y:.2f}<extra></extra>"
  };

  // Show the first recorded face velocity as a baseline.
  const baseline = records[0].faceVelocity;

  const shapes = [{
    type: "line",
    x0: records[0].date,
    x1: records[records.length - 1].date,
    y0: baseline,
    y1: baseline,
    line: {
      dash: "dash",
      width: 1,
      color: "#94a3b8"
    }
  }];

  const layout = {
    ...baseChartLayout(
      `${records[0].hood} — Face Velocity History`
    ),
    xaxis: {
      title: "Inspection date",
      type: "date"
    },
    yaxis: {
      title: "Face velocity",
      rangemode: "tozero"
    },
    shapes
  };

  Plotly.react(
    rankingChart,
    [trace],
    layout,
    plotConfig()
  );
}


function rankingValue(row, metric) {

  if (metric === "percent") {
    return Math.abs(
      row.achDropPercentBaseline ?? 0
    );
  }

  if (metric === "drop") {
    return Math.max(
      row.achDropBaseline,
      0
    );
  }

  return Math.abs(
    row.achDropBaseline
  );
}


/* ---------------------------------------------------------
   Heatmap
   --------------------------------------------------------- */

function updateHeatmapChart() {
  const selectedBuilding = buildingFilter.value;
   const metric = heatmapMetric.value || "ach";

   const metricField =
     metric === "faceVelocity" ? "faceVelocity" : "ach";
   
   const metricLabel =
     metric === "faceVelocity" ? "Face velocity" : "ACH";
   
   heatmapTitle.textContent =
     `${metricLabel} by Hood and Inspection Period`;
   
  if (!selectedBuilding) {
    Plotly.react(
      heatmapChart,
      [],
      {
        ...baseChartLayout(
          "Select a building to view hood ACH"
        ),
        xaxis: { visible: false },
        yaxis: { visible: false },
        annotations: [{
          text: "Please select a building first.",
          x: 0.5,
          y: 0.5,
          xref: "paper",
          yref: "paper",
          showarrow: false,
          font: { size: 14 }
        }]
      },
      plotConfig()
    );
    return;
  }

  // Keep every inspection, not just the latest per hood.
   
  const records = filteredRecords
     .filter(record =>
       record.building === selectedBuilding &&
       Number.isFinite(record[metricField])
     )
     .sort((a, b) =>
       naturalSort(a.hood, b.hood) ||
       a.date - b.date
     );

  if (!records.length) {
    Plotly.react(
      heatmapChart,
      [],
      baseChartLayout(
        `${selectedBuilding} — No matching inspections`
      ),
      plotConfig()
    );
    return;
  }

  /*
   * Assign each hood a stable categorical X position.
   * Include lab in the key so identically named hoods
   * in different labs are distinguishable.
   */
  const hoodKeys = [...new Set(
    records.map(record => record.hoodKey)
  )].sort(naturalSort);

  const hoodLabels = hoodKeys.map(key => {
    const record = records.find(r => r.hoodKey === key);
    return `${record.lab} · ${record.hood}`;
  });

  const hoodPosition = new Map(
    hoodKeys.map((key, index) => [key, index])
  );

  /*
   * Inspection-period color groups.
   * Adjust these boundaries or labels to match your
   * actual inspection schedule.
   */
    function getInspectionPeriod(date) {
    return String(date.getFullYear());
  }

  // Discover all years present in the filtered records.
  const periods = [...new Set(
    records.map(record => getInspectionPeriod(record.date))
  )].sort((a, b) => Number(a) - Number(b));

  // Generate distinct colors automatically for any number of years.
  function yearColor(year, index, total) {
    const palette = [
      "#f59e0b", // Orange
      "#22c55e", // Green
      "#3b82f6", // Blue
      "#a855f7", // Purple
      "#ef4444", // Red
      "#06b6d4", // Cyan
      "#ec4899", // Pink
      "#84cc16", // Lime
      "#f97316", // Dark orange
      "#6366f1"  // Indigo
    ];

    if (index < palette.length) {
      return palette[index];
    }

    // Continue generating colors for additional years.
    const hue = (Number(year) * 137.508) % 360;
    return `hsl(${hue}, 65%, 45%)`;
  }

  const traces = periods.map((period, periodIndex) => {
     const periodRecords = records.filter(
       record => getInspectionPeriod(record.date) === period
     );
   
     return {
       type: "scatter",
       mode: "markers",
       name: period,
       x: periodRecords.map(
         record => hoodPosition.get(record.hoodKey)
       ),
       y: periodRecords.map(record => record[metricField]),
       marker: {
         color: yearColor(period, periodIndex, periods.length),
         size: 10,
         opacity: 0.85,
         line: {
          color: yValues.map(y => {
            if (metric === "ach") {
              return (y < 200 || y > 400)
                ? "rgba(240, 71, 52, 0.8)"
                : "rgba(0, 0, 0, 0.15)";
            }
      
            if (metric === "face velocity") {
              return (y < 80 || y > 120)
                ? "rgba(240, 71, 52, 0.8)"
                : "rgba(0, 0, 0, 0.15)";
            }
      
            return "rgba(0, 0, 0, 0.15)";
          }),
          width: 1.5
        }
       },
      customdata: periodRecords.map(record => [
        record.building,
        record.lab,
        record.hood,
        record.date,
        record[metricField]
      ]),
      hovertemplate:
        "<b>%{customdata[2]}</b><br>" +
        "Building: %{customdata[0]}<br>" +
        "Lab: %{customdata[1]}<br>" +
        "Inspection date: %{customdata[3]|%b %d, %Y}<br>" +
        `${metricLabel}: %{customdata[4]:.2f}<br>` +
        "Period: " + period +
        "<extra></extra>"
    };
  });

  Plotly.react(
    heatmapChart,
    traces,
    {
      ...baseChartLayout(
        `${selectedBuilding} — ${metricLabel} by Hood and Inspection Period`
      ),
      margin: {
        l: 75,
        r: 25,
        t: 60,
        b: 160
      },
      xaxis: {
        title: "Hood name",
        tickmode: "array",
        tickvals: hoodKeys.map((_, index) => index),
        ticktext: hoodLabels,
        tickangle: -45,
        automargin: true,
        zeroline: false
      },
      yaxis: {
        title: metric === "faceVelocity"
          ? "Face velocity"
          : "Air changes per hour (ACH)",
        rangemode: "tozero",
        autorange: true
      },
       shapes:metric === "ach"
          ? [
              {
                type: "rect",
                xref: "paper",
                x0: 0,
                x1: 1,
                yref: "y",
                y0: 200,
                y1: 400,
                fillcolor: "rgba(34, 197, 94, 0.20)",
                line: {
                  color: "rgba(34, 197, 94, 0.5)",
                  width: 1
                },
                layer: "below"
              }
            ]
          : metric === "faceVelocity"
          ? [
              {
                type: "rect",
                xref: "paper",
                x0: 0,
                x1: 1,
                yref: "y",
                y0: 80,
                y1: 120,
                fillcolor: "rgba(34, 197, 94, 0.20)",
                line: {
                  color: "rgba(34, 197, 94, 0.5)",
                  width: 1
                },
                layer: "below"
              }
            ]
          : [],
      showlegend: true,
      legend: {
        orientation: "h",
        y: -1,
        x: 0,
        title: { text: "Hood Test Year:" }
      },
      hovermode: "closest"
    },
    plotConfig()
  );
}

heatmapMetric.addEventListener(
  "change",
  updateHeatmapChart
);

/* ---------------------------------------------------------
   Export
   --------------------------------------------------------- */

exportBtn.addEventListener("click", () => {

  if (!summary.length) {
    return;
  }


  const exportRows = summary.map(row => {

    const status =
      getDeteriorationStatus(row);


    return {
      "Hood Key": row.hoodKey,
      "Building ID": row.building,
      "Lab ID": row.lab,
      "Hood ID": row.hood,

      "First Inspection":
        dateKey(row.firstInspection),

      "Latest Inspection":
        dateKey(row.latestInspection),

      "Baseline ACH":
        row.baselineACH,

      "Latest ACH":
        row.latestACH,

      "Previous ACH":
        row.previousACH,

      "ACH Drop vs Previous":
        row.achDropPrevious,

      "ACH Drop % vs Previous":
        row.achDropPercentPrevious,

      "ACH Drop vs Baseline":
        row.achDropBaseline,

      "ACH Drop % vs Baseline":
        row.achDropPercentBaseline,

      "Inspection Count":
        row.inspectionCount,

      "Trend Available":
        row.trendAvailable,

      "Flagged":
        status.flagged,

      "Flag Reason":
        status.reasons.join("; ")
    };
  });


  const worksheet =
    XLSX.utils.json_to_sheet(exportRows);

  const workbook =
    XLSX.utils.book_new();

  XLSX.utils.book_append_sheet(
    workbook,
    worksheet,
    "Hood Summary"
  );


  const csv =
    XLSX.utils.sheet_to_csv(worksheet);


  const blob = new Blob(
    [csv],
    { type: "text/csv;charset=utf-8;" }
  );


  const url =
    URL.createObjectURL(blob);

  const link =
    document.createElement("a");

  link.href = url;
  link.download =
    `hoodwatch-filtered-${dateKey(new Date())}.csv`;

  document.body.appendChild(link);
  link.click();
  link.remove();

  URL.revokeObjectURL(url);
});


/* ---------------------------------------------------------
   Event listeners
   --------------------------------------------------------- */

buildingFilter.addEventListener("change", () => {

  updateLabOptions();

  applyFilters();
});


labFilter.addEventListener(
  "change",
  applyFilters
);


startDate.addEventListener(
  "change",
  applyFilters
);


endDate.addEventListener(
  "change",
  applyFilters
);

hoodFilter.addEventListener(
  "change",
  () => {

    selectedHoodKey = hoodFilter.value;

     updateTrendChart();
     updateFaceVelocityChart();
  }
);


resetFiltersBtn.addEventListener(
  "click",
  () => {

    if (!allRecords.length) {
      return;
    }

    buildingFilter.value = "";

    updateLabOptions();

    labFilter.value = "";

    const dates = allRecords
      .map(record => record.date)
      .sort((a, b) => a - b);

    startDate.value =
      dateKey(dates[0]);

    endDate.value =
      dateKey(dates[dates.length - 1]);

    selectedHoodKey = "";
    hoodFilter.value = "";

    applyFilters();
  }
);


clearBtn.addEventListener(
  "click",
  clearData
);


/* ---------------------------------------------------------
   Clear
   --------------------------------------------------------- */

function clearData() {

  heatmapMetric.value = "ach";
   heatmapMetric.disabled = true;
   heatmapTitle.textContent =
     "ACH by Hood and Inspection Period";
   
  allRecords = [];
  filteredRecords = [];
  summary = [];
  selectedHoodKey = "";

  fileInput.value = "";

  fileStatus.textContent =
    "No file loaded";

  recordInfo.textContent =
    "Upload your inspection history to begin.";

  clearBtn.disabled = true;
  exportBtn.disabled = true;
  resetFiltersBtn.disabled = true;

  enableControls(false);

  buildingFilter.innerHTML =
    `<option value="">All buildings</option>`;

  labFilter.innerHTML =
    `<option value="">All labs</option>`;

  hoodFilter.innerHTML =
    `<option value="">Select a hood</option>`;

  kpiHoods.textContent = "—";
  kpiFlagged.textContent = "—";
  kpiBuildings.textContent = "—";
  kpiLabs.textContent = "—";
  kpiVAVFailing.textContent = "—";
   
  flagCount.textContent =
    "0 flagged";

  tableCount.textContent = "";

  hoodList.innerHTML = `
    <div class="empty-state">
      Upload inspection data to see flagged hoods.
    </div>
  `;

  resultsBody.innerHTML = `
    <tr>
      <td colspan="10" class="empty-cell">
        Upload a file to view results.
      </td>
    </tr>
  `;


  Plotly.react(
    trendChart,
    [],
    baseChartLayout("Select a hood to view ACH history"),
    plotConfig()
  );

  Plotly.react(
    rankingChart,
    [],
    baseChartLayout("Select a hood to view Face Velocity History"),
    plotConfig()
  );

  Plotly.react(
    heatmapChart,
    [],
    baseChartLayout("Building and lab heatmap"),
    plotConfig()
  );

  clearError();
}


/* ---------------------------------------------------------
   Controls
   --------------------------------------------------------- */

function enableControls(enabled) {
  buildingFilter.disabled = false;
  labFilter.disabled = false;
  startDate.disabled = false;
  endDate.disabled = false;
  hoodFilter.disabled = false;

  heatmapMetric.disabled = !enabled;
  clearBtn.disabled = !enabled;
  exportBtn.disabled = !enabled;
  resetFiltersBtn.disabled = !enabled;
}


/* ---------------------------------------------------------
   Plotly helpers
   --------------------------------------------------------- */

function baseChartLayout(title) {

  return {
    title: {
      text: title,
      font: {
        size: 15
      }
    },

    margin: {
      l: 70,
      r: 30,
      t: 55,
      b: 60
    },

    paper_bgcolor: "rgba(0,0,0,0)",

    plot_bgcolor: "rgba(0,0,0,0)",

    font: {
      family:
        "Inter, system-ui, -apple-system, BlinkMacSystemFont, " +
        "\"Segoe UI\", sans-serif"
    },

    hoverlabel: {
      align: "left"
    }
  };
}


function plotConfig() {

  return {
    responsive: true,
    displaylogo: false,

    modeBarButtonsToRemove: [
      "lasso2d",
      "select2d",
      "autoScale2d"
    ]
  };
}


/* ---------------------------------------------------------
   General helpers
   --------------------------------------------------------- */

function unique(values) {
  return [...new Set(values)];
}


function naturalSort(a, b) {

  return String(a).localeCompare(
    String(b),
    undefined,
    {
      numeric: true,
      sensitivity: "base"
    }
  );
}


function showError(message) {

  errorMessage.textContent = message;
  errorMessage.hidden = false;

  fileStatus.textContent =
    "Error loading file";
}


function clearError() {

  errorMessage.textContent = "";
  errorMessage.hidden = true;
}


/* ---------------------------------------------------------
   Initial chart state
   --------------------------------------------------------- */

Plotly.newPlot(
  trendChart,
  [],
  baseChartLayout("Select a hood to view ACH history"),
  plotConfig()
);


Plotly.newPlot(
  rankingChart,
  [],
  baseChartLayout("Select a hood to view Face Velocity History"),
  plotConfig()
);


Plotly.newPlot(
  heatmapChart,
  [],
  baseChartLayout("Building and lab heatmap"),
  plotConfig()
);
