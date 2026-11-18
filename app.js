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

let sortColumn = "building";
let sortDirection = "asc";


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


const trendChart = document.getElementById("trendChart");
const rankingChart = document.getElementById("rankingChart");
const heatmapChart = document.getElementById("heatmapChart");

const resultsBody = document.getElementById("resultsBody");
const tableCount = document.getElementById("tableCount");

const heatmapMetric = document.getElementById("heatmapMetric");
const heatmapTitle = document.getElementById("heatmapTitle");

const greenBoxRanges = {
    ach: { min: 200, max: 400 },
    faceVelocity: { min: 80, max: 120 }
};

/* ---------------------------------------------------------
   Required columns
   --------------------------------------------------------- */

const REQUIRED_COLUMNS = [
  "Building",
  "Lab",
  "Hood",
  "Date",
  "Min ACH",
  "FaceV Avg"
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

function parsePercentage(value) {
  if (value === null || value === undefined || value === "") {
    return NaN;
  }

  if (typeof value === "number") {
    return Number.isFinite(value) ? value * 100 : NaN;
  }

  const cleaned = String(value)
    .replace(/,/g, "")
    .replace(/[^\d.+\-eE]/g, "");

  const number = Number(cleaned);

  return Number.isFinite(number) ? number * 100 : NaN;
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
    setGreenBoxDefaults()
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

  // Find optional spreadsheet columns even if their
  // capitalization, spacing, or underscores differ.
  function getOptionalColumn(row, names) {
    const normalize = value =>
      String(value).toLowerCase().replace(/[^a-z0-9]/g, "");

    for (const name of names) {
      const key = Object.keys(row).find(
        column => normalize(column) === normalize(name)
      );

      if (key !== undefined) {
        return normalizeValue(row[key]);
      }
    }

    return "";
  }

  rows.forEach((row, index) => {
    const building = normalizeValue(
      row["Building"],
      "Unknown Building"
    );

    const lab = normalizeValue(
      row["Lab"],
      "Unknown Lab"
    );

    const hood = normalizeValue(
      row["Hood"],
      "Unknown Hood"
    );

    const date = parseDate(row["Date"]);
    const ach = parseACH(row["Min ACH"]);
    const cov = parsePercentage(row["Average_Zone_COV"]);
    const faceVelocity = parseACH(row["FaceV Avg"]);
    const testRating = getOptionalColumn(row, ["Test Rating", "Test_Rating", "TestRating"]);

    // Read metadata from the original inspection sheet.
    const exhaustType = getOptionalColumn(row, [
      "Exhaust Type",
      "Exhaust_Type",
      "ExhaustType"
    ]);

    const totalArea = getOptionalColumn(row, [
      "Total Area",
      "Total_Area",
      "TotalArea"
    ]);

    // Keep the same record-validation rules.
    if (
      !date ||
      (!Number.isFinite(ach) &&
       !Number.isFinite(faceVelocity))
    ) {
      return;
    }

    // Unique hood identifier.
    const hoodKey = `${building} | ${lab} | ${hood}`;

    parsed.push({
      rowNumber: index + 2,
      building,
      lab,
      hood,
      date,
      dateString: dateKey(date),

      ach: Number.isFinite(ach) ? ach : null,
      faceVelocity: Number.isFinite(faceVelocity)
        ? faceVelocity
        : null,
      cov: Number.isFinite(cov) ? cov : null,

      exhaustType,
      totalArea,
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
   * Group ALL inspection records by Hood Key.
   * Do not exclude records with missing ACH or face velocity.
   */

  const sorted = [...records].sort((a, b) => {
    const hoodCompare = naturalSort(a.hoodKey, b.hoodKey);

    if (hoodCompare !== 0) {
      return hoodCompare;
    }

    return a.date - b.date || a.rowNumber - b.rowNumber;
  });

  const groups = new Map();

  sorted.forEach(record => {
    if (!groups.has(record.hoodKey)) {
      groups.set(record.hoodKey, []);
    }

    groups.get(record.hoodKey).push(record);
  });


  /*
   * Build one summary row per hood.
   */

  const results = [];

  for (const [hoodKey, hoodRecords] of groups) {

    // All inspection records are already chronological.
    const first = hoodRecords[0];
    const latest = hoodRecords[hoodRecords.length - 1];

    /*
     * Find all valid ACH measurements.
     */
    const achRecords = hoodRecords.filter(
      record => Number.isFinite(record.ach)
    );

    /*
     * Find all valid face velocity measurements.
     */
    const faceVelocityRecords = hoodRecords.filter(
      record => Number.isFinite(record.faceVelocity)
    );

    /*
     * Latest valid measurement for each metric.
     */
    const latestACHRecord =
      achRecords.length
        ? achRecords[achRecords.length - 1]
        : null;

    const latestFaceVelocityRecord =
      faceVelocityRecords.length
        ? faceVelocityRecords[faceVelocityRecords.length - 1]
        : null;

    /*
     * Calculate ACH comparisons using valid ACH records only.
     */
    const previousACHRecord =
      achRecords.length >= 2
        ? achRecords[achRecords.length - 2]
        : null;

    const previousACH =
      previousACHRecord?.ach ?? null;

    let baselineACH = null;

    if (achRecords.length >= 2) {

      const previousRecords = achRecords.slice(0, -1);

      const previousSum = previousRecords.reduce(
        (sum, record) => sum + record.ach,
        0
      );

      baselineACH = previousSum / previousRecords.length;
    }

    const latestACH =
      latestACHRecord?.ach ?? null;

    const achDropBaseline =
      latestACH !== null && baselineACH !== null
        ? latestACH - baselineACH
        : null;

    const achDropPrevious =
      latestACH !== null && previousACH !== null
        ? latestACH - previousACH
        : null;

    const achDropPercentBaseline =
      baselineACH !== null && baselineACH > 0
        ? 100 * (latestACH - baselineACH) / baselineACH
        : null;

    const achDropPercentPrevious =
      previousACH !== null && previousACH > 0
        ? 100 * (latestACH - previousACH) / previousACH
        : null;


    /*
     * Store one summary row for this hood.
     */
    results.push({

      hoodKey,

      building: latest.building,
      lab: latest.lab,
      hood: latest.hood,

      // First and last inspection dates, regardless of metric availability.
      firstTestDate: first.date,
      lastTestDate: latest.date,

      // Preserve existing field names for table compatibility.
      firstInspection: first.date,
      latestInspection: latest.date,

      // Latest available measurements, independently.
      latestACH,
      latestFaceVelocity:
        latestFaceVelocityRecord?.faceVelocity ?? null,

      // ACH comparison fields retained for existing dashboard logic.
      baselineACH,
      previousACH,

      achDropBaseline,
      achDropPercentBaseline,

      achDropPrevious,
      achDropPercentPrevious,

      // Total inspections, including records with neither measurement.
      inspectionCount: hoodRecords.length,

      // Number of valid measurements for each metric.
      achInspectionCount: achRecords.length,
      faceVelocityInspectionCount: faceVelocityRecords.length,

      trendAvailable: achRecords.length >= 2
    });
  }


  /*
   * Sort highest latest ACH first.
   * Hoods without ACH measurements appear at the bottom.
   */
  return results.sort((a, b) => {

    if (a.latestACH === null) return 1;
    if (b.latestACH === null) return -1;

    return b.latestACH - a.latestACH;
  });
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
   Get Face Velocity Deterioration
   ---------------------------------------------------------
*/

function getFaceVelocityDeteriorationStatus(hoodKey) {

  const min = greenBoxRanges.faceVelocity.min;
  const max = greenBoxRanges.faceVelocity.max;

  // Get valid face velocity measurements for this hood.
  const records = filteredRecords
    .filter(record =>
      record.hoodKey === hoodKey &&
      Number.isFinite(record.faceVelocity)
    )
    .sort((a, b) =>
      a.date - b.date ||
      a.rowNumber - b.rowNumber
    );

  if (records.length < 2) {
    return { flagged: false };
  }

  // Keep only the final valid measurement of each year.
  const latestByYear = new Map();

  records.forEach(record => {
    const year = record.date.getFullYear();
    latestByYear.set(year, record);
  });

  const annualRecords = [...latestByYear.values()];

  // Must have measurements from at least two years.
  if (annualRecords.length < 2) {
    return { flagged: false };
  }

  // Most recent valid face velocity measurement.
  const latest = annualRecords[annualRecords.length - 1];

  // Only compare against earlier years.
  const previousYears = annualRecords.slice(0, -1);

  const latestOutsideGreen =
    latest.faceVelocity < min ||
    latest.faceVelocity > max;

  const previouslyGreen = previousYears.some(record =>
    record.faceVelocity >= min &&
    record.faceVelocity <= max
  );

  return {
    flagged: latestOutsideGreen && previouslyGreen,
    latestFaceVelocity: latest.faceVelocity,
    latestYear: latest.date.getFullYear(),
    previouslyGreen
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
  updateHoodSelector();
  updateTable();
  updateBuildingComparison();
  updateTrendChart();
  updateFaceVelocityChart();
  updateHeatmapChart();
  updateYearlyPassSummary();
  updateWorstPerformers();
  updateAnalysis();
}

function updateHeatmap() {
  updateHeatmapChart();
}


/* ---------------------------------------------------------
   Building Performance Comparison
   --------------------------------------------------------- */

let buildingComparisonSort = "building";
let buildingComparisonDirection = "asc";

function calculateBuildingComparison() {
  // Compare every building, regardless of building/lab selections.
  // Date limits still apply to the comparison.
  const start = startDate.value
    ? new Date(`${startDate.value}T00:00:00`) : null;
  const end = endDate.value
    ? new Date(`${endDate.value}T23:59:59.999`) : null;

  const records = allRecords.filter(record =>
    (!start || record.date >= start) &&
    (!end || record.date <= end)
  ).sort((a, b) => a.date - b.date || a.rowNumber - b.rowNumber);

  const hoods = new Map();
  const buildings = new Map();
  for (const record of records) {
    if (!hoods.has(record.hoodKey)) hoods.set(record.hoodKey, []);
    hoods.get(record.hoodKey).push(record);
    if (!buildings.has(record.building)) {
      buildings.set(record.building, {
        building: record.building, hoodCount: 0,
        velocityTotal: 0, velocityCount: 0, lowVelocityCount: 0,
        covTotal: 0, covCount: 0,
        achCount: 0, unsatisfactoryAchCount: 0
      });
    }
  }

  for (const hoodRecords of hoods.values()) {
    const stats = buildings.get(hoodRecords[0].building);
    stats.hoodCount++;
    const latest = field => {
      for (let i = hoodRecords.length - 1; i >= 0; i--) {
        if (Number.isFinite(hoodRecords[i][field])) {
          return hoodRecords[i][field];
        }
      }
      return null;
    };

    const velocity = latest("faceVelocity");
    if (velocity !== null) {
      stats.velocityTotal += velocity;
      stats.velocityCount++;
      if (velocity < greenBoxRanges.faceVelocity.min) stats.lowVelocityCount++;
    }

    const cov = latest("cov");
    if (cov !== null) {
      stats.covTotal += cov;
      stats.covCount++;
    }

    const ach = latest("ach");
    if (ach !== null) {
      stats.achCount++;
      if (ach < greenBoxRanges.ach.min || ach > greenBoxRanges.ach.max) {
        stats.unsatisfactoryAchCount++;
      }
    }
  }

  return [...buildings.values()].map(s => ({
    building: s.building,
    hoodCount: s.hoodCount,
    avgVelocity: s.velocityCount ? s.velocityTotal / s.velocityCount : null,
    lowVelocityPct: s.velocityCount ? 100 * s.lowVelocityCount / s.velocityCount : null,
    vavCov: s.covCount ? s.covTotal / s.covCount : null,
    unsatisfactoryAchPct: s.achCount ? 100 * s.unsatisfactoryAchCount / s.achCount : null
  }));
}

function updateBuildingComparison() {
  const body = document.getElementById("buildingComparisonBody");
  if (!body) return;
  const rows = calculateBuildingComparison();
  rows.sort((a, b) => {
    const av = a[buildingComparisonSort];
    const bv = b[buildingComparisonSort];
    if (av == null || bv == null) {
      if (av == null && bv != null) return 1;
      if (bv == null && av != null) return -1;
    }
    const cmp = typeof av === "string" ? naturalSort(av, bv) : (av - bv);
    return (buildingComparisonDirection === "asc" ? cmp : -cmp) ||
      naturalSort(a.building, b.building);
  });

  const number = (value, digits = 1) =>
    Number.isFinite(value) ? formatNumber(value, digits) : "—";
  const percent = value => Number.isFinite(value) ? `${number(value)}%` : "—";
  const risk = value => !Number.isFinite(value) ? "" :
    value >= 25 ? "building-risk-high" :
    value >= 10 ? "building-risk-medium" : "building-risk-low";

  body.innerHTML = rows.length ? rows.map(r => `
    <tr>
      <td><strong>${escapeHTML(r.building)}</strong></td>
      <td>${r.hoodCount}</td>
      <td>${number(r.avgVelocity)}</td>
      <td class="${risk(r.lowVelocityPct)}">${percent(r.lowVelocityPct)}</td>
      <td>${percent(r.vavCov)}</td>
      <td class="${risk(r.unsatisfactoryAchPct)}">${percent(r.unsatisfactoryAchPct)}</td>
    </tr>
  `).join("") :
    '<tr><td colspan="6" class="empty-cell">No buildings match the selected date range.</td></tr>';

  document.querySelectorAll("[data-building-sort]").forEach(th => {
    const active = th.dataset.buildingSort === buildingComparisonSort;
    th.classList.toggle("sort-asc", active && buildingComparisonDirection === "asc");
    th.classList.toggle("sort-desc", active && buildingComparisonDirection === "desc");
    th.setAttribute("aria-sort", active ?
      (buildingComparisonDirection === "asc" ? "ascending" : "descending") : "none");
  });
}

document.querySelectorAll("[data-building-sort]").forEach(th => {
  th.addEventListener("click", () => {
    const column = th.dataset.buildingSort;
    if (buildingComparisonSort === column) {
      buildingComparisonDirection = buildingComparisonDirection === "asc" ? "desc" : "asc";
    } else {
      buildingComparisonSort = column;
      buildingComparisonDirection = "asc";
    }
    updateBuildingComparison();
  });
});

/* ---------------------------------------------------------
   KPIs
   --------------------------------------------------------- */

function updateKPIs() {

  const monitoredHoods = new Set(
    filteredRecords.map(record => record.hoodKey)
  );

  const buildings = new Set(
    filteredRecords.map(record => record.building)
  );

  const labs = new Set(
    filteredRecords.map(
      record => `${record.building} | ${record.lab}`
    )
  );

  // ACH analytics still use summary.
  const flagged = summary.filter(row =>
     getFaceVelocityDeteriorationStatus(row.hoodKey).flagged
   );

  const failingVAVSystems = new Set(
    summary
      .filter(row => getVAVFailureStatus(row).failing)
      .map(row => row.hoodKey)
  );

  kpiHoods.textContent =
    monitoredHoods.size.toLocaleString();

  kpiBuildings.textContent =
    buildings.size.toLocaleString();

  kpiLabs.textContent =
    labs.size.toLocaleString();

  kpiFlagged.textContent =
    flagged.length.toLocaleString();

  kpiVAVFailing.textContent =
    failingVAVSystems.size.toLocaleString();
}



/* ---------------------------------------------------------
   Hood selector
   --------------------------------------------------------- */

function updateHoodSelector() {

  /*const previous = selectedHoodKey;

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
  } */
   const previous = selectedHoodKey;

  // Get every unique hood from filtered records.
  const hoods = new Map();

  filteredRecords.forEach(record => {
    if (!hoods.has(record.hoodKey)) {
      hoods.set(record.hoodKey, record);
    }
  });

  const sorted = [...hoods.entries()].sort(
    (a, b) => naturalSort(a[0], b[0])
  );

  hoodFilter.innerHTML =
    `<option value="">Select a hood</option>` +
    sorted.map(([key, record]) => `
      <option value="${escapeHTML(key)}">
        ${escapeHTML(key)}
      </option>
    `).join("");

  if (
    previous &&
    hoods.has(previous)
  ) {
    hoodFilter.value = previous;
  } else {
    selectedHoodKey = "";
    hoodFilter.value = "";
  }
}


/* ---------------------------------------------------------
   Inspection table
   --------------------------------------------------------- */
/*Sorting table*/

function sortSummary(data) {

  return [...data].sort((a, b) => {

    let valueA;
    let valueB;

    switch (sortColumn) {

      case "status":
        valueA = getFaceVelocityDeteriorationStatus(a.hoodKey).flagged ? 1 : 0;
        valueB = getFaceVelocityDeteriorationStatus(b.hoodKey).flagged ? 1 : 0;
        break;

      case "building":
      case "lab":
      case "hood":
        valueA = a[sortColumn] || "";
        valueB = b[sortColumn] || "";
        break;

      case "firstInspection":
      case "latestInspection":
        valueA = a[sortColumn]?.getTime() ?? -Infinity;
        valueB = b[sortColumn]?.getTime() ?? -Infinity;
        break;

      default:
        valueA = a[sortColumn] ?? -Infinity;
        valueB = b[sortColumn] ?? -Infinity;
    }

    let comparison;

    if (
      typeof valueA === "string" &&
      typeof valueB === "string"
    ) {
      comparison = naturalSort(valueA, valueB);
    } else {
      comparison = valueA - valueB;
    }

    return sortDirection === "asc"
      ? comparison
      : -comparison;
  });
}

function updateTable() {

  tableCount.textContent =
    `${summary.length} hoods`;

  if (!summary.length) {
    resultsBody.innerHTML = `
      <tr>
        <td colspan="7" class="empty-cell">
          No inspection records match the current filters.
        </td>
      </tr>
    `;
    return;
  }

  resultsBody.innerHTML = sortSummary(summary)
    .map(row => {

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

          <td>${formatDate(row.firstTestDate)}</td>

          <td>${formatDate(row.lastTestDate)}</td>

          <td>${formatNumber(row.latestACH, 1)}</td>

          <td>${formatNumber(row.latestFaceVelocity, 2)}</td>
        </tr>
      `;

    }).join("");

  resultsBody
    .querySelectorAll(".table-hood-button")
    .forEach(button => {

      button.addEventListener("click", () => {

        selectHoodAndShowHistory(button.dataset.hoodKey);
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
      type: "rect",
      x0: records[0].date,
      x1: records[records.length - 1].date,
      y0: greenBoxRanges.ach.min,
      y1: greenBoxRanges.ach.max,
      fillcolor: "rgba(34, 197, 94, 0.20)",
      line: {
         color: "rgba(34, 197, 94, 0.5)",
         width: 1
      },
      layer: "below"
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

  const shapes = [
     {
       type: "rect",
       xref: "paper",
       x0: 0,
       x1: 1,
       yref: "y",
       y0: greenBoxRanges.faceVelocity.min,
       y1: greenBoxRanges.faceVelocity.max,
       fillcolor: "rgba(34, 197, 94, 0.20)",
       line: {
         color: "rgba(34, 197, 94, 0.5)",
         width: 1
       },
       layer: "below"
     }
   ];

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

function setGreenBoxDefaults() {
    const metric = heatmapMetric.value || "ach";

    document.getElementById("greenY0").value =
        greenBoxRanges[metric].min;

    document.getElementById("greenY1").value =
        greenBoxRanges[metric].max;
}

function updateHeatmapChart() {
  const selectedBuilding = buildingFilter.value;
  const metric = heatmapMetric.value || "ach";

  const y0Input = document.getElementById("greenY0");
  const y1Input = document.getElementById("greenY1");

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
   
   // Assign a distinct marker shape to each year.
   function yearShape(index) {
     const shapes = [
       "circle",
       "square",
       "diamond",
       "triangle-up",
       "cross",
       "star",
       "hexagon",
       "triangle-down"
     ];
   
     return shapes[index % shapes.length];
   }

   // Get user-defined minimum and maximum values.
   const minValue = Number(y0Input.value);
   const maxValue = Number(y1Input.value);
   
   let traces = periods.map((period, periodIndex) => {
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
         symbol: yearShape(periodIndex),
   
         // Red if outside user-defined limits; green otherwise.
         color: periodRecords.map(record => {
           const value = Number(record[metricField]);
   
           return (value < minValue || value > maxValue)
             ? "#ef4444"
             : "#22c55e";
         }),
   
         size: 10,
         opacity: 0.85,
   
         line: {
           color: "rgba(0, 0, 0, 0.25)",
           width: 1
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
          shapes:
           metric === "ach"
             ? [
                 {
                   type: "rect",
                   xref: "paper",
                   x0: 0,
                   x1: 1,
                   yref: "y",
                   y0: Number(document.getElementById("greenY0").value),
                   y1: Number(document.getElementById("greenY1").value),
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
                   y0: Number(document.getElementById("greenY0").value),
                   y1: Number(document.getElementById("greenY1").value),
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

heatmapMetric.addEventListener("change", () => {
   setGreenBoxDefaults();
   updateHeatmapChart();
   updateYearlyPassSummary();
   updateTrendChart();
   updateFaceVelocityChart();
});

/* -----------------------------

*/

function updateYearlyPassSummary() {
  const container = document.getElementById("yearlyPassContent");
  const building = buildingFilter.value;
  const metric = heatmapMetric.value || "ach";

  const y0Input = document.getElementById("greenY0");
  const y1Input = document.getElementById("greenY1");

  if (!building) {
    container.textContent = "Select a building to view yearly results.";
    return;
  }

  const metricField =
    metric === "faceVelocity" ? "faceVelocity" : "ach";

  const metricLabel =
    metric === "faceVelocity" ? "Face Velocity" : "ACH";

  const lowerLimit = Number(y0Input.value);
  const upperLimit = Number(y1Input.value);

  const start = startDate.value
    ? new Date(`${startDate.value}T00:00:00`)
    : null;

  const end = endDate.value
    ? new Date(`${endDate.value}T23:59:59.999`)
    : null;

  // Keep the summary consistent with the active dashboard filters.
  const records = allRecords.filter(record =>
    record.building === building &&
    (!labFilter.value || record.lab === labFilter.value) &&
    (!start || record.date >= start) &&
    (!end || record.date <= end) &&
    Number.isFinite(record[metricField])
  );

  // Keep the latest valid measurement for each hood in each year.
  const latestByHoodYear = new Map();

  records.forEach(record => {
    const year = record.date.getFullYear();
    const key = `${year}|${record.hoodKey}`;
    const existing = latestByHoodYear.get(key);

    if (!existing || record.date > existing.date ||
        (record.date.getTime() === existing.date.getTime() &&
         record.rowNumber > existing.rowNumber)) {
      latestByHoodYear.set(key, record);
    }
  });

  const yearly = new Map();

  latestByHoodYear.forEach(record => {
    const year = record.date.getFullYear();

    if (!yearly.has(year)) {
      yearly.set(year, {
        below: 0,
        within: 0,
        above: 0,
        total: 0
      });
    }

    const stats = yearly.get(year);
    const value = record[metricField];

    stats.total++;

    if (value < lowerLimit) {
      stats.below++;
    } else if (value > upperLimit) {
      stats.above++;
    } else {
      stats.within++;
    }
  });

  if (!yearly.size) {
    container.textContent =
      `No ${metricLabel} measurements match the current filters.`;
    return;
  }

  const percentage = (count, total) =>
    total ? `${(count / total * 100).toFixed(1)}%` : "—";

  const years = [...yearly.keys()].sort((a, b) => a - b);

  container.innerHTML = `
    <p>
      ${escapeHTML(building)} · ${metricLabel} ·
      Passing range: ${lowerLimit}–${upperLimit}
    </p>
    <div style="overflow-x:auto">
      <table class="yearly-pass-table">
        <thead>
          <tr>
            <th>Year</th>
            <th>Hoods Tested</th>
            <th>Below Range</th>
            <th>Within Range</th>
            <th>Above Range</th>
          </tr>
        </thead>
        <tbody>
          ${years.map(year => {
            const s = yearly.get(year);

            return `
              <tr>
                <td>${year}</td>
                <td>${s.total}</td>
                <td class="below-range">
                  ${percentage(s.below, s.total)}
                  (${s.below})
                </td>
                <td class="within-range">
                  ${percentage(s.within, s.total)}
                  (${s.within})
                </td>
                <td class="above-range">
                  ${percentage(s.above, s.total)}
                  (${s.above})
                </td>
              </tr>
            `;
          }).join("")}
        </tbody>
      </table>
    </div>
  `;
}


/* ---------------------------------------------------------
   Export
   --------------------------------------------------------- */

// Calculate statistics for one hood and one measurement.
function exportMetricStats(records, metric) {
  const valid = records
    .filter(r => Number.isFinite(r[metric]))
    .sort((a, b) =>
      a.date - b.date || a.rowNumber - b.rowNumber
    );

  if (!valid.length) {
    return {
      mean: "",
      stdDev: "",
      latest: "",
      deviation: "",
      yearComparison: ""
    };
  }

  const latest = valid[valid.length - 1];

  // Historical baseline: every valid inspection except
  // the latest, matching Performance Analysis.
  const history = valid.slice(0, -1);

  const mean = history.length
    ? history.reduce(
        (sum, r) => sum + r[metric],
        0
      ) / history.length
    : null;

  // Historical sample standard deviation (n - 1).
  const stdDev = history.length >= 2
    ? Math.sqrt(
        history.reduce(
          (sum, r) => sum + (r[metric] - mean) ** 2,
          0
        ) / (history.length - 1)
      )
    : null;

  const deviation = mean !== null
    ? latest[metric] - mean
    : null;

  // Keep the latest valid measurement in each year.
  const latestByYear = new Map();

  valid.forEach(record => {
    latestByYear.set(
      record.date.getFullYear(),
      record
    );
  });

  const years = [...latestByYear.keys()]
    .sort((a, b) => b - a);

  // Compare the most recent year with the preceding
  // available year, not two tests in the same year.
  const yearComparison = years.length >= 2
    ? latestByYear.get(years[0])[metric] -
      latestByYear.get(years[1])[metric]
    : "";

  return {
    mean: mean ?? "",
    stdDev: stdDev ?? "",
    latest: latest[metric],
    deviation: deviation ?? "",
    yearComparison
  };
}

// Export one summary row per hood.
exportBtn.addEventListener("click", () => {
  if (!summary.length) {
    return;
  }

  // Group currently filtered inspections by hood.
  const grouped = new Map();

  filteredRecords.forEach(record => {
    if (!grouped.has(record.hoodKey)) {
      grouped.set(record.hoodKey, []);
    }

    grouped.get(record.hoodKey).push(record);
  });

  const exportRows = summary.map(row => {
    const records = (grouped.get(row.hoodKey) || [])
      .slice()
      .sort((a, b) =>
        a.date - b.date || a.rowNumber - b.rowNumber
      );

    const ach = exportMetricStats(records, "ach");
    const velocity = exportMetricStats(
      records,
      "faceVelocity"
    );
    const cov = exportMetricStats(records, "cov");

    // Use the latest nonblank metadata for this hood.
    const latestMetadata = field => {
      const record = [...records].reverse().find(
        r => r[field] !== null &&
             r[field] !== undefined &&
             r[field] !== ""
      );

      return record ? record[field] : "";
    };

    return {
      // 1-5: Identification
      "Hood ID": row.hoodKey,
      "Building ID": row.building,
      "Lab ID": row.lab,
      "First Test Date": dateKey(row.firstInspection),
      "Latest Test Date": dateKey(row.latestInspection),

      // 6-9: ACH
      "Average ACH": ach.mean,
      "Stnd of ACH": ach.stdDev,
      "Latest ACH": ach.latest,
      "Mean vs Latest ACH":
        ach.deviation,

      // 10-13: Face velocity
      "Avg Face Vel.": velocity.mean,
      "Stnd of Face Vel.":
        velocity.stdDev,
      "Latest Face Vel.": velocity.latest,
      "Mean vs Latest Face Vel.":
        velocity.deviation,

      // 14-17: COV
      "Average COV": cov.mean,
      "Stnd of COV": cov.stdDev,
      "Latest COV": cov.latest,
      "Mean vs Latest COV":
        cov.deviation,

      // 18-20: Latest year versus previous available year
      "Current ACH vs Previous ACH":
        ach.yearComparison,
      "Current Face Vel. vs Previous Face Vel.":
        velocity.yearComparison,
      "Current COV vs Previous COV":
        cov.yearComparison,

      // 21-22: Uploaded metadata
      "Exhaust Type": latestMetadata("exhaustType"),
      "Total Area": latestMetadata("totalArea")
    };
  });

  // Generate CSV using SheetJS.
  const worksheet = XLSX.utils.json_to_sheet(exportRows);
  const csv = XLSX.utils.sheet_to_csv(worksheet);

  const blob = new Blob(
    [csv],
    { type: "text/csv;charset=utf-8;" }
  );

  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");

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

document.getElementById("updateGreenBox").onclick = () => {
    const metric = heatmapMetric.value || "ach";

    const y0 = Number(document.getElementById("greenY0").value);
    const y1 = Number(document.getElementById("greenY1").value);

    // Remember the user's range for this metric
    greenBoxRanges[metric].min = y0;
    greenBoxRanges[metric].max = y1;

    updateKPIs();
    updateTrendChart();
    updateFaceVelocityChart();
    updateHeatmapChart();
    updateYearlyPassSummary();
    updateWorstPerformers();
    updateAnalysis();
    updateBuildingComparison(); 
};


startDate.addEventListener(
  "change",
  applyFilters
);


endDate.addEventListener(
  "change",
  applyFilters
);

hoodFilter.addEventListener("change", () => {
    selectedHoodKey = hoodFilter.value;
     updateTrendChart();
     updateFaceVelocityChart();
     updateHeatmapChart();
  }
);

const resultsTable = resultsBody.closest("table");

resultsTable.querySelectorAll("thead th[data-sort]")
  .forEach(header => {

    header.addEventListener("click", () => {

      const column = header.dataset.sort;

      if (sortColumn === column) {
        sortDirection =
          sortDirection === "asc" ? "desc" : "asc";
      } else {
        sortColumn = column;
        sortDirection = "asc";
      }

      updateTable();

      // Update sorting indicators.
      resultsTable.querySelectorAll("th[data-sort]")
        .forEach(th => {
          th.classList.remove("sort-asc", "sort-desc");
        });

      header.classList.add(
        sortDirection === "asc" ? "sort-asc" : "sort-desc"
      );
    });
  });

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

   document.getElementById("yearlyPassContent").textContent =
  "Select a building to view yearly results.";

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
   
  tableCount.textContent = "";

  resultsBody.innerHTML = `
    <tr>
      <td colspan="7" class="empty-cell">
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

/* ---------------------------------------------------------
   Tabs, worst performers, and performance analysis
   --------------------------------------------------------- */
const analysisMetric = document.getElementById("analysisMetric");
const analysisHood = document.getElementById("analysisHood");
let activeTab = "summary";
let analysisSortColumn = "pct";
let analysisSortDirection = "asc";

document.querySelectorAll(".dashboard-tab").forEach(button => {
  button.addEventListener("click", () => {
    activeTab = button.dataset.tab;
    document.querySelectorAll(".dashboard-tab").forEach(tab => {
      const selected = tab === button;
      tab.classList.toggle("active", selected);
      tab.setAttribute("aria-selected", String(selected));
    });
    for (const name of ["summary", "history"]) {
      const element = document.getElementById(`${name}Tab`);
      element.hidden = name !== activeTab;
      element.classList.toggle("active", name === activeTab);
    }
    if (activeTab === "history") {
      requestAnimationFrame(() => {
        [trendChart, rankingChart, document.getElementById("analysisHistoryChart"),
          document.getElementById("analysisFleetChart")].forEach(chart => {
          if (chart && chart.data) Plotly.Plots.resize(chart);
        });
      });
    } else {
      requestAnimationFrame(() => {
        if (heatmapChart.data) Plotly.Plots.resize(heatmapChart);
      });
    }
  });
});

// Keep the ACH, face velocity, and performance analysis selections synchronized.
function selectHoodAndShowHistory(hoodKey) {
  selectedHoodKey = hoodKey;
  hoodFilter.value = hoodKey;
  analysisHood.value = hoodKey;
  document.querySelector('[data-tab="history"]').click();
  document.querySelector(".analysis-panel").open = true;
  updateTrendChart();
  updateFaceVelocityChart();
  updateAnalysis();
  requestAnimationFrame(() => {
    document.querySelector(".analysis-panel").scrollIntoView({behavior: "smooth", block: "start"});
  });
}

function updateWorstPerformers() {
  const { min, max } = greenBoxRanges.faceVelocity;
  const velocity = summary
    .filter(row => Number.isFinite(row.latestFaceVelocity))
    .map(row => ({
      ...row,
      velocityGroup: row.latestFaceVelocity < min ? 0 :
        row.latestFaceVelocity > max ? 1 : 2
    }))
    .sort((a, b) => {
      if (a.velocityGroup !== b.velocityGroup) return a.velocityGroup - b.velocityGroup;
      const difference = a.velocityGroup === 1
        ? b.latestFaceVelocity - a.latestFaceVelocity
        : a.latestFaceVelocity - b.latestFaceVelocity;
      return difference || naturalSort(a.hoodKey, b.hoodKey);
    });

  // Use the latest inspection containing a valid COV for each hood.
  const covByHood = new Map();
  [...filteredRecords]
    .sort((a, b) => a.date - b.date || a.rowNumber - b.rowNumber)
    .forEach(row => {
      if (Number.isFinite(row.cov)) covByHood.set(row.hoodKey, row);
    });
  const cov = [...covByHood.values()]
    .sort((a, b) => b.cov - a.cov || naturalSort(a.hoodKey, b.hoodKey));

  function render(id, rows, field, digits, unit, colored = false) {
    document.getElementById(id).innerHTML = rows.length ? rows.map(row => {
      const groupClass = colored
        ? ["velocity-low", "velocity-high", "velocity-good"][row.velocityGroup]
        : "";
      return `
        <button type="button" class="worst-row ${groupClass}" data-hood="${escapeHTML(row.hoodKey)}">
          <span>
            <strong>${escapeHTML(row.hood)}</strong>
            <small>${escapeHTML(row.building)} · ${escapeHTML(row.lab)} · ${formatDate(row.date || row.latestInspection)}</small>
          </span>
          <span class="worst-value">${formatNumber(row[field], digits)} ${unit}</span>
        </button>`;
    }).join("") : '<p class="muted">No matching measurements.</p>';

    document.querySelectorAll(`#${id} .worst-row`).forEach(button => {
      button.addEventListener("click", () => {
        selectHoodAndShowHistory(button.dataset.hood);
      });
    });
  }

  render("worstVelocity", velocity, "latestFaceVelocity", 1, "fpm", true);
  render("worstCOV", cov, "cov", 1, "%");
}

function getAnalysisRows() {
  const metric = analysisMetric.value;
  const grouped = new Map();
  [...filteredRecords].sort((a,b) => a.date - b.date || a.rowNumber - b.rowNumber)
    .forEach(record => {
      if (!Number.isFinite(record[metric])) return;
      if (!grouped.has(record.hoodKey)) grouped.set(record.hoodKey, []);
      grouped.get(record.hoodKey).push(record);
    });
  return [...grouped].map(([key, records]) => {
    const latest = records[records.length - 1];
    const previous = records.slice(0, -1);
    const mean = previous.length ? previous.reduce((sum,r) => sum + r[metric],0) / previous.length : null;
    const value = latest[metric];
    const deviation = mean !== null ? value - mean : null;
    const pct = mean !== null && mean !== 0 ? 100 * deviation / Math.abs(mean) : null;
    const sampleStdDev = previous.length >= 2 ? Math.sqrt(
      previous.reduce((sum,r) => sum + (r[metric] - mean) ** 2,0) / (previous.length - 1)
    ) : null;
    return {key, latest, records, mean, value, deviation, pct, sampleStdDev};
  });
}


function updateAnalysis() {
  const metric = analysisMetric.value;
  const rows = getAnalysisRows();
  const old = analysisHood.value || selectedHoodKey;

  // ---------------------------------------------------------
  // Hood Selection
  // ---------------------------------------------------------

  analysisHood.innerHTML =
    '<option value="">Select a hood</option>' +
    rows
      .slice()
      .sort((a, b) => naturalSort(a.key, b.key))
      .map(r =>
        `<option value="${escapeHTML(r.key)}">${escapeHTML(r.key)}</option>`
      )
      .join("");

  if (rows.some(r => r.key === old)) {
    analysisHood.value = old;
  }

  const current = rows.find(r => r.key === analysisHood.value);

  // ---------------------------------------------------------
  // Measurement Formatting
  // ---------------------------------------------------------

  const unit =
    metric === "ach"
      ? "ACH"
      : metric === "faceVelocity"
        ? "fpm"
        : "%";

  const fmt = value =>
    Number.isFinite(value)
      ? `${formatNumber(value, 2)} ${unit}`
      : "—";

  // ---------------------------------------------------------
  // Performance Analysis KPIs
  // ---------------------------------------------------------

  document.getElementById("analysisLatest").textContent =
    current ? fmt(current.value) : "—";

  document.getElementById("analysisMean").textContent =
    current ? fmt(current.mean) : "—";

  document.getElementById("analysisDeviation").textContent =
    current && Number.isFinite(current.pct)
      ? `${current.pct > 0 ? "+" : ""}${formatNumber(current.pct, 1)}% (${fmt(current.deviation)})`
      : "—";

  document.getElementById("analysisStdDev").textContent =
    current ? fmt(current.sampleStdDev) : "—";

  // Rating belongs to the latest inspection for this hood,
  // independent of which numeric metric is selected.
  const latestInspection = current
    ? filteredRecords
        .filter(record => record.hoodKey === current.key)
        .sort((a, b) => b.date - a.date || b.rowNumber - a.rowNumber)[0]
    : null;
  const ratingElement = document.getElementById("analysisTestRating");
  const rating = latestInspection?.testRating || "";
  ratingElement.textContent = rating || "—";
  ratingElement.classList.toggle("test-rating-pass", rating.toLowerCase() === "pass");
  ratingElement.classList.toggle("test-rating-fail", rating.toLowerCase() === "fail");

  document.getElementById("analysisInfo").textContent =
    current
      ? `${current.records.length} valid ${unit} measurements; ` +
        `${Math.max(0, current.records.length - 1)} used for historical mean. ` +
        `${formatDate(current.latest.date)} is the latest valid test.`
      : "Select a hood to view its historical statistics.";

  // ---------------------------------------------------------
  // Sort All Hoods Table
  // ---------------------------------------------------------

  const sortedRows = rows.slice().sort((a, b) => {
    const col = analysisSortColumn;

    let av =
      col === "date"
        ? a.latest.date.getTime()
        : ["building", "lab", "hood"].includes(col)
          ? a.latest[col]
          : a[col];

    let bv =
      col === "date"
        ? b.latest.date.getTime()
        : ["building", "lab", "hood"].includes(col)
          ? b.latest[col]
          : b[col];

    const aMissing =
      av === null ||
      av === undefined ||
      (typeof av === "number" && !Number.isFinite(av));

    const bMissing =
      bv === null ||
      bv === undefined ||
      (typeof bv === "number" && !Number.isFinite(bv));

    if (aMissing !== bMissing) {
      return aMissing ? 1 : -1;
    }

    let cmp = aMissing
      ? 0
      : typeof av === "string"
        ? naturalSort(av, bv)
        : av - bv;

    return (
      (analysisSortDirection === "asc" ? cmp : -cmp) ||
      naturalSort(a.key, b.key)
    );
  });

  // ---------------------------------------------------------
  // Populate All Hoods Table
  // ---------------------------------------------------------

  document.getElementById("analysisBody").innerHTML =
    rows.length
      ? sortedRows
          .map(
            r => `
              <tr>
                <td>${escapeHTML(r.latest.building)}</td>
                <td>${escapeHTML(r.latest.lab)}</td>

                <td>
                  <button
                    type="button"
                    class="table-hood-button analysis-select"
                    data-hood="${escapeHTML(r.key)}"
                  >
                    ${escapeHTML(r.latest.hood)}
                  </button>
                </td>

                <td>${fmt(r.value)}</td>
                <td>${fmt(r.mean)}</td>

                <td>
                  ${
                    Number.isFinite(r.pct)
                      ? `${r.pct > 0 ? "+" : ""}${formatNumber(r.pct, 1)}%`
                      : "—"
                  }
                </td>

                <td>${formatDate(r.latest.date)}</td>
              </tr>
            `
          )
          .join("")
      : `
          <tr>
            <td colspan="7" class="empty-cell">
              No measurements match the current filters.
            </td>
          </tr>
        `;

  // ---------------------------------------------------------
  // Update Table Sorting Indicators
  // ---------------------------------------------------------

  document.querySelectorAll("[data-analysis-sort]").forEach(th => {
    const isSelected =
      th.dataset.analysisSort === analysisSortColumn;

    th.classList.toggle(
      "sort-asc",
      isSelected && analysisSortDirection === "asc"
    );

    th.classList.toggle(
      "sort-desc",
      isSelected && analysisSortDirection === "desc"
    );

    th.setAttribute(
      "aria-sort",
      isSelected
        ? analysisSortDirection === "asc"
          ? "ascending"
          : "descending"
        : "none"
    );
  });

  // ---------------------------------------------------------
  // Hood Selection from All Hoods Table
  // ---------------------------------------------------------

  document.querySelectorAll(".analysis-select").forEach(button => {
    button.addEventListener("click", () => {
      analysisHood.value = button.dataset.hood;
      selectedHoodKey = button.dataset.hood;
      hoodFilter.value = selectedHoodKey;

      updateTrendChart();
      updateFaceVelocityChart();
      updateAnalysis();
    });
  });

  // ---------------------------------------------------------
  // Chart Elements
  // ---------------------------------------------------------

  const hist = document.getElementById("analysisHistoryChart");
  const fleet = document.getElementById("analysisFleetChart");

  // ---------------------------------------------------------
  // Selected Hood History Chart
  // ---------------------------------------------------------

  const historyTraces = current
    ? [
        {
          type: "scatter",
          mode: "lines+markers",
          name: "Measurements",
          x: current.records.map(r => r.date),
          y: current.records.map(r => r[metric])
        }
      ]
    : [];

  // Add historical mean reference line.
  if (current && Number.isFinite(current.mean)) {
    historyTraces.push({
      type: "scatter",
      mode: "lines",
      name: "Historical mean",

      x: [
        current.records[0].date,
        current.latest.date
      ],

      y: [
        current.mean,
        current.mean
      ],

      line: {
        dash: "dash",
        color: "#2864dc"
      }
    });
  }

  // ---------------------------------------------------------
  // Green Passing Range
  // ---------------------------------------------------------

  // Use the same user-configured green ranges as the
  // original ACH and face velocity trend charts.
  //
  // COV has no defined passing range and is not shaded.

  const range =
    metric === "ach" || metric === "faceVelocity"
      ? greenBoxRanges[metric]
      : null;

  const validRange =
    range &&
    Number.isFinite(range.min) &&
    Number.isFinite(range.max) &&
    range.min < range.max;

  const rangeNote = document.getElementById("analysisRangeNote");

  rangeNote.textContent = validRange
    ? `Green band: ${formatNumber(range.min, 2)}–${formatNumber(range.max, 2)} ${unit} (configured on Summary).`
    : metric === "cov"
      ? "No green band for COV: a passing range has not been defined."
      : "No valid green range is configured for this measurement.";

  const rangeShapes = validRange
    ? [
        {
          type: "rect",

          xref: "paper",
          x0: 0,
          x1: 1,

          yref: "y",
          y0: range.min,
          y1: range.max,

          fillcolor: "rgba(34, 197, 94, 0.20)",

          line: {
            color: "rgba(34, 197, 94, 0.5)",
            width: 1
          },

          layer: "below"
        }
      ]
    : [];

  // ---------------------------------------------------------
  // Render Selected Hood History
  // ---------------------------------------------------------

  Plotly.react(
    hist,
    historyTraces,

    {
      ...baseChartLayout(
        current
          ? `${current.latest.hood} — ${unit} history`
          : "Select a hood"
      ),

      xaxis: {
        title: "Inspection date",
        type: "date"
      },

      yaxis: {
        title: unit
      },

      shapes: rangeShapes
    },

    plotConfig()
  );

  // =========================================================
  // Fleet Comparison Chart
  // =========================================================

  const comparison =
    document.getElementById("fleetComparisonMetric").value;

  // Keep only hoods with valid comparison data.
  const fleetRows = rows
    .filter(r => Number.isFinite(r[comparison]))
    .sort(
      (a, b) =>
        a[comparison] - b[comparison] ||
        naturalSort(a.key, b.key)
    );

  // Keep the chart scrollable rather than allowing
  // the surrounding panel to grow indefinitely.
  const plotHeight = Math.max(
    360,
    120 + fleetRows.length * 29
  );

  // Currently selected hood.
  const selectedKey =
    analysisHood.value || selectedHoodKey;

  // ---------------------------------------------------------
  // Fleet Comparison Colors
  // ---------------------------------------------------------

  // Change these two values to customize the chart colors.

  const selectedHoodColor = "#8050c8"; // Purple
  const otherHoodColor = "#d2e1f3";    // Blue

  // ---------------------------------------------------------
  // Render Fleet Comparison
  // ---------------------------------------------------------

  Plotly.react(
    fleet,

    fleetRows.length
      ? [
          {
            type: "bar",
            orientation: "h",

            // Comparison values.
            x: fleetRows.map(r => r[comparison]),

            // Building, lab and hood labels.
            y: fleetRows.map(
              r =>
                `${r.latest.building} · ${r.latest.lab} · ${r.latest.hood}`
            ),

            // Purple for selected hood.
            // Blue for all other hoods.
            marker: {
              color: fleetRows.map(r =>
                r.key === selectedKey
                  ? selectedHoodColor
                  : otherHoodColor
              )
            },

            // Hover information.
            hovertext: fleetRows.map(r => r.key),
            hoverinfo: "x+text"
          }
        ]
      : [],

    {
      ...baseChartLayout(
        {
          value: "Latest measurement",
          mean: "Historical mean",
          pct: "Deviation from historical mean"
        }[comparison]
      ),

      height: plotHeight,

      margin: {
        l: 175,
        r: 25,
        t: 55,
        b: 65
      },

      xaxis: {
        title:
          comparison === "pct"
            ? "Deviation (%)"
            : unit
      },

      yaxis: {
        automargin: true,
        autorange: "reversed"
      }
    },

    {
      ...plotConfig(),
      responsive: false
    }
  );
}

// ===========================================================
// Performance Analysis Event Listeners
// ===========================================================

// Update analysis when measurement changes.
analysisMetric.addEventListener("change", updateAnalysis);

// Update fleet chart when comparison metric changes.
document
  .getElementById("fleetComparisonMetric")
  .addEventListener("change", updateAnalysis);

// -----------------------------------------------------------
// All Hoods Table Sorting
// -----------------------------------------------------------

document.querySelectorAll("[data-analysis-sort]").forEach(th => {
  th.tabIndex = 0;
  th.setAttribute("role", "button");

  const sort = () => {
    if (analysisSortColumn === th.dataset.analysisSort) {
      analysisSortDirection =
        analysisSortDirection === "asc"
          ? "desc"
          : "asc";
    } else {
      analysisSortColumn = th.dataset.analysisSort;
      analysisSortDirection = "asc";
    }

    updateAnalysis();
  };

  th.addEventListener("click", sort);

  th.addEventListener("keydown", event => {
    if (event.key === "Enter" || event.key === " ") {
      event.preventDefault();
      sort();
    }
  });
});

// -----------------------------------------------------------
// Hood Selector
// -----------------------------------------------------------

analysisHood.addEventListener("change", () => {
  selectedHoodKey = analysisHood.value;
  hoodFilter.value = selectedHoodKey;

  updateTrendChart();
  updateFaceVelocityChart();
  updateAnalysis();
});

// -----------------------------------------------------------
// Resize Charts When Performance Analysis Is Expanded
// -----------------------------------------------------------

document
  .querySelector(".analysis-panel")
  .addEventListener("toggle", event => {
    if (event.target.open && activeTab === "history") {
      requestAnimationFrame(() => {
        for (const id of [
          "analysisHistoryChart",
          "analysisFleetChart"
        ]) {
          const chart = document.getElementById(id);

          if (chart.data) {
            Plotly.Plots.resize(chart);
          }
        }
      });
    }
  });
