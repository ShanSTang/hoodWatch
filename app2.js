  const rowKeys = Object.keys(row);
  const normalizedKeys = new Map(
    rowKeys.map(key => [key.trim().toLowerCase(), key])
  );

  for (const candidate of candidates) {
    const actualKey = normalizedKeys.get(candidate.toLowerCase());

    if (actualKey !== undefined) {
      return row[actualKey];
    }
  }

  return null;
}


function parseRows(rows) {
  const parsed = [];

  rows.forEach((row, index) => {
    const building = normalizeValue(row["Building"], "Unknown Building");
    const lab = normalizeValue(row["Lab"], "Unknown Lab");
    const hood = normalizeValue(row["Hood"], "Unknown Hood");

    const date = parseDate(row["Date"]);

    if (!date) {
      return;
    }

    const ach = parseACH(row["Min ACH"]);
    const faceVelocity = parseACH(row["FaceV Avg"]);

    const cov = parsePercentage(
      getOptionalField(row, [
        "Average_Zone_COV",
        "Average COV",
        "Average_COV",
        "Average Zone COV"
      ])
    );

    const cycle = normalizeValue(
      getOptionalField(row, [
        "Cycle",
        "Test Cycle",
        "Cycle Number",
        "HoodTestCycle_CycleNumber"
      ]),
      "—"
    );

    const exhaustType = normalizeValue(
      getOptionalField(row, [
        "Exhaust Type",
        "ExhaustType",
        "Exhaust",
        "Exhaust_Type"
      ]),
      "—"
    );

    const hoodKey = `${building} | ${lab} | ${hood}`;

    parsed.push({
      rowNumber: index,
      building,
      lab,
      hood,
      hoodKey,
      date,
      dateString: dateKey(date),
      cycle,
      exhaustType,
      cov,
      faceVelocity,
      ach,
      previousACH: NaN,
      baselineACH: NaN,
      achDrop: NaN,
      achDropPercent: NaN
    });
  });

  return parsed;
}


/* ---------------------------------------------------------
   Filters
   --------------------------------------------------------- */

function initializeFilters() {
  const buildings = [...new Set(
    allRecords.map(record => record.building)
  )].sort();

  const labs = [...new Set(
    allRecords.map(record => record.lab)
  )].sort();

  buildingFilter.innerHTML =
    `<option value="">All Buildings</option>` +
    buildings.map(value =>
      `<option value="${escapeHTML(value)}">${escapeHTML(value)}</option>`
    ).join("");

  labFilter.innerHTML =
    `<option value="">All Labs</option>` +
    labs.map(value =>
      `<option value="${escapeHTML(value)}">${escapeHTML(value)}</option>`
    ).join("");

  if (hoodFilter) {
    const hoods = [...new Set(
      allRecords.map(record => record.hood)
    )].sort();

    hoodFilter.innerHTML =
      `<option value="">All Hoods</option>` +
      hoods.map(value =>
        `<option value="${escapeHTML(value)}">${escapeHTML(value)}</option>`
      ).join("");
  }

  startDate.value = "";
  endDate.value = "";
}


function applyFilters() {
  const selectedBuilding = buildingFilter.value;
  const selectedLab = labFilter.value;
  const selectedHood = hoodFilter ? hoodFilter.value : "";
  const fromDate = startDate.value;
  const toDate = endDate.value;

  filteredRecords = allRecords.filter(record => {
    if (selectedBuilding && record.building !== selectedBuilding) {
      return false;
    }

    if (selectedLab && record.lab !== selectedLab) {
      return false;
    }

    if (selectedHood && record.hood !== selectedHood) {
      return false;
    }

    if (fromDate && record.dateString < fromDate) {
      return false;
    }

    if (toDate && record.dateString > toDate) {
      return false;
    }

    return true;
  });

  summary = calculateSummary(filteredRecords);

  updateDashboard();
}


buildingFilter.addEventListener("change", () => {
  updateLabFilter();
  applyFilters();
});


labFilter.addEventListener("change", applyFilters);

if (hoodFilter) {
  hoodFilter.addEventListener("change", applyFilters);
}

startDate.addEventListener("change", applyFilters);
endDate.addEventListener("change", applyFilters);


function updateLabFilter() {
  const selectedBuilding = buildingFilter.value;

  const labs = [...new Set(
    allRecords
      .filter(record =>
        !selectedBuilding || record.building === selectedBuilding
      )
      .map(record => record.lab)
  )].sort();

  const currentLab = labFilter.value;

  labFilter.innerHTML =
    `<option value="">All Labs</option>` +
    labs.map(value =>
      `<option value="${escapeHTML(value)}">${escapeHTML(value)}</option>`
    ).join("");

  if (labs.includes(currentLab)) {
    labFilter.value = currentLab;
  }
}


resetFiltersBtn.addEventListener("click", () => {
  buildingFilter.value = "";
  labFilter.value = "";

  if (hoodFilter) {
    hoodFilter.value = "";
  }

  startDate.value = "";
  endDate.value = "";

  updateLabFilter();
  applyFilters();
});


/* ---------------------------------------------------------
   Summary calculations
   --------------------------------------------------------- */

function calculateSummary(records) {
  const groups = new Map();

  records.forEach(record => {
    if (!groups.has(record.hoodKey)) {
      groups.set(record.hoodKey, []);
    }

    groups.get(record.hoodKey).push(record);
  });

  const summaries = [];

  groups.forEach((hoodRecords, hoodKey) => {
    const sorted = [...hoodRecords].sort((a, b) => {
      const dateDifference = a.date - b.date;

      return dateDifference !== 0
        ? dateDifference
        : a.rowNumber - b.rowNumber;
    });

    const achRecords = sorted.filter(record =>
      Number.isFinite(record.ach)
    );

    const faceVelocityRecords = sorted.filter(record =>
      Number.isFinite(record.faceVelocity)
    );

    let baselineACH = NaN;
    let previousACH = NaN;

    achRecords.forEach(record => {
      if (!Number.isFinite(baselineACH)) {
        baselineACH = record.ach;
      }

      record.baselineACH = baselineACH;
      record.previousACH = previousACH;

      if (Number.isFinite(previousACH)) {
        record.achDrop = previousACH - record.ach;
        record.achDropPercent =
          previousACH !== 0
            ? (record.achDrop / previousACH) * 100
            : NaN;
      }

      previousACH = record.ach;
    });

    const firstInspection = sorted[0] || null;
    const latestInspection = sorted[sorted.length - 1] || null;

    const firstACHRecord = achRecords[0] || null;
    const latestACHRecord = achRecords[achRecords.length - 1] || null;

    const latestFaceVelocityRecord =
      faceVelocityRecords[faceVelocityRecords.length - 1] || null;

    const latestCOVRecord = [...sorted].reverse().find(record =>
      Number.isFinite(record.cov)
    ) || null;

    const latestExhaustRecord = [...sorted].reverse().find(record =>
      record.exhaustType !== "—"
    ) || null;

    const latestCycleRecord = [...sorted].reverse().find(record =>
      record.cycle !== "—"
    ) || null;

    const latestACH = latestACHRecord
      ? latestACHRecord.ach
      : NaN;

    const latestFaceVelocity = latestFaceVelocityRecord
      ? latestFaceVelocityRecord.faceVelocity
      : NaN;

    const achDrop = latestACHRecord
      ? latestACHRecord.achDrop
      : NaN;

    const achDropPercent = latestACHRecord
      ? latestACHRecord.achDropPercent
      : NaN;

    summaries.push({
      hoodKey,
      building: latestInspection.building,
      lab: latestInspection.lab,
      hood: latestInspection.hood,

      records: sorted,
      achRecords,
      faceVelocityRecords,

      firstInspection,
      latestInspection,
      firstInspectionDate: firstInspection.date,
      latestInspectionDate: latestInspection.date,

      firstACHRecord,
      latestACHRecord,
      latestFaceVelocityRecord,

      latestACH,
      latestFaceVelocity,
      latestACHDate: latestACHRecord
        ? latestACHRecord.date
        : null,
      latestFaceVelocityDate: latestFaceVelocityRecord
        ? latestFaceVelocityRecord.date
        : null,

      latestCOV: latestCOVRecord
        ? latestCOVRecord.cov
        : NaN,

      latestExhaustType: latestExhaustRecord
        ? latestExhaustRecord.exhaustType
        : "—",

      latestCycle: latestCycleRecord
        ? latestCycleRecord.cycle
        : "—",

      firstACH: firstACHRecord
        ? firstACHRecord.ach
        : NaN,

      achDrop,
      achDropPercent,

      inspectionCount: sorted.length,
      achCount: achRecords.length,
      faceVelocityCount: faceVelocityRecords.length
    });
  });

  return summaries;
}

  const keys = Object.keys(row);
  for (const candidate of candidates) {
    const match = keys.find(
      key => key.trim().toLowerCase() === candidate.toLowerCase()
    );
    if (match !== undefined && row[match] !== null && row[match] !== "") {
      return row[match];
    }
  }
  return null;
}

function parseRows(rows) {
  const parsed = [];

  rows.forEach((row, index) => {
    const building = normalizeValue(row["Building"], "Unknown Building");
    const lab = normalizeValue(row["Lab"], "Unknown Lab");
    const hood = normalizeValue(row["Hood"], "Unknown Hood");

    const date = parseDate(row["Date"]);
    const ach = parseACH(row["Min ACH"]);
    const covValue = getOptionalField(row, [
      "Average_Zone_COV", "Average COV", "Average_COV", "Average Zone COV"
    ]);
    const cov = parsePercentage(covValue);
    const faceVelocity = parseACH(row["FaceV Avg"]);

    const cycleValue = getOptionalField(row, [
      "Cycle", "Test Cycle", "Cycle Number", "HoodTestCycle_CycleNumber"
    ]);
    const exhaustValue = getOptionalField(row, [
      "Exhaust Type", "ExhaustType", "Exhaust", "Exhaust_Type"
    ]);

    // Keep dated inspection rows even when one or both measurement types are absent.
    if (!date) return;

    // Identity is the combination of hood, lab, and building names.
    const hoodKey = `${building} | ${lab} | ${hood}`;

    parsed.push({
      rowNumber: index + 2,
      building,
      lab,
      hood,
      date,
      dateString: dateKey(date),
      cycle: cycleValue === null ? null : normalizeValue(cycleValue),
      exhaustType: exhaustValue === null ? null : normalizeValue(exhaustValue),
      cov: Number.isFinite(cov) ? cov : null,
      faceVelocity: Number.isFinite(faceVelocity) ? faceVelocity : null,
      ach: Number.isFinite(ach) ? ach : null,
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
  // Group every inspection (including face-velocity-only inspections) by hood.
  const groups = new Map();

  records.forEach(record => {
    if (!groups.has(record.hoodKey)) groups.set(record.hoodKey, []);
    groups.get(record.hoodKey).push(record);
  });

  const results = [];

  for (const [hoodKey, hoodRecords] of groups) {
    const sortedAll = [...hoodRecords].sort(
      (a, b) => a.date - b.date || a.rowNumber - b.rowNumber
    );

    const achRecords = sortedAll.filter(record => Number.isFinite(record.ach));
    const faceRecords = sortedAll.filter(
      record => Number.isFinite(record.faceVelocity)
    );

    let previousSum = 0;
    const enrichedACH = achRecords.map((record, index) => {
      const baselineACH = index > 0 ? previousSum / index : null;
      const previousACH = index > 0 ? achRecords[index - 1].ach : null;

      const achDropBaseline =
        baselineACH !== null ? record.ach - baselineACH : null;
      const achDropPrevious =
        previousACH !== null ? record.ach - previousACH : null;

      const enriched = {
        ...record,
        baselineACH,
        previousACH,
        achDropBaseline,
        achDropPrevious,
        achDropPercentBaseline:
          baselineACH !== null && baselineACH > 0
            ? 100 * (record.ach - baselineACH) / baselineACH
            : null,
        achDropPercentPrevious:
          previousACH !== null && previousACH > 0
            ? 100 * (record.ach - previousACH) / previousACH
            : null
      };

      previousSum += record.ach;
      return enriched;
    });

    const first = sortedAll[0];
    const latest = sortedAll[sortedAll.length - 1];
    const latestACHRecord = enrichedACH.length
      ? enrichedACH[enrichedACH.length - 1]
      : null;
    const latestFaceRecord = faceRecords.length
      ? faceRecords[faceRecords.length - 1]
      : null;

    results.push({
      hoodKey,
      hood: latest.hood,
      building: latest.building,
      lab: latest.lab,

      // First/latest inspection dates include all valid inspection records.
      firstInspection: first.date,
      latestInspection: latest.date,
      firstRecord: first,
      latestRecord: latest,
      inspectionCount: sortedAll.length,

      // Metric-specific latest values may come from different inspection dates.
      latestACH: latestACHRecord?.ach ?? null,
      latestACHDate: latestACHRecord?.date ?? null,
      latestACHRecord,
      latestFaceVelocity: latestFaceRecord?.faceVelocity ?? null,
      latestFaceVelocityDate: latestFaceRecord?.date ?? null,
      latestFaceVelocityRecord: latestFaceRecord,

      latestCov: latest.cov,
      latestExhaustType: latest.exhaustType,
      latestCycle: latest.cycle,

      baselineACH: latestACHRecord?.baselineACH ?? null,
      previousACH: latestACHRecord?.previousACH ?? null,
      achDropBaseline: latestACHRecord?.achDropBaseline ?? null,
      achDropPercentBaseline:
        latestACHRecord?.achDropPercentBaseline ?? null,
      achDropPrevious: latestACHRecord?.achDropPrevious ?? null,
      achDropPercentPrevious:
        latestACHRecord?.achDropPercentPrevious ?? null,

      achInspectionCount: achRecords.length,
      trendAvailable: achRecords.length >= 2
    });
  }

  return results.sort((a, b) => {
    const achA = Number.isFinite(a.latestACH) ? a.latestACH : -Infinity;
    const achB = Number.isFinite(b.latestACH) ? b.latestACH : -Infinity;
    return achB - achA || naturalSort(a.hoodKey, b.hoodKey);
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
      row.latestACH < 200 ||       row.latestACH > 400
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
  updateYearlyPassSummary();
}

function updateHeatmap() {
  updateHeatmapChart();
}


/* ---------------------------------------------------------
   KPIs
   --------------------------------------------------------- */

function updateKPIs() {

  /*const flagged = summary.filter(
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
    `${flagged.length} flagged`;*/

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
  const flagged = summary.filter(
    row => getDeteriorationStatus(row).flagged
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

  flagCount.textContent =
    `${flagged.length} flagged`; 
}


/* ---------------------------------------------------------
   Flagged hood list
   --------------------------------------------------------- */

function updateFlaggedHoods() {
  const flagged = summary
    .map(row => ({ row, status: getDeteriorationStatus(row) }))
    .filter(item => item.status.flagged);

  if (!flagged.length) {
    hoodList.innerHTML = `
      <div class="empty-state">
        No hoods currently meet the deterioration rules.
      </div>
    `;
    return;
  }

  hoodList.innerHTML = flagged.map(({ row, status }) => `
    <button
      class="hood-item"
      data-hood-key="${escapeHTML(row.hoodKey)}"
      type="button"
    >
      <div class="hood-item-main">
        <strong>${escapeHTML(row.hood)}</strong>
        <span>
          ${escapeHTML(row.lab)} · ${escapeHTML(row.building)}
        </span>
      </div>

      <div class="hood-item-reason">
        <strong>Reason:</strong>
        ${status.reasons.map(reason => escapeHTML(reason)).join(" · ")}
      </div>

      <div class="hood-item-reason">
        <strong>Flagged by test:</strong>
        ${formatDate(row.latestACHDate)}
      </div>
    </button>
  `).join("");

  hoodList.querySelectorAll(".hood-item").forEach(button => {
    button.addEventListener("click", () => {
      selectedHoodKey = button.dataset.hoodKey;
      hoodFilter.value = selectedHoodKey;
      updateTrendChart();
      updateFaceVelocityChart();
      trendChart.scrollIntoView({ behavior: "smooth", block: "center" });
    });
  });
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
        valueA = getDeteriorationStatus(a).flagged ? 1 : 0;
        valueB = getDeteriorationStatus(b).flagged ? 1 : 0;
        break;

      case "building":
      case "lab":
      case "hood":
        valueA = a[sortColumn] || "";
        valueB = b[sortColumn] || "";
        break;

      case "firstInspection":
      case "latestInspection":
      case "latestACHDate":
      case "latestFaceVelocityDate":
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
  tableCount.textContent = `${summary.length} hoods`;

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

  resultsBody.innerHTML = sortSummary(summary).map(row => `
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
      <td>${formatNumber(row.latestACH, 1)}</td>
      <td>${formatNumber(row.latestFaceVelocity, 2)}</td>
    </tr>
  `).join("");
    resultsBody.querySelectorAll(".table-hood-button").forEach(button => {
    button.addEventListener("click", () => {
      selectedHoodKey = button.dataset.hoodKey;
      hoodFilter.value = selectedHoodKey;
      updateTrendChart();
      updateFaceVelocityChart();
      trendChart.scrollIntoView({ behavior: "smooth", block: "center" });
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
    .filter(record =>
      record.hoodKey === selectedHoodKey &&
      Number.isFinite(record.ach)
    )
    .sort((a, b) => a.date - b.date || a.rowNumber - b.rowNumber);


  if (!records.length) {
    Plotly.react(
      trendChart,
      [],
      baseChartLayout("No ACH data available for this hood"),
      plotConfig()
    );
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

  if (Number.isFinite(selectedSummary?.baselineACH)) {

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

function setGreenBoxDefaults() {
  const y0Input = document.getElementById("greenY0");
  const y1Input = document.getElementById("greenY1");
  const metric = heatmapMetric.value || "ach";

  if (metric === "ach") {
    y0Input.value = 200;
    y1Input.value = 400;
  } else if (metric === "faceVelocity") {
    y0Input.value = 80;
    y1Input.value = 120;
  }
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

      "Latest ACH Date":
        dateKey(row.latestACHDate),

      "Latest ACH":
        row.latestACH,

      "Latest Face Velocity Date":
        dateKey(row.latestFaceVelocityDate),

      "Latest Face Velocity":
        row.latestFaceVelocity,

      "Latest Average COV":
        row.latestCov,

      "Latest Exhaust Type":
        row.latestExhaustType,

      "Latest Cycle":
        row.latestCycle,

      "Baseline ACH":
                row.baselineACH,

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

document.getElementById("updateGreenBox").onclick = () => {
   const y0 = Number(document.getElementById("greenY0").value);
   const y1 = Number(document.getElementById("greenY1").value);

   updateHeatmapChart();
   updateYearlyPassSummary();
}


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
