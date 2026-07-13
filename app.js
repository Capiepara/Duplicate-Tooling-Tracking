(() => {
  "use strict";

  const WEEK_WIDTH = 42;
  const DAY_WIDTH = WEEK_WIDTH / 7;
  const MS_DAY = 86400000;
  const DEFAULT_REPORT_START = new Date(2026, 5, 1);

  const STAGES = [
    {
      key: "opm",
      name: "OPM Kick Off",
      color: "#12A0AA",
      aliases: ["opm kick off", "opm kickoff", "opm"],
    },
    {
      key: "drawing",
      name: "2D/3D Drawing",
      color: "#900C3F",
      aliases: ["2d/3d drawing", "2d 3d drawing", "2d", "3d"],
    },
    {
      key: "toolingSpec",
      name: "Tooling Spec Date",
      color: "#182B55",
      aliases: ["tooling spec date", "tooling spec"],
    },
    {
      key: "quotation",
      name: "Quotation",
      color: "#4871CC",
      aliases: ["quotation"],
    },
    {
      key: "car",
      name: "CAR",
      color: "#5F4E94",
      aliases: ["car"],
    },
    {
      key: "por",
      name: "POR",
      color: "#A291C7",
      aliases: ["por"],
    },
    {
      key: "toolingBuild",
      name: "Tooling Build",
      color: "#82CBEC",
      aliases: ["tooling build", "tooling buid"],
    },
    {
      key: "fai",
      name: "FAI",
      color: "#D94F21",
      aliases: ["fai"],
    },
    {
      key: "qbBuilding",
      name: "QB Building",
      color: "#FEBD2B",
      aliases: ["qb building"],
    },
    {
      key: "qbTesting",
      name: "QB Testing Complete",
      color: "#9AAB4B",
      aliases: [
        "qb testing complete date",
        "qb testing complete",
        "qb testing",
      ],
    },
    {
      key: "pr",
      name: "PR Building",
      color: "#D6568C",
      aliases: ["pr building"],
    },
    {
      key: "ser",
      name: "SER Fully Release",
      color: "#009999",
      aliases: [
        "ser fully release date",
        "ser fully release",
        "ser release",
      ],
    },
  ];

  const state = {
    workbook: null,
    rows: [],
    parts: [],
    weeks: [],
    timelineStart: null,
    timelineEnd: null,
  };

  const dom = {
    excelFile: get("excelFile"),
    fileName: get("fileName"),
    sheetSelect: get("sheetSelect"),
    reportStart: get("reportStart"),
    loadSheetBtn: get("loadSheetBtn"),
    partFilter: get("partFilter"),
    stageFilter: get("stageFilter"),
    dateFrom: get("dateFrom"),
    dateTo: get("dateTo"),
    delayOnly: get("delayOnly"),
    resetFiltersBtn: get("resetFiltersBtn"),
    legend: get("legend"),
    summaryText: get("summaryText"),
    emptyState: get("emptyState"),
    ganttViewport: get("ganttViewport"),
    ganttCanvas: get("ganttCanvas"),
    tooltip: get("tooltip"),
    kpiPart: get("kpiPart"),
    kpiStage: get("kpiStage"),
    kpiStageStrip: get("kpiStageStrip"),
    kpiFinish: get("kpiFinish"),
    kpiDelay: get("kpiDelay"),
    kpiDelayFoot: get("kpiDelayFoot"),
    kpiProgress: get("kpiProgress"),
    progressBar: get("progressBar"),
  };

  init();

  function init() {
    renderLegend();
    renderStageFilter();

    dom.excelFile.addEventListener("change", handleExcelImport);
    dom.loadSheetBtn.addEventListener("click", loadSelectedSheet);
    dom.reportStart.addEventListener("change", () => {
      if (state.parts.length) {
        rebuildTimeline();
        render();
      }
    });

    [
      dom.partFilter,
      dom.stageFilter,
      dom.dateFrom,
      dom.dateTo,
      dom.delayOnly,
    ].forEach((element) => {
      element.addEventListener("change", render);
    });

    dom.resetFiltersBtn.addEventListener("click", resetFilters);
  }

  function handleExcelImport(event) {
    const file = event.target.files && event.target.files[0];

    if (!file) {
      return;
    }

    dom.fileName.textContent = file.name;

    const reader = new FileReader();

    reader.onerror = () => {
      alert("Cannot read this Excel file.");
    };

    reader.onload = (loadEvent) => {
      try {
        const bytes = new Uint8Array(loadEvent.target.result);

        state.workbook = XLSX.read(bytes, {
          type: "array",
          cellDates: true,
        });

        dom.sheetSelect.innerHTML = "";

        state.workbook.SheetNames.forEach((sheetName) => {
          const option = document.createElement("option");
          option.value = sheetName;
          option.textContent = sheetName;
          dom.sheetSelect.appendChild(option);
        });

        dom.sheetSelect.disabled = false;
        dom.loadSheetBtn.disabled = false;
      } catch (error) {
        console.error(error);
        alert("The Excel file could not be parsed.");
      }
    };

    reader.readAsArrayBuffer(file);
  }

  function loadSelectedSheet() {
    if (!state.workbook) {
      return;
    }

    const sheetName = dom.sheetSelect.value;
    const worksheet = state.workbook.Sheets[sheetName];

    if (!worksheet || !worksheet["!ref"]) {
      alert("The selected sheet is empty.");
      return;
    }

    try {
      const data = worksheetToMatrix(worksheet);
      const header = locateHeaderRows(data);

      if (!header) {
        throw new Error(
          "Cannot find the stage header and Target / Estimate / Actual row."
        );
      }

      const columnMap = buildColumnMap(
        data[header.stageRowIndex],
        data[header.subHeaderRowIndex]
      );

      const rawRows = parseDataRows(
        data,
        header.subHeaderRowIndex + 1,
        columnMap
      );

      state.parts = buildPartSchedules(rawRows);
      populatePartFilter();
      rebuildTimeline();
      render();
    } catch (error) {
      console.error(error);
      alert(error.message || "Cannot build the Gantt chart.");
    }
  }

  function worksheetToMatrix(worksheet) {
    const sourceRange = XLSX.utils.decode_range(worksheet["!ref"]);

    sourceRange.s.c = 0;
    sourceRange.e.c = Math.min(35, sourceRange.e.c);

    return XLSX.utils.sheet_to_json(worksheet, {
      header: 1,
      range: XLSX.utils.encode_range(sourceRange),
      defval: "",
      raw: true,
    });
  }

  function locateHeaderRows(data) {
    const maxRows = Math.min(data.length, 20);

    for (let rowIndex = 0; rowIndex < maxRows - 1; rowIndex += 1) {
      const stageText = data[rowIndex].map(normalize).join(" ");
      const subHeaderText = data[rowIndex + 1].map(normalize).join(" ");

      const stageMatches = STAGES.filter((stage) =>
        stage.aliases.some((alias) => stageText.includes(alias))
      ).length;

      const hasSubHeaders =
        subHeaderText.includes("target") &&
        (subHeaderText.includes("actual") ||
          subHeaderText.includes("estimate"));

      if (stageMatches >= 3 && hasSubHeaders) {
        return {
          stageRowIndex: rowIndex,
          subHeaderRowIndex: rowIndex + 1,
        };
      }
    }

    return null;
  }

  function buildColumnMap(stageRow, subHeaderRow) {
    const map = {};
    let activeStage = null;

    for (let columnIndex = 0; columnIndex < stageRow.length; columnIndex += 1) {
      const stageText = normalize(stageRow[columnIndex]);
      const matchedStage = findStage(stageText);

      if (matchedStage) {
        activeStage = matchedStage;
      }

      if (!activeStage) {
        continue;
      }

      const subHeader = normalize(subHeaderRow[columnIndex]);

      if (!map[activeStage.key]) {
        map[activeStage.key] = {
          stage: activeStage,
          target: null,
          estimate: null,
          actual: null,
        };
      }

      if (subHeader.includes("target")) {
        map[activeStage.key].target = columnIndex;
      } else if (subHeader.includes("estimate")) {
        map[activeStage.key].estimate = columnIndex;
      } else if (subHeader.includes("actual")) {
        map[activeStage.key].actual = columnIndex;
      } else if (
        activeStage.key === "opm" &&
        map[activeStage.key].target === null
      ) {
        map[activeStage.key].target = columnIndex;
      }
    }

    return map;
  }

  function parseDataRows(data, firstDataRow, columnMap) {
    const parsed = [];
    let currentPart = "";

    for (let rowIndex = firstDataRow; rowIndex < data.length; rowIndex += 1) {
      const row = data[rowIndex] || [];
      const partValue = clean(row[0]);
      const rowType = titleCase(clean(row[1]));

      if (partValue) {
        currentPart = partValue;
      }

      if (!currentPart || (rowType !== "Original" && rowType !== "Actual")) {
        continue;
      }

      const milestones = STAGES.map((stage) => {
        const columns = columnMap[stage.key];

        if (!columns) {
          return null;
        }

        return {
          stage,
          target: parseExcelDate(
            columns.target === null ? null : row[columns.target]
          ),
          estimate: parseExcelDate(
            columns.estimate === null ? null : row[columns.estimate]
          ),
          actual: parseExcelDate(
            columns.actual === null ? null : row[columns.actual]
          ),
        };
      }).filter(Boolean);

      parsed.push({
        partNo: currentPart,
        type: rowType,
        milestones,
      });
    }

    if (!parsed.length) {
      throw new Error(
        "No Original / Actual rows were found. Check Part No. and Original/Actual columns."
      );
    }

    return parsed;
  }

  function buildPartSchedules(rawRows) {
    const grouped = new Map();

    rawRows.forEach((row) => {
      if (!grouped.has(row.partNo)) {
        grouped.set(row.partNo, {
          partNo: row.partNo,
          originalRaw: null,
          actualRaw: null,
        });
      }

      const part = grouped.get(row.partNo);

      if (row.type === "Original") {
        part.originalRaw = row;
      } else {
        part.actualRaw = row;
      }
    });

    const parts = [];

    grouped.forEach((part) => {
      const originalRaw = part.originalRaw || part.actualRaw;
      const actualRaw = part.actualRaw || part.originalRaw;

      const original = makeOriginalSchedule(originalRaw);
      const actual = makeActualForecastSchedule(actualRaw, original);

      parts.push({
        partNo: part.partNo,
        original,
        actual,
      });
    });

    return parts;
  }

  function makeOriginalSchedule(row) {
    const milestones = [];
    let previousDate = null;

    row.milestones.forEach((item) => {
      const date = cloneDate(item.target || item.estimate || item.actual);

      if (!date) {
        return;
      }

      const safeDate =
        previousDate && date < previousDate
          ? cloneDate(previousDate)
          : date;

      milestones.push({
        stage: item.stage,
        date: cloneDate(safeDate),
        target: cloneDate(item.target),
        estimate: cloneDate(item.estimate),
        actual: cloneDate(item.actual),
        source: "Target",
        delay: 0,
      });

      previousDate = cloneDate(safeDate);
    });

    return milestones;
  }

  function makeActualForecastSchedule(row, originalSchedule) {
    const originalByKey = new Map(
      originalSchedule.map((item) => [item.stage.key, item])
    );

    const milestones = [];
    let previousDate = null;

    row.milestones.forEach((item) => {
      const original = originalByKey.get(item.stage.key);
      const baselineDate = original ? original.date : item.target;

      /*
       * Required business rule:
       * 1. Use Actual when Actual exists.
       * 2. If Actual is blank, use Estimate.
       * 3. If both are blank, fall back to Target.
       * Do not push later stages automatically.
       */
      const candidate = cloneDate(
        item.actual || item.estimate || item.target
      );

      if (!candidate) {
        return;
      }

      const safeDate =
        previousDate && candidate < previousDate
          ? cloneDate(previousDate)
          : candidate;

      const source = item.actual
        ? "Actual"
        : item.estimate
          ? "Estimate"
          : "Target";

      const delay =
        baselineDate && safeDate
          ? Math.max(0, diffDays(baselineDate, safeDate))
          : 0;

      milestones.push({
        stage: item.stage,
        date: cloneDate(safeDate),
        target: cloneDate(item.target),
        estimate: cloneDate(item.estimate),
        actual: cloneDate(item.actual),
        source,
        delay,
      });

      previousDate = cloneDate(safeDate);
    });

    return milestones;
  }

  function rebuildTimeline() {
    const requestedStart = parseHtmlDate(dom.reportStart.value);
    const reportStart = getMonday(requestedStart || DEFAULT_REPORT_START);

    let maxFinish = addDays(reportStart, 7 * 52);

    state.parts.forEach((part) => {
      [...part.original, ...part.actual].forEach((item) => {
        if (item.date && item.date > maxFinish) {
          maxFinish = item.date;
        }
      });
    });

    state.timelineStart = reportStart;
    state.timelineEnd = getMonday(addDays(maxFinish, 7));
    state.weeks = [];

    for (
      let cursor = cloneDate(state.timelineStart);
      cursor <= state.timelineEnd;
      cursor = addDays(cursor, 7)
    ) {
      state.weeks.push(cursor);
    }
  }

  function render() {
    if (!state.parts.length || !state.weeks.length) {
      return;
    }

    const filteredParts = getFilteredParts();

    dom.emptyState.classList.add("hidden");
    dom.ganttViewport.classList.remove("hidden");
    dom.ganttCanvas.innerHTML = "";

    renderHeader();

    filteredParts.forEach((part) => {
      renderScheduleRow(part.partNo, "Original", part.original);
      renderScheduleRow(part.partNo, "Actual", part.actual);
    });

    renderTodayLine(filteredParts.length * 2);
    updateDashboard(filteredParts);

    dom.summaryText.textContent =
      `${filteredParts.length} part(s) · ` +
      `${filteredParts.length * 2} row(s) · ` +
      `${formatDate(state.timelineStart)} to ${formatDate(state.timelineEnd)}`;
  }

  function getFilteredParts() {
    const partValue = dom.partFilter.value;
    const stageKey = dom.stageFilter.value;
    const from = parseHtmlDate(dom.dateFrom.value);
    const to = parseHtmlDate(dom.dateTo.value);
    const delayOnly = dom.delayOnly.checked;

    return state.parts.filter((part) => {
      if (partValue !== "ALL" && part.partNo !== partValue) {
        return false;
      }

      const actualSchedule = part.actual;

      if (delayOnly && getPartDelay(part) <= 0) {
        return false;
      }

      if (stageKey !== "ALL" || from || to) {
        const relevant = actualSchedule.filter((item) => {
          if (stageKey !== "ALL" && item.stage.key !== stageKey) {
            return false;
          }

          if (from && item.date < from) {
            return false;
          }

          if (to && item.date > to) {
            return false;
          }

          return true;
        });

        if (!relevant.length) {
          return false;
        }
      }

      return true;
    });
  }

  function renderHeader() {
    const totalWidth = state.weeks.length * WEEK_WIDTH;
    const yearGroups = makeYearGroups(state.weeks);

    const yearRow = document.createElement("div");
    yearRow.className = "year-row";

    yearRow.innerHTML = `
      <div class="year-left">Part No.</div>
      <div class="year-type">Original / Actual</div>
      <div class="year-timeline" style="width:${totalWidth}px">
        ${yearGroups
          .map(
            (group) => `
              <div
                class="year-block"
                style="width:${group.count * WEEK_WIDTH}px"
              >
                ${group.year}
              </div>
            `
          )
          .join("")}
      </div>
    `;

    const weekRow = document.createElement("div");
    weekRow.className = "week-row";

    weekRow.innerHTML = `
      <div class="week-left"></div>
      <div class="week-type"></div>
      <div class="week-timeline" style="width:${totalWidth}px">
        ${state.weeks
          .map(
            (week) => `
              <div class="week-cell">
                <strong>WK ${getIsoWeek(week)}</strong>
                <span>${formatDate(week)}</span>
              </div>
            `
          )
          .join("")}
      </div>
    `;

    dom.ganttCanvas.appendChild(yearRow);
    dom.ganttCanvas.appendChild(weekRow);
  }

  function renderScheduleRow(partNo, type, schedule) {
    const selectedStage = dom.stageFilter.value;
    const row = document.createElement("div");
    row.className = "gantt-row";

    const partCell = document.createElement("div");
    partCell.className = "part-cell";
    partCell.textContent = type === "Original" ? partNo : "";

    const typeCell = document.createElement("div");
    typeCell.className =
      `type-cell ${type === "Original" ? "type-original" : "type-actual"}`;
    typeCell.textContent = type;

    const timeline = document.createElement("div");
    timeline.className = "timeline-row";
    timeline.style.width = `${state.weeks.length * WEEK_WIDTH}px`;

    schedule.forEach((item, index) => {
      if (selectedStage !== "ALL" && item.stage.key !== selectedStage) {
        return;
      }

      /*
       * A stage owns the time BEFORE its milestone:
       * previous milestone -> current milestone.
       *
       * Example:
       * POR = 07/01
       * Tooling Build = 12/28
       * The 180-day interval is colored Tooling Build, not POR.
       */
      const previousItem = schedule[index - 1];
      const segmentStart = previousItem
        ? previousItem.date
        : state.timelineStart;
      const segmentFinish = item.date;

      if (!segmentStart || !segmentFinish) {
        return;
      }

      const visibleStart =
        segmentStart < state.timelineStart
          ? state.timelineStart
          : segmentStart;

      const visibleFinish =
        segmentFinish > state.timelineEnd
          ? state.timelineEnd
          : segmentFinish;

      if (
        visibleFinish < state.timelineStart ||
        visibleStart > state.timelineEnd
      ) {
        return;
      }

      const segment = document.createElement("div");
      segment.className = "stage-segment";
      segment.style.background = item.stage.color;
      segment.style.left = `${dateOffset(visibleStart)}px`;

      const durationDays = Math.max(
        1,
        diffDays(visibleStart, visibleFinish)
      );

      segment.style.width = `${Math.max(
        DAY_WIDTH,
        durationDays * DAY_WIDTH
      )}px`;

      segment.addEventListener("mouseenter", (event) => {
        showTooltip(
          event,
          partNo,
          type,
          item,
          segmentStart,
          segmentFinish
        );
      });

      segment.addEventListener("mousemove", moveTooltip);
      segment.addEventListener("mouseleave", hideTooltip);

      timeline.appendChild(segment);
    });

    row.appendChild(partCell);
    row.appendChild(typeCell);
    row.appendChild(timeline);

    dom.ganttCanvas.appendChild(row);
  }

  function renderTodayLine(rowCount) {
    const today = stripTime(new Date());

    if (today < state.timelineStart || today > state.timelineEnd) {
      return;
    }

    const line = document.createElement("div");
    line.className = "today-line";
    line.style.left =
      `${82 + 94 + dateOffset(today)}px`;
    line.style.top = "28px";
    line.style.height = `${48 + rowCount * 30}px`;

    dom.ganttCanvas.appendChild(line);
  }

  function updateDashboard(filteredParts) {
    const selectedPart =
      dom.partFilter.value !== "ALL"
        ? filteredParts.find((part) => part.partNo === dom.partFilter.value)
        : filteredParts.length === 1
          ? filteredParts[0]
          : null;

    if (!selectedPart) {
      dom.kpiPart.textContent =
        dom.partFilter.value === "ALL" ? "All" : dom.partFilter.value;
      dom.kpiStage.textContent = filteredParts.length ? "Multiple" : "—";
      dom.kpiStageStrip.style.background = "#dce3ea";
      dom.kpiFinish.textContent = "—";
      dom.kpiDelay.textContent = "—";
      dom.kpiDelay.className = "kpi-value";
      dom.kpiProgress.textContent = "0%";
      dom.progressBar.style.width = "0%";
      return;
    }

    const today = stripTime(new Date());
    const schedule = selectedPart.actual;

    let current = schedule[schedule.length - 1] || null;

    for (let index = 0; index < schedule.length; index += 1) {
      const item = schedule[index];
      const next = schedule[index + 1];

      if (!next || today < next.date) {
        current = item;
        break;
      }
    }

    const completedActual = schedule.filter(
      (item) => item.actual && item.actual <= today
    ).length;

    const progress = Math.round(
      (completedActual / Math.max(1, schedule.length)) * 100
    );

    const delay = getPartDelay(selectedPart);
    const finish = schedule.length
      ? schedule[schedule.length - 1].date
      : null;

    dom.kpiPart.textContent = selectedPart.partNo;
    dom.kpiStage.textContent = current ? current.stage.name : "—";
    dom.kpiStageStrip.style.background = current
      ? current.stage.color
      : "#dce3ea";
    dom.kpiFinish.textContent = formatDate(finish);
    dom.kpiDelay.textContent =
      delay > 0 ? `+${delay} days` : "On time";
    dom.kpiDelay.className =
      `kpi-value ${delay > 0 ? "delay-positive" : "delay-good"}`;
    dom.kpiDelayFoot.textContent =
      delay > 0 ? "Behind Original finish" : "Compared with Original";
    dom.kpiProgress.textContent = `${progress}%`;
    dom.progressBar.style.width = `${progress}%`;
    dom.progressBar.style.background = current
      ? current.stage.color
      : "#009999";
  }

  function getPartDelay(part) {
    const originalFinish = part.original.length
      ? part.original[part.original.length - 1].date
      : null;

    const actualFinish = part.actual.length
      ? part.actual[part.actual.length - 1].date
      : null;

    if (!originalFinish || !actualFinish) {
      return 0;
    }

    return Math.max(0, diffDays(originalFinish, actualFinish));
  }

  function populatePartFilter() {
    dom.partFilter.innerHTML =
      `<option value="ALL">All Parts</option>` +
      state.parts
        .map(
          (part) =>
            `<option value="${escapeHtml(part.partNo)}">${escapeHtml(
              part.partNo
            )}</option>`
        )
        .join("");

    dom.partFilter.disabled = false;
  }

  function renderStageFilter() {
    dom.stageFilter.innerHTML =
      `<option value="ALL">All Stages</option>` +
      STAGES.map(
        (stage) =>
          `<option value="${stage.key}">${stage.name}</option>`
      ).join("");
  }

  function renderLegend() {
    dom.legend.innerHTML = STAGES.map(
      (stage) => `
        <div class="legend-item">
          <span
            class="legend-swatch"
            style="background:${stage.color}"
          ></span>
          <span>${stage.name}</span>
        </div>
      `
    ).join("");
  }

  function resetFilters() {
    dom.partFilter.value = "ALL";
    dom.stageFilter.value = "ALL";
    dom.dateFrom.value = "";
    dom.dateTo.value = "";
    dom.delayOnly.checked = false;
    render();
  }

  function showTooltip(
    event,
    partNo,
    type,
    item,
    segmentStart,
    segmentFinish
  ) {
    const delay = item.delay || 0;
    const duration = Math.max(
      1,
      diffDays(segmentStart, segmentFinish)
    );

    dom.tooltip.innerHTML = `
      <strong>${escapeHtml(partNo)} · ${escapeHtml(item.stage.name)}</strong>
      Row: ${type}<br>
      Target: ${formatDate(item.target)}<br>
      Estimate: ${formatDate(item.estimate)}<br>
      Actual: ${formatDate(item.actual)}<br>
      Milestone: ${formatDate(item.date)}<br>
      Source: ${escapeHtml(item.source)}<br>
      Delay: ${delay > 0 ? `+${delay} days` : "On time"}<br>
      Stage span: ${duration} day(s)
    `;

    dom.tooltip.classList.remove("hidden");
    moveTooltip(event);
  }

  function moveTooltip(event) {
    const gap = 14;
    const tooltipRect = dom.tooltip.getBoundingClientRect();

    let left = event.clientX + gap;
    let top = event.clientY + gap;

    if (left + tooltipRect.width > window.innerWidth - 8) {
      left = event.clientX - tooltipRect.width - gap;
    }

    if (top + tooltipRect.height > window.innerHeight - 8) {
      top = event.clientY - tooltipRect.height - gap;
    }

    dom.tooltip.style.left = `${left}px`;
    dom.tooltip.style.top = `${top}px`;
  }

  function hideTooltip() {
    dom.tooltip.classList.add("hidden");
  }

  function makeYearGroups(weeks) {
    const groups = [];

    weeks.forEach((week) => {
      const year = week.getFullYear();
      const last = groups[groups.length - 1];

      if (last && last.year === year) {
        last.count += 1;
      } else {
        groups.push({ year, count: 1 });
      }
    });

    return groups;
  }

  function findStage(value) {
    if (!value) {
      return null;
    }

    return STAGES.find((stage) =>
      stage.aliases.some(
        (alias) =>
          value === alias ||
          value.includes(alias)
      )
    ) || null;
  }

  function parseExcelDate(value) {
    if (value === null || value === undefined || value === "") {
      return null;
    }

    /*
     * Excel Date objects and Excel serial numbers already contain the full year.
     * Text values must use MM/DD/YY or MM/DD/YYYY.
     */
    if (value instanceof Date && !Number.isNaN(value.getTime())) {
      return stripTime(value);
    }

    if (typeof value === "number" && Number.isFinite(value)) {
      const parsed = XLSX.SSF.parse_date_code(value);

      if (!parsed) {
        return null;
      }

      return new Date(parsed.y, parsed.m - 1, parsed.d);
    }

    const text = clean(value);

    if (!text) {
      return null;
    }

    const match = text.match(
      /^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2}|\d{4})$/
    );

    if (!match) {
      return null;
    }

    const month = Number(match[1]);
    const day = Number(match[2]);
    let year = Number(match[3]);

    if (year < 100) {
      year += 2000;
    }

    const date = new Date(year, month - 1, day);

    if (
      Number.isNaN(date.getTime()) ||
      date.getFullYear() !== year ||
      date.getMonth() !== month - 1 ||
      date.getDate() !== day
    ) {
      return null;
    }

    return stripTime(date);
  }

  function parseHtmlDate(value) {
    if (!value) {
      return null;
    }

    const parts = value.split("-").map(Number);

    if (parts.length !== 3) {
      return null;
    }

    return new Date(parts[0], parts[1] - 1, parts[2]);
  }

  function getMonday(date) {
    const result = stripTime(date);
    const day = result.getDay();
    const offset = day === 0 ? -6 : 1 - day;

    result.setDate(result.getDate() + offset);
    return result;
  }

  function getIsoWeek(date) {
    const copy = new Date(
      Date.UTC(date.getFullYear(), date.getMonth(), date.getDate())
    );

    const day = copy.getUTCDay() || 7;
    copy.setUTCDate(copy.getUTCDate() + 4 - day);

    const yearStart = new Date(Date.UTC(copy.getUTCFullYear(), 0, 1));

    return Math.ceil(
      (((copy - yearStart) / MS_DAY) + 1) / 7
    );
  }

  function dateOffset(date) {
    return diffDays(state.timelineStart, date) * DAY_WIDTH;
  }

  function diffDays(start, end) {
    return Math.round(
      (stripTime(end) - stripTime(start)) / MS_DAY
    );
  }

  function addDays(date, days) {
    const result = cloneDate(date);
    result.setDate(result.getDate() + days);
    return result;
  }

  function stripTime(date) {
    const result = cloneDate(date);
    result.setHours(0, 0, 0, 0);
    return result;
  }

  function cloneDate(date) {
    return date ? new Date(date.getTime()) : null;
  }

  function formatDate(date) {
    if (!date) {
      return "—";
    }

    return [
      String(date.getMonth() + 1).padStart(2, "0"),
      String(date.getDate()).padStart(2, "0"),
    ].join("/");
  }

  function normalize(value) {
    return clean(value)
      .toLowerCase()
      .replace(/&/g, "and")
      .replace(/\s+/g, " ");
  }

  function clean(value) {
    return String(value ?? "").trim();
  }

  function titleCase(value) {
    const normalized = value.toLowerCase();

    if (normalized === "original") {
      return "Original";
    }

    if (normalized === "actual") {
      return "Actual";
    }

    return value;
  }

  function escapeHtml(value) {
    return clean(value)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#039;");
  }

  function get(id) {
    const element = document.getElementById(id);

    if (!element) {
      throw new Error(`Missing element #${id}`);
    }

    return element;
  }
})();
