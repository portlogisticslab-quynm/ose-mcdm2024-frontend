
"use strict";

// ============================================================
// APPLICATION STATE
// ============================================================

const state = {
  workbook: null,
  fileName: "",
  model: null,
  ranking: null,
  weightsResult: null,
  selectedCriterion: 0
};

const charts = {
  summary: null,
  detail: null,
  critic: null,
  entropy: null,
  extra: null
};

const normalizationNames = [
  "Max",
  "Sum",
  "Vector",
  "Max-Min",
  "DEA"
];

const methodNames = [
  "SAW",
  "CODAS",
  "COPRAS",
  "TOPSIS-SC",
  "TOPSIS-Eu",
  "TOPSIS-Inf",
  "GRAt",
  "PROMETHEE"
];

// ============================================================
// DOM
// ============================================================

const fileInput =
  document.getElementById("fileInput");

const loadButton =
  document.getElementById("loadButton");

const fileNameBox =
  document.getElementById("fileName");

const sheetSelect =
  document.getElementById("sheetSelect");

const methodSelect =
  document.getElementById("methodSelect");

const demoButton =
  document.getElementById("demoButton");

const templateButton =
  document.getElementById("templateButton");

const calculateButton =
  document.getElementById("calculateButton");

const weightsButton =
  document.getElementById("weightsButton");

const exportButton =
  document.getElementById("exportButton");

const resetButton =
  document.getElementById("resetButton");

const extraOutput =
  document.getElementById("extraOutput");

const criteriaFilter =
  document.getElementById("criteriaFilter");

const criteriaList =
  document.getElementById("criteriaList");

const statusBox =
  document.getElementById("statusBox");

const matrixTable =
  document.getElementById("matrixTable");

const criteriaTableBody =
  document.querySelector("#criteriaTable tbody");

const rankingTableBody =
  document.querySelector("#rankingTable tbody");

const weightsTable =
  document.getElementById("weightsTable");

// ============================================================
// TABS
// ============================================================

document
  .querySelectorAll(".tab-button")
  .forEach(button => {
    button.addEventListener("click", () => {
      openTab(button.dataset.tab);
    });
  });

function openTab(tabId) {
  document
    .querySelectorAll(".tab-button")
    .forEach(button => {
      button.classList.toggle(
        "active",
        button.dataset.tab === tabId
      );
    });

  document
    .querySelectorAll(".tab-content")
    .forEach(content => {
      content.classList.toggle(
        "active",
        content.id === tabId
      );
    });
}

// ============================================================
// FILE LOADING
// ============================================================

loadButton.addEventListener("click", () => {
  fileInput.value = "";
  fileInput.click();
});

fileInput.addEventListener("change", async event => {
  const file = event.target.files[0];

  if (!file) {
    return;
  }

  try {
    setStatus(`Reading ${file.name}...`);

    const buffer = await file.arrayBuffer();

    state.workbook = XLSX.read(buffer, {
      type: "array",
      cellDates: true
    });

    state.fileName = file.name;
    fileNameBox.textContent = file.name;

    populateSheetSelector();
    loadSelectedSheet();

    setStatus(
      `Loaded workbook: ${file.name}\nSelect a worksheet if needed, review the input tables, and calculate ranking.`
    );
  } catch (error) {
    setStatus(`Load error: ${error.message}`);
  }
});

sheetSelect.addEventListener("change", loadSelectedSheet);

function populateSheetSelector() {
  sheetSelect.innerHTML = "";

  for (const sheetName of state.workbook.SheetNames) {
    const option = document.createElement("option");
    option.value = sheetName;
    option.textContent = sheetName;
    sheetSelect.appendChild(option);
  }

  sheetSelect.disabled = false;

  const firstValid =
    state.workbook.SheetNames.find(name => {
      try {
        parseWorksheet(
          state.workbook.Sheets[name]
        );
        return true;
      } catch {
        return false;
      }
    });

  if (firstValid) {
    sheetSelect.value = firstValid;
  }
}

function loadSelectedSheet() {
  if (!state.workbook || !sheetSelect.value) {
    return;
  }

  try {
    const worksheet =
      state.workbook.Sheets[sheetSelect.value];

    const model = parseWorksheet(worksheet);

    state.model = model;
    state.ranking = null;
    state.weightsResult = null;
    state.selectedCriterion = 0;

    renderInputTables();
    renderCriteriaList();
    clearResults();

    enableModelControls();

    setStatus(
      `Worksheet: ${sheetSelect.value}\nAlternatives: ${model.alternatives.length}\nCriteria: ${model.criteria.length}\nReview or edit the data, then click Calculate Ranking.`
    );
  } catch (error) {
    state.model = null;
    clearInputTables();
    disableModelControls();

    setStatus(
      `Worksheet "${sheetSelect.value}" is not in the expected MCDM matrix format.\n${error.message}`
    );
  }
}

function parseWorksheet(worksheet) {
  const rows = XLSX.utils.sheet_to_json(
    worksheet,
    {
      header: 1,
      defval: null,
      raw: true
    }
  );

  if (rows.length < 2) {
    throw new Error("The sheet has insufficient data.");
  }

  const headerRowIndex = rows.findIndex(row =>
    row.some(
      value =>
        String(value ?? "")
          .trim()
          .toLowerCase() === "criteria"
    )
  );

  if (headerRowIndex < 0) {
    throw new Error(
      'A header cell named "Criteria" was not found.'
    );
  }

  const header = rows[headerRowIndex].map(
    value => String(value ?? "").trim()
  );

  const lower = header.map(
    value => value.toLowerCase()
  );

  const criteriaIndex =
    lower.findIndex(
      value => value === "criteria"
    );

  let notationIndex =
    lower.findIndex(
      value => value === "notation"
    );

  if (notationIndex < 0) {
    notationIndex = criteriaIndex + 1;
  }

  const weightIndex =
    lower.findIndex(
      value =>
        value === "weight" ||
        value === "weight factor"
    );

  const attributionIndex =
    lower.findIndex(
      value =>
        value === "attribution" ||
        value === "attribute" ||
        value.includes("positive") ||
        value.includes("negative")
    );

  const availableIndex =
    lower.findIndex(
      value => value === "available"
    );

  if (weightIndex < 0) {
    throw new Error(
      'A "Weight" column was not found.'
    );
  }

  const firstAlternativeColumn =
    notationIndex + 1;

  const lastAlternativeColumn =
    weightIndex - 1;

  if (
    lastAlternativeColumn <
    firstAlternativeColumn
  ) {
    throw new Error(
      "No alternative columns were found between Notation and Weight."
    );
  }

  const alternativeColumns = [];

  for (
    let column = firstAlternativeColumn;
    column <= lastAlternativeColumn;
    column++
  ) {
    const name = header[column];

    if (name) {
      alternativeColumns.push({
        index: column,
        name
      });
    }
  }

  if (alternativeColumns.length < 2) {
    throw new Error(
      "At least two alternatives are required."
    );
  }

  const criterionRows = rows
    .slice(headerRowIndex + 1)
    .filter(row => {
      const criterion =
        String(row[criteriaIndex] ?? "").trim();

      if (!criterion) {
        return false;
      }

      return alternativeColumns.some(
        item =>
          Number.isFinite(
            Number(row[item.index])
          )
      );
    });

  if (criterionRows.length < 2) {
    throw new Error(
      "At least two valid criteria rows are required."
    );
  }

  const criteria = criterionRows.map(
    (row, index) => {
      const rawWeight =
        Number(row[weightIndex]);

      const attribution =
        attributionIndex >= 0
          ? String(
              row[attributionIndex] ??
              "Positive"
            ).trim()
          : "Positive";

      const available =
        availableIndex >= 0
          ? parseBoolean(
              row[availableIndex],
              true
            )
          : true;

      return {
        name:
          String(
            row[criteriaIndex] ??
            `C${index + 1}`
          ).trim(),
        notation:
          String(
            row[notationIndex] ??
            `C${index + 1}`
          ).trim(),
        weight:
          Number.isFinite(rawWeight)
            ? rawWeight
            : 1,
        sign:
          isPositiveText(attribution)
            ? 1
            : -1,
        available
      };
    }
  );

  const alternatives =
    alternativeColumns.map(
      item => item.name
    );

  const matrix =
    alternatives.map(
      (_, alternativeIndex) =>
        criterionRows.map(row => {
          const value = Number(
            row[
              alternativeColumns[
                alternativeIndex
              ].index
            ]
          );

          return Number.isFinite(value)
            ? value
            : 0;
        })
    );

  return {
    alternatives,
    criteria,
    matrix,
    sourceSheet:
      sheetSelect.value || "Sheet"
  };
}

function parseBoolean(value, fallback) {
  if (
    value === true ||
    value === 1
  ) {
    return true;
  }

  if (
    value === false ||
    value === 0
  ) {
    return false;
  }

  const text =
    String(value ?? "")
      .trim()
      .toLowerCase();

  if (
    ["true", "yes", "y", "1", "available"].includes(text)
  ) {
    return true;
  }

  if (
    ["false", "no", "n", "0", "unavailable"].includes(text)
  ) {
    return false;
  }

  return fallback;
}

function isPositiveText(value) {
  const text =
    String(value ?? "")
      .trim()
      .toLowerCase();

  return !(
    text.startsWith("neg") ||
    text.startsWith("cost") ||
    text === "-1"
  );
}

// ============================================================
// BUILT-IN DEMO
// ============================================================

demoButton.addEventListener("click", () => {
  state.workbook = null;
  state.fileName = "Built-in car selection demo";
  fileNameBox.textContent = state.fileName;
  sheetSelect.innerHTML =
    `<option value="Demo">Demo</option>`;
  sheetSelect.disabled = true;

  state.model = {
    sourceSheet: "Demo",
    alternatives: [
      "Nissani",
      "Mazdan",
      "Hondan",
      "Toyotan",
      "Suzukin"
    ],
    criteria: [
      {
        name: "Price",
        notation: "Cost",
        weight: 0.2030845,
        sign: -1,
        available: true
      },
      {
        name: "Driving life",
        notation: "Life",
        weight: 0.2747733,
        sign: 1,
        available: true
      },
      {
        name: "Safety",
        notation: "Safe",
        weight: 0.2286536,
        sign: 1,
        available: true
      },
      {
        name: "Looks",
        notation: "Look",
        weight: 0.2934886,
        sign: 1,
        available: true
      }
    ],
    matrix: [
      [260, 16, 16, 4],
      [275, 16, 15, 3],
      [300, 18, 18, 5],
      [275, 20, 18, 4],
      [250, 12, 12, 4]
    ]
  };

  state.ranking = null;
  state.weightsResult = null;
  state.selectedCriterion = 0;

  renderInputTables();
  renderCriteriaList();
  clearResults();
  enableModelControls();

  setStatus(
    "Built-in car-selection demo loaded.\nYou can edit the matrix and criterion settings before calculation."
  );
});

// ============================================================
// INPUT TABLES
// ============================================================

function renderInputTables() {
  renderMatrixTable();
  renderCriteriaTable();
}

function renderMatrixTable() {
  const head =
    matrixTable.querySelector("thead");

  const body =
    matrixTable.querySelector("tbody");

  head.innerHTML = `
    <tr>
      <th>Alternative</th>
      ${state.model.criteria.map(
        criterion =>
          `<th title="${escapeHtml(criterion.name)}">${escapeHtml(criterion.notation || criterion.name)}</th>`
      ).join("")}
    </tr>
  `;

  body.innerHTML = "";

  state.model.alternatives.forEach(
    (alternative, rowIndex) => {
      const row =
        document.createElement("tr");

      row.innerHTML = `
        <td>${escapeHtml(alternative)}</td>
        ${state.model.criteria.map(
          (_, columnIndex) => `
            <td>
              <input
                type="number"
                step="any"
                data-matrix-row="${rowIndex}"
                data-matrix-column="${columnIndex}"
                value="${state.model.matrix[rowIndex][columnIndex]}"
              >
            </td>
          `
        ).join("")}
      `;

      body.appendChild(row);
    }
  );

  body
    .querySelectorAll("input")
    .forEach(input => {
      input.addEventListener(
        "change",
        event => {
          const row =
            Number(
              event.target.dataset.matrixRow
            );

          const column =
            Number(
              event.target.dataset.matrixColumn
            );

          const value =
            Number(event.target.value);

          state.model.matrix[row][column] =
            Number.isFinite(value)
              ? value
              : 0;

          invalidateResults();
        }
      );
    });
}

function renderCriteriaTable() {
  criteriaTableBody.innerHTML = "";

  state.model.criteria.forEach(
    (criterion, index) => {
      const row =
        document.createElement("tr");

      row.innerHTML = `
        <td title="${escapeHtml(criterion.name)}">
          ${escapeHtml(criterion.notation || criterion.name)}
        </td>

        <td>
          <input
            type="number"
            step="any"
            min="0"
            data-setting="weight"
            data-index="${index}"
            value="${criterion.weight}"
          >
        </td>

        <td>
          <select
            data-setting="sign"
            data-index="${index}"
          >
            <option value="1" ${criterion.sign === 1 ? "selected" : ""}>
              Positive
            </option>

            <option value="-1" ${criterion.sign === -1 ? "selected" : ""}>
              Negative
            </option>
          </select>
        </td>

        <td>
          <input
            type="checkbox"
            data-setting="available"
            data-index="${index}"
            ${criterion.available ? "checked" : ""}
          >
        </td>
      `;

      criteriaTableBody.appendChild(row);
    }
  );

  criteriaTableBody
    .querySelectorAll(
      "input, select"
    )
    .forEach(control => {
      control.addEventListener(
        "change",
        event => {
          const index =
            Number(
              event.target.dataset.index
            );

          const setting =
            event.target.dataset.setting;

          if (setting === "weight") {
            const value =
              Number(event.target.value);

            state.model.criteria[index].weight =
              Number.isFinite(value) &&
              value >= 0
                ? value
                : 0;
          }

          if (setting === "sign") {
            state.model.criteria[index].sign =
              Number(event.target.value);
          }

          if (setting === "available") {
            state.model.criteria[index].available =
              event.target.checked;
          }

          invalidateResults();
        }
      );
    });
}

function clearInputTables() {
  matrixTable
    .querySelector("thead")
    .innerHTML = "";

  matrixTable
    .querySelector("tbody")
    .innerHTML = "";

  criteriaTableBody.innerHTML = "";
  criteriaList.innerHTML = "";
}

// ============================================================
// CRITERIA FILTER / HIGHLIGHT
// ============================================================

criteriaFilter.addEventListener("input", renderCriteriaList);

function renderCriteriaList() {
  if (!state.model) {
    criteriaList.innerHTML = "";
    return;
  }

  const query =
    criteriaFilter.value
      .trim()
      .toLowerCase();

  criteriaList.innerHTML = "";

  state.model.criteria.forEach(
    (criterion, index) => {
      const label =
        `${criterion.notation} – ${criterion.name}`;

      if (
        query &&
        !label.toLowerCase().includes(query)
      ) {
        return;
      }

      const button =
        document.createElement("button");

      button.type = "button";
      button.className =
        `criterion-item ${
          state.selectedCriterion === index
            ? "active"
            : ""
        }`;

      button.textContent = label;

      button.addEventListener("click", () => {
        state.selectedCriterion = index;
        renderCriteriaList();

        if (state.weightsResult) {
          updateWeightCharts();
        }
      });

      criteriaList.appendChild(button);
    }
  );
}

// ============================================================
// RANKING CALCULATION
// ============================================================

calculateButton.addEventListener("click", async () => {
  try {
    if (!state.model) throw new Error("Load a matrix first.");
    const prepared = prepareActiveModel();
    const method = methodSelect.value;
    setStatus(`Calculating ${method} in your browser...`);
    await nextFrame();
    state.ranking=runEngine("calculateRanking",prepared,method); renderRanking(); exportButton.disabled=false; extraOutput.disabled=false; resetButton.disabled=false; openTab("rankingTab"); setStatus(buildRankingSummary(state.ranking));
  } catch(error) { setStatus(`Calculation error: ${error.message}`); }
});

function prepareActiveModel() {
  const activeIndices = [];

  state.model.criteria.forEach(
    (criterion, index) => {
      if (
        criterion.available &&
        Number(criterion.weight) > 0
      ) {
        activeIndices.push(index);
      }
    }
  );

  if (activeIndices.length === 0) {
    throw new Error(
      "No available criterion with positive weight remains."
    );
  }

  if (
    state.model.alternatives.length < 2
  ) {
    throw new Error(
      "At least two alternatives are required."
    );
  }

  const D =
    state.model.matrix.map(
      row =>
        activeIndices.map(
          index =>
            finiteOrZero(row[index])
        )
    );

  const rawWeights =
    activeIndices.map(
      index =>
        Math.max(
          0,
          finiteOrZero(
            state.model.criteria[index].weight
          )
        )
    );

  const weightTotal =
    sum(rawWeights);

  if (weightTotal <= 0) {
    throw new Error(
      "The active criterion weights sum to zero."
    );
  }

  const weights =
    rawWeights.map(
      value => value / weightTotal
    );

  const signs =
    activeIndices.map(
      index =>
        state.model.criteria[index].sign === -1
          ? -1
          : 1
    );

  const criteria =
    activeIndices.map(
      index =>
        state.model.criteria[index]
    );

  return {
    D,
    weights,
    signs,
    criteria,
    alternatives:
      [...state.model.alternatives]
  };
}

function calculateSingleMethod(
  prepared,
  method
) {
  const scoreSets =
    computeMethodScoreSets(
      prepared,
      method
    );

  const aggregate =
    aggregateScoreSets(
      scoreSets,
      prepared.alternatives
    );

  return {
    type: "single",
    method,
    prepared,
    scoreSets,
    ...aggregate
  };
}

function calculateAllMethods(prepared) {
  const methodResults =
    methodNames.map(method => {
      const scoreSets =
        computeMethodScoreSets(
          prepared,
          method
        );

      const aggregate =
        aggregateScoreSets(
          scoreSets,
          prepared.alternatives
        );

      return {
        method,
        scoreSets,
        ...aggregate
      };
    });

  const averageRanks =
    prepared.alternatives.map(
      (_, alternativeIndex) =>
        mean(
          methodResults.map(
            result =>
              result.averageRank[
                alternativeIndex
              ]
          )
        )
    );

  const order =
    sortedIndices(
      averageRanks,
      true
    );

  const finalRank =
    rankVectorFromOrder(order);

  return {
    type: "all",
    method: "All_methods",
    prepared,
    methodResults,
    averageRank: averageRanks,
    rankSum:
      averageRanks.map(
        value =>
          value *
          methodResults.length
      ),
    order,
    finalRank,
    rankCounts:
      Array.from(
        {
          length:
            prepared.alternatives.length
        },
        () =>
          new Array(
            prepared.alternatives.length
          ).fill(0)
      ),
    rankPercentages:
      prepared.alternatives.map(
        () => [0, 0, 0]
      )
  };
}

function computeMethodScoreSets(
  prepared,
  method
) {
  const m = prepared.D.length;
  const n = prepared.D[0].length;

  const normalizations = [1, 2, 3, 4, 5]
    .map(parameter => {
      const V =
        normalizeDecisionMatrix(
          prepared.D,
          parameter
        );

      const inverse =
        applyCostTransformation(
          V,
          prepared.D,
          prepared.signs,
          parameter
        );

      return {
        parameter,
        name:
          normalizationNames[
            parameter - 1
          ],
        V,
        inverse
      };
    });

  if (method === "PROMETHEE") {
    const profiles = [
      {
        name: "V-shape",
        types:
          new Array(n).fill(3)
      },
      {
        name: "Linear",
        types:
          new Array(n).fill(5)
      },
      {
        name: "Gaussian",
        types:
          new Array(n).fill(6)
      },
      {
        name: "Mixed",
        types:
          Array.from(
            { length: n },
            (_, index) =>
              index % 2 === 1
                ? 6
                : 5
          )
      }
    ];

    return profiles.map(profile => ({
      label: profile.name,
      scores:
        prometheeScores(
          prepared.D,
          prepared.weights,
          prepared.signs,
          profile.types
        )
    }));
  }

  const scoreSets =
    normalizations.map(item => ({
      label: item.name,
      scores:
        scoresForMethod(
          method,
          prepared,
          item
        )
    }));

  if (method === "GRAt") {
    scoreSets.push({
      label: "Raw matrix",
      scores:
        graScores(
          prepared.D,
          prepared.weights,
          prepared.signs,
          0.5
        )
    });
  }

  return scoreSets;
}

function scoresForMethod(
  method,
  prepared,
  normalization
) {
  const D = prepared.D;
  const weights = prepared.weights;
  const signs = prepared.signs;
  const V = normalization.V;
  const inverse = normalization.inverse;

  if (method === "SAW") {
    return matrixVectorProduct(
      inverse,
      weights
    );
  }

  if (method === "CODAS") {
    return codasScores(
      inverse,
      weights,
      0.02
    );
  }

  if (method === "COPRAS") {
    return coprasScores(
      V,
      weights,
      signs
    );
  }

  if (method === "TOPSIS-SC") {
    return topsisScores(
      V,
      weights,
      signs,
      1
    );
  }

  if (method === "TOPSIS-Eu") {
    return topsisScores(
      V,
      weights,
      signs,
      2
    );
  }

  if (method === "TOPSIS-Inf") {
    return topsisScores(
      V,
      weights,
      signs,
      Infinity
    );
  }

  if (method === "GRAt") {
    return graScores(
      V,
      weights,
      signs,
      0.5
    );
  }

  throw new Error(
    `Unsupported method: ${method}`
  );
}

function aggregateScoreSets(
  scoreSets,
  alternatives
) {
  const m = alternatives.length;
  const numberOfSets =
    scoreSets.length;

  const orders =
    scoreSets.map(
      item =>
        sortedIndices(
          item.scores,
          false
        )
    );

  const rankVectors =
    orders.map(rankVectorFromOrder);

  const rankCounts =
    Array.from(
      { length: m },
      () =>
        new Array(m).fill(0)
    );

  rankVectors.forEach(ranks => {
    ranks.forEach(
      (rank, alternativeIndex) => {
        rankCounts[
          alternativeIndex
        ][rank - 1]++;
      }
    );
  });

  const rankSum =
    rankCounts.map(
      row =>
        row.reduce(
          (total, count, rankIndex) =>
            total +
            count *
            (rankIndex + 1),
          0
        )
    );

  const averageRank =
    rankSum.map(
      value =>
        value / numberOfSets
    );

  const order =
    sortedIndices(
      averageRank,
      true
    );

  const finalRank =
    rankVectorFromOrder(order);

  const rankPercentages =
    rankCounts.map(row => [
      100 *
        (row[0] || 0) /
        numberOfSets,
      100 *
        (row[1] || 0) /
        numberOfSets,
      100 *
        (row[2] || 0) /
        numberOfSets
    ]);

  return {
    orders,
    rankVectors,
    rankCounts,
    rankSum,
    averageRank,
    order,
    finalRank,
    rankPercentages
  };
}

// ============================================================
// NORMALIZATION
// ============================================================

function normalizeDecisionMatrix(
  D,
  parameter
) {
  const m = D.length;
  const n = D[0].length;

  const mins =
    columnMinimums(D);

  const maxs =
    columnMaximums(D);

  const sums =
    columnSums(D);

  const vectors =
    Array.from(
      { length: n },
      (_, column) =>
        Math.sqrt(
          D.reduce(
            (total, row) =>
              total +
              row[column] ** 2,
            0
          )
        )
    );

  return D.map(row =>
    row.map((value, column) => {
      if (parameter === 1) {
        return safeDivide(
          value,
          maxs[column]
        );
      }

      if (parameter === 2) {
        return safeDivide(
          value,
          sums[column]
        );
      }

      if (parameter === 3) {
        return safeDivide(
          value,
          vectors[column]
        );
      }

      if (parameter === 4) {
        return safeDivide(
          value - mins[column],
          maxs[column] -
          mins[column]
        );
      }

      if (parameter === 5) {
        const denominator =
          m * maxs[column] -
          sums[column];

        if (
          Math.abs(denominator) <
          1e-15
        ) {
          return 1;
        }

        return (
          1 -
          (
            maxs[column] -
            value
          ) /
          denominator
        );
      }

      return value;
    })
  );
}

function applyCostTransformation(
  V,
  D,
  signs,
  parameter
) {
  const result =
    V.map(row => [...row]);

  const m = D.length;
  const n = D[0].length;

  const mins =
    columnMinimums(D);

  const maxs =
    columnMaximums(D);

  for (let column = 0; column < n; column++) {
    if (signs[column] !== -1) {
      continue;
    }

    if (parameter === 1) {
      for (let row = 0; row < m; row++) {
        result[row][column] =
          safeDivide(
            mins[column],
            D[row][column]
          );
      }
    }

    if (parameter === 2) {
      const inverses =
        D.map(row =>
          safeReciprocal(
            row[column]
          )
        );

      const total =
        sum(inverses);

      for (let row = 0; row < m; row++) {
        result[row][column] =
          safeDivide(
            inverses[row],
            total
          );
      }
    }

    if (parameter === 3) {
      const inverses =
        D.map(row =>
          safeReciprocal(
            row[column]
          )
        );

      const vector =
        Math.sqrt(
          sum(
            inverses.map(
              value => value ** 2
            )
          )
        );

      for (let row = 0; row < m; row++) {
        result[row][column] =
          safeDivide(
            inverses[row],
            vector
          );
      }
    }

    if (parameter === 4) {
      const denominator =
        safeReciprocal(maxs[column]) -
        safeReciprocal(mins[column]);

      for (let row = 0; row < m; row++) {
        result[row][column] =
          safeDivide(
            safeReciprocal(maxs[column]) -
            safeReciprocal(D[row][column]),
            denominator
          );
      }
    }

    if (parameter === 5) {
      const minReciprocal =
        safeReciprocal(
          mins[column]
        );

      const maxReciprocal =
        safeReciprocal(
          maxs[column]
        );

      for (let row = 0; row < m; row++) {
        result[row][column] =
          1 -
          safeDivide(
            safeReciprocal(D[row][column]) -
            minReciprocal,
            maxReciprocal
          );
      }
    }
  }

  return result;
}

// ============================================================
// MCDM METHODS
// ============================================================

function codasScores(
  V,
  weights,
  threshold
) {
  const weighted =
    applyColumnWeights(
      V,
      weights
    );

  const negativeIdeal =
    columnMinimums(weighted);

  const taxi = [];
  const euclidean = [];

  weighted.forEach(row => {
    const differences =
      row.map(
        (value, column) =>
          Math.abs(
            negativeIdeal[column] -
            value
          )
      );

    taxi.push(sum(differences));

    euclidean.push(
      Math.sqrt(
        sum(
          differences.map(
            value => value ** 2
          )
        )
      )
    );
  });

  return weighted.map(
    (_, first) => {
      let score = 0;

      weighted.forEach(
        (__, second) => {
          const switchValue =
            Math.abs(
              euclidean[first] -
              euclidean[second]
            ) >= threshold
              ? 1
              : 0;

          score +=
            (
              euclidean[first] -
              euclidean[second]
            ) +
            switchValue *
            (
              taxi[first] -
              taxi[second]
            );
        }
      );

      return score;
    }
  );
}

function coprasScores(
  V,
  weights,
  signs
) {
  const weighted =
    applyColumnWeights(
      V,
      weights
    );

  const positive =
    weighted.map(row =>
      row.reduce(
        (total, value, column) =>
          total +
          (
            signs[column] === 1
              ? value
              : 0
          ),
        0
      )
    );

  const negative =
    weighted.map(row =>
      row.reduce(
        (total, value, column) =>
          total +
          (
            signs[column] === -1
              ? value
              : 0
          ),
        0
      )
    );

  const negativeSum =
    sum(negative);

  const reciprocalSum =
    sum(
      negative
        .filter(
          value =>
            Math.abs(value) >
            1e-15
        )
        .map(
          value => 1 / value
        )
    );

  if (
    Math.abs(reciprocalSum) <
    1e-15
  ) {
    return positive;
  }

  return positive.map(
    (value, index) =>
      Math.abs(negative[index]) >
      1e-15
        ? value +
          negativeSum /
          (
            negative[index] *
            reciprocalSum
          )
        : value
  );
}

function topsisScores(
  V,
  weights,
  signs,
  metric
) {
  const weighted =
    applyColumnWeights(
      V,
      weights
    );

  const maximums =
    columnMaximums(weighted);

  const minimums =
    columnMinimums(weighted);

  const positiveIdeal =
    signs.map(
      (sign, column) =>
        sign === 1
          ? maximums[column]
          : minimums[column]
    );

  const negativeIdeal =
    signs.map(
      (sign, column) =>
        sign === 1
          ? minimums[column]
          : maximums[column]
    );

  return weighted.map(row => {
    const distancePositive =
      lpDistance(
        row,
        positiveIdeal,
        metric
      );

    const distanceNegative =
      lpDistance(
        row,
        negativeIdeal,
        metric
      );

    return safeDivide(
      distanceNegative,
      distancePositive +
      distanceNegative
    );
  });
}

function graScores(
  V,
  weights,
  signs,
  coefficient
) {
  const m = V.length;
  const n = V[0].length;

  const maximums =
    columnMaximums(V);

  const minimums =
    columnMinimums(V);

  const ideal =
    signs.map(
      (sign, column) =>
        sign === 1
          ? maximums[column]
          : minimums[column]
    );

  const antiIdeal =
    signs.map(
      (sign, column) =>
        sign === 1
          ? minimums[column]
          : maximums[column]
    );

  const differenceIdeal =
    V.map(row =>
      row.map(
        (value, column) =>
          Math.abs(
            ideal[column] -
            value
          )
      )
    );

  const differenceAnti =
    V.map(row =>
      row.map(
        (value, column) =>
          Math.abs(
            antiIdeal[column] -
            value
          )
      )
    );

  const rowMaxIdeal =
    differenceIdeal.map(
      row => Math.max(...row)
    );

  const rowMinIdeal =
    differenceIdeal.map(
      row => Math.min(...row)
    );

  const rowMaxAnti =
    differenceAnti.map(
      row => Math.max(...row)
    );

  const rowMinAnti =
    differenceAnti.map(
      row => Math.min(...row)
    );

  const globalMinIdeal =
    Math.min(...rowMinIdeal);

  const globalMaxIdeal =
    Math.max(...rowMaxIdeal);

  const globalMinAnti =
    Math.min(...rowMinAnti);

  const globalMaxAnti =
    Math.max(...rowMaxAnti);

  const relationIdeal =
    differenceIdeal.map(row =>
      row.map(value =>
        safeDivide(
          globalMinIdeal +
          coefficient *
          globalMaxIdeal,
          value +
          coefficient *
          globalMaxIdeal
        )
      )
    );

  const relationAnti =
    differenceAnti.map(row =>
      row.map(value =>
        safeDivide(
          globalMinAnti +
          coefficient *
          globalMaxAnti,
          value +
          coefficient *
          globalMaxAnti
        )
      )
    );

  const gradeIdeal =
    matrixVectorProduct(
      relationIdeal,
      weights
    );

  const gradeAnti =
    matrixVectorProduct(
      relationAnti,
      weights
    );

  return gradeIdeal.map(
    (value, index) =>
      safeDivide(
        value,
        gradeAnti[index]
      )
  );
}

function prometheeScores(
  D,
  weights,
  signs,
  preferenceTypes
) {
  const m = D.length;
  const n = D[0].length;

  const ranges =
    columnMaximums(D).map(
      (maximum, column) =>
        maximum -
        columnMinimums(D)[column]
    );

  const p = Array.from(
    { length: 7 },
    () =>
      new Array(n).fill(0)
  );

  const q = Array.from(
    { length: 7 },
    () =>
      new Array(n).fill(0)
  );

  for (let column = 0; column < n; column++) {
    const base =
      Math.ceil(
        ranges[column] / m
      );

    p[2][column] = base;
    p[3][column] = base;
    p[4][column] = base;
    p[5][column] = base;

    q[4][column] =
      Math.ceil(
        1.41 *
        ranges[column] /
        m
      );

    q[5][column] =
      q[4][column];

    p[6][column] =
      (
        p[5][column] +
        q[5][column]
      ) / 2;
  }

  const preference =
    Array.from(
      { length: m },
      () =>
        new Array(m).fill(0)
    );

  for (let criterion = 0; criterion < n; criterion++) {
    for (let first = 0; first < m; first++) {
      for (let second = 0; second < m; second++) {
        const difference =
          D[first][criterion] -
          D[second][criterion];

        const type =
          preferenceTypes[criterion];

        preference[first][second] +=
          weights[criterion] *
          preferenceFunction(
            difference,
            type,
            p[type][criterion],
            q[type][criterion],
            signs[criterion]
          );
      }
    }
  }

  const positiveFlow =
    preference.map(
      (row, index) =>
        row.reduce(
          (total, value, column) =>
            total +
            (
              column === index
                ? 0
                : value
            ),
          0
        )
    );

  const negativeFlow =
    Array.from(
      { length: m },
      (_, column) =>
        preference.reduce(
          (total, row, rowIndex) =>
            total +
            (
              rowIndex === column
                ? 0
                : row[column]
            ),
          0
        )
    );

  return positiveFlow.map(
    (value, index) =>
      value -
      negativeFlow[index]
  );
}

function preferenceFunction(
  difference,
  type,
  p,
  q,
  sign
) {
  const directed =
    difference * sign;

  if (directed <= 0) {
    return 0;
  }

  const absolute =
    Math.abs(difference);

  if (type === 1) {
    return 1;
  }

  if (type === 2) {
    return absolute <= p
      ? 0
      : 1;
  }

  if (type === 3) {
    if (p <= 0) {
      return 1;
    }

    return absolute < p
      ? absolute / p
      : 1;
  }

  if (type === 4) {
    if (absolute <= p) {
      return 0;
    }

    if (absolute < q) {
      return 0.5;
    }

    return 1;
  }

  if (type === 5) {
    if (absolute <= p) {
      return 0;
    }

    if (q <= p) {
      return 1;
    }

    if (absolute <= q) {
      return (
        absolute - p
      ) /
      (
        q - p
      );
    }

    return 1;
  }

  if (type === 6) {
    if (p <= 0) {
      return 1;
    }

    return (
      1 -
      Math.exp(
        -(difference ** 2) /
        (
          2 *
          p ** 2
        )
      )
    );
  }

  return 0;
}

// ============================================================
// RENDER RANKING
// ============================================================

function renderRanking() {
  renderRankingTable();
  updateRankingCharts();
}

function renderRankingTable() {
  const ranking = state.ranking;
  const alternatives =
    ranking.prepared.alternatives;

  rankingTableBody.innerHTML = "";

  ranking.order.forEach(
    alternativeIndex => {
      const row =
        document.createElement("tr");

      const percentages =
        ranking.rankPercentages[
          alternativeIndex
        ] || [0, 0, 0];

      row.innerHTML = `
        <td>${escapeHtml(alternatives[alternativeIndex])}</td>
        <td>${ranking.finalRank[alternativeIndex]}</td>
        <td>${formatNumber(ranking.averageRank[alternativeIndex])}</td>
        <td>${formatNumber(ranking.rankSum[alternativeIndex])}</td>
        <td>${formatNumber(percentages[0])}</td>
        <td>${formatNumber(percentages[1])}</td>
        <td>${formatNumber(percentages[2])}</td>
      `;

      rankingTableBody.appendChild(row);
    }
  );
}

function updateRankingCharts() {
  destroyChart("summary");
  destroyChart("detail");

  const ranking = state.ranking;
  const alternatives =
    ranking.prepared.alternatives;

  if (ranking.type === "all") {
    charts.summary = new Chart(
      document.getElementById("summaryChart"),
      {
        type: "bar",
        data: {
          labels: alternatives,
          datasets: [{
            label:
              "Average rank across methods",
            data:
              ranking.averageRank
          }]
        },
        options:
          categoryChartOptions(
            "Alternative",
            "Average rank (lower is better)"
          )
      }
    );
  } else {
    charts.summary = new Chart(
      document.getElementById("summaryChart"),
      {
        type: "bar",
        data: {
          labels: alternatives,
          datasets: [
            {
              label: "Rank I",
              data:
                ranking.rankPercentages.map(
                  row => row[0]
                )
            },
            {
              label: "Rank II",
              data:
                ranking.rankPercentages.map(
                  row => row[1]
                )
            },
            {
              label: "Rank III",
              data:
                ranking.rankPercentages.map(
                  row => row[2]
                )
            }
          ]
        },
        options:
          categoryChartOptions(
            "Alternative",
            "Rank score (%)"
          )
      }
    );
  }

  const orderedLabels =
    ranking.order.map(
      index => alternatives[index]
    );

  const orderedValues =
    ranking.order.map(
      index =>
        ranking.averageRank[index]
    );

  charts.detail = new Chart(
    document.getElementById("detailChart"),
    {
      type: "bar",
      data: {
        labels: orderedLabels,
        datasets: [{
          label:
            "Average rank",
          data: orderedValues
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        animation: false,
        indexAxis: "y",
        scales: {
          x: {
            beginAtZero: true,
            title: {
              display: true,
              text:
                "Average rank (lower is better)"
            }
          },
          y: {
            title: {
              display: true,
              text: "Alternative"
            }
          }
        }
      }
    }
  );
}

function buildRankingSummary(ranking) {
  const lines = [
    `Method: ${ranking.method}`,
    `Source: ${state.fileName || "Built-in demo"} / ${state.model.sourceSheet}`,
    `Alternatives: ${ranking.prepared.alternatives.length}`,
    `Active criteria: ${ranking.prepared.criteria.length}`,
    "",
    "Top alternatives (best to worst):",
    "Rank  Alternative                     Average rank"
  ];

  ranking.order
    .slice(0, 12)
    .forEach(
      (alternativeIndex, rankIndex) => {
        const name =
          ranking.prepared.alternatives[
            alternativeIndex
          ];

        lines.push(
          `${String(rankIndex + 1).padStart(4)}  ${name.padEnd(30).slice(0, 30)}  ${formatNumber(ranking.averageRank[alternativeIndex])}`
        );
      }
    );

  if (ranking.type === "single") {
    lines.push(
      "",
      `Ranking was aggregated from ${ranking.scoreSets.length} normalization / preference configurations.`
    );
  } else {
    lines.push(
      "",
      `Final ranking is the average of ${ranking.methodResults.length} MCDM methods.`
    );
  }

  return lines.join("\n");
}

// ============================================================
// CRITIC AND ENTROPY WEIGHTS
// ============================================================

weightsButton.addEventListener("click", async () => {
  try {
    const prepared=prepareActiveModel(); setStatus("Calculating CRITIC / Entropy weights in your browser...");
    await nextFrame();
    state.weightsResult=runEngine("calculateObjectiveWeights",prepared); renderWeights(); exportButton.disabled=false; openTab("weightsTab"); setStatus(buildWeightsSummary(state.weightsResult));
  } catch(error) { setStatus(`Weight calculation error: ${error.message}`); }
});

function calculateObjectiveWeights(prepared) {
  const critic = [];
  const entropy = [];

  for (
    let parameter = 1;
    parameter <= 5;
    parameter++
  ) {
    const V =
      normalizeDecisionMatrix(
        prepared.D,
        parameter
      );

    const inverse =
      applyCostTransformation(
        V,
        prepared.D,
        prepared.signs,
        parameter
      );

    const normalized =
      V.map((row, rowIndex) =>
        row.map(
          (value, column) =>
            prepared.signs[column] === 1
              ? value
              : inverse[rowIndex][column]
        )
      );

    critic.push(
      criticWeights(normalized)
    );

    entropy.push(
      entropyWeights(normalized)
    );
  }

  return {
    prepared,
    critic,
    entropy,
    inputWeights:
      [...prepared.weights],
    methods:
      [...normalizationNames]
  };
}

function criticWeights(matrix) {
  const n = matrix[0].length;

  const standardDeviations =
    Array.from(
      { length: n },
      (_, column) =>
        sampleStandardDeviation(
          matrix.map(
            row => row[column]
          )
        )
    );

  const correlation =
    correlationMatrix(matrix);

  const information =
    Array.from(
      { length: n },
      (_, column) => {
        const conflict =
          correlation[column]
            .reduce(
              (total, value) =>
                total +
                (
                  1 -
                  finiteOrZero(value)
                ),
              0
            );

        return (
          standardDeviations[column] *
          conflict
        );
      }
    );

  return normalizeNonnegative(information);
}

function entropyWeights(matrix) {
  const m = matrix.length;
  const n = matrix[0].length;
  const coefficient =
    m > 1
      ? 1 / Math.log(m)
      : 0;

  const diversification =
    new Array(n).fill(0);

  for (let column = 0; column < n; column++) {
    let sumXLogX = 0;

    for (let row = 0; row < m; row++) {
      const value =
        Math.max(
          finiteOrZero(
            matrix[row][column]
          ),
          Number.EPSILON
        );

      sumXLogX +=
        value *
        Math.log(value);
    }

    const entropy =
      -coefficient *
      sumXLogX;

    diversification[column] =
      Math.max(
        0,
        1 - entropy
      );
  }

  return normalizeNonnegative(
    diversification
  );
}

function renderWeights() {
  updateWeightCharts();
  renderWeightsTable();
}

function updateWeightCharts() {
  const result =
    state.weightsResult;

  if (!result) {
    return;
  }

  destroyChart("critic");
  destroyChart("entropy");

  const selected =
    state.selectedCriterion;

  const criterionLabels =
    result.prepared.criteria.map(
      criterion =>
        criterion.notation ||
        criterion.name
    );

  const methods = [
    ...result.methods,
    "Input"
  ];

  const criticRows = [
    ...result.critic,
    result.inputWeights
  ];

  const entropyRows = [
    ...result.entropy,
    result.inputWeights
  ];

  const criticDatasets =
    criterionLabels.map(
      (label, criterionIndex) => ({
        label,
        data:
          criticRows.map(
            row =>
              row[criterionIndex]
          ),
        borderWidth:
          criterionIndex === selected
            ? 3
            : 0.5
      })
    );

  const entropyDatasets =
    criterionLabels.map(
      (label, criterionIndex) => ({
        label,
        data:
          entropyRows.map(
            row =>
              row[criterionIndex]
          ),
        borderWidth:
          criterionIndex === selected
            ? 3
            : 0.5
      })
    );

  charts.critic = new Chart(
    document.getElementById("criticChart"),
    {
      type: "bar",
      data: {
        labels: methods,
        datasets: criticDatasets
      },
      options:
        stackedWeightOptions()
    }
  );

  charts.entropy = new Chart(
    document.getElementById("entropyChart"),
    {
      type: "bar",
      data: {
        labels: methods,
        datasets: entropyDatasets
      },
      options:
        stackedWeightOptions()
    }
  );
}

function stackedWeightOptions() {
  return {
    responsive: true,
    maintainAspectRatio: false,
    animation: false,
    scales: {
      x: {
        stacked: true,
        title: {
          display: true,
          text:
            "Normalization method"
        }
      },
      y: {
        stacked: true,
        beginAtZero: true,
        title: {
          display: true,
          text: "Weight"
        }
      }
    }
  };
}

function renderWeightsTable() {
  const result =
    state.weightsResult;

  const criteria =
    result.prepared.criteria;

  const head =
    weightsTable.querySelector("thead");

  const body =
    weightsTable.querySelector("tbody");

  head.innerHTML = `
    <tr>
      <th>Weight set</th>
      ${criteria.map(
        criterion =>
          `<th>${escapeHtml(criterion.notation || criterion.name)}</th>`
      ).join("")}
    </tr>
  `;

  const rows = [];

  result.methods.forEach(
    (method, index) => {
      rows.push({
        name:
          `CRITIC – ${method}`,
        values:
          result.critic[index]
      });
    }
  );

  result.methods.forEach(
    (method, index) => {
      rows.push({
        name:
          `Entropy – ${method}`,
        values:
          result.entropy[index]
      });
    }
  );

  rows.push({
    name: "Input weights",
    values:
      result.inputWeights
  });

  body.innerHTML =
    rows.map(row => `
      <tr>
        <td>${escapeHtml(row.name)}</td>
        ${row.values.map(
          value =>
            `<td>${formatNumber(value)}</td>`
        ).join("")}
      </tr>
    `).join("");
}

function buildWeightsSummary(result) {
  const lines = [
    "CRITIC / Entropy weights",
    `Active criteria: ${result.prepared.criteria.length}`,
    "",
    "Input weights:"
  ];

  result.prepared.criteria.forEach(
    (criterion, index) => {
      lines.push(
        `${String(criterion.notation || criterion.name).padEnd(18)} : ${formatNumber(result.inputWeights[index])}`
      );
    }
  );

  lines.push(
    "",
    "CRITIC and Entropy weights were calculated for Max, Sum, Vector, Max-Min, and DEA normalization."
  );

  return lines.join("\n");
}

// ============================================================
// EXTRA OUTPUT
// ============================================================

extraOutput.addEventListener("change", () => {
  if (
    !state.ranking ||
    !extraOutput.value
  ) {
    return;
  }

  renderExtraOutput(
    extraOutput.value
  );

  openTab("extraTab");
});

function renderExtraOutput(type) {
  destroyChart("extra");

  const ranking =
    state.ranking;

  const title =
    document.getElementById("extraTitle");

  if (
    ["Max", "Sum", "Vector", "Max-Min", "DEA"].includes(type)
  ) {
    if (ranking.type !== "single") {
      setStatus(
        "Normalization-specific ranks are not shown for All methods. Select one MCDM method first."
      );
      return;
    }

    const scoreSet =
      ranking.scoreSets.find(
        item => item.label === type
      );

    if (!scoreSet) {
      setStatus(
        `${type} output is not available for ${ranking.method}.`
      );
      return;
    }

    const order =
      sortedIndices(
        scoreSet.scores,
        false
      );

    const ranks =
      rankVectorFromOrder(order);

    title.textContent =
      `${type} normalization – rank position`;

    charts.extra = new Chart(
      document.getElementById("extraChart"),
      {
        type: "bar",
        data: {
          labels:
            ranking.prepared.alternatives,
          datasets: [{
            label: "Rank",
            data: ranks
          }]
        },
        options:
          categoryChartOptions(
            "Alternative",
            "Rank (1 is best)"
          )
      }
    );

    return;
  }

  if (type === "Alt ranks") {
    title.textContent =
      "Alternative rank probabilities";

    charts.extra = new Chart(
      document.getElementById("extraChart"),
      {
        type: "bar",
        data: {
          labels:
            ranking.prepared.alternatives,
          datasets: [
            {
              label: "Rank I",
              data:
                ranking.rankPercentages.map(
                  row => row[0]
                )
            },
            {
              label: "Rank II",
              data:
                ranking.rankPercentages.map(
                  row => row[1]
                )
            },
            {
              label: "Rank III",
              data:
                ranking.rankPercentages.map(
                  row => row[2]
                )
            }
          ]
        },
        options:
          categoryChartOptions(
            "Alternative",
            "Probability (%)"
          )
      }
    );

    return;
  }

  if (type === "Rank score") {
    title.textContent =
      "Aggregate rank score";

    charts.extra = new Chart(
      document.getElementById("extraChart"),
      {
        type: "bar",
        data: {
          labels:
            ranking.order.map(
              index =>
                ranking.prepared.alternatives[index]
            ),
          datasets: [{
            label: "Average rank",
            data:
              ranking.order.map(
                index =>
                  ranking.averageRank[index]
              )
          }]
        },
        options:
          categoryChartOptions(
            "Alternative",
            "Average rank"
          )
      }
    );
  }
}

// ============================================================
// EXCEL EXPORT
// ============================================================

exportButton.addEventListener("click", () => {
  try {
    if (
      !state.ranking &&
      !state.weightsResult
    ) {
      throw new Error(
        "Calculate ranking or objective weights first."
      );
    }

    const workbook =
      XLSX.utils.book_new();

    appendSheet(
      workbook,
      decisionMatrixRows(),
      "DecisionMatrix"
    );

    appendSheet(
      workbook,
      criteriaSettingRows(),
      "CriteriaSettings"
    );

    if (state.ranking) {
      appendSheet(
        workbook,
        rankingSummaryRows(),
        "RankingSummary"
      );

      appendSheet(
        workbook,
        rankCountRows(),
        "RankCounts"
      );

      appendSheet(
        workbook,
        methodRankRows(),
        "MethodRanks"
      );
    }

    if (state.weightsResult) {
      appendSheet(
        workbook,
        weightRows(
          state.weightsResult.critic,
          "CRITIC"
        ),
        "CRITICWeights"
      );

      appendSheet(
        workbook,
        weightRows(
          state.weightsResult.entropy,
          "Entropy"
        ),
        "EntropyWeights"
      );
    }

    XLSX.writeFile(
      workbook,
      `MCDM_results_${timestamp()}.xlsx`,
      {
        compression: true
      }
    );

    setStatus(
      "Results workbook exported successfully."
    );
  } catch (error) {
    setStatus(
      `Export error: ${error.message}`
    );
  }
});

function decisionMatrixRows() {
  return state.model.alternatives.map(
    (alternative, rowIndex) => {
      const row = {
        Alternative: alternative
      };

      state.model.criteria.forEach(
        (criterion, columnIndex) => {
          row[
            criterion.notation ||
            criterion.name
          ] =
            state.model.matrix[
              rowIndex
            ][columnIndex];
        }
      );

      return row;
    }
  );
}

function criteriaSettingRows() {
  return state.model.criteria.map(
    criterion => ({
      Criteria: criterion.name,
      Notation: criterion.notation,
      Weight: criterion.weight,
      Attribution:
        criterion.sign === 1
          ? "Positive"
          : "Negative",
      Available:
        criterion.available
    })
  );
}

function rankingSummaryRows() {
  const ranking =
    state.ranking;

  return ranking.order.map(
    alternativeIndex => ({
      Alternative:
        ranking.prepared.alternatives[
          alternativeIndex
        ],
      FinalRank:
        ranking.finalRank[
          alternativeIndex
        ],
      AverageRank:
        ranking.averageRank[
          alternativeIndex
        ],
      RankSum:
        ranking.rankSum[
          alternativeIndex
        ],
      RankI_Percent:
        ranking.rankPercentages[
          alternativeIndex
        ]?.[0] ?? 0,
      RankII_Percent:
        ranking.rankPercentages[
          alternativeIndex
        ]?.[1] ?? 0,
      RankIII_Percent:
        ranking.rankPercentages[
          alternativeIndex
        ]?.[2] ?? 0,
      Method:
        ranking.method
    })
  );
}

function rankCountRows() {
  const ranking =
    state.ranking;

  return ranking.prepared.alternatives.map(
    (alternative, alternativeIndex) => {
      const row = {
        Alternative: alternative
      };

      ranking.rankCounts[
        alternativeIndex
      ].forEach(
        (count, rankIndex) => {
          row[
            `Rank_${rankIndex + 1}`
          ] = count;
        }
      );

      return row;
    }
  );
}

function methodRankRows() {
  const ranking =
    state.ranking;

  if (ranking.type === "single") {
    return ranking.prepared.alternatives.map(
      (alternative, alternativeIndex) => {
        const row = {
          Alternative: alternative
        };

        ranking.scoreSets.forEach(
          (set, setIndex) => {
            row[
              set.label
            ] =
              ranking.rankVectors[
                setIndex
              ][alternativeIndex];

            row[
              `${set.label}_Score`
            ] =
              set.scores[
                alternativeIndex
              ];
          }
        );

        return row;
      }
    );
  }

  return ranking.prepared.alternatives.map(
    (alternative, alternativeIndex) => {
      const row = {
        Alternative: alternative
      };

      ranking.methodResults.forEach(
        result => {
          row[
            result.method
          ] =
            result.averageRank[
              alternativeIndex
            ];
        }
      );

      return row;
    }
  );
}

function weightRows(matrix, prefix) {
  const result =
    state.weightsResult;

  const rows =
    result.methods.map(
      (method, methodIndex) => {
        const row = {
          WeightType: prefix,
          Normalization: method
        };

        result.prepared.criteria.forEach(
          (criterion, criterionIndex) => {
            row[
              criterion.notation ||
              criterion.name
            ] =
              matrix[
                methodIndex
              ][criterionIndex];
          }
        );

        return row;
      }
    );

  const inputRow = {
    WeightType: "Input",
    Normalization: "Input"
  };

  result.prepared.criteria.forEach(
    (criterion, criterionIndex) => {
      inputRow[
        criterion.notation ||
        criterion.name
      ] =
        result.inputWeights[
          criterionIndex
        ];
    }
  );

  rows.push(inputRow);

  return rows;
}

function appendSheet(
  workbook,
  rows,
  name
) {
  const worksheet =
    XLSX.utils.json_to_sheet(
      rows.length > 0
        ? rows
        : [{
            Message: "No data"
          }]
    );

  XLSX.utils.book_append_sheet(
    workbook,
    worksheet,
    name
  );
}

// ============================================================
// EXCEL TEMPLATE
// ============================================================

templateButton.addEventListener("click", () => {
  const workbook =
    XLSX.utils.book_new();

  const rows = [
    [
      "Criteria",
      "Notation",
      "Alternative A",
      "Alternative B",
      "Alternative C",
      "Weight",
      "Attribution",
      "Available"
    ],
    [
      "Cost",
      "C1",
      260,
      275,
      300,
      0.30,
      "Negative",
      true
    ],
    [
      "Capacity",
      "C2",
      16,
      18,
      20,
      0.40,
      "Positive",
      true
    ],
    [
      "Safety",
      "C3",
      15,
      18,
      17,
      0.30,
      "Positive",
      true
    ]
  ];

  const worksheet =
    XLSX.utils.aoa_to_sheet(rows);

  XLSX.utils.book_append_sheet(
    workbook,
    worksheet,
    "DecisionMatrix"
  );

  XLSX.writeFile(
    workbook,
    "MCDM_decision_matrix_template.xlsx",
    {
      compression: true
    }
  );
});

// ============================================================
// RESET / ENABLEMENT
// ============================================================

resetButton.addEventListener("click", () => {
  invalidateResults();
  setStatus(
    "Calculated results were cleared. Input data remain available."
  );
});

function enableModelControls() {
  calculateButton.disabled = false;
  weightsButton.disabled = false;
  resetButton.disabled = false;
  criteriaFilter.disabled = false;
}

function disableModelControls() {
  calculateButton.disabled = true;
  weightsButton.disabled = true;
  exportButton.disabled = true;
  resetButton.disabled = true;
  extraOutput.disabled = true;
  criteriaFilter.disabled = true;
}

function invalidateResults() {
  state.ranking = null;
  state.weightsResult = null;
  clearResults();
}

function clearResults() {
  rankingTableBody.innerHTML = "";

  weightsTable
    .querySelector("thead")
    .innerHTML = "";

  weightsTable
    .querySelector("tbody")
    .innerHTML = "";

  Object.keys(charts).forEach(
    destroyChart
  );

  exportButton.disabled = true;
  extraOutput.disabled = true;
  extraOutput.value = "";

  document
    .getElementById("extraTitle")
    .textContent =
      "Select an extra output from the left panel.";
}

// ============================================================
// CHART HELPERS
// ============================================================

function destroyChart(name) {
  if (charts[name]) {
    charts[name].destroy();
    charts[name] = null;
  }
}

function categoryChartOptions(
  xLabel,
  yLabel
) {
  return {
    responsive: true,
    maintainAspectRatio: false,
    animation: false,
    scales: {
      x: {
        title: {
          display: true,
          text: xLabel
        },
        ticks: {
          maxRotation: 80,
          minRotation: 0
        }
      },
      y: {
        beginAtZero: true,
        title: {
          display: true,
          text: yLabel
        }
      }
    }
  };
}

// ============================================================
// NUMERIC / MATRIX HELPERS
// ============================================================

function applyColumnWeights(
  matrix,
  weights
) {
  return matrix.map(row =>
    row.map(
      (value, column) =>
        value *
        weights[column]
    )
  );
}

function matrixVectorProduct(
  matrix,
  vector
) {
  return matrix.map(row =>
    row.reduce(
      (total, value, column) =>
        total +
        value *
        vector[column],
      0
    )
  );
}

function columnMinimums(matrix) {
  return Array.from(
    { length: matrix[0].length },
    (_, column) =>
      Math.min(
        ...matrix.map(
          row => row[column]
        )
      )
  );
}

function columnMaximums(matrix) {
  return Array.from(
    { length: matrix[0].length },
    (_, column) =>
      Math.max(
        ...matrix.map(
          row => row[column]
        )
      )
  );
}

function columnSums(matrix) {
  return Array.from(
    { length: matrix[0].length },
    (_, column) =>
      sum(
        matrix.map(
          row => row[column]
        )
      )
  );
}

function lpDistance(
  first,
  second,
  metric
) {
  const differences =
    first.map(
      (value, index) =>
        Math.abs(
          value -
          second[index]
        )
    );

  if (metric === 1) {
    return sum(differences);
  }

  if (metric === 2) {
    return Math.sqrt(
      sum(
        differences.map(
          value => value ** 2
        )
      )
    );
  }

  if (metric === Infinity) {
    return Math.max(...differences);
  }

  return Math.pow(
    sum(
      differences.map(
        value =>
          value ** metric
      )
    ),
    1 / metric
  );
}

function sortedIndices(
  values,
  ascending
) {
  return values
    .map(
      (value, index) => ({
        value:
          Number.isFinite(value)
            ? value
            : ascending
              ? Infinity
              : -Infinity,
        index
      })
    )
    .sort(
      (first, second) =>
        ascending
          ? (
              first.value -
              second.value ||
              first.index -
              second.index
            )
          : (
              second.value -
              first.value ||
              first.index -
              second.index
            )
    )
    .map(item => item.index);
}

function rankVectorFromOrder(order) {
  const ranks =
    new Array(order.length);

  order.forEach(
    (alternativeIndex, rankIndex) => {
      ranks[alternativeIndex] =
        rankIndex + 1;
    }
  );

  return ranks;
}

function correlationMatrix(matrix) {
  const n =
    matrix[0].length;

  const columns =
    Array.from(
      { length: n },
      (_, column) =>
        matrix.map(
          row => row[column]
        )
    );

  return Array.from(
    { length: n },
    (_, first) =>
      Array.from(
        { length: n },
        (_, second) =>
          pearsonCorrelation(
            columns[first],
            columns[second]
          )
      )
  );
}

function pearsonCorrelation(
  first,
  second
) {
  const meanFirst =
    mean(first);

  const meanSecond =
    mean(second);

  let numerator = 0;
  let denominatorFirst = 0;
  let denominatorSecond = 0;

  for (
    let index = 0;
    index < first.length;
    index++
  ) {
    const dx =
      first[index] -
      meanFirst;

    const dy =
      second[index] -
      meanSecond;

    numerator += dx * dy;
    denominatorFirst += dx ** 2;
    denominatorSecond += dy ** 2;
  }

  const denominator =
    Math.sqrt(
      denominatorFirst *
      denominatorSecond
    );

  if (
    denominator <= 1e-15
  ) {
    return first === second
      ? 1
      : 0;
  }

  return numerator / denominator;
}

function sampleStandardDeviation(values) {
  if (values.length < 2) {
    return 0;
  }

  const average =
    mean(values);

  return Math.sqrt(
    sum(
      values.map(
        value =>
          (
            value -
            average
          ) ** 2
      )
    ) /
    (
      values.length -
      1
    )
  );
}

function normalizeNonnegative(values) {
  const cleaned =
    values.map(
      value =>
        Number.isFinite(value) &&
        value > 0
          ? value
          : 0
    );

  const total =
    sum(cleaned);

  if (total <= 1e-15) {
    return cleaned.map(
      () =>
        cleaned.length > 0
          ? 1 / cleaned.length
          : 0
    );
  }

  return cleaned.map(
    value => value / total
  );
}

function safeDivide(
  numerator,
  denominator
) {
  if (
    !Number.isFinite(numerator) ||
    !Number.isFinite(denominator) ||
    Math.abs(denominator) <
    1e-15
  ) {
    return 0;
  }

  return numerator / denominator;
}

function safeReciprocal(value) {
  return safeDivide(1, value);
}

function finiteOrZero(value) {
  const number =
    Number(value);

  return Number.isFinite(number)
    ? number
    : 0;
}

function sum(values) {
  return values.reduce(
    (total, value) =>
      total +
      finiteOrZero(value),
    0
  );
}

function mean(values) {
  return values.length > 0
    ? sum(values) /
      values.length
    : 0;
}

function formatNumber(value) {
  if (!Number.isFinite(value)) {
    return "NaN";
  }

  const absolute =
    Math.abs(value);

  if (
    absolute !== 0 &&
    (
      absolute >= 100000 ||
      absolute < 0.0001
    )
  ) {
    return value.toExponential(5);
  }

  return Number(value)
    .toFixed(6)
    .replace(/\.?0+$/, "");
}

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function timestamp() {
  const date = new Date();

  return [
    date.getFullYear(),
    String(
      date.getMonth() + 1
    ).padStart(2, "0"),
    String(
      date.getDate()
    ).padStart(2, "0"),
    "_",
    String(
      date.getHours()
    ).padStart(2, "0"),
    String(
      date.getMinutes()
    ).padStart(2, "0"),
    String(
      date.getSeconds()
    ).padStart(2, "0")
  ].join("");
}

function setStatus(text) {
  statusBox.textContent =
    String(text ?? "");
}

// ============================================================
// INITIAL SCREEN
// ============================================================

disableModelControls();

// ============================================================
// CALCULATION ENGINE: runs in the browser (js/mcdm-engine.js), no backend
// ============================================================

/* Same input/output as the former /api/ranking and /api/weights:
   the JSON round-trip keeps results identical (NaN/Infinity -> null). */
function runEngine(fn, ...args) {
  if (!window.MCDMEngine) throw new Error("Calculation engine not loaded (js/mcdm-engine.js). Reload the page.");
  const input = args.map((a) => (a && typeof a === "object" ? JSON.parse(JSON.stringify(a)) : a));
  return JSON.parse(JSON.stringify(window.MCDMEngine[fn](...input)));
}

function nextFrame() {
  return new Promise((resolve) => requestAnimationFrame(() => setTimeout(resolve, 0)));
}

(function initEngineBadge() {
  const dot = document.getElementById("backendDot");
  const txt = document.getElementById("backendText");
  if (!dot || !txt) return;
  const ok = !!window.MCDMEngine;
  dot.className = ok ? "dot ok" : "dot bad";
  txt.textContent = ok ? "Runs in your browser \u00b7 no server" : "Engine not loaded";
})();
