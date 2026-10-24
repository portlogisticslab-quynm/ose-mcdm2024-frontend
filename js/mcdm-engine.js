/* MCDM calculation engine (browser). Ported 1:1 from ose-mcdm2024-backend/src/mcdm.js — no server needed. */
(function (global) {
"use strict";
const normalizationNames=["Max","Sum","Vector","Max-Min","DEA"];
const methodNames=["SAW","CODAS","COPRAS","TOPSIS-SC","TOPSIS-Eu","TOPSIS-Inf","GRAt","PROMETHEE"];
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
function validatePrepared(p){if(!p||!Array.isArray(p.D)||p.D.length<2)throw new Error("At least two alternatives are required.");if(!Array.isArray(p.D[0])||p.D[0].length<1)throw new Error("At least one active criterion is required.");}
function calculateRanking(prepared,method){validatePrepared(prepared);return method==="All_methods"?calculateAllMethods(prepared):calculateSingleMethod(prepared,method);}
global.MCDMEngine = Object.freeze({ calculateRanking, calculateObjectiveWeights });
})(window);
