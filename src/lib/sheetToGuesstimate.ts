import * as ExcelJS from "exceljs";
import { escapeRegExp, zip } from "lodash";

interface CellData {
  address: string; // "A1"
  value: string;
  description?: string;
  formula?: string;
  rowNum: number;
  colNum: number;
}

export interface Grid {
  cells: { [address: string]: CellData };
  sheetNames: string[];
}

// Guesstimate types
interface Guesstimate {
  metric: string;
  input: null;
  expression: string;
  guesstimateType: "FUNCTION" | "POINT";
  description: string;
}

interface Metric {
  id: string;
  readableId: string;
  name: string;
  location: {
    row: number;
    column: number;
  };
}

interface Graph {
  metrics: Metric[];
  guesstimates: Guesstimate[];
}

export async function xlsxToGrid(data: ArrayBuffer): Promise<Grid> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(data);
  const grid = makeGrid(workbook);
  return grid;
}

export function gridToGuesstimateScript(grid: Grid): string {
  const graph = makeGraph(grid);
  return makeScript(graph);
}

function makeGrid(workbook: ExcelJS.Workbook): Grid {
  const grid: Grid = {
    sheetNames: workbook.worksheets.map((sheet) => sheet.name),
    cells: {},
  };

  let maxColumn = 0;
  workbook.eachSheet(function (sheet) {
    // Put multiple sheets to the right of each other
    const startColumn = maxColumn + 1;
    sheet.eachRow(function (row, rowNum) {
      let previousCell: CellData | undefined = undefined;
      row.eachCell(function (cell, colNum) {
        colNum += startColumn;
        maxColumn = Math.max(maxColumn, colNum);
        const cellData = getCellData(cell, grid.sheetNames, rowNum, colNum);
        // Use text in cell to the left as description of current cell,
        // if it's not a formula and doesn't have a value
        if (
          (cellData.value || cellData.formula) &&
          previousCell &&
          previousCell.colNum === cellData.colNum - 1 &&
          !previousCell.formula &&
          previousCell.description &&
          previousCell.value === ""
        ) {
          cellData.description = previousCell.description;
          delete grid.cells[previousCell.address];
        }
        grid.cells[cellData.address] = cellData;
        previousCell = cellData;
      });
    });
  });
  return grid;
}

function getCellData(
  cell: ExcelJS.Cell,
  sheetNames: string[],
  rowNum: number,
  colNum: number
): CellData {
  let address = cell.fullAddress.address;
  let text = ""
  try {
    text = cell.text || "";
  } catch (e) {
    text = "";
  }
  if (typeof text === "object") {
    text = (text as any).richText.map((e: any) => e.text).join("");
  }
  if (typeof cell.value === "number") {
    text = cell.value.toString();
  }
  text = text.toString().trim();

  if (/^-?\s?\d+\.?\d*%$/.test(text)) {
    text = (parseFloat(text) / 100).toString();
  }

  if (!isNaN(text.replaceAll(",", ".") as any)) {
    text = text.replaceAll(",", ".");
  }

  // isNaN is used to check if the text is not a number string
  const description = isNaN(text as any) ? text : "";
  const value = isNaN(text as any) ? "" : text;
  let formula = cell.formula ? `=${cell.formula}` : "";
  if (formula.startsWith("=HYPERLINK")) formula = "";
  formula = formula.replaceAll("$", "") || "";
  formula = evalPercentage(formula);
  formula = inlineSUMPRODUCT(formula);
  formula = inlineAverage(formula);
  formula = inlineSum(formula);
  formula = replaceLN(formula);
  formula = replacePV(formula);

  if (sheetNames.length > 1) {
    const makePrefix = (sheetName: string) =>
      sheetName.replace(/\W/g, "") + "!";
    address = makePrefix(cell.fullAddress.sheetName) + address;
    formula = formula.replaceAll(
      /(?<!!)([A-Z][0-9]+)/g,
      `${makePrefix(cell.fullAddress.sheetName)}$1`
    );
    for (let sheetName of sheetNames) {
      if (sheetName.includes(" ")) {
        sheetName = `'${sheetName.replace("'", "''")}'`;
      }
      const newSheetPrefix = makePrefix(sheetName);
      formula = formula.replaceAll(
        new RegExp(`${sheetName}!`, "gi"),
        newSheetPrefix
      );
      formula = formula.replaceAll(
        new RegExp(`(${newSheetPrefix}[A-Z][0-9]+)`, "g"),
        "${metric:$1}"
      );
    }
  } else {
    formula = formula.replaceAll(/([A-Z][0-9]+)/g, "${metric:$1}");
  }

  return {
    address,
    value,
    description,
    formula,
    rowNum,
    colNum,
  };
}

function makeGraph(grid: Grid): Graph {
  const graph = {
    metrics: [],
    guesstimates: [],
  };
  for (let cell of Object.values(grid.cells)) {
    let id = cell.address;
    const guesstimateType = cell.formula ? "FUNCTION" : "POINT";
    const expression = cell.formula || cell.value;
    let name = cell.description || "";
    let description = "";
    if (name.length > 20 && cell.value !== "") {
      description = name;
      name = name.slice(0, 17) + "...";
    }
    graph.metrics.push({
      id,
      readableId: id,
      name,
      location: {
        row: cell.rowNum - 1,
        column: cell.colNum - 1,
      },
    });
    graph.guesstimates.push({
      metric: id,
      input: null,
      expression,
      guesstimateType,
      description,
    });
  }
  return graph;
}

function makeScript(graph: Graph): string {
  const body = JSON.stringify({
    space: {
      graph,
    },
  });

  return `const modelId = window.location.pathname.split('/')[2];
const authorization = "Bearer " + JSON.parse(localStorage.getItem("Guesstimate-Testme")).token;
const payload = {
  headers: {
    accept: "application/json, text/javascript, */*; q=0.01",
    authorization: authorization,
    "cache-control": "no-cache",
    "content-type": "application/json",
  },
  referrer: "https://www.getguesstimate.com/",
  referrerPolicy: "strict-origin-when-cross-origin",
  body: JSON.stringify(${body}),
  method: "PATCH",
};
fetch("https://guesstimate.herokuapp.com/spaces/" + modelId, payload).then(() => location.reload());`;
}

export function gridToMermaidGraph(grid: Grid): string {
  const seen = new Set();
  let lines = [];
  function nodeText(cell: CellData): string {
    let description = cell.description;
    if (description.length > 40) {
      description = description.slice(0, 37) + "...";
    }
    if (!description) {
      return "";
    }
    return `${cell.address}[${description.replaceAll(
      /[^ a-zA-Z0-9,.\-$%]/g,
      "_"
    )}]`;
  }

  const addresses = Object.keys(grid.cells);
  for (let cell of Object.values(grid.cells)) {
    if (!cell.formula) continue;
    const matchedAddresses = addresses.filter((address) =>
      cell.formula.match(new RegExp(escapeRegExp(address) + "}"))
    );
    for (let address of matchedAddresses) {
      if (!seen.has(address)) {
        lines.push(nodeText(grid.cells[address]));
      }
      if (!seen.has(cell.address)) {
        lines.push(nodeText(cell));
      }
      seen.add(address);
      seen.add(cell.address);
      lines.push(`${address} --> ${cell.address}`);
    }
  }
  return `graph LR
  ${lines.join("\n")}`;
}

function deRange(start: string, end: string): string[] {
  const results = [];
  const letterStart = start[0];
  const letterEnd = end[0];
  const numberStart = parseInt(start.slice(1), 10);
  const numberEnd = parseInt(end.slice(1), 10);
  function nextLetter(letter: string): string {
    return String.fromCharCode(letter.charCodeAt(0) + 1);
  }
  for (
    let letter = letterStart;
    letter <= letterEnd;
    letter = nextLetter(letter)
  ) {
    for (let number = numberStart; number <= numberEnd; number++) {
      results.push(`${letter}${number}`);
    }
  }
  return results;
}

function inlineSum(formula: string): string {
  // Replaces =SUM(A10:A13) with =A10+A11+A12+A13
  // Replaces =SUM(A10, B1, C3) with =sum(A10, B1, C3)
  // Otherwise replace sum with _SUM, to aviod messy results
  // TODO use deRange

  if (!formula.toUpperCase().includes("SUM")) {
    return formula;
  }

  formula = formula.toUpperCase().replaceAll("SUM", "_SUM");
  const sumsNumber = formula.match(/SUM/g).length;
  const commaSums = formula.match(/_SUM\(([A-Z][0-9]+,?)+\)/g) || [];
  // sum(A1, B4, C5, D3) apparently works fine, just needs to be lowercase
  if (commaSums.length == sumsNumber) {
    return formula.replaceAll("_SUM", "sum");
  }
  const sums = formula.match(/_SUM\([A-Z][0-9]+:[A-Z][0-9]+\)/g);
  if (sums === null || sums.length != sumsNumber) {
    return formula;
  }
  for (let sum of sums) {
    const contents = sum.slice(5, -1);
    const [start, end] = contents.split(":");
    if (start[0] !== end[0]) {
      return formula;
    }
    const letter = start[0];
    const startNumber = Number(start.slice(1));
    const endNumber = Number(end.slice(1));
    if (isNaN(startNumber) || isNaN(endNumber)) {
      return formula;
    }
    let newSum = `${letter}${startNumber}`;
    for (let n = startNumber + 1; n <= endNumber; n++) {
      newSum += `+${letter}${n}`;
    }
    formula = formula.replaceAll(sum, `(${newSum})`);
  }
  return formula;
}

function inlineSUMPRODUCT(formula: string): string {
  const matches = formula.matchAll(
    /SUMPRODUCT\(([A-Z][0-9]+):([A-Z][0-9]+),([A-Z][0-9]+):([A-Z][0-9]+)\)/gi
  );
  if (!matches) {
    return formula;
  }
  for (let match of matches) {
    const [startSum, endSum, startProd, endProd] = [
      match[1],
      match[2],
      match[3],
      match[4],
    ];
    const sumRange = deRange(startSum, endSum);
    const prodRange = deRange(startProd, endProd);
    if (sumRange.length === 0 || sumRange.length != prodRange.length) {
      return formula;
    }
    let inlinedSumProd = "";
    for (let [sum, prod] of zip(sumRange, prodRange)) {
      inlinedSumProd += `+${sum}*${prod}`;
    }
    formula = formula.replaceAll(match[0], `(${inlinedSumProd})`);
  }
  return formula;
}

function replaceLN(formula: string): string {
  return formula.replaceAll(
    /LN\(([^()]+)\)/gi,
    "(1 / log10(2.71828) * log10($1))"
  );
}

function replacePV(formula: string): string {
  // PV($1,$2,$3) --> (-$3)*(1-(1+$1)^(-$2))/$1
  return formula.replaceAll(
    /PV\(([^,]+),([^,]+),([^,]+)(?:,,1)?\)/gi,
    "(-($3)*(1-(1+($1))^(-($2)))/($1))"
  );
}

function inlineAverage(formula: string): string {
  const matches = formula.matchAll(/AVERAGE\(([^()]+)\)/gi);
  if (!matches) {
    return formula;
  }
  for (let match of matches) {
    const args = match[1]
      .split(",")
      .flatMap((arg) =>
        arg.includes(":") ? deRange(arg.split(":")[0], arg.split(":")[1]) : arg
      );
    formula = formula.replaceAll(
      match[0],
      `((${args.join("+")})/${args.length})`
    );
  }
  return formula;
}

function evalPercentage(formula: string): string {
  const matches = formula.matchAll(/([0-9.]+)%/g);
  if (!matches) {
    return formula;
  }
  for (let match of matches) {
    formula = formula.replaceAll(match[0], `${parseFloat(match[1]) / 100}`);
  }
  return formula;
}
