import xlsx from "xlsx";
import fs from "fs";
import path from "path";
import {
  sumArray,
  getValueByTitle,
  getValuesBelowTitleAuto,
  mergeToObjectArray,
  buildMeasFormula,
  calculateFromFormula,
  filterColorsInCodes,
  getValuesBelowTitleAutoVoLumn_n_CBM,
  filterOnlyNumericValues,
  collapseIfSame,
  isArray,
  removeAfterString,
  convertCmFormulaToMeter,
  keepOnlyDimensionLxWxH,
  normalizeCartonArray,
  extractAllColorCodesFromWorkbook,
  smartSumFloat,
  removeAfterStringValue,
  smartCountSum,
  trimArrayByArray,
  removeAllBlockTotals,
  removeItemFromArray,
  removeItemsByMeasString,
  extractCountArray,
  extractItemArray,
  mergeToObjectArrayLevel2,
  arrayContainsString,
  getAllPOBlocks,
  hasAtLeastNEmptyCells,
  calcSum2Arr,
  getValuesBelowTitleAutoForTotalActiveBULK,
  getValuesBelowTitleAutoForColorACTIVE_BULK,
  removeBeforeTitle,
  extractPurchaseOrderNo,
  getValuesBelowTitleAutoForSizeCBM,
  buildCartonDimensionFormulas,
  getValuesBelowTitleExact,
  pickNonEmptyArray,
  getValuesBelowTitleAutoForCartonDimension,
  fillCartonDimension,
  normalizeCartonDimension,
  getValuesBelowTitleAuto3Cols,
  findRowIndexesByTitle,
  filterRowsByExcludedTitles,
  getValueForGrossNet_INNOVATION,
  getValuesBelowTitleAutoExcludeKeywords,
  removeFromStringValue,
  getGroupedValuesBelowTitle,
  getValuesBelowTitleAutoForCartonDimension_MAMMUT,
  removeItemsByKeywords,
  hasExactMatch,
  extractFirst5Digits,
  isSheetMatchTemplate,
  cutRowsFromTitleInFirstColumn,
  filter2DArrayByKeyword
} from "../utils/math.util.js";

const OUTPUT_PATH = path.join(process.cwd(), "uploads", "GRAND_TOTAL_consolidated.xlsx");

export function readExcelFile() {
  if (!fs.existsSync(OUTPUT_PATH)) return [];

  const workbook = xlsx.readFile(OUTPUT_PATH);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];

  return xlsx.utils.sheet_to_json(sheet, { defval: "" });
}

export function writeResultToJsonFile(result) {
  const outputPath = path.join(process.cwd(), "uploads", "preview.json");

  fs.writeFileSync(
    outputPath,
    JSON.stringify(result, null, 2), // pretty JSON
    "utf-8"
  );

  return outputPath;
}

const VALID_KEYWORDS = [
  "sheet",
  "article",
  "style",
  "po",
  "color",
  "count",
  "total",
  "net",
  "gross",
  "volume",
  "cbm",
  "carton",
  "dimension",
  "style name",
  "style no",
  "qty of ctns",
  "meas pr ctn",
  "total pcs",
  "g.w./kgs",
  "n.w./kgs",
  "colour",
  "qty",
  "total gw",
  "total nw",
  "total",
  "số kg n.w",
  "số kg g.w",
  "kích thước thùng",
  "tổng cộng",
  "tổng thùng",
  "carton no",
  "number carton",
  "net weight (kg)",
  "gross weight (kg)",
  "cmb",
  "q'ty (pcs)",
  "packing destination",
  "cbm pr ctn",
  "size(cm)",
  "l",
  "w",
  "h",
  "ctns.",
  "art no. color",
  "q'ty",
  "order no",
  "ctns",
  "order number style no. + description",
  "dims",
  "nw(kgs)",
  "gw(kgs)",
  "drn#",
];

function normalize(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/\s+/g, " ")
    .trim();
}

function isValidColumn(rows, columnKey) {
  return rows.some((row) => {
    const cellValue = normalize(row[columnKey]);
    return VALID_KEYWORDS.some((kw) => cellValue.includes(kw));
  });
}

function filterInvalidColumns(rows) {
  if (rows.length === 0) return rows;

  const allKeys = Object.keys(
    rows.reduce((acc, row) => ({ ...acc, ...row }), {})
  );

  const validKeys = allKeys.filter((key) =>
    isValidColumn(rows, key)
  );

  return rows.map((row) =>
    Object.fromEntries(
      Object.entries(row).filter(([key]) =>
        validKeys.includes(key)
      )
    )
  );
}

//==================== new code ====================

function findColumnKeyByTitle(rows, title) {
  if (!rows.length) return null;

  const keys = Object.keys(
    rows.reduce((acc, r) => ({ ...acc, ...r }), {})
  );

  for (const key of keys) {
    for (const row of rows) {
      const cell = normalize(row[key]);
      if (cell === title.toLowerCase()) {
        return key;
      }
    }
  }
  return null;
}

function extractColumnCellArrays(rows, sheetName) {
  if (!rows.length) return [];

  const keys = Object.keys(
    rows.reduce((acc, row) => ({ ...acc, ...row }), {})
  );

  return keys.map((key) => ({
    sheet: sheetName,
    column: key,
    cells: rows
      .map((row) => row[key])
      .filter((v) => String(v || "").trim() !== ""),
  }));
}

function isZeroInCountColumn(row, countKey) {
  if (!countKey) return false;

  const value = row[countKey];

  // null / undefined
  if (value === null || value === undefined) return true;

  // string rỗng
  if (typeof value === "string" && value.trim() === "") return true;

  // số 0 hoặc "0"
  if (
    value === 0 ||
    value === "0" ||
    String(value).trim() === "0"
  )
    return true;

  return false;
}

// hàm lấy value tiếp theo
function getNextValueByKey(cells, key) {
  for (const [index, value] of cells.entries()) {
    if (
      typeof value === "string" &&
      value.trim().toLowerCase() === key.toLowerCase() &&
      index + 1 < cells.length
    ) {
      return cells[index + 1];
    }
  }
  return null;
}

function normalizeHeader(str) {
  return String(str || "")
    .replace(/\n/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .toUpperCase();
}

function hasFromToCells(row) {
  const values = Object.values(row);

  for (let i = 0; i < values.length - 1; i++) {
    const cur = normalizeHeader(values[i]);
    const next = normalizeHeader(values[i + 1]);

    if (cur === "FROM" && next === "TO") {
      return true;
    }
  }

  return false;
}

const EXCEL_TEMPLATES = [
  {
    type: "EDDIE BAUER BULK",
    requiredHeaders: [
      "STYLE NO.",
      "PURCHASE ORDER",
      "COLOR",
      "CARTON NO",
      "TOTAL",
      "TOTAL NW",
      "TOTAL GW",
      "TOTAL CBM"
    ],
  },
  {
    type: "EDDIE BAUER USA SMS",
    requiredHeaders: [
      "DRN#",
      "PO#",
      "COLOR",
      "CARTON NO.",
      "QTY (PCS)",
      "GW (KG)",
    ],
  },
  {
    type: "HAGLOFS TEMPLATE 2",
    requiredHeaders: [
      "STYLE NUMBER",
      "PO-NUMBER",
      "COLOR CODE",
      "TOTAL CTN",
      "TOTAL QUANTITY",
      "NET WEIGHT KG",
      "GROSS WEIGHT KG",
      "CARTON TYPE",
    ],
  },
  {
    type: "HAGLOFS TEMPLATE 1",
    requiredHeaders: [
      "STYLE #",
      "PO",
      "COLOR CODE",
      "NUMBER CARTON",
      "Q'TY (PCS)",
      "NET WEIGHT (KG)",
      "GROSS WEIGHT (KG)",
      "CBM",
      "CARTON DIMENSION(L,W,H)",
    ],
  },
  {
    type: "ACTIVE SMS",
    requiredHeaders: [
      "STYLE NO",
      "PACKING DESTINATION",
      "COLOUR",
      "CARTON NO",
      "QTY",
      "N.W./KGS",
      "G.W./KGS",
      "VOLUMN (CBM)",
    ],
  },
  {
    type: "ACTIVE BULK",
    requiredHeaders: [
      "STYLE NO",
      "PURCHASE ORDER NO:",
      "COLOR",
      "QTY OF CTNS",
      "TOTAL PCS",
      "NET WEIGHT PR CTN",
      "GROSS WEIGHT PR CTN",
      "CBM PR CTN",
    ],
  },
  {
    type: "BARBOUR BULK",
    requiredHeaders: [
      "STYLE",
      "PO",
      "COLOR",
      "CARTON NO",
      "TOTAL",
      "TOTAL NW",
      "TOTAL GW",
      "TOTAL CBM",
      "SIZE(CM)",
    ],
  },
  {
    type: "BARBOUR SMS",
    requiredHeaders: [
      "STYLE",
      "PO",
      "COLOR",
      "CARTON NO",
      "TOTAL",
      "TOTAL NW",
      "TOTAL GW",
      "TOTAL CBM",
      "SIZE(CM)",
    ],
  },
  { // asics bulk, henri, odlo bulk
    type: "ASICS BULK",
    requiredHeaders: [
      "ARTICLE / STYLE",
      "PO",
      "COLOR",
      "COUNT",
      "TOTAL",
      "NET",
      "GROSS",
      "VOLUME (CBM)",
      "CARTON DIMENSION (CM)",
    ],
  },
  {
    type: "ASICS SMS",
    requiredHeaders: [
      "ARTICLE / STYLE",
      "COLOR",
      "COUNT",
      "TOTAL",
      "G.W",
      "N.W",
      "CBM",
      "CARTON DIMENSION (CM)",
    ],
  },
  {
    type: "HENRI BULK",
    requiredHeaders: [
      "ARTICLE / STYLE",
      "ORDER NO:",
      "COLOR",
      "COUNT",
      "TOTAL QTY",
      "GROSS",
      "NET",
      "VOLUME (CBM)",
      "CARTON DIMENSION (CM)",
    ],
  },
  {
    type: "INNOVATION",
    requiredHeaders: [
      "STYLE",
      "PO. NO.",
      "COLOUR",
      "CTNS.",
      "TOTAL QTY.",
    ],
  },
  {
    type: "WITTT THIEN BINH",
    requiredHeaders: [
      "STYLE NO",
      "ORDER NO",
      "ART NO. COLOR",
      "CARTONS",
      "TOTAL PCS",
      "N.W",
      "G.W",
      "CBM (CM)"
    ],
  },
  {
    type: "WITTT BMPY",
    requiredHeaders: [
      "STYLE NO",
      "ORDER NO",
      "ART NO. COLOR",
      "CARTONS",
      "TOTAL PCS",
      "N.W",
      "G.W",
      "(CM)"
    ],
  },
  {
    type: "ODLO SMS",
    requiredHeaders: [
      "STYLE",
      "PO",
      "COLOR",
      "COUNT",
      "Q'TY",
      "NW KG",
      "GW KG",
      "CBM",
      "KÍCH THÙNG"
    ],
  },
  {
    type: "POC BULK",
    requiredHeaders: [
      "ART. NO",
      "PO NO",
      "COLOR",
      "CTNS",
      "TOTAL (PCS)",
      "TTL N.W.",
      "TTL G.W.",
      "CBM",
      "DIMS"
    ],
  },
  {
    type: "REISS",
    requiredHeaders: [
      "STYLE NO & NAME",
      "ORDER NO",
      "COLOUR",
      "ITEM",
      "CTNS",
      "TOTAL",
      "N.W.(KG)",
      "G.W.(KG)",
      "DIMENSION(CM)"
    ],
  },
  {
    type: "MAMMUT",
    requiredHeaders: [
      "DETAILS AS PER MAMMUT ORDER:",
      "COL.+CODE",
      "CARTON TOTAL",
      "TOTAL",
      "TOTAL NW",
      "TOTAL GW",
      "TOTAL CBM",
      "CARTON",
    ],
  },
  {
    type: "MAMMUT GOP PO",
    requiredHeaders: [
      "DETAILS AS PER MAMMUT ORDER:",
      "CARTON TOTAL",
      "TOTAL",
      "TOTAL NW",
      "TOTAL GW",
      "TOTAL CBM",
      "CARTON",
    ],
  },
  {
    type: "ROSSIGNOL BULK",
    requiredHeaders: [
      "PRODUCT CODE",
      "CUSTOMER PO#",
      "NB+COLOR",
      "NB",
      "TOTAL",
      "NET",
      "GROSS",
      "VOLUME (CBM)",
      "CARTON DIMENSION (CM)"
    ],
  },
  {
    type: "ROSSIGNOL SMS",
    requiredHeaders: [
      "PRODUCT CODE SMS",
      "PO",
      "COLOR",
      "COUNT",
      "TOTAL QTY",
      "GW (KGS)",
      "CBM",
      "CARTON DIMENSION (CM)"
    ],
  },
];

export function detectExcelTypeByCell(inputPath) {
  const workbook = xlsx.readFile(inputPath);

  // ✅ ƯU TIÊN CHECK SHEET1 (KỂ CẢ HIDDEN)
  for (const name of workbook.SheetNames) {
    if (name.trim().toLowerCase() === "sheet1") {
      return "INNOVATION";
    }
  }

  // ✅ CHECK THEO TEMPLATE → CHECK TẤT CẢ SHEET
  for (const template of EXCEL_TEMPLATES) {

    for (const sheetName of workbook.SheetNames) {
      if (!sheetName) continue;

      const sheet = workbook.Sheets[sheetName];
      if (!sheet) continue;

      const allText = [];

      for (const addr in sheet) {
        if (addr.startsWith("!")) continue;

        const text = String(sheet[addr]?.v || "")
          .toUpperCase()
          .replace(/\s+/g, " ")
          .trim();

        if (text) allText.push(text);
      }

      const isMatch = template.requiredHeaders.every(header =>
        allText.some(cellText => cellText.includes(header))
      );

      // 🎯 CHỈ CẦN 1 SHEET MATCH
      if (isMatch) {
        return template.type;
      }
    }
  }

  // ❌ KHÔNG SHEET NÀO MATCH
  return "UNKNOWN";
}


// export function detectExcelTypeByCell(inputPath) {
//   const workbook = xlsx.readFile(inputPath);

//   // check Sheet1 trước (kể cả hidden)
//   for (const name of workbook.SheetNames) {
//     if (name.trim().toLowerCase() === "sheet1") {
//       return "INNOVATION";
//     }
//   }

//   for (const template of EXCEL_TEMPLATES) {
//     // chọn sheet index theo template
//     const sheetIndex = template.type === "ACTIVE SMS" ? 2 : 0;
//     const sheetName = workbook.SheetNames[sheetIndex];

//     if (!sheetName) continue;

//     const sheet = workbook.Sheets[sheetName];
//     const allText = [];



//     for (const addr in sheet) {
//       if (addr.startsWith("!")) continue;

//       const text = String(sheet[addr]?.v || "")
//         .toUpperCase()
//         .replace(/\s+/g, " ")
//         .trim();

//       if (text) allText.push(text);
//     }

//     const isMatch = template.requiredHeaders.every(header =>
//       allText.some(cellText => cellText.includes(header))
//     );

//     if (isMatch) {
//       return template.type;
//     }
//   }

//   return "UNKNOWN";
// }

const HAGLOFS_template_1 = [
  "STYLE #",
  "PO",
  "COLOR CODE",
  "NUMBER CARTON",
  "Q'TY (PCS)",
  "NET WEIGHT (KG)",
  "GROSS WEIGHT (KG)",
  "CBM",
  "CARTON DIMENSION(L,W,H)",
];
function previewHAGLOFS_tempplate_1(workbook) {
  const result = {
    sheetNames: [],
    data: {},
    summary: [],
  };

  workbook.SheetNames.forEach((sheetName) => {


    const sheet = workbook.Sheets[sheetName];
    // 🛑 GUARD: nếu sheet KHÔNG thỏa ROSSIGNOL BULK → BỎ QUA
    if (!isSheetMatchTemplate(sheet, HAGLOFS_template_1)) {
      console.log(`⏭️ Skip sheet: ${sheetName}`);
      return;
    }

    const cellMatrix = []; //  mỗi sheet 1 matrix
    const rawRows = xlsx.utils.sheet_to_json(sheet, { defval: "" });

    // 1️ lọc cột
    const columnCleaned = filterInvalidColumns(rawRows);

    // 2️ tìm cột COUNT
    const countColumnKey = findColumnKeyByTitle(columnCleaned, "count");

    // 3️ lọc row COUNT = 0
    const rowCleaned = columnCleaned.filter((row) => {
      // giữ row hợp lệ theo COUNT
      const validByCount = isValidColumn([row], countColumnKey)
        ? true
        : !isZeroInCountColumn(row, countColumnKey);

      if (!validByCount) return false;

      // loại row có ≥ 7 ô trống
      if (hasAtLeastNEmptyCells(row, 8)) return false;

      return true;
    });

    //  GIỮ NGUYÊN DATA CHO FRONTEND
    result.sheetNames.push(sheetName);
    result.data[sheetName] = rowCleaned.map((row) => ({
      SHEET: sheetName,
      ...row,
    }));

    // 4️ build cellMatrix cho sheet này
    const columnCells = extractColumnCellArrays(rowCleaned, sheetName);
    cellMatrix.push(...columnCells);

    // ===================== TÍNH TOÁN =====================

    // sheet name
    console.log("sheet name: ", sheetName)

    // style name
    const getStyles = getValuesBelowTitleAuto(cellMatrix, "Style #");
    const styleNameArr = filterOnlyNumericValues(getStyles);

    console.log("style: ", styleNameArr)

    // color
    const getColors = getValuesBelowTitleAuto(cellMatrix, "Color code");
    console.log("COLOR : ", getColors);

    // po
    const getPo = getValuesBelowTitleAuto(cellMatrix, "PO");
    const poArr = filterOnlyNumericValues(getPo);
    console.log("PO: ", poArr);

    // kiểm duyệt style name, color và po
    const styleNameCheck = collapseIfSame(styleNameArr); // string or arr
    const colorCheck = collapseIfSame(getColors); // string or arr
    const poCheck = collapseIfSame(poArr); // string or arr

    let styleName = "";
    let color = "";
    let po = "";

    // nếu cả 3 ko phải arr thì tức là string, và chỉ cần ghi 1 dòng, nếu 1 trong 3 vừa có arr hoặc string thì phải trả cả 3 về mảng, ko collapse
    // 
    if (!isArray(styleNameCheck) && !isArray(colorCheck) && !isArray(poCheck)) {
      styleName = styleNameCheck;
      color = colorCheck;
      po = poCheck;
    } else {
      styleName = styleNameArr;
      color = getColors;
      po = poArr;
    }

    console.log("--------- --------")
    console.log("style: ", styleName);
    console.log("color: ", color);
    console.log("po: ", po)

    // count (carton no)
    const getCount = getValuesBelowTitleAuto(cellMatrix, "number carton");
    const countArr = normalizeCartonArray(getCount);
    console.log("count arr: ", countArr);
    const count = sumArray(countArr);
    console.log("count: ", count);


    // total qty
    let total = "";
    const getTotal = getValuesBelowTitleAuto(cellMatrix, "Q'ty (pcs)");
    const totalArr = removeAfterString(getTotal);

    // xử lý total 
    const getSeason = getValuesBelowTitleAuto(cellMatrix, "Season");

    const checkSeason = arrayContainsString(getSeason, "TOTAL");
    if (checkSeason === true) {
      const seasonCleaned = removeItemFromArray(getSeason, "Season");
      const seasonAndTotal = mergeToObjectArrayLevel2(seasonCleaned, totalArr);
      console.log("season and total trước khi cleaned: ", seasonAndTotal);
      const seasonAndTotalCleaned = removeItemsByMeasString(seasonAndTotal, "TOTAL");
      console.log("season and total sau khi cleaned", seasonAndTotalCleaned);
      const totalArrCleaned = extractItemArray(seasonAndTotalCleaned);
      total = sumArray(totalArrCleaned);
      console.log("total: ", total)
    }
    else {
      total = smartCountSum(totalArr);
    }

    // gross
    const grossArr = getValuesBelowTitleAuto(cellMatrix, "Gross Weight (kg)");
    const grossArrCleaned = removeAfterString(grossArr);
    const gross = smartSumFloat(grossArrCleaned);

    //net
    const netArr = getValuesBelowTitleAuto(cellMatrix, "Net Weight (kg)");
    const netCleaned = removeAfterString(netArr);
    const net = smartSumFloat(netCleaned);
    console.log("net: ", net);

    // volumn 
    const volumeArr = getValuesBelowTitleAutoVoLumn_n_CBM(cellMatrix, "CBM");
    const volume = sumArray(volumeArr);
    console.log("total volume: ", volume);


    // carton dimension
    const getCartonDimension = getValuesBelowTitleAuto(cellMatrix, "Carton Dimension(L,W,H)");
    console.log("carton dimension: ", getCartonDimension);
    const cartonDimesionArr = keepOnlyDimensionLxWxH(getCartonDimension);
    const cartonAndCount = mergeToObjectArray(cartonDimesionArr, countArr);

    const cartonDimension = buildMeasFormula(cartonAndCount);
    console.log('cartion dimension: ', cartonDimension);

    const tmp = calculateFromFormula(cartonDimension);
    console.log("TMP: ", tmp);

    console.log("==============================================");

    result.summary.push({
      sheetName,
      article: styleName,
      po: po,
      color: color,
      countTotal: count,
      total: total,
      gross: gross,
      net: net,
      volume: volume,
      cartonFormula: cartonDimension,
    });

  });
  console.log(result)

  return result;
}

const HAGLOFS_template_2 = [
  "STYLE NUMBER",
  "PO-NUMBER",
  "COLOR CODE",
  "TOTAL CTN",
  "TOTAL QUANTITY",
  "NET WEIGHT KG",
  "GROSS WEIGHT KG",
  "CARTON TYPE",
];

function previewHAGLOFS_template_2(workbook) {
  const result = {
    sheetNames: [],
    data: {},
    summary: [],
  };

  workbook.SheetNames.forEach((sheetName) => {

    const sheet = workbook.Sheets[sheetName];

    // 🛑 GUARD: nếu sheet KHÔNG thỏa ROSSIGNOL BULK → BỎ QUA
    if (!isSheetMatchTemplate(sheet, HAGLOFS_template_2)) {
      console.log(`⏭️ Skip sheet: ${sheetName}`);
      return;
    }
    const cellMatrix = []; //  mỗi sheet 1 matrix
    const rawRows = xlsx.utils.sheet_to_json(sheet, { defval: "" });

    // 1️ lọc cột
    const columnCleaned = filterInvalidColumns(rawRows);

    // 2️ tìm cột COUNT
    const countColumnKey = findColumnKeyByTitle(columnCleaned, "total ctn");


    // 3️ lọc row COUNT = 0
    const rowCleaned = columnCleaned.filter((row) => {
      if (isValidColumn([row], countColumnKey)) return true;
      return !isZeroInCountColumn(row, countColumnKey);
    });

    //  GIỮ NGUYÊN DATA CHO FRONTEND
    result.sheetNames.push(sheetName);
    result.data[sheetName] = rowCleaned.map((row) => ({
      SHEET: sheetName,
      ...row,
    }));

    // 4️ build cellMatrix cho sheet này
    const columnCells = extractColumnCellArrays(rowCleaned, sheetName);
    cellMatrix.push(...columnCells);

    // ===================== TÍNH TOÁN =====================

    // sheet name
    console.log("sheet name: ", sheetName)

    // style name
    const getStyleNumber = getValuesBelowTitleAuto(cellMatrix, "Style Number");
    const styleNumberArr = removeAfterStringValue(getStyleNumber, "Style Number");

    console.log("style: ", styleNumberArr);

    // color
    const getColors = getValuesBelowTitleAuto(cellMatrix, "Color code");

    const colorArr = removeAfterStringValue(getColors, "Color code")


    // po
    const getPo = getValuesBelowTitleAuto(cellMatrix, "PO-Number");
    const poArr = filterOnlyNumericValues(getPo);
    console.log("PO: ", poArr);

    // kiểm duyệt style name, color và po
    const styleNameCheck = collapseIfSame(styleNumberArr); // string or arr
    const colorCheck = collapseIfSame(colorArr); // string or arr
    const poCheck = collapseIfSame(poArr); // string or arr

    let styleName = "";
    let color = "";
    let po = "";

    // nếu cả 3 ko phải arr thì tức là string, và chỉ cần ghi 1 dòng, nếu 1 trong 3 vừa có arr hoặc string thì phải trả cả 3 về mảng, ko collapse
    // 
    if (!isArray(styleNameCheck) && !isArray(colorCheck) && !isArray(poCheck)) {
      styleName = styleNameCheck;
      color = colorCheck;
      po = poCheck;
    } else {
      styleName = styleNumberArr;
      color = colorArr;
      po = poArr;
    }

    console.log("--------- --------")
    console.log("style: ", styleName);
    console.log("color: ", color);
    console.log("po: ", po)


    // count (carton no)
    const getCount = getValuesBelowTitleAutoVoLumn_n_CBM(cellMatrix, "Total Ctn");
    const countFiltered = filterOnlyNumericValues(getCount);
    const countArr = countFiltered;
    // console.log('count arr: ', countArr);

    // tmp (xử lý count arr)
    const getCartonNumber = getValuesBelowTitleAuto(cellMatrix, "Carton Number");
    console.log("get carton number: ", getCartonNumber);
    const cartonNumberCleaned = removeItemFromArray(getCartonNumber, "Carton Number");
    const cartionNumberAndCount = mergeToObjectArray(cartonNumberCleaned, countArr);
    console.log("trước khi cleaned: ", cartionNumberAndCount);
    const cartionNumberAndCountCleaned = removeItemsByMeasString(cartionNumberAndCount, "TOTAL CARTONS");
    console.log("carton number and count: ", cartionNumberAndCountCleaned);
    const countArrCleaned = extractCountArray(cartionNumberAndCountCleaned);
    // console.log("count arr cleaned: ",countArrCleaned);

    const count = sumArray(countArrCleaned);
    console.log("count: ", count);

    // total qty
    const getTotal = getValuesBelowTitleAuto(cellMatrix, "Total Quantity");
    const totalFiltered = removeAfterString(getTotal);
    const totalArr = filterOnlyNumericValues(totalFiltered);

    // xử lý total arr
    const cartonNumberAndTotal = mergeToObjectArrayLevel2(cartonNumberCleaned, totalArr);
    const cartonNumberAndTotalCleaned = removeItemsByMeasString(cartonNumberAndTotal, "TOTAL CARTONS");
    console.log("carton and total: ", cartonNumberAndTotalCleaned);
    const totalArrCleaned = extractItemArray(cartonNumberAndTotalCleaned);
    const total = sumArray(totalArrCleaned);
    console.log("total: ", total);



    // gross
    const grossArr = getValuesBelowTitleAuto(cellMatrix, "Gross Weight kg");
    const cartonNumberAndGross = mergeToObjectArrayLevel2(cartonNumberCleaned, grossArr);
    const cartonAndGrossCleaned = removeItemsByMeasString(cartonNumberAndGross, "TOTAL CARTONS");
    const grossArrCleaned = extractItemArray(cartonAndGrossCleaned);


    const gross = sumArray(grossArrCleaned);
    console.log("gross: ", gross);


    // net
    const netArr = getValuesBelowTitleAuto(cellMatrix, "Net Weight kg");
    const cartonNumberAndNet = mergeToObjectArrayLevel2(cartonNumberCleaned, netArr);
    const cartonNumberAndNetCleaned = removeItemsByMeasString(cartonNumberAndNet, "TOTAL CARTONS");
    const netArrCleaned = extractItemArray(cartonNumberAndNetCleaned);
    const net = sumArray(netArrCleaned);
    console.log("net: ", net);

    // carton dimension
    const getCartonDimension = getValuesBelowTitleAuto(cellMatrix, "Carton  Type");
    const cartonDimesionArr = keepOnlyDimensionLxWxH(getCartonDimension);

    const cartonAndCount = mergeToObjectArray(cartonDimesionArr, countArrCleaned);

    const cartonDimension = buildMeasFormula(cartonAndCount);
    console.log('cartion dimension: ', cartonDimension);

    const volume = calculateFromFormula(cartonDimension);
    console.log("volume: ", volume);

    console.log("==============================================");

    result.summary.push({
      sheetName,
      article: styleName,
      po: po,
      color: color,
      countTotal: count,
      total: total,
      gross: gross,
      net: net,
      volume: volume,
      cartonFormula: cartonDimension,
    });

  });
  console.log(result)

  return result;
}

// active sms
const Active_SMS = [
  "STYLE NO",
  "PACKING DESTINATION",
  "COLOUR",
  "CARTON NO",
  "QTY",
  "N.W./KGS",
  "G.W./KGS",
  "VOLUMN (CBM)",
];
function previewActive_SMS(workbook) {
  const result = {
    sheetNames: [],
    data: {},
    summary: [],
  };

  workbook.SheetNames.forEach((sheetName, index) => {
    if (index !== 2) return;


    const sheet = workbook.Sheets[sheetName];
    // 🛑 GUARD: nếu sheet KHÔNG thỏa ROSSIGNOL BULK → BỎ QUA
    if (!isSheetMatchTemplate(sheet, Active_SMS)) {
      console.log(`⏭️ Skip sheet: ${sheetName}`);
      return;
    }
    const cellMatrix = []; //  mỗi sheet 1 matrix

    const rawRows = xlsx.utils.sheet_to_json(sheet, { defval: "" });

    // 1️ lọc cột
    const columnCleaned = filterInvalidColumns(rawRows);

    // 2️ tìm cột COUNT
    const countColumnKey = findColumnKeyByTitle(columnCleaned, "count");

    // 3️ lọc row COUNT = 0
    const rowCleaned = columnCleaned.filter((row) => {
      // giữ row hợp lệ theo COUNT
      const validByCount = isValidColumn([row], countColumnKey)
        ? true
        : !isZeroInCountColumn(row, countColumnKey);

      if (!validByCount) return false;

      // loại row có ≥ 7 ô trống
      if (hasAtLeastNEmptyCells(row, 8)) return false;

      return true;
    });

    //  GIỮ NGUYÊN DATA CHO FRONTEND
    result.sheetNames.push(sheetName);
    result.data[sheetName] = rowCleaned.map((row) => ({
      SHEET: sheetName,
      ...row,
    }));


    // 4️ build cellMatrix cho sheet này
    const columnCells = extractColumnCellArrays(rowCleaned, sheetName);
    cellMatrix.push(...columnCells);

    // ===================== TÍNH TOÁN =====================

    // gross
    const getGross = getValuesBelowTitleAuto(cellMatrix, "G.W./KGS");
    const grossArr = filterOnlyNumericValues(getGross);
    // console.log("gross arr: ", grossArr);
    const gross = sumArray(grossArr.slice(0, grossArr.length - 1));
    console.log("gross: ", gross)
    //net
    const getNet = getValuesBelowTitleAuto(cellMatrix, "N.W./KGS");
    const netArr = filterOnlyNumericValues(getNet);
    const net = sumArray(netArr.slice(0, netArr.length - 1));
    console.log("net: ", net);

    // article / style
    const getStyleNameArr = getValuesBelowTitleAuto(cellMatrix, "Style No");
    console.log('style arr: ', getStyleNameArr)

    // color
    const getColourArr = getValuesBelowTitleAuto(cellMatrix, "Colour");
    console.log("color: ", getColourArr);

    // PO (Packing Destination)
    const getPackingDestinationArr = getValuesBelowTitleAuto(cellMatrix, "Packing Destination");
    console.log("Packing Destination: ", getPackingDestinationArr);

    // count
    const getCartonNoArr = getValuesBelowTitleAuto(cellMatrix, "Carton No");
    const cartonNoCleaned = removeItemFromArray(getCartonNoArr, "Carton No");
    const count = sumArray(cartonNoCleaned.slice(0, cartonNoCleaned.length - 1));
    console.log("count arr: ", cartonNoCleaned);
    console.log("count: ", count);

    // kiểm duyệt style name, color và po
    const styleNameCheck = collapseIfSame(getStyleNameArr); // string or arr
    const colorCheck = collapseIfSame(getColourArr); // string or arr
    const poCheck = collapseIfSame(getPackingDestinationArr); // string or arr

    let styleName = "";
    let color = "";
    let po = "";

    // nếu cả 3 ko phải arr thì tức là string, và chỉ cần ghi 1 dòng, nếu 1 trong 3 vừa có arr hoặc string thì phải trả cả 3 về mảng, ko collapse
    // 
    if (!isArray(styleNameCheck) && !isArray(colorCheck) && !isArray(poCheck)) {
      styleName = styleNameCheck;
      color = colorCheck;
      po = poCheck;
    } else {
      styleName = getStyleNameArr;
      color = getColourArr;
      po = getPackingDestinationArr;
    }

    // console.log("--------- --------")
    // console.log("style: ", styleName);
    // console.log("color: ", color);
    // console.log("po: ", po)


    // total qty
    const totalQty = getValuesBelowTitleAuto(cellMatrix, "QTY");
    const totalQtyArr = filterOnlyNumericValues(totalQty);
    const totalQtySum = sumArray(totalQtyArr.slice(0, totalQtyArr.length - 1));
    console.log("total: ", totalQtySum);

    // volumn 
    const getVolume = getValuesBelowTitleAuto(cellMatrix, "Volumn (CBM)");
    const volumeArr = filterOnlyNumericValues(getVolume);
    const totalVolumn = sumArray(volumeArr.slice(0, volumeArr.length - 1));
    console.log("volume (CBM): ", totalVolumn);

    // carton dimension


    result.summary.push({
      sheetName,
      article: styleName,
      po: po,
      color: color,
      countTotal: count,
      total: totalQtySum,
      gross: gross,
      net: net,
      volume: totalVolumn,
      cartonFormula: "", // chưa có
    });
  });

  console.log(result.summary);

  return result;
}

// active bulk
const Active_BULK = [
  "STYLE NO",
  "PURCHASE ORDER NO:",
  "COLOR",
  "QTY OF CTNS",
  "TOTAL PCS",
  "NET WEIGHT PR CTN",
  "GROSS WEIGHT PR CTN",
  "CBM PR CTN",
];
function previewActive_BULK(workbook) {
  const result = {
    sheetNames: [],
    data: {},
    summary: [],
  };

  workbook.SheetNames.forEach((sheetName, index) => {

    let purchase_order_no = ''



    const sheet = workbook.Sheets[sheetName];
    // 🛑 GUARD: nếu sheet KHÔNG thỏa ROSSIGNOL BULK → BỎ QUA
    if (!isSheetMatchTemplate(sheet, Active_BULK)) {
      console.log(`⏭️ Skip sheet: ${sheetName}`);
      return;
    }
    const cellMatrix = []; //  mỗi sheet 1 matrix

    // PO
    purchase_order_no = extractPurchaseOrderNo(sheet, "PURCHASE ORDER NO");

    console.log("Purchase Order No:", purchase_order_no);

    const rawRows = xlsx.utils.sheet_to_json(sheet, { defval: "" });

    // 1️ lọc cột
    const columnCleaned = filterInvalidColumns(rawRows);

    // 2️ tìm cột COUNT
    const countColumnKey = findColumnKeyByTitle(columnCleaned, "count");

    // 3️ lọc row COUNT = 0
    const rowCleaned = columnCleaned.filter((row) => {
      // giữ row hợp lệ theo COUNT
      const validByCount = isValidColumn([row], countColumnKey)
        ? true
        : !isZeroInCountColumn(row, countColumnKey);

      if (!validByCount) return false;

      // loại row có ≥ 7 ô trống
      if (hasAtLeastNEmptyCells(row, 8)) return false;

      return true;
    });

    //  GIỮ NGUYÊN DATA CHO FRONTEND
    result.sheetNames.push(sheetName);
    result.data[sheetName] = rowCleaned.map((row) => ({
      SHEET: sheetName,
      ...row,
    }));


    // 4️ build cellMatrix cho sheet này
    const columnCells = extractColumnCellArrays(rowCleaned, sheetName);
    cellMatrix.push(...columnCells);

    // ===================== TÍNH TOÁN =====================

    // GROSS WEIGHT PR CTN * QTY OF CTNS

    // QTY OF CTNS
    const getQTY_OF_CTNS = getValuesBelowTitleAuto(cellMatrix, "QTY OF CTNS");
    const QTY_OF_CTNS_ARR = filterOnlyNumericValues(getQTY_OF_CTNS);
    const qty_of_ctns_cleaned = removeItemFromArray(QTY_OF_CTNS_ARR, "QTY OF CTNS");
    // console.log("QTY OF CTNS arr: ", qty_of_ctns_cleaned);
    const count = sumArray(qty_of_ctns_cleaned);
    console.log("count: ", count);

    // gross
    const getGross = getValuesBelowTitleAuto(cellMatrix, "GROSS WEIGHT PR CTN");
    const grossArr = filterOnlyNumericValues(getGross);
    // console.log("gross arr: ", grossArr);
    console.log("gross arr: ", grossArr),
      console.log("qty_of_ctns_cleaned", qty_of_ctns_cleaned);
    const gross = calcSum2Arr(qty_of_ctns_cleaned, grossArr);
    console.log("gross: ", gross);

    //net
    const getNet = getValuesBelowTitleAuto(cellMatrix, "NET WEIGHT PR CTN");
    const netArr = filterOnlyNumericValues(getNet);
    const net = calcSum2Arr(qty_of_ctns_cleaned, netArr);
    console.log("net: ", net);

    // article / style
    const getStyleNoArr = getValuesBelowTitleAuto(cellMatrix, "STYLE NO");
    const styleNoArrCleaned = removeBeforeTitle(getStyleNoArr, 'STYLE NO');
    // console.log('style arr: ', styleNoArrCleaned);

    // color
    const getColourArr = getValuesBelowTitleAutoForColorACTIVE_BULK(cellMatrix, "COLOR");
    // console.log("color arr: ", getColourArr);

    // kiểm duyệt style name, color và po
    const styleNameCheck = collapseIfSame(styleNoArrCleaned); // string or arr
    const colorCheck = collapseIfSame(getColourArr); // string or arr

    let styleName = "";
    let color = "";

    // nếu cả 3 ko phải arr thì tức là string, và chỉ cần ghi 1 dòng, nếu 1 trong 3 vừa có arr hoặc string thì phải trả cả 3 về mảng, ko collapse
    // 
    if (!isArray(styleNameCheck) && !isArray(colorCheck)) {
      styleName = styleNameCheck;
      color = colorCheck;
    } else {
      styleName = styleNoArrCleaned;
      color = getColourArr;
    }

    // console.log("--------- --------")
    // console.log("style: ", styleName);
    // console.log("color: ", color);
    // console.log("po: ", po)


    // total qty
    const totalQty = getValuesBelowTitleAutoForTotalActiveBULK(cellMatrix, "Total pcs");
    const totalQtyArr = filterOnlyNumericValues(totalQty);
    const totalQtySum = sumArray(totalQtyArr);
    // console.log("total arr: ",totalQtyArr);
    console.log("total: ", totalQtySum);

    // volumn 

    // get CBM PR CTN
    const getCBM_PR_CTN = getValuesBelowTitleAuto(cellMatrix, "CBM PR CTN");
    const cbm_pr_ctn_cleaned = filterOnlyNumericValues(getCBM_PR_CTN);
    // console.log("CBM PR CTN", getCBM_PR_CTN);
    // qty_of_ctns_cleaned
    const volume = calcSum2Arr(qty_of_ctns_cleaned, cbm_pr_ctn_cleaned);
    console.log("volume: ", volume);

    // carton dimension
    const getMEAS_PR_CTN = getValuesBelowTitleAuto(cellMatrix, "MEAS PR CTN");
    const meas_pr_ctn_cleaned = removeItemFromArray(getMEAS_PR_CTN, "MEAS PR CTN");
    // console.log("count: ", qty_of_ctns_cleaned);
    // console.log("carton: ", getMEAS_PR_CTN);
    const MeasPrAndCount = mergeToObjectArray(meas_pr_ctn_cleaned, qty_of_ctns_cleaned);
    // console.log("arr 2: ", MeasPrAndCount);
    const cartonDimension = buildMeasFormula(MeasPrAndCount);
    console.log("carton dimension: ", cartonDimension);

    result.summary.push({
      sheetName,
      article: styleName,
      po: purchase_order_no,
      color: color,
      countTotal: count,
      total: totalQtySum,
      gross: gross,
      net: net,
      volume: volume,
      cartonFormula: cartonDimension,
    });
  });

  console.log(result.summary);

  return result;
}

// barbour bulk
const BARBOUR_BULK = [
  "STYLE",
  "PO",
  "COLOR",
  "CARTON NO",
  "TOTAL",
  "TOTAL NW",
  "TOTAL GW",
  "TOTAL CBM",
  "SIZE(CM)",
];
function previewBARBOUR_BULK(workbook) {
  const result = {
    sheetNames: [],
    data: {},
    summary: [],
  };

  workbook.SheetNames.forEach((sheetName) => {


    const sheet = workbook.Sheets[sheetName];

    // 🛑 GUARD: nếu sheet KHÔNG thỏa ROSSIGNOL BULK → BỎ QUA
    if (!isSheetMatchTemplate(sheet, BARBOUR_BULK)) {
      console.log(`⏭️ Skip sheet: ${sheetName}`);
      return;
    }

    const cellMatrix = []; //  mỗi sheet 1 matrix

    const rawRows = xlsx.utils.sheet_to_json(sheet, { defval: "" });

    // 1️ lọc cột
    const columnCleaned = filterInvalidColumns(rawRows);

    // 2️ tìm cột COUNT
    const countColumnKey = findColumnKeyByTitle(columnCleaned, "carton no");

    // 3️ lọc row COUNT = 0
    const rowCleaned = columnCleaned.filter((row) => {
      // giữ row hợp lệ theo COUNT
      const validByCount = isValidColumn([row], countColumnKey)
        ? true
        : !isZeroInCountColumn(row, countColumnKey);

      if (!validByCount) return false;

      // loại row có ≥ 7 ô trống
      if (hasAtLeastNEmptyCells(row, 7)) return false;

      return true;
    });

    //  GIỮ NGUYÊN DATA CHO FRONTEND
    result.sheetNames.push(sheetName);
    result.data[sheetName] = rowCleaned.map((row) => ({
      SHEET: sheetName,
      ...row,
    }));

    // 4️ build cellMatrix cho sheet này
    const columnCells = extractColumnCellArrays(rowCleaned, sheetName);
    cellMatrix.push(...columnCells);

    // ===================== TÍNH TOÁN =====================

    // sheet name
    console.log("sheet name: ", sheetName)
    // gross
    const grossArr = getValuesBelowTitleAuto(cellMatrix, "TOTAL GW");
    const grossArrCleaned = filterOnlyNumericValues(grossArr);
    const gross = sumArray(grossArrCleaned);
    console.log("gross: ", gross)

    //net
    const netArr = getValuesBelowTitleAuto(cellMatrix, "TOTAL NW");
    const netArrCleaned = filterOnlyNumericValues(netArr);
    const net = sumArray(netArrCleaned);
    console.log("net: ", net)
    // style name
    const styleNameArr = getValuesBelowTitleAuto(cellMatrix, "STYLE");
    console.log("style: ", styleNameArr);

    // color
    const getColors = getValuesBelowTitleAuto(cellMatrix, "COLOR");
    console.log("COLOR : ", getColors);

    // po
    const getPo = getValuesBelowTitleAuto(cellMatrix, "PO");
    const poArr = filterOnlyNumericValues(getPo);
    console.log("PO: ", poArr);

    // kiểm duyệt style name, color và po
    const styleNameCheck = collapseIfSame(styleNameArr); // string or arr
    const colorCheck = collapseIfSame(getColors); // string or arr
    const poCheck = collapseIfSame(poArr); // string or arr

    let styleName = "";
    let color = "";
    let po = "";

    // nếu cả 3 ko phải arr thì tức là string, và chỉ cần ghi 1 dòng, nếu 1 trong 3 vừa có arr hoặc string thì phải trả cả 3 về mảng, ko collapse
    // 
    if (!isArray(styleNameCheck) && !isArray(colorCheck) && !isArray(poCheck)) {
      styleName = styleNameCheck;
      color = colorCheck;
      po = poCheck;
    } else {
      styleName = styleNameArr;
      color = getColors;
      po = poArr;
    }

    console.log("--------- --------")
    console.log("style: ", styleName);
    console.log("color: ", color);
    console.log("po: ", po)


    // total qty
    const getTotal = getValuesBelowTitleAuto(cellMatrix, "Total");
    const totalArr = removeAfterString(getTotal);
    const totalArrCleaned = filterOnlyNumericValues(totalArr);
    const total = sumArray(totalArrCleaned);
    console.log("total: ", total)

    // count (carton no)
    const getCount = getValuesBelowTitleAutoVoLumn_n_CBM(cellMatrix, "carton no");
    const countArr = filterOnlyNumericValues(getCount);
    const count = sumArray(countArr);
    console.log("count: ", count);

    // volumn 
    const getVolumnArr = getValuesBelowTitleAutoVoLumn_n_CBM(cellMatrix, "TOTAL CBM");
    const volumeArr = filterOnlyNumericValues(getVolumnArr);
    const volume = sumArray(volumeArr);
    console.log("total volume: ", volume)


    // carton dimension
    const getSizeCBMS = getValuesBelowTitleAutoForSizeCBM(cellMatrix, "Size(cm)");
    // console.log("size arr: ", getSizeCBMS);
    const buildCartonDimension = buildCartonDimensionFormulas(getSizeCBMS);
    // console.log("size cbm: ", buildCartonDimension);

    const cartonAndCount = mergeToObjectArray(buildCartonDimension, countArr);
    console.log(cartonAndCount)
    const cartonDimension = buildMeasFormula(cartonAndCount);
    console.log(cartonDimension);


    result.summary.push({
      sheetName,
      article: styleName,
      po,
      color,
      countTotal: count,
      total,
      gross,
      net,
      volume: volume,
      cartonFormula: cartonDimension,
    });

  });

  console.log(result)

  return result;
}
// barbour sms 
const BARBOUR_SMS = [
  "STYLE",
  "PO",
  "COLOR",
  "CARTON NO",
  "TOTAL",
  "TOTAL NW",
  "TOTAL GW",
  "TOTAL CBM",
  "SIZE(CM)",
];
function previewBARBOUR_SMS(workbook) {
  const result = {
    sheetNames: [],
    data: {},
    summary: [],
  };

  workbook.SheetNames.forEach((sheetName) => {
    const sheet = workbook.Sheets[sheetName];

    // 🛑 GUARD: nếu sheet KHÔNG thỏa ROSSIGNOL BULK → BỎ QUA
    if (!isSheetMatchTemplate(sheet, BARBOUR_SMS)) {
      console.log(`⏭️ Skip sheet: ${sheetName}`);
      return;
    }

    const cellMatrix = []; //  mỗi sheet 1 matrix

    const rawRows = xlsx.utils.sheet_to_json(sheet, { defval: "" });

    // 1️ lọc cột
    const columnCleaned = filterInvalidColumns(rawRows);

    // 2️ tìm cột COUNT
    const countColumnKey = findColumnKeyByTitle(columnCleaned, "carton no");

    // 3️ lọc row COUNT = 0
    const rowCleaned = columnCleaned.filter((row) => {
      // giữ row hợp lệ theo COUNT
      const validByCount = isValidColumn([row], countColumnKey)
        ? true
        : !isZeroInCountColumn(row, countColumnKey);

      if (!validByCount) return false;

      // loại row có ≥ 7 ô trống
      if (hasAtLeastNEmptyCells(row, 7)) return false;

      return true;
    });

    //  GIỮ NGUYÊN DATA CHO FRONTEND
    result.sheetNames.push(sheetName);
    result.data[sheetName] = rowCleaned.map((row) => ({
      SHEET: sheetName,
      ...row,
    }));

    // 4️ build cellMatrix cho sheet này
    const columnCells = extractColumnCellArrays(rowCleaned, sheetName);
    cellMatrix.push(...columnCells);

    // ===================== TÍNH TOÁN =====================

    // sheet name
    console.log("sheet name: ", sheetName)
    // gross
    const grossArr = getValuesBelowTitleAuto(cellMatrix, "TOTAL GW");
    const grossArrCleaned = filterOnlyNumericValues(grossArr);
    const gross = sumArray(grossArrCleaned);
    console.log("gross: ", gross)

    //net
    const netArr = getValuesBelowTitleAuto(cellMatrix, "TOTAL NW");
    const netArrCleaned = filterOnlyNumericValues(netArr);
    const net = sumArray(netArrCleaned);
    console.log("net: ", net)
    // style name
    const styleNameArr = getValuesBelowTitleAuto(cellMatrix, "STYLE");
    console.log("style: ", styleNameArr);

    // color
    const getColors = getValuesBelowTitleAuto(cellMatrix, "COLOR");
    console.log("COLOR : ", getColors);

    // po
    const getPo = getValuesBelowTitleAuto(cellMatrix, "PO");
    const poArr = filterOnlyNumericValues(getPo);
    console.log("PO: ", poArr);

    // kiểm duyệt style name, color và po
    const styleNameCheck = collapseIfSame(styleNameArr); // string or arr
    const colorCheck = collapseIfSame(getColors); // string or arr
    const poCheck = collapseIfSame(poArr); // string or arr

    let styleName = "";
    let color = "";
    let po = "";

    // nếu cả 3 ko phải arr thì tức là string, và chỉ cần ghi 1 dòng, nếu 1 trong 3 vừa có arr hoặc string thì phải trả cả 3 về mảng, ko collapse
    // 
    if (!isArray(styleNameCheck) && !isArray(colorCheck) && !isArray(poCheck)) {
      styleName = styleNameCheck;
      color = colorCheck;
      po = poCheck;
    } else {
      styleName = styleNameArr;
      color = getColors;
      po = poArr;
    }

    console.log("--------- --------")
    console.log("style: ", styleName);
    console.log("color: ", color);
    console.log("po: ", po)


    // total qty
    const getTotal = getValuesBelowTitleAuto(cellMatrix, "Total");
    const totalArr = removeAfterString(getTotal);
    const totalArrCleaned = filterOnlyNumericValues(totalArr);
    const total = sumArray(totalArrCleaned);
    console.log("total: ", total)

    // count (carton no)
    const getCount = getValuesBelowTitleAutoVoLumn_n_CBM(cellMatrix, "carton no");
    const countArr = filterOnlyNumericValues(getCount);
    const count = sumArray(countArr);
    console.log("count: ", count);

    // volumn 
    const getVolumnArr = getValuesBelowTitleAutoVoLumn_n_CBM(cellMatrix, "TOTAL CBM");
    const volumeArr = filterOnlyNumericValues(getVolumnArr);
    const volume = sumArray(volumeArr);
    console.log("total volume: ", volume)


    // carton dimension
    const getSizeCBMS = getValuesBelowTitleAutoForSizeCBM(cellMatrix, "Size(cm)");
    // console.log("size arr: ", getSizeCBMS);
    const buildCartonDimension = buildCartonDimensionFormulas(getSizeCBMS);
    // console.log("size cbm: ", buildCartonDimension);

    const cartonAndCount = mergeToObjectArray(buildCartonDimension, countArr);
    console.log(cartonAndCount)
    const cartonDimension = buildMeasFormula(cartonAndCount);
    console.log(cartonDimension);


    result.summary.push({
      sheetName,
      article: styleName,
      po,
      color,
      countTotal: count,
      total,
      gross,
      net,
      volume: volume,
      cartonFormula: cartonDimension,
    });

  });

  console.log(result)

  return result;
}

// asics bulk
function detectPaletValue(row) {
  const text = normalize_(Object.values(row).join(" "));
  const match = text.match(/pallet\s*so\s*(\d+)/) || text.match(/palet\s*so\s*(\d+)/);
  return match ? Number(match[1]) : 0;
}


function normalize_(str = "") {
  return String(str)
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

const ASICS_BULK = [
  "ARTICLE / STYLE",
  "PO",
  "COLOR",
  "COUNT",
  "TOTAL",
  "NET",
  "GROSS",
  "VOLUME (CBM)",
  "CARTON DIMENSION (CM)",
];
function previewASICS_BULK(workbook) {
  const result = {
    sheetNames: [],
    data: {},
    summary: [],
  };

  // phần này xử lý cho ONE IBIS 021W:
  workbook.SheetNames.forEach((sheetName) => {

    const sheet = workbook.Sheets[sheetName];
    // 🛑 GUARD: nếu sheet KHÔNG thỏa ROSSIGNOL BULK → BỎ QUA
    if (!isSheetMatchTemplate(sheet, ASICS_BULK)) {
      console.log(`⏭️ Skip sheet: ${sheetName}`);
      return;
    }
    const cellMatrix = []; //  mỗi sheet 1 matrix

    // PO
    const purchase_order_no = extractPurchaseOrderNo(sheet, "ORDER NO:");
    console.log("Order no:", purchase_order_no);

    const rawRows = xlsx.utils.sheet_to_json(sheet, { defval: "" });

    // 1️ lọc cột
    const columnCleaned = filterInvalidColumns(rawRows);

    // 2️ tìm cột COUNT
    const countColumnKey = findColumnKeyByTitle(columnCleaned, "count_");

    // 3️ lọc row COUNT = 0
    // const rowCleaned = columnCleaned.filter((row) => {
    //   if (isValidColumn([row], countColumnKey)) return true;
    //   return !isZeroInCountColumn(row, countColumnKey);
    // });

    // let palet = 0;
    // hãy viết kiểm tra chỗ này nếu có 1 Palet thì palet = palet + 20, cứ có một palet xuất hiện thì cột 20 được không

    const excludedTitles = [
      "grand total",
      "PALET SỐ 1",
      "PALET SỐ 2",
      "PALET SỐ 3",
      "PALET SỐ 4",
      "PALET SỐ 5",
      "PALET SỐ 6",
      "PALET SỐ 7",
      "PALET SỐ 8",
      "PALET SỐ 9",
      "PALET SỐ 10",
      "PALLET SỐ 1",
      "PALLET SỐ 2",
      "PALLET SỐ 3",
      "PALLET SỐ 4",
      "PALLET SỐ 5",
      "PALLET SỐ 6",
      "PALLET SỐ 7",
      "PALLET SỐ 8",
      "PALLET SỐ 9",
      "PALLET SỐ 10",
      "TOTAL:",
      "PACKING",
      "TỔNG",
      "COLOR SUMMARY:",
      "+/-",
    ];

    let palet = 0;


    const rowCleaned = columnCleaned.filter(row => {
      const count = detectPaletValue(row);
      if (count > 0) {
        palet += 20;
        return false;
      }

      if (filterRowsByExcludedTitles([row], excludedTitles).length === 0) {
        return false;
      }

      if (isValidColumn([row], countColumnKey)) return true;
      return !isZeroInCountColumn(row, countColumnKey);
    });


    //  GIỮ NGUYÊN DATA CHO FRONTEND
    result.sheetNames.push(sheetName);
    result.data[sheetName] = rowCleaned.map((row) => ({
      SHEET: sheetName,
      ...row,
    }));


    console.log("palet: ", palet);

    // 4️ build cellMatrix cho sheet này
    const columnCells = extractColumnCellArrays(rowCleaned, sheetName);
    cellMatrix.push(...columnCells);

    // ===================== TÍNH TOÁN =====================

    console.log("------------------------------------------------------");


    // color 
    const getColor = getValuesBelowTitleAuto(cellMatrix, "COLOR");
    const colorArr = removeAfterStringValue(getColor, "COLOR");
    console.log("color arr: ", colorArr);

    // article / style
    const getArticleStyle = getValuesBelowTitleAutoForColorACTIVE_BULK(cellMatrix, "ARTICLE / STYLE");
    const articleStyleArr1 = removeAfterStringValue(getArticleStyle, "TOTAL")
    const articleStyleArr2 = removeAfterStringValue(articleStyleArr1, "ARTICLE / STYLE")
    const articleStyleArr3 = removeAfterStringValue(articleStyleArr2, "PO")
    const articleStyleArr4 = removeAfterStringValue(articleStyleArr3, "COLOR")
    const articleStyleArr5 = removeAfterStringValue(articleStyleArr4, "SHADE ")
    const articleStyleArr = removeAfterStringValue(articleStyleArr5, "total")

    console.log("ARTICLE / STYLE: ", articleStyleArr);


    // kiểm duyệt color, article / style, po
    const styleCheck = collapseIfSame(articleStyleArr);
    const colorCheck = collapseIfSame(colorArr);

    let styleName = "";
    let color = "";

    if (!isArray(styleCheck) || !isArray(colorCheck)) {
      styleName = styleCheck;
      color = colorCheck;
    }
    else {
      styleName = articleStyleArr;
      color = colorArr;
    }

    // COUNT
    const getCount = getValuesBelowTitleAuto(cellMatrix, "COUNT");
    const countArr = filterOnlyNumericValues(getCount);
    const count = sumArray(countArr);
    // console.log("count arr: ", countArrCleaned);
    console.log("count: ", count);

    // TOTAL
    // total ( do QTY bị ẩn, nếu lấy theo TOTAL thì sẽ có thêm value là 'QTY' )
    const getTotals = getValuesBelowTitleExact(cellMatrix, "TOTAL");
    // const totalArrFiltered = filterOnlyNumericValues(getTotals);
    const totalArrCleaned = removeItemFromArray(getTotals, "QTY");
    const totalArr = removeAfterString(totalArrCleaned);
    const totalArrFiltered = filterOnlyNumericValues(totalArr);
    console.log("total arr: ", totalArrFiltered);
    const total = sumArray(totalArrFiltered);
    console.log("total: ", total);

    // GROSS 
    const getGross = getValuesBelowTitleAuto(cellMatrix, "GROSS");
    console.log("gorss arr: ", getGross);
    const grossArr = removeItemFromArray(getGross, "WEIGHT");
    const gross = sumArray(grossArr) + palet;
    console.log("gorss: ", gross);

    // NET
    const getNet = getValuesBelowTitleAuto(cellMatrix, "NET");
    const netArr = removeItemFromArray(getNet, "WEIGHT");
    const net = sumArray(netArr);
    console.log("net: ", net);

    // VOLUME
    const getVolume = getValuesBelowTitleAuto(cellMatrix, "VOLUME (CBM)");
    const volumeArr = getVolume;
    const volume = sumArray(volumeArr);
    console.log("volume: ", volume);

    // CARTON FORMULA
    const getCartonDimension = getValuesBelowTitleAuto(cellMatrix, "CARTON DIMENSION (CM)");
    const cartonDimensionArr = removeItemFromArray(getCartonDimension, "CARTON DIMENSION (CM)")
    console.log("CARTON DIMENSION (CM): ", cartonDimensionArr);
    const cartonAndCount = mergeToObjectArray(cartonDimensionArr, countArr);
    console.log("carton and count: ", cartonAndCount);
    const cartonDimension = buildMeasFormula(cartonAndCount);
    console.log("carton dimension: ", cartonDimension);

    // ===================== GẮN SUMMARY =====================
    result.summary.push({
      sheetName,
      article: styleName,
      po: purchase_order_no,
      color: color,
      countTotal: count,
      total: total,
      gross: gross,
      net: net,
      volume: volume,
      cartonFormula: cartonDimension,
    });

  });
  console.log(result.summary);

  return result;
}

// asics sms
const ASICS_SMS = [
  "ARTICLE / STYLE",
  "COLOR",
  "COUNT",
  "TOTAL",
  "G.W",
  "N.W",
  "CBM",
  "CARTON DIMENSION (CM)",
];
function previewASICS_SMS(workbook) {
  const result = {
    sheetNames: [],
    data: {},
    summary: [],
  };

  // phần này xử lý cho ONE IBIS 021W:
  workbook.SheetNames.forEach((sheetName) => {


    const sheet = workbook.Sheets[sheetName];
    // 🛑 GUARD: nếu sheet KHÔNG thỏa ROSSIGNOL BULK → BỎ QUA
    if (!isSheetMatchTemplate(sheet, ASICS_SMS)) {
      console.log(`⏭️ Skip sheet: ${sheetName}`);
      return;
    }
    const cellMatrix = []; //  mỗi sheet 1 matrix
    // PO
    // const purchase_order_no = extractPurchaseOrderNo(sheet, "ORDER NO:");
    // console.log("Order no:", purchase_order_no);

    const rawRows = xlsx.utils.sheet_to_json(sheet, { defval: "" });

    // 1️ lọc cột
    const columnCleaned = filterInvalidColumns(rawRows);

    // 2️ tìm cột COUNT
    const countColumnKey = findColumnKeyByTitle(columnCleaned, "count_");

    // get L, W, H trước khi cleaned
    // getL

    // 3️ lọc row COUNT = 0
    const rowCleaned = columnCleaned.filter((row) => {
      // loại row có From – To nằm trong cell
      if (hasFromToCells(row)) return false;

      // logic cũ
      if (isValidColumn([row], countColumnKey)) return true;
      return !isZeroInCountColumn(row, countColumnKey);
    });

    //  GIỮ NGUYÊN DATA CHO FRONTEND
    result.sheetNames.push(sheetName);
    result.data[sheetName] = rowCleaned.map((row) => ({
      SHEET: sheetName,
      ...row,
    }));

    // 4️ build cellMatrix cho sheet này
    const columnCells = extractColumnCellArrays(rowCleaned, sheetName);
    cellMatrix.push(...columnCells);


    // console.log(
    //   cellMatrix.map(col =>
    //     col.cells.map(v => normalizeForSizeCBM(v))
    //   )
    // );

    // ===================== TÍNH TOÁN =====================

    console.log("------------------------------------------------------");



    // sheet name
    console.log("sheet name: ", sheetName);

    // color 
    const getColor = getValuesBelowTitleAutoForColorACTIVE_BULK(cellMatrix, "COLOR");
    console.log("color arr: ", getColor);

    // article / style
    const getArticleStyle = getValuesBelowTitleAutoForColorACTIVE_BULK(cellMatrix, "ARTICLE / STYLE");
    const articleStyleArr1 = removeAfterStringValue(getArticleStyle, "TOTAL")
    const articleStyleArr2 = removeAfterStringValue(articleStyleArr1, "ARTICLE / STYLE")
    const articleStyleArr3 = removeAfterStringValue(articleStyleArr2, "PO")
    const articleStyleArr4 = removeAfterStringValue(articleStyleArr3, "COLOR")
    const articleStyleArr5 = removeAfterStringValue(articleStyleArr4, "SHADE ")
    const articleStyleArr = removeAfterStringValue(articleStyleArr5, "total")

    console.log("ARTICLE / STYLE: ", articleStyleArr);


    // kiểm duyệt color, article / style, po
    const styleCheck = collapseIfSame(articleStyleArr);
    const colorCheck = collapseIfSame(getColor);

    let styleName = "";
    let color = "";

    if (!isArray(styleCheck) || !isArray(colorCheck)) {
      styleName = styleCheck;
      color = colorCheck;
    }
    else {
      styleName = articleStyleArr;
      color = getColor;
    }

    // COUNT
    const getCount = getValuesBelowTitleAuto(cellMatrix, "COUNT");
    const countArr = filterOnlyNumericValues(getCount);
    const countArrCleaned = countArr.slice(0, countArr.length - 1);
    const count = sumArray(countArrCleaned);
    // console.log("count arr: ", countArrCleaned);
    console.log("count: ", count);

    // TOTAL
    // total ( do QTY bị ẩn, nếu lấy theo TOTAL thì sẽ có thêm value là 'QTY' )
    const getTotals = getValuesBelowTitleExact(cellMatrix, "TOTAL");
    const totalArr = removeAfterString(getTotals);
    const totalArrFiltered = filterOnlyNumericValues(totalArr);


    const getTotals_1 = getValuesBelowTitleAuto(cellMatrix, "TOTAL QTY");
    const totalArr_1 = removeAfterString(getTotals_1);
    const totalArrFiltered_1 = filterOnlyNumericValues(totalArr_1);
    const totalArrCleaned_1 = removeItemFromArray(totalArrFiltered_1, "QTY");

    const totalArrFinal = pickNonEmptyArray(totalArrFiltered, totalArrCleaned_1);

    // console.log("total arr: ", totalArr);
    const total = sumArray(totalArrFinal.slice(0, totalArrFinal.length - 1));
    console.log("total: ", total);

    // GROSS 
    const getGross = getValuesBelowTitleAuto(cellMatrix, "G.W");
    // console.log("gorss arr: ", getGross);
    const grossArr = removeItemFromArray(getGross, "G.W");
    console.log("gross arr: ", grossArr);
    const gross = smartCountSum(grossArr);
    console.log("gross: ", gross);

    // NET
    const getNet = getValuesBelowTitleAuto(cellMatrix, "N.W");
    const netArr = removeItemFromArray(getNet, "N.W");
    const net = smartCountSum(netArr);
    console.log("net: ", net);

    // VOLUME
    const getVolume = getValuesBelowTitleAuto(cellMatrix, "CBM");
    const volumeArr = getVolume.slice(0, getVolume.length - 1);
    const volume = sumArray(volumeArr);
    console.log("volume: ", volume);





    // CARTON FORMULA
    const getL = getValuesBelowTitleAuto(cellMatrix, "CARTON DIMENSION (CM)");
    const getW = cellMatrix[10].cells
    const getH = cellMatrix[11].cells
    const L_cleaned = filterOnlyNumericValues(getL);
    const W_cleaned = filterOnlyNumericValues(getW);
    const H_cleaned = filterOnlyNumericValues(getH);
    const setCarton = [L_cleaned, W_cleaned, H_cleaned];
    const buildCartonDimension = buildCartonDimensionFormulas(setCarton);
    console.log("CARTON DIMENSION (CM): ", buildCartonDimension);

    const cartonAndCount = mergeToObjectArray(buildCartonDimension, countArr);
    console.log("carton and count: ", cartonAndCount);
    const cartonDimension = buildMeasFormula(cartonAndCount);
    console.log("carton dimension: ", cartonDimension);

    // po 
    const po = sheetName;

    // ===================== GẮN SUMMARY =====================
    result.summary.push({
      sheetName,
      article: styleName,
      po: po,
      color: color,
      countTotal: count,
      total: total,
      gross: gross,
      net: net,
      volume: volume,
      cartonFormula: cartonDimension,
    });

  });
  console.log(result.summary);
  return result;
}

// henri bulk
const HENRI_BULK = [
  "ARTICLE / STYLE",
  "ORDER NO:",
  "COLOR",
  "COUNT",
  "TOTAL QTY",
  "GROSS",
  "NET",
  "VOLUME (CBM)",
  "CARTON DIMENSION (CM)",
];
function previewHENRI_BULK(workbook) {
  const result = {
    sheetNames: [],
    data: {},
    summary: {},
  };

  // phần này xử lý cho ONE IBIS 021W:
  workbook.SheetNames.forEach((sheetName) => {

    const sheet = workbook.Sheets[sheetName];

    // 🛑 GUARD: nếu sheet KHÔNG thỏa ROSSIGNOL BULK → BỎ QUA
    if (!isSheetMatchTemplate(sheet, HENRI_BULK)) {
      console.log(`⏭️ Skip sheet: ${sheetName}`);
      return;
    }
    const cellMatrix = []; //  mỗi sheet 1 matrix

    const rawRows = xlsx.utils.sheet_to_json(sheet, { defval: "" });

    // 1️ lọc cột
    const columnCleaned = filterInvalidColumns(rawRows);

    // 2️ tìm cột COUNT
    const countColumnKey = findColumnKeyByTitle(columnCleaned, "count");

    // 3️ lọc row COUNT = 0
    const rowCleaned = columnCleaned.filter((row) => {
      if (isValidColumn([row], countColumnKey)) return true;
      return !isZeroInCountColumn(row, countColumnKey);
    });

    //  GIỮ NGUYÊN DATA CHO FRONTEND
    result.sheetNames.push(sheetName);
    result.data[sheetName] = rowCleaned.map((row) => ({
      SHEET: sheetName,
      ...row,
    }));

    // 4️ build cellMatrix cho sheet này
    const columnCells = extractColumnCellArrays(rowCleaned, sheetName);
    cellMatrix.push(...columnCells);

    // ===================== TÍNH TOÁN =====================
    // article
    const getArticle = getValuesBelowTitleAuto(cellMatrix, "ARTICLE / STYLE");
    const article = getArticle[0];
    console.log("article: ", article);

    // po 
    const getPo = getValuesBelowTitleAuto(cellMatrix, "PO");
    const po = getPo[0];
    console.log("po: ", po);

    // color
    const getColor = getValuesBelowTitleAuto(cellMatrix, "COLOR");
    const colorCodes = [
      "999 Black",
    ];
    const colorArr = filterColorsInCodes(getColor, colorCodes);
    console.log("COLOR : ", colorArr);
    const color = collapseIfSame(colorArr);

    // COUNT
    const getCount = getValuesBelowTitleAuto(cellMatrix, "COUNT");
    const countArr = filterOnlyNumericValues(getCount);
    console.log("count arr: ", countArr);
    const count = countArr[countArr.length - 1];
    console.log("count: ", count);

    // TOTAL
    const getTotal = getValuesBelowTitleAuto(cellMatrix, "TOTAL");
    const totalArr = filterOnlyNumericValues(getTotal);
    console.log("total arr: ", totalArr);
    const total = totalArr[totalArr.length - 1]
    console.log("total: ", total)

    // gross 
    const getGross = getValuesBelowTitleAuto(cellMatrix, "GROSS");
    const grossArr = filterOnlyNumericValues(getGross);
    const gross = sumArray(grossArr.slice(0, grossArr.length - 1));
    console.log("gross: ", gross);
    // net 
    const getNet = getValuesBelowTitleAuto(cellMatrix, "NET");
    const netArr = filterOnlyNumericValues(getNet);
    const net = sumArray(netArr.slice(0, netArr.length - 1));
    console.log("net: ", net);
    // volume
    const getVolume = getValuesBelowTitleAuto(cellMatrix, "VOLUME (CBM)");
    const volumeArr = filterOnlyNumericValues(getVolume);
    const volume = sumArray(volumeArr.slice(0, volumeArr.length - 1));
    console.log("volume: ", volume);

    // CARTON FORMULA
    const getDimension = getValuesBelowTitleAuto(cellMatrix, "CARTON DIMENSION (CM)");
    console.log("dimension: ", getDimension);

    const dimensions = getDimension
      .filter(v => typeof v === "string")
      .map(v => v.trim().toUpperCase())
      .filter(v => /^\d+X\d+X\d+$/.test(v));

    const countsArr = countArr.slice(0, -1);

    const formula = Object.entries(
      dimensions.map((d, i) => [d, countsArr[i]])
        .reduce((acc, [d, c]) => {
          acc[d] = (acc[d] || 0) + (c || 0);
          return acc;
        }, {})
    )
      .map(([d, t]) => `${d} * ${t}`)
      .join(", ");

    console.log("carton dimension: ", formula);

    // ===================== GẮN SUMMARY =====================
    result.summary[sheetName] = {
      sheetName,
      article: article,
      po: po,
      color: color,
      countTotal: count,
      total: total,
      gross: gross,
      net: net,
      volume: volume,
      cartonFormula: formula,
    };

  });
  console.log(result.summary)

  return result;
}

// innovation
const INNOVATION = [
  "STYLE",
  "PO. NO.",
  "COLOUR",
  "CTNS.",
  "TOTAL QTY.",
];
function previewINNOVATION(workbook) {
  const result = {
    sheetNames: [],
    data: {},
    summary: [],
  };

  let grossTmp = "";
  let netTmp = "";
  let volumeTmp = "";
  let cartonDimensionTmp = "";

  // Sheet 
  workbook.SheetNames.forEach((sheetName) => {


    const cellMatrix = []; //  mỗi sheet 1 matrix

    const sheet = workbook.Sheets[sheetName];

    const rawRows = xlsx.utils.sheet_to_json(sheet, { defval: "" });

    // 1️ lọc cột
    const columnCleaned = filterInvalidColumns(rawRows);

    // 2️ tìm cột COUNT
    const countColumnKey = findColumnKeyByTitle(columnCleaned, "count_");

    const rowCleaned = columnCleaned.filter(row => {

      const rowText = normalize_(Object.values(row).join(" "));

      // 1️ GIỮ row có TOTAL QTY.
      if (rowText.includes("total qty")) {
        return true;
      }

      if (rowText.includes("total gross")) {
        return true;
      }

      if (rowText.includes("total net")) {
        return true;
      }

      // 2️ LOẠI row có TOTAL nhưng KHÔNG phải TOTAL QTY.
      if (rowText.includes("total")) {
        return false;
      }

      if (isValidColumn([row], countColumnKey)) return true;
      return !isZeroInCountColumn(row, countColumnKey);
    });


    //  GIỮ NGUYÊN DATA CHO FRONTEND
    result.sheetNames.push(sheetName);
    result.data[sheetName] = rowCleaned.map((row) => ({
      SHEET: sheetName,
      ...row,
    }));


    // 4️ build cellMatrix cho sheet này
    const columnCells = extractColumnCellArrays(rowCleaned, sheetName);
    cellMatrix.push(...columnCells);


    // sheet 1

    if (sheetName === "Sheet1") {
      // gross
      const getGross = getValueForGrossNet_INNOVATION(sheet, "TOTAL GROSS WEIGHT");
      console.log("getGross:", getGross);

      // net
      const getNet = getValueForGrossNet_INNOVATION(sheet, "TOTAL NET WEIGHT");
      console.log("getNet:", getNet);

      const getVolume = getValueForGrossNet_INNOVATION(sheet, "MEAUSREMENT");
      console.log("volume: ", getVolume);

      grossTmp = getGross;
      netTmp = getNet;
      volumeTmp = getVolume;
      return;
    }
  });

  workbook.SheetNames.forEach((sheetName) => {

    if (sheetName === "Sheet1") {
      return;
    }

    const sheet = workbook.Sheets[sheetName];

    // 🛑 GUARD: nếu sheet KHÔNG thỏa ROSSIGNOL BULK → BỎ QUA
    if (!isSheetMatchTemplate(sheet, INNOVATION)) {
      console.log(`⏭️ Skip sheet: ${sheetName}`);
      return;
    }

    const cellMatrix = []; //  mỗi sheet 1 matrix

    const rawRows = xlsx.utils.sheet_to_json(sheet, { defval: "" });

    // 1️ lọc cột
    const columnCleaned = filterInvalidColumns(rawRows);

    // 2️ tìm cột COUNT
    const countColumnKey = findColumnKeyByTitle(columnCleaned, "count_");

    const rowCleaned = columnCleaned.filter(row => {

      const rowText = normalize_(Object.values(row).join(" "));

      // 1️ GIỮ row có TOTAL QTY.
      if (rowText.includes("total qty")) {
        return true;
      }

      if (rowText.includes("total gross")) {
        return true;
      }

      if (rowText.includes("total net")) {
        return true;
      }

      // 2️ LOẠI row có TOTAL nhưng KHÔNG phải TOTAL QTY.
      if (rowText.includes("total")) {
        return false;
      }

      if (rowText.includes("ttl q'ty")) {
        return false;
      }

      if (hasAtLeastNEmptyCells(row, 8)) return false;

      if (isValidColumn([row], countColumnKey)) return true;
      return !isZeroInCountColumn(row, countColumnKey);
    });

    //  GIỮ NGUYÊN DATA CHO FRONTEND
    result.sheetNames.push(sheetName);
    result.data[sheetName] = rowCleaned.map((row) => ({
      SHEET: sheetName,
      ...row,
    }));


    // 4️ build cellMatrix cho sheet này
    const columnCells = extractColumnCellArrays(rowCleaned, sheetName);
    cellMatrix.push(...columnCells);

    // ===================== TÍNH TOÁN =====================

    console.log("------------------------------------------------------");

    // color 
    const getColor = getValuesBelowTitleAuto(cellMatrix, "COLOUR");
    const colorArr = getColor;
    // console.log("color arr: ", colorArr);

    // article / style
    const getArticleStyle = getValuesBelowTitleAuto(cellMatrix, "STYLE");
    const articleStyleArr1 = removeAfterStringValue(getArticleStyle, "TOTAL")
    const articleStyleArr2 = removeAfterStringValue(articleStyleArr1, "ARTICLE / STYLE")
    const articleStyleArr3 = removeAfterStringValue(articleStyleArr2, "PO")
    const articleStyleArr4 = removeAfterStringValue(articleStyleArr3, "COLOR")
    const articleStyleArr5 = removeAfterStringValue(articleStyleArr4, "SHADE ")
    const articleStyleArr = removeAfterStringValue(articleStyleArr5, "total")

    // console.log("ARTICLE / STYLE: ", articleStyleArr);

    // po
    const getPO = getValuesBelowTitleAuto(cellMatrix, "PO. No.");
    const POArrCleaned = removeAfterStringValue(getPO, "TOTAL");
    const POArr = POArrCleaned;
    // console.log("po: ", POArr)


    // kiểm duyệt color, article / style, po
    const styleCheck = collapseIfSame(articleStyleArr);
    const colorCheck = collapseIfSame(colorArr);
    const POCheck = collapseIfSame(POArr);

    let styleName = "";
    let color = "";
    let PO = "";

    if (!isArray(styleCheck) || !isArray(colorCheck) || !isArray(POCheck)) {
      styleName = styleCheck;
      color = colorCheck;
      PO = POCheck;
    }
    else {
      styleName = articleStyleArr;
      color = colorArr;
      PO = POArr;
    }

    // COUNT
    const getCount = getValuesBelowTitleAuto(cellMatrix, "CTNS.");
    const countArr = filterOnlyNumericValues(getCount);
    const count = sumArray(countArr);
    // console.log("count arr: ", countArrCleaned);
    console.log("count: ", count);

    // TOTAL
    // total ( do QTY bị ẩn, nếu lấy theo TOTAL thì sẽ có thêm value là 'QTY' )

    const getTotals = getValuesBelowTitleAuto(cellMatrix, "TOTAL QTY.");
    const totalArrFiltered = filterOnlyNumericValues(getTotals);
    // const totalArrCleaned = removeItemFromArray(getTotals, "QTY");
    // const totalArr = removeAfterString(totalArrCleaned);
    // const totalArrFiltered = filterOnlyNumericValues(totalArr);
    // console.log("total arr: ", totalArrFiltered);
    // const total = sumArray(totalArrFiltered.slice(0, totalArrFiltered.length - 1));
    // console.log("total: ", total);

    const totalArr = trimArrayByArray(getTotals, colorArr);
    console.log(totalArr);
    const total = sumArray(totalArr);

    // GROSS 
    // const getGross = getValuesBelowTitleAuto(cellMatrix, "GROSS");
    // console.log("gorss arr: ", getGross);
    // const grossArr = removeItemFromArray(getGross, "WEIGHT");
    const gross = grossTmp;
    // console.log("gorss: ", gross);

    // NET
    // const getNet = getValuesBelowTitleAuto(cellMatrix, "NET");
    // const netArr = removeItemFromArray(getNet, "WEIGHT");
    const net = netTmp;
    // console.log("net: ", net);

    // VOLUME
    // const getVolume = getValuesBelowTitleAuto(cellMatrix, "VOLUME (CBM)");
    // const volumeArr = getVolume;
    const volume = 0;
    // console.log("volume: ", volume);

    // CARTON FORMULA
    const getCartonDimension = getValuesBelowTitleAuto(cellMatrix, "CARTON DIMENSION (CM)");
    const cartonDimensionArr = removeItemFromArray(getCartonDimension, "CARTON DIMENSION (CM)")
    console.log("CARTON DIMENSION (CM): ", cartonDimensionArr);
    const cartonAndCount = mergeToObjectArray(cartonDimensionArr, countArr);
    console.log("carton and count: ", cartonAndCount);
    const cartonDimension = buildMeasFormula(cartonAndCount);
    console.log("carton dimension: ", cartonDimension);

    // ===================== GẮN SUMMARY =====================
    result.summary.push({
      sheetName,
      article: styleName,
      po: PO,
      color: color,
      countTotal: count,
      total: total,
      gross: gross,
      net: net,
      volume: volume,
      cartonFormula: cartonDimension,
    });

  });
  console.log(result.summary);

  return result;
}
//WITTT THIEN BINH
const WITTT_THIEN_BINH = [
  "STYLE NO",
  "ORDER NO",
  "ART NO. COLOR",
  "CARTONS",
  "TOTAL PCS",
  "N.W",
  "G.W",
  "CBM (CM)"
];
function previewWITTT_THIEN_BINH(workbook) {
  const result = {
    sheetNames: [],
    data: {},
    summary: [],
  };

  workbook.SheetNames.forEach((sheetName) => {

    const sheet = workbook.Sheets[sheetName];
    // 🛑 GUARD: nếu sheet KHÔNG thỏa ROSSIGNOL BULK → BỎ QUA
    if (!isSheetMatchTemplate(sheet, WITTT_THIEN_BINH)) {
      console.log(`⏭️ Skip sheet: ${sheetName}`);
      return;
    }
    const cellMatrix = []; //  mỗi sheet 1 matrix


    const rawRows = xlsx.utils.sheet_to_json(sheet, { defval: "" });

    // 1️ lọc cột
    const columnCleaned = filterInvalidColumns(rawRows);

    // 2️ tìm cột COUNT
    const countColumnKey = findColumnKeyByTitle(columnCleaned, "count_");

    const rowCleaned = columnCleaned.filter(row => {

      const rowText = normalize_(Object.values(row).join(" "));

      // 1️ GIỮ row có 
      // if (rowText.includes("total qty")) {
      //   return true;
      // }

      // if (rowText.includes("total gross")) {
      //   return true;
      // }

      // if (rowText.includes("total net")) {
      //   return true;
      // }

      // 2️ LOẠI row có TOTAL 
      // if (rowText.includes("total")) {
      //   return false;
      // }

      // if(rowText.includes("ttl q'ty")){
      //   return false;
      // }

      // if (hasAtLeastNEmptyCells(row, 6)) return false;

      if (isValidColumn([row], countColumnKey)) return true;
      return !isZeroInCountColumn(row, countColumnKey);
    });


    //  GIỮ NGUYÊN DATA CHO FRONTEND
    result.sheetNames.push(sheetName);
    result.data[sheetName] = rowCleaned.map((row) => ({
      SHEET: sheetName,
      ...row,
    }));


    // 4️ build cellMatrix cho sheet này
    const columnCells = extractColumnCellArrays(rowCleaned, sheetName);
    cellMatrix.push(...columnCells);

    // ===================== TÍNH TOÁN =====================

    console.log("------------------------------------------------------");

    // article / style
    const getArticleStyle = extractPurchaseOrderNo(sheet, "STYLE NO.:");
    // console.log("ARTICLE / STYLE: ", articleStyleArr);

    // po
    const getPO = extractPurchaseOrderNo(sheet, "ORDER NO. :")
    const POArr = getPO;
    console.log("po: ", POArr)


    // color 
    const getColor = getValuesBelowTitleAuto(cellMatrix, "Art no. Color");
    const colorFilter = filterOnlyNumericValues(getColor);
    const colorArr = colorFilter;
    // console.log("color arr: ", colorArr);


    // kiểm duyệt color, article / style, po
    const styleCheck = collapseIfSame(getArticleStyle);
    const colorCheck = collapseIfSame(colorArr);
    const POCheck = collapseIfSame(POArr);

    let styleName = "";
    let color = "";
    let PO = "";

    if (!isArray(styleCheck) || !isArray(colorCheck) || !isArray(POCheck)) {
      styleName = styleCheck;
      color = colorCheck;
      PO = POCheck;
    }
    else {
      styleName = getArticleStyle;
      color = colorArr;
      PO = POArr;
    }

    // COUNT
    const getCount = getValuesBelowTitleAuto(cellMatrix, "Cartons");
    const countArr = filterOnlyNumericValues(getCount);

    const count = sumArray(countArr.slice(0, countArr.length - 1));
    // console.log("count arr: ", countArrCleaned);
    console.log("count: ", count);

    // TOTAL
    // total ( do QTY bị ẩn, nếu lấy theo TOTAL thì sẽ có thêm value là 'QTY' )

    const getTotals = getValuesBelowTitleAuto(cellMatrix, "Total Pcs");
    const totalArrFiltered = filterOnlyNumericValues(getTotals);
    // const totalArrCleaned = removeItemFromArray(getTotals, "QTY");
    // const totalArr = removeAfterString(totalArrCleaned);
    // const totalArrFiltered = filterOnlyNumericValues(totalArr);
    // console.log("total arr: ", totalArrFiltered);
    // const total = sumArray(totalArrFiltered.slice(0, totalArrFiltered.length - 1));
    // console.log("total: ", total);

    // const totalArr = trimArrayByArray(getTotals, colorArr);

    const totalArr = trimArrayByArray(totalArrFiltered, countArr);

    const total = sumArray(totalArr.slice(0, totalArr.length - 1));
    console.log("total: ", total);
    // GROSS 
    const getGross = getValuesBelowTitleAuto(cellMatrix, "G.W");
    const grossFiltered = filterOnlyNumericValues(getGross);
    // console.log("gorss arr: ", getGross);
    // const grossArr = removeItemFromArray(getGross, "WEIGHT");
    const gross = grossFiltered[0];
    // console.log("gorss: ", gross);

    // NET
    const getNet = getValuesBelowTitleAuto(cellMatrix, "N.W");
    const netFiltered = filterOnlyNumericValues(getNet);
    // const netArr = removeItemFromArray(getNet, "WEIGHT");
    const net = netFiltered[0];
    // console.log("net: ", net);

    // VOLUME
    // const getVolume = getValuesBelowTitleAuto(cellMatrix, "VOLUME (CBM)");
    // const volumeArr = getVolume;
    const volume = 0;
    // console.log("volume: ", volume);

    // CARTON FORMULA
    const getCartonDimension = getValuesBelowTitleAuto(cellMatrix, "CBM (CM)");
    const cartonDimensionArr = removeItemFromArray(getCartonDimension, "CBM\n(CM)");
    console.log("CARTON DIMENSION (CM): ", cartonDimensionArr);
    const cartonAndCount = mergeToObjectArray(cartonDimensionArr, countArr);
    console.log("carton and count: ", cartonAndCount);
    const cartonDimension = buildMeasFormula(cartonAndCount);
    console.log("carton dimension: ", cartonDimension);

    // ===================== GẮN SUMMARY =====================
    result.summary.push({
      sheetName,
      article: styleName,
      po: PO,
      color: color,
      countTotal: count,
      total: total,
      gross: gross,
      net: net,
      volume: volume,
      cartonFormula: cartonDimension,
    });

  });
  console.log(result.summary);

  return result;
}

//WITTT BMPY
const WITTT_BMPY = [
  "STYLE NO",
  "ORDER NO",
  "ART NO. COLOR",
  "CARTONS",
  "TOTAL PCS",
  "N.W",
  "G.W",
  "(CM)"
];
function previewWITTT_BMPY(workbook) {
  const result = {
    sheetNames: [],
    data: {},
    summary: [],
  };

  workbook.SheetNames.forEach((sheetName) => {

    const sheet = workbook.Sheets[sheetName];
    // 🛑 GUARD: nếu sheet KHÔNG thỏa ROSSIGNOL BULK → BỎ QUA
    if (!isSheetMatchTemplate(sheet, WITTT_BMPY)) {
      console.log(`⏭️ Skip sheet: ${sheetName}`);
      return;
    }

    const cellMatrix = []; //  mỗi sheet 1 matrix


    const rawRows = xlsx.utils.sheet_to_json(sheet, { defval: "" });

    // 1️ lọc cột
    const columnCleaned = filterInvalidColumns(rawRows);

    // 2️ tìm cột COUNT
    const countColumnKey = findColumnKeyByTitle(columnCleaned, "count_");

    const rowCleaned = columnCleaned.filter(row => {

      const rowText = normalize_(Object.values(row).join(" "));

      // 1️ GIỮ row có 
      // if (rowText.includes("total qty")) {
      //   return true;
      // }

      // if (rowText.includes("total gross")) {
      //   return true;
      // }

      // if (rowText.includes("total net")) {
      //   return true;
      // }

      // 2️ LOẠI row có TOTAL 
      // if (rowText.includes("total")) {
      //   return false;
      // }

      // if(rowText.includes("ttl q'ty")){
      //   return false;
      // }

      // if (hasAtLeastNEmptyCells(row, 6)) return false;

      if (isValidColumn([row], countColumnKey)) return true;
      return !isZeroInCountColumn(row, countColumnKey);
    });


    //  GIỮ NGUYÊN DATA CHO FRONTEND
    result.sheetNames.push(sheetName);
    result.data[sheetName] = rowCleaned.map((row) => ({
      SHEET: sheetName,
      ...row,
    }));


    // 4️ build cellMatrix cho sheet này
    const columnCells = extractColumnCellArrays(rowCleaned, sheetName);
    cellMatrix.push(...columnCells);

    // ===================== TÍNH TOÁN =====================

    console.log("------------------------------------------------------");

    // article / style
    const getArticleStyle = extractPurchaseOrderNo(sheet, "STYLE NO.:");
    // console.log("ARTICLE / STYLE: ", articleStyleArr);

    // po
    const getPO = extractPurchaseOrderNo(sheet, "ORDER NO. :")
    const POArr = getPO;
    console.log("po: ", POArr)


    // color 
    const getColor = getValuesBelowTitleAuto(cellMatrix, "Art no. Color");
    const colorFilter = filterOnlyNumericValues(getColor);
    const colorArr = colorFilter;
    // console.log("color arr: ", colorArr);


    // kiểm duyệt color, article / style, po
    const styleCheck = collapseIfSame(getArticleStyle);
    const colorCheck = collapseIfSame(colorArr);
    const POCheck = collapseIfSame(POArr);

    let styleName = "";
    let color = "";
    let PO = "";

    if (!isArray(styleCheck) || !isArray(colorCheck) || !isArray(POCheck)) {
      styleName = styleCheck;
      color = colorCheck;
      PO = POCheck;
    }
    else {
      styleName = getArticleStyle;
      color = colorArr;
      PO = POArr;
    }

    // COUNT
    const getCount = getValuesBelowTitleAuto(cellMatrix, "Cartons");
    const countArr = filterOnlyNumericValues(getCount);

    const count = sumArray(countArr.slice(0, countArr.length - 1));
    // console.log("count arr: ", countArrCleaned);
    console.log("count: ", count);

    // TOTAL
    // total ( do QTY bị ẩn, nếu lấy theo TOTAL thì sẽ có thêm value là 'QTY' )

    const getTotals = getValuesBelowTitleAuto(cellMatrix, "Total Pcs");
    const totalArrFiltered = filterOnlyNumericValues(getTotals);
    // const totalArrCleaned = removeItemFromArray(getTotals, "QTY");
    // const totalArr = removeAfterString(totalArrCleaned);
    // const totalArrFiltered = filterOnlyNumericValues(totalArr);
    // console.log("total arr: ", totalArrFiltered);
    // const total = sumArray(totalArrFiltered.slice(0, totalArrFiltered.length - 1));
    // console.log("total: ", total);

    // const totalArr = trimArrayByArray(getTotals, colorArr);

    const totalArr = trimArrayByArray(totalArrFiltered, countArr);

    const total = sumArray(totalArr.slice(0, totalArr.length - 1));
    console.log("total: ", total);
    // GROSS 
    const getGross = getValuesBelowTitleAuto(cellMatrix, "G.W");
    const grossFiltered = filterOnlyNumericValues(getGross);
    // console.log("gorss arr: ", getGross);
    // const grossArr = removeItemFromArray(getGross, "WEIGHT");
    const gross = grossFiltered[0];
    // console.log("gorss: ", gross);

    // NET
    const getNet = getValuesBelowTitleAuto(cellMatrix, "N.W");
    const netFiltered = filterOnlyNumericValues(getNet);
    // const netArr = removeItemFromArray(getNet, "WEIGHT");
    const net = netFiltered[0];
    // console.log("net: ", net);

    // VOLUME
    // const getVolume = getValuesBelowTitleAuto(cellMatrix, "VOLUME (CBM)");
    // const volumeArr = getVolume;
    const volume = 0;
    // console.log("volume: ", volume);

    // CARTON FORMULA
    const getCartonDimension = getValuesBelowTitleAuto(cellMatrix, "(CM)");
    const cartonDimensionArr = removeItemFromArray(getCartonDimension, "CBM\n(CM)");
    console.log("CARTON DIMENSION (CM): ", cartonDimensionArr);
    const cartonDimensionCleaned = trimArrayByArray(cartonDimensionArr, countArr.slice(0, countArr.length - 1));
    const cartonAndCount = mergeToObjectArray(cartonDimensionCleaned, countArr);
    console.log("carton and count: ", cartonAndCount);
    const cartonDimension = buildMeasFormula(cartonAndCount);
    console.log("carton dimension: ", cartonDimension);

    // ===================== GẮN SUMMARY =====================
    result.summary.push({
      sheetName,
      article: styleName,
      po: PO,
      color: color,
      countTotal: count,
      total: total,
      gross: gross,
      net: net,
      volume: volume,
      cartonFormula: cartonDimension,
    });

  });
  console.log(result.summary);

  return result;
}

// ODLO SMS
const ODLO_SMS = [
  "STYLE",
  "PO",
  "COLOR",
  "COUNT",
  "Q'TY",
  "NW KG",
  "GW KG",
  "CBM",
  "KÍCH THÙNG"
];
function previewODLO_SMS(workbook) {
  const result = {
    sheetNames: [],
    data: {},
    summary: [],
  };

  workbook.SheetNames.forEach((sheetName) => {
    const sheet = workbook.Sheets[sheetName];

    if (!isSheetMatchTemplate(sheet, ODLO_SMS)) {
      console.log(`⏭️ Skip sheet: ${sheetName}`);
      return;
    }

    const cellMatrix = []; //  mỗi sheet 1 matrix

    const rawRows = xlsx.utils.sheet_to_json(sheet, { defval: "" });

    // 1️ lọc cột
    const columnCleaned = filterInvalidColumns(rawRows);

    // 2️ tìm cột COUNT
    const countColumnKey = findColumnKeyByTitle(columnCleaned, "count_");

    const rowCleaned = columnCleaned.filter(row => {

      const rowText = normalize_(Object.values(row).join(" "));

      // 1️ GIỮ row có 
      // if (rowText.includes("total qty")) {
      //   return true;
      // }

      // if (rowText.includes("total gross")) {
      //   return true;
      // }

      // if (rowText.includes("total net")) {
      //   return true;
      // }

      // 2️ LOẠI row có TOTAL 
      if (rowText.includes("total")) {
        return false;
      }

      // if(rowText.includes("ttl q'ty")){
      //   return false;
      // }

      // if (hasAtLeastNEmptyCells(row, 6)) return false;

      if (isValidColumn([row], countColumnKey)) return true;
      return !isZeroInCountColumn(row, countColumnKey);
    });


    //  GIỮ NGUYÊN DATA CHO FRONTEND
    result.sheetNames.push(sheetName);
    result.data[sheetName] = rowCleaned.map((row) => ({
      SHEET: sheetName,
      ...row,
    }));


    // 4️ build cellMatrix cho sheet này
    const columnCells = extractColumnCellArrays(rowCleaned, sheetName);
    cellMatrix.push(...columnCells);

    // ===================== TÍNH TOÁN =====================

    console.log("------------------------------------------------------");

    // // // article / style
    const getArticleStyle = getValuesBelowTitleAuto(cellMatrix, "Style");
    const articleStyleArr = removeItemFromArray(getArticleStyle, "Style");
    console.log("ARTICLE / STYLE: ", articleStyleArr);

    // // // po
    const getPO = getValuesBelowTitleAuto(cellMatrix, "PO")
    const POArr = getPO;
    // console.log("po: ", POArr)


    // color 
    const getColor = getValuesBelowTitleAuto(cellMatrix, "color");
    // const colorFilter = filterOnlyNumericValues(getColor);
    const colorArr = getColor;
    // console.log("color arr: ", colorArr);


    // kiểm duyệt color, article / style, po
    const styleCheck = collapseIfSame(articleStyleArr);
    const colorCheck = collapseIfSame(colorArr);
    const POCheck = collapseIfSame(POArr);

    let styleName = "";
    let color = "";
    let PO = "";

    if (!isArray(styleCheck) || !isArray(colorCheck) || !isArray(POCheck)) {
      styleName = styleCheck;
      color = colorCheck;
      PO = POCheck;
    }
    else {
      styleName = articleStyleArr;
      color = colorArr;
      PO = POArr;
    }

    // COUNT
    const getCount = getValuesBelowTitleAuto(cellMatrix, "COUNT");
    const countArr = normalizeCartonArray(getCount)

    const count = sumArray(countArr);
    // console.log("count arr: ", countArrCleaned);
    console.log("count: ", count);

    // TOTAL
    // total ( do QTY bị ẩn, nếu lấy theo TOTAL thì sẽ có thêm value là 'QTY' )

    const getTotals = getValuesBelowTitleAuto(cellMatrix, "Q'ty");
    const totalArrFiltered = filterOnlyNumericValues(getTotals);
    // const totalArrCleaned = removeItemFromArray(getTotals, "QTY");
    // const totalArr = removeAfterString(totalArrCleaned);
    // const totalArrFiltered = filterOnlyNumericValues(totalArr);
    // console.log("total arr: ", totalArrFiltered);
    // const total = sumArray(totalArrFiltered.slice(0, totalArrFiltered.length - 1));
    // console.log("total: ", total);

    // const totalArr = trimArrayByArray(getTotals, colorArr);

    const total = sumArray(totalArrFiltered);
    console.log("total: ", total);
    // GROSS 
    const getGross = getValuesBelowTitleAuto(cellMatrix, "GW kg");
    const grossFiltered = filterOnlyNumericValues(getGross);
    // console.log("gorss arr: ", getGross);
    // const grossArr = removeItemFromArray(getGross, "WEIGHT");
    const gross = sumArray(grossFiltered)
    // console.log("gorss: ", gross);

    // NET
    const getNet = getValuesBelowTitleAuto(cellMatrix, "NW kg");
    const netFiltered = filterOnlyNumericValues(getNet);
    // const netArr = removeItemFromArray(getNet, "WEIGHT");
    const net = sumArray(netFiltered);
    // console.log("net: ", net);

    // VOLUME
    const getVolume = getValuesBelowTitleAuto(cellMatrix, "CBM");
    const volumeArr = getVolume;
    const volume = sumArray(volumeArr);
    // console.log("volume: ", volume);

    // CARTON FORMULA
    const getCartonDimension = getValuesBelowTitleAuto(cellMatrix, "Kích thùng");
    const cartonDimensionArr = removeItemFromArray(getCartonDimension, "CBM\n(CM)");
    console.log("CARTON DIMENSION (CM): ", cartonDimensionArr);
    const cartonDimensionCleaned = trimArrayByArray(cartonDimensionArr, countArr);
    const cartonAndCount = mergeToObjectArray(cartonDimensionCleaned, countArr);
    console.log("carton and count: ", cartonAndCount);
    const cartonDimension = buildMeasFormula(cartonAndCount);
    console.log("carton dimension: ", cartonDimension);

    // ===================== GẮN SUMMARY =====================
    result.summary.push({
      sheetName,
      article: styleName,
      po: PO,
      color: color,
      countTotal: count,
      total: total,
      gross: gross,
      net: net,
      volume: volume,
      cartonFormula: cartonDimension,
    });

  });
  console.log(result.summary);

  return result;
}

// POC BULK
const POC_BULK = [
  "ART. NO",
  "PO NO",
  "COLOR",
  "CTNS",
  "TOTAL (PCS)",
  "TTL N.W.",
  "TTL G.W.",
  "CBM",
  "DIMS"
];
function previewPOC_BULK(workbook) {
  const result = {
    sheetNames: [],
    data: {},
    summary: [],
  };

  workbook.SheetNames.forEach((sheetName) => {
    const sheet = workbook.Sheets[sheetName];

    if (!isSheetMatchTemplate(sheet, POC_BULK)) {
      console.log(`⏭️ Skip sheet: ${sheetName}`);
      return;
    }

    const cellMatrix = []; //  mỗi sheet 1 matrix

    const rawRows = xlsx.utils.sheet_to_json(sheet, { defval: "" });

    // 1️ lọc cột
    const columnCleaned = filterInvalidColumns(rawRows);

    // 2️ tìm cột COUNT
    const countColumnKey = findColumnKeyByTitle(columnCleaned, "count_)");


    // const find_ART_NO = cellText

    let index = 0;
    columnCleaned.filter(row => {

      const zeroKey = Object.keys(row)[0];
      const zeroCell = String(row[zeroKey]);

      const firstKey = Object.keys(row)[1];
      const firstCell = String(row[firstKey]);

      const secondKey = Object.keys(row)[2];
      const secondCell = String(row[secondKey]);

      if (firstCell === "ART. NO") {
        index = 1;
        return;
      }
      else if (secondCell === "ART. NO") {
        index = 2;
        return;
      }
      else if (zeroCell === "ART. NO") {
        index = 0;
        return;
      }
    })

    const rowCleaned = columnCleaned.filter(row => {

      const rowText = normalize_(Object.values(row).join(" "));

      // 🔴 FIX: lấy đúng cột đầu tiên
      const firstKey = Object.keys(row)[index];
      const firstCell = String(row[firstKey] || "")
        .toUpperCase()
        .trim();


      if (firstCell === "") {
        return false;
      }


      if (hasAtLeastNEmptyCells(row, 15)) return false;

      if (isValidColumn([row], countColumnKey)) return true;
      return !isZeroInCountColumn(row, countColumnKey);
    });


    //  GIỮ NGUYÊN DATA CHO FRONTEND
    result.sheetNames.push(sheetName);
    result.data[sheetName] = rowCleaned.map((row) => ({
      SHEET: sheetName,
      ...row,
    }));


    // 4️ build cellMatrix cho sheet này
    const columnCells = extractColumnCellArrays(rowCleaned, sheetName);
    cellMatrix.push(...columnCells);



    // ===================== TÍNH TOÁN =====================

    console.log("------------------------------------------------------");

    // // // article / style
    const getArticleStyle = getValuesBelowTitleAuto(cellMatrix, "ART. NO");
    const articleStyleArrCleaned = removeItemFromArray(getArticleStyle, "ART. NO");
    const articleStyleArr = extractFirst5Digits(articleStyleArrCleaned);
    console.log("ARTICLE / STYLE: ", articleStyleArr);

    // // // po
    const getPO = getValuesBelowTitleAuto(cellMatrix, "PO NO");
    const POCleaned = removeItemFromArray(getPO, "PO NO");
    const POArr = trimArrayByArray(POCleaned, articleStyleArr);
    console.log("po: ", POArr)


    // color 
    const getColor = getValuesBelowTitleAuto(cellMatrix, "Color");
    const colorArrCleaned = removeItemFromArray(getColor, "Color");
    // const colorFilter = filterOnlyNumericValues(getColor);
    // xử lý color
    const colorArr = trimArrayByArray(colorArrCleaned, articleStyleArr);
    // console.log("color arr: ", colorArr);


    // kiểm duyệt color, article / style, po
    const styleCheck = collapseIfSame(articleStyleArr);
    const colorCheck = collapseIfSame(colorArr);
    const POCheck = collapseIfSame(POArr);

    let styleName = "";
    let color = "";
    let PO = "";

    if (!isArray(styleCheck) || !isArray(colorCheck) || !isArray(POCheck)) {
      styleName = styleCheck;
      color = colorCheck;
      PO = POCheck;
    }
    else {
      styleName = articleStyleArr;
      color = colorArr;
      PO = POArr;
    }

    // COUNT
    const getCount = getValuesBelowTitleAuto(cellMatrix, "CTNS");
    // xử lý count
    const countArr = getCount
    console.log("count arr: ", countArr);
    const count = sumArray(countArr);
    // console.log("count arr: ", countArrCleaned);
    console.log("count: ", count);

    // TOTAL
    // total ( do QTY bị ẩn, nếu lấy theo TOTAL thì sẽ có thêm value là 'QTY' )

    const getTotals = getValuesBelowTitleAuto(cellMatrix, "TOTAL (PCS)");
    const totalArrFiltered = filterOnlyNumericValues(getTotals);
    // const totalArrCleaned = removeItemFromArray(getTotals, "QTY");
    // const totalArr = removeAfterString(totalArrCleaned);
    // const totalArrFiltered = filterOnlyNumericValues(totalArr);
    // console.log("total arr: ", totalArrFiltered);
    // const total = sumArray(totalArrFiltered.slice(0, totalArrFiltered.length - 1));
    // console.log("total: ", total);

    // const totalArr = trimArrayByArray(getTotals, colorArr);

    const total = sumArray(totalArrFiltered);
    console.log("total: ", total);
    // GROSS 
    const getGross = getValuesBelowTitleAuto(cellMatrix, "TTL G.W.");
    const grossArr = removeItemFromArray(getGross, "(KGS)");
    const grossFiltered = filterOnlyNumericValues(grossArr);
    // console.log("gorss arr: ", getGross);
    // const grossArr = removeItemFromArray(getGross, "WEIGHT");
    const gross = sumArray(grossFiltered)
    // console.log("gorss: ", gross);

    // NET
    const getNet = getValuesBelowTitleAuto(cellMatrix, "TTL N.W.");
    const netArr = removeItemFromArray(getNet, "(KGS)");
    const netFiltered = filterOnlyNumericValues(netArr);
    // const netArr = removeItemFromArray(getNet, "WEIGHT");
    const net = sumArray(netFiltered);
    // console.log("net: ", net);

    // VOLUME
    const getVolume = getValuesBelowTitleAuto(cellMatrix, "CBM");
    const volumeArr = getVolume;
    const volume = sumArray(volumeArr);
    // console.log("volume: ", volume);

    // CARTON FORMULA
    const getCartonDimension = getValuesBelowTitleAuto(cellMatrix, "Dims");
    const cartonDimensionArr = removeItemFromArray(getCartonDimension, "CBM\n(CM)");
    console.log("CARTON DIMENSION (CM): ", cartonDimensionArr);
    const cartonDimensionCleaned = trimArrayByArray(cartonDimensionArr, countArr);
    const cartonAndCount = mergeToObjectArray(cartonDimensionCleaned, countArr);
    console.log("carton and count: ", cartonAndCount);
    const cartonDimension = buildMeasFormula(cartonAndCount);
    console.log("carton dimension: ", cartonDimension);

    // ===================== GẮN SUMMARY =====================
    result.summary.push({
      sheetName,
      article: styleName,
      po: PO,
      color: color,
      countTotal: count,
      total: total,
      gross: gross,
      net: net,
      volume: volume,
      cartonFormula: cartonDimension,
    });

  });
  console.log(result.summary);

  return result;
}

// REISS
const POC_REISS = [
  "STYLE NO & NAME",
  "ORDER NO",
  "COLOUR",
  "ITEM",
  "CTNS",
  "TOTAL",
  "N.W.(KG)",
  "G.W.(KG)",
  "DIMENSION(CM)"
];

function previewPOC_REISS(workbook) {
  const result = {
    sheetNames: [],
    data: {},
    summary: [],
  };

  workbook.SheetNames.forEach((sheetName) => {

    const sheet = workbook.Sheets[sheetName];

    // 🛑 GUARD: nếu sheet KHÔNG thỏa ROSSIGNOL BULK → BỎ QUA
    if (!isSheetMatchTemplate(sheet, POC_REISS)) {
      console.log(`⏭️ Skip sheet: ${sheetName}`);
      return;
    }

    const cellMatrix = []; //  mỗi sheet 1 matrix

    const rawRows = xlsx.utils.sheet_to_json(sheet, { defval: "" });

    // 1️ lọc cột
    const columnCleaned = filterInvalidColumns(rawRows);

    // 2️ tìm cột COUNT
    const countColumnKey = findColumnKeyByTitle(columnCleaned, "count_)");

    const rowCleaned = columnCleaned.filter(row => {

      const rowText = normalize_(Object.values(row).join(" "));

      // 1️ GIỮ row có 
      // if (rowText.includes("total qty")) {
      //   return true;
      // }

      // if (rowText.includes("total gross")) {
      //   return true;
      // }

      // if (rowText.includes("total net")) {
      //   return true;
      // }

      // 2️ LOẠI row có TOTAL 
      if (rowText.includes("summary")) {
        return false;
      }

      // if(rowText.includes("ttl q'ty")){
      //   return false;
      // }

      if (hasAtLeastNEmptyCells(row, 15)) return false;

      if (isValidColumn([row], countColumnKey)) return true;
      return !isZeroInCountColumn(row, countColumnKey);
    });


    //  GIỮ NGUYÊN DATA CHO FRONTEND
    result.sheetNames.push(sheetName);
    result.data[sheetName] = rowCleaned.map((row) => ({
      SHEET: sheetName,
      ...row,
    }));


    // 4️ build cellMatrix cho sheet này
    const columnCells = extractColumnCellArrays(rowCleaned, sheetName);
    cellMatrix.push(...columnCells);



    // ===================== TÍNH TOÁN =====================

    console.log("------------------------------------------------------");

    // // // article / style
    const getArticleStyle = getValuesBelowTitleAuto(cellMatrix, "STYLE NO & Name");
    console.log("ARTICLE / STYLE: ", getArticleStyle);
    const articleStyleArr = removeFromStringValue(getArticleStyle, "STYLE NO");
    console.log("ARTICLE / STYLE: ", articleStyleArr);

    // // // po
    const getPO = getValuesBelowTitleAuto(cellMatrix, "ORDER NO");
    const POArr = trimArrayByArray(getPO, articleStyleArr);
    console.log("po: ", POArr);

    // color 
    const getColor = getValuesBelowTitleAuto(cellMatrix, "COLOUR");
    // const colorFilter = filterOnlyNumericValues(getColor);
    // xử lý color
    const colorArr = trimArrayByArray(getColor, articleStyleArr);
    // console.log("color arr: ", colorArr);


    // kiểm duyệt color, article / style, po
    const styleCheck = collapseIfSame(articleStyleArr);
    const colorCheck = collapseIfSame(colorArr);
    const POCheck = collapseIfSame(POArr);

    let styleName = "";
    let color = "";
    let PO = "";

    if (!isArray(styleCheck) || !isArray(colorCheck) || !isArray(POCheck)) {
      styleName = styleCheck;
      color = colorCheck;
      PO = POCheck;
    }
    else {
      styleName = articleStyleArr;
      color = colorArr;
      PO = POArr;
    }

    // COUNT
    const getCount = getValuesBelowTitleExact(cellMatrix, "CTNS");
    // xử lý count
    const countArr = getCount
    console.log("count arr: ", countArr);
    const count = sumArray(countArr.slice(0, countArr.length - 1));
    // console.log("count arr: ", countArrCleaned);
    console.log("count: ", count);

    // TOTAL
    // total ( do QTY bị ẩn, nếu lấy theo TOTAL thì sẽ có thêm value là 'QTY' )

    const getTotals = getValuesBelowTitleAutoForColorACTIVE_BULK(cellMatrix, "TOTAL");

    const totalCleaned = removeAfterStringValue(getTotals, "total");
    const totalArrFiltered = filterOnlyNumericValues(totalCleaned);
    console.log("total arr: ", totalArrFiltered);
    // const totalArrCleaned = removeItemFromArray(getTotals, "QTY");
    // const totalArr = removeAfterString(totalArrCleaned);
    // const totalArrFiltered = filterOnlyNumericValues(totalArr);
    // console.log("total arr: ", totalArrFiltered);
    // const total = sumArray(totalArrFiltered.slice(0, totalArrFiltered.length - 1));
    // console.log("total: ", total);

    // const totalArr = trimArrayByArray(getTotals, colorArr);

    const total = sumArray(totalArrFiltered.slice(1, totalArrFiltered.length - 1));
    console.log("total: ", total);
    // GROSS 
    const getGross = getValuesBelowTitleAuto(cellMatrix, "G.W.(KG)");
    console.log("gross arr: ", getGross);
    // const grossArr = removeItemFromArray(getGross, "(KGS)");
    const grossFiltered = filterOnlyNumericValues(getGross);
    // console.log("gorss arr: ", getGross);
    // const grossArr = removeItemFromArray(getGross, "WEIGHT");
    const gross = sumArray(grossFiltered.slice(0, grossFiltered.length - 1));
    // console.log("gorss: ", gross);

    // NET
    const getNet = getValuesBelowTitleAuto(cellMatrix, "N.W.(KG)");
    // const netArr = removeItemFromArray(getNet, "(KGS)");
    const netFiltered = filterOnlyNumericValues(getNet);
    // const netArr = removeItemFromArray(getNet, "WEIGHT");
    const net = sumArray(netFiltered.slice(0, netFiltered.length - 1));
    // console.log("net: ", net);

    // VOLUME
    const getVolume = getValuesBelowTitleAuto(cellMatrix, "CBM");
    const volumeArr = getVolume;
    const volume = sumArray(volumeArr);
    // console.log("volume: ", volume);

    // CARTON FORMULA
    const getCartonDimension = getValuesBelowTitleAuto(cellMatrix, "Dimension(cm)");
    const cartonDimensionFiltered = keepOnlyDimensionLxWxH(getCartonDimension);
    // const cartonDimensionArr = removeItemFromArray(getCartonDimension, "CBM\n(CM)");
    console.log("CARTON DIMENSION (CM): ", cartonDimensionFiltered);
    const cartonDimensionCleaned = trimArrayByArray(cartonDimensionFiltered, countArr);
    const cartonAndCount = mergeToObjectArray(cartonDimensionCleaned, countArr);
    console.log("carton and count: ", cartonAndCount);
    const cartonDimension = buildMeasFormula(cartonAndCount);
    console.log("carton dimension: ", cartonDimension);

    // ===================== GẮN SUMMARY =====================
    result.summary.push({
      sheetName,
      article: styleName,
      po: PO,
      color: color,
      countTotal: count,
      total: total,
      gross: gross,
      net: net,
      volume: volume,
      cartonFormula: cartonDimension,
    });
  });
  console.log(result.summary);

  return result;
}

// MAMMUT VSNT
const MAMMUT_VSNT = [
  "DETAILS AS PER MAMMUT ORDER:",
  "COL.+CODE",
  "CARTON TOTAL",
  "TOTAL",
  "TOTAL NW",
  "TOTAL GW",
  "TOTAL CBM",
  "CARTON",
];

function previewMAMMUT_VSNT(workbook) {
  const result = {
    sheetNames: [],
    data: {},
    summary: [],
  };

  // phần này xử lý cho ONE IBIS 021W:
  workbook.SheetNames.forEach((sheetName) => {


    const sheet = workbook.Sheets[sheetName];

    // 🛑 GUARD: nếu sheet KHÔNG thỏa ROSSIGNOL BULK → BỎ QUA
    if (!isSheetMatchTemplate(sheet, MAMMUT_VSNT)) {
      console.log(`⏭️ Skip sheet: ${sheetName}`);
      return;
    }

    const cellMatrix = []; //  mỗi sheet 1 matrix

    // PO
    const getPO_VSNT = extractPurchaseOrderNo(sheet, "DETAILS AS PER MAMMUT ORDER:")
    console.log("VSNT PO:", getPO_VSNT);

    // style 
    const getStyle_VSNT = extractPurchaseOrderNo(sheet, "STYLE NUMBER:");
    console.log("VSNT STYLE: ", getStyle_VSNT);


    const rawRows = xlsx.utils.sheet_to_json(sheet, { defval: "" });

    // 1️ lọc cột
    const columnCleaned = filterInvalidColumns(rawRows);

    // 2️ tìm cột COUNT
    const countColumnKey = findColumnKeyByTitle(columnCleaned, "count_");

    // 3️ lọc row COUNT = 0
    // const rowCleaned = columnCleaned.filter((row) => {
    //   if (isValidColumn([row], countColumnKey)) return true;
    //   return !isZeroInCountColumn(row, countColumnKey);
    // });

    // let palet = 0;
    // hãy viết kiểm tra chỗ này nếu có 1 Palet thì palet = palet + 20, cứ có một palet xuất hiện thì cột 20 được không

    const excludedTitles = [
      "PACKING",
      "TỔNG",
      "COLOR SUMMARY:",
      "+/-",
      "TOTAL DEEP TEAL",
      "SUBTOTAL",
      "PO SHEET",
      "PKL",
      "TOTAL BLACK",
    ];

    let palet = 0;

    // const rowCleaned = columnCleaned.filter(row => {
    //   const count = detectPaletValue(row);
    //   if (count > 0) {
    //     palet += 20;
    //     return false;
    //   }

    //   if (filterRowsByExcludedTitles([row], excludedTitles).length === 0) {
    //     return false;
    //   }

    //   if (isValidColumn([row], countColumnKey)) return true;
    //   return !isZeroInCountColumn(row, countColumnKey);
    // });
    const rowCleaned = columnCleaned.filter(row => {

      const count = detectPaletValue(row);
      if (count > 0) {
        palet += 25;
        return false;
      }

      // 🔴 FIX: lấy đúng cột đầu tiên
      const firstKey = Object.keys(row)[0];
      const firstCell = String(row[firstKey] || "")
        .toUpperCase()
        .trim();

      const secondKey = Object.keys(row)[1];
      const secondCell = String(row[secondKey] || "")
        .toUpperCase()
        .trim();

      if (firstCell.includes("TOTAL") || firstCell.includes("TOATL") || firstCell.includes("NHÓM") || firstCell.includes("MÃ HÀNG") || firstCell.includes("PALLET SỐ") || firstCell.includes("PALET SỐ") || firstCell === "") {
        return false;
      }

      if (secondCell.includes("TOTAL") || secondCell.includes("TOATL")) {
        return false;
      }



      if (filterRowsByExcludedTitles([row], excludedTitles).length === 0) {
        return false;
      }

      if (isValidColumn([row], countColumnKey)) return true;
      return !isZeroInCountColumn(row, countColumnKey);
    });


    //  GIỮ NGUYÊN DATA CHO FRONTEND
    result.sheetNames.push(sheetName);
    result.data[sheetName] = rowCleaned.map((row) => ({
      SHEET: sheetName,
      ...row,
    }));

    console.log("palet: ", palet);

    // 4️ build cellMatrix cho sheet này
    const columnCells = extractColumnCellArrays(rowCleaned, sheetName);
    cellMatrix.push(...columnCells);

    // ===================== TÍNH TOÁN =====================

    console.log("-------------------------------------------------------------------------------");

    // xử lý phần PO và style
    const getPO_VIETDUC = getValuesBelowTitleAuto(cellMatrix, "Order Number Style No. + description");
    console.log("VIET DUC PO:", getPO_VIETDUC[1]);

    const getStyle_VIETDUC = getValuesBelowTitleAuto(cellMatrix, "Order Number Style No. + description");
    console.log("VIET DUC STYLE: ", getStyle_VIETDUC[1]);

    let PO = "";
    let style = "";

    // po
    if (getPO_VIETDUC.length >= 1) {
      PO = getPO_VIETDUC[1];
    }
    else {
      PO = getPO_VSNT;
    }

    // style
    if (getStyle_VIETDUC.length >= 1) {
      style = getStyle_VIETDUC[1];;
    }
    else {
      style = getStyle_VSNT;
    }

    // color 
    const getColor = getValuesBelowTitleAuto(cellMatrix, "Col.+Code");
    const colorArrFiltered = removeItemFromArray(getColor, "Col.+Code");
    const colorArrFiltered_first = removeAfterStringValue(colorArrFiltered, "SUMMARY");
    const colorArrFiltered_second = removeAfterStringValue(colorArrFiltered_first, "PO");
    const colorArrFiltered_third = removeAfterStringValue(colorArrFiltered_second, "TOTAL CARTON");
    const colorArrFiltered_fourth = removeAfterStringValue(colorArrFiltered_third, "SIZE");
    const colorArrFiltered_fifth = removeAfterStringValue(colorArrFiltered_fourth, "TOTAL");
    const colorArr = colorArrFiltered_fifth;
    console.log("color arr: ", colorArr);

    // // article / style
    // const getArticleStyle = getValuesBelowTitleAutoForColorACTIVE_BULK(cellMatrix, "ARTICLE / STYLE");
    // const articleStyleArr1 = removeAfterStringValue(getArticleStyle, "TOTAL")
    // const articleStyleArr2 = removeAfterStringValue(articleStyleArr1, "ARTICLE / STYLE")
    // const articleStyleArr3 = removeAfterStringValue(articleStyleArr2, "PO")
    // const articleStyleArr4 = removeAfterStringValue(articleStyleArr3, "COLOR")
    // const articleStyleArr5 = removeAfterStringValue(articleStyleArr4, "SHADE ")
    // const articleStyleArr = removeAfterStringValue(articleStyleArr5, "total")

    // console.log("ARTICLE / STYLE: ", articleStyleArr);


    // kiểm duyệt color, article / style, po
    // const styleCheck = collapseIfSame(articleStyleArr);
    // const colorCheck = collapseIfSame(colorArr);

    // let styleName = "";
    // let color = "";

    // if (!isArray(styleCheck) || !isArray(colorCheck)) {
    //   styleName = styleCheck;
    //   color = colorCheck;
    // }
    // else {
    //   styleName = articleStyleArr;
    //   color = colorArr;
    // }


    // COUNT
    const getCount = getValuesBelowTitleAuto(cellMatrix, "Carton total");
    const countArr = filterOnlyNumericValues(getCount);
    // xử lý count
    const getCartonNumber = getValuesBelowTitleAuto(cellMatrix, "Carton Number");
    const cartonNumberCleaned = filterOnlyNumericValues(getCartonNumber);
    const countArrCleaned = trimArrayByArray(countArr, cartonNumberCleaned);
    const count = sumArray(countArrCleaned);
    console.log("count arr: ", countArrCleaned);
    console.log("count: ", count);

    // TOTAL
    // total ( do QTY bị ẩn, nếu lấy theo TOTAL thì sẽ có thêm value là 'QTY' )
    const getTotals = getValuesBelowTitleExact(cellMatrix, "TOTAL");
    // const totalArrFiltered = filterOnlyNumericValues(getTotals);
    const totalArrCleaned = removeItemFromArray(getTotals, "QTY");
    const totalArr = removeAfterString(totalArrCleaned);
    const totalArrFiltered = filterOnlyNumericValues(totalArr);
    console.log("total arr: ", totalArrFiltered);
    const total = sumArray(totalArrFiltered);
    console.log("total: ", total);

    // GROSS 
    const getGross = getValuesBelowTitleAuto(cellMatrix, "Total GW");
    console.log("gorss arr: ", getGross);
    // const grossArr = removeItemFromArray(getGross, "WEIGHT");
    const gross = sumArray(getGross) + palet;
    console.log("gorss: ", gross);

    // NET
    const getNet = getValuesBelowTitleAuto(cellMatrix, "Total NW");
    // const netArr = removeItemFromArray(getNet, "WEIGHT");
    const net = sumArray(getNet);
    console.log("net: ", net);

    // VOLUME
    const getVolume = getValuesBelowTitleAuto(cellMatrix, "Total CBM");
    const volumeArr = getVolume;
    const volume = sumArray(volumeArr);
    console.log("volume: ", volume);

    // CARTON FORMULA
    const getCartonDimension = getValuesBelowTitleAutoForCartonDimension_MAMMUT(cellMatrix, "Carton");
    console.log("cbm: ", getCartonDimension);
    const cartonDimensionArr = buildCartonDimensionFormulas(getCartonDimension);
    console.log("CARTON DIMENSION (CM): ", cartonDimensionArr);
    const cartonAndCount = mergeToObjectArray(cartonDimensionArr, countArrCleaned);
    console.log("carton and count: ", cartonAndCount);
    const cartonDimension = buildMeasFormula(cartonAndCount);
    console.log("carton dimension: ", cartonDimension);


    // ===================== GẮN SUMMARY =====================
    result.summary.push({
      sheetName,
      article: style,
      po: PO,
      color: colorArr,
      countTotal: count,
      total: total,
      gross: gross,
      net: net,
      volume: volume,
      cartonFormula: cartonDimension,
    });
  });
  console.log(result.summary);
  return result;
}

// MAMMUT GỘP PO
const MAMMUT_GOP_PO = [
  "DETAILS AS PER MAMMUT ORDER:",
  "CARTON TOTAL",
  "TOTAL",
  "TOTAL NW",
  "TOTAL GW",
  "TOTAL CBM",
  "CARTON",
];
function previewMAMMUT_GOP_PO(workbook) {
  const result = {
    sheetNames: [],
    data: {},
    summary: [],
  };

  // phần này xử lý cho ONE IBIS 021W:
  workbook.SheetNames.forEach((sheetName) => {

    const sheet = workbook.Sheets[sheetName];

    // 🛑 GUARD: nếu sheet KHÔNG thỏa ROSSIGNOL BULK → BỎ QUA
    if (!isSheetMatchTemplate(sheet, MAMMUT_GOP_PO)) {
      console.log(`⏭️ Skip sheet: ${sheetName}`);
      return;
    }

    const cellMatrix = []; //  mỗi sheet 1 matrix


    // PO
    const getPO_VSNT = extractPurchaseOrderNo(sheet, "DETAILS AS PER MAMMUT ORDER:")
    console.log("VSNT PO:", getPO_VSNT);

    // style 
    const getStyle_VSNT = extractPurchaseOrderNo(sheet, "STYLE NUMBER:");
    console.log("VSNT STYLE: ", getStyle_VSNT);


    const rawRows = xlsx.utils.sheet_to_json(sheet, { defval: "" });

    // 1️ lọc cột
    const columnCleaned = filterInvalidColumns(rawRows);

    // 2️ tìm cột COUNT
    const countColumnKey = findColumnKeyByTitle(columnCleaned, "count_");

    // 3️ lọc row COUNT = 0
    // const rowCleaned = columnCleaned.filter((row) => {
    //   if (isValidColumn([row], countColumnKey)) return true;
    //   return !isZeroInCountColumn(row, countColumnKey);
    // });

    // let palet = 0;
    // hãy viết kiểm tra chỗ này nếu có 1 Palet thì palet = palet + 20, cứ có một palet xuất hiện thì cột 20 được không

    const excludedTitles = [
      "PACKING",
      "TỔNG",
      "COLOR SUMMARY:",
      "+/-",
      "TOTAL DEEP TEAL",
      "SUBTOTAL",
      "PO SHEET",
      "PKL",
      "TOTAL BLACK",
    ];

    let palet = 0;

    // const rowCleaned = columnCleaned.filter(row => {
    //   const count = detectPaletValue(row);
    //   if (count > 0) {
    //     palet += 20;
    //     return false;
    //   }

    //   if (filterRowsByExcludedTitles([row], excludedTitles).length === 0) {
    //     return false;
    //   }

    //   if (isValidColumn([row], countColumnKey)) return true;
    //   return !isZeroInCountColumn(row, countColumnKey);
    // });
    const rowCleaned = columnCleaned.filter(row => {

      const count = detectPaletValue(row);
      if (count > 0) {
        palet += 25;
        return false;
      }

      // 🔴 FIX: lấy đúng cột đầu tiên
      const firstKey = Object.keys(row)[0];
      const firstCell = String(row[firstKey] || "")
        .toUpperCase()
        .trim();

      const secondKey = Object.keys(row)[1];
      const secondCell = String(row[secondKey] || "")
        .toUpperCase()
        .trim();

      if (firstCell.includes("TOTAL") || firstCell.includes("TOATL") || firstCell.includes("NHÓM") || firstCell.includes("MÃ HÀNG") || firstCell.includes("PALLET SỐ") || firstCell.includes("PALET SỐ") || firstCell === "") {
        return false;
      }

      if (secondCell.includes("TOTAL") || secondCell.includes("TOATL")) {
        return false;
      }



      if (filterRowsByExcludedTitles([row], excludedTitles).length === 0) {
        return false;
      }

      if (isValidColumn([row], countColumnKey)) return true;
      return !isZeroInCountColumn(row, countColumnKey);
    });


    //  GIỮ NGUYÊN DATA CHO FRONTEND
    result.sheetNames.push(sheetName);
    result.data[sheetName] = rowCleaned.map((row) => ({
      SHEET: sheetName,
      ...row,
    }));

    console.log("palet: ", palet);

    // 4️ build cellMatrix cho sheet này
    const columnCells = extractColumnCellArrays(rowCleaned, sheetName);
    cellMatrix.push(...columnCells);

    // ===================== TÍNH TOÁN =====================

    console.log("-------------------------------------------------------------------------------");

    // xử lý phần PO và style
    const getPO_VIETDUC = getValuesBelowTitleAuto(cellMatrix, "Order Number Style No. + description");
    console.log("VIET DUC PO:", getPO_VIETDUC[1]);

    const getStyle_VIETDUC = getValuesBelowTitleAuto(cellMatrix, "Order Number Style No. + description");
    console.log("VIET DUC STYLE: ", getStyle_VIETDUC[1]);

    let PO = "";
    let style = "";

    // po
    if (getPO_VIETDUC.length >= 1) {
      PO = getPO_VIETDUC[1];
    }
    else {
      PO = getPO_VSNT;
    }

    // style
    if (getStyle_VIETDUC.length >= 1) {
      style = getStyle_VIETDUC[1];;
    }
    else {
      style = getStyle_VSNT;
    }

    // color 
    // const getColor = getValuesBelowTitleAuto(cellMatrix, "Col.+Code");
    // const colorArrFiltered = removeItemFromArray(getColor, "Col.+Code");
    // const colorArrFiltered_first = removeAfterStringValue(colorArrFiltered, "SUMMARY");
    // const colorArrFiltered_second = removeAfterStringValue(colorArrFiltered_first, "PO");
    // const colorArrFiltered_third = removeAfterStringValue(colorArrFiltered_second, "TOTAL CARTON");
    // const colorArrFiltered_fourth = removeAfterStringValue(colorArrFiltered_third, "SIZE");
    // const colorArrFiltered_fifth = removeAfterStringValue(colorArrFiltered_fourth, "TOTAL");
    // const colorArr = colorArrFiltered_fifth;
    // console.log("color arr: ", colorArr);

    const getColor = cellMatrix[0].cells;
    const colorArr = removeItemsByKeywords(getColor, ["Style number:", "Style name:", "Details as per MAMMUT order:", "PO-"])
    console.log("cell matrix 0: ", colorArr);

    // // article / style
    // const getArticleStyle = getValuesBelowTitleAutoForColorACTIVE_BULK(cellMatrix, "ARTICLE / STYLE");
    // const articleStyleArr1 = removeAfterStringValue(getArticleStyle, "TOTAL")
    // const articleStyleArr2 = removeAfterStringValue(articleStyleArr1, "ARTICLE / STYLE")
    // const articleStyleArr3 = removeAfterStringValue(articleStyleArr2, "PO")
    // const articleStyleArr4 = removeAfterStringValue(articleStyleArr3, "COLOR")
    // const articleStyleArr5 = removeAfterStringValue(articleStyleArr4, "SHADE ")
    // const articleStyleArr = removeAfterStringValue(articleStyleArr5, "total")

    // console.log("ARTICLE / STYLE: ", articleStyleArr);


    // kiểm duyệt color, article / style, po
    // const styleCheck = collapseIfSame(articleStyleArr);
    // const colorCheck = collapseIfSame(colorArr);

    // let styleName = "";
    // let color = "";

    // if (!isArray(styleCheck) || !isArray(colorCheck)) {
    //   styleName = styleCheck;
    //   color = colorCheck;
    // }
    // else {
    //   styleName = articleStyleArr;
    //   color = colorArr;
    // }


    // COUNT
    const getCount = getValuesBelowTitleAuto(cellMatrix, "Carton total");
    const countArr = filterOnlyNumericValues(getCount);
    // xử lý count
    const getCartonNumber = getValuesBelowTitleAuto(cellMatrix, "Carton Number");
    const cartonNumberCleaned = filterOnlyNumericValues(getCartonNumber);
    const countArrCleaned = trimArrayByArray(countArr, cartonNumberCleaned);
    const count = sumArray(countArrCleaned);
    console.log("count arr: ", countArrCleaned);
    console.log("count: ", count);

    // TOTAL
    // total ( do QTY bị ẩn, nếu lấy theo TOTAL thì sẽ có thêm value là 'QTY' )
    const getTotals = getValuesBelowTitleExact(cellMatrix, "TOTAL");
    // const totalArrFiltered = filterOnlyNumericValues(getTotals);
    // const totalArrCleaned = removeItemFromArray(getTotals, "TOTAL");
    // const totalArr = removeAfterString(totalArrCleaned);
    const totalArrFiltered = filterOnlyNumericValues(getTotals);
    console.log("total arr: ", totalArrFiltered);
    const total = sumArray(totalArrFiltered);
    console.log("total: ", total);

    // GROSS 
    const getGross = getValuesBelowTitleAuto(cellMatrix, "Total GW");
    console.log("gorss arr: ", getGross);
    // const grossArr = removeItemFromArray(getGross, "WEIGHT");
    const gross = sumArray(getGross) + palet;
    console.log("gorss: ", gross);

    // NET
    const getNet = getValuesBelowTitleAuto(cellMatrix, "Total NW");
    // const netArr = removeItemFromArray(getNet, "WEIGHT");
    const net = sumArray(getNet);
    console.log("net: ", net);

    // VOLUME
    const getVolume = getValuesBelowTitleAuto(cellMatrix, "Total CBM");
    const volumeArr = getVolume;
    const volume = sumArray(volumeArr);
    console.log("volume: ", volume);

    // CARTON FORMULA
    const getCartonDimension = getValuesBelowTitleAutoForCartonDimension_MAMMUT(cellMatrix, "Carton");
    console.log("cbm: ", getCartonDimension);
    const cartonDimensionArr = buildCartonDimensionFormulas(getCartonDimension);
    console.log("CARTON DIMENSION (CM): ", cartonDimensionArr);
    const cartonAndCount = mergeToObjectArray(cartonDimensionArr, countArrCleaned);
    console.log("carton and count: ", cartonAndCount);
    const cartonDimension = buildMeasFormula(cartonAndCount);
    console.log("carton dimension: ", cartonDimension);


    // ===================== GẮN SUMMARY =====================
    result.summary.push({
      sheetName,
      article: style,
      po: PO,
      color: colorArr,
      countTotal: count,
      total: total,
      gross: gross,
      net: net,
      volume: volume,
      cartonFormula: cartonDimension,
    });
  });
  console.log(result.summary);
  return result;
}


// ROSSIGNOL BULK

const ROSSIGNOL__BULK_REQUIRED_HEADERS = [
  "PRODUCT CODE",
  "CUSTOMER PO#",
  "NB+COLOR",
  "NB",
  "TOTAL",
  "NET",
  "GROSS",
  "VOLUME (CBM)",
  "CARTON DIMENSION (CM)"
];

function previewROSSIGNOL_BULK(workbook) {
  const result = {
    sheetNames: [],
    data: {},
    summary: [],
  };

  // phần này xử lý cho ONE IBIS 021W:
  workbook.SheetNames.forEach((sheetName) => {


    const sheet = workbook.Sheets[sheetName];

    // 🛑 GUARD: nếu sheet KHÔNG thỏa ROSSIGNOL BULK → BỎ QUA
    if (!isSheetMatchTemplate(sheet, ROSSIGNOL__BULK_REQUIRED_HEADERS)) {
      console.log(`⏭️ Skip sheet: ${sheetName}`);
      return;
    }

    const cellMatrix = []; //  mỗi sheet 1 matrix

    // // PO
    // const purchase_order_no = extractPurchaseOrderNo(sheet, "ORDER NO:");
    // console.log("Order no:", purchase_order_no);

    const rawRows = xlsx.utils.sheet_to_json(sheet, { defval: "" });

    // 1️ lọc cột
    const columnCleaned = filterInvalidColumns(rawRows);

    // 2️ tìm cột COUNT
    const countColumnKey = findColumnKeyByTitle(columnCleaned, "count_");

    // 3️ lọc row COUNT = 0
    // const rowCleaned = columnCleaned.filter((row) => {
    //   if (isValidColumn([row], countColumnKey)) return true;
    //   return !isZeroInCountColumn(row, countColumnKey);
    // });

    // let palet = 0;
    // hãy viết kiểm tra chỗ này nếu có 1 Palet thì palet = palet + 20, cứ có một palet xuất hiện thì cột 20 được không

    const excludedTitles = [
      "grand total",
      "TỔNG PO",
    ];

    let palet = 0;

    // 🔴 CẮT TỪ COLOR SUMMARY TRỞ ĐI
    const rowsAfterCut = cutRowsFromTitleInFirstColumn(
      columnCleaned,
      "COLOR SUMMARY:"
    );

    const rowCleaned = rowsAfterCut.filter(row => {
      if (filterRowsByExcludedTitles([row], excludedTitles).length === 0) {
        return false;
      }

      if (isValidColumn([row], countColumnKey)) return true;
      return !isZeroInCountColumn(row, countColumnKey);
    });


    //  GIỮ NGUYÊN DATA CHO FRONTEND
    result.sheetNames.push(sheetName);
    result.data[sheetName] = rowCleaned.map((row) => ({
      SHEET: sheetName,
      ...row,
    }));


    console.log("palet: ", palet);

    // 4️ build cellMatrix cho sheet này
    const columnCells = extractColumnCellArrays(rowCleaned, sheetName);
    cellMatrix.push(...columnCells);

    // ===================== TÍNH TOÁN =====================

    console.log("------------------------------------------------------");


    // color 
    const getColor = getValuesBelowTitleAutoForColorACTIVE_BULK(cellMatrix, "NB+COLOR");
    const colorFiltered = removeItemFromArray(getColor, "NB+COLOR");
    const colorFiltered1 = removeAfterStringValue(colorFiltered, "Product code");
    const colorFiltered2 = removeAfterStringValue(colorFiltered1, "MODEL NAME");
    const colorFiltered3 = removeAfterStringValue(colorFiltered2, "TỔNG PO");
    const colorFiltered4 = removeAfterStringValue(colorFiltered3, "SHADE");
    const colorFiltered5 = removeAfterStringValue(colorFiltered4, "NB+COLOR");
    const colorFiltered6 = removeAfterStringValue(colorFiltered5, "PO");
    const colorArr = colorFiltered6;
    console.log("color arr: ", colorArr);

    // article / style
    const getArticleStyle = getValuesBelowTitleAutoForColorACTIVE_BULK(cellMatrix, "Product code");
    const articleStyleArr1 = removeItemFromArray(getArticleStyle, "Product code");
    const articleStyleArr2 = removeAfterStringValue(articleStyleArr1, "MODEL NAME")
    const articleStyleArr3 = removeAfterStringValue(articleStyleArr2, "TỔNG PO")
    const articleStyleArr4 = removeAfterStringValue(articleStyleArr3, "NB+COLOR")
    const articleStyleArr5 = removeAfterStringValue(articleStyleArr4, "SHADE ")
    const articleStyleArr = removeAfterStringValue(articleStyleArr5, "PO");

    console.log("ARTICLE / STYLE: ", articleStyleArr);

    // po 
    const getPO = getValuesBelowTitleAuto(cellMatrix, "Customer PO#");
    const POFiltered = removeItemFromArray(getPO, "Customer PO#");
    const POFiltered1 = removeAfterStringValue(POFiltered, "Product code");
    const POFiltered2 = removeAfterStringValue(POFiltered1, "PO");
    const POFiltered3 = removeAfterStringValue(POFiltered2, "TỔNG PO");
    const POFiltered4 = removeAfterStringValue(POFiltered3, "NB+COLOR");
    const POFiltered5 = removeAfterStringValue(POFiltered4, "SHADE");
    const POFiltered6 = removeAfterStringValue(POFiltered5, "MODEL NAME");
    const POArr = POFiltered6;
    console.log("PO ARR: ", POArr);


    // kiểm duyệt color, article / style, po
    const styleCheck = collapseIfSame(articleStyleArr);
    const colorCheck = collapseIfSame(colorArr);
    const POCheck = collapseIfSame(POArr);

    let styleName = "";
    let color = "";
    let PO = ""

    if (!isArray(styleCheck) || !isArray(colorCheck) || !isArray(POCheck)) {
      styleName = styleCheck;
      color = colorCheck;
      PO = POCheck;
    }
    else {
      styleName = articleStyleArr;
      color = colorArr;
      PO = POArr;
    }

    // COUNT
    const getCount = getValuesBelowTitleAuto(cellMatrix, "NB");
    const countArrFiltered = removeAfterStringValue(getCount, "Product code");
    const countArr = filterOnlyNumericValues(countArrFiltered);
    const count = sumArray(countArr);
    // console.log("count arr: ", countArrCleaned);
    console.log("count: ", count);

    // TOTAL
    // total ( do QTY bị ẩn, nếu lấy theo TOTAL thì sẽ có thêm value là 'QTY' )
    const getTotals = getValuesBelowTitleExact(cellMatrix, "TOTAL");
    // const totalArrFiltered = filterOnlyNumericValues(getTotals);
    const totalArrCleaned = removeItemFromArray(getTotals, "QTY");
    const totalArrCleaned1 = removeItemFromArray(totalArrCleaned, "TOTAL");
    const totalArrCleaned2 = removeAfterStringValue(totalArrCleaned1, "total");
    const totalArrFiltered = totalArrCleaned2;
    const totalArr = filterOnlyNumericValues(totalArrFiltered);
    console.log("total arr: ", totalArr);
    const total = sumArray(totalArr);
    console.log("total: ", total);

    // GROSS 1
    const getGross = getValuesBelowTitleAuto(cellMatrix, "GROSS");
    console.log("gorss arr: ", getGross);
    const grossArr = removeItemFromArray(getGross, "WEIGHT");
    const gross = sumArray(grossArr) + palet;
    console.log("gorss: ", gross);

    // NET
    const getNet = getValuesBelowTitleAuto(cellMatrix, "NET");
    const netArr = removeItemFromArray(getNet, "WEIGHT");
    const net = sumArray(netArr);
    console.log("net: ", net);

    // VOLUME
    const getVolume = getValuesBelowTitleAuto(cellMatrix, "VOLUME (CBM)");
    const volumeArr = getVolume;
    const volume = sumArray(volumeArr);
    console.log("volume: ", volume);

    // CARTON FORMULA
    const getCartonDimension = getValuesBelowTitleAuto(cellMatrix, "CARTON DIMENSION (CM)");
    const cartonDimensionArr = removeItemFromArray(getCartonDimension, "CARTON DIMENSION (CM)")
    console.log("CARTON DIMENSION (CM): ", cartonDimensionArr);
    const cartonAndCount = mergeToObjectArray(cartonDimensionArr, countArr);
    console.log("carton and count: ", cartonAndCount);
    const cartonDimension = buildMeasFormula(cartonAndCount);
    console.log("carton dimension: ", cartonDimension);

    // ===================== GẮN SUMMARY =====================
    result.summary.push({
      sheetName,
      article: styleName,
      po: PO,
      color: color,
      countTotal: count,
      total: total,
      gross: gross,
      net: net,
      volume: volume,
      cartonFormula: cartonDimension,
    });

  });
  console.log(result.summary);

  return result;
}

// ROSSIGNOL SMS
const ROSSIGNOL__SMS_REQUIRED_HEADERS = [
  "PRODUCT CODE SMS",
  "PO",
  "COLOR",
  "COUNT",
  "TOTAL QTY",
  "GW (KGS)",
  "CBM",
  "CARTON DIMENSION (CM)"
];

function previewROSSIGNOL_SMS(workbook) {
  const result = {
    sheetNames: [],
    data: {},
    summary: [],
  };

  // phần này xử lý cho ONE IBIS 021W:
  workbook.SheetNames.forEach((sheetName) => {


    const sheet = workbook.Sheets[sheetName];

    // 🛑 GUARD: nếu sheet KHÔNG thỏa ROSSIGNOL BULK → BỎ QUA
    if (!isSheetMatchTemplate(sheet, ROSSIGNOL__SMS_REQUIRED_HEADERS)) {
      console.log(`⏭️ Skip sheet: ${sheetName}`);
      return;
    }
    const cellMatrix = []; //  mỗi sheet 1 matrix

    // // PO
    // const purchase_order_no = extractPurchaseOrderNo(sheet, "ORDER NO:");
    // console.log("Order no:", purchase_order_no);

    const rawRows = xlsx.utils.sheet_to_json(sheet, { defval: "" });

    // 1️ lọc cột
    const columnCleaned = filterInvalidColumns(rawRows);

    // 2️ tìm cột COUNT
    const countColumnKey = findColumnKeyByTitle(columnCleaned, "count_");

    // 3️ lọc row COUNT = 0
    // const rowCleaned = columnCleaned.filter((row) => {
    //   if (isValidColumn([row], countColumnKey)) return true;
    //   return !isZeroInCountColumn(row, countColumnKey);
    // });

    // let palet = 0;
    // hãy viết kiểm tra chỗ này nếu có 1 Palet thì palet = palet + 20, cứ có một palet xuất hiện thì cột 20 được không

    const excludedTitles = [
      "grand total",
      "TỔNG PO",
    ];

    let palet = 0;

    // 🔴 CẮT TỪ COLOR SUMMARY TRỞ ĐI
    const rowsAfterCut = cutRowsFromTitleInFirstColumn(
      columnCleaned,
      "COLOR SUMMARY:"
    );

    const rowCleaned = rowsAfterCut.filter(row => {

      const firstKey = Object.keys(row)[1];
      const firstCell = String(row[firstKey] || "")
        .toUpperCase()
        .trim();

      if (firstCell === "") {
        return false;
      }

      if (filterRowsByExcludedTitles([row], excludedTitles).length === 0) {
        return false;
      }

      if (isValidColumn([row], countColumnKey)) return true;
      return !isZeroInCountColumn(row, countColumnKey);
    });


    //  GIỮ NGUYÊN DATA CHO FRONTEND
    result.sheetNames.push(sheetName);
    result.data[sheetName] = rowCleaned.map((row) => ({
      SHEET: sheetName,
      ...row,
    }));


    // console.log("palet: ", palet);

    // 4️ build cellMatrix cho sheet này
    const columnCells = extractColumnCellArrays(rowCleaned, sheetName);
    cellMatrix.push(...columnCells);

    // ===================== TÍNH TOÁN =====================

    console.log("------------------------------------------------------");


    // color 
    const getColor = getValuesBelowTitleAuto(cellMatrix, "Color");
    const colorFiltered = removeItemFromArray(getColor, "Color");
    // const colorFiltered1 = removeAfterStringValue(colorFiltered, "Product code");
    // const colorFiltered2 = removeAfterStringValue(colorFiltered1, "MODEL NAME");
    // const colorFiltered3 = removeAfterStringValue(colorFiltered2, "TỔNG PO");
    // const colorFiltered4 = removeAfterStringValue(colorFiltered3, "SHADE");
    // const colorFiltered5 = removeAfterStringValue(colorFiltered4, "NB+COLOR");
    // const colorFiltered6 = removeAfterStringValue(colorFiltered5, "PO");
    const colorArr = colorFiltered;
    console.log("color arr: ", colorArr);

    // article / style
    const getArticleStyle = getValuesBelowTitleAutoForColorACTIVE_BULK(cellMatrix, "Product code");
    const articleStyleArr1 = removeItemFromArray(getArticleStyle, "Product code");
    const articleStyleArr2 = removeAfterStringValue(articleStyleArr1, "MODEL NAME")
    const articleStyleArr3 = removeAfterStringValue(articleStyleArr2, "TỔNG PO")
    const articleStyleArr4 = removeAfterStringValue(articleStyleArr3, "NB+COLOR")
    const articleStyleArr5 = removeAfterStringValue(articleStyleArr4, "SHADE ")
    const articleStyleArr = removeAfterStringValue(articleStyleArr5, "PO");

    console.log("ARTICLE / STYLE: ", articleStyleArr);

    // po 
    const getPO = extractPurchaseOrderNo(sheet, "PO");
    // const POFiltered = removeItemFromArray(getPO, "PO");
    // const POFiltered1 = removeAfterStringValue(POFiltered, "Product code");
    // const POFiltered2 = removeAfterStringValue(POFiltered1, "PO");
    // const POFiltered3 = removeAfterStringValue(POFiltered2, "TỔNG PO");
    // const POFiltered4 = removeAfterStringValue(POFiltered3, "NB+COLOR");
    // const POFiltered5 = removeAfterStringValue(POFiltered4, "SHADE");
    // const POFiltered6 = removeAfterStringValue(POFiltered5, "MODEL NAME");
    const PO = getPO;
    console.log("PO ARR: ", PO);


    // kiểm duyệt color, article / style, po
    const styleCheck = collapseIfSame(articleStyleArr);
    const colorCheck = collapseIfSame(colorArr);
    // const POCheck = collapseIfSame(POArr);

    let styleName = "";
    let color = "";
    // let PO = "";

    if (!isArray(styleCheck) || !isArray(colorCheck)) {
      styleName = styleCheck;
      color = colorCheck;
      // PO = POCheck;
    }
    else {
      styleName = articleStyleArr;
      color = colorArr;
      // PO = POArr;
    }

    // COUNT
    const getCount = getValuesBelowTitleAuto(cellMatrix, "Count");
    // const countArrFiltered = removeAfterStringValue(getCount, "Product code");
    const countArr = filterOnlyNumericValues(getCount);
    const count = sumArray(countArr);
    // console.log("count arr: ", countArrCleaned);
    console.log("count: ", count);

    // TOTAL
    // total ( do QTY bị ẩn, nếu lấy theo TOTAL thì sẽ có thêm value là 'QTY' )
    const getTotals = getValuesBelowTitleAuto(cellMatrix, "TOTAL QTY");
    // const totalArrFiltered = filterOnlyNumericValues(getTotals);
    const totalArrCleaned = removeItemFromArray(getTotals, "TOTAL QTY");
    // const totalArrCleaned1 = removeItemFromArray(totalArrCleaned, "TOTAL");
    // const totalArrCleaned2 = removeAfterStringValue(totalArrCleaned1, "total");
    const totalArrFiltered = totalArrCleaned;
    const totalArr = filterOnlyNumericValues(totalArrFiltered);
    console.log("total arr: ", totalArr);
    const total = sumArray(totalArr);
    console.log("total: ", total);

    // GROSS 1
    const getGross = getValuesBelowTitleAuto(cellMatrix, "GW (KGS)");
    console.log("gorss arr: ", getGross);
    const grossArr = removeItemFromArray(getGross, "WEIGHT");
    const gross = sumArray(grossArr) + palet;
    console.log("gorss: ", gross);

    // NET
    let getNet = [];

    const netWithSpace = getValuesBelowTitleAuto(cellMatrix, "NW (KGS)");

    if (Array.isArray(netWithSpace) && netWithSpace.length > 0) {
      getNet = netWithSpace;
    } else {
      getNet = getValuesBelowTitleAuto(cellMatrix, "NW(KGS)");
    }


    const netArr = removeItemFromArray(getNet, "WEIGHT");
    const net = sumArray(netArr);
    console.log("net: ", net);

    // VOLUME
    const getVolume = getValuesBelowTitleAuto(cellMatrix, "CBM");
    const volumeArr = getVolume;
    const volume = sumArray(volumeArr);
    console.log("volume: ", volume);

    // CARTON FORMULA
    const getCartonDimension = getValuesBelowTitleAuto(cellMatrix, "CARTON DIMENSION (CM)");
    const cartonDimensionArr = removeItemFromArray(getCartonDimension, "CARTON DIMENSION (CM)")
    console.log("CARTON DIMENSION (CM): ", cartonDimensionArr);
    const cartonAndCount = mergeToObjectArray(cartonDimensionArr, countArr);
    console.log("carton and count: ", cartonAndCount);
    const cartonDimension = buildMeasFormula(cartonAndCount);
    console.log("carton dimension: ", cartonDimension);

    // ===================== GẮN SUMMARY =====================
    result.summary.push({
      sheetName,
      article: styleName,
      po: PO,
      color: color,
      countTotal: count,
      total: total,
      gross: gross,
      net: net,
      volume: volume,
      cartonFormula: cartonDimension,
    });

  });
  console.log(result.summary);

  return result;
}

// EDDIE BAUER BULK

const EDDIE_BAUER_BULK_REQUIRED_HEADERS = [
  "STYLE NO.",
  "PURCHASE ORDER",
  "COLOR",
  "CARTON NO",
  "TOTAL",
  "TOTAL NW",
  "TOTAL GW",
  "TOTAL CBM"
];

function previewEDDIE_BAUER_BULK(workbook) {
  const result = {
    sheetNames: [],
    data: {},
    summary: [],
  };

  // phần này xử lý cho ONE IBIS 021W:
  workbook.SheetNames.forEach((sheetName) => {

    const sheet = workbook.Sheets[sheetName];

    // 🛑 GUARD: nếu sheet KHÔNG thỏa ROSSIGNOL BULK → BỎ QUA
    if (!isSheetMatchTemplate(sheet, EDDIE_BAUER_BULK_REQUIRED_HEADERS)) {
      console.log(`⏭️ Skip sheet: ${sheetName}`);
      return;
    }
    const cellMatrix = []; //  mỗi sheet 1 matrix

    //  PO
    const PO = extractPurchaseOrderNo(sheet, "PURCHASE ORDER");
    console.log("Order no:", PO);
    //style no
    const styleNo = extractPurchaseOrderNo(sheet, "STYLE NO.");
    console.log("Order no:", styleNo);

    const rawRows = xlsx.utils.sheet_to_json(sheet, { defval: "" });

    // 1️ lọc cột
    const columnCleaned = filterInvalidColumns(rawRows);

    // 2️ tìm cột COUNT
    const countColumnKey = findColumnKeyByTitle(columnCleaned, "count_");

    // 3️ lọc row COUNT = 0
    // const rowCleaned = columnCleaned.filter((row) => {
    //   if (isValidColumn([row], countColumnKey)) return true;
    //   return !isZeroInCountColumn(row, countColumnKey);
    // });

    // let palet = 0;
    // hãy viết kiểm tra chỗ này nếu có 1 Palet thì palet = palet + 20, cứ có một palet xuất hiện thì cột 20 được không

    const excludedTitles = [
      "grand total",
      "TỔNG PO",
    ];

    let palet = 0;

    // 🔴 CẮT TỪ COLOR SUMMARY TRỞ ĐI
    const rowsAfterCut = cutRowsFromTitleInFirstColumn(
      columnCleaned,
      "COLOR SUMMARY:"
    );

    const rowCleaned = rowsAfterCut.filter(row => {

      const firstKey = Object.keys(row)[0];
      const firstCell = String(row[firstKey] || "")
        .toUpperCase()
        .trim();

      if (firstCell === "" || firstCell === "TOTAL") {
        return false;
      }

      if (filterRowsByExcludedTitles([row], excludedTitles).length === 0) {
        return false;
      }

      if (isValidColumn([row], countColumnKey)) return true;
      return !isZeroInCountColumn(row, countColumnKey);
    });


    //  GIỮ NGUYÊN DATA CHO FRONTEND
    result.sheetNames.push(sheetName);
    result.data[sheetName] = rowCleaned.map((row) => ({
      SHEET: sheetName,
      ...row,
    }));


    // console.log("palet: ", palet);

    // 4️ build cellMatrix cho sheet này
    const columnCells = extractColumnCellArrays(rowCleaned, sheetName);
    cellMatrix.push(...columnCells);

    // ===================== TÍNH TOÁN =====================

    console.log("------------------------------------------------------");


    // color 
    const getColor = getValuesBelowTitleAuto(cellMatrix, "COLOR");
    const colorFiltered = removeItemFromArray(getColor, "COLOR");
    // const colorFiltered1 = removeAfterStringValue(colorFiltered, "Product code");
    // const colorFiltered2 = removeAfterStringValue(colorFiltered1, "MODEL NAME");
    // const colorFiltered3 = removeAfterStringValue(colorFiltered2, "TỔNG PO");
    // const colorFiltered4 = removeAfterStringValue(colorFiltered3, "SHADE");
    // const colorFiltered5 = removeAfterStringValue(colorFiltered4, "NB+COLOR");
    // const colorFiltered6 = removeAfterStringValue(colorFiltered5, "PO");
    const colorArr = colorFiltered;
    console.log("color arr: ", colorArr);

    // article / style
    // const getArticleStyle = getValuesBelowTitleAuto(cellMatrix, "STYLE NO.");
    // const articleStyleArr1 = removeItemFromArray(getArticleStyle, "Product code");
    // const articleStyleArr2 = removeAfterStringValue(articleStyleArr1, "MODEL NAME")
    // const articleStyleArr3 = removeAfterStringValue(articleStyleArr2, "TỔNG PO")
    // const articleStyleArr4 = removeAfterStringValue(articleStyleArr3, "NB+COLOR")
    // const articleStyleArr5 = removeAfterStringValue(articleStyleArr4, "SHADE ")
    // const articleStyleArr = removeAfterStringValue(articleStyleArr5, "PO");

    // console.log("ARTICLE / STYLE: ", articleStyleArr);

    // po 
    // const getPO = extractPurchaseOrderNo(sheet, "PO");
    // const POFiltered = removeItemFromArray(getPO, "PO");
    // const POFiltered1 = removeAfterStringValue(POFiltered, "Product code");
    // const POFiltered2 = removeAfterStringValue(POFiltered1, "PO");
    // const POFiltered3 = removeAfterStringValue(POFiltered2, "TỔNG PO");
    // const POFiltered4 = removeAfterStringValue(POFiltered3, "NB+COLOR");
    // const POFiltered5 = removeAfterStringValue(POFiltered4, "SHADE");
    // const POFiltered6 = removeAfterStringValue(POFiltered5, "MODEL NAME");
    // const PO = getPO;
    // console.log("PO ARR: ", PO);


    // kiểm duyệt color, article / style, po
    // const styleCheck = collapseIfSame(articleStyleArr);
    const colorCheck = collapseIfSame(colorArr);
    // const POCheck = collapseIfSame(POArr);

    let styleName = "";
    let color = "";
    // let PO = "";

    if (!isArray(colorCheck)) {
      // styleName = styleCheck;
      color = colorCheck;
      // PO = POCheck;
    }
    else {
      // styleName = articleStyleArr;
      color = colorArr;
      // PO = POArr;
    }

    // COUNT
    const getCount = getValuesBelowTitleAuto(cellMatrix, "carton no");
    // const countArrFiltered = removeAfterStringValue(getCount, "Product code");
    const countArr = filterOnlyNumericValues(getCount);
    const count = sumArray(countArr);
    // console.log("count arr: ", countArrCleaned);
    console.log("count: ", count);

    // TOTAL
    // total ( do QTY bị ẩn, nếu lấy theo TOTAL thì sẽ có thêm value là 'QTY' )
    const getTotals = getValuesBelowTitleAuto(cellMatrix, "Total");
    // const totalArrFiltered = filterOnlyNumericValues(getTotals);
    const totalArrCleaned = removeItemFromArray(getTotals, "Total");
    // const totalArrCleaned1 = removeItemFromArray(totalArrCleaned, "TOTAL");
    // const totalArrCleaned2 = removeAfterStringValue(totalArrCleaned1, "total");
    const totalArrFiltered = totalArrCleaned;
    const totalArr = filterOnlyNumericValues(totalArrFiltered);
    console.log("total arr: ", totalArr);
    const total = sumArray(totalArr);
    console.log("total: ", total);

    // GROSS 1
    const getGross = getValuesBelowTitleAuto(cellMatrix, "TOTAL GW");
    console.log("gross arr: ", getGross);
    const grossArr = removeItemFromArray(getGross, "kg");
    const gross = sumArray(grossArr) + palet;
    console.log("gross: ", gross);

    // NET

    const getNet = getValuesBelowTitleAuto(cellMatrix, "TOTAL NW");
    const netArr = removeItemFromArray(getNet, "kg");
    const net = sumArray(netArr);
    console.log("net: ", net);

    // VOLUME
    const getVolume = getValuesBelowTitleAuto(cellMatrix, "TOTAL CBM");
    const volumeArr = getVolume;
    const volume = sumArray(volumeArr);
    console.log("volume: ", volume);

    // CARTON FORMULA
    const getCartonDimension = getValuesBelowTitleAutoForSizeCBM(cellMatrix, "Size(cm)");
    // console.log("CARTON DIMENSION (CM) ABC: ", getCartonDimension);
    const cartonDimensionArr = filter2DArrayByKeyword(getCartonDimension, "Size(cm)");
    console.log("CARTON DIMENSION (CM): ", cartonDimensionArr);
    const cartonAndCount = mergeToObjectArray(cartonDimensionArr, countArr);
    console.log("carton and count: ", cartonAndCount);
    const cartonDimension = buildMeasFormula(cartonAndCount);
    console.log("carton dimension: ", cartonDimension);

    // ===================== GẮN SUMMARY =====================
    result.summary.push({
      sheetName,
      article: styleNo,
      po: PO,
      color: color,
      countTotal: count,
      total: total,
      gross: gross,
      net: net,
      volume: volume,
      cartonFormula: cartonDimension,
    });

  });
  console.log(result.summary);

  return result;
}

// EDDIE BAUER USA SMS
const EDDIE_BAUER_USA_SMS_REQUIRED_HEADERS = [
  "DRN#",
  "PO#",
  "COLOR",
  "CARTON NO.",
  "QTY (PCS)",
  "GW (KG)",
];

function previewEDDIE_BAUER_USA_SMS(workbook) {
  const result = {
    sheetNames: [],
    data: {},
    summary: [],
  };

  // phần này xử lý cho ONE IBIS 021W:
  workbook.SheetNames.forEach((sheetName) => {

    const sheet = workbook.Sheets[sheetName];

    // 🛑 GUARD: nếu sheet KHÔNG thỏa ROSSIGNOL BULK → BỎ QUA
    if (!isSheetMatchTemplate(sheet, EDDIE_BAUER_USA_SMS_REQUIRED_HEADERS)) {
      console.log(`⏭️ Skip sheet: ${sheetName}`);
      return;
    }
    const cellMatrix = []; //  mỗi sheet 1 matrix

    // //  PO
    // const PO = extractPurchaseOrderNo(sheet, "PURCHASE ORDER");
    // console.log("Order no:", PO);
    // //style no
    // const styleNo = extractPurchaseOrderNo(sheet, "STYLE NO.");
    // console.log("Order no:", styleNo);

    const rawRows = xlsx.utils.sheet_to_json(sheet, { defval: "" });

    // 1️ lọc cột
    const columnCleaned = filterInvalidColumns(rawRows);

    // 2️ tìm cột COUNT
    const countColumnKey = findColumnKeyByTitle(columnCleaned, "count_");

    // 3️ lọc row COUNT = 0
    // const rowCleaned = columnCleaned.filter((row) => {
    //   if (isValidColumn([row], countColumnKey)) return true;
    //   return !isZeroInCountColumn(row, countColumnKey);
    // });

    // let palet = 0;
    // hãy viết kiểm tra chỗ này nếu có 1 Palet thì palet = palet + 20, cứ có một palet xuất hiện thì cột 20 được không

    const excludedTitles = [
      "grand total",
      "TỔNG PO",
    ];

    let palet = 0;

    // 🔴 CẮT TỪ COLOR SUMMARY TRỞ ĐI
    const rowsAfterCut = cutRowsFromTitleInFirstColumn(
      columnCleaned,
      "	REMARKS :"
    );

    const rowCleaned = rowsAfterCut.filter(row => {

      const firstKey = Object.keys(row)[1];
      const firstCell = String(row[firstKey] || "")
        .toUpperCase()
        .trim();

      if (firstCell === "" || firstCell === "TOTAL") {
        return false;
      }

      if (filterRowsByExcludedTitles([row], excludedTitles).length === 0) {
        return false;
      }

      if (isValidColumn([row], countColumnKey)) return true;
      return !isZeroInCountColumn(row, countColumnKey);
    });


    //  GIỮ NGUYÊN DATA CHO FRONTEND
    result.sheetNames.push(sheetName);
    result.data[sheetName] = rowCleaned.map((row) => ({
      SHEET: sheetName,
      ...row,
    }));


    // console.log("palet: ", palet);

    // 4️ build cellMatrix cho sheet này
    const columnCells = extractColumnCellArrays(rowCleaned, sheetName);
    cellMatrix.push(...columnCells);

    // ===================== TÍNH TOÁN =====================

    console.log("------------------------------------------------------");


    // color 
    const getColor = getValuesBelowTitleAuto(cellMatrix, "COLOR");
    const colorFiltered = removeItemFromArray(getColor, "COLOR");
    // const colorFiltered1 = removeAfterStringValue(colorFiltered, "Product code");
    // const colorFiltered2 = removeAfterStringValue(colorFiltered1, "MODEL NAME");
    // const colorFiltered3 = removeAfterStringValue(colorFiltered2, "TỔNG PO");
    // const colorFiltered4 = removeAfterStringValue(colorFiltered3, "SHADE");
    // const colorFiltered5 = removeAfterStringValue(colorFiltered4, "NB+COLOR");
    // const colorFiltered6 = removeAfterStringValue(colorFiltered5, "PO");
    const colorArr = colorFiltered;
    console.log("color arr: ", colorArr);

    // article / style
    const getArticleStyle = getValuesBelowTitleAuto(cellMatrix, "DRN#");
    // const articleStyleArr1 = removeItemFromArray(getArticleStyle, "Product code");
    // const articleStyleArr2 = removeAfterStringValue(articleStyleArr1, "MODEL NAME")
    // const articleStyleArr3 = removeAfterStringValue(articleStyleArr2, "TỔNG PO")
    // const articleStyleArr4 = removeAfterStringValue(articleStyleArr3, "NB+COLOR")
    // const articleStyleArr5 = removeAfterStringValue(articleStyleArr4, "SHADE ")
    const articleStyleArr = getArticleStyle;

    console.log("ARTICLE / STYLE: ", articleStyleArr);

    // po 
    const getPO = extractPurchaseOrderNo(sheet, "PO#");
    // const POFiltered = removeItemFromArray(getPO, "PO");
    // const POFiltered1 = removeAfterStringValue(POFiltered, "Product code");
    // const POFiltered2 = removeAfterStringValue(POFiltered1, "PO");
    // const POFiltered3 = removeAfterStringValue(POFiltered2, "TỔNG PO");
    // const POFiltered4 = removeAfterStringValue(POFiltered3, "NB+COLOR");
    // const POFiltered5 = removeAfterStringValue(POFiltered4, "SHADE");
    // const POFiltered6 = removeAfterStringValue(POFiltered5, "MODEL NAME");
    const POArr = getPO;
    console.log("PO ARR: ", POArr);


    // kiểm duyệt color, article / style, po
    const styleCheck = collapseIfSame(articleStyleArr);
    const colorCheck = collapseIfSame(colorArr);
    const POCheck = collapseIfSame(POArr);

    let styleName = "";
    let color = "";
    let PO = "";

    if (!isArray(colorCheck) && !isArray(styleCheck) && !isArray(POCheck)) {
      styleName = styleCheck;
      color = colorCheck;
      PO = POCheck;
    }
    else {
      styleName = articleStyleArr;
      color = colorArr;
      PO = POArr;
    }

    // COUNT
    const getCount = getValuesBelowTitleAuto(cellMatrix, "Carton No.");
    const countArrFiltered = normalizeCartonArray(getCount);
    const countArr = filterOnlyNumericValues(countArrFiltered);
    const count = sumArray(countArr);
    // console.log("count arr: ", countArrCleaned);
    console.log("count: ", count);

    // TOTAL
    // total ( do QTY bị ẩn, nếu lấy theo TOTAL thì sẽ có thêm value là 'QTY' )
    const getTotals = getValuesBelowTitleAuto(cellMatrix, "QTY (pcs)");
    // const totalArrFiltered = filterOnlyNumericValues(getTotals);
    const totalArrCleaned = removeItemFromArray(getTotals, "QTY (pcs)");
    // const totalArrCleaned1 = removeItemFromArray(totalArrCleaned, "TOTAL");
    // const totalArrCleaned2 = removeAfterStringValue(totalArrCleaned1, "total");
    const totalArrFiltered = totalArrCleaned;
    const totalArr = filterOnlyNumericValues(totalArrFiltered);
    console.log("total arr: ", totalArr);
    const total = sumArray(totalArr);
    console.log("total: ", total);

    // GROSS 1
    const getGross = getValuesBelowTitleAuto(cellMatrix, "GW (kg)");
    console.log("gross arr: ", getGross);
    const grossArr = removeItemFromArray(getGross, "kg");
    const gross = sumArray(grossArr) + palet;
    console.log("gross: ", gross);

    // NET
    const getNet = getValuesBelowTitleAuto(cellMatrix, "NW (kg)");
    const netArr = removeItemFromArray(getNet, "kg");
    const net = sumArray(netArr);
    console.log("net: ", net);

    // VOLUME
    const getVolume = getValuesBelowTitleAuto(cellMatrix, "TOTAL CBM");
    const volumeArr = getVolume;
    const volume = sumArray(volumeArr);
    console.log("volume: ", volume);

    // CARTON FORMULA
    const getCartonDimension = getValuesBelowTitleAuto(cellMatrix, "DIMENSION (CM)");
    // console.log("CARTON DIMENSION (CM) ABC: ", getCartonDimension);
    // const cartonDimensionArr = filter2DArrayByKeyword(getCartonDimension, "Size(cm)");
    console.log("CARTON DIMENSION (CM): ", getCartonDimension);
    const cartonAndCount = mergeToObjectArray(getCartonDimension, countArr);
    console.log("carton and count: ", cartonAndCount);
    const cartonDimension = buildMeasFormula(cartonAndCount);
    console.log("carton dimension: ", cartonDimension);

    // ===================== GẮN SUMMARY =====================
    result.summary.push({
      sheetName,
      article: styleName,
      po: PO,
      color: color,
      countTotal: count,
      total: total,
      gross: gross,
      net: net,
      volume: volume,
      cartonFormula: cartonDimension,
    });

  });
  console.log(result.summary);

  return result;
}

export function previewExcelWithSheets(inputPath) {
  const workbook = xlsx.readFile(inputPath);

  const excelType = detectExcelTypeByCell(inputPath);

  console.log("EXCEL TYPE =", excelType);

  if (excelType === "EDDIE BAUER USA SMS") {
    console.log("EDDIE BAUER USA SMS");
    return previewEDDIE_BAUER_USA_SMS(workbook);
  }

  if (excelType === "EDDIE BAUER BULK") {
    console.log("EDDIE BAUER BULK");
    return previewEDDIE_BAUER_BULK(workbook);
  }

  if (excelType === "ROSSIGNOL SMS") {
    console.log("ROSSIGNOL SMS");
    return previewROSSIGNOL_SMS(workbook);
  }

  if (excelType === "ROSSIGNOL BULK") {
    console.log("ROSSIGNOL BULK");
    return previewROSSIGNOL_BULK(workbook);
  }

  if (excelType === "MAMMUT GOP PO") {
    console.log("MAMNUT GOP PO");
    return previewMAMMUT_GOP_PO(workbook);
  }

  if (excelType === "MAMMUT") {
    console.log("MAMMUT");
    return previewMAMMUT_VSNT(workbook);
  }

  if (excelType === "REISS") {
    console.log("REISS");
    return previewPOC_REISS(workbook);
  }

  if (excelType === "POC BULK") { // chưa hoàn thiện
    console.log("POC BULK");
    return previewPOC_BULK(workbook);
  }

  if (excelType === "ODLO SMS") {
    console.log("ODLO SMS");
    return previewODLO_SMS(workbook);
  }

  if (excelType === "WITTT THIEN BINH") {
    console.log("WITTT THIEN BINH");
    return previewWITTT_THIEN_BINH(workbook);
  }

  if (excelType === "WITTT BMPY") {
    console.log("WITTT BMPY");
    return previewWITTT_BMPY(workbook);
  }

  if (excelType === "INNOVATION") {
    console.log("INNOVATION");
    return previewINNOVATION(workbook);
  }

  if (excelType === "HENRI BULK") {
    console.log("HENRI BULK")
    return previewHENRI_BULK(workbook);
  }

  if (excelType === "ASICS SMS") {
    console.log("ASICS SMS");
    return previewASICS_SMS(workbook);
  }

  if (excelType === "ASICS BULK") {
    console.log("ASICS BULK");
    return previewASICS_BULK(workbook);
  }

  if (excelType === "BARBOUR SMS") {
    console.log("BARBOUR SMS");
    return previewBARBOUR_SMS(workbook);
  }

  if (excelType === "BARBOUR BULK") {
    console.log("BARBOUR BULK");
    return previewBARBOUR_BULK(workbook);
  }

  if (excelType === "ACTIVE BULK") {
    console.log("ACTIVE BULK");
    return previewActive_BULK(workbook);
  }

  if (excelType === "ACTIVE SMS") {
    console.log("ACTIVE SMS")
    return previewActive_SMS(workbook);
  }

  if (excelType === "HAGLOFS TEMPLATE 2") {
    console.log("HAGLOFS TEMPLATE 2");
    return previewHAGLOFS_template_2(workbook);
  }


  if (excelType === "HAGLOFS TEMPLATE 1") {
    console.log("HAGLOFS TEMPLATE 1");
    return previewHAGLOFS_tempplate_1(workbook);
  }

  throw new Error("Unsupported Excel format");
}



