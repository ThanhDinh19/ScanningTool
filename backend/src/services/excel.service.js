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
  pickArray,
  isArray,
  removeAfterString,
  convertCmFormulaToMeter,
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

export function detectExcelTypeByCell(inputPath) {
  const workbook = xlsx.readFile(inputPath);

  let foundSMS = false;
  let foundActiveBrands = false;

  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    const rows = xlsx.utils.sheet_to_json(sheet, { defval: "" });

    for (const row of rows) {
      for (const value of Object.values(row)) {
        const text = String(value).toUpperCase();

        // ưu tiên cao nhất
        if (text.includes("ONE IBIS 021W")) {
          return "ONE_IBIS_021W";
        }

        if (text.includes("BARBOUR")) {
          return "BARBOUR";
        }

        // ghi nhận SMS (KHÔNG return ngay)
        if (/\bSMS\b/.test(text)) {
          foundSMS = true;
        }

        // ghi nhận ACTIVE BRANDS
        if (text.includes("ACTIVE BRANDS AS")) {
          foundActiveBrands = true;
        }

        // miền Bắc
        if (text.includes("MIỀN BẮC")) {
          return "MIỀN BẮC";
        }

        // miền Nam
        if (text.includes("MIỀN NAM")) {
          return "MIỀN NAM";
        }

        // henri
        if (text.includes("VIETSUN INVESTMENT CORPORATION")) {
          return "VIETSUN INVESTMENT CORPORATION";
        }
      }
    }
  }

  // 🔚 quyết định SAU KHI QUÉT HẾT FILE
  if (foundSMS) return "SMS";
  if (foundActiveBrands) return "ACTIVE_BRANDS_AS";

  return "UNKNOWN";
}

// ascis

function previewOneIbis(workbook) {
  const result = {
    sheetNames: [],
    data: {},
    summary: {},
  };

  // phần này xử lý cho ONE IBIS 021W:
  workbook.SheetNames.forEach((sheetName) => {
    const cellMatrix = []; //  mỗi sheet 1 matrix

    const sheet = workbook.Sheets[sheetName];
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

    const article = cellMatrix[2]?.cells?.[4] ?? "";
    const po = getNextValueByKey(cellMatrix[2]?.cells ?? [], "PO");
    const color = getNextValueByKey(cellMatrix[3]?.cells ?? [], "COLOR");

    // COUNT
    const counts = (cellMatrix[1]?.cells ?? []).filter(
      v => typeof v === "number" && !Number.isNaN(v)
    );
    const countTotal =
      counts.length > 1
        ? counts.slice(0, -1).reduce((a, b) => a + b, 0)
        : 0;

    // TOTAL
    const totals = (cellMatrix[4]?.cells ?? []).filter(v => typeof v === "number");
    const total =
      totals.length > 4
        ? totals.slice(0, -4).reduce((a, b) => a + b, 0)
        : 0;

    // HÀM TÍNH GROSS / NET / VOLUME
    const calcSmartSum = (arr) => {
      if (arr.length === 2) return arr[0];
      if (arr.length < 2) return 0;

      for (let i = arr.length - 1; i > 0; i--) {
        const sum = arr.slice(0, i).reduce((a, b) => a + b, 0);
        if (sum === arr[i]) return sum;
      }
      return arr.slice(0, -1).reduce((a, b) => a + b, 0);
    };

    const gross = calcSmartSum(
      (cellMatrix[6]?.cells ?? []).filter(v => typeof v === "number")
    );

    const net = calcSmartSum(
      (cellMatrix[7]?.cells ?? []).filter(v => typeof v === "number")
    );

    const volume = calcSmartSum(
      (cellMatrix[8]?.cells ?? []).filter(v => typeof v === "number")
    );

    // CARTON FORMULA
    const dimensions = (cellMatrix[5]?.cells ?? [])
      .filter(v => typeof v === "string")
      .map(v => v.trim().toUpperCase())
      .filter(v => /^\d+X\d+X\d+$/.test(v));

    const countsArr = counts.slice(0, -1);

    const formula = Object.entries(
      dimensions.map((d, i) => [d, countsArr[i]])
        .reduce((acc, [d, c]) => {
          acc[d] = (acc[d] || 0) + (c || 0);
          return acc;
        }, {})
    )
      .map(([d, t]) => `${d} * ${t}`)
      .join(", ");

    // ===================== GẮN SUMMARY =====================
    result.summary[sheetName] = {
      sheetName,
      article,
      po,
      color,
      countTotal,
      total,
      gross,
      net,
      volume,
      cartonFormula: formula,
    };

  });
  console.log(result.summary)

  return result;
}
//active
function previewActiveBrandsPKL(workbook) {
  const result = {
    sheetNames: [],
    data: {},
    summary: {},
  };

  workbook.SheetNames.forEach((sheetName) => {
    const cellMatrix = []; //  mỗi sheet 1 matrix

    const sheet = workbook.Sheets[sheetName];
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

    console.log(sheetName)

    // gross
    const grossWeight = getValueByTitle(rawRows, "GROSS WEIGHT:");
    console.log("Gross: ", grossWeight); // 4156.6990000000005
    const gross = grossWeight;
    // net
    const netWeight = getValueByTitle(rawRows, "NET WEIGHT:");
    console.log("Net: ", netWeight); // 4156.6990000000005
    const net = netWeight;

    // formula
    const measValues = getValuesBelowTitleAuto(
      cellMatrix,
      "MEAS PR CTN"
    );

    const countArr = getValuesBelowTitleAuto(
      cellMatrix,
      "QTY OF CTNS"
    );

    const countFormat = countArr.slice(0, countArr.length - 1);

    const measAndCount = mergeToObjectArray(measValues, countFormat);
    // console.log(measAndCount)

    const measFormula = buildMeasFormula(measAndCount);
    console.log("Fomula: ", measFormula);

    // volume total
    const volumeTotal = calculateFromFormula(measFormula);
    console.log("Volume total: ", volumeTotal);

    // count
    const countTotal = sumArray(countFormat);
    console.log("Count total: ", countTotal);

    // total qty
    const totalQty = getValuesBelowTitleAuto(cellMatrix, "Total pcs");
    const totalQtyFormat = totalQty.slice(0, totalQty.length - 1);
    const totalQtySum = sumArray(totalQtyFormat);
    console.log("Total qty: ", totalQtySum);

    // color
    const colorArr = getValuesBelowTitleAuto(cellMatrix, "COLOR");

    const colorCodes = [
      // BASIC
      "BLACK", "WHITE", "IVORY", "CREAM", "OFFWHITE", "OFF WHITE",
      "GREY", "GRAY", "CHARCOAL", "SILVER",

      // BLUE
      "NAVY", "NAVY BLUE", "ROYAL", "ROYAL BLUE", "SKY", "SKY BLUE",
      "DENIM", "INDIGO", "TEAL", "TURQUOISE", "WAVE",

      // GREEN
      "GREEN", "OLIVE", "ARMY", "KHAKI", "SAGE", "MOSS",
      "THYME", "FOREST", "EMERALD",

      // RED / PINK
      "RED", "WINE", "BURGUNDY", "MAROON",
      "PINK", "ROSE", "BLUSH", "SPINK", "HOT PINK",

      // BROWN
      "BROWN", "MOCHA", "COFFEE", "CHOCOLATE", "CAMEL", "TAN",

      // YELLOW / ORANGE
      "YELLOW", "MUSTARD", "GOLD",
      "ORANGE", "CORAL", "PEACH",

      // PURPLE
      "PURPLE", "LAVENDER", "LILAC", "VIOLET",

      // BEIGE / NATURAL
      "BEIGE", "SAND", "NUDE", "NATURAL", "ECRU",

      // SPECIAL / TREND
      "ROYA", "WAVE", "STONE", "SMOKE", "ASH", "ICE"
    ];

    const validColors = filterColorsInCodes(colorArr, colorCodes);
    console.log("COLOR: ", validColors);
    const color = validColors;

    // article / style
    const styleNameArr = getValuesBelowTitleAuto(cellMatrix, "STYLE NAME");
    console.log("ARTICLE / STYLE: ", styleNameArr[0])
    const article = styleNameArr[0]

    const po = ""

    result.summary[sheetName] = {
      sheetName,
      article: article,
      po,
      color: color,
      countTotal: countTotal,
      total: totalQtySum,
      gross: gross,
      net: net,
      volume: volumeTotal,
      cartonFormula: measFormula,
    };

  });

  // console.log("----------------------------- Summary ---------------------------")
  // console.log(result.summary);

  return result;
}

function previewActiveBrandsAW_SMS(workbook) {
  const result = {
    sheetNames: [],
    data: {},
    summary: {},
  };

  workbook.SheetNames.forEach((sheetName) => {

    if (sheetName.includes("SMS") && sheetName.includes("PL")) {
      const cellMatrix = []; //  mỗi sheet 1 matrix

      const sheet = workbook.Sheets[sheetName];
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

      // gross
      const grossArr = getValuesBelowTitleAuto(cellMatrix, "G.W./KGS");
      const gross = sumArray(grossArr.slice(0, grossArr.length - 1));
      console.log(gross)
      //net
      const netArr = getValuesBelowTitleAuto(cellMatrix, "N.W./KGS");
      const net = sumArray(netArr.slice(0, netArr.length - 1));
      console.log(net)
      // style name
      const styleNameArr = getValuesBelowTitleAuto(cellMatrix, "Style Name");
      console.log(styleNameArr)
      const colourArr = getValuesBelowTitleAuto(cellMatrix, "Colour");
      console.log(colourArr)
      // total qty
      const totalQty = getValuesBelowTitleAuto(cellMatrix, "QTY");
      const totalQtySum = sumArray(totalQty.slice(0, totalQty.length - 1))
      console.log(totalQtySum)
      // volumn 
      const volumnArr = getValuesBelowTitleAutoVoLumn_n_CBM(cellMatrix, "Volumn (CBM)");
      const totalVolumn = sumArray(volumnArr.slice(0, volumnArr.length - 1));
      console.log("total volume: ", totalVolumn)

      result.summary[sheetName] = {
        sheetName,
        article: styleNameArr,
        po: "",
        color: colourArr,
        countTotal: "",
        total: totalQtySum,
        gross: gross,
        net: net,
        volume: totalVolumn,
        cartonFormula: "",
      };
    }

  });

  return result;
}


// barbour
function previewBARBOUR(workbook) {
  const result = {
    sheetNames: [],
    data: {},
    summary: {},
  };

  workbook.SheetNames.forEach((sheetName) => {

    const cellMatrix = []; //  mỗi sheet 1 matrix

    const sheet = workbook.Sheets[sheetName];
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

    // sheet name
    console.log("sheet name: ", sheetName)
    // gross
    const grossArr = getValuesBelowTitleAuto(cellMatrix, "TOTAL GW");
    const gross = sumArray(grossArr.slice(0, grossArr.length - 1));
    console.log("gross: ", gross)
    console.log("gross arr: ", grossArr)
    //net
    const netArr = getValuesBelowTitleAuto(cellMatrix, "TOTAL NW");
    const net = sumArray(netArr.slice(0, netArr.length - 1));
    console.log("net: ", net)
    // style name
    const styleNameArr = getValuesBelowTitleAuto(cellMatrix, "STYLE");
    console.log("style: ", styleNameArr)

    // color
    const getColors = getValuesBelowTitleAuto(cellMatrix, "COLOR");
    const colorCodes = [
      "BK11", "BE33", "Aluminium ", "NY91 - Navy", "BK11 - BLACK", "BK1 - BLACK",
      "SG71 - Sage", "GN91 - Olive", "BE33 - Light Trench", "GN91 - Olive/Ancient",
      "Grey", "Navy", "Sage", "Dark Navy", "Asphalt", "Black", "Fir Green", "Wine", "Black/Classic",
      "Mist/Dress", "ClassicTartan", "Mist/Dress", "Fir Green", "BE34", "BE35"
    ];
    const colorArr = filterColorsInCodes(getColors, colorCodes);
    console.log("COLOR : ", colorArr);
    // color code
    const colorCodeArr = getValuesBelowTitleAuto(cellMatrix, "COLOR CODE");
    console.log("COLOR CODE: ", colorCodeArr)

    // nếu tồn tại "color" thì kiểm tra "color code" có tồn tại không
    // TH1: nếu tồn tại thì lấy color code
    // TH2: nếu không thì lấy color

    // color result: 
    const colorPick = pickArray(colorArr, colorCodeArr);

    // po
    const getPo = getValuesBelowTitleAuto(cellMatrix, "PO");
    const poArr = filterOnlyNumericValues(getPo);
    console.log("PO: ", poArr);

    // kiểm duyệt style name, color và po
    const styleNameCheck = collapseIfSame(styleNameArr); // string or arr
    const colorCheck = collapseIfSame(colorPick); // string or arr
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
      color = colorPick;
      po = poArr;
    }
    console.log("======================== result ======================");
    console.log("style: ", styleName);
    console.log("color: ", color);
    console.log("po: ", po)


    // total qty
    const getTotal = getValuesBelowTitleAuto(cellMatrix, "Total");
    const totalArr = removeAfterString(getTotal);
    const total = sumArray(totalArr.slice(0, totalArr.length - 1));
    console.log("total: ", total)


    // volumn 
    const volumnArr = getValuesBelowTitleAutoVoLumn_n_CBM(cellMatrix, "Volumn (CBM)");
    const totalVolumn = sumArray(volumnArr.slice(0, volumnArr.length - 1));
    console.log("total volume: ", totalVolumn)

    result.summary[sheetName] = {
      sheetName,
      article: styleName,
      po: po,
      color: color,
      countTotal: "",// chưa có
      total: total,
      gross: gross,
      net: net,
      volume: totalVolumn, // chưa có
      cartonFormula: "", // chưa có
    };

  });

  return result;
}

// canifa
function previewMienBac(workbook) {
  const result = {
    sheetNames: [],
    data: {},
    summary: {},
  };

  workbook.SheetNames.forEach((sheetName) => {

    const cellMatrix = []; //  mỗi sheet 1 matrix

    const sheet = workbook.Sheets[sheetName];
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

    // sheet name
    console.log("sheet name: ", sheetName)
    // gross
    const grossArr = getValuesBelowTitleAutoVoLumn_n_CBM(cellMatrix, "SỐ KG G.W");
    const gross = sumArray(grossArr);
    console.log("gross: ", gross);
    //net
    const netArr = getValuesBelowTitleAutoVoLumn_n_CBM(cellMatrix, "SỐ KG N.W");
    const net = sumArray(netArr);
    console.log("net: ", net);

    // color
    const getColors = getValuesBelowTitleAuto(cellMatrix, "Color:");
    const colorCodes = [
      "BK11", "BE33", "Aluminium ", "NY91 - Navy", "BK11 - BLACK", "BK1 - BLACK",
      "SG71 - Sage", "GN91 - Olive", "BE33 - Light Trench", "GN91 - Olive/Ancient",
      "Grey", "Navy", "Sage", "Dark Navy", "Asphalt", "Black", "Fir Green", "Wine", "Black/Classic",
      "Mist/Dress", "ClassicTartan", "Mist/Dress", "Fir Green", "BE34", "BE35", "Gray 405",
    ];
    const colorArr = filterColorsInCodes(getColors, colorCodes);
    const color = collapseIfSame(colorArr);
    console.log("color: ", color);

    // total 
    const getTotal = getValuesBelowTitleAutoVoLumn_n_CBM(cellMatrix, "Tổng cộng");
    const totalArr = filterOnlyNumericValues(getTotal);
    const total = totalArr[totalArr.length - 1];
    console.log('total: ', total);


    // count
    const getCount = getValuesBelowTitleAutoVoLumn_n_CBM(cellMatrix, "Tổng thùng");
    const countArr = filterOnlyNumericValues(getCount);
    const count = countArr[countArr.length - 1];
    console.log("count: ", count);

    // carton dimension
    const getCartonDimension = getValuesBelowTitleAuto(cellMatrix, "KÍCH THƯỚC THÙNG");
    const cartionDimension = convertCmFormulaToMeter(getCartonDimension[0]);
    console.log("carton dimension: ", cartionDimension);

    result.summary[sheetName] = {
      sheetName,
      article: "", // chưa có
      po: "", // chưa có
      color: color,
      countTotal: count,
      total: total,
      gross: gross,
      net: net,
      volume: 0, // chưa có
      cartonFormula: cartionDimension,
    };

  });

  return result;
}

function previewMienNam(workbook) {
  const result = {
    sheetNames: [],
    data: {},
    summary: {},
  };

  workbook.SheetNames.forEach((sheetName) => {

    const cellMatrix = []; //  mỗi sheet 1 matrix

    const sheet = workbook.Sheets[sheetName];
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

    // sheet name
    console.log("sheet name: ", sheetName)
    // gross
    const grossArr = getValuesBelowTitleAutoVoLumn_n_CBM(cellMatrix, "SỐ KG G.W");
    const gross = sumArray(grossArr);
    console.log("gross: ", gross);
    //net
    const netArr = getValuesBelowTitleAutoVoLumn_n_CBM(cellMatrix, "SỐ KG N.W");
    const net = sumArray(netArr);
    console.log("net: ", net);

    // color
    const getColors = getValuesBelowTitleAuto(cellMatrix, "Color:");
    const colorCodes = [
      "BK11", "BE33", "Aluminium ", "NY91 - Navy", "BK11 - BLACK", "BK1 - BLACK",
      "SG71 - Sage", "GN91 - Olive", "BE33 - Light Trench", "GN91 - Olive/Ancient",
      "Grey", "Navy", "Sage", "Dark Navy", "Asphalt", "Black", "Fir Green", "Wine", "Black/Classic",
      "Mist/Dress", "ClassicTartan", "Mist/Dress", "Fir Green", "BE34", "BE35", "Gray 405",
    ];
    const colorArr = filterColorsInCodes(getColors, colorCodes);
    const color = collapseIfSame(colorArr);
    console.log("color: ", color);

    // total 
    const getTotal = getValuesBelowTitleAutoVoLumn_n_CBM(cellMatrix, "Tổng cộng");
    const totalArr = filterOnlyNumericValues(getTotal);
    const total = totalArr[totalArr.length - 1];
    console.log('total: ', total);


    // count
    const getCount = getValuesBelowTitleAutoVoLumn_n_CBM(cellMatrix, "Tổng thùng");
    const countArr = filterOnlyNumericValues(getCount);
    const count = countArr[countArr.length - 1];
    console.log("count: ", count);

    // carton dimension
    const getCartonDimension = getValuesBelowTitleAuto(cellMatrix, "KÍCH THƯỚC THÙNG");
    const cartionDimension = convertCmFormulaToMeter(getCartonDimension[0]);
    console.log("carton dimension: ", cartionDimension);

    result.summary[sheetName] = {
      sheetName,
      article: "", // chưa có
      po: "", // chưa có
      color: color,
      countTotal: count,
      total: total,
      gross: gross,
      net: net,
      volume: 0, // chưa có
      cartonFormula: cartionDimension,
    };

  });

  return result;
}

// henri
function previewHENRI(workbook) {
  const result = {
    sheetNames: [],
    data: {},
    summary: {},
  };

  // phần này xử lý cho ONE IBIS 021W:
  workbook.SheetNames.forEach((sheetName) => {
    const cellMatrix = []; //  mỗi sheet 1 matrix

    const sheet = workbook.Sheets[sheetName];
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
    console.log("article: ",article);

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


export function previewExcelWithSheets(inputPath) {
  const workbook = xlsx.readFile(inputPath);

  const excelType = detectExcelTypeByCell(inputPath);

  console.log("EXCEL TYPE =", excelType);

  if (excelType === "ONE_IBIS_021W") {
    console.log("ONE_IBIS_021W")
    return previewOneIbis(workbook);
  }

  if (excelType === "BARBOUR") {
    console.log("BARBOUR")
    return previewBARBOUR(workbook);
  }

  if (excelType === "SMS") {
    console.log("SMS")
    return previewActiveBrandsAW_SMS(workbook);
  }

  if (excelType === "ACTIVE_BRANDS_AS") {
    console.log("ACTIVE_BRANDS_AS")
    return previewActiveBrandsPKL(workbook);
  }

  if (excelType === "MIỀN BẮC") {
    console.log("MIỀN BẮC");
    return previewMienBac(workbook);
  }

  if (excelType === "MIỀN NAM") {
    console.log("MIỀN NAM");
    return previewMienNam(workbook);
  }

  if (excelType === "VIETSUN INVESTMENT CORPORATION") {
    console.log("VIETSUN INVESTMENT CORPORATION");
    return previewHENRI(workbook);
  }



  throw new Error("Unsupported Excel format");
}



