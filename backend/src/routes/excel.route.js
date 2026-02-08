import express from "express";
import multer from "multer";
import path from "path";
import fs from "fs";
import xlsx from "xlsx";
import {
  readExcelFile,
  writeResultToJsonFile,
  previewExcelWithSheets,
} from "../services/excel.service.js";
import { resetGrandTotalExcel } from "../utils/math.util.js";


const router = express.Router();
const upload = multer({ dest: "uploads/" });


// ghi vào GRAND_TOTAL_consolidated.xlsx:
const GRAND_TOTAL_PATH = path.join(
  process.cwd(),
  "uploads",
  "GRAND_TOTAL_consolidated.xlsx"
);

function buildGrandTotalRows(summary) {
  const rows = [];

  summary.forEach(item => {
    const colors = Array.isArray(item.color) ? item.color : [item.color];
    const articles = Array.isArray(item.article) ? item.article : [item.article];
    const posRaw = Array.isArray(item.po) ? item.po : [item.po];
    const pos = posRaw.filter(v => String(v || "").trim() !== "");

    // =========================
    // CASE 1: style ↔ color ↔ po (1–1–1)
    // =========================
    if (
      articles.length === colors.length &&
      colors.length === pos.length &&
      articles.length > 1
    ) {
      articles.forEach((article, idx) => {
        rows.push(buildRow(item, article, colors[idx], pos[idx]));
      });
      return;
    }

    // =========================
    // CASE 2: style ↔ color (1–1) + PO đơn
    // =========================
    if (
      articles.length === colors.length &&
      articles.length > 1 &&
      pos.length === 1
    ) {
      articles.forEach((article, idx) => {
        rows.push(buildRow(item, article, colors[idx], pos[0]));
      });
      return;
    }

    // =========================
    // CASE 3: style ↔ color (1–1), không có PO
    // =========================
    if (
      articles.length === colors.length &&
      articles.length > 1 &&
      pos.length === 0
    ) {
      articles.forEach((article, idx) => {
        rows.push(buildRow(item, article, colors[idx], ""));
      });
      return;
    }

    // =========================
    // CASE X: color là mảng, style & PO đơn
    // =========================
    if (
      articles.length === 1 &&
      colors.length > 1 &&
      pos.length === 1
    ) {
      colors.forEach(color => {
        rows.push(buildRow(item, articles[0], color, pos[0]));
      });
      return;
    }

    // =========================
    // CASE 4: fallback cross (chỉ khi thật sự cần)
    // =========================
    articles.forEach(article => {
      colors.forEach(color => {
        pos.forEach(po => {
          rows.push(buildRow(item, article, color, po));
        });
      });
    });
  });

  return rows;
}
// >>> hãy thêm một case chỉ có color là mảng, po đơn, và style đơn

function buildRow(item, article, color, po) {
  return {
    SHEET: item.sheetName,
    "ARTICLE / STYLE": article,
    PO: po,
    COLOR: color,
    COUNT: item.countTotal,
    TOTAL: item.total,
    NET: item.net,
    GROSS: item.gross,
    "VOLUME (CBM)": Number(item.volume.toFixed(3)),
    "CARTON DIMENSION (CM)": item.cartonFormula,
  };
}


function writeGrandTotalExcel(summary) {
  const rows = buildGrandTotalRows(summary);

  let workbook;
  let worksheet;
  const sheetName = "GRAND_TOTAL";

  if (fs.existsSync(GRAND_TOTAL_PATH)) {
    // đọc file cũ
    workbook = xlsx.readFile(GRAND_TOTAL_PATH);
    worksheet = workbook.Sheets[sheetName];

    const existingData = worksheet
      ? xlsx.utils.sheet_to_json(worksheet, { defval: "" })
      : [];

    const newData = [...rows, ...existingData];

    const newSheet = xlsx.utils.json_to_sheet(newData);
    workbook.Sheets[sheetName] = newSheet;

    if (!workbook.SheetNames.includes(sheetName)) {
      workbook.SheetNames.push(sheetName);
    }
  } else {
    // tạo file mới
    workbook = xlsx.utils.book_new();
    worksheet = xlsx.utils.json_to_sheet(rows);
    xlsx.utils.book_append_sheet(workbook, worksheet, sheetName);
  }
  xlsx.writeFile(workbook, GRAND_TOTAL_PATH);
}

// >>> khi ghi row mới thì cho lên trên đầu được không

// TAB 1 – DATA
router.get("/data", (req, res) => {
  const data = readExcelFile();
  res.json(data);
});

// TAB 2 – PREVIEW IMPORT
router.post("/excel/preview", upload.single("file"), (req, res) => {

  const result = previewExcelWithSheets(req.file.path);

  writeGrandTotalExcel(result.summary);
  const jsonPath = writeResultToJsonFile(result);

  fs.unlinkSync(req.file.path);

  res.json({
    message: "Preview success",
    jsonFile: jsonPath,
    ...result,
  });
});

router.post('/excel/reset', (req, res) => {
  resetGrandTotalExcel();
})


export default router;



