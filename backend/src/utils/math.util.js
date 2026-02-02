import fs from "fs";
import path from "path";
import xlsx from "xlsx";
import { fileURLToPath } from "url";

const GRAND_TOTAL_PATH = path.join(
  process.cwd(),
  "uploads",
  "GRAND_TOTAL_consolidated.xlsx"
);

export function resetGrandTotalExcel() {
  const sheetName = "GRAND_TOTAL";

  // header chuẩn (theo buildGrandTotalRows)
  const emptyRows = [
    {
      SHEET: "",
      "ARTICLE / STYLE": "",
      PO: "",
      COLOR: "",
      COUNT: "",
      TOTAL: "",
      NET: "",
      GROSS: "",
      "VOLUME (CBM)": "",
      "CARTON DIMENSION (CM)": "",
    },
  ];

  let workbook;

  if (fs.existsSync(GRAND_TOTAL_PATH)) {
    // nếu file đã tồn tại → đọc lại
    workbook = xlsx.readFile(GRAND_TOTAL_PATH);

    const emptySheet = xlsx.utils.json_to_sheet([]);
    workbook.Sheets[sheetName] = emptySheet;

    if (!workbook.SheetNames.includes(sheetName)) {
      workbook.SheetNames.push(sheetName);
    }
  } else {
    // nếu file chưa tồn tại → tạo mới
    workbook = xlsx.utils.book_new();
    const emptySheet = xlsx.utils.json_to_sheet([]);
    xlsx.utils.book_append_sheet(workbook, emptySheet, sheetName);
  }

  xlsx.writeFile(workbook, GRAND_TOTAL_PATH);
}



export function sumArray(arr) {
  return arr
    .filter(v => typeof v === "number" && !Number.isNaN(v))
    .reduce((sum, v) => sum + v, 0);
}


// lấy value cột kế bên
export function getValueByTitle(rows, title) {
  const normalizedTitle = title.toLowerCase();

  for (const row of rows) {
    const values = Object.values(row);

    for (let i = 0; i < values.length - 1; i++) {
      if (
        typeof values[i] === "string" &&
        values[i].toLowerCase().includes(normalizedTitle)
      ) {
        return values[i + 1];
      }
    }
  }
  return null;
}


// lấy value dạng cột
// viết một hàm: truyền tham số title vào thì sẽ lấy các value nằm dưới nó

export function getValuesBelowTitleAuto(cellMatrix, title) {
  if (!title) return [];

  const target = String(title).trim();

  for (const column of cellMatrix) {
    const cells = column.cells;

    for (let i = 0; i < cells.length; i++) {
      if (
        typeof cells[i] === "string" &&
        cells[i].trim() === target
      ) {
        // tìm thấy title → lấy value bên dưới
        return cells
          .slice(i + 1)
          .filter(v => String(v || "").trim() !== "");
      }
    }
  }

  return [];
}


// cho volumn\n(CBM)

function normalizeTitle(text) {
  return String(text || "")
    .replace(/\s+/g, " ")   // gộp space + \n + \t
    .trim()
    .toLowerCase();
}

export function getValuesBelowTitleAutoVoLumn_n_CBM(cellMatrix, titles) {
  const targets = (Array.isArray(titles) ? titles : [titles])
    .map(normalizeTitle);

  for (const column of cellMatrix) {
    for (let i = 0; i < column.cells.length; i++) {
      if (typeof column.cells[i] !== "string") continue;

      const cellTitle = normalizeTitle(column.cells[i]);

      if (targets.some(t => cellTitle === t)) {
        return column.cells
          .slice(i + 1)
          .filter(v => String(v || "").trim() !== "");
      }
    }
  }

  return [];
}

// viết hàm gộp 2 mảng riêng biệt thành một mảng 2 chiều và các giá trị tương ứng với index, ví dụ: 'L.60 * W.40 * H.30' :  3
export function mergeToObjectArray(arr1, arr2) {
  const len = Math.min(arr1.length, arr2.length);

  return Array.from({ length: len }, (_, i) => ({
    meas: arr1[i],
    count: arr2[i]
  }));
}

// viết một hàm trả về công thức: 
// trước tiên là lọc các meas giống nhau và cột các count giống nhau đó lại
// và return, ví dụ return 'L.60 * W.40 * H.30 * 96, L.60 * W.40 * H.20 * 3'

export function buildMeasFormula(arr) {
  const grouped = arr.reduce((acc, { meas, count }) => {
    if (!meas || typeof count !== "number") return acc;

    acc[meas] = (acc[meas] || 0) + count;
    return acc;
  }, {});

  return Object.entries(grouped)
    .map(([meas, total]) => `${meas} * ${total}`)
    .join(", ");
}


// L.60 * W.40 * H.30 * 96, L.60 * W.40 * H.20 * 3 
// sau khi có được công thức trên thì hãy viết một phần tính công thức
// ví dụ result = 0.6 * 0.4 * 0.3 * 96 + 0.6 * 0.4 * 0.2 * 3
// return result

export function calculateFromFormula(formula) {
  if (!formula || typeof formula !== "string") return 0;

  return formula
    .split(",")
    .map(p => p.trim())
    .reduce((total, item) => {
      const numbers = item.match(/\d+(?:\.\d+)?/g)?.map(Number);
      if (!numbers || numbers.length < 2) return total;

      const qty = numbers[numbers.length - 1];
      const dims = numbers.slice(0, -1);

      // CM → M (ĐÚNG 1 LẦN)
      const volumePerCtn = dims.reduce(
        (v, d) => v * (d / 100),
        1
      );

      return total + volumePerCtn * qty;
    }, 0);
}


// viết hàm kiểm tra màu, nếu trong mảng không tồn tại màu hợp lệ thì loại khỏi mảng
// lọc màu

export function filterColorsInCodes(arr, codeColors = []) {
  const codeSet = new Set(
    codeColors.map(c => String(c).trim().toUpperCase())
  );

  return arr.filter(v => {
    if (typeof v !== "string") return false;

    const s = v.trim();
    if (!s) return false;

    const upper = s.toUpperCase();

    // ✅ chỉ giữ màu có trong codeColors
    return codeSet.has(upper);
  });
}



// lọc số 

export function filterOnlyNumericValues(arr) {
  return arr.filter(v =>
    typeof v === "number" ||
    (typeof v === "string" && /^\d+$/.test(v.trim()))
  );
}

//  hàm kiểm tra mảng, nếu trong mảng cùng một giá trị, 
// thì trả về một giá trị chuỗi trong mảng, ngược lại 
// nếu có một giá trị khác thì giữ nguyên mảng

export function collapseIfSame(arr) {
  if (!Array.isArray(arr) || arr.length === 0) return arr;

  const first = arr[0];

  const allSame = arr.every(v => v === first);

  return allSame ? first : arr;
}

// viết một hàm kiểm tra 2 arr, 
// arr1 và arr2, nếu arr1 không rỗng, thì kiểm tra arr2, nếu mảng arr2 không rỗng 
// thì lấy mảng return arr2, nếu arr2 rỗng thì return arr1

export function pickArray(arr1, arr2) {
  const isNonEmptyArray = (arr) =>
    Array.isArray(arr) && arr.length > 0;

  if (isNonEmptyArray(arr1)) {
    if (isNonEmptyArray(arr2)) {
      return arr2;
    }
    return arr1;
  }
  return arr2;
}


export function isArray(value) {
  return Array.isArray(value);
}


// total:  [
//   140,     212, 88,
//   20,      8,   3,
//   3,       3,   477,
//   'TOTAL', 477, 477
// ]

// viết hàm kiểm tra mảng số, nếu trong một mảng mà có string thì loại 
// các giá trị từ string đó, ví dụ như mảng trên các bỏ total, 477, 477
export function removeAfterString(arr) {
  const idx = arr.findIndex(v => typeof v === "string");
  return idx === -1 ? arr : arr.slice(0, idx);
}


// convert công thức
export function convertCmFormulaToMeter(formula) {
  if (typeof formula !== "string") return "";

  const numbers = formula.match(/\d+/g);
  if (!numbers) return "";

  return numbers
    .map(n => (Number(n) / 100).toString())
    .join("*");
}
