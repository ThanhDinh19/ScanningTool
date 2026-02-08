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


// cho total của active bulk dùng
export function getValuesBelowTitleAutoForTotalActiveBULK(cellMatrix, title) {
  if (!title) return [];

  const target = normalize(title);

  for (const column of cellMatrix) {
    const cells = column.cells.map(c => normalize(c));

    for (let i = 0; i < cells.length; i++) {
      // 1️⃣ match đúng 1 cell
      if (cells[i] === target) {
        return column.cells
          .slice(i + 1)
          .filter(v => String(v || "").trim() !== "");
      }

      // 2️⃣ match ghép 2 cell
      const join2 = `${cells[i]} ${cells[i + 1] || ""}`.trim();
      if (join2 === target) {
        return column.cells
          .slice(i + 2)
          .filter(v => String(v || "").trim() !== "");
      }

      // 3️⃣ match ghép 3 cell
      const join3 = `${join2} ${cells[i + 2] || ""}`.trim();
      if (join3 === target) {
        return column.cells
          .slice(i + 3)
          .filter(v => String(v || "").trim() !== "");
      }
    }
  }

  return [];
}

export function getValuesBelowTitleAutoForColorACTIVE_BULK(cellMatrix, title) {
  if (!title) return [];

  const target = normalize(title);
  const matchedColumns = [];

  for (const column of cellMatrix) {
    const cells = column.cells.map(c => normalize(c));
    let startIndex = -1;

    for (let i = 0; i < cells.length; i++) {
      // 1️⃣ match 1 cell
      if (cells[i] === target || cells[i].includes(target)) {
        startIndex = i + 1;
        break;
      }

      // 2️⃣ match ghép 2 cell
      const join2 = `${cells[i]} ${cells[i + 1] || ""}`.trim();
      if (join2.includes(target)) {
        startIndex = i + 2;
        break;
      }

      // 3️⃣ match ghép 3 cell
      const join3 = `${join2} ${cells[i + 2] || ""}`.trim();
      if (join3.includes(target)) {
        startIndex = i + 3;
        break;
      }
    }

    // nếu cột này match → lấy values
    if (startIndex > -1) {
      const values = column.cells
        .slice(startIndex)
        .filter(v => String(v || "").trim() !== "");

      matchedColumns.push(values);
    }
  }

  // ✅ ưu tiên mảng thứ 2
  if (matchedColumns.length >= 2) {
    return matchedColumns[1];
  }

  // fallback: mảng đầu hoặc []
  return matchedColumns[0] || [];
}


// >> viết hàm lấy đúng value theo đúng title, ko upper, ko lower, ko includes

export function getValuesBelowTitleExact(cellMatrix, title) {
  if (!title) return [];

  for (const column of cellMatrix) {
    const cells = column.cells;

    for (let i = 0; i < cells.length; i++) {
      // match CHÍNH XÁC title
      if (cells[i] === title) {
        return column.cells
          .slice(i + 1)
          .filter(v => String(v ?? "").trim() !== "");
      }
    }
  }

  return [];
}

export function getValuesBelowTitleAuto(cellMatrix, title) {
  if (!title) return [];

  const target = normalize(title);

  for (const column of cellMatrix) {
    const cells = column.cells.map(c => normalize(c));

    for (let i = 0; i < cells.length; i++) {
      // 1️ match 1 cell
      if (cells[i] === target || cells[i].includes(target)) {
        return column.cells
          .slice(i + 1)
          .filter(v => String(v || "").trim() !== "");
      }

      // 2️ match text dọc (ghép 2–3 cell)
      const join2 = `${cells[i]} ${cells[i + 1] || ""}`.trim();
      const join3 = `${join2} ${cells[i + 2] || ""}`.trim();

      if (join2.includes(target) || join3.includes(target)) {
        return column.cells
          .slice(i + 2)
          .filter(v => String(v || "").trim() !== "");
      }
    }
  }
  return [];
}

function normalize(str) {
  return String(str || "")
    .replace(/\n/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .toUpperCase();
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

export function mergeToObjectArrayLevel2(arr1, arr2) {
  const len = Math.min(arr1.length, arr2.length);

  return Array.from({ length: len }, (_, i) => ({
    meas: arr1[i],
    item: arr2[i]
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

export function removeAfterStringValue(arr, stopString) {
  if (!Array.isArray(arr)) return [];

  const target = String(stopString).trim().toUpperCase();

  const idx = arr.findIndex(v =>
    typeof v === "string" &&
    String(v).trim().toUpperCase().includes(target)
  );

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


// viết hàm lọc ra những cái ko cần thiết ra khỏi mảng ví dụ [ '60x40x30', '60x40x40', '60x40x30', '01/08/2025' ]
// bỏ đi '01/08/2025'

export function keepOnlyDimensionLxWxH(arr) {
  return arr.filter(item =>
    /^\d+\s*x\s*\d+\s*x\s*\d+$/i.test(String(item).trim())
  );
}


// [ '1 OF 4', '2 OF 4', '3 OF 4', '4 OF 4' ] >>> 
// viết hàm chuyển mảng này sang dạng [1, 1, 1, 1], 
// hoặc nếu có trường hợp mảng toàn số [1, 2, 3] thì giữ nguyên

export function normalizeCartonArray(arr) {
  if (!Array.isArray(arr)) return arr;

  const ofPattern = /^\s*\d+\s*OF\s*\d+\s*$/i;

  const numberItems = arr.filter(v => typeof v === "number");
  const ofItems = arr.filter(v => ofPattern.test(String(v)));

  // 1️⃣ toàn số → giữ nguyên
  if (numberItems.length === arr.length) {
    return arr;
  }

  // 2️⃣ toàn "x OF y" → toàn 1
  if (ofItems.length === arr.length) {
    return ofItems.map(() => 1);
  }

  // 3️⃣ count arr bị lẫn → chỉ lấy "x OF y"
  if (ofItems.length > 0) {
    return ofItems.map(() => 1);
  }

  // 4️⃣ fallback
  return arr;
}



export function extractAllColorCodesFromWorkbook(workbook) {
  const colorCodes = [];

  // regex cho color code kiểu BK11, BE33, GN91, BK1...
  const colorCodeRegex = /^[A-Z]{1,3}\d{1,3}$/i;

  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];

    for (const cellAddress in sheet) {
      if (cellAddress.startsWith("!")) continue;

      const value = sheet[cellAddress]?.v;
      if (!value) continue;

      const text = String(value).trim();

      if (colorCodeRegex.test(text)) {
        colorCodes.push(text.toUpperCase());
      }
    }
  }

  // loại trùng
  return [...new Set(colorCodes)];
}

// có thể viết hàm nào mà có thể tính sum của một mảng trong nhiều trường hợp như:
// ví dụ trường hợp 1: mảng có tổng sẵn, [12, 13, 25] 25 là tổng, khi tính thì loại 25 ra rồi mới tính tổng
// trường hợp 2: mảng không có tổng sẵn [12, 13] tính tổng ko cần loại

export function smartSumFloat(arr, epsilon = 1e-6) {
  const nums = arr.filter(v => typeof v === "number" && !isNaN(v));
  if (nums.length <= 1) return nums[0] || 0;

  const last = nums[nums.length - 1];
  const sumExceptLast = nums.slice(0, -1).reduce((a, b) => a + b, 0);

  if (Math.abs(last - sumExceptLast) < epsilon) {
    return sumExceptLast;
  }

  return nums.reduce((a, b) => a + b, 0);
}

// count arr:  [
//    1,   6, 10,  7, 3, 1, 1, 1,  30,   4,
//   19,  27, 19,  8, 1, 1, 1, 1,  81,   5,
//   21,  31, 21,  8, 1, 1, 1, 1,  90,  18,
//   70, 101, 69, 29, 3, 1, 1, 1, 293, 494
// ]

// có cách nào để tính được giá trị tổng của arr với trường hợp này ko,
// ví dụ mảng như trên, thì từ  1, 6, 10, 7, 3, 1, 1, 1 = 30. 
// từ 4, 19, 27, 19,  8, 1, 1, 1, 1 = 81.
// từ 5, 21,  31, 21,  8, 1, 1, 1, 1 = 90.
// 18, 70, 101, 69, 29, 3, 1, 1, 1 = 293.
// tổng 81 + 90 + 293 = 494.
// >>> return 494 
// có thể viết hàm tính kiểu v này được k

export function smartCountSum(arr) {
  if (!Array.isArray(arr)) return 0;

  const nums = arr.filter(v => typeof v === "number" && !isNaN(v));
  if (nums.length === 0) return 0;

  //  Rule 0: chỉ có 1 phần tử
  if (nums.length === 1) {
    return nums[0];
  }

  // Rule 1: nếu số cuối >= mọi số trước → đó là kết quả
  const last = nums[nums.length - 1];
  const maxBefore = Math.max(...nums.slice(0, -1));

  if (last >= maxBefore) {
    return last;
  }

  // Rule 2: fallback – cộng toàn bộ
  return nums.reduce((a, b) => a + b, 0);
}



// viết thêm hàm bỏ các tổng ra khỏi mảng được không 
export function removeAllBlockTotals(arr) {
  if (!Array.isArray(arr)) return [];

  const nums = arr.filter(v => typeof v === "number" && !isNaN(v));
  const result = [];

  let buffer = [];

  for (const n of nums) {
    const sumBuffer = buffer.reduce((a, b) => a + b, 0);

    // nếu n đúng bằng tổng buffer → đây là subtotal → bỏ
    if (buffer.length > 0 && n === sumBuffer) {
      buffer = []; // reset block
      continue;
    }

    buffer.push(n);
    result.push(n);
  }

  return result;
}



// arrA arr:  [
//    10,  60, 100,   70,   30,   9,    8,    9,
//   296,  40, 190,  270,  190,  80,   10,    7,
//     7,   7, 801,   50,  210, 310,  210,   80,
//    10,   7,   9,    8,  894, 180,  700, 1010,
//   690, 290,  30,   10,    9,  10, 2929, 4920,
//   296, 801, 894, 2929, 4920
// ]

// arrB arr:  [
//    1,   6, 10,  7, 3, 1, 1, 1,  30,   4,
//   19,  27, 19,  8, 1, 1, 1, 1,  81,   5,
//   21,  31, 21,  8, 1, 1, 1, 1,  90,  18,
//   70, 101, 69, 29, 3, 1, 1, 1, 293, 494
// ]

// viết hàm lọc mảng bởi mảng được ko
// ví dụ như 2 mảng trên, khi truyền vào 2 mảng arrA, arrB.
// thì sẽ lọc theo arrB, tức mảng arrB bao nhiêu phần tử thì return về mảng A cũng bấy nhiêu
// tực mảng A sẽ bỏ từ  296, 801, 894, 2929, 4920
// viết hàm này được không

export function trimArrayByArray(arrA, arrB) {
  if (!Array.isArray(arrA) || !Array.isArray(arrB)) return [];

  return arrA.slice(0, arrB.length);
}


export function removeItemFromArray(arr, removeValue) {
  if (!Array.isArray(arr)) return [];

  // nếu là string → so sánh không phân biệt hoa thường, trim
  if (typeof removeValue === "string") {
    const target = removeValue.trim().toUpperCase();

    return arr.filter(v =>
      typeof v !== "string" ||
      v.trim().toUpperCase() !== target
    );
  }

  // nếu là number → so sánh trực tiếp
  if (typeof removeValue === "number") {
    return arr.filter(v => v !== removeValue);
  }

  return arr;
}


// carton number and count:  [
//   { meas: 1, count: 1 },
//   { meas: 2, count: 6 },
//   { meas: 8, count: 10 },
//   { meas: 18, count: 7 },
//   { meas: 25, count: 3 },
//   { meas: 28, count: 1 },
//   { meas: 29, count: 1 },
//   { meas: 30, count: 1 },
//   { meas: 'TOTAL CARTONS ', count: 30 },
//   { meas: 31, count: 4 },
//   { meas: 35, count: 19 },
//   { meas: 54, count: 27 },
//   { meas: 81, count: 19 },
//   { meas: 100, count: 8 },
//   { meas: 108, count: 1 },
//   { meas: 109, count: 1 },
//   { meas: 110, count: 1 },
//   { meas: 111, count: 1 },
//   { meas: 'TOTAL CARTONS ', count: 81 },
//   { meas: 112, count: 5 },
//   { meas: 117, count: 21 },
//   { meas: 138, count: 31 },
//   { meas: 169, count: 21 },
//   { meas: 190, count: 8 },
//   { meas: 198, count: 1 },
//   { meas: 199, count: 1 },
//   { meas: 200, count: 1 },
//   { meas: 201, count: 1 },
//   { meas: 'TOTAL CARTONS ', count: 90 },
//   { meas: 202, count: 18 },
//   { meas: 220, count: 70 },
//   { meas: 290, count: 101 },
//   { meas: 391, count: 69 },
//   { meas: 460, count: 29 },
//   { meas: 489, count: 3 },
//   { meas: 492, count: 1 },
//   { meas: 493, count: 1 },
//   { meas: 494, count: 1 },
//   { meas: 'TOTAL CARTONS ', count: 293 },
//   { meas: 'TOTAL CARTONS ', count: 494 }
// ] 

// viết hàm loại những phần tử nào mà được truyền string vào


export function removeItemsByMeasString(arr, removeString) {
  if (!Array.isArray(arr)) return [];

  const target = String(removeString).trim().toUpperCase();

  return arr.filter(item => {
    if (!item || typeof item !== "object") return true;

    if (typeof item.meas === "string") {
      return item.meas.trim().toUpperCase() !== target;
    }

    return true;
  });
}


// hàm truyền vào matrix return về 1 arr count
export function extractCountArray(arr) {
  if (!Array.isArray(arr)) return [];

  return arr
    .filter(item => item && typeof item.count === "number" && !isNaN(item.count))
    .map(item => item.count);
}

export function extractItemArray(arr) {
  if (!Array.isArray(arr)) return [];

  return arr
    .filter(i => i && typeof i.item === "number" && !isNaN(i.item))
    .map(i => i.item);
}



// viết hàm kiểm tra mảng, truyền vào arr và một string, 
// nếu lọc trong mảng mà ko có từ string được truyền vào thì trả về false

export function arrayContainsString(arr, keyword) {
  if (!Array.isArray(arr) || !keyword) return false;

  const target = String(keyword).trim().toUpperCase();

  return arr.some(item =>
    String(item).toUpperCase().includes(target)
  );
}


// trả về mảng các mảng PO
export function getAllPOBlocks(cellMatrix, title) {
  const target = normalize(title);
  const blocks = [];

  for (const column of cellMatrix) {
    const rawCells = column.cells;
    const cells = rawCells.map(c => normalize(c));

    for (let i = 0; i < cells.length; i++) {
      if (cells[i] === target || cells[i].includes(target)) {
        const block = [];

        for (let j = i + 1; j < rawCells.length; j++) {
          const val = rawCells[j];
          const text = normalize(val);

          if (!val || text.includes("TOTAL")) break;
          block.push(val);
        }

        if (block.length) blocks.push(block);
      }
    }
  }

  return blocks;
}


export function hasAtLeastNEmptyCells(row, n = 7) {
  const emptyCount = Object.values(row).filter(
    (v) => v === "" || v === null || v === undefined
  ).length;

  return emptyCount >= n;
}


// 
export function calcSum2Arr(qtyArr, valueArr) {
  if (!Array.isArray(qtyArr) || !Array.isArray(valueArr)) return 0;

  const len = Math.min(qtyArr.length, valueArr.length);

  let total = 0;

  for (let i = 0; i < len; i++) {
    const qty = Number(qtyArr[i]) || 0;
    const value = Number(valueArr[i]) || 0;

    total += qty * value;
  }

  return Number(total.toFixed(3)); // làm tròn cho đẹp
}


// lấy value từ header cuối
export function removeBeforeTitle(arr, title) {
  if (!Array.isArray(arr) || !title) return arr || [];

  const target = normalizeForRemoveBeforeTitle(title);

  const index = arr.findIndex(
    v => normalizeForRemoveBeforeTitle(v) === target
  );

  // nếu không tìm thấy title → giữ nguyên
  if (index === -1) return arr;

  // cắt bỏ từ đầu tới sau title
  return arr.slice(index + 1);
}

function normalizeForRemoveBeforeTitle(str) {
  return String(str || "")
    .replace(/\n/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .toUpperCase();
}


// get purchase order no ở active bulk
export function extractPurchaseOrderNo(sheet, target) {
  const TARGET = normalizeForPO(target);

  for (const addr in sheet) {
    if (addr.startsWith("!")) continue;

    const cell = sheet[addr];
    const text = normalizeForPO(cell?.v);

    // ✅ CASE 1 + CASE 2
    if (text.startsWith(TARGET) || text === TARGET) {
      const { row, col } = splitCellAddress(addr);

      // 👉 quét ngang cùng row, từ cột kế bên
      for (let c = col + 1; c < col + 10; c++) {
        const nextAddr = makeCellAddress(c, row);
        const nextCell = sheet[nextAddr];

        if (nextCell && String(nextCell.v || "").trim() !== "") {
          return String(nextCell.v).trim();
        }
      }
    }
  }
  return "";
}


function splitCellAddress(addr) {
  const m = addr.match(/^([A-Z]+)(\d+)$/);
  if (!m) return {};

  return {
    col: columnToNumber(m[1]),
    row: Number(m[2]),
  };
}

function makeCellAddress(col, row) {
  return `${numberToColumn(col)}${row}`;
}

function columnToNumber(col) {
  let n = 0;
  for (let i = 0; i < col.length; i++) {
    n = n * 26 + (col.charCodeAt(i) - 64);
  }
  return n;
}

function numberToColumn(n) {
  let col = "";
  while (n > 0) {
    const r = (n - 1) % 26;
    col = String.fromCharCode(65 + r) + col;
    n = Math.floor((n - 1) / 26);
  }
  return col;
}

function normalizeForPO(v) {
  return String(v || "")
    .replace(/\n/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .toUpperCase();
}



export function getValuesBelowTitleAutoForSizeCBM(cellMatrix, title) {
  if (!title) return [];

  const target = normalizeForSizeCBM(title);

  for (let colIndex = 0; colIndex < cellMatrix.length; colIndex++) {
    const cells = cellMatrix[colIndex].cells.map(c =>
      normalizeForSizeCBM(c)
    );

    for (let i = 0; i < cells.length; i++) {
      let startIndex = -1;

      if (cells[i] === target || cells[i].includes(target)) {
        startIndex = i + 1;
      } else {
        const join2 = `${cells[i]} ${cells[i + 1] || ""}`.trim();
        const join3 = `${join2} ${cells[i + 2] || ""}`.trim();

        if (join2.includes(target)) startIndex = i + 2;
        if (join3.includes(target)) startIndex = i + 3;
      }

      if (startIndex > -1) {
        // ⚠️ CHỈ cột chứa title dùng startIndex
        const col1 = cellMatrix[colIndex].cells.slice(startIndex);

        // ⚠️ cột kế bên: KHÔNG bỏ dòng đầu
        const col2 = cellMatrix[colIndex + 1]?.cells.slice(startIndex - 1) || [];
        const col3 = cellMatrix[colIndex + 2]?.cells.slice(startIndex - 1) || [];

        const len = Math.max(col1.length, col2.length, col3.length);
        const result = [[], [], []];

        for (let r = 0; r < len; r++) {
          result[0].push(cleanCell(col1[r]));
          result[1].push(cleanCell(col2[r]));
          result[2].push(cleanCell(col3[r]));
        }

        return result;
      }
    }
  }

  return [];
}


function cleanCell(v) {
  const s = String(v ?? "").trim();
  return s === "" ? null : v;
}

function normalizeForSizeCBM(str) {
  return String(str || "")
    .replace(/\n/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .toUpperCase();
}


//========================= for asics sms

export function getValuesBelowTitleAutoForCartonDimension(cellMatrix, title) {
  if (!title) return [];

  const target = normalizeForSizeCBM(title);

  for (let colIndex = 0; colIndex < cellMatrix.length; colIndex++) {
    const cells = cellMatrix[colIndex].cells.map(c =>
      normalizeForSizeCBM(c)
    );

    for (let i = 0; i < cells.length; i++) {
      let startIndex = -1;

      if (cells[i] === target) {
        startIndex = i + 1;
      } else {
        const join2 = `${cells[i]} ${cells[i + 1] || ""}`.trim();
        const join3 = `${join2} ${cells[i + 2] || ""}`.trim();

        if (join2 === target) startIndex = i + 2;
        if (join3 === target) startIndex = i + 3;
      }

      if (startIndex > -1) {
        const col1 = cellMatrix[colIndex]?.cells.slice(startIndex) || [];
        const col2 = cellMatrix[colIndex + 1]?.cells.slice(startIndex) || [];
        const col3 = cellMatrix[colIndex + 2]?.cells.slice(startIndex) || [];

        const len = Math.max(col1.length, col2.length, col3.length);
        const result = [[], [], []];

        for (let r = 0; r < len; r++) {
          result[0].push(cleanCell(col1[r]));
          result[1].push(cleanCell(col2[r]));
          result[2].push(cleanCell(col3[r]));
        }

        return result;
      }
    }
  }

  return [];
}


export function fillCartonDimension(cartonDimensionArr) {
  if (
    !Array.isArray(cartonDimensionArr) ||
    cartonDimensionArr.length !== 3
  ) {
    return cartonDimensionArr;
  }

  const [L, W, H] = cartonDimensionArr;

  const filledL = [];
  const filledW = [];
  const filledH = [];

  let lastL = null;
  let lastW = null;
  let lastH = null;

  const firstL = L.find(v => v != null);
  const firstW = W.find(v => v != null);
  const firstH = H.find(v => v != null);

  const len = Math.max(L.length, W.length, H.length);

  for (let i = 0; i < len; i++) {
    const l = L[i] ?? lastL ?? firstL;
    const w = W[i] ?? lastW ?? firstW;
    const h = H[i] ?? lastH ?? firstH;

    if (l != null) lastL = l;
    if (w != null) lastW = w;
    if (h != null) lastH = h;

    filledL.push(Number(l));
    filledW.push(Number(w));
    filledH.push(Number(h));
  }

  return [filledL, filledW, filledH];
}




export function normalizeCartonDimension(arr) {
  const [L, W, H] = arr;

  return [
    L.map(Number),
    W.map(Number),
    H.map(Number),
  ];
}


function normalizeForCartonDimension(str) {
  return String(str || "")
    .replace(/\n/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .toUpperCase();
}

function isNumberCell(v) {
  if (v === null || v === undefined) return false;
  const s = String(v).trim();
  return s !== "" && !isNaN(s);
}

// ===============================

//=================================

// size cbm:  [
//   [
//     60, 60, 60, 60,
//     60, 60, 60, 60
//   ],
//   [
//     40, 40, 40, 40,
//     40, 40, 40, 40
//   ],
//   [
//     30, 30, 30, 30,
//     30, 30, 30, 30
//   ]
// ]

// >> từ 3 cột này tui muốn làm thành công thức tương ứng giữa 3 cột:
// ví dụ: return [ 60x40x30, ... ]
export function buildCartonDimensionFormulas(sizeCBM) {
  if (
    !Array.isArray(sizeCBM) ||
    sizeCBM.length !== 3
  ) {
    return [];
  }

  const [L, W, H] = sizeCBM;

  const len = Math.min(L.length, W.length, H.length);

  const result = [];

  for (let i = 0; i < len; i++) {
    result.push(`${L[i]}x${W[i]}x${H[i]}`);
  }

  return result;
}


// viết hàm truyền vào 2 arr, kiểm tra nếu mảng nào không rỗng thì lấy
export function pickNonEmptyArray(arr1, arr2) {
  if (Array.isArray(arr1) && arr1.length > 0) return arr1;
  if (Array.isArray(arr2) && arr2.length > 0) return arr2;
  return [];
}


export function getValuesBelowTitleAuto3Cols(cellMatrix, title) {
  if (!title) return [];

  const target = normalize(title);

  for (let colIndex = 0; colIndex < cellMatrix.length; colIndex++) {
    const col = cellMatrix[colIndex];
    const cells = col.cells.map(c => normalize(c));

    for (let row = 0; row < cells.length; row++) {
      const join3 = `${cells[row]} ${cells[row + 1] || ""} ${cells[row + 2] || ""}`.trim();

      if (cells[row].includes(target) || join3.includes(target)) {

        // 🔥 tìm row data thực sự (row đầu tiên có số ở 1 trong 3 cột)
        let dataRow = -1;

        for (let r = row + 1; r < cells.length; r++) {
          const values = [0, 1, 2].map(offset => {
            const c = cellMatrix[colIndex + offset];
            return c?.cells[r];
          });

          if (values.some(v => String(v || "").trim() !== "")) {
            dataRow = r;
            break;
          }
        }

        if (dataRow === -1) return [[], [], []];

        return [0, 1, 2].map(offset => {
          const c = cellMatrix[colIndex + offset];
          if (!c) return [];

          const v = c.cells[dataRow];
          return v && String(v).trim() !== "" ? [v] : [];
        });
      }
    }
  }

  return [];
}



// const cellMatrix = []; //  mỗi sheet 1 matrix

//     const sheet = workbook.Sheets[sheetName];
//     const rawRows = xlsx.utils.sheet_to_json(sheet, { defval: "" });

//     // 1️ lọc cột
//     const columnCleaned = filterInvalidColumns(rawRows);

//     // 2️ tìm cột COUNT
//     const countColumnKey = findColumnKeyByTitle(columnCleaned, "count");

//     // 3️ lọc row COUNT = 0
//     const rowCleaned = columnCleaned.filter((row) => {
//       if (isValidColumn([row], countColumnKey)) return true;
//       return !isZeroInCountColumn(row, countColumnKey);
//     });

//     //  GIỮ NGUYÊN DATA CHO FRONTEND
//     result.sheetNames.push(sheetName);
//     result.data[sheetName] = rowCleaned.map((row) => ({
//       SHEET: sheetName,
//       ...row,
//     }));

//     // 4️ build cellMatrix cho sheet này
//     const columnCells = extractColumnCellArrays(rowCleaned, sheetName);
//     cellMatrix.push(...columnCells);

// >>>> ở cellMatrix[0].cells >> tui muốn tìm những ô nào của cột này mà với title mà tui nhập vào khớp 
// với ô đó thì xóa nguyên hàng đó được ko

export function findRowIndexesByTitle(cellMatrix, title) {
  const target = normalizeTitleF(title);
  const removeIndexes = new Set();

  cellMatrix.forEach((col) => {
    col.cells.forEach((cell, rowIndex) => {
      const v = normalizeTitleF(cell);
      if (v && v.includes(target)) {
        removeIndexes.add(rowIndex);
      }
    });
  });

  return [...removeIndexes];
}


function normalizeTitleF(str) {
  return String(str || "")
    .replace(/\n/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .toUpperCase();
}

// =======
export function filterRowsByExcludedTitles(rows, excludedTitles = []) {
  if (!excludedTitles.length) return rows;

  const targets = excludedTitles.map(t => normalize(t));

  return rows.filter(row => {
    //  ghép toàn bộ cell trong row
    const rowText = normalize(
      Object.values(row).join(" ")
    );

    return !targets.some(t => rowText.includes(t));
  });
}
// 


export function getValueForGrossNet_INNOVATION(sheet, target) {
  const TARGET = normalizeForPO(target);

  for (const addr in sheet) {
    if (addr.startsWith("!")) continue;

    const cell = sheet[addr];
    const text = normalizeForPO(cell?.v || "");

    // ✅ includes chuỗi liền nhau
    if (text.includes(TARGET)) {
      const { row, col } = splitCellAddress(addr);

      // 👉 quét ngang cùng row, từ cột kế bên
      for (let c = col + 1; c < col + 10; c++) {
        const nextAddr = makeCellAddress(c, row);
        const nextCell = sheet[nextAddr];

        if (nextCell && String(nextCell.v || "").trim() !== "") {
          return String(nextCell.v).trim();
        }
      }
    }
  }

  return "";
}




export function getValuesBelowTitleAutoExcludeKeywords(
  cellMatrix,
  title,
  excludeKeywords = []
) {
  if (!title) return [];

  const target = normalize(title);
  const excludes = excludeKeywords.map(k => normalize(k));

  for (const column of cellMatrix) {
    const cellsNormalized = column.cells.map(c => normalize(c));

    for (let i = 0; i < cellsNormalized.length; i++) {

      // 1️⃣ match 1 cell
      if (
        cellsNormalized[i] === target ||
        cellsNormalized[i].includes(target)
      ) {
        return column.cells
          .slice(i + 1)
          .filter(v => {
            const text = normalize(v);
            return (
              text !== "" &&
              !excludes.some(k => text === k || text.includes(k))
            );
          });
      }

      // 2️⃣ match text dọc (2–3 cell)
      const join2 = `${cellsNormalized[i]} ${cellsNormalized[i + 1] || ""}`.trim();
      const join3 = `${join2} ${cellsNormalized[i + 2] || ""}`.trim();

      if (join2.includes(target) || join3.includes(target)) {
        return column.cells
          .slice(i + 2)
          .filter(v => {
            const text = normalize(v);
            return (
              text !== "" &&
              !excludes.some(k => text === k || text.includes(k))
            );
          });
      }
    }
  }

  return [];
}


// viết thêm một hàm, truyền vào arr, và keyword, xóa từ keyword trở về sau

export function removeFromStringValue(arr, keyword) {
  if (!Array.isArray(arr)) return [];

  const target = String(keyword).trim().toUpperCase();

  const idx = arr.findIndex(v =>
    typeof v === "string" &&
    String(v).trim().toUpperCase().includes(target)
  );

  // nếu không tìm thấy keyword → giữ nguyên
  if (idx === -1) return arr;

  // xóa từ keyword trở đi
  return arr.slice(0, idx);
}


// >>> viết thêm một hàm, truyền vào title ví dụ như Carton như hình, 
// return về [[60, 60, 60, 60], [40, 40, 40, 40], [40, 40, 40, 40]], 
// ko lấy L, W, H. hàm này viết được không

export function getGroupedValuesBelowTitle(cellMatrix, title) {
  if (!title || !Array.isArray(cellMatrix)) return [];

  // 1️⃣ tìm cột có title "Carton"
  const startColIndex = cellMatrix.findIndex(col =>
    col.cells.some(cell => String(cell).trim() === title)
  );

  if (startColIndex === -1) return [];

  // 2️⃣ lấy 3 cột: L, W, H
  const targetCols = cellMatrix.slice(startColIndex, startColIndex + 3);

  return targetCols.map(col => {
    // 🔑 tìm index của L / W / H trong từng cột
    const startRowIndex = col.cells.findIndex(v =>
      ["L", "W", "H"].includes(String(v).trim())
    );

    if (startRowIndex === -1) return [];

    // 👉 lấy dữ liệu SAU L/W/H
    return col.cells
      .slice(startRowIndex + 1)
      .filter(v => String(v).trim() !== "")
      .map(Number);
  });
}


// hàm lấy carton dimension của MAMMUT chính xác không include, không lower, không upper

export function getValuesBelowTitleAutoForCartonDimension_MAMMUT(cellMatrix, title) {
  if (!title) return [];

  const target = normalizeForSizeCBM(title);

  for (let colIndex = 0; colIndex < cellMatrix.length; colIndex++) {
    const rawCells = cellMatrix[colIndex].cells;
    const cells = rawCells.map(c => normalizeForSizeCBM(c));

    for (let i = 0; i < cells.length; i++) {
      let startIndex = -1;

      // ✅ match chính xác title
      if (cells[i] === target) {
        startIndex = i + 1;
      } else {
        const join2 = `${cells[i]} ${cells[i + 1] || ""}`.trim();
        const join3 = `${join2} ${cells[i + 2] || ""}`.trim();

        if (join2 === target) startIndex = i + 2;
        else if (join3 === target) startIndex = i + 3;
      }

      if (startIndex > -1) {
        const col1 = rawCells.slice(startIndex);
        const col2 = cellMatrix[colIndex + 1]?.cells.slice(startIndex - 1) || [];
        const col3 = cellMatrix[colIndex + 2]?.cells.slice(startIndex - 1) || [];

        const len = Math.max(col1.length, col2.length, col3.length);
        const result = [[], [], []];

        for (let r = 0; r < len; r++) {
          const v1 = Number(cleanCell(col1[r]));
          const v2 = Number(cleanCell(col2[r]));
          const v3 = Number(cleanCell(col3[r]));

          // ✅ CHỈ push số hợp lệ
          if (!Number.isNaN(v1)) result[0].push(v1);
          if (!Number.isNaN(v2)) result[1].push(v2);
          if (!Number.isNaN(v3)) result[2].push(v3);
        }

        return result;
      }
    }
  }

  return [];
}



export function removeItemsByKeywords(arr, keywords = []) {
  if (!Array.isArray(arr) || !Array.isArray(keywords)) return arr;

  return arr.filter(item => {
    const text = String(item ?? "").toUpperCase();

    return !keywords.some(keyword =>
      text.includes(String(keyword).toUpperCase())
    );
  });
}


// > viết một hàm truyền vào một mảng và keyword, nếu mảng có phần tử bằng với phần tử
// của keyword thì return về true, so sách chính xác, không include, không lower, không upper
export function hasExactMatch(arr, keywords = []) {
  if (!Array.isArray(arr) || !Array.isArray(keywords)) return false;

  return arr.some(item =>
    keywords.some(keyword => item === keyword)
  );
}


export function extractFirst5Digits(arr) {
  if (!Array.isArray(arr)) return [];

  return arr
    .map(v => {
      const match = String(v).match(/\d{5}/);
      return match ? match[0] : null;
    })
    .filter(v => v !== null);
}



// 
export function isSheetMatchTemplate(sheet, requiredHeaders = []) {
  const allText = [];

  for (const addr in sheet) {
    if (addr.startsWith("!")) continue;

    const text = String(sheet[addr]?.v || "")
      .toUpperCase()
      .replace(/\s+/g, " ")
      .trim();

    if (text) allText.push(text);
  }

  return requiredHeaders.every(header =>
    allText.some(cellText => cellText.includes(header))
  );
}



export function cutRowsFromTitleInFirstColumn(rows, title) {
  if (!Array.isArray(rows) || !title) return rows;

  const target = String(title).trim().toUpperCase();

  const cutIndex = rows.findIndex(row => {
    const firstKey = Object.keys(row)[0];
    const firstCell = String(row[firstKey] || "")
      .trim()
      .toUpperCase();

    return firstCell === target;
  });

  // ❌ không tìm thấy → giữ nguyên
  if (cutIndex === -1) return rows;

  // ✅ cắt từ row có title trở đi
  return rows.slice(0, cutIndex);
}

// [[60, 60, 60 "sbc"], [40, 40, 40 "abc"], [30, 30, 30]]
export function filter2DArrayByKeyword(arr, keyword) {
  if (!Array.isArray(arr) || !keyword) return arr;

  return arr.map(subArr =>
    subArr.filter(item => {
      if (typeof item !== "string") return true;
      return !item.includes(keyword);
    })
  );
}
