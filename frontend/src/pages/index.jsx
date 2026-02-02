import { useEffect, useState } from "react";
import { getData } from "../services/data.service";
import './index.css'
import * as XLSX from "xlsx";

export default function IndexPage() {
  const [tab, setTab] = useState("data");

  // TAB 1: DATA
  const [dataRows, setDataRows] = useState([]);

  // TAB 2: IMPORT PREVIEW
  const [sheetNames, setSheetNames] = useState([]);
  const [sheetData, setSheetData] = useState({});
  const [currentSheet, setCurrentSheet] = useState("");
  const [showResetConfirm, setShowResetConfirm] = useState(false);
  const [loading, setLoading] = useState(false);

  // load DATA tab
  useEffect(() => {
    if (tab === "data") loadData();
  }, [tab]);

  const loadData = async () => {
    setLoading(true);
    const data = await getData();
    setDataRows(Array.isArray(data) ? data : []);
    setLoading(false);
  };

  // PREVIEW IMPORT
  const handlePreviewImport = async (e) => {
    const file = e.target.files[0];
    if (!file) return;

    setLoading(true);

    const formData = new FormData();
    formData.append("file", file);

    const res = await fetch("http://localhost:5000/api/excel/preview", {
      method: "POST",
      body: formData,
    });

    const result = await res.json();

    setSheetNames(result.sheetNames || []);
    setSheetData(result.data || {});

    if (result.sheetNames?.length > 0) {
      setCurrentSheet(result.sheetNames[0]);
    }

    setLoading(false);
  };

  const handleExportExcel = () => {
    if (!dataRows || dataRows.length === 0) {
      alert("Không có dữ liệu Grand Total để xuất");
      return;
    }

    const worksheet = XLSX.utils.json_to_sheet(dataRows);
    const workbook = XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(workbook, worksheet, "Grand Total");

    const today = new Date().toISOString().slice(0, 10);
    XLSX.writeFile(workbook, `grand_total_${today}.xlsx`);
  };

  const handlResetExcel = () => {
    setShowResetConfirm(true);
  };

  const confirmResetExcel = async () => {
    setShowResetConfirm(false);
    setLoading(true);

    try {
      const res = await fetch("http://localhost:5000/api/excel/reset", {
        method: "POST",
      });

      if (!res.ok) throw new Error("Reset failed");

      await loadData();
      alert("Đã reset dữ liệu Grand Total");
    } catch (err) {
      alert("Reset thất bại, vui lòng thử lại");
    } finally {
      setLoading(false);
    }
  };



  const renderTable = (rows) => (
    <table border="1" cellPadding="6">
      <thead>
        <tr>
          {rows.length > 0 &&
            Object.keys(rows[0]).map((k) => <th key={k}>{k}</th>)}
        </tr>
      </thead>
      <tbody>
        {rows.map((row, i) => (
          <tr key={i}>
            {Object.values(row).map((v, j) => (
              <td key={j}>{v}</td>
            ))}
          </tr>
        ))}
      </tbody>
    </table>
  );

  const previewRows = sheetData[currentSheet] || [];

  return (
    <div className="container">
      <h2 className="title">Scanning Tool</h2>

      {/* TAB BUTTONS */}
      <div className="tabs">
        <button
          className={tab === "data" ? "tab active" : "tab"}
          onClick={() => setTab("data")}
        >
          📊 Grand Total
        </button>

        <button
          className={tab === "import" ? "tab active" : "tab"}
          onClick={() => setTab("import")}
        >
          🔍 Import & Scan
        </button>

        <button
          className={tab === "guide" ? "tab active" : "tab"}
          onClick={() => setTab("guide")}
        >
          📘 Hướng dẫn sử dụng
        </button>
      </div>

      {loading && <p className="loading">Loading...</p>}

      {/* TAB 1 */}
      {tab === "data" && !loading && (
        <div>
          <div className="grand-header">
            <h3>📊 Grand Total</h3>

            <div className="grand-actions">
              <button
                className="reset-btn"
                onClick={handlResetExcel}
                disabled={dataRows.length === 0}
              >
                🔄 Reset Excel
              </button>

              <button
                className="export-btn"
                onClick={handleExportExcel}
                disabled={dataRows.length === 0}
              >
                ⬇ Export Excel
              </button>

            </div>
          </div>

          <div className="table-wrapper">
            {renderTable(dataRows)}
          </div>
        </div>
      )}

      {/* TAB 2 */}
      {tab === "import" && (
        <div className="import-box">
          <input
            className="file-input"
            type="file"
            accept=".xlsx,.xls"
            onChange={handlePreviewImport}
          />

          {sheetNames.length > 0 && (
            <select
              className="sheet-select"
              value={currentSheet}
              onChange={(e) => setCurrentSheet(e.target.value)}
            >
              {sheetNames.map((name) => (
                <option key={name} value={name}>
                  {name}
                </option>
              ))}
            </select>
          )}

          {!loading && previewRows.length > 0 && (
            <div className="table-wrapper">{renderTable(previewRows)}</div>
          )}
        </div>
      )}

      {/* TAB 3: GUIDE */}
      {tab === "guide" && (
        <div className="guide">
          <h3>📘 Hướng dẫn sử dụng Scanning Tool</h3>
          <ol>
            <li>
              <b>🔍 Import & Scan</b>
              <p>
                Chọn file Excel (<code>.xls</code> hoặc <code>.xlsx</code>) để hệ thống
                tự động quét các chỉ số cần thiết.
              </p>
              <p>
                Sau khi upload, dữ liệu sẽ được xử lý và tổng hợp tự động.
              </p>
            </li>

            <li>
              <b>📊 Grand Total</b>
              <p>
                Sau khi quá trình import & scan hoàn tất, tab <b>Grand Total</b> sẽ được
                cập nhật với dữ liệu mới nhất.
              </p>
              <p>
                Người dùng có thể xem bảng dữ liệu tổng hợp và thực hiện xuất báo cáo.
              </p>
            </li>

            <li>
              <b>📤 Xuất Excel</b>
              <p>
                Tại tab <b>Grand Total</b>, nhấn nút <b>Export Excel</b> để tải dữ liệu
                tổng hợp ra file Excel.
              </p>
              <p>
                File xuất ra phục vụ cho việc báo cáo, lưu trữ hoặc chia sẻ dữ liệu.
              </p>
            </li>

            <li>
              <b>🔎 Kiểm tra dữ liệu</b>
              <p>
                Kiểm tra các chỉ số đã được quét và dữ liệu trong file Excel xuất ra
                để đảm bảo độ chính xác trước khi sử dụng cho báo cáo chính thức.
              </p>
            </li>
          </ol>

          <div className="guide-note">
            ⚠️ <b>Lưu ý:</b> Mỗi lần import sẽ <b>tự động cập nhật</b> dữ liệu trong
            Grand Total. Chức năng Export chỉ xuất dữ liệu đang hiển thị.
          </div>
        </div>
      )}


      {showResetConfirm && (
        <div className="modal-overlay">
          <div className="modal">
            <h3>⚠️ Xác nhận reset dữ liệu</h3>

            <p>
              Bạn có chắc chắn muốn <b>reset toàn bộ dữ liệu Grand Total</b> không?
            </p>
            <p className="modal-warning">
              Hành động này không thể hoàn tác.
            </p>

            <div className="modal-actions">
              <button
                className="btn-cancel"
                onClick={() => setShowResetConfirm(false)}
              >
                Hủy
              </button>

              <button
                className="btn-danger"
                onClick={confirmResetExcel}
              >
                Reset
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
