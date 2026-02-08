import { useEffect, useState } from "react";
import { getData, resetData, importExcel } from "../services/data.service";
import { ArrowUp } from "lucide-react";
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
  const [showScrollTop, setShowScrollTop] = useState(false);

  useEffect(() => {
    const handleScroll = () => {
      setShowScrollTop(window.scrollY > 300);
    };

    window.addEventListener("scroll", handleScroll);
    return () => window.removeEventListener("scroll", handleScroll);
  }, []);

  const scrollToTop = () => {
    window.scrollTo({
      top: 0,
      behavior: "smooth",
    });
  };


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

    try {
      const formData = new FormData();
      formData.append("file", file);

      const result = await importExcel(formData);

      setSheetNames(result.sheetNames || []);
      setSheetData(result.data || {});

      if (result.sheetNames?.length > 0) {
        setCurrentSheet(result.sheetNames[0]);
      }
    } catch (err) {
      console.error("Import excel error:", err);
    } finally {
      setLoading(false);
    }
  };

  const handleImportInGrandTotal = async (e) => {
    const file = e.target.files[0];
    if (!file) return;

    setLoading(true);

    try {
      const formData = new FormData();
      formData.append("file", file);

      const result = await importExcel(formData);

      // sau khi import xong → reload Grand Total
      await loadData();

      // alert("Import Excel thành công & Grand Total đã được cập nhật");
    } catch (err) {
      console.error(err);
      alert("Import Excel thất bại");
    } finally {
      setLoading(false);
      e.target.value = ""; // reset input để chọn lại file
    }
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
    setTab("data");
    setShowResetConfirm(true);
  };

  const confirmResetExcel = async () => {
    setShowResetConfirm(false);
    setTab("data");
    setDataRows([]);
    setLoading(true);

    try {
      // const res = await fetch("http://10.0.0.236:5000/api/excel/reset", {
      //   method: "POST",
      // });
      const res = await resetData();
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
      <h2 className="title">SCANNING TOOL</h2>

      {/* TAB BUTTONS */}
      <div className="tabs">
        <button
          className={tab === "data" ? "tab active" : "tab"}
          onClick={() => setTab("data")}
        >
          📊 Grand Total
        </button>

        {/* <button
          className={tab === "import" ? "tab active" : "tab"}
          onClick={() => setTab("import")}
        >
          🔍 Import & Scan
        </button> */}

        <button
          className={tab === "guide" ? "tab active" : "tab"}
          onClick={() => setTab("guide")}
        >
          📘 Hướng dẫn sử dụng
        </button>
      </div>

      {/* {loading && <p className="loading">Loading...</p>} */}

      {/* TAB 1 */}
      {tab === "data" && (
        <div>
          <div className="grand-header">
            <h3>📊 Grand Total</h3>

            <div className="grand-actions">

              <label className="import-btn">
                📥 Import Excel
                <input
                  type="file"
                  accept=".xlsx,.xls"
                  hidden
                  onChange={handleImportInGrandTotal}
                />
              </label>

              <button
                className="export-btn"
                onClick={handleExportExcel}
                disabled={dataRows.length === 0}
              >
                ⬇ Export Excel
              </button>

              <button
                className="reset-btn"
                onClick={handlResetExcel}
                disabled={dataRows.length === 0}
              >
                🔄 Reset Excel
              </button>

            </div>
          </div>

          <div className="table-wrapper">
            {dataRows.length === 0 ? (
              <p className="empty">📭 Chưa có dữ liệu Grand Total</p>
            ) : (
              renderTable(dataRows)
            )}
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
              <b>📊 Grand Total</b>
              <p>
                Tab <b>Grand Total</b> là nơi hiển thị toàn bộ dữ liệu đã được tổng hợp từ
                các file Excel đã import.
              </p>
              <p>
                Tại đây, bạn có thể xem bảng dữ liệu tổng hợp và theo dõi các chỉ
                số quan trọng.
              </p>
            </li>

            <li>
              <b>📥 Import Excel</b>
              <p>
                Trong tab <b>Grand Total</b>, nhấn nút <b>Import Excel</b> để chọn file Excel
                (<code>.xls</code> hoặc <code>.xlsx</code>) cần xử lý.
              </p>
              <p>
                Hệ thống sẽ tự động quét, xử lý dữ liệu và cập nhật vào bảng
                <b> Grand Total</b> sau khi import hoàn tất.
              </p>
            </li>

            <li>
              <b>📤 Export Excel</b>
              <p>
                Sau khi dữ liệu đã được cập nhật, nhấn nút <b>Export Excel</b> để tải toàn
                bộ dữ liệu Grand Total ra file Excel.
              </p>
              <p>
                File xuất ra dùng cho việc báo cáo, lưu trữ hoặc chia sẻ dữ liệu.
              </p>
            </li>

            <li>
              <b>🔄 Reset dữ liệu</b>
              <p>
                Nhấn nút <b>Reset Excel</b> để xoá toàn bộ dữ liệu hiện có trong
                <b> Grand Total</b>.
              </p>
              <p>
                Hệ thống sẽ yêu cầu xác nhận trước khi thực hiện. Dữ liệu sau khi reset
                <b> không thể khôi phục</b>.
              </p>
            </li>

            <li>
              <b>🔎 Kiểm tra dữ liệu</b>
              <p>
                Kiểm tra lại các chỉ số và dữ liệu trong bảng Grand Total để đảm bảo độ
                chính xác trước khi sử dụng cho báo cáo chính thức.
              </p>
            </li>
          </ol>

          <div className="guide-note">
            ⚠️ <b>Lưu ý:</b> Mỗi lần import Excel sẽ <b>tự động cập nhật</b> dữ liệu trong
            Grand Total. Chức năng Export chỉ xuất dữ liệu đang hiển thị tại thời điểm
            xuất.
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
            {/* <p className="modal-warning">
              Hành động này không thể hoàn tác.
            </p> */}

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

      {showScrollTop && (
        <button
          className="scroll-top-btn"
          onClick={scrollToTop}
          title="Lên đầu trang"
        >
          <ArrowUp size={22} strokeWidth={2.5} />
        </button>
      )}
    </div>
  );
}
