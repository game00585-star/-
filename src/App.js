import React, { useMemo, useState } from "react";
import * as XLSX from "xlsx";
import "./App.css";

const moneyHeaders = [
  "หนี้/ความเสียหาย",
  "ชำระแล้ว",
  "คงเหลือ",
  "หนังสือรับสภาพหนี้",
  "ยอดหนี้",
  "ยอดชำระ",
  "จำนวนเงิน",
  "รวม",
];

const wrapHeaders = ["หมายเหตุ", "ความคืบหน้า", "พนักงาน", "อัยการ", "สภ", "ศาล", "ชำระล่าสุด"];

const findHeaderRowIndex = (rows) => {
  const index = rows.findIndex(
    (row) => row.includes("ที่") && row.some((cell) => String(cell).includes("ชื่อ"))
  );
  return index >= 0 ? index : 0;
};

const formatDateValue = (value) => {
  if (value instanceof Date) return value.toLocaleDateString("th-TH");
  return String(value ?? "");
};

const toNumber = (value) => {
  if (value === "" || value === null || value === undefined) return 0;
  const text = String(value).replace(/,/g, "").replace(/-/g, "0").trim();
  const num = Number(text);
  return Number.isFinite(num) ? num : 0;
};

const formatMoneyPreview = (value) => {
  if (value === "" || value === null || value === undefined) return "";
  const textValue = String(value).trim();
  const numericValue = toNumber(value);
  if (!numericValue && textValue !== "0" && textValue !== "-") return String(value);
  return numericValue.toLocaleString("th-TH", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
};

const normalizeRows = (rows) => {
  const maxLength = Math.max(0, ...rows.map((row) => row.length));
  return rows.map((row) => {
    const nextRow = [...row];
    while (nextRow.length < maxLength) nextRow.push("");
    return nextRow;
  });
};

const findColumnIndex = (headerRow, keywords) => {
  return headerRow.findIndex((header) =>
    keywords.some((keyword) => String(header || "").includes(keyword))
  );
};

const isRealCaseRow = (row, rowIndex, headerRowIndex) => {
  if (rowIndex <= headerRowIndex) return false;
  const firstCell = String(row[0] ?? "").trim();
  if (!firstCell) return false;
  if (firstCell.includes("รวม")) return false;
  return !Number.isNaN(Number(firstCell));
};

const convertCellForExport = (cell) => {
  const text = String(cell ?? "").trim();
  const numericText = text.replace(/,/g, "");
  if (text !== "" && !Number.isNaN(Number(numericText))) return Number(numericText);
  return cell;
};

const getColumnWidthFromHeader = (header) => {
  const text = String(header || "");
  if (text === "ที่") return { min: 62, max: 72 };
  if (text.includes("ชื่อ")) return { min: 230, max: 320 };
  if (text.includes("สาขา") || text.includes("ภาค")) return { min: 220, max: 320 };
  if (text.includes("ศาล")) return { min: 170, max: 250 };
  if (text.includes("เลขคดี")) return { min: 130, max: 170 };
  if (text.includes("ทนาย")) return { min: 130, max: 190 };
  if (text.includes("ความคืบหน้า") || text.includes("หมายเหตุ")) return { min: 260, max: 430 };
  if (text.includes("ชำระล่าสุด")) return { min: 125, max: 175 };
  if (text.includes("วัน") || text.includes("เดือน") || text.includes("ปี")) return { min: 135, max: 170 };
  if (moneyHeaders.some((moneyHeader) => text.includes(moneyHeader))) return { min: 130, max: 180 };
  if (text.includes("สภ") || text.includes("พนักงาน") || text.includes("อัยการ")) return { min: 190, max: 330 };
  return { min: 120, max: 220 };
};

const measureTextWidth = (value) => {
  const text = String(value ?? "").replace(/\s+/g, " ").trim();
  if (!text) return 48;
  let width = 34;
  for (const char of text) {
    if (/\d|[A-Za-z]/.test(char)) width += 7;
    else if (/[,./:-]/.test(char)) width += 5;
    else width += 12;
  }
  return width;
};

const clamp = (value, min, max) => Math.max(min, Math.min(max, value));

export default function App() {
  const [fileName, setFileName] = useState("");
  const [sheets, setSheets] = useState([]);
  const [activeSheet, setActiveSheet] = useState("");
  const [sheetRows, setSheetRows] = useState({});

  const rows = sheetRows[activeSheet] || [];
  const headerRowIndex = useMemo(() => findHeaderRowIndex(rows), [rows]);
  const headerRow = rows[headerRowIndex] || [];

  const columnStyles = useMemo(() => {
    const columnCount = Math.max(0, ...rows.map((row) => row.length));
    return Array.from({ length: columnCount }, (_, cellIndex) => {
      const header = headerRow[cellIndex] || "";
      const limit = getColumnWidthFromHeader(header);
      const sampleRows = rows.slice(0, 120);
      const contentWidth = Math.max(
        measureTextWidth(header),
        ...sampleRows.map((row) => measureTextWidth(row[cellIndex]))
      );
      const width = clamp(contentWidth + 20, limit.min, limit.max);
      return { width: `${width}px`, minWidth: `${width}px`, maxWidth: `${width}px` };
    });
  }, [rows, headerRow]);

  const debtCol = useMemo(() => findColumnIndex(headerRow, ["หนี้/ความเสียหาย", "ยอดหนี้"]), [headerRow]);
  const paidCol = useMemo(() => findColumnIndex(headerRow, ["ชำระแล้ว", "ยอดชำระ"]), [headerRow]);
  const remainCol = useMemo(() => findColumnIndex(headerRow, ["คงเหลือ"]), [headerRow]);
  const documentCol = useMemo(() => findColumnIndex(headerRow, ["หนังสือรับสภาพหนี้"]), [headerRow]);

  const summary = useMemo(() => {
    const result = { caseCount: 0, debt: 0, paid: 0, remain: 0, document: 0, diff: 0, isBalanced: true };
    rows.forEach((row, rowIndex) => {
      if (!isRealCaseRow(row, rowIndex, headerRowIndex)) return;
      result.caseCount += 1;
      if (debtCol >= 0) result.debt += toNumber(row[debtCol]);
      if (paidCol >= 0) result.paid += toNumber(row[paidCol]);
      if (remainCol >= 0) result.remain += toNumber(row[remainCol]);
      if (documentCol >= 0) result.document += toNumber(row[documentCol]);
    });
    result.diff = result.debt - result.paid - result.remain;
    result.isBalanced = Math.abs(result.diff) < 0.01;
    return result;
  }, [rows, headerRowIndex, debtCol, paidCol, remainCol, documentCol]);

  const handleFileUpload = async (event) => {
    const file = event.target.files?.[0];
    if (!file) return;
    const data = await file.arrayBuffer();
    const workbook = XLSX.read(data, { type: "array", cellDates: true });
    const nextSheetRows = {};
    workbook.SheetNames.forEach((sheetName) => {
      const sheet = workbook.Sheets[sheetName];
      const jsonRows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", raw: false });
      nextSheetRows[sheetName] = normalizeRows(jsonRows);
    });
    setFileName(file.name);
    setSheets(workbook.SheetNames);
    setActiveSheet(workbook.SheetNames[0] || "");
    setSheetRows(nextSheetRows);
  };

  const exportExcel = () => {
    if (!sheets.length) {
      alert("กรุณาเลือกไฟล์ Excel ก่อน");
      return;
    }
    const workbook = XLSX.utils.book_new();
    sheets.forEach((sheetName) => {
      const cleanRows = (sheetRows[sheetName] || []).map((row) => row.map(convertCellForExport));
      const worksheet = XLSX.utils.aoa_to_sheet(cleanRows);
      XLSX.utils.book_append_sheet(workbook, worksheet, sheetName.slice(0, 31));
    });
    const outputName = fileName ? fileName.replace(/\.xlsx?$/i, "") + "_export.xlsx" : "case_export.xlsx";
    XLSX.writeFile(workbook, outputName);
  };

  const exportActiveSheet = () => {
    if (!activeSheet) return;
    const workbook = XLSX.utils.book_new();
    const cleanRows = (sheetRows[activeSheet] || []).map((row) => row.map(convertCellForExport));
    const worksheet = XLSX.utils.aoa_to_sheet(cleanRows);
    XLSX.utils.book_append_sheet(workbook, worksheet, activeSheet.slice(0, 31));
    XLSX.writeFile(workbook, `${activeSheet}_export.xlsx`);
  };

  const getCellClassName = (cellIndex) => {
    const header = String(headerRow[cellIndex] || "");
    const classes = [];
    if (cellIndex === 0) classes.push("case-no");
    if (header.includes("ชื่อ")) classes.push("case-name");
    if (header.includes("สาขา") || header.includes("ภาค")) classes.push("case-branch");
    if (header.includes("ศาล")) classes.push("case-court");
    if (header.includes("ทนาย")) classes.push("case-lawyer");
    if (wrapHeaders.some((key) => header.includes(key))) classes.push("wrap-text");
    if (moneyHeaders.some((moneyHeader) => header.includes(moneyHeader))) classes.push("money");
    return classes.join(" ");
  };

  const shouldShowFormattedMoney = (cellIndex) => {
    const header = String(headerRow[cellIndex] || "");
    return moneyHeaders.some((moneyHeader) => header.includes(moneyHeader));
  };

  return (
    <div className="case-page">
      <div className="case-header-card">
        <div>
          <h2>นำเข้าไฟล์คดี</h2>
          <p>เลือกไฟล์ Excel เพื่อดูข้อมูลหลายชีตแบบอ่านง่าย และ Export กลับออกมาเป็น Excel ได้</p>
        </div>
        <div className="case-actions">
          <label className="case-file-button">
            เลือกไฟล์ Excel
            <input type="file" accept=".xlsx,.xls" onChange={handleFileUpload} />
          </label>
          <button className="case-primary-button" onClick={exportExcel} disabled={!sheets.length}>Export Excel ทุกชีต</button>
          <button className="case-secondary-button" onClick={exportActiveSheet} disabled={!activeSheet}>Export เฉพาะชีตนี้</button>
        </div>
      </div>

      {fileName && <div className="case-file-name">ไฟล์: {fileName}</div>}

      {sheets.length > 0 && (
        <div className="case-tabs">
          {sheets.map((sheetName) => (
            <button key={sheetName} className={sheetName === activeSheet ? "active" : ""} onClick={() => setActiveSheet(sheetName)}>
              {sheetName}
            </button>
          ))}
        </div>
      )}

      {rows.length > 0 && (
        <>
          <div className="case-summary-grid">
            <div className="case-summary-card"><span>จำนวนรายการ</span><b>{summary.caseCount}</b></div>
            <div className="case-summary-card"><span>หนี้/ความเสียหาย</span><b>{formatMoneyPreview(summary.debt)}</b></div>
            <div className="case-summary-card"><span>ชำระแล้ว</span><b>{formatMoneyPreview(summary.paid)}</b></div>
            <div className="case-summary-card"><span>คงเหลือ</span><b>{formatMoneyPreview(summary.remain)}</b></div>
            <div className="case-summary-card"><span>รับสภาพหนี้</span><b>{formatMoneyPreview(summary.document)}</b></div>
            <div className={summary.isBalanced ? "case-summary-card ok" : "case-summary-card bad"}>
              <span>ตรวจสมดุล</span><b>{summary.isBalanced ? "สมดุล" : `ต่าง ${formatMoneyPreview(summary.diff)}`}</b>
            </div>
          </div>

          <div className="case-toolbar">
            <div>ชีตปัจจุบัน: <b>{activeSheet}</b> | จำนวนแถวทั้งหมด: <b>{rows.length}</b> | แถวข้อมูลจริง: <b>{summary.caseCount}</b></div>
          </div>

          <div className="case-table-wrap">
            <table className="case-table read-only-table">
              <tbody>
                {rows.map((row, rowIndex) => {
                  const filledCellCount = row.filter((cell) => String(cell).trim() !== "").length;
                  const isTitleRow = rowIndex < headerRowIndex && filledCellCount <= 2;
                  const isHeaderRow = rowIndex === headerRowIndex;
                  const isTotalRow = row.some((cell) => String(cell).includes("รวม"));

                  if (isTitleRow) {
                    return <tr key={rowIndex} className="case-title-row"><td colSpan={row.length}>{row.join(" ")}</td></tr>;
                  }
                  if (isHeaderRow) {
                    return (
                      <tr key={rowIndex}>
                        {row.map((cell, cellIndex) => <th key={cellIndex} style={columnStyles[cellIndex]}>{formatDateValue(cell)}</th>)}
                      </tr>
                    );
                  }
                  return (
                    <tr key={rowIndex} className={isTotalRow ? "case-total-row" : ""}>
                      {row.map((cell, cellIndex) => {
                        const isMoney = shouldShowFormattedMoney(cellIndex);
                        const value = isMoney ? formatMoneyPreview(cell) : formatDateValue(cell);
                        return (
                          <td key={cellIndex} className={getCellClassName(cellIndex)} style={columnStyles[cellIndex]} title={value}>
                            <span className="cell-text">{value}</span>
                          </td>
                        );
                      })}
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </>
      )}
    </div>
  );
}
