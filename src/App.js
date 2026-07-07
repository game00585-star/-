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

const formatMoneyPreview = (value) => {
  if (value === "" || value === null || value === undefined) return "";
  const numericValue = Number(String(value).replace(/,/g, ""));
  if (Number.isNaN(numericValue)) return String(value);
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

export default function App() {
  const [fileName, setFileName] = useState("");
  const [sheets, setSheets] = useState([]);
  const [activeSheet, setActiveSheet] = useState("");
  const [sheetRows, setSheetRows] = useState({});

  const rows = sheetRows[activeSheet] || [];
  const headerRowIndex = useMemo(() => findHeaderRowIndex(rows), [rows]);
  const headerRow = rows[headerRowIndex] || [];

  const handleFileUpload = async (event) => {
    const file = event.target.files?.[0];
    if (!file) return;

    const data = await file.arrayBuffer();
    const workbook = XLSX.read(data, { type: "array", cellDates: true });

    const nextSheetRows = {};
    workbook.SheetNames.forEach((sheetName) => {
      const sheet = workbook.Sheets[sheetName];
      const jsonRows = XLSX.utils.sheet_to_json(sheet, {
        header: 1,
        defval: "",
        raw: false,
      });
      nextSheetRows[sheetName] = normalizeRows(jsonRows);
    });

    setFileName(file.name);
    setSheets(workbook.SheetNames);
    setActiveSheet(workbook.SheetNames[0] || "");
    setSheetRows(nextSheetRows);
  };

  const updateCell = (rowIndex, cellIndex, value) => {
    setSheetRows((prev) => {
      const nextRows = (prev[activeSheet] || []).map((row) => [...row]);
      nextRows[rowIndex][cellIndex] = value;
      return { ...prev, [activeSheet]: nextRows };
    });
  };

  const addRow = () => {
    setSheetRows((prev) => {
      const currentRows = prev[activeSheet] || [];
      const columnCount = Math.max(1, ...currentRows.map((row) => row.length));
      const nextRows = [...currentRows, Array(columnCount).fill("")];
      return { ...prev, [activeSheet]: nextRows };
    });
  };

  const deleteRow = (rowIndex) => {
    if (!window.confirm("ต้องการลบแถวนี้ใช่ไหม?")) return;
    setSheetRows((prev) => {
      const nextRows = (prev[activeSheet] || []).filter((_, index) => index !== rowIndex);
      return { ...prev, [activeSheet]: nextRows };
    });
  };

  const exportExcel = () => {
    if (!sheets.length) {
      alert("กรุณาเลือกไฟล์ Excel ก่อน");
      return;
    }

    const workbook = XLSX.utils.book_new();

    sheets.forEach((sheetName) => {
      const cleanRows = (sheetRows[sheetName] || []).map((row) =>
        row.map((cell) => {
          const text = String(cell ?? "").trim();
          const numericText = text.replace(/,/g, "");
          if (text !== "" && !Number.isNaN(Number(numericText))) {
            return Number(numericText);
          }
          return cell;
        })
      );

      const worksheet = XLSX.utils.aoa_to_sheet(cleanRows);
      XLSX.utils.book_append_sheet(workbook, worksheet, sheetName.slice(0, 31));
    });

    const outputName = fileName
      ? fileName.replace(/\.xlsx?$/i, "") + "_แก้ไขแล้ว.xlsx"
      : "case_export_แก้ไขแล้ว.xlsx";

    XLSX.writeFile(workbook, outputName);
  };

  const exportActiveSheet = () => {
    if (!activeSheet) return;
    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.aoa_to_sheet(sheetRows[activeSheet] || []);
    XLSX.utils.book_append_sheet(workbook, worksheet, activeSheet.slice(0, 31));
    XLSX.writeFile(workbook, `${activeSheet}_แก้ไขแล้ว.xlsx`);
  };

  const getCellClassName = (cellIndex) => {
    const header = String(headerRow[cellIndex] || "");
    const classes = [];
    if (cellIndex === 0) classes.push("case-no");
    if (header.includes("ชื่อ")) classes.push("case-name");
    if (header.includes("สาขา") || header.includes("ภาค")) classes.push("case-branch");
    if (header.includes("หมายเหตุ") || header.includes("สภ") || header.includes("พนักงาน")) classes.push("case-note");
    if (moneyHeaders.some((moneyHeader) => header.includes(moneyHeader))) classes.push("money");
    return classes.join(" ");
  };

  return (
    <div className="case-page">
      <div className="case-header-card">
        <div>
          <h2>นำเข้าและแก้ไขไฟล์คดี</h2>
          <p>เลือกไฟล์ Excel แล้วแก้ไขข้อมูลในตาราง จากนั้น Export กลับออกมาเป็น Excel ได้</p>
        </div>

        <div className="case-actions">
          <label className="case-file-button">
            เลือกไฟล์ Excel
            <input type="file" accept=".xlsx,.xls" onChange={handleFileUpload} />
          </label>
          <button className="case-primary-button" onClick={exportExcel} disabled={!sheets.length}>
            Export Excel ทุกชีต
          </button>
          <button className="case-secondary-button" onClick={exportActiveSheet} disabled={!activeSheet}>
            Export เฉพาะชีตนี้
          </button>
        </div>
      </div>

      {fileName && <div className="case-file-name">ไฟล์: {fileName}</div>}

      {sheets.length > 0 && (
        <div className="case-tabs">
          {sheets.map((sheetName) => (
            <button
              key={sheetName}
              className={sheetName === activeSheet ? "active" : ""}
              onClick={() => setActiveSheet(sheetName)}
            >
              {sheetName}
            </button>
          ))}
        </div>
      )}

      {rows.length > 0 && (
        <>
          <div className="case-toolbar">
            <div>
              ชีตปัจจุบัน: <b>{activeSheet}</b> | จำนวนแถว: <b>{rows.length}</b>
            </div>
            <button className="case-secondary-button" onClick={addRow}>+ เพิ่มแถว</button>
          </div>

          <div className="case-table-wrap">
            <table className="case-table">
              <tbody>
                {rows.map((row, rowIndex) => {
                  const filledCellCount = row.filter((cell) => String(cell).trim() !== "").length;
                  const isTitleRow = rowIndex < headerRowIndex && filledCellCount <= 2;
                  const isHeaderRow = rowIndex === headerRowIndex;

                  if (isTitleRow) {
                    return (
                      <tr key={rowIndex} className="case-title-row">
                        <td colSpan={row.length + 1}>{row.join(" ")}</td>
                      </tr>
                    );
                  }

                  if (isHeaderRow) {
                    return (
                      <tr key={rowIndex}>
                        <th className="case-delete-col">จัดการ</th>
                        {row.map((cell, cellIndex) => (
                          <th key={cellIndex}>{formatDateValue(cell)}</th>
                        ))}
                      </tr>
                    );
                  }

                  return (
                    <tr key={rowIndex}>
                      <td className="case-delete-col">
                        <button className="case-delete-button" onClick={() => deleteRow(rowIndex)}>
                          ลบ
                        </button>
                      </td>
                      {row.map((cell, cellIndex) => {
                        const header = String(headerRow[cellIndex] || "");
                        const isMoney = moneyHeaders.some((moneyHeader) => header.includes(moneyHeader));

                        return (
                          <td key={cellIndex} className={getCellClassName(cellIndex)}>
                            <textarea
                              value={isMoney ? formatMoneyPreview(cell) : formatDateValue(cell)}
                              onChange={(event) => updateCell(rowIndex, cellIndex, event.target.value)}
                              rows={1}
                            />
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
