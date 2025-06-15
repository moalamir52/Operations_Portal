// ContractVlookup.tsx - Enhanced UI Matching Reference Image
import React, { useState, useEffect } from "react";
import * as XLSX from "xlsx";

function formatExcelDate(value) {
  if (typeof value === 'number') {
    const date = new Date(Math.round((value - 25569) * 86400 * 1000));
    const day = String(date.getDate()).padStart(2, '0');
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const year = date.getFullYear();
    return `${day}/${month}/${year}`;
  }
  return value || "❌";
}

export default function ContractVlookup() {
  const yeloColors = {
    primary: "#ffde38",
    secondary: "#7b1fa2",
    darkPurple: "#5d1789",
    lightPurple: "#9c4dcc",
    tertiary: "#fffce6",
    textDark: "#333333",
    white: "#FFFFFF",
    offWhite: "#f5f5f5",
    success: "#4CAF50",
    error: "#E53935",
    warning: "#FF9800",
    lightYellow: "#fff8dd",
  };

  const [refData, setRefData] = useState([]);
  const [uploadedData, setUploadedData] = useState([]);
  const [results, setResults] = useState(() => {
    const saved = localStorage.getItem('vlookupResults');
    return saved ? JSON.parse(saved) : [];
  });
  const [selectedColumns, setSelectedColumns] = useState(() => {
    const saved = localStorage.getItem('selectedColumns');
    return saved ? JSON.parse(saved) : ['plate', 'model', 'chassis', 'regExpiry', 'insExpiry', 'color'];
  });
  const [copyMessage, setCopyMessage] = useState("");
  const [error, setError] = useState("");

  useEffect(() => {
    localStorage.setItem('vlookupResults', JSON.stringify(results));
  }, [results]);

  useEffect(() => {
    localStorage.setItem('selectedColumns', JSON.stringify(selectedColumns));
  }, [selectedColumns]);

  useEffect(() => {
    const url =
      "https://docs.google.com/spreadsheets/d/1sHvEQMtt3suuxuMA0zhcXk5TYGqZzit0JvGLk1CQ0LI/export?format=csv&gid=804568597";
    fetch(url)
      .then((res) => res.text())
      .then((csv) => {
        const workbook = XLSX.read(csv, { type: "string" });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(sheet);
        setUploadedData(json);
      })
      .catch((err) => {
        console.error(err);
        setError("Failed to load Google Sheet data");
      });
  }, []);

  const handleUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = new Uint8Array(evt.target.result);
        const workbook = XLSX.read(data, { type: "array" });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(sheet);
        setRefData(json);
        performVlookup(json);
      } catch (err) {
        setError("Error processing file: " + err.message);
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const performVlookup = (refDataFromUpload) => {
    const refMap = new Map();
    refDataFromUpload.forEach(row => {
      const key = row["plate"]?.toString().trim();
      if (key) refMap.set(key, row);
    });

    const result = uploadedData.map((sheetRow) => {
      const plateNo = sheetRow["Plate No"]?.toString().trim();
      const match = refMap.get(plateNo);
      return {
        plate: plateNo,
        model: match?.["model"] || "❌",
        chassis: match?.["chassis"] || "❌",
        regExpiry: formatExcelDate(match?.["regExpiry"]),
        insExpiry: formatExcelDate(match?.["insExpiry"]),
        color: match?.["color"] || "❌",
      };
    });
    setResults(result);
  };

  const toggleColumn = (col) => {
    setSelectedColumns(prev => prev.includes(col)
      ? prev.filter(c => c !== col)
      : [...prev, col]);
  };

  const copySelectedColumns = () => {
    const content = results.map((row, i) => [i + 1, ...selectedColumns.map(col => row[col])].join("\t")).join("\n");
    navigator.clipboard.writeText(content).then(() => {
      setCopyMessage("Copied selected columns!");
      setTimeout(() => setCopyMessage(""), 2000);
    });
  };

  const exportToExcel = () => {
    const ws = XLSX.utils.json_to_sheet(results.map((r, i) => ({ "#": i + 1, ...r })));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Results");
    XLSX.writeFile(wb, "YELO_vlookup_results.xlsx");
  };

  const missingCount = results.filter(r => Object.values(r).includes("❌")).length;

  return (
    <div style={{ padding: 20, fontFamily: "'Segoe UI', sans-serif", background: yeloColors.tertiary }}>
      <div style={{ textAlign: "center", marginBottom: 20 }}>
      </div>

      <div style={{ textAlign: "center", marginBottom: 20 }}>
        <div style={{ display: "inline-block", background: yeloColors.primary, color: yeloColors.secondary, padding: "15px 40px", fontWeight: "bold", fontSize: 24, borderRadius: 10, boxShadow: `0 4px 0 ${yeloColors.darkPurple}` }}>
          YELO Fleet Lookup Tool With RTA
        </div>
      </div>

      <div style={{ textAlign: "center", marginBottom: 15, display: "flex", justifyContent: "center", flexWrap: "wrap", gap: "10px" }}>
        <label style={{ background: yeloColors.primary, padding: "10px 20px", borderRadius: 5, cursor: "pointer", fontWeight: "bold", boxShadow: `0 3px 0 ${yeloColors.secondary}` }}>
          Upload Excel File
          <input type="file" accept=".xlsx,.xls,.csv" onChange={handleUpload} style={{ display: "none" }} />
        </label>
        <button onClick={() => setResults([])} style={{ background: yeloColors.error, color: yeloColors.white, padding: "10px 20px", borderRadius: 5, fontWeight: "bold", boxShadow: `0 3px 0 #b71c1c`, border: "none" }}>Clear Results</button>
        <button onClick={copySelectedColumns} style={{ background: yeloColors.primary, color: yeloColors.secondary, padding: "10px 20px", borderRadius: 5, fontWeight: "bold", boxShadow: `0 3px 0 ${yeloColors.darkPurple}`, border: "none" }}>Copy Selected Columns</button>
        <button onClick={exportToExcel} style={{ background: yeloColors.secondary, color: yeloColors.white, padding: "10px 20px", borderRadius: 5, fontWeight: "bold", boxShadow: `0 3px 0 ${yeloColors.darkPurple}`, border: "none" }}>Export to Excel</button>
      </div>

      <div style={{ marginBottom: 15, textAlign: "center" }}>
        <div style={{ background: yeloColors.white, padding: 15, borderRadius: 8, display: "inline-block", boxShadow: "0 1px 3px rgba(0,0,0,0.1)" }}>
          <div style={{ fontWeight: "bold", fontSize: 16, color: yeloColors.secondary }}>Total Records: {results.length}</div>
          <div style={{ display: "flex", justifyContent: "center", marginTop: 10, gap: 10 }}>
            <div style={{ background: yeloColors.success, color: yeloColors.white, padding: "5px 15px", borderRadius: 5 }}>Complete Records: {results.length - missingCount}</div>
            <div style={{ background: yeloColors.warning, color: yeloColors.white, padding: "5px 15px", borderRadius: 5 }}>Missing Data Records: {missingCount}</div>
          </div>
        </div>
      </div>

      {copyMessage && <div style={{ color: yeloColors.success, textAlign: "center", fontWeight: "bold" }}>{copyMessage}</div>}
      {error && <div style={{ color: yeloColors.error, textAlign: "center" }}>{error}</div>}

      <div style={{ overflowX: "auto" }}>
        <table style={{ width: "100%", borderCollapse: "collapse", background: yeloColors.white }}>
          <thead style={{ background: yeloColors.secondary, color: yeloColors.white }}>
            <tr>
              <th>#</th>
              {['plate', 'model', 'chassis', 'regExpiry', 'insExpiry', 'color'].map(col => (
                <th key={col} style={{ cursor: "pointer", padding: "10px" }} onClick={() => toggleColumn(col)}>
                  <input type="checkbox" checked={selectedColumns.includes(col)} readOnly style={{ marginRight: 5 }} />
                  {col}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {results.map((row, i) => {
              const isMissing = Object.values(row).includes("❌");
              return (
                <tr key={i} style={{ backgroundColor: isMissing ? yeloColors.lightYellow : (i % 2 === 0 ? yeloColors.white : yeloColors.offWhite) }}>
                  <td>{i + 1}</td>
                  {['plate', 'model', 'chassis', 'regExpiry', 'insExpiry', 'color'].map(col => (
                    <td key={col} style={{ padding: "8px", color: row[col] === "❌" ? yeloColors.error : undefined }}>{row[col]}</td>
                  ))}
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
}
