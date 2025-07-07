import React, { useState, useEffect } from "react";
import * as XLSX from "xlsx";

export default function ContractVlookup() {
  // YELO Official Colors - Removed Gray
  const yeloColors = {
    primary: "#ffde38", // اللون الأصفر المعتمد
    secondary: "#7b1fa2", // اللون البنفسجي المعتمد
    darkPurple: "#5d1789", // بنفسجي أغمق للتأثيرات (مشتق من اللون المعتمد)
    lightPurple: "#9c4dcc", // بنفسجي أفتح للتأثيرات (مشتق من اللون المعتمد)
    tertiary: "#f8f8f8", // لون خلفية فاتح
    textDark: "#333333", // لون داكن للنصوص بدلاً من الرمادي
    white: "#FFFFFF",
    offWhite: "#f5f5f5", // أبيض مائل للدفء بدلاً من الرمادي الفاتح
    success: "#4CAF50",
    error: "#E53935",
    warning: "#FF9800",
    lightYellow: "#fff8dd", // أصفر فاتح للتأثيرات (مشتق من اللون المعتمد)
  };

  const [refData, setRefData] = useState([]);
  const [uploadedData, setUploadedData] = useState([]);
  const [results, setResults] = useState(() => {
    const savedResults = localStorage.getItem('vlookupResults');
    return savedResults ? JSON.parse(savedResults) : [];
  });
  const [copyMessage, setCopyMessage] = useState("");
  const [selectedColumns, setSelectedColumns] = useState(() => {
    const savedColumns = localStorage.getItem('selectedColumns');
    return savedColumns ? JSON.parse(savedColumns) : [];
  });
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState("");
  
  // Counter for missing data
  const getMissingRecords = () => {
    if (!results.length) return 0;
    
    return results.filter(record => 
      record.plate === "❌" || 
      record.model === "❌" || 
      record.pickup === "❌" || 
      record.dropoff === "❌"
    ).length;
  };
  
  const getCompleteRecords = () => {
    if (!results.length) return 0;
    
    return results.length - getMissingRecords();
  };

  // Styles based on YELO official colors
  const buttonStyle = {
    marginLeft: 10,
    padding: "10px 20px",
    color: yeloColors.white,
    border: "none",
    borderRadius: "4px",
    cursor: "pointer",
    fontWeight: "bold",
    boxShadow: "0 2px 5px rgba(0,0,0,0.1)",
    transition: "all 0.3s ease",
  };

  // حفظ النتائج في localStorage كلما تتغير
  useEffect(() => {
    if (results.length > 0) {
      localStorage.setItem('vlookupResults', JSON.stringify(results));
    }
  }, [results]);

  // حفظ الأعمدة المحددة في localStorage كلما تتغير
  useEffect(() => {
    localStorage.setItem('selectedColumns', JSON.stringify(selectedColumns));
  }, [selectedColumns]);

  useEffect(() => {
    const url =
      "https://docs.google.com/spreadsheets/d/1XwBko5v8zOdTdv-By8HK_DvZnYT2T12mBw_SIbCfMkE/export?format=csv&gid=769459790";
    
    setIsLoading(true);
    fetch(url)
      .then((res) => res.text())
      .then((csv) => {
        const workbook = XLSX.read(csv, { type: "string" });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(sheet);
        setRefData(json);
      })
      .catch((err) => {
        console.error(err);
        setError("Failed to load reference data");
      })
      .finally(() => setIsLoading(false));
  }, []);

  const handleUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    
    setIsLoading(true);
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = new Uint8Array(evt.target.result);
        const workbook = XLSX.read(data, { type: "array" });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(sheet);
        setUploadedData(json);
        
        // حفظ البيانات المحملة في localStorage
        localStorage.setItem('uploadedFile', JSON.stringify(json));
        
        // Perform VLOOKUP automatically after upload
        performVlookup(json);
      } catch (err) {
        setError("Error processing file: " + err.message);
      } finally {
        setIsLoading(false);
      }
    };
    reader.onerror = () => {
      setError("Failed to read file");
      setIsLoading(false);
    };
    reader.readAsArrayBuffer(file);
  };

  const performVlookup = (uploadedData) => {
    // Create lookup map for better performance
    const uploadedMap = new Map();
    uploadedData.forEach(row => {
      const key = row["Contract No."]?.toString().trim().toLowerCase();
      if (key) uploadedMap.set(key, row);
    });

    const result = refData.map((refRow) => {
      const contractNoRaw = refRow["Contract No."];
      const contractNo = contractNoRaw?.toString().trim().toLowerCase();
      const match = uploadedMap.get(contractNo);
      return {
        contract: contractNoRaw?.toString().trim() || "❌", // نعرض الرقم الأصلي بدون تغيير حالة الأحرف
        plate: match?.["Plate No."] || "❌",
        model: match?.["Model"] || "❌",
        pickup: match?.["Pick-up Date"] || "❌",
        dropoff: match?.["Drop-off Date"] || "❌",
      };
    });
    setResults(result);
  };

  const clearResults = () => {
    setResults([]);
    setError("");
    // مسح البيانات من localStorage أيضًا
    localStorage.removeItem('vlookupResults');
    localStorage.removeItem('uploadedFile');
  };

  const toggleColumnSelection = (columnKey) => {
    setSelectedColumns((prev) =>
      prev.includes(columnKey)
        ? prev.filter((key) => key !== columnKey)
        : [...prev, columnKey]
    );
  };

  const copySelectedColumns = () => {
    if (selectedColumns.length === 0) {
      setCopyMessage("Please select columns first!");
      setTimeout(() => setCopyMessage(""), 2000);
      return;
    }
    
    const columnData = results
      .map((row) =>
        selectedColumns.map((col) => row[col]).join("\t") // Use tab as separator
      )
      .join("\n");
    
    navigator.clipboard.writeText(columnData).then(() => {
      setCopyMessage("Copied Selected Columns!");
      setSelectedColumns([]); // إلغاء تحديد الأعمدة بعد النسخ
      setTimeout(() => setCopyMessage(""), 2000); // Hide message after 2 seconds
    });
  };

  const exportToExcel = () => {
    if (results.length === 0) {
      setCopyMessage("No data to export!");
      setTimeout(() => setCopyMessage(""), 2000);
      return;
    }
    
    const ws = XLSX.utils.json_to_sheet(results);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Results");
    XLSX.writeFile(wb, "YELO_vlookup_results.xlsx");
  };

  const toggleAllColumns = () => {
    if (selectedColumns.length === 5) { // all columns are selected
      setSelectedColumns([]);
    } else {
      setSelectedColumns(['contract', 'plate', 'model', 'pickup', 'dropoff']);
    }
  };

  // محاولة تحميل البيانات المحفوظة عند تحميل المكون للمرة الأولى
  useEffect(() => {
    const savedUploadedFile = localStorage.getItem('uploadedFile');
    if (savedUploadedFile && refData.length > 0) {
      try {
        const parsedData = JSON.parse(savedUploadedFile);
        setUploadedData(parsedData);
        // لا داعي لاستدعاء performVlookup هنا لأن النتائج سيتم تحميلها من localStorage
      } catch (err) {
        console.error("Failed to parse saved data:", err);
      }
    }
  }, [refData]);

  return (
    <div style={{ 
      padding: 20, 
      fontFamily: "'Segoe UI', Arial, sans-serif", 
      maxWidth: "1200px", 
      margin: "0 auto", 
      backgroundColor: yeloColors.tertiary,
      borderRadius: "8px",
      boxShadow: "0 4px 6px rgba(0,0,0,0.05)"
    }}>
      <div style={{ 
        display: "flex", 
        alignItems: "center", 
        justifyContent: "center", 
        marginBottom: "20px" 
      }}>
        <div style={{ 
          background: `linear-gradient(45deg, ${yeloColors.secondary}, ${yeloColors.darkPurple})`,
          padding: "5px",
          borderRadius: "6px",
          display: "inline-block",
          boxShadow: "0 4px 10px rgba(0,0,0,0.15)"
        }}>
          <div style={{
            backgroundColor: yeloColors.primary, 
            padding: "15px 25px",
            borderRadius: "3px",
          }}>
            <h1 style={{ 
              color: yeloColors.secondary, 
              margin: 0,
              fontWeight: "bold",
              letterSpacing: "1px"
            }}>
              YELO Contract Lookup Tool
            </h1>
          </div>
        </div>
      </div>
      
      <div style={{ 
        marginBottom: 30, 
        textAlign: "center", 
        display: "flex",
        flexWrap: "wrap",
        justifyContent: "center",
        alignItems: "center",
        gap: "15px"
      }}>
        <label style={{
          display: "inline-block",
          padding: "12px 20px",
          backgroundColor: yeloColors.primary,
          color: yeloColors.secondary,
          borderRadius: "4px",
          cursor: "pointer",
          fontWeight: "bold",
          boxShadow: "0 2px 5px rgba(0,0,0,0.1)",
          transition: "all 0.3s ease",
        }}
        onMouseOver={(e) => e.currentTarget.style.backgroundColor = "#fff693"}
        onMouseOut={(e) => e.currentTarget.style.backgroundColor = yeloColors.primary}>
          <span>Upload Excel File</span>
          <input
            type="file"
            accept=".xlsx,.xls"
            onChange={handleUpload}
            style={{
              position: "absolute",
              opacity: 0,
              width: "1px",
              height: "1px"
            }}
          />
        </label>
        
        <button
          onClick={clearResults}
          style={{
            ...buttonStyle,
            backgroundColor: yeloColors.error,
          }}
          onMouseOver={(e) => e.currentTarget.style.backgroundColor = "#C62828"}
          onMouseOut={(e) => e.currentTarget.style.backgroundColor = yeloColors.error}
        >
          Clear Results
        </button>
        <button
          onClick={copySelectedColumns}
          style={{
            ...buttonStyle,
            backgroundColor: yeloColors.primary,
            color: yeloColors.secondary,
          }}
          onMouseOver={(e) => e.currentTarget.style.backgroundColor = "#fff693"}
          onMouseOut={(e) => e.currentTarget.style.backgroundColor = yeloColors.primary}
        >
          Copy Selected Columns
        </button>
        <button
          onClick={exportToExcel}
          style={{
            ...buttonStyle,
            backgroundColor: yeloColors.secondary,
          }}
          onMouseOver={(e) => e.currentTarget.style.backgroundColor = yeloColors.darkPurple}
          onMouseOut={(e) => e.currentTarget.style.backgroundColor = yeloColors.secondary}
        >
          Export to Excel
        </button>
      </div>
      
      {isLoading && (
        <div style={{ 
          textAlign: "center", 
          padding: "20px", 
          backgroundColor: yeloColors.white, 
          borderRadius: "5px",
          marginBottom: "20px",
          boxShadow: "0 2px 5px rgba(0,0,0,0.05)"
        }}>
          <div style={{ 
            display: "inline-block", 
            borderRadius: "50%", 
            borderTop: `3px solid ${yeloColors.primary}`, 
            borderRight: "3px solid transparent",
            width: "24px", 
            height: "24px", 
            animation: "spin 1s linear infinite" 
          }}></div>
          <style>{`
            @keyframes spin {
              0% { transform: rotate(0deg); }
              100% { transform: rotate(360deg); }
            }
          `}</style>
          <span style={{ marginLeft: "10px", color: yeloColors.secondary }}>Loading data, please wait...</span>
        </div>
      )}
      
      {error && (
        <div style={{ 
          textAlign: "center", 
          padding: "15px", 
          backgroundColor: "#FFEBEE", 
          color: yeloColors.error, 
          borderRadius: "5px",
          marginBottom: "20px",
          boxShadow: "0 2px 5px rgba(0,0,0,0.05)"
        }}>
          {error}
        </div>
      )}
      
      {copyMessage && (
        <div
          style={{
            position: "fixed",
            top: 20,
            right: 20,
            backgroundColor: yeloColors.success,
            color: yeloColors.white,
            padding: "12px 20px",
            borderRadius: "4px",
            boxShadow: "0 2px 8px rgba(0,0,0,0.2)",
            zIndex: 1000
          }}
        >
          {copyMessage}
        </div>
      )}
      
      {results.length > 0 && (
        <div style={{ 
          marginBottom: 20, 
          padding: "15px", 
          backgroundColor: yeloColors.white,
          borderRadius: "4px",
          boxShadow: "0 2px 5px rgba(0,0,0,0.05)",
        }}>
          <div style={{
            display: "flex",
            justifyContent: "space-between",
            alignItems: "center",
            marginBottom: "10px"
          }}>
            <span style={{ 
              fontWeight: "bold", 
              color: yeloColors.secondary, 
              fontSize: "1.1em" 
            }}>
              Total Records: {results.length}
            </span>
            <span style={{
              color: yeloColors.secondary,
              fontSize: "0.9em"
            }}>
              YELO Car Rental - Contract Verification
            </span>
          </div>
          
          <div style={{
            display: "flex",
            justifyContent: "flex-start",
            gap: "30px",
            marginTop: "10px"
          }}>
            <div style={{
              display: "flex",
              alignItems: "center",
              gap: "5px",
              backgroundColor: yeloColors.success,
              color: yeloColors.white,
              padding: "8px 12px",
              borderRadius: "4px",
              fontSize: "0.9em"
            }}>
              <span style={{ fontWeight: "bold" }}>Complete Records:</span>
              <span>{getCompleteRecords()}</span>
            </div>
            
            <div style={{
              display: "flex",
              alignItems: "center",
              gap: "5px",
              backgroundColor: yeloColors.warning,
              color: yeloColors.white,
              padding: "8px 12px",
              borderRadius: "4px",
              fontSize: "0.9em"
            }}>
              <span style={{ fontWeight: "bold" }}>Missing Data Records:</span>
              <span>{getMissingRecords()}</span>
            </div>
          </div>
        </div>
      )}
      
      {results.length > 0 && (
        <div style={{ 
          overflowX: "auto", 
          backgroundColor: yeloColors.white, 
          borderRadius: "6px", 
          boxShadow: "0 4px 6px rgba(0,0,0,0.05)" 
        }}>
          <table
            style={{
              width: "100%",
              textAlign: "center",
              borderCollapse: "collapse",
              borderRadius: "6px",
              overflow: "hidden"
            }}
          >
            <thead style={{ 
              background: `linear-gradient(to right, ${yeloColors.secondary}, ${yeloColors.darkPurple})`, 
              color: yeloColors.white 
            }}>
              <tr>
                <th style={{ padding: "14px", borderBottom: `2px solid ${yeloColors.primary}`, cursor: "pointer" }} onClick={toggleAllColumns}>
                  <div style={{ display: "flex", alignItems: "center", justifyContent: "center" }}>
                    <input
                      type="checkbox"
                      checked={selectedColumns.length === 5}
                      onChange={toggleAllColumns}
                      style={{ marginRight: 5 }}
                    />
                    <span>#</span>
                  </div>
                </th>
                <th style={{ padding: "14px", borderBottom: `2px solid ${yeloColors.primary}`, cursor: "pointer" }} onClick={() => toggleColumnSelection("contract")}>
                  <div style={{ display: "flex", alignItems: "center", justifyContent: "center" }}>
                    <input
                      type="checkbox"
                      checked={selectedColumns.includes("contract")}
                      onChange={() => toggleColumnSelection("contract")}
                      style={{ marginRight: 5 }}
                    />
                    <span>Contract No</span>
                  </div>
                </th>
                <th style={{ padding: "14px", borderBottom: `2px solid ${yeloColors.primary}`, cursor: "pointer" }} onClick={() => toggleColumnSelection("plate")}>
                  <div style={{ display: "flex", alignItems: "center", justifyContent: "center" }}>
                    <input
                      type="checkbox"
                      checked={selectedColumns.includes("plate")}
                      onChange={() => toggleColumnSelection("plate")}
                      style={{ marginRight: 5 }}
                    />
                    <span>Plate No</span>
                  </div>
                </th>
                <th style={{ padding: "14px", borderBottom: `2px solid ${yeloColors.primary}`, cursor: "pointer" }} onClick={() => toggleColumnSelection("model")}>
                  <div style={{ display: "flex", alignItems: "center", justifyContent: "center" }}>
                    <input
                      type="checkbox"
                      checked={selectedColumns.includes("model")}
                      onChange={() => toggleColumnSelection("model")}
                      style={{ marginRight: 5 }}
                    />
                    <span>Model</span>
                  </div>
                </th>
                <th style={{ padding: "14px", borderBottom: `2px solid ${yeloColors.primary}`, cursor: "pointer" }} onClick={() => toggleColumnSelection("pickup")}>
                  <div style={{ display: "flex", alignItems: "center", justifyContent: "center" }}>
                    <input
                      type="checkbox"
                      checked={selectedColumns.includes("pickup")}
                      onChange={() => toggleColumnSelection("pickup")}
                      style={{ marginRight: 5 }}
                    />
                    <span>Pick-up Date</span>
                  </div>
                </th>
                <th style={{ padding: "14px", borderBottom: `2px solid ${yeloColors.primary}`, cursor: "pointer" }} onClick={() => toggleColumnSelection("dropoff")}>
                  <div style={{ display: "flex", alignItems: "center", justifyContent: "center" }}>
                    <input
                      type="checkbox"
                      checked={selectedColumns.includes("dropoff")}
                      onChange={() => toggleColumnSelection("dropoff")}
                      style={{ marginRight: 5 }}
                    />
                    <span>Drop-off Date</span>
                  </div>
                </th>
              </tr>
            </thead>
            <tbody>
              {results.map((r, i) => {
                // Check if this record has missing data
                const hasMissingData = r.plate === "❌" || r.model === "❌" || 
                                      r.pickup === "❌" || r.dropoff === "❌";
                
                return (
                  <tr 
                    key={i} 
                    style={{ 
                      backgroundColor: hasMissingData ? "#fff3e0" : (i % 2 === 0 ? yeloColors.white : yeloColors.offWhite),
                      transition: "background-color 0.2s"
                    }}
                    onMouseOver={(e) => e.currentTarget.style.backgroundColor = yeloColors.lightYellow}
                    onMouseOut={(e) => e.currentTarget.style.backgroundColor = hasMissingData ? "#fff3e0" : (i % 2 === 0 ? yeloColors.white : yeloColors.offWhite)}
                  >
                    <td style={{ padding: "12px", borderBottom: `1px solid ${yeloColors.offWhite}` }}>{i + 1}</td>
                    <td style={{ padding: "12px", borderBottom: `1px solid ${yeloColors.offWhite}`, fontWeight: "bold", color: yeloColors.secondary }}>{r.contract}</td>
                    <td style={{ padding: "12px", borderBottom: `1px solid ${yeloColors.offWhite}`, color: r.plate === "❌" ? yeloColors.error : "inherit" }}>{r.plate}</td>
                    <td style={{ padding: "12px", borderBottom: `1px solid ${yeloColors.offWhite}`, color: r.model === "❌" ? yeloColors.error : "inherit" }}>{r.model}</td>
                    <td style={{ padding: "12px", borderBottom: `1px solid ${yeloColors.offWhite}`, color: r.pickup === "❌" ? yeloColors.error : "inherit" }}>{r.pickup}</td>
                    <td style={{ padding: "12px", borderBottom: `1px solid ${yeloColors.offWhite}`, color: r.dropoff === "❌" ? yeloColors.error : "inherit" }}>{r.dropoff}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}
      
    </div>
  );
}