import React, { useState } from "react";
import * as XLSX from "xlsx";
import ContractVlookup from "./ContractVlookup.tsx";
import Fleet from "./Fleet.tsx";
import KilometerTracker from './KM.tsx';

function ReminderDue14Days() {
  const [dueContracts, setDueContracts] = useState([]);
  const [emailTarget, setEmailTarget] = useState("dubai"); // ✅ مهم جداً


  const handleFileUpload = (e) => {
    const file = e.target.files[0];
    const reader = new FileReader();

    reader.onload = (evt) => {
      const bstr = evt.target.result;
      const workbook = XLSX.read(bstr, { type: "binary" });
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];
      const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: "" });

      const processed = jsonData.map((row, index) => {
  const pickupRaw = row["Pick-up Date"];
  let pickupDate;

  // Better date parsing for Excel dates
  if (typeof pickupRaw === "number") {
    // Excel date number
    const parsed = XLSX.SSF.parse_date_code(pickupRaw);
    pickupDate = new Date(parsed.y, parsed.m - 1, parsed.d);
  } else if (typeof pickupRaw === "string") {
    // String date - try multiple formats
    const parts = pickupRaw.split(/[\s/:.-]+/);
    if (parts.length >= 3) {
      const [day, month, year] = parts.map((p) => parseInt(p));
      if (!isNaN(day) && !isNaN(month) && !isNaN(year)) {
        // Handle 2-digit years
        const fullYear = year < 100 ? (year < 50 ? 2000 + year : 1900 + year) : year;
        pickupDate = new Date(fullYear, month - 1, day);
      }
    }
  }

  if (!pickupDate || isNaN(pickupDate)) {
    console.warn(`Invalid pickup date for row ${index + 1}:`, pickupRaw);
    return null;
  }

  const today = new Date();
  const pickup = new Date(pickupDate.getFullYear(), pickupDate.getMonth(), pickupDate.getDate());
  const now = new Date(today.getFullYear(), today.getMonth(), today.getDate());

  const diff = Math.floor((now - pickup) / (1000 * 60 * 60 * 24));

  // Ensure all fields are properly extracted and formatted
  const processedRow = {
    contract: row["Contract No."] || "",
    customer: row["Customer"] || "",
    pickupDate: pickup.toLocaleDateString("en-GB"),
    dropDate: pickup.toLocaleDateString("en-GB"),
    days: diff,
    closedBy: row["Closed By"] || "",
    branch: row["Pick-up Branch"] || row["Branch"] || "",
  };

  // Debug log for the first few rows
  if (index < 3) {
    console.log(`Processing row ${index + 1}:`, {
      original: row,
      processed: processedRow
    });
  }

  return processedRow;
}).filter(Boolean);


      const due = processed.filter((r) => r.days === 13);
      setDueContracts(due);
    };

    reader.readAsBinaryString(file);
  };

  const handleSendEmail = () => {
    // Create formatted table for clipboard
    const tableData = dueContracts.map((row, i) => {
      return {
        no: i + 1,
        contract: row.contract || "",
        customer: row.customer || "",
        pickupDate: row.pickupDate || "",
        days: row.days || 0,
        closedBy: row.closedBy || "",
        branch: row.branch || ""
      };
    });

    // Create HTML table for clipboard
    const htmlTable = `
      <table style="border-collapse: collapse; width: 100%; margin: 20px 0; font-family: Arial, sans-serif;">
        <thead>
          <tr style="background-color: #ffd54f; color: #6a1b9a;">
            <th style="padding: 12px; border: 1px solid #ddd; text-align: center; font-weight: bold;">No.</th>
            <th style="padding: 12px; border: 1px solid #ddd; text-align: center; font-weight: bold;">Contract No.</th>
            <th style="padding: 12px; border: 1px solid #ddd; text-align: center; font-weight: bold;">Customer</th>
            <th style="padding: 12px; border: 1px solid #ddd; text-align: center; font-weight: bold;">Pick-up Date</th>
            <th style="padding: 12px; border: 1px solid #ddd; text-align: center; font-weight: bold;">Days</th>
            <th style="padding: 12px; border: 1px solid #ddd; text-align: center; font-weight: bold;">Closed By</th>
            <th style="padding: 12px; border: 1px solid #ddd; text-align: center; font-weight: bold;">Branch</th>
          </tr>
        </thead>
        <tbody>
          ${tableData.map((row, idx) => `
            <tr style="background-color: ${idx % 2 === 0 ? '#fff' : '#f8f9fa'};">
              <td style="padding: 10px; border: 1px solid #ddd; text-align: center; font-weight: bold; color: #6a1b9a;">${row.no}</td>
              <td style="padding: 10px; border: 1px solid #ddd; text-align: center; font-weight: bold; color: #2e7d32;">${row.contract}</td>
              <td style="padding: 10px; border: 1px solid #ddd; text-align: center;">${row.customer}</td>
              <td style="padding: 10px; border: 1px solid #ddd; text-align: center; color: #1976d2;">${row.pickupDate}</td>
              <td style="padding: 10px; border: 1px solid #ddd; text-align: center; font-weight: bold; color: ${row.days === 13 ? '#d32f2f' : '#388e3c'};">${row.days}</td>
              <td style="padding: 10px; border: 1px solid #ddd; text-align: center;">${row.closedBy}</td>
              <td style="padding: 10px; border: 1px solid #ddd; text-align: center; color: #5d4037;">${row.branch}</td>
            </tr>
          `).join('')}
        </tbody>
      </table>
    `;

    // Create plain text version as fallback
    const plainText = tableData.map((row, i) => 
      `${row.no}. Contract: ${row.contract} | Customer: ${row.customer} | Pick-up: ${row.pickupDate} | Days: ${row.days} | Branch: ${row.branch}`
    ).join('\n');

    // Copy to clipboard
    const copyToClipboard = async () => {
      try {
        // Try to copy HTML first
        const clipboardItem = new ClipboardItem({
          'text/html': new Blob([htmlTable], { type: 'text/html' }),
          'text/plain': new Blob([plainText], { type: 'text/plain' })
        });
        
        await navigator.clipboard.write([clipboardItem]);
        
        // Open email client
        const header = `Dear Team,%0D%0A%0D%0AThe following contracts were opened 13 days ago. Kindly review them, ensure all dues are settled, and update the status accordingly..%0D%0A%0D%0A(Note: If you find any cash deposit, please ignore it.)%0D%0A%0D%0A`;
        const footer = `%0D%0A%0D%0ABest regards,%0D%0ABusiness Bay Team`;

        // تحديد الجهة المستلمة بناءً على الاختيار
        let to = "";
        let cc = "a.naseer@iyelo.com";

        if (emailTarget === "dubai") {
          to = "dubaiair@iyelo.com,dubaiair2@iyelo.com";
          cc += ",k.alhamawi@iyelo.com";
        } else if (emailTarget === "oman") {
          to = "m.muscatair@iyelo.com,muscatair@iyelo.com";
        }

        const mailtoLink = `mailto:${to}?cc=${cc}&subject=Reminder: Contracts Pickup 13 Days Ago&body=${header}${footer}`;
        window.location.href = mailtoLink;
        
      } catch (err) {
        // Fallback to plain text copy
        await navigator.clipboard.writeText(plainText);
        
        // Open email client
        const header = `Dear Team,%0D%0A%0D%0AThe following contracts were opened 13 days ago. Kindly review them, ensure all dues are settled, and update the status accordingly..%0D%0A%0D%0A(Note: If you find any cash deposit, please ignore it.)%0D%0A%0D%0A`;
        const footer = `%0D%0A%0D%0ABest regards,%0D%0ABusiness Bay Team`;

        let to = "";
        let cc = "a.naseer@iyelo.com";

        if (emailTarget === "dubai") {
          to = "dubaiair@iyelo.com,dubaiair2@iyelo.com";
          cc += ",k.alhamawi@iyelo.com";
        } else if (emailTarget === "oman") {
          to = "m.muscatair@iyelo.com,muscatair@iyelo.com";
        }

        const mailtoLink = `mailto:${to}?cc=${cc}&subject=Reminder: Contracts Pickup 13 Days Ago&body=${header}${footer}`;
        window.location.href = mailtoLink;
      }
    };

    copyToClipboard();
  };


  const styles = {
    container: {
      fontFamily: "Arial",
      backgroundColor: "#fefce8",
      padding: "40px 20px",
      minHeight: "100vh",
    },
    topBar: {
      backgroundColor: "#FFD700",
      border: "2px solid #6a1b9a",
      color: "#6a1b9a",
      borderRadius: "16px",
      padding: "15px 25px",
      margin: "0 auto 30px",
      maxWidth: "950px",
      display: "flex",
      justifyContent: "center",
      alignItems: "center",
      boxShadow: "0 6px 12px rgba(105, 27, 154, 0.65)",
    },
    backBtn: {
      backgroundColor: "#6a1b9a",
      color: "#fff",
      padding: "8px 14px",
      borderRadius: "8px",
      fontWeight: "bold",
      textDecoration: "none",
      fontSize: "14px",
    },
    title: {
      fontSize: "18px",
      fontWeight: "bold",
    },
    content: {
      maxWidth: "1000px",
      margin: "0 auto",
    },
    input: {
      marginBottom: "20px",
      padding: "8px",
      border: "1px solid #aaa",
      borderRadius: "4px",
      width: "100%",
    },
    table: {
      width: "100%",
      borderCollapse: "collapse",
      marginBottom: "20px",
      borderRadius: "8px",
      overflow: "hidden",
      boxShadow: "0 4px 8px rgba(0,0,0,0.1)",
    },
    th: {
      backgroundColor: "#ffd54f",
      color: "#6a1b9a",
      padding: "15px 12px",
      border: "none",
      textAlign: "center",
      fontWeight: "bold",
      fontSize: "14px",
      textTransform: "uppercase",
      letterSpacing: "0.5px",
    },
    td: {
      padding: "12px 10px",
      border: "none",
      borderBottom: "1px solid #e0e0e0",
      textAlign: "center",
      color: "#333",
      fontSize: "13px",
    },
    emailBtn: {
      backgroundColor: "#6a1b9a",
      color: "#ffd54f",
      padding: "15px 30px",
      fontWeight: "bold",
      border: "none",
      borderRadius: "8px",
      cursor: "pointer",
      fontSize: "16px",
      boxShadow: "0 4px 8px rgba(0,0,0,0.2)",
      transition: "all 0.3s ease",
      display: "block",
      margin: "0 auto",
    },
  };

  return (
    <div style={styles.container}>
      <div style={styles.topBar}>
        <div style={styles.title}>📢 Reminder: Contracts Opened 13 Days Ago</div>
      </div>
      <div style={styles.content}>
        <input type="file" accept=".xlsx, .xls" onChange={handleFileUpload} style={styles.input} />

        {dueContracts.length > 0 ? (
          <>
          <div style={{ marginBottom: "20px" }}>
  <strong>Select Email Target:</strong><br />
  <div style={{ marginBottom: "30px", textAlign: "center" }}>
  <button
    onClick={() => setEmailTarget("dubai")}
    style={{
      margin: "10px",
      padding: "10px 20px",
      backgroundColor: emailTarget === "dubai" ? "#6a1b9a" : "#ffd54f",
      color: emailTarget === "dubai" ? "#ffd54f" : "#6a1b9a",
      border: "none",
      borderRadius: "8px",
      fontWeight: "bold",
      fontSize: "16px",
      cursor: "pointer",
      display: "inline-flex",
      alignItems: "center",
      transition: "all 0.3s ease"
    }}
  >
    <img
      src="https://flagcdn.com/w40/ae.png"
      alt="UAE"
      width="24"
      style={{ marginRight: "8px", borderRadius: "4px" }}
    />
    Dubai
  </button>

  <button
    onClick={() => setEmailTarget("oman")}
    style={{
      margin: "10px",
      padding: "10px 20px",
      backgroundColor: emailTarget === "oman" ? "#6a1b9a" : "#ffd54f",
      color: emailTarget === "oman" ? "#ffd54f" : "#6a1b9a",
      border: "none",
      borderRadius: "8px",
      fontWeight: "bold",
      fontSize: "16px",
      cursor: "pointer",
      display: "inline-flex",
      alignItems: "center",
      transition: "all 0.3s ease"
    }}
  >
    <img
      src="https://flagcdn.com/w40/om.png"
      alt="Oman"
      width="24"
      style={{ marginRight: "8px", borderRadius: "4px" }}
    />
    Oman
  </button>
</div>

</div>

            <table style={styles.table}>
              <thead>
                <tr>
                  <th style={styles.th}>No.</th>
                  <th style={styles.th}>Contract No.</th>
                  <th style={styles.th}>Customer</th>
                  <th style={styles.th}>Pick-up Date</th>
                  <th style={styles.th}>Days</th>
                  <th style={styles.th}>Closed By</th>
                  <th style={styles.th}>Branch</th>
                </tr>
              </thead>
              <tbody>
                {dueContracts.map((row, idx) => (
                  <tr key={idx} style={{
                    backgroundColor: idx % 2 === 0 ? "#fff" : "#f8f9fa",
                    transition: "background-color 0.2s ease"
                  }}
                  onMouseEnter={(e) => e.currentTarget.style.backgroundColor = "#fff3e0"}
                  onMouseLeave={(e) => e.currentTarget.style.backgroundColor = idx % 2 === 0 ? "#fff" : "#f8f9fa"}>
                    <td style={{...styles.td, fontWeight: "bold", color: "#6a1b9a"}}>{idx + 1}</td>
                    <td style={{...styles.td, fontWeight: "bold", color: "#2e7d32"}}>{row.contract}</td>
                    <td style={styles.td}>{row.customer}</td>
                    <td style={{...styles.td, color: "#1976d2"}}>{row.dropDate}</td>
                    <td style={{...styles.td, fontWeight: "bold", color: row.days === 13 ? "#d32f2f" : "#388e3c"}}>{row.days}</td>
                    <td style={styles.td}>{row.closedBy}</td>
                    <td style={{...styles.td, color: "#5d4037"}}>{row.branch}</td>
                  </tr>
                ))}
              </tbody>
            </table>
            
            <button 
              style={styles.emailBtn} 
              onClick={handleSendEmail}
              onMouseEnter={(e) => {
                e.currentTarget.style.backgroundColor = "#4a148c";
                e.currentTarget.style.transform = "translateY(-2px)";
                e.currentTarget.style.boxShadow = "0 6px 12px rgba(0,0,0,0.3)";
              }}
              onMouseLeave={(e) => {
                e.currentTarget.style.backgroundColor = "#6a1b9a";
                e.currentTarget.style.transform = "translateY(0)";
                e.currentTarget.style.boxShadow = "0 4px 8px rgba(0,0,0,0.2)";
              }}
            >
              📧 Send Email
            </button>
          </>
        ) : (
          <p>No contracts reached day 13 yet.</p>
        )}
      </div>
    </div>
  );
}

function App() {
  const [view, setView] = useState("home");

  const containerStyle = {
    fontFamily: "Arial, sans-serif",
    backgroundColor: "#fffde7",
    minHeight: "100vh",
    padding: "40px 20px",
    textAlign: "center",
  };

  const cardStyle = {
    margin: "0 auto",
    padding: "20px",
    width: "80%",
    backgroundColor: "#ffd54f",
    borderRadius: "15px",
    boxShadow: "0 4px 8px rgba(0, 0, 0, 0.2)",
    fontFamily: "Arial, sans-serif",
    color: "#4a148c",
    fontWeight: "bold",
    fontSize: "24px",
    display: "flex",
    justifyContent: "center",
    alignItems: "center",
    position: "relative",
  };

  const backBtnStyle = {
    backgroundColor: "#6a1b9a",
    color: "#fff",
    padding: "10px 20px",
    borderRadius: "8px",
    fontWeight: "bold",
    textDecoration: "none",
    fontSize: "16px",
    cursor: "pointer",
    position: "absolute",
    left: "20px",
  };

  const buttonStyle = {
    padding: "15px 30px",
    margin: "15px",
    fontSize: "16px",
    fontWeight: "bold",
    borderRadius: "10px",
    border: "none",
    cursor: "pointer",
    backgroundColor: "#ffd54f",
    color: "#4a148c",
    boxShadow: "0 4px 6px rgba(0, 0, 0, 0.2)",
    transition: "transform 0.3s ease, background-color 0.3s ease, box-shadow 0.3s ease",
  };

  const buttonHoverStyle = {
    backgroundColor: "#4a148c",
    color: "#ffd54f",
    transform: "scale(1.1)",
    boxShadow: "0 6px 12px rgba(0, 0, 0, 0.3)",
  };

  return (
    <div style={containerStyle}>
      {view === "home" && (
        <>
          <div style={cardStyle}>
            <a href="https://moalamir52.github.io/Yelo/#dashboard" style={backBtnStyle}>← Back to Dashboard</a>
            🎯 Welcome Team! Please Choose a Project
          </div>
          <button
            style={buttonStyle}
            onMouseEnter={(e) => Object.assign(e.target.style, buttonHoverStyle)}
            onMouseLeave={(e) => Object.assign(e.target.style, buttonStyle)}
            onClick={() => setView("reminder")}
          >
            📢 Reminder
          </button>
          <button
            style={buttonStyle}
            onMouseEnter={(e) => Object.assign(e.target.style, buttonHoverStyle)}
            onMouseLeave={(e) => Object.assign(e.target.style, buttonStyle)}
            onClick={() => setView("vlookup")}
          >
            🔍 Contracts
          </button>
          <button
            style={buttonStyle}
            onMouseEnter={(e) => Object.assign(e.target.style, buttonHoverStyle)}
            onMouseLeave={(e) => Object.assign(e.target.style, buttonStyle)}
            onClick={() => setView("fleet")}
          >
            🚗 Fleet
          </button>
          <button
            style={buttonStyle}
            onMouseEnter={(e) => Object.assign(e.target.style, buttonHoverStyle)}
            onMouseLeave={(e) => Object.assign(e.target.style, buttonStyle)}
            onClick={() => setView("kilometer")}
          >
            🧮 Mileage Calculator
          </button>

        </>
      )}

      {view === "reminder" && (
        <>
          <button onClick={() => setView("home")} style={{
            padding: "15px 30px",
            margin: "15px",
            fontSize: "16px",
            fontWeight: "bold",
            borderRadius: "10px",
            border: "none",
            cursor: "pointer",
            backgroundColor: "#ffd54f",
            borderBottom: "4px solid #6a1b9a",
            color: "#4a148c",
          }}>⬅ Back</button>
          <ReminderDue14Days />
        </>
      )}

      {view === "vlookup" && (
        <>
          <button onClick={() => setView("home")} style={{
            padding: "15px 30px",
            margin: "15px",
            fontSize: "16px",
            fontWeight: "bold",
            borderRadius: "10px",
            border: "none",
            cursor: "pointer",
            backgroundColor: "#ffd54f",
            borderBottom: "4px solid #6a1b9a",
            color: "#4a148c",
          }}>⬅ Back</button>
          <ContractVlookup />
        </>
      )}
      {view === "fleet" && (
  <>
    <button onClick={() => setView("home")} style={{
      padding: "15px 30px",
      margin: "15px",
      fontSize: "16px",
      fontWeight: "bold",
      borderRadius: "10px",
      border: "none",
      cursor: "pointer",
      backgroundColor: "#ffd54f",
      borderBottom: "4px solid #6a1b9a",
      color: "#4a148c",
    }}>⬅ Back</button>
    <Fleet />
  </>
)}
      {view === "kilometer" && (
        <>
          <button onClick={() => setView("home")} style={{
            padding: "15px 30px",
            margin: "15px",
            fontSize: "16px",
            fontWeight: "bold",
            borderRadius: "10px",
            border: "none",
            cursor: "pointer",
            backgroundColor: "#ffd54f",
            borderBottom: "4px solid #6a1b9a",
            color: "#4a148c",
          }}>⬅ Back</button>
          <KilometerTracker />
        </>
      )}
<div style={{
  textAlign: "center",
  marginTop: 40,
  padding: 16,
  fontSize: 14,
  color: "#888",
  borderTop: "1px solid #eee"
}}>
  © {new Date().getFullYear()} Mohamed Alamir. All rights reserved.
</div>

    </div>
  );
}

export default App;
