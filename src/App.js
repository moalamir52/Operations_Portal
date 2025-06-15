import React, { useState } from "react";
import * as XLSX from "xlsx";
import ContractVlookup from "./ContractVlookup.tsx";
import Fleet from "./Fleet.tsx";

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

      const processed = jsonData.map((row) => {
  const pickupRaw = row["Pick-up Date"];
  let pickupDate;

  if (typeof pickupRaw === "number") {
    const parsed = XLSX.SSF.parse_date_code(pickupRaw);
    pickupDate = new Date(parsed.y, parsed.m - 1, parsed.d);
  } else if (typeof pickupRaw === "string") {
    const parts = pickupRaw.split(/[\s/:.-]+/);
    if (parts.length >= 3) {
      const [day, month, year] = parts.map((p) => parseInt(p));
      if (!isNaN(day) && !isNaN(month) && !isNaN(year)) {
        pickupDate = new Date(year, month - 1, day);
      }
    }
  }

  if (!pickupDate || isNaN(pickupDate)) return null;

  const today = new Date();
  const pickup = new Date(pickupDate.getFullYear(), pickupDate.getMonth(), pickupDate.getDate());
  const now = new Date(today.getFullYear(), today.getMonth(), today.getDate());

  const diff = Math.floor((now - pickup) / (1000 * 60 * 60 * 24));

  return {
    contract: row["Contract No."],
    customer: row["Customer"],
    pickupDate: pickup.toLocaleDateString("en-GB"),
dropDate: pickup.toLocaleDateString("en-GB"), // العرض ما زال بنفس الاسم
    days: diff,
    closedBy: row["Closed By"],
    branch: row["Pick-up Branch"] || "",
  };
}).filter(Boolean);


      const due = processed.filter((r) => r.days === 13);
      setDueContracts(due);
    };

    reader.readAsBinaryString(file);
  };

  const handleSendEmail = () => {
  const header = `Dear Team,%0D%0A%0D%0AThe following contracts were Opened 13 days ago. Please review and ensure dues are settled.%0D%0A%0D%0A(Note: If you find any cash deposit, please ignore it.)%0D%0A%0D%0A`;
  const tableHeader = `No.  Contract No.           Pick-up Date   Days  Branch%0D%0A`;

  const tableBody = dueContracts.map((row, i) => {
    const num = (i + 1).toString().padEnd(4, " ");
    const contract = (row.contract || "").padEnd(22, " ");
    const pickup = (row.pickupDate || "").padEnd(16, " ");
    const days = row.days.toString().padEnd(6, " ");
    const branch = (row.branch || "").padEnd(15, " ");
    return `${num}${contract}${pickup}${days}${branch}`;
  }).join("%0D%0A");

  const footer = `%0D%0A%0D%0ABest regards,%0D%0ABusiness Bay Team`;

  // تحديد الجهة المستلمة بناءً على الاختيار
  let to = "";
  const cc = "a.naseer@iyelo.com";

  if (emailTarget === "dubai") {
    to = "dubaiair@iyelo.com,dubaiair2@iyelo.com";
  } else if (emailTarget === "oman") {
    to = "m.muscatair@iyelo.com,muscatair@iyelo.com";
  }

  const mailtoLink = `mailto:${to}?cc=${cc}&subject=Reminder: Contracts Closed 13 Days Ago&body=${header}${tableHeader}${tableBody}${footer}`;
  window.location.href = mailtoLink;
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
      color: "#333",
      borderRadius: "16px",
      padding: "15px 25px",
      margin: "0 auto 30px",
      maxWidth: "950px",
      display: "flex",
      justifyContent: "space-between",
      alignItems: "center",
      boxShadow: "0 4px 12px rgba(106, 27, 154, 0.2)",
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
    },
    th: {
      backgroundColor: "#ede7f6",
      color: "#6a1b9a",
      padding: "10px",
      border: "1px solid #ccc",
      textAlign: "center",
    },
    td: {
      padding: "10px",
      border: "1px solid #eee",
      textAlign: "center",
      color: "#333",
    },
    emailBtn: {
      backgroundColor: "#6a1b9a",
      color: "#fff",
      padding: "12px 20px",
      fontWeight: "bold",
      border: "none",
      borderRadius: "6px",
      cursor: "pointer",
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
                  <tr key={idx}>
                    <td style={styles.td}>{idx + 1}</td>
                    <td style={styles.td}>{row.contract}</td>
                    <td style={styles.td}>{row.customer}</td>
                    <td style={styles.td}>{row.dropDate}</td>
                    <td style={styles.td}>{row.days}</td>
                    <td style={styles.td}>{row.closedBy}</td>
                    <td style={styles.td}>{row.branch}</td>
                  </tr>
                ))}
              </tbody>
            </table>
            <button style={styles.emailBtn} onClick={handleSendEmail}>📧 Send Email</button>
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

    </div>
  );
}

export default App;
