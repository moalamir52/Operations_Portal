import React, { useState } from "react";
import * as XLSX from "xlsx";
import { BrowserRouter as Router, useNavigate } from "react-router-dom";

interface Contract {
  contract: string;
  customer: string;
  dropDate: string;
  closedBy: string;
  days: number;
}

function ReminderDue14Days() {
  const [dueContracts, setDueContracts] = useState<Contract[]>([]);
  const navigate = useNavigate();

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();

    reader.onload = (evt) => {
      const bstr = evt.target?.result;
      if (!bstr) return;

      const workbook = XLSX.read(bstr, { type: "binary" });
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];
      const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: "" });

      const processed = jsonData.map((row: any): Contract | null => {
        const dropRaw = row["Drop-off Date"];
        let dropDate: Date | undefined;

        if (typeof dropRaw === 'number') {
          const parsed = XLSX.SSF.parse_date_code(dropRaw);
          if (!parsed) return null;
          dropDate = new Date(parsed.y, parsed.m - 1, parsed.d);
        } else if (typeof dropRaw === 'string') {
          const parts = dropRaw.split(/[\s/:.-]+/);
          const [day, month, year] = parts.map(p => parseInt(p));
          if (!isNaN(day) && !isNaN(month) && !isNaN(year)) {
            dropDate = new Date(year, month - 1, day);
          }
        }

        if (!dropDate || isNaN(dropDate.getTime())) return null;

        const today = new Date();
        const drop = new Date(dropDate.getFullYear(), dropDate.getMonth(), dropDate.getDate());
        const now = new Date(today.getFullYear(), today.getMonth(), today.getDate());

        const diff = Math.floor((now.getTime() - drop.getTime()) / (1000 * 60 * 60 * 24));

        return {
          contract: row["Contract No."],
          customer: row["Customer"],
          dropDate: drop.toLocaleDateString('en-GB'),
          closedBy: row["Closed By"],
          days: diff,
        };
      }).filter((row): row is Contract => row !== null);

      const due = processed.filter((r) => r.days === 13);
      setDueContracts(due);
    };

    reader.readAsBinaryString(file);
  };

  const handleSendEmail = () => {
    const message = `Dear Team,\n\nPlease find below the list of customers whose contracts were closed 13 days ago. Kindly review their accounts and ensure any pending dues are settled promptly.\n\n` +
      dueContracts.map(c => `Contract: ${c.contract} | Customer: ${c.customer} | Drop-off Date: ${c.dropDate}`).join("\n") +
      "\n\nBest regards,\nYELO Team";

    const mailtoLink = `mailto:?subject=Reminder: Contracts Closed 13 Days Ago&body=${encodeURIComponent(message)}`;
    window.location.href = mailtoLink;
  };

  return (
    <div className="min-h-screen bg-gradient-to-br from-yellow-50 via-white to-purple-50 px-6 py-8">
      <div className="max-w-6xl mx-auto bg-white rounded-2xl shadow-lg p-6">
        <div className="flex items-center justify-between mb-6">
          <button onClick={() => navigate("/dashboard")} className="bg-yellow-400 hover:bg-yellow-500 text-black px-4 py-2 rounded-xl shadow">
            ← Back to Dashboard
          </button>
          <h1 className="text-3xl font-extrabold text-purple-900 flex items-center gap-2">
            <span>📢</span> Contracts Closed 13 Days Ago
          </h1>
          <div></div>
        </div>

        <input type="file" accept=".xlsx, .xls" onChange={handleFileUpload} className="mb-6 text-sm border border-gray-300 rounded px-3 py-2" />

        {dueContracts.length > 0 ? (
          <>
            <h2 className="text-xl font-semibold mb-3 text-green-800">✅ Total Contracts: {dueContracts.length}</h2>
            <div className="overflow-x-auto">
              <table className="table-auto border w-full text-sm mb-6 shadow-sm">
                <thead className="bg-purple-100 text-purple-900">
                  <tr>
                    <th className="border px-3 py-2">Contract No.</th>
                    <th className="border px-3 py-2">Customer</th>
                    <th className="border px-3 py-2">Drop-off Date</th>
                    <th className="border px-3 py-2">Closed By</th>
                    <th className="border px-3 py-2">Days</th>
                  </tr>
                </thead>
                <tbody>
                  {dueContracts.map((row, idx) => (
                    <tr key={idx} className="text-center bg-white hover:bg-yellow-50">
                      <td className="border px-3 py-1">{row.contract}</td>
                      <td className="border px-3 py-1">{row.customer}</td>
                      <td className="border px-3 py-1">{row.dropDate}</td>
                      <td className="border px-3 py-1">{row.closedBy}</td>
                      <td className="border px-3 py-1">{row.days}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>

            <button onClick={handleSendEmail} className="bg-purple-700 hover:bg-purple-800 text-white px-6 py-2 rounded-xl shadow">
              📧 Send Email with Contract Details
            </button>
          </>
        ) : (
          <p className="text-green-700 text-sm">✅ No contracts reached day 13 yet.</p>
        )}
      </div>
    </div>
  );
}

export default function WrappedReminder() {
  return (
    <Router>
      <ReminderDue14Days />
    </Router>
  );
}
