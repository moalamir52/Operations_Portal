import React, { useState, useEffect } from 'react';
import Papa from 'papaparse';
import * as XLSX from 'xlsx';
import html2canvas from 'html2canvas';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';

interface LogEntry {
  date: string;
  out: number;
  inVal: number;
}

interface ContractRow {
  [key: string]: any;
}

// Parse date in format DD/MM/YYYY HH:mm to YYYY-MM-DD
function parseCustomDate(dateStr: string): string | null {
  const match = dateStr.match(/^(\d{2})\/(\d{2})\/(\d{4})/);
  if (match) {
    const [_, day, month, year] = match;
    return `${year}-${month}-${day}`;
  }
  return null;
}

// Format date from YYYY-MM-DD to DD/MM/YYYY
function formatDateToDMY(dateStr: string): string {
  if (!dateStr) return '';
  const [year, month, day] = dateStr.split('-');
  return `${day}/${month}/${year}`;
}

function KilometerTracker() {
  const [logs, setLogs] = useState<LogEntry[]>([]);
  const [out, setOut] = useState('');
  const [inVal, setInVal] = useState('');
  const [date, setDate] = useState('');
  const [lastDate, setLastDate] = useState('');
  const [dateLocked, setDateLocked] = useState(false);
  const [booking, setBooking] = useState('');
  const [contractData, setContractData] = useState<ContractRow | null>(null);
  const [data, setData] = useState<ContractRow[]>([]);
  const [error, setError] = useState('');
  const [inputError, setInputError] = useState('');
  const [toastMsg, setToastMsg] = useState('');
  const showToast = (msg: string) => {
    setToastMsg(msg);
    setTimeout(() => setToastMsg(''), 2000);
  };

  const outInputRef = React.useRef<HTMLInputElement>(null);

  useEffect(() => {
    const csvUrl = 'https://docs.google.com/spreadsheets/d/1XwBko5v8zOdTdv-By8HK_DvZnYT2T12mBw_SIbCfMkE/export?format=csv&gid=769459790';
    Papa.parse(csvUrl, {
      download: true,
      header: true,
      complete: (results) => setData(results.data)
    });
  }, []);

  useEffect(() => {
    if (booking.trim() === '') {
      setContractData(null);
      setError('');
      return;
    }
    const match = data.find(row => row['Booking Number']?.toString().trim() === booking.trim());
    if (match) {
      setContractData(match);
      setError('');
    } else {
      setContractData(null);
      setError('❌ No data found for the entered number');
    }
  }, [booking, data]);

  useEffect(() => {
    if (contractData && contractData['Pick-up Date']) {
      const rawDate = contractData['Pick-up Date'];
      let formattedDate = parseCustomDate(rawDate);
      if (formattedDate) {
        setLastDate(formattedDate);
        setDateLocked(true);
      } else {
        setDateLocked(false);
        setLastDate('');
      }
    } else {
      setDateLocked(false);
      setLastDate('');
    }
  }, [contractData]);

  useEffect(() => {
    // حفظ البيانات في LocalStorage عند كل تغيير
    const dataToSave = {
      logs,
      out,
      inVal,
      date,
      lastDate,
      dateLocked,
      booking,
      contractData
    };
    localStorage.setItem('km-tracker-data', JSON.stringify(dataToSave));
  }, [logs, out, inVal, date, lastDate, dateLocked, booking, contractData]);

  // استرجاع البيانات من LocalStorage عند تحميل الصفحة
  useEffect(() => {
    const saved = localStorage.getItem('km-tracker-data');
    if (saved) {
      try {
        const data = JSON.parse(saved);
        if (data.logs) setLogs(data.logs);
        if (data.out) setOut(data.out);
        if (data.inVal) setInVal(data.inVal);
        if (data.date) setDate(data.date);
        if (data.lastDate) setLastDate(data.lastDate);
        if (typeof data.dateLocked === 'boolean') setDateLocked(data.dateLocked);
        if (data.booking) setBooking(data.booking);
        if (data.contractData) setContractData(data.contractData);
      } catch {}
    }
  }, []);

  // عند تغيير رقم البوكينج، امسح السجلات والحقول
  useEffect(() => {
    setLogs([]);
    setOut('');
    setInVal('');
    setDate('');
    setLastDate('');
    setDateLocked(false);
    localStorage.removeItem('km-tracker-data');
  }, [booking]);

  const handleAddLog = () => {
    const logDate = date || lastDate;
    if (!logDate || !out || !inVal) {
      setInputError('Please enter all fields.');
      return;
    }
    const outNum = Number(out);
    const inNum = Number(inVal);
    if (isNaN(outNum) || isNaN(inNum) || outNum < 0 || inNum < 0) {
      setInputError('OUT and IN must be positive numbers.');
      return;
    }
    if (outNum > inNum) {
      setInputError('OUT cannot be greater than IN.');
      return;
    }
    setLogs([...logs, { date: logDate, out: outNum, inVal: inNum }]);
    setOut(''); setInVal('');
    if (!dateLocked) {
      setLastDate(logDate);
      setDateLocked(true);
    }
    setDate('');
    setInputError('');
    if (outInputRef.current) outInputRef.current.focus();
  };

  const totalUsedKm = logs.reduce((acc, log) => acc + (log.inVal - log.out), 0);

  const getFirstDate = () => {
    if (logs.length === 0) return null;
    const sorted = [...logs].sort((a, b) => new Date(a.date) - new Date(b.date));
    return sorted[0].date;
  };

  const getDaysSinceFirst = () => {
    const firstDate = getFirstDate();
    if (!firstDate) return 0;
    const start = new Date(firstDate);
    const today = new Date();
    return Math.floor((today - start) / (1000 * 60 * 60 * 24));
  };

  const allowedKm = Math.floor((getDaysSinceFirst() / 30) * 2500);
  const exceeded = Math.max(0, totalUsedKm - allowedKm);

  // دالة تصدير السجلات إلى ملف Excel
  async function exportToExcel() {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Records');

    let rowIdx = 1;
    // 1. Contract Start Date block (if exists)
    if (dateLocked && lastDate) {
      const row = sheet.addRow([`Contract Start Date: ${formatDateToDMY(lastDate)}`]);
      sheet.mergeCells(`A${rowIdx}:D${rowIdx}`);
      row.font = { bold: true, color: { argb: 'FFB28704' }, size: 16 };
      row.alignment = { horizontal: 'center', vertical: 'middle' };
      row.height = 28;
      row.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFDE7' } };
      rowIdx++;
      sheet.addRow([]); rowIdx++;
    }

    // 2. بيانات العميل (لو موجودة)
    if (contractData) {
      const block = [
        [`📘 Booking:`, contractData['Booking Number'] || ''],
        [`📄 Contract:`, contractData['Contract No.'] || ''],
        [`👤 Customer:`, contractData['Customer'] || '']
      ];
      block.forEach(([label, value]) => {
        const row = sheet.addRow([label, value]);
        row.font = { bold: true, color: { argb: 'FF6a1b9a' }, size: 13 };
        row.alignment = { vertical: 'middle' };
        row.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEFF0FF' } };
        row.getCell(2).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEFF0FF' } };
        row.height = 20;
        rowIdx++;
      });
      sheet.addRow([]); rowIdx++;
    }

    // 3. عنوان السجلات
    {
      const row = sheet.addRow(['📂 Records']);
      sheet.mergeCells(`A${rowIdx}:D${rowIdx}`);
      row.font = { bold: true, size: 15, color: { argb: 'FF6a1b9a' } };
      row.alignment = { horizontal: 'left', vertical: 'middle' };
      row.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF3E5F5' } };
      row.height = 22;
      rowIdx++;
    }

    // 4. جدول السجلات
    if (logs.length > 0) {
      const headerRow = sheet.addRow(['#', 'OUT', 'IN', 'Distance']);
      headerRow.eachCell(cell => {
        cell.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 13 };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF6a1b9a' } };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
      });
      headerRow.height = 20;
      rowIdx++;
      logs.forEach((log, i) => {
        const row = sheet.addRow([i + 1, log.out, log.inVal, log.inVal - log.out]);
        row.eachCell(cell => {
          cell.alignment = { horizontal: 'center', vertical: 'middle' };
        });
        if (i % 2 === 0) {
          row.eachCell(cell => {
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF3E5F5' } };
          });
        }
        row.height = 18;
        rowIdx++;
      });
      sheet.addRow([]); rowIdx++;
    }

    // 5. Days since contract start
    {
      const row = sheet.addRow([`📅 Days since contract start: ${getDaysSinceFirst()} days`]);
      sheet.mergeCells(`A${rowIdx}:D${rowIdx}`);
      row.font = { bold: true, color: { argb: 'FF4b2991' }, size: 13 };
      row.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF0E6FF' } };
      row.alignment = { horizontal: 'left', vertical: 'middle' };
      row.height = 20;
      rowIdx++;
    }
    // 6. Allowed KM
    {
      const row = sheet.addRow([`✅ Allowed KM: ${allowedKm} km`]);
      sheet.mergeCells(`A${rowIdx}:D${rowIdx}`);
      row.font = { bold: true, color: { argb: 'FF256029' }, size: 13 };
      row.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE6F4EA' } };
      row.alignment = { horizontal: 'left', vertical: 'middle' };
      row.height = 20;
      rowIdx++;
    }
    // 7. Used KM
    {
      const row = sheet.addRow([`📌 Used KM: ${totalUsedKm} km`]);
      sheet.mergeCells(`A${rowIdx}:D${rowIdx}`);
      row.font = { bold: true, color: { argb: 'FF0d47a1' }, size: 13 };
      row.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE3F2FD' } };
      row.alignment = { horizontal: 'left', vertical: 'middle' };
      row.height = 20;
      rowIdx++;
    }
    // 8. Exceeded KM
    {
      const row = sheet.addRow([`⚠️ Exceeded KM: ${exceeded} km`]);
      sheet.mergeCells(`A${rowIdx}:D${rowIdx}`);
      row.font = { bold: true, color: { argb: 'FFb71c1c' }, size: 13 };
      row.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFEBEE' } };
      row.alignment = { horizontal: 'left', vertical: 'middle' };
      row.height = 20;
      rowIdx++;
    }

    // ضبط عرض الأعمدة
    sheet.columns.forEach(col => {
      col.width = 18;
    });

    // اسم الملف
    let fileName = '';
    if (contractData?.['Booking Number']) {
      fileName = `Booking-${contractData['Booking Number']}.xlsx`;
    } else if (lastDate) {
      fileName = `${formatDateToDMY(lastDate)}-records.xlsx`;
    } else {
      const today = new Date();
      const todayStr = today.toISOString().slice(0,10).split('-').reverse().join('-');
      fileName = `${todayStr}-records.xlsx`;
    }

    // حفظ الملف
    const buffer = await workbook.xlsx.writeBuffer();
    saveAs(new Blob([buffer]), fileName);
    // showToast('File exported successfully!'); // تم إلغاء الرسالة
  }

  // دالة تصدير البيانات كصورة
  function exportAsImage() {
    const element = document.getElementById('export-section');
    if (!element) return;
    html2canvas(element).then(canvas => {
      const link = document.createElement('a');
      let fileName = '';
      if (contractData?.['Booking Number']) {
        fileName = `Booking-${contractData['Booking Number']}.png`;
      } else if (lastDate) {
        fileName = `${formatDateToDMY(lastDate)}-records.png`;
      } else {
        const today = new Date();
        const todayStr = today.toISOString().slice(0,10).split('-').reverse().join('-');
        fileName = `${todayStr}-records.png`;
      }
      link.download = fileName;
      link.href = canvas.toDataURL();
      link.click();
      // showToast('Image exported successfully!'); // تم إلغاء الرسالة
    });
  }

  const handleReset = () => {
    setLogs([]);
    setOut('');
    setInVal('');
    setDate('');
    setLastDate('');
    setDateLocked(false);
    setBooking('');
    setContractData(null);
    setError('');
    showToast('Reset completed!');
    localStorage.removeItem('km-tracker-data');
  };

  const isMobile = typeof window !== 'undefined' && window.innerWidth <= 600;

  const containerStyle = {
    fontFamily: 'Arial',
    padding: isMobile ? '8px' : '20px',
    maxWidth: isMobile ? '100%' : '600px',
    margin: 'auto',
    backgroundColor: '#fffbe7',
    borderRadius: '8px',
    width: '100%',
    boxSizing: 'border-box' as const
  };

  const cardStyle = color => ({
    backgroundColor: color,
    color: 'white',
    padding: '15px',
    marginBottom: '10px',
    borderRadius: '6px'
  });

  const inputStyle = {
    margin: isMobile ? '4px 0' : '5px',
    padding: isMobile ? '8px' : '10px',
    width: isMobile ? '100%' : 'calc(100% - 22px)',
    borderRadius: '4px',
    border: '1px solid #ccc',
    fontSize: isMobile ? '15px' : undefined
  };

  const buttonStyle = {
    padding: isMobile ? '10px' : '10px 20px',
    border: 'none',
    borderRadius: '4px',
    cursor: 'pointer',
    marginTop: isMobile ? '8px' : '10px',
    width: isMobile ? '100%' : undefined,
    fontSize: isMobile ? '16px' : undefined
  };

  return (
    <div style={containerStyle}>
      {/* عنوان كبير وجذاب في الأعلى */}
      <div
        style={{
          background: '#ffe066',
          color: '#6a1b9a',
          fontWeight: 'bold',
          fontSize: '32px',
          textAlign: 'center',
          borderRadius: '12px',
          padding: '18px 0',
          margin: '32px 0 28px 0',
          boxShadow: '0 4px 12px rgba(124,58,237,0.10), 0 2px 0 #fff59d',
          letterSpacing: '1px',
          textShadow: '0 1px 0 #fffde7'
        }}
      >
        📊 YELO - Mileage calculation
      </div>

      <input
        type="text"
        placeholder="🔍 Booking Number"
        value={booking}
        onChange={e => setBooking(e.target.value)}
        style={inputStyle}
      />

      {error && <p style={{ color: 'red' }}>{error}</p>}

      {inputError && (
        <div style={{ color: '#e53935', fontWeight: 'bold', margin: '8px 0', fontSize: '15px' }}>{inputError}</div>
      )}

      {/* بيانات العقد تظهر فقط إذا لم توجد سجلات */}
      {contractData && logs.length === 0 && (
        <div style={{ marginBottom: '15px', background: '#eef0ff', padding: '10px', borderRadius: '6px' }}>
          <p><strong>📘 Booking:</strong> {contractData['Booking Number']}</p>
          <p><strong>📄 Contract:</strong> {contractData['Contract No.']}</p>
          <p><strong>👤 Customer:</strong> {contractData['Customer']}</p>
        </div>
      )}

      {/* احذف هذا الجزء: تاريخ البداية العلوي إذا لم توجد سجلات */}
      {/* {dateLocked && lastDate && logs.length === 0 && (
        <div style={{
          background: '#fffde7',
          color: '#b28704',
          fontSize: '24px',
          fontWeight: 'bold',
          borderRadius: '8px',
          padding: '12px 20px',
          margin: '24px 0 18px 0',
          letterSpacing: '1px',
          boxShadow: '0 1px 4px rgba(178,135,4,0.07)'
        }}>
          Contract Start Date: {formatDateToDMY(lastDate)}
        </div>
      )} */}

      {!dateLocked && (
        <>
          <input
            type="date"
            placeholder="📅 Contract Start Date"
            value={date}
            onChange={e => setDate(e.target.value)}
            style={inputStyle}
            onKeyDown={e => { if (e.key === 'Enter') handleAddLog(); }}
          />
          {contractData && (
            <p style={{ color: '#888', fontSize: '13px' }}>
              Contract start date not found, please enter it manually.
            </p>
          )}
        </>
      )}

      <input
        type="number"
        placeholder="🚗 OUT (Start KM)"
        value={out}
        onChange={e => setOut(e.target.value)}
        style={inputStyle}
        onKeyDown={e => { if (e.key === 'Enter') handleAddLog(); }}
        ref={outInputRef}
      />
      <input
        type="number"
        placeholder="🚙 IN (End KM)"
        value={inVal}
        onChange={e => setInVal(e.target.value)}
        style={inputStyle}
        onKeyDown={e => { if (e.key === 'Enter') handleAddLog(); }}
      />
      {/* ضع id="export-section" على القسم الذي تريد تصديره كصورة */}
      {/* الأزرار خارج export-section */}
      <div style={{
        display: isMobile ? 'block' : 'flex',
        gap: isMobile ? '0' : '12px',
        margin: isMobile ? '10px 0' : '18px 0',
        justifyContent: 'center'
      }}>
        <button
          style={{
            ...buttonStyle,
            background: '#4CAF50',
            color: '#fff',
          }}
          onClick={handleAddLog}
        >
          Add Entry
        </button>
        <button
          style={{
            ...buttonStyle,
            background: '#e53935',
            color: '#fff',
          }}
          onClick={handleReset}
        >
          Reset
        </button>
        <button
          style={{
            ...buttonStyle,
            background: '#7c3aed',
            color: '#fff',
          }}
          onClick={exportToExcel}
        >
          Export to Excel
        </button>
        <button
          style={{
            ...buttonStyle,
            background: '#ffb300',
            color: '#fff',
          }}
          onClick={exportAsImage}
        >
          Export as Image
        </button>
      </div>

      {/* النتائج فقط داخل export-section */}
      <div id="export-section">
        {/* تاريخ البداية العلوي بشكل واضح ومفصول */}
        {dateLocked && lastDate && (
          <div style={{
            background: '#fffde7',
            color: '#b28704',
            fontSize: '24px',
            fontWeight: 'bold',
            borderRadius: '8px',
            padding: '12px 20px',
            margin: '24px 0 18px 0',
            letterSpacing: '1px',
            boxShadow: '0 1px 4px rgba(178,135,4,0.07)'
          }}>
            Contract Start Date: {formatDateToDMY(lastDate)}
          </div>
        )}

        {logs.length > 0 && (
          <>
            {contractData && (
              <div style={{ marginBottom: '15px', background: '#eef0ff', padding: '10px', borderRadius: '6px' }}>
                <p><strong>📘 Booking:</strong> {contractData['Booking Number']}</p>
                <p><strong>📄 Contract:</strong> {contractData['Contract No.']}</p>
                <p><strong>👤 Customer:</strong> {contractData['Customer']}</p>
              </div>
            )}
            {/* احذف عرض تاريخ البداية هنا */}
            {/* {dateLocked && lastDate && (
              <div style={{
                background: '#fffde7',
                color: '#b28704',
                fontSize: '24px',
                fontWeight: 'bold',
                borderRadius: '8px',
                padding: '12px 20px',
                margin: '24px 0 18px 0',
                letterSpacing: '1px',
                boxShadow: '0 1px 4px rgba(178,135,4,0.07)'
              }}>
                Contract Start Date: {formatDateToDMY(lastDate)}
              </div>
            )} */}
            <h3 style={{ marginTop: '20px' }}>📂 Records</h3>
            {logs.map((log, i) => (
              <div
                key={i}
                style={{
                  display: 'flex',
                  alignItems: 'center',
                  justifyContent: 'space-between',
                  backgroundColor: '#fff',
                  padding: '12px 16px',
                  marginBottom: '10px',
                  borderRadius: '8px',
                  border: '1px solid #ddd',
                  boxShadow: '0 1px 2px rgba(0,0,0,0.03)'
                }}
              >
                <span style={{ fontWeight: 'bold', color: '#1565c0', minWidth: 30 }}>{i + 1}.</span>
                <span style={{ margin: '0 10px', color: '#333' }}>🚗 OUT: <strong>{log.out}</strong></span>
                ➡️<span style={{ margin: '0 10px', color: '#333' }}>🚙 IN: <strong>{log.inVal}</strong></span>
                <span style={{
                  background: '#e3f2fd',
                  color: '#0d47a1',
                  fontWeight: 'bold',
                  fontSize: '22px',
                  borderRadius: '6px',
                  padding: '4px 16px',
                  marginLeft: '10px',
                  display: 'flex',
                  alignItems: 'center'
                }}>
                  📍 {log.inVal - log.out} km
                </span>
              </div>
            ))}
            <div style={{
              background: '#f0e6ff',
              color: '#4b2991',
              fontWeight: 'bold',
              fontSize: '18px',
              borderRadius: '8px',
              padding: '12px 0',
              marginBottom: '10px',
              boxShadow: '0 1px 4px rgba(75,41,145,0.07)'
            }}>
              <span style={{marginRight: 8}}>📅</span>
              Days since contract start: {getDaysSinceFirst()} days
            </div>
            <div style={{
              background: '#e6f4ea',
              color: '#256029',
              fontWeight: 'bold',
              fontSize: '18px',
              borderRadius: '8px',
              padding: '12px 0',
              marginBottom: '10px',
              boxShadow: '0 1px 4px rgba(37,96,41,0.07)'
            }}>
              <span style={{marginRight: 8}}>✅</span>
              Allowed KM: {allowedKm} km
            </div>
            <div style={{
              background: '#e3f2fd',
              color: '#0d47a1',
              fontWeight: 'bold',
              fontSize: '18px',
              borderRadius: '8px',
              padding: '12px 0',
              marginBottom: '10px',
              boxShadow: '0 1px 4px rgba(13,71,161,0.07)'
            }}>
              <span style={{marginRight: 8}}>📌</span>
              Used KM: {totalUsedKm} km
            </div>
            <div style={{
              background: '#ffebee',
              color: '#b71c1c',
              fontWeight: 'bold',
              fontSize: '18px',
              borderRadius: '8px',
              padding: '12px 0',
              marginBottom: '10px',
              boxShadow: '0 1px 4px rgba(183,28,28,0.07)'
            }}>
              <span style={{marginRight: 8}}>⚠️</span>
              Exceeded KM: {exceeded} km
            </div>
          </>
        )}
      </div>
      {/* Toast للإشعارات */}
      {toastMsg && (
        <div style={{
          position: 'fixed',
          top: 24,
          left: '50%',
          transform: 'translateX(-50%)',
          background: '#323232',
          color: '#fff',
          padding: '14px 32px',
          borderRadius: '8px',
          fontWeight: 'bold',
          fontSize: '17px',
          zIndex: 9999,
          boxShadow: '0 2px 12px rgba(0,0,0,0.15)'
        }}>
          {toastMsg}
        </div>
      )}
    </div>
  );
}

export default KilometerTracker;
