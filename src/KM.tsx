import React, { useState, useEffect } from 'react';
import Papa from 'papaparse';
import * as XLSX from 'xlsx';
import html2canvas from 'html2canvas';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import jsPDF from 'jspdf';

interface LogEntry {
  date: string;
  out: number;
  inVal: number;
}

interface ContractRow {
  [key: string]: any;
}

// Parse date in multiple formats to YYYY-MM-DD
function parseCustomDate(dateStr: string): string | null {
  if (!dateStr) return null;
  
  // Format: DD/MM/YYYY HH:mm
  let match = dateStr.match(/^(\d{2})\/(\d{2})\/(\d{4})/);
  if (match) {
    const [_, day, month, year] = match;
    return `${year}-${month}-${day}`;
  }
  
  // Format: DD-MM-YYYY (with optional time)
  match = dateStr.match(/^(\d{2})-(\d{2})-(\d{4})/);
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
  const [closedData, setClosedData] = useState<ContractRow[]>([]);
  const [contractSource, setContractSource] = useState<string>('');
  const [error, setError] = useState('');
  const [inputError, setInputError] = useState('');
  const [toastMsg, setToastMsg] = useState('');
  const showToast = (msg: string) => {
    setToastMsg(msg);
    setTimeout(() => setToastMsg(''), 2000);
  };
  const [manualEndDate, setManualEndDate] = useState<string>('');
  const [endDateInputVisible, setEndDateInputVisible] = useState(true);
  const [showRefModal, setShowRefModal] = useState(false);
  const [refInput, setRefInput] = useState('');
  const [exportType, setExportType] = useState('both');

  const outInputRef = React.useRef<HTMLInputElement>(null);

  useEffect(() => {
    // Load Open Contracts
    const openCsvUrl = 'https://docs.google.com/spreadsheets/d/1XwBko5v8zOdTdv-By8HK_DvZnYT2T12mBw_SIbCfMkE/export?format=csv&gid=769459790';
    Papa.parse(openCsvUrl, {
      download: true,
      header: true,
      complete: (results) => setData(results.data)
    });
    
    // Load Closed Contracts
    const closedCsvUrl = 'https://docs.google.com/spreadsheets/d/1XwBko5v8zOdTdv-By8HK_DvZnYT2T12mBw_SIbCfMkE/export?format=csv&gid=1830448171';
    Papa.parse(closedCsvUrl, {
      download: true,
      header: true,
      complete: (results) => setClosedData(results.data)
    });
  }, []);

  useEffect(() => {
    if (booking.trim() === '') {
      setContractData(null);
      setError('');
      setContractSource('');
      return;
    }
    
    // Search in Open Contracts first
    let match = data.find(row => row['Booking Number']?.toString().trim() === booking.trim());
    if (match) {
      setContractData(match);
      setContractSource('Open Contract');
      setError('');
      return;
    }
    
    // If not found, search in Closed Contracts
    match = closedData.find(row => row['Booking Number']?.toString().trim() === booking.trim());
    if (match) {
      setContractData(match);
      setContractSource('Closed Contract');
      setError('');
      return;
    }
    
    // Not found in either
    setContractData(null);
    setContractSource('');
    setError('❌ No data found for the entered number');
  }, [booking, data, closedData]);

  useEffect(() => {
    if (contractData) {
      // قراءة تاريخ بداية العقد من Pick-up Date
      if (contractData['Pick-up Date']) {
        const rawDate = contractData['Pick-up Date'];
        let formattedDate = parseCustomDate(rawDate);
        if (formattedDate) {
          setLastDate(formattedDate);
          setDateLocked(true);
        }
      } else {
        // إذا لم يوجد تاريخ بداية، امسح القيمة
        setLastDate('');
        setDateLocked(false);
      }
      
      // للعقود المفتوحة: إذا لم يكن هناك تاريخ نهاية، ضع تاريخ اليوم
      if (contractSource === 'Open Contract') {
        if (!contractData['Close Date'] || contractData['Close Date'].trim() === '') {
          const today = new Date().toISOString().split('T')[0];
          setManualEndDate(today);
        } else {
          const closeDate = parseCustomDate(contractData['Close Date']);
          if (closeDate) {
            setManualEndDate(closeDate);
          }
        }
      }
      
      // للعقود المغلقة: استخدم تاريخ الإرجاع
      if (contractSource === 'Closed Contract' && contractData['Drop-Off Dte']) {
        const dropOffDate = parseCustomDate(contractData['Drop-Off Dte']);
        if (dropOffDate) {
          setManualEndDate(dropOffDate);
        }
      }
    } else {
      // إذا لم توجد بيانات عقد، امسح كل شيء
      setDateLocked(false);
      setLastDate('');
      setManualEndDate('');
    }
  }, [contractData, contractSource]);

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
        if (data.manualEndDate) setManualEndDate(data.manualEndDate);
        if (typeof data.endDateInputVisible === 'boolean') setEndDateInputVisible(data.endDateInputVisible);
      } catch {}
    }
  }, []);

  // حفظ البيانات في LocalStorage عند كل تغيير
  useEffect(() => {
    const dataToSave = {
      logs,
      out,
      inVal,
      date,
      lastDate,
      dateLocked,
      booking,
      contractData,
      manualEndDate,
      endDateInputVisible
    };
    localStorage.setItem('km-tracker-data', JSON.stringify(dataToSave));
  }, [logs, out, inVal, date, lastDate, dateLocked, booking, contractData, manualEndDate, endDateInputVisible]);

  // عند تغيير رقم البوكينج، امسح السجلات والحقول
  useEffect(() => {
    if (booking.trim() !== '') {
      setLogs([]);
      setOut('');
      setInVal('');
      setDate('');
      // لا تمسح lastDate و dateLocked هنا - دعها للـ useEffect الخاص بـ contractData
      setEndDateInputVisible(true);
      localStorage.removeItem('km-tracker-data');
    }
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
    if (logs.length === 0) setEndDateInputVisible(false); // Hide after first entry
  };

  const totalUsedKm = logs.reduce((acc, log) => acc + (log.inVal - log.out), 0);

  const getFirstDate = () => {
    if (logs.length === 0) return null;
    const sorted = [...logs].sort((a, b) => new Date(a.date) - new Date(b.date));
    return sorted[0].date;
  };

  // دالة لجلب تاريخ نهاية العقد
  const getContractEndDate = () => {
    if (manualEndDate) {
      return new Date(manualEndDate);
    }
    if (contractData) {
      // For closed contracts, use 'Drop-Off Dte'
      if (contractSource === 'Closed Contract' && contractData['Drop-Off Dte']) {
        const dropOffDate = parseCustomDate(contractData['Drop-Off Dte']);
        if (dropOffDate) return new Date(dropOffDate);
      }
      // For open contracts, use 'Close Date'
      if (contractData['Close Date']) {
        const closeDate = parseCustomDate(contractData['Close Date']);
        if (closeDate) return new Date(closeDate);
      }
    }
    return null;
  };

  const getDaysSinceFirst = () => {
    const firstDate = getFirstDate();
    if (!firstDate) return 0;
    const start = new Date(firstDate);
    const end = getContractEndDate();
    if (!end) return 0;
    return Math.floor((end.getTime() - start.getTime()) / (1000 * 60 * 60 * 24));
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

  // دالة إنشاء الفاتورة الضريبية كملف Word
  function generateTaxInvoice() {
    if (!contractData || exceeded <= 0) {
      showToast('No excess kilometers to invoice!');
      return;
    }

    // حساب التكلفة
    const pricePerKm = 1.0;
    const baseAmount = exceeded * pricePerKm;
    const vatAmount = baseAmount * 0.05;
    const totalAmount = baseAmount + vatAmount;

    // عرض النافذة المنبثقة
    setShowRefModal(true);
  }

  const handleRefSubmit = () => {
    // حساب التكلفة
    const pricePerKm = 1.0;
    const baseAmount = exceeded * pricePerKm;
    const vatAmount = baseAmount * 0.05;
    const totalAmount = baseAmount + vatAmount;

    // بيانات الفاتورة
    const today = new Date();
    const invoiceDate = today.toLocaleDateString('en-GB');
    const refNumber = refInput ? `ALWFQ-${refInput}` : 'ALWFQ-';
    const customerName = contractData['Customer'] || 'Customer';
    const bookingId = contractData['Booking Number'] || '';
    const contractNo = (contractData['Contract No.'] || '').split('-')[0] || '';
    const model = contractData['Car Model'] || contractData['Model'] || '';
    const plateNo = contractData['Plate Number'] || contractData['Plate No.'] || '';
    const vehicle = model && plateNo ? `${model} - ${plateNo}` : (model || plateNo || 'Vehicle');
    const startDate = lastDate ? formatDateToDMY(lastDate) : '';
    const endDate = manualEndDate ? formatDateToDMY(manualEndDate) : '';
    // إنشاء قائمة بكل الرحلات
    const tripDetails = logs.map((log, index) => 
      `Delivered Kilometer: ${log.out} KM\nCollected Kilometer: ${log.inVal} KM`
    ).join('\n\n');

    // إنشاء محتوى HTML للفاتورة
    const htmlContent = `
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Tax Invoice</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 2px; font-size: 14px; line-height: 1.2; }
        .header { text-align: center; font-size: 22px; font-weight: bold; margin-bottom: 10px; }
        .date-ref { display: flex; justify-content: space-between; align-items: center; margin-bottom: 8px; font-size: 14px; }
        .company-info { margin-bottom: 10px; font-size: 14px; line-height: 1.1; }
        .subject { font-weight: bold; text-decoration: underline; margin: 8px 0; font-size: 14px; }
        table { width: 100%; border-collapse: collapse; margin: 8px 0; font-size: 11px; }
        th, td { border: 1px solid black; padding: 4px; text-align: center; word-wrap: break-word; vertical-align: top; }
        th { background-color: #f0f0f0; font-weight: bold; height: 25px; }
        .description { text-align: left; font-size: 11px; width: 50%; padding: 4px; line-height: 1.1; }
        .main-table th:nth-child(1) { width: 6%; }
        .main-table th:nth-child(2) { width: 52%; }
        .main-table th:nth-child(3) { width: 14%; }
        .main-table th:nth-child(4) { width: 14%; }
        .main-table th:nth-child(5) { width: 14%; }
        .total-row { font-weight: bold; }
        .conditions { margin-top: 10px; font-size: 14px; }
        .signature { margin-top: 15px; font-size: 14px; }
    </style>
</head>
<body>
    <div class="header">Tax Invoice</div>
    
    <table style="width: 100%; margin-bottom: 1px; border: none;">
        <tr>
            <td style="text-align: left; border: none; padding: 2px;">${invoiceDate}</td>
            <td style="text-align: right; border: none; padding: 2px;">Ref: ${refNumber}</td>
        </tr>
    </table>
    
    <div style="text-align: right; margin-bottom: 8px; font-size: 13px;">TRN#: 100397403500003</div>
    
    <div class="company-info">
        <div>Invygo Tech FZ-LLC</div>
        <div>Dubai Internet City</div>
        <div>Dubai, U.A.E.</div>
    </div>
    
    <div class="subject">SUB: Micro Lease Cars</div>
    
    <p style="margin: 5px 0;">Dear Sir,</p>
    
    <p style="margin: 5px 0;">We thank you for your business renting the below vehicle;</p>
    
    <table class="main-table">
        <tr>
            <th>No.</th>
            <th>Description</th>
            <th>Exceed KM Amount</th>
            <th>VAT 5%</th>
            <th>Total Price</th>
        </tr>
        <tr>
            <td>1</td>
            <td class="description">
                ${customerName}<br>
                Booking ID: ${bookingId}<br>
                R/A: ${contractNo}<br>
                Vehicle: ${vehicle}<br>
                Start Date: ${startDate} - ${endDate}<br>
                ${tripDetails}<br>
                Total Used Kilometer = ${totalUsedKm} KM<br>
                Total Allowed Kilometer = ${allowedKm} KM<br>
                Exceeded KM: ${totalUsedKm - allowedKm} KM
            </td>
            <td>${exceeded.toLocaleString()}</td>
            <td>${vatAmount.toFixed(2)}</td>
            <td>${totalAmount.toFixed(2)}</td>
        </tr>
        <tr class="total-row">
    <td colspan="3" style="border: none;"></td>
    <td style="text-align: right; background-color: #bfbfbf;">TOTAL:</td>
    <td style="background-color: #bfbfbf;">AED ${totalAmount.toFixed(2)}</td>
</tr>
    </table>
    
    <div class="conditions">
        <div style="font-weight: bold; text-decoration: underline;">General Conditions:</div>
        <br>
        <div>Terms of Payment: within 7 days</div>
    </div>
    
    <div class="signature">
        <p>Thanking you and assuring you of our best co-operation and services at all times.</p>
        <br>
        <p>Best Regards,</p>
        <br><br>
        <p><strong>Saudian Alwefaq Rent A Car</strong></p>
    </div>
</body>
</html>`;

    // تصدير حسب الاختيار
    if (exportType === 'word' || exportType === 'both') {
      const blob = new Blob([htmlContent], { type: 'application/msword' });
      const fileName = `Tax-Invoice-${contractData['Booking Number'] || 'Unknown'}.doc`;
      saveAs(blob, fileName);
    }
    
    if (exportType === 'pdf' || exportType === 'both') {
      generatePDFWithBackground(htmlContent, customerName, bookingId);
    }
    
    showToast('Tax Invoice generated successfully!');
    
    // إغلاق النافذة ومسح المدخل
    setShowRefModal(false);
    setRefInput('');
  };

  // Helper function to fetch image and convert to data URL
  const toDataURL = (url: string): Promise<string> => fetch(url)
    .then(response => response.blob())
    .then(blob => new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onloadend = () => resolve(reader.result as string);
      reader.onerror = reject;
      reader.readAsDataURL(blob);
    }));

  // دالة إنشاء PDF مع خلفية Letterhead
  const generatePDFWithBackground = async (htmlContent: string, customerName: string, bookingId: string) => {
    try {
      // 1. Fetch the background image and convert to Base64
      const letterheadDataUrl = await toDataURL(process.env.PUBLIC_URL + '/Letterhead.jpg');
      
      // 2. Create the HTML structure with the embedded image
      const htmlWithBg = `
      <div style="
        width: 210mm;
        height: 297mm;
        position: relative;
        margin: 0;
        padding: 0;
        background-color: white; /* Ensure a white background */
      ">
        <img src="${letterheadDataUrl}" style="
          position: absolute;
          top: 0;
          left: 0;
          width: 210mm;
          height: 297mm;
          z-index: 1;
        " />
        <div style="
          position: absolute;
          top: 0;
          left: 0;
          right: 0;
          bottom: 0;
          z-index: 2;
          padding: 120px 30px 20px 30px;
          box-sizing: border-box;
        ">
          ${htmlContent.replace(/<\/?html>|<\/?body>/g, '')}
        </div>
      </div>`;
      
      // 3. Create a temporary element to render the HTML
      const tempDiv = document.createElement('div');
      tempDiv.innerHTML = htmlWithBg;
      tempDiv.style.position = 'absolute';
      tempDiv.style.left = '-9999px';
      tempDiv.style.top = '0';
      tempDiv.style.width = '210mm';
      tempDiv.style.height = '297mm';
      
      document.body.appendChild(tempDiv);
      
      // 4. Use html2canvas to capture the element
      // Use a small timeout to ensure the image is rendered
      setTimeout(() => {
        html2canvas(tempDiv, {
          scale: 2, // Higher scale for better quality
          useCORS: true,
          allowTaint: true,
          backgroundColor: null, // Use transparent background for canvas
        }).then(canvas => {
          // 5. Generate PDF
          const pdf = new jsPDF('p', 'mm', 'a4');
          const imgData = canvas.toDataURL('image/png');
          const imgWidth = 210;
          const imgHeight = (canvas.height * imgWidth) / canvas.width;
          
          pdf.addImage(imgData, 'PNG', 0, 0, imgWidth, imgHeight);
          pdf.save(`Tax-Invoice-${bookingId || 'Unknown'}.pdf`);
          
          // 6. Cleanup
          document.body.removeChild(tempDiv);
        }).catch(error => {
          console.error('PDF generation error:', error);
          document.body.removeChild(tempDiv);
        });
      }, 500); // A short delay can help ensure images are loaded

    } catch (error) {
      console.error('Failed to load background image for PDF:', error);
    }
  };
  

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
    setEndDateInputVisible(true); // أضفت هذا السطر ليظهر حقل نهاية العقد بعد الريسيت
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
    padding: isMobile ? '12px' : '14px',
    width: isMobile ? '100%' : 'calc(100% - 28px)',
    borderRadius: '18px',
    border: '2px solid #ffe066', // أصفر يلو
    fontSize: isMobile ? '15px' : '17px',
    background: '#fffbe7', // أصفر فاتح جداً
    boxShadow: '0 2px 12px rgba(106,27,154,0.07)', // بنفسجي خفيف
    outline: 'none',
    transition: 'box-shadow 0.2s, border-color 0.2s, background 0.2s',
    color: '#6a1b9a', // بنفسجي يلو
    fontWeight: 500,
  };

  // تأثير عند التركيز (focus) عبر style inline
  const handleInputFocus = (e: React.FocusEvent<HTMLInputElement>) => {
    e.target.style.boxShadow = '0 4px 16px rgba(106,27,154,0.18)';
    e.target.style.borderColor = '#6a1b9a'; // بنفسجي يلو
    e.target.style.background = '#fff';
  };
  const handleInputBlur = (e: React.FocusEvent<HTMLInputElement>) => {
    e.target.style.boxShadow = '0 2px 12px rgba(106,27,154,0.07)';
    e.target.style.borderColor = '#ffe066'; // أصفر يلو
    e.target.style.background = '#fffbe7';
  };

  // Helper to convert pasted date like '13/07/2025 16:54' to '2025-07-13'
  const handleDatePaste = (e: React.ClipboardEvent<HTMLInputElement>, setter: (val: string) => void) => {
    const pasted = e.clipboardData.getData('text');
    // Match DD/MM/YYYY or DD/MM/YYYY HH:mm
    const match = pasted.match(/^(\d{2})\/(\d{2})\/(\d{4})/);
    if (match) {
      const [_, day, month, year] = match;
      const formatted = `${year}-${month}-${day}`;
      e.preventDefault();
      setter(formatted);
    }
    // else allow default
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
          fontSize: '34px',
          textAlign: 'center',
          borderRadius: '28px',
          padding: '22px 0',
          margin: '32px 0 28px 0',
          boxShadow: '0 8px 32px 0 rgba(106,27,154,0.22), 0 2px 0 #ffe066',
          letterSpacing: '1.5px',
          textShadow: '0 2px 8px #fffde7, 0 1px 0 #fff',
          transition: 'transform 0.18s, box-shadow 0.18s',
          cursor: 'pointer',
        }}
        onMouseOver={e => {
          e.currentTarget.style.transform = 'scale(1.025)';
          e.currentTarget.style.boxShadow = '0 16px 48px 0 rgba(106,27,154,0.28), 0 2px 0 #ffe066';
        }}
        onMouseOut={e => {
          e.currentTarget.style.transform = 'scale(1)';
          e.currentTarget.style.boxShadow = '0 8px 32px 0 rgba(106,27,154,0.22), 0 2px 0 #ffe066';
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
        onFocus={handleInputFocus}
        onBlur={handleInputBlur}
      />

      {error && <p style={{ color: 'red' }}>{error}</p>}

      {inputError && (
        <div style={{ color: '#e53935', fontWeight: 'bold', margin: '8px 0', fontSize: '15px' }}>{inputError}</div>
      )}

      {/* بيانات العقد تظهر فقط إذا لم توجد سجلات */}
      {contractData && logs.length === 0 && (
        <div
          style={{
            marginBottom: '18px',
            background: '#fffbe7', // أصفر فاتح جداً
            borderRadius: '18px',
            boxShadow: '0 2px 12px rgba(106,27,154,0.10)',
            border: '1.5px solid #ffe066',
            padding: '16px 18px',
            display: 'flex',
            flexDirection: 'column',
            alignItems: 'center',
            width: isMobile ? '100%' : 'fit-content',
            maxWidth: isMobile ? '98vw' : '600px',
            minWidth: isMobile ? '90%' : undefined,
            alignSelf: 'center',
            marginLeft: 'auto',
            marginRight: 'auto',
            overflowX: 'auto',
          }}
        >
          {/* whiteSpace: nowrap لكل سطر */}
          <p style={{ margin: '0 0 8px 0', fontWeight: 700, color: '#6a1b9a', fontSize: 18, display: 'flex', alignItems: 'center', whiteSpace: 'nowrap' }}>
            <span style={{ fontSize: 20, marginRight: 6, color: '#29b6f6' }}>■</span>
            Booking: <span style={{ fontWeight: 400, color: '#222', marginLeft: 6 }}>{contractData['Booking Number']}</span>
          </p>
          <p style={{ margin: '0 0 8px 0', fontWeight: 700, color: '#6a1b9a', fontSize: 18, display: 'flex', alignItems: 'center', whiteSpace: 'nowrap' }}>
            <span style={{ fontSize: 20, marginRight: 6, color: '#b39ddb' }}>📄</span>
            Contract: <span style={{ fontWeight: 400, color: '#222', marginLeft: 6 }}>{contractData['Contract No.']}</span>
          </p>
          <p style={{ margin: '0 0 8px 0', fontWeight: 700, color: '#6a1b9a', fontSize: 18, display: 'flex', alignItems: 'center', whiteSpace: 'nowrap' }}>
            <span style={{ fontSize: 20, marginRight: 6, color: '#6a1b9a' }}>👤</span>
            Customer: <span style={{ fontWeight: 400, color: '#222', marginLeft: 6 }}>{contractData['Customer']}</span>
          </p>
          {contractSource && (
            <p style={{ margin: 0, fontWeight: 700, color: contractSource === 'Open Contract' ? '#4CAF50' : '#ff9800', fontSize: 16, display: 'flex', alignItems: 'center', whiteSpace: 'nowrap' }}>
              <span style={{ fontSize: 18, marginRight: 6 }}>{contractSource === 'Open Contract' ? '🟢' : '🟠'}</span>
              {contractSource}
            </p>
          )}
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

      {/* خانة تاريخ بداية العقد */}
      {dateLocked && lastDate ? (
        <div style={{ marginBottom: '8px' }}>
          <label style={{ fontWeight: 'bold', color: '#6a1b9a', fontSize: '16px', display: 'block', marginBottom: '4px' }}>
            📅 Contract Start Date
          </label>
          <div style={{
            background: '#fffde7',
            color: '#b28704',
            fontSize: '18px',
            fontWeight: 'bold',
            borderRadius: '8px',
            padding: '8px 16px',
            marginBottom: '4px',
            letterSpacing: '1px',
            boxShadow: '0 1px 4px rgba(178,135,4,0.07)'
          }}>
            {formatDateToDMY(lastDate)}
          </div>
        </div>
      ) : (
        <div style={{ marginBottom: '8px' }}>
          <label style={{ fontWeight: 'bold', color: '#6a1b9a', fontSize: '16px', display: 'block', marginBottom: '4px' }}>
            📅 Contract Start Date
          </label>
          <input
            type="date"
            placeholder="📅 Contract Start Date"
            value={date}
            onChange={e => setDate(e.target.value)}
            style={inputStyle}
            onKeyDown={e => { if (e.key === 'Enter') handleAddLog(); }}
            onFocus={handleInputFocus}
            onBlur={handleInputBlur}
            onPaste={e => handleDatePaste(e, setDate)}
          />
          {contractData && (
            <p style={{ color: '#888', fontSize: '13px' }}>
              Contract start date not found, please enter it manually.
            </p>
          )}
        </div>
      )}

      {/* خانة اختيارية لإدخال تاريخ نهاية العقد */}
      {endDateInputVisible && (
        <div style={{ marginBottom: '8px' }}>
          <label style={{ fontWeight: 'bold', color: '#b71c1c', fontSize: '16px', display: 'block', marginBottom: '4px' }}>
            🛑 Contract End Date (optional)
          </label>
          <input
            type="date"
            placeholder="📅 Contract End Date (optional)"
            value={manualEndDate}
            onChange={e => setManualEndDate(e.target.value)}
            style={inputStyle}
            onFocus={handleInputFocus}
            onBlur={handleInputBlur}
            onPaste={e => handleDatePaste(e, setManualEndDate)}
          />
          <div style={{ color: '#b71c1c', fontSize: '13px', marginTop: '2px' }}>
            If you enter this date, calculations will be up to this day only.
          </div>
        </div>
      )}

      <input
        type="number"
        placeholder="🚗 OUT (Start KM)"
        value={out}
        onChange={e => setOut(e.target.value)}
        style={inputStyle}
        onKeyDown={e => { if (e.key === 'Enter') handleAddLog(); }}
        ref={outInputRef}
        onFocus={handleInputFocus}
        onBlur={handleInputBlur}
      />
      <input
        type="number"
        placeholder="🚙 IN (End KM)"
        value={inVal}
        onChange={e => setInVal(e.target.value)}
        style={inputStyle}
        onKeyDown={e => { if (e.key === 'Enter') handleAddLog(); }}
        onFocus={handleInputFocus}
        onBlur={handleInputBlur}
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
        {exceeded > 0 && contractData && (
          <button
            style={{
              ...buttonStyle,
              background: '#d32f2f',
              color: '#fff',
            }}
            onClick={generateTaxInvoice}
          >
            📄 Generate Tax Invoice
          </button>
        )}
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
        
        {/* تاريخ الإغلاق إذا كان موجود */}
        {manualEndDate && (
          <div style={{
            background: '#ffebee',
            color: '#b71c1c',
            fontSize: '24px',
            fontWeight: 'bold',
            borderRadius: '8px',
            padding: '12px 20px',
            margin: '0 0 18px 0',
            letterSpacing: '1px',
            boxShadow: '0 1px 4px rgba(183,28,28,0.07)'
          }}>
            Contract End Date: {formatDateToDMY(manualEndDate)}
          </div>
        )}

        {logs.length > 0 && (
          <>
            {contractData && (
              <div
                style={{
                  marginBottom: '18px',
                  background: '#fffbe7',
                  borderRadius: '18px',
                  boxShadow: '0 2px 12px rgba(106,27,154,0.10)',
                  border: '1.5px solid #ffe066',
                  padding: '16px 18px',
                  display: 'flex',
                  flexDirection: 'column',
                  alignItems: 'center',
                  width: isMobile ? '100%' : 'fit-content',
                  maxWidth: isMobile ? '98vw' : '600px',
                  minWidth: isMobile ? '90%' : undefined,
                  alignSelf: 'center',
                  marginLeft: 'auto',
                  marginRight: 'auto',
                  overflowX: 'auto',
                }}
              >
                {/* whiteSpace: nowrap لكل سطر */}
                <p style={{ margin: '0 0 8px 0', fontWeight: 700, color: '#6a1b9a', fontSize: 18, display: 'flex', alignItems: 'center', whiteSpace: 'nowrap' }}>
                  <span style={{ fontSize: 20, marginRight: 6, color: '#29b6f6' }}>■</span>
                  Booking: <span style={{ fontWeight: 400, color: '#222', marginLeft: 6 }}>{contractData['Booking Number']}</span>
                </p>
                <p style={{ margin: '0 0 8px 0', fontWeight: 700, color: '#6a1b9a', fontSize: 18, display: 'flex', alignItems: 'center', whiteSpace: 'nowrap' }}>
                  <span style={{ fontSize: 20, marginRight: 6, color: '#b39ddb' }}>📄</span>
                  Contract: <span style={{ fontWeight: 400, color: '#222', marginLeft: 6 }}>{contractData['Contract No.']}</span>
                </p>
                <p style={{ margin: '0 0 8px 0', fontWeight: 700, color: '#6a1b9a', fontSize: 18, display: 'flex', alignItems: 'center', whiteSpace: 'nowrap' }}>
                  <span style={{ fontSize: 20, marginRight: 6, color: '#6a1b9a' }}>👤</span>
                  Customer: <span style={{ fontWeight: 400, color: '#222', marginLeft: 6 }}>{contractData['Customer']}</span>
                </p>
                {contractSource && (
                  <p style={{ margin: 0, fontWeight: 700, color: contractSource === 'Open Contract' ? '#4CAF50' : '#ff9800', fontSize: 16, display: 'flex', alignItems: 'center', whiteSpace: 'nowrap' }}>
                    <span style={{ fontSize: 18, marginRight: 6 }}>{contractSource === 'Open Contract' ? '🟢' : '🟠'}</span>
                    {contractSource}
                  </p>
                )}
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

      {/* نافذة منبثقة لرقم المرجع */}
      {showRefModal && (
        <div style={{
          position: 'fixed',
          top: 0,
          left: 0,
          width: '100%',
          height: '100%',
          background: 'rgba(0,0,0,0.5)',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          zIndex: 10000
        }}>
          <div style={{
            background: '#fff',
            padding: '30px',
            borderRadius: '12px',
            boxShadow: '0 8px 32px rgba(0,0,0,0.3)',
            maxWidth: '400px',
            width: '90%'
          }}>
            <h3 style={{ margin: '0 0 20px 0', color: '#6a1b9a', textAlign: 'center' }}>Generate Tax Invoice</h3>
            <input
              type="text"
              placeholder="Reference number (optional - will be added after ALWFQ-)"
              value={refInput}
              onChange={e => setRefInput(e.target.value)}
              style={{
                width: '100%',
                padding: '12px',
                border: '2px solid #ffe066',
                borderRadius: '8px',
                fontSize: '16px',
                marginBottom: '15px',
                boxSizing: 'border-box'
              }}
              autoFocus
            />
            <div style={{ marginBottom: '20px' }}>
              <p style={{ margin: '0 0 10px 0', color: '#666', fontWeight: 'bold' }}>Export Format:</p>
              <div style={{ display: 'flex', gap: '10px', justifyContent: 'center' }}>
                <label style={{ display: 'flex', alignItems: 'center', cursor: 'pointer' }}>
                  <input
                    type="radio"
                    name="exportType"
                    value="word"
                    checked={exportType === 'word'}
                    onChange={e => setExportType(e.target.value)}
                    style={{ marginRight: '5px' }}
                  />
                  Word Only
                </label>
                <label style={{ display: 'flex', alignItems: 'center', cursor: 'pointer' }}>
                  <input
                    type="radio"
                    name="exportType"
                    value="pdf"
                    checked={exportType === 'pdf'}
                    onChange={e => setExportType(e.target.value)}
                    style={{ marginRight: '5px' }}
                  />
                  PDF Only
                </label>
                <label style={{ display: 'flex', alignItems: 'center', cursor: 'pointer' }}>
                  <input
                    type="radio"
                    name="exportType"
                    value="both"
                    checked={exportType === 'both'}
                    onChange={e => setExportType(e.target.value)}
                    style={{ marginRight: '5px' }}
                  />
                  Both
                </label>
              </div>
            </div>
            <div style={{ display: 'flex', gap: '10px', justifyContent: 'center' }}>
              <button
                onClick={() => {
                  setShowRefModal(false);
                  setRefInput('');
                  handleRefSubmit();
                }}
                style={{
                  padding: '10px 20px',
                  background: '#4CAF50',
                  color: '#fff',
                  border: 'none',
                  borderRadius: '6px',
                  cursor: 'pointer',
                  fontSize: '16px'
                }}
              >
                Generate Invoice
              </button>
              <button
                onClick={() => {
                  setShowRefModal(false);
                  setRefInput('');
                }}
                style={{
                  padding: '10px 20px',
                  background: '#e53935',
                  color: '#fff',
                  border: 'none',
                  borderRadius: '6px',
                  cursor: 'pointer',
                  fontSize: '16px'
                }}
              >
                Cancel
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

export default KilometerTracker;