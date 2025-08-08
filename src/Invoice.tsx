import React, { useState } from 'react';
import * as XLSX from 'xlsx';
import { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell } from 'docx';
import { saveAs } from 'file-saver';

function ExcelToWord() {
  const [excelFile, setExcelFile] = useState<File | null>(null);
  const [templateFile, setTemplateFile] = useState<File | null>(null);
  const [status, setStatus] = useState('');
  const [selectedDate, setSelectedDate] = useState('');
  const [trnNumber, setTrnNumber] = useState('100397403500003');

  // قراءة ملف Excel
  const handleExcelUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      setExcelFile(e.target.files[0]);
    }
  };

  // قراءة ملف Word (القالب)
  const handleTemplateUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      setTemplateFile(e.target.files[0]);
    }
  };

  // استخراج متغيرات من نص القالب
  const extractVariables = (text: string) => {
    const regex = /{{(\w+)}}/g;
    const vars = [];
    let match;
    while ((match = regex.exec(text)) !== null) {
      vars.push(match[1]);
    }
    return vars;
  };

  // تحويل تاريخ Excel إلى نص مقروء
  const formatExcelDate = (excelDate: any) => {
    if (!excelDate) return '';
    
    // إذا كان التاريخ رقم (Excel date)
    if (typeof excelDate === 'number') {
      const date = new Date((excelDate - 25569) * 86400 * 1000);
      return date.toLocaleDateString('en-GB'); // تنسيق DD/MM/YYYY
    }
    
    // إذا كان التاريخ نص
    if (typeof excelDate === 'string') {
      return excelDate;
    }
    
    return '';
  };
// Convert each row to a Word file
const handleConvert = async () => {
  if (!excelFile) {
    setStatus('Please upload an Excel file first');
    return;
  }
  if (!selectedDate) {
    setStatus('Please select the date first');
    return;
  }
  if (!trnNumber) {
    setStatus('Please enter the TRN number first');
    return;
  }
  setStatus('Converting...');
    // قراءة بيانات Excel
    const data = await excelFile.arrayBuffer();
    const workbook = XLSX.read(data, { type: 'array' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows: any[] = XLSX.utils.sheet_to_json(sheet);

    // بناء الفاتورة لكل صف
    const sections = rows.map((row, i) => {
      // إعداد خصائص الخط
      const fontProps = { font: 'Calibri', size: 22, color: '000000' }; // 22 = 11pt
      // حساب رقم TRN للفاتورة الحالية
      const currentTrnNumber = (parseInt(trnNumber) + i).toString();
      // Determine Salik Date text
      let salikDateText = '';
      const startDate = formatExcelDate(row['Date']);
      const endDate = formatExcelDate(row['End Date']);
      if (endDate && endDate !== '') {
        salikDateText = `Salik Date: ${startDate} - ${endDate}`;
      } else {
        salikDateText = `Salik Date: ${startDate}`;
      }
      return {
        properties: {
          page: {
            margin: { top: 1440 }
          }
        },
        children: [
          new Paragraph({ children: [new TextRun({ text: '' , ...fontProps })] }),
          new Paragraph({ children: [new TextRun({ text: '' , ...fontProps })] }),
          new Paragraph({ children: [new TextRun({ text: '' , ...fontProps })] }),
          new Paragraph({ children: [new TextRun({ text: '' , ...fontProps })] }),
          new Paragraph({ children: [new TextRun({ text: '' , ...fontProps })] }),
          new Paragraph({ children: [new TextRun({ text: '' , ...fontProps })] }),
          new Paragraph({ children: [new TextRun({ text: 'Tax Invoice', font: 'Arial', size: 52, bold: true, color: '000000' })], heading: 'Heading1' }),
          new Paragraph({ children: [new TextRun({ text: '' , ...fontProps })] }),
          new Paragraph({
            children: [
              new TextRun({ text: `Date: ${selectedDate}`, ...fontProps }),
              new TextRun({ text: '                                                                                ', ...fontProps }),
              new TextRun({ text: 'Ref: ALWFQ', ...fontProps })
            ]
          }),
          new Paragraph({
            children: [
              new TextRun({ text: `TRN#: ${currentTrnNumber}`, ...fontProps })
            ]
          }),
          new Paragraph({ children: [new TextRun({ text: '' , ...fontProps })] }),
          new Paragraph({ children: [new TextRun({ text: 'Invygo Tech FZ-LLC', ...fontProps })] }),
          new Paragraph({ children: [new TextRun({ text: 'Dubai Internet City', ...fontProps })] }),
          new Paragraph({ children: [new TextRun({ text: 'Dubai, U.A.E.', ...fontProps })] }),
          new Paragraph({ children: [new TextRun({ text: '' , ...fontProps })] }),
          new Paragraph({ children: [new TextRun({ text: 'SUB: Micro Lease Cars', ...fontProps })] }),
          new Paragraph({ children: [new TextRun({ text: '' , ...fontProps })] }),
          new Paragraph({ children: [new TextRun({ text: 'Dear Sir,', ...fontProps })] }),
          new Paragraph({ children: [new TextRun({ text: 'We thank you for your business renting the below vehicle.', ...fontProps })] }),
          new Paragraph({ children: [new TextRun({ text: '' , ...fontProps })] }),
          // جدول الفاتورة
          new Table({
            rows: [
              new TableRow({
                children: [
                  new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: 'No.', ...fontProps, bold: true })] })], width: { size: 1000, type: 'dxa' } }),
                  new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: 'Description', ...fontProps, bold: true })] })], width: { size: 6000, type: 'dxa' } }),
                  new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: 'Salik Trips', ...fontProps, bold: true })], alignment: 'center' })], width: { size: 2000, type: 'dxa' }, shading: { fill: 'F0F0F0' } }),
                  new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: 'Total Price', ...fontProps, bold: true })], alignment: 'center' })], width: { size: 2000, type: 'dxa' }, shading: { fill: 'F0F0F0' } }),
                ]
              }),
              new TableRow({
                children: [
                  new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: '1', ...fontProps })], alignment: 'center' })], width: { size: 1000, type: 'dxa' }, verticalAlign: 'center' }),
                  new TableCell({
                    children: [
                      new Paragraph({ children: [new TextRun({ text: `Name: ${row['Customer'] || ''}`, ...fontProps })] }),
                      new Paragraph({ children: [new TextRun({ text: `Booking ID: ${row['Booking Number'] || ''}`, ...fontProps })] }),
                      new Paragraph({ children: [new TextRun({ text: `R/A: ${row['Contract No.'] || ''}`, ...fontProps })] }),
                      new Paragraph({ children: [new TextRun({ text: `Vehicle: ${row['Model'] || ''} ${row['Plate No.'] || ''}`, ...fontProps })] }),
                      new Paragraph({ children: [new TextRun({ text: salikDateText, ...fontProps })] }),
                    ],
                    width: { size: 6000, type: 'dxa' }
                  }),
                  new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: `${row['Salik Trips'] || '0'} Trips`, ...fontProps })], alignment: 'center' })], width: { size: 2000, type: 'dxa' }, verticalAlign: 'center' }),
                  new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: `${row['Total Price'] || '0.00'}`, ...fontProps })], alignment: 'center' })], width: { size: 2000, type: 'dxa' }, verticalAlign: 'center' }),
                ]
              }),
              new TableRow({
                children: [
                  new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: '', ...fontProps })] })], columnSpan: 2 }),
                  new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: 'TOTAL:', ...fontProps, bold: true })], alignment: 'center' })], shading: { fill: 'D9D9D9' } }),
                  new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: `AED ${row['Total Price'] || '0.00'}`, ...fontProps, bold: true })], alignment: 'center' })], shading: { fill: 'D9D9D9' } }),
                ]
              })
            ]
          }),
          new Paragraph({ children: [new TextRun({ text: '' , ...fontProps })] }),
          new Paragraph({ children: [new TextRun({ text: 'General Conditions:', ...fontProps, bold: true, underline: {} })] }),
          new Paragraph({ children: [new TextRun({ text: '' , ...fontProps })] }),
          new Paragraph({ 
            children: [
              new TextRun({ text: 'Terms of Payment', ...fontProps }),
              new TextRun({ text: ' : within 7 days', ...fontProps })
            ]
          }),
          new Paragraph({ children: [new TextRun({ text: '' , ...fontProps })] }),
          new Paragraph({ children: [new TextRun({ text: '' , ...fontProps })] }),
          new Paragraph({ children: [new TextRun({ text: '' , ...fontProps })] }),
          new Paragraph({ children: [new TextRun({ text: 'Thanking you and assuring you of our best co-operation and services at all times.', ...fontProps })] }),
          new Paragraph({ children: [new TextRun({ text: '' , ...fontProps })] }),
          new Paragraph({ children: [new TextRun({ text: '' , ...fontProps })] }),
          new Paragraph({ children: [new TextRun({ text: 'Best Regards,', ...fontProps })] }),
          new Paragraph({ children: [new TextRun({ text: '' , ...fontProps })] }),
          new Paragraph({ children: [new TextRun({ text: '' , ...fontProps })] }),
          new Paragraph({
            children: [
              new TextRun({ text: 'Saudian Alwefaq Rent A Car', ...fontProps, bold: true })
            ]
          }),
        ]
      };
    });
    const doc = new Document({ sections });
    const buffer = await Packer.toBlob(doc);
    saveAs(buffer, 'invoices.docx');
    setStatus('Done !');
  };

  return (
    <div style={{ maxWidth: 600, margin: '40px auto', padding: 40, background: 'linear-gradient(135deg, #f3e7ff 0%, #fffbe7 100%)', borderRadius: 24, boxShadow: '0 8px 32px rgba(106,27,154,0.12)', fontFamily: 'Segoe UI, Arial, sans-serif' }}>
      <h2 style={{ color: '#6a1b9a', fontSize: 36, fontWeight: 700, textAlign: 'center', marginBottom: 32, letterSpacing: 1 }}>Invoice Creation Software</h2>
      <div style={{ display: 'flex', flexDirection: 'column', gap: 24 }}>
        <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
          <label htmlFor="excel-upload" style={{ fontSize: 18, fontWeight: 500, color: '#333' }}>Excel file:</label>
          <input id="excel-upload" type="file" accept=".xlsx,.xls" onChange={handleExcelUpload} style={{ fontSize: 18, padding: '8px 12px', borderRadius: 8, border: '1px solid #ccc', background: '#fff' }} />
        </div>
        <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
          <label htmlFor="date-input" style={{ fontSize: 18, fontWeight: 500, color: '#333' }}>Select date:</label>
          <input id="date-input" type="date" value={selectedDate} onChange={(e) => setSelectedDate(e.target.value)} style={{ fontSize: 18, padding: '8px 12px', borderRadius: 8, border: '1px solid #ccc', background: '#fff' }} />
        </div>
        <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
          <label htmlFor="trn-input" style={{ fontSize: 18, fontWeight: 500, color: '#333' }}>TRN number:</label>
          <input id="trn-input" type="text" value={trnNumber} onChange={(e) => setTrnNumber(e.target.value)} style={{ fontSize: 18, padding: '8px 12px', borderRadius: 8, border: '1px solid #ccc', background: '#fff' }} />
        </div>
        <button onClick={handleConvert} style={{ background: 'linear-gradient(90deg, #6a1b9a 0%, #8e24aa 100%)', color: '#fff', padding: '16px 0', fontSize: 22, fontWeight: 600, border: 'none', borderRadius: 12, cursor: 'pointer', marginTop: 16, boxShadow: '0 2px 8px rgba(106,27,154,0.10)' }}>Start Generate Files</button>
        {status && <div style={{ marginTop: 18, color: '#b71c1c', fontWeight: 'bold', fontSize: 20, textAlign: 'center' }}>{status}</div>}
      </div>
    </div>
  );
}

export default ExcelToWord;
