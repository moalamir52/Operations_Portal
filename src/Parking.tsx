import React, { useState } from 'react';
import * as XLSX from 'xlsx';
import { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, ISectionOptions } from 'docx';
import { saveAs } from 'file-saver';


// Helper function to create an invoice section.
// This logic was moved here to avoid code duplication.
const createInvoiceSection = (row: any, invoiceDate: string, trnNumber: string): ISectionOptions => {
  const fontProps = { font: 'Calibri', size: 22, color: '000000' }; // 22 = 11pt

  // Helper function to format Excel's numeric date format
  const formatExcelDate = (excelDate: any): string => {
    if (!excelDate) return '';
    // If the date is a number (Excel's date serial number)
    if (typeof excelDate === 'number') {
      const date = new Date((excelDate - 25569) * 86400 * 1000);
      return date.toLocaleDateString('en-GB'); // Format: DD/MM/YYYY
    }
    // If the date is already a string
    if (typeof excelDate === 'string') {
      return excelDate;
    }
    return '';
  };



  const invoiceNumber = row['Tax_Invoice_No'] ? row['Tax_Invoice_No'] : '';

  // Function to calculate exit date
  const calculateExitDate = (entryDate: string, entryTime: string, exitTime: string): string => {
    if (!entryDate || !entryTime || !exitTime) return entryDate || '';
    
    // Convert times to 24-hour format for comparison
    const parseTime = (timeStr: string): number => {
      if (!timeStr || typeof timeStr !== 'string') return 0;
      const time = timeStr.toLowerCase().trim();
      let [hours, minutes] = time.replace(/[ap]m/, '').split(':').map(Number);
      if (time.includes('pm') && hours !== 12) hours += 12;
      if (time.includes('am') && hours === 12) hours = 0;
      return hours * 60 + (minutes || 0);
    };
    
    const entryMinutes = parseTime(entryTime);
    const exitMinutes = parseTime(exitTime);
    
    // If exit time is earlier than entry time, it's the next day
    if (exitMinutes < entryMinutes) {
      // Parse DD/MM/YYYY format
      const [day, month, year] = entryDate.split('/').map(Number);
      const date = new Date(year, month - 1, day);
      date.setDate(date.getDate() + 1);
      return date.toLocaleDateString('en-GB');
    }
    
    return entryDate;
  };

  const formattedDate = formatExcelDate(row['Date']);
  const exitDate = calculateExitDate(formattedDate, row['Time_In'] || '', row['Time_Out'] || '');

  return {
    properties: {
      page: {
        margin: { top: 1440 }, // Margin in DXA (twentieth of a point)
      },
    },
    children: [
        new Paragraph({ children: [new TextRun({ text: '', ...fontProps })] }),
        new Paragraph({ children: [new TextRun({ text: '', ...fontProps })] }),
        new Paragraph({ children: [new TextRun({ text: '', ...fontProps })] }),
        new Paragraph({ children: [new TextRun({ text: '', ...fontProps })] }),
        new Paragraph({ children: [new TextRun({ text: '', ...fontProps })] }),
        new Paragraph({ children: [new TextRun({ text: '', ...fontProps })] }),
        new Paragraph({ children: [new TextRun({ text: 'Tax Invoice', font: 'Arial', size: 52, bold: true, color: '000000' })], heading: 'Heading1' }),
        new Paragraph({ children: [new TextRun({ text: '', ...fontProps })] }),
        new Paragraph({
          children: [
            new TextRun({ text: `Date: ${invoiceDate}`, ...fontProps }),
            new TextRun({ text: '                                                                                ', ...fontProps }),
            new TextRun({ text: `Ref: ${invoiceNumber ? ' ' + invoiceNumber : ''}`, ...fontProps }),
          ],
        }),
        new Paragraph({
          children: [
            new TextRun({ text: `                                                                                                               TRN#: ${trnNumber}`, ...fontProps }),
          ],
        }),
        new Paragraph({ children: [new TextRun({ text: '', ...fontProps })] }),
        new Paragraph({ children: [new TextRun({ text: 'Invygo Tech FZ-LLC', ...fontProps })] }),
        new Paragraph({ children: [new TextRun({ text: 'Dubai Internet City', ...fontProps })] }),
        new Paragraph({ children: [new TextRun({ text: 'Dubai, U.A.E.', ...fontProps })] }),
        new Paragraph({ children: [new TextRun({ text: '', ...fontProps })] }),
        new Paragraph({ children: [new TextRun({ text: 'SUB: Micro Lease Cars', ...fontProps })] }),
        new Paragraph({ children: [new TextRun({ text: '', ...fontProps })] }),
        new Paragraph({ children: [new TextRun({ text: 'Dear Sir,', ...fontProps })] }),
        new Paragraph({ children: [new TextRun({ text: 'We thank you for your business renting the below vehicle.', ...fontProps })] }),
        new Paragraph({ children: [new TextRun({ text: '', ...fontProps })] }),
        // Invoice Table
        new Table({
          rows: [
            new TableRow({
              height: { value: 800, rule: 'exact' },
              children: [
                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: 'No.', ...fontProps, bold: true })], alignment: 'center' })], width: { size: 1000, type: 'dxa' }, verticalAlign: 'center' }),
                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: 'Description', ...fontProps, bold: true })], alignment: 'center' })], width: { size: 6000, type: 'dxa' }, verticalAlign: 'center' }),
                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: 'Duration', ...fontProps, bold: true })], alignment: 'center' })], width: { size: 2000, type: 'dxa' }, shading: { fill: 'F0F0F0' }, verticalAlign: 'center' }),
                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: 'Total Price', ...fontProps, bold: true })], alignment: 'center' })], width: { size: 2000, type: 'dxa' }, shading: { fill: 'F0F0F0' }, verticalAlign: 'center' }),
              ],
            }),
            new TableRow({
              children: [
                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: '1', ...fontProps })], alignment: 'center' })], width: { size: 1000, type: 'dxa' }, verticalAlign: 'center' }),
                new TableCell({
                  children: [
                    new Paragraph({ children: [new TextRun({ text: ` Name: ${row['Customer Name:'] || row['Customer Name'] || row['Customer'] || ''}`, ...fontProps })] }),
                    new Paragraph({ children: [new TextRun({ text: ` Booking ID: ${row['Dealer_Booking_Number'] || ''}`, ...fontProps })] }),
                    new Paragraph({ children: [new TextRun({ text: ` R/A: ${row['Contract'] || ''}`, ...fontProps })] }),
                    new Paragraph({ children: [new TextRun({ text: ` Vehicle: ${row['Model'] || ''} - ${row['Plate_Number'] || ''}`, ...fontProps })] }),
                    new Paragraph({ children: [new TextRun({ text: ` Entry: ${formatExcelDate(row['Date'])} - ${row['Time_In'] || ''}`, ...fontProps })] }),
                    new Paragraph({ children: [new TextRun({ text: ` Exit: ${exitDate} - ${row['Time_Out'] || ''}`, ...fontProps })] }),
                  ],
                  width: { size: 6000, type: 'dxa' },
                }),
                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: `${(row['Time'] || '0').toString().replace(/hrs?/gi, '').trim()} hours`, ...fontProps })], alignment: 'center' })], width: { size: 2000, type: 'dxa' }, verticalAlign: 'center' }),
                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: `${parseFloat(row['Amount'] || '0').toFixed(2)} `, ...fontProps })], alignment: 'right' })], width: { size: 2000, type: 'dxa' }, verticalAlign: 'center', margins: { right: 144 } }),
              ],
            }),
            new TableRow({
              children: [
                new TableCell({
                  children: [new Paragraph({ children: [new TextRun({ text: '', ...fontProps })] })],
                  columnSpan: 2,
                  borders: {
                    top: { style: 'single' },
                    left: { style: 'none', size: 4, color: '000000' },
                    right: { style: 'none', size: 4, color: '000000' },
                    bottom: { style: 'none', size: 4, color: '000000' },
                  },
                }),
                new TableCell({
                  children: [
                    new Paragraph({ children: [new TextRun({ text: 'TOTAL:', ...fontProps, bold: true, size: 30 })], alignment: 'center' }),
                  ],
                  shading: { fill: 'D9D9D9' },
                  rowSpan: 2,
                  verticalAlign: 'center',
                }),
                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: `AED ${parseFloat(row['Amount'] || '0').toFixed(2)} `, ...fontProps, bold: true, size: 30 })], alignment: 'right' })], shading: { fill: 'D9D9D9' }, margins: { right: 144 } }),
              ],
            }),
          ],
        }),
        new Paragraph({ children: [new TextRun({ text: '', ...fontProps })] }),
        new Paragraph({ children: [new TextRun({ text: 'General Conditions:', ...fontProps, bold: true, underline: {} })] }),
        new Paragraph({ children: [new TextRun({ text: '', ...fontProps })] }),
        new Paragraph({
          children: [
            new TextRun({ text: 'Terms of Payment', ...fontProps }),
            new TextRun({ text: '             : within 7 days', ...fontProps }),
          ],
        }),
        new Paragraph({ children: [new TextRun({ text: '', ...fontProps })] }),
        new Paragraph({ children: [new TextRun({ text: '', ...fontProps })] }),
        new Paragraph({ children: [new TextRun({ text: '', ...fontProps })] }),
        new Paragraph({ children: [new TextRun({ text: 'Thanking you and assuring you of our best co-operation and services at all times.', ...fontProps })] }),
        new Paragraph({ children: [new TextRun({ text: '', ...fontProps })] }),
        new Paragraph({ children: [new TextRun({ text: '', ...fontProps })] }),
        new Paragraph({ children: [new TextRun({ text: 'Best Regards,', ...fontProps })] }),
        new Paragraph({ children: [new TextRun({ text: '', ...fontProps })] }),
        new Paragraph({ children: [new TextRun({ text: '', ...fontProps })] }),
        new Paragraph({
          children: [
            new TextRun({ text: 'Saudian Alwefaq Rent A Car', ...fontProps, bold: true }),
          ],
        }),
    ],
  };
};

function ExcelToWord() {
  const [excelFile, setExcelFile] = useState<File | null>(null);
  const [status, setStatus] = useState('');
  const [selectedDate, setSelectedDate] = useState('');
  const [trnNumber, setTrnNumber] = useState('100397403500003');


  const handleExcelUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      setExcelFile(e.target.files[0]);
    }
  };

  const handleConvert = async () => {
    if (!excelFile) {
      setStatus('Please upload an Excel or CSV file first.');
      return;
    }

    let invoiceDate = selectedDate;
    if (!invoiceDate) {
      const today = new Date();
      invoiceDate = today.toISOString().split('T')[0]; // yyyy-mm-dd format
      setSelectedDate(invoiceDate);
    }

    if (!trnNumber) {
      setStatus('Please enter the TRN number first.');
      return;
    }

    setStatus('Converting...');

    let rows: any[] = [];
    
    // Check file type and process accordingly
    if (excelFile.name.toLowerCase().endsWith('.csv')) {
      // Handle CSV files
      const text = await excelFile.text();
      const workbook = XLSX.read(text, { type: 'string' });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      rows = XLSX.utils.sheet_to_json(sheet);
    } else {
      // Handle Excel files
      const data = await excelFile.arrayBuffer();
      const workbook = XLSX.read(data, { type: 'array' });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      rows = XLSX.utils.sheet_to_json(sheet);
    }

    const sections = rows.map((row) => {
      return createInvoiceSection(row, invoiceDate, trnNumber);
    });

    const doc = new Document({ sections });
    const buffer = await Packer.toBlob(doc);
    saveAs(buffer, 'invoices.docx');
    setStatus('Success!');
  };

  return (
    <div style={{ maxWidth: 600, margin: '40px auto', padding: 40, background: 'linear-gradient(135deg, #f3e7ff 0%, #fffbe7 100%)', borderRadius: 24, boxShadow: '0 8px 32px rgba(106,27,154,0.12)', fontFamily: 'Segoe UI, Arial, sans-serif' }}>
      <h2 style={{ color: '#6a1b9a', fontSize: 36, fontWeight: 700, textAlign: 'center', marginBottom: 32, letterSpacing: 1 }}>Parking Invoice Creation </h2>
      <div style={{ display: 'flex', flexDirection: 'column', gap: 24 }}>
        <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
          <label htmlFor="excel-upload" style={{ fontSize: 18, fontWeight: 500, color: '#333' }}>Excel or CSV file:</label>
          <input id="excel-upload" type="file" accept=".xlsx,.xls,.csv" onChange={handleExcelUpload} style={{ fontSize: 18, padding: '8px 12px', borderRadius: 8, border: '1px solid #ccc', background: '#fff' }} />
        </div>
        <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
          <label htmlFor="date-input" style={{ fontSize: 18, fontWeight: 500, color: '#333' }}>Select date:</label>
          <input id="date-input" type="date" value={selectedDate} onChange={(e) => setSelectedDate(e.target.value)} style={{ fontSize: 18, padding: '8px 12px', borderRadius: 8, border: '1px solid #ccc', background: '#fff' }} />
        </div>



        <button onClick={handleConvert} style={{ background: 'linear-gradient(90deg, #6a1b9a 0%, #8e24aa 100%)', color: '#fff', padding: '16px 0', fontSize: 22, fontWeight: 600, border: 'none', borderRadius: 12, cursor: 'pointer', marginTop: 16, boxShadow: '0 2px 8px rgba(106,27,154,0.10)' }}>Start Generating Files</button>
        {status && <div style={{ marginTop: 18, color: '#b71c1c', fontWeight: 'bold', fontSize: 20, textAlign: 'center' }}>{status}</div>}
      </div>
    </div>
  );
}

export default ExcelToWord;