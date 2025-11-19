import React, { useState } from 'react';
import * as XLSX from 'xlsx';
import { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, ISectionOptions } from 'docx';
import { saveAs } from 'file-saver';

// Component to preview invoice
const InvoicePreview = ({ data }: { data: any }) => {
  const { row, invoiceDate, trnNumber } = data;
  
  const formatExcelDate = (excelDate: any): string => {
    if (!excelDate) return '';
    if (typeof excelDate === 'number') {
      const date = new Date((excelDate - 25569) * 86400 * 1000);
      return date.toLocaleDateString('en-GB');
    }
    if (typeof excelDate === 'string') {
      return excelDate;
    }
    return '';
  };

  const calculateExitDate = (entryDate: string, entryTime: string, exitTime: string): string => {
    if (!entryDate || !entryTime || !exitTime) return entryDate || '';
    
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
    
    if (exitMinutes < entryMinutes) {
      const [day, month, year] = entryDate.split('/').map(Number);
      const date = new Date(year, month - 1, day);
      date.setDate(date.getDate() + 1);
      return date.toLocaleDateString('en-GB');
    }
    
    return entryDate;
  };

  const formattedDate = formatExcelDate(row['Date']);
  const exitDate = calculateExitDate(formattedDate, row['Time_In'] || '', row['Time_Out'] || '');
  const invoiceNumber = row['Tax_Invoice_No'] ? row['Tax_Invoice_No'] : '';

  return (
    <div style={{ padding: 40, fontFamily: 'Calibri, Arial, sans-serif', fontSize: 11, lineHeight: 1.4, color: '#000', background: '#fff', minWidth: 600 }}>
      <div style={{ textAlign: 'center', marginBottom: 30 }}>
        <h1 style={{ fontSize: 26, fontWeight: 'bold', margin: 0, fontFamily: 'Arial' }}>Tax Invoice</h1>
      </div>
      
      <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: 20 }}>
        <span>Date: {invoiceDate}</span>
        <span>Ref: {invoiceNumber}</span>
      </div>
      
      <div style={{ textAlign: 'right', marginBottom: 20 }}>
        <span>TRN#: {trnNumber}</span>
      </div>
      
      <div style={{ marginBottom: 20 }}>
        <div>Invygo Tech FZ-LLC</div>
        <div>Dubai Internet City</div>
        <div>Dubai, U.A.E.</div>
      </div>
      
      <div style={{ marginBottom: 20 }}>
        <div>SUB: Micro Lease Cars</div>
      </div>
      
      <div style={{ marginBottom: 20 }}>
        <div>Dear Sir,</div>
        <div>We thank you for your business renting the below vehicle.</div>
      </div>
      
      <table style={{ width: '100%', borderCollapse: 'collapse', marginBottom: 20, border: '1px solid #000' }}>
        <thead>
          <tr style={{ backgroundColor: '#f0f0f0' }}>
            <th style={{ border: '1px solid #000', padding: 8, textAlign: 'center', width: '8%' }}>No.</th>
            <th style={{ border: '1px solid #000', padding: 8, textAlign: 'center', width: '54%' }}>Description</th>
            <th style={{ border: '1px solid #000', padding: 8, textAlign: 'center', width: '19%', backgroundColor: '#f0f0f0' }}>Duration</th>
            <th style={{ border: '1px solid #000', padding: 8, textAlign: 'center', width: '19%', backgroundColor: '#f0f0f0' }}>Total Price</th>
          </tr>
        </thead>
        <tbody>
          <tr>
            <td style={{ border: '1px solid #000', padding: 8, textAlign: 'center' }}>1</td>
            <td style={{ border: '1px solid #000', padding: 8 }}>
              <div>Name: {row['Customer Name:'] || row['Customer Name'] || row['Customer'] || ''}</div>
              <div>Booking ID: {row['Dealer_Booking_Number'] || ''}</div>
              <div>R/A: {row['Contract'] || ''}</div>
              <div>Vehicle: {row['Model'] || ''} - {row['Plate_Number'] || ''}</div>
              <div>Entry: {formattedDate} - {row['Time_In'] || ''}</div>
              <div>Exit: {exitDate} - {row['Time_Out'] || ''}</div>
            </td>
            <td style={{ border: '1px solid #000', padding: 8, textAlign: 'center' }}>{(row['Time'] || '0').toString().replace(/hrs?/gi, '').trim()} hours</td>
            <td style={{ border: '1px solid #000', padding: 8, textAlign: 'right' }}>{parseFloat(row['Amount'] || '0').toFixed(2)}</td>
          </tr>
          <tr>
            <td colSpan={2} style={{ border: '1px solid #000', borderTop: 'none', padding: 8 }}></td>
            <td style={{ border: '1px solid #000', padding: 8, textAlign: 'center', backgroundColor: '#d9d9d9', fontWeight: 'bold', fontSize: 15 }}>TOTAL:</td>
            <td style={{ border: '1px solid #000', padding: 8, textAlign: 'right', backgroundColor: '#d9d9d9', fontWeight: 'bold', fontSize: 15 }}>AED {parseFloat(row['Amount'] || '0').toFixed(2)}</td>
          </tr>
        </tbody>
      </table>
      
      <div style={{ marginBottom: 20 }}>
        <div style={{ fontWeight: 'bold', textDecoration: 'underline' }}>General Conditions:</div>
      </div>
      
      <div style={{ marginBottom: 20 }}>
        <div>Terms of Payment: within 7 days</div>
      </div>
      
      <div style={{ marginTop: 40 }}>
        <div>Thanking you and assuring you of our best co-operation and services at all times.</div>
      </div>
      
      <div style={{ marginTop: 40 }}>
        <div>Best Regards,</div>
        <div style={{ marginTop: 20, fontWeight: 'bold' }}>Saudian Alwefaq Rent A Car</div>
      </div>
    </div>
  );
};


// دالة لتحميل تمبليت Excel بالأعمدة المطلوبة
const downloadExcelTemplate = () => {
  const headers = [
    [
      'Tax_Invoice_No',
      'Customer Name:',
      'Dealer_Booking_Number',
      'Contract',
      'Model',
      'Plate_Number',
      'Date',
      'Time_In',
      'Time_Out',
      'Time',
      'Amount',
    ],
  ];
  const worksheet = XLSX.utils.aoa_to_sheet(headers);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Template');
  const wbout = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
  const blob = new Blob([wbout], { type: 'application/octet-stream' });
  saveAs(blob, 'Parking-Invoice-Template.xlsx');
};


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
  const [showPreview, setShowPreview] = useState(false);
  const [previewData, setPreviewData] = useState<any>(null);


  const handleExcelUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      setExcelFile(e.target.files[0]);
    }
  };

  const handlePreview = async () => {
    if (!excelFile) {
      setStatus('Please upload an Excel or CSV file first.');
      return;
    }

    let invoiceDate = selectedDate;
    if (!invoiceDate) {
      const today = new Date();
      invoiceDate = today.toISOString().split('T')[0];
      setSelectedDate(invoiceDate);
    }

    if (!trnNumber) {
      setStatus('Please enter the TRN number first.');
      return;
    }

    setStatus('Loading preview...');

    let rows: any[] = [];
    
    if (excelFile.name.toLowerCase().endsWith('.csv')) {
      const text = await excelFile.text();
      const workbook = XLSX.read(text, { type: 'string' });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      rows = XLSX.utils.sheet_to_json(sheet);
    } else {
      const data = await excelFile.arrayBuffer();
      const workbook = XLSX.read(data, { type: 'array' });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      rows = XLSX.utils.sheet_to_json(sheet);
    }

    if (rows.length > 0) {
      setPreviewData({ row: rows[0], invoiceDate, trnNumber });
      setShowPreview(true);
      setStatus('');
    } else {
      setStatus('No data found in file.');
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
            <button
              type="button"
              onClick={downloadExcelTemplate}
              style={{
                marginTop: 10,
                background: 'linear-gradient(90deg, #43cea2 0%, #185a9d 100%)',
                color: '#fff',
                padding: '10px 0',
                fontSize: 18,
                fontWeight: 600,
                border: 'none',
                borderRadius: 8,
                cursor: 'pointer',
                boxShadow: '0 2px 8px rgba(24,90,157,0.10)'
              }}
            >
              Download Excel Template
            </button>
        </div>
        <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
          <label htmlFor="date-input" style={{ fontSize: 18, fontWeight: 500, color: '#333' }}>Select date:</label>
          <input id="date-input" type="date" value={selectedDate} onChange={(e) => setSelectedDate(e.target.value)} style={{ fontSize: 18, padding: '8px 12px', borderRadius: 8, border: '1px solid #ccc', background: '#fff' }} />
        </div>



        <div style={{ display: 'flex', gap: 12, marginTop: 16 }}>
          <button onClick={handlePreview} style={{ background: 'linear-gradient(90deg, #ff9800 0%, #f57c00 100%)', color: '#fff', padding: '16px 24px', fontSize: 18, fontWeight: 600, border: 'none', borderRadius: 12, cursor: 'pointer', flex: 1, boxShadow: '0 2px 8px rgba(255,152,0,0.10)' }}>Preview Invoice</button>
          <button onClick={handleConvert} style={{ background: 'linear-gradient(90deg, #6a1b9a 0%, #8e24aa 100%)', color: '#fff', padding: '16px 24px', fontSize: 18, fontWeight: 600, border: 'none', borderRadius: 12, cursor: 'pointer', flex: 1, boxShadow: '0 2px 8px rgba(106,27,154,0.10)' }}>Generate Invoices</button>
        </div>
        {status && <div style={{ marginTop: 18, color: '#b71c1c', fontWeight: 'bold', fontSize: 20, textAlign: 'center' }}>{status}</div>}
      </div>
      
      {showPreview && previewData && (
        <div style={{ position: 'fixed', top: 0, left: 0, right: 0, bottom: 0, background: 'rgba(0,0,0,0.8)', zIndex: 1000, display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 20 }}>
          <div style={{ background: '#fff', borderRadius: 12, maxWidth: '90%', maxHeight: '90%', overflow: 'auto', position: 'relative' }}>
            <button onClick={() => setShowPreview(false)} style={{ position: 'absolute', top: 10, right: 15, background: '#f44336', color: '#fff', border: 'none', borderRadius: '50%', width: 30, height: 30, cursor: 'pointer', fontSize: 16, zIndex: 1001 }}>×</button>
            <InvoicePreview data={previewData} />
          </div>
        </div>
      )}
    </div>
  );
}

export default ExcelToWord;