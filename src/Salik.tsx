import React, { useState } from 'react';
import * as XLSX from 'xlsx';
import { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, ISectionOptions } from 'docx';
import { saveAs } from 'file-saver';

// Component to preview Salik invoice
const SalikInvoicePreview = ({ data }: { data: any }) => {
  const { row, invoiceDate, trnNumber } = data;
  
  const formatExcelDate = (excelDate: any): string => {
    if (!excelDate) return '';
    if (typeof excelDate === 'number') {
      const date = new Date((excelDate - 25569) * 86400 * 1000);
      return date.toLocaleDateString('en-GB');
    }
    if (typeof excelDate === 'string') {
      const parsedDate = new Date(excelDate);
      if (!isNaN(parsedDate.getTime())) {
        return parsedDate.toLocaleDateString('en-GB');
      }
      return excelDate;
    }
    return '';
  };

  const finalInvoiceDate = row['Invoice_Date'] ? formatExcelDate(row['Invoice_Date']) : invoiceDate;
  const invoiceNumber = row['INVOICE'] ? row['INVOICE'] : '';
  
  let salikDateText = '';
  if (row['Month'] && row['Month'].toString().trim() !== '') {
    salikDateText = `Salik Month: ${row['Month']}`;
  } else {
    const startDate = formatExcelDate(row['Date']);
    const endDate = formatExcelDate(row['End Date']);
    if (startDate) {
      if (endDate && endDate !== '') {
        salikDateText = `Salik Date: ${startDate} - ${endDate}`;
      } else {
        salikDateText = `Salik Date: ${startDate}`;
      }
    }
  }

  const hasSalikTrips = row['Salik Trips'] && row['Salik Trips'].toString().trim() !== '';
  const formatPrice = (price: any): string => {
    const numPrice = parseFloat(price) || 0;
    return numPrice.toFixed(2);
  };

  return (
    <div style={{ padding: 40, fontFamily: 'Calibri, Arial, sans-serif', fontSize: 11, lineHeight: 1.4, color: '#000', background: '#fff', minWidth: 600 }}>
      <div style={{ textAlign: 'center', marginBottom: 30 }}>
        <h1 style={{ fontSize: 26, fontWeight: 'bold', margin: 0, fontFamily: 'Arial' }}>Tax Invoice</h1>
      </div>
      
      <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: 20 }}>
        <span>Date: {finalInvoiceDate}</span>
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
            <th style={{ border: '1px solid #000', padding: 8, textAlign: 'center', width: hasSalikTrips ? '54%' : '73%' }}>Description</th>
            {hasSalikTrips && <th style={{ border: '1px solid #000', padding: 8, textAlign: 'center', width: '19%', backgroundColor: '#f0f0f0' }}>Salik Trips</th>}
            <th style={{ border: '1px solid #000', padding: 8, textAlign: 'center', width: '19%', backgroundColor: '#f0f0f0' }}>Total Price</th>
          </tr>
        </thead>
        <tbody>
          <tr>
            <td style={{ border: '1px solid #000', padding: 8, textAlign: 'center' }}>1</td>
            <td style={{ border: '1px solid #000', padding: 8 }}>
              <div>Name: {row['Customer'] || ''}</div>
              <div>Booking ID: {row['Booking Number'] || ''}</div>
              <div>R/A: {row['Contract No.'] || ''}</div>
              <div>Vehicle: {row['Model'] || ''} - {row['Plate No.'] || ''}</div>
              <div>{salikDateText}</div>
            </td>
            {hasSalikTrips && <td style={{ border: '1px solid #000', padding: 8, textAlign: 'center' }}>{row['Salik Trips']} Trips</td>}
            <td style={{ border: '1px solid #000', padding: 8, textAlign: 'right' }}>{formatPrice(row['Total Price'])}</td>
          </tr>
          <tr>
            <td colSpan={hasSalikTrips ? 2 : 2} style={{ border: '1px solid #000', borderTop: 'none', padding: 8 }}></td>
            {hasSalikTrips && <td style={{ border: '1px solid #000', padding: 8, textAlign: 'center', backgroundColor: '#d9d9d9', fontWeight: 'bold', fontSize: 15 }}>TOTAL:</td>}
            <td style={{ border: '1px solid #000', padding: 8, textAlign: 'right', backgroundColor: '#d9d9d9', fontWeight: 'bold', fontSize: 15 }}>AED {formatPrice(row['Total Price'])}</td>
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

// دالة لتحميل تمبليت Excel بالأعمدة المطلوبة لهذا الكود
const downloadExcelTemplate = () => {
  const headers = [
    [
      'INVOICE',
      'Customer',
      'Booking Number',
      'Contract No.',
      'Model',
      'Plate No.',
      'Date',
      'End Date',
      'Month',
      'Salik Trips',
      'Total Price',
      'Invoice_Date',
    ],
  ];
  const worksheet = XLSX.utils.aoa_to_sheet(headers);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Template');
  const wbout = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
  const blob = new Blob([wbout], { type: 'application/octet-stream' });
  saveAs(blob, 'Salik-Invoice-Template.xlsx');
};


// Helper function to format price with two decimal places
const formatPrice = (price: any): string => {
  const numPrice = parseFloat(price) || 0;
  return numPrice.toFixed(2);
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
    // If the date is already a string, return as is
    if (typeof excelDate === 'string') {
      // Try to parse as date and format, if fails return original string
      const parsedDate = new Date(excelDate);
      if (!isNaN(parsedDate.getTime())) {
        return parsedDate.toLocaleDateString('en-GB');
      }
      return excelDate;
    }
    return '';
  };

  // Use Invoice_Date from Excel if available, otherwise use the passed invoiceDate
  const finalInvoiceDate = row['Invoice_Date'] ? formatExcelDate(row['Invoice_Date']) : invoiceDate;

  // Determine the Salik Date text - Month has priority over Date range
  let salikDateText = '';
  
  // Check for Month column - Month has priority over Date range
  if (row['Month'] && row['Month'].toString().trim() !== '') {
    salikDateText = ` Salik Month: ${row['Month']}`;
  } else {
    // Fallback to Date columns if available
    const startDate = formatExcelDate(row['Date']);
    const endDate = formatExcelDate(row['End Date']);
    if (startDate) {
      if (endDate && endDate !== '') {
        salikDateText = ` Salik Date: ${startDate} - ${endDate}`;
      } else {
        salikDateText = ` Salik Date: ${startDate}`;
      }
    }
  }

  // Check if Salik Trips column has data
  const hasSalikTrips = row['Salik Trips'] && row['Salik Trips'].toString().trim() !== '';

  const invoiceNumber = row['INVOICE'] ? row['INVOICE'] : '';

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
            new TextRun({ text: `Date: ${finalInvoiceDate}`, ...fontProps }),
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
        // Invoice Table - Dynamic based on available data
        new Table({
          rows: [
            // Header Row
            new TableRow({
              height: { value: 800, rule: 'exact' },
              children: [
                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: 'No.', ...fontProps, bold: true })], alignment: 'center' })], width: { size: 1000, type: 'dxa' }, verticalAlign: 'center' }),
                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: 'Description', ...fontProps, bold: true })], alignment: 'center' })], width: { size: hasSalikTrips ? 6000 : 8000, type: 'dxa' }, verticalAlign: 'center' }),
                ...(hasSalikTrips ? [new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: 'Salik Trips', ...fontProps, bold: true })], alignment: 'center' })], width: { size: 2000, type: 'dxa' }, shading: { fill: 'F0F0F0' }, verticalAlign: 'center' })] : []),
                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: 'Total Price', ...fontProps, bold: true })], alignment: 'center' })], width: { size: 2000, type: 'dxa' }, shading: { fill: 'F0F0F0' }, verticalAlign: 'center' }),
              ],
            }),
            // Data Row
            new TableRow({
              children: [
                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: '1', ...fontProps })], alignment: 'center' })], width: { size: 1000, type: 'dxa' }, verticalAlign: 'center' }),
                new TableCell({
                  children: [
                    new Paragraph({ children: [new TextRun({ text: ` Name: ${row['Customer'] || ''}`, ...fontProps })] }),
                    new Paragraph({ children: [new TextRun({ text: ` Booking ID: ${row['Booking Number'] || ''}`, ...fontProps })] }),
                    new Paragraph({ children: [new TextRun({ text: ` R/A: ${row['Contract No.'] || ''}`, ...fontProps })] }),
                    new Paragraph({ children: [new TextRun({ text: ` Vehicle: ${row['Model'] || ''} - ${row['Plate No.'] || ''}`, ...fontProps })] }),
                    new Paragraph({ children: [new TextRun({ text: salikDateText, ...fontProps })] }),
                  ],
                  width: { size: hasSalikTrips ? 6000 : 8000, type: 'dxa' },
                }),
                ...(hasSalikTrips ? [new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: `${row['Salik Trips']} Trips`, ...fontProps })], alignment: 'center' })], width: { size: 2000, type: 'dxa' }, verticalAlign: 'center' })] : []),
                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: `${formatPrice(row['Total Price'])}`, ...fontProps })], alignment: 'right' })], width: { size: 2000, type: 'dxa' }, verticalAlign: 'center', margins: { right: 144 } }),
              ],
            }),
            // Total Row
            new TableRow({
              children: [
                new TableCell({
                  children: [new Paragraph({ children: [new TextRun({ text: '', ...fontProps })] })],
                  columnSpan: hasSalikTrips ? 2 : 2,
                  borders: {
                    top: { style: 'single' },
                    left: { style: 'none', size: 4, color: '000000' },
                    right: { style: 'none', size: 4, color: '000000' },
                    bottom: { style: 'none', size: 4, color: '000000' },
                  },
                }),
                ...(hasSalikTrips ? [new TableCell({
                  children: [
                    new Paragraph({ children: [new TextRun({ text: 'TOTAL:', ...fontProps, bold: true, size: 30 })], alignment: 'center' }),
                  ],
                  shading: { fill: 'D9D9D9' },
                  verticalAlign: 'center',
                })] : []),
                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: `AED ${formatPrice(row['Total Price'])}`, ...fontProps, bold: true, size: 30 })], alignment: 'right' })], shading: { fill: 'D9D9D9' }, margins: { right: 144 } }),
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
      setStatus('Please upload an Excel file first.');
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

    const data = await excelFile.arrayBuffer();
    const workbook = XLSX.read(data, { type: 'array' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows: any[] = XLSX.utils.sheet_to_json(sheet);

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
      setStatus('Please upload an Excel file first.');
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

    const data = await excelFile.arrayBuffer();
    const workbook = XLSX.read(data, { type: 'array' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows: any[] = XLSX.utils.sheet_to_json(sheet);

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
      <h2 style={{ color: '#6a1b9a', fontSize: 36, fontWeight: 700, textAlign: 'center', marginBottom: 32, letterSpacing: 1 }}>Salik Invoice Creation </h2>
      <div style={{ display: 'flex', flexDirection: 'column', gap: 24 }}>
        <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
          <label htmlFor="excel-upload" style={{ fontSize: 18, fontWeight: 500, color: '#333' }}>Excel file:</label>
          <input id="excel-upload" type="file" accept=".xlsx,.xls" onChange={handleExcelUpload} style={{ fontSize: 18, padding: '8px 12px', borderRadius: 8, border: '1px solid #ccc', background: '#fff' }} />
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
            <SalikInvoicePreview data={previewData} />
          </div>
        </div>
      )}
    </div>
  );
}

export default ExcelToWord;