import React, { useState } from 'react';
import * as XLSX from 'xlsx';
import { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, ISectionOptions } from 'docx';
import { saveAs } from 'file-saver';

// Component to preview Traffic Fine invoice
const TrafficInvoicePreview = ({ data }: { data: any }) => {
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

  const invoiceNumber = row['Tax_Invoice_No'] ? row['Tax_Invoice_No'] : '';
  const formattedDate = formatExcelDate(row['Date']);

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
        <div>Traffic fine details given below;</div>
      </div>
      
      <table style={{ width: '100%', borderCollapse: 'collapse', marginBottom: 20, border: '1px solid #000' }}>
        <thead>
          <tr>
            <th style={{ border: '1px solid #000', padding: 8, textAlign: 'center', width: '8%' }}>No.</th>
            <th style={{ border: '1px solid #000', padding: 8, textAlign: 'center', width: '73%' }}>Description</th>
            <th style={{ border: '1px solid #000', padding: 8, textAlign: 'center', width: '19%' }}>Total Price</th>
          </tr>
        </thead>
        <tbody>
          <tr>
            <td style={{ border: '1px solid #000', padding: 8, textAlign: 'center' }}>1</td>
            <td style={{ border: '1px solid #000', padding: 8 }}>
              <div style={{ fontWeight: 'bold' }}>{row['Customer'] || row['Customer_Name'] || 'N/A'}</div>
              <div>Traffic Fine No: {row['TFINE No.'] || ''}</div>
              <div>Date: {formattedDate}</div>
              <div>Time: {row['Time'] || ''}</div>
              <div>Booking ID: {row['Booking_ID'] || row['Dealer_Booking_Number'] || ''}</div>
              <div>R/A: {row['Dealer_Booking_Number'] || ''}</div>
              <div>Plate No. {row['Plate_Number'] || ''}</div>
              <div>Description: {row['Description'] || ''}</div>
            </td>
            <td style={{ border: '1px solid #000', padding: 8, textAlign: 'right' }}>{parseFloat(row['Amount'] || '0').toFixed(2)} AED</td>
          </tr>
          <tr>
            <td style={{ border: '1px solid #000', borderTop: 'none', padding: 8 }}></td>
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

// Function to download CSV template with required columns for traffic fines
const downloadCSVTemplate = () => {
  const headers = 'TFINE No.,Plate_Number,Date,Time,Amount,Description,Dealer_Booking_Number,Booking_ID,Invoice_Date,Tax_Invoice_No,Customer';
  const blob = new Blob([headers], { type: 'text/csv' });
  saveAs(blob, 'Traffic-Fines-Template.csv');
};




// Helper function to create a traffic fine invoice section
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
  const formattedDate = formatExcelDate(row['Date']);

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
        new Paragraph({ children: [new TextRun({ text: 'SUB: Micro Lease Cars ', ...fontProps })] }),
        new Paragraph({ children: [new TextRun({ text: '', ...fontProps })] }),
        new Paragraph({ children: [new TextRun({ text: 'Traffic fine details given below;.', ...fontProps })] }),
        new Paragraph({ children: [new TextRun({ text: '', ...fontProps })] }),
        // Invoice Table
        new Table({
          borders: {
            top: { style: 'none' },
            bottom: { style: 'none' },
            left: { style: 'none' },
            right: { style: 'none' },
            insideHorizontal: { style: 'none' },
            insideVertical: { style: 'none' }
          },
          rows: [
            new TableRow({
              height: { value: 800, rule: 'exact' },
              children: [
                new TableCell({ 
                  children: [new Paragraph({ children: [new TextRun({ text: 'No.', ...fontProps, bold: true })], alignment: 'center' })], 
                  width: { size: 1000, type: 'dxa' }, 
                  verticalAlign: 'center',
                  borders: {
                    top: { style: 'single', size: 1 },
                    bottom: { style: 'single' },
                    left: { style: 'single', size: 1 },
                    right: { style: 'single', size: 1 }
                  }
                }),
                new TableCell({ 
                  children: [new Paragraph({ children: [new TextRun({ text: 'Description', ...fontProps, bold: true })], alignment: 'center' })], 
                  width: { size: 7000, type: 'dxa' }, 
                  verticalAlign: 'center',
                  borders: {
                    top: { style: 'single', size: 1 },
                    bottom: { style: 'single', size: 1 },
                    left: { style: 'single', size: 1 },
                    right: { style: 'single', size: 1 }
                  }
                }),
                new TableCell({ 
                  children: [new Paragraph({ children: [new TextRun({ text: 'Total Price', ...fontProps, bold: true })], alignment: 'center' })], 
                  width: { size: 2000, type: 'dxa' }, 
                  verticalAlign: 'center',
                  borders: {
                    top: { style: 'single', size: 1 },
                    bottom: { style: 'single', size: 1 },
                    left: { style: 'single', size: 1 },
                    right: { style: 'single', size: 1 }
                  }
                }),
              ],
            }),
            new TableRow({
              children: [
                new TableCell({ 
                  children: [new Paragraph({ children: [new TextRun({ text: '1', ...fontProps })], alignment: 'center' })], 
                  width: { size: 1000, type: 'dxa' }, 
                  verticalAlign: 'center',
                  borders: {
                    top: { style: 'none' },
                    bottom: { style: 'none' },
                    left: { style: 'single', size: 1 },
                    right: { style: 'single', size: 1 }
                  }
                }),
                new TableCell({
                  children: [
                    new Paragraph({ children: [new TextRun({ text: `${row['Customer'] || row['Customer_Name'] || 'N/A'}`, ...fontProps, bold: true })] }),
                    new Paragraph({ children: [new TextRun({ text: `Traffic Fine No: ${row['TFINE No.'] || ''}`, ...fontProps })] }),
                    new Paragraph({ children: [new TextRun({ text: `Date: ${formattedDate}`, ...fontProps })] }),
                    new Paragraph({ children: [new TextRun({ text: `Time: ${row['Time'] || ''}`, ...fontProps })] }),
                    new Paragraph({ children: [new TextRun({ text: `Booking ID: ${row['Booking_ID'] || row['Dealer_Booking_Number'] || ''}`, ...fontProps })] }),
                    new Paragraph({ children: [new TextRun({ text: `R/A: ${row['Dealer_Booking_Number'] || ''}`, ...fontProps })] }),
                    new Paragraph({ children: [new TextRun({ text: `Plate No. ${row['Plate_Number'] || ''}`, ...fontProps })] }),
                    new Paragraph({ children: [new TextRun({ text: `Description: ${row['Description'] || ''}`, ...fontProps })] }),
                    new Paragraph({ children: [new TextRun({ text: '', ...fontProps })] }),
                  ],
                  width: { size: 7000, type: 'dxa' },
                  borders: {
                    top: { style: 'single', size: 1 },
                    bottom: { style: 'single', size: 1 },
                    left: { style: 'single', size: 1 },
                    right: { style: 'single', size: 1 }
                  }
                }),
                new TableCell({ 
                  children: [new Paragraph({ children: [new TextRun({ text: `${parseFloat(row['Amount'] || '0').toFixed(2)} AED`, ...fontProps })], alignment: 'right' })], 
                  width: { size: 2000, type: 'dxa' }, 
                  verticalAlign: 'center', 
                  margins: { right: 144 },
                  borders: {
                    top: { style: 'single', size: 1 },
                    bottom: { style: 'single', size: 1 },
                    left: { style: 'single', size: 1 },
                    right: { style: 'single', size: 1 }
                  }
                }),
              ],
            }),
            new TableRow({
              children: [
                new TableCell({
                  children: [new Paragraph({ children: [new TextRun({ text: '', ...fontProps })] })],
                  borders: {
                    top: { style: 'none' },
                    bottom: { style: 'single', size: 1 },
                    left: { style: 'single', size: 1 },
                    right: { style: 'single', size: 1 }
                  },
                }),
                new TableCell({
                  children: [new Paragraph({ children: [new TextRun({ text: 'TOTAL:', ...fontProps, bold: true, size: 30 })], alignment: 'center' })],
                  shading: { fill: 'D9D9D9' },
                  verticalAlign: 'center',
                  borders: {
                    top: { style: 'single', size: 1 },
                    bottom: { style: 'single', size: 1 },
                    left: { style: 'single', size: 1 },
                    right: { style: 'single', size: 1 }
                  }
                }),
                new TableCell({ 
                  children: [new Paragraph({ children: [new TextRun({ text: `AED ${parseFloat(row['Amount'] || '0').toFixed(2)} `, ...fontProps, bold: true, size: 30 })], alignment: 'right' })], 
                  shading: { fill: 'D9D9D9' }, 
                  margins: { right: 144 },
                  borders: {
                    top: { style: 'single', size: 1 },
                    bottom: { style: 'single', size: 1 },
                    left: { style: 'single', size: 1 },
                    right: { style: 'single', size: 1 }
                  }
                }),
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

function TrafficFines() {
  const [excelFile, setExcelFile] = useState<File | null>(null);
  const [status, setStatus] = useState('');
  const [selectedDate, setSelectedDate] = useState('');
  const [trnNumber, setTrnNumber] = useState('100397403500003');
  const [uploadedData, setUploadedData] = useState<any[]>([]);
  const [showPreview, setShowPreview] = useState(false);
  const [previewData, setPreviewData] = useState<any>(null);


  const handleExcelUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      setExcelFile(e.target.files[0]);
      
      // Read and store the uploaded data
      const file = e.target.files[0];
      const text = await file.text();
      const workbook = XLSX.read(text, { type: 'string' });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet);
      setUploadedData(rows);
    }
  };

  const generateCSVMirror = () => {
    if (uploadedData.length === 0) {
      setStatus('Please upload a CSV file first.');
      return;
    }

    // Convert data back to CSV format
    const headers = 'TFINE No.,Plate_Number,Date,Time,Amount,Description,Dealer_Booking_Number,Booking_ID,Invoice_Date,Tax_Invoice_No,Customer';
    const csvRows = uploadedData.map(row => 
      `${row['TFINE No.'] || ''},${row['Plate_Number'] || ''},${row['Date'] || ''},${row['Time'] || ''},${row['Amount'] || ''},"${row['Description'] || ''}",${row['Dealer_Booking_Number'] || ''},${row['Booking_ID'] || ''},${row['Invoice_Date'] || ''},${row['Tax_Invoice_No'] || ''},"${row['Customer'] || ''}"`
    );
    
    const csvContent = headers + '\n' + csvRows.join('\n');
    const blob = new Blob([csvContent], { type: 'text/csv' });
    saveAs(blob, 'Invygo Upload.csv');
    setStatus('Invygo Upload generated successfully!');
  };

  const handlePreview = async () => {
    if (!excelFile) {
      setStatus('Please upload a CSV file first.');
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

    const text = await excelFile.text();
    const workbook = XLSX.read(text, { type: 'string' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet);

    if (rows.length > 0) {
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

      let finalInvoiceDate = invoiceDate;
      if (rows[0]['Invoice_Date']) {
        const csvInvoiceDate = formatExcelDate(rows[0]['Invoice_Date']);
        if (csvInvoiceDate) {
          finalInvoiceDate = csvInvoiceDate;
        }
      } else {
        finalInvoiceDate = new Date(invoiceDate + 'T00:00:00').toLocaleDateString('en-GB');
      }

      setPreviewData({ row: rows[0], invoiceDate: finalInvoiceDate, trnNumber });
      setShowPreview(true);
      setStatus('');
    } else {
      setStatus('No data found in file.');
    }
  };

  const handleConvert = async () => {
    if (!excelFile) {
      setStatus('Please upload a CSV file first.');
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

    // Helper function to format date (same as in createInvoiceSection)
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

    let rows: any[] = [];
    
    // Handle CSV files
    const text = await excelFile.text();
    const workbook = XLSX.read(text, { type: 'string' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    rows = XLSX.utils.sheet_to_json(sheet);

    const sections = rows.map((row) => {
      // Use Invoice_Date from CSV if available, otherwise use selected date
      let finalInvoiceDate = invoiceDate;
      
      if (row['Invoice_Date']) {
        // If Invoice_Date exists in CSV, use it
        const csvInvoiceDate = formatExcelDate(row['Invoice_Date']);
        if (csvInvoiceDate) {
          return createInvoiceSection(row, csvInvoiceDate, trnNumber);
        }
      }
      
      // Otherwise use selected date, converted to DD/MM/YYYY format
      const formattedInvoiceDate = finalInvoiceDate ? 
        new Date(finalInvoiceDate + 'T00:00:00').toLocaleDateString('en-GB') : 
        new Date().toLocaleDateString('en-GB');
      
      return createInvoiceSection(row, formattedInvoiceDate, trnNumber);
    });

    const doc = new Document({ sections });
    const buffer = await Packer.toBlob(doc);
    saveAs(buffer, 'traffic-fines-invoices.docx');
    setStatus('Success!');
  };

  return (
    <div style={{ maxWidth: 600, margin: '40px auto', padding: 40, background: 'linear-gradient(135deg, #f3e7ff 0%, #fffbe7 100%)', borderRadius: 24, boxShadow: '0 8px 32px rgba(106,27,154,0.12)', fontFamily: 'Segoe UI, Arial, sans-serif' }}>
      <h2 style={{ color: '#6a1b9a', fontSize: 36, fontWeight: 700, textAlign: 'center', marginBottom: 32, letterSpacing: 1 }}>Traffic Fines Invoice Generator</h2>
      <div style={{ display: 'flex', flexDirection: 'column', gap: 24 }}>
        <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
          <label htmlFor="csv-upload" style={{ fontSize: 18, fontWeight: 500, color: '#333' }}>Upload Traffic Fines CSV file:</label>
          <input id="csv-upload" type="file" accept=".csv" onChange={handleExcelUpload} style={{ fontSize: 18, padding: '8px 12px', borderRadius: 8, border: '1px solid #ccc', background: '#fff' }} />
            <button
              type="button"
              onClick={downloadCSVTemplate}
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
              Download Empty Template
            </button>
        </div>
        <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
          <label htmlFor="date-input" style={{ fontSize: 18, fontWeight: 500, color: '#333' }}>Select date:</label>
          <input id="date-input" type="date" value={selectedDate} onChange={(e) => setSelectedDate(e.target.value)} style={{ fontSize: 18, padding: '8px 12px', borderRadius: 8, border: '1px solid #ccc', background: '#fff' }} />
        </div>



        <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
          <label htmlFor="trn-input" style={{ fontSize: 18, fontWeight: 500, color: '#333' }}>TRN Number:</label>
          <input id="trn-input" type="text" value={trnNumber} onChange={(e) => setTrnNumber(e.target.value)} style={{ fontSize: 18, padding: '8px 12px', borderRadius: 8, border: '1px solid #ccc', background: '#fff' }} />
        </div>

        <div style={{ display: 'flex', gap: 12, marginTop: 16 }}>
          <button onClick={handlePreview} style={{ background: 'linear-gradient(90deg, #ff9800 0%, #f57c00 100%)', color: '#fff', padding: '16px 12px', fontSize: 16, fontWeight: 600, border: 'none', borderRadius: 12, cursor: 'pointer', flex: 1, boxShadow: '0 2px 8px rgba(255,152,0,0.10)' }}>Preview Invoice</button>
          <button onClick={handleConvert} style={{ background: 'linear-gradient(90deg, #6a1b9a 0%, #8e24aa 100%)', color: '#fff', padding: '16px 12px', fontSize: 16, fontWeight: 600, border: 'none', borderRadius: 12, cursor: 'pointer', boxShadow: '0 2px 8px rgba(106,27,154,0.10)', flex: 1 }}>Generate Word Invoices</button>
          <button onClick={generateCSVMirror} style={{ background: 'linear-gradient(90deg, #e17055 0%, #d63031 100%)', color: '#fff', padding: '16px 12px', fontSize: 16, fontWeight: 600, border: 'none', borderRadius: 12, cursor: 'pointer', boxShadow: '0 2px 8px rgba(214,48,49,0.10)', flex: 1 }}>Generate Invygo Upload</button>
        </div>
        {status && <div style={{ marginTop: 18, color: '#b71c1c', fontWeight: 'bold', fontSize: 20, textAlign: 'center' }}>{status}</div>}
      </div>
      
      {showPreview && previewData && (
        <div style={{ position: 'fixed', top: 0, left: 0, right: 0, bottom: 0, background: 'rgba(0,0,0,0.8)', zIndex: 1000, display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 20 }}>
          <div style={{ background: '#fff', borderRadius: 12, maxWidth: '90%', maxHeight: '90%', overflow: 'auto', position: 'relative' }}>
            <button onClick={() => setShowPreview(false)} style={{ position: 'absolute', top: 10, right: 15, background: '#f44336', color: '#fff', border: 'none', borderRadius: '50%', width: 30, height: 30, cursor: 'pointer', fontSize: 16, zIndex: 1001 }}>×</button>
            <TrafficInvoicePreview data={previewData} />
          </div>
        </div>
      )}
    </div>
  );
}

export default TrafficFines;