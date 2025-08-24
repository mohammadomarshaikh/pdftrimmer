const express = require('express');
const multer = require('multer');
const { PDFDocument } = require('pdf-lib');

const app = express();
const port = 3005;

// Configure multer to handle file uploads
const upload = multer({
  storage: multer.memoryStorage(),
  limits: {
    fileSize: 10 * 1024 * 1024 // 10MB limit
  }
});

// Configure middleware for parsing form data and JSON
app.use(express.urlencoded({ extended: true }));
app.use(express.json({ limit: '15mb' })); // Increased limit for base64 data

// Configure CORS headers for testing
app.use((req, res, next) => {
  res.header('Access-Control-Allow-Origin', '*');
  res.header('Access-Control-Allow-Methods', 'GET, POST, PUT, DELETE, OPTIONS');
  res.header('Access-Control-Allow-Headers', 'Origin, X-Requested-With, Content-Type, Accept, Authorization');
  next();
});

// Health check endpoint
app.get('/health', (req, res) => {
  res.json({ status: 'Server is running', port: port });
});

// Helper function to process PDF trimming
async function processPdfTrimming(pdfBuffer, pageCount) {
  // Load the original PDF
  const originalPdf = await PDFDocument.load(pdfBuffer);
  const totalPages = originalPdf.getPageCount();

  console.log(`Original PDF has ${totalPages} pages`);

  // Validate page count against total pages
  if (pageCount > totalPages) {
    throw new Error(`Cannot trim to ${pageCount} pages. PDF only has ${totalPages} pages.`);
  }

  // Create new PDF with trimmed pages
  const newPdf = await PDFDocument.create();
  const pagesToCopy = Math.min(pageCount, totalPages);

  // Copy the first N pages
  const copiedPages = await newPdf.copyPages(originalPdf, [...Array(pagesToCopy).keys()]);
  copiedPages.forEach(page => newPdf.addPage(page));

  // Generate the new PDF
  const newPdfBytes = await newPdf.save();

  console.log(`Successfully trimmed PDF to ${pagesToCopy} pages`);
  return newPdfBytes;
}

// JSON endpoint for Power Automate (base64)
app.post('/api/trim-pdf', async (req, res) => {
  try {
    console.log('JSON API request received');
    console.log('Body keys:', Object.keys(req.body));

    // Validate required fields
    if (!req.body.pdfData) {
      return res.status(400).json({
        error: 'Missing required field: pdfData (base64 encoded PDF content)'
      });
    }

    if (!req.body.pages) {
      return res.status(400).json({
        error: 'Missing required field: pages (number of pages to keep)'
      });
    }

    // Get and validate page count
    const pageCount = parseInt(req.body.pages);
    if (isNaN(pageCount) || pageCount < 1) {
      return res.status(400).json({
        error: 'Please specify a valid number of pages to trim (must be a positive integer)'
      });
    }

    // Decode base64 PDF data
    let pdfBuffer;
    try {
      // Remove data URL prefix if present
      const base64Data = req.body.pdfData.replace(/^data:application\/pdf;base64,/, '');
      pdfBuffer = Buffer.from(base64Data, 'base64');
    } catch (err) {
      return res.status(400).json({
        error: 'Invalid base64 PDF data. Please ensure the PDF is properly encoded.'
      });
    }

    console.log(`Processing PDF with ${pdfBuffer.length} bytes, trimming to ${pageCount} pages`);

    // Process the PDF
    const newPdfBytes = await processPdfTrimming(pdfBuffer, pageCount);

    // Return response based on requested format
    const returnFormat = req.body.returnFormat || 'base64';

    if (returnFormat === 'base64') {
      // Return as JSON with base64 encoded PDF
      const base64Result = Buffer.from(newPdfBytes).toString('base64');
      res.json({
        success: true,
        pdfData: base64Result,
        originalSize: pdfBuffer.length,
        trimmedSize: newPdfBytes.length,
        pagesTrimmed: pageCount
      });
    } else {
      // Return as binary PDF file
      res.set({
        'Content-Type': 'application/pdf',
        'Content-Disposition': 'attachment; filename=trimmed.pdf',
        'Content-Length': newPdfBytes.length
      });
      res.send(Buffer.from(newPdfBytes));
    }

  } catch (err) {
    console.error('Error processing PDF:', err);
    res.status(500).json({
      error: 'Failed to process PDF',
      details: err.message
    });
  }
});

// Multipart form endpoint (for backward compatibility)
app.post('/trim-pdf', upload.single('file'), async (req, res) => {
  try {
    console.log('Multipart form request received');
    console.log('File:', req.file ? 'Present' : 'Missing');
    console.log('Body:', req.body);

    // Validate file upload
    if (!req.file) {
      return res.status(400).json({
        error: 'No PDF file uploaded. Please upload a file with field name "file"'
      });
    }

    // Validate file type
    if (req.file.mimetype !== 'application/pdf') {
      return res.status(400).json({
        error: 'Invalid file type. Please upload a PDF file.'
      });
    }

    // Get and validate page count
    const pageCount = parseInt(req.body.pages);
    if (isNaN(pageCount) || pageCount < 1) {
      return res.status(400).json({
        error: 'Please specify a valid number of pages to trim (must be a positive integer)'
      });
    }

    console.log(`Processing PDF with ${req.file.size} bytes, trimming to ${pageCount} pages`);

    // Process the PDF
    const newPdfBytes = await processPdfTrimming(req.file.buffer, pageCount);

    // Set response headers
    res.set({
      'Content-Type': 'application/pdf',
      'Content-Disposition': 'attachment; filename=trimmed.pdf',
      'Content-Length': newPdfBytes.length
    });

    // Send the trimmed PDF
    res.send(Buffer.from(newPdfBytes));

  } catch (err) {
    console.error('Error processing PDF:', err);
    res.status(500).json({
      error: 'Failed to process PDF',
      details: err.message
    });
  }
});

// Endpoint to extract and return last 2 pages in base64 for Power Automate
app.post('/api/last-2-pages', async (req, res) => {
  try {
    console.log('JSON API request to extract last 2 pages');
    console.log('Body keys:', Object.keys(req.body));

    // Validate required field
    if (!req.body.pdfData) {
      return res.status(400).json({
        error: 'Missing required field: pdfData (base64 encoded PDF content)'
      });
    }

    // Decode base64 PDF data
    let pdfBuffer;
    try {
      const base64Data = req.body.pdfData.replace(/^data:application\/pdf;base64,/, '');
      pdfBuffer = Buffer.from(base64Data, 'base64');
    } catch (err) {
      return res.status(400).json({
        error: 'Invalid base64 PDF data. Please ensure the PDF is properly encoded.'
      });
    }

    // Load the original PDF
    const originalPdf = await PDFDocument.load(pdfBuffer);
    const totalPages = originalPdf.getPageCount();

    if (totalPages < 2) {
      return res.status(400).json({
        error: `PDF only has ${totalPages} page(s). Cannot extract last 2 pages.`
      });
    }

    // Create a new PDF with the last 2 pages
    const newPdf = await PDFDocument.create();
    const last2PageIndexes = [totalPages - 2, totalPages - 1];
    const copiedPages = await newPdf.copyPages(originalPdf, last2PageIndexes);
    copiedPages.forEach((page) => newPdf.addPage(page));
    const newPdfBytes = await newPdf.save();

    const base64Result = Buffer.from(newPdfBytes).toString('base64');
    res.json({
      success: true,
      pdfData: base64Result,
      originalPageCount: totalPages,
      returnedPages: 2
    });

  } catch (err) {
    console.error('Error extracting last 2 pages:', err);
    res.status(500).json({
      error: 'Failed to extract last 2 pages',
      details: err.message
    });
  }
});

// Endpoint to extract last N pages from a base64 PDF
app.post('/api/last-n-pages', async (req, res) => {
  try {
    console.log('JSON API request to extract last N pages');
    console.log('Body keys:', Object.keys(req.body));

    // Validate required fields
    const base64Input = req.body.pdfData;
    const numPages = parseInt(req.body.pages); // pages to extract from the end

    if (!base64Input) {
      return res.status(400).json({
        error: 'Missing required field: pdfData (base64 encoded PDF content)'
      });
    }

    if (!numPages || isNaN(numPages) || numPages < 1) {
      return res.status(400).json({
        error: 'Please provide a valid "pages" field (positive integer)'
      });
    }

    // Decode base64 PDF
    let pdfBuffer;
    try {
      const base64Data = base64Input.replace(/^data:application\/pdf;base64,/, '');
      pdfBuffer = Buffer.from(base64Data, 'base64');
    } catch (err) {
      return res.status(400).json({
        error: 'Invalid base64 PDF data. Please ensure the PDF is properly encoded.'
      });
    }

    // Load PDF and get page count
    const originalPdf = await PDFDocument.load(pdfBuffer);
    const totalPages = originalPdf.getPageCount();

    if (numPages > totalPages) {
      return res.status(400).json({
        error: `PDF only has ${totalPages} page(s). Cannot extract last ${numPages} pages.`
      });
    }

    // Copy last N pages
    const newPdf = await PDFDocument.create();
    const startIndex = totalPages - numPages;
    const pageIndexes = [...Array(numPages).keys()].map(i => startIndex + i);
    const copiedPages = await newPdf.copyPages(originalPdf, pageIndexes);
    copiedPages.forEach(page => newPdf.addPage(page));

    // Save and return base64
    const newPdfBytes = await newPdf.save();
    const base64Result = Buffer.from(newPdfBytes).toString('base64');

    res.json({
      success: true,
      pdfData: base64Result,
      originalPageCount: totalPages,
      returnedPages: numPages
    });

  } catch (err) {
    console.error('Error extracting last N pages:', err);
    res.status(500).json({
      error: 'Failed to extract pages',
      details: err.message
    });
  }
});

// NEW: Endpoint to merge first N and last M pages from a base64 PDF
app.post('/api/merge-first-last-pages', async (req, res) => {
  try {
    console.log('JSON API request to merge first and last pages');
    console.log('Body keys:', Object.keys(req.body));

    // Validate required fields
    const base64Input = req.body.pdfData;
    const firstPages = parseInt(req.body.firstPages);
    const lastPages = parseInt(req.body.lastPages);

    if (!base64Input) {
      return res.status(400).json({
        error: 'Missing required field: pdfData (base64 encoded PDF content)'
      });
    }

    if (!firstPages || isNaN(firstPages) || firstPages < 0) {
      return res.status(400).json({
        error: 'Please provide a valid "firstPages" field (non-negative integer)'
      });
    }

    if (!lastPages || isNaN(lastPages) || lastPages < 0) {
      return res.status(400).json({
        error: 'Please provide a valid "lastPages" field (non-negative integer)'
      });
    }

    if (firstPages === 0 && lastPages === 0) {
      return res.status(400).json({
        error: 'At least one of "firstPages" or "lastPages" must be greater than 0'
      });
    }

    // Decode base64 PDF
    let pdfBuffer;
    try {
      const base64Data = base64Input.replace(/^data:application\/pdf;base64,/, '');
      pdfBuffer = Buffer.from(base64Data, 'base64');
    } catch (err) {
      return res.status(400).json({
        error: 'Invalid base64 PDF data. Please ensure the PDF is properly encoded.'
      });
    }

    // Load PDF and get page count
    const originalPdf = await PDFDocument.load(pdfBuffer);
    const totalPages = originalPdf.getPageCount();

    // Validate page counts
    if (firstPages > totalPages) {
      return res.status(400).json({
        error: `PDF only has ${totalPages} page(s). Cannot extract ${firstPages} pages from the start.`
      });
    }

    if (lastPages > totalPages) {
      return res.status(400).json({
        error: `PDF only has ${totalPages} page(s). Cannot extract ${lastPages} pages from the end.`
      });
    }

    // Check for potential overlap (if firstPages + lastPages > totalPages)
    const maxNonOverlappingFirstPages = Math.min(firstPages, totalPages - lastPages);
    const effectiveFirstPages = Math.max(0, maxNonOverlappingFirstPages);
    const effectiveLastPages = Math.min(lastPages, totalPages);

    if (effectiveFirstPages + effectiveLastPages === 0) {
      return res.status(400).json({
        error: 'No pages to extract after accounting for overlap and total page count.'
      });
    }

    // Create new PDF
    const newPdf = await PDFDocument.create();
    const pageIndexes = [];

    // Add first N pages
    for (let i = 0; i < effectiveFirstPages; i++) {
      pageIndexes.push(i);
    }

    // Add last M pages
    for (let i = totalPages - effectiveLastPages; i < totalPages; i++) {
      if (i >= effectiveFirstPages) { // Avoid duplicating pages
        pageIndexes.push(i);
      }
    }

    // Copy the selected pages
    const copiedPages = await newPdf.copyPages(originalPdf, pageIndexes);
    copiedPages.forEach(page => newPdf.addPage(page));

    // Save and return base64
    const newPdfBytes = await newPdf.save();
    const base64Result = Buffer.from(newPdfBytes).toString('base64');

    res.json({
      success: true,
      pdfData: base64Result,
      originalPageCount: totalPages,
      returnedPages: pageIndexes.length,
      firstPagesExtracted: effectiveFirstPages,
      lastPagesExtracted: effectiveLastPages
    });

  } catch (err) {
    console.error('Error merging first and last pages:', err);
    res.status(500).json({
      error: 'Failed to merge pages',
      details: err.message
    });
  }
});

// Error handling middleware
app.use((error, req, res, next) => {
  if (error instanceof multer.MulterError) {
    if (error.code === 'LIMIT_FILE_SIZE') {
      return res.status(400).json({ error: 'File too large. Maximum size is 10MB.' });
    }
    if (error.code === 'LIMIT_UNEXPECTED_FILE') {
      return res.status(400).json({ error: 'Unexpected file field. Use field name "file".' });
    }
    return res.status(400).json({ error: `Upload error: ${error.message}` });
  }
  res.status(500).json({ error: 'Internal server error' });
});

app.listen(port, () => {
  console.log(`PDF server running on http://localhost:${port}`);
  console.log('Available endpoints:');
  console.log('  GET  /health        - Health check');
  console.log('  POST /api/trim-pdf  - Trim PDF (JSON with base64) - Power Automate friendly');
  console.log('  POST /trim-pdf      - Trim PDF (multipart form) - Original endpoint');
  console.log('  POST /api/last-2-pages - Extract last 2 pages (JSON with base64)');
  console.log('  POST /api/last-n-pages - Extract last N pages (JSON with base64)');
  console.log('  POST /api/merge-first-last-pages - Merge first N and last M pages (JSON with base64)');
  console.log('');
  console.log('For Power Automate, use /api/merge-first-last-pages with JSON body:');
  console.log('  {');
  console.log('    "pdfData": "base64-encoded-pdf-content",');
  console.log('    "firstPages": 2,');
  console.log('    "lastPages": 8');
  console.log('  }');
});