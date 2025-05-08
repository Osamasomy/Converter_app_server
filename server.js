const express = require('express');
const path = require('path');
const multer = require('multer');
const fs = require('fs');
const fetch = require('node-fetch');
const PDFDocument = require('pdf-lib').PDFDocument;
const pdfParse = require('pdf-parse');
const docx = require('docx');
const { Document, Paragraph, TextRun } = docx;
const poppler = require('pdf-poppler');
const { mkdirp } = require('mkdirp');
const PDFKit = require('pdfkit');
const sharp = require('sharp');
const { createWorker } = require('tesseract.js'); // Add Tesseract.js v3 for OCR

// API key for PDF.co
const PDFCO_API_KEY = 'khawajaabdullah688@gmail.com_uNwC3CPFIqb2XnHHnRfRU6QiKDxdR6KeWsFNh805GeN4r4ugYcZqJtzUn95v2uC6';

// Create directories if they don't exist
const uploadDir = path.join(__dirname, 'images');
const wordDir = path.join(__dirname, 'document', 'word');
const pdfImagesDir = path.join(__dirname, 'pdf-images');
const pdfDir = path.join(__dirname, 'document', 'pdf'); // New directory for generated PDFs
const textOutputDir = path.join(__dirname, 'document', 'text'); // New directory for text output

[uploadDir, wordDir, pdfImagesDir, pdfDir, textOutputDir].forEach(dir => {
  if (!fs.existsSync(dir)) {
    fs.mkdirSync(dir, { recursive: true });
  }
});

const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    // Choose destination based on file type/route
    if (req.path === '/convert-pdf-to-images') {
      cb(null, pdfImagesDir);
    } else if (file.fieldname === 'pdf') {
      cb(null, wordDir);
    } else if (req.path === '/convert-images-to-pdf') {
      cb(null, uploadDir); // Store images in upload directory before processing
    } else if (req.path === '/convert-image-to-text') {
      cb(null, uploadDir); // Store image in upload directory for OCR processing
    } else {
      cb(null, uploadDir);
    }
  },
  filename: (req, file, cb) => {
    cb(null, `${file.fieldname}_${Date.now()}${path.extname(file.originalname)}`);
  }
});

const upload = multer({ storage });

const app = express();
const port = process.env.PORT || 3000;

app.use(express.json());
app.use('/images', express.static('images'));
app.use('/document/word', express.static(path.join(__dirname, 'document', 'word')));
app.use('/pdf-images', express.static(pdfImagesDir));
app.use('/document/pdf', express.static(pdfDir)); // Serve PDF files
app.use('/document/text', express.static(textOutputDir)); // Serve text files

app.get('/', (req, res) => {
  res.send('Welcome to my server!');
});

app.post('/image', upload.single("img"), (req, res) => {
  const file = req.file;

  if (!file) {
    return res.status(400).json({ success: 0, message: 'No file uploaded' });
  }

  return res.json({
    success: 1,
    url: `http://192.168.0.106:${port}/images/${file.filename}`
  });
});

// New API: Convert Image to Text using Tesseract.js
app.post('/convert-image-to-text', upload.single('image'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ success: false, error: 'No image file uploaded' });
    }

    console.log(`Processing image for OCR: ${req.file.path}`);

    // Optional parameters from request
    const lang = req.body.lang || 'eng'; // Default language is English
    const outputFormat = req.body.format || 'txt'; // Default output format is txt

    // Generate output filename
    const textFileName = `text_${Date.now()}.${outputFormat}`;
    const textFilePath = path.join(textOutputDir, textFileName);

    console.log(`Starting OCR process with language: ${lang}`);

    // Perform OCR using Tesseract.js v3
    // Initialize worker
    const worker = await createWorker(lang);

    // Recognize text in image
    const { data } = await worker.recognize(req.file.path);

    // Terminate worker when done
    await worker.terminate();

    // Write the extracted text to a file
    fs.writeFileSync(textFilePath, outputFormat === 'json' ?
      JSON.stringify(data, null, 2) :
      data.text
    );

    console.log(`OCR completed successfully. Text saved to: ${textFilePath}`);

    // Create response
    const response = {
      success: true,
      message: 'Image converted to text successfully',
      text: data.text.substring(0, 500) + (data.text.length > 500 ? '...' : ''), // Preview of text (first 500 chars)
      confidence: data.confidence,
      textUrl: `http://192.168.0.106:${port}/document/text/${textFileName}`,
      imageUrl: `http://192.168.0.106:${port}/images/${req.file.filename}`,
      language: lang
    };

    // If JSON format was requested, include full data structure
    if (outputFormat === 'json') {
      response.data = data;
    }

    res.json(response);
  } catch (error) {
    console.error('Error converting image to text:', error);
    res.status(500).json({
      success: false,
      message: 'Failed to convert image to text',
      error: error.message
    });
  }
});

// New route: Convert Images to PDF
app.post('/convert-images-to-pdf', upload.array('images', 20), async (req, res) => {
  try {
    if (!req.files || req.files.length === 0) {
      return res.status(400).json({ success: false, error: 'No image files uploaded' });
    }

    console.log(`Received ${req.files.length} images for PDF conversion`);

    // Generate PDF filename
    const pdfFileName = `pdf_${Date.now()}.pdf`;
    const pdfFilePath = path.join(pdfDir, pdfFileName);

    // Create PDF document using PDFKit
    const doc = new PDFKit({ autoFirstPage: false });
    const writeStream = fs.createWriteStream(pdfFilePath);
    doc.pipe(writeStream);


    for (const file of req.files) {
      try {
        console.log(`Processing image: ${file.path}`);

        // Get image dimensions using sharp
        const metadata = await sharp(file.path).metadata();

        // Calculate PDF page dimensions based on image
        const pdfWidth = 612; // Default US Letter width in points
        const pdfHeight = 792; // Default US Letter height in points

        // Create new page with proper dimensions
        doc.addPage({ size: [pdfWidth, pdfHeight] });

        // Calculate scaling to fit image properly on page (with margins)
        const margin = 40;
        const maxWidth = pdfWidth - (margin * 2);
        const maxHeight = pdfHeight - (margin * 2);

        let scale = Math.min(
          maxWidth / metadata.width,
          maxHeight / metadata.height
        );

        // Calculate dimensions for centered image
        const imgWidth = metadata.width * scale;
        const imgHeight = metadata.height * scale;
        const x = (pdfWidth - imgWidth) / 2;
        const y = (pdfHeight - imgHeight) / 2;

        // Add image to PDF page
        doc.image(file.path, x, y, {
          width: imgWidth,
          height: imgHeight
        });

      } catch (imageError) {
        console.error(`Error processing image ${file.path}:`, imageError);
        // Continue with other images even if one fails
      }
    }

    // Finalize PDF
    doc.end();

    // Wait for PDF to finish writing
    await new Promise((resolve, reject) => {
      writeStream.on('finish', resolve);
      writeStream.on('error', reject);
    });

    console.log(`PDF created successfully: ${pdfFilePath}`);

    res.json({
      success: true,
      message: `${req.files.length} images converted to PDF successfully`,
      pdfUrl: `http://192.168.0.106:${port}/document/pdf/${pdfFileName}`,
      imageCount: req.files.length
    });

  } catch (error) {
    console.error('Error converting images to PDF:', error);
    res.status(500).json({
      success: false,
      message: 'Failed to convert images to PDF',
      error: error.message
    });
  }
});

// Single image to PDF conversion
app.post('/convert-image-to-pdf', upload.single('image'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ success: false, error: 'No image file uploaded' });
    }

    console.log(`Converting single image to PDF: ${req.file.path}`);

    // Generate PDF filename
    const pdfFileName = `pdf_${Date.now()}.pdf`;
    const pdfFilePath = path.join(pdfDir, pdfFileName);

    // Get image dimensions using sharp
    const metadata = await sharp(req.file.path).metadata();

    // Create PDF document using PDFKit
    const doc = new PDFKit({
      size: [metadata.width, metadata.height],
      margins: { top: 0, bottom: 0, left: 0, right: 0 }
    });

    const writeStream = fs.createWriteStream(pdfFilePath);
    doc.pipe(writeStream);

    // Add image to PDF with full dimensions
    doc.image(req.file.path, 0, 0, {
      width: metadata.width,
      height: metadata.height,
      align: 'center',
      valign: 'center'
    });

    // Finalize PDF
    doc.end();

    // Wait for PDF to finish writing
    await new Promise((resolve, reject) => {
      writeStream.on('finish', resolve);
      writeStream.on('error', reject);
    });

    console.log(`PDF created successfully: ${pdfFilePath}`);

    res.json({
      success: true,
      message: 'Image converted to PDF successfully',
      pdfUrl: `http://192.168.0.106:${port}/document/pdf/${pdfFileName}`,
      originalImage: req.file.filename
    });

  } catch (error) {
    console.error('Error converting image to PDF:', error);
    res.status(500).json({
      success: false,
      message: 'Failed to convert image to PDF',
      error: error.message
    });
  }
});

app.post('/convert-pdf-to-long-image', upload.single('pdf'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ success: false, error: 'No PDF file uploaded' });
    }

    const pdfPath = req.file.path;
    const tempOutputFolder = path.join(pdfImagesDir, `temp_${Date.now()}`);

    // Create temp directory for the individual page images
    if (!fs.existsSync(tempOutputFolder)) {
      fs.mkdirSync(tempOutputFolder);
    }

    // Convert PDF to individual page images first
    const options = {
      format: 'png',
      out_dir: tempOutputFolder,
      out_prefix: 'page',
      page: null,
      resolution: 300
    };

    await poppler.convert(pdfPath, options);

    // Get all generated images in order
    const imageFiles = fs.readdirSync(tempOutputFolder)
      .filter(file => file.startsWith('page') && file.endsWith('.png'))
      .sort((a, b) => {
        const numA = parseInt(a.replace('page-', '').replace('.png', ''));
        const numB = parseInt(b.replace('page-', '').replace('.png', ''));
        return numA - numB;
      });

    // Read all images to get dimensions
    const images = await Promise.all(
      imageFiles.map(file =>
        sharp(path.join(tempOutputFolder, file)).metadata()
      )
    );

    // Calculate total height
    const width = images[0].width;
    const totalHeight = images.reduce((sum, img) => sum + img.height, 0);

    // Create a new image with the combined dimensions
    const composite = sharp({
      create: {
        width: width,
        height: totalHeight,
        channels: 4,
        background: { r: 255, g: 255, b: 255, alpha: 1 }
      }
    });

    // Prepare overlays for the composite
    const overlays = [];
    let currentY = 0;

    for (let i = 0; i < imageFiles.length; i++) {
      overlays.push({
        input: path.join(tempOutputFolder, imageFiles[i]),
        top: currentY,
        left: 0
      });
      currentY += images[i].height;
    }

    // Generate the long image
    const longImageName = `long_pdf_${Date.now()}.png`;
    const longImagePath = path.join(pdfImagesDir, longImageName);

    await composite
      .composite(overlays)
      .toFile(longImagePath);

    // Clean up temp files if desired
    fs.rmSync(tempOutputFolder, { recursive: true, force: true });

    res.json({
      success: true,
      message: `PDF converted to long image successfully`,
      imageUrl: `http://192.168.0.106:${port}/pdf-images/${longImageName}`,
      pageCount: imageFiles.length
    });

  } catch (error) {
    console.error('Error converting PDF to long image:', error);
    res.status(500).json({
      success: false,
      message: 'Failed to convert PDF to long image',
      error: error.message
    });
  }
});

app.post('/convert-pdf-to-images', upload.single('pdf'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ success: false, error: 'No PDF file uploaded' });
    }

    const pdfPath = req.file.path;
    const outputFolder = path.join(pdfImagesDir, `pdf_${Date.now()}`);

    // Create output directory for this specific PDF
    if (!fs.existsSync(outputFolder)) {
      fs.mkdirSync(outputFolder);
    }

    // Options for pdf-poppler
    const options = {
      format: 'png',
      out_dir: outputFolder,
      out_prefix: 'page',
      page: null, // Convert all pages
      resolution: 300
    };

    // Convert PDF to images using pdf-poppler
    await poppler.convert(pdfPath, options);

    // Get all generated images
    const images = fs.readdirSync(outputFolder)
      .filter(file => file.startsWith('page') && file.endsWith('.png'))
      .sort((a, b) => {
        // Sort pages numerically
        const numA = parseInt(a.replace('page-', '').replace('.png', ''));
        const numB = parseInt(b.replace('page-', '').replace('.png', ''));
        return numA - numB;
      });

    // Create URLs for all images
    const imageUrls = images.map(imageName =>
      `http://192.168.0.106:${port}/pdf-images/${path.basename(outputFolder)}/${imageName}`
    );

    res.json({
      success: true,
      message: `PDF converted to ${images.length} images successfully`,
      pdfName: req.file.originalname,
      imageCount: images.length,
      images: imageUrls
    });

  } catch (error) {
    console.error('Error converting PDF to images:', error);
    res.status(500).json({
      success: false,
      message: 'Failed to convert PDF to images',
      error: error.message
    });
  }
});

app.post('/convert-pdf-to-word', upload.single('pdf'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ success: false, error: 'No PDF file uploaded' });

    const pdfPath = req.file.path;
    const pdfBuffer = fs.readFileSync(pdfPath);

    // Extract text from PDF
    const pdfData = await pdfParse(pdfBuffer);
    const pdfText = pdfData.text;

    // Create paragraphs from text (split by double newlines)
    const paragraphs = pdfText.split(/\n\s*\n/).filter(para => para.trim().length > 0);

    // Create a new Word document
    const doc = new Document({
      sections: [{
        properties: {},
        children: paragraphs.map(para =>
          new Paragraph({
            children: [
              new TextRun({
                text: para.trim(),
                size: 24 // 12pt font
              })
            ]
          })
        )
      }]
    });

    // Generate Word document
    const wordFileName = `word_${Date.now()}.docx`;
    const wordFilePath = path.join(wordDir, wordFileName);

    // Write the Word document to file
    const buffer = await docx.Packer.toBuffer(doc);
    fs.writeFileSync(wordFilePath, buffer);

    res.json({
      success: true,
      message: 'PDF converted to Word successfully',
      localUrl: `http://192.168.0.106:${port}/document/word/${wordFileName}`,
      filename: wordFileName
    });
  } catch (err) {
    console.error('Error in PDF to Word conversion process:', err);
    res.status(500).json({
      success: false,
      message: 'Failed to convert PDF to Word',
      error: err.message
    });
  }
});

app.post('/convert-word-to-pdf', upload.single('word'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ success: false, error: 'No Word file uploaded' });
    }
    
    // Check if file is a Word document
    const fileExtension = path.extname(req.file.originalname).toLowerCase();
    if (!['.doc', '.docx'].includes(fileExtension)) {
      return res.status(400).json({ 
        success: false, 
        error: 'Uploaded file is not a Word document. Only .doc and .docx files are supported.'
      });
    }

    const wordPath = req.file.path;
    console.log('Processing Word document:', wordPath);
    
    // Make sure the PDF directory exists
    try {
      if (!fs.existsSync(pdfDir)) {
        fs.mkdirSync(pdfDir, { recursive: true });
      }
    } catch (dirError) {
      console.error('Error creating PDF directory:', dirError);
      return res.status(500).json({
        success: false,
        message: 'Failed to create PDF directory',
        error: dirError.message
      });
    }
    
    // Generate PDF filename
    const pdfFileName = `pdf_${Date.now()}.pdf`;
    const pdfFilePath = path.join(pdfDir, pdfFileName);
    
    // Try each conversion method in sequence, with proper error handling
    let conversionSuccessful = false;
    let conversionError = null;
    
    // Method 2: Use docx-pdf as another approach for .docx files
    if (fileExtension === '.docx' && !conversionSuccessful) {
      try {
        // Check if required package is installed
        let docxToPdf;
        try {
          docxToPdf = require('docx-pdf');
        } catch (requireError) {
          console.error('Error requiring docx-pdf:', requireError);
          throw new Error('docx-pdf package not available: ' + requireError.message);
        }
        
        await new Promise((resolve, reject) => {
          docxToPdf(wordPath, pdfFilePath, (err, result) => {
            if (err) {
              console.error('Error converting using docx-pdf:', err);
              reject(err);
            } else {
              resolve(result);
            }
          });
        });
        
        console.log('PDF created successfully using docx-pdf:', pdfFilePath);
        conversionSuccessful = true;
        
        return res.json({
          success: true,
          message: 'Word document converted to PDF successfully using docx-pdf',
          pdfUrl: `http://192.168.0.106:${port}/document/pdf/${pdfFileName}`,
          originalName: req.file.originalname,
          fileName: pdfFileName
        });
      } catch (docxPdfError) {
        console.error('Error converting using docx-pdf:', docxPdfError);
        conversionError = conversionError || docxPdfError;
        // Fall through to Method 3
      }
    }
    
    if (!conversionSuccessful) {
      let errorMessage = 'Failed to convert document after trying multiple methods.';
      
      if (fileExtension === '.doc') {
        errorMessage += ' Only .docx will convert to PDF.';
      } else {
        errorMessage += ' Please try with a different document or format.';
      }
      
      return res.status(500).json({
        success: false,
        message: 'Failed to convert Word to PDF',
        error: errorMessage,
        technicalDetails: conversionError ? conversionError.message : 'Unknown error'
      });
    }
  } catch (err) {
    console.error('Error in Word to PDF conversion process:', err);
    return res.status(500).json({ 
      success: false, 
      message: 'Failed to convert Word to PDF', 
      error: err.message 
    });
  }
});

// app.post('/convert-word-to-pdf', upload.single('word'), async (req, res) => {
//   try {
//     if (!req.file) {
//       return res.status(400).json({ success: false, error: 'No Word file uploaded' });
//     }

//     // Check if file is a Word document
//     const fileExtension = path.extname(req.file.originalname).toLowerCase();
//     if (!['.doc', '.docx'].includes(fileExtension)) {
//       return res.status(400).json({ 
//         success: false, 
//         error: 'Uploaded file is not a Word document. Only .doc and .docx files are supported.'
//       });
//     }

//     const wordPath = req.file.path;
//     console.log('Processing Word document:', wordPath);

//     // Use docx library to convert Word to PDF
//     // For .docx files, we'll use docx library's capabilities
//     // Note: This method works best with .docx files, not .doc files
//     try {
//       // Generate PDF filename
//       const pdfFileName = `pdf_${Date.now()}.pdf`;
//       const pdfFilePath = path.join(pdfDir, pdfFileName);

//       // For .docx files using docx-to-pdf approach
//       if (fileExtension === '.docx') {
//         // Create a PDF document
//         const pdfDoc = new PDFKit();
//         const writeStream = fs.createWriteStream(pdfFilePath);
//         pdfDoc.pipe(writeStream);

//         // Read the Word document
//         const content = await fs.promises.readFile(wordPath);

//         // Parse the content (this is a simplified approach - in reality this would require
//         // more sophisticated Word document parsing)
//         try {
//           const textContent = await extractTextFromDocx(wordPath);

//           // Add text content to PDF
//           pdfDoc.fontSize(12);
//           pdfDoc.text(textContent);

//           // Finalize the PDF file
//           pdfDoc.end();

//           // Wait for the write stream to finish
//           await new Promise((resolve, reject) => {
//             writeStream.on('finish', resolve);
//             writeStream.on('error', reject);
//           });

//           console.log('PDF created successfully:', pdfFilePath);

//           // Return success response
//           return res.json({
//             success: true,
//             message: 'Word document converted to PDF successfully',
//             pdfUrl: `http://192.168.0.106:${port}/document/pdf/${pdfFileName}`,
//             originalName: req.file.originalname,
//             fileName: pdfFileName
//           });
//         } catch (parseError) {
//           console.error('Error parsing Word document:', parseError);
//           throw new Error(`Failed to parse Word document: ${parseError.message}`);
//         }
//       } else {
//         // For .doc files
//         throw new Error('Direct conversion of .doc files is not supported. Please convert to .docx first.');
//       }
//     } catch (conversionError) {
//       console.error('Error in direct conversion:', conversionError);

//       // If direct conversion fails, we can use PDF.co API as a fallback
//       // But we need to uncomment and implement the PDF.co code from your original implementation
//       throw new Error(`Direct conversion failed: ${conversionError.message}. Please implement PDF.co API as fallback.`);
//     }
//   } catch (err) {
//     console.error('Error in Word to PDF conversion process:', err);
//     res.status(500).json({ 
//       success: false, 
//       message: 'Failed to convert Word to PDF', 
//       error: err.message 
//     });
//   }
// });

// // Helper function to extract text from a .docx file
// async function extractTextFromDocx(filePath) {
//   // This is a simplified implementation
//   // For a full implementation, consider using a library like mammoth.js
//   // which is better at preserving formatting

//   try {
//     // Read the .docx file
//     const content = await fs.promises.readFile(filePath);

//     // Create a new Document instance
//     const doc = new docx.Document({
//       sections: []
//     });

//     // Extract text (simplified approach)
//     // In a real implementation, we would parse the docx structure
//     // and extract text with formatting

//     // For now, we'll just return a placeholder message
//     // return "This is a simplified text extraction from the Word document. For better results, implement a proper Word document parser like mammoth.js";

//     // Alternatively, if you add mammoth.js to your project:
//     const mammoth = require('mammoth');
//     const result = await mammoth.extractRawText({path: filePath});
//     return result.value;
//   } catch (error) {
//     console.error('Error extracting text from Word document:', error);
//     throw error;
//   }
// }

// app.post('/convert-pdf-to-excel', upload.single('pdf'), async (req, res) => {
//   try {
//     if (!req.file) return res.status(400).json({ success: false, error: 'No file uploaded' });

//     const pdfPath = req.file.path;
//     const fileBuffer = fs.readFileSync(pdfPath);
//     const base64Data = fileBuffer.toString('base64');

//     console.log('Uploading PDF to PDF.co...');

//     // Upload PDF to PDF.co
//     const uploadRes = await fetch('https://api.pdf.co/v1/file/upload/base64', {
//       method: 'POST',
//       headers: {
//         'x-api-key': PDFCO_API_KEY,
//         'Content-Type': 'application/json'
//       },
//       body: JSON.stringify({
//         name: req.file.originalname,
//         file: base64Data
//       })
//     });

//     if (!uploadRes.ok) {
//       const errorText = await uploadRes.text();
//       console.error('Upload response not OK:', uploadRes.status, errorText);
//       throw new Error(`Upload failed with status ${uploadRes.status}: ${errorText}`);
//     }

//     const uploadJson = await uploadRes.json();
//     if (uploadJson.error) {
//       console.error('Upload API error:', uploadJson);
//       throw new Error(uploadJson.message || 'Error uploading file to PDF.co');
//     }

//     const uploadedUrl = uploadJson.url;
//     console.log('File uploaded successfully, URL:', uploadedUrl);

//     // Convert PDF to Excel
//     console.log('Converting PDF to Excel...');
//     const convertRes = await fetch('https://api.pdf.co/v1/pdf/convert/to/xls', {
//       method: 'POST',
//       headers: {
//         'x-api-key': PDFCO_API_KEY,
//         'Content-Type': 'application/json'
//       },
//       body: JSON.stringify({
//         url: uploadedUrl,
//         name: `converted_${Date.now()}.xls`,
//         async: false
//       })
//     });

//     if (!convertRes.ok) {
//       const errorText = await convertRes.text();
//       console.error('Convert response not OK:', convertRes.status, errorText);
//       throw new Error(`Conversion failed with status ${convertRes.status}: ${errorText}`);
//     }

//     const convertJson = await convertRes.json();
//     if (convertJson.error) {
//       console.error('Convert API error:', convertJson);
//       throw new Error(convertJson.message || 'Error converting file with PDF.co');
//     }

//     console.log('File converted successfully, URL:', convertJson.url);

//     // Download the converted file and save it locally
//     console.log('Downloading converted file...');
//     const excelFileRes = await fetch(convertJson.url);
//     if (!excelFileRes.ok) {
//       throw new Error(`Failed to download the converted file: ${excelFileRes.status}`);
//     }

//     const excelFileName = `excel_${Date.now()}.xls`;
//     const excelFilePath = path.join(wordDir, excelFileName);

//     const fileStream = fs.createWriteStream(excelFilePath);
//     excelFileRes.body.pipe(fileStream);

//     await new Promise((resolve, reject) => {
//       fileStream.on('finish', resolve);
//       fileStream.on('error', reject);
//     });

//     console.log('Excel file saved locally:', excelFilePath);

//     // Return both the external URL and local URL
//     res.json({ 
//       success: true,
//       externalUrl: convertJson.url,
//       localUrl: `http://192.168.0.106:${port}/document/word/${excelFileName}`
//     });
//   } catch (err) {
//     console.error('Error in conversion process:', err);
//     res.status(500).json({ success: false, error: err.message });
//   }
// });

// Word to PDF conversion route

// app.post('/convert-word-to-pdf', upload.single('word'), async (req, res) => {
//   try {
//     if (!req.file) {
//       return res.status(400).json({ success: false, error: 'No Word file uploaded' });
//     }

//     // Check if file is a Word document
//     const fileExtension = path.extname(req.file.originalname).toLowerCase();
//     if (!['.doc', '.docx'].includes(fileExtension)) {
//       return res.status(400).json({ 
//         success: false, 
//         error: 'Uploaded file is not a Word document. Only .doc and .docx files are supported.'
//       });
//     }

//     const wordPath = req.file.path;
//     const fileBuffer = fs.readFileSync(wordPath);
//     const base64Data = fileBuffer.toString('base64');

//     console.log('Uploading Word document to PDF.co...');

//     // Upload Word file to PDF.co
//     const uploadRes = await fetch('https://api.pdf.co/v1/file/upload/base64', {
//       method: 'POST',
//       headers: {
//         'x-api-key': PDFCO_API_KEY,
//         'Content-Type': 'application/json'
//       },
//       body: JSON.stringify({
//         name: req.file.originalname,
//         file: base64Data
//       })
//     });

//     if (!uploadRes.ok) {
//       const errorText = await uploadRes.text();
//       console.error('Upload response not OK:', uploadRes.status, errorText);
//       throw new Error(`Upload failed with status ${uploadRes.status}: ${errorText}`);
//     }

//     const uploadJson = await uploadRes.json();
//     if (uploadJson.error) {
//       console.error('Upload API error:', uploadJson);
//       throw new Error(uploadJson.message || 'Error uploading file to PDF.co');
//     }

//     const uploadedUrl = uploadJson.url;
//     console.log('Word file uploaded successfully, URL:', uploadedUrl);

//     // Convert Word to PDF
//     console.log('Converting Word to PDF...');
//     const convertRes = await fetch('https://api.pdf.co/v1/pdf/convert/from/doc', {
//       method: 'POST',
//       headers: {
//         'x-api-key': PDFCO_API_KEY,
//         'Content-Type': 'application/json'
//       },
//       body: JSON.stringify({
//         url: uploadedUrl,
//         name: `pdf_${Date.now()}.pdf`,
//         async: false
//       })
//     });

//     if (!convertRes.ok) {
//       const errorText = await convertRes.text();
//       console.error('Convert response not OK:', convertRes.status, errorText);
//       throw new Error(`Conversion failed with status ${convertRes.status}: ${errorText}`);
//     }

//     const convertJson = await convertRes.json();
//     if (convertJson.error) {
//       console.error('Convert API error:', convertJson);
//       throw new Error(convertJson.message || 'Error converting file with PDF.co');
//     }

//     console.log('File converted to PDF successfully, URL:', convertJson.url);

//     // Download the converted PDF file and save it locally
//     console.log('Downloading converted PDF file...');
//     const pdfFileRes = await fetch(convertJson.url);
//     if (!pdfFileRes.ok) {
//       throw new Error(`Failed to download the converted PDF file: ${pdfFileRes.status}`);
//     }

//     const pdfFileName = `pdf_${Date.now()}.pdf`;
//     const pdfFilePath = path.join(pdfDir, pdfFileName);

//     const fileStream = fs.createWriteStream(pdfFilePath);
//     pdfFileRes.body.pipe(fileStream);

//     await new Promise((resolve, reject) => {
//       fileStream.on('finish', resolve);
//       fileStream.on('error', reject);
//     });

//     console.log('PDF file saved locally:', pdfFilePath);

//     // Return the result with URLs
//     res.json({ 
//       success: true,
//       message: 'Word document converted to PDF successfully',
//       externalUrl: convertJson.url,
//       pdfUrl: `http://192.168.0.106:${port}/document/pdf/${pdfFileName}`,
//       originalName: req.file.originalname,
//       fileName: pdfFileName
//     });
//   } catch (err) {
//     console.error('Error in Word to PDF conversion process:', err);
//     res.status(500).json({ 
//       success: false, 
//       message: 'Failed to convert Word to PDF', 
//       error: err.message 
//     });
//   }
// });

// // Fix for PDF to Word conversion function
// app.post('/convert-pdf-to-word', upload.single('pdf'), async (req, res) => {
//   try {
//     if (!req.file) return res.status(400).json({ success: false, error: 'No PDF file uploaded' });

//     const pdfPath = req.file.path;
//     const fileBuffer = fs.readFileSync(pdfPath);
//     const base64Data = fileBuffer.toString('base64');

//     console.log('Uploading PDF to PDF.co for Word conversion...');

//     // Upload PDF to PDF.co
//     const uploadRes = await fetch('https://api.pdf.co/v1/file/upload/base64', {
//       method: 'POST',
//       headers: {
//         'x-api-key': PDFCO_API_KEY,
//         'Content-Type': 'application/json'
//       },
//       body: JSON.stringify({
//         name: req.file.originalname,
//         file: base64Data
//       })
//     });

//     if (!uploadRes.ok) {
//       const errorText = await uploadRes.text();
//       console.error('Upload response not OK:', uploadRes.status, errorText);
//       throw new Error(`Upload failed with status ${uploadRes.status}: ${errorText}`);
//     }

//     const uploadJson = await uploadRes.json();
//     if (uploadJson.error) {
//       console.error('Upload API error:', uploadJson);
//       throw new Error(uploadJson.message || 'Error uploading file to PDF.co');
//     }

//     const uploadedUrl = uploadJson.url;
//     console.log('PDF file uploaded successfully, URL:', uploadedUrl);

//     // Convert PDF to Word (DOCX)
//     console.log('Converting PDF to Word...');

//     // Use the correct endpoint for PDF to Word conversion
//     const convertRes = await fetch('https://api.pdf.co/v1/pdf/convert/to/doc', {
//       method: 'POST',
//       headers: {
//         'x-api-key': PDFCO_API_KEY,
//         'Content-Type': 'application/json'
//       },
//       body: JSON.stringify({
//         url: uploadedUrl,
//         name: `word_${Date.now()}.doc`,
//         async: false
//       })
//     });

//     if (!convertRes.ok) {
//       const errorText = await convertRes.text();
//       console.error('Convert response not OK:', convertRes.status, errorText);
//       throw new Error(`Conversion failed with status ${convertRes.status}: ${errorText}`);
//     }

//     const convertJson = await convertRes.json();
//     if (convertJson.error) {
//       console.error('Convert API error:', convertJson);
//       throw new Error(convertJson.message || 'Error converting file with PDF.co');
//     }

//     console.log('File converted to Word successfully, URL:', convertJson.url);

//     // Download the converted Word file and save it locally
//     console.log('Downloading converted Word file...');
//     const wordFileRes = await fetch(convertJson.url);
//     if (!wordFileRes.ok) {
//       throw new Error(`Failed to download the converted Word file: ${wordFileRes.status}`);
//     }

//     const wordFileName = `word_${Date.now()}.doc`;
//     const wordFilePath = path.join(wordDir, wordFileName);

//     const fileStream = fs.createWriteStream(wordFilePath);
//     wordFileRes.body.pipe(fileStream);

//     await new Promise((resolve, reject) => {
//       fileStream.on('finish', resolve);
//       fileStream.on('error', reject);
//     });

//     console.log('Word file saved locally:', wordFilePath);

//     // Return both the external URL and local URL
//     res.json({ 
//       success: true,
//       message: 'PDF converted to Word successfully',
//       externalUrl: convertJson.url,
//       localUrl: `http://192.168.0.106:${port}/document/word/${wordFileName}`,
//       filename: wordFileName
//     });
//   } catch (err) {
//     console.error('Error in PDF to Word conversion process:', err);
//     res.status(500).json({ 
//       success: false, 
//       message: 'Failed to convert PDF to Word', 
//       error: err.message 
//     });
//   }
// });

app.listen(port, () => {
  console.log(`Server is running on port ${port}`);
});