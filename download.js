const express = require('express');
const { google } = require('googleapis');
const cors = require('cors');
const bodyParser = require('body-parser');
require('dotenv').config();
const fs = require('fs');
const path = require('path');
const archiver = require('archiver');
const nodemailer = require('nodemailer');

const app = express();
const port = 3001;

// CORS configuration
app.use(cors({
  origin: '*',
  methods: ['GET', 'POST'],
  allowedHeaders: ['Content-Type', 'Authorization']
}));

app.use(bodyParser.json());

// Google Sheets and Drive credentials
const oauth2Client = new google.auth.OAuth2(
  process.env.CLIENT_ID,
  process.env.CLIENT_SECRET,
  process.env.REDIRECT_URI
);

oauth2Client.setCredentials({ refresh_token: process.env.REFRESH_TOKEN });

const sheets = google.sheets({ version: 'v4', auth: oauth2Client });
const drive = google.drive({ version: 'v3', auth: oauth2Client });
const slides = google.slides({ version: 'v1', auth: oauth2Client });

// Nodemailer configuration
const transporter = nodemailer.createTransport({
  service: 'gmail',
  auth: {
    user: process.env.EMAIL,
    pass: process.env.PASSWORD
  }
});

// Log file path
const logFilePath = path.join(__dirname, 'download.json');

// Initialize the log file on server start
fs.writeFileSync(logFilePath, JSON.stringify({
  progress: 0,
  totalCertificates: 0,
  generatedCount: 0,
  generating: false
}));

// Route to generate certificates
app.post('/generate-certificates', async (req, res) => {
  const { sheetId, sheetName, date, todate, school } = req.body;

  if (!sheetId || !sheetName || !date || !todate) {
    return res.status(400).json({ status: 'error', message: 'One or more parameters are missing.' });
  }

  try {
    const sheetData = await getSheetData(sheetId, sheetName);
    const filteredData = school ? sheetData.filter(row => row[2]?.toString().toUpperCase() === school.toUpperCase()) : sheetData;
    const totalCertificates = filteredData.length - 1; // Subtract 1 for the header row

    fs.writeFileSync(logFilePath, JSON.stringify({
      progress: 0,
      totalCertificates,
      generatedCount: 0,
      generating: true
    }));

    if (!school) {
      await sendCertificates(filteredData,date,todate);
      await new Promise(resolve => setTimeout(resolve, 4000));
      await generateCertificatesAsSinglePDF(filteredData, date, todate);
    } else {
      // Specific school selected, generate certificates as a ZIP file
      await generateCertificates(filteredData, date, todate, (generatedCount) => {
        const progress = calculatePercentage(generatedCount, totalCertificates);
        fs.writeFileSync(logFilePath, JSON.stringify({
          progress,
          totalCertificates,
          generatedCount,
          generating: true
        }));
      });
    }

    fs.writeFileSync(logFilePath, JSON.stringify({
      progress: 100,
      totalCertificates,
      generatedCount: totalCertificates,
      generating: false
    }));

    res.status(200).json({ status: 'success', message: 'Certificates generation started successfully! You can download the file once it is ready.' });

  } catch (error) {
    console.error('Error in /generate-certificates:', error);
    res.status(500).json({ status: 'error', message: 'An error occurred while generating certificates. Please check the server logs.' });
  }
});

// Route to fetch unique school names
app.post('/unique-schools', async (req, res) => {
  const { sheetId, sheetName } = req.body;

  if (!sheetId || !sheetName) {
    return res.status(400).json({ status: 'error', message: 'Sheet ID and Sheet Name are required.' });
  }

  try {
    const sheetData = await getSheetData(sheetId, sheetName);
    const schools = [...new Set(sheetData.slice(1).map(row => row[2]?.toString().toUpperCase()))];

    res.json({ schools });
  } catch (error) {
    console.error('Error in /unique-schools:', error);
    res.status(500).json({ status: 'error', message: 'An error occurred while fetching school names. Please check the server logs.' });
  }
});

// Route to fetch progress
app.get('/fetch-progress', (req, res) => {
  if (fs.existsSync(logFilePath)) {
    const logData = JSON.parse(fs.readFileSync(logFilePath, 'utf8'));
    res.json(logData);
  } else {
    res.json({ progress: 0, totalCertificates: 0, generatedCount: 0, generating: false }); // Return 0 values if file does not exist
  }
});

// Route to download the file (ZIP or PDF)
app.get('/download-file', (req, res) => {
  const zipFilePath = path.join(__dirname, 'certificates.zip');
  const pdfFilePath = path.join(__dirname, 'certificates.pdf');

  if (fs.existsSync(zipFilePath)) {
    res.download(zipFilePath, 'certificates.zip', (err) => {
      if (err) {
        console.error('Error downloading zip file:', err);
        res.status(500).json({ status: 'error', message: 'Error downloading zip file.' });
      }
      fs.unlink(zipFilePath, (err) => {
        if (err) console.error('Error deleting zip file:', err);
      });
    });
  } else if (fs.existsSync(pdfFilePath)) {
    res.download(pdfFilePath, 'certificates.pdf', (err) => {
      if (err) {
        console.error('Error downloading PDF file:', err);
        res.status(500).json({ status: 'error', message: 'Error downloading PDF file.' });
      }
      fs.unlink(pdfFilePath, (err) => {
        if (err) console.error('Error deleting PDF file:', err);
      });
    });
  } else {
    res.status(404).json({ status: 'error', message: 'File not found.' });
  }
});

async function sendCertificates(sheetData, date, todate) {
  const templateId = process.env.TEMPLATE_ID;
  const folderId = process.env.FOLDER_ID;

  for (let i = 1; i < sheetData.length; i++) {
    const row = sheetData[i];
    const name = row[0]?.toString() || '';
    const email = row[1]?.toString() || '';
    const schoolName = row[2]?.toString() || '';
    const domain = row[3]?.toString() || '';
    const certificateNumber = row[4]?.toString().toUpperCase() || '';
    const formattedDate = formatDateToReadable(new Date(date));
    const formattedtoDate = formatDateToReadable(new Date(todate));

    const copyFile = await drive.files.copy({
      fileId: templateId,
      requestBody: {
        name: `${name} - Certificate`,
        parents: [folderId]
      }
    });

    const copyId = copyFile.data.id;

    function capitalizeWords(domain) {
      return domain.replace(/\b\w/g, char => char.toUpperCase());
    }

    let formattedWebinarName = capitalizeWords(domain);

    function capitalizesch(schoolName) {
      return schoolName.replace(/\b\w/g, char => char.toUpperCase());
    }

    let formattedsch = capitalizesch(schoolName);

    await slides.presentations.batchUpdate({
      presentationId: copyId,
      requestBody: {
        requests: [
          { replaceAllText: { containsText: { text: '{{Name}}' }, replaceText: name } },
          { replaceAllText: { containsText: { text: '{{SchoolName}}' }, replaceText: formattedsch } },
          { replaceAllText: { containsText: { text: '{{WebinarName}}' }, replaceText: formattedWebinarName } },
          { replaceAllText: { containsText: { text: '{{Date}}' }, replaceText: formattedDate } },
          { replaceAllText: { containsText: { text: '{{Dateto}}' }, replaceText: formattedtoDate } },
          { replaceAllText: { containsText: { text: '{{CERT-NUMBER}}' }, replaceText: certificateNumber } }
        ],
      },
    });

    const exportUrl = `https://www.googleapis.com/drive/v3/files/${copyId}/export?mimeType=application/pdf`;
    const response = await drive.files.export({
      fileId: copyId,
      mimeType: 'application/pdf',
    }, { responseType: 'stream' });

    const filename = `${name}_${certificateNumber}.pdf`;
await sendEmailWithAttachment(
  email,
  `GeniusHub Internship Completion Certificate`,
  `Dear ${name},<br><br>
   Greetings of the day!!<br><br>
   Thank you for participating in the GeniusHub Internship Program.We wish you all the best in your future endeavors.Please find your certificate attached.<br><br>
   Warm regards,<br><br>
   <b>Nisha Jain</b><br>
   <b>Internships Program Manager</b><br>
   <b>9873331785</b><br>
   <img src="https://upload.wikimedia.org/wikipedia/commons/a/a5/Instagram_icon.png" alt="Instagram" width="20" height="20" style="vertical-align:middle;">: 
   <a href="https://www.instagram.com/geniushub_internships" target="_blank"> 
     @geniushub_internships
   </a><br>
   <a href="https://www.geniushub.in/" target="_blank">https://www.geniushub.in</a>`,
  response.data,
  filename
);

    await drive.files.update({
      fileId: copyId,
      requestBody: { trashed: true }
    });

    await new Promise(resolve => setTimeout(resolve, 5000));
  }
}

// Function to generate certificates as a single PDF
async function generateCertificatesAsSinglePDF(sheetData, date, todate) {
  const templateId = process.env.TEMPLATE_ID;
  const folderId = process.env.FOLDER_ID;

  // Dynamically import PDFMerger
  const { default: PDFMerger } = await import('pdf-merger-js');
  const merger = new PDFMerger();

  for (let i = 1; i < sheetData.length; i++) {
    const row = sheetData[i];
    const name = row[0]?.toString() || '';
    const email = row[1]?.toString() || '';
    const schoolName = row[2]?.toString() || '';
    const domain = row[3]?.toString() || '';
    const certificateNumber = row[4]?.toString().toUpperCase() || '';
    const formattedDate = formatDateToReadable(new Date(date));
    const formattedtoDate = formatDateToReadable(new Date(todate));

    const copyFile = await drive.files.copy({
      fileId: templateId,
      requestBody: {
        name: `${name} - Certificate`,
        parents: [folderId]
      }
    });

    const copyId = copyFile.data.id;

    function capitalizeWords(domain) {
      return domain.replace(/\b\w/g, char => char.toUpperCase());
    }

    let formattedWebinarName = capitalizeWords(domain);

    function capitalizesch(schoolName) {
      return schoolName.replace(/\b\w/g, char => char.toUpperCase());
    }

    let formattedsch = capitalizesch(schoolName);

    await slides.presentations.batchUpdate({
      presentationId: copyId,
      requestBody: {
        requests: [
          { replaceAllText: { containsText: { text: '{{Name}}' }, replaceText: name } },
          { replaceAllText: { containsText: { text: '{{SchoolName}}' }, replaceText: formattedsch } },
          { replaceAllText: { containsText: { text: '{{WebinarName}}' }, replaceText: formattedWebinarName } },
          { replaceAllText: { containsText: { text: '{{Date}}' }, replaceText: formattedDate } },
          { replaceAllText: { containsText: { text: '{{Dateto}}' }, replaceText: formattedtoDate } },
          { replaceAllText: { containsText: { text: '{{CERT-NUMBER}}' }, replaceText: certificateNumber } }
        ],
      },
    });

    const exportUrl = `https://www.googleapis.com/drive/v3/files/${copyId}/export?mimeType=application/pdf`;
    const response = await drive.files.export({
      fileId: copyId,
      mimeType: 'application/pdf',
    }, { responseType: 'stream' });

    const filePath = path.join(__dirname, `temp_${name}_${certificateNumber}.pdf`);
    console.log(`Processing certificate for: ${name} - ${certificateNumber}`);
    const writeStream = fs.createWriteStream(filePath);
    response.data.pipe(writeStream);

    await new Promise(resolve => writeStream.on('finish', resolve));

    try {
      await merger.add(filePath);
      fs.unlinkSync(filePath);
    } catch (err) {
      console.error(`Failed to process file ${filePath}:`, err);
    }

    try {
      await drive.files.update({
        fileId: copyId,
        requestBody: { trashed: true }
      });
    } catch (err) {
      console.error(`Failed to update file ${copyId} status:`, err);
    }

    // Update progress incrementally
    fs.writeFileSync(logFilePath, JSON.stringify({
      progress: calculatePercentage(i, sheetData.length - 1),
      totalCertificates: sheetData.length - 1,
      generatedCount: i,
      generating: true
    }));
  }

  await merger.save(path.join(__dirname, 'certificates.pdf'));
}

// Function to generate certificates as a ZIP file
async function generateCertificates(sheetData, date, todate, updateGeneratedCount) {
  if (!Array.isArray(sheetData) || sheetData.length === 0) {
    throw new Error('No data found in the Google Sheet.');
  }

  const templateId = process.env.TEMPLATE_ID;
  const folderId = process.env.FOLDER_ID;
  let generatedCount = 0;

  const zipFilePath = path.join(__dirname, 'certificates.zip');
  const output = fs.createWriteStream(zipFilePath);
  const archive = archiver('zip', { zlib: { level: 9 } });

  // Ensure the ZIP archive is finalized properly
  output.on('close', () => {
    console.log(`Created zip file with ${archive.pointer()} total bytes.`);
  });

  // Log any errors that occur during archiving
  archive.on('error', (err) => {
    throw err;
  });

  // Pipe archive data to the output file
  archive.pipe(output);

  const totalCertificates = sheetData.length;
  
  for (let i = 0; i <totalCertificates; i++) {
    const row = sheetData[i];
    const name = row[0]?.toString() || '';
    const schoolName = row[2]?.toString() || '';
    const domain = row[3]?.toString() || '';
    const certificateNumber = row[4]?.toString().toUpperCase() || '';

    if (!name || !schoolName || !certificateNumber) {
      console.log(`Skipping row ${i + 1} due to missing data.`);
      continue;
    }

    const formattedDate = formatDateToReadable(new Date(date));
    const formattedtoDate = formatDateToReadable(new Date(todate));

    console.log(`Processing certificate for: ${name} - ${certificateNumber}`);

    try {
      const copyFile = await drive.files.copy({
        fileId: templateId,
        requestBody: {
          name: `${name} - Certificate`,
          parents: [folderId]
        }
      });

      const copyId = copyFile.data.id;

      function capitalizeWord(domain) {
        return domain.replace(/\b\w/g, char => char.toUpperCase());
      }
  
      let formattedWebinar = capitalizeWord(domain);

      function capitalizeschool(schoolName) {
        return schoolName.replace(/\b\w/g, char => char.toUpperCase());
      }
  
      let formattedschool = capitalizeschool(schoolName);

      await slides.presentations.batchUpdate({
        presentationId: copyId,
        requestBody: {
          requests: [
            { replaceAllText: { containsText: { text: '{{Name}}' }, replaceText: name } },
            { replaceAllText: { containsText: { text: '{{SchoolName}}' }, replaceText: formattedschool } },
            { replaceAllText: { containsText: { text: '{{WebinarName}}' }, replaceText: formattedWebinar } },
            { replaceAllText: { containsText: { text: '{{Date}}' }, replaceText: formattedDate } },
            { replaceAllText: { containsText: { text: '{{Dateto}}' }, replaceText: formattedtoDate } },
            { replaceAllText: { containsText: { text: '{{CERT-NUMBER}}' }, replaceText: certificateNumber } }
          ],
        },
      });

      const exportUrl = `https://www.googleapis.com/drive/v3/files/${copyId}/export?mimeType=application/pdf`;
      const response = await drive.files.export({
        fileId: copyId,
        mimeType: 'application/pdf',
      }, { responseType: 'stream' });

      const filePath = path.join(__dirname, `${name}_${certificateNumber}.pdf`);
      const writeStream = fs.createWriteStream(filePath);

      console.log(`Writing file: ${filePath}`);
      response.data.pipe(writeStream);

      await new Promise((resolve, reject) => {
        writeStream.on('finish', resolve);
        writeStream.on('error', reject);
      });

      console.log(`Adding file to ZIP: ${filePath}`);
      archive.file(filePath, { name: `${name}_${certificateNumber}.pdf` });

      try {
        await drive.files.update({
          fileId: copyId,
          requestBody: { trashed: true }
        });
      } catch (err) {
        console.error(`Failed to update file ${copyId} status:`, err);
      }

      try {
        fs.unlinkSync(filePath);
      } catch (err) {
        console.error(`Failed to delete file ${filePath}:`, err);
      }

      generatedCount++;
      if (updateGeneratedCount) {
        updateGeneratedCount(generatedCount);
      }

      // Update progress after processing each certificate
      fs.writeFileSync(logFilePath, JSON.stringify({
        progress: calculatePercentage(generatedCount, totalCertificates),
        totalCertificates,
        generatedCount,
        generating: true
      }));

    } catch (err) {
      console.error(`Error processing certificate for ${name} - ${certificateNumber}:`, err);
    }
  }

  console.log('Finalizing ZIP archive');
  await archive.finalize();
  
  // Final progress update to 100%
  fs.writeFileSync(logFilePath, JSON.stringify({
    progress: 100,
    totalCertificates,
    generatedCount,
    generating: false
  }));
}
async function getSheetData(sheetId, sheetName) {
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId: sheetId,
    range: sheetName,
  });
  return response.data.values;
}

// Helper function to format the date
function formatDateToReadable(date) {
  const monthNames = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
  ];

  const day = date.getDate();
  const year = date.getFullYear();
  const month = monthNames[date.getMonth()];

  let suffix = "th";
  if (day === 1 || day === 21 || day === 31) {
    suffix = "st";
  } else if (day === 2 || day === 22) {
    suffix = "nd";
  } else if (day === 3 || day === 23) {
    suffix = "rd";
  }

  return `${day}${suffix} ${month}, ${year}`;
}

// Helper function to calculate the percentage
function calculatePercentage(completed, total) {
  if (total === 0) return 0;
  return Math.round((completed / total) * 100);
}

async function sendEmailWithAttachment(to, subject, htmlContent, pdfStream, filename) {
  const mailOptions = {
    from: '"GeniusHub - Unleash your Genius" <Certificate@geniushub.in>',
    to,
    subject,
    html: htmlContent,
    attachments: [
      {
        filename,
        content: pdfStream,
        contentType: 'application/pdf'
      }
    ]
  };

  try {
    await transporter.sendMail(mailOptions);
    console.log(`Email sent successfully to ${to}`);
  } catch (error) {
    console.error(`Error sending email to ${to}:`, error);
  }
}

app.listen(port, () => {
  console.log(`Server is running on port ${port}`);
});
