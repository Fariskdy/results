const express = require('express');
const app = express();
const { google } = require('googleapis');
const credentials = require('./himaya-course-ee66c11de0e1.json');
const spreadsheetId = '1H3v8tQ-INg_p0-CbqGhhbpt7wDhu0UHmCUVSb6nsWN4'; // Update with your spreadsheet ID

// Set up Google Sheets API client

// Set EJS as the view engine
app.set('view engine', 'ejs');

// Middleware to parse JSON and URL-encoded bodies
app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(express.static('public'));

const requestQueue = []; // Queue to store incoming requests
let isProcessing = false; // Flag to track if a request is currently being processed

// Function to process a single request
async function processRequest(req, res) {
  try {
    const admissionNumber = req.body.admissionNumber;

    // Create a new instance of the client inside the route handler
    const client = new google.auth.JWT(
      credentials.client_email,
      null,
      credentials.private_key,
      ['https://www.googleapis.com/auth/spreadsheets']
    );

    const sheets = google.sheets({ version: 'v4', auth: client });

    // Write the admission number to cell D8:E8 in the "IND STATUS" sheet
    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: 'D8:E8',
      valueInputOption: 'USER_ENTERED',
      resource: {
        values: [[admissionNumber]],
      },
    });

    // Read specific cell values from the first sheet ("IND STATUS")
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: 'A1:M52', // Update with your desired range
    });

    const values = response.data.values;

    // Extract specific cell values with error handling
    const name = values && values[9] && values[9][3] ? values[9][3] : 'N/A'; // Assuming the name is in D10:L10
    const term = values && values[11] && values[11][3] ? values[11][3] : 'N/A'; // Assuming the term is in D12:L12
    const college = values && values[12] && values[12][3] ? values[12][3] : 'N/A'; // Assuming the college is in D13:L13
    const admission = values && values[7] && values[7][3] ? values[7][3] : 'N/A';
    const qfaValue = values && values[19] && values[19][4] ? values[19][4] : 'N/A';
    const hfaValue = values && values[20] && values[20][4] ? values[20][4] : 'N/A';
    const ffaValue = values && values[21] && values[21][4] ? values[21][4] : 'N/A';
    const tfaValue = values && values[22] && values[22][4] ? values[22][4] : 'N/A';
    const qstValue = values && values[19] && values[19][5] ? values[19][5] : 'N/A';
const hstValue = values && values[20] && values[20][5] ? values[20][5] : 'N/A';
const fstValue = values && values[21] && values[21][5] ? values[21][5] : 'N/A';
const tstValue = values && values[22] && values[22][5] ? values[22][5] : 'N/A';
const qsaValue = values && values[19] && values[19][7] ? values[19][7] : 'N/A';
const hsaValue = values && values[20] && values[20][7] ? values[20][7] : 'N/A';
const fsaValue = values && values[21] && values[21][7] ? values[21][7] : 'N/A';
const tsaValue = values && values[22] && values[22][7] ? values[22][7] : 'N/A';

const qsastValue = values && values[19] && values[19][8] ? values[19][8] : 'N/A';
const hsastValue = values && values[20] && values[20][8] ? values[20][8] : 'N/A';
const fsastValue = values && values[21] && values[21][8] ? values[21][8] : 'N/A';
const tsastValue = values && values[22] && values[22][8] ? values[22][8] : 'N/A';

const qtoValue = values && values[19] && values[19][10] ? values[19][10] : 'N/A';
const htoValue = values && values[20] && values[20][10] ? values[20][10] : 'N/A';
const ftoValue = values && values[21] && values[21][10] ? values[21][10] : 'N/A';
const ttoValue = values && values[22] && values[22][10] ? values[22][10] : 'N/A';

const qpassValue = values && values[19] && values[19][11] ? values[19][11] : 'N/A';
const hpassValue = values && values[20] && values[20][11] ? values[20][11] : 'N/A';
const fpassValue = values && values[21] && values[21][11] ? values[21][11] : 'N/A';
const tpassValue = values && values[22] && values[22][11] ? values[22][11] : 'N/A';


const tofaValue = values && values[37] && values[37][4] ? values[37][4] : 'N/A';
const tosaValue = values && values[37] && values[37][7] ? values[37][7] : 'N/A';
const totaValue = values && values[37] && values[37].slice(10, 13).join(' ') ? values[37].slice(10, 13).join(' ') : 'N/A';


const grtotalValue = values && values[42] && values[42][3] ? values[42][3] : 'N/A';
const failsubValue = values && values[44] && values[44][3] ? values[44][3] : 'N/A';
const passubValue = values && values[44] && values[44][4] ? values[44][4] : 'N/A';
const gradeValue = values && values[45] && values[45].slice(3, 5).join(' ') ? values[45].slice(3, 5).join(' ') : 'N/A';
const statusValue = values && values[46] && values[46].slice(4, 6).join(' ') ? values[46].slice(4, 6).join(' ') : 'N/A';
const percentageValue = values && values[38] && values[38].slice(3, 5).join(' ') ? values[38].slice(3, 5).join(' ') : 'N/A';
const perfailedValue = values && values[39] && values[39].slice(3, 5).join(' ') ? values[39].slice(3, 5).join(' ') : 'N/A';
const dateValue = values && values[4] && values[4].slice(11, 13).join(' ') ? values[4].slice(11, 13).join(' ') : 'N/A';
// Render the EJS template and pass the extracted values
res.render('index', { name, term, college,admission, qfaValue, hfaValue, ffaValue, tfaValue, qstValue, hstValue, fstValue, tstValue,
  qsaValue, hsaValue, fsaValue, tsaValue, qsastValue, hsastValue, fsastValue, tsastValue, qtoValue, htoValue, ftoValue, ttoValue, qpassValue, hpassValue, fpassValue, tpassValue,
  tofaValue, tosaValue, totaValue, grtotalValue, failsubValue, passubValue, gradeValue, statusValue, perfailedValue, dateValue, percentageValue
       });

        // Process the next request in the queue
    if (requestQueue.length > 0) {
      const nextRequest = requestQueue.shift();
      await processRequest(nextRequest.req, nextRequest.res);
    } else {
      isProcessing = false; // Mark that no request is being processed
    }
  } catch (error) {
    console.error('Error:', error);
    res.status(500).send('An error occurred.');
  }
}


// Route to render the form page
app.get('/', (req, res) => {
  res.render('form');
});




// Route to handle form submission
app.post('/submit', async (req, res) => {
  try {
    requestQueue.push({ req, res }); // Add the incoming request to the queue

    // If no request is being processed, start processing the first request in the queue
    if (!isProcessing) {
      isProcessing = true;
      const nextRequest = requestQueue.shift();
      await processRequest(nextRequest.req, nextRequest.res);
    }
  } catch (error) {
    console.error('Error:', error);
    res.status(500).send('An error occurred.');
  }
});

const port = process.env.PORT || 3000;

app.listen(port, () => {
  console.log(`Server is running on port ${port}`);
});
