// api/load-draft.js
// GET /api/load-draft?email=xxx
// Load user draft data from Google Sheets

const { getSheetClient, findUserRow, SPREADSHEET_ID, DATA_SHEET } = require('../lib/googleSheets');

module.exports = async function handler(req, res) {
  // CORS headers
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();

  if (req.method !== 'GET') {
    return res.status(405).json({ success: false, message: 'Method not allowed' });
  }

  const { email } = req.query;
  if (!email) {
    return res.status(400).json({ success: false, message: 'Missing email parameter' });
  }

  try {
    const sheets = await getSheetClient();
    const row = await findUserRow(sheets, email);

    if (row < 0) {
      return res.status(200).json({ success: false, message: 'Khong tim thay du lieu cu.' });
    }

    // Get row data (columns A to J = 1 to 10)
    const rangeRes = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: DATA_SHEET + '!A' + row + ':J' + row,
    });

    const data = (rangeRes.data.values || [[]])[0];
    return res.status(200).json({
      success: true,
      email: data[0] || '',
      name: data[1] || '',
      role: data[2] || '',
      school: data[3] || '',
      status: data[4] || '',
      draftJson: data[6] || null,
    });
  } catch (err) {
    console.error('load-draft error:', err);
    return res.status(500).json({ success: false, error: err.message });
  }
};
