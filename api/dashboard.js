// api/dashboard.js
// GET /api/dashboard
// Get all user stats for dashboard (admin view)

const { getSheetClient, SPREADSHEET_ID, DATA_SHEET } = require('../lib/googleSheets');

module.exports = async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();

  if (req.method !== 'GET') {
    return res.status(405).json({ success: false, message: 'Method not allowed' });
  }

  try {
    const sheets = await getSheetClient();
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: DATA_SHEET + '!A:J',
    });

    const rows = response.data.values || [];
    const stats = [];

    // Skip header row (index 0)
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      if (row[0]) {
        stats.push({
          email: row[0] || '',
          name: row[1] || '',
          role: row[2] || '',
          school: row[3] || '',
          status: row[4] || '',
          updateTime: row[5] || '',
          scoreStr: row[7] || '',
          classification: row[8] || '',
        });
      }
    }

    return res.status(200).json({ success: true, list: stats });
  } catch (err) {
    console.error('dashboard error:', err);
    return res.status(500).json({ success: false, error: err.message });
  }
};
