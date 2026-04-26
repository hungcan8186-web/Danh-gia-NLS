// api/save-draft.js
// POST /api/save-draft
// Save user draft data to Google Sheets

const { getSheetClient, findUserRow, SPREADSHEET_ID, DATA_SHEET } = require('../lib/googleSheets');

module.exports = async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();

  if (req.method !== 'POST') {
    return res.status(405).json({ success: false, message: 'Method not allowed' });
  }

  const { email, name, role, school, draftData } = req.body || {};
  if (!email || !name || !role) {
    return res.status(400).json({ success: false, message: 'Missing required fields' });
  }

  try {
    const sheets = await getSheetClient();
    const row = await findUserRow(sheets, email);
    const dateStr = new Date().toLocaleString('vi-VN');
    const jsonStr = JSON.stringify(draftData || []);

    if (row > -1) {
      // Update existing row (columns B to G = name, role, school, status, date, json)
      await sheets.spreadsheets.values.update({
        spreadsheetId: SPREADSHEET_ID,
        range: DATA_SHEET + '!B' + row + ':G' + row,
        valueInputOption: 'RAW',
        resource: { values: [[name, role, school, 'Dang danh gia', dateStr, jsonStr]] },
      });
    } else {
      // Append new row
      await sheets.spreadsheets.values.append({
        spreadsheetId: SPREADSHEET_ID,
        range: DATA_SHEET + '!A:J',
        valueInputOption: 'RAW',
        insertDataOption: 'INSERT_ROWS',
        resource: { values: [[email, name, role, school, 'Dang danh gia', dateStr, jsonStr, '', '', '']] },
      });
    }

    return res.status(200).json({ success: true });
  } catch (err) {
    console.error('save-draft error:', err);
    return res.status(500).json({ success: false, error: err.message });
  }
};
