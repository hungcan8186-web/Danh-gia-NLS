// api/submit-final.js
// POST /api/submit-final
// Submit final assessment - save to Google Sheets and generate report URL

const { getSheetClient, findUserRow, SPREADSHEET_ID, DATA_SHEET } = require('../lib/googleSheets');
const { google } = require('googleapis');

module.exports = async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();

  if (req.method !== 'POST') {
    return res.status(405).json({ success: false, message: 'Method not allowed' });
  }

  const { email, name, role, school, draftData, stats, evaluation } = req.body || {};
  if (!email || !name || !role) {
    return res.status(400).json({ success: false, message: 'Missing required fields' });
  }

  try {
    const sheets = await getSheetClient();
    const row = await findUserRow(sheets, email);
    const dateStr = new Date().toLocaleString('vi-VN');
    const jsonStr = JSON.stringify(draftData || []);
    const scoreStr = (stats && stats.totalScore != null) ? stats.totalScore + ' / ' + stats.maxScore : '';
    const classStr = (stats && stats.classification) ? stats.classification : '';

    // Build a Google Sheets link as the 'report URL'
    const sheetsUrl = 'https://docs.google.com/spreadsheets/d/' + SPREADSHEET_ID;

    const recordValues = [name, role, school, 'Hoan thanh', dateStr, jsonStr, scoreStr, classStr, sheetsUrl];

    if (row > -1) {
      await sheets.spreadsheets.values.update({
        spreadsheetId: SPREADSHEET_ID,
        range: DATA_SHEET + '!B' + row + ':J' + row,
        valueInputOption: 'RAW',
        resource: { values: [recordValues] },
      });
    } else {
      await sheets.spreadsheets.values.append({
        spreadsheetId: SPREADSHEET_ID,
        range: DATA_SHEET + '!A:J',
        valueInputOption: 'RAW',
        insertDataOption: 'INSERT_ROWS',
        resource: { values: [[email, ...recordValues]] },
      });
    }

    return res.status(200).json({ success: true, docUrl: sheetsUrl });
  } catch (err) {
    console.error('submit-final error:', err);
    return res.status(500).json({ success: false, error: err.message });
  }
};
