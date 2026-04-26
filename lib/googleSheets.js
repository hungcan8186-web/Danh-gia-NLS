// lib/googleSheets.js
// Shared Google Sheets API helper for Vercel serverless functions
// Uses service account credentials stored in Vercel environment variables

const { google } = require('googleapis');

const SPREADSHEET_ID = process.env.SPREADSHEET_ID;
const DATA_SHEET = 'DanhSachTheoDoi';

async function getSheetClient() {
  const auth = new google.auth.GoogleAuth({
    credentials: {
      client_email: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
      private_key: (process.env.GOOGLE_PRIVATE_KEY || '').replace(/\\n/g, '\n'),
    },
    scopes: [
      'https://www.googleapis.com/auth/spreadsheets',
      'https://www.googleapis.com/auth/drive',
    ],
  });
  return google.sheets({ version: 'v4', auth });
}

async function findUserRow(sheets, email) {
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: DATA_SHEET + '!A:A',
  });
  const rows = res.data.values || [];
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] && rows[i][0].toString().toLowerCase() === email.toLowerCase()) {
      return i + 1;
    }
  }
  return -1;
}

module.exports = { getSheetClient, findUserRow, SPREADSHEET_ID, DATA_SHEET };
