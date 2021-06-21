const { google } = require('googleapis');
const readXlsxFile = require('read-excel-file/node');

const getGoogleAuth = () => new google.auth.GoogleAuth({
  scopes: [
    'https://www.googleapis.com/auth/drive.readonly',
    'https://www.googleapis.com/auth/drive.metadata.readonly',
    'https://www.googleapis.com/auth/spreadsheets'
  ]
});

async function downloadFile(fileId, auth) {
  const drive = google.drive({ version: 'v3', auth });
  
  try {
    const file = await drive.files
      .get({ fileId, alt: 'media' }, { responseType: 'stream' });
    return file.data;
  } catch (error) {
    console.error(error);
    throw error;
  }
}

async function writeDestinationSheet (spreadsheetId, auth, sheetTitle, values) {
  const sheets = google.sheets({ version: 'v4', auth });

  const resource = {
    values
  };

  try {
    await sheets.spreadsheets.values.clear({
      spreadsheetId,
      range: sheetTitle,
    });
    
    await sheets.spreadsheets.values.append({
      spreadsheetId,
      range: `${sheetTitle}!A1`,
      valueInputOption: 'RAW',
      resource
    });
  } catch (error) {
    console.error(error);
    throw error;
  }
}

async function doProcess () {
  const auth = getGoogleAuth();
  const originFileId = '1N4V-qqp6q8exuYAZgZ79GT9GazsSCHUK';
  const destinationFileId = '1eblBeozGn1soDGXbOIicwyEDkUqNMzzpJoAKw84TTA4';

  const originFile = await downloadFile(originFileId, auth);  
  const originRows = await readXlsxFile(originFile, { sheet: 3 });

  await writeDestinationSheet(destinationFileId, auth, 'Reporte diario', originRows);
}

/**
 * Responds to any HTTP request.
 *
 * @param {!express:Request} req HTTP request context.
 * @param {!express:Response} res HTTP response context.
 */
exports.convert = async (req, res) => {
  try {
    await doProcess();
    res.status(200).end();
  } catch (error) {
    res.status(500).send(error);
  }
};
