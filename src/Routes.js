const express = require('express');
const router = express.Router();
const { authorize, sheet, googleDrive } = require('./GoogleAuth/auth.js');

router.get('/', (req, res) => {
    res.send('Hello World!');
});

router.post('/create-sheet', async (req, res) => {
    const { sheetName, email } = req.body;
    const client = await authorize();
    const sheets = await sheet(client);
    const spreadsheetResource = {
        properties: {
            title: sheetName || 'Sheet-' + Date.now(),
        },
    };
    const spreadsheetResponse = await sheets.spreadsheets.create({
        resource: spreadsheetResource,
        fields: 'spreadsheetId',
    });
    const spreadsheetId = spreadsheetResponse.data.spreadsheetId;

    const resource = {
        role: 'writer',
        type: 'user',
        emailAddress: email,
    };
    const drive = await googleDrive(client);
    await drive.permissions.create({
        fileId: spreadsheetId,
        resource: resource,
        fields: 'id',
    });
    const result = {
        spreadsheetId: spreadsheetId,
        sheetName: sheetName,
        sheetUrl: `https://docs.google.com/spreadsheets/d/${spreadsheetId}`,
        email: email,
    };
    return res.send(result);
});

router.post('/share-sheet-with-email', async (req, res) => {
    const { sheetId, email } = req.body;
    const client = await authorize();
    const resource = {
        role: 'writer',
        type: 'user',
        emailAddress: email,
    };
    const drive = await google.drive({ version: 'v3', auth: client });
    const permission = await drive.permissions.create({
        fileId: sheetId,
        resource: resource,
        fields: 'id',
    });
    return res.send(permission);
});

module.exports = router;