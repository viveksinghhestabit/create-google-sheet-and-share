const express = require('express');
const router = express.Router();
const { authorize, sheet, googleDrive } = require('./GoogleAuth/auth.js');
const path = require('path');
const fs = require('fs');
const stream = require('stream');
const { google } = require('googleapis');


router.get('/', (req, res) => {
    res.send('Hello World!');
});

router.post('/create-sheet', async (req, res) => {
    const { sheetName, email } = req.body;
    console.log(sheetName, email);
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
    const headerRow = ['Name', 'Email', 'Phone', 'Country'];
    const headerRowIndex = 0;
    const headerRowValueRange = {
        values: [headerRow],
    };
    const params = {
        spreadsheetId: spreadsheetResponse.data.spreadsheetId,
        range: 'A1',
        valueInputOption: 'RAW',
        resource: headerRowValueRange,
    };
    const data = [
        ['John', 'john@gmail.com', '1234567890', 'USA'],
        ['Jane', 'jane@gmail.com', '0987654321', 'UK'],
    ];
    await sheets.spreadsheets.values.update(params);
    const valueRangeBody = {
        values: data,
    };
    const valueRangeParams = {
        spreadsheetId: spreadsheetResponse.data.spreadsheetId,
        range: 'A2',
        valueInputOption: 'RAW',
        resource: valueRangeBody,
    };
    await sheets.spreadsheets.values.update(valueRangeParams);
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

router.post('/create-sheet-with-data', async (req, res) => {
    const { sheetName, email } = req.body;
    const client = await authorize();
    const sheets = await sheet(client);
    const claimPath = path.join(process.cwd(), 'claim.json');
    const content = await fs.readFile(claimPath);
    const claims = JSON.parse(content);
    // return res.send(claims);
    const sheetResource = {
        properties: {
            title: sheetName || 'Sheet-' + Date.now(),
        },
    };
    const claim = [claims];
    const data = [];
    data.push(['CONTENTS REPORT']);
    data.push(['Insured', claim[0]?.client?.name]);
    data.push(['Address', claim[0]?.location]);
    data.push(['City', '']);
    data.push(['State', '']);
    data.push(['Zip', '']);
    data.push(['Type of Loss', claim[0]?.lossType]);
    data.push(['Sales Tax (future field)', claim[0]?.salesTax]);
    data.push(['Content Claim Replacement Value', '$' + claim[0]?.replacementCostValue]);
    data.push(['Content Claim Actual Cash Value (Number Provided by Insurance Carrier)', '$' + claim[0]?.actualCashValue]);
    data.push(['Insurance Company', claim[0]?.insuranceProvider]);
    data.push(['Adjuster', claim[0]?.adjustor?.name]);
    data.push(['To Be completed by Carrie Adjuster']);

    data.push(['No.', 'Room', 'Type', 'Description', 'Brand', 'Model', 'Condition', 'Item Age (Year)', 'Item Age (Months)', 'Qty', 'Replacement / Clean Cost', 'Total Replacement / Clean Cost', 'Tax', 'Depreciation', 'Actual Cash Value', 'Original Image', 'Replacement Source Link', 'Image Link to Replacement Image', 'Confidence', 'Receipt (link)']);
    let counter = 1;
    claim[0].rooms.forEach((room) => {
        let s_no = counter;
        data.push([s_no, room?.name, room?.images?.content, room?.images?.description, room?.images?.brand || '', room?.images?.model || '', room?.images?.conditions, room?.images?.year, room?.images?.month, room?.images?.quantity, room?.images?.responsePrice, room?.images?.responsePrice, '', '', '', room?.images?.selectedImageUrl, '', '', '', '']);
        counter++;
    });
    data.push(['Total Contents', '$' + claim[0]?.replacementCostValue]);
    const resource = {
        values: data,
    };
    const range = 'A1:W' + data.length;
    const valueInputOption = 'RAW';
    const requestBody = { valueInputOption, resource };
    const spreadsheetResponse = await sheets.spreadsheets.create({
        resource: sheetResource,
        fields: 'spreadsheetId',
    });
    const spreadsheetId = spreadsheetResponse.data.spreadsheetId;
    const params = {
        spreadsheetId: spreadsheetId,
        range: range,
        valueInputOption: 'RAW',
        resource: resource,
    };
    await sheets.spreadsheets.values.update(params);
    const resource1 = {
        role: 'writer',
        type: 'user',
        emailAddress: email,
    };
    const result = {
        spreadsheetId: spreadsheetId,
        sheetName: sheetName,
        sheetUrl: `https://docs.google.com/spreadsheets/d/${spreadsheetId}`,
        email: email,
    };
    return res.send(result);

});

router.post('/import-excel-sheet-in-google-sheet', async (req, res) => {
    const { sheetName, email } = req.body;
    const client = await authorize();
    const drive = await googleDrive(client);

    var folderId = process.env.FOLDER_ID || '';
    if (folderId === '') {
        const folderMetadata = {
            name: 'MyExcelFolder',
            mimeType: 'application/vnd.google-apps.folder',
        };
        const folderExists = await drive.files.list({
            q: "mimeType='application/vnd.google-apps.folder' and name='MyExcelFolder'",
            fields: 'files(id, name)',
        });
        if (folderExists.data.files.length > 0) {
            folderId = folderExists.data.files[0].id;
        } else {
            const folderResponse = await drive.files.create({
                resource: folderMetadata,
                fields: 'id',
            });
            folderId = folderResponse.data.id;
        }
    }

    const sheetPath = path.join(process.cwd() + '/mern', './../sheet.xlsx');
    console.log('sheetPath', sheetPath);
    const file = await fs.readFileSync(sheetPath);
    const fileMetadata = {
        name: 'sheet.xlsx',
        mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        parents: [folderId],
    };
    const media = {
        mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        body: stream.PassThrough().end(file),
    };

    const fileResponse = await drive.files.create({
        resource: fileMetadata,
        media: media,
        fields: 'id',
    });
    const fileId = fileResponse.data.id;
    const resource = {
        role: 'writer',
        type: 'user',
        emailAddress: email,
    };
    await drive.permissions.create({
        fileId: fileId,
        resource: resource,
        fields: 'id',
    });
    const result = {
        fileId: fileId,
        sheetName: sheetName,
        sheetUrl: `https://docs.google.com/spreadsheets/d/${fileId}`,
        email: email,
    };

    return res.send(result);
});

module.exports = router;