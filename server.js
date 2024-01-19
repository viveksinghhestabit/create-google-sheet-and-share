require('dotenv').config();
const serverless = require('serverless-http');
const bodyParser = require('body-parser');
const express = require('express');
const cors = require('cors');
const app = express();
const Routes = require('./src/Routes');


app.use(cors());
app.use(
	bodyParser.json({
		limit: '255mb',
		extended: true,
	}),
);
app.use(
	bodyParser.urlencoded({
		limit: '255mb',
		extended: true,
	}),
);
app.use(Routes);

app.listen(5000, () => {
	console.log('Server is running...');
});

module.exports.handler = serverless(app);
