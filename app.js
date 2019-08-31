const express = require('express');
const logger = require('morgan');
const bodyParser = require('body-parser');

const getJSON = require('./ImportExcel').getJSON

// Set up the express app
const app = express();

// Log requests to the console.
app.use(logger('dev'));

// Parse incoming requests data (https://github.com/expressjs/body-parser)
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({
    extended: false
}));

app.post('/api/:filename', async (req, res) => {
    let list = req.body.list;
    try {
        let result = await getJSON(req.params.filename,0,list)
        res.status(200).send(result)
    } catch (err) {
        console.log(err)
        res.status(500).send(err)
    }
})

// Setup a default catch-all route that sends back a welcome message in JSON format.
app.get('*', (req, res) => res.status(200).send({
    message: 'Welcome to the beginning of nothingness.',
}));

app.listen(3000, () => {
    console.log("Server run on localhost:3000");
})

module.exports = app;