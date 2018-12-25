const express = require('express');
const bodyParser = require('body-parser');
const app = express();
app.use(bodyParser.json());

const getData = require('./scripts/get-data')


app.get('/', (req, res) => {
    res.send('InstaGraph v1');
});

app.post('/instagraph', async (req, res) => {
    const ticker = 'OSB'; // prompt('Ticker:');
    const exchange = 'NASDAQ'; // prompt('Exchange:');

    const { msg, err } = await getData(ticker, exchange);
    if (err) {
        res.status(400).send(err.message);
        return;
    }
    
    res.send(msg);
});



app.listen(3000, () => console.log('Stock Price InstaGraph Running'));
