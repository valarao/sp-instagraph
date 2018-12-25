const puppeteer = require('puppeteer');
const navigate = require('./navigate');

async function getData(ticker, exchange) {

    const browser = await puppeteer.launch();
    const page = await browser.newPage();

    try {
        await navigate(page, ticker, exchange);
    } catch (e) {
        return { err: e};
    }
}


module.exports = getData;