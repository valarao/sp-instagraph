const QUOTE_URL = 'https://ca.finance.yahoo.com/quote/';


module.exports = async function navigate(page, ticker, exchange) {
    const exchangeC = convertExchange(exchange);
    await page.goto(`${QUOTE_URL}/${ticker}${exchangeC}`);
}

function convertExchange(exchange) {
    if (exchange === 'TO') {
        return '.TO';
    } else {
        return '';
    }
}