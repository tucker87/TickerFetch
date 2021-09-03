const { URL, URLSearchParams } = require('url');
const fetch = require('node-fetch')
const Excel = require('exceljs');
const fs = require('fs')

const startRow = 2

const excelPath = "D:/Home/Desktop/Stocksv2.xlsx"
const tickerSheetName = 'Stock Breakdown'
const tickerSymbolCol = 1
const tickerPriceCol = 8

const cryptoSheetName = 'Crypto Breakdown'
const cryptoSymbolCol = 1
const cryptoPriceCol = 8

let keys = {} 

fs.readFile('./keys.json', function(err, data) {
    keys = JSON.parse(data)
	console.log(keys)
})

const getSymbols = async (excelPath, sheetName, row, column) => {
	const tickers = []

	let workbook = new Excel.Workbook();
	await workbook.xlsx.readFile(excelPath);
	let worksheet = workbook.getWorksheet(sheetName);

	let r = worksheet.getRow(row).values;
	while (r[column] != undefined) {
		tickers.push(r[column]);
		row++;
		r = worksheet.getRow(row).values;
	}
	return tickers;
}

const getQuotes = async tickers => {
	const getQuotesUrl = new URL("https://apidojo-yahoo-finance-v1.p.rapidapi.com/market/v2/get-quotes")
	const params = { region: 'US', symbols: tickers }
	getQuotesUrl.search = new URLSearchParams(params).toString();
		
	let response = await fetch(getQuotesUrl, {
		"method": "GET",
		"headers": {
			"x-rapidapi-key": keys.yahoo,
			"x-rapidapi-host": "apidojo-yahoo-finance-v1.p.rapidapi.com"
		}
	}).catch(err => {
		console.error(err);
	});

	let json = await response.json()

	let results = json.quoteResponse.result
	let values = results.map(r => ({ symbol: r.symbol, price: r.postMarketPrice || r.regularMarketPrice}))
	return values
}

const writeQuotes = async (excelPath, sheetName, symbolColumn, priceColumn, quotes) => {
	let row = startRow;
	let workbook = new Excel.Workbook();
	await workbook.xlsx.readFile(excelPath);
	let worksheet = workbook.getWorksheet(sheetName);

	let r = worksheet.getRow(row).values;
	while (r[priceColumn] != undefined) {
		let symbol = worksheet.getRow(row).getCell(symbolColumn).value
		worksheet.getRow(row).getCell(priceColumn).value = quotes.find(q => symbol.includes(q.symbol)).price;
		row++;
		r = worksheet.getRow(row).values;
	}
	
	return workbook.xlsx.writeFile(excelPath)
}

const getCrypto = async coins => {
	const getCryptoQuotesUrl = new URL('https://pro-api.coinmarketcap.com/v1/cryptocurrency/quotes/latest')
	const params = { symbol: coins }
	getCryptoQuotesUrl.search = new URLSearchParams(params).toString();
	
	let response = await fetch(getCryptoQuotesUrl, {
		method: "GET",
		headers: {
			'X-CMC_PRO_API_KEY': keys.coinMarketCap
		}
	}).catch(err => {
		console.error(err);
	});

	let json = await response.json()

	let values = Object.values(json.data).map(d => ({ symbol: d.symbol, price: d.quote.USD.price }))
	return values
}

(async function() {
	

	let tickers = await getSymbols(excelPath, tickerSheetName, startRow, tickerSymbolCol)
	
	let quotes = await getQuotes(tickers)
	console.log(quotes)
	
	await writeQuotes(excelPath, tickerSheetName, tickerSymbolCol, tickerPriceCol, quotes)

	//Crypto
	let coins = await getSymbols(excelPath, cryptoSheetName, startRow, cryptoSymbolCol)

	//filter out bad columns
	coins = coins.filter(c => !c.includes('('))

	quotes = await getCrypto(coins.join(','))
	console.log(quotes)

	await writeQuotes(excelPath, cryptoSheetName, cryptoSymbolCol, cryptoPriceCol, quotes)
})()