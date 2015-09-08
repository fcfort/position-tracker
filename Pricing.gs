/**
 * @OnlyCurrentDoc Limits the script to only accessing the current spreadsheet.
 */
 
/*********************
 * PRICE QUERIES
 *********************/

/**
 * Returns the price of a ticker on a given date
 *
 * @param {String} ticker Stock ticker 
 * @param {Date} date Date to query prices on
 * @return {Number} Price
 * @customfunction
 */
function TICKER_PRICE(ticker, date) {
  return getPrice(ticker, date);
}

function getPrice(ticker, date) {
  var priceMap = getPrices([ticker],date);
  if(ticker in priceMap) {
    return priceMap[ticker];
  }
  
  return 0;
}

function _updateDictionary(old_dict, new_dict) { 
  for (var key in new_dict) {
    if (!old_dict.hasOwnProperty(key)) {
      old_dict[key] = new_dict[key];
    }
  }
}

function getPrices(tickers, date) {
  var priceMap = {};
  
  if(tickers.length == 0) { return priceMap; }
  
  // First special case out cash prices
  var cashTicker = 'USD'
  var index = tickers.indexOf(cashTicker)
  if (index >= 0) {
    tickers.splice(index, 1);
    priceMap[cashTicker] = 1
  }
  
  // Check Yahoo first
  _updateDictionary(priceMap, getYahooPrices(tickers,date));
 
  // If any left, look them up against the table data
  var missingTickers = getMissingKeys(tickers, priceMap);
  
  if(missingTickers.length > 0) {
    Logger.log('Got missing tickers %s', missingTickers);
  
    for(var len = missingTickers.length, i = 0; i < len; i++) {
      var ticker = missingTickers[i];
      priceMap[ticker] = getTablePrice(ticker, date);
    }
  }
  
  return priceMap;
}

function getMissingKeys(keys, map) {
  var missingKeys = [];
  for(var len = keys.length, i = 0; i < len; i++) {
    if(!(keys[i] in map)) {
      missingKeys.push(keys[i])
    }
  }
  return missingKeys;
}

function getTablePrice(ticker, date) {
  Logger.log('Retrieving price from table data for ticker %s and date %s', ticker, date);
  var priceData = getPriceData();
  for (var len = priceData.length, i=0; i<len; ++i) {
    //Logger.log('At row %s, %s, %s, next row %s, %s',row_ticker, row_date, row_price, next_ticker, next_date);
    var row_ticker = priceData[i][0];
    var row_date = priceData[i][1];
    var row_price = priceData[i][2];
    
    var next_ticker = "";
    var next_date = 0;
    
    if(i < len - 1) {
      next_ticker =  priceData[i+1][0];
      next_date = priceData[i+1][1];
    }
    
    // Logger.log('ticker %s == %s, date %s == %s', ticker, row_ticker, date, row_date);
    if(ticker === row_ticker) {
      //Logger.log('Found ticker, row %s, %s, %s, next row %s, %s',row_ticker, row_date, row_price, next_ticker, next_date);
      // handles dates before price data and matching dates
      if(+date <= +row_date) {
        return row_price;
      }
      // handles dates after all price points
      if(next_ticker != row_ticker || (next_ticker === row_ticker && +next_date > +date)) {
        return row_price;
      }
    }
  }
  
  return;
}


/* 
 * Queries Yahoo finance API for historical prices for a given list of tickers
 *
 * @param {Array} tickers A list of tickers
 * @param {Date} date Date to query for. Must be a date the markets were open
 * @return A map of ticker to adjusted close prices
 */
function getYahooPrices(tickers, date) {
  if(tickers.length == 0) {
    return {};
  }
  
  Logger.log('Getting prices for %s for date %s', tickers, date)
  var yql_query = createPriceQuery(tickers, getDaysAgo(date, 5), date);
  var url = 'http://query.yahooapis.com/v1/public/yql?q=' + encodeURI(yql_query) + 
    '&format=json' + '&diagnostics=false' + '&env=store://datatables.org/alltableswithkeys' + '&callback=';
  Logger.log(url);
  var data = getJsonData(url)
  Logger.log('Got data back %s', data);
  /*
    {"query":{
      "count":2,"created":"2015-06-01T14:21:43Z","lang":"en-US",
      "results":{
        "quote":[
          {"Symbol":"GOOG","Adj_Close":"548.90253"},
          {"Symbol":"AAPL","Adj_Close":"105.42101"}]}}}
  */  
  if(!data || !data.query.results) {
    Logger.log("Unable to find prices for " + tickers + " and date " + date + ' Got JSON data ' + data);
    return {};
  } 
  
  Utilities.sleep(1000); // avoid Yahoo errors
  
  // adj close accounts for corporate actions on the security.
  // better for historical analysis 
  var priceList = [];
  if(data.query.count == 1) {
    priceList = [data.query.results.quote];
  } else {
    priceList = data.query.results.quote;
  }
  Logger.log('Got quotes %s', priceList);

  var priceMap = {};
  
  for(var len = priceList.length, i = 0; i < len; i++) {
    var quote = priceList[i];
    var symbol = quote['Symbol'];
    var date = quote['Date'];
    var price = quote['Adj_Close'];
    if(symbol in priceMap) {
      if(date > priceMap[symbol]['date']) {
        priceMap[symbol]['price'] = price;
        priceMap[symbol]['date'] = date;
      }
    } else {
      priceMap[symbol] = {'price': price, 'date': date}; 
    }
  }
  
  // maintain old dict format of symbol -> price
  for(var k in priceMap) {
    priceMap[k] = priceMap[k]['price'];
  }
  
  return priceMap;
}

function getDaysAgo(date, numDaysAgo) {
  var a = new Date(date);
  a.setDate(date.getDate()-numDaysAgo);
  return a;
}

function getYahooPrice(ticker, date) { 
  var priceMap = getYahooPrices([ticker], date);
  return priceMap[ticker];
}

function getJsonData(url) {
  var response = UrlFetchApp.fetch(url);
  var json = response.getContentText();
  var data = JSON.parse(json);
  return data;
}


// http://stackoverflow.com/questions/25097779/getting-stocks-historical-data
function createPriceQuery(tickers, startDate, endDate) {
  endDate = endDate || startDate;

  var tickerStrings = [];
  for(var len = tickers.length, i = 0; i < len; i++) {
    tickerStrings.push('"' + tickers[i] + '"');
  }
  
  var yql_query = 'select Symbol, Adj_Close, Date from yahoo.finance.historicaldata where symbol in (' + 
    tickerStrings.join(',') + ') and startDate = "' + dateToString(startDate) + '" and endDate = "' + dateToString(endDate) + '"';
  Logger.log(yql_query);
  return yql_query;
}


// http://stackoverflow.com/questions/25097779/getting-stocks-historical-data
function createPriceQueryOld(ticker, date) {
 var date_string = dateToString(date)
 var yql_query = 'select Adj_Close from yahoo.finance.historicaldata where  symbol = "' + 
   ticker + '" and startDate = "' + date_string + '" and endDate = "' + date_string + '"';
 Logger.log(yql_query);
 return yql_query;
}


function dateToString(date) {
  var mon = padDate(date.getMonth()+1);
  var day = padDate(date.getDate());  
  return date.getYear() + '-' + mon + '-' + day;
}


function padDate(number) {
  return number < 10 ? '0' + number : number;
}


function getCurrentYahooPrice(ticker) { 
  getCurrentYahooPrices([ticker])[ticker];
}

function getCurrentYahooPrices(tickers) {
  var yql_query = createCurrentPriceQuery(tickers); 
  
  Utilities.sleep(1000); // avoid Yahoo errors  
  
  var data = _getUrlData(_getYahooQueryUrl(yql_query));
  
  if(!data) { return {}; }
  
  var priceList = [];
  if(data.query.count == 1) {
    priceList = [data.query.results.quote];
  } else {
    priceList = data.query.results.quote;
  }
  Logger.log('Got quotes %s', priceList);
  
  var priceMap = {};
  
  for(var len = priceList.length, i = 0; i < len; i++) {
    priceMap[priceList[i]['Symbol']] = priceList[i]['Ask'];
  }
  
  return priceMap;
}

function _getYahooQueryUrl(yql_query) {
  var url = 'http://query.yahooapis.com/v1/public/yql?q=' + encodeURI(yql_query) + 
    '&format=json' + '&diagnostics=false' + '&env=store://datatables.org/alltableswithkeys' + '&callback=';
  return url;
}

function _getUrlData(query_url) {
  Logger.log(query_url);
  var response = UrlFetchApp.fetch(query_url);
  Logger.log(response);
  var json = response.getContentText();
  var data = JSON.parse(json);  
  return data;
}

// http://stackoverflow.com/questions/25097779/getting-stocks-historical-data
function createCurrentPriceQuery(tickers) { 
  var tickerStrings = [];
  for(var len = tickers.length, i = 0; i < len; i++) {
    tickerStrings.push('"' + tickers[i] + '"');
  }
  var yql_query = 'select Symbol, Ask from yahoo.finance.quotes where symbol in (' + tickerStrings.join(',') + ')';
  Logger.log(yql_query);
  return yql_query;
}


function getCurrentPrice(ticker, date) {
  // https://query.yahooapis.com/v1/public/yql?q=select%20Ask%20from%20yahoo.finance.quotes%20where%20symbol%20in%20(%22YHOO%22%2C%22AAPL%22%2C%22GOOG%22%2C%22MSFT%22)&format=json&env=store%3A%2F%2Fdatatables.org%2Falltableswithkeys&callback=
  YAHOO_API = 'https://query.yahooapis.com/v1/public/yql?q='
  QUERY = 'select%20Ask%20from%20yahoo.finance.quotes%20where%20symbol%20in%20(%22YHOO%22%2C%22AAPL%22%2C%22GOOG%22%2C%22MSFT%22)'
  PARAMETERS = '&format=json&env=store%3A%2F%2Fdatatables.org%2Falltableswithkeys&callback='  
  url = YAHOO_API + QUERY + PARAMETERS
  var response = UrlFetchApp.fetch(url);
  Logger.log(response);
  var json = response.getContentText();
  var data = JSON.parse(json);
  Logger.log(data);
  /*  {query={
         results={
             quote=[{Ask=43.35}, {Ask=131.890}, {Ask=539.50}, {Ask=47.60}]
         }, count=4, created=2015-05-27T21:56:49Z, lang=en-US}} 
  */
  if(data) {
    var quotes = data.query.results.quote;
    var count = data.query.count;
    for(var i = 0; i < count; i++) {
       Logger.log(quotes[i]);
    }
  }
  return 0;
}


/*********************
 * TESTS
 *********************/ 

function testTablePrice() {
  var d = new Date(2013,3-1,29); // 2013-03-29

  Logger.log("Got %s for date %s", getTablePrice('SSgA:MidCap',d), d);

}

function testGetDaysAgo() {
  var d = new Date(2014,11-2,28);
  Logger.log("Date %s, five days ago %s", d, getDaysAgo(d,5));
}

function testYahooPrice() {
  var price = getYahooPrice("AAPL",new Date(2014,11-2,28));
  Logger.log("Got price for ticker %s", price);
}


function testGetPrice() {
  var price = getPrice("GOOG",new Date(2014,11-2,28));
  Logger.log("Got price for ticker %s", price);
  var price = getPrice("CIT1659",new Date(2014,11-2,28));
  Logger.log("Got price for ticker %s", price);
}


function testGetPrices() {
  var price = getPrices(["GOOG","AAPL","CIT1659"],new Date(2014,11-2,28));
  Logger.log("Got price for ticker %s", price);
}


function testYahooPrices() {
  var price = getYahooPrices(["GOOG","AAPL","CIT1659"],new Date(2014,11-2,28));
  Logger.log("Got price for ticker %s", price);
}

function testTICKER_PRICE() {
  //Logger.log(TICKER_PRICE('VBTLX',new Date(2015,7-1,15)));
  //Logger.log(TICKER_PRICE('VBTLX',new Date(2015,7-1,19)));
  Logger.log(TICKER_PRICE('USD',new Date(2015,7-1,19)));
  
}
