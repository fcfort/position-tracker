/**
 * @OnlyCurrentDoc Limits the script to only accessing the current spreadsheet.
 */

/**
 * Returns the sum of positions for the ticker and date for the given
 * account.
 *
 * @param {Array} range Source data 
 * @param {Date} date Date to sum up to 
 * @param {String} ticker Ticker
 * @param {String} account Account (optional)
 * @return {Number} Sum
 * @customfunction
 */
function POSITION(range, date, ticker, account) {
  if (!(date instanceof Date)) {
    throw 'Invalid: Date input required';
  }
  
  var sum = 0;
  
  for (var i = 0; i < range.length; i++) {   
    row_data = parseRow(range[i])
    
    if(ticker.equals(row_data.ticker) && row_data.trade_date <= date) {      
      if(account && row_data.account != account) {
        continue;
      }      
      sum += row_data.quantity
    }
  }
  
  return sum;
}

/**
 * Returns the total value of positions for a given 
 * account as of a given date.
 *
 * @param {Array} range Source data 
 * @param {Date} date Date to sum up to 
 * @param {String} account Account (optional)
 * @return {Number} Total amount
 * @customfunction
 */
function POSITION_VALUE(range, date, account) {  
  var tickers = [];
  
  for (var i = 0; i < range.length; i++) {   
    row_data = parseRow(range[i]);
    
    if(account && row_data.account!= account) {
      continue
    }
    
    tickers.push(row_data.ticker)
  }
  
  var total = 0;
  
  for(var i = 0; i < tickers.length; i++) {
    sum = POSITION(range, date, tickers[i], account);
    total += getPrice(tickers[i], date)
  }
  
  return total;

}

/* 
 * Calculates total contributions made to our investments accounts made between
 * from_date and to_date
 *
 * @param {Array} array Source data 
 * @param {Date} from_date Start date, greater than or equal
 * @return {Date} to_date End date, less than
 * @customfunction
 */
function CONTRIBUTIONS(from_date, to_date) {
  var trans = getTransactions();
  
  // filter
  for(var len = trans.length, i = 0; i < len; i++) {
    var tran = trans[i];
  }
  
  
  // group

  var grouped_trans = groupTransactions(getTransactions(), ['account', 'ticker'], from_date, to_date);
  
  var total = 0;
  
  for(var len = grouped_trans.length, i = 0; i < len; i++) {
      
  }
}

function testQuery() {
  var query = new google.visualization.Query(getTransactions());
  query.setQuery('select dept, sum(salary) group by dept');
  query.send(handleQueryResponse);
}

/*
 * Groups transaction data based on the keys passed in between the 
 * dates specified.
 */
function groupTransactions(trans, key_list, from_date, to_date) {
  var grouped_results = {};
  
  for(var len = trans.length, i = 0; i < len; i++) {
    var row = parseRow(trans[i]);
    
    if(row.date < from_date || row.date >= to_date ) {
      continue;
    }
    
    // Create key from row
    var key_parts = [];
    for(var len = key_list.length, i = 0; i < len; i++) {
      key_parts.push(row[key_list[i]]);
    }
    var key = key_parts.join('|');
    
    if(!(key in grouped_results)) {
      grouped_results[key] = {
        quantity: row.quantity,
        amount: row.amount,
      };
    } else {
      grouped_results[key].quantity += row_data.quantity;
      grouped_results[key].amount += row_data.amount;
    }
  }
  
}

/* 
 * Takes in a transaction array and outputs
 * an array grouped by account/ticker up to the date
 * specified.
 * @param {Array} array Source data 
 * @param {Date} date Date to sum up to 
 * @return {Array} Grouped quantities
 * @customfunction
 */
function SUM_TO_DATE(array, date) {
  var grouped_results = {};

  Logger.log('Got %s transactions', array.length);
  for (var i = 0; i < array.length; i++) {   
    var row_data = parseRow(array[i]);
    if(row_data.trade_date > date) {
      continue;
    }
    
    var key = row_data.account + '_' + row_data.ticker;
    if(key in grouped_results) {
      grouped_results[key] += row_data.quantity;      
    } else {
      grouped_results[key] = row_data.quantity;      
    }
    
  }
  
  var output_array = [];
  for (var key in grouped_results) {
    if (grouped_results.hasOwnProperty(key)) {
      if(grouped_results[key] > 0.00001) {
        key_vals = key.split('_');
        output_array.push([key_vals[0], key_vals[1], grouped_results[key]]);
      }
    }
  }
  return output_array;
}

/* 
 * Takes in a transaction array and outputs
 * an array grouped by account/ticker up to the date
 * specified.
 * @param {Array} array Source data 
 * @param {Date} date Date to sum up to 
 * @return {Number} Total value in dollars
 * @customfunction
 */
function TOTAL_VALUE(array, date) {
  var grouped_results = SUM_TO_DATE(array, date);
  var total_value = 0;
  Logger.log('Got %s grouped results', grouped_results.length);
  
  var tickerSet = {};
  for (var i = 0; i < grouped_results.length; i++) { 
    tickerSet[grouped_results[i][1]] = true;
  }
  var priceMap = getPrices(Object.keys(tickerSet), date);
  Logger.log('Got price map %s', priceMap);
  
  for (var i = 0; i < grouped_results.length; i++) {   
    var holding = grouped_results[i];
    var ticker = holding[1];
    var quantity = holding[2];
    var price = priceMap[ticker];
    
    if(typeof price === 'undefined') {
      throw new Error("No price found for ticker:" + ticker);
    }
    
    var amount = quantity * price;
    
    Logger.log('for ticker %s, got quantity %s and price %s for total %s',
               ticker, quantity, price, amount);
    total_value += amount;
  }
  return total_value;
}


/*********************
 * SHEET DATA
 *********************/


function getTransactions() {
  var array = SpreadsheetApp.getActiveSpreadsheet().getRangeByName('Transactions').getValues();
  
  var trans = [];
  
  for(var len = array.length, i = 0; i < len; i++) {
    var tran = parseRow(array[i]);
    if(tran.account != '') {      
      trans.push(tran);
    }
  }
  
  return trans;
}

function testGetTransactions() {
  Logger.log('Got %s transactions', getTransactions().length);
}


function getPriceData() {
  var range = SpreadsheetApp.getActiveSpreadsheet().getRangeByName('PriceData');
  return range.getValues()
}


/**
 * A function that takes a Vanguard transaction row
 * and returns the relevant data from it.
 *
 * @param {Array} Row data
 * @return {Object} Parsed row data object
 */
function parseRow(row) {
  var row_account = row[0];
  var row_trade_date = row[2];   
  var row_ticker = row[3];
  var row_transaction_type = row[5];
  var row_quantity = row[6];
  var row_amount = row[9];
      
  // Fix en dash negative symbol in Vanguard transaction listings
  if(typeof row_quantity == 'string' && /^– /.test(row_quantity)) {
    row_quantity = +row_quantity.replace(/– /g,"-").replace(/,/g,"");
  }
  
  if(typeof row_quantity != 'number') {        
    row_quantity = 0;
  }
  
  return {  
    account: row_account,
    ticker: row_ticker,
    trade_date: row_trade_date,
    quantity: row_quantity,
    transaction_type: row_transaction_type,
    amount: row_amount
  };
}

function testTotalValue() {
  var range = SpreadsheetApp.getActiveSpreadsheet().getRangeByName('Transactions');
  var d = new Date(2013,3-1,29); // 2013-03-29
  var value = TOTAL_VALUE(range.getValues(), d);
  if (range != null) {
   Logger.log('Got range columns %s', range.getNumColumns());
  } else {
    Logger.log('Range is null %s', range);
  }
 
  Logger.log('Got total value ' + value);

}
