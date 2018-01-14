/**
* 集計開始  
*/
function startAggregate() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var aggregateSheet = ss.getSheetByName("aggregate");
  var summarySheet = ss.getSheetByName("summary");
  
  const LAST_ROW = aggregateSheet.getLastRow();
  const LAST_COLUMN = aggregateSheet.getLastColumn();

  const　FIXED_COLUMN_COUNT　= 5;
  const PURCHASER_START_INDEX = FIXED_COLUMN_COUNT + 1; 
  const PURCHASER_COUNT　= LAST_COLUMN - FIXED_COLUMN_COUNT;
  
  var purchaserMatrix = aggregateSheet.getRange(1, PURCHASER_START_INDEX,LAST_ROW, PURCHASER_COUNT).getValues();
  var fixedFormatMatrix = aggregateSheet.getRange(1, 1, LAST_ROW, FIXED_COLUMN_COUNT).getValues();

  var purchaserTotalPriceArray = [];
  purchaserMatrix[0].forEach(function(purchaserName,purchaserIndex,array) {

    // 購入者に対し、購入した商品を取得するループ
    var sumPurchasePrice = 0;
    for (var row = 2; row <= LAST_ROW; row++) {
      var rowIndex = row -1;
      
      var purchaseCount = purchaserMatrix[rowIndex][purchaserIndex];
      if (purchaseCount > 0) {
        var itemPrice　= fixedFormatMatrix[rowIndex][3];
        var sumPurchaserCount = fixedFormatMatrix[rowIndex][4];
        var purchasePrice = itemPrice * (purchaseCount / sumPurchaserCount);
        sumPurchasePrice = sumPurchasePrice + purchasePrice;        
      }
    }
    purchaserTotalPriceArray[purchaserName] = sumPurchasePrice;
  });
  postscriptSummarySheet(purchaserTotalPriceArray);
  
}

/**
* サマリーシートに書き出し  
*/
function postscriptSummarySheet(purchaserTotalPriceArray) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var summarySheet = ss.getSheetByName("summary");

  var i = 2;
  Object.keys(purchaserTotalPriceArray).forEach(function(key) {
    summarySheet.getRange("A" + i).setValue(key);
    summarySheet.getRange("B" + i).setValue(purchaserTotalPriceArray[key]);
    i = i+ 1;
  });
  
}
