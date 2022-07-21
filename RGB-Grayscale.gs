var ui = SpreadsheetApp.getUi();
var ss = SpreadsheetApp.getActiveSpreadsheet();
var rgbSheet = ss.getSheetByName('RGB');
var convSheet = ss.getSheetByName('Converted');


function onOpen() {
  ui.createMenu('Convert to Grayscale')
  .addItem('Lightness', 'lgtRGB')
  .addItem('Average', 'avgRGB')
  .addItem('Luminosity', 'lumRGB')
  .addToUi();
};

function convertRGB(type){
  convSheet.getDataRange().clearContent();

  var rangeData = rgbSheet.getDataRange();
  var rangeNotate = rangeData.getA1Notation();
  var pasteRange = convSheet.getRange(rangeNotate);
  var rangeValues = rangeData.getValues();

  rangeValues.forEach(rowFunc);

  function rowFunc(value) {
    value.forEach(cellFunc);
  };

  function cellFunc(clr, index, array) {
    var countComma = (clr.match(/,/g) || []).length;

    if(clr.length <= 11 && clr.length >= 5 && countComma == 2){
      let firstComma = clr.indexOf(",");
      let secondComma = clr.indexOf(",",firstComma + 1);
      let r = clr.slice(0,firstComma);
      let g = clr.slice(firstComma + 1, secondComma);
      let b = clr.slice(secondComma + 1);
      
      switch (type){
        case 'lgt':
          array[index] = Math.round((Math.min(r,g,b) + Math.max(r,g,b)) / 2);
          break;
        case 'avg':
          array[index] = Math.round((1 * r + 1 * g + 1 * b) / 3);
          break;
        case 'lum':
          array[index] = Math.round(0.3 * r + 0.59 * g + 0.11 * b);
          break;
      };
    }else {
      array[index] = '';
    };
  };

  pasteRange.setValues(rangeValues);

};

function lgtRGB() {
  convertRGB('lgt');
};

function avgRGB() {
  convertRGB('avg');
};

function lumRGB() {
  convertRGB('lum');
};