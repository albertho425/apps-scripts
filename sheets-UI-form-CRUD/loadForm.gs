/*************************************************************
*  Load the form in the middle of screen and set the dimensions
**************************************************************/

function loadMainForm() {

  const htmlServ = HtmlService.createTemplateFromFile("main");
  const html = htmlServ.evaluate();
  html.setWidth(1280).setHeight(720);
  const ui = SpreadsheetApp.getUi();
  ui.showModalDialog(html, "MEL UI Form");

}

/*************************************************************
*  Set the menu in the the spreadsheet
**************************************************************/


function createMenu_(){

  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu("Custom Menu");
  menu.addItem("Open Form", "loadMainForm");
  menu.addToUi();

}

/*************************************************************
*  When the sreadsheet opens, run the function
**************************************************************/


function onOpen(){

  createMenu_();

}

/*************************************************************
*  Load the form on the side and set the dimensions
**************************************************************/


function loadMainFormOnSide() {

  const htmlServ = HtmlService.createTemplateFromFile("main");
  const html = htmlServ.evaluate();
  html.setWidth(1280).setHeight(720);
  const ui = SpreadsheetApp.getUi();
  ui.showModalDialog(html, "Edit MEL Data");

}
