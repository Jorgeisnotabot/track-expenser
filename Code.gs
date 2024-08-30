function addExpense(date, description, category, amount, paymentMethod) {
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    sheet.appendRow([date, description, category, amount, paymentMethod]);
}

function showSidebar() {
    let html = HtmlService.createHtmlOutputFromFile('Sidebar')
        .setTitle('Expense Tracker');
    SpreadsheetApp.getUi().showSidebar(html);
}

function onOpen() {
    let ui = SpreadsheetApp.getUi();
    ui.createMenu('Expense Tracker')
        .addItem('Add Expense', 'showSidebar')
        .addToUi();
}