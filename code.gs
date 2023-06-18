function onOpen() {
  let ui = SpreadsheetApp.getUi();
  ui.createMenu('Crecer')
      .addItem('Crear pago', 'createPayments')
      .addToUi();
}

const sheetPaymentsName = 'Pagos';
const sheetClientsName = 'Clientes';

const addDays = (date, days) => {
  var result = new Date(date);
  result.setDate(result.getDate() + days);
  return result;
}

const getClient = (clients) => {
  clients.values
}

const createPayments = () => {
  let ui = SpreadsheetApp.getUi();

  // Get person
  let sheet = SpreadsheetApp.getActiveSheet();
  let currentCell = sheet.getActiveCell();
  let index = currentCell.getRowIndex();

  // for (let index = 2; index <= 63; index++) {

    let idCredit = sheet.getRange('A' + index).getValue();
    let idClient = sheet.getRange('C' + index).getValue();

    // let response = ui.prompt('Estás seguro de crear la tabla de pago para ' + person + '?', ui.ButtonSet.OK_CANCEL);

    //if (response.getSelectedButton() == ui.Button.OK) {
      let sheets = SpreadsheetApp.getActiveSpreadsheet();
      let paymentsSheet = sheets.getSheetByName(sheetPaymentsName);

      // Get client data
      let clients = sheets.getSheetByName(sheetClientsName).getRange('B2:H').getValues();
      let clientIndex = clients.map(c => c[0]).indexOf(idClient);
      let clientName = clients[clientIndex][2] + ' ' + clients[clientIndex][3] + ' ' + clients[clientIndex][4];
      let phone = clients[clientIndex][5];
      let address = clients[clientIndex][6];

      // Get data
      let weeks = sheet.getRange('K' + index).getValue();
      let weeklyPayment = sheet.getRange('L' + index).getValue();
      let paymentDate = sheet.getRange('N' + index).getValue();

      for (let week = 1; week <= weeks; week++) {
        const sWeek = week + " de " + weeks;
        const summary = sWeek + " / $" + weeklyPayment.toFixed(2).replace(/\d(?=(\d{3})+\.)/g, '$&,') + " / No";
        const row = [idCredit, clientName, phone, address, sWeek, paymentDate, weeklyPayment, summary, false]
        paymentDate = addDays(paymentDate, 7);
        paymentsSheet.appendRow(row);
      }
    //}
  // }
}