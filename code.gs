function onOpen() {
  let ui = SpreadsheetApp.getUi();
  ui.createMenu('Crecer')
      .addItem('Crear pago', 'createPayments')
      .addItem('Cerrar semana', 'closeWeek')
      .addItem('Crear directorio de cliente', 'createFolder')
      .addItem('Crear pagaré', 'createPagare')
      .addToUi();
}

const sheetPaymentsName = 'Pagos';
const sheetClientsName = 'Clientes';
const directoryRoot = 'Crecer';

const addDays = (date, days) => {
  var result = new Date(date);
  result.setDate(result.getDate() + days);
  return result;
}

const createPayments = () => {
  let ui = SpreadsheetApp.getUi();

  // Get person
  let sheet = SpreadsheetApp.getActiveSheet();
  let currentCell = sheet.getActiveCell();
  let index = currentCell.getRowIndex();

  //for (let index = 2; index <= 113; index++) {

    let idCredit = sheet.getRange('A' + index).getValue();
    let idClient = sheet.getRange('E' + index).getValue();

    let response = ui.prompt('Estás seguro de crear la tabla de pago para ' + idCredit + '?', ui.ButtonSet.OK_CANCEL);

    if (response.getSelectedButton() == ui.Button.OK) {
      let sheets = SpreadsheetApp.getActiveSpreadsheet();
      let paymentsSheet = sheets.getSheetByName(sheetPaymentsName);

      // Get client data
      let clients = sheets.getSheetByName(sheetClientsName).getRange('A2:H').getValues();
      let clientIndex = clients.map(c => c[0]).indexOf(idClient);
      let clientName = clients[clientIndex][3] + ' ' + clients[clientIndex][4] + ' ' + clients[clientIndex][5];
      let phone = clients[clientIndex][6];
      let address = clients[clientIndex][7];
      let comission = clients[clientIndex][11] == "Si" ? 0.035 : 0.07;

      // Get data
      let weeks = sheet.getRange('I' + index).getValue();
      let weeklyPayment = sheet.getRange('T' + index).getValue();
      let paymentDate = sheet.getRange('U' + index).getValue();
      let frequency = sheet.getRange('V' + index).getValue();

      // Workarounds
      weeks = frequency == "Semanal" ? weeks : weeks / 2;
      weeklyPayment = frequency == "Semanal" ? weeklyPayment : weeklyPayment * 2;
      let fixedDate = new Date(2023, 6, 27)

      if (paymentDate != "" && weeklyPayment != "") {
        for (let week = 1; week <= weeks; week++) {
          const sWeek = week + " de " + weeks;

          startDate = new Date(paymentDate.getFullYear(), 0, 1);
          const days = Math.floor((paymentDate - startDate) / (24 * 60 * 60 * 1000));
          const weekNumber = Math.ceil((paymentDate.getDay() + 1 + days) / 7);

          const row = [idCredit, clientName, phone, address, sWeek, weekNumber, paymentDate, weeklyPayment, (paymentDate <= fixedDate ? "Si" : "No"), comission, comission * weeklyPayment]
          paymentDate = addDays(paymentDate, frequency == "Semanal" ? 7 : 14);
          paymentsSheet.appendRow(row);
        }
      }
    }
}

function createFolder() {
  // Get index
  let sheet = SpreadsheetApp.getActiveSheet();
  let currentCell = sheet.getActiveCell();
  let index = currentCell.getRowIndex();

  let root = DriveApp.getFoldersByName(directoryRoot);
  if (root.hasNext()) {
    let clientes = root.next().getFoldersByName("Clientes");
    if (clientes.hasNext()) {
      let folder = clientes.next();
      let idClient = sheet.getRange('A' + index).getValue();
      let name = sheet.getRange('D' + index).getValue();
      let lastName = sheet.getRange('E' + index).getValue();
      let secondLastName = sheet.getRange('F' + index).getValue();
      let folderName = idClient + " " + name + " " + lastName + " " + secondLastName;
      SpreadsheetApp.getUi().alert(folder.createFolder(folderName).getUrl());
    }
  }
}

function closeWeek() {
  let sheets = SpreadsheetApp.getActiveSpreadsheet();
  let paymentsSheet = sheets.getSheetByName(sheetPaymentsName);
  let allPayments = paymentsSheet.getRange("A:I").getValues();
  let payments = allPayments.filter(a => a[7] == true);

  let sheetWeek = sheets.insertSheet("New Sheet");
  sheetWeek.getRange(sheetWeek.getLastRow() + 1, 1, payments.length, payments[0].length).setValues(payments);

/*
  
  for (let pay of payments) {
    sheetWeek.appendRow(pay);
  }

  */
}

function createPagare() {
  const templateId = '195pg8r4BJR0pOWpduH4aUZO_xyogiyPWoEwVGfw94xM';
  const templateFile = DriveApp.getFileById(templateId);

}
