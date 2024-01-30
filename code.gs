function onOpen() {
  let ui = SpreadsheetApp.getUi();
  ui.createMenu('Crecer')
      .addItem('Crear pago', 'createPayments')
      .addItem('Cerrar semana', 'closeWeek')
      //.addItem('Crear directorio de cliente', 'createFolder')
      //.addItem('Crear pagaré', 'createPagare')
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

  let idCredit = sheet.getRange('A' + index).getValue();
  let idClient = sheet.getRange('E' + index).getValue();
  let idCreditRenewed = sheet.getRange('D' + index).getValue();

  let response = ui.prompt('Estás seguro de crear la tabla de pago para ' + idCredit + '?', ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() == ui.Button.OK) {
    let sheets = SpreadsheetApp.getActiveSpreadsheet();
    let paymentsSheet = sheets.getSheetByName(sheetPaymentsName);

    // Get client data
    let clients = sheets.getSheetByName(sheetClientsName).getRange('A2:L').getValues();
    let clientIndex = clients.map(c => c[0]).indexOf(idClient);
    updateCreditRenewed(paymentsSheet, idCreditRenewed);

    // Get data
    let weeks = sheet.getRange('I' + index).getValue();
    let weeklyPayment = sheet.getRange('T' + index).getValue();
    let paymentDate = sheet.getRange('U' + index).getValue();
    let frequency = sheet.getRange('V' + index).getValue();
    let disposalDate = sheet.getRange('H' + index).getValue();
    let withHeldPayment = sheet.getRange('Y' + index).getValue() == "Si";

    // Workarounds
    weeks = frequency == "Semanal" ? weeks : weeks / 2;
    weeklyPayment = frequency == "Semanal" ? weeklyPayment : weeklyPayment * 2;
    let fixedDate = new Date(2023, 6, 27)

    if (paymentDate != "" && weeklyPayment != "") {
      let weekIndex = 1;
      let accrual = weeklyPayment;
      const rows = [];
      let firstRow = paymentsSheet.getLastRow() + 1;
      let lastRow = firstRow;

      if (withHeldPayment) { // Si tiene pago retenido 
        const sWeek = weekIndex + " de " + weeks;
        const weekNumber = "=WEEKNUM(D" + lastRow + ")";
        const debt = "=F".concat(lastRow, "-SUM(J", firstRow, ":J", lastRow, ")");
        rows.push([idCredit, sWeek, weekNumber, disposalDate, weeklyPayment, accrual, debt, "Si", true, weeklyPayment]);
        accrual += weeklyPayment;
        weekIndex++;
        lastRow++;
      }

      for (let week = weekIndex; week <= weeks; week++, lastRow++) {
        const sWeek = week + " de " + weeks;
        const weekNumber = "=WEEKNUM(D" + lastRow + ")";
        const debt = "=F".concat(lastRow, "-SUM(J", firstRow, ":J", lastRow, ")");
        rows.push([idCredit, sWeek, weekNumber, paymentDate, weeklyPayment, accrual, debt, (paymentDate <= fixedDate ? "Si" : "No"), '', weeklyPayment]);
        accrual += weeklyPayment;
        paymentDate = addDays(paymentDate, frequency == "Semanal" ? 7 : 14);
      }

      paymentsSheet.getRange(paymentsSheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    }
  }
}

function getWeekNumber(date, row) {
  return "=WEEKNUM(D" + row + ")";
}

function updateCreditRenewed(paymentsSheet, idCreditRenewed) {
  if (idCreditRenewed != "") {
    let rows = paymentsSheet.createTextFinder(idCreditRenewed).findAll();
    for(let row of rows) {
      const a1notation = "H" + row.getRowIndex();
      const cell = paymentsSheet.getRange(a1notation);
      if (cell.getValue() == "No") {
        cell.setValue("Renovado");
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
  let ui = SpreadsheetApp.getUi();
  const weekNumber = getISOWeekNumber(new Date());
  let response = ui.prompt('Estás seguro de cerrar la semana ' + weekNumber +'?', ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() == ui.Button.OK) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Pagos");

    let data = sheet.getDataRange();
    let values = data.getValues();

    values.filter(item => item[2] == weekNumber && item[7] == "No" &&(item[8] == "" || item[8] == false)).forEach(elem => {
      elem[9] = 0;
    });

    sheet.getRange("J:J").setValues(values.map(c => [ c[9] ]));
  }
}

function getISOWeekNumber(date) {
  const d = new Date(date);
  d.setHours(0, 0, 0, 0);
  d.setDate(d.getDate() + 4 - (d.getDay() || 7));
  const yearStart = new Date(d.getFullYear(), 0, 1);
  return Math.ceil(((d - yearStart) / 86400000 + 1) / 7);
}

function createPagare() {
  const templateId = '195pg8r4BJR0pOWpduH4aUZO_xyogiyPWoEwVGfw94xM';
  const templateFile = DriveApp.getFileById(templateId);

}
