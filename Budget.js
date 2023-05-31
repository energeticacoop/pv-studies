function createBudget() {

  // Get chosen installation size and default budget
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const installationType = ss.getRangeByName("installationSize").getValue()
  const defaultBudget = ss.getSheetByName(installationType)

  // Show message
  Tools.showMessage(
    "ℹ️ Generación de presupuesto",
    `Generando presupuesto tipo de instalación ${installationType}...`, 3
  )

  // Get previous budget sheet
  const previousSheet = ss.getSheetByName("Presupuesto")
  if (previousSheet != null) previousSheet.setName("Viejo presupuesto")

  // Copy chosen budget sheet  
  const budgetSheet = defaultBudget.copyTo(ss)
  budgetSheet.setName("Presupuesto")
  budgetSheet.activate()

  // Move new budget sheet
  var index = SpreadsheetApp.getActive().getSheetByName("Flujos").getIndex()
  ss.moveActiveSheet(index + 1)

  // Set range names
  const data = budgetSheet.getDataRange().getValues()
  for (var row = 1; row < data.length; row++) {
    const concept = data[row][data[0].indexOf("Concepto")]
    if (concept == "PRESUPUESTO TOTAL SIN IVA") {
      const column = data[0].indexOf("Total Capítulo (€)")
      ss.setNamedRange("presupuestoSinIva", budgetSheet.getRange(row + 1, column + 1))
    }
    if (concept == "IVA (21%)") {
      const column = data[0].indexOf("Total Capítulo (€)")
      ss.setNamedRange("ivaPresupuesto", budgetSheet.getRange(row + 1, column + 1))
    }
    if (concept == "PRESUPUESTO TOTAL CON IVA") {
      const column = data[0].indexOf("Total Capítulo (€)")
      ss.setNamedRange("inversion", budgetSheet.getRange(row + 1, column + 1))
    }
    if (concept == "Tasa DROU") {
      const column = data[0].indexOf("Coste unitario")
      ss.setNamedRange("tasasDROU", budgetSheet.getRange(row + 1, column + 1))
    }

    const code = data[row][data[0].indexOf("Código")]
    if (code.toString() == "1") {
      const column = data[0].indexOf("Total Capítulo sin BI/CG (PEM) (€)")
      ss.setNamedRange("totalMontajeSuministro", budgetSheet.getRange(row + 1, column + 1))
    }
    if (code == "1.2") { // Modules
      const column = data[0].indexOf("Concepto")
      ss.setNamedRange("modulos1", budgetSheet.getRange(row + 2, column + 2))
      ss.setNamedRange("modulosNumero1", budgetSheet.getRange(row + 2, column + 3))
      ss.setNamedRange("modulos2", budgetSheet.getRange(row + 3, column + 2))
      ss.setNamedRange("modulosNumero2", budgetSheet.getRange(row + 3, column + 3))
      ss.setNamedRange("modulosNumeroTotal", budgetSheet.getRange(row + 1, column + 3))
    }
    if (code == "1.3") { // Inverters
      const column = data[0].indexOf("Concepto")
      ss.setNamedRange("inversor1", budgetSheet.getRange(row + 2, column + 2))
      ss.setNamedRange("inversorNumero1", budgetSheet.getRange(row + 2, column + 3))
      ss.setNamedRange("inversor2", budgetSheet.getRange(row + 3, column + 2))
      ss.setNamedRange("inversorNumero2", budgetSheet.getRange(row + 3, column + 3))
      ss.setNamedRange("inversorNumeroTotal", budgetSheet.getRange(row + 1, column + 3))
    }
  }
  ss.setNamedRange("potenciaPico", budgetSheet.getRange(1, data[0].indexOf("Potencia pico (Wp)") + 2))
  ss.setNamedRange("factorDimensionamiento", budgetSheet.getRange(1, data[0].indexOf("Factor dimensionamiento") + 2))
  ss.setNamedRange("potenciaTotalW", budgetSheet.getRange(1, data[0].indexOf("Potencia (W)") + 2))
  ss.setNamedRange("potenciaTotalKW", budgetSheet.getRange(1, data[0].indexOf("Potencia (kW)") + 2))


  // Remove previous budget sheet
  if (previousSheet != null) ss.deleteSheet(previousSheet)
}




function getBudget(getOptionalElements = false, getGuideTable = false) {

  function getByName(data, colName, row) {
    var col = data[0].indexOf(colName)
    if (col != -1) return data[row][col]
  }

  function getItemType(itemCode) {
    const matches = itemCode.match(new RegExp("\\.", 'g'))
    const count = matches ? matches.length : 0
    switch (count) {
      case 0: return "title"
      case 1: return "item"
      case 2: return "subitem"
    }
  }

  var data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Presupuesto").getDataRange().getDisplayValues()
  const budget = []
  const headers = data[0].slice(0, data[0].indexOf("Potencia pico (Wp)"))

  var lastEmptyElement = ""

  for (var i = 2; i < data.length; i++) {
    const row = Object.assign.apply({}, headers.map((v, j) => ({ [v]: data[i][j] })))
    const code = row.Código.toString()
    if (code.startsWith(lastEmptyElement + ".")) continue
    if (getGuideTable && !code.startsWith("1")) continue

    const quantity = row.Cantidad
    const itemType = getItemType(row.Código)
    row.type = itemType
    const zeroQuantity = (quantity == 0 || quantity == "" || quantity == "0")
    if ((itemType == "title" || itemType == "item") && zeroQuantity) lastEmptyElement = code

    // If non empty, add row to budget
    if (code != "" && !zeroQuantity && (getOptionalElements ? code.includes("O") : !code.includes("O"))) {

      // Insert empty rows after chapters (except SUMINISTRO and OPCIONALES)
      if (!getGuideTable && itemType == "title" && code != "1" && code != "O") budget.push({ "type": "empty" })
      if (getGuideTable && itemType == "item" && code != "1.1" && !zeroQuantity) budget.push({ "type": "empty" })
      budget.push(row)
    }
  }
  budget.push({ "type": "empty" })
  return budget
}



