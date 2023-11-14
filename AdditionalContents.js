
function replaceBudgetTable(budget, budgetTable, subtotalColumName = "Total Capítulo (€)") {
  // Get master TableRows
  const titleRow = budgetTable.getRow(0).copy()
  const itemRow = budgetTable.getRow(2).copy()
  const lastRow = budgetTable.getRow(4).copy()
  budgetTable.clear()

  // Add table rows to template
  budget.forEach((element, index) => {
    var row
    switch (element.type) {
      case "title":
        row = budgetTable.insertTableRow(index, titleRow.copy())
        row.getCell(0).editAsText().replaceText("title", element.Concepto)
        break
      case "item":
        row = budgetTable.insertTableRow(index, itemRow.copy())
        row.getCell(0).editAsText().replaceText("Concept", element.Concepto).replaceText("Description", element.Descripción)
        row.getCell(1).editAsText().replaceText("quantity", element.Cantidad)
        row.getCell(2).editAsText().replaceText("subtotal", element[subtotalColumName])
        break
      case "empty":
        budgetTable.insertTableRow(index, lastRow.copy())
        break
    }
  })
}




function createPemAdditionalContents(doc) {

  // Replace budget table
  const tables = doc.getBody().getTables()
  const budgetTable = tables[0] // Manually select budget table!

  // Get budgets and replace tables
  const budget = getBudget(getOptionalElements = false).filter(e => e.type != "subitem")
  // Take only chapter one and add empty line
  const chapter1Budget = budget.slice(0, budget.findIndex(e => e.Código == "2"))
  replaceBudgetTable(chapter1Budget, budgetTable, subtotalColumName="Total Capítulo sin BI/CG (PEM) (€)")
}


function createBillAdditionalContents(doc) {

  // Replace budget table
  const tables = doc.getBody().getTables()
  const budgetTable = tables[tables.length - 2] // Manually select budget table!

  // Get budgets and replace tables
  const budget = getBudget(getOptionalElements = false).filter(e => e.type != "subitem")
  replaceBudgetTable(budget, budgetTable)
}


function createInstallationGuideAdditionalContents(doc) {

  // Replace budget table
  const tables = doc.getBody().getTables()
  const budgetTable = tables[0] // Manually select budget table!

  // Get budgets and replace tables
  const budget = getBudget(getOptionalElements = false, getGuideTable = true)
  //  const filteredBudget = budget.filter(e => e.type != "title" && (e.type == "empty" || e.Código.startsWith("1"))).filter(element => element.type == "item" || element.type == "empty" || element["Subcat. Coste"] == "MAT")

  const filteredBudget = budget.filter(e => e.type == "item"|| (e.type == "subitem" && e["Subcat. Coste"] == "MAT") || e.type == "empty" )
  replaceGuideTable(filteredBudget, budgetTable)
}

function replaceGuideTable(budget, budgetTable) {
  // Get master TableRows
  const titleRow = budgetTable.getRow(0).copy()
  const itemRow = budgetTable.getRow(2).copy()
  const lastRow = budgetTable.getRow(3).copy()
  budgetTable.clear()

  // Add table rows to template
  budget.forEach((element, index) => {
    var row
    switch (element.type) {
      case "item":
        row = budgetTable.insertTableRow(index, titleRow.copy())
        row.getCell(0).editAsText()
          .replaceText("code", element.Código)
          .replaceText("item", element.Concepto)
        break
      case "subitem":
        row = budgetTable.insertTableRow(index, itemRow.copy())
        row.getCell(0).editAsText()
          .replaceText("ProviderReference", element["Referencia Proveedor"])
        row.getCell(1).editAsText()
          .replaceText("Description", element.Descripción)
        row.getCell(2).editAsText().replaceText("quantity", element.Cantidad)
        break
      case "empty":
        budgetTable.insertTableRow(index, lastRow.copy())
        break
    }
  })
}


function createFinalStudyAdditionalContents(doc) {

  // Replace budget tables
  const body = doc.getBody()
  const tables = body.getTables()
  const budgetTable = tables[tables.length - 3] // Manually select budget table!
  const optionalElementsTable = tables[tables.length - 1] // Manually select optional elements table!

  // Get budgets and replace tables
  const budget = getBudget(getOptionalElements = false).filter(e => e.type != "subitem")
  replaceBudgetTable(budget, budgetTable)

  const optionalBudget = getBudget(getOptionalElements = true).filter(e => e.type != "subitem")
  replaceBudgetTable(optionalBudget, optionalElementsTable)



  // Remove null production values lines in summary table
  const conventionalTotal = getValue("conventionalTotal")
  const recurringTotal = getValue("recurringTotal")
  const ashpTotal = getValue("ashpTotal")
  const saveTotal = getValue("saveTotal")
  // Get master TableRows
  const summaryTable = tables[2] // Manually select summary table!
  const conventionalTotalRow = summaryTable.getRow(1)
  const recurringTotalRow = summaryTable.getRow(2)
  const ashpTotalRow = summaryTable.getRow(3)
  const saveTotalRow = summaryTable.getRow(4)
  if (conventionalTotal == 0) conventionalTotalRow.removeFromParent()
  if (recurringTotal == 0) recurringTotalRow.removeFromParent()
  if (ashpTotal == 0) ashpTotalRow.removeFromParent()
  if (saveTotal == 0) saveTotalRow.removeFromParent()




  // Get charts
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const charts = ss.getSheetByName("Documentación").getCharts()
  const chartReplacementValues = ["<graficaConsumos>", "<graficaAutoconsumo>", "<graficaFacturaMensual>", "<graficaFacturaAnualImpuestos>",]
  const IMAGEWIDTH = 600

  chartReplacementValues.forEach((chartReplacementValue, index) => {
    // Get image through Slide (workaround to preserve chart properties)
    const slides = SlidesApp.create("temp")
    const imageBlob = slides
      .getSlides()[0]
      .insertSheetsChartAsImage(charts[index])
      .getAs("image/png")
    DriveApp.getFileById(slides.getId()).setTrashed(true)
    replaceImage(doc, chartReplacementValue, imageBlob, IMAGEWIDTH)
  })

}
