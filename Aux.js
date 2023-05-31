const sumVector = (sum, number) => sum + number

function rotateArray(array, displacement) {
  return [
    ...array.slice(array.length - displacement),
    ...array.slice(0, array.length - displacement),
  ]
}

function getParentFolderId() {
  return DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId())
    .getParents()
    .next()
    .getId()
}

function getParentFolder() {
  return DriveApp.getFolderById(getParentFolderId())
}

function downloadFile(fileURL, fileName, destinationFolder) {
  var response = UrlFetchApp.fetch(fileURL, { muteHttpExceptions: true })
  var rc = response.getResponseCode()
  if (rc == 200) {
    var fileBlob = response.getBlob()
    if (destinationFolder != null) {
      var file = destinationFolder.createFile(fileBlob)
      file.setName(fileName)
    }
  }
  var fileInfo = { rc: rc, file: file }
  return fileInfo
}

function eraseNamedFields() {
  var ui = SpreadsheetApp.getUi()

  var result = ui.alert(
    `Estás a punto de borrar todos los campos de entrada de la pestaña "${SpreadsheetApp.getActiveSheet().getName()}"`,
    "¿Seguro que quieres continuar?",
    ui.ButtonSet.YES_NO
  )

  const unerasableRanges = [
    "SIPS2",
    "SIPS3",
    "CSVtype",
    "CSVfilename",
    "REEconsumption",
    "surpassingValuesFactor",
    "ahorroTotal", "ahorroTotalConImpuestos", "installationSize"
  ]

  if (result == ui.Button.YES) {
    switch (SpreadsheetApp.getActive().getActiveSheet().getSheetName()) {

      case "Repositorio":
        SpreadsheetApp.getUi().alert("Hombre, el repositorio no, por favor 🤦‍♂️")
        break

      case "Presupuesto":
        break

      default:
        SpreadsheetApp.getActive()
          .getActiveSheet()
          .getNamedRanges()
          .forEach((range) => {
            if (!unerasableRanges.includes(range.getName()))
              range.getRange().clearContent()
          })

    }

    SpreadsheetApp.flush()
  }
}

function replaceImage(doc, replacementValue, imageBlob, imageWidth) {

  const searchResult = doc.getBody().findText(replacementValue)
  if (searchResult) {
    // Set image
    var imageContainer = searchResult.getElement().getParent().asParagraph()
    imageContainer.clear()
    const insertedImage = imageContainer.appendInlineImage(imageBlob)
    const width = insertedImage.getWidth()
    const height = insertedImage.getHeight()
    insertedImage.setWidth(imageWidth).setHeight(height * imageWidth / width)
    return insertedImage
  }
}


function importDb() {

  function cloneGoogleSheet(sheetName) {
    // source doc
    var sss = SpreadsheetApp.openById("1s11D27POxjYv0juY6gIvwbyb4pYs0EoPzqYmyxmZfBY")

    // source sheet
    var ss = sss.getSheetByName(sheetName)

    // Get full range of data
    var SRange = ss.getDataRange()

    // get the data values in range
    var SData = SRange.getValues()

    // target spreadsheet
    var tss = SpreadsheetApp.getActiveSpreadsheet()

    // target sheet
    var ts = tss.getSheetByName(sheetName)

    // Clear the Google Sheet before copy
    ts.clear({ contentsOnly: true })

    // set the target range to the values of the source data
    ts.getRange(SRange.getA1Notation()).setValues(SData)
  }

  const sheetNames = [
    "Fusibles CA",
    "Puesta a tierra",
    "Cableado",
    "Fusibles CC",
    "Cargador VE",
    "Magnetos",
    "CombiIGA+Sobretensiones",
    "Dif.",
    "Cajas y Cuadros",
    "Canales",
    "Descargador CA",
    "Eq. Genéricos",
    "Repositorio",
    "Listas",
    "Tablas",
    "Normativa",
    "Inversores",
    "Módulos",
    "Estructura",
    "Consumibles",
    "Descargador CC",
    "Meters",
    "Enphase Extra",
    "Control y otros",
    "InterruptorSeccionador",
    "Contadores distri",
    "Cargador VE",
  ]

  sheetNames.forEach(e=> cloneGoogleSheet(e))
  setValue("fechaImportacion", new Date())

}
