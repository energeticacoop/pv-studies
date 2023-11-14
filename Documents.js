
function createFinalStudy() {
  createDocuments(getFinalStudyTemplates(), "outputStudy")
}

function create00FolderDocumentation() {
  createDocuments(get00FolderTemplates(), "output00Folder")
}

function create01FolderDocumentation() {
  createDocuments(get01FolderTemplates(), "output01Folder")
}

function createMemoryAndGuide() {
  createDocuments(getGuideTemplates(), "outputGuide")
  createDocuments(getMemoryTemplates(), "outputMemory")
}

function createBill(){
    createDocuments(getBillTemplates(), "outputBill")
}

function create02And03FolderDocumentation() {
  createDocuments(get02FolderTemplates(), "output02Folder")
  createDocuments(get03FolderTemplates(), "output03Folder")
}

function create04FolderDocumentation() {
  createDocuments(get04FolderTemplates(), "output04Folder")
}

function createAllDocuments() {
  createFinalStudy()
  create00FolderDocumentation()
  create01FolderDocumentation()
  createMemoryAndGuide()
  create02And03FolderDocumentation()
  create04FolderDocumentation()
}

function createTestDocument() {
  createDocuments(getTestTemplates(), "test")
}

function createDocuments(templates, outputRangeName) {

  // Clear output range
  const outputRange = getRangeByName(outputRangeName)
  clearRange(outputRangeName)

  // Replace values in all templates
  templates.forEach((template, templateIndex) => {

    // Create document and replace values
    const copy = createDocumentFromTemplate(template)

    // Output link to document
    setURL(
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Documentación").getRange(outputRange.getRow() + templateIndex, outputRange.getColumn()),
      copy.getUrl(),
      template.templateName
    )
    SpreadsheetApp.flush()
  })
}





function createDocumentFromTemplate(template) {

  // Get template values
  const destinationFolder = getDestinationFolder(template.folder)
  const filename = `${getValue("Nombre")} ${getValue("Apellidos")} - ${template.templateName}`
  const templateId = template.templateId
  const exportToPDF = template.exportToPDF
  const copyComments = template.copyComments

  // Remove all matching files on destination folder to avoid duplicates
  Tools.deleteFile(filename, destinationFolder)
  const templateFile = DriveApp.getFileById(templateId)
  const mimeType = templateFile.getMimeType()
  const copy = templateFile.makeCopy(filename, destinationFolder)

  // Form search pattern
  const leftDelimiter = "<"
  const rightDelimiter = ">"
  const searchPattern = `${leftDelimiter}.*?${rightDelimiter}`

  // Copy comments and replies
  function copyCommentsAndReplies(copy, templateId) {
    var newDocId = copy.getId()
    var commentList = Drive.Comments.list(templateId, { 'maxResults': 100 })
    commentList.items.forEach(item => {
      //if (!item.status == "resolved") {
      var replies = item.replies
      delete item.replies
      var commentId = Drive.Comments.insert(item, newDocId).commentId
      replies.forEach(reply => Drive.Replies.insert(reply, newDocId, commentId))
      //}
    })
  }

  // Create doc or excel from template 

  // Documents
  if (mimeType == "application/vnd.google-apps.document") {

    const doc = DocumentApp.openById(copy.getId())

    // Replace signature images
    const signatureText = "<firmaIngeniera>"
    if (doc.getBody().findText(signatureText) != null) {
      const signature = DriveApp.getFileById(getValue("firmaIngeniera")).getBlob()
      replaceImage(doc, signatureText, signature, 300)
    }

    // Replace values in body, headers and footers
    const parent = doc.getBody().getParent()

    for (var i = 0; i < parent.getNumChildren(); i++) {
      try {
        // Get all values to be replaced in current child
        const child = parent.getChild(i)

        var range = child.findText(searchPattern)
        const valuesToBeReplaced = []
        while (range) {
          const matches = [...range.getElement().asText().getText().matchAll(searchPattern)].map(e => e[0])
          matches.forEach(match => {
            if (!valuesToBeReplaced.includes(match)) valuesToBeReplaced.push(match)
          })
          range = child.findText(searchPattern, range)
        }

        // Replace values
        valuesToBeReplaced.forEach(valueToBeReplaced => {
          const namedRange = SpreadsheetApp.getActiveSpreadsheet().getRangeByName(valueToBeReplaced.split(leftDelimiter).pop().split(rightDelimiter)[0])
          if (namedRange != null) {
            const namedRangeValue = namedRange.getDisplayValue()
            child.replaceText(valueToBeReplaced, namedRangeValue)
          }
        })

      }
      catch (err) { }
    }

    // Copy comments and replies
    if (copyComments) copyCommentsAndReplies(copy, templateId)

    // Manage additional actions for templates with special tables
    if (templateId == "1Q7aRFNjqB_Kc9daRPDRhs2kqt7XuZMYmhGRQjLVHBY8") createPemAdditionalContents(doc)
    if (templateId == "1xttGTW5xY0mnyuCEU8gEQwQ926WH4vTG0mZS4CEKuJQ") createBillAdditionalContents(doc)
    if (templateId == "13ldY9Q8bK7ijZSauJYD2piVbyYfYKtTPjf7qdhWTE6k") createFinalStudyAdditionalContents(doc)
    if (templateId == "1k9lYTmxMsINC6U49w1NWOu0-3dGkYUBF5yrhKmkGb04") createInstallationGuideAdditionalContents(doc)

    doc.saveAndClose()

    // Create PDF version
    if (exportToPDF) {
      // Remove all matching pdf files on destination folder to avoid duplicates
      const pdfFilename = filename + ".pdf"
      Tools.deleteFile(pdfFilename, destinationFolder)

      var pdfVersion = DriveApp.createFile(doc.getAs('application/pdf'))
      pdfVersion.moveTo(destinationFolder)
      pdfVersion.setName(pdfFilename)
    }
  }

  // Spreadsheets
  else if (mimeType == "application/vnd.google-apps.spreadsheet") {
    const excel = SpreadsheetApp.openById(copy.getId())

    // Get patterns to be replaced
    const textFinder = excel.createTextFinder(searchPattern).useRegularExpression(true)
    const allMatches = textFinder.findAll().map(e => e.getValue())
    const replacementValues = []
    allMatches.forEach(row => {
      replacementValues.push(...[...row.toString().matchAll(searchPattern)].map(e => e[0]))
    })

    // Replace variables in template
    replacementValues.forEach(value => {
      const textFinder = excel.createTextFinder(value)
      const namedRangeName = value.split(leftDelimiter).pop().split(rightDelimiter)[0]
      const namedRange = SpreadsheetApp.getActiveSpreadsheet().getRangeByName(namedRangeName)
      if (namedRange != null) {
        textFinder.replaceAllWith(getValue(namedRangeName))
      }
    })

    // Copy comments and replies
    copyCommentsAndReplies(copy, templateId)
  }

  return copy

}