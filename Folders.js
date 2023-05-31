function getDestinationFolder(destinationFolderName) {

  const folder01 = DriveApp.getFolderById(getParentFolderId())
  const clientFolder = folder01.getParents().next()
  const folder02 = clientFolder.getFoldersByName("02 - Tramitación").next()
  const folder03 = clientFolder.getFoldersByName("03 - Ejecución").next()

  switch (destinationFolderName) {
    case "clientFolder":
      return clientFolder
    case "folder01":
      return folder01
    case "folder02":
      return folder02
    case "folder0200":
      return folder02.getFoldersByName("00 - Documentación para firma").next()
    case "folder0201":
      return folder02.getFoldersByName("01 - DROU").next()
    case "folder0202":
      return folder02.getFoldersByName("02 - BOEL").next()
    case "folder0203":
      return folder02.getFoldersByName("03 - Registro autoconsumo").next()
    case "folder0204":
      return folder02.getFoldersByName("04 - Documentación proyecto").next()
    case "folder03":
      return folder03
  }

}


function forgetFolders() {

  const erasableRangesDocumentation = [
    "clientFolder",
    "outputStudy",
    "adminFolder",
    "outputGuide",
    "output00Folder",
    "output01Folder",
    "output02Folder",
    "output03Folder",
    "output04Folder",
    "outputMemory",
    "outputBill",
    "folder01",
    "folder02",
    "folder0200",
    "folder0201",
    "folder0202",
    "folder0203",
    "folder0204",
    "folder03",
    "helpscoutEmail",
  ]

  SpreadsheetApp.getActiveSpreadsheet().getNamedRanges().forEach((range) => {
    if (erasableRangesDocumentation.includes(range.getName()))
      range.getRange().clearContent()
  })
}


function printFolderLink(rangename) {
  const folder = getDestinationFolder(rangename)
  setURL(getRangeByName(rangename), folder.getUrl(), folder.getName())
}

function printAllFoldersLinks(){
  // Print folder links
  const clientFolder = getDestinationFolder("clientFolder")
  setURL(getRangeByName("clientFolder"), clientFolder.getUrl(), `Directorio madre: ${clientFolder.getName()}`)
  printFolderLink("folder01")
  printFolderLink("folder02")
  printFolderLink("folder0200")
  printFolderLink("folder0201")
  printFolderLink("folder0202")
  printFolderLink("folder0203")
  printFolderLink("folder0204")
  printFolderLink("folder03")

}
