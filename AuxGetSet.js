function setURL(range, url, text) {
  var richValue = SpreadsheetApp.newRichTextValue()
    .setText(text)
    .setLinkUrl(url)
    .build()
  range
    .setRichTextValue(richValue)
}

function setColumnURLS(rangeName, urls, texts) {
  const range = getRangeByName(rangeName)
  range.getValues().forEach(function (row, i) {
    var richValue = SpreadsheetApp.newRichTextValue()
      .setText(texts[i])
      .setLinkUrl(urls[i])
      .build()
    range.getCell(i + 1, 1).setRichTextValue(richValue)
  })
}

function setValue(name, value) {
  getRangeByName(name).setValue(value)
}

function setValues(name, values) {
  getRangeByName(name).setValues(values)
}

function setColumn(name, values) {
  setValues(
    name,
    values.map((element) => [element])
  )
}

function getValue(rangeName) {
  return getRangeByName(rangeName).getValue()
}

function getValues(rangeName) {
  return getRangeByName(rangeName).getValues()
}

function getColumn(rangeName) {
  return getValues(rangeName).map((element) => element[0])
}

function clearRange(rangeName){
  SpreadsheetApp.getActiveSpreadsheet().getRangeByName(rangeName).clearContent()
}

function getRangeByName(rangeName){
  return SpreadsheetApp.getActiveSpreadsheet().getRangeByName(rangeName)
}


function getTariffPrices(tariff) {
  if (tariff == "2.0TD") return getColumn("energeticaTariff20")
  if (tariff == "3.0TD") return getColumn("energeticaTariff30")
  if (tariff == "6.1TD") return getColumn("energeticaTariff61")
}
