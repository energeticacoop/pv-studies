function Borraroutputproduccion() {
  const outputRanges = [
    "tmyFileepw",
    "tmyFilecsv",
    "satelliteImages",
    "scales",
    "UTMcoordinates",
    "CP"
  ]
  SpreadsheetApp.getActive()
    .getSheetByName("Producción")
    .getNamedRanges()
    .forEach((range) => {
      if (outputRanges.includes(range.getName()))
        range.getRange().clearContent()
    })
}

function Borraroutputconsumoconvencional() {
  const outputRanges = [
    "yearlyConsumption",
    "csvCUPS",
    "monthlyConsumption",
    "monthlyHourlyPeak",
    "hourlyConsumptionMeans",
    "normalizedDates",
    "normalizedCCH",
    "normalizedComments",
    "weeklyConsumptionMeans",
    "exceedingValues",
  ]
  SpreadsheetApp.getActive()
    .getSheetByName("Consumo base")
    .getNamedRanges()
    .forEach((range) => {
      if (outputRanges.includes(range.getName()))
        range.getRange().clearContent()
    })
}
