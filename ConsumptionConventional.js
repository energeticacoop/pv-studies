function processConventionalConsumptionAndProduction() {
  Borraroutputconsumoconvencional()

  // Get normalized curve
  const normalizedCCH = (() => {
    const csvType = getValue("CSVtype")
    switch (csvType) {
      case "Infoenergia":
        return processInfoenergia()
      case "Datadis API":
        return processDatadisAPI()
      case "Datadis CSV":
        return processDatadisCsv()
      case "Iberdrola":
        return processCNMC()
      case "Unión Fenosa":
        return processCNMC()
      case "REE":
        return processREE()
      case "SIPS 2.0":
        return processSIPS2()
      case "SIPS 3.0":
        return processSIPS3()
    }
  })()

  // Output dates and cch data
  Tools.showMessage(
    `ℹ️ Consumos convencionales y producción`,
    `Procesando datos de consumo...`
  )
  const dates = normalizedCCH.map((e) => e.date)
  const cch = normalizedCCH.map((e) => e.value)
  const comments = normalizedCCH.map((e) => e.comment)

  // Match production and consumption values
  const production = getColumn("productionSAM")
  const numberOfDates = production.length
  const normalizedProduction = new Array(numberOfDates)
  const dateDisplacement =
    numberOfDates -
    dates.findIndex(
      (e) => e.getMonth() == 0 && e.getDate() == 1 && e.getHours() == 0
    )
  var negativeProduction = 0
  for (var i = 0; i < numberOfDates; i++) {
    var normalizedProductionIndex = (i + dateDisplacement) % numberOfDates
    // Asign production value and accumulate negative values (inverter nightly consumption)
    normalizedProduction[i] = Math.max(0, production[normalizedProductionIndex])
    if (production[normalizedProductionIndex] < 0) {
      negativeProduction += -production[normalizedProductionIndex]
    }
  }

  // Output production data
  const totalProduction = normalizedProduction.reduce(sumVector)
  const inverterConsumptionOverTotalProduction = (
    Math.round((negativeProduction / totalProduction + Number.EPSILON) * 100) /
    100
  ).toFixed(2)
  setColumn("normalizedProduction", normalizedProduction)
  setValue("totalProduction", totalProduction)
  setValue(
    "negativeProduction",
    `Consumo nocturno anual inversor: ${Math.round(
      negativeProduction
    )} kWh (${inverterConsumptionOverTotalProduction}% de producción)`
  )
  setColumn("normalizedDates", dates)
  setColumn("normalizedCCH", cch)
  setColumn("normalizedComments", comments)
  setValue("conventionalTotal", cch.reduce(sumVector))
  setColumn("dates", dates)
  setColumn("conventionalCCH", cch)

  if (inverterConsumptionOverTotalProduction >= 1)
    try {
      SpreadsheetApp.getUi().alert(
        `Consumo nocturno del inversor superior al 1% del total de la producción.`
      )
    } catch (error) {
      console.error(error)
    }

  // Calculate and output relevant data: yearly consumption, monthly consumption and monthly peak consumption
  Tools.showMessage(
    `ℹ️ Consumos convencionales y producción`,
    `Calculando valores para el análisis...`,
    3
  )
  const monthlyConsumptionValues = Array.from(Array(12), () => new Array())
  const hourlyConsumptionValues = Array.from(Array(12), () =>
    Array.from(Array(24), () => new Array())
  )
  const daylyConsumption = Array.from(Array(12), () =>
    Array.from(Array(7), () => new Array(31).fill(0))
  )

  dates.forEach((date, index) => {
    var month = date.getMonth()
    var dayOfWeek = date.getDay()
    var day = date.getDate() - 1 // Zero-based index
    var hour = date.getHours()
    daylyConsumption[month][dayOfWeek][day] += cch[index]
    monthlyConsumptionValues[month].push(cch[index])
    hourlyConsumptionValues[month][hour].push(cch[index]) // Array of hourly values per month
  })

  const monthlyConsumption = monthlyConsumptionValues.map((e) =>
    e.reduce(sumVector)
  )
  const yearlyConsumption = monthlyConsumption.reduce(sumVector)
  const monthlyHourlyPeak = monthlyConsumptionValues.map((element) =>
    Math.max(...element)
  )
  const weeklyConsumptionMeans = daylyConsumption.map((month, monthIndex) =>
    daylyConsumption[monthIndex].map(
      (week, weekIndex) =>
        daylyConsumption[monthIndex][weekIndex].reduce(sumVector) /
        daylyConsumption[monthIndex][weekIndex].filter((x) => x != 0).length
    )
  )

  const surpassingValuesFactor = getValue("surpassingValuesFactor")
  const hourlyConsumptionMeans = hourlyConsumptionValues.map(
    (month, monthIndex) =>
      hourlyConsumptionValues[monthIndex].map(
        (hour, hourIndex) =>
          hourlyConsumptionValues[monthIndex][hourIndex].reduce(sumVector) /
          hourlyConsumptionValues[monthIndex][hourIndex].length
      )
  )
  const exceedingValues = hourlyConsumptionValues.map((month, monthIndex) =>
    hourlyConsumptionValues[monthIndex].map(
      (hour, hourIndex) =>
        hourlyConsumptionValues[monthIndex][hourIndex].filter(
          (x) =>
            x >
            hourlyConsumptionMeans[monthIndex][hourIndex] *
            surpassingValuesFactor
        ).length
    )
  )

  // Output data
  setValue("yearlyConsumption", yearlyConsumption)
  setColumn("monthlyConsumption", monthlyConsumption)
  setColumn("monthlyHourlyPeak", monthlyHourlyPeak)
  setValues("hourlyConsumptionMeans", hourlyConsumptionMeans)
  setValues("weeklyConsumptionMeans", weeklyConsumptionMeans)
  setValues("exceedingValues", exceedingValues)
}

function removeFebruary29th(cch) {
  // Remove data from February 29th, if there is any (we do not consider leap years)
  cch = cch.filter(
    (element) => !(element.date.getMonth() == 1 && element.date.getDate() == 29)
  )
}

function trimToLastNaturalYear(cch) {
  // Trim array to last natural year values
  const endingDate = new Date(cch[cch.length - 1].date)
  const startingDate = new Date(endingDate)
  startingDate.setFullYear(startingDate.getFullYear() - 1)
  startingDate.setTime(startingDate.getTime() + 3600000)
  const startingDateIndex = cch.findIndex(
    (element) => element.date.getTime() >= startingDate.getTime()
  )
  cch.splice(0, startingDateIndex)
}

function normalizeCNMCdst(cch) {
  const dstStartDateIndex = cch.findIndex((element) => isDstStartDate(element.date))
  const dstEndDateIndex = cch.findIndex((element) => isDstEndDate(element.date))

  const dstStartDate = cch[dstStartDateIndex].date
  const dstEndDate = cch[dstEndDateIndex].date

  // Take hours 2..23 of DST start day and move to next hour
  for (
    let index = dstStartDateIndex + 2, j = 1;
    index < dstStartDateIndex + 23;
    index++, j++
  ) {
    cch[index].date = new Date(dstStartDate.getTime() + 3600000 * (j + 1))
  }

  // Take hours 3..25 of DST start day and move to previous hour
  for (
    let index = dstEndDateIndex + 3, j = 1;
    index < dstEndDateIndex + 25;
    index++, j++
  ) {
    cch[index].date = new Date(dstEndDate.getTime() + 3600000 * (j + 2))
  }

  return cch
}

function normalizeInfoenergiaDst(cch) {
  // Sanitize DST end hour
  const startingYear = cch[1].date.getFullYear()
  const dstEndHour = getPreviousHour(
    new Date(getlastSunday((year = startingYear), (month = 10)).setHours(2))
  )
  const dstEndHourIndex = cch.findIndex(
    (element) => element.date.getTime() == dstEndHour.getTime()
  )
  if (cch[dstEndHourIndex + 1].date.getTime() == dstEndHour.getTime())
    cch[dstEndHourIndex + 1].date.setTime(
      cch[dstEndHourIndex + 1].date.getTime() + 3600000
    )
  return cch
}

function completeMissingValues(cch) {
  // Complete missing values for a full 8760-hours yearly cch
  let completedCCH = []
  for (var i = 0, j = 0; i < 8760; i++) {
    var currentHour = new Date(cch[0].date.getTime() + 3600000 * i)
    // Skip Febrary 29th
    if (currentHour.getMonth() == 1 && currentHour.getDate() == 29) continue

    var row
    if (j < cch.length && cch[j].date.getTime() == currentHour.getTime()) {
      row = { date: cch[j].date, value: cch[j].value, comment: "" }
      j++
    } else {
      row = {
        date: new Date(currentHour.getTime()),
        value: 0,
        comment: "Ausente en CSV",
      }
    }
    completedCCH.push(row)
  }
  return completedCCH
}

function getCsvMatrix(delimiter = ";") {
  const workingFolder = DriveApp.getFolderById(getParentFolderId())
  const csv = workingFolder
    .getFilesByName(getValue("CSVfilename"))
    .next()
    .getBlob()
    .getDataAsString()
  // Parse CSV, remove header and pre-trim to last year-ish values
  return Utilities.parseCsv(csv, delimiter).slice(1).splice(-8785)
}

function processInfoenergia() {
  const delimiter = ","
  const matrix = getCsvMatrix(delimiter)

  // Shift first element (belongs to 23-24h of previous day)
  matrix.shift()

  setValue("csvCUPS", matrix[0][0])

  // Create array of date and consumption (in kWh) pairs
  const cch = matrix.map((e) => {
    return {
      date: getPreviousHour(
        new Date(
          e[1].split("/")[2],
          e[1].split("/")[1] - 1, // MonthIndex is zero-based (January == 0)
          e[1].split("/")[0],
          e[2]
        )
      ),
      value: normalizeValue(e[3]),
    }
  })

  // Normalize CCH
  removeFebruary29th(cch)
  trimToLastNaturalYear(cch)
  normalizeInfoenergiaDst(cch)
  return completeMissingValues(cch)
}

function processDatadisCsv() {
  const delimiter = ";"
  const matrix = getCsvMatrix(delimiter)

  setValue("csvCUPS", matrix[0][0])

  // Create array of date and consumption (in kWh) pairs
  const cch = matrix.map((e) => {
    return {
      date: getPreviousHour(
        new Date(
          e[1].split("/")[0],
          e[1].split("/")[1] - 1, // MonthIndex is zero-based (January == 0)
          e[1].split("/")[2],
          e[2].slice(0, 2) // Get first two characters (original format is "01:00")
        )
      ),
      value: normalizeValue(e[3]),
    }
  })

  // Normalize CCH
  removeFebruary29th(cch)
  trimToLastNaturalYear(cch)
  normalizeInfoenergiaDst(cch) // Same end-DST normalization needed
  return completeMissingValues(cch)

}

function processDatadisAPI() {

  const nif = getValue("NIFTitular")
  const clientCups = getValue("CUPScliente")
  const startDate = getValue("datadisStartDate")
  const endDate = getValue("datadisEndDate")

  if (!(nif && clientCups && startDate && endDate)) {
    throw new Error('Los valores de NIF, CUPS y fechas de inicio y fin de Datadis no pueden estar vacías.')
  }


  // Get authentication token from DATADIS
  var auth = {
    username: DATADISUSER,
    password: DATADISPASSWORD
  }
  var loginOptions = {
    method: "POST",
    payload: auth,
    muteHttpExceptions: true,
    validateHttpsCertificates: false,
  }
  var loginResponse = UrlFetchApp.fetch("https://datadis.es/nikola-auth/tokens/login", loginOptions)
  if (loginResponse.getResponseCode() != 200) {
    throw new Error("Error en la API de Datadis al solicitar el token de autenticación: " + loginResponse.getContent())
  }
  var token = loginResponse.getContentText()



  // Get supplies from authorized NIF
  var headers = {
    "Authorization": "Bearer " + token,
    'Content-Type': "application/json", 'Accept': "application/json",
  }
  var payload = {
    authorizedNif: nif
  }
  var suppliesOptions = {
    method: "GET",
    headers: headers,
    muteHttpExceptions: true,
    validateHttpsCertificates: false,
  }
  var suppliesResponse = UrlFetchApp.fetch("https://datadis.es/api-private/api/get-supplies?authorizedNif=" + payload.authorizedNif, suppliesOptions)
  if (suppliesResponse.getResponseCode() != 200) {
    throw new Error("Error en la API de Datadis al solicitar los datos del contrato: " + suppliesResponse.getContent())
  }
  var datadisSupplies = JSON.parse(suppliesResponse.getContentText())

  // Select supply from list with same CUPS
  try {
    const suppliesList = datadisSupplies.map(e => e["cups"].slice(0, 20))
    var supplyIndex = suppliesList.indexOf(clientCups.slice(0, 20))
  } catch {
    throw "El listado de contratos obtenido no contiene el CUPS " + clientCups
  }
  const supply = datadisSupplies[supplyIndex]


  // Get consumption data from supply
  payload.cups = supply.cups
  payload.distributorCode = supply.distributorCode
  payload.startDate = startDate
  payload.endDate = endDate
  payload.measurementType = "0"
  payload.pointType = supply.pointType

  var consumptionOptions = {
    method: "get",
    headers: headers,
    muteHttpExceptions: true
  }

  var consumptionResponse = UrlFetchApp.fetch("https://datadis.es/api-private/api/get-consumption-data?authorizedNif=" + nif + "&cups=" + payload.cups + "&distributorCode=" + payload.distributorCode + "&startDate=2022/01&endDate=2022/12&measurementType=0&pointType=" + payload.pointType, consumptionOptions)

  if (consumptionResponse.getResponseCode() != 200) {
    throw new Error("Error en la API de Datadis al solicitar los datos de consumo: " + consumptionResponse.getContent())
  }
  var datadisCch = JSON.parse(consumptionResponse.getContentText())


  // Select CCH from list with same CUPS
  try {
    const cupsList = datadisCch.map(e => e["cups"].slice(0, 20))
    var cupsIndex = cupsList.indexOf(clientCups.slice(0, 20))
  } catch {
    throw "El listado de CUPS obtenido no contiene el CUPS " + clientCups
  }


  // Create array of date and consumption (in kWh) pairs
  setValue("csvCUPS", datadisCch[cupsIndex]["cups"])
  const cch = datadisCch.map((e) => {
    return {
      date: getPreviousHour(
        new Date(
          e["date"].split("/")[0],
          e["date"].split("/")[1] - 1, // MonthIndex is zero-based (January == 0)
          e["date"].split("/")[2],
          e["time"].slice(0, 2) // Get first two characters (original format is "01:00")
        )
      ),
      value: normalizeValue(e["consumptionKWh"]),
    }
  })

  // Normalize CCH
  removeFebruary29th(cch)
  trimToLastNaturalYear(cch)
  normalizeInfoenergiaDst(cch) // Same end-DST normalization needed
  return completeMissingValues(cch)

}



function processCNMC() {
  const matrix = getCsvMatrix()

  setValue("csvCUPS", matrix[0][0])

  // Create array of date (shifted to one less) and consumption (in kWh) pairs
  const cch = matrix.map((e) => {
    return {
      date: new Date(
        e[1].split("/")[2],
        e[1].split("/")[1] - 1, // MonthIndex is zero-based (January == 0)
        e[1].split("/")[0],
        e[2]
      ),
      value: normalizeValue(e[3]),
    }
  })

  normalizeDateConvention(cch)
  removeFebruary29th(cch)
  trimToLastNaturalYear(cch)
  normalizeCNMCdst(cch)
  return cch
}

function processREE() {
  const anualConsumption = getValue("REEconsumption")
  const normalizedCCH = getValues("REE").map((e) => {
    return {
      date: getPreviousHour(e[0]),
      value: normalizeValue(e[1]) * anualConsumption,
      comment: "",
    }
  })
  return normalizedCCH
}

function processSIPS2() {
  const additionalValues = getValues("SIPS2additional24")
  const sips2 = getValues("SIPS2").concat(
    additionalValues[0][0] != "" ? additionalValues : []
  )
  const normalizedCCH = sips2.map((e) => {
    return {
      date: getPreviousHour(e[0]),
      value: normalizeValue(e[1]),
      comment: "",
    }
  })

  removeFebruary29th(normalizedCCH)
  return normalizedCCH
}

function processSIPS3() {
  const additionalValues = getValues("SIPS3additional24")
  const sips3 = getValues("SIPS3").concat(
    additionalValues[0][0] != "" ? additionalValues : []
  )
  const normalizedCCH = sips3.map((e) => {
    return {
      date: getPreviousHour(e[0]),
      value: normalizeValue(e[1]),
      comment: "",
    }
  })
  removeFebruary29th(normalizedCCH)
  return normalizedCCH
}
