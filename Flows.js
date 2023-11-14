function processEnergyFlows() {
  // Comment
  Tools.showMessage(
    `ℹ️ Flujos energéticos y económicos`,
    `Calculando flujos energéticos...`
  )

  // Get CCH and Production curves
  const dates = getColumn("dates")
  const conventionalCCH = getColumn("conventionalCCH")
  const recurringCCH = getColumn("recurringCCH")
  const ashpCCH = getColumn("ashpCCH")
  const saveCCH = getColumn("saveCCH")
  const totalCCH = conventionalCCH.map(
    (e, index) => e + +recurringCCH[index] + +ashpCCH[index] + +saveCCH[index]
  )
  const production = getColumn("normalizedProduction")

  // Calculate self-consumption, surplus and grid demand
  const selfConsumption = totalCCH.map((e, i) => Math.min(e, production[i]))
  const surplus = totalCCH.map((e, i) => Math.max(0, production[i] - e))
  const gridDemand = totalCCH.map((e, i) => Math.max(0, e - production[i]))

  // Output data
  setColumn("totalCCH", totalCCH)
  setColumn("selfConsumption", selfConsumption)
  setColumn("surplus", surplus)
  setValue("totalConsumption", totalCCH.reduce(sumVector))
  setValue("totalSelfConsumption", selfConsumption.reduce(sumVector))
  setValue("totalSurplus", surplus.reduce(sumVector))
  setColumn("gridDemand", gridDemand)
  setValue("totalGridDemand", gridDemand.reduce(sumVector))

  SpreadsheetApp.flush()

  // Process and output anual values

  function calculateHourlyMean(arrayName, hourlyOutputName, monthlyOutputName) {
    const array = getColumn(arrayName)
    const arrayHourlyValues = Array.from(Array(24), () => new Array())
    const arrayMonthlyValues = Array.from(Array(12), () => new Array())
    dates.forEach((date, index) => {
      arrayHourlyValues[new Date(date).getHours()].push(array[index])
      arrayMonthlyValues[new Date(date).getMonth()].push(array[index])
    })
    const arrayHourlyMeans = arrayHourlyValues.map((row) =>
      row.reduce(sumVector)
    )
    const arrayMonthlyMeans = arrayMonthlyValues.map((row) =>
      row.reduce(sumVector)
    )
    setColumn(hourlyOutputName, arrayHourlyMeans)
    setColumn(monthlyOutputName, arrayMonthlyMeans)
  }

  const columns = [
    ["conventionalCCH", "conventionalMeans", "conventionalMonthlyMeans"],
    ["recurringCCH", "recurringMeans", "recurringMonthlyMeans"],
    ["ashpCCH", "ashpMeans", "ashpMonthlyMeans"],
    ["saveCCH", "saveMeans", "saveMonthlyMeans"],
    ["totalCCH", "totalCCHmeans", "totalCCHMonthlymeans"],
    ["normalizedProduction", "productionMeans", "productionMonthlyMeans"],
    ["selfConsumption", "selfConsumptionMeans", "selfConsumptionMonthlyMeans"],
    ["surplus", "surplusMeans", "surplusMonthlyMeans"],
    ["gridDemand", "gridDemandMeans", "gridDemandMonthlyMeans"],
  ]

  columns.forEach((e) => calculateHourlyMean(...e))
}

function processEconomicFlows() {
  Tools.showMessage(
    `ℹ️ Flujos energéticos y económicos`,
    `Calculando flujos económicos...`,
    5
  )

  const dates = getColumn("dates")
  const totalCCH = getColumn("totalCCH")

  // Get tariffs
  const tariff = getValue("tariff")
  const nationalHolidays = getColumn("nationalHolidays").map((e) => new Date(e))
  const tariffPeriods = dates.map((date) => getTariffPeriod(date, tariff, nationalHolidays))
  const tariffPrices = getTariffPrices(tariff)
  const TAXES = (1 + getValue("IVA")) * (1 + getValue("IEE"))

  // Regular demand, without PV
  const gridDemandCost = tariffPeriods.map(
    (period, i) => totalCCH[i] * tariffPrices[period - 1] * TAXES
  )
  const monthlyValues = Array.from(Array(12), () => new Array())
  dates.forEach((date, index) => {
    monthlyValues[new Date(date).getMonth()].push(gridDemandCost[index])
  })
  const monthlyBill = monthlyValues.map((e) => e.reduce(sumVector))
  setColumn("monthlyBill", monthlyBill)

  // Demand with PV
  const gridDemandPV = getColumn("gridDemand")
  const gridDemandCostPV = tariffPeriods.map(
    (period, i) => gridDemandPV[i] * tariffPrices[period - 1] * TAXES
  )
  const monthlyValuesPV = Array.from(Array(12), () => new Array())
  dates.forEach((date, index) => {
    monthlyValuesPV[new Date(date).getMonth()].push(gridDemandCostPV[index])
  })
  const monthlyBillPV = monthlyValuesPV.map((e) => e.reduce(sumVector))
  setColumn("monthlyBillPV", monthlyBillPV)

  // Demand with PV and surplus compensation
  const energeticaTariffCompensation = getValue("energeticaTariffCompensation")
  const surplusMonthlyMeans = getColumn("surplusMonthlyMeans")
  var compensableSurplusMeans = []
  const monthlyCompensation = monthlyBillPV.map((monthlyCost, index) =>
    Math.min(monthlyCost, surplusMonthlyMeans[index] * energeticaTariffCompensation * TAXES))

  const monthlyBillComplete = monthlyBillPV.map((monthlyCost, index) => {
    compensableSurplusMeans[index] = monthlyCompensation[index] / (energeticaTariffCompensation * TAXES)
    return monthlyCost - monthlyCompensation[index]
  })
  setColumn("monthlyBillComplete", monthlyBillComplete)
  setColumn("compensableSurplusMeans", compensableSurplusMeans)

  // Demand with PV and limitless surplus
  const monthlyBillLimitless = monthlyBillPV.map((monthlyCost, index) => {
    const monthlyCompensation =
      surplusMonthlyMeans[index] * energeticaTariffCompensation * TAXES
    return monthlyCost - monthlyCompensation
  })
  setColumn("monthlyBillLimitless", monthlyBillLimitless)

  SpreadsheetApp.flush()


  // Demand with PV, surplus compensation and Som Energia's "Flux Solar"
  const totalBillBefore = getValue("facturaAntes")
  const fluxCoefficient = getValue("fluxCoefficient")
  const monthlyBillBeforeFlux = [...monthlyBillComplete]
  const monthlyConvertedToSoles = monthlyBillComplete.map((monthlyCost, index) => (monthlyCost - monthlyBillLimitless[index]))
  setColumn("monthlyConvertedToSoles", monthlyConvertedToSoles)
  const monthlySolesInput = monthlyConvertedToSoles.map(e => e * fluxCoefficient)

  var solesQueue = []
  var anualSavings = []

  var currentBill = []
  for (let year = 0; year < 25; year++) { // For each year in 25 years

    currentBill = [...monthlyBillBeforeFlux]
    for (let month = 0; month < 12; month++) { // For each month

      const currentMonth = month + 12 * year
      console.log(`******** Paso ${currentMonth.toString()}: mes ${month}, año ${year} ********`)
      console.log("   Cola de soles: " + JSON.stringify(solesQueue))
      console.log("   Factura mensual antes de Flux: " + currentBill[month])

      // Extract soles and generate monthly discout
      while (solesQueue.length > 0) { // While there are soles in queue
        const oldestSoles = solesQueue.shift() // Take oldest soles
        if (currentMonth - oldestSoles.monthOfGeneration < 60) { // If those soles are still valid

          if (currentBill[month] >= oldestSoles.value) { // If all soles are discountable
            currentBill[month] -= oldestSoles.value
            console.log("   Consumidos soles totales: " + oldestSoles.value.toString())
          } else { // If there is an excess of soles
            oldestSoles.value -= currentBill[month]
            console.log("   Consumidos soles parciales: " + currentBill[month])
            solesQueue.unshift(oldestSoles)
            currentBill[month] -= currentBill[month]
            break
          }
          console.log(` Factura mensual: ${currentBill[month]}`)

          if (currentBill[month] == 0) break // If bill is already zero, stop 

        } else { // If soles not valid anymore
          console.log("   Soles caducados, se eliminan: " + oldestSoles.value)
        }

      }

      // Generate soles for this month and add them to the queue
      if (monthlySolesInput[month] > 0) {
        const soles = {
          "value": monthlySolesInput[month],
          "monthOfGeneration": currentMonth
        }
        solesQueue.push(soles)
        console.log("   Generados soles: " + soles.value.toString())
      }

    }

    anualSavings.push(totalBillBefore - currentBill.reduce(sumVector)) // Ahorro anual: factura antes - total de factura después

  }

  setColumn("monthlyBillFlux", currentBill)
  setColumn("ahorrosAnuales", anualSavings)

}


function processAllFlows() {
  processEnergyFlows()
  processEconomicFlows()
}