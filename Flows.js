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
  const monthlyBillComplete = monthlyBillPV.map((monthlyCost, index) => {
    const monthlyCompensation = Math.min(
      monthlyCost,
      surplusMonthlyMeans[index] * energeticaTariffCompensation * TAXES
    )
    compensableSurplusMeans[index] = monthlyCompensation / (energeticaTariffCompensation * TAXES)
    return monthlyCost - monthlyCompensation
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

}


function processAllFlows() {
  processEnergyFlows()
  processEconomicFlows()
}