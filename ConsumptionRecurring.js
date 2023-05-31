function processRecurringConsumption() {
  Tools.showMessage(
    `ℹ️ Consumos adicionales`,
    `Procesando consumos recurrentes...`
  )

  const dates = getColumn("dates")
  const recurringDayly = getColumn("recurringDayly")
  const recurringMonthly = getValues("recurringMonthly")
  const recurringWeekly = getValues("recurringWeekly")
  const recurringSeasonly = getValues("recurringSeasonly")

  const nationalHolidays = getColumn("nationalHolidays").map((e) => new Date(e))

  const recurringCCH = dates.map((date) => [
    +recurringDayly[date.getHours()] +
    +recurringMonthly[date.getHours()][date.getMonth()] +
    +recurringWeekly[date.getHours()][(date.getDay() + 6) % 7] +
    +recurringSeasonly[date.getHours()][
    isWeekendOrNationalHoliday(date, nationalHolidays)
      ? getSeason(date) * 2 + 1
      : getSeason(date) * 2
    ]
  ])

  // Output data
  setValues("recurringCCH", recurringCCH)
  setValue("recurringTotal", recurringCCH.map((e) => e[0]).reduce(sumVector))
}

function processASHP() {
  Tools.showMessage(
    `ℹ️ Consumos adicionales`,
    `Procesando consumos de aerotermia...`
  )

  // Calculate ASHP load curve
  const dates = getColumn("dates")
  const ashpCurves = getValues("ASHPcurves")
  const ashpCurveName = getValue("ASHPcurveName")
  const ashpIndex = ashpCurves[0].findIndex((e) => e == ashpCurveName)
  const ashpCurve = ashpCurves.slice(1).map((e) => e[ashpIndex])
  const ashpConsumption = getValue("ASHPconsumption")
  const ashpCCH = ashpCurve.map((e) => +e * ashpConsumption)

  // Displace load curve to match consumption dates
  const firstDateIndex = dates.findIndex(
    (e) => e.getMonth() == 0 && e.getDate() == 1 && e.getHours() == 0
  )
  const rotatedASHPcch = rotateArray(ashpCCH, firstDateIndex)

  // Output date
  setColumn("ashpCCH", rotatedASHPcch)
  setValue("ashpTotal", rotatedASHPcch.reduce(sumVector))
}

function processSAVE() {
  Tools.showMessage(`ℹ️ Consumos adicionales`, `Procesando consumos de SAVE...`)

  // Get data
  const dates = getColumn("dates")
  const gridUsage = getValues("SAVEgridUsage")
  const seasonsKm = getValues("SAVEseasonsKm")
  const batteryCapacity = getValue("SAVEbatteryCapacity")
  const consumptionPerKm = getValue("SAVEconsumptionPer100Km") / 100
  const nationalHolidays = getColumn("nationalHolidays").map((e) => new Date(e))

  // Construct save load availability curve and zeroes curve
  var zeroes = {}
  const saveLoadAvailability = dates.map((date) => {
    const gridUsageValue =
      gridUsage[date.getHours()][
      isWeekendOrNationalHoliday(date, nationalHolidays)
        ? getSeason(date) * 2 + 1
        : getSeason(date) * 2
      ]
    const zeroesIndex = new Date(
      date.getFullYear(),
      date.getMonth(),
      date.getDate()
    ).getTime()
    if (0 == gridUsageValue)
      zeroes[zeroesIndex] = isNaN(zeroes[zeroesIndex])
        ? 1
        : zeroes[zeroesIndex] + 1
    return gridUsageValue
  })

  // Construct battery usage curve
  const seasonsConsumption = seasonsKm[0].map((e) => +e * consumptionPerKm)
  const batteryUsage = dates.map((date, index) => {
    const zeroesIndex = new Date(
      date.getFullYear(),
      date.getMonth(),
      date.getDate()
    ).getTime()
    const numberOfZeroes = zeroes[zeroesIndex]

    const saveUsageValue = saveLoadAvailability[index]
      ? 0
      : isWeekendOrNationalHoliday(date, nationalHolidays)
        ? seasonsConsumption[getSeason(date) * 2 + 1] / numberOfZeroes
        : seasonsConsumption[getSeason(date) * 2] / numberOfZeroes
    return saveUsageValue
  })

  // Get partial grid load and surplus curves
  const conventionalCCH = getColumn("conventionalCCH")
  const recurringCCH = getColumn("recurringCCH")
  const ashpCCH = getColumn("ashpCCH")
  const partialCCH = conventionalCCH.map(
    (e, index) => +e + +recurringCCH[index] + +ashpCCH[index]
  )
  const production = getColumn("normalizedProduction")
  const partialSurplus = partialCCH.map((e, index) =>
    Math.max(0, production[index] - e)
  )
  const partialGridDemand = partialCCH.map((e, index) =>
    Math.max(0, e - production[index])
  )

  // Get maximum charge power for each tariff period as max of contracted power, SAVE power and normalized power
  const electricalInstallationType = getValue("electricalInstallationType")
  const saveContractedPowers = getValues("SAVEcontractedPowers")
  const maxSAVEpower = getValue("maxSAVEpower")

  const normalizedPowersArray =
    getValue("saveType") == "Orbis"
      ? getValues("normalizedPowersOrbis")
      : getValues("normalizedPowersFronius")
  const normalizedPowers =
    electricalInstallationType == "Monofásica"
      ? normalizedPowersArray.map((e) => e[0])
      : normalizedPowersArray.map((e) => e[1])
  const maxChargePowers = saveContractedPowers.map((saveContractedPower) => {
    const inferiorNormalizedPowerIndex =
      normalizedPowers.findIndex((e) => e > saveContractedPower) - 1
    const inferiorNormalizedPower =
      inferiorNormalizedPowerIndex < 0
        ? Number.MAX_VALUE
        : normalizedPowers[inferiorNormalizedPowerIndex]
    return Math.min(saveContractedPower, maxSAVEpower, inferiorNormalizedPower)
  })

  // Calculate SAVE CCH curve
  let batteryStatus = batteryCapacity
  let batteryNegativeEnergy = 0
  let batteryNegativeOccurrencies = 0
  const tariff = getValue("tariff")
  const saveCCH = new Array(8760)
  batteryUsage.forEach((hourlyLoad, index) => {
    if (hourlyLoad) {
      // Spend energy from EV battery by taking the hourly load
      batteryStatus -= Math.max(0, hourlyLoad)
      // Register when battery losses all energy
      if (hourlyLoad < 0) {
        batteryNegativeEnergy += hourlyLoad
        batteryNegativeOccurrencies++
      }
      saveCCH[index] = 0
    } else {
      if (saveLoadAvailability[index] == 1) {
        // Recharge EV battery
        // In cheap periods, charge from the grid. Otherwise, charge from FV surplus
        const datePeriod = getTariffPeriod(
          new Date(dates[index]),
          tariff,
          nationalHolidays
        )
        const nightPeriod =
          ((tariff == "2.0TD" && datePeriod == 3) || datePeriod == 6)
        const availableEnergy = nightPeriod
          ? Math.max(
            0,
            saveContractedPowers[datePeriod - 1] - partialGridDemand[index]
          )
          : Math.min(partialSurplus[index], maxSAVEpower)
        const actualDemand = Math.min(
          batteryCapacity - batteryStatus,
          availableEnergy,
          maxChargePowers[datePeriod - 1]
        )
        batteryStatus += actualDemand
        saveCCH[index] = actualDemand
      } else {
        saveCCH[index] = 0
      }
    }
  })

  //saveCCH
  setColumn("saveCCH", saveCCH)
  setValue("saveTotal", saveCCH.reduce(sumVector))
  setValue(
    "batteryNegativeEnergy",
    `El V.E. se queda sin carga ${batteryNegativeOccurrencies} horas al año. Faltan un total de ${Math.round(
      batteryNegativeEnergy
    )
      .toString()
      .replace(".", ",")} kWh.`
  )
}


function processAllConsumptions() {
  processConventionalConsumptionAndProduction()
  processRecurringConsumption()
  processASHP()
  processSAVE()
}