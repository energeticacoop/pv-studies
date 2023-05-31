function getFinalStudyTemplates() {
  return [
    {
      templateName: getValue("nombreFichero"),
      templateId: "1ZAxidSJsBKZvLUDMJ5mJBvkRVmmG5OakZk9--3g43vg",
      exportToPDF: false,
      copyComments: true,
      folder: "folder01",
    },
  ]
}

function get00FolderTemplates() {
  const templates = [
    {
      templateName: "DROU - Acreditación de la representación",
      templateId: "1yGMRXtlaPxqdQqUzeaJO8PFFleWaWgUzxm6Zva1u2GU",
      exportToPDF: true,
      folder: "folder0200",
    },
    {
      templateName: "Contrato de instalación",
      templateId: "1ElIx5TtbT7y7iUawq8IC6W1KpR7e9hLeABwHkaD15S8",
      exportToPDF: false,
      folder: "folder0200"
    },
    {
      templateName: "BIE  - Acreditación de la representación",
      templateId: "1EdUm7oKmbVTkHy-wS2VdLkA4MjJNxPw1AbOk-nEKaN4",
      exportToPDF: true,
      folder: "folder0200"
    },
  ]

  if (inValladolid()) templates.push(
    {
      templateName: "VLL - Solicitud inscripción Industria ST",
      templateId: "1cC0bJJjOynwr_3dQq0H26SILowwXflj4Kv6OzD2rp6g",
      exportToPDF: true,
      folder: "folder0200"
    }
  )

  return templates
}


function get01FolderTemplates() {
  const templates = [
    {
      templateName: "Presupuesto Ejecución Material - Ayto",
      templateId: "1Q7aRFNjqB_Kc9daRPDRhs2kqt7XuZMYmhGRQjLVHBY8",
      exportToPDF: true,
      copyComments: false,
      folder: "folder0201",
    },
    {
      templateName: "Compromiso dirección de obra",
      templateId: "1Pk3IS0Uz1Jj7R-AyNff5F0rqPJrBulJTYsQU2DYET5k",
      exportToPDF: true,
      folder: "folder0201"
    }, {
      templateName: "Declaración responsable Técnico",
      templateId: "1Yz2U2J5zfDCPm_abDMUKVnmZyW6AFVHKUMwYTdFwyCE",
      exportToPDF: true,
      folder: "folder0201"
    },
    {
      templateName: "Solicitud bonificación ICIO",
      templateId: "1TPeCE6wXHIQuVxVvjZyYcME6H83hPfO6QJV9WNmBPnk",
      exportToPDF: true,
      folder: "folder0201"
    },
    {
      templateName: "Comunicación ambiental",
      templateId: "1DAlTAw-pZk6Rd8-lT818PFkaIthy0PLSKKvrzEIocTA",
      exportToPDF: false,
      copyComments: true,
      folder: "folder0201"
    },
  ]

  if (!withProject()) templates.push(
    {
      templateName: "Certificado final de obra",
      templateId: "1Pgytz6cEZlbotL2St31JJDTbir0gUi4rP2ipGb4qPac",
      exportToPDF: false,
      folder: "folder0201"
    },
  )
  return templates
}


function getMemoryTemplates() {
  return [
    {
      templateName: "Memoria técnica",
      templateId: "1uxoE7-2P41jCHrwnVDRCFHFreeU7NgsI7uonWOVDGAc",
      exportToPDF: false,
      copyComments: true,
      folder: "folder0201",
    },
  ]
}

function getGuideTemplates() {
  return [
    {
      templateName: "Guía de instalación",
      templateId: "1k9lYTmxMsINC6U49w1NWOu0-3dGkYUBF5yrhKmkGb04",
      exportToPDF: false,
      folder: "folder03",
    },
  ]
}

function getBillTemplates() {
  return [
    {
      templateName: "Factura",
      templateId: "1xttGTW5xY0mnyuCEU8gEQwQ926WH4vTG0mZS4CEKuJQ",
      exportToPDF: false,
      folder: "folder02",
    },
  ]
}

function get02FolderTemplates() {
  const templates = [
    {
      templateName: "Manual usuario",
      templateId: "1J_RAwFGlkYCnlTqICXu3i0ogmf0xN48n8bjw1mH4S0o",
      exportToPDF: true,
      copyComments: false,
      folder: "folder0202"
    },
    {
      templateName: "Certificado reconocimiento instalación",
      templateId: "1NZhggc5SLkA2HxnjpamfgoPqCSYuKk1h0KLOJN2H-48",
      exportToPDF: false,
      copyComments: true,
      folder: "folder0202"
    }
  ]
  if (inValladolid()) {
    templates.push(
      {
        templateName: "VLL - Anexo I - Instrucción autoconsumo abril 2020",
        templateId: "1AuSnyW8ypVFlTCSj3dcW61a4T8NptULR4UBWKpuoD3E",
        exportToPDF: false,
        copyComments: true,
        folder: "folder0202"
      }
    )
  } else {
    templates.push(
      {
        templateName: "Anexo I datos de autoconsumo v.3",
        templateId: "1POtr51vwXX5qXUotJU5gD3jge2AMPcvJTr_QQfJ65RE",
        exportToPDF: false,
        copyComments: true,
        folder: "folder0202",
      }
    )
  }

  return templates
}



function get03FolderTemplates() {
  const templates = []
  if (inValladolid()) {
    templates.push(
      {
        templateName: "VLL - Plantilla A103 ABRIL - Remisión de información AC",
        templateId: "1F-JUhAtSpSxNoWs25Afe6XZb4-5IkzeGJJyxXXB7fHw",
        exportToPDF: false,
        copyComments: true,
        folder: "folder0203",
      }
    )
  } else {
    templates.push(
      {
        templateName: "Formulario Autoconsumo",
        templateId: "1icBoMi4MCBoNVakzXIg_2WneNDGXmTAbJOXX3RY3ZqI",
        exportToPDF: false,
        copyComments: true,
        folder: "folder0203"
      }
    )
  }

  if (inLeon()) {
    templates.push(
      {
        templateName: "Datos establecimiento industrial León",
        templateId: "11iaYzBMwd1QiniKKt9HdI6FRpo7E41ZVj0gijEujV1k",
        exportToPDF: false,
        copyComments: true,
        folder: "folder0203"
      }
    )
  }

  return templates
}




function get04FolderTemplates() {
  const templates = []
  if (withProject()) templates.push(...[
    {
      templateName: "Certificado final de obra COIIM",
      templateId: "1Yi3ORUrh_-HO6eFVoi1kftE3PsHDIe6lTvvSkGBa084",
      exportToPDF: false,
      folder: "folder0204"
    },
    {
      templateName: "Asume Coordinación Seguridad y Salud",
      templateId: "16VQsGA3z-Q--cdLx4U-Nqye1IOAIbq1k86vWGfkkOco",
      exportToPDF: false,
      folder: "folder0204"
    },
    {
      templateName: "Asume Dirección Técnica",
      templateId: "1j7Ozyh6P8-Mj-pW1noTKTVPTwNco89jy-eJ0HYSbMJQ",
      exportToPDF: false,
      folder: "folder0204"
    },
    {
      templateName: "Acta Aprobación Plan SyS",
      templateId: "1QZ0q4Bibj_rRAUvOpAK0fW29VfjdlhiyCCSaT9RYFsU",
      exportToPDF: true,
      folder: "folder0204"
    },
    {
      templateName: "Plan de Seguridad y Salud",
      templateId: "1DjSwmbcsmTj2mhn7o4381f5X4t90K7cVPwxzCV9hjG8",
      exportToPDF: false,
      folder: "folder0204"
    },
    {
      templateName: "Requisitos para obras de proyectos",
      templateId: "1yTK-NfmzCabqI6M8-BaDRfiNENB2KNQrsqViJYECzT8",
      exportToPDF: false,
      folder: "folder0204"
    },
    {
      templateName: "Coordinador de actividades preventivas",
      templateId: "1AdFMygEZLIILYHOTsSEvYDFNy5QxBWOqEMUFFTOAgMM",
      exportToPDF: false,
      folder: "folder0204"
    },
    {
      templateName: "Coordinador de actividades preventivas empresas concurrentes",
      templateId: "1FXxZLvTLiieihKDwIeEvLv2jHi6K82GR6DdaJpVwfPk",
      exportToPDF: false,
      folder: "folder0204"
    },
    {
      templateName: "En caso de accidente, acudir a",
      templateId: "18CcpHMTNTcigSFjPun4Kkp_8rYuRKKPu8gr8MEjzGWs",
      exportToPDF: false,
      folder: "folder0204"
    },
    {
      templateName: "Itinerario evacuaciones de accidentados",
      templateId: "19UpA6bcso9_v00Kwg7VXOQofgvVKmWi0fGFBmvB6dJE",
      exportToPDF: false,
      folder: "folder0204"
    },
    {
      templateName: "Normas de autorización uso maquinaria",
      templateId: "1TMws2DavUsAJx2TTA_iQ6Vo8tKu06ko5Qy6bQ53wFhM",
      exportToPDF: false,
      folder: "folder0204"
    },
    {
      templateName: "Recurso preventivo - Actividades que desempeñar",
      templateId: "1597H4SXlMAm7621FdNQETlro1Muy7KGOitI8FXmu4Qo",
      exportToPDF: false,
      folder: "folder0204"
    },
    {
      templateName: "Verificación de entrega EPIs",
      templateId: "1CJ2Nj0Py7QFUJr5r8t_fo5u14fw2KpJmssHWoMfE7Ug",
      exportToPDF: false,
      folder: "folder0204"
    },
  ]
  )
  return templates
}


function getTestTemplates() {
  return [
    {
      templateName: "Test",
      templateId: "1QXPT8zrOuKfwfVPMBqLP0SKJNW3ldRJgvDpSui8qfsM",
      exportToPDF: false,
      folder: "folder03",
    },
  ]
}




function inValladolid() {
  return SpreadsheetApp.getActiveSpreadsheet().getRangeByName("province").getValue().trim() == "Valladolid"
}

function inLeon() {
  return SpreadsheetApp.getActiveSpreadsheet().getRangeByName("province").getValue().trim() == "León"
}

function withProject() {
  return SpreadsheetApp.getActiveSpreadsheet().getRangeByName("withproject").getValue()
}
