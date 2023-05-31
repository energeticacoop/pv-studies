function generateEmail() {
  clearRange("helpscoutEmail")
  const name = getValue("Nombre")
  const emailAddress = getValue("email")
  const emailBody = getEmailBody(name)
  const helpscoutURL = getHelpscoutURL(name, emailAddress, "☀️ Estudio definitivo de instalación fotovoltaica", sanitizeEmailString(emailBody))
  setURL(getRangeByName("helpscoutEmail"), helpscoutURL, "Enlace a email automatizado en Helpscout")
}


function getEmailBody(name) {

return `Hola, ${name}:<br><br>

Te escribimos, después de haber realizado la correspondiente visita técnica, para remitirte tu estudio definitivo, que incluye la propuesta de instalación fotovoltaica "llave en mano", el presupuesto con explicaciones detalladas acerca de los equipos y materiales que empleamos, trámites, tasas e impuestos incluidos, y aclaraciones sobre las garantías de la instalación y los equipos.<br><br>

Hemos realizado este estudio considerando tu consumo eléctrico actual y <strong>[‼️👁️✏️‼️]</strong><br><br>

Si esta propuesta te parece interesante, contesta por favor a este correo para que podamos comenzar a preparar tu expediente y la documentación necesaria para iniciar la tramitación de tu futura instalación fotovoltaica.<br><br>

¡Gracias por tu confianza en Energética Coop! Recibe un cordial saludo,<br><br>`

}


function sanitizeEmailString(text){
  return text.toString().replace("+","%2b").replace("+","%2b")
}


/**
 * Get a Helpscout URL for a new preformed ticket in Energética Autoconsumo's mailbox with the provided data 
 * @param {string} name of the recipient
 * @param {string} email address of the recipient
 * @param {string} subject for the email
 * @param {string} body content for the email
 * @return {string} a URL for the Helpscout preformed ticket
 */
function getHelpscoutURL(name, emailAddress, subject, body){
  const MAILBOX_SLUG = "238693defaf37e75";
  const formattedName = name.trim().replace(" ", "+");
  const formattedSubject = subject.trim().replace(" ", "+");
  return `https://secure.helpscout.net/mailbox/${MAILBOX_SLUG}/new-ticket/?name=${formattedName}&email=${emailAddress.trim()}&subject=${formattedSubject}&body=${body}`
}


