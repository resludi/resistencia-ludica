function getNextDate() {
  const MONTHS = [
    'Enero',
    'Febrero',
    'Marzo',
    'Abril',
    'Mayo',
    'Junio',
    'Julio',
    'Agosto',
    'Septiembre',
    'Octubre',
    'Noviembre',
    'Diciembre',
  ];
  const now = new Date();
  now.setMonth(now.getMonth() + 1);
  return `${MONTHS[now.getMonth()]}_${now.getFullYear().toString().substring(2)}`;
}

function formatDate(dateObj) {
  if (typeof dateObj === 'string') return dateObj;
  return dateObj.toLocaleDateString(
    'es-ES', 
    {timeZone: 'Europe/Madrid', day: '2-digit', month: '2-digit', year: 'numeric'}
  )
}

function formatTime(dateObj) {
  if (typeof dateObj === 'string') return dateObj.replaceAll(' ','');
  return dateObj.toLocaleTimeString(
    'es-ES',
    {timeZone:'Europe/Madrid', hour12: false, hour: '2-digit', minute: '2-digit'}
  );
}

function cestOrCet(dateObj) {
  dateStr = dateObj.toLocaleString(
    'es-ES',
    {timeZone: 'Europe/Madrid', timeZoneName: "short"}
  )
  // FIXME: esto es una chapuza para sacar el timezone del string
  return dateStr.split(' ')[2];
}

function createDoc() {
  const VERSION = '1.2.0';
  const SPREADSHEET_RANGE = 'A:U';
  const CUSTOMER_SPREADSHEET = '15r8Bg2x5piFcq-jDGm9EDB1CJEJUrmmAZfJRwRZLuKU';
  const data = SpreadsheetApp.openById(CUSTOMER_SPREADSHEET);

  const sheet = data.getSheets()[0];

  const dataRange = sheet.getRange(SPREADSHEET_RANGE);
  const values = dataRange.getValues();
  
  // Remove empty rows
  var content = values
    .filter(item => {
      return item.join('') !== '';
    });
  
  content.shift();
  const docTitle = `RL_${getNextDate()}`;

  const doc = DocumentApp.create(docTitle);
  var body = doc.getBody();
  
  for (let i=0; i<content.length; i++) {
    const item = content[i];
    const title = `${item[6]} (${item[8]})`;
    
    if (i === 0) {
      body.getChild(0).asParagraph().setText(title);
      body.getChild(0).setBold(true);
    } else {
      body.appendParagraph(title).setBold(true);
    }

    body.appendParagraph('').setBold(false);
    body.appendParagraph(`<strong>Sinopsis</strong>: ${item[13]}`);
    body.appendParagraph(`<strong>Ambientación</strong>: ${item[7]}`);
    body.appendParagraph(`<strong>Sistema de juego</strong>: ${item[8]}`);
    body.appendParagraph(`<strong>Plazas</strong>: mínimo ${item[9]}, máximo ${item[10]}`);
    body.appendParagraph(`<strong>Idioma/s</strong>: ${item[11]}`);
    body.appendParagraph(`<strong>Aviso de contenido</strong>: ${item[12]}`);
    body.appendParagraph(`<strong>Día</strong>: ${item[1]} ${formatDate(item[0])} de ${formatTime(item[2])} a ${formatTime(item[3])} (${cestOrCet(item[0])})`);

    if (item[14]) {
      body.appendParagraph(`<strong>Experto</strong>: Se requieren conocimientos previos de sistema y ambientación.`);
    }

    if (item[15]) {
      body.appendParagraph(`<strong>Trasfondo</strong>: Contacta con el organizador a través de Discord antes de la partida.`);
    }

    if (item[16]) {
      body.appendParagraph(`<strong>Mecánicas</strong>: Hay cambios sustanciales en las mecánicas de juego.`);
    }

    if (item[17]) {
      body.appendParagraph(`<strong>Múltiple</strong>: Esta partida se juega en varias sesiones. Atento al título y a la descripción.`);
    }

    if (item[18]) {
      body.appendParagraph(`<strong>Campaña</strong>: Esta partida forma parte de una campaña. Atento al icono del título y a la descripción.`);
    }

    if (item[19]) {
      body.appendParagraph(`<strong>Grabación</strong>: La partida se grabará.`);
    }

    if (item[20]) {
      body.appendParagraph(`<strong>Emisión</strong>: La partida se emitirá.`);
    }
    
    body.appendParagraph(`${item[4]} (Discord: ${item[5]})`);
    body.appendPageBreak();
  }
}

function onOpen(e) {
  SpreadsheetApp.getUi()
      .createMenu('Exportar partidas')
      .addItem('Google DOC', 'createDoc')
      .addToUi();
}


