const MONTHS = [
  'enero',
  'febrero',
  'marzo',
  'abril',
  'mayo',
  'junio',
  'julio',
  'agosto',
  'septiembre',
  'octubre',
  'noviembre',
  'diciembre',
];

function getNextDate() {
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

function getGamePublishDate(dateObj) {
  // 👇️ clone date object, so we don't mutate it
  const date = new Date(dateObj);
  const day = date.getDay(); // 👉️ get day of week

  // 👇️ day of month - day of week (-6 if Sunday), otherwise +1
  const diff = date.getDate() - day + (day === 0 ? +1 : 1) - 7;

  return new Date(date.setDate(diff));
}

/*function createDoc() {
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
  
  for (let i=1; i<content.length; i++) {
    const item = content[i];
    const title = `${item[6]} (${item[8]})`;
    const eventDate = formatDate(item[0]);
    const eventStartTime = formatTime(item[2]);
    const eventEndTime = formatTime(item[3]);
    const eventTz = cestOrCet(item[0]);
    
    if (i === 0) {
      body.getChild(0).asParagraph().setText(title);
      body.getChild(0).setBold(true);
    } else {
      body.appendParagraph(title).setBold(true);
    }

    body.appendParagraph('').setBold(false);
    body.appendParagraph(`<strong>Sinopsis</strong>: ${item[14]}`);
    body.appendParagraph(`<strong>Ambientación</strong>: ${item[7]}`);
    body.appendParagraph(`<strong>Sistema de juego</strong>: ${item[8]}`);
    body.appendParagraph(`<strong>Plazas</strong>: mínimo ${item[9]}, máximo ${item[10]}`);

    const sessions = item[12] ? item[12] : 1;
    body.appendParagraph(`<strong>Número de sesiones</strong>: ${sessions}`);

    if (item[15]) {
      body.appendParagraph(`<strong>Plataforma</strong>: ${item[15]}`);
    }

    const language = item[11] ? item[11] : 'ESP';
    body.appendParagraph(`<strong>Idioma/s</strong>: ${language}`);

    const contentWarning = item[13] ? item[13] : 'Sin contenido sensible';
    body.appendParagraph(`<strong>Aviso de contenido</strong>: ${contentWarning}`);
    
    body.appendParagraph(`<strong>Día</strong>: ${item[1]} ${eventDate} de ${eventStartTime} a ${eventEndTime} (${eventTz})`);

    const principiantes = ((item[16] == true) ? 'Sí' : 'No');
    body.appendParagraph(`<strong>Apta para principiantes</strong>: ` + principiantes);

    if (item[17]) {
      body.appendParagraph(`<strong>Trasfondo</strong>: Contacta con el organizador a través de Discord antes de la partida.`);
    }

    if (item[18]) {
      body.appendParagraph(`<strong>Mecánicas</strong>: Hay cambios sustanciales en las mecánicas de juego.`);
    }

    // if (item[18]) {
    //   body.appendParagraph(`<strong>Múltiple</strong>: Esta partida se juega en varias sesiones. Atento al título y a la descripción.`);
    // }

    if (item[19]) {
      body.appendParagraph(`<strong>Campaña</strong>: Esta partida forma parte de una campaña. Atento al icono del título y a la descripción.`);
    }

    if (item[20]) {
      body.appendParagraph(`<strong>Grabación</strong>: La partida se grabará.`);
    }

    if (item[21]) {
      body.appendParagraph(`<strong>Emisión</strong>: La partida se emitirá.`);
    }
    
    const publishDate = getGamePublishDate(item[0])
    body.appendParagraph(
      `¡Atención! Las inscripciones para esta partida se abrirán el próximo lunes ${publishDate.getDate()} de ${MONTHS[publishDate.getMonth()]} a las 22:00 (${eventTz}).`
    );
    body.appendParagraph('');
    body.appendParagraph(`<!-- ${item[4]} (Discord: ${item[5]}) -->`);
    body.appendPageBreak();
  }
}*/

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
  
  for (let i=1; i<content.length; i++) {
    const item = content[i];
    const title = `${item[6]} (${item[8]})`;
    const eventDate = formatDate(item[0]);
    const eventStartTime = formatTime(item[2]);
    const eventEndTime = formatTime(item[3]);
    const eventTz = cestOrCet(item[0]);
    
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

    if (item[14]) {
      body.appendParagraph(`<strong>Plataforma</strong>: ${item[14]}`);
    }

    const language = item[11] ? item[11] : 'ESP';
    body.appendParagraph(`<strong>Idioma/s</strong>: ${language}`);

    const contentWarning = item[12] ? item[12] : 'Sin contenido sensible';
    body.appendParagraph(`<strong>Aviso de contenido</strong>: ${contentWarning}`);
    
    body.appendParagraph(`<strong>Día</strong>: ${item[1]} ${eventDate} de ${eventStartTime} a ${eventEndTime} (${eventTz})`);

    const principiantes = ((item[15] == true) ? 'Sí' : 'No');
    body.appendParagraph(`<strong>Apta para principiantes</strong>: ` + principiantes);

    if (item[16]) {
      body.appendParagraph(`<strong>Trasfondo</strong>: Contacta con el organizador a través de Discord antes de la partida.`);
    }

    if (item[17]) {
      body.appendParagraph(`<strong>Mecánicas</strong>: Hay cambios sustanciales en las mecánicas de juego.`);
    }

    if (item[18]) {
      body.appendParagraph(`<strong>Múltiple</strong>: Esta partida se juega en varias sesiones. Atento al título y a la descripción.`);
    }

    if (item[19]) {
      body.appendParagraph(`<strong>Campaña</strong>: Esta partida forma parte de una campaña. Atento al icono del título y a la descripción.`);
    }

    if (item[20]) {
      body.appendParagraph(`<strong>Grabación</strong>: La partida se grabará.`);
    }

    if (item[21]) {
      body.appendParagraph(`<strong>Emisión</strong>: La partida se emitirá.`);
    }
    
    const publishDate = getGamePublishDate(item[0])
    body.appendParagraph(
      `¡Atención! Las inscripciones para esta partida se abrirán el próximo lunes ${publishDate.getDate()} de ${MONTHS[publishDate.getMonth()]} a las 22:00 (${eventTz}).`
    );
    body.appendParagraph('');
    body.appendParagraph(`<!-- ${item[4]} (Discord: ${item[5]}) -->`);
    body.appendPageBreak();
  }
}

function onOpen(e) {
  SpreadsheetApp.getUi()
      .createMenu('Exportar partidas')
      .addItem('Google DOC', 'createDoc')
      .addToUi();
}


