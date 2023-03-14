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

function getlangPlural(langs) {
  return langs.split(',').length > 1 ? 's' : '';
}

function formatDate(dateObj) {
  if (typeof dateObj === 'string') return dateObj;
  return Utilities.formatDate(dateObj, "GMT+1", "dd/MM/yyyy");
}

function formatTime(dateObj) {
  if (typeof dateObj === 'string') return dateObj.replaceAll(' ','');
  return dateObj.toLocaleTimeString('en',
    {timeZone:'Europe/Madrid',hour12:true,hour:'numeric',minute:'numeric'}
  );
}

function cestOrCet(dateObj) {
  /*
  if (typeof dateObj === 'string') return '--';
  const correctDateStr = Utilities.formatDate(dateObj, "GMT+1", "dd-MM-yyyy");
  const correctDate = new Date(correctDateStr)
  console.log(correctDate.getTimezoneOffset())
  return dateObj.getTimezoneOffset() === -60 ? 'CET' : 'CEST';
  */
  return 'Horario de España';
}

function getAdditionaInfo(item) {
  const ADDITIONAL_INFO = [
    {
      index: 14,
      label: 'Experto',
      message: 'Se requieren conocimientos previos de sistema y ambientación.'
    },
    {
      index: 15,
      label: 'Trasfondo',
      message: 'Contacta con el organizador a través de Discord antes de la partida.'
    },
    {
      index: 16,
      label: 'Mecánicas',
      message: 'Hay cambios sustanciales en las mecánicas de juego.'
    },
    {
      index: 17,
      label: 'Múltiple',
      message: 'Esta partida se juega en varias sesiones. Atento al título y a la descripción.'
    },
    {
      index: 18,
      label: 'Campaña',
      message: 'Esta partida forma parte de una campaña. Atento al icono del título y a la descripción.'
    },
    {
      index: 19,
      label: 'Grabación',
      message: 'La partida se grabará.'
    },
    {
      index: 20,
      label: 'Emisión',
      message: 'La partida se emitirá.'
    }
  ];

  const result = [];
  ADDITIONAL_INFO.forEach(info => {
    if (item[info.index]) {
      result.push(info.message);
    }
  })
  return result;
}

function onOpen(e) {
  SpreadsheetApp.getUi()
      .createMenu('Exportar partidas')
      .addItem('Google DOC', 'createDoc')
      .addToUi();
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

  // Remove red warning
  if (content[content.length - 1].join('').includes('loguitos')) {
    content.pop();
  }
  
  const columns = content[0].length;
  const header = [...content[0]];
  content.shift();
  const docTitle = `RL_${getNextDate()}`;

  const doc = DocumentApp.create(docTitle);
  var body = doc.getBody();
  
  for (let i=0; i<content.length; i++) {
    const item = content[i];
    const title = `${item[6]} (${item[8]})`;
    const gameDate = `<strong>Día</strong>: ${item[1]} ${formatDate(item[0])} de ${formatTime(item[2])} a ${formatTime(item[3])} (${cestOrCet(item[0])})`;
    
    if (i === 0) {
      body.getChild(0).asParagraph().setText(title);
      body.getChild(0).setBold(true);
    } else {
      body.appendParagraph(title).setBold(true);
    }

    const aditionaInfo = getAdditionaInfo(item);

    body.appendParagraph('').setBold(false);
    body.appendParagraph(`<strong>Sinopsis</strong>: ${item[13]}`);

    // Consideraciones adicionales
    if (aditionaInfo.length) {
      body.appendParagraph('');
      body.appendParagraph(`<strong>Importante</strong>:`);
    }
    aditionaInfo.forEach(info => {
      body.appendParagraph(info);  
    });
    
    body.appendParagraph('');
    body.appendParagraph(`<strong>Ambientación</strong>: ${item[7]}`);

    body.appendParagraph('');
    body.appendParagraph(`<strong>Sistema de juego</strong>: ${item[8]}`);

    body.appendParagraph('');
    body.appendParagraph(`<strong>Jugadores</strong>: mínimo ${item[9]}, máximo ${item[10]}`);

    body.appendParagraph('');
    body.appendParagraph(`<strong>Idioma${getlangPlural(item[11])}</strong>: ${item[11]}`);

    body.appendParagraph('');
    body.appendParagraph(`<strong>Aviso de contenido</strong>: ${item[12]}`);

    body.appendParagraph('');
    body.appendParagraph(gameDate);

    body.appendParagraph('');
    body.appendParagraph(`${item[4]} (Discord: ${item[5]})`);

    body.appendPageBreak();
    
  }
}



