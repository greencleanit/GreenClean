const APP_NAME = 'Green Clean - Anagrafiche Clienti';
const DB_SHEET_NAME = 'CLIENTI';
const DB_PROPERTY_KEY = 'GREEN_CLEAN_CLIENTI_DB_ID';

function doGet() {
  return HtmlService
    .createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Green Clean - Anagrafica Cliente')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getOrCreateDatabase_() {
  const props = PropertiesService.getScriptProperties();
  let spreadsheetId = props.getProperty(DB_PROPERTY_KEY);

  let ss;

  if (spreadsheetId) {
    try {
      ss = SpreadsheetApp.openById(spreadsheetId);
    } catch (e) {
      spreadsheetId = null;
    }
  }

  if (!spreadsheetId) {
    ss = SpreadsheetApp.create(APP_NAME + ' - Database');
    props.setProperty(DB_PROPERTY_KEY, ss.getId());
  }

  let sheet = ss.getSheetByName(DB_SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(DB_SHEET_NAME);
  }

  setupSheet_(sheet);

  return sheet;
}

function setupSheet_(sheet) {
  const headers = [
    'Timestamp',
    'Ragione Sociale',
    'Forma Giuridica',
    'Partita IVA',
    'Codice Fiscale',
    'Codice SDI',
    'PEC',
    'Sito Web',

    'Sede Legale - Indirizzo',
    'Sede Legale - CAP',
    'Sede Legale - Comune',
    'Sede Legale - Provincia',
    'Sede Legale - Nazione',

    'Sede Operativa Coincide',
    'Sede Operativa - Indirizzo',
    'Sede Operativa - CAP',
    'Sede Operativa - Comune',
    'Sede Operativa - Provincia',
    'Sede Operativa - Nazione',

    'Acquisti - Nome',
    'Acquisti - Telefono',
    'Acquisti - Cellulare',
    'Acquisti - Email',

    'Amministrazione - Nome',
    'Amministrazione - Telefono',
    'Amministrazione - Cellulare',
    'Amministrazione - Email',

    'Magazzino - Nome',
    'Magazzino - Telefono',
    'Magazzino - Cellulare',
    'Magazzino - Email',

    'Note Operative'
  ];

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#12372A')
      .setFontColor('#FFFFFF');
    sheet.autoResizeColumns(1, headers.length);
  }
}

function saveCustomerData(data) {
  try {
    validateData_(data);

    const sheet = getOrCreateDatabase_();

    const row = [
      new Date(),

      data.ragioneSociale || '',
      data.formaGiuridica || '',
      data.partitaIva || '',
      data.codiceFiscale || '',
      data.codiceSdi || '',
      data.pec || '',
      data.sitoWeb || '',

      data.legaleIndirizzo || '',
      data.legaleCap || '',
      data.legaleComune || '',
      data.legaleProvincia || '',
      data.legaleNazione || '',

      data.sedeCoincide ? 'Sì' : 'No',
      data.operativaIndirizzo || '',
      data.operativaCap || '',
      data.operativaComune || '',
      data.operativaProvincia || '',
      data.operativaNazione || '',

      data.acquistiNome || '',
      data.acquistiTelefono || '',
      data.acquistiCellulare || '',
      data.acquistiEmail || '',

      data.amministrazioneNome || '',
      data.amministrazioneTelefono || '',
      data.amministrazioneCellulare || '',
      data.amministrazioneEmail || '',

      data.magazzinoNome || '',
      data.magazzinoTelefono || '',
      data.magazzinoCellulare || '',
      data.magazzinoEmail || '',

      data.noteOperative || ''
    ];

    sheet.appendRow(row);
    sendNotificationEmail_(data);

    return {
      success: true,
      message: "Grazie. L'anagrafica è stata acquisita correttamente."
    };

  } catch (err) {
    return {
      success: false,
      message: err.message || 'Errore durante il salvataggio.'
    };
  }
}

function validateData_(data) {
  if (!data) throw new Error('Dati non ricevuti.');

  if (!data.ragioneSociale) {
    throw new Error('La Ragione Sociale è obbligatoria.');
  }

  if (!data.partitaIva) {
    throw new Error('La Partita IVA è obbligatoria.');
  }

  if (!data.codiceSdi && !data.pec) {
    throw new Error('È obbligatorio inserire almeno Codice SDI o PEC.');
  }
}

function inizializzaDatabase() {
  getOrCreateDatabase_();
}

function sendNotificationEmail_(data) {
  const destinatario = 'a.pelosi@greencleanitalia.it';

  const oggetto = 'Nuova anagrafica cliente - ' + data.ragioneSociale;

  const corpo =
    'È stata inserita una nuova anagrafica cliente.\n\n' +
    'Ragione Sociale: ' + (data.ragioneSociale || '') + '\n' +
    'Partita IVA: ' + (data.partitaIva || '') + '\n' +
    'Codice SDI: ' + (data.codiceSdi || '') + '\n' +
    'PEC: ' + (data.pec || '') + '\n' +
    'Comune sede legale: ' + (data.legaleComune || '') + '\n\n' +
    'Note operative:\n' + (data.noteOperative || '');

  MailApp.sendEmail(destinatario, oggetto, corpo);
}
