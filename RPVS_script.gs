// ============================================================
// RPVS – hlavný skript
// ============================================================

// ---------- Funkcia volaná z bočnej lišty ----------
function searchByIco(ico) {
  var result = RPVS_BY_ICO(ico);
  return result;          // vráti pole polí (vrátane hlavičky)
}

// ---------- Otvorí bočnú lištu ----------
function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('RPVS – Vyhľadávanie')
      .setWidth(320);
  SpreadsheetApp.getUi().showSidebar(html);
}

// ---------- Pridá menu po otvorení súboru ----------
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('RPVS')
      .addItem('Otvoriť vyhľadávanie', 'showSidebar')
      .addToUi();
}

// ---------- Export výsledkov do CSV ----------
// Zostaví CSV na serveri a vráti ho ako Base64 string + názov súboru.
// Stiahnutie prebieha priamo v prehliadači (sidebar) cez data: URI – bez Drive.
function exportToCsv(ico, rows) {
  if (!rows || rows.length === 0) {
    return { error: 'Žiadne dáta na export.' };
  }

  var csvContent = rows.map(function(row) {
    return row.map(function(cell) {
      var val = (cell === null || cell === undefined) ? '' : cell.toString();
      if (val.indexOf(',') !== -1 || val.indexOf('"') !== -1 || val.indexOf('\n') !== -1) {
        val = '"' + val.replace(/"/g, '""') + '"';
      }
      return val;
    }).join(',');
  }).join('\n');

  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd');
  var fileName = 'RPVS_' + ico.toString().trim() + '_' + today + '.csv';

  // BOM + obsah → Base64 (sidebar ho dekóduje a stiahne)
  var blob = Utilities.newBlob('\uFEFF' + csvContent, 'text/csv', fileName);
  var base64 = Utilities.base64Encode(blob.getBytes());

  return { fileName: fileName, base64: base64 };
}

// ---------- Zapíše výsledky do aktívneho sheetu ----------
function writeResultsToSheet(rows) {
  if (!rows || rows.length === 0) return;
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.clearContents();
  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  // Formátovanie hlavičky
  sheet.getRange(1, 1, 1, rows[0].length)
       .setBackground('#1a73e8')
       .setFontColor('#ffffff')
       .setFontWeight('bold');
  sheet.autoResizeColumns(1, rows[0].length);
}

// ============================================================
// RPVS_BY_ICO – pôvodná funkcia (nezmenená)
// ============================================================

function RPVS_BY_ICO(ico) {

  if (!ico) {
    return [["Chýba IČO"]];
  }

  ico = ico.toString().trim();

  var url = "https://rpvs.gov.sk/rpvs/Partner/Partner/VyhladavaniePodlaFyzickejOsobyData";

  var payload = {
    "draw": "1",
    "start": "0",
    "length": "100",

    "columns[0][data]": "",
    "columns[0][name]": "",
    "columns[0][searchable]": "true",
    "columns[0][orderable]": "true",
    "columns[0][search][value]": "",
    "columns[0][search][regex]": "false",

    "columns[1][data]": "DatumNarodeniaFyzickejOsoby",
    "columns[1][name]": "",
    "columns[1][searchable]": "true",
    "columns[1][orderable]": "true",
    "columns[1][search][value]": "",
    "columns[1][search][regex]": "false",

    "columns[2][data]": "Adresa",
    "columns[2][name]": "",
    "columns[2][searchable]": "true",
    "columns[2][orderable]": "false",
    "columns[2][search][value]": "",
    "columns[2][search][regex]": "false",

    "columns[3][data]": "CisloVlozky",
    "columns[3][name]": "",
    "columns[3][searchable]": "true",
    "columns[3][orderable]": "true",
    "columns[3][search][value]": "",
    "columns[3][search][regex]": "false",

    "columns[4][data]": "MenoPartnera",
    "columns[4][name]": "",
    "columns[4][searchable]": "true",
    "columns[4][orderable]": "true",
    "columns[4][search][value]": "",
    "columns[4][search][regex]": "false",

    "columns[5][data]": "IcoPartnera",
    "columns[5][name]": "",
    "columns[5][searchable]": "true",
    "columns[5][orderable]": "true",
    "columns[5][search][value]": "",
    "columns[5][search][regex]": "false",

    "order[0][column]": "0",
    "order[0][dir]": "asc",

    "search[value]": "",
    "search[regex]": "false",

    "filter[MenoFyzickejOsoby]": "",
    "filter[MenoPartnera]": "",
    "filter[DatumNarodeniaFyzickejOsoby]": "",
    "filter[IcoPartnera]": ico,
    "filter[CisloVlozky]": ""
  };

  var options = {
    "method": "post",
    "payload": payload,
    "headers": {
      "X-Requested-With": "XMLHttpRequest",
      "Referer": "https://rpvs.gov.sk/rpvs/Partner/Partner/VyhladavaniePartnera"
    },
    "muteHttpExceptions": false
  };

  var response = UrlFetchApp.fetch(url, options);
  var json = JSON.parse(response.getContentText());

  if (!json.data || json.data.length === 0) {
    return [["Nenájdené"]];
  }

  var output = [];

  // Hlavička
  output.push([
    "PartnerId",
    "CisloVlozky",
    "IcoPartnera",
    "MenoPartnera",
    "MenoFyzickejOsoby",
    "DatumNarodenia",
    "TypOsoby",
    "Platny",
    "PlatnyOd",
    "PlatnyDo",
    "Adresa"
  ]);

  json.data.forEach(function(row) {
    output.push([
      row.PartnerId || "",
      row.CisloVlozky || "",
      (row.IcoPartnera || "").trim(),
      (row.MenoPartnera || "").trim(),
      (row.MenoFyzickejOsoby || "").trim(),
      formatDate(row.DatumNarodeniaFyzickejOsoby),
      row.TypOsoby || "",
      row.Platny === true ? "Áno" : "Nie",
      formatDate(row.PlatnyOd),
      formatDate(row.PlatnyDo),
      row.Adresa || ""
    ]);
  });

  return output;
}

function formatDate(dateString) {
  if (!dateString) return "";
  var date = new Date(dateString);
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "dd.MM.yyyy");
}
