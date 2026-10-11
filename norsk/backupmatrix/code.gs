/**
 * Hovedfunksjon som tar imot forespørsler fra spillet.
 * Denne versjonen samler alle rader som har samme kode i kolonne A.
 */
function doGet(e) {
  if (!e || !e.parameter) {
    return ContentService.createTextOutput("Scriptet må kalles fra spillet.")
      .setMimeType(ContentService.MimeType.TEXT);
  }

  var action = e.parameter.action;
  var code = e.parameter.code;
  var callback = e.parameter.callback;
  var result = {};
  
  try {
    if (action === 'getList') {
      result = getListByCode(code);
    }
  } catch (err) {
    result = { error: err.toString() };
  }
  
  var jsonString = JSON.stringify(result);
  
  if (callback) {
    return ContentService.createTextOutput(callback + "(" + jsonString + ")")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  
  return ContentService.createTextOutput(jsonString)
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Leter etter alle rader med riktig kode i kolonne A.
 */
function getListByCode(code) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0]; 
  var data = sheet.getDataRange().getValues();
  
  var searchCode = code.toUpperCase().trim();
  var formattedWords = [];
  var listName = "";

  // Vi går gjennom alle rader (hopper over overskriftsraden)
  for (var i = 1; i < data.length; i++) {
    var rowCode = data[i][0] ? data[i][0].toString().toUpperCase().trim() : "";
    
    if (rowCode === searchCode) {
      // Vi lagrer navnet fra kolonne B (bruker det fra den første raden vi finner)
      if (!listName) {
        listName = data[i][1] || "Min liste";
      }
      
      // Henter Spørsmål fra kolonne C og Svar fra kolonne D
      var question = data[i][2] ? data[i][2].toString().trim() : "";
      var answer = data[i][3] ? data[i][3].toString().trim() : "";
      
      if (question && answer) {
        formattedWords.push({
          q: question,
          a: answer
        });
      }
    }
  }
  
  if (formattedWords.length > 0) {
    return {
      name: listName,
      words: formattedWords
    };
  } else {
    return { error: "Fant ingen ord med koden " + code };
  }
}
