/**
 * Haalt BOD contacten op uit de Configuratie sheet.
 */
function getBODContacts(configSheet) {
  var data = configSheet.getDataRange().getValues();
  var contacts = {};
  var startRow = -1;
  
  for (var i = 0; i < data.length; i++) {
    var key = data[i][0] ? data[i][0].toString().trim() : '';
    if (key === 'Contactpersonen BOD' || key === 'Contactpersoon BOD') {
      startRow = i + 1; // start from the next row
      break;
    }
  }
  
  if (startRow !== -1) {
    for (var j = startRow; j < data.length; j++) {
      var afk = data[j][0] ? data[j][0].toString().trim() : '';
      if (afk) {
        contacts[afk] = {
          naam: data[j][1] ? data[j][1].toString().trim() : '',
          email: data[j][2] ? data[j][2].toString().trim() : '',
          mobiel: data[j][3] ? data[j][3].toString().trim() : ''
        };
      }
    }
  }
  return contacts;
}

/**
 * Helper functie om configuratie als key-value paar uit te lezen
 */
function getConfiguratie(ss) {
  var configSheet = ss.getSheetByName('Configuratie');
  if (!configSheet) return null;
  
  var configData = configSheet.getDataRange().getValues();
  var config = {};
  for (var i = 0; i < configData.length; i++) {
    var key = configData[i][0] ? configData[i][0].toString().trim() : '';
    var value = configData[i][1] ? configData[i][1].toString().trim() : '';
    // We stoppen met het inlezen van de platte config als we de contacten-tabel bereiken
    if (key === 'Contactpersonen BOD' || key === 'Contactpersoon BOD') break;
    
    if (key) {
      config[key] = value;
    }
  }
  return {
    config: config,
    sheet: configSheet
  };
}

/**
 * Hoofdfunctie om het accreditatieproces te draaien voor alle rijen.
 */
function runAccreditatieProces(optSs) {
  _processAccreditatie(false, optSs);
}

/**
 * Functie om het accreditatieproces te draaien voor alleen de geselecteerde rij.
 */
function runAccreditatieProcesGeselecteerd(optSs) {
  _processAccreditatie(true, optSs);
}

/**
 * Interne processor voor het genereren van spreadsheets.
 */
function _processAccreditatie(alleenGeselecteerd, optSs) {
  var ui;
  try {
    ui = SpreadsheetApp.getUi();
  } catch (e) {
    ui = null;
  }
  
  var ss = optSs || SpreadsheetApp.getActiveSpreadsheet();
  var configObj = getConfiguratie(ss);
  
  if (!configObj) {
    if (ui) ui.alert('Fout: Tabblad "Configuratie" niet gevonden.');
    Logger.log('Fout: Tabblad "Configuratie" niet gevonden.');
    return;
  }
  
  var config = configObj.config;
  
  var templateId = config['Template ID'];
  var folderId = config['Doelmap ID'];
  var nameTemplate = config['Bestandsnaam'];
  var outputKolomNaam = config['Output Kolom'];
  var tabbladenString = config['Tabbladen'];
  
  if (!templateId || !folderId || !nameTemplate || !outputKolomNaam || !tabbladenString) {
    if (ui) ui.alert('Fout: Ontbrekende configuratie. Controleer of Template ID, Doelmap ID, Bestandsnaam, Output Kolom en Tabbladen zijn ingevuld.');
    return;
  }
  
  var targetFolder;
  try {
    targetFolder = DriveApp.getFolderById(folderId);
  } catch (e) {
    if (ui) ui.alert('Fout: Doelmap ID is onjuist of niet toegankelijk.');
    return;
  }

  var templateFile;
  try {
    templateFile = DriveApp.getFileById(templateId);
  } catch (e) {
    if (ui) ui.alert('Fout: Template ID is onjuist of niet toegankelijk.');
    return;
  }
  
  var tabbladen = tabbladenString.split(',').map(function(s) { return s.trim(); });
  var totalProcessed = 0;
  
  // Als alleenGeselecteerd waar is, bepalen we welke rij geselecteerd is
  var activeSheet = ss.getActiveSheet();
  var activeRowIndex = alleenGeselecteerd ? activeSheet.getActiveCell().getRow() : -1;
  var activeSheetName = alleenGeselecteerd ? activeSheet.getName() : null;
  
  // 2. Loop door de dynamische tabbladen
  for (var t = 0; t < tabbladen.length; t++) {
    var sheetName = tabbladen[t];
    var sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) continue;
    if (alleenGeselecteerd && sheetName !== activeSheetName) continue;
    
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow < 2 || lastCol < 1) continue;
    
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) { return h.toString().trim(); });
    
    // 3. Slimme Output Kolom zoeken / toevoegen
    var outputColIndex = headers.indexOf(outputKolomNaam);
    if (outputColIndex === -1) {
      outputColIndex = headers.length;
      sheet.getRange(1, outputColIndex + 1).setValue(outputKolomNaam);
      headers.push(outputKolomNaam);
    }
    
    var bedrijfsnaamIndex = headers.indexOf('Bedrijfsnaam');
    var dataRange = sheet.getRange(2, 1, lastRow - 1, headers.length);
    var data = dataRange.getDisplayValues();
    var formulas = dataRange.getFormulas();
    
    // 4. Loop door rijen
    for (var r = 0; r < data.length; r++) {
      var rowNum = r + 2;
      
      if (alleenGeselecteerd && rowNum !== activeRowIndex) continue;
      
      var rowData = data[r];
      var formulaOutput = formulas[r][outputColIndex] || '';
      var stringOutput = rowData[outputColIndex] ? rowData[outputColIndex].toString() : '';
      
      // Controleer of er al een link is
      if (formulaOutput.toUpperCase().indexOf('HYPERLINK(') !== -1 || stringOutput.indexOf('http') === 0) {
        continue;
      }
      
      var fileName = nameTemplate;
      
      // Bepaal de dynamische bestandsnaam door placeholders (headers) te vervangen
      for (var c = 0; c < headers.length; c++) {
        var headerName = headers[c];
        if (headerName) {
          var placeholder = '\\{\\{' + headerName + '\\}\\}';
          var regex = new RegExp(placeholder, 'gi');
          fileName = fileName.replace(regex, rowData[c]);
        }
      }
      
      // Kopieer template spreadsheet naar de doelmap
      var newFile = templateFile.makeCopy(fileName, targetFolder);
      var newSsId = newFile.getId();
      var newSs = SpreadsheetApp.openById(newSsId);
      
      // 5. Dynamische Kolommen (Merge Tags)
      var newSheets = newSs.getSheets();
      
      for (var s = 0; s < newSheets.length; s++) {
        var currentSheet = newSheets[s];
        
        // Optioneel snelle controle of het een lege sheet is:
        if (currentSheet.getLastRow() === 0 && currentSheet.getLastColumn() === 0) continue;
        
        var targetRange = currentSheet.getDataRange();
        
        for (var c2 = 0; c2 < headers.length; c2++) {
          var hName = headers[c2];
          if (hName) {
            var tag = '{{' + hName + '}}';
            var value = rowData[c2] !== undefined && rowData[c2] !== null ? rowData[c2].toString() : '';
            targetRange.createTextFinder(tag).replaceAllWith(value);
          }
        }
      }
      
      // 6. Link terugschrijven naar de output kolom met dynamische naam
      var docUrl = newFile.getUrl();
      var bNaam = bedrijfsnaamIndex !== -1 ? rowData[bedrijfsnaamIndex] : '';
      var linkTekst = bNaam ? 'Accreditatie ' + bNaam : 'Accreditatie';
      
      sheet.getRange(rowNum, outputColIndex + 1).setFormula('=HYPERLINK("' + docUrl + '"; "' + linkTekst + '")');
      SpreadsheetApp.flush();
      totalProcessed++;
    }
  }
  
  if (ui) {
    if (totalProcessed > 0) {
      ui.alert('Succes! ' + totalProcessed + ' document(en) gegenereerd en gekoppeld.');
    } else {
      if (!alleenGeselecteerd) {
        ui.alert('Klaar. Er waren geen te verwerken rijen zonder document gevonden.');
      } else {
        ui.alert('Klaar. De geselecteerde rij heeft al een document of is leeg.');
      }
    }
  }
}

/**
 * Hoofdfunctie om mails te sturen voor alle rijen.
 */
function mailAccreditatieProcesAlle(optSs) {
  mailAccreditatieProces(false, optSs);
}

/**
 * Functie om de mail te sturen voor de geselecteerde rij.
 */
function mailAccreditatieProcesGeselecteerd(optSs) {
  mailAccreditatieProces(true, optSs);
}

/**
 * Interne processor voor het verzenden van e-mails.
 */
function mailAccreditatieProces(alleenGeselecteerd, optSs) {
  var ui;
  try {
    ui = SpreadsheetApp.getUi();
  } catch (e) {
    ui = null;
  }
  
  var ss = optSs || SpreadsheetApp.getActiveSpreadsheet();
  var configObj = getConfiguratie(ss);
  
  if (!configObj) {
    if (ui) ui.alert('Fout: Tabblad "Configuratie" niet gevonden.');
    return;
  }
  
  var config = configObj.config;
  var configSheet = configObj.sheet;
  
  var formatEmail = config['Format Email'];
  var outputKolomSend = config['Output Kolom Send'];
  var outputKolomNaam = config['Output Kolom']; // Bevat de gegenereerde link
  var tabbladenString = config['Tabbladen'];
  
  if (!formatEmail || !outputKolomSend || !outputKolomNaam || !tabbladenString) {
    if (ui) ui.alert('Fout: Zorg dat Format Email, Output Kolom Send, Output Kolom en Tabbladen zijn ingevuld.');
    return;
  }
  
  var bodContacts = getBODContacts(configSheet);
  var tabbladen = tabbladenString.split(',').map(function(s) { return s.trim(); });
  var totalProcessed = 0;
  
  var activeSheet = ss.getActiveSheet();
  var activeRowIndex = alleenGeselecteerd ? activeSheet.getActiveCell().getRow() : -1;
  var activeSheetName = alleenGeselecteerd ? activeSheet.getName() : null;
  
  for (var t = 0; t < tabbladen.length; t++) {
    var sheetName = tabbladen[t];
    var sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) continue;
    if (alleenGeselecteerd && sheetName !== activeSheetName) continue;
    
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow < 2 || lastCol < 1) continue;
    
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) { return h.toString().trim(); });
    
    var outputColIndex = headers.indexOf(outputKolomNaam);
    var emailIndex = headers.indexOf('Email');
    var bNaamIndex = headers.indexOf('Bedrijfsnaam');
    var contactIndex = headers.indexOf('Contactpersoon');
    var bodContactIndex = headers.indexOf('Contactpersoon BOD');
    
    if (outputColIndex === -1 || emailIndex === -1) {
      Logger.log('Waarschuwing: Tabblad mist de kolom Email of de Accreditatie output kolom.');
      continue;
    }
    
    var sendColIndex = headers.indexOf(outputKolomSend);
    // Voeg sendColIndex toe als deze niet bestaat
    if (sendColIndex === -1) {
      sendColIndex = headers.length;
      sheet.getRange(1, sendColIndex + 1).setValue(outputKolomSend);
      headers.push(outputKolomSend);
    }
    
    var dataRange = sheet.getRange(2, 1, lastRow - 1, headers.length);
    var data = dataRange.getDisplayValues();
    var formulas = dataRange.getFormulas();
    
    for (var r = 0; r < data.length; r++) {
      var rowNum = r + 2;
      if (alleenGeselecteerd && rowNum !== activeRowIndex) continue;
      
      var rowData = data[r];
      var formulaLink = formulas[r][outputColIndex] || '';
      var stringLink = rowData[outputColIndex] || '';
      var sentStatus = rowData[sendColIndex] ? rowData[sendColIndex].toString().trim() : '';
      var targetEmail = rowData[emailIndex] ? rowData[emailIndex].toString().trim() : '';
      
      // Extract URL uit HYPERLINK of text
      var docUrl = '';
      var match = formulaLink.match(/HYPERLINK\("([^"]+)"/i);
      if (match) {
        docUrl = match[1];
      } else if (stringLink.indexOf('http') === 0) {
        docUrl = stringLink;
      }
      
      // Controleer condities: Moet een URL hebben, moet onverzonden zijn, en moet een e-mail hebben
      if (!docUrl || sentStatus !== '' || !targetEmail) {
        continue;
      }
      
      var bNaam = bNaamIndex !== -1 ? rowData[bNaamIndex] : '';
      var cPersoon = contactIndex !== -1 ? rowData[contactIndex] : '';
      var bodAfkorting = bodContactIndex !== -1 ? rowData[bodContactIndex] : '';
      
      var bodDetails = bodContacts[bodAfkorting] || { naam: '', email: '', mobiel: '' };
      
      // Bouw E-mail content
      var bodyHtml = formatEmail;
      
      // Behoud enters uit de cel door \n te vervangen door <br>
      bodyHtml = bodyHtml.replace(/\n/g, '<br>');
      
      // Replace basis tags (case-insensitive in geval van typfouten)
      bodyHtml = bodyHtml.replace(new RegExp('\\{\\{Bedrijfsnaam\\}\\}', 'gi'), bNaam);
      bodyHtml = bodyHtml.replace(new RegExp('\\{\\{Contactpersoon\\}\\}', 'gi'), cPersoon);
      bodyHtml = bodyHtml.replace(new RegExp('\\{\\{Contactpersoon BOD\\}\\}', 'gi'), bodDetails.naam);
      bodyHtml = bodyHtml.replace(new RegExp('\\{\\{Email Contactpersoon BOD\\}\\}', 'gi'), bodDetails.email);
      bodyHtml = bodyHtml.replace(new RegExp('\\{\\{Mobiel Contactpersoon BOD\\}\\}', 'gi'), bodDetails.mobiel);
      
      // Voeg twee knoppen toe: één voor desktop (Sheet) en één voor mobiel (Web App)
      var webAppUrl = config['Web App URL'] || config['Web App url'] || config['web app url'];
      var primaryColor = config['Kleur Primair'] || config['kleur primair'] || '#1a73e8';
      var ssId = '';
      var ssIdMatch = docUrl.match(/\/d\/([a-zA-Z0-9-_]+)/);
      if (ssIdMatch) {
        ssId = ssIdMatch[1];
      }
      
      var buttonHtml = '<br><br>';
      
      // Button 1: Desktop (Direct link to Google Sheet)
      buttonHtml += '<a href="' + docUrl + '" style="' +
        'background-color: ' + primaryColor + '; ' +
        'color: white; ' +
        'padding: 12px 24px; ' +
        'text-decoration: none; ' +
        'border-radius: 4px; ' +
        'display: inline-block; ' +
        'margin-bottom: 10px; ' +
        'font-family: Arial, sans-serif; ' +
        'font-weight: bold;">' +
        'Beste voor laptop: Bekijk Accreditatie ' + bNaam + '</a><br>';
        
      // Button 2: Mobile (Link to Web App)
      if (webAppUrl && bNaam) {
        var mobileUrl = webAppUrl + '?bedrijf=' + encodeURIComponent(bNaam);
        buttonHtml += '<br><a href="' + mobileUrl + '" style="' +
          'background-color: ' + primaryColor + '; ' +
          'color: white; ' +
          'padding: 12px 24px; ' +
          'text-decoration: none; ' +
          'border-radius: 4px; ' +
          'display: inline-block; ' +
          'font-family: Arial, sans-serif; ' +
          'font-weight: bold;">' +
          'Beste voor mobiel: Bekijk Accreditatie ' + bNaam + ' als webapp</a><br><br>';
      } else {
        buttonHtml += '<br>';
      }
        
      bodyHtml += buttonHtml;
      
      var subject = 'Accreditatie ' + bNaam;
      
      try {
        var mailOptions = {
          htmlBody: bodyHtml
        };
        
        var afzenderNaam = config['Afzender Naam'];
        if (afzenderNaam) {
          mailOptions.name = afzenderNaam;
        }
        
        if (bodDetails.email) {
          mailOptions.cc = bodDetails.email;
        }
        
        // GmailApp zorgt ervoor dat het ook in Verzonden Items komt in de Google Workspace
        GmailApp.sendEmail(targetEmail, subject, 'Bekijk de accreditatie hier: ' + docUrl, mailOptions);
        
        // Noteer verzendmoment
        var now = new Date();
        var timestamp = Utilities.formatDate(now, ss.getSpreadsheetTimeZone(), "dd-MM HH:mm");
        sheet.getRange(rowNum, sendColIndex + 1).setValue(timestamp);
        SpreadsheetApp.flush();
        totalProcessed++;
        
      } catch (err) {
        Logger.log('Fout bij verzenden email naar ' + targetEmail + ': ' + err.message);
      }
    }
  }
  
  if (ui) {
    if (totalProcessed > 0) {
      ui.alert('Succes! ' + totalProcessed + ' e-mail(s) verstuurd.');
    } else {
      if (!alleenGeselecteerd) {
        ui.alert('Klaar. Er waren geen te verzenden e-mails gevonden (link ontbreekt of mail was al verstuurd).');
      } else {
        ui.alert('Klaar. Voor de geselecteerde rij ontbreekt de link of is de mail al verstuurd.');
      }
    }
  }
}

/**
 * Web App Entry Point
 */
function doGet(e) {
  var bedrijfsnaam = e.parameter.bedrijf;
  if (!bedrijfsnaam) {
    return HtmlService.createHtmlOutput('<h3 style="color:red;font-family:sans-serif;text-align:center;margin-top:50px;">Fout: Geen bedrijfsnaam opgegeven in URL.</h3>');
  }
  
  var activeSs = SpreadsheetApp.getActiveSpreadsheet();
  if (!activeSs) {
    return HtmlService.createHtmlOutput('<h3 style="color:red;font-family:sans-serif;text-align:center;margin-top:50px;">Fout: Kan hoofdspreadsheet niet openen. Zorg dat de Web App gekoppeld is aan de spreadsheet.</h3>');
  }
  
  var accreditatieSheet = activeSs.getSheetByName('Accreditatie');
  if (!accreditatieSheet) {
    return HtmlService.createHtmlOutput('<h3 style="color:red;font-family:sans-serif;text-align:center;margin-top:50px;">Fout: Tabblad "Accreditatie" niet gevonden.</h3>');
  }
  
  var configObj = getConfiguratie(activeSs);
  if (!configObj) {
    return HtmlService.createHtmlOutput('<h3 style="color:red;font-family:sans-serif;text-align:center;margin-top:50px;">Fout: Tabblad "Configuratie" niet gevonden.</h3>');
  }
  
  var outputKolomNaam = configObj.config['Output Kolom'];
  var lastRow = accreditatieSheet.getLastRow();
  var lastCol = accreditatieSheet.getLastColumn();
  
  if (lastRow < 2 || lastCol < 1 || !outputKolomNaam) {
    return HtmlService.createHtmlOutput('<h3 style="color:red;font-family:sans-serif;text-align:center;margin-top:50px;">Fout: Accreditatie tabblad is leeg of configuratie mist Output Kolom.</h3>');
  }
  
  var headers = accreditatieSheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) { return h.toString().trim(); });
  var bNaamIndex = headers.indexOf('Bedrijfsnaam');
  var outputColIndex = headers.indexOf(outputKolomNaam);
  
  if (bNaamIndex === -1 || outputColIndex === -1) {
    return HtmlService.createHtmlOutput('<h3 style="color:red;font-family:sans-serif;text-align:center;margin-top:50px;">Fout: Kolom Bedrijfsnaam of Output Kolom niet gevonden in Accreditatie tabblad.</h3>');
  }
  
  var dataRange = accreditatieSheet.getRange(2, 1, lastRow - 1, headers.length);
  var data = dataRange.getDisplayValues();
  var formulas = dataRange.getFormulas();
  
  var targetUrl = '';
  for (var r = 0; r < data.length; r++) {
    if (data[r][bNaamIndex] == bedrijfsnaam) {
      var formulaLink = formulas[r][outputColIndex] || '';
      var stringLink = data[r][outputColIndex] || '';
      
      var match = formulaLink.match(/HYPERLINK\("([^"]+)"/i);
      if (match) {
        targetUrl = match[1];
      } else if (stringLink.indexOf('http') === 0) {
        targetUrl = stringLink;
      }
      break;
    }
  }
  
  if (!targetUrl) {
    return HtmlService.createHtmlOutput('<h3 style="color:red;font-family:sans-serif;text-align:center;margin-top:50px;">Fout: Geen accreditatiedocument gevonden voor bedrijf: ' + bedrijfsnaam + '</h3>');
  }
  
  var idMatch = targetUrl.match(/\/d\/([a-zA-Z0-9-_]+)/);
  if (!idMatch) {
    return HtmlService.createHtmlOutput('<h3 style="color:red;font-family:sans-serif;text-align:center;margin-top:50px;">Fout: Ongeldige document URL gevonden.</h3>');
  }
  
  var id = idMatch[1];
  
  var template = HtmlService.createTemplateFromFile('WebApp');
  template.sheetId = id;
  
  // Standaardwaarden voor fallback
  template.eventName = 'BOD Accreditatie';
  template.colorPrimary = '#1a73e8';
  template.colorSecondary = '#6c757d';
  template.colorBackground = '#f8f9fa';
  template.ssUrl = 'https://docs.google.com/spreadsheets/d/' + id;
  
  try {
    var ss = SpreadsheetApp.openById(id);
    template.ssUrl = ss.getUrl();
    
    // Gebruik de bestaande getConfiguratie functie
    var configObj = getConfiguratie(ss);
    if (configObj) {
      var config = configObj.config;
      if (config['Naam Evenement']) template.eventName = config['Naam Evenement'];
      if (config['Kleur Primair']) template.colorPrimary = config['Kleur Primair'];
      if (config['Kleur Secundair']) template.colorSecondary = config['Kleur Secundair'];
      if (config['Kleur Achtergrond']) template.colorBackground = config['Kleur Achtergrond'];
    }
    
    var sheet = ss.getSheets()[0];
    var lastCol = sheet.getLastColumn();
    if (lastCol > 0) {
      var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) { return h.toString().trim(); });
      template.headers = headers;
    } else {
      template.headers = [];
    }
  } catch (err) {
    return HtmlService.createHtmlOutput('<div style="font-family:sans-serif;text-align:center;margin-top:50px;">' +
      '<h2 style="color:#d32f2f;">Toegang Geweigerd of Fout</h2>' +
      '<p>Kan de bijbehorende accreditatie spreadsheet niet openen.</p>' +
      '<p style="color:#666;font-size:0.9em;">(Details: ' + err.message + ')</p>' +
      '</div>');
  }
  
  try {
    return template.evaluate()
        .setTitle('Accreditatie ' + template.eventName)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } catch (err) {
    return HtmlService.createHtmlOutput('<div style="font-family:sans-serif;text-align:center;margin-top:50px;">' +
      '<h2 style="color:#d32f2f;">Fout in Weergave</h2>' +
      '<p>Er is een probleem opgetreden bij het laden van het mobiele formulier.</p>' +
      '<p style="color:#666;font-size:0.9em;">(Details: ' + err.message + ')</p>' +
      '</div>');
  }
}

/**
 * Backend functie aangeroepen vanuit Web App
 */
function verwerkFormulier(sheetId, formDataArray) {
  var ss = SpreadsheetApp.openById(sheetId);
  var sheet = ss.getSheets()[0];
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function(h) { return h.toString().trim(); });
  
  for (var i = 0; i < formDataArray.length; i++) {
    var rowData = [];
    var personData = formDataArray[i];
    
    for (var j = 0; j < headers.length; j++) {
      var hName = headers[j];
      rowData.push(personData[hName] || '');
    }
    
    sheet.appendRow(rowData);
  }
  
  return true;
}
