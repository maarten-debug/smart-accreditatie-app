/**
 * Haalt BOD contacten op uit de Contactpersonen sheet.
 */
function getBODContacts(ss) {
  var contactSheet = ss.getSheetByName('Contactpersonen');
  if (!contactSheet) return {};
  
  var data = contactSheet.getDataRange().getValues();
  var contacts = {};
  
  for (var i = 1; i < data.length; i++) {
    var afk = data[i][0] ? data[i][0].toString().trim() : '';
    if (afk) {
      contacts[afk] = {
        naam: data[i][1] ? data[i][1].toString().trim() : '',
        email: data[i][2] ? data[i][2].toString().trim() : '',
        mobiel: data[i][3] ? data[i][3].toString().trim() : ''
      };
    }
  }
  return contacts;
}

/**
 * Haalt een lijst met alle BOD emails op uit de Contactpersonen sheet.
 */
function getBODEmails(ss) {
  var contactSheet = ss.getSheetByName('Contactpersonen');
  if (!contactSheet) return [];
  
  var data = contactSheet.getDataRange().getValues();
  var emails = [];
  
  for (var i = 1; i < data.length; i++) {
    var email = data[i][2] ? data[i][2].toString().trim() : '';
    if (email && email.indexOf('@') !== -1) {
      emails.push(email);
    }
  }
  return emails;
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
      
      // Verfijn bladbeveiliging vanuit template (als aanwezig)
      var bodEmails = getBODEmails(ss);
      var startRij = parseInt(config['Beveilig template tot rij']) || 10;
      
      var newSheets = newSs.getSheets();
      for (var s = 0; s < newSheets.length; s++) {
        var currentSheet = newSheets[s];
        
        // Bladbeveiligingen
        var sheetProtections = currentSheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
        for (var p = 0; p < sheetProtections.length; p++) {
          var protection = sheetProtections[p];
          protection.removeEditors(protection.getEditors());
          if (bodEmails.length > 0) {
            protection.addEditors(bodEmails);
          }
          
          var maxRows = currentSheet.getMaxRows();
          var maxCols = currentSheet.getMaxColumns();
          if (maxRows >= startRij) {
            var unprotectedRange = currentSheet.getRange(startRij, 1, maxRows - startRij + 1, maxCols);
            protection.setUnprotectedRanges([unprotectedRange]);
          }
        }
        
        // Range beveiligingen
        var rangeProtections = currentSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
        for (var p = 0; p < rangeProtections.length; p++) {
          var protection = rangeProtections[p];
          protection.removeEditors(protection.getEditors());
          if (bodEmails.length > 0) {
            protection.addEditors(bodEmails);
          }
        }
      }
      
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
  mailAccreditatieProces(false, optSs, false);
}

/**
 * Functie om de mail te sturen voor de geselecteerde rij.
 */
function mailAccreditatieProcesGeselecteerd(optSs) {
  mailAccreditatieProces(true, optSs, false);
}

/**
 * Functie om een concept mail klaar te zetten voor de geselecteerde rij.
 */
function mailAccreditatieProcesGeselecteerdConcept(optSs) {
  mailAccreditatieProces(true, optSs, true);
}

/**
 * Interne processor voor het verzenden (of klaarzetten) van e-mails.
 */
function mailAccreditatieProces(alleenGeselecteerd, optSs, isConcept) {
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
  
  var formatEmail = config['Format Email'];
  var outputKolomSend = config['Output Kolom Send'];
  var outputKolomNaam = config['Output Kolom']; // Bevat de gegenereerde link
  var tabbladenString = config['Tabbladen'];
  
  if (!formatEmail || !outputKolomSend || !outputKolomNaam || !tabbladenString) {
    if (ui) ui.alert('Fout: Zorg dat Format Email, Output Kolom Send, Output Kolom en Tabbladen zijn ingevuld.');
    return;
  }
  
  var bodContacts = getBODContacts(ss);
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
      
      // Voeg enkele knop toe die DIRECT naar de Google Sheet gaat
      var buttonHtml = '<br><br><a href="' + docUrl + '" style="' +
        'background-color: #1a73e8; ' +
        'color: white; ' +
        'padding: 12px 24px; ' +
        'text-decoration: none; ' +
        'border-radius: 4px; ' +
        'display: inline-block; ' +
        'font-family: Arial, sans-serif; ' +
        'font-weight: bold;">' +
        'Open Accreditatiesheet ' + bNaam + '</a><br><br>';
        
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
        
        if (isConcept) {
          GmailApp.createDraft(targetEmail, subject, 'Bekijk de accreditatie hier: ' + docUrl, mailOptions);
          
          var now = new Date();
          var timestamp = Utilities.formatDate(now, ss.getSpreadsheetTimeZone(), "dd-MM HH:mm");
          var statusMessage = timestamp + ' concept ' + bodAfkorting;
          sheet.getRange(rowNum, sendColIndex + 1).setValue(statusMessage);
          
          if (ss && ss.toast) {
            ss.toast('Concept aangemaakt voor ' + bNaam, 'Succes', 5);
          }
        } else {
          // GmailApp zorgt ervoor dat het ook in Verzonden Items komt in de Google Workspace
          GmailApp.sendEmail(targetEmail, subject, 'Bekijk de accreditatie hier: ' + docUrl, mailOptions);
          
          // Noteer verzendmoment
          var now = new Date();
          var timestamp = Utilities.formatDate(now, ss.getSpreadsheetTimeZone(), "dd-MM HH:mm");
          sheet.getRange(rowNum, sendColIndex + 1).setValue(timestamp);
        }
        
        SpreadsheetApp.flush();
        totalProcessed++;
        
      } catch (err) {
        Logger.log('Fout bij verzenden/klaarzetten email naar ' + targetEmail + ': ' + err.message);
      }
    }
  }
  
  if (ui) {
    if (totalProcessed > 0) {
      if (isConcept) {
        ui.alert('Succes! ' + totalProcessed + ' concept e-mail(s) klaargezet.');
      } else {
        ui.alert('Succes! ' + totalProcessed + ' e-mail(s) verstuurd.');
      }
    } else {
      if (!alleenGeselecteerd) {
        ui.alert('Klaar. Er waren geen te verwerken e-mails gevonden (link ontbreekt of reeds verwerkt).');
      } else {
        ui.alert('Klaar. Voor de geselecteerde rij ontbreekt de link of is de mail al verstuurd.');
      }
    }
  }
}

/**
 * Synchroniseert data van alle sheets in de 'Doelmap' naar een externe 'Master Sheet'
 * met intelligente kolom-mapping.
 */
function syncAllToMaster(optSs) {
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
  
  var masterId = config['Master Sheet ID'];
  var folderId = config['Doelmap ID'];
  var startRow = parseInt(config['Beveilig template tot rij']) || 10;
  
  if (!masterId || !folderId) {
    if (ui) ui.alert('Fout: Zorg dat Master Sheet ID en Doelmap ID zijn ingevuld in de configuratie.');
    return;
  }
  
  var targetFolder;
  try {
    targetFolder = DriveApp.getFolderById(folderId);
  } catch (e) {
    if (ui) ui.alert('Fout: Doelmap ID is onjuist of niet toegankelijk.');
    return;
  }
  
  var masterSs;
  try {
    masterSs = SpreadsheetApp.openById(masterId);
  } catch (e) {
    if (ui) ui.alert('Fout: Master Sheet ID is onjuist of niet toegankelijk.');
    return;
  }
  
  // 2. Scan Proces
  var files = targetFolder.getFilesByType(MimeType.GOOGLE_SHEETS);
  var totalProcessedRows = 0;
  var totalNewRows = 0;
  var totalChangedRows = 0;
  var totalDeletedRows = 0;
  
  while (files.hasNext()) {
    var file = files.next();
    var sourceSs = SpreadsheetApp.openById(file.getId());
    var sourceSheets = sourceSs.getSheets();
    
    for (var s = 0; s < sourceSheets.length; s++) {
      var sourceSheet = sourceSheets[s];
      var sheetName = sourceSheet.getName();
      var masterSheet = masterSs.getSheetByName(sheetName);
      
      if (!masterSheet) continue; // Match op exacte tabblad naam
      
      // 3. Slimme Kolom-Mapping
      var masterLastCol = masterSheet.getLastColumn();
      var sourceLastCol = sourceSheet.getLastColumn();
      
      if (masterLastCol === 0 || sourceLastCol === 0) continue;
      
      var masterHeaders = masterSheet.getRange(9, 1, 1, masterLastCol).getValues()[0];
      var sourceHeaders = sourceSheet.getRange(9, 1, 1, sourceLastCol).getValues()[0];
      
      // Controleer/Voeg SyncKey en SyncStatus kolommen toe in Master
      var syncKeyIndex = masterHeaders.indexOf('SyncKey');
      if (syncKeyIndex === -1) {
        syncKeyIndex = masterLastCol;
        masterSheet.getRange(9, syncKeyIndex + 1).setValue('SyncKey');
        masterHeaders.push('SyncKey');
        masterLastCol++;
      }
      
      var statusIndex = masterHeaders.indexOf('SyncStatus');
      if (statusIndex === -1) {
        statusIndex = masterLastCol;
        masterSheet.getRange(9, statusIndex + 1).setValue('SyncStatus');
        masterHeaders.push('SyncStatus');
        masterLastCol++;
      }
      
      var colMap = {}; // source col index -> master col index
      var hasMapping = false;
      for (var sc = 0; sc < sourceHeaders.length; sc++) {
        var sh = sourceHeaders[sc];
        if (sh) {
          var mc = masterHeaders.indexOf(sh);
          if (mc !== -1) {
            colMap[sc] = mc;
            hasMapping = true;
          }
        }
      }
      
      if (!hasMapping) continue;
      
      // 4. Synchronisatie Logica per Rij
      var sourceLastRow = sourceSheet.getLastRow();
      if (sourceLastRow < startRow) continue;
      
      var sourceData = sourceSheet.getRange(startRow, 1, sourceLastRow - startRow + 1, sourceLastCol).getDisplayValues();
      var sourceFileId = file.getId();
      
      var masterLastRow = masterSheet.getLastRow();
      var masterData = [];
      var masterColors = [];
      if (masterLastRow >= startRow) {
        var masterRange = masterSheet.getRange(startRow, 1, masterLastRow - startRow + 1, masterLastCol);
        masterData = masterRange.getDisplayValues();
        masterColors = masterRange.getBackgrounds();
      }
      
      var masterKeyMap = {}; // SyncKey -> rowIndex in masterData
      for (var mr = 0; mr < masterData.length; mr++) {
        var key = masterData[mr][syncKeyIndex];
        if (key) {
          masterKeyMap[key] = mr;
        }
      }
      
      var newRows = [];
      var newColors = [];
      var sourceKeysPresent = {};
      var updatesNeeded = false;
      
      for (var r = 0; r < sourceData.length; r++) {
        var row = sourceData[r];
        
        // Controleer of de rij leeg is in de bron
        var isEmpty = true;
        for (var c = 0; c < row.length; c++) {
          if (row[c] !== '') {
            isEmpty = false;
            break;
          }
        }
        
        var key = sourceFileId + '_' + sheetName + '_' + (startRow + r);
        
        if (isEmpty) {
          if (masterKeyMap.hasOwnProperty(key)) {
            var mrIndex = masterKeyMap[key];
            if (masterData[mrIndex][statusIndex] !== 'Verwijderd') {
              masterData[mrIndex][statusIndex] = 'Verwijderd';
              updatesNeeded = true;
              totalDeletedRows++;
            }
          }
          continue;
        }
        
        sourceKeysPresent[key] = true;
        totalProcessedRows++;
        
        if (masterKeyMap.hasOwnProperty(key)) {
          // Bestaande regel
          var mrIndex = masterKeyMap[key];
          var masterRow = masterData[mrIndex];
          var changed = false;
          
          for (var sc in colMap) {
            var mc = colMap[sc];
            if (row[sc] !== masterRow[mc]) {
              masterData[mrIndex][mc] = row[sc];
              changed = true;
            }
          }
          
          if (changed) {
            masterData[mrIndex][statusIndex] = 'Gewijzigd';
            // Kleur rij lichtoranje
            for (var c = 0; c < masterLastCol; c++) {
              masterColors[mrIndex][c] = '#FFE5B4';
            }
            updatesNeeded = true;
            totalChangedRows++;
          }
        } else {
          // Nieuwe regel
          var newMasterRow = new Array(masterLastCol);
          for (var i = 0; i < masterLastCol; i++) newMasterRow[i] = '';
          
          for (var sc in colMap) {
            var mc = colMap[sc];
            newMasterRow[mc] = row[sc];
          }
          newMasterRow[syncKeyIndex] = key;
          newMasterRow[statusIndex] = 'Nieuw';
          
          newRows.push(newMasterRow);
          
          var rowColor = new Array(masterLastCol);
          for (var i = 0; i < masterLastCol; i++) rowColor[i] = null;
          newColors.push(rowColor);
          
          totalNewRows++;
        }
      }
      
      // Check voor rijen die verwijderd zijn uit de bron maar wel in de master staan
      for (var k in masterKeyMap) {
        if (k.indexOf(sourceFileId + '_' + sheetName + '_') === 0 && !sourceKeysPresent[k]) {
          var mrIndex = masterKeyMap[k];
          if (masterData[mrIndex][statusIndex] !== 'Verwijderd') {
            masterData[mrIndex][statusIndex] = 'Verwijderd';
            updatesNeeded = true;
            totalDeletedRows++;
          }
        }
      }
      
      // 5. Efficiëntie: Batch updates
      if (updatesNeeded && masterData.length > 0) {
        var updateRange = masterSheet.getRange(startRow, 1, masterData.length, masterLastCol);
        updateRange.setValues(masterData);
        updateRange.setBackgrounds(masterColors);
      }
      
      if (newRows.length > 0) {
        var targetStartRow = (masterLastRow >= startRow) ? (startRow + masterData.length) : startRow;
        var newRange = masterSheet.getRange(targetStartRow, 1, newRows.length, masterLastCol);
        newRange.setValues(newRows);
        newRange.setBackgrounds(newColors);
      }
      
      SpreadsheetApp.flush();
    }
  }
  
  if (ui) {
    ui.alert('Synchronisatie voltooid!\n\n' +
      'Rijen verwerkt: ' + totalProcessedRows + '\n' +
      'Nieuw: ' + totalNewRows + '\n' +
      'Gewijzigd: ' + totalChangedRows + '\n' +
      'Verwijderd: ' + totalDeletedRows);
  }
}
