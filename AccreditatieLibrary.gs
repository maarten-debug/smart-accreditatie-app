/**
 * Haalt BOD contacten op uit de Contactpersonen sheet.
 */
function getBODContacts(ss) {
  var contactSheet = ss.getSheetByName('Contactpersonen BOD');
  if (!contactSheet) contactSheet = ss.getSheetByName('Contactpersonen');
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
  var contactSheet = ss.getSheetByName('Contactpersonen BOD');
  if (!contactSheet) contactSheet = ss.getSheetByName('Contactpersonen');
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
        var protection;
        if (sheetProtections.length > 0) {
          protection = sheetProtections[0];
          // Verwijder eventuele extra overbodige beveiligingen
          for (var p = 1; p < sheetProtections.length; p++) {
            sheetProtections[p].remove();
          }
        } else {
          protection = currentSheet.protect().setDescription('Systeem Beveiliging');
        }
        
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
  var configOnderwerp = config['Onderwerp'];
  
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
      
      var subject = configOnderwerp ? configOnderwerp : ('Accreditatie ' + bNaam);
      subject = subject.replace(new RegExp('\\{\\{Bedrijfsnaam\\}\\}', 'gi'), bNaam);
      
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
  
  var scriptProperties = PropertiesService.getScriptProperties();
  var lastSyncTimeString = scriptProperties.getProperty('lastSyncTime');
  var lastSyncTime = lastSyncTimeString ? parseInt(lastSyncTimeString, 10) : 0;
  
  // 1. Lookup Tabel maken
  var bedrijfInfo = {};
  var accSheet = ss.getSheetByName('Accreditatie');
  if (accSheet) {
    var accData = accSheet.getDataRange().getValues();
    if (accData.length > 0) {
      var accHeaders = accData[0];
      var bNameIdx = accHeaders.indexOf('Bedrijfsnaam');
      var cPersoonIdx = accHeaders.indexOf('Contactpersoon');
      var cBodIdx = accHeaders.indexOf('Contactpersoon BOD');
      var bandjeIdx = accHeaders.indexOf('Standaard Bandje');
      var cateringIdx = accHeaders.indexOf('Standaard Catering');
      
      if (bNameIdx !== -1) {
        for (var i = 1; i < accData.length; i++) {
          var bn = accData[i][bNameIdx] ? accData[i][bNameIdx].toString().trim() : '';
          if (bn) {
            bedrijfInfo[bn] = {
              'Contactpersoon': cPersoonIdx !== -1 ? accData[i][cPersoonIdx] : '',
              'Contactpersoon BOD': cBodIdx !== -1 ? accData[i][cBodIdx] : '',
              'Standaard Bandje': bandjeIdx !== -1 ? accData[i][bandjeIdx] : '',
              'Standaard Catering': cateringIdx !== -1 ? accData[i][cateringIdx] : ''
            };
          }
        }
      }
    }
  }

  var configObj = getConfiguratie(ss);
  
  if (!configObj) {
    if (ui) ui.alert('Fout: Tabblad "Configuratie" niet gevonden.');
    return;
  }
  
  var config = configObj.config;
  
  var masterId = config['Master Sheet ID'];
  var folderId = config['Doelmap ID'];
  var sourceStartRow = parseInt(config['Beveilig template tot rij']) || 10;
  var masterStartRow = 2;
  var outputKolomNaam = config['Output Kolom'];
  var outputKolomFilledNaam = config['Output Kolom Filled'] || 'Ingevuld';
  
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
  
  var files = targetFolder.getFilesByType(MimeType.GOOGLE_SHEETS);
  var totalProcessedRows = 0;
  var totalNewRows = 0;
  var totalChangedRows = 0;
  var totalDeletedRows = 0;
  
  // GLOBAL CACHE FOR MASTER SHEETS
  var masterCache = {};
  var processedFileIds = {};
  var filesWithChanges = {};
  
  // 2. Scan Proces: Inlezen en updaten in geheugen
  while (files.hasNext()) {
    var file = files.next();
    var sourceFileId = file.getId();
    
    var fileLastUpdated = file.getLastUpdated().getTime();
    if (lastSyncTime > 0 && fileLastUpdated <= lastSyncTime) {
      continue; // Bestand is niet gewijzigd sinds laatste sync, sla over
    }
    
    processedFileIds[sourceFileId] = true;
    
    // Bedrijfsnaam extraheren
    var fileName = file.getName();
    var nameParts = fileName.split(' - ');
    var extractedBedrijfsnaam = nameParts.length > 1 ? nameParts[1].trim() : '';
    
    var spreadsheet;
    try {
      spreadsheet = Sheets.Spreadsheets.get(sourceFileId);
    } catch (e) {
      Logger.log("Kan file " + sourceFileId + " niet ophalen via API.");
      continue;
    }
    var sourceSheets = spreadsheet.sheets || [];
    
    for (var s = 0; s < sourceSheets.length; s++) {
      var sheetName = sourceSheets[s].properties.title;
      
      var masterSheet = masterSs.getSheetByName(sheetName);
      if (!masterSheet) continue; // Match op exacte tabblad naam
      
      // Initialize cache for this sheetName if not exists
      if (!masterCache[sheetName]) {
        var m_lastCol = masterSheet.getLastColumn();
        if (m_lastCol === 0) continue;
        
        var m_headers = masterSheet.getRange(1, 1, 1, m_lastCol).getValues()[0];
        
        // Controleer/Voeg SyncKey en SyncStatus kolommen toe in Master
        var syncKeyIndex = m_headers.indexOf('SyncKey');
        if (syncKeyIndex === -1) {
          syncKeyIndex = m_lastCol;
          masterSheet.getRange(1, syncKeyIndex + 1).setValue('SyncKey');
          m_headers.push('SyncKey');
          m_lastCol++;
        }
        
        var statusIndex = m_headers.indexOf('SyncStatus');
        if (statusIndex === -1) {
          statusIndex = m_lastCol;
          masterSheet.getRange(1, statusIndex + 1).setValue('SyncStatus');
          m_headers.push('SyncStatus');
          m_lastCol++;
        }
        
        var m_lastRow = masterSheet.getLastRow();
        var m_data = [];
        if (m_lastRow >= masterStartRow) {
          var masterRange = masterSheet.getRange(masterStartRow, 1, m_lastRow - masterStartRow + 1, m_lastCol);
          m_data = masterRange.getDisplayValues();
        }
        
        var m_keyMap = {};
        for (var mr = 0; mr < m_data.length; mr++) {
          var key = m_data[mr][syncKeyIndex];
          if (key) {
            m_keyMap[key] = mr;
          }
        }
        
        masterCache[sheetName] = {
          sheet: masterSheet,
          lastCol: m_lastCol,
          lastRow: m_lastRow,
          headers: m_headers,
          data: m_data,
          colors: m_colors,
          keyMap: m_keyMap,
          syncKeyIndex: syncKeyIndex,
          statusIndex: statusIndex,
          bNameMasterIdx: m_headers.indexOf('Bedrijfsnaam'),
          cPersoonMasterIdx: m_headers.indexOf('Contactpersoon bedrijf'),
          cBodMasterIdx: m_headers.indexOf('Contactpersoon BOD'),
          bandjeMasterIdx: m_headers.indexOf('Bandje_Accr'),
          cateringMasterIdx: m_headers.indexOf('Catering_Accr'),
          newRows: [],
          coloredRows: [],
          updatesNeeded: false,
          sourceKeysPresent: {}
        };
      }
      
      var cache = masterCache[sheetName];
      var sourceDataResponse;
      try {
        sourceDataResponse = Sheets.Spreadsheets.Values.get(sourceFileId, "'" + sheetName + "'");
      } catch (e) {
        continue;
      }
      var allValues = sourceDataResponse.values;
      if (!allValues || allValues.length < 9) continue; // Headers staan op rij 9 (index 8)
      
      var sourceHeaders = allValues[8];
      var sourceLastCol = sourceHeaders.length;
      if (sourceLastCol === 0) continue;
      
      var colMap = {}; // source col index -> master col index
      var hasMapping = false;
      for (var sc = 0; sc < sourceHeaders.length; sc++) {
        var sh = sourceHeaders[sc];
        if (sh) {
          var mc = cache.headers.indexOf(sh);
          if (mc !== -1) {
            colMap[sc] = mc;
            hasMapping = true;
          }
        }
      }
      
      if (!hasMapping) continue;
      
      if (allValues.length < sourceStartRow) continue;
      
      var sourceData = allValues.slice(sourceStartRow - 1);
      var voornaamIdx = sourceHeaders.indexOf('Voornaam');
      var achternaamIdx = sourceHeaders.indexOf('Achternaam');
      
      for (var r = 0; r < sourceData.length; r++) {
        var row = sourceData[r];
        
        // Controleer of de rij leeg is in de bron
        var vNaam = (voornaamIdx !== -1 && row[voornaamIdx]) ? row[voornaamIdx].toString().trim() : '';
        var aNaam = (achternaamIdx !== -1 && row[achternaamIdx]) ? row[achternaamIdx].toString().trim() : '';

        if (vNaam === '' && aNaam === '') {
          continue; // Sla deze rij over, het is een lege (of alleen checkbox) rij
        }
        
        var key = sourceFileId + '_' + sheetName + '_' + (sourceStartRow + r);
        
        cache.sourceKeysPresent[key] = true;
        totalProcessedRows++;
        
        if (cache.keyMap.hasOwnProperty(key)) {
          // Bestaande regel
          var mrIndex = cache.keyMap[key];
          var masterRow = cache.data[mrIndex];
          var changed = false;
          
          for (var sc in colMap) {
            var mc = colMap[sc];
            
            var valSource = row[sc] ? row[sc].toString().trim() : '';
            var valMaster = masterRow[mc] ? masterRow[mc].toString().trim() : '';
            
            var vS_low = valSource.toLowerCase();
            var vM_low = valMaster.toLowerCase();
            
            // Normaliseer checkboxes
            if (vS_low === 'waar') vS_low = 'true';
            if (vS_low === 'onwaar') vS_low = 'false';
            if (vM_low === 'waar') vM_low = 'true';
            if (vM_low === 'onwaar') vM_low = 'false';
            
            // Normaliseer voorloopnullen (getallen)
            if (vS_low !== vM_low && !isNaN(valMaster) && valMaster !== '') {
              if (vS_low.replace(/^0+/, '') === vM_low.replace(/^0+/, '')) {
                vS_low = vM_low; // Ze zijn wiskundig hetzelfde
              }
            }
            
            if (vS_low !== vM_low) {
              cache.data[mrIndex][mc] = row[sc];
              changed = true;
            }
          }
          
          // Speciale Kolommen injecteren
          if (cache.bNameMasterIdx !== -1 && extractedBedrijfsnaam) {
            var bMaster = masterRow[cache.bNameMasterIdx] ? masterRow[cache.bNameMasterIdx].toString().trim() : '';
            var bSource = extractedBedrijfsnaam.toString().trim();
            if (bSource.toLowerCase() !== bMaster.toLowerCase()) {
              cache.data[mrIndex][cache.bNameMasterIdx] = extractedBedrijfsnaam;
              changed = true;
            }
          }
          
          var info = extractedBedrijfsnaam ? bedrijfInfo[extractedBedrijfsnaam] : null;
          if (info) {
            if (cache.cPersoonMasterIdx !== -1) {
              var cMaster = masterRow[cache.cPersoonMasterIdx] ? masterRow[cache.cPersoonMasterIdx].toString().trim() : '';
              var cSource = info['Contactpersoon'] ? info['Contactpersoon'].toString().trim() : '';
              if (cSource.toLowerCase() !== cMaster.toLowerCase()) {
                cache.data[mrIndex][cache.cPersoonMasterIdx] = info['Contactpersoon'];
                changed = true;
              }
            }
            if (cache.cBodMasterIdx !== -1) {
              var bodMaster = masterRow[cache.cBodMasterIdx] ? masterRow[cache.cBodMasterIdx].toString().trim() : '';
              var bodSource = info['Contactpersoon BOD'] ? info['Contactpersoon BOD'].toString().trim() : '';
              if (bodSource.toLowerCase() !== bodMaster.toLowerCase()) {
                cache.data[mrIndex][cache.cBodMasterIdx] = info['Contactpersoon BOD'];
                changed = true;
              }
            }
            if (cache.bandjeMasterIdx !== -1) {
              var bandjeMaster = masterRow[cache.bandjeMasterIdx] ? masterRow[cache.bandjeMasterIdx].toString().trim() : '';
              if (bandjeMaster === '') {
                var bandjeSource = info['Standaard Bandje'] ? info['Standaard Bandje'].toString().trim() : '';
                if (bandjeSource !== '') {
                  cache.data[mrIndex][cache.bandjeMasterIdx] = info['Standaard Bandje'];
                  changed = true;
                }
              }
            }
            if (cache.cateringMasterIdx !== -1) {
              var cateringMaster = masterRow[cache.cateringMasterIdx] ? masterRow[cache.cateringMasterIdx].toString().trim() : '';
              if (cateringMaster === '') {
                var cateringSource = info['Standaard Catering'] ? info['Standaard Catering'].toString().trim() : '';
                if (cateringSource !== '') {
                  cache.data[mrIndex][cache.cateringMasterIdx] = info['Standaard Catering'];
                  changed = true;
                }
              }
            }
          }
          
          if (changed) {
            cache.data[mrIndex][cache.statusIndex] = 'Gewijzigd';
            cache.coloredRows.push(masterStartRow + mrIndex); // Opslaan voor API batchUpdate (1-based rijnummering wordt later 0-based index)
            cache.updatesNeeded = true;
            totalChangedRows++;
            filesWithChanges[sourceFileId] = true;
          }
        } else {
          // Nieuwe regel
          var newMasterRow = new Array(cache.lastCol);
          for (var i = 0; i < cache.lastCol; i++) newMasterRow[i] = '';
          
          for (var sc in colMap) {
            var mc = colMap[sc];
            newMasterRow[mc] = row[sc];
          }
          
          // Speciale Kolommen injecteren
          if (cache.bNameMasterIdx !== -1 && extractedBedrijfsnaam) {
            newMasterRow[cache.bNameMasterIdx] = extractedBedrijfsnaam;
          }
          var info = extractedBedrijfsnaam ? bedrijfInfo[extractedBedrijfsnaam] : null;
          if (info) {
            if (cache.cPersoonMasterIdx !== -1) {
              newMasterRow[cache.cPersoonMasterIdx] = info['Contactpersoon'] || '';
            }
            if (cache.cBodMasterIdx !== -1) {
              newMasterRow[cache.cBodMasterIdx] = info['Contactpersoon BOD'] || '';
            }
            if (cache.bandjeMasterIdx !== -1) {
              newMasterRow[cache.bandjeMasterIdx] = info['Standaard Bandje'] || '';
            }
            if (cache.cateringMasterIdx !== -1) {
              newMasterRow[cache.cateringMasterIdx] = info['Standaard Catering'] || '';
            }
          }
          
          newMasterRow[cache.syncKeyIndex] = key;
          newMasterRow[cache.statusIndex] = 'Nieuw';
          
          cache.newRows.push(newMasterRow);
          // Nieuwe rijen worden later tijdens de save-fase in coloredRows gestoken
          
          totalNewRows++;
          filesWithChanges[sourceFileId] = true;
        }
      }
    }
  }
  
  // 3. Batch Updates uitvoeren
  var batchUpdateRequests = [];
  
  for (var sheetName in masterCache) {
    var cache = masterCache[sheetName];
    
    // Check voor rijen die verwijderd zijn uit de bron maar wel in de master staan
    for (var k in cache.keyMap) {
      // Verwijder alleen rijen van bestanden die nog steeds in de doelmap staan
      var sourceFileId = k.split('_')[0];
      if (processedFileIds[sourceFileId] && !cache.sourceKeysPresent[k]) {
        var mrIndex = cache.keyMap[k];
        if (cache.data[mrIndex][cache.statusIndex] !== 'Verwijderd') {
          cache.data[mrIndex][cache.statusIndex] = 'Verwijderd';
          cache.updatesNeeded = true;
          totalDeletedRows++;
          filesWithChanges[sourceFileId] = true;
        }
      }
    }
    
    // Schrijf alle aanpassingen in 1x naar de sheet
    if (cache.updatesNeeded && cache.data.length > 0) {
      var updateRange = cache.sheet.getRange(masterStartRow, 1, cache.data.length, cache.lastCol);
      updateRange.setValues(cache.data);
    }
    
    if (cache.newRows.length > 0) {
      var targetStartRow = (cache.lastRow >= masterStartRow) ? (masterStartRow + cache.data.length) : masterStartRow;
      var newRange = cache.sheet.getRange(targetStartRow, 1, cache.newRows.length, cache.lastCol);
      newRange.setValues(cache.newRows);
      
      // Bereken de index voor de nieuwe rijen en sla op voor kleuring
      for (var i = 0; i < cache.newRows.length; i++) {
        cache.coloredRows.push(targetStartRow + i);
      }
    }
    
    // Bouw de batchUpdate API payload op voor deze specifieke tabblad
    if (cache.coloredRows.length > 0) {
      var sheetId = cache.sheet.getSheetId();
      for (var i = 0; i < cache.coloredRows.length; i++) {
        var rIdx = cache.coloredRows[i] - 1; // Advanced Sheets API is 0-based
        batchUpdateRequests.push({
          repeatCell: {
            range: {
              sheetId: sheetId,
              startRowIndex: rIdx,
              endRowIndex: rIdx + 1,
              startColumnIndex: 0,
              endColumnIndex: cache.lastCol
            },
            cell: {
              userEnteredFormat: {
                backgroundColor: {
                  red: 1.0,
                  green: 0.898,
                  blue: 0.706
                }
              }
            },
            fields: "userEnteredFormat.backgroundColor"
          }
        });
      }
    }
  }
  
  if (batchUpdateRequests.length > 0) {
    Sheets.Spreadsheets.batchUpdate({
      requests: batchUpdateRequests
    }, masterId);
  }
  
  // 4. Update 'Ingevuld' status in het 'Accreditatie' tabblad
  if (accSheet && outputKolomNaam && Object.keys(filesWithChanges).length > 0) {
    var accLastRow = accSheet.getLastRow();
    var accLastCol = accSheet.getLastColumn();
    if (accLastRow >= 2 && accLastCol >= 1) {
      var accHeadersRange = accSheet.getRange(1, 1, 1, accLastCol);
      var accHeaders = accHeadersRange.getValues()[0].map(function(h) { return h.toString().trim(); });
      
      var linkColIdx = accHeaders.indexOf(outputKolomNaam);
      if (linkColIdx !== -1) {
        var filledColIdx = accHeaders.indexOf(outputKolomFilledNaam);
        if (filledColIdx === -1) {
          filledColIdx = accLastCol;
          accSheet.getRange(1, filledColIdx + 1).setValue(outputKolomFilledNaam);
          accLastCol++;
        }
        
        var accDataRange = accSheet.getRange(2, 1, accLastRow - 1, accLastCol);
        var accData = accDataRange.getDisplayValues();
        var accFormulas = accDataRange.getFormulas();
        var accChanged = false;
        
        var now = new Date();
        var timestamp = Utilities.formatDate(now, ss.getSpreadsheetTimeZone(), "dd-MM HH:mm");
        
        for (var i = 0; i < accData.length; i++) {
          var rowData = accData[i];
          var stringLink = rowData[linkColIdx] || '';
          var formulaLink = accFormulas[i][linkColIdx] || '';
          var docUrl = '';
          
          var match = formulaLink.match(/HYPERLINK\("([^"]+)"/i);
          if (match) {
            docUrl = match[1];
          } else if (stringLink.indexOf('http') === 0) {
            docUrl = stringLink;
          }
          
          if (docUrl) {
            var fileIdMatch = docUrl.match(/[-\w]{25,}/);
            if (fileIdMatch && fileIdMatch[0]) {
              var fId = fileIdMatch[0];
              if (filesWithChanges[fId]) {
                var currentStatus = rowData[filledColIdx];
                if (currentStatus !== timestamp) {
                  accData[i][filledColIdx] = timestamp;
                  accChanged = true;
                }
              }
            }
          }
        }
        
        if (accChanged) {
          var filledColValues = [];
          for (var i = 0; i < accData.length; i++) {
            filledColValues.push([accData[i][filledColIdx] !== undefined ? accData[i][filledColIdx] : '']);
          }
          accSheet.getRange(2, filledColIdx + 1, filledColValues.length, 1).setValues(filledColValues);
        }
      }
    }
  }

  SpreadsheetApp.flush();
  
  var currentSyncTime = new Date().getTime();
  scriptProperties.setProperty('lastSyncTime', currentSyncTime.toString());
  
  if (ui) {
    ui.alert('Synchronisatie voltooid!\n\n' +
      'Rijen verwerkt: ' + totalProcessedRows + '\n' +
      'Nieuw: ' + totalNewRows + '\n' +
      'Gewijzigd: ' + totalChangedRows + '\n' +
      'Verwijderd: ' + totalDeletedRows);
  }
}
