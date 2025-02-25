let previousMessageDate = PropertiesService.getScriptProperties().getProperty('previousMessageDate')
let previousMessageID = PropertiesService.getScriptProperties().getProperty('previousMessageID');
let latestEmailReport = {};
let webhookUrl = PropertiesService.getScriptProperties().getProperty('webhookUrl');
let emailSearchString = PropertiesService.getScriptProperties().getProperty('emailSearchString')
let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PIAS UK sales last week")

function main() {

  try {
    latestEmailReport = searchForNewEmailReport()
    if (previousMessageID == latestEmailReport.messageID){
      Logger.log("No new report email since last update on " + previousMessageDate);
    }

    else {
      data = extractDataFromReport(latestEmailReport.xlsxBlob);
      writeDataToSheet(data, sheet);
      sendToGoogleChat(data);
      GmailApp.getMessageById(latestEmailReport.messageID).markRead();
      PropertiesService.getScriptProperties().setProperty('previousMessageID', latestEmailReport.messageID);
      PropertiesService.getScriptProperties().setProperty('previousMessageDate', latestEmailReport.messageDate);
    }
  }
  catch(error) {
    Logger.log('Error: ' + error);
  }  
}  

function searchForNewEmailReport(){

  let threads = GmailApp.search(emailSearchString);
  let message = threads[0].getMessages()[0];
  let messageDate = message.getDate();
  let messageSubject = message.getSubject();
  let messageID = message.getId();
  let xlsxBlob = message.getAttachments()[0];

  return {
    messageDate: messageDate,
    messageSubject: messageSubject,
    messageID: messageID,
    xlsxBlob: xlsxBlob,
  }
}

function extractDataFromReport(xlsxBlob){
  
    let tempSpreadsheetId = Drive.Files.create({mimeType: MimeType.GOOGLE_SHEETS}, xlsxBlob).id;
    Logger.log(tempSpreadsheetId);

    let tempSheet = SpreadsheetApp.openById(tempSpreadsheetId).getSheets()[0];
    let tempData = tempSheet.getRange(4,1,tempSheet.getLastRow(),5).getValues();
    Drive.Files.remove(tempSpreadsheetId); 

    let matchText = "TOTAL";
    let lastDataRow = tempData.findIndex(row => row[0] === matchText);
    let processedData = tempSheet.getRange(4,1,lastDataRow,5).getValues();
    let aggregatedData = {};

    processedData.forEach(row =>{
      if (row[4] === 0) return;

      let key = row[0] + "|" + row[0] + "|" + row[1] + "|" + row[2] + "|" +row[3];
      if (!aggregatedData[key]) {
        aggregatedData[key] = {
          article: row[0],
          title: row[1],
          ean: row[2],
          format: row[3],
          salesQty: 0
        };
        }
        aggregatedData[key].salesQty += row[4];
    });

    let outputData = Object.values(aggregatedData).map(item => [
      item.article,
      item.title,
      // item.ean,
      item.format,
      item.salesQty
    ]);

    outputData.sort((a, b) => b[3] - a[3]);
    outputData.filter

    return outputData;
}

function writeDataToSheet(data, sheet){

    let clearRange = sheet.getRange(3,1,sheet.getLastRow(),sheet.getLastColumn());
      clearRange.clearContent();

    let sheetRange = sheet.getRange(3, 1, data.length, data[0].length);
    sheetRange.setValues(data);

    sheet.getRange("A1").setValue("Last report: " + latestEmailReport.messageDate.toDateString());
}

function sendToGoogleChat(data){

  let top15Sales = "";

  for (let i = 0; i < 15; i++) {

    top15Sales += data[i][3] + " x " + data[i][0] + " - " + data[i][1] + "\n";  
  
  }

  Logger.log(top15Sales);
  let url = PropertiesService.getScriptProperties().getProperty('spreadsheetURL')
  let options = {"method": "post","headers": {"Content-Type": "application/json; charset=UTF-8"},"payload": JSON.stringify(jsonPayload(top15Sales, url))};

  UrlFetchApp.fetch(webhookUrl, options);  
};
