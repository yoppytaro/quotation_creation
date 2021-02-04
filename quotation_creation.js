function getInfo() {
    var ssForm = SpreadsheetApp.openById('1ERFTlYNZxAUsRP6JhmZ6z4OTA9IKWBs2GfXRXMkhjc8');
    var sheetForm = ssForm.getSheetByName('記入情報');
    
    var lastRow = sheetForm.getLastRow();  
    
    var timestamp = sheetForm.getRange(lastRow, 1).getValue();
    var companyName = sheetForm.getRange(lastRow, 2).getValue();
    var person = sheetForm.getRange(lastRow, 3).getValue();
    var email = sheetForm.getRange(lastRow, 4).getValue();
    
    var documentNumber = lastRow - 1;
    
    //商品情報取得
    var item1 = sheetForm.getRange(lastRow, 5, 1, 4).getValues();
    var item2 = sheetForm.getRange(lastRow, 9,1,4).getValues();
    var item3 = sheetForm.getRange(lastRow, 13,1,4).getValues();
    
    item1[0].splice(1, 0,'','');
    item2[0].splice(1, 0,'','');
    item3[0].splice(1, 0,'','');
    
    //見積書テンプレート
    var ssQuotation = SpreadsheetApp.openById('1gFbyLBgRZ26ezjTgu3dH3gYT0SggwQerqRpSiWY3Jp0');
    //var sheetQuotation = ssQuotation.getSheetByName('見積書_テンプレート');
    
    //シート複製
    var now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_hhmm');
    var sheetName = now+'_'+companyName;
    var sheetQuotation = ssQuotation.duplicateActiveSheet().setName(sheetName);
    var sheetId = sheetQuotation.getSheetId();
    sheetQuotation.activate();
    
    //書き込み
    sheetQuotation.getRange('A4').setValue(companyName);
    sheetQuotation.getRange('A5').setValue(person+'様');
    sheetQuotation.getRange('F12').setValue(documentNumber);
    
    sheetQuotation.getRange('A17:F17').setValues(item1);
    sheetQuotation.getRange('A18:F18').setValues(item2);
    sheetQuotation.getRange('A19:F19').setValues(item3);
    
    //PDF出力
    
    SpreadsheetApp.flush();
    
    var url = 'https://docs.google.com/spreadsheets/d/1gFbyLBgRZ26ezjTgu3dH3gYT0SggwQerqRpSiWY3Jp0/export?exportFormat=pdf&gid=SID'.replace('SID',sheetId);
    var token = ScriptApp.getOAuthToken();
    var response = UrlFetchApp.fetch(url,{
      headers:{
        'Authorization':'Bearer '+ token
      }
    });
    
    var date = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd');
    var blob = response.getBlob().setName(date+'_'+companyName+'御中.pdf');
    
    var folder = DriveApp.getFolderById('1kicrVaBIeV_Iap3KOK0rTBRRrRnssyQ7');
    var file = folder.createFile(blob);
    
    var to = email;
    var subject = '見積書';
    var body = '見積書の添付ファイルです';
    var options = {
      attachments: [file]
    };
    
    GmailApp.sendEmail(
      to,
      subject,
      body,
      options
    );
    
  }
    