//useful resource: https://www.youtube.com/watch?v=QTySwuhpHG0
//Author: Jamal Seder  jamalseder@gmail.com
//Date: 1/5/2024

function onOpen(){
    const ui = SpreadsheetApp.getUi();
    const menu = ui.createMenu('AutoFill Docs');
    menu.addItem('Create Muna Student Reports', 'createNewGoogleDocsMuna');
    menu.addItem('Create Mostafa Student Reports', 'createNewGoogleDocsMostafa');
    menu.addItem('Create Mubarak Student Reports', 'createNewGoogleDocsMubarak');
    menu.addItem('Create Yaser Student Reports', 'createNewGoogleDocsYaser');
    menu.addToUi();
  }
  
  function createNewGoogleDocsMuna() {
    const googleDocTemplate = DriveApp.getFileById('1gAT0y3hKA6C2039_92BbagbdP0w2Z6O7lMjxJuAX2e8');
    const googleDocTemplate_Alphabet = DriveApp.getFileById('12PQJq2t8AxCJu5wv5BOcVUrpZEIuPahmuO5tJ6yr89E');
    const destinationFolder = DriveApp.getFolderById('1lVsXDf8DMgBDZRYi_KZVabk4Bq6-iIbN');
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Mona_Class_Roll');
    const rows = sheet.getDataRange().getValues();
    //Logger.log(rows);
    rows.forEach(function(row, index) {
      if(index === 0) return;
      if (row[6]) return;
      //const studentSheet = SpreadsheetApp.openById('1Gwm-znij837LbKWimAcYPiUjGnhppB2LiZqgByleKSk').getUrl();
  
      const studentSheet = SpreadsheetApp.openById('1Gwm-znij837LbKWimAcYPiUjGnhppB2LiZqgByleKSk').getSheetByName(`${row[1]}`);
      Logger.log(studentSheet);
      Logger.log(row[3]);
      if(row[3] === `Alphabet`){
        var copy = googleDocTemplate_Alphabet.makeCopy(`${row[1]} Student Report Card`, destinationFolder);
      }else{
        var copy = googleDocTemplate.makeCopy(`${row[1]} Student Report Card`, destinationFolder);
      }
  
      const doc = DocumentApp.openById(copy.getId())
      const body = doc.getBody();
      // const friendlyDate = new Date(row[3]).toLocaleDateString();
      body.replaceText('{{StudentName}}', row[1]);
      body.replaceText('{{TeacherName}}', 'Sheikha Muna');
      body.replaceText('{{OverallGrade}}', (Math.round(row[2] * 100) / 100).toFixed(2));
      
      studentRows = studentSheet.getDataRange().getValues();
      // Get the index of the last item in the list
      // Start a for loop from the index of the last item to the index of the third-last item
      var counter = 1;
      if(row[3] === `Alphabet`){
        studentRows.forEach(function(row, index){
        if((studentRows.length-1 === index || studentRows.length-2 === index || studentRows.length-3 === index)){
          
          const friendlyDate = new Date(row[0]).toLocaleDateString();
          body.replaceText(`{{Date_${counter}}}`, friendlyDate);
          body.replaceText(`{{AlphabetEvaluation_${counter}}}`, row[18]);
          body.replaceText(`{{Behavior_${counter}}}`, row[19]);
          body.replaceText(`{{Comments_${counter}}}`, row[20]);
          counter++;
        }
        //Do the code here that should go in the second half of the sheet. Check what kind of student it is, alphabet or memorization
      })
      }else{
        studentRows.forEach(function(row, index){
          if((studentRows.length-1 === index || studentRows.length-2 === index || studentRows.length-3 === index)){
            
            const friendlyDate = new Date(row[0]).toLocaleDateString();
            body.replaceText(`{{Date_${counter}}}`, friendlyDate);
            // body.replaceText(`{{RevSurahStart_${counter}}}`, row[9]);
            // body.replaceText(`{{RevAyahStart_${counter}}}`, row[10]);
            // body.replaceText(`{{RevSurahEnd_${counter}}}`, row[11]);
            // body.replaceText(`{{RevAyahEnd_${counter}}}`, row[12]);
            body.replaceText(`{{RevEval_${counter}}}`, row[13]);
            
            body.replaceText(`{{ReinEval_${counter}}}`, row[15]);
  
            // body.replaceText(`{{MemSurahStart_${counter}}}`, row[4]);
            // body.replaceText(`{{MemAyahStart_${counter}}}`, row[5]);
            // body.replaceText(`{{MemSurahEnd_${counter}}}`, row[6]);
            // body.replaceText(`{{MemAyahEnd_${counter}}}`, row[7]);
            body.replaceText(`{{MemEval_${counter}}}`, row[8]);
            
            //[{{NMemSurahStart_1}}:{{NMemAyahStart_1}}] - [{{NMemSurahEnd_1}}:{{NMemAyahEnd_1}}]
            
            body.replaceText(`{{NMemEval_${counter}}}`, row[17]);
  
            body.replaceText(`{{Behavior_${counter}}}`, row[19]);
            body.replaceText(`{{Comments_${counter}}}`, row[20]);
   
            counter++;
          }
        })
      }
      doc.saveAndClose();
      const url = doc.getUrl();
      sheet.getRange(index + 1, 7).setValue(url);
  
    })
  }
  
  
  function createNewGoogleDocsYaser() {
    const googleDocTemplate = DriveApp.getFileById('1gAT0y3hKA6C2039_92BbagbdP0w2Z6O7lMjxJuAX2e8');
    const googleDocTemplate_Alphabet = DriveApp.getFileById('12PQJq2t8AxCJu5wv5BOcVUrpZEIuPahmuO5tJ6yr89E');
    const destinationFolder = DriveApp.getFolderById('1u8zki-qsesUoE_PKWSMtZqHUPmH8Ffte');
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Yaser_Class_Roll');
    const rows = sheet.getDataRange().getValues();
    //Logger.log(rows);
    rows.forEach(function(row, index) {
      if(index === 0) return;
      if (row[6]) return;
      //const studentSheet = SpreadsheetApp.openById('1Gwm-znij837LbKWimAcYPiUjGnhppB2LiZqgByleKSk').getUrl();
  
      const studentSheet = SpreadsheetApp.openById('1ZELsv3xht3EfvQufs2wbRmUQgA-SbOpo2_IQpys5wts').getSheetByName(`${row[1]}`);
      Logger.log(studentSheet);
      Logger.log(row[3]);
      if(row[3] === `Alphabet`){
        var copy = googleDocTemplate_Alphabet.makeCopy(`${row[1]} Student Report Card`, destinationFolder);
      }else{
        var copy = googleDocTemplate.makeCopy(`${row[1]} Student Report Card`, destinationFolder);
      }
  
      const doc = DocumentApp.openById(copy.getId())
      const body = doc.getBody();
      // const friendlyDate = new Date(row[3]).toLocaleDateString();
      body.replaceText('{{StudentName}}', row[1]);
      body.replaceText('{{TeacherName}}', 'Sheikh Yaser');
      body.replaceText('{{OverallGrade}}', (Math.round(row[2] * 100) / 100).toFixed(2));
      
      studentRows = studentSheet.getDataRange().getValues();
      // Get the index of the last item in the list
      // Start a for loop from the index of the last item to the index of the third-last item
      var counter = 1;
      if(row[3] === `Alphabet`){
        studentRows.forEach(function(row, index){
        if((studentRows.length-1 === index || studentRows.length-2 === index || studentRows.length-3 === index)){
          
          const friendlyDate = new Date(row[0]).toLocaleDateString();
          body.replaceText(`{{Date_${counter}}}`, friendlyDate);
          body.replaceText(`{{AlphabetEvaluation_${counter}}}`, row[18]);
          body.replaceText(`{{Behavior_${counter}}}`, row[19]);
          body.replaceText(`{{Comments_${counter}}}`, row[20]);
          counter++;
        }
        //Do the code here that should go in the second half of the sheet. Check what kind of student it is, alphabet or memorization
      })
      }else{
        studentRows.forEach(function(row, index){
          if((studentRows.length-1 === index || studentRows.length-2 === index || studentRows.length-3 === index)){
            
            const friendlyDate = new Date(row[0]).toLocaleDateString();
            body.replaceText(`{{Date_${counter}}}`, friendlyDate);
            // body.replaceText(`{{RevSurahStart_${counter}}}`, row[9]);
            // body.replaceText(`{{RevAyahStart_${counter}}}`, row[10]);
            // body.replaceText(`{{RevSurahEnd_${counter}}}`, row[11]);
            // body.replaceText(`{{RevAyahEnd_${counter}}}`, row[12]);
            body.replaceText(`{{RevEval_${counter}}}`, row[13]);
            
            body.replaceText(`{{ReinEval_${counter}}}`, row[15]);
  
            // body.replaceText(`{{MemSurahStart_${counter}}}`, row[4]);
            // body.replaceText(`{{MemAyahStart_${counter}}}`, row[5]);
            // body.replaceText(`{{MemSurahEnd_${counter}}}`, row[6]);
            // body.replaceText(`{{MemAyahEnd_${counter}}}`, row[7]);
            body.replaceText(`{{MemEval_${counter}}}`, row[8]);
            
            //[{{NMemSurahStart_1}}:{{NMemAyahStart_1}}] - [{{NMemSurahEnd_1}}:{{NMemAyahEnd_1}}]
            
            body.replaceText(`{{NMemEval_${counter}}}`, row[17]);
  
            body.replaceText(`{{Behavior_${counter}}}`, row[19]);
            body.replaceText(`{{Comments_${counter}}}`, row[20]);
   
            counter++;
          }
        })
      }
      doc.saveAndClose();
      const url = doc.getUrl();
      sheet.getRange(index + 1, 7).setValue(url);
  
    })
  }
  
  function createNewGoogleDocsMubarak() {
    const googleDocTemplate = DriveApp.getFileById('1gAT0y3hKA6C2039_92BbagbdP0w2Z6O7lMjxJuAX2e8');
    const googleDocTemplate_Alphabet = DriveApp.getFileById('12PQJq2t8AxCJu5wv5BOcVUrpZEIuPahmuO5tJ6yr89E');
    const destinationFolder = DriveApp.getFolderById('1vhkdwATQJLTWSfayWk1z8nXzYotxICwb');
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Mubarak_Class_Roll');
    const rows = sheet.getDataRange().getValues();
    //Logger.log(rows);
    rows.forEach(function(row, index) {
      if(index === 0) return;
      if (row[6]) return;
      //const studentSheet = SpreadsheetApp.openById('1Gwm-znij837LbKWimAcYPiUjGnhppB2LiZqgByleKSk').getUrl();
  
      const studentSheet = SpreadsheetApp.openById('1xnwjsxMHStBlMWTYChGjFuHDXfufe70kLIaaM9pieCU').getSheetByName(`${row[1]}`);
      Logger.log(studentSheet);
      Logger.log(row[3]);
      if(row[3] === `Alphabet`){
        var copy = googleDocTemplate_Alphabet.makeCopy(`${row[1]} Student Report Card`, destinationFolder);
      }else{
        var copy = googleDocTemplate.makeCopy(`${row[1]} Student Report Card`, destinationFolder);
      }
  
      const doc = DocumentApp.openById(copy.getId())
      const body = doc.getBody();
      // const friendlyDate = new Date(row[3]).toLocaleDateString();
      body.replaceText('{{StudentName}}', row[1]);
      body.replaceText('{{TeacherName}}', 'Sheikh Mubarak');
      body.replaceText('{{OverallGrade}}', (Math.round(row[2] * 100) / 100).toFixed(2));
      
      studentRows = studentSheet.getDataRange().getValues();
      // Get the index of the last item in the list
      // Start a for loop from the index of the last item to the index of the third-last item
      var counter = 1;
      if(row[3] === `Alphabet`){
        studentRows.forEach(function(row, index){
        if((studentRows.length-1 === index || studentRows.length-2 === index || studentRows.length-3 === index)){
          
          const friendlyDate = new Date(row[0]).toLocaleDateString();
          body.replaceText(`{{Date_${counter}}}`, friendlyDate);
          body.replaceText(`{{AlphabetEvaluation_${counter}}}`, row[18]);
          body.replaceText(`{{Behavior_${counter}}}`, row[19]);
          body.replaceText(`{{Comments_${counter}}}`, row[20]);
          counter++;
        }
        //Do the code here that should go in the second half of the sheet. Check what kind of student it is, alphabet or memorization
      })
      }else{
        studentRows.forEach(function(row, index){
          if((studentRows.length-1 === index || studentRows.length-2 === index || studentRows.length-3 === index)){
            
            const friendlyDate = new Date(row[0]).toLocaleDateString();
            body.replaceText(`{{Date_${counter}}}`, friendlyDate);
            body.replaceText(`{{RevSurahStart_${counter}}}`, row[9]);
            body.replaceText(`{{RevAyahStart_${counter}}}`, row[10]);
            body.replaceText(`{{RevSurahEnd_${counter}}}`, row[11]);
            body.replaceText(`{{RevAyahEnd_${counter}}}`, row[12]);
            body.replaceText(`{{RevEval_${counter}}}`, row[13]);
            
            body.replaceText(`{{ReinEval_${counter}}}`, row[15]);
  
            body.replaceText(`{{MemSurahStart_${counter}}}`, row[4]);
            body.replaceText(`{{MemAyahStart_${counter}}}`, row[5]);
            body.replaceText(`{{MemSurahEnd_${counter}}}`, row[6]);
            body.replaceText(`{{MemAyahEnd_${counter}}}`, row[7]);
            body.replaceText(`{{MemEval_${counter}}}`, row[8]);
            
            //[{{NMemSurahStart_1}}:{{NMemAyahStart_1}}] - [{{NMemSurahEnd_1}}:{{NMemAyahEnd_1}}]
            
            body.replaceText(`{{NMemEval_${counter}}}`, row[17]);
  
            body.replaceText(`{{Behavior_${counter}}}`, row[19]);
            body.replaceText(`{{Comments_${counter}}}`, row[20]);
   
            counter++;
          }
        })
      }
      doc.saveAndClose();
      const url = doc.getUrl();
      sheet.getRange(index + 1, 7).setValue(url);
  
    })
  }
  
  function createNewGoogleDocsMostafa() {
    const googleDocTemplate = DriveApp.getFileById('1gAT0y3hKA6C2039_92BbagbdP0w2Z6O7lMjxJuAX2e8');
    const googleDocTemplate_Alphabet = DriveApp.getFileById('12PQJq2t8AxCJu5wv5BOcVUrpZEIuPahmuO5tJ6yr89E');
    const destinationFolder = DriveApp.getFolderById('1SkaVDgMRRYpEWeyC2mLUrRl8tVTpZvkw');
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Mostafa_Class_Roll');
    const rows = sheet.getDataRange().getValues();
    //Logger.log(rows);
    rows.forEach(function(row, index) {
      if(index === 0) return;
      if (row[6]) return;
      //const studentSheet = SpreadsheetApp.openById('1Gwm-znij837LbKWimAcYPiUjGnhppB2LiZqgByleKSk').getUrl();
  
      const studentSheet = SpreadsheetApp.openById('1prZpPAhkA2D6TVgKyPg5-vKppGzGzQDeprUzlK2PmWw').getSheetByName(`${row[1]}`);
      Logger.log(studentSheet);
      Logger.log(row[3]);
      if(row[3] === `Alphabet`){
        var copy = googleDocTemplate_Alphabet.makeCopy(`${row[1]} Student Report Card`, destinationFolder);
      }else{
        var copy = googleDocTemplate.makeCopy(`${row[1]} Student Report Card`, destinationFolder);
      }
  
      const doc = DocumentApp.openById(copy.getId())
      const body = doc.getBody();
      // const friendlyDate = new Date(row[3]).toLocaleDateString();
      body.replaceText('{{StudentName}}', row[1]);
      body.replaceText('{{TeacherName}}', 'Sheikh Mostafa');
      body.replaceText('{{OverallGrade}}', (Math.round(row[2] * 100) / 100).toFixed(2));
      
      studentRows = studentSheet.getDataRange().getValues();
      // Get the index of the last item in the list
      // Start a for loop from the index of the last item to the index of the third-last item
      var counter = 1;
      if(row[3] === `Alphabet`){
        studentRows.forEach(function(row, index){
        if((studentRows.length-1 === index || studentRows.length-2 === index || studentRows.length-3 === index)){
          
          const friendlyDate = new Date(row[0]).toLocaleDateString();
          body.replaceText(`{{Date_${counter}}}`, friendlyDate);
          body.replaceText(`{{AlphabetEvaluation_${counter}}}`, row[18]);
          body.replaceText(`{{Behavior_${counter}}}`, row[19]);
          body.replaceText(`{{Comments_${counter}}}`, row[20]);
          counter++;
        }
        //Do the code here that should go in the second half of the sheet. Check what kind of student it is, alphabet or memorization
      })
      }else{
        studentRows.forEach(function(row, index){
          if((studentRows.length-1 === index || studentRows.length-2 === index || studentRows.length-3 === index)){
            
            const friendlyDate = new Date(row[0]).toLocaleDateString();
            body.replaceText(`{{Date_${counter}}}`, friendlyDate);
            body.replaceText(`{{RevSurahStart_${counter}}}`, row[9]);
            body.replaceText(`{{RevAyahStart_${counter}}}`, row[10]);
            body.replaceText(`{{RevSurahEnd_${counter}}}`, row[11]);
            body.replaceText(`{{RevAyahEnd_${counter}}}`, row[12]);
            body.replaceText(`{{RevEval_${counter}}}`, row[13]);
            
            body.replaceText(`{{ReinEval_${counter}}}`, row[15]);
  
            body.replaceText(`{{MemSurahStart_${counter}}}`, row[4]);
            body.replaceText(`{{MemAyahStart_${counter}}}`, row[5]);
            body.replaceText(`{{MemSurahEnd_${counter}}}`, row[6]);
            body.replaceText(`{{MemAyahEnd_${counter}}}`, row[7]);
            body.replaceText(`{{MemEval_${counter}}}`, row[8]);
            
            //[{{NMemSurahStart_1}}:{{NMemAyahStart_1}}] - [{{NMemSurahEnd_1}}:{{NMemAyahEnd_1}}]
            
            body.replaceText(`{{NMemEval_${counter}}}`, row[17]);
  
            body.replaceText(`{{Behavior_${counter}}}`, row[19]);
            body.replaceText(`{{Comments_${counter}}}`, row[20]);
   
            counter++;
          }
        })
      }
      doc.saveAndClose();
      const url = doc.getUrl();
      sheet.getRange(index + 1, 7).setValue(url);
  
    })
  }
  