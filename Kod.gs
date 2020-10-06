const ss = SpreadsheetApp;
const drive = DriveApp;
const ui = ss.getUi();

const activeSs = ss.getActiveSpreadsheet();
const sheetsList = activeSs.getSheets();

function initMenu() {
  const menu = ui.createMenu('Magic');
  
  menu.addItem('Kopiera & byt datum', 'copyFile');
  menu.addItem('Sätt datum', 'setDates');
  menu.addToUi();
}
  
function onOpen() {
    initMenu();
}
  
function copyFile() {
  const newTitle = ui.prompt('Ange titel på kopian \n (ex. V11. 1/1 - 1/1)').getResponseText();
  
  if (newTitle) {
    const createTitle = ui.alert('Vill du skapa en kopia med titeln: ' + newTitle + '?', ui.ButtonSet.YES_NO);
    
    if (createTitle == 'YES') {
      drive.getFilesByName(activeSs.getName()).next().makeCopy(newTitle);
      ui.alert('Klart!');
    }
  }
}

function setDates() {
  const startDateString = ui.prompt('Ange startdatum \n (ex. 24):').getResponseText();
  const startDateInt = parseInt(startDateString); 
  
  let numberOfDaysInt = 0;
  
  if (startDateString) {
    const currentMonth  = ui.prompt('Vilken månad startar veckan i?').getResponseText().toUpperCase();
    
    switch(currentMonth) {
      case 'FEBRUARI':
        const isLeapyear = ui.alert('Är det skottår?', ui.ButtonSet.YES_NO);
      
        if (isLeapyear == 'YES') {
          numberOfDaysInt = 29;
        } else if (isLeapyear == 'NO'){
          numberOfDaysInt = 28;
        } break;
      case 'JANUARI' || 'MARS' || 'MAJ' || 'JULI' || 'AUGUSTI' || 'OKTOBER' || 'DECEMBER':
        numberOfDaysInt = 31;
        break;
      case 'APRIL' || 'JUNI' || 'SEPTEMBER' || 'NOVEMBER':
        numberOfDaysInt = 30;
        break;
      default:
        numberOfDaysInt = 0;
    }
    
    let dayOfNewMonth = 0;
    let nextMonth = false;
  
    for (let i = 0; i < sheetsList.length; i++) {
      const sheet = sheetsList[i];

      if (startDateInt + i > numberOfDaysInt) {
        dayOfNewMonth = dayOfNewMonth + 1;
        if (nextMonth) {
          sheet.getRange(1, 4).setValue(dayOfNewMonth);
          sheet.getRange(1, 5).setValue(nextMonth);
        } else {
          //nextMonth = ui.prompt('Vilken är nästa månad?').getResponseText();
          switch(currentMonth) {
            case 'JANUARI':
              nextMonth = 'FEBRUARI';
              break;
            case 'FEBRUARI':
              nextMonth = 'MARS';
              break;
            case 'MARS':
              nextMonth = 'APRIL';
              break;
            case 'APRIL':
              nextMonth = 'MAJ';
              break;
            case 'MAJ':
              nextMonth = 'JUNI';
              break;
            case 'JUNI':
              nextMonth = 'JULI';
              break;
            case 'JULI':
              nextMonth = 'AUGUSTI';
              break;
            case 'AUGUSTI':
              nextMonth = 'SEPTEMBER';
              break;
            case 'SEPTEMBER':
              nextMonth = 'OKTOBER';
              break;
            case 'OKTOBER':
              nextMonth = 'NOVEMBER';
              break;
            case 'NOVEMBER':
              nextMonth = 'DECEMBER';
              break;
            default:
              nextMonth = 'JANUARI';
          }
          
          sheet.getRange(1, 4).setValue(dayOfNewMonth);
          sheet.getRange(1, 5).setValue(nextMonth);
         }   
   
      } else {
        sheet.getRange(1, 4).setValue(startDateInt + i);
        sheet.getRange(1, 5).setValue(currentMonth);
      }
    }
  }
}