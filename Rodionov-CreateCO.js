const ss = SpreadsheetApp.openById("1A98m1NIEmJHX526soHKv403nf2DtmoJh4sj5ly_IjDg");
const choseOfgoods = ss.getSheetByName("Выбор товаров");
const orders = ss.getSheetByName("Заказы");
const constructorPage = ss.getSheetByName("ListConstructor");
const activeCell = SpreadsheetApp.getActive().getActiveCell();
const lr = orders.getLastRow();
const lc = orders.getLastColumn();

function onEdit() {
  const clientsName = activeCell.getValue();
  if (activeCell.getColumn() === 2 && activeCell.getRow() === 1) {
    choseOfgoods.getRange(3, 2, lr, lc).clearContent();
    choseOfgoods.getRange(3, 1, lr, lc).clear();
    choseOfgoods.getRange(3, 1, lr).setDataValidation(null);
    showClientsOrders(clientsName);
  }
  if (activeCell.getColumn() === 3 && activeCell.getRow() > 200) {
    createListOfClients();
  }
}

function onOpen() {
  SpreadsheetApp
    .getUi()
    .createMenu('Сохранить как pdf')
    .addItem('Сохранить как pdf', 'pdfAndExcelDownloader')
    .addToUi();
}
//

function createListOfClients() {
  const validationRange = constructorPage
    .getRange(1, 1, lr - 199)
    .setValues(orders.getRange(200, 3, lr - 199).getValues())
    .sort(1);
  const validationRule = SpreadsheetApp.newDataValidation().requireValueInRange(validationRange).build();
  const selectListCell = choseOfgoods.getRange('B1');
  
  selectListCell.setDataValidation(validationRule);
}


//

function showClientsOrders(clientsName) {
  const allOrders = orders.getRange(200, 1, lr - 199, lc).getValues();
  const chekBox = SpreadsheetApp.newDataValidation().requireCheckbox().build()
  let result = [];

  for (let i = 0; i < allOrders.length; i++) {
    if (allOrders[i][2][0] === "0"
      && ("0" + String(clientsName)) === allOrders[i][2]
      || ("00" + String(clientsName)) === allOrders[i][2]
      || ("000" + String(clientsName)) === allOrders[i][2]) {
      result.push(allOrders.indexOf(allOrders[i]));
    }
    if (allOrders[i][2] === clientsName) {
      result.push(allOrders.indexOf(allOrders[i]));   //узнаём номер строки в таблице "заказы"
    }
  }
  for (let j = 0; j < result.length; j++) {
    orders.getRange(200 + result[j], 2, 1, lc).copyTo(choseOfgoods.getRange(3 + j, 2, 1));
    choseOfgoods.getRange(3 + j, lc + 1).setValue(result[j] + 200); // добавляем столбец с номером строки в таблице "заказы"
    choseOfgoods.getRange(3 + j, 1).setDataValidation(chekBox).check();
  }
}

//


const constructorList = ss.getSheetByName("Конструктор КП");
const commissionOfferList = ss.getSheetByName("КП");
const footerConstructor = ss.getSheetByName("Footer constructor");

function constructorDataFilling() {
  const constructorLR = choseOfgoods.getLastRow();
  const listOrders = choseOfgoods.getRange(3, 1, constructorLR, lc).getValues();

  constructorList.getRange(3, 1, 1010, 7).clear();

  let result = [];
  let increment = 1;
  let itemIncrementForRec = 0;
  let sumPart = 0;
  let sumFastCargo = 0;
  let sumSlowCargo = 0;
  const showList = listOrders.filter(order => order[0] === true)  //собираем отмеченные заказы

  for (let i = 0; i < showList.length; i++) {
    //готовим ячейку для изображени
    const mergeCellForImg = constructorList.getRange(15 + itemIncrementForRec, 2, 13);
    //готовим ячейку для заголовка Тех. хар.
    const mergeForHeaderTech = constructorList.getRange(15 + itemIncrementForRec, 5, 1, 3);
    //готовим ячейку для заголовка Стоимость
    const mergeCellForHeaderCost = constructorList.getRange(20 + itemIncrementForRec, 5, 1, 3);
    const mergeCellsNumbers = [16, 17, 18, 19, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30];

    let dataObj = {}

    //if (listOrders[i][0] === true) {
      choseOfgoods
        .getRange(3 + i, 4)
        .copyTo(constructorList.getRange(15 + itemIncrementForRec, 2)); //копируем изображение
      //корректируем стили ячеек:
      constructorList.getRange(15 + itemIncrementForRec, 2)
        .setBorder(true, true, true, true, true, true, "white", SpreadsheetApp.BorderStyle.SOLID);
      constructorList.getRange(27 + itemIncrementForRec, 2)
        .setBorder(true, true, true, true, true, true, "white", SpreadsheetApp.BorderStyle.SOLID);

      dataObj.increment = increment;                                           //инкремент
      dataObj.quantity = choseOfgoods.getRange(3 + i, 14).getValue();          //количество
      dataObj.titleOfCargoVolume = choseOfgoods.getRange('S2').getValue();     //название ячейки объем груза куб.м
      dataObj.cargoVolume = choseOfgoods.getRange(3 + i, 19).getValue();       //объем груза куб.м
      dataObj.titleOfCargoMass = choseOfgoods.getRange('T2').getValue();       //название ячейки Масса груза, кг
      dataObj.cargoMass = choseOfgoods.getRange(3 + i, 20).getValue();         //Масса груза, кг
      dataObj.titleOfProtecSheath = choseOfgoods.getRange('U2').getValue();    //название ячейки Масса обрешетки, кг
      dataObj.protecSheathMass = choseOfgoods.getRange(3 + i, 21).getValue();  //Масса обрешетки, кг
      dataObj.titleTotalWeight = choseOfgoods.getRange('V2').getValue();       //название Общая масса с учетом обрешетки, кг
      dataObj.totalWeight = choseOfgoods.getRange(3 + i, 22).getValue();       //Общая масса с учетом обрешетки, кг
      dataObj.titleCargoFastPrice = choseOfgoods.getRange('AA2').getValue();   //название Ставка карго $/кг 18-24 дней
      dataObj.cargoFastPrice = choseOfgoods.getRange(3 + i, 27).getValue();    //Ставка карго $/кг 18-24 дней
      dataObj.titleCargoSlowPrice = choseOfgoods.getRange('AB2').getValue();   //название Ставка карго $/кг 27-30 дней
      dataObj.cargoSlowPrice = choseOfgoods.getRange(3 + i, 28).getValue();    //Ставка карго $/кг 27-30 дней
      dataObj.titleCargoPackPrice = choseOfgoods.getRange('AD2').getValue();   //название Стоимость упаковки карго
      dataObj.cargoPackPrice = choseOfgoods.getRange(3 + i, 30).getValue();    //Стоимость упаковки карго
      dataObj.titleUnloadMscCost = choseOfgoods.getRange('AE2').getValue();    //название Стоимость разгрузки в Москве
      dataObj.unloadMscCost = choseOfgoods.getRange(3 + i, 31).getValue();     //Стоимость разгрузки в Москве
      dataObj.titleInsurPrice = choseOfgoods.getRange('AT2').getValue();       //название Страховка
      dataObj.insurPrice = choseOfgoods.getRange(3 + i, 46).getValue();        //Страховка
      dataObj.titleCommissionPrice = choseOfgoods.getRange('R2').getValue();   //назв стоимость единицы с учетом комиссии 5%
      dataObj.commissionPrice = choseOfgoods.getRange(3 + i, 18).getValue();   //стоимость единицы с учетом комиссии 5%
      dataObj.titleCargoPrice = choseOfgoods.getRange('O2').getValue();        //название цена доставки по Китаю за партию
      dataObj.cargoPrice = choseOfgoods.getRange(3 + i, 15).getValue();        //цена доставки по Китаю за партию  
      dataObj.titleCargoFastTotal = choseOfgoods.getRange('AL2').getValue();   //название ИТОГО (скорость перевозки - БЫСТРАЯ)
      dataObj.cargoFastTotal = choseOfgoods.getRange(3 + i, 38).getValue();    //ИТОГО (скорость перевозки - БЫСТРАЯ):
      dataObj.titleCargoSlowTotal = choseOfgoods.getRange('AM2').getValue();   //название ИТОГО (скор перевоз - МЕДЛЕННАЯ)
      dataObj.cargoSlowTotal = choseOfgoods.getRange(3 + i, 39).getValue();    //ИТОГО (скор перевоз - МЕДЛЕННАЯ) 
      dataObj.titlePurchaseCost = choseOfgoods.getRange('AK2').getValue();     //название СТОИМОСТЬ ЗАКУПКИ
      dataObj.purchaseCost = choseOfgoods.getRange(3 + i, 37).getValue();      //СТОИМОСТЬ ЗАКУПКИ
      dataObj.rateYuanCurrensy = choseOfgoods.getRange(3 + i, 36).getValue();  //курс юаня

      result.push(dataObj);

      sumPart += dataObj.purchaseCost          //Итоговая Стоимость Партии
      //Итоговая Стоимость Быстрой Перевозки:
      dataObj.cargoFastTotal === '#VALUE!'
        ? dataObj.cargoFastTotal = 0
        : sumFastCargo += dataObj.cargoFastTotal;
      //Итоговая Стоимость Медленной Перевозки:
      dataObj.cargoSlowTotal === '#VALUE!'
        ? dataObj.cargoSlowTotal = 0
        : sumSlowCargo += dataObj.cargoSlowTotal;
      //объём груза:
      dataObj.cargoVolume === '#VALUE!'
        ? dataObj.cargoVolume = 0
        : sumSlowCargo += dataObj.cargoVolume;
      //масса обрешетки:
      dataObj.protecSheathMass === '#VALUE!'
        ? dataObj.protecSheathMass = 0
        : sumSlowCargo += dataObj.protecSheathMass;
      //общая масса:
      dataObj.totalWeight === '#VALUE!'
        ? dataObj.totalWeight = 0
        : sumSlowCargo += dataObj.totalWeight;

      function ConcatMerge() {
        mergeCellForImg.merge();         //объединяем ячейки для изображения
        mergeForHeaderTech.merge();      //объединяем ячейки для заголовка Технические характеристики партии товара
        mergeCellForHeaderCost.merge();  //объединяем ячейки для заголовка Стоимость
        //объединяем ячейки для отображение характеристек:
        for (let i = 0; i < mergeCellsNumbers.length; i++) {
          const mergeCell = constructorList.getRange(mergeCellsNumbers[i] + itemIncrementForRec, 5, 1, 2);
          mergeCell.merge();
        }
      };
      ConcatMerge();
      increment++;
      //itemIncrementForRec += 18;

      if (i === 0 || i % 2 === 0) {
        itemIncrementForRec += 18;
      } else if (i === 1) {
        itemIncrementForRec += 33;
      } else {
        itemIncrementForRec += 46;
      }
    //}
  }

  let itemIncrementForRead = 0;

  for (let i = 0; i < result.length; i++) {
    constructorList.getRange(1, 3).setValue(result[i].rateYuanCurrensy);                  //устанавливаем курс юаня
    constructorList.getRange(14 + itemIncrementForRead, 1).setValue(result[i].increment);
    constructorList.getRange(14 + itemIncrementForRead, 5).setValue("кол-во");
    constructorList.getRange(14 + itemIncrementForRead, 6).setValue(result[i].quantity);
    constructorList.getRange(15 + itemIncrementForRead, 5).setValue("Технические характеристики партии товара");
    constructorList.getRange(16 + itemIncrementForRead, 5).setValue(result[i].titleOfCargoVolume);
    constructorList.getRange(16 + itemIncrementForRead, 7).setValue(result[i].cargoVolume);
    constructorList.getRange(17 + itemIncrementForRead, 5).setValue(result[i].titleOfCargoMass);
    constructorList.getRange(17 + itemIncrementForRead, 7).setValue(result[i].cargoMass);
    constructorList.getRange(18 + itemIncrementForRead, 5).setValue(result[i].titleOfProtecSheath);
    constructorList.getRange(18 + itemIncrementForRead, 7).setValue(result[i].protecSheathMass);
    constructorList.getRange(19 + itemIncrementForRead, 5).setValue(result[i].titleTotalWeight);
    constructorList.getRange(19 + itemIncrementForRead, 7).setValue(result[i].totalWeight);
    constructorList.getRange(20 + itemIncrementForRead, 5).setValue("Стоимость");
    constructorList.getRange(21 + itemIncrementForRead, 5).setValue(result[i].titleCargoFastPrice);
    constructorList.getRange(21 + itemIncrementForRead, 7).setValue(result[i].cargoFastPrice);
    constructorList.getRange(22 + itemIncrementForRead, 5).setValue(result[i].titleCargoSlowPrice);
    constructorList.getRange(22 + itemIncrementForRead, 7).setValue(result[i].cargoSlowPrice);
    constructorList.getRange(23 + itemIncrementForRead, 5).setValue(result[i].titleCargoPackPrice);
    constructorList.getRange(23 + itemIncrementForRead, 7).setValue(result[i].cargoPackPrice);
    constructorList.getRange(24 + itemIncrementForRead, 5).setValue(result[i].titleUnloadMscCost);
    constructorList.getRange(24 + itemIncrementForRead, 7).setValue(result[i].unloadMscCost);
    constructorList.getRange(25 + itemIncrementForRead, 5).setValue(result[i].titleInsurPrice);
    constructorList.getRange(25 + itemIncrementForRead, 7).setValue(result[i].insurPrice);
    constructorList.getRange(26 + itemIncrementForRead, 5).setValue(result[i].titleCommissionPrice);
    constructorList.getRange(26 + itemIncrementForRead, 7).setValue(result[i].commissionPrice);
    constructorList.getRange(27 + itemIncrementForRead, 5).setValue(result[i].titleCargoPrice);
    constructorList.getRange(27 + itemIncrementForRead, 7).setValue(result[i].cargoPrice);
    constructorList.getRange(28 + itemIncrementForRead, 5).setValue(result[i].titleCargoFastTotal);
    constructorList.getRange(28 + itemIncrementForRead, 7).setValue(result[i].cargoFastTotal);
    constructorList.getRange(29 + itemIncrementForRead, 5).setValue(result[i].titleCargoSlowTotal);
    constructorList.getRange(29 + itemIncrementForRead, 7).setValue(result[i].cargoSlowTotal);
    constructorList.getRange(30 + itemIncrementForRead, 5).setValue(result[i].titlePurchaseCost);
    constructorList.getRange(30 + itemIncrementForRead, 7).setValue(result[i].purchaseCost);

   // itemIncrementForRead += 18;
    
    if (i === 0 || i % 2 === 0) {
      itemIncrementForRead += 18;
    } else if (i === 1) {
      itemIncrementForRead += 33;
    } else {
      itemIncrementForRead += 46;
    }
  }
  const now = new Date();
  const currentDate = Utilities.formatDate(now, 'Russia/Moscow', 'dd.MM.YYYY')
  constructorList.getRange(2, 6).setValue(sumPart);       //запись Итоговая Стоимость Партии
  constructorList.getRange(3, 6).setValue(sumFastCargo);  //запись Итоговая Стоимость Быстрой Перевозки
  constructorList.getRange(4, 6).setValue(sumSlowCargo);  //запись Итоговая Стоимость Медленной Перевозки
  constructorList.getRange(5, 6).setValue(currentDate);
  constructorList.getRange(1, 2).setValue(`курс юаня на ${currentDate}`);

  constructorHeader();
  constructorBody();
  ss.setActiveSheet(commissionOfferList);
}


//

function constructorHeader() {

  constructorList.getRange('B1').copyTo(commissionOfferList.getRange('B1')); //копируем дату
  constructorList.getRange('C1').copyTo(commissionOfferList.getRange('C1')); //копируем фразу курс валюты Юань
  commissionOfferList.getRange('B3').setValue('ИТОГОВАЯ СТОИМОСТЬ ЗАКУПКИ');
  commissionOfferList.getRange('B4').setValue('(с учетом комиссии 5%)');
  constructorList.getRange('C1').copyTo(commissionOfferList.getRange('C1')); //копируем курс валюты Юань
  commissionOfferList.getRange('C4').setValue(constructorList.getRange('F2').getValue()) //ИТОГОВАЯ СТОИМОСТЬ ЗАКУПКИ
  commissionOfferList.getRange('B6').setValue('итоговая стоимость доставки (быстрая)');
  commissionOfferList.getRange('B7').setValue('итоговая стоимость доставки (медленная)');
  commissionOfferList.getRange('C6').setValue(constructorList.getRange('F3').getValue()) //ИТОГОВАЯ СТОИМОСТЬ БЫСТРАЯ
  commissionOfferList.getRange('C7').setValue(constructorList.getRange('F4').getValue()) //ИТОГОВАЯ СТОИМОСТЬ МЕДЛЕННАЯ
  
  const mergeCellForPhone = commissionOfferList.getRange(8, 5, 1, 3);

  function ConcatMerge() {
    mergeCellForPhone.merge();         //объединяем ячейки для номера телефона
  };

  ConcatMerge();

  commissionOfferList.getRange('E8').setValue('+7(951) 652-61-01');
}

//

function constructorBody() {
  const lr = constructorList.getLastRow();           //последняя строка конструктора

  constructorList.getRange(14, 1, 998, 7).copyTo(commissionOfferList.getRange(14, 1, 998, 7));
  commissionOfferList.getRange(lr + 1, 1, 1013 - lr, 7).clear();

  //commissionOfferList.getRange(lr + 1, 1, 1013 - lr).getRow(lr + 1).setMinimumHeight(21);

  footerConstructor.getRange(1, 1).copyTo(commissionOfferList.getRange(lr + 1, 1));

  const mergeCellForFooterG = commissionOfferList.getRange(lr + 1, 1, 6, 7);
  //const mergeCellForFooterV = commissionOfferList.getRange(lr + 1, 1, 6, 1);

  function ConcatMerge() {
    mergeCellForFooterG.merge();         //объединяем ячейки для футера
    //mergeCellForFooterV.merge();         //объединяем ячейки для футера
  };

  ConcatMerge();

}

//

function pdfAndExcelDownloader() {
  const headFolder = DriveApp.getFoldersByName("20230907 Создание КП Родионов");
  const arrFolders = DriveApp.getFoldersByName("20230907 Создание КП Родионов").next().getFolders()
  const foldersArr = headFolder.next().getFolders();
  const formattedDate = Utilities.formatDate(new Date(), "CET", "yyyy-MM-dd' 'HH:mm");
  const currentCustomerName = choseOfgoods.getRange('B1').getValue().toString().trim();

  function makeCopyPdf(folder) {

    const destination = DriveApp.getFolderById(folder);
    const blob = createblobpdf('КП', 'КП_' + currentCustomerName + '_' + formattedDate + '.pdf');
    destination.createFile(blob);
  }

  let folderNames = [];
  function getInfoFolders(array) {
    while (array.hasNext()) {
      folderNames.push(array.next().getName());
    }
    return folderNames;
  }
  let customerFolderId = [];
  function getIdsFolders(arr) {
    while (arr.hasNext()) {
      customerFolderId.push(arr.next().getId());
    }
  }

  getInfoFolders(foldersArr) //получаем имена папок
  getIdsFolders(arrFolders)  //получаем ID папок

  console.log(currentCustomerName);

  if (folderNames.includes(currentCustomerName)) {
    const indexId = folderNames.indexOf(currentCustomerName);

    makeCopyPdf(customerFolderId[indexId]);
    console.log('pdf +')
  }
  else {
    const newFolder = DriveApp.getFolderById('1nb_6bUY6EfPljfO57lp4vKo6huEMzT-H').createFolder(currentCustomerName).getId();
    console.log('нет такой папки...', newFolder)
    makeCopyPdf(newFolder);
  }
}

//

function generatePdf() {
  SpreadsheetApp.flush();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSpreadsheet = SpreadsheetApp.getActive(); // Get active spreadsheet.
  var sheets = sourceSpreadsheet.getSheets(); // Get active sheet.
  var sheetName = sourceSpreadsheet.getActiveSheet().getName();
  var sourceSheet = sourceSpreadsheet.getSheetByName(sheetName);
  var pdfName = sheetName + ".pdf"; // Set the output filename as SheetName.
  var parents = DriveApp.getFileById(sourceSpreadsheet.getId()).getParents(); // Get folder containing spreadsheet to save pdf in.
  if (parents.hasNext()) {
    var folder = parents.next();
  } else {
    folder = DriveApp.getRootFolder();
  }
  var theBlob = createblobpdf(sheetName, pdfName);
  var newFile = folder.createFile(theBlob);
  var email = Session.getActiveUser().getEmail() || 'admin@gmail.com';
  var custemail = sourceSheet.getRange('A1').getValue();
  email = email + "," + custemail;
  // Subject of email message
  const subject = `Your subject Attachement: ${sheetName}`;
  // Email Body can  be HTML too with your image
  const body = "body";
  if (MailApp.getRemainingDailyQuota() > 0)
    GmailApp.sendEmail(email, subject, body, {
      htmlBody: body,
      attachments: [theBlob]
    });
  // delete pdf if already exists
  var files = folder.getFilesByName(pdfName);
  while (files.hasNext()) {
    files.next().setTrashed(true);
  }
  sourceSpreadsheet.toast("Emailed to " + email, "Success");

}

function createblobpdf(sheetName, pdfName) {
  var sourceSpreadsheet = SpreadsheetApp.getActive();
  var sourceSheet = sourceSpreadsheet.getSheetByName(sheetName);
  var url = 'https://docs.google.com/spreadsheets/d/' + sourceSpreadsheet.getId() + '/export?exportFormat=pdf&format=pdf' // export as pdf / csv / xls / xlsx
    + '&size=A4' // paper size legal / letter / A4
    + '&portrait=true' // orientation, false for landscape
    + '&fitw=true' // fit to page width, false for actual size
    + '&sheetnames=true&printtitle=false' // hide optional headers and footers
    + '&pagenum=RIGHT&gridlines=false' // hide page numbers and gridlines
    + '&fzr=false' // do not repeat row headers (frozen rows) on each page
    + '&horizontal_alignment=CENTER' //LEFT/CENTER/RIGHT
    + '&vertical_alignment=TOP' //TOP/MIDDLE/BOTTOM
    + '&gid=' + sourceSheet.getSheetId(); // the sheet's Id
  var token = ScriptApp.getOAuthToken();
  // request export url
  var response = UrlFetchApp.fetch(url, {
    headers: {
      'Authorization': 'Bearer ' + token
    }
  });
  var theBlob = response.getBlob().setName(pdfName);
  return theBlob;
};



