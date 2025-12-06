// ========== КОНСТАНТЫ СИСТЕМЫ ==========
var SCHEDULE_SHEET = 'График_смен';
var PRODUCT_CATALOG_SHEET = 'Справочник_товаров';
var EMPLOYEE_DIRECTORY = 'Справочник_работников';
var COMPONENTS_CATALOG_SHEET = 'Справочник_комплектующих';
var CONSUMPTION_RATES_SHEET = 'Нормы_расхода';
var PURCHASES_SHEET = 'Закупки';
var PRODUCTION_WAREHOUSE_SHEET = 'Склад_Производства';
var MAIN_WAREHOUSE_SHEET = 'Склад_Главный';
var TRANSFERS_SHEET = 'Перемещения';
var PLANNING_SHEET = 'Планирование';
var MATERIALS_WRITEOFF_SHEET = 'Списание_материалов';
var PURCHASE_ID_PREFIX = 'PUR-';                          // Генерация id закупки
var TRANSFER_ID_PREFIX = 'TRF-';                          // Генерация id перемещения
var PLANNING_ID_PREFIX = 'PLN-';                          // Генерация id планирования
var TRANSFER_CHAIN_CACHE = {};                            // Утилита для работы с цепочками перемещений

var NEW_WAREHOUSE_STRUCTURE = {
  TRANSFER_ID: 0,     // A - ID перемещения
  PURCHASE_ID: 1,     // B - ID Закупки
  DATE: 2,            // C - Дата поступления
  COMPONENT: 3,       // D - Комплектующие
  CURRENT_STOCK: 4,   // E - Остаток
  STATUS: 5,          // F - Статус
};

// 🔥 НОВАЯ СТРУКТУРА ДЛЯ СПИСАНИЙ МАТЕРИАЛОВ
var NEW_WRITEOFF_STRUCTURE = {
  PLANNING_ID: 0,     // A - ID планирования
  TRANSFER_ID: 1,     // B - ID перемещения
  PURCHASE_ID: 2,     // C - ID Закупки
  DATE: 3,            // D - Дата
  DELIVERY: 4,        // E - Поставка
  PRODUCT: 5,         // F - Товар
  SHIFT_DATE: 6,      // G - Дата смены
  WORKER: 7,          // H - Работник
  ITEMS_COUNT: 8,     // I - Кол-во ед.
  COMPONENT: 9,       // J - Комплектующее
  WRITEOFF: 10        // K - Списание
};

// Глобальные переменные для защиты от повторного выполнения
var lastProcessedPurchase = {
  row: 0,
  timestamp: 0,
  status: ''
};

var lastProcessedPlanning = {
  row: 0,
  timestamp: 0,
  quantity: 0,
  product: ''
};

function onEdit(e){
  try{                                                    //Начало блока обработки ошибок
    var sheet = e.source.getActiveSheet();                //Получаем объект листа, который редактировали
    var range = e.range;                                  //Получаем объект диапазона (может быть одна ячейка или несколько)
    var sheetName = sheet.getName();                      //Получаем название листа
    var row = range.getRow();                             //Получаем номер строки, где произошло изменение
    var col = range.getColumn();                          //Получаем номер столбца, где произошло изменение

 // 🔵🔵🔵  Обработка листа "Планирование"  🔵🔵🔵
 if (sheet.getName() === PLANNING_SHEET) {
  if (range.getRow() < 2) return;
      
      var row = range.getRow();
      var col = range.getColumn();

      console.log('✏️ Изменение в Планировании: строка ' + row + ', столбец ' + col);

      // 🔥 НОВАЯ ФУНКЦИОНАЛЬНОСТЬ: ГЕНЕРАЦИЯ ID ПЛАНИРОВАНИЯ
      if ((col === 4 || col === 5 || col === 8) && row >= 2) {
        var idCell = sheet.getRange(row, 11); // K - ID планирования
        var currentId = idCell.getValue();
        
        if (!currentId && isPlanningRowFilled(sheet, row)) {
          var newId = generatePlanningId();
          idCell.setValue(newId);
          console.log('✅ Сгенерирован ID для планирования: ' + newId + ' в строке ' + row);
        }
      }

      // 🔥 НОВАЯ ФУНКЦИОНАЛЬНОСТЬ: УДАЛЕНИЕ ЗАПИСЕЙ СПИСАНИЙ ПРИ УДАЛЕНИИ СТРОКИ
      if (isPlanningRowDeleted(sheet, row)) {
        console.log('🗑️ Обнаружено удаление строки в Планировании: ' + row);
        
        // Получаем ID планирования перед удалением
        var planningId = sheet.getRange(row, 11).getValue(); // K - ID планирования
        
        if (planningId) {
          // Удаляем соответствующие записи списаний
          deleteWriteOffRecordsByPlanningId(planningId);
          console.log('✅ Записи списаний удалены для ID: ' + planningId);
        }
        
        Utilities.sleep(500);
        updateScheduleWriteOffs();
        // Также при удалении строки может измениться сумма по поставкам, поэтому проверяем лимит
        checkDeliveryLimit(sheet);
      }

      if (col === 4) { // Товар (D)
        updateDates(sheet, row);
      }

      if (col === 5) { // Дата смены (E)
        updateShifts(sheet, row);
      }

      if (col === 8) { // Кол-во коробок (H)
        // 🔥 ДОПОЛНИТЕЛЬНАЯ ПРОВЕРКА: ВСЯ ЛИ СТРОКА ЗАПОЛНЕНА
        if (!isPlanningRowFilled(sheet, row)) {
          console.log('🛑 Строка ' + row + ' заполнена не полностью, процесс не запускается');
          sheet.getRange(row, 9).setValue('❌ Заполните все поля'); // I - Статус
          sheet.getRange(row, 9).setBackground('#FFCCCC');
          return;
        }
        
        // ✅ ЗАЩИТА ОТ ПОВТОРНОГО ВЫПОЛНЕНИЯ ДЛЯ СПИСАНИЙ
        var now = new Date().getTime();
        var currentQuantity = sheet.getRange(row, 8).getValue() || 0;
        var currentProduct = sheet.getRange(row, 4).getValue() || '';
        var currentStatus = sheet.getRange(row, 9).getValue() || '';
        
        if (lastProcessedPlanning.row === row && 
            lastProcessedPlanning.product === currentProduct &&
            lastProcessedPlanning.quantity === currentQuantity &&
            (now - lastProcessedPlanning.timestamp) < 3000) {
          console.log('🛑 Защита: пропускаем повторное выполнение для строки ' + row);
          return;
        }
        
        // Обновляем информацию о последней обработанной строке планирования
        lastProcessedPlanning.row = row;
        lastProcessedPlanning.timestamp = now;
        lastProcessedPlanning.quantity = currentQuantity;
        lastProcessedPlanning.product = currentProduct;
        
        updateStatus(sheet, row);
        updateScheduleWriteOffs();
        checkDeliveryLimit(sheet);
        // ✅ РАСХОД МАТЕРИАЛОВ при выполнении поставки
        updateMaterialsConsumption(sheet, row);
      }

      if (col === 2) { // Поставки (B) - добавляем обработку для столбца B
        checkDeliveryLimit(sheet);
      }

      // ПРОВЕРЯЕМ УДАЛЕНИЕ СТРОКИ ПРИ ИЗМЕНЕНИИ ЛЮБОГО СТОЛБЦА
      if (isPlanningRowDeleted(sheet, row)) {
        console.log('🗑️ Обнаружено удаление строки в Планировании: ' + row);
        Utilities.sleep(500);
        updateScheduleWriteOffs();
        // Также при удалении строки может измениться сумма по поставкам, поэтому проверяем лимит
        checkDeliveryLimit(sheet);
      }

 }

 // 🔵🔵🔵  Обработка листа "График_смен" 🔵🔵🔵
 if (sheet.getName() === SCHEDULE_SHEET){
  if (range.getRow() < 2) return;

 // ЕСЛИ ИЗМЕНЯЕМ ВЫРАБОТКУ (СТОЛБЕЦ F) - ПЕРЕСЧИТЫВАЕМ ОСТАТОК
 if (col === 6){                                          // F - Выработка
   console.log('Изменена выработка в строке ' + row + ', пересчитываем остаток...');
   updateBalanceForRow(sheet, row);
 }
 // Если редактируем разъединение или дату переноса
      if (col === 10 || col === 11) { // J, K - Разъединить, Перенос с/на дату
        processCompleteTransfer(sheet, row);
      }
 // Если очищаются столбцы I, J, K, L
      if (col >= 9 && col <= 12) {
        handleClearConnectionWithParsing(sheet, row, col);
      }
      
      // ДОПОЛНИТЕЛЬНО: Если удаляется вся строка (проверяем по нескольким пустым ячейкам)
      if (isRowBeingDeleted(sheet, row)) {
        console.log('🗑️ Обнаружено удаление строки: ' + row);
        handleRowDeletion(sheet, row);
      }
 }

 // 🔵🔵🔵  Обработка листа "Закупки" 🔵🔵🔵

if (sheetName === PURCHASES_SHEET) {
  if (row < 2) return;

  console.log('✏️ Изменение в Закупках: строка ' + row + ', столбец ' + col);

  // 🔄 АВТОМАТИЧЕСКАЯ ГЕНЕРАЦИЯ ID ДЛЯ НОВЫХ ЗАКУПОК
  if ((col === 2 || col === 3 || col === 4) && row >= 2) {
    var idCell = sheet.getRange(row, 1);
    var currentId = idCell.getValue();
    
    if (!currentId && isPurchaseRowFilled(sheet, row)) {
      var newId = generatePurchaseId();
      idCell.setValue(newId);
      console.log('✅ Сгенерирован ID для новой закупки: ' + newId + ' в строке ' + row);
    }
  }

  // 🔢 АВТОМАТИЧЕСКИЙ РАСЧЕТ СУММЫ
  if (col === 4 || col === 5) {
    console.log('🔄 Расчет суммы для строки ' + row);
    autoCalculateSum(sheet, row);
  }

  // 📦 ОБРАБОТКА СТАТУСА "ПОЛУЧЕНО"
  if (col === 9 && range.getValue() === 'Получено') {
    // ✅ ЗАЩИТА ОТ ПОВТОРНОГО ВЫПОЛНЕНИЯ
    var now = new Date().getTime();
    if (lastProcessedPurchase.row === row && 
        lastProcessedPurchase.status === 'Получено' &&
        (now - lastProcessedPurchase.timestamp) < 5000) {
      console.log('🛑 Защита: пропускаем повторное выполнение для строки ' + row);
      return;
    }
    
    // Обновляем информацию о последней обработанной закупке
    lastProcessedPurchase.row = row;
    lastProcessedPurchase.timestamp = now;
    lastProcessedPurchase.status = 'Получено';
    
    console.log('✅ Обрабатываем закупку в строке ' + row);
    updateWarehouseFromPurchase(sheet, row);
  }
}

  }catch (error) {                                        //Если что-то пойдет не так, код перейдет в catch
  console.error('Ошибка в onEdit: ' + error.toString());
}
}

/**
 * 🔄 МЕНЮ ДЛЯ ДОБАВЛЕНИЯ НОВЫХ ФУНКЦИЙ
 */
function onOpen(){
  try {
    console.log('🔄 Запуск функции onOpen...');

    var ui = SpreadsheetApp.getUi();

    // Создаем основное меню
    var menu = ui.createMenu('🔄 Меню')

    // 📋 Планирование
    var planningMenu = ui.createMenu('📋 Планирование');
    planningMenu.addItem('🗑️ Удалить строку планирования', 'removePlanningRowDirect');
    menu.addSubMenu(planningMenu);

    // 📦 Управление складами
    var warehouseMenu = ui.createMenu('📦 Управление складами');
    warehouseMenu.addItem('📋 Создать перемещение', 'showTransferDialog');
    warehouseMenu.addItem('✅ Выполнить перемещение', 'showPendingTransfers');
    warehouseMenu.addSeparator();
    warehouseMenu.addItem('👁️ Показать скрытые строки', 'showAllHiddenRowsAllWarehouses');
    warehouseMenu.addItem('🙈 Скрыть строки с нулевым остатком', 'updateStatusAndHideZeroStockAllWarehouses');
    warehouseMenu.addSeparator();
    warehouseMenu.addItem('🔄 Безопасный пересчет остатков', 'safeRecalculateWarehouseBalances');
    warehouseMenu.addSeparator();
    warehouseMenu.addItem('🔄 Обновить все статусы складов', 'updateAllWarehouseStatuses');
    warehouseMenu.addItem('📊 Проверить статусы (отчет)', 'checkWarehouseStatusesReport');
    warehouseMenu.addItem('🔧 Диагностика статусов складов', 'diagnoseStatusProblem');
    warehouseMenu.addItem('💡 Исправить подсказки статусов', 'forceUpdateStatusesAfterSorting');
    warehouseMenu.addSeparator();
    warehouseMenu.addItem('🔍 Проверить остатки', 'showMultiWarehouseMaterialsCheck');
    warehouseMenu.addSeparator();
    
    // 🔥 ДОБАВЛЯЕМ НОВЫЕ ФУНКЦИИ ДЛЯ ЦЕПОЧЕК ПЕРЕМЕЩЕНИЙ
    var transferChainMenu = ui.createMenu('🔗 Цепочки перемещений');
    transferChainMenu.addItem('🔍 Проверить цепочку перемещений', 'showTransferChainDialog');
    transferChainMenu.addItem('🔄 Восстановить связи перемещений', 'restoreTransferLinks');
    transferChainMenu.addItem('📊 Сгенерировать отчет по цепочке', 'showTransferReportDialog');
    warehouseMenu.addSubMenu(transferChainMenu);
    
    warehouseMenu.addItem('📋 Обновить минимальные запасы', 'updateMinStocksFromCatalogSmart');
    menu.addSubMenu(warehouseMenu);

    // 📥 Закупки
    var purchasesMenu = ui.createMenu('📥 Закупки');
    purchasesMenu.addItem('🗑️ Удалить строку закупки', 'removePurchaseRowSmart');
    menu.addSubMenu(purchasesMenu);

    // Добавляем меню
    menu.addToUi();
    
    console.log('✅ Меню успешно создано');
  }catch (error) {
    console.error('❌ Критическая ошибка в onOpen: ' + error.toString());
  }
}

// ================= Функции  ===================
/**
 *  🔍 ДОПОЛНИТЕЛЬНАЯ ОПТИМИЗАЦИЯ: ПЕРЕСЧЕТ ОСТАТКА С ПРОВЕРКОЙ ИЗМЕНЕНИЙ
 */
function updateBalanceForRow(sheet, row) {
  try {
    console.log('🔢 Пересчет остатка для строки ' + row);

    // Прямое чтение значений из ячеек
    var production = sheet.getRange(row, 6).getValue() || 0;    // F - Выработка
    var writeoff = sheet.getRange(row, 8).getValue() || 0;      // H - Списано (из планирования)
    var connect = sheet.getRange(row, 9).getValue() || 0;       // I - Соединить (ручное)
    var disconnect = sheet.getRange(row, 10).getValue() || 0;   // J - Разъединить (ручное)
    
    // Получаем текущий остаток для проверки изменений
    var currentBalance = sheet.getRange(row, 7).getValue() || 0;

    // ФОРМУЛА вычисления: Остаток = Выработка - Списано + Соединить - Разъединить
    var newBalance = Number(production) - Number(writeoff) + Number(connect) - Number(disconnect);

    // 🔥 ОБНОВЛЯЕМ ТОЛЬКО ЕСЛИ ЗНАЧЕНИЕ ИЗМЕНИЛОСЬ
    if (Math.abs(currentBalance - newBalance) > 0.001) {
      sheet.getRange(row, 7).setValue(newBalance);
      console.log('✅ Остаток обновлен для строки ' + row + ': ' + newBalance + 
                 ' (Выработка: ' + production + ', Списано: ' + writeoff + 
                 ', Соединить: ' + connect + ', Разъединить: ' + disconnect + ')');
    } else {
      console.log('ℹ️ Остаток не изменился для строки ' + row + ': ' + newBalance);
    }
    
    return newBalance;

  } catch(error){
    console.error('❌ Ошибка в updateBalanceForRow для строки ' + row + ': ' + error.toString());
    return 0;
  }
}

/**
 * ✅ БРАБОТКА ПЕРЕНОСА
 */
function processCompleteTransfer(sheet, row) {
  try {
    // Прямое чтение значений из ячеек
    var disconnect = sheet.getRange(row, 10).getValue() || 0; // J - Разъединить  
    var transferDate = sheet.getRange(row, 11).getValue();   // K - Перенос с/на дату
    
    console.log('🔄 Обработка переноса: строка ' + row);
    console.log('Разъединить: ' + disconnect + ', Дата: ' + transferDate);
    
    // Если указан разъединение и дата переноса
    if (disconnect > 0 && transferDate) {
      processCompleteDisconnection(sheet, row, disconnect, transferDate);
    }
    
    // ВСЕГДА пересчитываем остаток для этой строки
    updateBalanceForRow(sheet, row);
    
  } catch (error) {
    console.error('❌ Ошибка в processCompleteTransfer: ' + error.toString());
  }
}

/**
 * ✅ ПОЛНАЯ ОБРАБОТКА РАЗЪЕДИНЕНИЯ
 */
function processCompleteDisconnection(sheet, row, disconnectAmount, transferDate) {
  try {
    // Прямое чтение значений из ячеек
    var sourceDate = sheet.getRange(row, 2).getValue();    // B - Дата
    var sourceWorker = sheet.getRange(row, 3).getValue();  // C - Работник
    var sourceProduct = sheet.getRange(row, 4).getValue(); // D - Товар
    var sourceMP = sheet.getRange(row, 5).getValue() || ''; // E - МП (маркетплейс,WB, OZ, MM)

    console.log('🔁 Разъединение: ' + sourceWorker + ', ' + sourceProduct + ', МП: ' + sourceMP);
    
    // ПРОВЕРКА: НЕЛЬЗЯ ПЕРЕНОСИТЬ НА ТУ ЖЕ ДАТУ
    if (isSameDate(sourceDate, transferDate)) {
      console.log('❌ Ошибка: попытка переноса на ту же дату');
      SpreadsheetApp.getUi().alert('❌ Ошибка: Нельзя переносить на ту же дату!');
      
      // Очищаем поля переноса
      sheet.getRange(row, 10).setValue(''); // J - Разъединить
      sheet.getRange(row, 11).setValue(''); // K - Перенос с/на дату
      
      // Пересчитываем остаток
      updateBalanceForRow(sheet, row);
      return;
    }

    // Ищем ЛЮБУЮ строку с указанной датой и товаром
    var targetRow = findAnyTargetRow(transferDate, sourceProduct);
    
    if (targetRow) {
      // Получаем данные целевой строки
      var targetMP = sheet.getRange(targetRow, 5).getValue() || '';
      var targetWorker = sheet.getRange(targetRow, 3).getValue() || '';
      
      // ОБНОВЛЯЕМ ИСХОДНУЮ СТРОКУ
      updateSourceRow(sheet, row, disconnectAmount, transferDate, targetWorker, targetMP);
      
      // ОБНОВЛЯЕМ ЦЕЛЕВУЮ СТРОКУ
      updateTargetRow(sheet, targetRow, disconnectAmount, sourceDate, sourceWorker, sourceMP);
      
      // ПЕРЕСЧИТЫВАЕМ ОСТАТКИ ДЛЯ ОБЕИХ СТРОК
      updateBalanceForRow(sheet, row);
      updateBalanceForRow(sheet, targetRow);
      
      console.log('✅ Перенос завершен: ' + disconnectAmount + ' кор. от ' + sourceWorker + 
                 ' к ' + targetWorker);
    } else {
      console.log('❌ Не найдена смена с товаром "' + sourceProduct + '" на дату ' + transferDate);
      SpreadsheetApp.getUi().alert('❌ Не найдена смена с товаром "' + sourceProduct + '" на дату ' + 
                                  formatDateShort(transferDate));

      // Очищаем поля переноса
      sheet.getRange(row, 10).setValue(''); // J - Разъединить
      sheet.getRange(row, 11).setValue(''); // K - Перенос с/на дату
      
      // Пересчитываем остаток
      updateBalanceForRow(sheet, row);
      return;
    }
    
  } catch (error) {
    console.error('❌ Ошибка в processCompleteDisconnection: ' + error.toString());
  }
}

/**
 * ✅ СРАВНЕНИЕ ДАТ В ФОРМАТЕ "dd.MM.yy"
 */
function isSameDate(date1, date2) {
  try {
    var formatDate1 = Utilities.formatDate(date1, Session.getScriptTimeZone(), "dd.MM.yy");
    var formatDate2 = Utilities.formatDate(date2, Session.getScriptTimeZone(), "dd.MM.yy");
    return formatDate1 === formatDate2;
  } catch (error) {
    return false;
  }
}

/**
 * ✅ ПОИСК ЛЮБОЙ ЦЕЛЕВОЙ СТРОКИ
 */
function findAnyTargetRow(transferDate, product) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('График_смен');
  var data = sheet.getRange('B2:D' + sheet.getLastRow()).getValues(); // B:D - Дата, Работник, Товар
  
  for (var i = 0; i < data.length; i++) {
    var rowDate = data[i][0];
    var rowProduct = data[i][2];
    
    if (rowDate && rowProduct && 
        isSameDate(rowDate, transferDate) &&
        rowProduct === product) {
      
      console.log('Найдена целевая строка: ' + (i+2) + ', работник: ' + data[i][1]);
      return i + 2;
    }
  }
  
  return null;
}

/**
 * ✅ ФОРМАТИРОВАНИЕ ДАТЫ В ФОРМАТ "dd.MM.yy"
 */
function formatDateShort(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "dd.MM.yy");
}

/**
 * ✅ ОБНОВЛЕНИЕ ИСХОДНОЙ СТРОКИ (ОТКУДА ЗАБИРАЕМ)
 */
function updateSourceRow(sheet, row, amount, transferDate, targetWorker, targetMP) {
  var sourceData = sheet.getRange(row, 1, 1, 12).getValues()[0];
  var sourceMP = sourceData[4] || '';
  
  // Записываем количество в столбец Разъединить (J)
  sheet.getRange(row, 10).setValue(amount);
  
  // Записываем дату в столбец Перенос с/на дату (K)
  sheet.getRange(row, 11).setValue(transferDate);
  
  // Формируем комментарий для столбца L
  var comment = 'Перенос ' + amount + ' кор. к ' + targetWorker + ' на ' + 
               formatDateShort(transferDate);
  
  // Добавляем примечание о переклейке если МП разные
  if (sourceMP && targetMP && sourceMP !== targetMP) {
    comment += '; Переклеить ШК ' + sourceMP + '>' + targetMP;
  }
  
  // Добавляем комментарий в столбец L
  addNoteToRow(row, comment);
  
  console.log('✅ Исходная строка обновлена: ' + comment);
}

/**
 * ✅ ДОБАВЛЕНИЕ ПРИМЕЧАНИЯ К СТРОКЕ
 */
function addNoteToRow(row, note) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('График_смен');
  var currentNote = sheet.getRange(row, 12).getValue() || ''; // L - Примечание
  
  if (currentNote) {
    // Если примечание уже есть, добавляем новое через точку с запятой
    if (!currentNote.includes(note)) {
      sheet.getRange(row, 12).setValue(currentNote + '; ' + note);
    }
  } else {
    // Если примечания нет, создаем новое
    sheet.getRange(row, 12).setValue(note);
  }
}

/**
 * ✅ ОБНОВЛЕНИЕ ЦЕЛЕВОЙ СТРОКИ (КУДА ДОБАВЛЯЕМ)
 */
function updateTargetRow(sheet, targetRow, amount, sourceDate, sourceWorker, sourceMP) {
  var targetData = sheet.getRange(targetRow, 1, 1, 12).getValues()[0];
  var targetMP = targetData[4] || '';
  
  // Записываем количество в столбец Соединить (I)
  var currentConnect = sheet.getRange(targetRow, 9).getValue() || 0;
  var newConnect = currentConnect + amount;
  sheet.getRange(targetRow, 9).setValue(newConnect);
  
  // Записываем дату откуда перенесли в столбец Перенос с/на дату (K)
  sheet.getRange(targetRow, 11).setValue(sourceDate);
  
  // Формируем комментарий для столбец L
  var comment = 'Соединение ' + amount + ' кор. от ' + sourceWorker + ' с ' + 
               formatDateShort(sourceDate);
  
  // Добавляем примечание о переклейке если МП разные
  if (sourceMP && targetMP && sourceMP !== targetMP) {
    comment += '; Переклеить ШК ' + sourceMP + '>' + targetMP;
  }
  
  // Добавляем комментарий в столбец L
  addNoteToRow(targetRow, comment);
  
  console.log('✅ Целевая строка обновлена: ' + comment);
}

/**
 * ✅ ФИНАЛЬНАЯ ВЕРСИЯ: ОЧИСТКА С УДАЛЕНИЕМ КОММЕНТАРИЯ
 */
function handleClearConnectionWithParsing(sheet, row, col) {
  try {
    var value = sheet.getRange(row, col).getValue();
    
    if (value === '' || value === 0) {
      console.log('🧹 Очистка значения в строке ' + row + ', столбец ' + col);
      
      // 1. ПАРСИМ КОММЕНТАРИЙ ПЕРЕД ОЧИСТКОЙ
      var comment = sheet.getRange(row, 12).getValue();
      console.log('💬 Комментарий для парсинга: "' + comment + '"');
      
      var parsedData = parseComment(comment);
      
      // 2. ОЧИЩАЕМ ТЕКУЩУЮ СТРОКУ (включая комментарий)
      sheet.getRange(row, 12).setValue(''); // L - Комментарий
      console.log('🗑️ Очищен комментарий в текущей строке ' + row);
      
      if (col === 9) {
        sheet.getRange(row, 11).setValue(''); // K - Перенос с/на дату
      } else if (col === 10) {
        sheet.getRange(row, 11).setValue(''); // K - Перенос с/на дату
      }
      
      // 3. ЕСЛИ ЕСТЬ ДАТЫ - ИЩЕМ И ОЧИЩАЕМ СВЯЗАННЫЕ СТРОКИ
      if (parsedData && parsedData.dates.length > 0) {
        console.log('📅 Найдены даты для поиска: ' + parsedData.dates.join(', '));
        
        parsedData.dates.forEach(function(date) {
          console.log('🔎 Ищем строки с датой: ' + date);
          var relatedRows = findRowsByDate(sheet, date);
          
          if (relatedRows.length > 0) {
            relatedRows.forEach(function(relatedRow) {
              if (relatedRow.row !== row) {
                console.log('🔄 Очищаем связанную строку: ' + relatedRow.row);
                clearConnectionRow(sheet, relatedRow.row);
                console.log('✅ Очищена связанная строка: ' + relatedRow.row);
              }
            });
          } else {
            console.log('ℹ️ Связанные строки не найдены для даты: ' + date);
          }
        });
      } else {
        console.log('ℹ️ Даты в комментарии не найдены');
      }
      
      // 4. ПЕРЕСЧИТЫВАЕМ БАЛАНС
      updateBalanceForRow(sheet, row);
      console.log('✅ Полная очистка завершена для строки ' + row);
    }
    
  } catch (error) {
    console.error('❌ Ошибка: ' + error.toString());
  }
}

/**
 * ✅ ОБНОВЛЕННЫЙ ПАРСИНГ КОММЕНТАРИЯ
 */
function parseComment(comment) {
  if (!comment || comment === '') {
    console.log('💬 Комментарий пустой');
    return null;
  }
  
  try {
    console.log('🔧 Парсим комментарий: "' + comment + '"');
    
    // ПАТТЕРН ДЛЯ ДАТ В ФОРМАТЕ DD.MM.YY
    var datePattern = /\b(\d{1,2}\.\d{1,2}\.\d{2})\b/g;
    var dates = [];
    var match;
    
    while ((match = datePattern.exec(comment)) !== null) {
      var rawDate = match[1];
      console.log('📅 Найдена дата: "' + rawDate + '"');
      dates.push(rawDate);
    }
    
    // Ищем количество
    var quantityPattern = /(\d+\.?\d*)\s*кор?/i;
    var quantityMatch = comment.match(quantityPattern);
    var quantity = quantityMatch ? parseFloat(quantityMatch[1]) : null;
    
    // Ищем имя
    var namePattern = /(?:от|к)\s+([А-Я][а-я]+\s+[А-Я]\.?)/i;
    var nameMatch = comment.match(namePattern);
    var name = nameMatch ? nameMatch[1] : null;
    
    var result = {
      dates: dates,
      quantity: quantity,
      name: name,
      originalComment: comment
    };
    
    console.log('📊 Результат парсинга: ' + JSON.stringify(result));
    return result;
    
  } catch (error) {
    console.error('❌ Ошибка парсинга комментария: ' + error);
    return null;
  }
}


/**
 * ✅ УЛУЧШЕННЫЙ ПОИСК С ПРАВИЛЬНЫМ ФОРМАТИРОВАНИЕМ ДАТ
 */

function findRowsByDate(sheet, targetDate) {
  var foundRows = [];
  var lastRow = sheet.getLastRow();
  
  if (lastRow < 2) return foundRows;
  
  console.log('🔍 Быстрый поиск даты "' + targetDate + '"');
  
  // ЗАГРУЖАЕМ ВСЕ ДАННЫЕ ОДНИМ ЗАПРОСОМ
  var range = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
  var allData = range.getValues();
  
  // Преобразуем целевую дату один раз
  var targetDateObj = parseTargetDate(targetDate);
  
  for (var i = 0; i < allData.length; i++) {
    var row = i + 2;
    var dateCell = allData[i][1]; // Столбец B (индекс 1)
    
    if (dateCell && isDateMatchFast(dateCell, targetDateObj)) {
      foundRows.push({
        row: row,
        column: 2,
        value: dateCell,
        formatted: Utilities.formatDate(dateCell, Session.getScriptTimeZone(), 'dd.MM.yy')
      });
    }
  }
  
  console.log('📋 Найдено строк: ' + foundRows.length);
  return foundRows;
}

// БЫСТРАЯ ПРОВЕРКА ДАТ
function isDateMatchFast(cellValue, targetDateObj) {
  if (!(cellValue instanceof Date)) return false;
  
  try {
    var cellDate = new Date(cellValue);
    return cellDate.getDate() === targetDateObj.day &&
           cellDate.getMonth() === targetDateObj.month &&
           cellDate.getFullYear() === targetDateObj.year;
  } catch (e) {
    return false;
  }
}

// ПРЕОБРАЗОВАНИЕ ЦЕЛЕВОЙ ДАТЫ ОДИН РАЗ
function parseTargetDate(dateString) {
  var parts = dateString.split('.');
  return {
    day: parseInt(parts[0]),
    month: parseInt(parts[1]) - 1, // Месяцы в JS: 0-11
    year: parseInt(parts[2]) + (parts[2].length === 2 ? 2000 : 0)
  };
}

/**
 * ✅ ПРАВИЛЬНОЕ СРАВНЕНИЕ ОБЪЕКТОВ DATE С СТРОКОЙ ДАТЫ
 */
function isDateMatch(cellDate, targetDate) {
  try {
    // Если cellDate - это объект Date (из столбца B)
    if (cellDate instanceof Date) {
      var cellDateFormatted = Utilities.formatDate(cellDate, Session.getScriptTimeZone(), 'dd.MM.yy');
      console.log('   🔄 Преобразовано: "' + cellDate + '" → "' + cellDateFormatted + '"');
      
      if (cellDateFormatted === targetDate) {
        console.log('   ✅ Совпадение после преобразования: "' + cellDateFormatted + '" = "' + targetDate + '"');
        return true;
      }
    }
    
    // Если cellDate - это строка (из других мест)
    var cellDateStr = cellDate.toString().trim();
    
    // Прямое сравнение строк
    if (cellDateStr === targetDate) {
      console.log('   ✅ Прямое совпадение: "' + cellDateStr + '" = "' + targetDate + '"');
      return true;
    }
    
    // Пробуем преобразовать обе даты к одному формату
    var date1 = tryParseDate(cellDateStr);
    var date2 = tryParseDate(targetDate);
    
    if (date1 && date2) {
      var date1Formatted = Utilities.formatDate(date1, Session.getScriptTimeZone(), 'dd.MM.yy');
      var date2Formatted = Utilities.formatDate(date2, Session.getScriptTimeZone(), 'dd.MM.yy');
      
      if (date1Formatted === date2Formatted) {
        console.log('   ✅ Совпадение после нормализации: "' + date1Formatted + '" = "' + date2Formatted + '"');
        return true;
      }
    }
    
    console.log('   ❌ Не совпадает: "' + cellDateStr + '" ≠ "' + targetDate + '"');
    return false;
    
  } catch (error) {
    console.log('   ❌ Ошибка сравнения: ' + error);
    return false;
  }
}

/**
 * ✅ ПАРСИНГ ДАТЫ ИЗ РАЗНЫХ ФОРМАТОВ
 */
function tryParseDate(dateStr) {
  try {
    // Для формата DD.MM.YY (12.11.25)
    if (dateStr.match(/^\d{1,2}\.\d{1,2}\.\d{2}$/)) {
      var parts = dateStr.split('.');
      var day = parseInt(parts[0]);
      var month = parseInt(parts[1]) - 1; // месяцы 0-11
      var year = 2000 + parseInt(parts[2]); // 25 → 2025
      return new Date(year, month, day);
    }
    
    // Для формата DD.MM.YYYY (12.11.2025)
    if (dateStr.match(/^\d{1,2}\.\d{1,2}\.\d{4}$/)) {
      var parts = dateStr.split('.');
      var day = parseInt(parts[0]);
      var month = parseInt(parts[1]) - 1;
      var year = parseInt(parts[2]);
      return new Date(year, month, day);
    }
    
    // Пробуем стандартный парсинг
    return new Date(dateStr);
    
  } catch (error) {
    return null;
  }
}

/**
 * ✅ ПОЛНАЯ ОЧИСТКА СТРОКИ СОЕДИНЕНИЯ
 */
function clearConnectionRow(sheet, row) {
  try {
    // Очищаем столбцы I, J, K, L
    var columnsToClear = [9, 10, 11, 12]; // I, J, K, L
    
    columnsToClear.forEach(function(col) {
      sheet.getRange(row, col).setValue('');
    });
    
    // Пересчитываем баланс
    updateBalanceForRow(sheet, row);
    
    console.log('🧹 Очищена строка ' + row + ' (столбцы I,J,K,L)');
    
  } catch (error) {
    console.error('❌ Ошибка очистки строки ' + row + ': ' + error);
  }
}

/**
 * ✅ ПРОВЕРКА - УДАЛЯЕТСЯ ЛИ СТРОКА
 */
function isRowBeingDeleted(sheet, row) {
  // Проверяем несколько ключевых ячеек в строке
  var keyCells = sheet.getRange(row, 1, 1, 8).getValues()[0]; // A-H
  var emptyCount = keyCells.filter(function(cell) {
    return cell === '' || cell === null || cell === undefined;
  }).length;
  
  // Если большинство ключевых ячеек пустые - строка удаляется
  return emptyCount >= 6; // Настроить порог по необходимости
}

/**
 * ✅ ОБРАБОТКА УДАЛЕНИЯ СТРОКИ
 */
function handleRowDeletion(sheet, row) {
  var comment = sheet.getRange(row, 12).getValue(); // L - Комментарий
  var parsedData = parseComment(comment);
  
  if (parsedData && parsedData.dates.length > 0) {
    // Ищем и очищаем связанные строки
    parsedData.dates.forEach(function(date) {
      var relatedRows = findRowsByDate(sheet, date);
      relatedRows.forEach(function(relatedRow) {
        if (relatedRow.row !== row) {
          clearConnectionRow(sheet, relatedRow.row);
        }
      });
    });
  }
}

/**
 * ✅ ПРОВЕРКА ЗАПОЛНЕННОСТИ СТРОКИ ЗАКУПКИ
 */
function isPurchaseRowFilled(sheet, row) {
  try {
    var date = sheet.getRange(row, 2).getValue();        // B - Дата
    var component = sheet.getRange(row, 3).getValue();   // C - Комплектующие
    var quantity = sheet.getRange(row, 4).getValue();    // D - Количество
    
    // Считаем строку заполненной если есть дата И комплектующие И количество
    return !!(date && component && quantity > 0);
  } catch (error) {
    console.error('Ошибка в isPurchaseRowFilled: ' + error.toString());
    return false;
  }
}

/**
 * ✅ АВТОМАТИЧЕСКИЙ РАСЧЕТ СУММЫ В ЗАКУПКАХ (ИСПРАВЛЕННАЯ)
 */
function autoCalculateSum(sheet, row) {
  try {
    var quantity = sheet.getRange(row, 4).getValue(); // D - Количество (было 3)
    var price = sheet.getRange(row, 5).getValue();    // E - Цена (было 4)
    
    console.log('🔢 Расчет суммы: количество=' + quantity + ', цена=' + price);
    
    if (quantity > 0 && price > 0) {
      var sum = quantity * price;
      sheet.getRange(row, 6).setValue(sum); // F - Сумма (было 5)
      console.log('✅ Сумма рассчитана: ' + quantity + ' × ' + price + ' = ' + sum);
    } else {
      // Если данные неполные - очищаем сумму
      sheet.getRange(row, 6).setValue('');
      console.log('🔄 Сумма очищена - неполные данные');
    }
  } catch (error) {
    console.error('❌ Ошибка в autoCalculateSum: ' + error.toString());
  }
}

/**
 *  🔥 ОБНОВЛЕННАЯ ФУНКЦИЯ: ОБНОВЛЕНИЕ СКЛАДА ОТ ЗАКУПОК С СОЗДАНИЕМ НОВЫХ ЗАПИСЕЙ
 */
function updateWarehouseFromPurchase(sheet, row) {
  try {
    var status = sheet.getRange(row, 9).getValue(); // I - Статус
    var purchaseId = sheet.getRange(row, 1).getValue(); // A - ID закупки
    var component = sheet.getRange(row, 3).getValue(); // C - Комплектующие
    var quantity = sheet.getRange(row, 4).getValue(); // D - Количество
    var targetWarehouse = sheet.getRange(row, 8).getValue(); // H - Склад назначения
    
    // Добавляем материалы только если статус "Получено"
    if (status === 'Получено' && component && quantity > 0 && purchaseId) {
      
      if (!targetWarehouse) {
        SpreadsheetApp.getUi().alert('❌ Ошибка', 'Не выбран склад назначения', SpreadsheetApp.getUi().ButtonSet.OK);
        return;
      }
      
      console.log('📥 Приход материалов на ' + targetWarehouse + ': ' + component + ', ' + quantity + ' ед., ID: ' + purchaseId);
      
      // 🔥 ГЕНЕРИРУЕМ ID ПЕРЕМЕЩЕНИЯ ДЛЯ ЗАКУПКИ
      var transferId = generateTransferId();
      
      //  ПЕРЕДАЕМ УПРАВЛЕНИЕ В ФУНКЦИЮ СОЗДАНИЯ ПЕРЕМЕЩЕНИЯ
      // которая создаст новую запись на складе
      createAutoTransferFromPurchase(purchaseId, component, quantity, targetWarehouse, transferId);
      
      SpreadsheetApp.getUi().alert('✅ Материалы получены', 'ID: ' + purchaseId + '\n' + component + ' × ' + quantity + ' ед.\nСклад: ' + targetWarehouse + '\nID перемещения: ' + transferId, SpreadsheetApp.getUi().ButtonSet.OK);
    }
    
  } catch (error) {
    console.error('Ошибка в updateWarehouseFromPurchase: ' + error.toString());
    SpreadsheetApp.getUi().alert('❌ Ошибка обработки закупки', error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 *  🔥 ОБНОВЛЕННАЯ ФУНКЦИЯ: ДОБАВЛЕНИЕ МАТЕРИАЛА НА ГЛАВНЫЙ СКЛАД
 * Теперь всегда создает новую запись, как при перемещениях
 */
function addMaterialToMainWarehouse(component, purchaseId, quantity, transferId) {
  try {
    var mainSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MAIN_WAREHOUSE_SHEET);
    var now = new Date();
    
    var roundedQuantity = roundToTwo(quantity);
    
    console.log('📥 Добавление на главный склад: ' + component + ' × ' + roundedQuantity + ' ед. (ID перемещения: ' + transferId + ')');
    
    //  🔥 ВСЕГДА СОЗДАЕМ НОВУЮ ЗАПИСЬ, КАК ПРИ ПЕРЕМЕЩЕНИЯХ
    var status = '';
    if (roundedQuantity < 1) {
      status = '🟡 В наличии: ' + roundedQuantity.toFixed(2) + ' ед. | Меньше единицы';
    } else if (roundedQuantity === 1) {
      status = '✅ В наличии: 1.00 ед. | Последняя';
    } else if (roundedQuantity <= 5) {
      status = '✅ В наличии: ' + roundedQuantity.toFixed(2) + ' ед. | Мало осталось';
    } else {
      status = '✅ В наличии: ' + roundedQuantity.toFixed(2) + ' ед.';
    }
    
    var newRow = [
      transferId,
      purchaseId || '',
      now,
      component,
      roundedQuantity,
      status
    ];
    
    mainSheet.appendRow(newRow);
    
    var newRowIndex = mainSheet.getLastRow();
    mainSheet.getRange(newRowIndex, 3).setNumberFormat('dd.mm.yyyy HH:mm');
    
    console.log('✅ Создана новая запись на главном складе с ID перемещения: ' + transferId);
    
    updateWarehouseStatusDirect(mainSheet);
    
  } catch (error) {
    console.error('❌ Ошибка в addMaterialToMainWarehouse: ' + error.toString());
    throw error;
  }
}


/**
 * 🔥 НОВАЯ ФУНКЦИЯ: ДОБАВЛЕНИЕ МАТЕРИАЛА НА ГЛАВНЫЙ СКЛАД ДЛЯ ПЕРЕМЕЩЕНИЙ
 * Создает новую запись с уникальным transferId
 */
function addMaterialToMainWarehouseNew(component, purchaseId, quantity, transferId) {
  try {
    var mainSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MAIN_WAREHOUSE_SHEET);
    var now = new Date();
    
    var roundedQuantity = roundToTwo(quantity);
    
    console.log('📥 Добавление на главный склад (перемещение): ' + component + ' × ' + roundedQuantity + ' ед. (ID перемещения: ' + transferId + ')');
    
    // Создаем новую запись с уникальным transferId
    var status = '';
    if (roundedQuantity < 1) {
      status = '🟡 В наличии: ' + roundedQuantity.toFixed(2) + ' ед. | Меньше единицы';
    } else if (roundedQuantity === 1) {
      status = '✅ В наличии: 1.00 ед. | Последняя';
    } else if (roundedQuantity <= 5) {
      status = '✅ В наличии: ' + roundedQuantity.toFixed(2) + ' ед. | Мало осталось';
    } else {
      status = '✅ В наличии: ' + roundedQuantity.toFixed(2) + ' ед.';
    }
    
    var newRow = [
      transferId,
      purchaseId || '',
      now,
      component,
      roundedQuantity,
      status
    ];
    
    mainSheet.appendRow(newRow);
    
    var newRowIndex = mainSheet.getLastRow();
    mainSheet.getRange(newRowIndex, 3).setNumberFormat('dd.mm.yyyy HH:mm');
    
    console.log('✅ Создана новая партия на главном складе с ID перемещения: ' + transferId);
    
    updateWarehouseStatusDirect(mainSheet);
    
  } catch (error) {
    console.error('❌ Ошибка в addMaterialToMainWarehouseNew: ' + error.toString());
    throw error;
  }
}


/**
 *  🔥 ОБНОВЛЕННАЯ ФУНКЦИЯ: ДОБАВЛЕНИЕ МАТЕРИАЛА НА СКЛАД ПРОИЗВОДСТВА
 * Теперь всегда создает новую запись, как при перемещениях
 */
function addMaterialToProductionWarehouse(component, purchaseId, quantity, transferId) {
  try {
    var productionSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PRODUCTION_WAREHOUSE_SHEET);
    var now = new Date();
    
    console.log('📥 Добавление на производство: ' + component + ' × ' + quantity + ' ед. (ID перемещения: ' + transferId + ')');
    
    //  🔥 ВСЕГДА СОЗДАЕМ НОВУЮ ЗАПИСЬ, КАК ПРИ ПЕРЕМЕЩЕНИЯХ
    var status = '';
    if (quantity < 1) {
      status = '🟡 В наличии: ' + quantity.toFixed(2) + ' ед. | Меньше единицы';
    } else if (quantity === 1) {
      status = '✅ В наличии: 1.00 ед. | Последняя';
    } else if (quantity <= 5) {
      status = '✅ В наличии: ' + quantity.toFixed(2) + ' ед. | Мало осталось';
    } else {
      status = '✅ В наличии: ' + quantity.toFixed(2) + ' ед.';
    }
    
    var newRow = [
      transferId,
      purchaseId || '',
      now,
      component,
      quantity,
      status
    ];
    
    productionSheet.appendRow(newRow);
    
    var newRowIndex = productionSheet.getLastRow();
    productionSheet.getRange(newRowIndex, 3).setNumberFormat('dd.mm.yyyy HH:mm');
    
    console.log('✅ Создана новая запись на производстве с ID перемещения: ' + transferId);
    
    updateWarehouseStatusDirect(productionSheet);
    
  } catch (error) {
    console.error('❌ Ошибка в addMaterialToProductionWarehouse: ' + error.toString());
    throw error;
  }
}


/**
 * 🔥 НОВАЯ ФУНКЦИЯ: ДОБАВЛЕНИЕ МАТЕРИАЛА НА СКЛАД ПРОИЗВОДСТВА ДЛЯ ПЕРЕМЕЩЕНИЙ
 * Создает новую запись с уникальным transferId
 */
function addMaterialToProductionWarehouseNew(component, purchaseId, quantity, transferId) {
  try {
    var productionSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PRODUCTION_WAREHOUSE_SHEET);
    var now = new Date();
    
    console.log('📥 Добавление на производство (перемещение): ' + component + ' × ' + quantity + ' ед. (ID перемещения: ' + transferId + ')');
    
    // Создаем новую запись с уникальным transferId
    var status = '';
    if (quantity < 1) {
      status = '🟡 В наличии: ' + quantity.toFixed(2) + ' ед. | Меньше единицы';
    } else if (quantity === 1) {
      status = '✅ В наличии: 1.00 ед. | Последняя';
    } else if (quantity <= 5) {
      status = '✅ В наличии: ' + quantity.toFixed(2) + ' ед. | Мало осталось';
    } else {
      status = '✅ В наличии: ' + quantity.toFixed(2) + ' ед.';
    }
    
    var newRow = [
      transferId,
      purchaseId || '',
      now,
      component,
      quantity,
      status
    ];
    
    productionSheet.appendRow(newRow);
    
    var newRowIndex = productionSheet.getLastRow();
    productionSheet.getRange(newRowIndex, 3).setNumberFormat('dd.mm.yyyy HH:mm');
    
    console.log('✅ Создана новая партия на производстве с ID перемещения: ' + transferId);
    
    updateWarehouseStatusDirect(productionSheet);
    
  } catch (error) {
    console.error('❌ Ошибка в addMaterialToProductionWarehouseNew: ' + error.toString());
    throw error;
  }
}

/**
 *  🔥 ОБНОВЛЕННАЯ ФУНКЦИЯ АВТОПЕРЕМЕЩЕНИЯ С ПЕРЕДАЧЕЙ ID НА СКЛАД
 */
function createAutoTransferFromPurchase(purchaseId, component, quantity, targetWarehouse, transferId) {
  try {
    var transfersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TRANSFERS_SHEET);
    
    if (!transfersSheet) {
      transfersSheet = createTransfersSheet(SpreadsheetApp.getActiveSpreadsheet());
    }
    
    var lastRow = transfersSheet.getLastRow() + 1;
    
    // 🔑 ИСПОЛЬЗУЕМ ПЕРЕДАННЫЙ TRANSFER_ID
    var transferId = transferId || generateTransferId();
    
    // СОЗДАЕМ ПОЛНУЮ ЗАПИСЬ СО ВСЕМИ ДАННЫМИ ПО НОВОЙ СТРУКТУРЕ
    var newRow = [
      transferId,                              // A - ID перемещения
      purchaseId,                              // B - ID Закупки
      new Date(),                              // C - Дата
      'Поставщик',                             // D - От склада >
      targetWarehouse,                         // E - > На склад
      component,                               // F - Комплектующее
      quantity,                                // G - Количество
      'Выполнено',                             // H - Статус
      'Автоперемещение при закупке'            // I - Примечание
    ];
    
    // Записываем всю строку сразу
    transfersSheet.getRange(lastRow, 1, 1, 9).setValues([newRow]);
    
    // Форматируем дату
    transfersSheet.getRange(lastRow, 3).setNumberFormat('dd.mm.yyyy HH:mm'); // C - Дата
    
    //  🎯 ВЫЗЫВАЕМ ФУНКЦИИ ДОБАВЛЕНИЯ НА СКЛАД С ПЕРЕДАЧЕЙ TRANSFER_ID
    //  🔥 ТЕПЕРЬ ВСЕГДА СОЗДАЕМ НОВЫЕ ЗАПИСИ
    if (targetWarehouse === 'Главный') {
      addMaterialToMainWarehouse(component, purchaseId, quantity, transferId);
    } else if (targetWarehouse === 'Производство') {
      addMaterialToProductionWarehouse(component, purchaseId, quantity, transferId);
    }
    
    console.log(`📦 Создано автоматическое перемещение: ${component} × ${quantity} от Поставщика на ${targetWarehouse}`);
    console.log(`✅ ID перемещения: ${transferId} записан в таблицу перемещений и передан на склад`);
    
  } catch (error) {
    console.error('Ошибка в createAutoTransferFromPurchase: ' + error.toString());
  }
}

/**
 * ✅ ФУНКЦИЯ: СКРЫТИЕ СТРОК С НУЛЕВЫМ ОСТАТКОМ
 */
function hideZeroStockRows(sheet, rowsToHide) {
  try {
    console.log('👻 Скрытие строк с нулевым остатком: ' + rowsToHide.length + ' строк');
    
    // Сортируем строки в обратном порядке (снизу вверх) чтобы не сбивать индексы при скрытии
    rowsToHide.sort(function(a, b) { return b - a; });
    
    var hiddenCount = 0;
    
    for (var i = 0; i < rowsToHide.length; i++) {
      var row = rowsToHide[i];
      
      // Проверяем, что строка еще не скрыта
      if (!sheet.isRowHiddenByUser(row)) {
        sheet.hideRows(row);
        hiddenCount++;
        console.log('👻 Скрыта строка ' + row);
      }
    }
    
    console.log('✅ Скрыто строк с нулевым остатком: ' + hiddenCount);
    
  } catch (error) {
    console.error('❌ Ошибка в hideZeroStockRows: ' + error.toString());
  }
}

/**
 * ✅ ГЕНЕРАЦИЯ УНИКАЛЬНОГО ID ДЛЯ ЗАКУПКИ
 */
function generatePurchaseId() {
  var timestamp = new Date().getTime();
  var random = Math.floor(Math.random() * 1000);
  return PURCHASE_ID_PREFIX + timestamp + '-' + random;
}

/**
 * ✅ ГЕНЕРАЦИЯ УНИКАЛЬНОГО ID ДЛЯ ПЕРЕМЕЩЕНИЯ
 */
function generateTransferId() {
  var timestamp = new Date().getTime();
  var random = Math.floor(Math.random() * 1000);
  return TRANSFER_ID_PREFIX + timestamp + '-' + random;
}

/**
 * ✅ ПОЛУЧЕНИЕ МИНИМАЛЬНОГО ЗАПАСА ИЗ СПРАВОЧНИКА
 */
function getMinStockFromReference(componentName, warehouseType) {
  try {
    var referenceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Справочник_комплектующих');
    if (!referenceSheet) return 0;
    
    var lastRow = referenceSheet.getLastRow();
    if (lastRow < 2) return 0;
    
    var data = referenceSheet.getRange('A2:C' + lastRow).getValues();
    
    for (var i = 0; i < data.length; i++) {
      var referenceComponent = data[i][0]; // A - Наименование
      
      if (referenceComponent && referenceComponent.toString().trim() === componentName.toString().trim()) {
        if (warehouseType === 'Главный') {
          return parseFloat(data[i][1]) || 0; // B - Минимальный запас (Гл.)
        } else if (warehouseType === 'Производство') {
          return parseFloat(data[i][2]) || 0; // C - Минимальный запас (Пр.)
        }
      }
    }
    
    return 0; // Если компонент не найден
    
  } catch (error) {
    console.error('❌ Ошибка в getMinStockFromReference: ' + error.toString());
    return 0;
  }
}

/**
 * ✅ УЛУЧШЕННАЯ ФУНКЦИЯ: ОКРУГЛЕНИЕ ДО 2 ЗНАКОВ ПОСЛЕ ЗАПЯТОЙ С ВЫСОКОЙ ТОЧНОСТЬЮ
 */
function roundToTwo(num) {
  if (isNaN(num)) return 0;
  // Используем точное округление с учетом погрешности чисел с плавающей точкой
  return Math.round((Number(num) + Number.EPSILON) * 100) / 100;
}

/**
 * 🔥 ФУНКЦИЯ: РАСЧЕТ ОБЩИХ ОСТАТКОВ С 2 ЗНАКАМИ
 */
function calculateComponentTotalsWithDecimals(sheet) {
  try {
    var lastRow = sheet.getLastRow();
    
    if (lastRow <= 1) {
      return {};
    }
    
    var data = sheet.getRange('A2:F' + lastRow).getValues();
    var totals = {};
    
    for (var i = 0; i < data.length; i++) {
      var component = data[i][NEW_WAREHOUSE_STRUCTURE.COMPONENT];
      var currentStock = roundToTwo(data[i][NEW_WAREHOUSE_STRUCTURE.CURRENT_STOCK] || 0);
      
      if (component && component.toString().trim() !== '') {
        if (!totals[component]) {
          totals[component] = 0;
        }
        totals[component] = roundToTwo(totals[component] + currentStock);
      }
    }
    
    return totals;
    
  } catch (error) {
    console.error('❌ Ошибка в calculateComponentTotalsWithDecimals: ' + error.toString());
    return {};
  }
}

/**
 * ✅ ФУНКЦИЯ: РАСЧЕТ ОБЩИХ ОСТАТКОВ ПО КОМПОНЕНТАМ
 */
function calculateComponentTotals(sheet) {
  try {
    var data = sheet.getRange('A2:F' + sheet.getLastRow()).getValues();
    var totals = {};
    
    for (var i = 0; i < data.length; i++) {
      var component = data[i][NEW_WAREHOUSE_STRUCTURE.COMPONENT];
      var currentStock = data[i][NEW_WAREHOUSE_STRUCTURE.CURRENT_STOCK] || 0;
      
      if (component) {
        if (!totals[component]) {
          totals[component] = 0;
        }
        totals[component] += currentStock;
      }
    }
    
    return totals;
    
  } catch (error) {
    console.error('❌ Ошибка в calculateComponentTotals: ' + error.toString());
    return {};
  }
}

/**
 * ✅ ОБНОВЛЕННАЯ ФУНКЦИЯ УМНОГО УДАЛЕНИЯ ДЛЯ ОБОИХ СКЛАДОВ
 */
function removePurchaseRowSmart() {
  try {
    var purchasesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PURCHASES_SHEET);
    var transfersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TRANSFERS_SHEET);
    var mainSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MAIN_WAREHOUSE_SHEET);
    var productionSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PRODUCTION_WAREHOUSE_SHEET);
    var ui = SpreadsheetApp.getUi();
    
    var response = ui.prompt(
      '🗑️ Умное удаление закупки',
      'Введите номер строки закупки для удаления:',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (response.getSelectedButton() !== ui.Button.OK) return;
    
    var row = parseInt(response.getResponseText());
    
    if (isNaN(row) || row < 2) {
      ui.alert('❌ Ошибка', 'Введите корректный номер строки (начиная с 2)', ui.ButtonSet.OK);
      return;
    }
    
    var lastRow = purchasesSheet.getLastRow();
    if (row > lastRow) {
      ui.alert('❌ Ошибка', 'Строка ' + row + ' не существует', ui.ButtonSet.OK);
      return;
    }
    
    // Получаем данные закупки
    var purchaseData = purchasesSheet.getRange(row, 1, 1, 9).getValues()[0];
    var purchaseId = purchaseData[0]; // A - ID
    var component = purchaseData[2]; // C - Комплектующие
    var quantity = purchaseData[3]; // D - Количество
    var targetWarehouse = purchaseData[7]; // H - Склад назначения
    var status = purchaseData[8]; // I - Статус
    
    var message = '🗑️ УМНОЕ УДАЛЕНИЕ ЗАКУПКИ\n\n';
    message += '📋 ИНФОРМАЦИЯ:\n';
    message += '• ID: ' + purchaseId + '\n';
    message += '• Материал: ' + component + '\n';
    message += '• Количество: ' + quantity + ' ед.\n';
    message += '• Склад: ' + targetWarehouse + '\n';
    message += '• Статус: ' + status + '\n\n';
    
    // ПРОВЕРЯЕМ СВЯЗАННЫЕ ДАННЫЕ НА ВСЕХ СКЛАДАХ
    var relatedTransfers = findTransferIdForPurchase(purchaseId);
    var relatedMain = findMaterialsByPurchaseId(mainSheet, purchaseId);
    var relatedProduction = findMaterialsByPurchaseId(productionSheet, purchaseId);
    
    if (relatedTransfers.length > 0) {
      message += '🔄 СВЯЗАННЫЕ ПЕРЕМЕЩЕНИЯ:\n';
      relatedTransfers.forEach(function(transfer) {
        message += '• Строка ' + transfer.row + ': ' + transfer.component + ' × ' + transfer.quantity + ' ед.\n';
      });
      message += '\n';
    }
    
    if (relatedMain.length > 0) {
      message += '🏢 МАТЕРИАЛЫ НА ГЛАВНОМ СКЛАДЕ:\n';
      relatedMain.forEach(function(item) {
        message += '• Строка ' + item.row + ': остаток ' + item.currentStock + ' из ' + item.initialQty + ' ед.\n';
      });
      message += '\n';
    }
    
    if (relatedProduction.length > 0) {
      message += '🏭 МАТЕРИАЛЫ НА ПРОИЗВОДСТВЕ:\n';
      relatedProduction.forEach(function(item) {
        message += '• Строка ' + item.row + ': остаток ' + item.currentStock + ' из ' + item.initialQty + ' ед.\n';
      });
      message += '\n';
    }
    
    message += '⚠️ При удалении закупки будут также удалены все связанные записи!\n';
    message += 'Продолжить?';
    
    var confirmResponse = ui.alert(
      'Подтверждение удаления',
      message,
      ui.ButtonSet.YES_NO
    );
    
    if (confirmResponse !== ui.Button.YES) {
      ui.alert('ℹ️ Отменено', 'Удаление отменено', ui.ButtonSet.OK);
      return;
    }
    
    // УДАЛЯЕМ СВЯЗАННЫЕ ДАННЫЕ
    if (status === 'Получено') {
      // 1. Удаляем материалы со складов (в обратном порядке чтобы номера строк не сдвигались)
      relatedProduction.sort((a, b) => b.row - a.row).forEach(function(item) {
        productionSheet.deleteRow(item.row);
        console.log('✅ Удалена запись на производстве: строка ' + item.row);
      });
      
      relatedMain.sort((a, b) => b.row - a.row).forEach(function(item) {
        mainSheet.deleteRow(item.row);
        console.log('✅ Удалена запись на главном складе: строка ' + item.row);
      });
      
      // 2. Удаляем перемещения (в обратном порядке)
      relatedTransfers.sort((a, b) => b.row - a.row).forEach(function(transfer) {
        transfersSheet.deleteRow(transfer.row);
        console.log('✅ Удалено перемещение: строка ' + transfer.row);
      });
    }
    
    // 3. Удаляем саму закупку
    purchasesSheet.deleteRow(row);
    
    ui.alert('✅ Успех', 'Закупка и все связанные данные удалены', ui.ButtonSet.OK);
    
  } catch (error) {
    console.error('❌ Ошибка в removePurchaseRowSmart: ' + error.toString());
    ui.alert('❌ Ошибка', 'Не удалось удалить закупку: ' + error.toString(), ui.ButtonSet.OK);
  }
}

/**
 * ✅ ИСПРАВЛЕННЫЙ ПОИСК ПЕРЕМЕЩЕНИЙ ПО ID ЗАКУПКИ
 */
function findTransferIdForPurchase(purchaseId) {
  try {
    var transfersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TRANSFERS_SHEET);
    if (!transfersSheet) return [];
    
    var data = transfersSheet.getDataRange().getValues();
    var relatedTransfers = [];
    
    for (var i = 1; i < data.length; i++) { // i=1 пропускаем заголовок
      var transferPurchaseId = data[i][1]; // B - ID Закупки (БЫЛО 5, СТАЛО 1)
      
      if (transferPurchaseId === purchaseId) {
        relatedTransfers.push({
          row: i + 1,
          component: data[i][5], // F - Комплектующее (БЫЛО 4, СТАЛО 5)
          quantity: data[i][6],  // G - Количество (БЫЛО 6, ОСТАЛОСЬ 6)
          transferId: data[i][0] // A - ID перемещения
        });
      }
    }
    
    console.log('🔍 Найдено перемещений для ' + purchaseId + ': ' + relatedTransfers.length);
    return relatedTransfers;
    
  } catch (error) {
    console.error('Ошибка в findTransferIdForPurchase: ' + error.toString());
    return [];
  }
}

/**
 * ✅ УНИВЕРСАЛЬНАЯ ФУНКЦИЯ ПОИСКА МАТЕРИАЛОВ ПО ID ЗАКУПКИ
 */
function findMaterialsByPurchaseId(sheet, purchaseId) {
  try {
    var data = sheet.getDataRange().getValues();
    var relatedItems = [];
    
    for (var i = 1; i < data.length; i++) {
      var itemPurchaseId = data[i][1]; // B - ID Закупки
      if (itemPurchaseId === purchaseId) {
        relatedItems.push({
          row: i + 1,
          component: data[i][0], // A - Комплектующее
          initialQty: data[i][3], // D - Количество
          currentStock: data[i][4] // E - Остаток
        });
      }
    }
    
    return relatedItems;
  } catch (error) {
    console.error('Ошибка в findMaterialsByPurchaseId: ' + error.toString());
    return [];
  }
}

/**
 * ✅ ОБНОВЛЕННАЯ ФУНКЦИЯ: ОБНОВЛЕНИЕ СТАТУСА ДЛЯ ОДНОЙ СТРОКИ С МИН. ЗАПАСОМ ИЗ СПРАВОЧНИКА
 */
function updateSingleRowStatus(sheet, row) {
  try {
    var data = sheet.getRange(row, 1, 1, 6).getValues()[0];
    var component = data[0]; // A - Комплектующие
    var purchaseId = data[1]; // B - ID Закупки
    var receiptDate = data[2]; // C - Дата поступления
    var currentStock = data[3] || 0; // D - Остаток
    
    if (!component) return;
    
    // 🎯 ПОЛУЧАЕМ МИН. ЗАПАС ИЗ СПРАВОЧНИКА
    var warehouseType = sheet.getName() === 'Склад_Главный' ? 'Главный' : 'Производство';
    var minStock = getMinStockFromReference(component, warehouseType);
    
    // Рассчитываем общий остаток по компоненту
    var componentTotals = calculateComponentTotals(sheet);
    var totalStock = componentTotals[component] || 0;
    
    var status = '';
    var backgroundColor = '';
    var statusType = '';
    
    // ИНФОРМАТИВНАЯ ЛОГИКА СТАТУСОВ:
    if (currentStock === 0) {
      status = '❌ Отсутствует';
      backgroundColor = '#FFCCCC';
      statusType = 'absent';
    } else if (minStock > 0 && totalStock < minStock) {
      var deficit = minStock - totalStock;
      status = '⚠️ Мало: ' + currentStock + '/' + minStock + ' ед. | -' + deficit;
      backgroundColor = '#FFF2CC';
      statusType = 'low';
    } else if (currentStock === 1) {
      status = '✅ В наличии: ' + currentStock + ' ед. | Последняя';
      backgroundColor = '#FFF2CC';
      statusType = 'last';
    } else if (currentStock <= 5) {
      status = '✅ В наличии: ' + currentStock + ' ед. | Мало осталось';
      backgroundColor = '#CCFFCC';
      statusType = 'few';
    } else {
      status = '✅ В наличии: ' + currentStock + ' ед.';
      backgroundColor = '#CCFFCC';
      statusType = 'normal';
    }
    
    // Обновляем статус (столбец F)
    var statusCell = sheet.getRange(row, 6);
    statusCell.setValue(status);
    statusCell.setBackground(backgroundColor);
    
    // 🔥 ОБНОВЛЯЕМ ПОДСКАЗКУ
    var tooltip = createTooltip(component, purchaseId, receiptDate, currentStock, minStock, totalStock, statusType);
    setCellNote(statusCell, tooltip);
    
  } catch (error) {
    console.error('❌ Ошибка в updateSingleRowStatus: ' + error.toString());
  }
}

/**
 * 🔥 ФУНКЦИЯ: СОЗДАНИЕ ПОДСКАЗКИ С 2 ЗНАКАМИ
 */
function createTooltip(component, purchaseId, date, currentStock, minStock, totalStock, statusType) {
  // 🔥 ОКРУГЛЯЕМ ВСЕ ЧИСЛА ДО 2 ЗНАКОВ
  currentStock = roundToTwo(currentStock || 0);
  minStock = roundToTwo(minStock || 0);
  totalStock = roundToTwo(totalStock || 0);
  
  var tooltip = '';
  
  // Основная информация
  tooltip += '📦 ' + (component || 'Неизвестный компонент') + '\n';
  tooltip += '🏷️ ID: ' + (purchaseId || 'Неизвестен') + '\n';
  
  if (date) {
    var dateStr = formatDateTime(date);
    tooltip += '📅 Поступление: ' + dateStr + '\n';
  }
  
  tooltip += '📊 В этой партии: ' + currentStock.toFixed(2) + ' ед.\n';
  tooltip += '📈 Общий остаток: ' + totalStock.toFixed(2) + ' ед.\n';
  tooltip += '🎯 Мин. запас: ' + minStock.toFixed(2) + ' ед.\n\n';
  
  // Статус запаса
  if (statusType === 'absent') {
    tooltip += '❌ ЗАПАС ОТСУТСТВУЕТ\n\n';
  } else if (totalStock < minStock) {
    var deficit = roundToTwo(minStock - totalStock);
    tooltip += '🟡 ВНИМАНИЕ: Ниже минимального запаса (дефицит: ' + deficit.toFixed(2) + ' ед.)\n\n';
  } else if (currentStock < 1) {
    tooltip += '🟡 ВНИМАНИЕ: МЕНЬШЕ ЕДИНИЦЫ\n\n';
  } else if (statusType === 'last') {
    tooltip += '🔴 ПОСЛЕДНЯЯ ЕДИНИЦА\n\n';
  } else if (statusType === 'few') {
    tooltip += '🟡 МАЛО ОСТАТКОВ\n\n';
  } else {
    tooltip += '✅ Запас достаточен\n\n';
  }
  
  // 🔥 РЕКОМЕНДАЦИИ С 2 ЗНАКАМИ
  tooltip += '💡 РЕКОМЕНДАЦИИ:\n';
  
  if (statusType === 'absent') {
    tooltip += '• Закупить: ' + minStock.toFixed(2) + ' ед.\n';
    tooltip += '• Срочность: 🔴 КРИТИЧЕСКАЯ\n';
  } else if (totalStock < minStock) {
    var needed = roundToTwo(minStock - totalStock);
    tooltip += '• Закупить: ' + needed.toFixed(2) + ' ед.\n';
    
    if (needed <= 5) {
      tooltip += '• Срочность: 🟡 Средняя\n';
    } else {
      tooltip += '• Срочность: 🔴 Высокая\n';
    }
  } else if (currentStock < 1) {
    tooltip += '• Контролировать остаток\n';
    tooltip += '• Срочность: 🟡 Средняя\n';
  } else if (statusType === 'last' || statusType === 'few') {
    tooltip += '• Рекомендуется докупить\n';
    tooltip += '• Срочность: 🟡 Средняя\n';
  } else {
    tooltip += '• Докупка не требуется\n';
    tooltip += '• Срочность: ✅ Низкая\n';
  }
  
  // Дополнительная информация
  tooltip += '\n📋 ДЕТАЛИ:\n';
  tooltip += '• Остаток в партии: ' + currentStock.toFixed(2) + ' ед.\n';
  tooltip += '• Всего на складе: ' + totalStock.toFixed(2) + ' ед.\n';
  tooltip += '• Минимальный запас: ' + minStock.toFixed(2) + ' ед.';
  
  if (totalStock < minStock) {
    var deficit = roundToTwo(minStock - totalStock);
    tooltip += '\n• Дефицит: ' + deficit.toFixed(2) + ' ед.';
  }
  
  return tooltip;
}

/**
 * ✅ ФОРМАТИРОВАНИЕ ДАТЫ И ВРЕМЕНИ ДЛЯ ЛОГИРОВАНИЯ
 */
function formatDateTime(date) {
  try {
    if (!date) return 'Не указана';
    return Utilities.formatDate(date, Session.getScriptTimeZone(), "dd.MM.yyyy HH:mm:ss");
  } catch (error) {
    return date.toString();
  }
}

/**
 * ✅ ФУНКЦИЯ: УСТАНОВКА КОММЕНТАРИЯ В ЯЧЕЙКУ
 */
function setCellNote(cell, note) {
  try {
    // Очищаем старый комментарий и устанавливаем новый
    cell.setNote(note);
  } catch (error) {
    console.error('❌ Ошибка установки комментария: ' + error.toString());
  }
}

/**
 * ✅ ИСПРАВЛЕННАЯ ФУНКЦИЯ: ПОКАЗ ДИАЛОГА СОЗДАНИЯ ПЕРЕМЕЩЕНИЯ (КОМПЛЕКТУЮЩИЕ ИЗ ВЫБРАННОГО СКЛАДА)
 */
function showTransferDialog() {
  try {
    var html = HtmlService.createHtmlOutput(`
      <!DOCTYPE html>
      <html>
      <head>
        <base target="_top">
        <style>
          body { font-family: Arial, sans-serif; padding: 20px; background: #f5f5f5; }
          .container { background: white; padding: 20px; border-radius: 10px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
          h3 { color: #333; margin-bottom: 20px; }
          label { display: block; margin: 10px 0 5px; font-weight: bold; color: #555; }
          select, input, textarea { width: 100%; padding: 10px; margin: 5px 0 15px; border: 1px solid #ddd; border-radius: 5px; box-sizing: border-box; }
          button { background: #4CAF50; color: white; padding: 12px 25px; border: none; border-radius: 5px; cursor: pointer; font-size: 16px; width: 100%; }
          button:hover { background: #45a049; }
          .stock-info { background: #e7f3ff; padding: 10px; border-radius: 5px; margin: 10px 0; font-size: 14px; }
          .warning { background: #fff3cd !important; color: #856404; }
          .loading { color: #666; font-style: italic; }
          .batch-info { background: #f0f8ff; padding: 8px; border-radius: 5px; margin: 5px 0; font-size: 13px; }
          .success { background: #d4edda !important; color: #155724; }
          .hidden { display: none; }
        </style>
      </head>
      <body>
        <div class="container">
          <h3>📦 Создание перемещения</h3>
          <form id="transferForm">
            <label>От склада:</label>
            <select name="fromWarehouse" id="fromWarehouse" onchange="updateComponentsList()">
              <option value="Главный">🏢 Главный склад</option>
              <option value="Производство">🏭 Склад производства</option>
            </select>
            
            <label>На склад:</label>
            <select name="toWarehouse" id="toWarehouse" onchange="validateWarehouses()">
              <option value="Производство">🏭 Склад производства</option>
              <option value="Главный">🏢 Главный склад</option>
            </select>
            
            <label>Комплектующее:</label>
            <select name="component" id="component" onchange="updatePurchaseIdsList()">
              <option value="" class="loading">⏳ Загрузка комплектующих...</option>
            </select>
            
            <div id="componentStockInfo" class="stock-info hidden">
              📊 Общий остаток на складе: <span id="totalStock">0</span> ед.
            </div>
            
            <label>ID Закупки (партия):</label>
            <select name="purchaseId" id="purchaseId" onchange="updateBatchInfo()">
              <option value="">-- Сначала выберите комплектующее --</option>
            </select>
            
            <div id="batchInfo" class="batch-info hidden">
              📦 Информация о партии:<br>
              • Доступно: <span id="availableStock">0</span> ед.<br>
              • Дата поступления: <span id="receiptDate">-</span><br>
              • Изначальное количество: <span id="initialQuantity">-</span>
            </div>
            
            <label>Количество для перемещения:</label>
            <input type="number" name="quantity" id="quantity" placeholder="10" min="1" step="1">
            
            <label>Примечание:</label>
            <textarea name="notes" placeholder="Причина перемещения..."></textarea>
            
            <button type="button" onclick="createTransfer()">📋 Создать перемещение</button>
          </form>
        </div>
        
        <script>
          // 🔥 ИСПРАВЛЕНО: Загружаем комплектующие при загрузке страницы для выбранного склада
          document.addEventListener('DOMContentLoaded', function() {
            updateComponentsList();
          });
          
          // Обновляем список комплектующих при смене склада
          function updateComponentsList() {
            var warehouse = document.getElementById('fromWarehouse').value;
            var componentSelect = document.getElementById('component');
            var purchaseIdSelect = document.getElementById('purchaseId');
            var componentStockInfo = document.getElementById('componentStockInfo');
            var batchInfo = document.getElementById('batchInfo');
            
            // Сбрасываем зависимые поля
            componentSelect.innerHTML = '<option value="" class="loading">⏳ Загрузка комплектующих...</option>';
            purchaseIdSelect.innerHTML = '<option value="">-- Сначала выберите комплектующее --</option>';
            componentStockInfo.classList.add('hidden');
            batchInfo.classList.add('hidden');
            document.getElementById('quantity').value = '';
            
            google.script.run
              .withSuccessHandler(function(components) {
                componentSelect.innerHTML = '<option value="">-- Выберите комплектующее --</option>';
                if (components.length > 0) {
                  components.forEach(function(comp) {
                    componentSelect.innerHTML += '<option value="' + comp + '">' + comp + '</option>';
                  });
                } else {
                  componentSelect.innerHTML = '<option value="">❌ На складе нет комплектующих</option>';
                }
              })
              .withFailureHandler(function(error) {
                componentSelect.innerHTML = '<option value="">❌ Ошибка загрузки</option>';
                console.error('Ошибка загрузки компонентов:', error);
              })
              .getUniqueWarehouseComponents(warehouse);
          }
          
          // Обновляем список ID закупок при выборе комплектующего
function updatePurchaseIdsList() {
  var component = document.getElementById('component').value;
  var warehouse = document.getElementById('fromWarehouse').value;
  var purchaseIdSelect = document.getElementById('purchaseId');
  var componentStockInfo = document.getElementById('componentStockInfo');
  var batchInfo = document.getElementById('batchInfo');
  
  if (!component) {
    purchaseIdSelect.innerHTML = '<option value="">-- Сначала выберите комплектующее --</option>';
    componentStockInfo.classList.add('hidden');
    batchInfo.classList.add('hidden');
    return;
  }
  
  // Сбрасываем поля
  purchaseIdSelect.innerHTML = '<option value="" class="loading">⏳ Загрузка партий...</option>';
  componentStockInfo.classList.add('hidden');
  batchInfo.classList.add('hidden');
  document.getElementById('quantity').value = '';
  
  google.script.run
    .withSuccessHandler(function(result) {
      purchaseIdSelect.innerHTML = '<option value="">-- Выберите партию --</option>';
      
      if (result && result.batches && result.batches.length > 0) {
        // Показываем общий остаток
        document.getElementById('totalStock').textContent = result.totalStock;
        componentStockInfo.classList.remove('hidden');
        
        if (result.totalStock > 0) {
          componentStockInfo.className = 'stock-info success';
        } else {
          componentStockInfo.className = 'stock-info warning';
        }
        
        // 🔥 ЗАПОЛНЯЕМ СПИСОК ПАРТИЙ С УЧЕТОМ ВСЕХ СТРОК
        result.batches.forEach(function(batch) {
          // 🔥 ФОРМИРУЕМ УНИКАЛЬНОЕ ОПИСАНИЕ ДЛЯ КАЖДОЙ СТРОКИ
          var optionText = batch.purchaseId;
          if (batch.transferId && batch.transferId !== 'Без ID') {
            optionText += ' | ID пер.: ' + batch.transferId;
          }
          optionText += ' | ' + batch.availableStock + ' ед. | ' + batch.receiptDate;
          
          var option = document.createElement('option');
          option.value = batch.purchaseId;
          option.textContent = optionText;
          option.setAttribute('data-stock', batch.availableStock);
          option.setAttribute('data-date', batch.receiptDate);
          option.setAttribute('data-initial', batch.initialQuantity);
          option.setAttribute('data-transferid', batch.transferId);
          option.setAttribute('data-purchaseid', batch.purchaseId);
          
          purchaseIdSelect.appendChild(option);
        });
      } else {
        purchaseIdSelect.innerHTML = '<option value="">❌ Нет доступных партий</option>';
        // Все равно показываем общий остаток (0)
        document.getElementById('totalStock').textContent = 0;
        componentStockInfo.classList.remove('hidden');
        componentStockInfo.className = 'stock-info warning';
      }
    })
    .withFailureHandler(function(error) {
      purchaseIdSelect.innerHTML = '<option value="">❌ Ошибка загрузки</option>';
      console.error('Ошибка загрузки партий:', error);
    })
    .getAvailableBatchesWithTotal(warehouse, component);
}
          
          function updateBatchInfo() {
  var purchaseIdSelect = document.getElementById('purchaseId');
  var selectedOption = purchaseIdSelect.options[purchaseIdSelect.selectedIndex];
  var batchInfo = document.getElementById('batchInfo');
  var quantityInput = document.getElementById('quantity');
  
  if (selectedOption.value) {
    var availableStock = selectedOption.getAttribute('data-stock');
    var receiptDate = selectedOption.getAttribute('data-date');
    var initialQuantity = selectedOption.getAttribute('data-initial');
    var transferId = selectedOption.getAttribute('data-transferid');
    var purchaseId = selectedOption.getAttribute('data-purchaseid');
    
    document.getElementById('availableStock').textContent = availableStock;
    document.getElementById('receiptDate').textContent = receiptDate;
    document.getElementById('initialQuantity').textContent = initialQuantity;
    
    //  🔥 ДОБАВЛЯЕМ ИНФОРМАЦИЮ О ID ПЕРЕМЕЩЕНИЯ И ЗАКУПКИ
    var transferInfo = document.getElementById('transferInfo');
    if (!transferInfo) {
      // Создаем элемент если его нет
      var batchInfoDiv = document.getElementById('batchInfo');
      transferInfo = document.createElement('div');
      transferInfo.id = 'transferInfo';
      transferInfo.className = 'batch-info';
      transferInfo.innerHTML = '🔑 ID перемещения: <span id="transferId">-</span><br>🏷️ ID закупки: <span id="purchaseIdValue">-</span>';
      batchInfoDiv.appendChild(transferInfo);
    }
    
    document.getElementById('transferId').textContent = transferId || 'Без ID';
    document.getElementById('purchaseIdValue').textContent = purchaseId || 'Не указан';
    
    batchInfo.classList.remove('hidden');
    
    // Устанавливаем максимальное количество и предлагаемое значение
    if (availableStock > 0) {
      quantityInput.max = availableStock;
      if (!quantityInput.value || parseInt(quantityInput.value) > availableStock) {
        quantityInput.value = Math.min(availableStock, 10);
      }
    } else {
      quantityInput.value = '';
      quantityInput.max = '';
    }
  } else {
    batchInfo.classList.add('hidden');
    // Если партия не выбрана, сбрасываем ограничения по количеству
    quantityInput.max = '';
  }
}
          
          // Проверяем чтобы склады были разными
          function validateWarehouses() {
            var fromWarehouse = document.getElementById('fromWarehouse').value;
            var toWarehouse = document.getElementById('toWarehouse').value;
            
            if (fromWarehouse === toWarehouse) {
              // Автоматически меняем целевой склад на противоположный
              var newToWarehouse = fromWarehouse === 'Главный' ? 'Производство' : 'Главный';
              document.getElementById('toWarehouse').value = newToWarehouse;
            }
          }
          
          function createTransfer() {
            var form = document.getElementById('transferForm');
            var component = form.component.value;
            var purchaseId = form.purchaseId.value;
            var quantity = parseInt(form.quantity.value);
            var fromWarehouse = form.fromWarehouse.value;
            var toWarehouse = form.toWarehouse.value;
            var notes = form.notes.value;
            
            if (!component) {
              alert('❌ Выберите комплектующее');
              return;
            }
            
            if (!purchaseId) {
              alert('❌ Выберите ID закупки (партию)');
              return;
            }
            
            if (!quantity || quantity <= 0) {
              alert('❌ Введите корректное количество');
              return;
            }
            
            if (fromWarehouse === toWarehouse) {
              alert('❌ Склады должны быть разными');
              return;
            }
            
            // Проверяем остаток в выбранной партии
            var selectedOption = document.getElementById('purchaseId').options[document.getElementById('purchaseId').selectedIndex];
            var availableStock = parseInt(selectedOption.getAttribute('data-stock'));
            
            if (quantity > availableStock) {
              alert('❌ Недостаточно единиц в выбранной партии! Доступно: ' + availableStock);
              return;
            }
            
            var formData = {
              fromWarehouse: fromWarehouse,
              toWarehouse: toWarehouse,
              component: component,
              purchaseId: purchaseId,
              quantity: quantity,
              notes: notes
            };
            
            var button = document.querySelector('button');
            var originalText = button.innerHTML;
            button.innerHTML = '⏳ Создание...';
            button.disabled = true;
            
            google.script.run
              .withSuccessHandler(function() {
                alert('✅ Перемещение создано!');
                google.script.host.close();
              })
              .withFailureHandler(function(error) {
                alert('❌ Ошибка: ' + error);
                button.innerHTML = originalText;
                button.disabled = false;
              })
              .createTransferManual(formData);
          }
        </script>
      </body>
      </html>
    `).setWidth(500).setHeight(700);
    
    SpreadsheetApp.getUi().showModalDialog(html, '📦 Создание перемещения');
    
  } catch (error) {
    console.error('Ошибка в showTransferDialog: ' + error.toString());
    SpreadsheetApp.getUi().alert('❌ Ошибка при создании диалога: ' + error.toString());
  }
}

/**
 * ✅ ОБНОВЛЕНО: ПОЛУЧЕНИЕ УНИКАЛЬНЫХ КОМПЛЕКТУЮЩИХ СКЛАДА (НОВАЯ СТРУКТУРА)
 */
function getUniqueWarehouseComponents(warehouseName) {
  try {
    var warehouseSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
      warehouseName === 'Главный' ? MAIN_WAREHOUSE_SHEET : PRODUCTION_WAREHOUSE_SHEET
    );
    
    if (!warehouseSheet) {
      console.error('❌ Склад не найден: ' + warehouseName);
      return [];
    }
    
    var lastRow = warehouseSheet.getLastRow();
    if (lastRow < 2) {
      return [];
    }
    
    // ОБНОВЛЕНО: Получаем данные из колонки D (Комплектующие) вместо A
    var data = warehouseSheet.getRange('D2:D' + lastRow).getValues();
    var componentsSet = new Set();
    
    for (var i = 0; i < data.length; i++) {
      // ОБНОВЛЕНО: Проверяем колонку D (Комплектующие)
      if (data[i][0] && data[i][0].toString().trim() !== '') {
        componentsSet.add(data[i][0].toString().trim());
      }
    }
    
    var uniqueComponents = Array.from(componentsSet).sort();
    console.log('📋 Уникальные комплектующие на ' + warehouseName + ': ' + uniqueComponents.length + ' наименований');
    
    return uniqueComponents;
    
  } catch (error) {
    console.error('Ошибка в getUniqueWarehouseComponents: ' + error.toString());
    return [];
  }
}

/**
 * ✅ ИСПРАВЛЕННАЯ ФУНКЦИЯ: ПОЛУЧЕНИЕ ПАРТИЙ С ОБЩИМ ОСТАТКОМ (ПРАВИЛЬНАЯ СТРУКТУРА)
 * Теперь показывает отдельные строки для каждой партии с уникальными ID перемещений
 */
function getAvailableBatchesWithTotal(warehouseName, component) {
  try {
    console.log('🔍 Поиск партий для:', component, 'на складе:', warehouseName);
    
    var warehouseSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
      warehouseName === 'Главный' ? MAIN_WAREHOUSE_SHEET : PRODUCTION_WAREHOUSE_SHEET
    );
    
    if (!warehouseSheet) {
      console.error('❌ Склад не найден: ' + warehouseName);
      return { batches: [], totalStock: 0 };
    }
    
    var lastRow = warehouseSheet.getLastRow();
    if (lastRow < 2) {
      console.log('ℹ️ На складе нет данных');
      return { batches: [], totalStock: 0 };
    }
    
    // 🔥 ПРАВИЛЬНАЯ СТРУКТУРА: 
    // A: ID перемещения, B: ID Закупки, C: Дата поступления, D: Комплектующие, E: Остаток, F: Статус
    var data = warehouseSheet.getRange('A2:F' + lastRow).getValues();
    console.log('📊 Найдено строк:', data.length);
    
    var batches = [];
    var totalStock = 0;
    
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var transferId = row[0];    // A - ID перемещения
      var purchaseId = row[1];    // B - ID Закупки
      var receiptDate = row[2];   // C - Дата поступления
      var rowComponent = row[3];  // D - Комплектующие
      var availableStock = row[4];// E - Остаток
      var status = row[5];        // F - Статус
      
      // Проверяем совпадение и доступность
      if (rowComponent && 
          rowComponent.toString().trim().toLowerCase() === component.toString().trim().toLowerCase() &&
          purchaseId && 
          purchaseId.toString().trim() !== '' &&
          availableStock > 0) {
        
        // 🔥 ПОЛУЧАЕМ ПРАВИЛЬНОЕ ИЗНАЧАЛЬНОЕ КОЛИЧЕСТВО
        var initialQuantity = getInitialQuantityFromPurchase(purchaseId, component, transferId);
        
        // Форматируем дату для отображения
        var formattedDate = 'Не указана';
        if (receiptDate) {
          if (Object.prototype.toString.call(receiptDate) === '[object Date]') {
            formattedDate = Utilities.formatDate(receiptDate, Session.getScriptTimeZone(), "dd.MM.yyyy");
          } else if (typeof receiptDate === 'string' && receiptDate.trim() !== '') {
            formattedDate = receiptDate;
          }
        }
        
        // 🔥 СОЗДАЕМ ОТДЕЛЬНУЮ ЗАПИСЬ ДЛЯ КАЖДОЙ СТРОКИ С ПРАВИЛЬНЫМ ИЗНАЧАЛЬНЫМ КОЛИЧЕСТВОМ
        var batch = {
          transferId: transferId ? transferId.toString().trim() : 'Без ID',
          purchaseId: purchaseId.toString().trim(),
          availableStock: Number(availableStock),
          receiptDate: formattedDate,
          initialQuantity: initialQuantity, // 🔥 ПРАВИЛЬНОЕ ИЗНАЧАЛЬНОЕ КОЛИЧЕСТВО
          rowIndex: i + 2 // Номер строки в таблице для отладки
        };
        
        batches.push(batch);
        totalStock += Number(availableStock);
        
        console.log('✅ Найдена партия:', batch.purchaseId, 'ID перемещения:', batch.transferId, 
                   'Остаток:', batch.availableStock, 'Изначально:', batch.initialQuantity);
      }
    }
    
    // 🔥 УБИРАЕМ ОБЪЕДИНЕНИЕ ДУБЛИКАТОВ - ТЕПЕРЬ КАЖДАЯ СТРОКА ОТОБРАЖАЕТСЯ ОТДЕЛЬНО
    var uniqueBatches = batches; // Просто используем все найденные партии
    
    // 🔥 СОРТИРУЕМ ПАРТИИ ПО ДАТЕ (СТАРЫЕ ПЕРВЫМИ - FIFO) И ПО ID ПЕРЕМЕЩЕНИЯ
    uniqueBatches.sort(function(a, b) {
      try {
        // Сначала по дате
        var dateCompare = new Date(a.receiptDate) - new Date(b.receiptDate);
        if (dateCompare !== 0) return dateCompare;
        
        // Если даты одинаковые, то по ID перемещения
        return a.transferId.localeCompare(b.transferId);
      } catch (e) {
        return 0;
      }
    });
    
    console.log('📦 Итоговые партии:', uniqueBatches.length, 'Общий остаток:', totalStock);
    
    return {
      batches: uniqueBatches,
      totalStock: totalStock
    };
    
  } catch (error) {
    console.error('❌ Ошибка в getAvailableBatchesWithTotal: ' + error.toString());
    return { batches: [], totalStock: 0 };
  }
}

// 🔥 УЛУЧШЕННАЯ ФУНКЦИЯ: ПОЛУЧЕНИЕ ИЗНАЧАЛЬНОГО КОЛИЧЕСТВА
function getInitialQuantityFromPurchase(purchaseId, component, transferId) {
  try {
    // Сначала пробуем получить из закупок
    var purchaseDetails = getPurchaseDetails(purchaseId);
    
    if (purchaseDetails.success) {
      // Проверяем совпадение компонента
      if (purchaseDetails.component && 
          purchaseDetails.component.toString().trim().toLowerCase() === component.toString().trim().toLowerCase()) {
        
        console.log('✅ Найдена закупка:', purchaseId, 'Компонент:', component, 'Количество:', purchaseDetails.quantity);
        return Number(purchaseDetails.quantity);
      }
    }
    
    // Если не нашли в закупках, пробуем в перемещениях
    if (transferId && transferId !== 'Без ID') {
      var transferQuantity = getInitialQuantityFromTransfer(transferId, component);
      if (transferQuantity > 0) {
        return transferQuantity;
      }
    }
    
    console.log('⚠️ Не удалось получить изначальное количество для:', purchaseId, 'Компонент:', component);
    return 0;
    
  } catch (error) {
    console.error('❌ Ошибка получения изначального количества: ' + error.toString());
    return 0;
  }
}

// 🔥 НОВАЯ ФУНКЦИЯ: ПОЛУЧЕНИЕ ПОДРОБНОЙ ИНФОРМАЦИИ О ЗАКУПКЕ
function getPurchaseDetails(purchaseId) {
  try {
    var purchasesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PURCHASES_SHEET);
    if (!purchasesSheet) {
      return { success: false, error: 'Лист закупок не найден' };
    }
    
    var lastRow = purchasesSheet.getLastRow();
    if (lastRow < 2) {
      return { success: false, error: 'Лист закупок пуст' };
    }
    
    // Ищем закупку по ID
    var data = purchasesSheet.getRange('A2:I' + lastRow).getValues();
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var rowPurchaseId = row[0]; // A - ID закупки
      
      if (rowPurchaseId && rowPurchaseId.toString().trim() === purchaseId.toString().trim()) {
        return {
          success: true,
          purchaseId: rowPurchaseId,
          date: row[1],           // B - Дата
          component: row[2],      // C - Комплектующие
          quantity: row[3],       // D - Количество
          price: row[4],         // E - Цена
          sum: row[5],           // F - Сумма
          supplier: row[6],      // G - Поставщик
          targetWarehouse: row[7], // H - Склад назначения
          status: row[8]         // I - Статус
        };
      }
    }
    
    return { success: false, error: 'Закупка не найдена: ' + purchaseId };
    
  } catch (error) {
    console.error('❌ Ошибка получения информации о закупке: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * ✅ ВСПОМОГАТЕЛЬНАЯ ФУНКЦИЯ: ОКРУГЛЕНИЕ ДО 2 ЗНАКОВ
 */
function roundToTwo(num) {
  return Math.round(num * 100) / 100;
}

/**
 * ✅ АДАПТИРОВАННАЯ ФУНКЦИЯ: СОЗДАНИЕ ПЕРЕМЕЩЕНИЯ (ПОД СТРУКТУРУ ПЕРЕМЕЩЕНИЙ)
 */
function createTransferManual(formData) {
  try {
    smartDelay(300);
    
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var transfersSheet = spreadsheet.getSheetByName(TRANSFERS_SHEET);
    
    if (!transfersSheet) {
      transfersSheet = createTransfersSheet(spreadsheet);
    }
    
    var lastRow = transfersSheet.getLastRow() + 1;
    
    // ✅ АВТОМАТИЧЕСКАЯ ДАТА - ТЕКУЩЕЕ ВРЕМЯ
    var transferDate = new Date();
    
    // 🔥 АДАПТИРОВАННАЯ СТРУКТУРА ПОД ВАШ ЛИСТ ПЕРЕМЕЩЕНИЙ:
    // A: ID перемещения | B: ID Закупки | C: Дата | D: От склада > | E: На склад | F: Комплектующее | G: Количество | H: Статус | I: Примечание
    var newRow = [
      generateTransferId(),     // A - ID перемещения
      formData.purchaseId || '', // B - ID Закупки
      transferDate,             // C - Дата (автоматически)
      formData.fromWarehouse,   // D - От склада >
      formData.toWarehouse,     // E - На склад
      formData.component,       // F - Комплектующее
      formData.quantity,        // G - Количество
      'В процессе',             // H - Статус
      formData.notes || ''      // I - Примечание
    ];
    
    // Записываем всю строку сразу
    transfersSheet.getRange(lastRow, 1, 1, 9).setValues([newRow]);
    
    // Форматируем дату
    transfersSheet.getRange(lastRow, 3).setNumberFormat('dd.mm.yyyy HH:mm'); // Колонка C - Дата
    
    console.log(`📦 Создано перемещение: ${formData.component} × ${formData.quantity} из ${formData.fromWarehouse} в ${formData.toWarehouse}`);
    if (formData.purchaseId) {
      console.log(`🏷️ ID закупки: ${formData.purchaseId}`);
    }
    console.log('✅ Все столбцы перемещения заполнены по новой структуре');
    
    return true;
    
  } catch (error) {
    console.error('❌ Ошибка в createTransferManual: ' + error.toString());
    throw error;
  }
}

/**
 * ✅ ФУНКЦИЯ ДЛЯ УМНОЙ ЗАДЕРЖКИ
 */
function smartDelay(ms) {
  try {
    Utilities.sleep(ms || 500); // Задержка по умолчанию 500мс
  } catch (error) {
    // Игнорируем ошибки задержки
  }
}

/**
 * ✅ ИСПРАВЛЕННАЯ ФУНКЦИЯ ПОКАЗА ОЖИДАЮЩИХ ПЕРЕМЕЩЕНИЙ (ПРАВИЛЬНЫЕ ИНДЕКСЫ)
 */
function showPendingTransfers() {
  try {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var transfersSheet = spreadsheet.getSheetByName(TRANSFERS_SHEET);
    var ui = SpreadsheetApp.getUi();
    
    if (!transfersSheet) {
      throw new Error('Лист перемещений не найден: ' + TRANSFERS_SHEET);
    }
    
    var data = transfersSheet.getDataRange().getValues();
    var pendingTransfers = [];
    
    console.log('🔍 Поиск перемещений со статусом "В процессе"...');
    console.log('📊 Структура листа Перемещений:');
    console.log('A: ID перемещения, B: ID Закупки, C: Дата, D: От склада >, E: На склад, F: Комплектующее, G: Количество, H: Статус, I: Примечание');
    
    // 🔥 ИСПРАВЛЕНИЕ: Правильные индексы столбцов
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      
      // 🔥 ПРАВИЛЬНЫЕ ИНДЕКСЫ ПО РЕАЛЬНОЙ СТРУКТУРЕ:
      var transferId = row[0];        // A - ID перемещения (индекс 0)
      var purchaseId = row[1];        // B - ID Закупки (индекс 1)
      var date = row[2];              // C - Дата (индекс 2)
      var fromWarehouse = row[3];     // D - От склада > (индекс 3)
      var toWarehouse = row[4];       // E - На склад (индекс 4)
      var component = row[5];         // F - Комплектующее (индекс 5)
      var quantity = row[6];          // G - Количество (индекс 6)
      var status = row[7];            // H - Статус (индекс 7)
      var notes = row[8];             // I - Примечание (индекс 8)
      
      console.log('Строка ' + (i + 1) + ': ', component, 'Статус:', status, 'Количество:', quantity);
      
      // 🔥 ИСПРАВЛЕНИЕ: Правильная проверка статуса
      if (status && status.toString().trim() === 'В процессе' && component && quantity > 0) {
        var transferInfo = 'Строка ' + (i + 1) + ': ' + component + ' × ' + quantity + ' ед. (' + fromWarehouse + ' → ' + toWarehouse + ')';
        if (purchaseId) {
          transferInfo += ' | ID: ' + purchaseId;
        }
        
        pendingTransfers.push({
          row: i + 1,
          info: transferInfo,
          component: component,
          quantity: quantity,
          fromWarehouse: fromWarehouse,
          toWarehouse: toWarehouse,
          purchaseId: purchaseId,
          transferId: transferId
        });
        console.log('✅ Добавлено: ' + transferInfo);
      }
    }
    
    console.log('📊 Найдено перемещений: ' + pendingTransfers.length);
    
    if (pendingTransfers.length === 0) {
      ui.alert('ℹ️ Нет ожидающих перемещений', 'Все перемещения уже обработаны или нет перемещений со статусом "В процессе"', ui.ButtonSet.OK);
      return;
    }
    
    // 🔥 ИСПРАВЛЕНИЕ: Безопасное формирование списка
    var transferList = '';
    for (var j = 0; j < pendingTransfers.length; j++) {
      transferList += pendingTransfers[j].info + '\n';
    }
    
    var response = ui.prompt(
      '🔄 Выполнение перемещения',
      'Выберите перемещение для выполнения:\n\n' + transferList + '\n\nВведите номер строки:',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (response.getSelectedButton() !== ui.Button.OK) {
      console.log('❌ Пользователь отменил выбор');
      return;
    }
    
    var selectedText = response.getResponseText();
    console.log('🔢 Пользователь ввел: "' + selectedText + '"');
    
    var selectedRow = extractRowNumber(selectedText);
    console.log('🔢 Извлечен номер строки: ' + selectedRow);
    
    var selectedTransfer = null;
    for (var k = 0; k < pendingTransfers.length; k++) {
      if (pendingTransfers[k].row == selectedRow) {
        selectedTransfer = pendingTransfers[k];
        break;
      }
    }
    
    if (!selectedTransfer) {
      console.log('❌ Перемещение не найдено для строки: ' + selectedRow);
      ui.alert('❌ Ошибка', 'Перемещение в строке ' + selectedRow + ' не найдено или уже обработано', ui.ButtonSet.OK);
      return;
    }
    
    console.log('🎯 Выбрано перемещение: строка ' + selectedTransfer.row);
    console.log('📋 Детали:', selectedTransfer.component, selectedTransfer.quantity, selectedTransfer.fromWarehouse, '→', selectedTransfer.toWarehouse);
    
    var confirmResponse = ui.alert(
      '✅ Подтверждение выполнения',
      'Выполнить перемещение?\n\n' + 
      'Комплектующее: ' + selectedTransfer.component + '\n' +
      'Количество: ' + selectedTransfer.quantity + ' ед.\n' +
      'Из: ' + selectedTransfer.fromWarehouse + '\n' +
      'В: ' + selectedTransfer.toWarehouse + '\n' +
      (selectedTransfer.purchaseId ? 'ID закупки: ' + selectedTransfer.purchaseId + '\n' : '') +
      '\nЭто действие нельзя отменить.',
      ui.ButtonSet.YES_NO
    );
    
    if (confirmResponse !== ui.Button.YES) {
      console.log('❌ Пользователь отменил выполнение');
      ui.alert('ℹ️ Отменено', 'Выполнение перемещения отменено', ui.ButtonSet.OK);
      return;
    }
    
    console.log('🚀 Запуск выполнения перемещения...');
    completeTransferSimple(selectedTransfer.row);
    
  } catch (error) {
    console.error('❌ Ошибка в showPendingTransfers:', error.toString());
    SpreadsheetApp.getUi().alert(
      '❌ Ошибка выполнения перемещения', 
      'Не удалось выполнить перемещение:\n\n' + error.toString(),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}



/**
 * ✅ ФУНКЦИЯ: ИЗВЛЕЧЕНИЕ НОМЕРА СТРОКИ
 */
function extractRowNumber(inputText) {
  try {
    if (/^\d+$/.test(inputText.trim())) {
      return parseInt(inputText.trim());
    }
    
    var match = inputText.match(/строка\s*(\d+)/i);
    if (match && match[1]) {
      return parseInt(match[1]);
    }
    
    var numbers = inputText.match(/\d+/g);
    if (numbers && numbers.length > 0) {
      return parseInt(numbers[0]);
    }
    
    return null;
    
  } catch (error) {
    console.error('❌ Ошибка извлечения номера:', error.toString());
    return null;
  }
}

/**
 * ✅ ИСПРАВЛЕННАЯ ФУНКЦИЯ: ВЫПОЛНЕНИЕ ПЕРЕМЕЩЕНИЯ С СОЗДАНИЕМ НОВЫХ ЗАПИСЕЙ
 */
function completeTransferSimple(row) {
  try {
    var transfersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TRANSFERS_SHEET);
    
    // 🔥 ИСПРАВЛЕНИЕ: Безопасное получение данных строки
    if (!transfersSheet) {
      throw new Error('Лист перемещений не найден');
    }
    
    var lastRow = transfersSheet.getLastRow();
    if (row > lastRow) {
      throw new Error('Строка ' + row + ' не существует в листе перемещений');
    }
    
    // Получаем данные строки перемещения
    var data = transfersSheet.getRange(row, 1, 1, 9).getValues()[0];
    
    // 🔥 ПРАВИЛЬНЫЕ ИНДЕКСЫ ДЛЯ ВАШЕЙ СТРУКТУРЫ:
    var transferId = data[0] || '';      // A - ID перемещения
    var purchaseId = data[1] || '';      // B - ID Закупки
    var date = data[2];                  // C - Дата
    var fromWarehouse = data[3] || '';   // D - От склада >
    var toWarehouse = data[4] || '';     // E - На склад
    var component = data[5] || '';       // F - Комплектующее
    var quantity = data[6] || 0;         // G - Количество
    var status = data[7] || '';          // H - Статус
    var notes = data[8] || '';           // I - Примечание
    
    console.log('🔄 Выполнение перемещения:', component, quantity, 'из', fromWarehouse, 'в', toWarehouse);
    console.log('📋 Используем существующий ID перемещения:', transferId);
    
    // Проверяем обязательные поля
    if (!component || !fromWarehouse || !toWarehouse || quantity <= 0) {
      throw new Error('Неполные данные перемещения: компонент=' + component + ', из=' + fromWarehouse + ', в=' + toWarehouse + ', количество=' + quantity);
    }
    
    // Проверяем статус
    if (status === 'Выполнено') {
      SpreadsheetApp.getUi().alert('ℹ️ Информация', 'Перемещение уже выполнено', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    // Определяем листы складов
    var fromSheet = fromWarehouse === 'Главный' 
      ? spreadsheet.getSheetByName(MAIN_WAREHOUSE_SHEET)
      : spreadsheet.getSheetByName(PRODUCTION_WAREHOUSE_SHEET);
    
    var toSheet = toWarehouse === 'Главный' 
      ? spreadsheet.getSheetByName(MAIN_WAREHOUSE_SHEET)
      : spreadsheet.getSheetByName(PRODUCTION_WAREHOUSE_SHEET);
    
    if (!fromSheet || !toSheet) {
      throw new Error('Склады не найдены: ' + fromWarehouse + ' или ' + toWarehouse);
    }
    
    // 1. СПИСЫВАЕМ С ИСХОДНОГО СКЛАДА
    console.log('📤 Списание с исходного склада...');
    
    // 🔥 ИСПРАВЛЕНИЕ: Правильный вызов функции с учетом новой структуры
    var fromResult = consumeMaterialsFIFOSpecific(fromWarehouse, component, purchaseId, quantity);
    if (!fromResult.success) {
      throw new Error('Не удалось списать материалы: ' + fromResult.message);
    }
    
    // Обновляем статус склада-источника
    updateWarehouseStatusDirect(fromSheet);
    Utilities.sleep(500); // Пауза для стабильности
    
    // 2. СОЗДАЕМ НОВУЮ ЗАПИСЬ НА ЦЕЛЕВОМ СКЛАДЕ
    console.log('📥 Создание новой записи на целевом складе...');
    
    if (toWarehouse === 'Главный') {
      addMaterialToMainWarehouse(component, purchaseId, quantity, transferId);
    } else {
      addMaterialToProductionWarehouse(component, purchaseId, quantity, transferId);
    }
    
    // 3. ОБНОВЛЯЕМ СТАТУС ПЕРЕМЕЩЕНИЯ
    transfersSheet.getRange(row, 8).setValue('Выполнено'); // Колонка H - Статус
    
    // 4. ФОРМИРУЕМ ОТЧЕТ ДЛЯ ПОЛЬЗОВАТЕЛЯ
    var message = '✅ ПЕРЕМЕЩЕНИЕ ВЫПОЛНЕНО\n\n';
    message += '📦 ' + component + ' × ' + quantity.toFixed(2) + ' ед.\n';
    message += '🔄 ' + fromWarehouse + ' → ' + toWarehouse + '\n';
    if (purchaseId && purchaseId.toString().trim() !== '') {
      message += '🏷️ ID закупки: ' + purchaseId + '\n';
    }
    message += '📋 ID перемещения: ' + transferId + '\n\n';
    message += '✅ Использован существующий ID перемещения для лучшего отслеживания';

    console.log('✅ Перемещение успешно выполнено с ID:', transferId);
    SpreadsheetApp.getUi().alert('✅ Перемещение выполнено!', message, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (error) {
    console.error('❌ Ошибка выполнения перемещения: ' + error.toString());
    SpreadsheetApp.getUi().alert('❌ Ошибка', 'Не удалось выполнить перемещение:\n\n' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
    throw error;
  }
}



/**
 * ✅ ИСПРАВЛЕННАЯ ФУНКЦИЯ: ОБНОВЛЕНИЕ СТАТУСА ДЛЯ ОДНОЙ СТРОКИ С МИН. ЗАПАСОМ ИЗ СПРАВОЧНИКА
 */
function updateSingleRowStatus(sheet, row) {
  try {
    var data = sheet.getRange(row, 1, 1, 6).getValues()[0];
    var component = data[0]; // A - Комплектующие
    var purchaseId = data[1]; // B - ID Закупки
    var receiptDate = data[2]; // C - Дата поступления
    var currentStock = data[3] || 0; // D - Остаток
    
    if (!component) return;
    
    // 🎯 ПОЛУЧАЕМ МИН. ЗАПАС ИЗ СПРАВОЧНИКА
    var warehouseType = sheet.getName() === 'Склад_Главный' ? 'Главный' : 'Производство';
    var minStock = getMinStockFromReference(component, warehouseType);
    
    // Рассчитываем общий остаток по компоненту
    var componentTotals = calculateComponentTotals(sheet);
    var totalStock = componentTotals[component] || 0;
    
    var status = '';
    var backgroundColor = '';
    var statusType = '';
    
    // ИНФОРМАТИВНАЯ ЛОГИКА СТАТУСОВ:
    if (currentStock === 0) {
      status = '❌ Отсутствует';
      backgroundColor = '#FFCCCC';
      statusType = 'absent';
    } else if (minStock > 0 && totalStock < minStock) {
      var deficit = minStock - totalStock;
      status = '⚠️ Мало: ' + currentStock + '/' + minStock + ' ед. | -' + deficit;
      backgroundColor = '#FFF2CC';
      statusType = 'low';
    } else if (currentStock === 1) {
      status = '✅ В наличии: ' + currentStock + ' ед. | Последняя';
      backgroundColor = '#FFF2CC';
      statusType = 'last';
    } else if (currentStock <= 5) {
      status = '✅ В наличии: ' + currentStock + ' ед. | Мало осталось';
      backgroundColor = '#CCFFCC';
      statusType = 'few';
    } else {
      status = '✅ В наличии: ' + currentStock + ' ед.';
      backgroundColor = '#CCFFCC';
      statusType = 'normal';
    }
    
    // Обновляем статус (столбец F)
    var statusCell = sheet.getRange(row, 6);
    statusCell.setValue(status);
    statusCell.setBackground(backgroundColor);
    
    // 🔥 ОБНОВЛЯЕМ ПОДСКАЗКУ
    var tooltip = createTooltip(component, purchaseId, receiptDate, currentStock, minStock, totalStock, statusType);
    setCellNote(statusCell, tooltip);
    
  } catch (error) {
    console.error('❌ Ошибка в updateSingleRowStatus: ' + error.toString());
  }
}
/**
 * ✅ АДАПТИРОВАННАЯ ФУНКЦИЯ: ПОЛУЧЕНИЕ ОБЩЕГО ОСТАТКА КОМПОНЕНТА НА СКЛАДЕ (ПОД ВАШУ СТРУКТУРУ)
 */
function getTotalStockForComponentDirect(sheet, component) {
  try {
    if (!sheet) {
      console.error('❌ Лист не передан');
      return 0;
    }
    
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return 0;
    }
    
    // 🔥 ПРАВИЛЬНАЯ СТРУКТУРА ДЛЯ ВАШИХ СКЛАДОВ:
    // A: ID перемещения, B: ID Закупки, C: Дата поступления, D: Комплектующие, E: Остаток, F: Статус
    var data = sheet.getRange('A2:F' + lastRow).getValues();
    var totalStock = 0;
    
    console.log('🔍 Поиск остатков для:', component, 'на складе:', sheet.getName());
    
    for (var i = 0; i < data.length; i++) {
      // 🔥 ПРАВИЛЬНЫЕ ИНДЕКСЫ:
      var idПеремещения = data[i][0];    // A - ID перемещения
      var idЗакупки = data[i][1];        // B - ID Закупки
      var датаПоступления = data[i][2];   // C - Дата поступления
      var rowComponent = data[i][3];     // D - Комплектующие ✅
      var currentStock = data[i][4] || 0; // E - Остаток ✅ (индекс 4!)
      var статус = data[i][5];           // F - Статус
      
      // Проверяем совпадение по названию комплектующего
      if (rowComponent && rowComponent.toString().trim().toLowerCase() === component.toString().trim().toLowerCase()) {
        totalStock += Number(currentStock);
        console.log('✅ Найден остаток в строке', i + 2, ':', currentStock, 'ед.');
      }
    }
    
    totalStock = roundToTwo(totalStock);
    
    console.log('📊 Общий остаток ' + component + ' на складе ' + sheet.getName() + ': ' + totalStock + ' ед.');
    return totalStock;
    
  } catch (error) {
    console.error('❌ Ошибка в getTotalStockForComponentDirect: ' + error.toString());
    return 0;
  }
}

/**
 * ✅ УПРОЩЕННАЯ ФУНКЦИЯ: ПОКАЗАТЬ ВСЕ СКРЫТЫЕ СТРОКИ НА ВСЕХ СКЛАДАХ
 */
function showAllHiddenRowsAllWarehouses() {
  try {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = [
      spreadsheet.getSheetByName(MAIN_WAREHOUSE_SHEET),
      spreadsheet.getSheetByName(PRODUCTION_WAREHOUSE_SHEET)
    ];
    
    var totalShown = 0;
    
    for (var s = 0; s < sheets.length; s++) {
      var sheet = sheets[s];
      if (sheet) {
        console.log('🔍 Поиск скрытых строк на: ' + sheet.getName());
        
        var lastRow = sheet.getLastRow();
        var sheetShownCount = 0;
        
        // Если в таблице только заголовок или нет данных
        if (lastRow >= 2) {
          for (var row = 2; row <= lastRow; row++) {
            if (sheet.isRowHiddenByUser(row)) {
              sheet.showRows(row);
              sheetShownCount++;
              console.log('👁️ Показана строка ' + row + ' на ' + sheet.getName());
            }
          }
        }
        
        console.log('✅ Показано скрытых строк на ' + sheet.getName() + ': ' + sheetShownCount);
        totalShown += sheetShownCount;
      }
    }
    
    console.log('✅ Всего показано скрытых строк: ' + totalShown);
    
    // Показываем уведомление пользователю
    if (totalShown > 0) {
      SpreadsheetApp.getUi().alert(
        '👁️ Показаны скрытые строки', 
        'Было показано ' + totalShown + ' скрытых строк на всех складах',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    } else {
      SpreadsheetApp.getUi().alert(
        'ℹ️ Нет скрытых строк', 
        'На складах нет скрытых строк для отображения',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    
    return totalShown;
    
  } catch (error) {
    console.error('❌ Ошибка в showAllHiddenRowsAllWarehouses: ' + error.toString());
    SpreadsheetApp.getUi().alert(
      '❌ Ошибка', 
      'Не удалось показать скрытые строки:\n' + error.toString(),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return 0;
  }
}

/**
 * ✅ ФУНКЦИЯ: ОБНОВИТЬ СТАТУСЫ И СКРЫТЬ НУЛЕВЫЕ ОСТАТКИ НА ВСЕХ СКЛАДАХ
 */
function updateStatusAndHideZeroStockAllWarehouses() {
  try {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = [
      spreadsheet.getSheetByName(MAIN_WAREHOUSE_SHEET),
      spreadsheet.getSheetByName(PRODUCTION_WAREHOUSE_SHEET)
    ];
    
    var totalHidden = 0;
    
    for (var s = 0; s < sheets.length; s++) {
      var sheet = sheets[s];
      if (sheet) {
        console.log('🔄 Обновление статусов и скрытие нулевых остатков на: ' + sheet.getName());
        
        // Обновляем статусы
        updateWarehouseStatusDirect(sheet, true);
        
        // Скрываем нулевые остатки
        var hiddenCount = hideZeroStockRowsOnSheet(sheet);
        totalHidden += hiddenCount;
      }
    }
    
    console.log('✅ Все операции завершены. Скрыто строк: ' + totalHidden);
    return totalHidden;
    
  } catch (error) {
    console.error('❌ Ошибка в updateStatusAndHideZeroStockAllWarehouses: ' + error.toString());
    return 0;
  }
}

/**
 * ✅ ФУНКЦИЯ: ПОКАЗАТЬ ВСЕ СКРЫТЫЕ СТРОКИ (для отладки или ручного управления)
 */
function showAllHiddenRows(sheet) {
  try {
    var lastRow = sheet.getLastRow();
    var shownCount = 0;
    
    for (var row = 2; row <= lastRow; row++) {
      if (sheet.isRowHiddenByUser(row)) {
        sheet.showRows(row);
        shownCount++;
        console.log('👁️ Показана строка ' + row);
      }
    }
    
    console.log('✅ Показано скрытых строк: ' + shownCount);
    
  } catch (error) {
    console.error('❌ Ошибка в showAllHiddenRows: ' + error.toString());
  }
}


/**
 * ✅ НОВАЯ ФУНКЦИЯ: УДАЛЕНИЕ МАТЕРИАЛА С КОНКРЕТНОГО СКЛАДА (ПО ID ПЕРЕМЕЩЕНИЯ)
 */
function removeMaterialFromWarehouse(sheet, component, purchaseId, quantity, transferId) {
  try {
    console.log('🗑️ Удаление материала со склада:', sheet.getName(), component, quantity, 'ID перемещения:', transferId);
    
    var data = sheet.getDataRange().getValues();
    
    // Ищем запись по ID перемещения (самый точный способ)
    for (var i = 1; i < data.length; i++) {
      var rowTransferId = data[i][0];    // A - ID перемещения
      var rowPurchaseId = data[i][1];    // B - ID Закупки
      var rowComponent = data[i][3];     // D - Комплектующие
      var currentStock = data[i][4];     // E - Остаток
      
      // Ищем по ID перемещения (самый точный) или по комбинации компонент + ID закупки
      if ((rowTransferId && rowTransferId.toString().trim() === transferId.toString().trim()) ||
          (rowComponent && rowComponent.toString().trim() === component.toString().trim() && 
           rowPurchaseId && rowPurchaseId.toString().trim() === purchaseId.toString().trim())) {
        
        console.log('✅ Найдена запись для удаления в строке', i + 1, 'Остаток:', currentStock);
        
        if (currentStock >= quantity) {
          // Уменьшаем остаток
          var newStock = currentStock - quantity;
          sheet.getRange(i + 1, 5).setValue(newStock); // E - Остаток
          
          // Если остаток стал нулевым, удаляем строку
          if (newStock <= 0) {
            sheet.deleteRow(i + 1);
            console.log('🗑️ Удалена строка', i + 1, 'так как остаток стал нулевым');
          } else {
            // Обновляем статус строки
            updateSingleRowStatus(sheet, i + 1);
          }
          
          console.log('✅ Материал удален со склада. Новый остаток:', newStock);
          return true;
        } else {
          console.log('❌ Недостаточно материала для удаления. Доступно:', currentStock, 'Требуется:', quantity);
          return false;
        }
      }
    }
    
    console.log('❌ Запись не найдена для удаления');
    return false;
    
  } catch (error) {
    console.error('❌ Ошибка в removeMaterialFromWarehouse:', error.toString());
    return false;
  }
}

/**
 * ✅ ВСПОМОГАТЕЛЬНАЯ ФУНКЦИЯ: ПОЛУЧЕНИЕ ЛИСТА СКЛАДА ПО НАЗВАНИЮ
 */
function getWarehouseSheetByName(warehouseName) {
  try {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    if (warehouseName === 'Главный') {
      return spreadsheet.getSheetByName(MAIN_WAREHOUSE_SHEET);
    } else if (warehouseName === 'Производство') {
      return spreadsheet.getSheetByName(PRODUCTION_WAREHOUSE_SHEET);
    } else if (warehouseName === 'Поставщик') {
      // Для поставщика возвращаем null, так как это виртуальный склад
      return null;
    }
    
    return null;
    
  } catch (error) {
    console.error('❌ Ошибка в getWarehouseSheetByName:', error.toString());
    return null;
  }
}

/**
 * ✅ ОБНОВЛЕННАЯ ФУНКЦИЯ: ПОИСК ЗАПИСИ НА СКЛАДЕ ПО ID ПЕРЕМЕЩЕНИЯ
 */
function findWarehouseRecordByTransferId(sheet, transferId) {
  try {
    if (!sheet || !transferId) return null;
    
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return null;
    
    //  🔥 ИСПРАВЛЕННЫЕ ИНДЕКСЫ ПО НОВОЙ СТРУКТУРЕ:
    // A: ID перемещения (0), B: ID Закупки (1), C: Дата (2), D: Комплектующие (3), E: Остаток (4), F: Статус (5)
    var data = sheet.getRange('A2:F' + lastRow).getValues();
    
    for (var i = 0; i < data.length; i++) {
      var rowTransferId = data[i][0]; // A - ID перемещения
      
      if (rowTransferId && rowTransferId.toString().trim() === transferId.toString().trim()) {
        return {
          row: i + 2,
          transferId: rowTransferId,
          purchaseId: data[i][1], // B - ID Закупки
          component: data[i][3], // D - Комплектующие
          currentStock: data[i][4] // E - Остаток
        };
      }
    }
    
    return null;
    
  } catch (error) {
    console.error('❌ Ошибка в findWarehouseRecordByTransferId:', error.toString());
    return null;
  }
}

/**
 * ✅ НОВАЯ ФУНКЦИЯ: ПОИСК ВСЕХ ЗАПИСЕЙ НА СКЛАДЕ ПО ID ЗАКУПКИ
 */
function findAllWarehouseRecordsByPurchaseId(sheet, purchaseId) {
  try {
    if (!sheet || !purchaseId) return [];
    
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    
    var data = sheet.getRange('A2:F' + lastRow).getValues();
    var records = [];
    
    for (var i = 0; i < data.length; i++) {
      var rowPurchaseId = data[i][1]; // B - ID Закупки
      
      if (rowPurchaseId && rowPurchaseId.toString().trim() === purchaseId.toString().trim()) {
        records.push({
          row: i + 2,
          transferId: data[i][0], // A - ID перемещения
          purchaseId: rowPurchaseId,
          component: data[i][3], // D - Комплектующие
          currentStock: data[i][4] // E - Остаток
        });
      }
    }
    
    return records;
    
  } catch (error) {
    console.error('❌ Ошибка в findAllWarehouseRecordsByPurchaseId:', error.toString());
    return [];
  }
}

/**
 * ✅ ОБНОВЛЕНИЕ ДАТ В ПЛАНИРОВАНИИ
 */
function updateDates(sheet, row) {
  try {
    var product = sheet.getRange(row, 4).getValue(); // D - Выбор товара
    if (!product || product === 'Выбор товара') {
      clearRow(sheet, row);
      return;
    }
    
    var scheduleSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('График_смен');
    var dates = getUniqueDatesForProduct(product, scheduleSheet);
    
    var dateCell = sheet.getRange(row, 5); // E - Дата смены
    if (dates.length > 0) {
      var dateRule = SpreadsheetApp.newDataValidation().requireValueInList(dates, true).build();
      dateCell.setDataValidation(dateRule);
      dateCell.setValue('');
    } else {
      dateCell.setDataValidation(null);
      dateCell.setValue('Нет доступных дат для ' + product);
    }
    
    // Очищаем столбец F (Выбор смены) и G (Остаток)
    sheet.getRange(row, 6).setDataValidation(null);
    sheet.getRange(row, 6).setValue('');
    sheet.getRange(row, 7).setValue('');
    
  } catch (error) {
    console.error('Ошибка в updateDates: ' + error.toString());
  }
}

/**
 * ✅ ОБНОВЛЕНИЕ СМЕН В ПЛАНИРОВАНИИ
 */
function updateShifts(sheet, row) {
  try {
    var product = sheet.getRange(row, 4).getValue(); // D - Выбор товара
    var date = sheet.getRange(row, 5).getValue();    // E - Дата смены
    
    if (!product || product === 'Выбор товара' || !date) {
      sheet.getRange(row, 6).setValue('Сначала выберите товар и дату');
      return;
    }
    
    var dateFormatted = Utilities.formatDate(date, Session.getScriptTimeZone(), "dd.MM.yy");
    var scheduleSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('График_смен');
    var shifts = getShiftsForProductAndDate(product, dateFormatted, scheduleSheet);
    
    var shiftCell = sheet.getRange(row, 6); // F - Выбор смены с этим товаром
    if (shifts.length > 0) {
      var shiftRule = SpreadsheetApp.newDataValidation().requireValueInList(shifts, true).build();
      shiftCell.setDataValidation(shiftRule);
      shiftCell.setValue('');
      // Обновляем отображение остатка
      if (shifts[0]) {
        updateBalanceDisplay(sheet, row, shifts[0]);
      }
    } else {
      shiftCell.setDataValidation(null);
      shiftCell.setValue('Нет смен для ' + product + ' на ' + dateFormatted);
      sheet.getRange(row, 7).setValue(''); // G - Остаток
    }
    
  } catch (error) {
    console.error('Ошибка в updateShifts: ' + error.toString());
  }
}

/**
 * ✅ ОБНОВЛЕННАЯ ФУНКЦИЯ: ОБНОВЛЕНИЕ СТАТУСА С ОПТИМИЗАЦИЕЙ
 */
function updateStatus(sheet, row) {
  try {
    var quantity = sheet.getRange(row, 8).getValue(); // H - Кол-во коробок
    var balance = sheet.getRange(row, 7).getValue();  // G - Остаток
    var statusCell = sheet.getRange(row, 9);          // I - Статус
    var product = sheet.getRange(row, 4).getValue();  // D - Товар
    var shiftInfo = sheet.getRange(row, 6).getValue(); // F - Выбор смены

    console.log('🔄 Обновление статуса для строки ' + row + ': товар=' + product + ', количество=' + quantity);

    if (!quantity || quantity === 0) {
      statusCell.setValue('🟡 В работе');
      statusCell.setBackground('#FFF2CC'); // Желтый
      return;
    }

    // Проверяем достаточно ли остатка
    if (quantity > balance) {
      statusCell.setValue('❌ Превышение! Доступно: ' + balance);
      statusCell.setBackground('#FFCCCC'); // Красный
      return;
    }

    // Проверяем что выбрана смена
    if (!shiftInfo || !shiftInfo.includes(' | Остаток:')) {
      statusCell.setValue('❌ Не выбрана смена');
      statusCell.setBackground('#FFCCCC');
      return;
    }

    //  🔵 ЭТАП 1: Проверка материалов
    statusCell.setValue('🔵 Проверка материалов...');
    statusCell.setBackground('#CCE5FF'); // Синий
    SpreadsheetApp.flush();
    
    console.log('🔵 Этап 1: Проверка материалов для ' + product + ' × ' + quantity + ' кор.');
    
    // Получаем количество товаров в групповой коробке
    var itemsPerBox = getItemsPerBox(product);
    var totalItems = quantity * itemsPerBox;
    
    // 🔥 ОБНОВЛЕНО: Используем улучшенную проверку с уведомлениями
    var availabilityCheck = checkMaterialsAvailabilityQuickSimple(product, totalItems);

    if (availabilityCheck.error) {
      statusCell.setValue('❌ Ошибка проверки материалов');
      statusCell.setBackground('#FFCCCC');
      return;
    }
    
    if (!availabilityCheck.isAvailable) {
      statusCell.setValue('❌ Недостаточно материалов');
      statusCell.setBackground('#FFCCCC');
      return;
    }

    //  🔄 ЭТАП 2: Списание материалов
    statusCell.setValue('🔄 Списание материалов...');
    statusCell.setBackground('#FFE0B2'); // Оранжевый
    SpreadsheetApp.flush();
    
    console.log('🔄 Этап 2: Списание материалов для ' + totalItems + ' шт.');
    
    // Выполняем списание материалов
    var consumptionResult = executeMaterialsConsumptionOptimized(sheet, row, product, quantity, totalItems);
    
    if (!consumptionResult.success) {
      statusCell.setValue('❌ Ошибка списания материалов: ' + (consumptionResult.error || 'Неизвестная ошибка'));
      statusCell.setBackground('#FFCCCC');
      return;
    }

    // ✅ ЭТАП 3: Устанавливаем статус "Выполнено"
    statusCell.setValue('✅ Выполнено');
    statusCell.setBackground('#CCFFCC'); // Зеленый
    SpreadsheetApp.flush();
    
    // 🔥 ОПТИМИЗАЦИЯ: Обновляем списания ТОЛЬКО ДЛЯ ИЗМЕНЕННЫХ СТРОК
    updateScheduleWriteOffs();
    
    console.log('✅ Списания обновлены в графике смен');

    // Проверяем лимит поставки
    checkDeliveryLimit(sheet);
    
  } catch (error) {
    console.error('❌ Критическая ошибка в updateStatus: ' + error.toString());
    sheet.getRange(row, 9).setValue('❌ Ошибка обработки');
    sheet.getRange(row, 9).setBackground('#FFCCCC');
  }
}

/**
 * 🔍 ОПТИМИЗИРОВАННАЯ ФУНКЦИЯ: ОБНОВЛЕНИЕ СПИСАНИЙ ТОЛЬКО ДЛЯ ИЗМЕНЕННЫХ СТРОК
 */
function updateScheduleWriteOffs() {
  try {
    var planningSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Планирование');
    var scheduleSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('График_смен');
    
    console.log('=== ОПТИМИЗИРОВАННОЕ ОБНОВЛЕНИЕ СПИСАНИЙ ===');
    
    // Получаем данные из планирования
    var planningData = planningSheet.getRange('A2:J' + planningSheet.getLastRow()).getValues();
    console.log('📝 Строк в планировании:', planningData.length);
    
    // Создаем объект для суммирования списаний по сменам
    var writeoffs = {};
    var affectedRows = new Set(); // Для отслеживания строк, которые нужно обновить
    
    // Суммируем списания из планирования
    for (var i = 0; i < planningData.length; i++) {
      var rowNum = i + 2;
      var date = planningData[i][0];           // A - Дата поставки
      var deliveryNumber = planningData[i][1]; // B - Поставки
      var warehouse = planningData[i][2];      // C - Склад
      var product = planningData[i][3];        // D - Выбор товара
      var shiftDate = planningData[i][4];      // E - Дата смены
      var shiftInfo = planningData[i][5];      // F - Выбор смены с этим товаром
      var balance = planningData[i][6];        // G - Остаток
      var quantity = planningData[i][7] || 0;  // H - Кол-во коробок
      var status = planningData[i][8];         // I - Статус
      var limitControl = planningData[i][9];   // J - Контроль лимита
      
      // Учитываем только строки с "Выполнено" статусом и заполненными данными
      if (status === '✅ Выполнено' && 
          product && product !== 'Выбор товара' && 
          shiftDate && Object.prototype.toString.call(shiftDate) === '[object Date]' && 
          shiftInfo && shiftInfo.includes(' | Остаток:') && 
          quantity > 0) {
        
        // Извлекаем работника из информации о смене
        var worker = shiftInfo.split(' | ')[0];
        
        // Форматируем дату для ключа
        var dateFormatted = Utilities.formatDate(shiftDate, Session.getScriptTimeZone(), "dd.MM.yy");
        
        // Создаем уникальный ключ: дата|работник|товар
        var key = dateFormatted + '|' + worker + '|' + product;
        
        if (!writeoffs[key]) {
          writeoffs[key] = 0;
        }
        writeoffs[key] += quantity;
        
        console.log('✅ СПИСАНИЕ: ' + key + ' = ' + quantity + ' кор.');
      }
    }
    
    console.log('📦 Списаний для обновления: ' + Object.keys(writeoffs).length);
    
    // Если нет списаний для обновления, выходим
    if (Object.keys(writeoffs).length === 0) {
      console.log('ℹ️ Нет списаний для обновления');
      return;
    }
    
    // 🔥 ОПТИМИЗАЦИЯ: Получаем только строки графика смен, которые могут быть затронуты
    var scheduleData = scheduleSheet.getRange('B2:D' + scheduleSheet.getLastRow()).getValues();
    var scheduleBalances = scheduleSheet.getRange('H2:H' + scheduleSheet.getLastRow()).getValues(); // Только столбец H (Списано)
    
    console.log('👥 Проверяем ' + scheduleData.length + ' строк в графике смен');
    
    var rowsToUpdate = [];
    var updatedRows = 0;
    
    for (var i = 0; i < scheduleData.length; i++) {
      var rowNum = i + 2;
      var date = scheduleData[i][0];    // B - Дата
      var worker = scheduleData[i][1];  // C - Работник
      var product = scheduleData[i][2]; // D - Товар
      
      if (date && worker && product) {
        var dateFormatted = Utilities.formatDate(date, Session.getScriptTimeZone(), "dd.MM.yy");
        var key = dateFormatted + '|' + worker + '|' + product;
        
        var writeoff = writeoffs[key] || 0;
        var currentWriteoff = scheduleBalances[i][0] || 0;
        
        //  🔥 ОБНОВЛЯЕМ ТОЛЬКО ЕСЛИ ЗНАЧЕНИЕ ИЗМЕНИЛОСЬ
        if (writeoff !== currentWriteoff) {
          rowsToUpdate.push({
            row: rowNum,
            writeoff: writeoff,
            key: key
          });
        }
      }
    }
    
    console.log('📊 Строк для обновления: ' + rowsToUpdate.length);
    
    // 🔥 ОБНОВЛЯЕМ СПИСАНИЯ БАТЧЕМ
    if (rowsToUpdate.length > 0) {
      for (var j = 0; j < rowsToUpdate.length; j++) {
        var update = rowsToUpdate[j];
        scheduleSheet.getRange(update.row, 8).setValue(update.writeoff); // Столбец H - Списано
        console.log('📝 ОБНОВЛЕНО: ' + update.key + ' -> ' + update.writeoff + ' кор. (строка ' + update.row + ')');
        updatedRows++;
        
        // 🔥 ПЕРЕСЧИТЫВАЕМ ОСТАТОК ТОЛЬКО ДЛЯ ИЗМЕНЕННОЙ СТРОКИ
        updateBalanceForRow(scheduleSheet, update.row);
      }
    }
    
    console.log('✅ Обновлено строк: ' + updatedRows);
    console.log('=== ОПТИМИЗИРОВАННОЕ ОБНОВЛЕНИЕ ЗАВЕРШЕНО ===');
    
  } catch (error) {
    console.error('❌ Ошибка в updateScheduleWriteOffs: ' + error.toString());
  }
}

/**
 * ✅ КОНТРОЛЬ ЛИМИТА 20 КОРОБОК
 */
function checkDeliveryLimit(sheet) {
  try {
    var data = sheet.getRange('B2:H' + sheet.getLastRow()).getValues();
    
    var deliverySums = {};
    
    // Считаем сумму по каждому номеру поставки
    for (var i = 0; i < data.length; i++) {
      var deliveryNumber = data[i][0]; // Столбец B - Поставки
      var quantity = data[i][6] || 0;  // Столбец H - Кол-во коробок
      
      if (deliveryNumber) {
        if (!deliverySums[deliveryNumber]) {
          deliverySums[deliveryNumber] = 0;
        }
        deliverySums[deliveryNumber] += quantity;
      }
    }
    
    // Обновляем статусы лимитов
    for (var i = 0; i < data.length; i++) {
      var deliveryNumber = data[i][0];
      if (deliveryNumber && deliverySums[deliveryNumber]) {
        var limitCell = sheet.getRange(i + 2, 10); // Столбец J - Контроль лимита
        var total = deliverySums[deliveryNumber];
        
        if (total > 20) {
          limitCell.setValue('⚠️ Лимит 20 кор. превышен!');
        } else {
          limitCell.setValue('ОК - ' + total + '/20 кор.');
        }
      }
    }
    
    console.log('✅ Лимиты проверены');
    
  } catch (error) {
    console.error('Ошибка в checkDeliveryLimit: ' + error.toString());
  }
}

/**
 * ✅ ОБНОВЛЕННАЯ ФУНКЦИЯ: РАСХОД МАТЕРИАЛОВ ПРИ ВЫПОЛНЕНИИ ПОСТАВКИ С ЗАЩИТОЙ ОТ ДУБЛИРОВАНИЯ
 */
function updateMaterialsConsumption(sheet, row) {
  try {
    var status = sheet.getRange(row, 9).getValue(); // I - Статус
    var product = sheet.getRange(row, 4).getValue(); // D - Товар
    var quantity = sheet.getRange(row, 8).getValue(); // H - Кол-во коробок
    
    // ✅ ДОПОЛНИТЕЛЬНАЯ ПРОВЕРКА ЗАЩИТЫ ОТ ДУБЛИРОВАНИЯ
    var now = new Date().getTime();
    if (lastProcessedPlanning.row === row && 
        lastProcessedPlanning.product === product &&
        lastProcessedPlanning.quantity === quantity &&
        (now - lastProcessedPlanning.timestamp) < 3000) {
      console.log('🛑 Защита: пропускаем повторное списание для строки ' + row);
      return;
    }
    
    if (status === '✅ Выполнено' && product && product !== 'Выбор товара' && quantity > 0) {
      console.log('⚡ Быстрая проверка материалов для: ' + product + ' × ' + quantity + ' кор.');
      
      var itemsPerBox = getItemsPerBox(product);
      var totalItems = quantity * itemsPerBox;
      
      //  🔥 ИСПОЛЬЗУЕМ ОБНОВЛЕННУЮ ФУНКЦИЮ С ПЕРЕДАЧЕЙ ДАННЫХ ИЗ ПЛАНИРОВАНИЯ
      var result = executeMaterialsConsumptionOptimized(sheet, row, product, quantity, totalItems);
      
      if (!result.success) {
        console.error('❌ Ошибка списания материалов: ' + result.error);
        sheet.getRange(row, 9).setValue('❌ Ошибка списания');
        sheet.getRange(row, 9).setBackground('#FFCCCC');
      } else {
        console.log('✅ Данные успешно переданы в журнал списаний');
      }
    }
    
  } catch (error) {
    console.error('❌ Ошибка в updateMaterialsConsumption: ' + error.toString());
  }
}

/**
 * ✅ ПРОВЕРКА УДАЛЕНИЯ СТРОКИ В ПЛАНИРОВАНИИ
 */
function isPlanningRowDeleted(sheet, row) {
  try {
    var data = sheet.getRange(row, 1, 1, 10).getValues()[0]; // A:J - все поля
    
    var product = data[3]; // D - Выбор товара
    var shiftDate = data[4]; // E - Дата смены
    var shiftInfo = data[5]; // F - Выбор смены
    var quantity = data[7]; // H - Кол-во коробок
    var status = data[8]; // I - Статус
    
    // Проверяем, является ли строка пустой или заголовком
    var isEmptyRow = (!product || product === 'Выбор товара') && 
                     !shiftDate && 
                     (!shiftInfo || shiftInfo === 'Выбор смены с этим товаром') && 
                     !quantity && 
                     (!status || status === 'Статус');
    
    if (isEmptyRow) {
      console.log('🔄 Обнаружена пустая строка в Планировании: ' + row);
      return true;
    }
    
    return false;
    
  } catch (error) {
    console.error('Ошибка в isPlanningRowDeleted: ' + error.toString());
    return false;
  }
}

/**
 * ✅ ПОЛУЧЕНИЕ УНИКАЛЬНЫХ ДАТ ДЛЯ ТОВАРА
 */
function getUniqueDatesForProduct(product, scheduleSheet) {
  var data = scheduleSheet.getRange('B2:D' + scheduleSheet.getLastRow()).getValues();
  var balances = scheduleSheet.getRange('G2:G' + scheduleSheet.getLastRow()).getValues();
  
  var dates = [];
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] && data[i][2] === product && balances[i][0] > 0) {
      var date = Utilities.formatDate(data[i][0], Session.getScriptTimeZone(), "dd.MM.yy");
      if (dates.indexOf(date) === -1) {
        dates.push(date);
      }
    }
  }
  return dates.sort();
}

/**
 * ✅ ПОЛУЧЕНИЕ СМЕН ДЛЯ ТОВАРА И ДАТЫ
 */
function getShiftsForProductAndDate(product, date, scheduleSheet) {
  try {
    var data = scheduleSheet.getRange('B2:D' + scheduleSheet.getLastRow()).getValues();
    var workers = scheduleSheet.getRange('C2:C' + scheduleSheet.getLastRow()).getValues();
    var balances = scheduleSheet.getRange('G2:G' + scheduleSheet.getLastRow()).getValues();
    
    var shifts = [];
    for (var i = 0; i < data.length; i++) {
      if (data[i][0] && Object.prototype.toString.call(data[i][0]) === '[object Date]') {
        var rowDate = Utilities.formatDate(data[i][0], Session.getScriptTimeZone(), "dd.MM.yy");
        var rowProduct = data[i][2];
        var rowBalance = balances[i][0];
        
        if (rowProduct === product && rowDate === date && rowBalance > 0) {
          var shift = workers[i][0] + " | Остаток: " + rowBalance + " кор.";
          shifts.push(shift);
        }
      }
    }
    return shifts;
    
  } catch (error) {
    console.error('Ошибка в getShifts: ' + error.toString());
    return [];
  }
}

/**
 * ✅ ОБНОВЛЕНИЕ ОТОБРАЖЕНИЯ ОСТАТКА
 */
function updateBalanceDisplay(sheet, row, shiftInfo) {
  var balanceMatch = shiftInfo.match(/Остаток: (\d+\.?\d*) кор\./);
  if (balanceMatch) {
    var balance = parseFloat(balanceMatch[1]);
    sheet.getRange(row, 7).setValue(balance); // G - Остаток
  }
}

/**
 * 🚀 ИСПРАВЛЕННАЯ ФУНКЦИЯ: ПОЛУЧЕНИЕ КОЛИЧЕСТВА ТОВАРОВ В КОРОБКЕ ИЗ ВАШЕГО СПРАВОЧНИКА
 */
function getItemsPerBox(productName) {
  try {
    console.log('🔍 Поиск количества в коробке для товара: "' + productName + '"');
    
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var productCatalogSheet = spreadsheet.getSheetByName('Справочник_товаров');
    
    if (!productCatalogSheet) {
      console.error('❌ Лист "Справочник_товаров" не найден!');
      return 16; // Значение по умолчанию
    }
    
    var lastRow = productCatalogSheet.getLastRow();
    if (lastRow <= 1) {
      console.error('❌ Справочник товаров пуст!');
      return 16;
    }
    
    // 🔥 ПРАВИЛЬНЫЕ СТОЛБЦЫ ИЗ ВАШЕЙ СТРУКТУРЫ:
    // A - Код (пират. черн.)
    // B - Название товара  
    // C - Групповая коробка
    var data = productCatalogSheet.getRange('A2:C' + lastRow).getValues();
    
    console.log('📋 Всего товаров в справочнике: ' + data.length);
    
    for (var i = 0; i < data.length; i++) {
      var code = data[i][0]; // A - Код (пират. черн.)
      var name = data[i][1]; // B - Название товара
      var itemsPerBox = data[i][2]; // C - Групповая коробка
      
      // 🔥 ИЩЕМ ПО КОДУ (столбец A)
      if (code && code.toString().trim().toLowerCase() === productName.toString().trim().toLowerCase()) {
        console.log('✅ Найдено в справочнике: ' + code + ' - ' + itemsPerBox + ' шт. в коробке');
        return Number(itemsPerBox);
      }
      
      // 🔥 ДОПОЛНИТЕЛЬНО ИЩЕМ ПО НАЗВАНИЮ (столбец B)
      if (name && name.toString().trim().toLowerCase() === productName.toString().trim().toLowerCase()) {
        console.log('✅ Найдено в справочнике по названию: ' + name + ' - ' + itemsPerBox + ' шт. в коробке');
        return Number(itemsPerBox);
      }
    }
    
    // 🔥 ЕСЛИ НЕ НАШЛИ - ВЫВОДИМ ДЛЯ ОТЛАДКИ ВСЕ ТОВАРЫ
    console.log('❌ Товар "' + productName + '" не найден в справочнике. Доступные товары:');
    for (var j = 0; j < Math.min(data.length, 10); j++) {
      if (data[j][0]) {
        console.log('   - ' + data[j][0] + ' (' + data[j][1] + ') - ' + data[j][2] + ' шт.');
      }
    }
    
    console.log('⚠️ Используем значение по умолчанию: 16');
    return 16;
    
  } catch (error) {
    console.error('❌ Ошибка получения количества в коробке: ' + error.toString());
    return 16;
  }
}

/**
 * 🚀 ФУНКЦИЯ: УПРОЩЕННАЯ ПРОВЕРКА МАТЕРИАЛОВ (ИСПРАВЛЕННАЯ)
 */
function checkMaterialsAvailabilityQuickSimple(product, totalItems) {
  try {
    var availability = checkMaterialsAvailabilityQuick(product, totalItems);
    
    // 🔥 ИСПРАВЛЕНИЕ: проверяем что availability существует
    if (!availability) {
      return { isAvailable: false, error: 'Ошибка проверки доступности' };
    }
    
    return {
      isAvailable: availability.isAvailable,
      missingComponents: availability.missingComponents || [],
      componentDetails: availability.componentDetails || {}
    };
    
  } catch (error) {
    console.error('❌ Ошибка в упрощенной функции: ' + error.toString());
    return { 
      isAvailable: false, 
      missingComponents: [],
      componentDetails: {},
      error: error.toString()
    };
  }
}

/**
 *  🔥 ОБНОВЛЕННАЯ ФУНКЦИЯ: УПРОЩЕННАЯ ПРОВЕРКА МАТЕРИАЛОВ С УВЕДОМЛЕНИЕМ
 */
function checkMaterialsAvailabilityQuickSimple(product, totalItems) {
  try {
    var availability = checkMaterialsAvailabilityQuick(product, totalItems);
    
    // 🔥 ИСПРАВЛЕНИЕ: проверяем что availability существует
    if (!availability) {
      return { isAvailable: false, error: 'Ошибка проверки доступности' };
    }
    
    //  🔥 ДОБАВЛЕНО: ПОКАЗ УВЕДОМЛЕНИЯ ПОЛЬЗОВАТЕЛЮ
    if (!availability.isAvailable && availability.missingComponents && availability.missingComponents.length > 0) {
      var quantity = Math.ceil(totalItems / getItemsPerBox(product));
      showMaterialsShortageNotification(product, quantity, availability.missingComponents, availability.componentDetails);
    }
    
    return {
      isAvailable: availability.isAvailable,
      missingComponents: availability.missingComponents || [],
      componentDetails: availability.componentDetails || {}
    };
    
  } catch (error) {
    console.error('❌ Ошибка в упрощенной функции: ' + error.toString());
    return { 
      isAvailable: false, 
      missingComponents: [],
      componentDetails: {},
      error: error.toString()
    };
  }
}

/**
 * ✅ ОПТИМИЗИРОВАННАЯ ФУНКЦИЯ: БЫСТРАЯ ПРОВЕРКА ДОСТУПНОСТИ МАТЕРИАЛОВ
 */
function checkMaterialsAvailabilityQuick(product, totalItems) {
  try {
    console.log('⚡ БЫСТРАЯ проверка материалов для: ' + product + ' × ' + totalItems + ' шт.');
    
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var productionSheet = spreadsheet.getSheetByName(PRODUCTION_WAREHOUSE_SHEET);
    
    if (!productionSheet) {
      return { isAvailable: false, error: 'Лист склада не найден' };
    }
    
    // 🔥 ЗАГРУЖАЕМ ВСЕ ДАННЫЕ СКЛАДА ОДНИМ ЗАПРОСОМ
    var lastRow = productionSheet.getLastRow();
    if (lastRow <= 1) {
      return { isAvailable: false, error: 'Склад пуст' };
    }
    
    var warehouseData = productionSheet.getRange('A2:F' + lastRow).getValues();
    var componentStocks = {};
    
    // 🔥 ПРЕДВАРИТЕЛЬНО СЧИТАЕМ ОСТАТКИ ВСЕХ КОМПОНЕНТОВ
    for (var i = 0; i < warehouseData.length; i++) {
      var component = warehouseData[i][3]; // D - Комплектующие
      var stock = warehouseData[i][4] || 0;  // E - Остаток
      
      if (component) {
        if (!componentStocks[component]) {
          componentStocks[component] = 0;
        }
        componentStocks[component] += Number(stock);
      }
    }
    
    // Получаем нормы расхода
    var consumptionRates = getConsumptionRatesForProduct(product);
    
    var isAvailable = true;
    var missingComponents = [];
    var componentDetails = {};
    
    // Проверяем ВСЕ компоненты в памяти
    for (var component in consumptionRates) {
      if (consumptionRates.hasOwnProperty(component)) {
        var rate = consumptionRates[component];
        var requiredQuantity = totalItems * rate;
        var availableQuantity = componentStocks[component] || 0;
        
        componentDetails[component] = {
          required: requiredQuantity,
          available: availableQuantity,
          sufficient: availableQuantity >= requiredQuantity
        };
        
        if (availableQuantity < requiredQuantity) {
          isAvailable = false;
          missingComponents.push({
            component: component,
            required: requiredQuantity,
            available: availableQuantity,
            deficit: requiredQuantity - availableQuantity
          });
        }
      }
    }
    
    return {
      isAvailable: isAvailable,
      missingComponents: missingComponents,
      componentDetails: componentDetails
    };
    
  } catch (error) {
    console.error('❌ Ошибка быстрой проверки материалов: ' + error.toString());
    return { isAvailable: false, error: error.toString() };
  }
}

/**
 *  🔥 УЛУЧШЕННАЯ ФУНКЦИЯ: ПОЛУЧЕНИЕ НОРМ РАСХОДА С ДЕТАЛЬНЫМ ЛОГИРОВАНИЕМ
 */
function getConsumptionRatesForProduct(productName) {
  try {
    console.log('🔍 ===== ПОИСК НОРМ РАСХОДА ДЛЯ ТОВАРА =====');
    console.log('📦 Искомый товар: "' + productName + '"');
    
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var consumptionSheet = spreadsheet.getSheetByName(CONSUMPTION_RATES_SHEET);
    
    if (!consumptionSheet) {
      console.error('❌ Лист норм расхода не найден: ' + CONSUMPTION_RATES_SHEET);
      throw new Error('Лист норм расхода не найден: ' + CONSUMPTION_RATES_SHEET);
    }
    
    var lastRow = consumptionSheet.getLastRow();
    if (lastRow <= 1) {
      console.warn('⚠️ Лист норм расхода пуст!');
      return {};
    }
    
    // 🔥 ПРАВИЛЬНЫЕ СТОЛБЦЫ ИЗ ВАШЕЙ СТРУКТУРЫ:
    // A - Продукция, B - Комплектующие, C - Норма расхода
    var data = consumptionSheet.getRange('A2:C' + lastRow).getValues();
    var consumptionRates = {};
    var ratesFound = 0;
    
    console.log('📋 Всего строк в нормах расхода: ' + data.length);
    
    for (var i = 0; i < data.length; i++) {
      var product = data[i][0]; // A - Продукция
      var component = data[i][1]; // B - Комплектующие
      var rate = data[i][2]; // C - Норма расхода
      
      //  🔥 УЛУЧШЕННОЕ СРАВНЕНИЕ: регистронезависимое, с триммингом
      if (product && product.toString().trim().toLowerCase() === productName.toString().trim().toLowerCase()) {
        if (component && rate > 0) {
          //  🔥 ТОЧНОЕ ЗНАЧЕНИЕ БЕЗ ИЗМЕНЕНИЙ
          consumptionRates[component] = Number(rate);
          ratesFound++;
          console.log('✅ Найдена норма: "' + component + '" - ' + rate + ' ед./шт. (строка ' + (i + 2) + ')');
        }
      }
    }
    
    console.log('🔍 ===== РЕЗУЛЬТАТ ПОИСКА НОРМ =====');
    console.log('📦 Товар: "' + productName + '"');
    console.log('📋 Найдено норм расхода: ' + ratesFound);
    
    if (ratesFound === 0) {
      console.warn('⚠️ ВНИМАНИЕ: для товара "' + productName + '" не найдены нормы расхода!');
      console.log('🔍 Доступные товары в справочнике:');
      
      // Показываем доступные товары для отладки
      var availableProducts = {};
      for (var j = 0; j < data.length; j++) {
        var availableProduct = data[j][0];
        if (availableProduct && !availableProducts[availableProduct]) {
          availableProducts[availableProduct] = true;
          console.log('   - "' + availableProduct + '"');
        }
      }
    }
    
    return consumptionRates;
    
  } catch (error) {
    console.error('❌ КРИТИЧЕСКАЯ ОШИБКА получения норм расхода: ' + error.toString());
    return {};
  }
}

/**
 * 🔥 УЛУЧШЕННАЯ ФУНКЦИЯ: ПОЛУЧЕНИЕ ОБЩЕГО ОСТАТКА КОМПОНЕНТА С ДЕТАЛЬНЫМ ЛОГИРОВАНИЕМ
 */
function getTotalStockForComponent(sheet, component) {
  try {
    console.log('🔍 Поиск остатков компонента: "' + component + '" на складе: ' + sheet.getName());
    
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      console.log('ℹ️ Склад пуст, остаток: 0');
      return 0;
    }
    
    //  🔥 ПРАВИЛЬНАЯ СТРУКТУРА ДЛЯ ВАШИХ СКЛАДОВ:
    // A: ID перемещения, B: ID Закупки, C: Дата поступления, D: Комплектующие, E: Остаток, F: Статус
    var data = sheet.getRange('A2:F' + lastRow).getValues();
    var totalStock = 0;
    var foundRows = 0;
    
    console.log('📊 Проверяем ' + data.length + ' строк на складе...');
    
    for (var i = 0; i < data.length; i++) {
      var rowComponent = data[i][3];     // D - Комплектующие
      var currentStock = data[i][4] || 0; // E - Остаток
      
      // Проверяем совпадение по названию комплектующего (регистронезависимо)
      if (rowComponent && rowComponent.toString().trim().toLowerCase() === component.toString().trim().toLowerCase()) {
        totalStock += Number(currentStock);
        foundRows++;
        console.log('   ✅ Строка ' + (i + 2) + ': ' + currentStock + ' ед. (общий: ' + totalStock + ' ед.)');
      }
    }
    
    totalStock = roundToTwo(totalStock);
    
    console.log('📊 ИТОГ по компоненту "' + component + '":');
    console.log('    📋 Найдено строк: ' + foundRows);
    console.log('    📦 Общий остаток: ' + totalStock + ' ед.');
    
    return totalStock;
    
  } catch (error) {
    console.error('❌ ОШИБКА получения остатка компонента "' + component + '": ' + error.toString());
    return 0;
  }
}

/**
 * 🔥 ОБНОВЛЕННАЯ ФУНКЦИЯ: ВЫПОЛНЕНИЕ СПИСАНИЯ С ПРАВИЛЬНЫМ FIFO И ОКРУГЛЕНИЕМ ДО 2 ЗНАКОВ
 */
function executeMaterialsConsumptionOptimized(sheet, row, product, quantity, totalItems) {
  try {
    console.log('⚡ ОПТИМИЗИРОВАННОЕ списание для: ' + product + ' × ' + quantity + ' кор. (FIFO + округление)');

    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var productionSheet = spreadsheet.getSheetByName(PRODUCTION_WAREHOUSE_SHEET);
    var writeOffSheet = spreadsheet.getSheetByName('Списание_материалов');
    
    if (!productionSheet) {
      throw new Error('Лист склада производства не найден');
    }

    // 🔥 ИСПОЛЬЗУЕМ ЛОГИКУ, СООТВЕТСТВУЮЩУЮ ПОЛНОМУ ПЕРЕСЧЕТУ
    var consumptionRates = getConsumptionRatesForProduct(product);
    
    if (Object.keys(consumptionRates).length === 0) {
      throw new Error('Не найдены нормы расхода для товара: ' + product);
    }

    // 🔥 СПИСЫВАЕМ С УЧЕТОМ ВСЕХ ПАРТИЙ ПО ФИФО (СТАРЫЕ ПЕРВЫМИ) С ОКРУГЛЕНИЕМ
    var writeOffResults = [];
    
    for (var component in consumptionRates) {
      if (consumptionRates.hasOwnProperty(component)) {
        var requiredQuantity = roundToTwo(totalItems * consumptionRates[component]);
        var consumedQuantity = 0;
        
        // Получаем все партии компонента, отсортированные по дате (FIFO)
        var batches = getComponentBatchesSortedByDate(productionSheet, component);
        
        console.log('🔍 Списание ' + component + ': требуется ' + requiredQuantity + ' ед.');
        console.log('📦 Доступные партии в порядке FIFO: ' + batches.length + ' шт.');
        
        for (var i = 0; i < batches.length && consumedQuantity < requiredQuantity; i++) {
          var batch = batches[i];
          var availableInBatch = roundToTwo(batch.stock);
          var neededFromBatch = roundToTwo(Math.min(requiredQuantity - consumedQuantity, availableInBatch));
          
          if (neededFromBatch > 0) {
            //  🔥 ОБНОВЛЯЕМ ОСТАТОК В ПАРТИИ С ОКРУГЛЕНИЕМ ДО 2 ЗНАКОВ
            var newStock = roundToTwo(availableInBatch - neededFromBatch);
            productionSheet.getRange(batch.row, 5).setValue(newStock); // E - Остаток
            
            consumedQuantity = roundToTwo(consumedQuantity + neededFromBatch);
            writeOffResults.push({
              component: component,
              transferId: batch.transferId,
              purchaseId: batch.purchaseId,
              consumed: neededFromBatch,
              fromBatch: batch.row,
              batchDate: batch.date
            });
            
            console.log('➖ Списано ' + component + ': ' + neededFromBatch + ' из партии ' + 
                       batch.transferId + ' (дата: ' + 
                       (batch.date ? Utilities.formatDate(batch.date, Session.getScriptTimeZone(), "dd.MM.yyyy") : "нет даты") + 
                       ', новый остаток: ' + newStock + ')');
          }
        }
        
        if (roundToTwo(consumedQuantity) < roundToTwo(requiredQuantity)) {
          throw new Error('Недостаточно ' + component + ': требуется ' + requiredQuantity + ', доступно ' + consumedQuantity);
        }
      }
    }

    // 🔥 ОБНОВЛЯЕМ СТАТУСЫ ВСЕХ ЗАТРОНУТЫХ СТРОК
    updateWarehouseStatusDirect(productionSheet);

    //  🔥 ЗАПИСЫВАЕМ В ЖУРНАЛ СПИСАНИЙ
    if (writeOffSheet && writeOffResults.length > 0) {
      recordWriteOffToJournal(sheet, row, product, quantity, writeOffResults);
    }

    return {
      success: true,
      componentsProcessed: Object.keys(consumptionRates).length,
      totalConsumed: roundToTwo(writeOffResults.reduce((sum, item) => sum + item.consumed, 0))
    };

  } catch (error) {
    console.error('❌ Ошибка оптимизированного списания: ' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}



/**
 * 🔥 УЛУЧШЕННАЯ ФУНКЦИЯ: ПРОВЕРКА ДОСТАТОЧНОСТИ МАТЕРИАЛОВ С ДЕТАЛЬНЫМ ЛОГИРОВАНИЕМ
 */
function checkMaterialAvailability(product, plannedQuantity, componentsPerBox) {
  try {
    console.log('🔵 ===== НАЧАЛО ПРОВЕРКИ МАТЕРИАЛОВ =====');
    console.log('📦 Товар: ' + product);
    console.log('📊 Планируемое количество: ' + plannedQuantity + ' коробок');
    console.log('📦 Товаров в коробке: ' + componentsPerBox + ' шт.');
    
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var productionSheet = spreadsheet.getSheetByName(PRODUCTION_WAREHOUSE_SHEET);
    
    if (!productionSheet) {
      console.error('❌ Лист склада производства не найден: ' + PRODUCTION_WAREHOUSE_SHEET);
      throw new Error('Лист склада производства не найден');
    }
    
    var totalComponents = plannedQuantity * componentsPerBox;
    console.log('🔢 Общее количество товаров: ' + plannedQuantity + ' кор. × ' + componentsPerBox + ' шт./кор. = ' + totalComponents + ' шт.');
    
    // Получаем нормы расхода
    var consumptionRates = getConsumptionRatesForProduct(product);
    console.log('📋 Найдено норм расхода для "' + product + '": ' + Object.keys(consumptionRates).length);
    
    if (Object.keys(consumptionRates).length === 0) {
      console.warn('⚠️ Внимание: для товара "' + product + '" не найдены нормы расхода!');
    }
    
    var availability = {
      allAvailable: true,
      missingComponents: [],
      componentDetails: {},
      totalRequired: 0,
      calculationDetails: {
        product: product,
        boxes: plannedQuantity,
        itemsPerBox: componentsPerBox,
        totalItems: totalComponents
      }
    };
    
    // Проверяем каждый компонент
    for (var component in consumptionRates) {
      if (consumptionRates.hasOwnProperty(component)) {
        var rate = consumptionRates[component];
        
        // 🔥 ТОЧНЫЙ РАСЧЕТ БЕЗ ОКРУГЛЕНИЯ ДО ЦЕЛЫХ
        var requiredQuantity = totalComponents * rate;
        var availableQuantity = getTotalStockForComponent(productionSheet, component);
        
        // 🔥 ДЕТАЛЬНОЕ ЛОГИРОВАНИЕ РАСЧЕТА
        console.log('🔍 Компонент: ' + component);
        console.log('   📐 Норма расхода: ' + rate + ' ед./шт.');
        console.log('   📊 Расчет: ' + totalComponents + ' шт. × ' + rate + ' ед./шт. = ' + requiredQuantity + ' ед.');
        console.log('    📦 Доступно на складе: ' + availableQuantity + ' ед.');
        console.log('   ✅ Достаточно: ' + (availableQuantity >= requiredQuantity ? 'ДА' : 'НЕТ'));
        
        availability.componentDetails[component] = {
          required: requiredQuantity,
          available: availableQuantity,
          sufficient: availableQuantity >= requiredQuantity,
          rate: rate,
          deficit: Math.max(0, requiredQuantity - availableQuantity)
        };
        
        availability.totalRequired += requiredQuantity;
        
        if (availableQuantity < requiredQuantity) {
          availability.allAvailable = false;
          availability.missingComponents.push({
            component: component,
            required: requiredQuantity,
            available: availableQuantity,
            deficit: requiredQuantity - availableQuantity,
            rate: rate
          });
          
          console.log('❌ Недостаточно компонента: ' + component);
          console.log('    💡 Требуется: ' + requiredQuantity.toFixed(2) + ' ед.');
          console.log('    📊 Доступно: ' + availableQuantity.toFixed(2) + ' ед.');
          console.log('    ⚠️ Дефицит: ' + (requiredQuantity - availableQuantity).toFixed(2) + ' ед.');
        }
      }
    }
    
    console.log('🔵 ===== РЕЗУЛЬТАТ ПРОВЕРКИ =====');
    console.log('✅ Всего компонентов проверено: ' + Object.keys(consumptionRates).length);
    console.log('📊 Компонентов достаточно: ' + (availability.allAvailable ? 'ДА' : 'НЕТ'));
    console.log('❌ Проблемных компонентов: ' + availability.missingComponents.length);
    
    if (!availability.allAvailable) {
      console.log('📋 Список недостающих компонентов:');
      availability.missingComponents.forEach(function(missing, index) {
        console.log('   ' + (index + 1) + '. ' + missing.component + ': не хватает ' + missing.deficit.toFixed(2) + ' ед.');
      });
    }
    
    return availability;
    
  } catch (error) {
    console.error('❌ КРИТИЧЕСКАЯ ОШИБКА ПРОВЕРКИ МАТЕРИАЛОВ: ' + error.toString());
    return {
      allAvailable: false,
      missingComponents: [],
      componentDetails: {},
      error: error.toString(),
      calculationDetails: {
        product: product,
        boxes: plannedQuantity,
        itemsPerBox: componentsPerBox,
        error: true
      }
    };
  }
}

/**
 * 🔥 НОВАЯ ФУНКЦИЯ: ДИАГНОСТИКА ПРОБЛЕМЫ С МАТЕРИАЛАМИ
 */
function diagnoseMaterialProblem(productName, plannedQuantity) {
  try {
    console.log('🔧 ===== ДИАГНОСТИКА ПРОБЛЕМЫ С МАТЕРИАЛАМИ =====');
    
    var itemsPerBox = getItemsPerBox(productName);
    var totalItems = plannedQuantity * itemsPerBox;
    
    console.log('📊 Параметры расчета:');
    console.log('    📦 Товар: ' + productName);
    console.log('    🔢 Планируемое количество: ' + plannedQuantity + ' коробок');
    console.log('    📐 Товаров в коробке: ' + itemsPerBox + ' шт.');
    console.log('   📈 Общее количество товаров: ' + totalItems + ' шт.');
    
    // Проверяем наличие товара в справочнике
    var productExists = checkProductInCatalog(productName);
    console.log('📋 Товар в справочнике: ' + (productExists ? 'НАЙДЕН' : 'НЕ НАЙДЕН'));
    
    // Проверяем нормы расхода
    var consumptionRates = getConsumptionRatesForProduct(productName);
    console.log('📊 Норм расхода найдено: ' + Object.keys(consumptionRates).length);
    
    // Проверяем остатки на складе
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var productionSheet = spreadsheet.getSheetByName(PRODUCTION_WAREHOUSE_SHEET);
    
    if (!productionSheet) {
      console.error('❌ Склад производства не найден!');
      return;
    }
    
    console.log('📦 Остатки на складе производства:');
    for (var component in consumptionRates) {
      var available = getTotalStockForComponent(productionSheet, component);
      var required = totalItems * consumptionRates[component];
      console.log('    🔧 ' + component + ': требуется ' + required.toFixed(2) + ' ед., доступно ' + available.toFixed(2) + ' ед.');
    }
    
    console.log('🔧 ===== ЗАВЕРШЕНИЕ ДИАГНОСТИКИ =====');
    
  } catch (error) {
    console.error('❌ Ошибка диагностики: ' + error.toString());
  }
}

/**
 * 🔥 НОВАЯ ФУНКЦИЯ: ПРОВЕРКА НАЛИЧИЯ ТОВАРА В КАТАЛОГЕ
 */
function checkProductInCatalog(productName) {
  try {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var catalogSheet = spreadsheet.getSheetByName('Справочник_товаров');
    
    if (!catalogSheet) return false;
    
    var lastRow = catalogSheet.getLastRow();
    if (lastRow <= 1) return false;
    
    var data = catalogSheet.getRange('A2:B' + lastRow).getValues();
    
    for (var i = 0; i < data.length; i++) {
      var code = data[i][0]; // A - Код
      var name = data[i][1]; // B - Название
      
      if ((code && code.toString().trim().toLowerCase() === productName.toString().trim().toLowerCase()) ||
          (name && name.toString().trim().toLowerCase() === productName.toString().trim().toLowerCase())) {
        return true;
      }
    }
    
    return false;
    
  } catch (error) {
    console.error('Ошибка проверки товара в каталоге: ' + error.toString());
    return false;
  }
}

/**
 * ✅ ФУНКЦИЯ: УВЕДОМЛЕНИЕ О НЕДОСТАТКЕ МАТЕРИАЛОВ С ДЕТАЛЬНОЙ ИНФОРМАЦИЕЙ
 */
function showMaterialsShortageNotification(product, quantity, missingComponents, componentDetails) {
  try {
    var ui = SpreadsheetApp.getUi();
    var message = '⚠️ ПРОВЕРКА МАТЕРИАЛОВ ДЛЯ ПРОИЗВОДСТВА\n\n';
    message += '📦 Товар: ' + product + '\n';
    message += '📊 Планируемое количество: ' + quantity + ' коробок\n\n';
    
    if (missingComponents.length === 0) {
      message += '✅ Все материалы в наличии!\n\n';
      message += '📋 Детали по компонентам:\n';
      
      for (var component in componentDetails) {
        if (componentDetails.hasOwnProperty(component)) {
          var detail = componentDetails[component];
          message += '• ' + component + ': ' + detail.available.toFixed(2) + ' / ' + 
                    detail.required.toFixed(2) + ' ед. (' + 
                    (detail.sufficient ? '✅ Достаточно' : '❌ Недостаточно') + ')\n';
        }
      }
    } else {
      message += '❌ НЕДОСТАТОЧНО МАТЕРИАЛОВ!\n\n';
      message += '📋 Проблемные компоненты:\n';
      
      missingComponents.forEach(function(missing, index) {
        message += (index + 1) + '. ' + missing.component + ':\n';
        message += '    💡 Требуется: ' + missing.required.toFixed(2) + ' ед.\n';
        message += '    📊 Доступно: ' + missing.available.toFixed(2) + ' ед.\n';
        message += '    ⚠️ Дефицит: ' + missing.deficit.toFixed(2) + ' ед.\n\n';
      });
      
      message += '💡 РЕКОМЕНДАЦИИ:\n';
      message += '• Проверьте остатки на складе производства\n';
      message += '• Запланируйте закупку недостающих материалов\n';
      message += '• Скорректируйте количество коробок\n';
    }
    
    ui.alert('📦 Проверка материалов', message, ui.ButtonSet.OK);
    
  } catch (error) {
    console.error('❌ Ошибка показа уведомления: ' + error.toString());
  }
}

/**
 * 🚀 ФУНКЦИЯ: ВЫПОЛНЕНИЕ БАТЧЕВЫХ ОБНОВЛЕНИЙ
 */
function executeBatchUpdates(sheet, updates) {
  try {
    console.log('🔄 Начало выполнения батчевых обновлений: ' + updates.length + ' компонентов');
    
    var totalUpdates = 0;
    
    // 🔥 ВЫПОЛНЯЕМ ВСЕ ОБНОВЛЕНИЯ ПО КОМПОНЕНТАМ
    for (var i = 0; i < updates.length; i++) {
      var componentUpdate = updates[i];
      var component = componentUpdate.component;
      var plan = componentUpdate.plan;
      
      console.log('➖ Выполнение списания ' + component + ' из ' + plan.length + ' партий');
      
      // 🔥 ОБНОВЛЯЕМ ВСЕ ПАРТИИ ДЛЯ КОМПОНЕНТА
      for (var j = 0; j < plan.length; j++) {
        var batchUpdate = plan[j];
        sheet.getRange(batchUpdate.rowIndex, NEW_WAREHOUSE_STRUCTURE.CURRENT_STOCK + 1)
             .setValue(batchUpdate.newStock);
        
        console.log('   ➖ Партия ' + batchUpdate.rowIndex + ': ' + batchUpdate.oldStock + ' → ' + batchUpdate.newStock + ' ед. (списано: ' + batchUpdate.quantityToWriteOff + ' ед.)');
        totalUpdates++;
      }
    }
    
    console.log('✅ Батчевые обновления завершены: ' + totalUpdates + ' обновлений строк');
    
  } catch (error) {
    console.error('❌ Ошибка выполнения батчевых обновлений: ' + error.toString());
    throw error;
  }
}

/**
 * ✅ ИСПРАВЛЕННАЯ ФУНКЦИЯ: ОБНОВЛЕНИЕ СТАТУСОВ С ПРАВИЛЬНЫМ СООТВЕТСТВИЕМ СТРОК ПОСЛЕ СОРТИРОВКИ
 */
function updateWarehouseStatusDirect(sheet, addTooltips) {
  try {
    // 🔥 addTooltips по умолчанию = false (не добавляем комментарии)
    if (typeof addTooltips === 'undefined') {
      addTooltips = true;
    }
    
    console.log('🔄 Обновление статусов: ' + sheet.getName() + ' (комментарии: ' + (addTooltips ? 'ДА' : 'НЕТ') + ')');
    
    if (!sheet) {
      throw new Error('Лист не передан в функцию');
    }
    
    var lastRow = sheet.getLastRow();
    
    if (lastRow <= 1) {
      console.log('✅ ' + sheet.getName() + ' - склад пуст, пропускаем обновление статусов');
      return;
    }
    
    // 🔥 ПРОВЕРЯЕМ СТРУКТУРУ ЛИСТА
    var lastColumn = sheet.getLastColumn();
    if (lastColumn < 6) {
      console.error('❌ Неправильная структура листа ' + sheet.getName() + ': только ' + lastColumn + ' столбцов (требуется минимум 6)');
      throw new Error('Неправильная структура листа: требуется минимум 6 столбцов');
    }
    
    // 🔥 ВАЖНОЕ ИСПРАВЛЕНИЕ: Получаем ВСЕ данные ОДНИМ ЗАПРОСОМ для сохранения соответствия строк
    var dataRange = sheet.getRange('A2:F' + lastRow);
    var data = dataRange.getValues();
    var displayValues = dataRange.getDisplayValues();
    
    var componentTotals = calculateComponentTotals(sheet);
    
    for (var i = 0; i < data.length; i++) {
      var row = i + 2;
      var component = data[i][NEW_WAREHOUSE_STRUCTURE.COMPONENT];
      var purchaseId = data[i][NEW_WAREHOUSE_STRUCTURE.PURCHASE_ID];
      var receiptDate = data[i][NEW_WAREHOUSE_STRUCTURE.DATE];
      
      var currentStock = parseFloat(data[i][NEW_WAREHOUSE_STRUCTURE.CURRENT_STOCK]) || 0;
      
      // 🔥 ИСПРАВЛЕНИЕ: Получаем минимальный запас из справочника, а не из столбца склада
      var warehouseType = sheet.getName() === 'Склад_Главный' ? 'Главный' : 'Производство';
      var minStock = getMinStockFromReference(component, warehouseType);
      
      currentStock = roundToTwo(currentStock);
      minStock = roundToTwo(minStock);
      
      var currentStockFormatted = displayValues[i][NEW_WAREHOUSE_STRUCTURE.CURRENT_STOCK];
      var minStockFormatted = minStock.toFixed(2); // Форматируем minStock для отображения
      
      if (!component || component.toString().trim() === '') {
        console.log('⚠️ Пропущена пустая строка ' + row + ' на ' + sheet.getName());
        continue;
      }
      
      var status = '';
      var backgroundColor = '';
      var tooltip = '';
      
      var totalStock = roundToTwo(componentTotals[component] || 0);
      
      // Определяем статус
      if (currentStock === 0) {
        status = '❌ Отсутствует';
        backgroundColor = '#FFCCCC';
        tooltip = addTooltips ? createTooltip(component, purchaseId, receiptDate, currentStock, minStock, totalStock, 'absent') : '';
      } else if (minStock > 0 && totalStock < minStock) {
        var deficit = roundToTwo(minStock - totalStock);
        status = '⚠️ Мало: ' + currentStockFormatted + '/' + minStockFormatted + ' ед. | -' + deficit.toFixed(2);
        backgroundColor = '#FFF2CC';
        tooltip = addTooltips ? createTooltip(component, purchaseId, receiptDate, currentStock, minStock, totalStock, 'low') : '';
      } else if (currentStock === 1) {
        status = '✅ В наличии: ' + currentStockFormatted + ' ед. | Последняя';
        backgroundColor = '#FFF2CC';
        tooltip = addTooltips ? createTooltip(component, purchaseId, receiptDate, currentStock, minStock, totalStock, 'last') : '';
      } else if (currentStock <= 5) {
        status = '✅ В наличии: ' + currentStockFormatted + ' ед. | Мало осталось';
        backgroundColor = '#CCFFCC';
        tooltip = addTooltips ? createTooltip(component, purchaseId, receiptDate, currentStock, minStock, totalStock, 'few') : '';
      } else {
        status = '✅ В наличии: ' + currentStockFormatted + ' ед.';
        backgroundColor = '#CCFFCC';
        tooltip = addTooltips ? createTooltip(component, purchaseId, receiptDate, currentStock, minStock, totalStock, 'normal') : '';
      }
      
      var statusCell = sheet.getRange(row, NEW_WAREHOUSE_STRUCTURE.STATUS + 1);
      statusCell.setValue(status);
      statusCell.setBackground(backgroundColor);
      
      // 🔥 ДОБАВЛЯЕМ КОММЕНТАРИЙ ТОЛЬКО ЕСЛИ addTooltips = true
      if (addTooltips && tooltip) {
        setCellNote(statusCell, tooltip);
      } else {
        // 🔥 ИНАЧЕ ОЧИЩАЕМ КОММЕНТАРИЙ
        statusCell.clearNote();
      }
    }
    
    console.log('✅ Статусы обновлены: ' + sheet.getName());
    
  } catch (error) {
    console.error('❌ Ошибка в updateWarehouseStatusDirect: ' + error.toString());
    throw error;
  }
}

/**
 * ✅ ФУНКЦИЯ: МАССОВОЕ ОБНОВЛЕНИЕ СТАТУСОВ ВСЕХ СКЛАДОВ
 */
function updateAllWarehouseStatuses() {
  try {
    console.log('🔄 Запуск массового обновления статусов всех складов...');
    
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var ui = SpreadsheetApp.getUi();
    
    // Получаем все склады
    var mainSheet = spreadsheet.getSheetByName(MAIN_WAREHOUSE_SHEET);
    var productionSheet = spreadsheet.getSheetByName(PRODUCTION_WAREHOUSE_SHEET);
    
    var updatedCount = 0;
    
    console.log('🔍 Проверка доступности листов:');
    console.log('• Главный склад: ' + (mainSheet ? '✅ Найден' : '❌ Не найден'));
    console.log('• Склад производства: ' + (productionSheet ? '✅ Найден' : '❌ Не найден'));
    
    if (mainSheet) {
      console.log('📊 Обновление статусов главного склада...');
      console.log('📈 Строк на главном складе: ' + mainSheet.getLastRow());
      updateWarehouseStatusDirect(mainSheet, true);
      updatedCount++;
      console.log('✅ Главный склад обновлен');
    }
    
    if (productionSheet) {
      console.log('📊 Обновление статусов склада производства...');
      console.log('📈 Строк на производстве: ' + productionSheet.getLastRow());
      updateWarehouseStatusDirect(productionSheet, true);
      updatedCount++;
      console.log('✅ Склад производства обновлен');
    }
    
    console.log('✅ Массовое обновление завершено. Обновлено складов: ' + updatedCount);
    
    ui.alert(
      '✅ Статусы обновлены', 
      'Статусы складов успешно обновлены!\n\n' +
      '• Главный склад: ' + (mainSheet ? '✅' : '❌ Не найден') + '\n' +
      '• Склад производства: ' + (productionSheet ? '✅' : '❌ Не найден') + '\n\n' +
      'Всего обновлено: ' + updatedCount + ' складов',
      ui.ButtonSet.OK
    );
    
    return updatedCount;
    
  } catch (error) {
    console.error('❌ Ошибка массового обновления статусов: ' + error.toString());
    SpreadsheetApp.getUi().alert(
      '❌ Ошибка', 
      'Не удалось обновить статусы складов:\n\n' + error.toString(),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return 0;
  }
}

/**
 * ✅ ФУНКЦИЯ: ПРОВЕРКА СТАТУСОВ СКЛАДОВ С ВЫВОДОМ ОТЧЕТА
 */
function checkWarehouseStatusesReport() {
  try {
    console.log('🔍 Проверка статусов складов с отчетом...');
    
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var mainSheet = spreadsheet.getSheetByName(MAIN_WAREHOUSE_SHEET);
    var productionSheet = spreadsheet.getSheetByName(PRODUCTION_WAREHOUSE_SHEET);
    
    var report = '📊 ОТЧЕТ ПО СТАТУСАМ СКЛАДОВ\n\n';
    var problemComponents = [];
    
    // Проверяем главный склад
    if (mainSheet) {
      var mainStatus = analyzeWarehouseStatus(mainSheet, 'Главный склад');
      report += mainStatus.report;
      problemComponents = problemComponents.concat(mainStatus.problems);
    }
    
    // Проверяем склад производства
    if (productionSheet) {
      var productionStatus = analyzeWarehouseStatus(productionSheet, 'Склад производства');
      report += productionStatus.report;
      problemComponents = problemComponents.concat(productionStatus.problems);
    }
    
    // Добавляем сводку по проблемам
    if (problemComponents.length > 0) {
      report += '\n⚠️ ПРОБЛЕМНЫЕ КОМПОНЕНТЫ:\n';
      problemComponents.forEach(function(problem, index) {
        report += (index + 1) + '. ' + problem.component + ' (' + problem.warehouse + '): ' + 
                 problem.message + '\n';
      });
      
      report += '\n💡 РЕКОМЕНДАЦИИ:\n';
      report += '• Проверьте остатки проблемных компонентов\n';
      report += '• Запланируйте закупку при необходимости\n';
      report += '• Обновите минимальные запасы в справочнике\n';
    } else {
      report += '\n✅ ВСЕ КОМПОНЕНТЫ В НОРМЕ!\n';
      report += 'Все материалы имеют достаточный запас на складах.';
    }
    
    // Показываем отчет пользователю
    SpreadsheetApp.getUi().alert('📊 Отчет по статусам складов', report, SpreadsheetApp.getUi().ButtonSet.OK);
    
    console.log('✅ Отчет по статусам складов сформирован');
    return report;
    
  } catch (error) {
    console.error('❌ Ошибка формирования отчета: ' + error.toString());
    SpreadsheetApp.getUi().alert(
      '❌ Ошибка', 
      'Не удалось сформировать отчет:\n\n' + error.toString(),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return 'Ошибка формирования отчета';
  }
}

/**
 * ✅ ФУНКЦИЯ: АНАЛИЗ СТАТУСОВ КОНКРЕТНОГО СКЛАДА
 */
function analyzeWarehouseStatus(sheet, warehouseName) {
  try {
    console.log('🔍 Анализ статусов: ' + warehouseName);
    
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return {
        report: '📦 ' + warehouseName + ': ПУСТО\n\n',
        problems: []
      };
    }
    
    var data = sheet.getRange('A2:F' + lastRow).getValues();
    var componentTotals = calculateComponentTotals(sheet);
    
    var totalComponents = 0;
    var okComponents = 0;
    var warningComponents = 0;
    var errorComponents = 0;
    var problems = [];
    
    for (var i = 0; i < data.length; i++) {
      var component = data[i][NEW_WAREHOUSE_STRUCTURE.COMPONENT];
      var currentStock = parseFloat(data[i][NEW_WAREHOUSE_STRUCTURE.CURRENT_STOCK]) || 0;
      var minStock = parseFloat(data[i][NEW_WAREHOUSE_STRUCTURE.MIN_STOCK]) || 0;
      
      if (!component || component.toString().trim() === '') continue;
      
      totalComponents++;
      currentStock = roundToTwo(currentStock);
      minStock = roundToTwo(minStock);
      var totalStock = roundToTwo(componentTotals[component] || 0);
      
      // Анализируем статус
      if (currentStock === 0) {
        errorComponents++;
        problems.push({
          component: component,
          warehouse: warehouseName,
          message: 'Отсутствует на складе',
          current: currentStock,
          required: minStock,
          type: 'error'
        });
      } else if (minStock > 0 && totalStock < minStock) {
        warningComponents++;
        problems.push({
          component: component,
          warehouse: warehouseName,
          message: 'Ниже минимального запаса (дефицит: ' + roundToTwo(minStock - totalStock) + ' ед.)',
          current: totalStock,
          required: minStock,
          type: 'warning'
        });
      } else if (currentStock <= 1) {
        warningComponents++;
        problems.push({
          component: component,
          warehouse: warehouseName,
          message: 'Осталась последняя единица',
          current: currentStock,
          required: minStock,
          type: 'warning'
        });
      } else {
        okComponents++;
      }
    }
    
    var report = '📦 ' + warehouseName + ':\n';
    report += '   • Всего компонентов: ' + totalComponents + '\n';
    report += '   • ✅ В норме: ' + okComponents + '\n';
    report += '   • ⚠️ С предупреждением: ' + warningComponents + '\n';
    report += '   • ❌ С ошибкой: ' + errorComponents + '\n\n';
    
    return {
      report: report,
      problems: problems
    };
    
  } catch (error) {
    console.error('❌ Ошибка анализа статусов: ' + error.toString());
    return {
      report: '❌ Ошибка анализа ' + warehouseName + ': ' + error.toString() + '\n\n',
      problems: []
    };
  }
}

/**
 * ✅ ИСПРАВЛЕННАЯ ФУНКЦИЯ: ПОЛНЫЙ ПЕРЕСЧЕТ ОСТАТКОВ С ПРАВИЛЬНОЙ ЛОГИКОЙ
 */
function recalculateAllWarehouseBalances() {
  try {
    console.log('🔄 ЗАПУСК ИСПРАВЛЕННОГО ПЕРЕСЧЕТА ОСТАТКОВ...');
    
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    var mainWarehouseSheet = spreadsheet.getSheetByName(MAIN_WAREHOUSE_SHEET);
    var productionWarehouseSheet = spreadsheet.getSheetByName(PRODUCTION_WAREHOUSE_SHEET);
    var purchasesSheet = spreadsheet.getSheetByName(PURCHASES_SHEET);
    var planningSheet = spreadsheet.getSheetByName(PLANNING_SHEET);
    var transfersSheet = spreadsheet.getSheetByName(TRANSFERS_SHEET);
    
    if (!mainWarehouseSheet) throw new Error('Лист "' + MAIN_WAREHOUSE_SHEET + '" не найден!');
    if (!productionWarehouseSheet) throw new Error('Лист "' + PRODUCTION_WAREHOUSE_SHEET + '" не найден!');
    if (!purchasesSheet) throw new Error('Лист "' + PURCHASES_SHEET + '" не найден!');
    if (!planningSheet) throw new Error('Лист "' + PLANNING_SHEET + '" не найден!');
    
    //  🔥 ИСПРАВЛЕНИЕ: Используем правильную структуру для хранения остатков
    var warehouseBalances = {
      'Главный': {}, // Компонент -> общий остаток
      'Производство': {} // Компонент -> общий остаток
    };
    
    // ========== 1. ОБРАБАТЫВАЕМ ЗАКУПКИ ==========
    console.log('\n📥 === ЭТАП 1: ЗАКУПКИ ===');
    var purchasesLastRow = purchasesSheet.getLastRow();
    
    if (purchasesLastRow > 1) {
      var purchasesData = purchasesSheet.getRange('A2:I' + purchasesLastRow).getValues();
      
      for (var i = 0; i < purchasesData.length; i++) {
        var pRow = purchasesData[i];
        var purchaseId = pRow[0];
        var component = pRow[2];
        var quantity = roundToTwo(pRow[3] || 0);
        var targetWarehouseRaw = pRow[7];
        var status = pRow[8];
        
        if (!component || !purchaseId || status !== 'Получено' || quantity <= 0) {
          continue;
        }
        
        //  🔥 ИСПРАВЛЕНИЕ: Правильное определение склада назначения
        var warehouseName = 'Главный'; // По умолчанию
        if (targetWarehouseRaw) {
          var warehouseStr = targetWarehouseRaw.toString().trim().toLowerCase();
          if (warehouseStr === 'производство' || warehouseStr.indexOf('производ') !== -1) {
            warehouseName = 'Производство';
          }
        }
        
        // 🔥 ИСПРАВЛЕНИЕ: Суммируем остатки по компонентам, а не создаем отдельные записи
        if (!warehouseBalances[warehouseName][component]) {
          warehouseBalances[warehouseName][component] = 0;
        }
        warehouseBalances[warehouseName][component] += quantity;
        
        console.log('✅ Закупка: ' + component + ' × ' + quantity + ' на ' + warehouseName);
      }
    }
    
    // ========== 2. ОБРАБАТЫВАЕМ ПЕРЕМЕЩЕНИЯ ==========
    console.log('\n🔄 === ЭТАП 2: ПЕРЕМЕЩЕНИЯ ===');
    
    if (transfersSheet && transfersSheet.getLastRow() > 1) {
      var transfersData = transfersSheet.getRange('A2:I' + transfersSheet.getLastRow()).getValues();
      
      for (var t = 0; t < transfersData.length; t++) {
        var tRow = transfersData[t];
        var tComponent = tRow[5]; // F - Комплектующее
        var tQty = roundToTwo(tRow[6] || 0); // G - Количество
        var tStatus = tRow[7]; // H - Статус
        var fromW = tRow[3]; // D - От склада
        var toW = tRow[4]; // E - На склад
        
        if (!tComponent || tQty <= 0 || tStatus !== 'Выполнено') {
          continue;
        }
        
        // 🔥 ИСПРАВЛЕНИЕ: Пропускаем перемещения от Поставщика (уже учтены в закупках)
        var fromWarehouseStr = fromW ? fromW.toString().trim().toLowerCase() : '';
        if (fromWarehouseStr === 'поставщик' || fromWarehouseStr.indexOf('поставщик') !== -1) {
          continue;
        }
        
        // Определяем склады
        var fromWarehouse = fromWarehouseStr.indexOf('производ') !== -1 ? 'Производство' : 'Главный';
        var toWarehouse = toW && toW.toString().trim().toLowerCase().indexOf('производ') !== -1 ? 'Производство' : 'Главный';
        
        // 🔥 ИСПРАВЛЕНИЕ: Проверяем наличие компонента на складе-источнике
        if (warehouseBalances[fromWarehouse][tComponent] !== undefined && 
            warehouseBalances[fromWarehouse][tComponent] >= tQty) {
          
          // Снимаем с источника
          warehouseBalances[fromWarehouse][tComponent] -= tQty;
          
          // Добавляем к приемнику
          if (!warehouseBalances[toWarehouse][tComponent]) {
            warehouseBalances[toWarehouse][tComponent] = 0;
          }
          warehouseBalances[toWarehouse][tComponent] += tQty;
          
          console.log('✅ Перемещение: ' + tComponent + ' × ' + tQty + ' из ' + fromWarehouse + ' в ' + toWarehouse);
        } else {
          console.log('⚠️ Пропущено перемещение: недостаточно ' + tComponent + ' на ' + fromWarehouse);
        }
      }
    }
    
    // ========== 3. ОБРАБАТЫВАЕМ СПИСАНИЯ ==========
    console.log('\n📋 === ЭТАП 3: СПИСАНИЯ ===');
    
    var planningLastRow = planningSheet.getLastRow();
    
    if (planningLastRow > 1) {
      var planningData = planningSheet.getRange('A2:J' + planningLastRow).getValues();
      
      for (var i = 0; i < planningData.length; i++) {
        var prow = planningData[i];
        var product = prow[3]; // D - Товар
        var quantity = roundToTwo(prow[7] || 0); // H - Кол-во коробок
        var status = prow[8]; // I - Статус
        
        if (!product || quantity <= 0 || status !== '✅ Выполнено') {
          continue;
        }
        
        var itemsPerBox = getItemsPerBox(product);
        var totalItems = roundToTwo(quantity * itemsPerBox);
        var consumptionRates = getConsumptionRatesForProduct(product);
        
        console.log('📦 Списание для: ' + product + ' × ' + quantity + ' кор. (' + totalItems + ' шт.)');
        
        // Списываем с производства
        for (var componentName in consumptionRates) {
          if (consumptionRates.hasOwnProperty(componentName)) {
            var qtyPerUnit = roundToTwo(consumptionRates[componentName] || 0);
            var needed = roundToTwo(qtyPerUnit * totalItems);
            
            if (needed > 0 && warehouseBalances['Производство'][componentName] !== undefined) {
              var available = warehouseBalances['Производство'][componentName];
              if (available >= needed) {
                warehouseBalances['Производство'][componentName] -= needed;
                console.log('➖ Списано: ' + componentName + ' × ' + needed + ' ед.');
              } else {
                console.log('⚠️ Не хватило: ' + componentName + ' - нужно ' + needed + ', есть ' + available);
                warehouseBalances['Производство'][componentName] = 0; // Списываем все что есть
              }
            }
          }
        }
      }
    }
    
    // 🔥 ИСПРАВЛЕНИЕ: Удаляем нулевые остатки
    cleanZeroBalances(warehouseBalances);
    
    // ========== 4. ЗАПИСЫВАЕМ РЕЗУЛЬТАТЫ ==========
    console.log('\n💾 === ЭТАП 4: ЗАПИСЬ РЕЗУЛЬТАТОВ ===');
    
    // Сохраняем существующие ID перемещений и даты для главного склада
    var mainRecords = preserveExistingRecords(mainWarehouseSheet);
    updateWarehouseSheet(mainWarehouseSheet, warehouseBalances['Главный'], mainRecords, 'Главный');
    
    // Сохраняем существующие ID перемещений и даты для производства
    var productionRecords = preserveExistingRecords(productionWarehouseSheet);
    updateWarehouseSheet(productionWarehouseSheet, warehouseBalances['Производство'], productionRecords, 'Производство');
    
    // Обновляем статусы
    updateWarehouseStatusDirect(mainWarehouseSheet);
    updateWarehouseStatusDirect(productionWarehouseSheet);
    
    console.log('✅ ПЕРЕСЧЕТ ЗАВЕРШЕН УСПЕШНО');
    
    // Показываем итоговый отчет
    showRecalculationReport(warehouseBalances);
    
  } catch (error) {
    console.error('❌ ОШИБКА ПЕРЕСЧЕТА: ' + error.toString());
    SpreadsheetApp.getUi().alert('❌ Ошибка пересчета', error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 *  🔥 ИСПРАВЛЕННАЯ ФУНКЦИЯ: ОЧИСТКА НУЛЕВЫХ ОСТАТКОВ (ТЕПЕРЬ НЕ УДАЛЯЕТ, А ТОЛЬКО ОТМЕЧАЕТ)
 */
function cleanZeroBalances(warehouseBalances) {
  var zeroCount = 0;
  for (var warehouse in warehouseBalances) {
    for (var component in warehouseBalances[warehouse]) {
      if (warehouseBalances[warehouse][component] <= 0) {
        zeroCount++;
        // 🔥 ВАЖНОЕ ИЗМЕНЕНИЕ: Не удаляем нулевые остатки, а только помечаем их
        console.log('⚠️ Найден нулевой остаток: ' + component + ' на ' + warehouse + ' - сохраняем для истории');
      }
    }
  }
  console.log('📊 Найдено нулевых остатков: ' + zeroCount + ' (сохранены для истории)');
}

/**
 * 🔥 НОВАЯ ФУНКЦИЯ: СОХРАНЕНИЕ СУЩЕСТВУЮЩИХ ЗАПИСЕЙ (ID и даты)
 */
function preserveExistingRecords(sheet) {
  var records = {};
  
  if (sheet.getLastRow() > 1) {
    var data = sheet.getRange('A2:F' + sheet.getLastRow()).getValues();
    
    for (var i = 0; i < data.length; i++) {
      var component = data[i][3]; // D - Комплектующие
      var transferId = data[i][0]; // A - ID перемещения
      var purchaseId = data[i][1]; // B - ID Закупки
      var date = data[i][2]; // C - Дата
      
      if (component) {
        if (!records[component]) {
          records[component] = [];
        }
        records[component].push({
          transferId: transferId,
          purchaseId: purchaseId,
          date: date
        });
      }
    }
  }
  
  return records;
}


/**
 * ✅ ФУНКЦИЯ: ДИАГНОСТИКА ПРОБЛЕМЫ СО СТАТУСАМИ
 */
function diagnoseStatusProblem() {
  try {
    console.log('🔧 ДИАГНОСТИКА ПРОБЛЕМЫ СО СТАТУСАМИ СКЛАДОВ');
    
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var mainSheet = spreadsheet.getSheetByName(MAIN_WAREHOUSE_SHEET);
    var productionSheet = spreadsheet.getSheetByName(PRODUCTION_WAREHOUSE_SHEET);
    
    var report = '🔧 ДИАГНОСТИКА СТАТУСОВ СКЛАДОВ\n\n';
    
    // Проверяем доступность листов
    if (!mainSheet) {
      report += '❌ Главный склад не найден: ' + MAIN_WAREHOUSE_SHEET + '\n';
    } else {
      report += '✅ Главный склад: найден\n';
      report += '   📊 Строк: ' + mainSheet.getLastRow() + '\n';
      
      // Проверяем структуру главного склада
      var mainData = mainSheet.getDataRange().getValues();
      if (mainData.length > 0 && mainData[0].length >= 6) {
        report += '   ✅ Структура: правильная (минимум 6 столбцов)\n';
      } else {
        report += '   ❌ Структура: неправильная (' + (mainData[0] ? mainData[0].length : 0) + ' столбцов)\n';
      }
    }
    
    if (!productionSheet) {
      report += '❌ Склад производства не найден: ' + PRODUCTION_WAREHOUSE_SHEET + '\n';
    } else {
      report += '✅ Склад производства: найден\n';
      report += '   📊 Строк: ' + productionSheet.getLastRow() + '\n';
      
      // Проверяем структуру склада производства
      var productionData = productionSheet.getDataRange().getValues();
      if (productionData.length > 0 && productionData[0].length >= 6) {
        report += '   ✅ Структура: правильная (минимум 6 столбцов)\n';
      } else {
        report += '   ❌ Структура: неправильная (' + (productionData[0] ? productionData[0].length : 0) + ' столбцов)\n';
      }
    }
    
    // Запускаем тестовое обновление
    report += '\n🧪 ТЕСТОВОЕ ОБНОВЛЕНИЕ СТАТУСОВ:\n';
    
    if (mainSheet) {
      try {
        var beforeStatus = mainSheet.getRange('F2:F' + mainSheet.getLastRow()).getValues();
        updateWarehouseStatusDirect(mainSheet, true);
        var afterStatus = mainSheet.getRange('F2:F' + mainSheet.getLastRow()).getValues();
        
        var changedCount = 0;
        for (var i = 0; i < beforeStatus.length; i++) {
          if (beforeStatus[i][0] !== afterStatus[i][0]) {
            changedCount++;
          }
        }
        
        report += '✅ Главный склад: обновлено ' + changedCount + ' статусов\n';
      } catch (e) {
        report += '❌ Главный склад: ОШИБКА - ' + e.toString() + '\n';
      }
    }
    
    if (productionSheet) {
      try {
        var beforeStatus = productionSheet.getRange('F2:F' + productionSheet.getLastRow()).getValues();
        updateWarehouseStatusDirect(productionSheet, true);
        var afterStatus = productionSheet.getRange('F2:F' + productionSheet.getLastRow()).getValues();
        
        var changedCount = 0;
        for (var i = 0; i < beforeStatus.length; i++) {
          if (beforeStatus[i][0] !== afterStatus[i][0]) {
            changedCount++;
          }
        }
        
        report += '✅ Склад производства: обновлено ' + changedCount + ' статусов\n';
      } catch (e) {
        report += '❌ Склад производства: ОШИБКА - ' + e.toString() + '\n';
      }
    }
    
    // Показываем отчет
    SpreadsheetApp.getUi().alert('🔧 Диагностика статусов', report, SpreadsheetApp.getUi().ButtonSet.OK);
    console.log(report);
    
    return report;
    
  } catch (error) {
    console.error('❌ Ошибка диагностики: ' + error.toString());
    SpreadsheetApp.getUi().alert(
      '❌ Ошибка диагностики', 
      'Не удалось выполнить диагностику:\n\n' + error.toString(),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return 'Ошибка диагностики';
  }
}

/**
 * ✅ ВСПОМОГАТЕЛЬНАЯ ФУНКЦИЯ: ЛОГИРОВАНИЕ ОБЩИХ ОСТАТКОВ
 */
function logTotalBalances(balances) {
  var mainTotal = {};
  var productionTotal = {};
  
  // Считаем общие остатки по компонентам на главном складе
  for (var key in balances.main) {
    var batch = balances.main[key];
    if (!mainTotal[batch.component]) {
      mainTotal[batch.component] = 0;
    }
    mainTotal[batch.component] += batch.quantity;
  }
  
  // Считаем общие остатки по компонентам на производстве
  for (var key in balances.production) {
    var batch = balances.production[key];
    if (!productionTotal[batch.component]) {
      productionTotal[batch.component] = 0;
    }
    productionTotal[batch.component] += batch.quantity;
  }
  
  console.log('🏢 ГЛАВНЫЙ СКЛАД:');
  for (var component in mainTotal) {
    console.log('  • ' + component + ': ' + mainTotal[component].toFixed(2) + ' ед.');
  }
  
  console.log('🏭 СКЛАД ПРОИЗВОДСТВА:');
  for (var component in productionTotal) {
    console.log('  • ' + component + ': ' + productionTotal[component].toFixed(2) + ' ед.');
  }
}

/**
 * ✅ ФУНКЦИЯ: ОБНОВЛЕНИЕ ЛИСТА СКЛАДА ИЗ РАСЧЕТНЫХ ДАННЫХ
 */
function updateWarehouseSheetFromBalances(sheet, balances, warehouseType) {
  try {
    console.log('💾 Обновление листа: ' + sheet.getName());
    
    // Очищаем существующие данные (кроме заголовка)
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange('A2:F' + lastRow).clearContent();
    }
    
    var newRows = [];
    var rowIndex = 2;
    
    for (var key in balances) {
      var batch = balances[key];
      if (batch.quantity > 0) { // Добавляем только партии с положительным остатком
        var newRow = [
          generateTransferId(), // A - ID перемещения (новый для каждой партии)
          batch.purchaseId,     // B - ID Закупки
          batch.date,           // C - Дата поступления
          batch.component,      // D - Комплектующие
          batch.quantity,       // E - Остаток
          ''                    // F - Статус (заполнится при обновлении статусов)
        ];
        newRows.push(newRow);
        
        console.log('  ✅ Добавлена партия: ' + batch.component + ' × ' + batch.quantity + ' ед. (ID: ' + batch.purchaseId + ')');
      }
    }
    
    // Записываем все строки сразу
    if (newRows.length > 0) {
      sheet.getRange(2, 1, newRows.length, 6).setValues(newRows);
      
      // Форматируем даты
      sheet.getRange(2, 3, newRows.length, 1).setNumberFormat('dd.mm.yyyy HH:mm');
    }
    
    console.log('💾 Записано строк: ' + newRows.length);
    
  } catch (error) {
    console.error('❌ Ошибка обновления листа ' + sheet.getName() + ': ' + error.toString());
    throw error;
  }
}

/**
 * ✅ ПОЛУЧЕНИЕ МИНИМАЛЬНОГО ЗАПАСА ДЛЯ КОМПЛЕКТУЮЩЕГО
 */
function getMinStockForComponent(component, warehouseType) {
  try {
    var catalogSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(COMPONENTS_CATALOG_SHEET);
    if (!catalogSheet) {
      console.log('❌ Справочник комплектующих не найден');
      return 0;
    }
    
    var data = catalogSheet.getRange('A2:C' + catalogSheet.getLastRow()).getValues();
    
    for (var i = 0; i < data.length; i++) {
      var catalogComponent = data[i][0]; // A - Наименование
      var mainMinStock = data[i][1] || 0; // B - Минимальный запас Гл.
      var productionMinStock = data[i][2] || 0; // C - Минимальный запас Пр.
      
      if (catalogComponent && catalogComponent.toString().trim() === component.toString().trim()) {
        if (warehouseType === 'Главный') {
          return Number(mainMinStock);
        } else if (warehouseType === 'Производство') {
          return Number(productionMinStock);
        }
      }
    }
    
    console.log('⚠️ Комплектующее "' + component + '" не найдено в справочнике');
    return 0;
    
  } catch (error) {
    console.error('❌ Ошибка в getMinStockForComponent: ' + error.toString());
    return 0;
  }
}

/**
 * 🔥 НОВАЯ ФУНКЦИЯ: ОБНОВЛЕНИЕ ЛИСТА СКЛАДА С СОХРАНЕНИЕМ ИСТОРИИ
 */
function updateWarehouseSheet(sheet, balances, existingRecords, warehouseName) {
  try {
    console.log('💾 Обновление ' + warehouseName + ': ' + Object.keys(balances).length + ' компонентов');
    
    // Очищаем существующие данные (кроме заголовка)
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange('A2:F' + lastRow).clearContent();
    }
    
    var newRows = [];
    
    for (var component in balances) {
      var quantity = roundToTwo(balances[component]);
      
      //  🔥 ИСПРАВЛЕНИЕ: Используем существующие записи или создаем новые
      var recordInfo = { transferId: generateTransferId(), purchaseId: '', date: new Date() };
      
      if (existingRecords[component] && existingRecords[component].length > 0) {
        // Используем первую найденную запись для этого компонента
        recordInfo = existingRecords[component][0];
      }
      
      var newRow = [
        recordInfo.transferId, // A - ID перемещения
        recordInfo.purchaseId, // B - ID Закупки
        recordInfo.date,       // C - Дата поступления
        component,             // D - Комплектующие
        quantity,              // E - Остаток
        ''                     // F - Статус (заполнится позже)
      ];
      newRows.push(newRow);
      
      console.log('  ✅ ' + component + ': ' + quantity + ' ед.');
    }
    
    // Записываем все строки
    if (newRows.length > 0) {
      sheet.getRange(2, 1, newRows.length, 6).setValues(newRows);
      sheet.getRange(2, 3, newRows.length, 1).setNumberFormat('dd.mm.yyyy HH:mm');
    }
    
    console.log('💾 Записано строк: ' + newRows.length);
    
  } catch (error) {
    console.error('❌ Ошибка обновления ' + warehouseName + ': ' + error.toString());
    throw error;
  }
}

/**
 * 🔥 НОВАЯ ФУНКЦИЯ: ПОКАЗ ОТЧЕТА О ПЕРЕСЧЕТЕ
 */
function showRecalculationReport(warehouseBalances) {
  var report = '✅ ПЕРЕСЧЕТ ОСТАТКОВ ЗАВЕРШЕН\n\n';
  
  for (var warehouse in warehouseBalances) {
    report += '📦 ' + warehouse + ':\n';
    var components = Object.keys(warehouseBalances[warehouse]);
    
    if (components.length === 0) {
      report += '   • Склад пуст\n';
    } else {
      components.forEach(function(component) {
        report += '   • ' + component + ': ' + warehouseBalances[warehouse][component].toFixed(2) + ' ед.\n';
      });
    }
    report += '\n';
  }
  
  SpreadsheetApp.getUi().alert('✅ Пересчет завершен', report, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Безопасный запуск полного пересчета остатков
 */
function safeRecalculateWarehouseBalances() {
  var ui = SpreadsheetApp.getUi();
  
  // 1. Запрашиваем подтверждение
  var response = ui.alert('🔐 Полный пересчет остатков',
    'Вы уверены, что хотите выполнить полный пересчет остатков?\n\n' +
    '⚠️  ВНИМАНИЕ:\n' +
    '• Это может занять несколько минут\n' +
    '• Не закрывайте таблицу во время выполнения\n' +
    '• Рекомендуется создать резервную копию\n\n' +
    'Создать резервную копию перед началом?',
    ui.ButtonSet.YES_NO_CANCEL);
  
  if (response === ui.Button.CANCEL) {
    return { cancelled: true };
  }
  
  // 2. Создаем резервную копию, если пользователь согласен
  if (response === ui.Button.YES) {
    var backupResult = createBackupBeforeRecalculation();
    if (!backupResult.success) {
      ui.alert('⚠️  Внимание',
        'Не удалось создать резервную копию:\n' + backupResult.error + '\n\n' +
        'Продолжить без резервной копии?',
        ui.ButtonSet.YES_NO);
    }
  }
  
  // 3. Проверяем данные
  var validation = validateDataBeforeRecalculation();
  if (!validation.isValid) {
    ui.alert('❌ Проверка данных не пройдена',
      'Обнаружены проблемы:\n\n' + validation.issues.join('\n') + 
      '\n\nИсправьте проблемы и повторите попытку.',
      ui.ButtonSet.OK);
    return { success: false, issues: validation.issues };
  }
  
  // 4. Запускаем пересчет
  return fastRecalculateBalancesWithTransferChain();
}

/**
 * 📋 СОЗДАНИЕ РЕЗЕРВНОЙ КОПИИ СКЛАДОВ
 */
function createBackup() {
  try {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm');
    
    // Создаем лист для резервной копии
    var backupSheetName = 'Резерв_' + today;
    var backupSheet = spreadsheet.getSheetByName(backupSheetName);
    
    if (backupSheet) {
      spreadsheet.deleteSheet(backupSheet);
    }
    
    backupSheet = spreadsheet.insertSheet(backupSheetName);
    
    // Копируем данные главного склада
    var mainSheet = spreadsheet.getSheetByName(MAIN_WAREHOUSE_SHEET);
    if (mainSheet) {
      var mainData = mainSheet.getDataRange().getValues();
      backupSheet.getRange(1, 1, mainData.length, mainData[0].length).setValues(mainData);
      backupSheet.getRange(1, 1).setValue('Главный склад - ' + today);
    }
    
    // Копируем данные склада производства
    var productionSheet = spreadsheet.getSheetByName(PRODUCTION_WAREHOUSE_SHEET);
    if (productionSheet) {
      var startRow = backupSheet.getLastRow() + 2;
      var productionData = productionSheet.getDataRange().getValues();
      backupSheet.getRange(startRow, 1, productionData.length, productionData[0].length).setValues(productionData);
      backupSheet.getRange(startRow, 1).setValue('Склад производства - ' + today);
    }
    
    console.log('✅ Создана резервная копия: ' + backupSheetName);
    SpreadsheetApp.getUi().alert('✅ Резервная копия', 'Создана резервная копия на листе: ' + backupSheetName, ui.ButtonSet.OK);
    
  } catch (error) {
    console.error('❌ Ошибка создания резервной копии: ' + error.toString());
  }
}

/**
 * Полный пересчет остатков с учетом цепочек перемещений (исправленная версия)
 */
function fastRecalculateBalancesWithTransferChain() {
  try {
    var ui = SpreadsheetApp.getUi();
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    // Блокируем интерфейс на время выполнения
    ui.alert('🔄 Пересчет остатков', 
      'Начался полный пересчет остатков. Это может занять несколько минут.\n\n' +
      'Не закрывайте таблицу и не вносите изменения во время расчета.',
      ui.ButtonSet.OK);
    
    // 1. Получаем все данные
    var warehouseSheet = spreadsheet.getSheetByName(MAIN_WAREHOUSE_SHEET);
    var transfersSheet = spreadsheet.getSheetByName(TRANSFERS_SHEET);
    var purchasesSheet = spreadsheet.getSheetByName(PURCHASES_SHEET);
    
    var lastWarehouseRow = warehouseSheet.getLastRow();
    var lastTransferRow = transfersSheet.getLastRow();
    var lastPurchaseRow = purchasesSheet.getLastRow();
    
    // 2. Загружаем данные пачками
    var warehouseData = lastWarehouseRow > 1 ? 
      warehouseSheet.getRange(2, 1, lastWarehouseRow - 1, 6).getValues() : [];
    var transfersData = lastTransferRow > 1 ? 
      transfersSheet.getRange(2, 1, lastTransferRow - 1, 8).getValues() : [];
    var purchasesData = lastPurchaseRow > 1 ? 
      purchasesSheet.getRange(2, 1, lastPurchaseRow - 1, 10).getValues() : [];
    
    // 3. Создаем карту остатков
    var balancesMap = {};
    
    // 4. Обрабатываем каждое перемещение в хронологическом порядке
    // Сначала сортируем все перемещения по дате
    transfersData.sort(function(a, b) {
      return new Date(a[2]) - new Date(b[2]); // Дата в колонке C
    });
    
    // 5. Пересчитываем остатки
    transfersData.forEach(function(transfer, index) {
      var component = transfer[3]; // Название компонента
      var quantity = Number(transfer[4]) || 0; // Количество
      var type = transfer[5]; // Тип операции (приход/расход)
      var transferId = transfer[0]; // ID перемещения
      
      if (!component) return;
      
      // Инициализируем запись для компонента
      if (!balancesMap[component]) {
        balancesMap[component] = {
          balance: 0,
          transfers: [],
          lastTransferId: ''
        };
      }
      
      // Обновляем остаток
      if (type === 'Приход') {
        balancesMap[component].balance += quantity;
      } else if (type === 'Расход') {
        balancesMap[component].balance -= quantity;
      }
      
      // Сохраняем историю
      balancesMap[component].transfers.push({
        id: transferId,
        date: transfer[2],
        type: type,
        quantity: quantity,
        balanceAfter: balancesMap[component].balance
      });
      
      balancesMap[component].lastTransferId = transferId;
    });
    
    // 6. Обновляем складской лист
    var updates = [];
    var currentTime = new Date();
    
    warehouseData.forEach(function(row, index) {
      var component = row[NEW_WAREHOUSE_STRUCTURE.COMPONENT];
      
      if (component && balancesMap[component]) {
        // Обновляем остаток
        row[NEW_WAREHOUSE_STRUCTURE.CURRENT_STOCK] = balancesMap[component].balance;
        
        // Обновляем статус
        row[NEW_WAREHOUSE_STRUCTURE.STATUS] = 'Пересчитано ' + 
          Utilities.formatDate(currentTime, Session.getScriptTimeZone(), 'dd.MM.yyyy HH:mm');
        
        // Обновляем ID последнего перемещения
        if (balancesMap[component].lastTransferId) {
          row[NEW_WAREHOUSE_STRUCTURE.TRANSFER_ID] = balancesMap[component].lastTransferId;
        }
      }
      
      updates.push(row);
    });
    
    // 7. Записываем обновления пачкой
    if (updates.length > 0) {
      warehouseSheet.getRange(2, 1, updates.length, 6).setValues(updates);
    }
    
    // 8. Очищаем кэш цепочек
    TRANSFER_CHAIN_CACHE = {};
    CacheService.getScriptCache().remove('transfer_chains');
    
    // 9. Логируем результат
    var componentsCount = Object.keys(balancesMap).length;
    logToSheet('INFO', 'fastRecalculateBalancesWithTransferChain',
      'Успешный пересчет остатков для ' + componentsCount + ' компонентов');
    
    // 10. Показываем результат
    ui.alert('✅ Пересчет завершен',
      'Успешно пересчитаны остатки для ' + componentsCount + ' компонентов.\n\n' +
      'Обновлено записей: ' + updates.length + '\n' +
      'Время выполнения: ' + (new Date() - currentTime) / 1000 + ' сек.',
      ui.ButtonSet.OK);
    
    return {
      success: true,
      components: componentsCount,
      updated: updates.length
    };
    
  } catch (error) {
    logToSheet('ERROR', 'fastRecalculateBalancesWithTransferChain',
      'Критическая ошибка: ' + error.message + '\n' + error.stack);
    
    SpreadsheetApp.getUi().alert('❌ Ошибка пересчета',
      'Произошла ошибка при пересчете остатков:\n\n' +
      error.message + '\n\n' +
      'Проверьте логи для подробностей.',
      SpreadsheetApp.getUi().ButtonSet.OK);
    
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * 🔥 НОВАЯ ФУНКЦИЯ: СОРТИРОВКА СКЛАДА ПО ДАТЕ ПОСТУПЛЕНИЯ (ОТ СТАРОЙ К НОВОЙ)
 */
function sortWarehouseByDate(sheet) {
  try {
    console.log('📅 Сортировка ' + sheet.getName() + ' по дате поступления...');
    
    var lastRow = sheet.getLastRow();
    if (lastRow <= 2) {
      console.log('ℹ️ Нечего сортировать: меньше 2 строк');
      return;
    }
    
    // Получаем все данные (кроме заголовка)
    var dataRange = sheet.getRange('A2:F' + lastRow);
    var data = dataRange.getValues();
    var displayValues = dataRange.getDisplayValues();
    
    // Создаем массив объектов для сортировки
    var rowsForSorting = [];
    for (var i = 0; i < data.length; i++) {
      rowsForSorting.push({
        rowIndex: i + 2,
        date: data[i][2], // C - Дата поступления
        values: data[i],
        displayValues: displayValues[i]
      });
    }
    
    // Сортируем по дате (от старой к новой)
    rowsForSorting.sort(function(a, b) {
      var dateA = a.date ? new Date(a.date) : new Date(0);
      var dateB = b.date ? new Date(b.date) : new Date(0);
      return dateA - dateB;
    });
    
    // Создаем новые массивы данных в отсортированном порядке
    var sortedData = [];
    var sortedDisplayValues = [];
    
    for (var j = 0; j < rowsForSorting.length; j++) {
      sortedData.push(rowsForSorting[j].values);
      sortedDisplayValues.push(rowsForSorting[j].displayValues);
    }
    
    // Записываем отсортированные данные обратно
    dataRange.setValues(sortedData);
    
    // Восстанавливаем форматирование отображения
    for (var k = 0; k < sortedDisplayValues.length; k++) {
      for (var col = 0; col < sortedDisplayValues[k].length; col++) {
        if (col === 2) { // Столбец C - Дата поступления
          sheet.getRange(k + 2, col + 1).setNumberFormat('dd.mm.yyyy HH:mm');
        }
      }
    }
    
    console.log('✅ ' + sheet.getName() + ' отсортирован по дате: ' + rowsForSorting.length + ' строк');
    
  } catch (error) {
    console.error('❌ Ошибка сортировки ' + sheet.getName() + ': ' + error.toString());
  }
}



/**
 *  💾 ОБНОВЛЕНИЕ ЛИСТОВ СКЛАДОВ С СОХРАНЕНИЕМ ОРИГИНАЛЬНЫХ ID ПЕРЕМЕЩЕНИЙ
 */
function updateWarehouseSheetsWithOriginalIds(mainSheet, productionSheet, warehouseData) {
  try {
    // Обновляем главный склад
    updateSingleWarehouseSheetWithOriginalIds(mainSheet, warehouseData['Главный'], 'Главный');
    
    // Обновляем склад производства
    updateSingleWarehouseSheetWithOriginalIds(productionSheet, warehouseData['Производство'], 'Производство');
    
    console.log('💾 Листы складов успешно обновлены с сохранением оригинальных ID перемещений');
    
  } catch (error) {
    console.error('❌ Ошибка обновления листов: ' + error.toString());
    throw error;
  }
}

/**
 *  💾 ОБНОВЛЕННАЯ ФУНКЦИЯ: ОБНОВЛЕНИЕ ОДНОГО ЛИСТА СКЛАДА С СОХРАНЕНИЕМ ОРИГИНАЛЬНЫХ ID И НУЛЕВЫХ ОСТАТКОВ
 */
function updateSingleWarehouseSheetWithOriginalIds(sheet, data, warehouseName) {
  try {
    console.log('💾 Обновление ' + warehouseName + ': ' + Object.keys(data).length + ' партий с оригинальными ID');
    
    // Очищаем существующие данные (кроме заголовка)
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange('A2:F' + lastRow).clearContent();
    }
    
    var newRows = [];
    var rowIndex = 0;
    var zeroStockCount = 0;
    
    // 🔥 ВАЖНОЕ ИЗМЕНЕНИЕ: СОХРАНЯЕМ ВСЕ ЗАПИСИ, ДАЖЕ С НУЛЕВЫМ ОСТАТКОМ
    for (var transferId in data) {
      var batch = data[transferId];
      
      // 🔥 ИСПОЛЬЗУЕМ ОРИГИНАЛЬНЫЙ ID ПЕРЕМЕЩЕНИЯ ИЗ ДАННЫХ
      var originalTransferId = batch.originalTransferId || transferId;
      
      var newRow = [
        originalTransferId,    // A - ОРИГИНАЛЬНЫЙ ID перемещения
        batch.purchaseId,      // B - ID Закупки
        batch.date,            // C - Дата (правильно заполнена)
        batch.component,       // D - Комплектующие
        roundToTwo(batch.quantity), // E - Остаток (🔥 СОХРАНЯЕМ ДАЖЕ НУЛЕВОЙ)
        ''                     // F - Статус (заполнится позже)
      ];
      newRows.push(newRow);
      rowIndex++;
      
      if (batch.quantity <= 0) {
        zeroStockCount++;
      }
      
      if (rowIndex % 100 === 0) {
        console.log('   📦 Обработано партий: ' + rowIndex + ' (ID: ' + originalTransferId + ')');
      }
    }
    
    // Записываем все строки сразу для ускорения
    if (newRows.length > 0) {
      sheet.getRange(2, 1, newRows.length, 6).setValues(newRows);
      sheet.getRange(2, 3, newRows.length, 1).setNumberFormat('dd.mm.yyyy HH:mm');
    }
    
    console.log('💾 ' + warehouseName + ': записано ' + newRows.length + ' партий (' + zeroStockCount + ' с нулевым остатком)');
    
  } catch (error) {
    console.error('❌ Ошибка обновления ' + warehouseName + ': ' + error.toString());
    throw error;
  }
}


/**
 *  🔢 ПОДСЧЕТ КОЛИЧЕСТВА ЗАПИСЕЙ С НУЛЕВЫМ ОСТАТКОМ
 */
function countZeroBalances(data) {
  var zeroCount = 0;
  for (var transferId in data) {
    if (data[transferId].quantity <= 0) {
      zeroCount++;
    }
  }
  return zeroCount;
}

/**
 *  📊 ПОКАЗ ОТЧЕТА О ПЕРЕСЧЕТЕ С ОРИГИНАЛЬНЫМИ ID
 */
function showFastRecalculationReportWithOriginalIds(warehouseData, duration) {
  var report = '🚀 УСКОРЕННЫЙ ПЕРЕСЧЕТ С СОХРАНЕНИЕМ ОРИГИНАЛЬНЫХ ID ЗАВЕРШЕН\n\n';
  report += '⏱️ Время выполнения: ' + duration + ' секунд\n\n';
  report += '✅ ОСОБЕННОСТИ ПЕРЕСЧЕТА:\n';
  report += '• Сохранены оригинальные ID перемещений из листа перемещений\n';
  report += '• Сохранены оригинальные ID закупок\n';
  report += '• Учтены все закупки за весь период\n';
  report += '• Учтены все перемещения между складами\n';
  report += '• Правильно переданы даты поступления\n';
  report += '• Сортировка по дате (от старых к новым)\n';
  report += '• Нулевые остатки не удаляются\n\n';
  
  for (var warehouse in warehouseData) {
    report += '📦 ' + warehouse + ':\n';
    var batches = Object.keys(warehouseData[warehouse]);
    
    if (batches.length === 0) {
      report += '   • Склад пуст\n';
    } else {
      var totalBatches = batches.length;
      var zeroBatches = countZeroBalances(warehouseData[warehouse]);
      var nonZeroBatches = totalBatches - zeroBatches;
      
      report += '   • Всего партий: ' + totalBatches + '\n';
      report += '   • Партий с остатком: ' + nonZeroBatches + '\n';
      report += '   • Партий с нулевым остатком: ' + zeroBatches + '\n';
      
      // Показываем примеры ID
      var sampleIds = [];
      for (var transferId in warehouseData[warehouse]) {
        if (sampleIds.length < 3) {
          sampleIds.push(transferId);
        }
      }
      
      if (sampleIds.length > 0) {
        report += '   • Примеры ID перемещений: ' + sampleIds.join(', ') + '\n';
      }
    }
    report += '\n';
  }
  
  SpreadsheetApp.getUi().alert('✅ Пересчет завершен', report, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * 📝 ОБНОВЛЕННАЯ ФУНКЦИЯ: ЗАПИСЬ СПИСАНИЯ МАТЕРИАЛОВ В ЖУРНАЛ С НОВОЙ СТРУКТУРОЙ
 */
function recordMaterialsWriteOff(planningRowData, component, writeOffQuantity, transferId, purchaseId) {
  try {
    var writeoffSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MATERIALS_WRITEOFF_SHEET);
    
    if (!writeoffSheet) {
      console.error('❌ Лист "' + MATERIALS_WRITEOFF_SHEET + '" не найден');
      return false;
    }
    
    var lastRow = writeoffSheet.getLastRow() + 1;
    
    // Получаем данные из строки планирования
    var deliveryDate = planningRowData[0]; // A - Дата поставки
    var deliveryNumber = planningRowData[1]; // B - Поставки
    var product = planningRowData[3]; // D - Товар
    var shiftDate = planningRowData[4]; // E - Дата смены
    var shiftInfo = planningRowData[5]; // F - Выбор смены с этим товаром
    var quantityBoxes = planningRowData[7] || 0; // H - Кол-во коробок
    
    // Извлекаем работника из информации о смене
    var worker = '';
    if (shiftInfo && shiftInfo.includes(' | Остаток:')) {
      worker = shiftInfo.split(' | ')[0]; // Берем часть до " | Остаток:"
    }
    
    // Получаем количество единиц в коробке
    var itemsPerBox = getItemsPerBox(product);
    var totalItems = quantityBoxes * itemsPerBox;
    
    // Получаем норму расхода для компонента
    var consumptionRates = getConsumptionRatesForProduct(product);
    var rate = consumptionRates[component] || 0;
    
    // 🔥 НОВАЯ СТРУКТУРА СПИСАНИЙ:
    // A: ID перемещения | B: ID Закупки | C: Дата | D: Поставка | E: Товар | F: Дата смены | G: Работник | H: Кол-во ед. | I: Комплектующее | J: Списание
    var newRow = [
      transferId || generateTransferId(), // A - ID перемещения
      purchaseId || '',                   // B - ID Закупки
      new Date(),                         // C - Дата списания (текущая дата)
      deliveryNumber,                     // D - Поставка
      product,                            // E - Товар
      shiftDate,                          // F - Дата смены (НОВОЕ ПОЛЕ)
      worker,                             // G - Работник (НОВОЕ ПОЛЕ)
      roundToTwo(totalItems),             // H - Кол-во ед.
      component,                          // I - Комплектующее
      roundToTwo(writeOffQuantity)        // J - Списание
    ];
    
    // Записываем данные
    writeoffSheet.getRange(lastRow, 1, 1, 10).setValues([newRow]);
    
    // Форматируем даты
    writeoffSheet.getRange(lastRow, 3).setNumberFormat('dd.mm.yyyy HH:mm'); // C - Дата списания
    if (shiftDate) {
      writeoffSheet.getRange(lastRow, 6).setNumberFormat('dd.mm.yyyy'); // F - Дата смены
    }
    
    console.log('📝 Записано списание: ' + component + ' × ' + writeOffQuantity + 
               ' для ' + product + ' (работник: ' + worker + ', дата смены: ' + shiftDate + ')');
    return true;
    
  } catch (error) {
    console.error('❌ Ошибка записи списания: ' + error.toString());
    return false;
  }
}

/**
 * ✅ НОВАЯ ФУНКЦИЯ: ПРОВЕРКА СУЩЕСТВОВАНИЯ ЗАПИСИ СПИСАНИЯ
 */
function isWriteOffRecordExists(writeOffSheet, deliveryNumber, product, shiftDate, worker, quantity) {
  try {
    var lastRow = writeOffSheet.getLastRow();
    if (lastRow <= 1) return false;
    
    var data = writeOffSheet.getRange('A2:J' + lastRow).getValues();
    
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var existingDelivery = row[3]; // D - Поставка
      var existingProduct = row[4]; // E - Товар
      var existingShiftDate = row[5]; // F - Дата смены
      var existingWorker = row[6]; // G - Работник
      var existingQuantity = row[7]; // H - Кол-во ед.
      
      // Проверяем совпадение по ключевым полям
      if (existingDelivery === deliveryNumber &&
          existingProduct === product &&
          isSameDate(existingShiftDate, shiftDate) &&
          existingWorker === worker &&
          Math.abs(existingQuantity - (quantity * getItemsPerBox(product))) < 0.01) {
        console.log('🔍 Найдена дублирующая запись списания:');
        console.log('   Поставка: ' + deliveryNumber + ', Товар: ' + product);
        console.log('   Дата смены: ' + shiftDate + ', Работник: ' + worker);
        return true;
      }
    }
    
    return false;
    
  } catch (error) {
    console.error('❌ Ошибка проверки дублирования записей: ' + error.toString());
    return false;
  }
}

/**
 * 🔥 ИСПРАВЛЕННАЯ ФУНКЦИЯ: ПОЛУЧЕНИЕ ПАРТИЙ КОМПОНЕНТА, ОТСОРТИРОВАННЫХ ПО ДАТЕ (FIFO)
 * Теперь правильно сортирует от старых к новым (FIFO)
 */
function getComponentBatchesSortedByDate(sheet, component) {
  try {
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return [];
    
    var data = sheet.getRange('A2:F' + lastRow).getValues();
    var batches = [];
    
    for (var i = 0; i < data.length; i++) {
      var rowComponent = data[i][3]; // D - Комплектующие
      var stock = data[i][4] || 0; // E - Остаток
      var date = data[i][2]; // C - Дата
      var transferId = data[i][0]; // A - ID перемещения
      var purchaseId = data[i][1]; // B - ID Закупки
      
      if (rowComponent === component && stock > 0) {
        batches.push({
          row: i + 2,
          component: component,
          stock: stock,
          date: date,
          transferId: transferId,
          purchaseId: purchaseId
        });
      }
    }
    
    // 🔥 ИСПРАВЛЕНИЕ: Сортируем по дате (старые первыми - FIFO)
    // РАНЬШЕ БЫЛО: new Date(a.date) - new Date(b.date) - это неправильно
    // ТЕПЕРЬ: new Date(b.date) - new Date(a.date) для правильного FIFO
    batches.sort(function(a, b) {
      var dateA = a.date ? new Date(a.date) : new Date(0);
      var dateB = b.date ? new Date(b.date) : new Date(0);
      return dateA - dateB; // Отрицательное значение = a идет перед b (старые первыми)
    });
    
    console.log('📅 Партии для ' + component + ' (FIFO порядок):');
    batches.forEach(function(batch, index) {
      var dateStr = batch.date ? Utilities.formatDate(batch.date, Session.getScriptTimeZone(), "dd.MM.yyyy") : 'Нет даты';
      console.log('   ' + (index + 1) + '. ' + dateStr + ' - ' + batch.stock + ' ед. (ID: ' + batch.transferId + ')');
    });
    
    return batches;
    
  } catch (error) {
    console.error('❌ Ошибка получения партий: ' + error.toString());
    return [];
  }
}

/**
 * 🔥 ОБНОВЛЕННАЯ ФУНКЦИЯ: ЗАПИСЬ СПИСАНИЯ В ЖУРНАЛ С НОВОЙ СТРУКТУРОЙ
 */
function recordWriteOffToJournal(planningSheet, row, product, quantity, writeOffResults) {
  try {
    var writeOffSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Списание_материалов');
    if (!writeOffSheet) return;
    
    var planningData = planningSheet.getRange(row, 1, 1, 11).getValues()[0]; // 🔥 ОБНОВЛЕНО: теперь до столбца K
    var planningId = planningData[10]; // K - ID планирования
    var deliveryDate = planningData[0]; // A - Дата поставки
    var deliveryNumber = planningData[1]; // B - Поставки
    var shiftDate = planningData[4]; // E - Дата смены
    var shiftInfo = planningData[5]; // F - Выбор смены
    
    var worker = '';
    if (shiftInfo && shiftInfo.includes(' | Остаток:')) {
      worker = shiftInfo.split(' | ')[0];
    }
    
    var itemsPerBox = getItemsPerBox(product);
    var totalItems = quantity * itemsPerBox;
    
    var lastRow = writeOffSheet.getLastRow();
    var newRows = [];
    
    for (var i = 0; i < writeOffResults.length; i++) {
      var result = writeOffResults[i];
      
      // 🔥 НОВАЯ СТРУКТУРА СПИСАНИЙ:
      var newRow = [
        planningId,                 // A - ID планирования
        result.transferId,          // B - ID перемещения
        result.purchaseId,          // C - ID Закупки
        new Date(),                 // D - Дата списания
        deliveryNumber,             // E - Поставка
        product,                    // F - Товар
        shiftDate,                  // G - Дата смены
        worker,                     // H - Работник
        roundToTwo(totalItems),     // I - Кол-во ед.
        result.component,           // J - Комплектующее
        result.consumed             // K - Списание
      ];
      newRows.push(newRow);
    }
    
    if (newRows.length > 0) {
      writeOffSheet.getRange(lastRow + 1, 1, newRows.length, 11).setValues(newRows); // 🔥 ОБНОВЛЕНО: 11 столбцов
      
      // Форматируем даты
      writeOffSheet.getRange(lastRow + 1, 4, newRows.length, 1).setNumberFormat('dd.mm.yyyy HH:mm'); // D - Дата списания
      if (shiftDate) {
        writeOffSheet.getRange(lastRow + 1, 7, newRows.length, 1).setNumberFormat('dd.mm.yyyy'); // G - Дата смены
      }
    }
    
    console.log('📝 Записано ' + newRows.length + ' записей в журнал списаний с ID планирования: ' + planningId);
    
  } catch (error) {
    console.error('❌ Ошибка записи в журнал: ' + error.toString());
  }
}

/**
 * ✅ ОБНОВЛЕННАЯ ФУНКЦИЯ: УДАЛЕНИЕ СТРОКИ ПЛАНИРОВАНИЯ С УДАЛЕНИЕМ СООТВЕТСТВУЮЩИХ СПИСАНИЙ
 */
function removePlanningRowDirect() {
  try {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var ui = SpreadsheetApp.getUi();
    
    var response = ui.prompt(
      '🗑️ Удаление строки планирования',
      'Введите номер строки для удаления:',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (response.getSelectedButton() !== ui.Button.OK) return;
    
    var row = parseInt(response.getResponseText());
    if (isNaN(row) || row < 2) {
      ui.alert('❌ Ошибка', 'Введите корректный номер строки (начиная с 2)', ui.ButtonSet.OK);
      return;
    }
    
    var planningSheet = spreadsheet.getSheetByName('Планирование');
    if (!planningSheet) throw new Error('Лист "Планирование" не найден');
    
    var lastRow = planningSheet.getLastRow();
    if (row > lastRow) {
      ui.alert('❌ Ошибка', 'Строка ' + row + ' не существует', ui.ButtonSet.OK);
      return;
    }
    
    // 🔥 ПОЛУЧАЕМ ДАННЫЕ СТРОКИ ДО УДАЛЕНИЯ (ВКЛЮЧАЯ ID ПЛАНИРОВАНИЯ)
    var rowData = planningSheet.getRange(row, 1, 1, 11).getValues()[0]; // 🔥 ОБНОВЛЕНО: до столбца K
    var product = rowData[3]; // D - Товар
    var quantity = rowData[7]; // H - Количество
    var status = rowData[8]; // I - Статус
    var planningId = rowData[10]; // K - ID планирования
    
    console.log('🗑️ Удаление планирования: ' + product + ' × ' + quantity + ' кор., статус: ' + status + ', ID: ' + planningId);
    
    // 🔥 УДАЛЯЕМ СООТВЕТСТВУЮЩИЕ ЗАПИСИ СПИСАНИЙ ПО ID ПЛАНИРОВАНИЯ
    var writeoffDeleted = false;
    if (planningId) {
      writeoffDeleted = deleteWriteOffRecordsByPlanningId(planningId);
      console.log('📝 Удаление записей списаний для ID ' + planningId + ': ' + (writeoffDeleted ? '✅ Успешно' : '❌ Не найдено'));
    }
    
    // Удаляем строку планирования
    planningSheet.deleteRow(row);
    
    // Обновляем списания в графике смен
    updateScheduleWriteOffs();
    
    // 🔥 ПЕРЕСЧИТЫВАЕМ СКЛАД ПРОИЗВОДСТВА
    if (status === '✅ Выполнено' && product && quantity > 0) {
      console.log('🔄 Запуск пересчёта склада производства...');
      
      try {
        // Запускаем полный пересчёт всех складов
        safeRecalculateWarehouseBalances();
        
        console.log('✅ Склад производства успешно пересчитан');
        
        var message = 'Строка ' + row + ' удалена\n\n' +
          '📊 Склад производства пересчитан:\n' +
          '• Товар: ' + product + '\n' +
          '• Количество: ' + quantity + ' кор.\n' +
          '• ID планирования: ' + planningId;
          
        if (writeoffDeleted) {
          message += '\n• Записи списаний удалены';
        }
        
        ui.alert('✅ Удаление завершено', message, ui.ButtonSet.OK);
        
      } catch (recalcError) {
        console.error('❌ Ошибка пересчёта склада: ' + recalcError.toString());
        
        ui.alert(
          '⚠️ Предупреждение',
          'Строка удалена, но не удалось пересчитать склад:\n\n' + recalcError.toString() + '\n\n' +
          '💡 Запустите вручную: Действия → Управление складами → Полный пересчёт остатков',
          ui.ButtonSet.OK
        );
      }
    } else {
      var message = 'Строка ' + row + ' успешно удалена';
      if (writeoffDeleted) {
        message += '\n• Записи списаний удалены';
      }
      ui.alert('✅ Успех', message, ui.ButtonSet.OK);
    }
    
  } catch (error) {
    console.error('❌ Ошибка удаления:', error.toString());
    SpreadsheetApp.getUi().alert('❌ Ошибка: ' + error.toString());
  }
}

/**
 * ✅ ПРОВЕРКА ЗАПОЛНЕННОСТИ СТРОКИ ПЛАНИРОВАНИЯ
 */
function isPlanningRowFilled(sheet, row) {
  try {
    var date = sheet.getRange(row, 1).getValue();        // A - Дата поставки
    var deliveryNumber = sheet.getRange(row, 2).getValue(); // B - Поставки
    var warehouse = sheet.getRange(row, 3).getValue();   // C - Склад
    var product = sheet.getRange(row, 4).getValue();     // D - Выбор товара
    var shiftDate = sheet.getRange(row, 5).getValue();   // E - Дата смены
    var shiftInfo = sheet.getRange(row, 6).getValue();   // F - Выбор смены
    var quantity = sheet.getRange(row, 8).getValue();    // H - Кол-во коробок
    
    // Проверяем, что все обязательные поля заполнены
    var isFilled = !!(date && deliveryNumber && warehouse && 
                     product && product !== 'Выбор товара' && 
                     shiftDate && shiftInfo && shiftInfo !== 'Выбор смены с этим товаром' && 
                     quantity > 0);
    
    console.log('🔍 Проверка заполненности строки ' + row + ': ' + (isFilled ? '✅ Заполнена' : '❌ Не заполнена'));
    
    return isFilled;
    
  } catch (error) {
    console.error('Ошибка в isPlanningRowFilled: ' + error.toString());
    return false;
  }
}

/**
 * ✅ ГЕНЕРАЦИЯ УНИКАЛЬНОГО ID ДЛЯ ПЛАНИРОВАНИЯ
 */
function generatePlanningId() {
  var timestamp = new Date().getTime();
  var random = Math.floor(Math.random() * 1000);
  return PLANNING_ID_PREFIX + timestamp + '-' + random;
}

/**
 * ✅ НОВАЯ ФУНКЦИЯ: УДАЛЕНИЕ ЗАПИСЕЙ СПИСАНИЙ ПО ID ПЛАНИРОВАНИЯ
 */
function deleteWriteOffRecordsByPlanningId(planningId) {
  try {
    var writeoffSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Списание_материалов');
    if (!writeoffSheet) {
      console.error('❌ Лист "Списание_материалов" не найден');
      return false;
    }
    
    var lastRow = writeoffSheet.getLastRow();
    if (lastRow <= 1) {
      console.log('ℹ️ Лист списаний пуст, удалять нечего');
      return false;
    }
    
    // 🔥 ПОЛУЧАЕМ ВСЕ ДАННЫЕ ИЗ ЛИСТА СПИСАНИЙ
    var data = writeoffSheet.getRange('A2:K' + lastRow).getValues();
    var rowsToDelete = [];
    
    // 🔥 ИЩЕМ СТРОКИ С СООТВЕТСТВУЮЩИМ ID ПЛАНИРОВАНИЯ (СТОЛБЕЦ A)
    for (var i = 0; i < data.length; i++) {
      var rowPlanningId = data[i][0]; // A - ID планирования
      
      if (rowPlanningId && rowPlanningId.toString().trim() === planningId.toString().trim()) {
        rowsToDelete.push(i + 2); // +2 потому что данные начинаются со 2 строки
        console.log('🔍 Найдена запись списания для удаления: строка ' + (i + 2) + ', ID: ' + planningId);
      }
    }
    
    if (rowsToDelete.length === 0) {
      console.log('ℹ️ Не найдено записей списаний для ID планирования: ' + planningId);
      return false;
    }
    
    // 🔥 УДАЛЯЕМ СТРОКИ В ОБРАТНОМ ПОРЯДКЕ (ЧТОБЫ НЕ СБИВАЛИСЬ НОМЕРА)
    rowsToDelete.sort(function(a, b) { return b - a; });
    
    var deletedCount = 0;
    for (var j = 0; j < rowsToDelete.length; j++) {
      writeoffSheet.deleteRow(rowsToDelete[j]);
      deletedCount++;
      console.log('🗑️ Удалена запись списания: строка ' + rowsToDelete[j]);
    }
    
    console.log('✅ Удалено записей списаний: ' + deletedCount + ' для ID планирования: ' + planningId);
    return true;
    
  } catch (error) {
    console.error('❌ Ошибка удаления записей списаний: ' + error.toString());
    return false;
  }
}

/**
 * ✅ ИСПРАВЛЕННАЯ ФУНКЦИЯ: FIFO СПИСАНИЕ КОНКРЕТНОЙ ПАРТИИ С ПРАВИЛЬНЫМИ ИНДЕКСАМИ СТОЛБЦОВ
 */
function consumeMaterialsFIFOSpecific(warehouseName, component, purchaseId, requiredQuantity) {
  try {
    var warehouseSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
      warehouseName === 'Главный' ? MAIN_WAREHOUSE_SHEET : PRODUCTION_WAREHOUSE_SHEET
    );
    
    // 🔥 ИСПРАВЛЕНИЕ: Правильное получение данных с учетом структуры
    var data = warehouseSheet.getDataRange().getValues();
    
    // 🔥 ПРАВИЛЬНЫЕ ИНДЕКСЫ СТОЛБЦОВ ПО НОВОЙ СТРУКТУРЕ СКЛАДА:
    // A: ID перемещения (0), B: ID Закупки (1), C: Дата (2), D: Комплектующие (3), E: Остаток (4), F: Статус (5)
    for (var i = 1; i < data.length; i++) {
      var rowComponent = data[i][3]; // D - Комплектующие
      var rowPurchaseId = data[i][1]; // B - ID Закупки
      var currentStock = data[i][4] || 0; // E - Остаток
      
      console.log('🔍 Поиск партии: ищем ' + component + ' с ID ' + purchaseId);
      console.log('   Найдено: компонент=' + rowComponent + ', ID=' + rowPurchaseId + ', остаток=' + currentStock);
      
      if (rowComponent && rowComponent.toString().trim() === component.toString().trim() && 
          rowPurchaseId && rowPurchaseId.toString().trim() === purchaseId.toString().trim() && 
          currentStock > 0) {
        
        if (currentStock >= requiredQuantity) {
          // Списание из конкретной партии
          var newStock = currentStock - requiredQuantity;
          warehouseSheet.getRange(i + 1, 5).setValue(newStock); // E - Остаток (столбец 5)
          
          // 🔄 ОБНОВЛЯЕМ СТАТУС ДЛЯ ЭТОЙ СТРОКИ
          updateSingleRowStatus(warehouseSheet, i + 1);
          
          console.log('✅ Списано из партии ' + purchaseId + ': ' + requiredQuantity + ' ед. (остаток: ' + newStock + ' ед.)');
          
          return {
            success: true,
            consumed: requiredQuantity,
            batches: [{
              purchaseId: purchaseId,
              consumed: requiredQuantity,
              remaining: newStock
            }]
          };
        } else {
          throw new Error(`Недостаточно в партии ${purchaseId}: доступно ${currentStock}, требуется ${requiredQuantity}`);
        }
      }
    }
    
    throw new Error(`Партия ${purchaseId} не найдена или пуста для компонента ${component}`);
    
  } catch (error) {
    console.error('❌ Ошибка в consumeMaterialsFIFOSpecific: ' + error.toString());
    throw error;
  }
}


/**
 * ✅ ИСПРАВЛЕННАЯ ФУНКЦИЯ: ПРОСТОЕ ОБНОВЛЕНИЕ ОСТАТКА КОМПОНЕНТА НА СКЛАДЕ
 */
function updateStockSimple(sheet, component, quantityChange) {
  try {
    console.log('🔄 Простое обновление остатка: ' + component + ' на ' + quantityChange + ' ед.');
    
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      throw new Error('Склад пуст');
    }
    
    //  🔥 ИСПРАВЛЕННЫЕ ИНДЕКСЫ ПО НОВОЙ СТРУКТУРЕ:
    // A: ID перемещения (0), B: ID Закупки (1), C: Дата (2), D: Комплектующие (3), E: Остаток (4)
    var data = sheet.getRange('A2:F' + lastRow).getValues();
    
    var updated = false;
    
    for (var i = 0; i < data.length; i++) {
      var rowComponent = data[i][3]; // D - Комплектующие
      
      if (rowComponent && rowComponent.toString().trim().toLowerCase() === component.toString().trim().toLowerCase()) {
        var currentStock = data[i][4] || 0; // E - Остаток
        var newStock = currentStock + quantityChange;
        
        if (newStock < 0) {
          throw new Error(`Недостаточно ${component} для списания: доступно ${currentStock}, требуется ${-quantityChange}`);
        }
        
        sheet.getRange(i + 2, 5).setValue(newStock); // E - Остаток (столбец 5)
        updateSingleRowStatus(sheet, i + 2);
        
        console.log('✅ Обновлен остаток ' + component + ': ' + currentStock + ' → ' + newStock + ' ед.');
        updated = true;
        break;
      }
    }
    
    if (!updated) {
      throw new Error(`Компонент ${component} не найден на складе`);
    }
    
  } catch (error) {
    console.error('❌ Ошибка в updateStockSimple: ' + error.toString());
    throw error;
  }
}

/**
 * 🔥 НОВАЯ ФУНКЦИЯ: ПРИНУДИТЕЛЬНОЕ ОБНОВЛЕНИЕ СТАТУСОВ ПОСЛЕ СОРТИРОВКИ С ПРАВИЛЬНЫМИ ПОДСКАЗКАМИ
 */
function forceUpdateStatusesAfterSorting() {
  try {
    console.log('🔄 ПРИНУДИТЕЛЬНОЕ ОБНОВЛЕНИЕ СТАТУСОВ ПОСЛЕ СОРТИРОВКИ...');
    
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var mainSheet = spreadsheet.getSheetByName(MAIN_WAREHOUSE_SHEET);
    var productionSheet = spreadsheet.getSheetByName(PRODUCTION_WAREHOUSE_SHEET);
    
    var updatedCount = 0;
    
    if (mainSheet) {
      console.log('📊 Принудительное обновление главного склада...');
      // 🔥 ВАЖНО: Используем addTooltips = true для правильных подсказок
      updateWarehouseStatusDirect(mainSheet, true);
      updatedCount++;
      console.log('✅ Главный склад обновлен с правильными подсказками');
    }
    
    if (productionSheet) {
      console.log('📊 Принудительное обновление склада производства...');
      // 🔥 ВАЖНО: Используем addTooltips = true для правильных подсказок
      updateWarehouseStatusDirect(productionSheet, true);
      updatedCount++;
      console.log('✅ Склад производства обновлен с правильными подсказками');
    }
    
    console.log('✅ Принудительное обновление завершено. Обновлено складов: ' + updatedCount);
    
    SpreadsheetApp.getUi().alert(
      '✅ Статусы обновлены', 
      'Статусы складов успешно обновлены с правильными подсказками!\n\n' +
      'Теперь подсказки соответствуют своим строкам после сортировки.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
    return updatedCount;
    
  } catch (error) {
    console.error('❌ Ошибка принудительного обновления статусов: ' + error.toString());
    SpreadsheetApp.getUi().alert(
      '❌ Ошибка', 
      'Не удалось обновить статусы:\n\n' + error.toString(),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return 0;
  }
}

// 🔥 АЛЬТЕРНАТИВНАЯ ФУНКЦИЯ: ПОЛУЧЕНИЕ ИЗНАЧАЛЬНОГО КОЛИЧЕСТВА ИЗ ПЕРЕМЕЩЕНИЙ
function getInitialQuantityFromTransfer(transferId, component) {
  try {
    var transfersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TRANSFERS_SHEET);
    if (!transfersSheet) {
      console.error('❌ Лист перемещений не найден');
      return 0;
    }
    
    var lastRow = transfersSheet.getLastRow();
    if (lastRow < 2) {
      console.log('ℹ️ Лист перемещений пуст');
      return 0;
    }
    
    // Ищем перемещение по ID и компоненту
    var data = transfersSheet.getRange('A2:G' + lastRow).getValues();
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var rowTransferId = row[0]; // A - ID перемещения
      var rowComponent = row[5];  // F - Комплектующее
      var rowQuantity = row[6];   // G - Количество
      
      if (rowTransferId && rowTransferId.toString().trim() === transferId.toString().trim() &&
          rowComponent && rowComponent.toString().trim().toLowerCase() === component.toString().trim().toLowerCase()) {
        
        console.log('✅ Найдено перемещение:', transferId, 'Компонент:', component, 'Количество:', rowQuantity);
        return Number(rowQuantity);
      }
    }
    
    console.log('⚠️ Перемещение не найдено:', transferId, 'Компонент:', component);
    return 0;
    
  } catch (error) {
    console.error('❌ Ошибка получения изначального количества из перемещения: ' + error.toString());
    return 0;
  }
}

/**
 * 🔥 НОВАЯ ФУНКЦИЯ: ПРОВЕРКА СТРУКТУРЫ ЛИСТА ПЕРЕМЕЩЕНИЙ
 */
function checkTransfersSheetStructure() {
  try {
    var transfersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TRANSFERS_SHEET);
    if (!transfersSheet) {
      throw new Error('Лист перемещений не найден');
    }
    
    var headers = transfersSheet.getRange('A1:I1').getValues()[0];
    console.log('📋 Структура листа перемещений:');
    console.log('A: ' + headers[0] + ' (ID перемещения)');
    console.log('B: ' + headers[1] + ' (ID Закупки)');
    console.log('C: ' + headers[2] + ' (Дата)');
    console.log('D: ' + headers[3] + ' (От склада >)');
    console.log('E: ' + headers[4] + ' (На склад)');
    console.log('F: ' + headers[5] + ' (Комплектующее)');
    console.log('G: ' + headers[6] + ' (Количество)');
    console.log('H: ' + headers[7] + ' (Статус)');
    console.log('I: ' + headers[8] + ' (Примечание)');
    
    return headers;
  } catch (error) {
    console.error('❌ Ошибка проверки структуры: ' + error.toString());
    return null;
  }
}

/**
 * ✅ НОВАЯ ФУНКЦИЯ: СКРЫТИЕ СТРОК С НУЛЕВЫМ ОСТАТКОМ НА ОДНОМ ЛИСТЕ
 */
function hideZeroStockRowsOnSheet(sheet) {
  try {
    if (!sheet) {
      console.error('❌ Лист не передан для скрытия строк');
      return 0;
    }
    
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      console.log('ℹ️ Лист ' + sheet.getName() + ' пуст, нечего скрывать');
      return 0;
    }
    
    // 🔥 ПРАВИЛЬНАЯ СТРУКТУРА: E - Остаток (индекс 4)
    var data = sheet.getRange('A2:F' + lastRow).getValues();
    var rowsToHide = [];
    var hiddenCount = 0;
    
    console.log('🔍 Поиск строк с нулевым остатком на ' + sheet.getName() + '...');
    
    for (var i = 0; i < data.length; i++) {
      var row = i + 2;
      var currentStock = data[i][4] || 0; // E - Остаток
      
      if (currentStock <= 0) {
        rowsToHide.push(row);
        console.log('👻 Строка ' + row + ' имеет нулевой остаток: ' + currentStock);
      }
    }
    
    // Сортируем строки в обратном порядке для безопасного скрытия
    rowsToHide.sort(function(a, b) { return b - a; });
    
    for (var j = 0; j < rowsToHide.length; j++) {
      var rowToHide = rowsToHide[j];
      
      // Проверяем, что строка еще не скрыта
      if (!sheet.isRowHiddenByUser(rowToHide)) {
        sheet.hideRows(rowToHide);
        hiddenCount++;
        console.log('👻 Скрыта строка ' + rowToHide);
      }
    }
    
    console.log('✅ Скрыто строк с нулевым остатком на ' + sheet.getName() + ': ' + hiddenCount);
    return hiddenCount;
    
  } catch (error) {
    console.error('❌ Ошибка скрытия строк на ' + sheet.getName() + ': ' + error.toString());
    return 0;
  }
}

/**
* 🔥 НОВАЯ ФУНКЦИЯ: ПОИСК ID ИСХОДНОГО ПЕРЕМЕЩЕНИЯ ДЛЯ ЦЕПОЧКИ
*/
function findSourceTransferId(transfersData, purchaseId, component, currentTransferDate) {
  try {
    var earliestTransfer = null;
    for (var i = 0; i < transfersData.length; i++) {
      var tRow = transfersData[i];
      var transferId = tRow[0];
      var rowPurchaseId = tRow[1];
      var rowComponent = tRow[5];
      var transferDate = tRow[2];
      var status = tRow[7];
      
      // Ищем самое раннее перемещение для этой закупки и компонента
      if (rowPurchaseId === purchaseId && 
          rowComponent === component && 
          status === 'Выполнено' &&
          (!earliestTransfer || new Date(transferDate) < new Date(earliestTransfer.date))) {
        earliestTransfer = {
          transferId: transferId,
          date: transferDate
        };
      }
    }
    return earliestTransfer ? earliestTransfer.transferId : null;
  } catch (error) {
    console.error('❌ Ошибка поиска исходного ID перемещения: ' + error.toString());
    return null;
  }
}

/**
* 🔥 ОБНОВЛЕННАЯ ФУНКЦИЯ: СОЗДАНИЕ КАРТЫ ПЕРЕМЕЩЕНИЙ С ПРАВИЛЬНОЙ ЦЕПОЧКОЙ
*/
function createTransferChainMap(transfersData) {
  try {
    var transferChain = {};
    var purchaseTransferMap = {};
    
    console.log('🔗 Создание карты цепочки перемещений...');
    
    for (var i = 0; i < transfersData.length; i++) {
      var tRow = transfersData[i];
      var transferId = tRow[0];
      var purchaseId = tRow[1];
      var fromWarehouse = tRow[3];
      var toWarehouse = tRow[4];
      var component = tRow[5];
      var quantity = tRow[6];
      var status = tRow[7];
      var transferDate = tRow[2];
      
      if (status !== 'Выполнено' || !purchaseId || !component) continue;
      
      // Для перемещений от поставщика - это начало цепочки
      if (fromWarehouse && fromWarehouse.toString().trim().toLowerCase().indexOf('поставщик') !== -1) {
        if (!purchaseTransferMap[purchaseId]) {
          purchaseTransferMap[purchaseId] = [];
        }
        purchaseTransferMap[purchaseId].push({
          transferId: transferId,
          component: component,
          quantity: quantity,
          date: transferDate,
          type: 'initial'
        });
        console.log('   🔗 Начало цепочки: ' + purchaseId + ' → ' + transferId);
      }
      
      // Для перемещений между складами - отслеживаем цепочку
      else if (fromWarehouse && toWarehouse) {
        var fromWh = fromWarehouse.toString().trim().toLowerCase().indexOf('производ') !== -1 ? 'Производство' : 'Главный';
        var toWh = toWarehouse.toString().trim().toLowerCase().indexOf('производ') !== -1 ? 'Производство' : 'Главный';
        
        // Ищем предыдущее перемещение в цепочке
        var sourceTransferId = findSourceTransferId(transfersData, purchaseId, component, transferDate);
        
        if (sourceTransferId) {
          if (!transferChain[sourceTransferId]) {
            transferChain[sourceTransferId] = [];
          }
          transferChain[sourceTransferId].push({
            newTransferId: transferId,
            component: component,
            quantity: quantity,
            from: fromWh,
            to: toWh,
            date: transferDate
          });
          console.log('   🔗 Цепочка: ' + sourceTransferId + ' → ' + transferId);
        }
      }
    }
    
    return {
      purchaseMap: purchaseTransferMap,
      chainMap: transferChain
    };
    
  } catch (error) {
    console.error('❌ Ошибка создания карты цепочки: ' + error.toString());
    return { purchaseMap: {}, chainMap: {} };
  }
}

/**
 * 🔍 ПОЛУЧЕНИЕ ПОЛНОЙ ЦЕПОЧКИ ПЕРЕМЕЩЕНИЙ ДЛЯ ID ЗАКУПКИ
 */
function getFullTransferChainForPurchase(purchaseId, component) {
  try {
    var cacheKey = purchaseId + '_' + component;
    if (TRANSFER_CHAIN_CACHE[cacheKey]) {
      return TRANSFER_CHAIN_CACHE[cacheKey];
    }
    
    var transfersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Перемещения');
    if (!transfersSheet || transfersSheet.getLastRow() <= 1) {
      return [];
    }
    
    var data = transfersSheet.getRange('A2:I' + transfersSheet.getLastRow()).getValues();
    var chain = [];
    
    // Находим все перемещения для этой закупки и компонента
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      if (row[1] === purchaseId && row[5] === component && row[7] === 'Выполнено') {
        chain.push({
          transferId: row[0],
          date: row[2],
          from: row[3],
          to: row[4],
          quantity: row[6],
          notes: row[8]
        });
      }
    }
    
    // Сортируем по дате (от старых к новым)
    chain.sort(function(a, b) {
      return new Date(a.date) - new Date(b.date);
    });
    
    TRANSFER_CHAIN_CACHE[cacheKey] = chain;
    return chain;
    
  } catch (error) {
    console.error('❌ Ошибка получения цепочки перемещений: ' + error.toString());
    return [];
  }
}

/**
 * 🔗 ВОССТАНОВЛЕНИЕ СВЯЗЕЙ МЕЖДУ ПЕРЕМЕЩЕНИЯМИ
 */
function restoreTransferLinks() {
  try {
    console.log('🔗 Восстановление связей между перемещениями...');
    
    var transfersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Перемещения');
    if (!transfersSheet || transfersSheet.getLastRow() <= 1) {
      console.log('ℹ️ Лист перемещений пуст');
      return 0;
    }
    
    var data = transfersSheet.getRange('A2:I' + transfersSheet.getLastRow()).getValues();
    var updatedCount = 0;
    
    for (var i = 0; i < data.length; i++) {
      var row = i + 2;
      var purchaseId = data[i][1];
      var component = data[i][5];
      var fromWarehouse = data[i][3];
      var status = data[i][7];
      
      // Для перемещений не от поставщика ищем исходное перемещение
      if (status === 'Выполнено' && purchaseId && component && 
          fromWarehouse && fromWarehouse.toString().trim().toLowerCase().indexOf('поставщик') === -1) {
        
        var chain = getFullTransferChainForPurchase(purchaseId, component);
        if (chain.length > 1) {
          // Находим предыдущее перемещение в цепочке
          var currentTransferDate = data[i][2];
          var sourceTransfer = null;
          
          for (var j = 0; j < chain.length; j++) {
            if (new Date(chain[j].date) < new Date(currentTransferDate) && 
                (!sourceTransfer || new Date(chain[j].date) > new Date(sourceTransfer.date))) {
              sourceTransfer = chain[j];
            }
          }
          
          if (sourceTransfer) {
            // Добавляем примечание о связи
            var currentNotes = data[i][8] || '';
            var newNote = 'Источник: ' + sourceTransfer.transferId;
            
            if (!currentNotes.includes(newNote)) {
              var updatedNotes = currentNotes ? currentNotes + '; ' + newNote : newNote;
              transfersSheet.getRange(row, 9).setValue(updatedNotes);
              updatedCount++;
              console.log('✅ Связь: ' + sourceTransfer.transferId + ' → ' + data[i][0]);
            }
          }
        }
      }
    }
    
    console.log('🔗 Восстановлено связей: ' + updatedCount);
    return updatedCount;
    
  } catch (error) {
    console.error('❌ Ошибка восстановления связей: ' + error.toString());
    return 0;
  }
}

/**
 * 📊 ГЕНЕРАЦИЯ ОТЧЕТА ПО ЦЕПОЧКЕ ПЕРЕМЕЩЕНИЙ
 */
function generateTransferChainReport(purchaseId, component) {
  try {
    var chain = getFullTransferChainForPurchase(purchaseId, component);
    if (chain.length === 0) {
      return '❌ Цепочка перемещений не найдена для ' + purchaseId + ' - ' + component;
    }
    
    var report = '📊 ОТЧЕТ ПО ЦЕПОЧКЕ ПЕРЕМЕЩЕНИЙ\n\n';
    report += '🏷️ ID закупки: ' + purchaseId + '\n';
    report += '🔧 Комплектующее: ' + component + '\n';
    report += '📈 Количество этапов: ' + chain.length + '\n\n';
    
    report += '🔗 ЦЕПОЧКА ПЕРЕМЕЩЕНИЙ:\n';
    var totalMoved = 0;
    
    for (var i = 0; i < chain.length; i++) {
      var transfer = chain[i];
      report += (i + 1) + '. ' + transfer.transferId + '\n';
      report += '   📅 ' + Utilities.formatDate(transfer.date, Session.getScriptTimeZone(), "dd.MM.yyyy HH:mm") + '\n';
      report += '   🔄 ' + transfer.from + ' → ' + transfer.to + '\n';
      report += '   📦 ' + transfer.quantity + ' ед.\n';
      
      if (transfer.notes) {
        report += '   💬 ' + transfer.notes + '\n';
      }
      report += '\n';
      
      totalMoved += transfer.quantity;
    }
    
    report += '📊 ИТОГО перемещено: ' + totalMoved + ' ед.\n';
    
    // Проверяем конечный остаток
    var finalWarehouse = chain[chain.length - 1].to;
    var finalTransferId = chain[chain.length - 1].transferId;
    var finalQuantity = getCurrentStockByTransferId(finalTransferId, finalWarehouse);
    
    report += '📦 Текущий остаток: ' + (finalQuantity !== null ? finalQuantity + ' ед.' : 'не найден') + '\n';
    report += '🏢 Текущее местоположение: ' + finalWarehouse + '\n';
    
    return report;
    
  } catch (error) {
    console.error('❌ Ошибка генерации отчета: ' + error.toString());
    return '❌ Ошибка генерации отчета: ' + error.toString();
  }
}

/**
 * 🔍 ПОЛУЧЕНИЕ ТЕКУЩЕГО ОСТАТКА ПО ID ПЕРЕМЕЩЕНИЯ
 */
function getCurrentStockByTransferId(transferId, warehouseName) {
  try {
    var sheetName = warehouseName === 'Производство' ? 'Склад_Производства' : 'Склад_Главный';
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    
    if (!sheet || sheet.getLastRow() <= 1) {
      return null;
    }
    
    var data = sheet.getRange('A2:F' + sheet.getLastRow()).getValues();
    for (var i = 0; i < data.length; i++) {
      if (data[i][0] === transferId) {
        return data[i][4]; // Остаток
      }
    }
    
    return null;
    
  } catch (error) {
    console.error('❌ Ошибка поиска остатка: ' + error.toString());
    return null;
  }
}

/**
 * 🔍 ДИАЛОГ ПРОВЕРКИ ЦЕПОЧКИ ПЕРЕМЕЩЕНИЙ
 */
function showTransferChainDialog() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('TransferChainDialog')
      .setWidth(500)
      .setHeight(400);
    
    SpreadsheetApp.getUi().showModalDialog(html, '🔍 Проверка цепочки перемещений');
    
  } catch (error) {
    console.error('❌ Ошибка показа диалога цепочки: ' + error.toString());
    SpreadsheetApp.getUi().alert('❌ Ошибка', 'Не удалось открыть диалог проверки цепочки');
  }
}

/**
 * 📊 ДИАЛОГ ГЕНЕРАЦИИ ОТЧЕТА ПО ЦЕПОЧКЕ
 */
function showTransferReportDialog() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('TransferReportDialog')
      .setWidth(600)
      .setHeight(500);
    
    SpreadsheetApp.getUi().showModalDialog(html, '📊 Отчет по цепочке перемещений');
    
  } catch (error) {
    console.error('❌ Ошибка показа диалога отчета: ' + error.toString());
    SpreadsheetApp.getUi().alert('❌ Ошибка', 'Не удалось открыть диалог отчета');
  }
}

/**
 * 🔍 ПОЛУЧЕНИЕ ДАННЫХ ДЛЯ ЦЕПОЧКИ ПЕРЕМЕЩЕНИЙ (для диалога)
 */
function getTransferChainData(purchaseId, component) {
  try {
    if (!purchaseId || !component) {
      return { error: 'Укажите ID закупки и компонент' };
    }
    
    var chain = getFullTransferChainForPurchase(purchaseId, component);
    if (chain.length === 0) {
      return { error: 'Цепочка перемещений не найдена' };
    }
    
    // Форматируем данные для отображения
    var formattedChain = chain.map(function(transfer, index) {
      return {
        number: index + 1,
        transferId: transfer.transferId,
        date: Utilities.formatDate(transfer.date, Session.getScriptTimeZone(), "dd.MM.yyyy HH:mm"),
        from: transfer.from,
        to: transfer.to,
        quantity: transfer.quantity,
        notes: transfer.notes || ''
      };
    });
    
    // Получаем текущий остаток
    var finalTransfer = chain[chain.length - 1];
    var currentStock = getCurrentStockByTransferId(finalTransfer.transferId, finalTransfer.to);
    
    return {
      success: true,
      purchaseId: purchaseId,
      component: component,
      chain: formattedChain,
      currentStock: currentStock,
      currentWarehouse: finalTransfer.to,
      totalTransferred: chain.reduce((sum, t) => sum + t.quantity, 0)
    };
    
  } catch (error) {
    console.error('❌ Ошибка получения данных цепочки: ' + error.toString());
    return { error: 'Ошибка получения данных: ' + error.toString() };
  }
}

/**
 * 🔥 НОВАЯ ФУНКЦИЯ: СОРТИРОВКА ПЕРЕМЕЩЕНИЙ В ХРОНОЛОГИЧЕСКОМ ПОРЯДКЕ
 */
function getChronologicallySortedTransfers(transferMaps) {
  try {
    console.log('📅 Сортировка перемещений в хронологическом порядке...');
    
    var allTransfers = [];
    
    // Собираем все перемещения из purchaseMap
    for (var purchaseId in transferMaps.purchaseMap) {
      var transfers = transferMaps.purchaseMap[purchaseId];
      allTransfers = allTransfers.concat(transfers.map(function(t) {
        return {
          transferId: t.transferId,
          purchaseId: purchaseId,
          component: t.component,
          quantity: t.quantity,
          fromWarehouse: 'Поставщик',
          toWarehouse: t.type === 'initial' ? 'Главный' : 'Производство',
          transferDate: t.date,
          sourceTransferId: null,
          type: 'purchase'
        };
      }));
    }
    
    // Собираем все перемещения из chainMap
    for (var sourceId in transferMaps.chainMap) {
      var chainTransfers = transferMaps.chainMap[sourceId];
      for (var i = 0; i < chainTransfers.length; i++) {
        var chainTransfer = chainTransfers[i];
        allTransfers.push({
          transferId: chainTransfer.newTransferId,
          purchaseId: '', // Будет заполнено позже
          component: chainTransfer.component,
          quantity: chainTransfer.quantity,
          fromWarehouse: chainTransfer.from,
          toWarehouse: chainTransfer.to,
          transferDate: chainTransfer.date,
          sourceTransferId: sourceId,
          type: 'transfer'
        });
      }
    }
    
    // Сортируем по дате (от старых к новым)
    allTransfers.sort(function(a, b) {
      var dateA = a.transferDate ? new Date(a.transferDate) : new Date(0);
      var dateB = b.transferDate ? new Date(b.transferDate) : new Date(0);
      return dateA - dateB;
    });
    
    console.log('✅ Отсортировано перемещений: ' + allTransfers.length);
    return allTransfers;
    
  } catch (error) {
    console.error('❌ Ошибка сортировки перемещений: ' + error.toString());
    return [];
  }
}

/**
 * Возвращает хронологически отсортированные перемещения для конкретного компонента
 * @param {string} componentName - Название компонента
 * @param {Array} allTransfers - Все перемещения (если не передано, берется из листа)
 * @returns {Array} Отсортированные по дате перемещения
 */
function getChronologicallySortedTransfers(componentName, allTransfers) {
  try {
    // Если не переданы все перемещения, загружаем их
    if (!allTransfers || allTransfers.length === 0) {
      var transfersSheet = SpreadsheetApp.getActiveSpreadsheet()
        .getSheetByName(TRANSFERS_SHEET);
      var lastRow = transfersSheet.getLastRow();
      
      if (lastRow <= 1) return []; // Нет данных
      
      allTransfers = transfersSheet.getRange(2, 1, lastRow - 1, 8)
        .getValues()
        .filter(row => row[0]); // Фильтруем пустые строки
    }
    
    // Фильтруем по компоненту и сортируем по дате
    var componentTransfers = allTransfers.filter(function(row) {
      return row[3] === componentName; // COMPONENT в колонке D (индекс 3)
    });
    
    // Сортировка по дате (колонка C - индекс 2)
    componentTransfers.sort(function(a, b) {
      return new Date(a[2]) - new Date(b[2]);
    });
    
    return componentTransfers;
    
  } catch (error) {
    logToSheet('ERROR', 'getChronologicallySortedTransfers', 
      'Ошибка при сортировке перемещений: ' + error.message);
    return [];
  }
}

/**
 * Валидация данных перед пересчетом
 */
function validateDataBeforeRecalculation() {
  var issues = [];
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Проверяем существование листов
  var requiredSheets = [
    MAIN_WAREHOUSE_SHEET,
    TRANSFERS_SHEET,
    PURCHASES_SHEET
  ];
  
  requiredSheets.forEach(function(sheetName) {
    var sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      issues.push('Отсутствует лист: ' + sheetName);
    } else if (sheet.getLastRow() <= 1) {
      issues.push('Лист пустой: ' + sheetName);
    }
  });
  
  return {
    isValid: issues.length === 0,
    issues: issues
  };
}

/**
 * Создает резервную копию перед пересчетом
 */
function createBackupBeforeRecalculation() {
  try {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var backupName = spreadsheet.getName() + ' - Резервная копия ' + 
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH-mm');
    
    // Создаем копию файла
    var backupFile = DriveApp.getFileById(spreadsheet.getId()).makeCopy(backupName);
    
    logToSheet('INFO', 'createBackupBeforeRecalculation',
      'Создана резервная копия: ' + backupFile.getUrl());
    
    return {
      success: true,
      url: backupFile.getUrl(),
      name: backupName
    };
    
  } catch (error) {
    logToSheet('WARNING', 'createBackupBeforeRecalculation',
      'Не удалось создать резервную копию: ' + error.message);
    return { success: false, error: error.message };
  }
}

/**
 * Логирование событий в специальный лист
 * @param {string} level - Уровень (INFO, WARNING, ERROR)
 * @param {string} functionName - Название функции
 * @param {string} message - Сообщение
 * @param {Object} extraData - Дополнительные данные (опционально)
 */
function logToSheet(level, functionName, message, extraData) {
  try {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var logSheet = spreadsheet.getSheetByName('Системные_Логи');
    
    // Если листа нет - создаем
    if (!logSheet) {
      logSheet = spreadsheet.insertSheet('Системные_Логи');
      logSheet.getRange(1, 1, 1, 5).setValues([[
        'Дата/Время', 'Уровень', 'Функция', 'Сообщение', 'Доп. данные'
      ]]);
      logSheet.getRange(1, 1, 1, 5).setFontWeight('bold');
      logSheet.setFrozenRows(1);
    }
    
    // Подготавливаем данные для записи
    var timestamp = Utilities.formatDate(
      new Date(), 
      Session.getScriptTimeZone(), 
      'dd.MM.yyyy HH:mm:ss'
    );
    
    var extraDataStr = '';
    if (extraData) {
      try {
        extraDataStr = JSON.stringify(extraData);
      } catch (e) {
        extraDataStr = String(extraData);
      }
    }
    
    // Добавляем запись
    var lastRow = logSheet.getLastRow();
    logSheet.getRange(lastRow + 1, 1, 1, 5).setValues([[
      timestamp,
      level,
      functionName,
      message,
      extraDataStr
    ]]);
    
    // Автоматическое форматирование строк в зависимости от уровня
    var range = logSheet.getRange(lastRow + 1, 1, 1, 5);
    
    switch(level) {
      case 'ERROR':
        range.setBackground('#ffebee'); // Красный фон для ошибок
        range.setFontColor('#d32f2f');
        break;
      case 'WARNING':
        range.setBackground('#fff3e0'); // Оранжевый фон для предупреждений
        range.setFontColor('#f57c00');
        break;
      case 'INFO':
        range.setBackground('#e8f5e9'); // Зеленый фон для информации
        break;
    }
    
    // Ограничиваем количество записей (макс. 1000 строк)
    if (lastRow > 1000) {
      logSheet.deleteRow(2); // Удаляем самую старую запись
    }
    
    // Также выводим в консоль для отладки
    console.log('[' + level + '] ' + functionName + ': ' + message);
    
  } catch (error) {
    // Если логирование сломалось, пишем хотя бы в консоль
    console.error('Ошибка в logToSheet:', error);
    console.log('Исходное сообщение: [' + level + '] ' + functionName + ': ' + message);
  }
}
