function doGet(e) {
  const action = e.parameter.action;
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Fetch Ingredients and Packaging
  if (action === 'getIngredients') {
    const data = [];
    ['Ingredients', 'Packaging'].forEach(function(sheetName) {
      var sheet = ss.getSheetByName(sheetName);
      if (sheet && sheet.getLastRow() > 0) {
        var values = sheet.getRange(1, 1, sheet.getLastRow(), 4).getValues();
        values.forEach(function(row) {
          if (row[0]) {
            data.push({
              name: row[0],
              suppliers: row[1] || '',
              storageType: row[2] || (sheetName === 'Packaging' ? 'Packaging' : 'Ambient'),
              icon: row[3] || ''
            });
          }
        });
      }
    });
    return ContentService.createTextOutput(JSON.stringify({ingredients: data})).setMimeType(ContentService.MimeType.JSON);
  }

  // 2. Fetch Master Supplier List
  if (action === 'getSuppliers') {
    var sheet = ss.getSheetByName('Suppliers');
    if (!sheet) sheet = ss.insertSheet('Suppliers');
    var data = sheet.getLastRow() > 0 ? sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues().flat().filter(function(v) { return v; }) : [];
    return ContentService.createTextOutput(JSON.stringify({suppliers: data})).setMimeType(ContentService.MimeType.JSON);
  }

  // 3. Fetch all delivery records
  if (action === 'getDeliveries') {
    var records = [];
    var tz = Session.getScriptTimeZone();

    function fmtDate(v) {
      if (!v) return '';
      if (v instanceof Date) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
      return String(v);
    }

    var delSheet = ss.getSheetByName('Deliveries');
    if (delSheet && delSheet.getLastRow() > 0) {
      delSheet.getRange(1, 1, delSheet.getLastRow(), 12).getValues().forEach(function(r) {
        if (!r[0]) return;
        records.push({
          date: fmtDate(r[0]),
          invoice: String(r[1]),
          supplier: String(r[2]),
          staff: String(r[3]),
          ingredient: String(r[4]),
          batch: String(r[5]),
          ubd: fmtDate(r[6]),
          qty: String(r[7]),
          unit: String(r[8]),
          condition: String(r[9]),
          storage: String(r[10]),
          isPackaging: false,
          savedAt: String(r[11] || '')
        });
      });
    }

    var pkgSheet = ss.getSheetByName('Packaging Deliveries');
    if (pkgSheet && pkgSheet.getLastRow() > 0) {
      pkgSheet.getRange(1, 1, pkgSheet.getLastRow(), 11).getValues().forEach(function(r) {
        if (!r[0]) return;
        records.push({
          date: fmtDate(r[0]),
          invoice: String(r[1]),
          supplier: String(r[2]),
          staff: String(r[3]),
          ingredient: String(r[4]),
          batch: String(r[5]),
          ubd: '',
          qty: String(r[6]),
          unit: String(r[7]),
          condition: String(r[8]),
          storage: String(r[9]),
          isPackaging: true,
          savedAt: String(r[10] || '')
        });
      });
    }

    return ContentService.createTextOutput(JSON.stringify({result: 'success', records: records})).setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService.createTextOutput(JSON.stringify({status: 'ok'})).setMimeType(ContentService.MimeType.JSON);
}


function doPost(e) {
  const data = JSON.parse(e.postData.contents);
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Action: Add or Update an Ingredient/Packaging item
  if (data.action === 'saveIngredientObj') {
    const ingName = data.ingredient.name;
    const oldName = data.ingredient.oldName;
    const ingSups = data.ingredient.suppliers || '';
    const ingStorage = data.ingredient.storageType || 'Ambient';
    const ingIcon = data.ingredient.icon || '';

    const searchName = oldName ? oldName : ingName;

    ['Ingredients', 'Packaging'].forEach(function(sheetName) {
      var s = ss.getSheetByName(sheetName);
      if (s && s.getLastRow() > 0) {
        var values = s.getRange(1, 1, s.getLastRow(), 1).getValues().flat();
        for (var i = values.length - 1; i >= 0; i--) {
          if (values[i] === searchName) s.deleteRow(i + 1);
        }
      }
    });

    const targetSheetName = ingStorage === 'Packaging' ? 'Packaging' : 'Ingredients';
    var targetSheet = ss.getSheetByName(targetSheetName);
    if (!targetSheet) targetSheet = ss.insertSheet(targetSheetName);
    targetSheet.appendRow([ingName, ingSups, ingStorage, ingIcon]);
    return ContentService.createTextOutput(JSON.stringify({result: 'success'})).setMimeType(ContentService.MimeType.JSON);
  }

  // Action: Delete an item
  if (data.action === 'deleteIngredient') {
    ['Ingredients', 'Packaging'].forEach(function(sheetName) {
      var s = ss.getSheetByName(sheetName);
      if (s && s.getLastRow() > 0) {
        var values = s.getRange(1, 1, s.getLastRow(), 1).getValues().flat();
        for (var i = values.length - 1; i >= 0; i--) {
          if (values[i] === data.ingredient) s.deleteRow(i + 1);
        }
      }
    });
    return ContentService.createTextOutput(JSON.stringify({result: 'success'})).setMimeType(ContentService.MimeType.JSON);
  }

  // Action: Save a Supplier to Master List
  if (data.action === 'saveSupplier') {
    var sheet = ss.getSheetByName('Suppliers');
    if (!sheet) sheet = ss.insertSheet('Suppliers');
    var found = false;
    if (sheet.getLastRow() > 0) {
      var existing = sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues().flat();
      if (existing.includes(data.supplier)) found = true;
    }
    if (!found) sheet.appendRow([data.supplier]);
    return ContentService.createTextOutput(JSON.stringify({result: 'success'})).setMimeType(ContentService.MimeType.JSON);
  }

  // Action: Delete a Supplier from Master List
  if (data.action === 'deleteSupplier') {
    var sheet = ss.getSheetByName('Suppliers');
    if (!sheet) return ContentService.createTextOutput(JSON.stringify({result: 'success'})).setMimeType(ContentService.MimeType.JSON);
    if (sheet.getLastRow() > 0) {
      var values = sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues().flat();
      for (var i = values.length - 1; i >= 0; i--) {
        if (values[i] === data.supplier) sheet.deleteRow(i + 1);
      }
    }
    return ContentService.createTextOutput(JSON.stringify({result: 'success'})).setMimeType(ContentService.MimeType.JSON);
  }

  // Debug: Write raw received data to a Debug sheet
  var debugSheet = ss.getSheetByName('Debug');
  if (!debugSheet) debugSheet = ss.insertSheet('Debug');
  debugSheet.appendRow([new Date().toISOString(), JSON.stringify(data)]);

  // Save Delivery records
  if (data.rows && Array.isArray(data.rows)) {
    data.rows.forEach(function(r) {
      const isPkg = r.isPackaging === true;
      const targetSheetName = isPkg ? 'Packaging Deliveries' : 'Deliveries';

      var sheet = ss.getSheetByName(targetSheetName);
      if (!sheet) sheet = ss.insertSheet(targetSheetName);

      if (isPkg) {
        sheet.appendRow([
          r.date || '', r.invoice || '', r.supplier || '', r.staff || '',
          r.ingredient || '', r.batch || '', r.qty || '', r.unit || '',
          r.condition || '', r.storage || '', r.savedAt || ''
        ]);
      } else {
        sheet.appendRow([
          r.date || '', r.invoice || '', r.supplier || '', r.staff || '',
          r.ingredient || '', r.batch || '', r.ubd || '', r.qty || '',
          r.unit || '', r.condition || '', r.storage || '', r.savedAt || ''
        ]);
      }
    });
  }

  return ContentService.createTextOutput(JSON.stringify({result: 'success'})).setMimeType(ContentService.MimeType.JSON);
}
