/**
 * Merchant Product Type Export
 *
 * Google Ads Script для експорту product_type з Merchant Center
 * після застосування правил перетворення.
 *
 * Листи:
 * - Products: експорт ID + Product Type для перенесення між фідами.
 * - Settings: налаштування запуску.
 *
 * Потрібні Advanced APIs у Google Ads Scripts:
 * - Merchant API -> Products
 * - Merchant API -> Accounts, якщо auto_register_gcp_project = TRUE
 */


var SPREADSHEET_URL = 'PASTE_SPREADSHEET_URL_HERE';


var PRODUCTS_SHEET = 'Products';
var SETTINGS_SHEET = 'Settings';
var OLD_GENERATED_SHEETS = ['ProductTypeExport', 'ProductTypes', 'Brands'];


var START_TIME = new Date().getTime();


function main() {
  var ctx = null;
  try {
    if (!SPREADSHEET_URL || SPREADSHEET_URL === 'PASTE_SPREADSHEET_URL_HERE') {
      throw new Error('Вкажіть SPREADSHEET_URL на початку скрипта.');
    }


    var ss = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
    var sheets = ensureSheets_(ss);
    var settings = readSettings_(sheets.settings);
    ctx = { ss: ss, sheets: sheets, settings: settings, runId: runId_() };


    requireMerchantId_(settings);
    ensureMerchantApiServicesOn_(settings);
    if (!ensureMerchantApiDeveloperRegistration_(settings)) {
      Logger.log('GCP project зареєстровано в Merchant API. Запустіть main() ще раз після оновлення доступу.');
      return;
    }


    var exported = runProductTypeExport_(ctx);
    Logger.log('Product Type Export завершено. Рядків=' + exported + '.');
  } catch (e) {
    throw e;
  }
}


function resetAudit() {
  if (!SPREADSHEET_URL || SPREADSHEET_URL === 'PASTE_SPREADSHEET_URL_HERE') {
    throw new Error('Вкажіть SPREADSHEET_URL на початку скрипта.');
  }
  var ss = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  var sheets = ensureSheets_(ss);
  clearBelowHeader_(sheets.products);
}


function ensureSheets_(ss) {
  cleanupOldGeneratedSheets_(ss);
  var out = {
    products: ss.getSheetByName(PRODUCTS_SHEET) || ss.insertSheet(PRODUCTS_SHEET, 0),
    settings: ss.getSheetByName(SETTINGS_SHEET) || ss.insertSheet(SETTINGS_SHEET)
  };
  ss.setActiveSheet(out.products);
  ss.moveActiveSheet(1);
  ensureHeader_(out.products, productExportHeader_());
  ensureSettingsTemplate_(out.settings);
  return out;
}


function cleanupOldGeneratedSheets_(ss) {
  for (var i = 0; i < OLD_GENERATED_SHEETS.length; i++) {
    var sheet = ss.getSheetByName(OLD_GENERATED_SHEETS[i]);
    if (sheet && ss.getSheets().length > 1) ss.deleteSheet(sheet);
  }
}


function ensureSettingsTemplate_(sheet) {
  var existing = readSettingsMap_(sheet);
  var rows = [
    ['setting', 'value', 'description'],
    ['merchant_id', existing.merchant_id || '', 'ID акаунта Merchant Center.'],
    ['developer_email', existing.developer_email || 'bazhinalex05@gmail.com', 'Email розробника для реєстрації GCP project у Merchant API.'],
    ['auto_register_gcp_project', settingOr_(existing.auto_register_gcp_project, 'TRUE'), 'TRUE автоматично реєструє GCP project Google Ads Scripts у Merchant API.'],
    ['max_run_minutes', settingOr_(existing.max_run_minutes, '25'), 'М’який ліміт часу запуску в хвилинах.'],
    ['merchant_api_page_size', settingOr_(existing.merchant_api_page_size, '1000'), 'Кількість товарів в одному запиті до Merchant API.'],
    ['merchant_api_retry_count', settingOr_(existing.merchant_api_retry_count, '5'), 'Кількість повторних спроб після помилок Merchant API.'],
    ['merchant_api_retry_sleep_seconds', settingOr_(existing.merchant_api_retry_sleep_seconds, '10'), 'Базова пауза між повторними спробами в секундах.'],
    ['max_merchant_pages_per_run', settingOr_(existing.max_merchant_pages_per_run, '25'), 'Максимальна кількість сторінок Merchant API за один запуск.'],
    ['product_row_flush_size', settingOr_(existing.product_row_flush_size, '1000'), 'Кількість рядків, які записуються в Products одним пакетом.'],
    ['merchant_data_source_id_filter', existing.merchant_data_source_id_filter || '', 'Необов’язковий фільтр data source ID через кому. Порожньо = всі.'],
    ['merchant_feed_label_filter', existing.merchant_feed_label_filter || '', 'Необов’язковий фільтр feed label через кому. Порожньо = всі.'],
    ['merchant_content_language_filter', existing.merchant_content_language_filter || '', 'Необов’язковий фільтр мов контенту через кому. Порожньо = всі.']
  ];
  sheet.clear();
  sheet.getRange(1, 1, rows.length, 3).setValues(rows);
  sheet.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#c9daf8');
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 260);
  sheet.setColumnWidth(2, 220);
  sheet.setColumnWidth(3, 620);
  var boolRows = { auto_register_gcp_project: true };
  for (var i = 2; i <= rows.length; i++) {
    var key = String(sheet.getRange(i, 1).getValue() || '');
    if (boolRows[key]) sheet.getRange(i, 2).insertCheckboxes();
  }
}


function readSettings_(sheet) {
  var raw = readSettingsMap_(sheet);
  return {
    merchantId: String(raw.merchant_id || '').trim(),
    developerEmail: String(raw.developer_email || '').trim(),
    autoRegisterGcpProject: bool_(raw.auto_register_gcp_project, true),
    maxRunMinutes: num_(raw.max_run_minutes, 25),
    merchantPageSize: num_(raw.merchant_api_page_size, 1000),
    merchantApiRetryCount: num_(raw.merchant_api_retry_count, 5),
    merchantApiRetrySleepSeconds: num_(raw.merchant_api_retry_sleep_seconds, 10),
    maxMerchantPagesPerRun: num_(raw.max_merchant_pages_per_run, 25),
    productRowFlushSize: num_(raw.product_row_flush_size, 1000),
    merchantDataSourceIdFilter: listSetting_(raw.merchant_data_source_id_filter),
    merchantFeedLabelFilter: listSetting_(raw.merchant_feed_label_filter),
    merchantContentLanguageFilter: listSetting_(raw.merchant_content_language_filter)
  };
}


function runProductTypeExport_(ctx) {
  var token = '';
  var scanned = 0;
  var filteredOut = 0;
  var rowsById = {};
  var idOrder = [];


  var pages = 0;
  while (pages < ctx.settings.maxMerchantPagesPerRun && !shouldStopSoon_(ctx.settings)) {
    var page = fetchMerchantProductsPage_(ctx.settings.merchantId, token, ctx.settings.merchantPageSize, ctx.settings);
    for (var i = 0; i < page.items.length; i++) {
      scanned++;
      if (!merchantProductMatchesFilters_(page.items[i], ctx.settings)) {
        filteredOut++;
        continue;
      }
      mergeProductExportRow_(rowsById, idOrder, productExportRow_(page.items[i]));
    }
    token = page.nextPageToken || '';
    pages++;
    if (!token) break;
  }


  if (token) {
    throw new Error('У Merchant більше товарів, ніж скрипт обробив за один запуск. Збільшіть max_merchant_pages_per_run і запустіть ще раз. Проскановано=' + scanned + ', відфільтровано=' + filteredOut + '.');
  }
  var rows = productExportRows_(rowsById, idOrder);
  clearBelowHeader_(ctx.sheets.products);
  var flushSize = Math.max(1, ctx.settings.productRowFlushSize || 1000);
  for (var j = 0; j < rows.length; j += flushSize) {
    var batch = rows.slice(j, j + flushSize);
    ctx.sheets.products.getRange(j + 2, 1, batch.length, 2).setValues(batch);
  }
  SpreadsheetApp.flush();
  return rows.length;
}


function productExportHeader_() {
  return ['ID', 'Product Type'];
}


function productExportRow_(p) {
  var allProductTypes = allProductTypes_(p);
  var fullPath = normalizeProductType_(allProductTypes.length ? allProductTypes[0] : '');
  return [
    normOfferId_(p.offerId || p.id || ''),
    fullPath
  ];
}


function mergeProductExportRow_(rowsById, idOrder, row) {
  var id = String(row[0] || '').trim();
  if (!id) return;
  var productType = normalizeProductType_(row[1]);
  if (!rowsById[id]) {
    rowsById[id] = [id, productType];
    idOrder.push(id);
    return;
  }
  var existingProductType = normalizeProductType_(rowsById[id][1]);
  if (!existingProductType || productType.length > existingProductType.length) {
    rowsById[id] = [id, productType];
  }
}


function productExportRows_(rowsById, idOrder) {
  var rows = [];
  for (var i = 0; i < idOrder.length; i++) {
    var row = rowsById[idOrder[i]];
    if (row) rows.push(row);
  }
  return rows;
}


function ensureHeader_(sheet, header) {
  var finalHeader = header.slice();
  enforceSheetColumnCount_(sheet, finalHeader.length);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, finalHeader.length).setValues([finalHeader]);
  } else {
    sheet.getRange(1, 1, 1, finalHeader.length).setValues([finalHeader]);
  }
  sheet.getRange(1, 1, 1, finalHeader.length).setFontWeight('bold').setBackground('#c9daf8');
  sheet.setFrozenRows(1);
}


function fetchMerchantProductsPage_(merchantId, pageToken, pageSize, settings) {
  var service = getMerchantProductsService_();
  var parent = 'accounts/' + String(merchantId);
  var response = listMerchantProductsPageWithRetry_(service, parent, pageToken, pageSize, settings);
  return {
    items: response && response.products ? response.products : [],
    nextPageToken: response && response.nextPageToken ? response.nextPageToken : ''
  };
}


function listMerchantProductsPageWithRetry_(service, parent, pageToken, pageSize, settings) {
  var attempts = Math.max(1, Number(settings.merchantApiRetryCount || 0) + 1);
  var baseSleepMs = Math.max(1, Number(settings.merchantApiRetrySleepSeconds || 10)) * 1000;
  var lastMessage = '';
  for (var attempt = 1; attempt <= attempts; attempt++) {
    try {
      return service.Accounts.Products.list(parent, {
        pageToken: pageToken || undefined,
        pageSize: pageSize
      });
    } catch (e) {
      lastMessage = e && e.message ? e.message : String(e);
      if (lastMessage.indexOf('not registered with the merchant account') !== -1) {
        throw new Error('Merchant API не зареєстровано для цього Merchant ID. Увімкніть Merchant API -> Accounts або залиште auto_register_gcp_project=TRUE. Деталі: ' + lastMessage);
      }
      if (attempt >= attempts) break;
      Utilities.sleep(baseSleepMs * attempt);
    }
  }
  throw new Error('Merchant API Products.list не виконався після повторних спроб: ' + lastMessage);
}


function ensureMerchantApiServicesOn_(settings) {
  var productsService = getMerchantProductsService_();
  if (!productsService ||
      !productsService.Accounts ||
      !productsService.Accounts.Products ||
      typeof productsService.Accounts.Products.list !== 'function') {
    throw new Error('Увімкніть Advanced API у Google Ads Scripts: Merchant API -> Products.');
  }
  if (!settings.autoRegisterGcpProject) return;
  var accountsService = getMerchantAccountsService_();
  if (!accountsService ||
      !accountsService.Accounts ||
      !accountsService.Accounts.DeveloperRegistration ||
      typeof accountsService.Accounts.DeveloperRegistration.registerGcp !== 'function') {
    throw new Error('Увімкніть Advanced API у Google Ads Scripts: Merchant API -> Accounts або встановіть auto_register_gcp_project=FALSE після ручної реєстрації.');
  }
}


function ensureMerchantApiDeveloperRegistration_(settings) {
  if (!settings.autoRegisterGcpProject) return true;
  if (!settings.developerEmail) {
    throw new Error('Settings.developer_email обов’язковий, якщо auto_register_gcp_project=TRUE.');
  }
  var accountsService = getMerchantAccountsService_();
  var merchantId = String(settings.merchantId);
  try {
    accountsService.Accounts.DeveloperRegistration.getDeveloperRegistration(
      'accounts/' + merchantId + '/developerRegistration'
    );
    return true;
  } catch (e) {
    Logger.log('Реєстрацію розробника ще не підтверджено: ' + String(e));
  }
  accountsService.Accounts.DeveloperRegistration.registerGcp(
    { developerEmail: settings.developerEmail },
    'accounts/' + merchantId + '/developerRegistration'
  );
  return false;
}


function getMerchantProductsService_() {
  if (typeof MerchantApiProducts !== 'undefined') return MerchantApiProducts;
  if (typeof MerchantProducts !== 'undefined') return MerchantProducts;
  return null;
}


function getMerchantAccountsService_() {
  if (typeof MerchantApiAccounts !== 'undefined') return MerchantApiAccounts;
  if (typeof MerchantAccounts !== 'undefined') return MerchantAccounts;
  return null;
}


function merchantProductMatchesFilters_(p, settings) {
  var a = productAttributes_(p);
  var dataSource = String(p.dataSource || p.source || a.dataSource || '');
  var dataSourceId = extractDataSourceId_(dataSource);
  var feedLabel = String(p.feedLabel || a.feedLabel || '');
  var contentLanguage = String(p.contentLanguage || a.contentLanguage || '');
  if (!matchesAnyFilter_(dataSourceId, settings.merchantDataSourceIdFilter) &&
      !matchesAnyFilter_(dataSource, settings.merchantDataSourceIdFilter)) return false;
  if (!matchesAnyFilter_(feedLabel, settings.merchantFeedLabelFilter)) return false;
  if (!matchesAnyFilter_(contentLanguage, settings.merchantContentLanguageFilter)) return false;
  return true;
}


function readSettingsMap_(sheet) {
  var out = {};
  if (!sheet || sheet.getLastRow() < 2) return out;
  var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  for (var i = 0; i < values.length; i++) {
    var key = String(values[i][0] || '').trim();
    if (key) out[key] = values[i][1];
  }
  return out;
}


function appendRows_(sheet, rows) {
  if (!rows || !rows.length) return;
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
}


function clearBelowHeader_(sheet) {
  var last = sheet.getLastRow();
  if (last > 1) sheet.getRange(2, 1, last - 1, Math.max(1, sheet.getLastColumn())).clearContent();
}


function enforceSheetColumnCount_(sheet, columnCount) {
  var current = sheet.getMaxColumns();
  if (current < columnCount) {
    sheet.insertColumnsAfter(current, columnCount - current);
  } else if (current > columnCount) {
    sheet.deleteColumns(columnCount + 1, current - columnCount);
  }
}


function applyProductLinkColumns_(sheet) {
  applyRichTextLinks_(sheet, 5);
  applyRichTextLinks_(sheet, 14);
}


function applyRichTextLinks_(sheet, column) {
  var last = sheet.getLastRow();
  if (last <= 1) return;
  var range = sheet.getRange(2, column, last - 1, 1);
  var values = range.getValues();
  var richValues = [];
  for (var i = 0; i < values.length; i++) {
    var url = String(values[i][0] || '').trim();
    if (url) {
      richValues.push([SpreadsheetApp.newRichTextValue().setText(url).setLinkUrl(url).build()]);
    } else {
      richValues.push([SpreadsheetApp.newRichTextValue().setText('').build()]);
    }
  }
  range.setRichTextValues(richValues);
}


function requireMerchantId_(settings) {
  if (!settings.merchantId) throw new Error('Settings.merchant_id обов’язковий.');
}


function productAttributes_(p) {
  return (p && (p.productAttributes || p.attributes)) || {};
}


function allProductTypes_(p) {
  var out = [];
  var a = productAttributes_(p);
  if (a.productTypes && a.productTypes.length) out = a.productTypes;
  else if (p.productTypes && p.productTypes.length) out = p.productTypes;
  else if (a.productType) out = [a.productType];
  else if (p.productType) out = [p.productType];
  else {
    var customProductType = customAttributeValue_(p, 'product_type') || customAttributeValue_(p, 'product type');
    if (customProductType) out = [customProductType];
  }
  var clean = [];
  for (var i = 0; i < out.length; i++) {
    var v = normalizeProductType_(out[i]);
    if (v) clean.push(v);
  }
  return clean;
}


function customLabelValue_(p, index) {
  var a = productAttributes_(p);
  var camel = 'customLabel' + index;
  var snake = 'custom_label_' + index;
  return value_(a[camel], p[camel], customAttributeValue_(p, camel), customAttributeValue_(p, snake));
}


function customAttributeValue_(p, name) {
  var attrs = p && p.customAttributes ? p.customAttributes : [];
  var target = String(name || '').toLowerCase();
  for (var i = 0; i < attrs.length; i++) {
    var attrName = String(attrs[i].name || '').toLowerCase();
    if (attrName === target) return attrs[i].value || attrs[i].textValue || '';
  }
  return '';
}


function customAttributesBrief_(p) {
  var attrs = p && p.customAttributes ? p.customAttributes : [];
  var out = [];
  for (var i = 0; i < attrs.length; i++) {
    var name = String(attrs[i].name || '').trim();
    var value = value_(attrs[i].value, attrs[i].textValue, attrs[i].intValue, attrs[i].floatValue);
    if (name || value) out.push(name + '=' + value);
  }
  return out.join(' | ');
}


function productStatusBrief_(status) {
  if (!status) return '';
  return jsonBrief_({
    destinationStatusesCount: status.destinationStatuses ? status.destinationStatuses.length : 0,
    itemLevelIssuesCount: status.itemLevelIssues ? status.itemLevelIssues.length : 0,
    creationDate: status.creationDate || '',
    lastUpdateDate: status.lastUpdateDate || '',
    googleExpirationDate: status.googleExpirationDate || ''
  });
}


function destinationStatusesBrief_(status) {
  var items = status && status.destinationStatuses ? status.destinationStatuses : [];
  var out = [];
  for (var i = 0; i < items.length; i++) {
    out.push([
      value_(items[i].reportingContext, items[i].destination),
      value_(items[i].status),
      value_(items[i].approvedCountries || items[i].approvedLocations),
      value_(items[i].disapprovedCountries || items[i].disapprovedLocations),
      value_(items[i].pendingCountries || items[i].pendingLocations)
    ].join(':'));
  }
  return out.join(' | ');
}


function itemLevelIssuesBrief_(status) {
  var items = status && status.itemLevelIssues ? status.itemLevelIssues : [];
  var out = [];
  for (var i = 0; i < items.length; i++) {
    out.push([
      value_(items[i].code),
      value_(items[i].severity),
      value_(items[i].attribute),
      value_(items[i].description),
      value_(items[i].resolution)
    ].join(':'));
  }
  return out.join(' | ');
}


function jsonBrief_(value) {
  if (!value) return '';
  return truncate_(JSON.stringify(value), 2000);
}


function value_() {
  for (var i = 0; i < arguments.length; i++) {
    if (arguments[i] === null || typeof arguments[i] === 'undefined') continue;
    if (Object.prototype.toString.call(arguments[i]) === '[object Array]') {
      if (arguments[i].length) return arrayValue_(arguments[i]);
      continue;
    }
    var value = String(arguments[i]).trim();
    if (value) return value;
  }
  return '';
}


function arrayValue_(value) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) !== '[object Array]') return String(value);
  var out = [];
  for (var i = 0; i < value.length; i++) {
    if (value[i] !== null && typeof value[i] !== 'undefined' && String(value[i]).trim()) out.push(String(value[i]).trim());
  }
  return out.join(' | ');
}


function priceObject_(price) {
  if (!price) return { value: '', currency: '' };
  if (typeof price === 'string') {
    return { value: num_(price.replace(/[^\d.,-]/g, '').replace(',', '.'), ''), currency: currencyFromPriceString_(price) };
  }
  return {
    value: typeof price.value !== 'undefined' ? price.value : (typeof price.amountMicros !== 'undefined' ? price.amountMicros / 1000000 : ''),
    currency: price.currencyCode || price.currency || ''
  };
}


function currencyFromPriceString_(price) {
  var m = String(price || '').match(/[A-Z]{3}/);
  return m ? m[0] : '';
}


function normalizeProductType_(s) {
  return String(s || '')
    .replace(/\u00A0/g, ' ')
    .replace(/\s+/g, ' ')
    .replace(new RegExp('\\s*[\\/\\u2192]\\s*', 'g'), ' > ')
    .replace(/\s*>\s*/g, ' > ')
    .trim();
}


function splitProductType_(path) {
  var parts = String(path || '').split(/\s*>\s*/);
  var out = [];
  for (var i = 0; i < parts.length && i < 5; i++) {
    var v = String(parts[i] || '').trim();
    if (v) out.push(v);
  }
  while (out.length < 5) out.push('');
  return out;
}


function extractDataSourceId_(value) {
  var s = String(value || '').trim();
  var m = s.match(/dataSources\/([^\/]+)/);
  if (m && m[1]) return m[1];
  var parts = s.split('/');
  return parts.length ? parts[parts.length - 1] : s;
}


function matchesAnyFilter_(value, filters) {
  if (!filters || !filters.length) return true;
  var normalized = String(value || '').trim().toLowerCase();
  for (var i = 0; i < filters.length; i++) {
    if (normalized === String(filters[i]).trim().toLowerCase()) return true;
  }
  return false;
}


function listSetting_(value) {
  var text = String(value || '').trim();
  if (!text) return [];
  var parts = text.split(',');
  var out = [];
  for (var i = 0; i < parts.length; i++) {
    var v = String(parts[i] || '').trim();
    if (v) out.push(v);
  }
  return out;
}


function settingOr_(value, fallback) {
  return value === null || typeof value === 'undefined' || String(value) === '' ? fallback : value;
}


function normOfferId_(id) {
  return String(id || '')
    .replace(/[\u200B-\u200D\uFEFF]/g, '')
    .replace(/\u00A0/g, ' ')
    .replace(/[\u2010-\u2015\u2212]/g, '-')
    .replace(/^online:[^:]+:[^:]+:/i, '')
    .replace(/\s+/g, ' ')
    .trim();
}


function shouldStopSoon_(settings) {
  var maxMinutes = Math.max(1, num_(settings && settings.maxRunMinutes, 25));
  var workMinutes = Math.max(1, maxMinutes - 3);
  return new Date().getTime() - START_TIME > workMinutes * 60000;
}


function bool_(value, fallback) {
  if (value === true || value === false) return value;
  var s = String(value || '').toUpperCase();
  if (s === 'TRUE' || s === 'YES' || s === '1') return true;
  if (s === 'FALSE' || s === 'NO' || s === '0') return false;
  return fallback;
}


function num_(value, fallback) {
  var n = parseFloat(value);
  return isNaN(n) ? fallback : n;
}


function truncate_(value, maxChars) {
  var text = String(value || '');
  var max = Math.max(100, num_(maxChars, 1000));
  return text.length <= max ? text : text.substring(0, max) + '... [обрізано]';
}


function naturalCmp_(a, b) {
  return String(a).localeCompare(String(b), undefined, { numeric: true, sensitivity: 'base' });
}


function runId_() {
  return Utilities.formatDate(new Date(), tz_(), 'yyyyMMdd-HHmmss') + '-' + Math.floor(Math.random() * 100000);
}


function iso_(date) {
  return Utilities.formatDate(date, tz_(), "yyyy-MM-dd'T'HH:mm:ss");
}


function tz_() {
  try {
    return AdsApp.currentAccount().getTimeZone();
  } catch (e) {
    return 'Etc/UTC';
  }
}
