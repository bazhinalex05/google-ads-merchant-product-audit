# Merchant Product Type Export

Google Ads Script для експорту `product_type` з Merchant Center після застосування правил перетворення.

## Код

Актуальний файл скрипта: [`script.js`](script.js)

## Як встановити

1. Відкрийте `script.js` у GitHub.
2. Натисніть кнопку копіювання коду.
3. У Google Ads відкрийте `Tools -> Bulk actions -> Scripts`.
4. Створіть новий скрипт.
5. Вставте код.
6. У верхній частині скрипта замініть `SPREADSHEET_URL` на URL Google Sheets клієнта.
7. Увімкніть Advanced APIs у Google Ads Scripts:
   - `Merchant API -> Products`
   - `Merchant API -> Accounts`, якщо `auto_register_gcp_project = TRUE`
8. Запустіть скрипт, щоб він створив листи.
9. У листі `Settings` заповніть `merchant_id`.
10. За потреби налаштуйте фільтри `merchant_data_source_id_filter`, `merchant_feed_label_filter`, `merchant_content_language_filter`.
11. Запустіть скрипт ще раз і перевірте результат.

## Результат

Скрипт створює або оновлює листи:

- `Settings` - налаштування запуску.
- `ProductTypeExport` - проста таблиця для перенесення `product_type`: `ID` і `Product Type`.
