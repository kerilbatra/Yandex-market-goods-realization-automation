# Yandex-market-goods-realization-automation
Automation script for downloading Yandex Market "Goods Realization" reports, cleaning the Excel file (keeping only required columns/sheets), merging with SharePoint data, removing duplicates, and uploading back automatically. The project includes last-month auto-selection, API calls, file cleanup, SharePoint integration, and Excel transformation.

# Yandex Market Goods Realization Automation

This repository contains a fully automated Python workflow that:

✔ Generates a **Goods Realization** report from the Yandex Market API  
✔ Downloads the generated Excel file  
✔ Cleans and restructures the sheet (removes extra sheets, deletes rows, keeps specific columns)  
✔ Merges the cleaned data with existing SharePoint Excel files  
✔ Removes duplicates using selected columns  
✔ Uploads the updated file back to SharePoint  
✔ Deletes local temporary files  

The script is designed for **daily/monthly automation** and requires no manual Excel work.

---

## 🚀 Features

- Auto-detect previous month (month/year)
- Calls Yandex Market API `/reports/goods-realization`
- Downloads report using the `file` link
- Removes unwanted sheets
- Unmerges all merged cells and fills values
- Deletes first 16 rows
- Keeps only important columns:
  - Название товара  
  - Доставлено, шт.  
  - Дата оформления заказа  
  - Стоимость всех доставленных штук с НДС без учёта скидок, ₽  
- Removes empty rows
- Merges with existing SharePoint file
- Deduplicates data based on key columns
- Uploads final file back to SharePoint



