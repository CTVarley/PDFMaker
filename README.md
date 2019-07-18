This is a lightweight, bespoke Google App Script application for converting Google Sheets to PDF for each day of the week.

To add this script to your Google Sheet:
  1. In your Google Sheet file, in the top menu go to Tools >> Script Editor
  2. Copy and Paste the contents of PDFMaker.gs into the Script Editor Window

This should add a new menu item to your sheet, called PDF Publisher. The first time you run it, it will raise a security warning. You will have to click on "advanced" to allow all permissions.

This script currently expects the following information to exist in a specific cell:
  - A sheet named School_Delivery, a sheet named COLD_Delivery, and a sheet named Prep
  - The day of the week in cell E2 of a sheet named School_Delivery
  - The date in cell E3 of a sheet named School_Delivery
  - The number of the week in the cycle included in cell E4 of a sheet named School_Delivery, without any other numbers.
  - The tables in the Prep_Sheet to maintain their specific width in columns

If any of your own changes to the spreadsheet effect the pagination of the School_Delivery or COLD_Delivery sheets, just adjust the height of the blank row between pages.

If there are any issues, requests or suggestions, we urge you to raise them at https://github.com/lilithbuilds/PDFMaker

Thanks for using this PDF Maker!
