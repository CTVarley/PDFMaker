This is a lightweight, bespoke Google App Script application for converting Google Sheets to PDF for each day of the week.

To add this script to your Google Sheet:
  1. In your Google Sheet file, in the top menu go to Tools >> Script Editor
  2. Copy and Paste the contents of PDFMaker.gs into the Script Editor Window

This should add a new menu item to your sheet, called PDF Publisher. The first time you run it, it will raise a security warning. You will have to click on "advanced" to allow all permissions.

If there are any issues, we urge you to raise them at https://github.com/lilithbuilds/PDFMaker

Thanks for using this PDF Maker!

This script currently expects the following information to exist in a specific cell:
  - The day of the week in cell E2 of a sheet named School_Delivery
  - The date in cell E2 of a sheet named School_Delivery