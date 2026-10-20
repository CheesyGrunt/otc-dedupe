# OTC Dedupe

Drop in the Daily OTC Creations export (Salesforce) and get back one row per Client: Contact ID.

- Keeps the first row for each ID, or drops every row whose ID repeats (toggle)
- Shows the updated list and the removed rows on separate tabs
- Downloads an .xlsx with "Updated List" and "Removed Rows" sheets, or copies the list for pasting into Excel
- Runs entirely in the browser. The file never leaves your computer.

Static site on Render: no build command, publish directory `.`
