Blackboard Local Grade Tool (Browser)

How to use
1) Open index.html in a browser.
2) Build Moderation Workbook:
   - Upload Blackboard export (.xls/.csv)
   - Module code is auto-read from Blackboard headers (for example {001}{MTHA4007B-25-SEM2-B} -> MTHA4007B)
   - Download generated .xlsx workbook
3) Fill in moderation marks in Excel.
4) Merge Moderated Marks:
   - Upload original Blackboard export
   - Upload filled workbook
   - Download Blackboard upload file (.xls name, CSV content)

Notes
- All processing is local in your browser.
- No server upload is used.
- Child course sheets (if present) are merged into one upload file named *_upload.xls (CSV content).
- Coursework count is inferred from Blackboard headers like {001}{MODULECODE}; weighting cells are left blank for manual entry.
