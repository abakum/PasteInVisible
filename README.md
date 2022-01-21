# PasteInVisible
* By Ctrl+C, excel copies consecutive cells to the clipboard. It not excludes cells that are filtered, hidden, or grouped.
* By Ctrl+V, excel pastes from the clipboard into consecutive rows and columns not only visible but also those that are filtered, hidden or grouped. [Look Note: at bottom](https://support.microsoft.com/en-us/office/copy-visible-cells-only-6e3a1f01-2884-4332-b262-8b814412847e#:~:text=Note%3A%C2%A0Excel%20pastes%20the%20copied%20data%20into%20consecutive%20rows%20or%20columns.)
* To paste  from clipboard in only visible cells use this code.
# Usage: 
* Ctrl+C - copy connected (CR) or fragmented range (FR) to clipboard (CB)
* Shift+Ctr+C - convert the selected range (SR) from CR to possibly fragmented by grouping or filters FR and save it as (RR) see SelectVisible
* Ctrl+D - replicate the first row of SR to the whole SR including rows hidden by grouping or filters
* Shift+Ctr+C Ctrl+D - replicate the first row of SR to the entire RR, not including rows hidden by grouping or filters
* Ctrl+R - replicate the first column of SR to the entire SR including columns hidden by grouping
* Shift+Ctr+C Ctrl+R - replicate the first column of SR to the entire RR, not including columns hidden by grouping
* Ctrl+C Ctrl+V - insert CR or FR from CB into the selected CR with border expansion including cells hidden by grouping or filters
* Ctrl+C Ctr+Alt+V - insert CR or FR from CB into the selected CR with border expansion including cells hidden by grouping or filters and selecting the type of insertion
* Shift+Ctr+C Shift+Ctr+X - Paste RR into SR without extending borders, not including cells hidden by grouping or filters, see PasteX
* Shift+Ctr+C Shift+Ctr+V - Paste RR into SR without extending borders and pasting values, not including cells hidden by grouping or filters, see PasteV
* Alt+F8 SaveAsAddIn Run - Save and set ThisWorkbook as AddIn see SaveAsAddIn
# [Использование:](https://github.com/abakum/PasteInVisible/blob/master/usage.rus.txt)
