# PasteInVisible
* By Ctrl+C, excel copies only visible cells to the clipboard. It excludes cells in rows or columns that are filtered, hidden, or grouped.
* By Ctrl+V, excel pastes cells from the clipboard into consecutive rows and columns not only visible but also those that are filtered, hidden or grouped. [Note:](https://support.microsoft.com/en-us/office/copy-visible-cells-only-6e3a1f01-2884-4332-b262-8b814412847e#:~:text=Note%3A%C2%A0Excel%20pastes%20the%20copied%20data%20into%20consecutive%20rows%20or%20columns.)
* To paste cells from clipboard into only visible cells use this AddIn.
* Shift+Ctr+K is useful when consolidating parts of a table into a whole table.
## Usage: 
* Ctrl+C - copy consecutive (CR) or fragmented by hidding, grouping or filtering range (FR) of visible cells to clipboard (CB)
* Shift+Ctr+C - convert the selected range (SR) from CR or FR and save it (RR). Look [SelectVisible](https://github.com/abakum/PasteInVisible/blob/main/PasteInVisible.bas#:~:text=Sub%20SelectVisible(Optional%20hide_from_Macros_dialog_box%20As%20Boolean)), [Copy visible cells only](https://support.microsoft.com/en-us/office/copy-visible-cells-only-6e3a1f01-2884-4332-b262-8b814412847e)
* Ctrl+D - replicate the first row of SR to the whole SR including rows hidden by grouping or filters
* Shift+Ctr+C Ctrl+D - replicate the first row of SR to the entire RR, not including rows hidden by grouping or filters
* Ctrl+R - replicate the first column of SR to the entire SR including columns hidden by grouping
* Shift+Ctr+C Ctrl+R - replicate the first column of SR to the entire RR, not including columns hidden by grouping
* If FR is in CB Ctrl+V - inserts from it into SR including cells hidden by grouping or filtering values and formats
* If CR is in CB Ctrl + V - pastes from it into SR including cells hidden by grouping or filtering formulas and formats
* Ctrl+C Ctrl+Alt+V - pastes CR or FR from CB to SD including cells hidden by grouping or filtering with choice of insertion type
* Ctrl+C Shift+Ctr+X - Paste RR into SR without extending borders, not including cells hidden by grouping or filters. Look [PasteX](https://github.com/abakum/PasteInVisible/blob/main/PasteInVisible.bas#:~:text=Sub%20PasteX(Optional%20val%20As%20Boolean%20%3D%20False%2C%20Optional%20key%20As%20Boolean%20%3D%20False))
* Ctrl+C Shift+Ctr+V - Paste RR into SR without extending borders and pasting values, not including cells hidden by grouping or filters. Look [PasteV](https://github.com/abakum/PasteInVisible/blob/main/PasteInVisible.bas#:~:text=Sub%20PasteV(Optional%20hide_from_Macros_dialog_box%20As%20Boolean))
* Shift+Ctr+K -  same as in Shift+Ctr+V, but only empty cells (EC) are replaced and only if all key cells (not EC) are equal. Look [PasteK](https://github.com/abakum/PasteInVisible/blob/main/PasteInVisible.bas#:~:text=Sub%20PasteK(Optional%20hide_from_Macros_dialog_box%20As%20Boolean))
## Installation:
* Alt+F8 SaveAsAddIn Run - Save and set ThisWorkbook as AddIn. Look [SaveAsAddIn](https://github.com/abakum/PasteInVisible/blob/main/PasteInVisible.bas#:~:text=a/70916088/18055780-,Sub%20SaveAsAddIn(),-%27Alt%2BF8%20SaveAsAddIn)
## Сonsolidation example:
* There is a table `Whole` that needs to be filled in by different parts (sections, branches, subdivisions)
* Red `Part1&2` - the result of filling in the table `Whole`
* Blue `Part3` - the result of filling the table` Whole`
* If parts in different books open them,
* filter `Whole` by Part<3, filter `Part1&2` by Part<3, select from `Part1&2` Ctrl+C, paste into `Whole` Shift+Ctrl+K,
* filter `Whole` by Part=3, filter `Part3` by Part=3, select from `Part3` Ctrl+C paste into `Whole` Shift+Ctrl+K
## [Использование, установка, пример консолидации](https://github.com/abakum/PasteInVisible/blob/master/usage.rus.txt)
## Credits
* [Fastest Method to Copy Large Number of Values in Excel VBA](https://stackoverflow.com/questions/36272285/fastest-method-to-copy-large-number-of-values-in-excel-vba)
* [How do you take a vba addin and make an installer](https://stackoverflow.com/questions/18397251/how-do-you-take-a-vba-addin-and-make-an-installer)
