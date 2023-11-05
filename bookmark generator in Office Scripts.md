# Bookmark generator in Office Scripts

Office Scripts allows to automate tasks by recording, editing, and running scripts in Excel on both of web and desktop, it is typescript.

Bookmark generator is good demo to show how the common task works, the main difference than VBA on desktop is UI, which is not existing for Office Scripts, so I output generated bookmark file in new sheet or console as examples.

## How to run Bookmark-console
- run script from Excel Automate tab 
- input data from selection of 2 columns: name and url which starting with http or https
- generated bookmark under a folder which name is from active sheet name
- output the bookmark file on console

## How to run Bookmark-sheet
- run script from Excel Automate tab 
- input data from selection of 2 columns: name and url which starting with http or https
- generated bookmark under a folder which name is from active sheet name
- create new sheet 'bookmark' and place bookmark file on A1 cell
  
## Bookmark generator script
- [Bookmark-console.osts](https://github.com/robertluwang/hands-on-auto/blob/main/src/osts/Bookmark-console.osts)
- [Bookmark-sheet.osts](https://github.com/robertluwang/hands-on-auto/blob/main/src/osts/Bookmark-sheet.osts) 

