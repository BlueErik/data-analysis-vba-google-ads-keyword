## Overview
This is a small personal data analysis script. 

_Microsoft Excel Visual Basic for Applications (VBA) for Google Ads search terms and keywords data._

## Details

The VBA script functionalities:
- Read 2 input sheets.
- Create 2 new output sheets.
- Generate 2 pivot tables in each sheet.
- Set customized calculated fields in each pivot table.
- In each pivot table, do simple conditional formatting logic and highlight cells.
- The highlight cells use custom RGB colors. 
- Set basic formatting in all 4 sheets.
- Display a finished pop-up message with total runtime.

## How to use
- Download Google Ads search keyword raw data, and put in "Raw-Search keyword report".
- Download Google Ads search term raw data, and put in "Raw-Search terms report".
- Make sure the header row is at row 3.
- Make sure the header row contains required fields:
  - Raw-Search keyword report: Campaign, Ad group, Keyword, Cost, Impr., Clicks
  - Raw-Search terms report: Search term, keyword, Cost, Impr., Clicks
- Press Alt+F8 on the keyword to show the built in macro name window
- Choose "Search_Report" macro and click Run.
- Wait for the macro to finish and click OK.
- The CTR (Clicks/Impr.) and CPC (Cost/Clicks) fields are generated for each table.
- CTR threshold is < 5.17%, and CPC threshold is > $0.97. 

> [!NOTE]
> If you want to change the CTR and CPC threshold values, you can do so with the CTR_threshold and CPC_threshold values inside the module.
