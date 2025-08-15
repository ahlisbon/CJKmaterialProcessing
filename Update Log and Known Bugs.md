# Issues and Bug with Chinese Japanese Korean Material Processing (CJKmP)

##Bugs
Some Series RE titles have trailing text "; </font>" ^ 2024-11-25
Some Series RE titles have volume data ",; v."
Some Total Vols display data like "volumes : illustrations ; 22 cm"

##Features
Add new field in GUI to build search strings with OCLC#s so that if an item is a dupe in WorldCat, if can link back to the local record in the spreadsheet. ^ 2025-08-08

##Fixed/Implimented
- Updated .ini filename from "Bibdata to Spreadsheet" to "CJKmP - Settings" ^ 2025-05-29 ^ 1.03
- FirstSearch loads with "Home" screen instead of "Basic Search" or "Advanced Search" ^ 2025-05-29 ^ 1.03
- Formatting of total volume data for journals, should equal n/a ^ 2024-11-22 ^ 1.02
- FirstSeach Dupe column has a link to your local catalog. ^ 2024-11-22 ^ 1.02
- Can search with ISSNs in ISBN10 column, returns ISSN to ISBN-10 column. ^ 2024-11-22 ^ 1.02
- When Converting ISBN-13s to ISBN-10s, ISBN-10s that should end in 0 instead end in 11 ^ 2024-11-19 ^ 1.01
