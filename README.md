# reporting_macros
All the VB Script codes for custom functions and and macros in reporting excel sheets can be written here, so that they can be at one place

# Notes
You can format VB code at http://warp.senecac.on.ca/timothy.mckenna/CodeFormatter.asp

1. SCHUPD seperate sheets required for Gen SCH, NET, UI; Const SCH, NET, UI; Volt Max Min

# Todos
1. Use Exit For to exit for loop after match occurence -- done
2. Limit iteration to 1000 -- done

## Reporting facts
1. We have raw data
2. We derive base data from raw data
3. Presentation data will be derived from raw data and base data

## Philosophies for reporting
1. All the data values should be accessible via labels and not only by addresses
2. Base data should directly be derived from raw data. It should not be derived from derived data or another base data
3. While using formulas or functions that data input variables should not be repeating and should follow a pattern, so that we can drag down to get the series
