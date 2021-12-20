# Excel to JSON Converter
### VBA on Excel, converts an Excel table to .JSON file.

## Functionality
This tool converts your Excel table to a JSON file. 
It can be useful for Excel users who need a simple way to convert their Excel data to JSON (JavaScript Object Notation), rigth from Excel (no coding needed). 
JSON has become a standard for data interchange, APIs, and online applications. However, there's no built-in Excel conversion. 

This tool supports 1- or 2-level objects. An example of 2-level is `{"EmployeeRef":{"ID":"55","name":"Emily Platt"}}`.
However, in this version the second-level object can contain only 2 key-values (e.g., "ID" and "Name"). 
It outputs .JSON file to: `[your home dir]\JSON output\timesheets.json`, where [your home dir] is typically your \Documents folder.

## Instrucitons
The .xlsx workbook contains all the needed VBA code. Follow the instructions in the workbook to format your data, then just click the button to run. 
You can modify the VBA macro to fit your needs. 
The app is also saved separately as .bas file. For custom modifications, contact me at the email in my profile.

## Setup/Dependencies
I use the VBA-JSON JsonConverter, source: https://github.com/VBA-tools/VBA-JSON, as a function module.
If you're using my .bas code separately (not the complete Excel worksheet), then you'll need to install the JsonConverter module as well - follow the link above for instrucitons.
Also, you'll need to turn on the Microsoft Scripting Runtime library. 
