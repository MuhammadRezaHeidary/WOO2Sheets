# WOO2Sheets
Sending data from Woocommerce  to Google Sheets and processing some features

This is a simple google appsscript code for processing recieved data from woocommerce. 

Primary data is sent to google sheets by using api connector and woocommerce rest api in 2 independent GET requests! Each GET request should be saved in a specific sheet. 
So we should have a sheet for woocommerce orders and a sheet for woocommerce customers. 
We also should create a final sheet for processing and showing usefull data. 
Finally we will use <code>WOO2Sheets.gs</code> for processing and saving data from these two sheets in the final sheet.
