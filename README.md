Call-Detail-Record--CDR--Manager
================================

This is a VB 6.0 application which manages a Call Detail Record file.

It takes a CDR which is a text file, SPL_Output_24052012.txt in this case. The Call Detail Record (CDR) contains details of all Inbound and Outbound calls of a call center within a time duration (for eg. in a day).

 The CDR manager performs the following operations :-
 
 1) Browse : Selecting the text file from the host computer which is to be used.
 
 2) Read : Reads all the contents of the text file and displays them on the 'INPUT FILE' screen.
 
 3) Split : Separates each word of the file and displays them in the 'SPLIT ITEMS' screen.
 
 4) Find MS-Excel Format : Creates an MS-Excel file and writes the call detail record in the .xlsx sheet. The words that    were split in the previous stage is written in the excel sheet cells in this stage. 
 
 5) Outbound Calls : Creates an MS-Excel file, indentifies and seggregates the details of all the Outbound calls from the CDR text file, and writes them in the .xlxs sheet. Thus it creates a sepatate MS-Excel file for Outbound call records.
 
 6) Inbound Calls : Creates an MS-Excel file, indentifies and seggregates the details of all the Inbound calls from the CDR text file, and writes them in the .xlxs sheet. Thus it creates a sepatate MS-Excel file for Inbound call records.
 
 7) Manage Inbounds : Creates an MS-Excel file, indentifies and seggregates the details of all the Inbound calls from the CDR text file, further structures the data for better management of the Inbound Record, and writes them in the .xlxs sheet.
 
    The .xlxs sheets can be saved in the host computer for later use.
