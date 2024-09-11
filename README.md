# Konica_Counters_IMAP
Get Konica Minolta's counters using imap_tools

# How it works
Printer is sending an email message with its counters by configured schedule. The python script is reading mails by imap_tools then creating a data frame using pandas. The script includes a connection to a MySQL database. The code creates a table for each month to store the data (bellow you can see a example of the table). The table is exporting to a .csv file, then from .csv to the .xlsx (final file).

# How printers are configured to send mails with counters
1. Login to pirnter's web admin panel
2. Go to "Total Counter Notification Setting"
3. Enter the model name to easily identify the printer
4. Set when you want to receive mail with counters
5. You can turn on a test message to see if everything is working properly

Counters message example:
[Model Name], Pinter number 1<br />
[Serial Number], 123456789abc<br />
[Send Date], 11/09/2024<br />
[Total Counter], 139102<br />
[Total Color Counter], 0<br />
[Total Black Counter], 139102<br />
[Total Scan/Fax Counter], 14858

# Example of the table (in MySQL)

![image](https://github.com/user-attachments/assets/323362ce-c2c7-4e9f-b14d-67a6bf91733a)
