Google Apps Mail Merge
----------------
Author: btong34@gmail.com 

This script allows you to mail merge using a built-in email editor within your Google spreadsheet.

###Installation
+ In your Google Docs spreadsheet, click on  'Tools' --> 'Script Editor...' and paste the code
+ Save the script and return to your spreadsheet, then refresh your browser
+ After the page loads, wait a few seconds and you should see 'Mail Merge' appear as a menu item 

###Before you start
+ Make sure your contact list data is all in **Sheet1** 
+ Open a **blank Sheet2** and title it something like 'email-template' so you remember not to delete as this is where the email draft is stored before sending
+ In order for the script to know where to send the emails, the column header for email addresses **must** be titled **'email'** (not case-sensitive).

###Using the Editor
+ Click on 'Mail-Merge' --> 'Start App' to start mail merge!
+ Compose your email using the built-in editor and include any variable in your spreadsheet by using its header name like such $%header name% (not case-sensitive either)

