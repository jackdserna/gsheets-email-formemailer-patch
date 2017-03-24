---
title: "FormEmailer patch walkthrough"
author: "Jack D. Serna"
date: "3/24/2017"
output: html_document
---
Using FormEmailer by Henrique Abreu: https://sites.google.com/site/formemailer/form-emailer

Version 3.65 

"FormEmailer is a Google Apps Script for Google Forms and Spreadsheets. With it you're able to easily set up flexible e-mail merge in your spreadsheets and automatic emailing to your forms."

Patch walkthrough:

After installing "FormEmailer" into your desired google sheet in the script editor, adjusting your settings, and adding your trigger for automation, you may notice errors inside your spreadsheet as well as seemingly endless blank emails. 

If you're having emails sent to you that must mean you have set up an automatic trigger (time-driven) and you have set up the settings to send emails to your address (for testing, I assume?). If you have an email address column you want to send emails to, that should be the column in which a placeholder references.

To stop sending blank emails to yourself in the thousands, you have to adjust the code so that it knows when to trigger! The original code does not account for when to stop emailing on an automatic trigger.

Line 209:
```{r}
var status = c.fs.getRange(2,1,last-1,1).getValues();
```
Line 209 is defining a variable within your monitor function. This variable "status" is getting the values in a specified range. Specifically, it is getting the current form sheet range. Google apps script has 4 different ways of writing getRange(), and this case has 4 arguments like this: 
```{r}
getRange(row, column, numRows, numColumns)
```
The original code as is is saying, "From my current form sheet, start from the second row and first column, go all the way to the last row, and do all of this for an entire one column (column 1)".


But what if you have some data on rows 2:5 for columns B:Z (2:26 in javascript), and you have no data on rows 6 and beyond? You're emails will be sent for rows 2:5, and you'll receive error messages for lines 6:1000 or so (whatever "last" is) in column 1(javascript) or column 0 (excel/google sheet); in other words, column A will update with hundreds of error messages and depending on your settings you may have received hundreds of blank emails!

Update Line 209 precisely at the last argument of getRange(), where you specify how many columns you want to collect data. Choose up to the column where you know there will always be data, such as a Timestamp column if you're using google forms: 
```{r}
var status = c.fs.getRange(2,1,last-1,INSERT_REFERENCE_COLUMNS_HERE).getValues();
```

You've just set up (a) reference column(s) to check whether there is data to be emailed! Now, the next code is your for loop which checks status line by line, or row by row. In the code at Line 211 you are setting the condition when to run your email function called "doIt".

As is, Line 211 if( status[i][0] === '' ) says, "If status, at the first column, is equal to blank... then run function doIt". But column A is not the column where my data lives! Column A is where my email statuses and error messages live. This condition is only preventing resending emails that have already been sent, or from sending emails where there has been an error.

Here is what your if statement should look like at Line 211:
```{r}
if( status[i][0] === '' && status[i][YOUR REFERENCE COLUMN] !== '' )
  doIt_(i+2);
if( status[i][0] === '' && status[i][YOUR REFERENCE COLUMN] === '' )
  return;
```
This statement is saying, "If status at my first column is blank AND my reference column is not blank, then run my email function doIt...If status at my first column is blank AND my reference column is blank, then stop the for loop."

Save your code. You have just updated your code to accept automatic triggers! Set your time driven trigger and give it a test run!
