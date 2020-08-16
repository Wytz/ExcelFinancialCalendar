# ExcelFinancialCalendar
Create a financial calendar in Excel with this VBA script

I found that lots of mutations are actually very predictable, but yet I never felt in control. Hence me creating this script.

What it basically does is create a calendar in Excel, and inserts your planned income/expenditures for each day in this calendar. 
It also calculates a final balance at the end of each day, and your monthly total. 
This helps you predict your daily balance for the whole year. 
Every month is shown in a separate sheet, and every day is represented by a small table that shows the name of the mutation and the mutated amount.

The end result should look like the screenshot below.

[[/images/Example.png|Example of the outcome]]

##Why Excel/VBA?

While there are existing tools that pretty much do the same, those are often paid and are limited in the operations you can perform with the data. 
My main reason to program this in Excel are:

* Lots of tools to play around and visualize the data. You are in complete control as a user.
* Excel is ubiquitous, so lots of people should be able to use it. Once you have generated the sheet you can probably also upload it to the free online version of Excel (I haven't tested this yet).
* Excel updates all linked values real-time when you update a field, so you always have an up-to-date view.
* Excel actually performed surprisingly well when generating the sheets, given my crappy code. Creating a reasonably complex overview takes roughly a minute.
* Excel has funky features which automatically fills in the names of the months and days of the weeks in the language that you have configured it.


##HOW TO USE:

* Open the attached "Template.xlsx" workbook
* This contains two sheets, one named "Income", and one named "Expenditures" (These names can be changed in the code if you want to name them differently)
* Please make sure that the columns with the dates in the Income and Expenditure sheets are actually recognized as dates by Excel. This appears to be especially relevant for people who use the US style of notating dates. You can do so in the "Home" tab on the ribbon, above the "Number" field. Also feel free to try with a minimal example first. Mismatches between the notations can cause runtime errors (see comments). A similar issue may also occur with the Euro currency symbol.
* Please don't reshuffle the columns or the top rows. Those are hardcoded. However, it is fine if you add additional rows
* Enable the "Developer" tab in the ribbon. On the File tab, go to Options > Customize Ribbon. Under Customize the Ribbon and under Main Tabs, select the Developer check box.
* In the "Developer" tab on the ribbon, click on the "Visual Basic" button to the left.
* A new screen should pop up. To the left there is a frame which is titled "Project - VBAProject".
* Right-click within an empty space in this frame, click on "Insert" and then on "Module".
* Copy/paste the VBA code in this post, and click on the little play icon at the top, under "Debug".

If necessary, the next steps help you to change modify the calendar generation.

* Fill in the year for which you want to generate the calendar. This year should correspond with the year of the dates in your Income/Expenditure sheets.
* Estimate how many mutations you have as a maximum per day. The default has been set to 8, but if you need more you should increase this number.
* The other variables are names of sheets or descriptive stuff.
* Once you have clicked on the play button, the calendar should start to generate.

This is my first VBA project, so please don't crucify me for the crappy/inefficient code :) You're free to reuse and distribute it, but if you do please credit me.