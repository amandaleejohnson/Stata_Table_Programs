# Stata Table Programs
Programs to automate the production of tables from Stata to MS Excel.

If you want to present column percentages you can use “Table1catcol” and for row percentages use “Table1catrow”. These are only for presenting **non-weighted** estimates (right now). 

To use in Stata, run the following commands:

do “filepathway\table1catrow.do”
do “filepathway\table1catcol.do”

table1catrow names-of-row-variables, by(name-of-column-variable) replace
table1catcol names-of-row-variables, by(name-of-column-variable) replace

The above line will produce a table with the following title: “Table 1. Column percentages of categorical variables by [name-of-column-variable]”

Both programs can also incorporate “if statements”:

table1catrow names-of-row-variables **if female==1**, by(name-of-column-variable) replace

This will produce a table with the following title: “Table 1. Column percentages of categorical variables by [name-of-column-variable] **if [if-statement]**”


**NOTE – The Excel file name is generated with the “by variable”. So, if you type “table1catrow age sex race, by(smoking)”, the Excel file will be called “table1_bysmoking.xlsx”. If you change the list of covariates in the table but keep the same by variable, you will overwrite the Excel file each time you run it.**
