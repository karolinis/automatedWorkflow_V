# automatedWorkflow_V
A colleague of mine needed to automate her workflow where she receives 100s of PDF files, needs to look through each of them &amp; collect data, validate data and upload documents to their respective entries in "E-conomic" ERP system.

The way workflow works:

**
Lauch the PDF scraper first, it looks through all the PDFs in the current working directory and by using regex filters out the important data. 
It's very easy to filter out the data as the format of incoming invoices doesn't change.
This program gives us output.xlsx
**

**
We now have data from our scraper, but need to validate it. To do so, we can download original entry data from E-conomic. 
With output.xlsx and the file you just obtained, you can use difference finder ( a program I wrote to find differences in two Excel files that backfired with extra functionality than I thought )
Difference finder allows you to choose which files to work with. It then uses 3 ID columns (Date, Invoice, Account in our case) to identify entries in both files and check if their balances match.
If the balances do match - great! This means that the PDF scraper was correct and now the program can apply the information user chooses to display from BOTH files.
This is important, because in output.xlsx we are missing data in column Entry, but have File (filename). In original entry file from E-conomic we do have a column Entry, but are missing a column File.
We need those two columns to use our last program that will upload documents to E-conomic via API.

If the balances do not match - Difference column will be populated with the difference between balances.

Program makes a file Compared_output.xlsx
**

**
For post_invoices - a program that uploads documents to API, we need a file "working_file.xlsx". We can rename Compared_output to working_file and make sure that File and Entry columns are populated.
Entry is the entry number that will have File attached to it.
Simply run the program, enter agreement codes and watch the console outputting
                  
                  201 Created OK - entry + str(i)
                  201 Created OK - entry + str(i)
