# Transparency Project



1.	What is this?

> A program that generates Transparency data pages for Admissions, which can then be rendered as PDFs for the website.

2.	What do I need to run it?

> It’s written in Classic ASP so a version of IIS or a server to host it. Either a remote or local instance of SQL Server and SQL Management Studio. A free program called wkhtmltopdf and a knowledge of IIS and SQL to get the set-up running.
> To view a page you need to check the database for program codes to see what pages you can generate. For Example: http://localhost:8080/Program-Profile-Report.asp?ppc=IBHT-P where the query string ppc= ProgramPlanCode in the database table ProgramProfile.

3.	How do I create PDFs from the data
>	Download and install wkhtmltopdf https://wkhtmltopdf.org/ then pass the program a list of all the input pages and all the output files. Run this query against the DB.
SELECT Distinct ([ProgramPlanCode]),[DocumentName]  
FROM [transparency].[dbo].[ProgramProfile]

>	The query will return a list of all the pages and the PDF doc names. Put these into Excel or a text editor using delimited fields and use this as a template to build the command(s).

>	Run a cmd prompt from the wkhtmltopdf tool folder: e.g. C:\Program Files (x86)\wkhtmltopdf\bin 

>	Here is a sample of what your command should look like, you will need this line per every report... If you have prepared the data in a text editor you can copy the entire file and right-click into the command prompt and it will process every line (if you use the example below make sure you have a C:\temp folder that you can write to.

wkhtmltopdf -O landscape http://localhost:8080/program-profile-report.asp?ppc=ISM-DBBN C:\temp\ISM-DBBN.pdf

 
 
