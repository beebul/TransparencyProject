# Transparency Project

1.	What is this?

> A program that generates Transparency data pages for Admissions, which can then be rendered as PDFs for the website.

2.	What do I need to run it?

> It’s written in Classic ASP so a version of IIS or a server to host it. Either a remote or local instance of SQL Server and SQL Management Studio. A free program called wkhtmltopdf and a knowledge of IIS and SQL to get the set-up running.
> To view a page you need to check the database for program codes to see what pages you can generate. For Example: http://localhost:8080/Program-Profile-Report.asp?ppc=IBHT-P where the query string ppc= ProgramPlanCode in the database table ProgramProfile.