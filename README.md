# ExcelReader
A console-based application to load an Excel file dynamically,
allowing you to view the data and add rows.

## Given Requirements:

- [x] This is an application that will read data from an Excel 
spreadsheet into a database
- [x] When the application starts, it should delete the database 
if it exists, create a new one, create all tables, read from Excel, 
seed into the database.
- [x] You need to use EPPlus package
- [x] You shouldn't read into Json first.
- [x] You can use SQLite or SQL Server (or MySQL if you're 
using a Mac)
- [x] Once the database is populated, you'll fetch data from it 
and show it in the console.
- [x] You don't need any user input
- [x] You should print messages to the console letting the user 
know what the app is doing at that moment (i.e. reading from excel; 
creating tables, etc)
- [x] The application will be written for a known table, you don't 
need to make it dynamic.
- [x] When submitting the project for review, you need to include an 
xls file that can be read by your application.

## Features

- SQLite connection:

  - The data is stored in a SQLite database for the session.

- Console-based UI to show the application status:

  - imagelink

- Console-based UI to interact with the Excel

  - imageLinks

- Insert rows and save the Excel:

  - You can insert a row in the worksheet selected.
  - imageLink
  - You can save the Excel by entering a path.
  - imageLink

## Challenges

  - Reading an Excel file.
  - Querying metadata from SQLite for dynamic tables and fields.
  - Writing an Excel file.

## Lessons Learned

  - Working with Excel in .NET using EPPlus.
  - Retrieving tables from a SQLite database.
  - Retrieving columns from a table.
  - Validating paths.
  - Dynamic type conversion.

## Areas to Improve

  - Working with files.
  - LINQ.
  - SQLite metadata knowledge.

##  Resources used

  - StackOverflow posts
  - links