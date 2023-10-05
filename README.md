# Northern Ireland Shippments

## Overview
The "Northern Ireland Shippments" is a robust console application microservice designed to fetch data from a WMS database and Excel document, matching the data and present it in another Excel format template file distributed by email as per the customer's weekly requirement.

## Key Features
* Built with .NET 6, EF Core, Dapper, Excel InterOp, Clean Architecture, and CQRS
* Reads configuration from external .xml and .sql files
* Writes logs to a database table and a .txt file
* Uses SQL Database and Excel file as source of data
* Uses MS Excel as presentation layer
* Distributes .xlsm file as attachment using SMTP
* Measures the end-to-end execution time of the app

### License
The "Northern Ireland Shippments" is one of 70+ commercial microservice products created by Michal Zielinski for the warehouse and logistics ecosystem.
