# Description
    This is a simple console program that connects to quick books datasource and retreives all tables.
    This exports all tables in CSV format which can then be used to export to any other Accounting System

# Prerequisites
    Need QuickBooks ODBC Connector Installed
    http://www.qodbc.com/download.htm 


    Install DotNet Core on the machine you need to run this

    cd to folder you extracted this to.
    Change paths in Program.cs ( I've hardcoded for sake of demo .. Or change code to make it relative path)

    run
    
    dotnet build  (TO build Code)

    and

    dotnet run    (To RUn Code)