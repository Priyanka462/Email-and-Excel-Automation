*** Settings ***
Documentation     Dispatcher
...    adding queue

Library    RPA.Excel.Application
Library    RPA.Robocorp.WorkItems
Library    RPA.Tables
Library    RPA.Excel.Files

*** Variables ***


*** Tasks ***

adding queue

   ${table}=    Reading input Excel file
    FOR    ${item}    IN    @{table}
        Log    ${item}
        Create Output Work Item    ${item}
        Save Work Item
        Log To Console    ${item}
    END


*** Keywords ***
Reading input Excel file
    RPA.Excel.Files.Open Workbook    Copy of LOU.xlsx
    Read Worksheet    Sheet1
    ${table}=    Read Worksheet As Table    header=True
    RETURN    ${table}
    
