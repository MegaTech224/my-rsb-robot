*** Settings ***
Documentation     Template robot main suite, now also tracked through git.
Library     RPA.Browser.Selenium    auto_close=${FALSE}    
Library    RPA.HTTP
Library    RPA.Excel.Files
Library    RPA.PDF

*** Tasks ***
Insert the sales data for the week and export it as a pdf
    [Teardown]    Log out and close the browser
    Open the intranet website
    Log in
    Download The Excel File
    Fill the form using the data from the Excel file
    Collect the results
    Export the table as a PDF
    
*** Keywords ***
Open the intranet website
    Open Browser    https://robotsparebinindustries.com/    edge    

Log in
    Input Text    username    maria
    Input Password    password    thoushallnotpass
    Click Button    Log in
    Wait Until Page Contains Element    sales-form

Fill and submit the form for one person
    [Arguments]    ${voornaam}    ${achternaam}    ${verkoopdoel}    ${resultaat}
    Input Text    firstname    ${voornaam}
    Input Text    lastname    ${achternaam}
    Select From List By Value    salestarget    ${verkoopdoel}
    Input Text    salesresult    ${resultaat}
    Click Button    Submit

Download The Excel File
    Download    https://robotsparebinindustries.com/SalesData.xlsx    overwrite=${True}    target_file=${OUTPUT DIR}${/}SalesData.xlsx

Fill the form using the data from the Excel file
    Open Workbook    ${OUTPUT DIR}${/}SalesData.xlsx
    @{sales_reps}    Read Worksheet As Table    header=${True}
    Close Workbook
    FOR    ${row}    IN    @{sales_reps}
        Fill and submit the form for one person    ${row}[First Name]    ${row}[Last Name]    ${row}[Sales Target]    ${row}[Sales]
    END

Collect the results
    Screenshot    //div[contains(@class, 'sales-summary') and contains(.//span, 'Active sales people')]    ${OUTPUT DIR}${/}sales-summary.png

Export the table as a PDF
    Wait Until Element Is Visible    //div[contains(@class, 'sales-summary') and contains(.//span, 'Active sales people')]
    ${sales_results_html}    Get Element Attribute    //div[contains(@id,'sales-results')]    outerHTML
    Html To Pdf    ${sales_results_html}    ${OUTPUT DIR}${/}sales_results.pdf

Log out and close the browser
    Click Button    Log out
    Close Browser