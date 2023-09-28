*** Settings ***
Documentation     Inserir dados de vendas para a semana e exportar como PDF.

Library    RPA.Browser.Selenium
Library    RPA.HTTP
Library    RPA.Excel.Files
Library    RPA.PDF

*** Tasks ***
Inserir dados de vendas para a semana e exportar como PDF
    Abrir website intranet
    Logar
    Baixar arquivo Excel
    Preencher o formulario a partir do arquivo Excel
    Coletar os resultados
    Exportar a tabela como PDF
    [Teardown]    Deslogar e fechar o browser

*** Keywords ***
Abrir website intranet
    Open Available Browser    https://robotsparebinindustries.com/    headless=True

Logar
    Input Text    id:username    maria
    Input Password    id:password    thoushallnotpass
    Submit Form
    Wait Until Page Contains Element    id:sales-form

Baixar arquivo Excel
    Download    https://robotsparebinindustries.com/SalesData.xlsx    overwrite=True

Preencher e enviar o formulario para uma pessoa
    [Arguments]    ${sales_rep}
    Input Text    firstname    ${sales_rep}[First Name]
    Input Text    lastname    ${sales_rep}[Last Name]
    Input Text    salesresult    ${sales_rep}[Sales]
    Select From List By Value    salestarget    ${sales_rep}[Sales Target]
    Click Button    Submit

Preencher o formulario a partir do arquivo Excel
    Open Workbook    SalesData.xlsx
    ${sales_reps}=    Read Worksheet As Table    header=True
    Close Workbook
    FOR     ${sales_reps}    IN   @{sales_reps}
        Preencher e enviar o formulario para uma pessoa    ${sales_reps}
    END

Coletar os resultados
    Screenshot    css:div.sales-summary    ${OUTPUT_DIR}${/}sales_summary.png

Exportar a tabela como PDF
    Wait Until Element Is Visible    id:sales-results
    ${sales_results_html}=    Get Element Attribute    id:sales-results    outerHTML
    Html To Pdf    ${sales_results_html}    ${OUTPUT_DIR}${/}sales_results.pdf

Deslogar e fechar o browser
    Click Button    Log out
    Close Browser