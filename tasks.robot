*** Settings ***
Documentation       huuto.net/tori.fi hakusana ilmoituksien taltioija Excel-tiedostoon

Library             RPA.Browser.Selenium    auto_close=${FALSE}
Library             OperatingSystem
Library             RPA.Excel.Files
Library             Collections
Library             RPA.Desktop
Library             RPA.Excel.Files
Library             RPA.Netsuite


*** Variables ***
${selain}       Firefox
${url}          https://www.tori.fi/
${hakusana}     Nintendo Switch
@{otsikko}
@{hinta}
@{linkki}
@{päivä}
${column}       ${2}


*** Tasks ***
huuto.net/tori.fi hakusana ilmoituksien taltioija Excel-tiedostoon
    Open the intranet website
    Accept terms and services
    Hakusana haku
    Tallenna ilmoitukset
    Täytä Excel ilmoituksilla
    Robotti sulkeutuu


*** Keywords ***
Open the intranet website
    Open Available Browser    ${url}    ${selain}

Accept terms and services
    Sleep    2 seconds
    Select Frame    //iframe[@title="SP Consent Message"]
    ${button}=    Get WebElement    //button[@title="Hyväksy kaikki evästeet"]
    Click Element    ${button}
    Unselect Frame
    Sleep    1 seconds

Hakusana haku
    Input Text    id:frontSearch    ${hakusana}
    Click Button    id:searchFront

Tallenna ilmoitukset
    #Tallenna ilmoitukset nappaavat ensin ilmoituksen elementin, ja sitten ottavat siitå tekstit muuttijaan
    #Listataaan: otsikko, hinta, linkki, päivämäärä
    #Tästä alkaa for loopit, jotka täyttävät yllä luodut listat
    FOR    ${i}    IN RANGE    999
        Wait Until Page Contains Element    //a[@class=" item_row_flex odd_row"]
        ${ilmoitusOtsikko}=    Get WebElements    //div[@class="li-title"]
        FOR    ${element}    IN    @{ilmoitusOtsikko}
            ${otsikkoon}=    Get Text    ${element}
            Append To List    ${otsikko}    ${otsikkoon}
            Log    ${otsikko}
        END
        ${ilmoitusHinta}=    Get WebElements    //p[@class="list_price ineuros"]
        FOR    ${element}    IN    @{ilmoitusHinta}
            ${hintaan}=    Get Text    ${element}
            Append To List    ${hinta}    ${hintaan}
            Log    ${hinta}
        END
        ${ilmoitusLinkki}=    Get WebElements    //a[contains(@class, 'item_row_flex')]
        FOR    ${element}    IN    @{ilmoitusLinkki}
            ${linkkiin}=    Get Element Attribute    ${element}    href
            Append To List    ${linkki}    ${linkkiin}
            Log    ${linkki}
        END
        ${ilmoitusPäivä}=    Get WebElements    //div[@class="date_image"]
        FOR    ${element}    IN    @{ilmoitusPäivä}
            ${päivään}=    Get Text    ${element}
            Append To List    ${päivä}    ${päivään}
            Log    ${päivä}
        END
        #Painaa Seuraava sivu nappia ja tarkistaa, että onko se olemassa. Jos ei, loppuu loop.
        ${onkoSeuraavaNappi}=    Run Keyword And Return Status
        ...    Page Should Contain Element
        ...    //a[.="Seuraava sivu »"]
        Exit For Loop If    ${onkoSeuraavaNappi} == False
        IF    ${onkoSeuraavaNappi} == True
            Click Element    //a[.="Seuraava sivu »"]
        END
    END
    Close Browser

Täytä Excel ilmoituksilla
    #Luodaan uusi taulukko aina, tehdään headerit ja täytetään listan tavaroilla
    Create Workbook    ilmoitukset.xlsx
    Set Worksheet Value    1    1    Otsikko
    Set Worksheet Value    1    2    Hinta
    Set Worksheet Value    1    3    Linkki (Paina avataksesi)
    Set Worksheet Value    1    4    Päivämäärä
    FOR    ${element}    IN    @{otsikko}
        Set Worksheet Value    ${column}    A    ${element}
        ${column}=    Evaluate    ${column} + 1
    END
    ${column}=    Set Variable    2
    FOR    ${element}    IN    @{hinta}
        Set Worksheet Value    ${column}    B    ${element}
        ${column}=    Evaluate    ${column} + 1
    END
    ${column}=    Set Variable    2
    FOR    ${element}    IN    @{linkki}
        Set Worksheet Value    ${column}    C    ${element}
        ${column}=    Evaluate    ${column} + 1
    END
    ${column}=    Set Variable    2
    FOR    ${element}    IN    @{päivä}
        Set Worksheet Value    ${column}    D    ${element}
        ${column}=    Evaluate    ${column} + 1
    END
    Save Workbook

Robotti sulkeutuu
    Log To Console    Robotti on valmis!
