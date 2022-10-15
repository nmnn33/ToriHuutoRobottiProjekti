*** Settings ***
Documentation       huuto.net/tori.fi hakusana ilmoituksien taltioija Excel-tiedostoon

Library             RPA.Browser.Selenium    auto_close=${FALSE}
Library             OperatingSystem
Library             RPA.Excel.Files
Library             Collections
Library             RPA.Desktop
Library             RPA.Excel.Files


*** Variables ***
${selain}       Firefox
${url}          https://www.tori.fi/
${hakusana}     Nintendo Switch


*** Tasks ***
huuto.net/tori.fi hakusana ilmoituksien taltioija Excel-tiedostoon
    Open the intranet website
    Accept terms and services
    Hakusana haku
    Tallenna ilmoitukset
    #Laita ilmoitukset dictionaryyn
    Täytä Excel ilmoituksilla


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
    @{otsikko}=    Create List
    @{hinta}=    Create List
    @{linkki1}=    Create List
    @{linkki2}=    Create List
    @{päivä}=    Create List
    #Tästä alkaa for loopit, jotka täyttävät yllä luodut listat
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
    #Kaksi eri linkki listaa, sillä tori.fi jakaa ne kahtia
    ${ilmoitusLinkki1}=    Get WebElements    //a[@class=" item_row_flex"]
    FOR    ${element}    IN    @{ilmoitusLinkki1}
        ${linkkiin1}=    Get Element Attribute    ${element}    href
        Append To List    ${linkki1}    ${linkkiin1}
        Log    ${linkki1}
    END
    ${ilmoitusLinkki2}=    Get WebElements    //a[@class=" item_row_flex odd_row"]
    FOR    ${element}    IN    @{ilmoitusLinkki2}
        ${linkkiin2}=    Get Element Attribute    ${element}    href
        Append To List    ${linkki2}    ${linkkiin2}
        Log    ${linkki2}
    END
    ${ilmoitusPäivä}=    Get WebElements    //div[@class="date_image"]
    FOR    ${element}    IN    @{ilmoitusPäivä}
        ${päivään}=    Get Text    ${element}
        Append To List    ${päivä}    ${päivään}
        Log    ${päivä}
    END

#Laita ilmoitukset dictionaryyn
#    ${dict}=    create dictionary
#    FOR    ${element}    IN    @{dict}
#    Set To Dictionary    ${dict}    otsikko=    hinta=    päivä=    linkki=
#    END

Täytä Excel ilmoituksilla
    #Luodaan uusi taulukko aina, tehdään headerit ja täytetään listan tavaroilla
    Create Workbook    ilmoitukset.xlsx
    Set Worksheet Value    1    1    Otsikko
    Set Worksheet Value    1    2    Hinta
    Set Worksheet Value    1    3    Linkki (Paina avataksesi)
    Set Worksheet Value    1    4    Päivämäärä
    FOR    ${element}    IN    @{otsikko}
        Log    ${element}
    END
    Save Workbook
