*** Settings ***
Documentation       Hakee dataa Vantaan Energia Sähköverkkojen sivustolta sekä sähkön SPOT-hinta tietoja ja näiden perusteella
...                 yrittää laskea kuluvan päivän ja huomisen hinnan sähkölaskulle.
Library             RPA.Browser.Selenium    auto_close=${FALSE}
Library             RPA.HTTP
Library             OperatingSystem
Library             KeePassLibrary
Variables           credentials.py    #Koska en ole vielä saanut KeePass ominaisuutta toimimaan, tiedot ladataan erillisestä muuttujatiedostosta nyt toistaiseksi.
Library             scripts/lastSevenDays.py
Library             Process
Library             DateTime
Library             RPA.Excel.Files
Library             RPA.Tables

*** Variables ***
${pohja}               ${CURDIR}//output//Robottipohja.xlsx
${hinnat}              ${CURDIR}//output//data.xlsx    
${hinnathuomenna}      ${CURDIR}//output//huominen.xlsx    
${kulutus}             ${CURDIR}//output//Sähkö_12102022-18102022.csv

*** Tasks ***
Hae dataa sähkön kulutuksesta ja hinnasta
    Avaa verkkoselain ja mene Vantaan Energia Sähköverkkojen sivuille
    Kirjaudu sisään ja mene raportointi sivulle
    Etsi ja lataa kulutusdata
    Hae sähkön hinta

Liitä tiedot Exceliin
    Korvaa pohjan hinnat
    Korvaa huomiset hinnat
    Korvaa pohjan kulutus
    Hae tiedot

*** Keywords ***
Avaa verkkoselain ja mene Vantaan Energia Sähköverkkojen sivuille
    Set Download Directory    ${OUTPUT_DIR}
    Open Chrome Browser    https://online.vantaanenergiasahkoverkot.fi/eServices/Online/IndexNoAuth
    Maximize Browser Window    #Suurennetaan ikkuna, jotta kaikki elementit varmasti näykyvät botille.

Hae kirjautumistunnukset
    #Tämä osuus ei jostain syystä toimi. Sitä pitää tutkia lisää myöhemmin.
    Open Keepass Database    Database.kdbx    keyfile=Database.key
    ${entry}=    Get Entries    User Name    first=True
    ${username}=    Get Entry Username    ${entry}  
    ${password}=    Get Entry Password    ${entry}
    Log     Username is ${username}

Kirjaudu sisään ja mene raportointi sivulle
    Input Text    UserName    ${username}
    Input Password    Password    ${password}
    Submit Form
    Wait Until Page Contains Element    id:fpGraphHeader3

Etsi ja lataa kulutusdata
    Click Element    id:fpGraphHeader3
    ${startDate}=    Get Start Date
    Input Text When Element Is Visible    id:startDateSelector    ${startDate}
    ${endDate}=    Get End Date
    Input Text When Element Is Visible    id:endDateSelector    ${endDate}
    Click Button    id:updateInterval
    Sleep    2s    #Pieni paussi, jotta tiedot ehtiävät päivittyä.
    Click Button    id:ExportToExcel

Hae sähkön hinta
    Go To    https://www.vattenfall.fi/sahkosopimukset/porssisahko/tuntispot-hinnat-sahkoporssissa/
    Click Element When Visible    id:cmpbntnotxt    #Evästeet -_-
    Sleep    1s
    Run Keyword And Ignore Error    Scroll Element Into View    xpath:/html/body/main/section/main/div/div[1]/div/price-spot-fi/div[2]/div/div/div/button    #Skrollataan valmiiksi latausnapin luokse.
    Sleep    1s
    Execute Javascript    document.querySelector("#startDate").removeAttribute("readonly");    #Poistetaan readonly attribuutti, jotta voimme suoraan kirjoittaa päivämäärän.

    #Haetaan tämän päivän data:

    Sleep    1s
    ${currentDate}=    Get Current Date    exclude_millis=True
    ${startDate}=    Convert Date    ${currentDate}    result_format=%d.%m.%Y
    Input Text When Element Is Visible    id:startDate    ${startDate}
    sleep    1s
    Click Element When Visible    id:periodSelector
    Sleep    1s
    Click Element    xpath:/html/body/main/section/main/div/div[1]/div/price-spot-fi/div[1]/div/div/div[2]/div/div[2]/period-picker/div/form/div/ul/li[1]/a
    Sleep    1s
    Click Element    xpath:/html/body/main/section/main/div/div[1]/div/price-spot-fi/div[2]/div/div/div/button

    #Ja vielä huomisen data:

    Run Keyword And Ignore Error    Scroll Element Into View    xpath:/html/body/main/section/main/div/div[1]/div/price-spot-fi/div[2]/div/div/div/button
    #Pitää taas scrollata, koska Chromen latauspalkki saattaa estää elementin.
    ${currentDate}=    Get Current Date    exclude_millis=True
    ${tomorrow}=    Add Time To Date    ${currentDate}    1 days
    ${startDate}=    Convert Date    ${tomorrow}    result_format=%d.%m.%Y
    Input Text When Element Is Visible    id:startDate    ${startDate}
    sleep    1s
    Click Element    xpath:/html/body/main/section/main/div/div[1]/div/price-spot-fi/div[2]/div/div/div/button

# Leevin osuus
Hae hintatiedot
    [Arguments]    ${workbook}
    Open Workbook    ${workbook}
    ${hintapohja}=    Read Worksheet As Table   WorkSheet
    [Return]    ${hintapohja}

Hae kulutustiedot
    ${kulutustiedot}=    Read table from CSV     ${kulutus}     delimiters=;
    [Return]    ${kulutustiedot}
    
Korvaa pohjan hinnat
    ${hinnat}=    Hae hintatiedot   ${hinnat}
    Open Workbook    ${pohja}
    Set Active Worksheet   Hinnat

    ${index}=    Set Variable     1
    FOR    ${RIVI}    IN    @{hinnat}
        # Set running index
        FOR    ${SARAKE}    IN    @{RIVI}
            Log    ${SARAKE}
            
            Set Cell Value    ${index}    ${SARAKE}    ${RIVI}[${SARAKE}]
        END
        ${index}=    Evaluate   ${index} + 1
    END
    Save Workbook   ${pohja}

Korvaa huomiset hinnat
    ${hinnat}=    Hae hintatiedot  ${hinnathuomenna}
    Open Workbook    ${pohja}
    Set Active Worksheet   Hinnat huomenna

    ${index}=    Set Variable     1
    FOR    ${RIVI}    IN    @{hinnat}
        # Set running index
        FOR    ${SARAKE}    IN    @{RIVI}
            Log    ${SARAKE}
            
            Set Cell Value    ${index}    ${SARAKE}    ${RIVI}[${SARAKE}]
        END
        ${index}=    Evaluate   ${index} + 1
    END
    Save Workbook   ${pohja}
    
Korvaa pohjan kulutus
    ${kulutus}=    Hae kulutustiedot
    Open Workbook    ${pohja}
    Set Active Worksheet   KulutusRaaka

    ${index}=    Set Variable     2
    ${sarakeindex}=    Set Variable     1
    FOR    ${RIVI}    IN    @{kulutus}
        ${sarakeindex}=    Set Variable     1
        FOR    ${SARAKE}    IN    @{RIVI}            
            IF    "${SARAKE}" == "Energia yhteensä (kWh)"
                Set Cell Value    ${index}    ${sarakeindex}    ${RIVI}[${SARAKE}]  fmt="##.####"      
            ELSE 
                Set Cell Value    ${index}    ${sarakeindex}    ${RIVI}[${SARAKE}]
            END
            ${sarakeindex}=    Evaluate   ${sarakeindex} + 1
        END
        ${index}=    Evaluate   ${index} + 1
    END
    Save Workbook   ${pohja}
    
Hae tiedot
    Open Workbook    ${pohja}
    ${table}=     Read Worksheet As Table  Data
    ${kalleinhintatänään}=       Get Cell Value    30    G
    ${kalleinhintahuomenna}=     Get Cell Value    30    H    
    ${kalleintuntitänään}=       Get Cell Value    31    G
    ${kalleintuntihuomenna}=     Get Cell Value    31    H

    ${halvinhintatänään}       Get Cell Value    32    G
    ${halvinhintahuomenna}     Get Cell Value    32    H
    ${halvintuntitänään}       Get Cell Value    33    G
    ${halvintuntihuomenna}     Get Cell Value    33    H

    
    Log    Tämän päivän kallein tunti on ${kalleintuntitänään} ja se maksaa ${kalleinhintatänään}. Huomenna kallein tunti on ${kalleintuntihuomenna} ja se maksaa ${kalleinhintahuomenna}. 
    Log    Halvin tunti tänään on ${halvintuntitänään} ja se maksaa ${halvinhintatänään}. Huomenna halvin tunti on ${halvintuntihuomenna} ja se maksaa ${halvinhintahuomenna}.   