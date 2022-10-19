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

*** Tasks ***
Hae dataa sähkön kulutuksesta ja hinnasta
    Avaa verkkoselain ja mene Vantaan Energia Sähköverkkojen sivuille
    Kirjaudu sisään ja mene raportointi sivulle
    Etsi ja lataa kulutusdata
    Hae sähkön hinta

*** Keywords ***
Avaa verkkoselain ja mene Vantaan Energia Sähköverkkojen sivuille
    Set Download Directory    ${OUTPUT_DIR}
    Open Available Browser    https://online.vantaanenergiasahkoverkot.fi/eServices/Online/IndexNoAuth
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