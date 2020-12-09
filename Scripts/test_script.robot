*** Settings ***
Documentation     A test suite with a single test for undefined
...               Created by hats' Robotcorder
Library           Selenium2Library    timeout=10

*** Variables ***
${BROWSER}    chrome
${SLEEP}    3

*** Test Cases ***
undefined test
    Open Browser    undefined    ${BROWSER}
    Input Text    //input[@name="username"]    konkocho@in.ibm.com
    Input Text    //input[@name="password"]    ***
    Click Element    //button[@id="signin"]
    Click Element    //path[@class="svg-icon03"]
    Click Link    //a[@id="zr3:0:zt0:0:pt1:tt1:1::di"]

    Close Browser