*** Settings ***
Library    canoe_robot_lib.py
Library    tenma_robot_lib.py

*** Variables ***
${CFG_PATH}    D:\\CAN Simulater APK\\test\\UDS stack CI Automation\\CANoe Config File\\Config file\\UDS_Stack_Automation01.cfg
${TEST_MODULE}    test_capl
${TENMA_PORT}    COM10
${VOLTAGE}    12
${CURRENT}    3

*** Test Cases ***
Run UDS CAPL Test Module
    [Setup]    Setup Tenma PSU    ${TENMA_PORT}    ${VOLTAGE}    ${CURRENT}
    Open Canoe Configuration    ${CFG_PATH}
    Start Canoe Measurement
    Run Test Module    ${TEST_MODULE}
    Stop Canoe Measurement
    [Teardown]    Close Tenma

*** Keywords ***
Setup Tenma PSU
    [Arguments]    ${port}    ${voltage}    ${current}
    Connect Tenma    ${port}
    Set Tenma Voltage And Current    ${voltage}    ${current}
    Tenma Power On
