Attribute VB_Name = "modGlobal"
'############################################################
'Author: S.S. Ahmed
'Email: ss_ahmed1@hotmail.com
'Program: TAPI Sample Application
'Date: Jul 21, 2001
'Note: This sample code is provided without any support
'############################################################


'Declare the main function in this module

Declare Function tapiRequestMakeCall Lib "tapi32" _
    (ByVal lpszDestAddress As String, _
    ByVal lpszAppName As String, _
    ByVal lpszCalledParty As String, _
    ByVal lpszComment As String) As Long
    
'Global constants

Global Const TAPIERR_CONNECTED = 0&
Global Const TAPIERR_DROPPED = -1&
Global Const TAPIERR_NOREQUESTRECIPIENT = -2&
Global Const TAPIERR_REQUESTQUEUEFULL = -3&
Global Const TAPIERR_INVALDESTADDRESS = -4&
Global Const TAPIERR_INVALWINDOWHANDLE = -5&
Global Const TAPIERR_INVALDEVICECLASS = -6&
Global Const TAPIERR_INVALDEVICEID = -7&
Global Const TAPIERR_DEVICECLASSUNAVAIL = -8&
Global Const TAPIERR_DEVICEIDUNAVAIL = -9&
Global Const TAPIERR_DEVICEINUSE = -10&
Global Const TAPIERR_DESTBUSY = -11&
Global Const TAPIERR_DESTNOANSWER = 12&
Global Const TAPIERR_DESTUNAVAIL = -13&
Global Const TAPIERR_UNKNOWNWINHANDLE = -14&
Global Const TAPIERR_UNKNOWNREQUESTID = -15&
Global Const TAPIERR_REQUESTFAILED = -16&
Global Const TAPIERR_REQUESTCANCELLED = -17&
Global Const TAPIERR_INVALPOINTER = -18&



