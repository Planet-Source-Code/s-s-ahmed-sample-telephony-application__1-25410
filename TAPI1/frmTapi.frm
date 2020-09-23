VERSION 5.00
Begin VB.Form frmTapi 
   Caption         =   "TAPI "
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6240
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   6240
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   495
      Left            =   3240
      TabIndex        =   5
      Top             =   2640
      Width           =   1455
   End
   Begin VB.ListBox lstStatus 
      Height          =   1620
      Left            =   2400
      TabIndex        =   4
      Top             =   720
      Width           =   3735
   End
   Begin VB.CommandButton cmdStartSession 
      Caption         =   "Start Session"
      Height          =   495
      Left            =   1200
      TabIndex        =   2
      Top             =   2640
      Width           =   1455
   End
   Begin VB.TextBox txtPhonenum 
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   240
      Width           =   3735
   End
   Begin VB.Label lblStatus 
      Caption         =   "Status:"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   615
   End
   Begin VB.Label lblEnterPhone 
      Caption         =   "Enter phone number to dial:"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "frmTapi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'############################################################
'Author: S.S. Ahmed
'Email: ss_ahmed1@hotmail.com
'Program: TAPI Sample Application
'Date: Jul 21, 2001
'Note: This sample code is provided without any support
'############################################################


Private Sub cmdQuit_Click()
Unload Me
End
End Sub

Private Sub cmdStartSession_Click()

Dim lonTAPIStatus As Long

'Check to see if a phone number was entered
If RTrim(txtPhonenum.Text) = "" Then
    lstStatus.AddItem "Err: No phone number entered"
    Exit Sub
Else
    strphonenum = RTrim(txtPhonenum.Text)
End If

'Initialize the TAPI session with the tapiRequestMakeCall function
lonTAPIStatus = tapiRequestMakeCall(strphonenum, _
    "TAPI Sample", strphonenum, "")

'Report the status
Call TAPIStatus(lonTAPIStatus)

End Sub

Private Sub TAPIStatus(lonStatCode As Long)

'Based on the TAPI status code (passed to this procedure in lonStatCode),
' add an appropriate message to the lststatus listbox

Select Case lonStatCode
    Case TAPIERR_CONNECTED
        lstStatus.AddItem "OK"
    Case TAPIERR_DROPPED
        lstStatus.AddItem "Dropped"
    Case TAPIERR_NOREQUESTRECIPIENT
        lstStatus.AddItem "Err: No Request Recipient"
    Case TAPIERR_REQUESTQUEUEFULL
        lstStatus.AddItem "Err: Request Queue Full"
    Case TAPIERR_INVALDESTADDRESS
        lstStatus.AddItem "Err: Destination Address Invalid"
    Case TAPIERR_INVALWINDOWHANDLE
        lstStatus.AddItem "Err: Window Handle Invalid"
    Case TAPIERR_INVALDEVICECLASS
        lstStatus.AddItem "Err: Device Class Invalid"
    Case TAPIERR_INVALDEVICEID
        lstStatus.AddItem "Err: Device ID Invalid"
    Case TAPIERR_DEVICECLASSUNAVAIL
        lstStatus.AddItem "Err: Device Class Unavailable"
    Case TAPIERR_DEVICEIDUNAVAIL
        lstStatus.AddItem "Err: Device ID Unavailable"
    Case TAPIERR_DESTBUSY
        lstStatus.AddItem "Err: Destination Busy"
    Case TAPIERR_DESTUNAVAIL
        lstStatus.AddItem "Err: Destination Unavailable"
    Case TAPIERR_UNKNOWNWINHANDLE
        lstStatus.AddItem "Err: Unknown Windows Handle"
    Case TAPIERR_UNKNOWNREQUESTID
        lstStatus.AddItem "Err: Unknown Request ID"
    Case TAPIERR_REQUESTFAILED
        lstStatus.AddItem "Err: Request Failed"
    Case TAPIERR_REQUESTCANCELLED
        lstStatus.AddItem "Err: Request Cancelled"
    Case TAPIERR_INVALPOINTER
        lstStatus.AddItem "Err: Invalid Pointer"
    
End Select

End Sub
