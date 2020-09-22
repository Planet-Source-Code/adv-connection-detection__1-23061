VERSION 5.00
Begin VB.Form FrmMain 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   2520
      Width           =   1215
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim flags As Long
Dim result As Boolean

    result = InternetGetConnectedState(flags, 0)
    If result Then
        Print "Connected to the Internet"
    Else
        Print "Not Connected to the Internet"
    End If
     
    If flags And INTERNET_CONNECTION_MODEM Then Print "Connection Via Modem"
    If flags And INTERNET_CONNECTION_LAN Then Print "Connecion Via LAN"
    If flags And INTERNET_CONNECTION_PROXY Then Print "Connection uses a Proxy"
    If flags And INTERNET_CONNECTION_MODEM_BUSY Then Print "Connection Via Modem but modem is busy"

End Sub
