VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form netserver 
   BorderStyle     =   0  'None
   Caption         =   "netmessenger - server"
   ClientHeight    =   465
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2955
   Icon            =   "netserver.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   465
   ScaleWidth      =   2955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock nswinsock 
      Left            =   180
      Top             =   -315
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NetMessenger Server"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   90
      TabIndex        =   0
      Top             =   45
      Width           =   2760
   End
End
Attribute VB_Name = "netserver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim msgtype, msgbuttons As String
Dim msgtext, msgtitle As String

Private Sub Form_Load()
msgtype = vbQuestion
msgbuttons = vbYesNo
nswinsock.Protocol = sckUDPProtocol
nswinsock.LocalPort = 1183
nswinsock.Bind
End Sub

Private Sub Label1_Click()
End
End Sub

Private Sub nswinsock_DataArrival(ByVal bytesTotal As Long)
Dim msgreply As Variant
Dim strdata As String
nswinsock.GetData strdata
    Select Case strdata
        Case "Test Server"
            nswinsock.SendData "Server is online"
        Case "vbcritical"
            msgtype = vbCritical
            Exit Sub
        Case "vbquestion"
            msgtype = vbQuestion
            Exit Sub
        Case "vbinformation"
            msgtype = vbInformation
            Exit Sub
        Case "vbexclamation"
            msgtype = vbExclamation
            Exit Sub
        Case "vbabortretryignore"
            msgbuttons = vbAbortRetryIgnore
            Exit Sub
        Case "vbokcancel"
            msgbuttons = vbOKCancel
            Exit Sub
        Case "vbokonly"
            msgbuttons = vbOKOnly
            Exit Sub
        Case "vbretrycancel"
            msgbuttons = vbRetryCancel
            Exit Sub
        Case "vbyesno"
            msgbuttons = vbYesNo
            Exit Sub
        Case "vbyesnocancel"
            msgbuttons = vbYesNoCancel
            Exit Sub
        Case "domsg"
            msgreply = MsgBox(msgtext, msgbuttons + msgtype, msgtitle)
            If msgreply = vbYes Then nswinsock.SendData "...YES"
            If msgreply = vbNo Then nswinsock.SendData "...NO"
            If msgreply = vbOK Then nswinsock.SendData "...OK"
            If msgreply = vbCancel Then nswinsock.SendData "...CANCEL"
            If msgreply = vbRetry Then nswinsock.SendData "...RETRY"
            If msgreply = vbAbort Then nswinsock.SendData "...ABORT"
            If msgreply = vbIgnore Then nswinsock.SendData "...IGNORE"
            Exit Sub
        End Select
        If Left$(strdata, 2) = "> " Then msgtitle = strdata
        If Left$(strdata, 1) = " " Then msgtext = strdata
End Sub
