VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form netclient 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Net Messenger - by Nick Smith"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5970
   Icon            =   "client.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   5970
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdHELP 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Help"
      Height          =   1185
      Left            =   2205
      TabIndex        =   26
      Top             =   3285
      Width           =   1185
   End
   Begin VB.Frame frames 
      Caption         =   "Message Buttons"
      Height          =   2040
      Index           =   0
      Left            =   45
      TabIndex        =   22
      Top             =   2430
      Width           =   1995
      Begin VB.OptionButton btns_okonly 
         Caption         =   "OK only"
         Height          =   195
         Left            =   180
         TabIndex        =   9
         Top             =   270
         Value           =   -1  'True
         Width           =   915
      End
      Begin VB.OptionButton btns_okcancel 
         Caption         =   "OK , Cancel"
         Height          =   195
         Left            =   180
         TabIndex        =   10
         Top             =   540
         Width           =   1230
      End
      Begin VB.OptionButton btns_yesno 
         Caption         =   "Yes, No"
         Height          =   195
         Left            =   180
         TabIndex        =   11
         Top             =   810
         Width           =   915
      End
      Begin VB.OptionButton btns_yesnocancel 
         Caption         =   "Yes, No, Cancel"
         Height          =   195
         Left            =   180
         TabIndex        =   12
         Top             =   1080
         Width           =   1500
      End
      Begin VB.OptionButton btns_retrycancel 
         Caption         =   "Retry, Cancel"
         Height          =   195
         Left            =   180
         TabIndex        =   13
         Top             =   1350
         Width           =   1275
      End
      Begin VB.OptionButton btns_abortretryignore 
         Caption         =   "Abort, Retry, Ignore"
         Height          =   195
         Left            =   180
         TabIndex        =   14
         Top             =   1620
         Width           =   1725
      End
   End
   Begin VB.Frame frames 
      Caption         =   "Message Type"
      Height          =   2400
      Index           =   1
      Left            =   45
      TabIndex        =   20
      Top             =   0
      Width           =   1995
      Begin VB.OptionButton types_critical 
         Caption         =   "Critical"
         Height          =   195
         Left            =   630
         TabIndex        =   5
         Top             =   405
         Value           =   -1  'True
         Width           =   825
      End
      Begin VB.OptionButton types_question 
         Caption         =   "Question"
         Height          =   195
         Left            =   630
         TabIndex        =   6
         Top             =   900
         Width           =   960
      End
      Begin VB.OptionButton types_information 
         Caption         =   "Information"
         Height          =   195
         Left            =   630
         TabIndex        =   7
         Top             =   1395
         Width           =   1095
      End
      Begin VB.OptionButton types_exclamation 
         Caption         =   "Exclamation"
         Height          =   195
         Left            =   630
         TabIndex        =   8
         Top             =   1890
         Width           =   1185
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   135
         Picture         =   "client.frx":030A
         Top             =   270
         Width           =   480
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   135
         Picture         =   "client.frx":074C
         Top             =   765
         Width           =   480
      End
      Begin VB.Image Image3 
         Height          =   480
         Left            =   135
         Picture         =   "client.frx":0B8E
         Top             =   1260
         Width           =   480
      End
      Begin VB.Image Image4 
         Height          =   480
         Left            =   135
         Picture         =   "client.frx":0FD0
         Top             =   1755
         Width           =   480
      End
   End
   Begin VB.CommandButton cmdSEND 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Send"
      Height          =   1185
      Left            =   4725
      TabIndex        =   16
      Top             =   3285
      Width           =   1185
   End
   Begin VB.CommandButton cmdPREVIEW 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Preview"
      Height          =   1185
      Left            =   3465
      TabIndex        =   15
      Top             =   3285
      Width           =   1185
   End
   Begin VB.Frame frames 
      Caption         =   "Text Settings"
      Height          =   2400
      Index           =   2
      Left            =   2115
      TabIndex        =   17
      Top             =   0
      Width           =   3795
      Begin VB.CommandButton cmdANALYSESERVER 
         Caption         =   "Analyse Server"
         Height          =   285
         Left            =   2160
         TabIndex        =   1
         Top             =   315
         Width           =   1500
      End
      Begin VB.TextBox txtREPLY 
         Height          =   825
         Left            =   810
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         ToolTipText     =   "Will display a list of replies from the remote computer."
         Top             =   1395
         Width           =   2895
      End
      Begin VB.TextBox txtMESSAGE 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   285
         Left            =   810
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "The message that will be sent in the message box."
         Top             =   1035
         Width           =   2880
      End
      Begin VB.TextBox txtCAPTION 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   285
         Left            =   810
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "The caption that will appear in the title of the sent message."
         Top             =   675
         Width           =   2880
      End
      Begin VB.TextBox txtTARGET 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   285
         Left            =   810
         TabIndex        =   0
         TabStop         =   0   'False
         Text            =   "localhost"
         ToolTipText     =   "The target IP of the remote computer that the message will be sent to."
         Top             =   315
         Width           =   1260
      End
      Begin MSWinsockLib.Winsock nswinsock 
         Left            =   2070
         Top             =   180
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         Protocol        =   1
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "User Reply:"
         Height          =   390
         Left            =   90
         TabIndex        =   23
         Top             =   1575
         Width           =   510
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Target:"
         Height          =   195
         Left            =   90
         TabIndex        =   21
         Top             =   360
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Message:"
         Height          =   195
         Left            =   90
         TabIndex        =   19
         Top             =   1080
         Width           =   690
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Caption:"
         Height          =   195
         Left            =   90
         TabIndex        =   18
         Top             =   720
         Width           =   585
      End
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "This application needs a server program to be running on the target computer. Download HERE."
      ForeColor       =   &H00FF0000&
      Height          =   390
      Left            =   2205
      TabIndex        =   25
      Top             =   2520
      Width           =   3570
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Visit the homepage @ http://come.to/magikcube"
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   2205
      TabIndex        =   24
      Top             =   3015
      Width           =   3570
   End
End
Attribute VB_Name = "netclient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim receive As String
Private Sub cmdANALYSESERVER_Click()
MsgBox "This will test if the server program is active on the target computer. If no alert appears after 5 secs then the server program is not installed on the target machine. If an alert does occur, then the server is active, and Net Messenger will function properly.", vbOKOnly + vbInformation, "Analyse Server"
nswinsock.RemoteHost = txtTARGET.Text
nswinsock.SendData "Test Server"
End Sub

Private Sub cmdHELP_Click()
help.Show
End Sub

Private Sub cmdPREVIEW_Click()
Dim msgbuttonslocal As String
Dim msgtypelocal As String
    If types_critical.Value = True Then msgtypelocal = vbCritical
    If types_question.Value = True Then msgtypelocal = vbQuestion
    If types_information.Value = True Then msgtypelocal = vbInformation
    If types_exclamation.Value = True Then msgtypelocal = vbExclamation
    If btns_abortretryignore.Value = True Then msgbuttonslocal = vbAbortRetryIgnore
    If btns_okcancel.Value = True Then msgbuttonslocal = vbOKCancel
    If btns_okonly.Value = True Then msgbuttonslocal = vbOKOnly
    If btns_retrycancel.Value = True Then msgbuttonslocal = vbRetryCancel
    If btns_yesno.Value = True Then msgbuttonslocal = vbYesNo
    If btns_yesnocancel.Value = True Then msgbuttonslocal = vbYesNoCancel
    MsgBox " " & txtMESSAGE.Text, msgbuttonslocal + msgtypelocal, "> " & txtCAPTION.Text
End Sub

Private Sub cmdSEND_Click()
    nswinsock.RemoteHost = txtTARGET.Text
    nswinsock.SendData "> " & txtCAPTION.Text
    nswinsock.SendData " " & txtMESSAGE.Text
    If types_critical.Value = True Then nswinsock.SendData "vbcritical"
    If types_question.Value = True Then nswinsock.SendData "vbquestion"
    If types_information.Value = True Then nswinsock.SendData "vbinformation"
    If types_exclamation.Value = True Then nswinsock.SendData "vbexclamation"
    If btns_abortretryignore.Value = True Then nswinsock.SendData "vbabortretryignore"
    If btns_okcancel.Value = True Then nswinsock.SendData "vbokcancel"
    If btns_okonly.Value = True Then nswinsock.SendData "vbokonly"
    If btns_retrycancel.Value = True Then nswinsock.SendData "vbretrycancel"
    If btns_yesno.Value = True Then nswinsock.SendData "vbyesno"
    If btns_yesnocancel.Value = True Then nswinsock.SendData "vbyesnocancel"
    nswinsock.SendData "domsg"
End Sub

Private Sub Form_Load()
'help.Show
'netserver.Show
nswinsock.Protocol = sckUDPProtocol
nswinsock.RemotePort = 1183
nswinsock.RemoteHost = txtTARGET.Text
txtTARGET.Text = nswinsock.LocalIP
txtCAPTION.Text = "Caption used for the title."
txtMESSAGE.Text = "Main text used for the message."
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub nswinsock_DataArrival(ByVal bytesTotal As Long)
    nswinsock.GetData receive
    If Left$(receive, 3) = "..." Then txtREPLY.Text = txtREPLY.Text & receive & vbCrLf
    If receive = "Server is online" Then MsgBox "This is a reply from the installed server, indicating that " & vbCrLf & "the server on the target machine is active. " & vbCrLf & vbCrLf & "Net Messenger will now function properly.", vbOKOnly + vbInformation, "Server is online"
End Sub

Private Sub txtTARGET_Change()
nswinsock.RemoteHost = txtTARGET.Text
End Sub
