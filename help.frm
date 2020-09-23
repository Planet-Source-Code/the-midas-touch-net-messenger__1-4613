VERSION 5.00
Begin VB.Form help 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7605
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3870
   LinkTopic       =   "Form1"
   ScaleHeight     =   7605
   ScaleWidth      =   3870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   7575
      Left            =   45
      TabIndex        =   0
      Top             =   -45
      Width           =   3750
      Begin VB.CommandButton cmdCLOSE 
         Caption         =   "&Close"
         Height          =   285
         Left            =   1215
         TabIndex        =   14
         Top             =   7155
         Width           =   1230
      End
      Begin VB.Label Label14 
         Caption         =   "Preview: This will allow you to see your message before it is sent to the remote computer."
         Height          =   465
         Left            =   225
         TabIndex        =   15
         Top             =   5985
         Width           =   3435
      End
      Begin VB.Label Label13 
         Caption         =   "Analyse Server: This command will send a test message to the remote computer, which will then reply, if the server is active."
         Height          =   645
         Left            =   225
         TabIndex        =   13
         Top             =   6435
         Width           =   3345
      End
      Begin VB.Label Label12 
         Caption         =   "Message Buttons: This will affect the buttons that will be displayed when the user receives the sent message."
         Height          =   645
         Left            =   225
         TabIndex        =   12
         Top             =   5310
         Width           =   3390
      End
      Begin VB.Label Label11 
         Caption         =   "Message Type: This will affect the image type that the message will display, according to the user's selection."
         Height          =   690
         Left            =   225
         TabIndex        =   11
         Top             =   4635
         Width           =   3390
      End
      Begin VB.Label Label10 
         Caption         =   "User Reply: This will show which of the command buttons the user clicks on when the message appears on the remote computer."
         Height          =   645
         Left            =   225
         TabIndex        =   10
         Top             =   3960
         Width           =   3345
      End
      Begin VB.Label Label9 
         Caption         =   "Message: This is the text that will appear in the main section of the message when the user receives the sent message."
         Height          =   645
         Left            =   225
         TabIndex        =   9
         Top             =   3285
         Width           =   3390
      End
      Begin VB.Label Label8 
         Caption         =   "Caption: This is the text that will appear in the title bar of the message when the user receives the sent message."
         Height          =   645
         Left            =   225
         TabIndex        =   8
         Top             =   2610
         Width           =   3255
      End
      Begin VB.Label Label7 
         Caption         =   "Target: This is the target IP address of the remote computer you are trying to send a message to, in the format of 172.13.16.54."
         Height          =   645
         Left            =   225
         TabIndex        =   7
         Top             =   1935
         Width           =   3075
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Commands: "
         Height          =   195
         Left            =   90
         TabIndex        =   6
         Top             =   1710
         Width           =   870
      End
      Begin VB.Label Label5 
         Caption         =   "The obvious purpose of this program is to be able to send messages over a local area network or over the internet."
         Height          =   645
         Left            =   90
         TabIndex        =   5
         Top             =   1080
         Width           =   3480
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "essenger"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   1890
         TabIndex        =   4
         Top             =   585
         Width           =   960
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "M"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   26.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   555
         Left            =   1440
         TabIndex        =   3
         Top             =   360
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "et"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   1170
         TabIndex        =   2
         Top             =   450
         Width           =   195
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "N"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   26.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   555
         Left            =   810
         TabIndex        =   1
         Top             =   225
         Width           =   420
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   135
         Picture         =   "help.frx":0000
         Top             =   225
         Width           =   480
      End
   End
End
Attribute VB_Name = "help"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCLOSE_Click()
Me.Hide
End Sub
