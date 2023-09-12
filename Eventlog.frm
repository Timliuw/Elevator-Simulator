VERSION 5.00
Begin VB.Form EventLog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Event Log"
   ClientHeight    =   5760
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7470
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   7470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.TextBox txtLog 
      Height          =   5415
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   120
      Width           =   5775
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   10.5
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6120
      TabIndex        =   0
      Top             =   4920
      Width           =   1215
   End
End
Attribute VB_Name = "EventLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Load()
    Dim Str1 As String
    Dim Str2 As String
    Dim StrFileName As String

    StrFileName = App.Path & "\EventLog.txt" '
    Open StrFileName For Input As #1 'open file
    Do Until EOF(1) '
        Line Input #1, Str1 '
        Str2 = Str2 & Str1 & vbCrLf '
    Loop
    txtLog.Text = Str2
    Close #1
End Sub

Private Sub OKButton_Click()
    Unload Me
End Sub
