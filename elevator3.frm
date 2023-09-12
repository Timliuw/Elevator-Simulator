VERSION 5.00
Begin VB.Form Elevator 
   Caption         =   "Elevator Simulator"
   ClientHeight    =   8505
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10485
   LinkTopic       =   "Form1"
   ScaleHeight     =   8505
   ScaleWidth      =   10485
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.CommandButton btViewLog 
      Caption         =   "Logbook"
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
      Left            =   7680
      TabIndex        =   48
      Top             =   7320
      Width           =   1575
   End
   Begin VB.Timer high_low 
      Interval        =   100
      Left            =   9840
      Top             =   3120
   End
   Begin VB.Timer idle 
      Interval        =   1000
      Left            =   9840
      Top             =   2520
   End
   Begin VB.Timer bDoorClosing_timer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   9840
      Top             =   1920
   End
   Begin VB.Timer dw_timer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   9840
      Top             =   1320
   End
   Begin VB.Timer up_timer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   9840
      Top             =   720
   End
   Begin VB.CommandButton btDW 
      Caption         =   "DOWN"
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
      Index           =   10
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   700
      Width           =   1095
   End
   Begin VB.CommandButton btDW 
      Caption         =   "DOWN"
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
      Index           =   9
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   1400
      Width           =   1095
   End
   Begin VB.CommandButton btDW 
      Caption         =   "DOWN"
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
      Index           =   8
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   2100
      Width           =   1095
   End
   Begin VB.CommandButton btDW 
      Caption         =   "DOWN"
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
      Index           =   7
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   2800
      Width           =   1095
   End
   Begin VB.CommandButton btDW 
      Caption         =   "DOWN"
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
      Index           =   6
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   3500
      Width           =   1095
   End
   Begin VB.CommandButton btDW 
      Caption         =   "DOWN"
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
      Index           =   5
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton btDW 
      Caption         =   "DOWN"
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
      Index           =   4
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   4900
      Width           =   1095
   End
   Begin VB.CommandButton btDW 
      Caption         =   "DOWN"
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
      Index           =   3
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   5600
      Width           =   1095
   End
   Begin VB.PictureBox ele2 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   500
      Left            =   5880
      ScaleHeight     =   500
      ScaleMode       =   0  'User
      ScaleWidth      =   500
      TabIndex        =   28
      Top             =   7000
      Width           =   500
   End
   Begin VB.PictureBox ele1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   500
      Left            =   5400
      ScaleHeight     =   500
      ScaleMode       =   0  'User
      ScaleWidth      =   500
      TabIndex        =   27
      Top             =   7000
      Width           =   500
   End
   Begin VB.CommandButton btUP 
      Caption         =   "UP"
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
      Index           =   9
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   1400
      Width           =   1095
   End
   Begin VB.CommandButton btUP 
      Caption         =   "UP"
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
      Index           =   8
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   2100
      Width           =   1095
   End
   Begin VB.CommandButton btUP 
      Caption         =   "UP"
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
      Index           =   7
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   2800
      Width           =   1095
   End
   Begin VB.CommandButton btUP 
      Caption         =   "UP"
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
      Index           =   6
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   3500
      Width           =   1095
   End
   Begin VB.CommandButton btUP 
      Caption         =   "UP"
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
      Index           =   5
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton btUP 
      Caption         =   "UP"
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
      Index           =   4
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   4900
      Width           =   1095
   End
   Begin VB.CommandButton btUP 
      Caption         =   "UP"
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
      Index           =   3
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   5600
      Width           =   1095
   End
   Begin VB.CommandButton btUP 
      Caption         =   "UP"
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
      Index           =   2
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   6300
      Width           =   1095
   End
   Begin VB.CommandButton btUP 
      Caption         =   "UP"
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
      Index           =   1
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   7000
      Width           =   1095
   End
   Begin VB.Timer main_timer 
      Interval        =   1
      Left            =   9840
      Top             =   120
   End
   Begin VB.CommandButton btHOLD 
      Caption         =   "HOLD"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6360
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Height          =   4695
      Left            =   7200
      TabIndex        =   13
      Top             =   1560
      Width           =   2535
      Begin VB.CommandButton num_in 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   2
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   3840
         Width           =   615
      End
      Begin VB.CommandButton num_in 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   3
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   3000
         Width           =   615
      End
      Begin VB.CommandButton num_in 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   4
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   3000
         Width           =   615
      End
      Begin VB.CommandButton num_in 
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   5
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   2160
         Width           =   615
      End
      Begin VB.CommandButton num_in 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   6
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   2160
         Width           =   615
      End
      Begin VB.CommandButton num_in 
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   7
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   1320
         Width           =   615
      End
      Begin VB.CommandButton num_in 
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   8
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   1320
         Width           =   615
      End
      Begin VB.CommandButton num_in 
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   9
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   480
         Width           =   615
      End
      Begin VB.CommandButton num_in 
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   10
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   480
         Width           =   615
      End
      Begin VB.CommandButton num_in 
         Caption         =   "L"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   1
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   3840
         Width           =   615
      End
   End
   Begin VB.Frame Ele_frm 
      Height          =   7575
      Left            =   4920
      TabIndex        =   12
      Top             =   360
      Width           =   1935
   End
   Begin VB.CommandButton btDW 
      Caption         =   "DOWN"
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
      Index           =   2
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6300
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   7575
      Left            =   960
      TabIndex        =   10
      Top             =   360
      Width           =   3495
   End
   Begin VB.Label l4 
      Caption         =   "Label4"
      Height          =   495
      Left            =   8640
      TabIndex        =   38
      Top             =   240
      Width           =   975
   End
   Begin VB.Label l2 
      Caption         =   "Label2"
      Height          =   495
      Left            =   7200
      TabIndex        =   37
      Top             =   240
      Width           =   615
   End
   Begin VB.Label flour 
      Alignment       =   2  'Center
      BackColor       =   &H8000000B&
      Caption         =   "L"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   7440
      TabIndex        =   17
      Top             =   840
      Width           =   735
   End
   Begin VB.Label istate_text 
      Alignment       =   2  'Center
      BackColor       =   &H8000000B&
      Caption         =   "IDLE"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8280
      TabIndex        =   16
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000A&
      Caption         =   " L"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   9
      Left            =   240
      TabIndex        =   9
      Top             =   7000
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000A&
      Caption         =   " 2"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   8
      Left            =   240
      TabIndex        =   8
      Top             =   6300
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000A&
      Caption         =   " 3"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   7
      Left            =   240
      TabIndex        =   7
      Top             =   5600
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000A&
      Caption         =   " 4"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   240
      TabIndex        =   6
      Top             =   4900
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000A&
      Caption         =   " 5"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   240
      TabIndex        =   5
      Top             =   4200
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000A&
      Caption         =   " 6"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   240
      TabIndex        =   4
      Top             =   3500
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000A&
      Caption         =   " 7"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   240
      TabIndex        =   3
      Top             =   2800
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000A&
      Caption         =   " 8"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   240
      TabIndex        =   2
      Top             =   2100
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000A&
      Caption         =   " 9"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   1400
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000A&
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   700
      Width           =   495
   End
End
Attribute VB_Name = "Elevator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bDestFloor(10) As Boolean
Dim bup(10) As Boolean
Dim iState As Integer '0:idle, 1:up, 2:down
Dim highestFloor As Single
Dim lowestFloor As Single
Dim curFloor As Single
Dim bDoorClosing As Boolean
Dim tDoorAction As Integer
Dim iTimeIdel As Integer 'idle time
Dim floor_to_go(10) As Boolean
Dim bdw(10) As Boolean
Dim bHold As Boolean
Dim spare As Boolean
Dim fPath As String 'define the path of Event Log


Private Sub btHold_Click()
    bHold = Not (bHold)
    btHOLD.BackColor = IIf(bHold, &HFF&, &H8000000F) 'if it is pressed, light on
    If bHold Then
        SaveEvent Date & "  " & Time & vbTab & "HOLD button is pressed."
    Else
        SaveEvent Date & "  " & Time & vbTab & "HOLD button is released."
    End If

End Sub

Private Sub btViewLog_Click()
    EventLog.Show
End Sub

Private Sub Form_Load()
Dim i As Integer

    iState = 0
    highestFloor = 0
    lowestFloor = 100
    bHold = False
    spare = True
    curFloor = 1
    bDoorClosing = True ' door is closed
    tDoorAction = 12

    For i = 1 To 9
        bup(i) = False
        bdw(i + 1) = False
        bDestFloor(i) = False
    Next i
    bDestFloor(10) = False

    fPath = App.Path & "\EventLog.txt"
    SaveEvent vbCrLf & Date & "  " & Time & vbTab & "Elevator starts service."
End Sub

Private Sub btDW_Click(Index As Integer)
    btDW(Index).BackColor = &HFF&
    SaveEvent Date & "  " & Time & vbTab & "DOWN button is pressed on F" & Index
    If curFloor <> Index Then
        bdw(Index) = True
        'Light On
        If iState = 0 Then ' the lift is in ideling iState
            If curFloor < Index Then 'curFloorrent position is lower than the floor with up button clicked,then UP
                iState = 1
                highestFloor = Index
            Else
                iState = 2
            End If
        End If
        If iState <> 0 Then ' the lift is in travelling iState
            If Index < lowestFloor Then
                lowestFloor = Index
            ElseIf Index > highestFloor And curFloor < Index Then
                highestFloor = Index
            End If
        End If
    Else
        If bDoorClosing = True And ele1.Width = 500 Then
            bDoorClosing_timer.Enabled = True
            tDoorAction = 12
            tDoorAction = 12
            If Int((10 * Rnd) + 1) > 1 Then
                new_bdestfloor = 1
            Else
                new_bdestfloor = Int(((curFloor - 2) * Rnd) + 2)
'                new_bDestFloor = Int(((curFloor - 1) * Rnd) + 2)
            End If
            bDestFloor(new_bdestfloor) = True
            SaveEvent Date & "  " & Time & vbTab & "The destination floor is " & new_bdestfloor
            
            num_in(new_bdestfloor).BackColor = &HFF&
            If lowestFloor > new_bdestfloor Then
                lowestFloor = new_bdestfloor
            End If
        End If
        btDW(Index).BackColor = &H8000000F
    End If

End Sub


Private Sub btUP_Click(Index As Integer)
    btUP(Index).BackColor = &HFF&
    SaveEvent Date & "  " & Time & vbTab & "UP button is pressed on F" & Index
    If curFloor <> Index Then 'if the lift is not on the same floor
        bup(Index) = True
        'Light On
        If iState = 0 Then 'if the lift is on waiting iState
            If curFloor < Index Then
                iState = 1
            End If
            If curFloor > Index Then
                iState = 2
            End If
        End If
        If iState <> 0 Then
'            Findhigh_low
            If Index > highestFloor And curFloor < Index Then
                highestFloor = Index
            ElseIf Index < lowestFloor Then
                lowestFloor = Index
            End If

        End If
    Else ' the lift is on the same floor
        If bDoorClosing = True And ele1.Width = 500 Then
            tDoorAction = 12
            bDoorClosing_timer.Enabled = True
            new_bdestfloor = Int(((11 - curFloor) * Rnd) + curFloor) 'generate a num from curFloor+1 to 10
            If new_bdestfloor > 10 Then
                new_bdestfloor = 10
            End If
'            new_bdestfloor = Int(((11 - curFloor) * Rnd) + curFloor) 'generate a num from curFloor+1 to 10
            bDestFloor(new_bdestfloor) = True
            
            SaveEvent Date & "  " & Time & vbTab & "The destination floor is " & new_bdestfloor
            
            num_in(new_bdestfloor).BackColor = &HFF&
            If highestFloor < new_bdestfloor Then
                highestFloor = new_bdestfloor
            End If
            
            num_in(new_bdestfloor).BackColor = &HFF&
        End If

        btUP(Index).BackColor = &H8000000F
    End If
    
    
End Sub


Private Sub high_low_Timer()
    Dim a1 As Integer, a2 As Integer, i As Integer

    a1 = 0
    a2 = 100
    For i = 1 To 10
        If bup(i) = True Or bdw(i) = True Or bDestFloor(i) = True Then
            floor_to_go(i) = True
        Else
            floor_to_go(i) = False
        End If
    Next i
    For i = 1 To 10
        If floor_to_go(i) = True Then
            If i > a1 Then
                a1 = i
            End If
            If i < a2 Then
                a2 = i
            End If
        End If
    Next i
    highestFloor = a1
    If spare = True Then
        lowestFloor = a2
    Else
        lowestFloor = 1
    End If
    
End Sub

Private Sub idle_Timer()
    If iState = 0 And curFloor <> 1 Then
        iTimeIdel = iTimeIdel + 1
        
    Else
        iTimeIdel = 0
    End If
    If curFloor = 1 Then
        spare = True
    End If
    For i = 1 To 10
        If curFloor = Int(curFloor) And curFloor = num_in(i) Then
            num_in(i).BackColor = &H8000000F
        End If
    Next i
End Sub

Private Sub main_timer_Timer()
Dim bClear As Boolean ' no button selected
Dim i As Integer

    If iState = 1 And bDoorClosing = True And curFloor < highestFloor Then 'if the lift is UP and the bDoorClosing is closed and not reach the bDestFloorination
        up_timer.Enabled = True
    Else
        up_timer.Enabled = False
    End If
    If iState = 2 And bDoorClosing = True And curFloor > lowestFloor Then 'if the lift is Down and the bDoorClosing is closed and not reach the bDestFloorination
        dw_timer.Enabled = True
    Else
        dw_timer.Enabled = False
    End If

    '---------up reaches flour---------
    If iState = 1 And Int(curFloor) = curFloor Then
        If bup(curFloor) = True Then
            bup(curFloor) = False
            btUP(curFloor).BackColor = &H8000000F
            bDoorClosing = False
            bDoorClosing_timer.Enabled = True
            tDoorAction = 12
            new_bdestfloor = Int(((11 - curFloor) * Rnd) + curFloor) 'generate a num from curFloor+1 to 10
'            new_bdestfloor = Int(((11 - curFloor) * Rnd) + curFloor) 'generate a num from curFloor+1 to 10
            bDestFloor(new_bdestfloor) = True
            SaveEvent Date & "  " & Time & vbTab & "The destination floor is " & new_bdestfloor
            
            num_in(new_bdestfloor).BackColor = &HFF&
            If highestFloor < new_bdestfloor Then
                highestFloor = new_bdestfloor
            End If
            'up bt reach highestFloor no need to change direction
        ElseIf bDestFloor(curFloor) = True Then
            bDestFloor(curFloor) = False
            num_in(curFloor).BackColor = &H8000000F
            bDoorClosing = False
            bDoorClosing_timer.Enabled = True
            tDoorAction = 12
        ElseIf curFloor = highestFloor Then
            highestFloor = 0
            bDoorClosing = False
            bDoorClosing_timer.Enabled = True
            tDoorAction = 12
            If bdw(curFloor) = True Then
                If Int((10 * Rnd) + 1) > 1 Then
                    new_bdestfloor = 1
                Else
                    new_bdestfloor = Int(((curFloor - 2) * Rnd) + 2)
'                    new_bDestFloor = Int(((curFloor - 1) * Rnd) + 2)
                End If
                bDestFloor(new_bdestfloor) = True
                SaveEvent Date & "  " & Time & vbTab & "The destination floor is " & new_bdestfloor
                
                num_in(new_bdestfloor).BackColor = &HFF&
                If lowestFloor > new_bdestfloor Then
                    lowestFloor = new_bdestfloor
                End If
                bdw(curFloor) = False
                btDW(curFloor).BackColor = &H8000000F
            End If
            If lowestFloor <> 100 Then 'change direction
                iState = 2
            Else
                iState = 0
            End If
        
        End If
    End If
    
    '-------down reaches flour---------
    If iState = 2 And Int(curFloor) = curFloor Then
        If bdw(curFloor) = True Then
            bdw(curFloor) = False
            btDW(curFloor).BackColor = &H8000000F
            bDoorClosing = False
            bDoorClosing_timer.Enabled = True
            tDoorAction = 12
            If Int((10 * Rnd) + 1) > 1 Then
                    new_bdestfloor = 1
                Else
                    new_bdestfloor = Int(((curFloor - 2) * Rnd) + 2)
'                    new_bDestFloor = Int(((curFloor - 1) * Rnd) + 2)
            End If
            bDestFloor(new_bdestfloor) = True
            SaveEvent Date & "  " & Time & vbTab & "The destination floor is " & new_bdestfloor
            
            num_in(new_bdestfloor).BackColor = &HFF&
            If lowestFloor > new_bdestfloor Then
                lowestFloor = new_bdestfloor
            End If
        ElseIf bDestFloor(curFloor) = True Then
            bDestFloor(curFloor) = False
            num_in(curFloor).BackColor = &H8000000F
            bDoorClosing = False
            bDoorClosing_timer.Enabled = True
            tDoorAction = 12
        ElseIf curFloor = lowestFloor Then
            lowestFloor = 100
            bDoorClosing = False
            bDoorClosing_timer.Enabled = True
            tDoorAction = 12
            If bup(curFloor) = True Then
                new_bdestfloor = Int(((11 - curFloor) * Rnd) + curFloor) 'generate a num from curFloor+1 to 10
'                new_bdestfloor = Int(((11 - curFloor) * Rnd) + curFloor) 'generate a num from curFloor+1 to 10
                bDestFloor(new_bdestfloor) = True
                SaveEvent Date & "  " & Time & vbTab & "The destination floor is " & new_bdestfloor
                
                num_in(new_bdestfloor).BackColor = &HFF&
                If highestFloor < new_bdestfloor Then
                    highestFloor = new_bdestfloor
                End If
                bup(curFloor) = False
                btUP(curFloor).BackColor = &H8000000F
            End If
            If highestFloor <> 0 Then 'change direction
                iState = 1
            Else
                iState = 0
            End If
        End If
    End If
    If iState = 1 Then
        istate_text.Caption = " ¡ü "
        istate_text.ForeColor = &HFF00&
    ElseIf iState = 2 Then
        istate_text.Caption = " ¡ý "
        istate_text.ForeColor = &HFF&
    ElseIf iState = 0 Then
        istate_text.ForeColor = &H80000012
        istate_text.Caption = "IDLE"
    End If
    If curFloor = Int(curFloor) Then
        flour.Caption = curFloor
    End If
    l2.Caption = highestFloor
    l4.Caption = lowestFloor
    
    bClear = False
    For i = 10 To 2 Step -1
        bClear = bClear Or bup(i) Or bdw(i)
    Next
    idle.Enabled = Not bClear 'back to the ground floor if no button selected
    
    For k = 1 To 10
        If k <> 10 Then
            If btUP(k).BackColor = &HFF& Then
                bup(k) = True
            End If
        End If
        If k <> 1 Then
            If btDW(k).BackColor = &HFF& Then
                bdw(k) = True
            End If
        End If
    Next k
    If iTimeIdel = 60 Then
        lowestFloor = 1
        iState = 2
        iTimeIdel = 0
        spare = False
    ElseIf spare = True And bDoorClosing = True And dw_timer.Enabled = False And up_timer.Enabled = False Then
        iState = 0
    End If
End Sub

Private Sub up_timer_Timer()
    If ele1.Width = 500 Then
        ele1.Top = ele1.Top - 175
        ele2.Top = ele2.Top - 175
        curFloor = curFloor + 0.25
    End If
End Sub
Private Sub dw_timer_Timer()
    If ele1.Width = 500 Then
        ele1.Top = ele1.Top + 175
        ele2.Top = ele2.Top + 175
        curFloor = curFloor - 0.25
    End If
End Sub

Private Sub bDoorClosing_timer_Timer()
    If tDoorAction <= 4 And tDoorAction > 0 And bHold = False Then
        If tDoorAction = 4 Then SaveEvent Date & "  " & Time & vbTab & "Door starts closing on F" & curFloor
        If tDoorAction = 1 Then SaveEvent Date & "  " & Time & vbTab & "Elevator is fully closed"
        ele1.Width = ele1.Width + 100
        ele2.Width = ele2.Width + 100
        ele2.Left = ele2.Left - 100
        tDoorAction = tDoorAction - 1
    End If
    If tDoorAction = 0 Then
        bDoorClosing_timer.Enabled = False
        bDoorClosing = True
    End If
    If tDoorAction <= 12 And tDoorAction >= 9 And ele1.Width >= 101 Then
        If tDoorAction = 12 Then SaveEvent Date & "  " & Time & vbTab & "Elevator stops on F" & curFloor
        If tDoorAction = 9 Then SaveEvent Date & "  " & Time & vbTab & "Door is fully opened on F" & curFloor
        ele1.Width = ele1.Width - 100
        ele2.Width = ele2.Width - 100
        ele2.Left = ele2.Left + 100
        tDoorAction = tDoorAction - 1
    End If
    If tDoorAction >= 5 And tDoorAction <= 8 Then
        tDoorAction = tDoorAction - 1
    End If



    If iState = 0 And curFloor < highestFloor And ele1.Width = 500 Then
        iState = 1
    End If
    If iState = 0 And curFloor > lowestFloor And ele1.Width = 500 Then
        iState = 2
    End If

End Sub
                


Function SaveEvent(outString As String)
    Dim fN As Integer
    fN = FreeFile 'get an available fileNo.
    Open fPath For Append As #fN
        Print #fN, outString
    Close #fN
End Function

