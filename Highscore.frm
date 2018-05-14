VERSION 5.00
Begin VB.Form Highscore 
   BackColor       =   &H00400000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Highscore"
   ClientHeight    =   2580
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   2790
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   2790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label lblPoängHighÖverskrift 
      BackStyle       =   0  'Transparent
      Caption         =   "Poäng"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1560
      TabIndex        =   14
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblNamnHighÖverskrift 
      BackStyle       =   0  'Transparent
      Caption         =   "Spelare"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   13
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblHighPoäng 
      Alignment       =   2  'Center
      BackColor       =   &H00400000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   1680
      TabIndex        =   12
      Top             =   480
      Width           =   375
   End
   Begin VB.Label lblHighPoäng 
      Alignment       =   2  'Center
      BackColor       =   &H00400000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   1680
      TabIndex        =   11
      Top             =   720
      Width           =   375
   End
   Begin VB.Label lblHighPoäng 
      Alignment       =   2  'Center
      BackColor       =   &H00400000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   1680
      TabIndex        =   10
      Top             =   960
      Width           =   375
   End
   Begin VB.Label lblHighPoäng 
      Alignment       =   2  'Center
      BackColor       =   &H00400000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   1680
      TabIndex        =   9
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label lblHighPoäng 
      Alignment       =   2  'Center
      BackColor       =   &H00400000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   1680
      TabIndex        =   8
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label lblHighNamn 
      BackColor       =   &H00400000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   480
      TabIndex        =   7
      Top             =   480
      Width           =   975
   End
   Begin VB.Label lblHighNamn 
      BackColor       =   &H00400000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   480
      TabIndex        =   6
      Top             =   720
      Width           =   975
   End
   Begin VB.Label lblHighNamn 
      BackColor       =   &H00400000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   480
      TabIndex        =   5
      Top             =   960
      Width           =   975
   End
   Begin VB.Label lblHighNamn 
      BackColor       =   &H00400000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   480
      TabIndex        =   4
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label lblHighNamn 
      BackColor       =   &H00400000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   480
      TabIndex        =   3
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label lblHighPoäng 
      Alignment       =   2  'Center
      BackColor       =   &H00400000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   1680
      TabIndex        =   1
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label lblHighNamn 
      BackColor       =   &H00400000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   0
      Top             =   1680
      Width           =   975
   End
End
Attribute VB_Name = "Highscore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdOk_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  iFilnr = FreeFile
  
  Open "Highscore.dat" For Random As iFilnr Len = Len(aktHighScore)
    For i = 10 To 5 Step -1
      Get #iFilnr, i, aktHighScore
      lblHighNamn(i - 4).Caption = aktHighScore.Namn
      lblHighPoäng(i - 4).Caption = aktHighScore.Poäng
    Next i
  Close #iFilnr
End Sub
