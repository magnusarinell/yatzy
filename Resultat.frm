VERSION 5.00
Begin VB.Form frmResultat 
   BackColor       =   &H00400000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Resultat"
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
   Begin VB.CommandButton cmdHighScore 
      Caption         =   "Highscore"
      Height          =   375
      Left            =   1440
      TabIndex        =   11
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton CmdSlut 
      Caption         =   "Game Over"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label lblSummaResultatÖverskrift 
      BackStyle       =   0  'Transparent
      Caption         =   "Po�ng"
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
      Left            =   1680
      TabIndex        =   10
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblSummaResultat 
      BackStyle       =   0  'Transparent
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
      Left            =   1800
      TabIndex        =   9
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label lblSummaResultat 
      BackStyle       =   0  'Transparent
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
      Left            =   1800
      TabIndex        =   8
      Top             =   960
      Width           =   495
   End
   Begin VB.Label lblSummaResultat 
      BackStyle       =   0  'Transparent
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
      Left            =   1800
      TabIndex        =   7
      Top             =   720
      Width           =   495
   End
   Begin VB.Label lblSummaResultat 
      BackStyle       =   0  'Transparent
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
      Index           =   0
      Left            =   1800
      TabIndex        =   6
      Top             =   480
      Width           =   495
   End
   Begin VB.Label lblNamnResultatÖverskrift 
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
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblNamnResultat 
      BackStyle       =   0  'Transparent
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
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label lblNamnResultat 
      BackStyle       =   0  'Transparent
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
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   975
   End
   Begin VB.Label lblNamnResultat 
      BackStyle       =   0  'Transparent
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
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   975
   End
   Begin VB.Label lblNamnResultat 
      BackStyle       =   0  'Transparent
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
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   975
   End
End
Attribute VB_Name = "frmResultat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Temp As Integer
Dim Temp2 As String
Dim PostNamn(1 To 10) As String
Dim PostPo�ng(1 To 10) As Integer
Dim MaxPo�ng As Integer

Private Sub cmdHighScore_Click()
  Highscore.Show
End Sub

Private Sub CmdSlut_Click()
  End
End Sub

Private Sub Form_Load()
  iFilnr = FreeFile
  
  For i = 0 To AntalSpelare
    lblSummaResultat(i).Caption = IntResultat(i)
    lblNamnResultat(i).Caption = Spelare(i)
  Next i

  Open "Highscore.dat" For Random As iFilnr Len = Len(aktHighScore)
    For i = 10 To 1 Step -1
      Get #iFilnr, i, aktHighScore
      PostPoäng(i) = Val(aktHighScore.Po�ng)
      PostNamn(i) = aktHighScore.Namn
    Next i
  Close #iFilnr
  
  MaxPo�ng = PostPoäng(10)
  
  For i = 0 To AntalSpelare
    PostNamn(i + 1) = Spelare(i)
    PostPo�ng(i + 1) = IntResultat(i)
  Next i
  
  For j = 1 To 9
    For i = 1 To 9
      If PostPoäng(i) > PostPoäng(i + 1) Then
        Temp = PostPoäng(i + 1)
        PostPoäng(i + 1) = PostPoäng(i)
        PostPoäng(i) = Temp
        Temp2 = PostNamn(i + 1)
        PostNamn(i + 1) = PostNamn(i)
        PostNamn(i) = Temp2
      End If
    Next i
  Next j
  
  Open "highscore.dat" For Random As iFilnr Len = Len(aktHighScore)
    For i = 10 To 1 Step -1
      Spara
    Next i
  Close #iFilnr
      
  If MaxPoäng < PostPoäng(10) Then
    Highscore.Show
    frmResultat.Hide
  End If
End Sub

Public Sub Spara()
  aktHighScore.Poäng = PostPoäng(i)
  aktHighScore.Namn = PostNamn(i)
  Put #iFilnr, i, aktHighScore
End Sub

