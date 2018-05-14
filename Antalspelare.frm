VERSION 5.00
Begin VB.Form frmAntalspelare 
   BackColor       =   &H00400000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Antal Spelare"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3960
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MousePointer    =   99  'Custom
   Moveable        =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   3960
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkDatorSpelare 
      BackColor       =   &H00400000&
      Caption         =   "Datorspelare"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1440
      MaskColor       =   &H00400000&
      TabIndex        =   13
      Top             =   3480
      Width           =   1335
   End
   Begin VB.TextBox TxtSpelarnamn 
      Height          =   285
      Index           =   3
      Left            =   1920
      MaxLength       =   9
      TabIndex        =   11
      Text            =   "Spelare 4"
      Top             =   2880
      Width           =   1575
   End
   Begin VB.TextBox TxtSpelarnamn 
      Height          =   285
      Index           =   2
      Left            =   1920
      MaxLength       =   9
      TabIndex        =   10
      Text            =   "Spelare 3"
      Top             =   2520
      Width           =   1575
   End
   Begin VB.TextBox TxtSpelarnamn 
      Height          =   285
      Index           =   1
      Left            =   1920
      MaxLength       =   9
      TabIndex        =   9
      Text            =   "Spelare 2"
      Top             =   2160
      Width           =   1575
   End
   Begin VB.TextBox TxtSpelarnamn 
      Height          =   285
      Index           =   0
      Left            =   1920
      MaxLength       =   9
      TabIndex        =   8
      Text            =   "Spelare 1"
      Top             =   1800
      Width           =   1575
   End
   Begin VB.CommandButton cmdAvbryt 
      Caption         =   "&Avbryt"
      Height          =   495
      Left            =   2040
      TabIndex        =   5
      Top             =   4080
      Width           =   1695
   End
   Begin VB.CommandButton CmdStarta 
      Caption         =   "&Starta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   4080
      Width           =   1695
   End
   Begin VB.OptionButton optAntalSpelare 
      BackColor       =   &H00400000&
      Caption         =   "4 Spelare"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   1
      Top             =   2880
      Width           =   1335
   End
   Begin VB.OptionButton optAntalSpelare 
      BackColor       =   &H00400000&
      Caption         =   "3 Spelare"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   3
      Top             =   2520
      Width           =   1335
   End
   Begin VB.OptionButton optAntalSpelare 
      BackColor       =   &H00400000&
      Caption         =   "2 Spelare"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   2
      Top             =   2160
      Width           =   1335
   End
   Begin VB.OptionButton optAntalSpelare 
      BackColor       =   &H00400000&
      Caption         =   "1 Spelare"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   1800
      Value           =   -1  'True
      Width           =   1335
   End
   Begin VB.Label lblSpelarnamn 
      BackColor       =   &H00400000&
      Caption         =   "Namn:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1920
      TabIndex        =   12
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label lblavMagnusBroberg 
      BackStyle       =   0  'Transparent
      Caption         =   "Av Magnus Broberg"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1560
      TabIndex        =   7
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00400000&
      Caption         =   "Yatzy!"
      BeginProperty Font 
         Name            =   "News Gothic MT"
         Size            =   33.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   840
      TabIndex        =   6
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "frmAntalspelare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub chkDatorSpelare_Click()
  If chkDatorSpelare.Value = 1 Then
    optAntalSpelare(1).Value = True
    'TxtSpelarnamn(1).Text = "Mastah"
    optAntalSpelare(0).Enabled = False
    optAntalSpelare(2).Enabled = False
    optAntalSpelare(3).Enabled = False
  Else
    'TxtSpelarnamn(1).Text = "Spelare 2"
    optAntalSpelare(0).Enabled = True
    optAntalSpelare(2).Enabled = True
    optAntalSpelare(3).Enabled = True
  End If
End Sub

Private Sub CmdStarta_Click()
  For i = 0 To 3
    If optAntalSpelare(i).Value = True Then
      AntalSpelare = i
    End If
  Next i

  For i = 0 To AntalSpelare
    Spelare(i) = TxtSpelarnamn(i).Text
  Next i
  
  If chkDatorSpelare.Value = 1 Then
    BoolDatorSpelare = True
  End If
  
  Unload Me
  frmYatzy.Show
End Sub

Private Sub CmdAvbryt_Click()
  End
End Sub

Private Sub Form_Load()
  For i = 1 To 3
    TxtSpelarnamn(i).Enabled = False
  Next i
End Sub

Private Sub optAntalSpelare_Click(index As Integer)
  If index = 0 And chkDatorSpelare.Value = 1 Then
    optAntalSpelare(1).Value = True
    
  Else
    For i = 1 To index
      TxtSpelarnamn(i).Enabled = True
    Next i
    For i = (index + 1) To 3
      TxtSpelarnamn(i).Enabled = False
    Next i
  End If
End Sub
