VERSION 5.00
Begin VB.Form frmYatzy 
   BackColor       =   &H00400000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Yatzy"
   ClientHeight    =   6840
   ClientLeft      =   150
   ClientTop       =   420
   ClientWidth     =   7515
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00400000&
   Icon            =   "Yatzy.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   7515
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   71
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   102
      Top             =   6240
      Width           =   375
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   70
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   101
      Top             =   5880
      Width           =   375
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   69
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   100
      Top             =   5520
      Width           =   375
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   68
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   99
      Top             =   5160
      Width           =   375
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   67
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   98
      Top             =   4800
      Width           =   375
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   66
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   97
      Top             =   4440
      Width           =   375
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   65
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   96
      Top             =   4080
      Width           =   375
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   64
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   95
      Top             =   3720
      Width           =   375
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   63
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   94
      Top             =   3360
      Width           =   375
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   62
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   93
      Top             =   3000
      Width           =   375
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   61
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   92
      Top             =   5520
      Width           =   375
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   60
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   91
      Top             =   5160
      Width           =   375
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   59
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   90
      Top             =   4800
      Width           =   375
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   58
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   89
      Top             =   4440
      Width           =   375
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   57
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   88
      Top             =   4080
      Width           =   375
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   56
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   87
      Top             =   3720
      Width           =   375
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   55
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   86
      Top             =   3360
      Width           =   375
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   54
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   85
      Top             =   3000
      Width           =   375
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   53
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   84
      Top             =   6240
      Width           =   375
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   52
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   83
      Top             =   5880
      Width           =   375
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   51
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   82
      Top             =   5520
      Width           =   375
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   50
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   81
      Top             =   5160
      Width           =   375
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   49
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   80
      Top             =   4800
      Width           =   375
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   48
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   79
      Top             =   4440
      Width           =   375
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   47
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   78
      Top             =   4080
      Width           =   375
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   46
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   77
      Top             =   3720
      Width           =   375
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   45
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   76
      Top             =   3360
      Width           =   375
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   44
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   75
      Top             =   3000
      Width           =   375
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   43
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   74
      Top             =   5520
      Width           =   375
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   42
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   73
      Top             =   5160
      Width           =   375
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   41
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   72
      Top             =   4800
      Width           =   375
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   40
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   71
      Top             =   4440
      Width           =   375
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   39
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   70
      Top             =   4080
      Width           =   375
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   38
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   69
      Top             =   3720
      Width           =   375
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   37
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   68
      Top             =   3360
      Width           =   375
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   36
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   67
      Top             =   3000
      Width           =   375
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   35
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   66
      Top             =   6240
      Width           =   375
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   34
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   65
      Top             =   5880
      Width           =   375
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   33
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   64
      Top             =   5520
      Width           =   375
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   32
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   63
      Top             =   5160
      Width           =   375
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   31
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   62
      Top             =   4800
      Width           =   375
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   30
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   61
      Top             =   4440
      Width           =   375
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   29
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   60
      Top             =   4080
      Width           =   375
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   28
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   59
      Top             =   3720
      Width           =   375
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   27
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   58
      Top             =   3360
      Width           =   375
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   26
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   57
      Top             =   3000
      Width           =   375
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   25
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   56
      Top             =   5520
      Width           =   375
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   24
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   55
      Top             =   5160
      Width           =   375
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   23
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   54
      Top             =   4800
      Width           =   375
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   22
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   53
      Top             =   4440
      Width           =   375
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   21
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   52
      Top             =   4080
      Width           =   375
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   20
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   51
      Top             =   3720
      Width           =   375
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   19
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   50
      Top             =   3360
      Width           =   375
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   18
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   49
      Top             =   3000
      Width           =   375
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   17
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   48
      Top             =   6240
      Width           =   375
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   16
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   47
      Top             =   5880
      Width           =   375
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   15
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   46
      Top             =   5520
      Width           =   375
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   14
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   45
      Top             =   5160
      Width           =   375
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   13
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   44
      Top             =   4800
      Width           =   375
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   12
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   43
      Top             =   4440
      Width           =   375
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   11
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   42
      Top             =   4080
      Width           =   375
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   10
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   41
      Top             =   3720
      Width           =   375
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   9
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   40
      Top             =   3360
      Width           =   375
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   8
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   39
      Top             =   3000
      Width           =   375
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   7
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   34
      Top             =   5520
      Width           =   375
   End
   Begin VB.CommandButton cmdval 
      Caption         =   "YATZY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   -1  'True
      EndProperty
      Height          =   255
      Index           =   14
      Left            =   3120
      TabIndex        =   32
      Top             =   5880
      Width           =   735
   End
   Begin VB.CommandButton cmdval 
      Caption         =   "Chans"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   13
      Left            =   3120
      TabIndex        =   31
      Top             =   5520
      Width           =   735
   End
   Begin VB.CommandButton cmdval 
      Caption         =   "4tal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   3120
      TabIndex        =   30
      Top             =   4080
      Width           =   735
   End
   Begin VB.CommandButton cmdSortera 
      Caption         =   "&Sortera"
      BeginProperty Font 
         Name            =   "News Gothic MT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1680
      TabIndex        =   29
      Top             =   1680
      Width           =   1265
   End
   Begin VB.CommandButton cmdval 
      Caption         =   "Stor S"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   3120
      TabIndex        =   28
      Top             =   4800
      Width           =   735
   End
   Begin VB.CommandButton cmdval 
      Caption         =   "2 Par"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   3120
      TabIndex        =   27
      Top             =   3360
      Width           =   735
   End
   Begin VB.CommandButton cmdval 
      Caption         =   "1 Par"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   3120
      TabIndex        =   26
      Top             =   3000
      Width           =   735
   End
   Begin VB.CommandButton cmdval 
      Caption         =   "Liten S"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   3120
      TabIndex        =   25
      Top             =   4440
      Width           =   735
   End
   Begin VB.CommandButton cmdval 
      Caption         =   "3tal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   3120
      TabIndex        =   24
      Top             =   3720
      Width           =   735
   End
   Begin VB.CommandButton cmdval 
      Caption         =   "K�k"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   12
      Left            =   3120
      TabIndex        =   23
      Top             =   5160
      Width           =   735
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   6
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   5160
      Width           =   375
   End
   Begin VB.CommandButton cmdval 
      Caption         =   "Sexor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   19
      Top             =   4800
      Width           =   735
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   4800
      Width           =   375
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   4440
      Width           =   375
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   4080
      Width           =   375
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   3720
      Width           =   375
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   3360
      Width           =   375
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton cmdval 
      Caption         =   "Femmor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   12
      Top             =   4440
      Width           =   735
   End
   Begin VB.CommandButton cmdval 
      Caption         =   "Fyror"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   11
      Top             =   4080
      Width           =   735
   End
   Begin VB.CommandButton cmdval 
      Caption         =   "Treor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   10
      Top             =   3720
      Width           =   735
   End
   Begin VB.CommandButton cmdval 
      Caption         =   "Tv�or"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   9
      Top             =   3360
      Width           =   735
   End
   Begin VB.CommandButton cmdval 
      Caption         =   "Ettor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   8
      Top             =   3000
      Width           =   735
   End
   Begin VB.CommandButton cmdavsluta 
      Caption         =   "&Avsluta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   6120
      Width           =   1455
   End
   Begin VB.CommandButton cmdkasta 
      Caption         =   "&Kasta"
      BeginProperty Font 
         Name            =   "News Gothic MT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   5
      Top             =   1680
      Width           =   1265
   End
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   5
      Left            =   6000
      ScaleHeight     =   1155
      ScaleWidth      =   1200
      TabIndex        =   4
      Top             =   240
      Width           =   1265
   End
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   4
      Left            =   4560
      ScaleHeight     =   1155
      ScaleWidth      =   1200
      TabIndex        =   3
      Top             =   240
      Width           =   1265
   End
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   3
      Left            =   3120
      ScaleHeight     =   1155
      ScaleWidth      =   1200
      TabIndex        =   2
      Top             =   240
      Width           =   1265
   End
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   2
      Left            =   1680
      ScaleHeight     =   1155
      ScaleWidth      =   1200
      TabIndex        =   1
      Top             =   240
      Width           =   1265
   End
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   1
      Left            =   240
      ScaleHeight     =   1155
      ScaleWidth      =   1200
      TabIndex        =   0
      Top             =   240
      Width           =   1265
   End
   Begin VB.Label LblKastKvar 
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4440
      TabIndex        =   38
      Top             =   2280
      Width           =   135
   End
   Begin VB.Label lblNamn 
      BackStyle       =   0  'Transparent
      Caption         =   "Spelarnamn"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4440
      TabIndex        =   37
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Bonus"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   36
      Top             =   5520
      Width           =   735
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Spelare nr:"
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
      Left            =   3240
      TabIndex        =   35
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Summa"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3000
      TabIndex        =   33
      Top             =   6240
      Width           =   855
   End
   Begin VB.Image ImageT�rningChecked 
      Height          =   1200
      Index           =   6
      Left            =   6120
      Picture         =   "Yatzy.frx":030A
      Top             =   1680
      Width           =   1200
   End
   Begin VB.Image ImageT�rningChecked 
      Height          =   1200
      Index           =   5
      Left            =   6120
      Picture         =   "Yatzy.frx":100C
      Top             =   1680
      Width           =   1200
      Visible         =   0   'False
   End
   Begin VB.Image ImageT�rningChecked 
      Height          =   1200
      Index           =   4
      Left            =   6120
      Picture         =   "Yatzy.frx":1D0E
      Top             =   1680
      Width           =   1200
      Visible         =   0   'False
   End
   Begin VB.Image ImageT�rningChecked 
      Height          =   1200
      Index           =   3
      Left            =   6120
      Picture         =   "Yatzy.frx":2A10
      Top             =   1680
      Width           =   1200
      Visible         =   0   'False
   End
   Begin VB.Image ImageT�rningChecked 
      Height          =   1200
      Index           =   2
      Left            =   6120
      Picture         =   "Yatzy.frx":3712
      Top             =   1680
      Width           =   1200
      Visible         =   0   'False
   End
   Begin VB.Image ImageT�rningChecked 
      Height          =   1200
      Index           =   1
      Left            =   6120
      Picture         =   "Yatzy.frx":4414
      Top             =   1680
      Width           =   1200
      Visible         =   0   'False
   End
   Begin VB.Image ImageT�rning 
      Height          =   1200
      Index           =   6
      Left            =   6120
      Picture         =   "Yatzy.frx":5116
      Top             =   5400
      Width           =   1200
   End
   Begin VB.Image ImageT�rning 
      Height          =   1200
      Index           =   5
      Left            =   6120
      Picture         =   "Yatzy.frx":5E18
      Top             =   5400
      Width           =   1200
      Visible         =   0   'False
   End
   Begin VB.Image ImageT�rning 
      Height          =   1200
      Index           =   4
      Left            =   6120
      Picture         =   "Yatzy.frx":6B1A
      Top             =   5400
      Width           =   1200
      Visible         =   0   'False
   End
   Begin VB.Image ImageT�rning 
      Height          =   1200
      Index           =   3
      Left            =   6120
      Picture         =   "Yatzy.frx":781C
      Top             =   5400
      Width           =   1200
      Visible         =   0   'False
   End
   Begin VB.Image ImageT�rning 
      Height          =   1200
      Index           =   2
      Left            =   6120
      Picture         =   "Yatzy.frx":851E
      Top             =   5400
      Width           =   1200
      Visible         =   0   'False
   End
   Begin VB.Image ImageT�rning 
      Height          =   1200
      Index           =   1
      Left            =   6120
      Picture         =   "Yatzy.frx":9220
      Top             =   5400
      Width           =   1200
      Visible         =   0   'False
   End
   Begin VB.Label lblsumma 
      BackColor       =   &H80000017&
      BackStyle       =   0  'Transparent
      Caption         =   "Summa"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   5160
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8040
      TabIndex        =   20
      Top             =   4080
      Width           =   1695
   End
   Begin VB.Label lblKast_Kvar 
      BackStyle       =   0  'Transparent
      Caption         =   "Kast Kvar"
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
      Left            =   3240
      TabIndex        =   7
      Top             =   2280
      Width           =   975
   End
   Begin VB.Menu mnuarkiv 
      Caption         =   "Arkiv"
      Begin VB.Menu mnuHighscore 
         Caption         =   "&Highscore"
      End
      Begin VB.Menu mnuavsluta 
         Caption         =   "&Avsluta"
      End
   End
   Begin VB.Menu mnuoptions 
      Caption         =   "A&lternativ"
      Begin VB.Menu MnuDatorspelarensHastighet 
         Caption         =   "&Datorspelarens hastighet"
      End
      Begin VB.Menu mnut�rning 
         Caption         =   "&Antal T�rningssnurr"
      End
   End
   Begin VB.Menu mnuhj�lp 
      Caption         =   "Hj�lp"
      Begin VB.Menu MnuOm 
         Caption         =   "&Om Yatzy"
      End
   End
End
Attribute VB_Name = "frmYatzy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Aktuellspelare, KastKvar, HarKastat, AntalSnurr, Apan, PlusTv�, index%
Dim T�rning(1 To 5) As Single, L�sT�rning(1 To 5) As Integer, AntalT�rningar(0 To 6) As Integer
Dim T�rningTemp(1 To 5) As Single
Dim Temp(1 To 5) As Single                     'Diverse tempor�ra variabler
Dim IntVal(0 To 18, 0 To 3) As Integer         'Summan m�ttes i tidigare versioner mha denna tv�dimensionella vektor. Sparad f�r framtida bruk(?).
Dim B�rjan, Slut, Summan%
Dim SummanAv(1 To 2) As Integer
Dim DatorspelareKlar As Boolean
Dim DatorspelarensHastighet As Integer
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) 'ger en pausfunktion n�r den �beropas.

Private Sub cmdavsluta_Click()
    End
End Sub

Private Sub diska(index As Integer)   'Tar bort de spelare som �r fler �n vad man har valt.
    For j = 1 To 18
        txt(((18 * (index)) + j) - 1).Visible = False
    Next j
End Sub

Private Sub cmdkasta_Click()
    KastKvar = (KastKvar - 1)
    HarKastat = 1
    LblKastKvar.Caption = KastKvar
  
    For j = 0 To AntalSnurr
        For i = 1 To 5
            If L�sT�rning(i) = 0 Then
                T�rning(i) = Int(Rnd * 6) + 1                     'Samlar in fem Slumpv�rden.
                Picture1(i).Picture = ImageT�rning(T�rning(i))    'Ritar upp t�rningarna.
            End If
        Next i
    Next j
  
    If KastKvar > 0 Then
        cmdSortera.Enabled = True
    Else
        cmdkasta.Enabled = False
    End If
  
End Sub

Private Sub Rensa() '�terst�ller
  
    cmdSortera.Enabled = False
    KastKvar = 3
    LblKastKvar.Caption = KastKvar
    cmdkasta.Enabled = True
    HarKastat = 0
    PlusTv� = 0
    
    For i = 1 To 5
        Picture1(i).Picture = LoadPicture() 'Laddar in blanka bilder i rutorna
        L�sT�rning(i) = 0
        Temp(i) = 0
    Next i
 
    If Aktuellspelare = AntalSpelare Then  'Skriver namn p� aktuell spelare
        Aktuellspelare = 0
    Else
        Aktuellspelare = Aktuellspelare + 1
    End If
    
    lblNamn.Caption = Spelare(Aktuellspelare)
  
    For i = 0 To 14
        cmdval(i).Enabled = True
    Next i
    For i = (Aktuellspelare * 18) To (Aktuellspelare * 18) + 5       'Otillg�ngligg�r de
        If txt(i) > "" Then                                          'knappar som g�ller f�r
            cmdval(i - 18 * Aktuellspelare).Enabled = False          'de v�rden f�r vilka spelaren
        End If                                                       'redan valt.
    Next i
    For i = (Aktuellspelare * 18) + 8 To (Aktuellspelare * 18) + 16
        If txt(i) > "" Then
            cmdval(i - 18 * Aktuellspelare - 2).Enabled = False
        End If
    Next i
  
    DoEvents         'Beh�vs f�r att v�rden skall s�ttas ut innan datorspelaren b�rjar.
End Sub

Private Sub cmdsortera_Click() 'Funktion som sorterar T�rningarna
    Sortera
    Bilder
End Sub

Private Sub CmdVal_Click(index As Integer)
    If HarKastat = 1 Then
        For j = 6 To 14
            If index = j Then     'N�dv�ndig pga Summa och bonus d� index m�ste f�rkjutas 2 steg d� man kommer h�gre �n sexor.
                PlusTv� = 2
            End If
        Next j
        
        �terst�llb
        Antal
        Sortera
    
        For j = 0 To 5
            If index = j Then
                ettor_till_sexor (index)
            End If
        Next j
        
        Select Case index
            Case 6 'Ett Par
                For i = 1 To 6
                    If AntalT�rningar(i) >= 2 Then
                        IntVal(index, Aktuellspelare) = i * 2
                    End If
                Next i
            Case 7 'Tv� Par
                For i = 1 To 6
                    If AntalT�rningar(i) >= 2 Then
                        Temp(4) = Temp(4) + 1
                    End If
                Next i
             
                If Temp(4) = 2 Then
                    For i = 1 To 6
                        If AntalT�rningar(i) >= 2 Then
                            IntVal(index, Aktuellspelare) = IntVal(index, Aktuellspelare) + 2 * i
                        End If
                    Next i
                End If
            Case 8 'Tretal
                For i = 1 To 6
                    If AntalT�rningar(i) >= 3 Then
                        IntVal(index, Aktuellspelare) = IntVal(index, Aktuellspelare) + 3 * i
                    End If
                Next i
            Case 9 'Fyrtal
                For i = 1 To 6
                    If AntalT�rningar(i) >= 4 Then
                        IntVal(index, Aktuellspelare) = IntVal(index, Aktuellspelare) + 4 * i
                    End If
                Next i
            Case 10 'Liten Straight
                For i = 1 To 5
                    IntVal(index, Aktuellspelare) = AntalT�rningar(i) ^ 5 + IntVal(index, Aktuellspelare)
                Next i
                If IntVal(index, Aktuellspelare) = 5 Then
                    IntVal(index, Aktuellspelare) = IntVal(index, Aktuellspelare) + 10
                Else
                    IntVal(index, Aktuellspelare) = 0
                End If
            Case 11 'Stor Straight
                For i = 2 To 6
                    IntVal(index, Aktuellspelare) = AntalT�rningar(i) ^ 3 + IntVal(index, Aktuellspelare)
                Next i
                If IntVal(index, Aktuellspelare) = 5 Then
                    IntVal(index, Aktuellspelare) = IntVal(index, Aktuellspelare) + 15
                Else
                    IntVal(index, Aktuellspelare) = 0
                End If
            Case 12 'K�k
                If AntalT�rningar(1) ^ 3 + AntalT�rningar(2) ^ 3 + AntalT�rningar(3) ^ 3 + AntalT�rningar(4) ^ 3 + AntalT�rningar(5) ^ 3 + AntalT�rningar(6) ^ 3 = 35 Then
                    For i = 1 To 5
                        IntVal(index, Aktuellspelare) = IntVal(index, Aktuellspelare) + T�rning(i)
                    Next i
                End If
            Case 13 'Chans
                For i = 1 To 5
                    IntVal(index, Aktuellspelare) = IntVal(index, Aktuellspelare) + T�rning(i)
                Next i
            Case 14 'Yatzy
                If T�rning(1) = T�rning(5) Then
                    IntVal(index, Aktuellspelare) = 50
                End If
        End Select
    
        txt(index + (18 * Aktuellspelare) + PlusTv�).Text = IntVal(index, Aktuellspelare)
        Rensa
        Funktion_Summa
        If BoolDatorSpelare = True And Aktuellspelare = 1 Then
            DatorSpelare
        End If
    Else
        Beep
    End If
End Sub

Private Function ettor_till_sexor(index) 'funktion f�r ettor till sexor
  For i = 1 To 5
    If T�rning(i) = (index + 1) Then
      IntVal(index, Aktuellspelare) = IntVal(index, Aktuellspelare) + (index + 1)
    End If
  Next i
End Function

Private Sub Form_Load()
  For i = 1 To 3
    If i > AntalSpelare Then  'Diskar spelare
      diska (i)
    End If
  Next i
  Randomize
  KastKvar = 3
  HarKastat = 0
  AntalSnurr = 50
  lblNamn.Caption = Spelare(0)
  Aktuellspelare = 0
  DatorspelarensHastighet = 250
  
  For i = 1 To 5
    Picture1(i).BackColor = vbWhite
  Next i
  
  If Not BoolDatorSpelare = True Then
    MnuDatorspelarensHastighet.Enabled = False
  End If
End Sub

Private Sub mnuavsluta_Click()
  End
End Sub

Private Sub MnuDatorspelarensHastighet_Click()
   DatorspelarensHastighet = Val(InputBox("Ange Datorspelarens hastighet (1000 = L�ngsammast, 1 = Snabbast)", "Datorspelarens hastighet"))
   If DatorspelarensHastighet > 1000 Then
     DatorspelarensHastighet = 1000
   ElseIf DatorspelarensHastighet < 1 Then
     DatorspelarensHastighet = 1
   End If
End Sub

Private Sub mnuHighscore_Click()
  Highscore.Show
End Sub

Private Sub MnuOm_Click()
  frmOm.Show
End Sub

Private Sub mnut�rning_Click() 'V�ljer antal t�rningssnurr
  AntalSnurr = Val(InputBox("Hur m�nga varv skall t�rningarna snurra?", "Antal T�rningssnurr"))
  If AntalSnurr <= 0 Then
    AntalSnurr = 1
  End If
  If AntalSnurr >= 1000 Then
    AntalSnurr = 1000
  End If
End Sub

'Private Sub MSComm1_OnComm()
  
  'MSComm1.CommPort = Inst�lldPort
  'MSComm1.PortOpen = True
  'Select Case MinSpelare
  '  Case 1
  '    MSComm1.Output = index & " " & Aktuellspelare
  '  Do
  '    DoEvents
  '  buffer$ = buffer$ & MSComm1.Input
  '  Loop Until InStr(buffer$, "OK")
  'End Select
  


'End Sub

Private Sub Picture1_Click(index As Integer) 'Vad som h�nder n�r man klickar p� t�rningarna.
  If HarKastat = 1 Then
    If cmdkasta.Enabled = True Then
      If L�sT�rning(index) = 0 Then
        Picture1(index).Picture = ImageT�rningChecked(T�rning(index))
        L�sT�rning(index) = 1
        T�rning(index) = T�rning(index) + 0.1
      Else
        Temp(3) = T�rning(index) - 0.1
        Picture1(index).Picture = ImageT�rning(Temp(3))
        L�sT�rning(index) = 0
        T�rning(index) = T�rning(index) - 0.1
      End If
    End If
  End If
End Sub

Private Sub Sortera() 'Funktion som sorterar t�rningsv�rdena
  For j = 1 To 4
    For i = 1 To 4
      If T�rning(i) > T�rning(i + 1) Then
        Temp(1) = T�rning(i + 1)
        T�rning(i + 1) = T�rning(i)
        T�rning(i) = Temp(1)
      End If
    Next i
  Next j
End Sub

Private Sub Antal()  'Kollar antalet t�rningar
  For i = 1 To 6
    AntalT�rningar(i) = 0
  Next i
  
  For j = 1 To 6
    For i = 1 To 5
      If Int(T�rning(i)) = j Then
        AntalT�rningar(j) = AntalT�rningar(j) + 1
      End If
    Next i
  Next j
End Sub

Private Sub Bilder() 'Funktion som ritar upp bilder efter sortering
  For j = 1 To 6
    For i = 1 To 5
      Temp(2) = T�rning(i) - 0.1
      If T�rning(i) = (j + 0.1) Then
        Picture1(i).Picture = ImageT�rningChecked(Temp(2))
        L�sT�rning(i) = 1
      End If
      If T�rning(i) = (j) Then
        Picture1(i).Picture = ImageT�rning(j)
        L�sT�rning(i) = 0
      End If
    Next i
  Next j
End Sub

Private Sub �terst�llb() 'G�r om t�rningarna till heltal igen
  For i = 1 To 5
    T�rning(i) = Int(T�rning(i))
  Next i
End Sub

Private Sub Funktion_Summa()      'Denna funktion r�knar ut
  HarKastat = 0                   'po�ngsumman och har den f�rdelen
  Apan = 0                        'att man med ens kan se resultatet p� endera av
                                  'de b�da sektionerna. Den m� vara lite gr�tig..
  For k = 1 To 2
    If k = 1 Then
      B�rjan = 0
      Slut = 5
      Summan = 6
    Else                          'K = 1(ettor till sexor) eller 2(ett par osv)
      Apan = 0
      B�rjan = 8
      Slut = 16
      Summan = 17
    End If
    
    If Not SummanAv(k) = 1 Then    'Ifall en summa blivit utr�knad hoppar den �ver
      For j = B�rjan To Slut       'denna kod.
        For i = 0 To AntalSpelare
          If txt(j + 18 * i).Text > "" Then   'Kontrollerar ifall alla v�rden �r st�rre
            Apan = Apan + 1                   '�n noll.
          End If
        Next i
      Next j
  
      If Apan = (AntalSpelare + 1) * 6 And k = 1 Or Apan = (AntalSpelare + 1) * 9 And k = 2 Then 'Ifall n�gon sektion �r fylld.
        For j = B�rjan To Slut
          For i = 0 To AntalSpelare
            txt(Summan + 18 * i).Text = Val(txt(Summan + 18 * i).Text) + Val(txt(j + 18 * i))
          Next i
        Next j
        If k = 1 Then
          For i = 0 To AntalSpelare
            If Val(txt(Summan + 18 * i)) >= 63 Then
              txt(7 + 18 * i).Text = 50      'Vid =>63 utdelas 50 p bonus.
            End If
          Next i
        End If
        SummanAv(k) = 1
      End If
    End If
  Next k
  Apan = 0
  For i = 0 To 14
    If cmdval(i).Enabled = False Then    'Visar de totala summorna p� ett separat form
      Apan = Apan + 1
    End If
  Next i
  If Apan = 15 Then
    For i = 0 To AntalSpelare
      IntResultat(i) = Val(txt(6 + i * 18)) + Val(txt(7 + i * 18)) + Val(txt(17 + i * 18))
    Next i
  cmdkasta.Enabled = False
  frmResultat.Show
  End If
End Sub

Public Sub DatorSpelare()
  Dim max(1 To 2) As Integer
  Dim IntSlutF�ltEtt As Byte
  IntSlutF�ltEtt = 1
 
  Call Sleep(DatorspelarensHastighet * 2)
  
  Do Until KastKvar = 0
    
    cmdkasta_Click
    For i = 1 To 5
      T�rningTemp(i) = T�rning(i)
    Next i
    Call Sleep(DatorspelarensHastighet * 2)
    �terst�llTempb
    SorteraT�rningTemp
    Antal
           
    If T�rningTemp(1) = T�rningTemp(5) Then
      If Not txt(34) > "" Then
        Call Sleep(DatorspelarensHastighet * 4) 'Yatzy
        CmdVal_Click (14)
        GoTo Slut
      End If
    End If
    
    For i = 1 To 6
      If AntalT�rningar(i) >= 4 And (i * 4) >= Int(12 / IntSlutF�ltEtt) Then
        If Not txt(17 * Aktuellspelare + 12) = "" And txt(29) > "" Then
          Call Sleep(DatorspelarensHastighet * 4)  'Fyrtal
          CmdVal_Click (9)
          GoTo Slut
        End If
      End If
    Next i
    
    If T�rningTemp(5) = T�rningTemp(3) And T�rningTemp(2) = T�rningTemp(1) Or T�rningTemp(5) = T�rningTemp(4) And T�rningTemp(3) = T�rningTemp(1) Then 'K�k
      If T�rningTemp(1) + T�rningTemp(2) + T�rningTemp(3) + T�rningTemp(4) + T�rningTemp(5) >= Int(18 / IntSlutF�ltEtt) Then
        If Not txt(32) > "" Then
          Call Sleep(DatorspelarensHastighet * 4)
          CmdVal_Click (12)
          GoTo Slut
        End If
      End If
    End If
    
    If T�rningTemp(5) = T�rningTemp(4) And T�rningTemp(3) = T�rningTemp(2) And T�rningTemp(4) <> T�rningTemp(3) Then 'Tv�par
      If T�rningTemp(1) + T�rningTemp(2) + T�rningTemp(3) + T�rningTemp(4) + T�rningTemp(5) >= Int(15 / IntSlutF�ltEtt) Then
        If Not txt(27) > "" Then
          Call Sleep(DatorspelarensHastighet * 4)
          CmdVal_Click (7)
          GoTo Slut
        End If
      End If
    End If
    
    If AntalT�rningar(1) ^ 3 + AntalT�rningar(2) ^ 3 + AntalT�rningar(3) ^ 3 + AntalT�rningar(4) ^ 3 + AntalT�rningar(5) ^ 3 = 5 Then
      If Not txt(30) > "" Then
        Call Sleep(DatorspelarensHastighet * 4)
        CmdVal_Click (10)
        GoTo Slut
      End If
    End If
    
    If AntalT�rningar(6) ^ 3 + AntalT�rningar(2) ^ 3 + AntalT�rningar(3) ^ 3 + AntalT�rningar(4) ^ 3 + AntalT�rningar(5) ^ 3 = 5 Then
      If Not txt(31) > "" Then
        Call Sleep(DatorspelarensHastighet * 4)
        CmdVal_Click (11)
        GoTo Slut
      End If
    End If
    
    If txt(18) > "" And txt(19) > "" And txt(20) > "" And txt(21) > "" And txt(22) > "" And txt(23) > "" Then
      If T�rningTemp(1) + T�rningTemp(2) + T�rningTemp(3) + T�rningTemp(4) + T�rningTemp(5) >= Int(20 / IntSlutF�ltEtt) Then
        If Not txt(33) > "" Then
          Call Sleep(DatorspelarensHastighet * 4)
          CmdVal_Click (13)
          GoTo Slut
        End If
      End If
    End If
    
    For i = 1 To 6 'Tretal
      If AntalT�rningar(i) >= 3 And i * 3 >= Int(9 / IntSlutF�ltEtt) Then
        If Not txt(17 * Aktuellspelare + i) = "" Then
          If Not txt(28) > "" Then
            Call Sleep(DatorspelarensHastighet * 4)
            CmdVal_Click (8)
            GoTo Slut
          End If
        End If
      End If
    Next i
    
    If AntalT�rningar(6) = 2 And txt(26) = "" And KastKvar = 0 Then
      Call Sleep(DatorspelarensHastighet * 4)
      CmdVal_Click (6)
      GoTo Slut
    End If
    
    For i = 1 To 6
      If AntalT�rningar(i) >= AntalT�rningar(max(1)) And txt(i - 1 + 18).Text = "" Then
        max(1) = i
      End If
    Next i
            
    If IntSlutF�ltEtt = 2 And KastKvar = 0 Then
      For j = 34 To 26 Step -1
        If txt(j) = "" Then
          Call Sleep(DatorspelarensHastighet)
          CmdVal_Click (j - 20)
          GoTo Slut
        End If
      Next j
    End If
    
    If max(1) = 0 Then
      For i = 6 To 1 Step -1
        If AntalT�rningar(i) > AntalT�rningar(max(1)) Then
          max(1) = i
          IntSlutF�ltEtt = 2
        End If
      Next i
    End If

    For i = 1 To 5
      If Picture1(i).Picture = ImageT�rning(max(1)) Then
        Picture1_Click (i)
        Call Sleep(DatorspelarensHastighet)
      End If
    Next i
    For i = 1 To 5
      For j = 1 To 6
        If Picture1(i).Picture = ImageT�rningChecked(j) And max(1) <> Int(T�rning(i)) Then
          Picture1_Click (i)
          Call Sleep(DatorspelarensHastighet)
        End If
      Next j
    Next i
        
    Call Sleep(DatorspelarensHastighet)
  Loop
  
  Call Sleep(DatorspelarensHastighet * 4)
  CmdVal_Click (max(1) - 1)
Slut:
  max(1) = 0

End Sub
Private Sub SorteraT�rningTemp() 'Tempor�ra t�rningar beh�vs f�r att datorspelaren ska kunna kolla villkoren och dessutom v�lja vilka t�rningar han vill spara utan att han ska beh�va sortera dem.
  For j = 1 To 4
    For i = 1 To 4
      If T�rningTemp(i) > T�rningTemp(i + 1) Then
        Temp(5) = T�rningTemp(i + 1)
        T�rningTemp(i + 1) = T�rningTemp(i)
        T�rningTemp(i) = Temp(5)
      End If
    Next i
  Next j
End Sub

Private Sub �terst�llTempb()
  For i = 1 To 5
    T�rningTemp(i) = Int(T�rningTemp(i))
  Next i
End Sub

