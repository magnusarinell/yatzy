VERSION 5.00
Begin VB.Form frmOm 
   BackColor       =   &H00400000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Om..."
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   2220
   ForeColor       =   &H00400000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   2220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FEE7DE&
      Height          =   1160
      Left            =   600
      Picture         =   "Om.frx":0000
      ScaleHeight     =   1095
      ScaleWidth      =   870
      TabIndex        =   3
      Top             =   1080
      Width           =   935
   End
   Begin VB.CommandButton Okej 
      Caption         =   "&Ok"
      Height          =   330
      Left            =   600
      TabIndex        =   0
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label lblavMagnus 
      BackStyle       =   0  'Transparent
      Caption         =   "av Magnus Broberg"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1800
      TabIndex        =   5
      Top             =   360
      Width           =   1575
      Visible         =   0   'False
   End
   Begin VB.Label lblEmail 
      BackStyle       =   0  'Transparent
      Caption         =   ""
      BeginProperty Font 
         Name            =   "Westminster"
         Size            =   8.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "Version 10.23"
      BeginProperty Font 
         Name            =   "Westminster"
         Size            =   8.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   720
      Width           =   975
   End
   Begin VB.Label lblYatzy 
      BackStyle       =   0  'Transparent
      Caption         =   "Yatzy!"
      BeginProperty Font 
         Name            =   "Westminster"
         Size            =   18
         Charset         =   0
         Weight          =   500
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "frmOm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Okej_Click()
   Unload Me
End Sub
