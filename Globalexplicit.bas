Attribute VB_Name = "modExplicit"
Option Explicit

Global AntalSpelare, i, j, k%
Global IntResultat(0 To 3)
Global Spelare(0 To 3) As String

Global iFilnr As Integer
Global aktHighScore As Akt
Private Type Akt
    Namn As String * 10
    Poäng As String * 10
End Type

Global BoolDatorSpelare As Boolean
