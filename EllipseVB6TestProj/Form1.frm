VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   855
      Left            =   1080
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim a1 As Complex


Public Sub Command1_Click()
    Dim a1() As Double
    'a1 = GetPoints(1, 2, 3, 4, 5, 6, 7)
    'a1 = GetPoints(0, 0, 3, 4, 5, 6, 7)
    'a1 = GetPoints(0, 0, 1, 4, 5, 6, 7)
    'a1 = GetPoints(3, 0, 1, 4, 5, 2, 7)
    'a1 = GetPoints(1, 1, 3, 1, 1, 2, 7)
    'a1 = GetPoints(5, 0, 1, 4, 5, 2, 7)
    'a1 = GetPoints(0, 6, 3, 0, 1, 2, 7)
    'a1 = GetPoints(0, 5, 3, 0, 1, 2, 7)
    'a1 = GetPoints(0, 10, 3, 0, 1, 2, 7)
    a1 = GetPoints(0, 11, 3, 0, 1, 2, 7)
    'a1 = GetPoints(0, 12, 3, 0, 1, 2, 7)
    MsgBox "Pt1: " & a1(0) & "," & a1(1) & " // " & "Pt2: " & a1(2) & "," & a1(3) & " // " & "Pt3: " & a1(4) & "," & a1(5) & " // " & "Pt4: " & a1(6) & "," & a1(7)
End Sub



