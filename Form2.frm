VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "See The Code, You Will Know !"
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3900
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   3900
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Text            =   "See The Code, You Will Know ! Easy !"
      Top             =   120
      Width           =   2895
   End
   Begin VB.Timer Timer1 
      Left            =   240
      Top             =   120
   End
   Begin VB.Label Label1 
      Caption         =   $"Form2.frx":0000
      Height          =   1215
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   3735
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
Timer1.Interval = 1000
End Sub

Private Sub Timer1_Timer()
Text1.Text = Now
End Sub
