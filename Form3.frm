VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Now you are learning TIMER"
   ClientHeight    =   1620
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4395
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1620
   ScaleWidth      =   4395
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Left            =   120
      Top             =   1080
   End
   Begin VB.CheckBox Check2 
      Caption         =   "No"
      Height          =   495
      Left            =   2280
      TabIndex        =   3
      Top             =   840
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Yes"
      Height          =   495
      Left            =   2280
      TabIndex        =   2
      Top             =   480
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Caption         =   "Before I Start The Lesson, Let Me Ask You Something !"
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      Begin VB.Label Label1 
         Caption         =   "Does this project helps ?"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   1815
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim timebefore, timeafter  As Long
Dim totaltime As String

Private Sub Check1_Click()
timeafter = Timer 'log the ending timer
totaltime = timeafter - timebefore 'SIMPLE operation, even a moron could understand it!
Call MsgBox("You used " + totaltime + " seconds to answer this simple question!" + vbCrLf + "But at last you say yes, I am so happy(although I don't know), Thanks Guys !", vbOKOnly + vbInformation, "Thanks!!")
MsgBox "So now you should know the useful example of the TIMER function, see the code and you will understand !", vbOKOnly, "Congratulation..."
End Sub

Private Sub Check2_Click()
timeafter = Timer 'log the ending timer
totaltime = timeafter - timebefore 'SIMPLE operation, even a moron could understand it!
Call MsgBox("You used " + totaltime + " seconds to answer this simple question!" + vbCrLf + "Why this doesn't helps? Hoping that you could drop a comment for me, thanks!", vbOKOnly + vbInformation, "I hope this helps!!")
MsgBox "So now you should know the useful example of the TIMER function, see the code and you will understand !", vbOKOnly, "Congratulation..."
End Sub

Private Sub Form_Load()
timebefore = Timer ' log the initial timer
End Sub
