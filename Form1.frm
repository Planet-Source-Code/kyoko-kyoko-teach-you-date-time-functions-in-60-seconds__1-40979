VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Complete Date & Time Functions - Kyoko"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   5070
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text9 
      Height          =   285
      Left            =   240
      TabIndex        =   13
      Text            =   "Text9"
      Top             =   4080
      Width           =   1935
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   240
      TabIndex        =   12
      Text            =   "Text8"
      Top             =   3720
      Width           =   1935
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   240
      TabIndex        =   11
      Text            =   "Text7"
      Top             =   3360
      Width           =   1935
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   240
      TabIndex        =   10
      Text            =   "Text6"
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Frame Frame2 
      Caption         =   "Not So Common Functions - Refer to the codes !"
      Height          =   2295
      Left            =   120
      TabIndex        =   8
      Top             =   2280
      Width           =   4815
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Text            =   "Text5"
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Common Functions"
      Height          =   2055
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4815
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Text            =   "Text2"
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Text            =   "Text3"
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Text            =   "Text4"
         Top             =   1440
         Width           =   1935
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Wait, why my time stop there, it should update every second !!??"
         Height          =   615
         Left            =   2160
         TabIndex        =   3
         Top             =   480
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.CommandButton Command3 
         Caption         =   "<-----What the heck is this for??"
         Height          =   255
         Left            =   2160
         TabIndex        =   2
         Top             =   1440
         Visible         =   0   'False
         Width           =   2535
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Just Do It !"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   4680
      Width           =   4815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Command2.Visible = True 'ignore this
Command3.Visible = True 'ignore this
Text1.Text = Now 'show date & time at once
Text2.Text = Date 'show date
Text3.Text = Time 'show time
Text4.Text = Timer 'show number of seconds since midnight

'''Below are uncommon Date & Time Functions, don't worry, EASY!'''
Dim hour, minute, second As Integer
hour = 14 ' pls note that the output is not in 24 hour format, but is in 12 hour format, and it will automatically assigns PM or AM to the time
minute = 59
second = 59
Text5.Text = TimeSerial(hour, minute, second) 'return an internal data value for the three arguments.(Run the code, You will understand!)

Dim year, month, day As Integer
year = 2020
month = 12
day = 31
Text6.Text = DateSerial(year, month, day) 'see the output!!

'''Below are advance Date & Time Functions'''
'Before you use this functions, you should refer to the Table below
'   Interval   Description
'   h          Hour
'   d          Day
'   m          Month
'   n          Minute
'   q          Quarter
'   s          Second
'   y          Day of Year
'   w          Weekday
'   ww         Week
'   yyyy       Year
'Now you are going to learn dateadd(), datediff() and datepart() functions
'dateadd(interval, interval value to be added, date)
'datediff(interval, date1, date2)
'datepart(interval, date)
Text7.Text = DateAdd("d", 12, Date) ' add 12 d(Days) to the current Date
Text8.Text = DateDiff("m", 12 / 12 / 2020, Date) 'the output shows that how many m(Month) needed to reach 12/12/2020 compare to the current date
Text9.Text = DatePart("yyyy", Date) ' the output code should be the value of current year, for example, if this year is 2002, then text9.text = 2002
                                    ' datepart() function is good to strip of specific value from the date
End Sub

Private Sub Command2_Click()
Form2.Show
End Sub

Private Sub Command3_Click()
Form3.Show
End Sub
