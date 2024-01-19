VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form CALENDAR 
   BackColor       =   &H80000007&
   Caption         =   "CALENDAR"
   ClientHeight    =   5415
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14280
   LinkTopic       =   "Form1"
   ScaleHeight     =   5415
   ScaleWidth      =   14280
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   615
      Left            =   240
      TabIndex        =   7
      Text            =   "FINAL EXAMINATION"
      Top             =   240
      Width           =   4575
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   615
      Left            =   5400
      TabIndex        =   6
      Text            =   "SECTION: CEIT-37-303A"
      Top             =   4560
      Width           =   8415
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   615
      Left            =   5160
      TabIndex        =   5
      Text            =   "CREATED BY: JAYRALD LATO PELEGRINO"
      Top             =   4080
      Width           =   8655
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   6240
      Top             =   2760
   End
   Begin MSComCtl2.DTPicker CALENDAR 
      Height          =   615
      Left            =   4080
      TabIndex        =   0
      Top             =   2040
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1085
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   -2147483637
      CalendarForeColor=   0
      CalendarTitleBackColor=   0
      CalendarTitleForeColor=   65280
      CalendarTrailingForeColor=   12632256
      CustomFormat    =   "MM-dd-yyyy"
      Format          =   38141955
      CurrentDate     =   44909
   End
   Begin VB.Label Labelyear 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   735
      Left            =   6120
      TabIndex        =   4
      Top             =   2880
      Width           =   2775
   End
   Begin VB.Label Labelnumber 
      BackStyle       =   0  'Transparent
      Caption         =   "Day"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   69
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   6600
      TabIndex        =   3
      Top             =   1560
      Width           =   2895
   End
   Begin VB.Label Labelmonth 
      BackStyle       =   0  'Transparent
      Caption         =   "Month"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   735
      Left            =   6720
      TabIndex        =   2
      Top             =   1440
      Width           =   2775
   End
   Begin VB.Label Labelday 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Weekdays"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   735
      Left            =   3240
      TabIndex        =   1
      Top             =   1440
      Width           =   3375
   End
End
Attribute VB_Name = "CALENDAR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'CREATED BY: JAYRALD LATO PELEGRINO
'FROM: CEIT-37-303A
'STUDENT ID: 2021-106623
'FINAL EXAMINATION: STOPWATCH
'PROF: JULIUS CABAUATAN

Dim today As Variant 'To process and run the day month and year code

Private Sub Timer1_Timer()
today = Now
Labelday.Caption = Format(today, "ddd") 'Label day
Labelmonth.Caption = Format(today, "mmmm") 'Label month
Labelyear.Caption = Format(today, "yyyy") 'Label year
Labelnumber.Caption = Format(today, "d") 'Label highlight day
End Sub
