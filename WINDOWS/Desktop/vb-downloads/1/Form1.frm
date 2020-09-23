VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Weather"
   ClientHeight    =   3300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2250
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   2250
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Quit"
      Height          =   375
      Left            =   1080
      TabIndex        =   12
      Top             =   2880
      Width           =   1095
   End
   Begin Project1.Weather Weather1 
      Left            =   360
      Top             =   3360
      _ExtentX        =   1667
      _ExtentY        =   1879
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   960
      MaxLength       =   5
      TabIndex        =   1
      Text            =   "19116"
      ToolTipText     =   "Your zip code"
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Get Weather"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Temperature 
      Height          =   255
      Left            =   1080
      TabIndex        =   21
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Wind 
      Height          =   495
      Left            =   600
      TabIndex        =   20
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Dewpoint 
      Height          =   255
      Left            =   960
      TabIndex        =   19
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Humidity 
      Height          =   255
      Left            =   1200
      TabIndex        =   18
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Visibility 
      Height          =   255
      Left            =   840
      TabIndex        =   17
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Barometer 
      Height          =   255
      Left            =   960
      TabIndex        =   16
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Sunrise 
      Height          =   255
      Left            =   840
      TabIndex        =   15
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Sunset 
      Height          =   255
      Left            =   720
      TabIndex        =   14
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Overall 
      Height          =   255
      Left            =   720
      TabIndex        =   13
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label10 
      Caption         =   "Overall:"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label Label9 
      Caption         =   "Sunset:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label Label8 
      Caption         =   "Sunrise:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label Label7 
      Caption         =   "Barometer:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "Visibility:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "Dewpoint:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Rel. Humidity:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Wind:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Temperature:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Zip Code:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Command1.Enabled = False
Weather1.LoadWeather (Text1.text)
Command1.Enabled = True
Temperature.Caption = Weather1.GetTemp & "°F"
Wind.Caption = Weather1.GetWind
Dewpoint.Caption = Weather1.GetDewpoint & "°F"
Humidity.Caption = Weather1.GetHumidity
Visibility.Caption = Weather1.GetVisibility
Barometer.Caption = Weather1.GetBarometer
Sunrise.Caption = Weather1.GetSunrise
Sunset.Caption = Weather1.GetSunset
Overall.Caption = Weather1.GetDescription
End Sub
Private Sub Command2_Click()
Unload Me
End Sub
