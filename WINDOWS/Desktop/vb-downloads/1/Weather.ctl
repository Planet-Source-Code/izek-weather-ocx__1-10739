VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.UserControl Weather 
   CanGetFocus     =   0   'False
   ClientHeight    =   1770
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2085
   InvisibleAtRuntime=   -1  'True
   Picture         =   "Weather.ctx":0000
   ScaleHeight     =   1770
   ScaleWidth      =   2085
   ToolboxBitmap   =   "Weather.ctx":0674
   Begin InetCtlsObjects.Inet Inet 
      Left            =   1080
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
End
Attribute VB_Name = "Weather"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Type cWeather
    cHeat As String
    cWind As String
    cTemp As String
    cDewpoint As String
    cHumidity As String
    cBarometer As String
    cSunrise As String
    cSunset As String
    cVisibility As String
    cDescription As String
End Type
Private cWeather As cWeather
Sub LoadWeather(zipcode As String)
Dim text As String
Dim Search As String
Dim Spot As Integer, tempo As String
Dim Spot2 As Integer, Text2 As String
Dim pos1 As Long

text = Inet.OpenURL("http://www.weather.com/weather/us/zips/" & zipcode & ".html")
Search = "<FONT FACE=""Arial, Helvetica, Chicago, Sans Serif"" SIZE=3><B>"
Text2 = text
    
tempo = "Temp:</B></FONT></TD>" & Chr(10) & "<TD WIDTH=5><IMG WIDTH=5 HEIGHT=1 SRC=" & Chr(34) & "http://image.weather.com/pics/blank.gif" & Chr(34) & " ALT=" & Chr(34) & Chr(34) & "></TD>" & Chr(10) & "          <TD WIDTH=90><FONT FACE=" & Chr(34) & "Arial, Helvetica, Chicago, Sans Serif" & Chr(34) & " SIZE=2>"
pos1 = InStr(Text2, tempo)
cWeather.cTemp = Mid(Text2, pos1 + Len(tempo))
cWeather.cTemp = Mid(cWeather.cTemp, 1, InStr(cWeather.cTemp, "&") - 1)

tempo = "Heat Index:</B></FONT></TD>" & Chr(10) & "<TD WIDTH=5><IMG WIDTH=5 HEIGHT=1 SRC=" & Chr(34) & "http://image.weather.com/pics/blank.gif" & Chr(34) & " ALT=" & Chr(34) & Chr(34) & "></TD>" & Chr(10) & "          <TD WIDTH=90><FONT FACE=" & Chr(34) & "Arial, Helvetica, Chicago, Sans Serif" & Chr(34) & " SIZE=2>"
pos1 = InStr(Text2, tempo)
cWeather.cHeat = Mid(Text2, pos1 + Len(tempo))
tempo = InStr(cWeather.cHeat, "&")
cWeather.cHeat = Mid(cWeather.cHeat, 1, tempo - 1)

tempo = "Wind:</B></FONT></TD>" & Chr(10) & "<TD WIDTH=5><IMG WIDTH=5 HEIGHT=1 SRC=" & Chr(34) & "http://image.weather.com/pics/blank.gif" & Chr(34) & " ALT=" & Chr(34) & Chr(34) & "></TD>" & Chr(10) & "          <TD WIDTH=90><FONT FACE=" & Chr(34) & "Arial, Helvetica, Chicago, Sans Serif" & Chr(34) & " SIZE=2>"
pos1 = InStr(Text2, tempo)
cWeather.cWind = Mid(Text2, pos1 + Len(tempo))
tempo = InStr(cWeather.cWind, "<")
cWeather.cWind = Mid(cWeather.cWind, 1, tempo - 1)
    
tempo = "Dewpoint:</B></FONT></TD>" & Chr(10) & "<TD WIDTH=5><IMG WIDTH=5 HEIGHT=1 SRC=" & Chr(34) & "http://image.weather.com/pics/blank.gif" & Chr(34) & " ALT=" & Chr(34) & Chr(34) & "></TD>" & Chr(10) & "          <TD WIDTH=90><FONT FACE=" & Chr(34) & "Arial, Helvetica, Chicago, Sans Serif" & Chr(34) & " SIZE=2>"
pos1 = InStr(Text2, tempo)
cWeather.cDewpoint = Mid(Text2, pos1 + Len(tempo))
tempo = InStr(cWeather.cDewpoint, "&")
cWeather.cDewpoint = Mid(cWeather.cDewpoint, 1, tempo - 1)
    
tempo = "Rel. Humidity:</B></FONT></TD>" & Chr(10) & "<TD WIDTH=5><IMG WIDTH=5 HEIGHT=1 SRC=" & Chr(34) & "http://image.weather.com/pics/blank.gif" & Chr(34) & " ALT=" & Chr(34) & Chr(34) & "></TD>" & Chr(10) & "          <TD WIDTH=90><FONT FACE=" & Chr(34) & "Arial, Helvetica, Chicago, Sans Serif" & Chr(34) & " SIZE=2>"
pos1 = InStr(Text2, tempo)
cWeather.cHumidity = Mid(Text2, pos1 + Len(tempo))
tempo = InStr(cWeather.cHumidity, "<")
cWeather.cHumidity = Mid(cWeather.cHumidity, 1, tempo - 1)
    
tempo = "Visibility:</B></FONT></TD>" & Chr(10) & "<TD WIDTH=5><IMG WIDTH=5 HEIGHT=1 SRC=" & Chr(34) & "http://image.weather.com/pics/blank.gif" & Chr(34) & " ALT=" & Chr(34) & Chr(34) & "></TD>" & Chr(10) & "          <TD WIDTH=90><FONT FACE=" & Chr(34) & "Arial, Helvetica, Chicago, Sans Serif" & Chr(34) & " SIZE=2>"
pos1 = InStr(Text2, tempo)
cWeather.cVisibility = Mid(Text2, pos1 + Len(tempo))
tempo = InStr(cWeather.cVisibility, "<")
cWeather.cVisibility = Mid(cWeather.cVisibility, 1, tempo - 1)
    
tempo = "Barometer:</B></FONT></TD>" & Chr(10) & "<TD WIDTH=5><IMG WIDTH=5 HEIGHT=1 SRC=" & Chr(34) & "http://image.weather.com/pics/blank.gif" & Chr(34) & " ALT=" & Chr(34) & Chr(34) & "></TD>" & Chr(10) & "          <TD WIDTH=90><FONT FACE=" & Chr(34) & "Arial, Helvetica, Chicago, Sans Serif" & Chr(34) & " SIZE=2>"
pos1 = InStr(Text2, tempo)
cWeather.cBarometer = Mid(Text2, pos1 + Len(tempo))
tempo = InStr(cWeather.cBarometer, "<")
cWeather.cBarometer = Mid(cWeather.cBarometer, 1, tempo - 1)
    
tempo = "Sunrise:</B></FONT></TD>" & Chr(10) & "<TD WIDTH=5><IMG WIDTH=5 HEIGHT=1 SRC=" & Chr(34) & "http://image.weather.com/pics/blank.gif" & Chr(34) & " ALT=" & Chr(34) & Chr(34) & "></TD>" & Chr(10) & "          <TD WIDTH=90><FONT FACE=" & Chr(34) & "Arial, Helvetica, Chicago, Sans Serif" & Chr(34) & " SIZE=2>"
pos1 = InStr(Text2, tempo)
cWeather.cSunrise = Mid(Text2, pos1 + Len(tempo))
tempo = InStr(cWeather.cSunrise, "<")
cWeather.cSunrise = Mid(cWeather.cSunrise, 1, tempo - 1)
    
tempo = "Sunset:</B></FONT></TD>" & Chr(10) & "<TD WIDTH=5><IMG WIDTH=5 HEIGHT=1 SRC=" & Chr(34) & "http://image.weather.com/pics/blank.gif" & Chr(34) & " ALT=" & Chr(34) & Chr(34) & "></TD>" & Chr(10) & "          <TD WIDTH=90><FONT FACE=" & Chr(34) & "Arial, Helvetica, Chicago, Sans Serif" & Chr(34) & " SIZE=2>"
pos1 = InStr(Text2, tempo)
cWeather.cSunset = Mid(Text2, pos1 + Len(tempo))
tempo = InStr(cWeather.cSunset, "<")
cWeather.cSunset = Mid(cWeather.cSunset, 1, tempo - 1)
    
Spot = InStr(1, text, Search) + Len(Search)
Spot2 = InStr(Spot, text, "</B>")
cWeather.cDescription = Mid$(text, Spot, Spot2 - Spot)
    
End Sub
Function GetHeat() As String
GetHeat = cWeather.cHeat
End Function
Function GetTemp() As String
GetTemp = cWeather.cTemp
End Function
Function GetWind() As String
GetWind = cWeather.cWind
End Function
Function GetDewpoint() As String
GetDewpoint = cWeather.cDewpoint
End Function
Function GetHumidity() As String
GetHumidity = cWeather.cHumidity
End Function
Function GetBarometer() As String
GetBarometer = cWeather.cBarometer
End Function
Function GetSunrise() As String
GetSunrise = cWeather.cSunrise
End Function
Function GetSunset() As String
GetSunset = cWeather.cSunset
End Function
Function GetVisibility() As String
GetVisibility = cWeather.cVisibility
End Function
Function GetDescription() As String
GetDescription = cWeather.cDescription
End Function
Private Sub UserControl_Resize()
    Width = 945
    Height = 1065
End Sub
