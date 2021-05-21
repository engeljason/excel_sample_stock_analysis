VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Stock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Stock class for holding opening and closing prices for the year and totaling stock volume
Public Opening As Double
Public FirstOpen As Long
Public Closing As Double
Public LastClose As Long
Private mVolume As Double
Private mName As String

Private Sub Class_Initialize() ' default values of 0 just so nothing explodes
    Me.Opening = CDbl(0)
    Me.FirstOpen = CDbl(0)
    Me.Closing = CDbl(0)
    Me.LastClose = CDbl(0)
    Me.Volume = CDbl(0)
End Sub

Public Property Get Name() As String
    Name = mName
End Property

Public Property Let Name(sname As String)
    mName = sname
End Property

Public Property Get Volume() As Double
    Volume = mVolume
End Property

Public Property Let Volume(lvlm As Double)
    mVolume = lvlm
End Property

Sub Update(openprice As Double, closeprice As Double, dtg As Long, vlm As Double)
    
    ' if default, then always update, otherwise make sure its the first price
    If Me.Opening = 0 Or dtg < Me.FirstOpen Then
        Me.Opening = CDbl(openprice)
        Me.FirstOpen = dtg
    End If
    
    ' if default, then always update, otherwise make sure its the last price
    If Me.Closing = 0 Or dtg > Me.LastClose Then
        Me.Closing = CDbl(closeprice)
        Me.LastClose = dtg
    End If
    
    mVolume = CDbl(mVolume) + CDbl(vlm)
End Sub


Public Function ToString() As String
    ToString = mName + " opened at " + Str(Me.Opening) + " on " + Str(Me.FirstOpen) + " and closed at " + Str(Me.Closing) + " on " + Str(Me.LastClose) + " with a total volume of " + Str(mVolume)
End Function


Public Function TotalChange() As Double
    TotalChange = CDbl(Me.Closing) - CDbl(Me.Opening)
End Function

Public Function PercentChange() As Double
    If Me.Opening = 0 Then
        PercentChange = 0
    Else
        PercentChange = (CDbl(Me.Closing) / CDbl(Me.Opening)) - CDbl(1)
    End If
End Function
