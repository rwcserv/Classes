VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CBA_BTF_ForecastTypes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit       ' CBA_BTF_ForecastTypes @CBA_BTF

Private Sal As CBA_BTF_ForecastMetrics
Private MDol As CBA_BTF_ForecastMetrics
Private MPer As CBA_BTF_ForecastMetrics
Private Cos As CBA_BTF_ForecastMetrics
Private RCVRNet As Single
Private Lvl As String
''Private RCVC As Single
Private RCVR As Single
Function Sales(Optional ByVal OriginalForecast As Single, Optional ByVal ReForecast As Single, Optional ByVal Actuals As Single) As CBA_BTF_ForecastMetrics
    If Actuals <> 0 Then Sal.Actual = Actuals
    If OriginalForecast <> 0 Then Sal.OriginalForecast = OriginalForecast
    If ReForecast <> 0 Then Sal.ReForecast = ReForecast
    If Sal.OriginalForecast <> 0 And Sal.ReForecast = 0 Then Sal.ReForecast = Sal.OriginalForecast
    Sales = Sal
End Function
Function MarginDol(Optional ByVal OriginalForecast As Single, Optional ByVal ReForecast As Single, Optional ByVal Actuals As Single) As CBA_BTF_ForecastMetrics
    If Actuals <> 0 Then MDol.Actual = Actuals
    If OriginalForecast <> 0 Then MDol.OriginalForecast = OriginalForecast
    If ReForecast <> 0 Then MDol.ReForecast = ReForecast
    If MDol.OriginalForecast <> 0 And MDol.ReForecast = 0 Then MDol.ReForecast = MDol.OriginalForecast
    MarginDol = MDol
End Function
Function MarginPer(Optional ByVal OriginalForecast As Single, Optional ByVal ReForecast As Single, Optional ByVal Actuals As Single) As CBA_BTF_ForecastMetrics
    If Actuals <> 0 Then MPer.Actual = Actuals
    If OriginalForecast <> 0 Then MPer.OriginalForecast = OriginalForecast
    If ReForecast <> 0 Then MPer.ReForecast = ReForecast
    If MPer.OriginalForecast <> 0 And MPer.ReForecast = 0 Then MPer.ReForecast = MPer.OriginalForecast
    MarginPer = MPer
End Function
Function Cost(Optional ByVal OriginalForecast As Single, Optional ByVal ReForecast As Single, Optional ByVal Actuals As Single) As CBA_BTF_ForecastMetrics
    If Actuals <> 0 Then Cos.Actual = Actuals
    If OriginalForecast <> 0 Then Cos.OriginalForecast = OriginalForecast
    If ReForecast <> 0 Then Cos.ReForecast = ReForecast
    If Cos.OriginalForecast <> 0 And Cos.ReForecast = 0 Then Cos.ReForecast = Cos.OriginalForecast
    Cost = Cos
End Function
Property Set RCVRetailNet(ByVal RRNet)
    RCVRNet = RRNet
End Property
Property Set RCVRetail(ByVal rr)
    RCVR = rr
End Property
Property Get RCVRetailNet()
    RCVRetailNet = RCVRNet
End Property
Property Get RCVRetail()
    RCVRetail = RCVR
End Property
Property Get Level()
    Level = Lvl
End Property
Function setLevel(ByVal sLevel As String)
    Lvl = sLevel
End Function
Function SalesAcc(Optional ByVal OriginalForecast As Single, Optional ByVal ReForecast As Single, Optional ByVal Actuals As Single)
    If Actuals <> 0 Then Sal.Actual = Sal.Actual + Actuals
    If OriginalForecast <> 0 Then Sal.OriginalForecast = Sal.OriginalForecast + OriginalForecast
    If ReForecast <> 0 Then Sal.ReForecast = Sal.ReForecast + ReForecast
    If ReForecast = 0 And OriginalForecast <> 0 Then Sal.ReForecast = Sal.ReForecast + OriginalForecast
End Function
Function MarginDolAcc(Optional ByVal OriginalForecast As Single, Optional ByVal ReForecast As Single, Optional ByVal Actuals As Single)
    If Actuals <> 0 Then MDol.Actual = MDol.Actual + Actuals
    If OriginalForecast <> 0 Then MDol.OriginalForecast = MDol.OriginalForecast + OriginalForecast
    If ReForecast <> 0 Then MDol.ReForecast = MDol.ReForecast + ReForecast
    If ReForecast = 0 And OriginalForecast <> 0 Then MDol.ReForecast = MDol.ReForecast + OriginalForecast
End Function
Function CostAcc(Optional ByVal OriginalForecast As Single, Optional ByVal ReForecast As Single, Optional ByVal Actuals As Single)
    If Actuals <> 0 Then Cos.Actual = Cos.Actual + Actuals
    If OriginalForecast <> 0 Then Cos.OriginalForecast = Cos.OriginalForecast + OriginalForecast
    If ReForecast <> 0 Then Cos.ReForecast = Cos.ReForecast + ReForecast
    If ReForecast = 0 And OriginalForecast <> 0 Then Cos.ReForecast = Cos.ReForecast + OriginalForecast
End Function


