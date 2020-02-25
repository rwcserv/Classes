VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cCBA_Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pCN As ADODB.Connection
Private RS() As ADODB.Recordset
Private RSCopy() As ADODB.Recordset
Private bIsLive As Boolean
Private pRSCnt As Long
Private pOP As Scripting.Dictionary
Private pOPCpy As Scripting.Dictionary
Private Sub Class_Initialize()
    ReDim RS(1 To 999)
    ReDim RSCopy(1 To 999)
    Set op = New Scripting.Dictionary
    Set OPCpy = New Scripting.Dictionary
    Set CN = New ADODB.Connection
End Sub
Public Function SetConnection(ByVal SystemName As String, Optional ByVal isAccessNotSQLServer As Boolean = False) As Boolean
Dim sName As String
    
    Select Case UCase(SystemName)
        Case "CBIS"
            sName = "599DBL01"
        Case "COMRADE"
            sName = "0599Z0NDBREPL01"
        Case "TEN"
            sName = "599DBL01"
        Case "CAMERA"
            sName = "599DBL01"
        Case "ABI"
            sName = "599DBL01"
        Case "STAR"
            sName = "0599Z0NDBREPL01"
        Case "OTHER2"
            sName = "599DBL01"
        Case "OTHER3"
            sName = "599DBL01"
        Case "501", "502", "503", "504", "505", "506", "507", "509", "Purchase"
            If SystemName = "Purchase" Then SystemName = "501"
            sName = "0" & SystemName & "Z0IDBSRVL02"
    End Select
    
    If sName = "" Then Exit Function
    With CN
        .ConnectionTimeout = 100
        .CommandTimeout = 100
        If UCase(SystemName) = "TEN" Then
            .Open "Provider=" & CBA_MSAccess & ";DATA SOURCE=" & g_GetDB("Ten")
        ElseIf UCase(SystemName) = "CAMERA" Then
            .Open "Provider=" & CBA_MSAccess & ";DATA SOURCE=" & g_GetDB("Cam")
        ElseIf UCase(SystemName) = "ASYST" Then
            .Open "Provider=" & CBA_MSAccess & ";DATA SOURCE=" & g_GetDB("ASYST")
        Else
            .Open "Provider=" & IIf(isAccessNotSQLServer = False, "SQLNCLI10", CBA_MSAccess) & ";DATA SOURCE=" & CBA_BasicFunctions.TranslateServerName(sName, Date) & IIf(isAccessNotSQLServer = False, ";;INTEGRATED SECURITY=sspi;", ";")
        End If
    End With
    If CN.State > 0 Then IsLive = True: SetConnection = True
End Function
Public Function Query(ByVal strSQL As String, ByVal QueryName As String, Optional ByVal ForceRequery As Boolean = False) As Boolean
    If op.Exists(QueryName) Then If ForceRequery = True Then op.Remove QueryName Else Query = True: Exit Function
    RSCnt = RSCnt + 1
    Set RS(RSCnt) = New ADODB.Recordset
    If CN.State = 0 Then Query = False: Exit Function
    If Left(QueryName, 3) = "Upd" Then
        RS(RSCnt).Open strSQL, CN, adOpenDynamic, adLockOptimistic
    Else
        RS(RSCnt).Open strSQL, CN
    End If
    If RS(RSCnt).State = 0 Then Query = False: Exit Function
    If RS(RSCnt).EOF = True Then Query = False: Exit Function
    'Set RSCopy(RSCnt) = RS(RSCnt).Clone
    'OPCpy.Add QueryName, RS(RSCnt).Clone
    op.Add QueryName, RS(RSCnt)
    Query = True
End Function
Private Property Get CN() As ADODB.Connection: Set CN = pCN: End Property
Private Property Set CN(ByVal objNewValue As ADODB.Connection): Set pCN = objNewValue: End Property
Public Property Get IsLive() As Boolean: IsLive = bIsLive: End Property
Private Property Let IsLive(ByVal bNewValue As Boolean): bIsLive = bNewValue: End Property
Private Property Get RSCnt() As Long: RSCnt = pRSCnt: End Property
Private Property Let RSCnt(ByVal NewValue As Long): pRSCnt = NewValue: End Property
Public Property Get op() As Scripting.Dictionary: Set op = pOP: End Property
Public Property Set op(ByVal NewValue As Scripting.Dictionary): Set pOP = NewValue: End Property
Private Property Get OPCpy() As Scripting.Dictionary: Set OPCpy = pOPCpy: End Property
Private Property Set OPCpy(ByVal NewValue As Scripting.Dictionary): Set pOPCpy = NewValue: End Property

