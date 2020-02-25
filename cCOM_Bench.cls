VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cCOM_Bench"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pPCode As Long
Private pCG As Byte
Private pCompCode As String
Private pMatchType As String
Private pCompetitor As String
Private pPDesc As String
Private pMatchDescription As String
Public Property Get PCode() As Long: PCode = pPCode: End Property
Public Property Let PCode(ByVal NewValue As Long): pPCode = NewValue: End Property
Public Property Get CG() As Byte: CG = pCG: End Property
Public Property Let CG(ByVal NewValue As Byte): pCG = NewValue: End Property
Public Property Get CompCode() As String: CompCode = pCompCode: End Property
Public Property Let CompCode(ByVal NewValue As String): pCompCode = NewValue: End Property
Public Property Get MatchType() As String: MatchType = pMatchType: End Property
Public Property Let MatchType(ByVal NewValue As String): pMatchType = NewValue: End Property
Public Property Get Competitor() As String: Competitor = pCompetitor: End Property
Public Property Let Competitor(ByVal NewValue As String): pCompetitor = NewValue: End Property
Public Property Get PDesc() As String: PDesc = pPDesc: End Property
Public Property Let PDesc(ByVal NewValue As String): pPDesc = NewValue: End Property
Public Property Get MatchDescription() As String: MatchDescription = pMatchDescription: End Property
Public Property Let MatchDescription(ByVal NewValue As String): pMatchDescription = NewValue: End Property
