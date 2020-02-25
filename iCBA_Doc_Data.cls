VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "iCBA_Doc_Data"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit                             ' iCBA_Doc_Data
    ' Note: In implementations, there is no code allowed; the code that is run is held in the objects that implement this class
    ' These functions are prefixed with iCBA_Doc_Data_'Name of function' eg. iCBA_Doc_Data_DataForDocumentType
    
Public Function DocumentType() As e_DocuType
End Function

Public Function GetVBACode(Optional ByVal additionalCode As String) As Scripting.Dictionary
End Function

Public Function SetDocID(ByVal DocID As Long) As Boolean
End Function

'Public Function GetUserForms() As Scripting.Dictionary
'End Function

Public Function GetUDTData(ByRef cUDT As cCBA_UDT, Optional ByVal ReqType As String) As Boolean
End Function

Public Function GetGrpNo(ByVal ProcedureName As String) As Long
End Function


