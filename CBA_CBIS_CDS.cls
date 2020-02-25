VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CBA_CBIS_CDS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Pc As Long
Private PD As String
Private SNo As Long
Private sName As String
Private BR As Boolean
Private PK As Integer
Property Let productcode(ByVal PCode As Long)
    Pc = PCode
End Property
Property Let Description(ByVal PDesc As String)
    PD = PDesc
End Property
Property Let SupplierNo(ByVal SuppNo As Long)
    SNo = SuppNo
End Property
Property Let SupplierName(ByVal SuppName As String)
    sName = SuppName
End Property
Property Let Branded(ByVal Brand As Boolean)
    BR = Brand
End Property
Property Let Packsize(ByVal Pack As Integer)
    PK = Pack
End Property
Property Get productcode() As Long
    productcode = Pc
End Property
Property Get Description() As String
    Description = PD
End Property
Property Get SupplierNo() As Long
    SupplierNo = SNo
End Property
Property Get SupplierName() As String
    SupplierName = sName
End Property
Property Get Branded() As Boolean
    Branded = BR
End Property
Property Get Packsize() As Integer
    Packsize = PK
End Property
