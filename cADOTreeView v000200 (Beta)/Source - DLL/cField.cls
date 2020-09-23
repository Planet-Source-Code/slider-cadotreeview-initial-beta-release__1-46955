VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'===========================================================================
'
' Class Name:   cField
' Author:       Graeme Grant        (a.k.a. Slider)
' Date:         16/07/2003
' Version:      00.02.00 *BETA*
' Description:  Advanced ADO Handler for the TreeView  Control
' Edit History: 00.01.00 10/07/2003 Initial *BETA* Release
'               00.02.00 16/07/2003 Added Object persistance
'
' Compiling:   DLL : set class Instancing = 2 - PublicNotCreatable
'              APP : set class Instancing = 1 - Private
'
' Notes:       NODLL conditional compilation flag is used in the project's
'              'Conditional Compiling Arguments' property (see
'              Demo?_NODLL.PRJ properties to see usage).
'
'===========================================================================

Option Explicit

'===========================================================================
' Private: Variables and Declarations
'
Private msDesc As String
Private msKey  As String
Private msSQL  As String

'===========================================================================
' Public Properties
'
Public Property Get Contents() As Byte()
    '
    '## Returns the contents of the object as a serialized byte array.
    '
    Dim oBag As VBRUN.PropertyBag
    
    Set oBag = New VBRUN.PropertyBag
    With oBag
        .WriteProperty "D", msDesc, ""
        .WriteProperty "K", msKey, ""
        .WriteProperty "S", msSQL, ""
        Contents = .Contents
    End With

End Property

Public Property Let Contents(ByRef aValue() As Byte)
    '
    '## Sets the contents of the object as a serialized byte array.
    '
    Dim oBag As VBRUN.PropertyBag

    Set oBag = New VBRUN.PropertyBag
    With oBag
        .Contents = aValue
        msDesc = .ReadProperty("D", "")
        msKey = .ReadProperty("K", "")
        msSQL = .ReadProperty("S", "")
    End With

End Property

Public Property Get Desc() As String
Attribute Desc.VB_UserMemId = 0
Attribute Desc.VB_MemberFlags = "200"
    Desc = msDesc
End Property

Public Property Let Desc(ByVal sNewDesc As String)
    msDesc = Trim$(sNewDesc)
End Property

Public Property Get Key() As String
    Key = msKey
End Property

Public Property Let Key(ByVal sNewKey As String)
    msKey = Trim$(sNewKey)
End Property

Public Property Get SQL() As String
    SQL = msSQL
End Property

Public Property Let SQL(ByVal sNewSQL As String)
    sNewSQL = Trim$(sNewSQL)
    Select Case Len(sNewSQL)
        Case 0:    msSQL = msDesc
        Case Else: msSQL = sNewSQL
    End Select
End Property
