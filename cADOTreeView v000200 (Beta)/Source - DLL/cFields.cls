VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFields"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'===========================================================================
'
' Class Name:   cFields
' Author:       Graeme Grant        (a.k.a. Slider)
' Date:         16/07/2003
' Version:      00.02.00 *BETA*
' Description:  Advanced ADO Handler for the TreeView Control
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
Private moFields As Collection

'===========================================================================
' Internal Class Subroutines
'
Private Sub Class_Initialize()
    Set moFields = New Collection
End Sub

Private Sub Class_Terminate()
    Set moFields = Nothing
End Sub

'===========================================================================
' Internally shared properties/methods
'
Friend Function Add(ByVal Key As String) As cField

    Dim oData As cField

    Set oData = New cField
    With oData
        .Desc = Trim$(Key)
        .Key = Trim$(Key)
        moFields.Add oData, .Key
    End With

    Set Add = oData
    Set oData = Nothing

End Function

'===========================================================================
' Public Properties
'
Public Property Get Contents() As Byte()
    '
    '## Returns the contents of the collection as a serialized byte array.
    '
    Dim oBag As New VBRUN.PropertyBag
    Dim lNdx As Long

    With oBag
        .WriteProperty "Cnt", moFields.Count, 0
        For lNdx = 1 To moFields.Count
            With moFields(lNdx)
                #If DEBUGMODE = 1 Then
                    Debug.Print "--- Field < K" + CStr(lNdx), _
                                "[" + .Key + "]"
                #End If
                oBag.WriteProperty "K" + CStr(lNdx), .Key, ""
                oBag.WriteProperty "F" + CStr(lNdx), .Contents, ""
            End With
        Next
        Contents = .Contents
    End With

End Property

Public Property Let Contents(ByRef aValue() As Byte)
    '
    '## Sets the contents of the collection as a serialized byte array.
    '
    Dim oBag As New VBRUN.PropertyBag
    Dim lNdx As Long
    
    On Error GoTo ErrorHandler
    With oBag
        .Contents = aValue
        If Err.Number = 0 Then
            For lNdx = 1 To .ReadProperty("Cnt", 0)
                #If DEBUGMODE = 1 Then
                    Debug.Print "--- Field > K" + CStr(lNdx), _
                                "[" + .ReadProperty("K" + CStr(lNdx), "") + "]"
                #End If
                moFields(.ReadProperty("K" + CStr(lNdx), "")).Contents = .ReadProperty("F" + CStr(lNdx))
            Next
        End If
    End With
    Exit Property

ErrorHandler:
    Stop
End Property

Public Property Get Count() As Long
    Count = moFields.Count
End Property

Public Property Get Item(ByVal vntIndexKey As Variant) As cField
Attribute Item.VB_UserMemId = 0
  Set Item = moFields(vntIndexKey)
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = moFields.[_NewEnum]
End Property
