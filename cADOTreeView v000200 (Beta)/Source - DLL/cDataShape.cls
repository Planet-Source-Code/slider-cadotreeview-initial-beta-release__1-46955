VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDataShape"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'===========================================================================
'
' Class Name:   cDataShape
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
Private moDataObjs      As Collection, _
        msConnect       As String, _
        mbLoadOnDemand  As Boolean, _
        miParentKeySize As Integer, _
        msTopParent     As String

'===========================================================================
' Public Declarations
'
Public Enum eJetVersion
    ejvJet3 = 3
    ejvJet4 = 4
End Enum

'===========================================================================
' Internal Class Subroutines
'
Private Sub Class_Initialize()
    Set moDataObjs = New Collection
End Sub

Private Sub Class_Terminate()
    Set moDataObjs = Nothing
End Sub

'===========================================================================
' Internally shared properties/methods
'
Friend Property Get fConnectString() As String
    fConnectString = msConnect
End Property

Friend Property Get fGetItem(ByVal NodeKey As String) As cDataObj
    '
    '## Returns a DataObj (Table definition) from a Node's Key
    '
    Dim sIDTag As String, _
        oData  As cDataObj

    '
    '## Extract the IDTag from the Node key
    '
    sIDTag = fGetIDTag(NodeKey)
    '
    '## Loop through the Table Definitions to find the correct object to return
    '
    For Each oData In moDataObjs
        If oData.IDTag = sIDTag Then
            Set fGetItem = oData
            Exit For
        End If
    Next

End Property

Friend Property Get fGetIDTag(ByVal NodeKey As String) As String

    If miParentKeySize Then
        '
        '## Working with a custom (alphanumeric) record IDs
        '
        fGetIDTag = Left$(NodeKey, Len(NodeKey) - miParentKeySize)
    Else
        '
        '## Working with numerical record IDs
        '
        fGetIDTag = Mid$(NodeKey, Len(CStr(Val(NodeKey))) + 1)
    End If

End Property

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
        .WriteProperty "Con", msConnect, ""
        .WriteProperty "LOD", mbLoadOnDemand, True
        .WriteProperty "PKS", miParentKeySize, 0
        .WriteProperty "Cnt", moDataObjs.Count, 0

        For lNdx = 1 To Count
            With moDataObjs(lNdx)
                #If DEBUGMODE = 1 Then
                    Debug.Print "DataObj < K" + CStr(lNdx), _
                                "[" + .Key + "]", _
                                "[" + .TableName + "]", _
                                "[" + .IDTag + "]", _
                                "[" + .ParentKey + "]"
                #End If
                oBag.WriteProperty "K" + CStr(lNdx), .Key, ""
                oBag.WriteProperty "TN" + CStr(lNdx), .TableName, ""
                oBag.WriteProperty "ID" + CStr(lNdx), .IDTag, ""
                oBag.WriteProperty "PK" + CStr(lNdx), .ParentKey, ""
                oBag.WriteProperty "DO" + CStr(lNdx), .Contents, ""
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
            Set moDataObjs = New Collection
            
            msConnect = .ReadProperty("Con", "")
            mbLoadOnDemand = .ReadProperty("LOD", True)
            miParentKeySize = .ReadProperty("PKS", 0)
            For lNdx = 1 To .ReadProperty("Cnt", 0)
                #If DEBUGMODE = 1 Then
                    Debug.Print "DataObj > K" + CStr(lNdx), _
                                "[" + .ReadProperty("K" + CStr(lNdx), "") + "]", _
                                "[" + .ReadProperty("TN" + CStr(lNdx), "") + "]", _
                                "[" + .ReadProperty("ID" + CStr(lNdx), "") + "]", _
                                "[" + .ReadProperty("PK" + CStr(lNdx), "") + "]"
                #End If
                Add(.ReadProperty("K" + CStr(lNdx), ""), _
                    .ReadProperty("TN" + CStr(lNdx), ""), _
                    .ReadProperty("ID" + CStr(lNdx), ""), _
                    .ReadProperty("PK" + CStr(lNdx), "") _
                    ).Contents = oBag.ReadProperty("DO" + CStr(lNdx))
            Next
        End If
    End With
    Exit Property

ErrorHandler:
    With Err
        .Raise .Number, "cDataShape", .Description
    End With
End Property

Public Property Get LoadOnDemand() As Boolean
    LoadOnDemand = mbLoadOnDemand
End Property

Public Property Let LoadOnDemand(ByVal bState As Boolean)
    mbLoadOnDemand = bState
End Property

Public Property Get ParentKeySize() As Integer
    ParentKeySize = miParentKeySize
End Property

Public Property Let ParentKeySize(ByVal iNewSize As Integer)
    miParentKeySize = iNewSize
End Property

Public Property Get Item(ByVal vntIndexKey As Variant) As cDataObj
Attribute Item.VB_UserMemId = 0
    If InStr("LONG*INTEGER*", UCase$(TypeName(vntIndexKey)) + "*") Then
        Set Item = moDataObjs(msTopParent)
    Else
        Set Item = moDataObjs(vntIndexKey)
    End If
End Property

Public Property Get Count() As Long
    Count = moDataObjs.Count
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = moDataObjs.[_NewEnum]
End Property

'===========================================================================
' Public Subroutines and Functions
'
Public Function Add(ByVal Key As String, _
                    ByVal TableName As String, _
                    ByVal IDTag As String, _
                    ByVal ParentKey As String, _
           Optional ByVal Recursive As Boolean = False, _
           Optional ByVal ForeColor As OLE_COLOR = vbWindowText, _
           Optional ByVal BackColor As OLE_COLOR = vbWindowBackground, _
           Optional ByVal Bold As Boolean = False, _
           Optional ByVal vImage As Variant, _
           Optional ByVal vSelectedImage As Variant, _
           Optional ByVal vExpandedImage As Variant) As cDataObj

    Dim sGetKey As String, _
        sTmpKey As String, _
        oData   As cDataObj

    Set oData = New cDataObj

    With oData
        .Key = Trim$(Key)
        .TableName = Trim$(TableName)
        .IDTag = Trim$(IDTag)
        .ParentKey = Trim$(ParentKey)
        If miParentKeySize Then
            .ParentKey = Left$(.ParentKey + Space$(miParentKeySize), miParentKeySize)
        End If
        .Recursive = Recursive

        .ForeColor = ForeColor
        .BackColor = BackColor
        .Bold = Bold
        If Not IsMissing(vImage) Then .Image = vImage
        If Not IsMissing(vSelectedImage) Then .SelectedImage = vSelectedImage
        If Not IsMissing(vExpandedImage) Then .ExpandedImage = vExpandedImage

        '
        '## Setup top-down relationship tracing
        '
        If Len(.ParentKey) = 0 Then
            If Len(msTopParent) = 0 Then
                msTopParent = .Key
            Else
                sGetKey = msTopParent
                Do
                    sTmpKey = sGetKey
                    sGetKey = pCastObj(sGetKey).fSibling
                Loop Until Len(sGetKey) = 0
                pCastObj(sTmpKey).fSibling = .Key
            End If
        Else
            If Len(pCastObj(.ParentKey).fChild) = 0 Then
                pCastObj(.ParentKey).fChild = .Key
            Else
                sGetKey = pCastObj(.ParentKey).fChild
                Do
                    sTmpKey = sGetKey
                    sGetKey = pCastObj(sGetKey).fSibling
                Loop Until Len(sGetKey) = 0
                pCastObj(sTmpKey).fSibling = .Key
            End If
        End If

    End With

    moDataObjs.Add oData, Key

    Set Add = oData
    Set oData = Nothing

End Function

Public Sub Clear()
    Set moDataObjs = New Collection
End Sub

Public Sub Remove(ByVal vntIndexKey As Variant)
    moDataObjs.Remove vntIndexKey
End Sub

Public Sub ConnectString(ByVal FileName As String, _
                Optional ByVal User As String = "admin", _
                Optional ByVal Password As String = "", _
                Optional ByVal DefPath As String = "", _
                Optional ByVal JetVersion As eJetVersion = ejvJet3)
    '
    '## Defines the connection to the Database. Currently only MS Access is supported
    '
    If Len(Trim$(DefPath)) = 0 Then DefPath = App.Path + "\"
    Select Case JetVersion
        Case ejvJet3
            msConnect = "Driver={Microsoft Access Driver (*.mdb)};DBQ=" + _
                        Trim$(FileName) + ";DefaultDir=" + Trim$(DefPath) + ";UID=" + _
                        Trim$(User) + ";PWD=;" + Trim$(Password)
        Case ejvJet4
            msConnect = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                        "Data Source=" & Trim$(FileName) & ";" & _
                        "Jet OLEDB:Database Password=;" & _
                        "Jet OLEDB:Engine Type=5;"
    End Select

End Sub

'===========================================================================
' General: Private subroutines and functions
'
Private Function pCastObj(ByVal Key As String) As cDataObj
    '
    '## Required only for cDataObj 'Friend' properties. VB gags in the 'Add'
    '   function when working directly with the collection object for some reason!
    '
    Set pCastObj = moDataObjs(Key)
End Function
