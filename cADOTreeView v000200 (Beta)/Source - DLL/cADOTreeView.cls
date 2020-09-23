VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cADOTreeView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'===========================================================================
'
' Class Name:   cADOTreeView *BETA*
' Author:       Graeme Grant        (a.k.a. Slider)
' Date:         16/07/2003
' Version:      00.02.00 *BETA*
' Description:  Advanced ADO Handler for the TreeView Control
' Edit History: 00.01.00 10/07/2003 Initial *BETA* Release
'               00.02.00 16/07/2003 Added Object persistance
'
' Compiling:   DLL : set class Instancing = 5 - MultiUse
'              APP : set class Instancing = 1 - Private
'
' Notes:       NODLL conditional compilation flag is used in the project's
'              'Conditional Compiling Arguments' property (see
'              Demo?_NODLL.PRJ properties to see usage).
'
'===========================================================================

Option Explicit

'===========================================================================
' Debugging... To change goto Project > Properties > Make > Conditional
'              Compilation Arguments
'
#If DEBUGMODE = 1 Then
    Private dbgCtrlName  As String
#End If

'===========================================================================
' Private: Internal Error Declarations
'
Private Const csSOURCE_ERR    As String = "cADOTreeView"
Private Const clTREECOPY_ERR1 As Long = vbObjectError + 2048 + 201
Private Const csTREECOPY_ERR1 As String = "Destination child node does not exist."
Private Const clTREEHOOK_ERR  As Long = vbObjectError + 2048 + 203
Private Const csTREEHOOK_ERR  As String = "Invalid control! Must be a VB6 TreeView control"

'===========================================================================
' Private: Variables and Declarations
'
Private WithEvents oTree As MSComctlLib.TreeView
Attribute oTree.VB_VarHelpID = -1
Private WithEvents oLod  As cLODTreeview        '## Manages the Load-On-Demand
Attribute oLod.VB_VarHelpID = -1

Private moDataShape      As cDataShape          '## Hold the Table relationship definition
Private moCon            As ADODB.Connection

'===========================================================================
' Public: Event Declarations
'
Public Event BeforeLoading(ByVal Node As MSComctlLib.Node, ByRef Cancel As Boolean)
Public Event AfterLoading(ByVal Node As MSComctlLib.Node)

'===========================================================================
' Internal Class Subroutines
'
Private Sub Class_Initialize()
    Set moDataShape = New cDataShape
    Set oLod = New cLODTreeview
End Sub

Private Sub Class_Terminate()
    Set oTree = Nothing
    Set moDataShape = Nothing
    Set oLod = Nothing
End Sub

'===========================================================================
' Private Load-on-Demand cLODTreeView wrapper Event
'
Private Sub oLod_Expanding(ByVal Node As MSComctlLib.Node, ByRef Cancel As Boolean)
    '
    '## This event is only called if a Parent Node is not already loaded.
    '
    RaiseEvent BeforeLoading(Node, Cancel)
    If Not Cancel Then
        pExpandBranch Node.Key
        RaiseEvent AfterLoading(Node)
    End If

End Sub

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
        '
        '## DataShape
        '
        .WriteProperty "DS", moDataShape.Contents
        '
        '## Associated ImageList control
        '
        pImagelist2Propertybag oTree.ImageList, oBag

        Contents = .Contents
    End With

End Property

Public Property Let Contents(ByRef aValue() As Byte)
    '
    '## Sets the contents of the object as a serialized byte array.
    '
    Dim oBag As VBRUN.PropertyBag, _
        oIL  As MSComctlLib.ImageList, _
        lNdx As Long

    On Error GoTo ErrorHandler
    Set oBag = New VBRUN.PropertyBag
    With oBag
        .Contents = aValue
        '
        '## DataShape
        '
        Set moDataShape = New cDataShape
        moDataShape.Contents = .ReadProperty("DS")
        '
        '## Associated ImageList control
        '
        pPropertybag2Imagelist oBag, oTree.ImageList

    End With

Exit Property

ErrorHandler:
    With Err
        .Raise .Number, "cADOTreeView::" + .Source, .Description
    End With

End Property

Public Function LoadShape(ByVal FileName As String) As Boolean

    Dim fHndl    As Long, _
        aBits()  As Byte

    If pFileExist(FileName) Then
        fHndl = FreeFile(0)
        Open FileName For Binary As #fHndl
            ReDim aBits(0 To LOF(fHndl) - 1)
            Get #fHndl, , aBits
            Contents = aBits
        Close #fHndl
    Else
        '!!raise error!
    End If

End Function

Public Function SaveShape(ByVal FileName As String) As Boolean

    Dim fHndl As Long, _
    aBits()   As Byte

    If pFileExist(FileName) Then Kill FileName

    fHndl = FreeFile(0)
    aBits = Contents
    Open FileName For Binary As #fHndl
    Put #fHndl, , aBits
    Close #fHndl

End Function

'===========================================================================
' Public Subroutines and Functions
'
Public Function DataShape() As cDataShape
    Set DataShape = moDataShape
End Function

#If NODLL = 0 Then
    Public Sub HookCtrl(ByRef Ctrl As Object)

    If Not TypeName(Ctrl) = "TreeView" Then
        Err.Raise clTREEHOOK_ERR, csSOURCE_ERR, csTREEHOOK_ERR
        Exit Sub
    End If

#Else

    Public Sub HookCtrl(ByRef Ctrl As MSComctlLib.TreeView)

#End If

    Set oTree = Ctrl                                '## Capture the Treeview control.
                                                    '   Class won't work if this isn't called first.
    oLod.HookCtrl Ctrl
    With oTree
        #If DEBUGMODE = 1 Then
            dbgCtrlName = .Parent.Name + "." + .Name
        #End If
    End With

End Sub

Public Function KeyExist(ByVal vItemKey As Variant) As Boolean
    '
    '## Returns whether a Node Key Exists in the TreeView Control
    '
    On Error Resume Next
    KeyExist = (Len(oTree.Nodes(vItemKey).Key) > 0)

End Function

Public Function Node2ShapeKey(ByVal NodeKey As String) As String
    '
    '## Returns a DataObj key based on a TreeView Node Key
    '
    If KeyExist(NodeKey) Then
        Node2ShapeKey = moDataShape.fGetItem(NodeKey).Key
    End If

End Function

Public Sub Reload(Optional ByVal oNode As MSComctlLib.Node, _
                  Optional ByVal KeepState As Boolean)
    '
    '## Currently only Loads the Treeview with Data. Refreshing branches has not been fully implemented - you have been warned!
    '
    Dim oData As cDataObj

    On Error GoTo ErrorHandler

    If moCon Is Nothing Then
        Set moCon = New ADODB.Connection
        moCon.Open moDataShape.fConnectString
    End If

    #If DEBUGMODE = 1 Then
        For Each oData In moDataShape
            Debug.Print oData.Key + "(" + oData.TableName + ")  SIB:" + oData.fSibling + "  CHILD:" + oData.fChild
        Next
    #End If

    With oTree
        .Visible = False
        If oNode Is Nothing Then
            .Nodes.Clear
            pExpandBranch , True
        Else
            If oLod.ExpandIcon(oNode) And oNode.Children Then
                '!! 1. Remember key node states
                '!! 2. Delete Branch
                '!! 3. Reload Branch
                pExpandBranch oNode, True, KeepState
            Else
                '
            End If
        End If

        .Visible = True
    End With

Exit Sub

ErrorHandler:
    oTree.Visible = True
    With Err
        .Raise .Number, "cADOTreeView::" + .Source, .Description
    End With

End Sub

'Public Function Update()
'    '
'    '!! NOT YET IMPLEMENTED!
'    '
'End Function
'
'===========================================================================
' General: Private subroutines and functions
'
Private Function pAddNode(Optional ByVal vRelative As Variant, _
                          Optional ByVal eRelationship As MSComctlLib.TreeRelationshipConstants, _
                          Optional ByVal sKey As String, _
                          Optional ByVal sText As String, _
                          Optional ByVal vImage As Variant, _
                          Optional ByVal vSelectedImage As Variant, _
                          Optional ByVal ForeColor As OLE_COLOR, _
                          Optional ByVal BackColor As OLE_COLOR, _
                          Optional ByVal Bold As Boolean = False, _
                          Optional ByVal vExpandedImage As Variant) As MSComctlLib.Node
    '
    '## Adds a Node to the TreeView Control and sets multiple properties in a single call
    '
    #If DEBUGMODE = 1 Then
        Debug.Print "Add::  Name = " + sText + "  Key =" + sKey
    #End If

    Set pAddNode = oTree.Nodes.Add(vRelative, eRelationship, sKey, sText, vImage, vSelectedImage)

    With pAddNode
        If ForeColor = 0 Then
            .ForeColor = vbWindowText
        Else
            .ForeColor = ForeColor
        End If
        If BackColor = 0 Then
            .BackColor = vbWindowBackground
        Else
            .BackColor = BackColor
        End If
        .Bold = Bold

        If Not IsMissing(vExpandedImage) Then
            pAddNode.ExpandedImage = vExpandedImage
        End If

    End With

End Function

Private Function pCheckForKids(ByVal oData As cDataObj) As Boolean
    '
    '## Returns 'True' if the Table relationship definition has more branches
    '
    pCheckForKids = True
    If moDataShape.Count > 1 Then
        If Len(oData.fSibling) = 0 Then
            If Len(oData.fChild) = 0 Then
                pCheckForKids = False
            End If
        End If
    End If

End Function

Private Sub pExpandBranch(Optional ByVal NodeKey As String, _
                          Optional ByVal Initialize As Boolean = False, _
                          Optional ByVal KeepState As Boolean = False)
    '
    '## This is the main engine for loading data from the database into the TreeView Control
    '
    Dim oRS            As ADODB.Recordset, _
        oCmd           As ADODB.Command, _
        oNode          As MSComctlLib.Node, _
        eRelationship  As MSComctlLib.TreeRelationshipConstants, _
        oData          As cDataObj, _
        sData()        As String, _
        sParentID      As String, _
        sKey           As String, _
        lID            As Long, _
        lCntRS         As Long, _
        lCount         As Long, _
        lLoop          As Long, _
        jLoop          As Long, _
        lForeColor     As OLE_COLOR, _
        lBackColor     As OLE_COLOR, _
        lBold          As Boolean, _
        vImage         As Variant, _
        vSelectedImage As Variant, _
        vExpandedImage As Variant

    '
    '## Define data Layout
    '
    If Len(NodeKey) = 0 Then
        '
        '## No data displayed - start at top data layout layer
        '
        Set oData = moDataShape(-1)
        sParentID = "0"
        eRelationship = 0
    Else
        '
        '## Extract start Data Layout layer
        '
        Set oData = moDataShape.fGetItem(NodeKey)
        eRelationship = tvwChild
        sParentID = CStr(Val(NodeKey))
    End If
    '
    '## Check to see if were dealing with a non-recursive layer
    '
    If Not oData.Recursive Then
        '
        '## Get the next layer
        '
        If Len(oData.fChild) Then
            Set oData = moDataShape(oData.fChild)
        Else
            Exit Sub
        End If
    End If

    lCount = 1
    ReDim sData(1 To lCount)
    '
    '## Extract all the DataObj keys that we need to load
    '
    sData(lCount) = oData.Key
    Do
        If Not Initialize Then
            If Len(oData.fChild) Then
                Set oData = moDataShape(oData.fChild)
                lCount = lCount + 1
                ReDim Preserve sData(1 To lCount)
                sData(lCount) = oData.Key
            End If
        End If
        Do
            If Len(oData.fSibling) Then
                Set oData = moDataShape(oData.fSibling)
                lCount = lCount + 1
                ReDim Preserve sData(1 To lCount)
                sData(lCount) = oData.Key
            Else
                Exit Do
            End If
        Loop
        If Initialize Then
            Set oData = moDataShape(sData(1))
            Initialize = False
        Else
            Exit Do
        End If

    Loop

    #If DEBUGMODE = 1 Then
        '
        '## Display data layout for debugging purposes
        '
        Debug.Print "=========================="
        Debug.Print "Data Keys: ";
        For jLoop = 1 To lCount
            Debug.Print "["; moDataShape(sData(jLoop)).Key; "]",
        Next
        Debug.Print vbCrLf + "--------------------------"
    #End If

    Set oCmd = New ADODB.Command
    Set oCmd.ActiveConnection = moCon
    '
    '## Now step through data layout: extrapolate from DB and display
    '
    For jLoop = 1 To lCount
        Set oData = moDataShape(sData(jLoop))
        Set oRS = New ADODB.Recordset
        With oRS
            With oCmd
                .CommandType = adCmdText
                '
                '## Set either normal/recursive or parent 'Link ID' SQL
                '   command with 'Link ID'.
                '
                .CommandText = Replace(oData.SQL( _
                            (Not (moDataShape.fGetIDTag(NodeKey) = oData.IDTag)) _
                            And Len(sParentID) > 0), "@@@", sParentID)
            End With
            .CursorLocation = ADODB.adUseClient
            .CacheSize = 1

            #If DEBUGMODE = 1 Then
                Debug.Print "SQL = " + oCmd.CommandText
            #End If
            '
            '## Execute SQL command
            '
            .Open oCmd, , ADODB.adOpenForwardOnly

            #If DEBUGMODE = 1 Then
                Debug.Print "  >>  Records = " + CStr(.RecordCount)
            #End If
            '
            '## Apply the Filer Criteria to the Recorset
            '
            .Filter = oData.FilterCriteria
            lCntRS = .RecordCount
            '
            '## Now loop tru all the records and load the Nodes into the TreeView structure
            '
            For lLoop = 1 To lCntRS

                #If DEBUGMODE = 1 Then
                    Debug.Print "@@@ Table (" + oData.TableName + _
                                 ") = " + oRS(oData.Fields(efldDesc).Desc).Value
                #End If

                lID = oRS(oData.Fields("ID").Desc).Value
                sKey = CStr(lID) + oData.IDTag
                '
                '## Node Forecolor
                '
                Select Case (oData.SQLBuildFields And esqlForeColor) = esqlForeColor
                    Case True:  lForeColor = CLng(oRS(oData.Fields(efldForeColor).Desc).Value)
                    Case False: lForeColor = oData.ForeColor
                End Select
                '
                '## Node Backcolor
                '
                Select Case (oData.SQLBuildFields And esqlBackColor) = esqlBackColor
                    Case True:  lBackColor = CLng(oRS(oData.Fields(efldBackColor).Desc).Value)
                    Case False: lBackColor = oData.BackColor
                End Select
                '
                '## Node Bold
                '
                Select Case (oData.SQLBuildFields And esqlBold) = esqlBold
                    Case True:  lBold = CLng(oRS(oData.Fields(efldBold).Desc).Value)
                    Case False: lBold = oData.Bold
                End Select
                '
                '## Retrieve node Image
                '
                vImage = pGetImageKey(oRS, _
                                      oData.Fields(efldImage).Desc, _
                                      oData.SQLBuildFields, _
                                      esqlImage, _
                                      oData.Image)
                '
                '## Retrieve node Selected Image
                '
                vSelectedImage = pGetImageKey(oRS, _
                                              oData.Fields(efldSelectedImage).Desc, _
                                              oData.SQLBuildFields, _
                                              esqlSelectedImage, _
                                              oData.SelectedImage)
                '
                '## Retrieve node Expanded Image
                '
                vExpandedImage = pGetImageKey(oRS, _
                                              oData.Fields(efldExpandedImage).Desc, _
                                              oData.SQLBuildFields, _
                                              esqlExpandedImage, _
                                              oData.ExpandedImage)
                '
                '## Add Node to TreeView control
                '
                #If DEBUGMODE = 1 Then
                    Select Case Len(NodeKey)
                        Case 0          '## Root TreeView Node
                            Set oNode = pAddNode(, , sKey, _
                                                     sKey + " - " + oRS(oData.Fields(efldDesc).Desc).Value, _
                                                     vImage, _
                                                     vSelectedImage, _
                                                     lForeColor, _
                                                     lBackColor, _
                                                     lBold, _
                                                     vExpandedImage)
                        Case Else       '## Child TreeView Node
                            Set oNode = pAddNode(NodeKey, _
                                                 eRelationship, _
                                                 sKey, _
                                                 sKey + " - " + oRS(oData.Fields(efldDesc).Desc).Value, _
                                                 vImage, _
                                                 vSelectedImage, _
                                                 lForeColor, _
                                                 lBackColor, _
                                                 lBold, _
                                                 vExpandedImage)
                    End Select
                #Else
                    Select Case Len(NodeKey)
                        Case 0          '## Root TreeView Node
                            Set oNode = pAddNode(, , sKey, _
                                                     oRS(oData.Fields(efldDesc).Desc).Value, _
                                                     vImage, _
                                                     vSelectedImage, _
                                                     lForeColor, _
                                                     lBackColor, _
                                                     lBold, _
                                                     vExpandedImage)
                        Case Else       '## Child TreeView Node
                            Set oNode = pAddNode(NodeKey, _
                                                 eRelationship, _
                                                 sKey, _
                                                 oRS(oData.Fields(efldDesc).Desc).Value, _
                                                 vImage, _
                                                 vSelectedImage, _
                                                 lForeColor, _
                                                 lBackColor, _
                                                 lBold, _
                                                 vExpandedImage)
                    End Select
                #End If
                '
                '## Check Data layout for possible child/sibling data layers
                '
                If pCheckForKids(oData) Then
                    '
                    '## Check DB for records
                    '
                    If pHasKids(sKey) Then
                        '
                        '## There are more records to be loaded
                        '   - set Node expand '+' image
                        '
                        oLod.ExpandIcon(oNode) = True
                        If Not moDataShape.LoadOnDemand Then
                            '
                            '## Force the node to be expanded and automatically load
                            '   records from the defined data structure
                            '
                            oNode.Expanded = True
                        End If
                    End If
                End If
                '
                '## Ok, now the next record...
                '
                oRS.MoveNext
            Next
        End With
    Next

End Sub

Public Function pFileExist(ByVal FileName As String) As Boolean

    On Error GoTo ErrorHandler
    pFileExist = (GetAttr(FileName) And vbDirectory) = 0
ErrorHandler:
    ' if an error occurs, this function returns False
End Function

Private Function pGetImageKey(ByVal oRS As ADODB.Recordset, _
                              ByVal FldKey As String, _
                              ByVal BldFlds As eSQLBuildFieldTypes, _
                              ByVal BldFld As eSQLBuildFieldTypes, _
                              ByVal DefaultImage As Variant) As Variant
    '
    '## Returns a valid Image Key & stops errors for invalid keys and out-of-bound indexes
    '
    Select Case (BldFlds And BldFld) = BldFld
        Case True                                   '## A Recordset Field holds the Image key
            With oRS(FldKey)
                Select Case .Type
                    Case adInteger, adNumeric       '## Numeric Image Key
                        pGetImageKey = CLng(.Value)
                    Case Else                       '## Alpha Image Key
                        pGetImageKey = CStr(.Value)
                End Select
            End With
            If Not pImageExist(pGetImageKey) Then   '## Check if the Image Key is valid
                pGetImageKey = DefaultImage         '## Nope - so use the Default Image Key
            End If
        Case False: pGetImageKey = DefaultImage     '## Use the Default Image Key

    End Select
    If Not IsEmpty(pGetImageKey) Then               '## Do we have a Image Key?
        If Not pImageExist(pGetImageKey) Then       '## Check if the Image Key is valid
            pGetImageKey = vbEmpty                  '## Image Key isn't valid
        End If
    End If

End Function

Private Function pHasKids(ByVal NodeKey As String) As Boolean

    Dim oRS       As ADODB.Recordset, _
        oCmd      As ADODB.Command, _
        oData     As cDataObj, _
        oInitkey  As String, _
        sParentID As String

    sParentID = CStr(Val(NodeKey))
    If Left$(NodeKey, Len(sParentID)) = sParentID Then

        Set oData = moDataShape.fGetItem(NodeKey)
        oInitkey = oData.Key
        
        If Not oData.Recursive Then
            If Len(oData.fChild) Then
                Set oData = moDataShape(oData.fChild)
            Else
                Exit Function
            End If
        End If

        Set oCmd = New ADODB.Command
        Set oCmd.ActiveConnection = moCon
        '
        '## Step through all the DataObj (Tables) for a particular layer and
        '   check if there are any children for each.
        '
        Do
            Set oRS = New ADODB.Recordset
            With oRS
                With oCmd
                    .CommandType = adCmdText
                    .CommandText = Replace(oData.SQL, "@@@", sParentID)
                End With
                .CursorLocation = ADODB.adUseClient
                .CacheSize = 1

                #If DEBUGMODE = 1 Then
                    Debug.Print "SQL = " + oCmd.CommandText
                #End If

                .Open oCmd, , ADODB.adOpenForwardOnly
                .Filter = oData.FilterCriteria

                #If DEBUGMODE = 1 Then
                    Debug.Print "  >>  Records = " + CStr(.RecordCount)
                #End If

                If .RecordCount Then
                    #If DEBUGMODE = 1 Then
                        Debug.Print "  >>   Result = " + CStr(.Fields(0).Value)
                    #End If
                    '
                    '## Are there any children?
                    '
                    If .Fields(0).Value > 0 Then
                        pHasKids = True
                        Exit Do
                    End If
                End If
            End With
            '
            '## Get the next DataObj (Table) to check
            '
            If oData.Key = oInitkey Then
                If Len(oData.fChild) Then
                    Set oData = moDataShape(oData.fChild)
                Else
                    Exit Do
                End If
            Else
                If Len(oData.fSibling) Then
                    Set oData = moDataShape(oData.fSibling)
                Else
                    Exit Do
                End If
            End If

        Loop

    End If

End Function

Private Sub pImagelist2Propertybag(ByVal oList As MSComctlLib.ImageList, _
                                   ByRef oBag As VBRUN.PropertyBag)
    '
    '## Stores an Imagelist in a PropertyBag object
    '
    Dim lNdx As Long

    With oBag
        .WriteProperty "BC", oList.BackColor
        .WriteProperty "IH", oList.ImageHeight
        .WriteProperty "IW", oList.ImageWidth
        .WriteProperty "MC", oList.MaskColor
        .WriteProperty "UMC", oList.UseMaskColor
        .WriteProperty "Tag", oList.Tag
        .WriteProperty "Cnt", oList.ListImages.Count
        For lNdx = 1 To oList.ListImages.Count
            With oList.ListImages(lNdx)
                oBag.WriteProperty "K" + CStr(lNdx), .Key, ""
                oBag.WriteProperty "T" + CStr(lNdx), .Tag, ""
                oBag.WriteProperty "I" + CStr(lNdx), .Picture
            End With
        Next
        Contents = .Contents
    End With

End Sub

Private Function pImageExist(ByVal vKey As Variant) As Boolean
    '
    '## Checks to see if an Image exists in the ImageList
    '
    On Error Resume Next
    pImageExist = CBool(oTree.ImageList.ListImages(vKey).Index)

End Function

Private Sub pPropertybag2Imagelist(ByVal oBag As VBRUN.PropertyBag, _
                                   ByRef oList As MSComctlLib.ImageList)
    '
    '## Restores an Imagelist from a PropertyBag object
    '
    Dim lNdx As Long

    With oBag
        oList.BackColor = .ReadProperty("BC")
        oList.ImageHeight = .ReadProperty("IH")
        oList.ImageWidth = .ReadProperty("IW")
        oList.MaskColor = .ReadProperty("MC")
        oList.UseMaskColor = .ReadProperty("UMC")
        oList.Tag = .ReadProperty("Tag")
        With oList.ListImages
            .Clear
            For lNdx = 1 To oBag.ReadProperty("Cnt")
                .Add(lNdx, _
                     oBag.ReadProperty("K" + CStr(lNdx), ""), _
                     oBag.ReadProperty("I" + CStr(lNdx)) _
                     ).Tag = oBag.ReadProperty("T", "")
            Next
        End With
    End With

End Sub
