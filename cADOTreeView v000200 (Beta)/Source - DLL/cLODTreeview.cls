VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cLODTreeview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'===========================================================================
'
' Class Name:   cLODTreeview        [code subset of cTreeView v02.04.00]
' Author:       Graeme Grant        (a.k.a. Slider)
' Date:         10/07/2003
' Version:      00.01.00
' Description:  Load-On-Demand Management wrapper for the TreeView Control
' Edit History: 00.01.00 10/07/2003 Initial *BETA* Release
' Compiling:    DLL : set class Instancing = 5 - MultiUse
'               APP : set class Instancing = 1 - Private
'
' Notes:        NODLL conditional compilation flag is used in the project's
'               'Conditional Compiling Arguments' property (see
'               Demo?_NODLL.PRJ properties to see usage).
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
' Private: API Delclarations & variables
'
Private Type POINTAPI
  x As Long
  y As Long
End Type

Private Declare Function SendMessageAny Lib "user32" Alias "SendMessageA" _
                                        (ByVal hwnd As Long, _
                                         ByVal wMsg As Long, _
                                         ByVal wParam As Long, _
                                               lParam As Any) As Long

Private Declare Function SendMessageAAny Lib "user32" Alias "SendMessageA" _
                                        (ByVal hwnd As Long, _
                                         ByVal wMsg As Long, _
                                               wParam As Any, _
                                               lParam As Any) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
                                        (lpvDest As Any, _
                                         lpvSource As Any, _
                                         ByVal cbCopy As Long)

Private Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, ByVal lpRect As Long, ByVal bErase As Long) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Private Const TV_FIRST            As Long = &H1100

Private Const TVM_GETNEXTITEM     As Long = (TV_FIRST + 10)
Private Const TVM_GETITEM         As Long = (TV_FIRST + 12)
Private Const TVM_SETITEM         As Long = (TV_FIRST + 13)

Private Const TVN_FIRST           As Long = -400&
Private Const TVN_ITEMEXPANDINGA  As Long = (TVN_FIRST - 5)
Private Const TVN_ITEMEXPANDEDA   As Long = (TVN_FIRST - 6)

' TVM_GETNEXTITEM wParam values
Private Enum TVGN_Flags
    TVGN_ROOT = &H0
    TVGN_NEXT = &H1
    TVGN_PREVIOUS = &H2
    TVGN_PARENT = &H3
    TVGN_CHILD = &H4
    TVGN_FIRSTVISIBLE = &H5
    TVGN_NEXTVISIBLE = &H6
    TVGN_PREVIOUSVISIBLE = &H7
    TVGN_DROPHILITE = &H8
    TVGN_CARET = &H9
'#If (WIN32_IE >= &H400) Then   ' >= Comctl32.dll v4.71
    TVGN_LASTVISIBLE = &HA
'#End If
End Enum

Private Type TVITEM   ' was TV_ITEM
  Mask           As Long
  hItem          As Long
  State          As Long
  StateMask      As Long
  pszText        As Long    ' pointer
  cchTextMax     As Long
  iImage         As Long
  iSelectedImage As Long
  cChildren      As Long
  lParam         As Long
End Type

Private Enum TVITEM_mask
    TVIF_TEXT = &H1
    TVIF_IMAGE = &H2
    TVIF_PARAM = &H4
    TVIF_STATE = &H8
    TVIF_HANDLE = &H10
    TVIF_SELECTEDIMAGE = &H20
    TVIF_CHILDREN = &H40
#If (WIN32_IE >= &H400) Then   ' WIN32_IE = 1024 (>= Comctl32.dll v4.71)
    TVIF_INTEGRAL = &H80
#End If
    TVIF_DI_SETITEM = &H1000   ' Notification
End Enum

Private Type NMHDR
    hwndFrom As Long            '## Window handle of control sending message
    idfrom   As Long            '## Identifier of control sending message
    code     As Long            '## Specifies the notification code
End Type

Private Type NMTREEVIEW   ' was NM_TREEVIEW
  hdr As NMHDR
  ' Specifies a notification-specific action flag.
  ' Is NMTREEVIEW_action for TVN_SELCHANGING, TVN_SELCHANGED, TVN_SETDISPINFO
  ' Is TVM_EXPAND_wParam for TVN_ITEMEXPANDING, TVN_ITEMEXPANDED
  action As Long
  itemOld As TVITEM
  itemNew As TVITEM
  ptDrag As POINTAPI
End Type

'===========================================================================
' Private: Subclassing
'
Private Const WM_NOTIFY        As Long = &H4E

Private m_emr                  As EMsgResponse
Private moSubclass             As GSubclass
Private mlHwnd                 As Long
Private mlHWndParent           As Long

'===========================================================================
' Private: Internal Error Declarations
'
Private Const csSOURCE_ERR    As String = "cLODTreeview"
Private Const clTREEHOOK_ERR  As Long = vbObjectError + 2048 + 203
Private Const csTREEHOOK_ERR  As String = "Invalid control! Must be a VB6 TreeView control"

'===========================================================================
' Private: Variables and Declarations
'
Private WithEvents oTree As MSComctlLib.TreeView
Attribute oTree.VB_VarHelpID = -1

'===========================================================================
' Public: Event Declarations
'
Public Event Expanding(ByVal Node As MSComctlLib.Node, ByRef Cancel As Boolean)

'===========================================================================
' Subclassing
'
Friend Property Let MsgResponse(ByVal RHS As EMsgResponse)
    m_emr = RHS
End Property

Friend Property Get MsgResponse() As EMsgResponse
    MsgResponse = m_emr
End Property

Friend Function WindowProc(ByVal hwnd As Long, _
                           ByVal iMsg As Long, _
                           ByVal wParam As Long, _
                           ByVal lParam As Long) As Long

    Dim oNode  As MSComctlLib.Node, _
        Cancel As Boolean

    Select Case iMsg
        Case WM_NOTIFY
            Dim tNMH As NMHDR
            CopyMemory tNMH, ByVal lParam, Len(tNMH)
            If tNMH.hwndFrom = mlHwnd Then
                Select Case tNMH.code
                    Case TVN_ITEMEXPANDINGA
                        #If DEBUGMODE = 1 Then
                            Debug.Print "WM_NOTIFY::TVN_ITEMEXPANDING"
                        #End If
                        Dim tNMTV As NMTREEVIEW
                        CopyMemory tNMTV, ByVal lParam, Len(tNMTV)
                        If Not pHasRealKids(tNMTV.hdr.hwndFrom, tNMTV.itemNew.hItem) Then
                            Set oNode = pGetNodeFromhItem(tNMTV.hdr.hwndFrom, _
                                                          tNMTV.itemNew.lParam, _
                                                          tNMTV.itemNew.hItem)
                            RaiseEvent Expanding(oNode, Cancel)
                            If Cancel Then
                                m_emr = emrConsume
                            Else
                                m_emr = emrPostProcess
                            End If
                        End If

                    Case TVN_ITEMEXPANDEDA
                        #If DEBUGMODE = 1 Then
                            Debug.Print "WM_NOTIFY::TVN_ITEMEXPANDED"
                        #End If
                        m_emr = emrPostProcess
                    Case Else
                        m_emr = emrPostProcess
                End Select
                
            End If

    End Select

End Function

'===========================================================================
' Internal Class Subroutines
'
Private Sub Class_Initialize()
    '
End Sub

Private Sub Class_Terminate()
    UnHookCtrl
End Sub

'===========================================================================
' Public Subroutines and Functions
'
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
    With oTree
        mlHwnd = .hwnd
        mlHWndParent = .Parent.hwnd
        #If DEBUGMODE = 1 Then
            dbgCtrlName = .Parent.Name + "." + .Name
        #End If
    End With

    Set moSubclass = New GSubclass

    With moSubclass
        .AttachMessage Me, mlHWndParent, WM_NOTIFY
    End With

End Sub

Public Property Get ExpandIcon(Optional ByVal Node As MSComctlLib.Node) As Boolean

    Dim tvi   As TVITEM, _
        hItem As Long

    '## Get Node handle
    hItem = pGetNodeHandle(Node)
    '## Process if valid handle
    If hItem Then
        With tvi
            .hItem = hItem
            .Mask = TVIF_CHILDREN
            '.cChildren = Abs(hasChildNodes)
        End With
        '## Old docs say returns 0 on success, -1 on failure. New docs say
        '   returns TRUE if successful, or FALSE otherwise we'll go new...
        'ExpandIcon = SendMessageAny(oTree.hwnd, TVM_GETITEM, 0&, tvi)
        SendMessageAny oTree.hwnd, TVM_GETITEM, 0&, tvi
        ExpandIcon = (tvi.cChildren > 0)
    End If

End Property

Public Property Let ExpandIcon(Optional ByVal Node As MSComctlLib.Node, ByVal SetState As Boolean)

    Dim tvi   As TVITEM, _
        hItem As Long

    '## Get Node handle
    hItem = pGetNodeHandle(Node)
    '## Process if valid handle
    If hItem Then
        With tvi
            .hItem = hItem
            .Mask = TVIF_CHILDREN
            .cChildren = Abs(SetState)
        End With
        '## Old docs say returns 0 on success, -1 on failure. New docs say
        '   returns TRUE if successful, or FALSE otherwise we'll go new...
        SendMessageAny oTree.hwnd, TVM_SETITEM, 0&, tvi
    End If

End Property

Public Sub UnHookCtrl()

    '
    '## Release the hooked control
    '
    Set oTree = Nothing
    '
    '## Shut down subclassing
    '
    If Not (moSubclass Is Nothing) Then
        With moSubclass
            .DetachMessage Me, mlHWndParent, WM_NOTIFY
        End With
    End If

End Sub

'===========================================================================
' General: Private subroutines and functions
'
Private Function pGetItemFromNode(ByVal Node As MSComctlLib.Node) As Long
    '
    '## If successful, returns the treeview item handle represented by
    '   the specified Node, returns 0 otherwise.
    '
    Dim anSiblingPos() As Integer  '## Contains the sibling position of the node and all it's parents
    Dim iLevel         As Integer  '## Hierarchical level of the node
    Dim hItem          As Long, _
        iloop          As Integer, _
        iPos           As Integer

    '
    '## Continually work backwards from the current node to the current node's
    '   first sibling, caching the current node's sibling position in the one-based
    '   array. Then get the first sibling's parent node and start over. Keep going
    '   until the postion of the specified node's top level parent item is obtained...
    '
    Do While (Node Is Nothing) = False
        iLevel = iLevel + 1
        ReDim Preserve anSiblingPos(iLevel)
        anSiblingPos(iLevel) = pGetNodeSiblingPos(Node)
        Set Node = Node.Parent
    Loop

    '## Get the hItem of the first item in the treeview
    hItem = SendMessageAAny(oTree.hwnd, TVM_GETNEXTITEM, ByVal TVGN_ROOT, 0&)
    If hItem Then
        '
        '## Now work backwards through the cached node positions in the array
        '   (from the first treeview node to the specified node), obtaining the respective
        '   item handle for each node at the cached position. When we get to the
        '   specified node's position (the value of the first element in the array), we
        '   got it's hItem...
        '
        For iloop = iLevel To 1 Step -1
            iPos = anSiblingPos(iloop)
            Do While iPos > 1
                hItem = SendMessageAAny(oTree.hwnd, TVM_GETNEXTITEM, ByVal TVGN_NEXT, ByVal hItem)
                iPos = iPos - 1
            Loop
            If (iloop > 1) Then hItem = SendMessageAAny(oTree.hwnd, TVM_GETNEXTITEM, ByVal TVGN_CHILD, ByVal hItem)
        Next
        pGetItemFromNode = hItem
    End If

End Function

Private Function pGetNodeFromhItem(ByVal hwnd As Long, ByVal lParam As Long, ByVal hItem As Long) As MSComctlLib.Node

    Dim pNode  As Long, _
        oNode As MSComctlLib.Node

    If lParam Then
        CopyMemory pNode, ByVal lParam + 8, 4
        CopyMemory oNode, pNode, 4
        Set pGetNodeFromhItem = oNode
        CopyMemory oNode, 0&, 4
    End If

End Function

Private Function pGetNodeHandle(Optional ByVal Node As MSComctlLib.Node) As Long

    If (Node Is Nothing) Then
        pGetNodeHandle = SendMessageAny(oTree.hwnd, TVM_GETNEXTITEM, TVGN_CARET, ByVal 0&)
    Else
        pGetNodeHandle = pGetItemFromNode(Node)
    End If

End Function

Private Function pGetNodeSiblingPos(ByVal Node As MSComctlLib.Node) As Integer
    '
    '## Returns the one-base position of the specified node
    '   with respect to it's sibling order.
    '
    Dim nPos As Integer

    '## Keep counting up from one until the node has no more previous siblings
    Do While (Node Is Nothing) = False
        nPos = nPos + 1
        Set Node = Node.Previous
    Loop
    pGetNodeSiblingPos = nPos

End Function

Private Function pHasRealKids(ByVal hwnd As Long, ByVal hItem As Long) As Boolean
    '
    '## Tests to see if there is a child Node
    '
    pHasRealKids = Not (SendMessageAny(hwnd, TVM_GETNEXTITEM, ByVal TVGN_CHILD, ByVal hItem) = 0)

End Function
