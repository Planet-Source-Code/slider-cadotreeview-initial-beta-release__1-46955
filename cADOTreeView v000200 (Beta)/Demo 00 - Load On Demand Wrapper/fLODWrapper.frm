VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form fLODWrapper 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Demo 00 : Load On Demand (No ADO)"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5145
   Icon            =   "fLODWrapper.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   5145
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.TreeView tvwDialog 
      Height          =   6630
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   4950
      _ExtentX        =   8731
      _ExtentY        =   11695
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
End
Attribute VB_Name = "fLODWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===========================================================================
'
' Form Name:    fTreeHotel [Demo02 - 3 Layers]
' Author:       Graeme Grant        (a.k.a. Slider)
' Date:         10/07/2003
' Version:      00.01.00
' Description:  Test/Demo 2 - Hotels by City by Country
' Edit History: 00.01.00 10/07/2003 Initial *BETA* Release
'
'===========================================================================

Option Explicit

'===========================================================================
' Private: Variables and Declarations
'
#If NODLL = 0 Then
    Private WithEvents moTreeLOD As vbADOTree.cLODTreeview
Attribute moTreeLOD.VB_VarHelpID = -1
#Else
    Private WithEvents moTreeLOD As cLODTreeview
Attribute moTreeLOD.VB_VarHelpID = -1
#End If

'===========================================================================
' cLODTreeView Events
'
Private Sub moTreeLOD_Expanding(ByVal Node As MSComctlLib.Node, ByRef Cancel As Boolean)
    pLoadNodes Node
End Sub

'===========================================================================
' Form Events
'
Private Sub Form_Load()

    With tvwDialog
        '.Style = tvwTreelinesPlusMinusPictureText
        .Style = tvwTreelinesPlusMinusText
        .LineStyle = tvwRootLines
        .Indentation = 10
        '.ImageList = imgDialog
        .FullRowSelect = False
        .HideSelection = False
        .HotTracking = True
        .LabelEdit = tvwManual
    End With

    Set moTreeLOD = New cLODTreeview
    '
    '## Tell the wrapper which TreeView control to use
    '
    moTreeLOD.HookCtrl tvwDialog
    pLoadNodes

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set moTreeLOD = Nothing
End Sub

'===========================================================================
' Private subroutines and functions
'
Private Sub pLoadNodes(Optional ByRef oNode As MSComctlLib.Node)

    Dim lLoop As Long
    Dim lFrom As Long

    With tvwDialog.Nodes
        If .Count Then
            lFrom = .Count + 1
        Else
            lFrom = 1
        End If
        For lLoop = lFrom To lFrom + 9
            Select Case lFrom
                Case 1
                    moTreeLOD.ExpandIcon(.Add(, , CStr(lLoop) + "A", "Node " + CStr(lLoop))) = True
                Case Else
                    moTreeLOD.ExpandIcon(.Add(oNode, tvwChild, CStr(lLoop) + "A", "Node " + CStr(lLoop))) = True
            End Select
        Next
    End With
End Sub
