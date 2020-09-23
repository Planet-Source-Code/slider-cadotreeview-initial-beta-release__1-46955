VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form fAdoWrapper 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Demo 05 : Minimal Code - Predefined DataShape"
   ClientHeight    =   6810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5115
   Icon            =   "fAdoWrapper.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   5115
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList imgDialog 
      Left            =   3000
      Top             =   3600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fAdoWrapper.frx":12FA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwDialog 
      Height          =   6630
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4950
      _ExtentX        =   8731
      _ExtentY        =   11695
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   18
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      BorderStyle     =   1
      Appearance      =   0
   End
End
Attribute VB_Name = "fAdoWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===========================================================================
'
' Form Name:    fAdoWrapper [Demo04 - Advanced]
' Author:       Graeme Grant        (a.k.a. Slider)
' Date:         16/07/2003
' Version:      00.01.00
' Description:  Test/Demo 5 - Minimal code with predefined relationships
' Edit History: 00.01.00 16/07/2003 Initial *BETA* Release
'
'===========================================================================

Option Explicit

'===========================================================================
' Private: Variables and Declarations
'
#If NODLL = 0 Then
    Private WithEvents moTreeDB As vbADOTree.cADOTreeView
Attribute moTreeDB.VB_VarHelpID = -1
#Else
    Private WithEvents moTreeDB As cADOTreeView
Attribute moTreeDB.VB_VarHelpID = -1
#End If

'
'## Uncomment below the Datashape Relationship & Images file that you wish to view
'
'Private Const csSHAPELOAD As String = "\Stock.DS"
'Private Const csSHAPELOAD As String = "\Travel.DS"
Private Const csSHAPELOAD As String = "\Complex.DS"

Private Const csDATABASE  As String = "\..\Demo.mdb"

'===========================================================================
' Form Events
'
Private Sub Form_Load()

    On Error GoTo ErrorHandler

    tvwDialog.ImageList = imgDialog         '## Point to the Imagelist control
                                            '   (Must be initialised with at least one image)
    Set moTreeDB = New cADOTreeView         '## Initialise wrapper
    With moTreeDB
        .HookCtrl tvwDialog                 '## Tell the wrapper which TreeView control to use
        .LoadShape App.Path + csSHAPELOAD   '## Load from file the Relationships & Images to be used
        .DataShape.ConnectString App.Path + csDATABASE, , , , ejvJet4   '## Point to the Database
        .Reload                             '## Now load the TreeView with Data
    End With

Exit Sub

ErrorHandler:
    MsgBox "Problem encountered defining/connecting/loading data.", _
           vbCritical + vbOKOnly + vbDefaultButton1 + vbApplicationModal, _
           "Critical cADOTreeView (*BETA RELEASE*) Error!"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set moTreeDB = Nothing
End Sub
