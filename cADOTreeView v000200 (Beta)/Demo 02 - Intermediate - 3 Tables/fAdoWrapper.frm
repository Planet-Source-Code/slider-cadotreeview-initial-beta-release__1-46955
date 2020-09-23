VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form fAdoWrapper 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Demo02 : 3 Tables/Layers"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5145
   Icon            =   "fAdoWrapper.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   5145
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList imgDialog 
      Left            =   3045
      Top             =   3570
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fAdoWrapper.frx":12FA
            Key             =   "USA"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fAdoWrapper.frx":1454
            Key             =   "England"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fAdoWrapper.frx":15AE
            Key             =   "Australia"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fAdoWrapper.frx":1708
            Key             =   "City"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fAdoWrapper.frx":2412
            Key             =   "Hotel"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fAdoWrapper.frx":2864
            Key             =   "Select"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fAdoWrapper.frx":2CB6
            Key             =   "Country"
         EndProperty
      EndProperty
   End
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
Attribute VB_Name = "fAdoWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===========================================================================
'
' Form Name:    fADOWrapper [Demo02 - 3 Layers]
' Author:       Graeme Grant        (a.k.a. Slider)
' Date:         09/07/2003
' Version:      00.01.00 *BETA*
' Description:  Test/Demo 2 - Hotels by City by Country
' Edit History: 00.01.00 10/07/2003 Initial *BETA* Release
'
'===========================================================================

Option Explicit

'===========================================================================
' Private: Variables and Declarations
'
Private Const clCITYCOLOR  As Long = &H800080
Private Const clHOTELCOLOR As Long = &H40C0&

Private Const csDATABASE   As String = "\..\Demo.mdb"

#If NODLL = 0 Then
    Private moTreeDB       As vbADOTree.cADOTreeView
#Else
    Private moTreeDB       As cADOTreeView
#End If

'===========================================================================
' Form Events
'
Private Sub Form_Load()

    With tvwDialog
        '
        '## Setup Treeview control properties
        '
        .Style = tvwTreelinesPlusMinusPictureText
        .LineStyle = tvwRootLines
        .Indentation = 10
        .ImageList = imgDialog
        .FullRowSelect = False
        .HideSelection = False
        .HotTracking = True
        .LabelEdit = tvwManual
        .DragMode = vbManual
    End With

    Set moTreeDB = New cADOTreeView

    With moTreeDB
        .HookCtrl tvwDialog             '## Tell the wrapper which TreeView control to use
        pInit                           '## Define Tables to be used with relationships
        On Error GoTo ErrorHandler
        .Reload                         '## Now load the TreeView with Data
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

'===========================================================================
' Private subroutines and functions
'
Private Sub pInit()
        
    With moTreeDB
        '
        '## Define Tables to be used with relationships
        '
        '   Layer Country (Co) ........... [Recursive]
        '      |
        '      +- Layer City (Ci)
        '            |
        '            +- Layer Hotel (H)
        '
        ' Object         ID  MDB                 Sort   Parent         Build
        ' Key            Tag Table     Recursive Field  Object         Fields
        ' ==============+===+=========+=========+======+==============+==========================
        ' Layer Country  Co  Country      Yes    Desc   GRPSTK         Image
        ' Layer City     Ci  City         No     Desc   Layer Country  ---
        ' Layer Hotel    H   Hotel        No     Desc   Layer City     ---
        ' --------------+---+---------+---------+------+--------------+--------------------------
        '
        ' Object         |<-    Required Field/Column Names     ->|
        ' Key             efldID  efldDesc  efldLinkID  efldParent
        ' ===============+=======+=========+===========+==========+==============================
        ' Layer Country   PkID    Desc      LinkID      ---
        ' Layer City      PkID    Desc      LinkID      ---
        ' Layer Hotel     PkID    Desc      LinkID      ---
        ' ---------------+-------+---------+-----------+----------+------------------------------
        '
        ' Object Key     Property                   SQL Command (Pseudo Syntax)
        ' ==============+==========================+=============================================
        ' Layer Country  .Fields(efldImage).Sql     IIf(InStr([Desc],'Hemi'),'',[Desc]) AS NormImage
        ' --------------+--------------------------+---------------------------------------------
        '
        ' NOTES:
        ' ======
        '
        ' 1. Properties TableName, TableCritera, TableSort, and Fields(??).SQL all have standard
        '    SQL language syntax. These properties are applied to the SQL command text as
        '    follows:-
        '
        '    SELECT DISTINCTROW [Required: .Fields(efldID, efldDesc).SQL],
        '                       [Optional: .Fields(efldForeColor to efldExpandedImage).SQL]
        '           FROM        [.TableName]
        '           WHERE       [.Fields(efldLinkID or efldParent)).Desc]=@@@ AND [.TableCritera]
        '           ORDER BY    [.TableSort]
        '
        '    [.FilerCriteria] is then applied against the Recordset before the data is loaded
        '    into the TreeView control.
        '
        '    ** Currently the wrapper DOES NOT support JOINS. Dependant on feedback, this may be
        '       implemented in a later release. (I find votes @ PSC very incouraging ;) )
        '
        With .DataShape
            '
            '-----------------------------------------------------------------------------------
            '## Country Table (Recursive structure)
            '
            With .Add("Layer Country", "Country", "Co", "", True, , , True, "Country", "Select")
                '
                '## To minimise the loading delay of data, we don't use the TreeView
                '   control's Node sort property. By doing it at SQL level, it allows
                '   for custom sort orders without any complicated subclassing and API's.
                '   This property appears in the 'ORDER BY' clause of the automatically
                '   built SQL command text.
                '
                .TableSort = "[Desc]"
                '
                '## Advise which custom fields we'll be using. Please note that
                '   we're customising the Description (Node Text). So if you wish
                '   to custom columns in your SQL commands, then you *must* do it
                '   here.
                '
                .SQLBuildFields = esqlImage

                With .Fields
                    .Item("ID").Desc = "PkID"
                    .Item("Desc").Desc = "Desc"
                    .Item("Parent").Desc = "LinkID"
                    '
                    '## Assign Flags to each country. I've achieved this by using the
                    '   country name as the Image key - In a 'real world' application
                    '   you might want to use a seperate column in the table for this
                    '   purpose.
                    '
                    With .Item(efldImage)
                        .Desc = "NormImage"
                        .SQL = "IIf(InStr([Desc],'Hemi'),'',[Desc]) AS " + .Desc
                    End With
                End With
            End With
            '
            '-----------------------------------------------------------------------------------
            '## City Table (Links to 'Country Group' by Record ID)
            '
             With .Add("Layer City", "City", "Ci", "Layer Country", , clCITYCOLOR, , , "City", "Select")
                .TableSort = "[Desc]"

                With .Fields
                    .Item("ID").Desc = "PkID"
                    .Item("Desc").Desc = "Desc"
                    .Item("Parent").Desc = "LinkID"
                End With
            End With
            '
            '-----------------------------------------------------------------------------------
            '## Hotel Table (Links to 'City' by Record ID)
            '
             With .Add("Layer Hotel", "Hotel", "H", "Layer City", , clHOTELCOLOR, , , "Hotel", "Select")
                .TableSort = "[Desc]"

                With .Fields
                    .Item("ID").Desc = "PkID"
                    .Item("Desc").Desc = "Desc"
                    .Item("Parent").Desc = "LinkID"
                End With
            End With
            '
            '-----------------------------------------------------------------------------------
            '## Prepare connection to database
            '
            .ConnectString App.Path + csDATABASE, , , , ejvJet4
            '
            '-----------------------------------------------------------------------------------
            '## Comment the line below to turn off load on demand. It's worth doing just
            '   to see how much of an impact this feature has on the load time!
            '
            .LoadOnDemand = True

        End With
    End With

End Sub
