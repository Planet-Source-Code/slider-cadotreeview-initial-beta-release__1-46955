VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form fAdoWrapper 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Demo 03 : Intermediate - 4 Tables & Build Fields"
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
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fAdoWrapper.frx":12FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fAdoWrapper.frx":1894
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fAdoWrapper.frx":1E2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fAdoWrapper.frx":23C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fAdoWrapper.frx":26E2
            Key             =   ""
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
' Form Name:    fAdoWrapper [Demo02 - 3 Layers]
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
    Private moTreeDB As vbADOTree.cADOTreeView
#Else
    Private moTreeDB As cADOTreeView
#End If

'===========================================================================
' Form Events
'
Private Sub Form_Load()

    With tvwDialog
        .Style = tvwTreelinesPlusMinusPictureText
        .LineStyle = tvwRootLines
        .Indentation = 10
        .ImageList = imgDialog
        .FullRowSelect = False
        .HideSelection = False
        .HotTracking = True
        .LabelEdit = tvwManual
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
        '   +- GRPPROD (GP) ....................... [Recursive]
        '   |     |
        '   |     +- PROD (PR)
        '   |
        '   +- GRPSTK (GT) ........................ [Recursive]
        '   |     |
        '   |     +- STK (ST)
        '   |     |
        '   +- GRPSUP (GU) ........................ [Recursive]
        '         |
        '         +- SUP (SU)
        '
        ' Object         ID  MDB                 Sort   Parent         Build
        ' Key            Tag Table     Recursive Field  Object         Fields
        ' ==============+===+=========+=========+======+==============+==========================
        ' GRPPROD        GP  GroupB       Yes    Desc   ---            Desc, Fore/Back Color
        ' PROD           PR  ProductB     No     Desc   GRPPROD        ---
        ' GRPSTK         GT  GroupB       Yes    Desc   ---            ---
        ' STK            ST  StockB       No     Desc   GRPSTK         ---
        ' GRPSUP         GU  GroupB       Yes    Desc   GRPSTK         ---
        ' SUP            SU  SupplierB    No     Desc   GRPSUP         ---
        ' --------------+---+---------+---------+------+--------------+--------------------------
        '
        ' Object         |<-    Required Field/Column Names     ->|
        ' Key             efldID  efldDesc  efldLinkID  efldParent
        ' ===============+=======+=========+===========+==========+==============================
        ' GRPPROD         PkID    KeyDesc   GroupID     ---
        ' PROD            PkID    Desc      GroupID     ---
        ' GRPSTK          PkID    Desc      GroupID     ---
        ' STK             PkID    Desc      GroupID     ---
        ' GRPSUP          PkID    Desc      GroupID     ---
        ' SUP             PkID    Desc      GroupID     ---
        ' ---------------+-------+---------+-----------+----------+------------------------------
        '
        ' Object Key     Property                   SQL Command (Pseudo Syntax)
        ' ==============+==========================+=============================================
        ' GRPPROD        .TableCriteria             Type=0
        ' GRPPROD        .Fields(efldDesc).Sql      (CStr(PkID) + ' - ' + Desc) AS KeyDesc
        ' GRPPROD        .Fields(efldForeColor).Sql (IIf([GroupID] > 9, vbRed, vbMagenta) AS ForeColor
        ' GRPPROD        .Fields(efldBackColor).Sql (IIf([GroupID] > 9, vbYellow, vbCyan) AS BackColor
        ' GRPSTK         .TableCriteria             Type=1
        ' GRPSUP         .TableCriteria             Type=4
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
        ' 2. If you look closely at GRPPROD, GRPSTK, & GRPSUP, you'll notice that They all use
        '    the same table. Therefore, by using the TableCriteria property of a DataObj to break
        '    the table elements into seperate groups of data, it's possible to set an unlimited
        '    number of definitions against a single table - e.g. Alphabetize Phonebook entries;
        '    color coded ranges; Images as warning indicators; color coded regions, etc...
        '    "... a picture is worth a thousand words."
        '
        With .DataShape
            '
            '-----------------------------------------------------------------------------------
            '## Product Group Table (Recursive structure)
            '
            With .Add("GRPPROD", "Groupb", "GP", "", True, vbRed, , True, 1, 3, 2)
                '
                '## I've used several Group Types in the one table. So here I'm
                '   designating the group type (0 = Product, 1 = Stock, 4 = Supplier).
                '   This property appears in the 'WHERE' clause of the automatically
                '   built SQL command text.
                '
                .TableCriteria = "Type=0"
                '
                '## To minimise the loading delay of data, we don't use the TreeView
                '   control's Node sort property. By doing it at SQL level, it allows
                '   for custom sort orders without any complicated subclassing or APIs.
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
                .SQLBuildFields = esqlBackColor + esqlForeColor + esqlDesc

                With .Fields
                    '
                    '## Define the Fields in the table and any custom SQL requirements
                    '
                    .Item(efldID).Desc = "PkID"
                    With .Item(efldDesc)
                        '
                        '## We're creating a custom column (Field) name
                        '
                        .Desc = "KeyDesc"
                        '
                        '## Column data is based on joining two table columns
                        '
                        .SQL = "(CStr(PkID) + "" - "" + Desc) AS " + .Desc
                    End With
                    .Item(efldLinkID).Desc = "GroupID"
                    '
                    '## Conditional Node Fore/Back color using SQL
                    '
                    With .Item(efldForeColor)
                        .Desc = "ForeColor"
                        .SQL = "(IIf([GroupID] > 9, " + CStr(vbRed) + ", " + _
                                                        CStr(vbMagenta) + ")) AS " + .Desc
                    End With
                    With .Item(efldBackColor)
                        .Desc = "BackColor"
                        .SQL = "(IIf([GroupID] > 9, " + CStr(vbYellow) + ", " + _
                                                        CStr(vbCyan) + ")) AS " + .Desc
                    End With
                End With
            End With
            '
            '-----------------------------------------------------------------------------------
            '## Product Table (Links to 'Product Group' by Record ID)
            '
            With .Add("PROD", "ProductB", "PR", "GRPPROD", , , , , 4, 5)
                .TableSort = "[Desc]"

                With .Fields
                    .Item(efldID).Desc = "PkID"
                    .Item(efldDesc).Desc = "Desc"
                    .Item(efldLinkID).Desc = "GroupID"
                End With
            End With
            '
            '-----------------------------------------------------------------------------------
            '## Stock Group Table (Recursive structure)
            '
            With .Add("GRPSTK", "GroupB", "GT", "", True, vbBlue, , True, 1, 3, 2)
                .TableCriteria = "Type=1"
                .TableSort = "[Desc]"

                With .Fields
                    .Item(efldID).Desc = "PkID"
                    .Item(efldDesc).Desc = "Desc"
                    .Item(efldLinkID).Desc = "GroupID"
                End With
            End With
            '
            '-----------------------------------------------------------------------------------
            '## Stock Table (Links to 'Stock Group' by Record ID)
            '
            With .Add("STK", "StockB", "ST", "GRPSTK", , , , , 4, 5)
                .TableSort = "[Desc]"

                With .Fields
                    .Item(efldID).Desc = "PkID"
                    .Item(efldDesc).Desc = "Desc"
                    .Item(efldLinkID).Desc = "GroupID"
                End With
            End With
            '
            '-----------------------------------------------------------------------------------
            '## Supplier Group Table (Recursive structure) (Links to 'Stock Group' by Record ID)
            '
            With .Add("GRPSUP", "GroupB", "GU", "", True, vbGreen, , True, 1, 3, 2)
                .TableCriteria = "Type=4"
                .TableSort = "[Desc]"

                With .Fields
                    .Item(efldID).Desc = "PkID"
                    .Item(efldDesc).Desc = "Desc"
                    .Item(efldLinkID).Desc = "GroupID"
                End With
            End With
            '
            '-----------------------------------------------------------------------------------
            '## Supplier Table (Links to 'Supplier Group' by Record ID)
            '
            With .Add("SUP", "SupplierB", "SU", "GRPSUP", , , , , 4, 5)
                .TableSort = "[Desc]"

                With .Fields
                    .Item(efldID).Desc = "PkID"
                    .Item(efldDesc).Desc = "Desc"
                    .Item(efldLinkID).Desc = "GroupID"
                End With
            End With
            '
            '-----------------------------------------------------------------------------------
            '## Prepare connection to database
            '
            .ConnectString App.Path + "\..\Demo.mdb", , , , ejvJet4
            '
            '-----------------------------------------------------------------------------------
            '## Comment the line below to turn off load on demand. It's worth doing just
            '   to see how much of an impact this feature has on the load time!
            '
            .LoadOnDemand = True

        End With
    End With

End Sub
