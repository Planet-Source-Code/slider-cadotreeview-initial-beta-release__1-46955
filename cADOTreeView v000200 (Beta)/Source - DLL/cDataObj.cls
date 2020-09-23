VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDataObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'===========================================================================
'
' Class Name:   cDataObj
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
' Public Declarations
'
Public Enum eFields
    efldID = 1
    efldDesc = 2
    efldLinkID = 3
    efldParent = 4
    efldForeColor = 5
    efldBackColor = 6
    efldBold = 7
    efldImage = 8
    efldSelectedImage = 9
    efldExpandedImage = 10
End Enum

Public Enum eSQLBuildFieldTypes     '## Defines which node properties are
    esqlForeColor = 1               '   to be included in the SQL command text
    esqlBackColor = 2
    esqlBold = 4
    esqlImage = 8
    esqlSelectedImage = &H10
    esqlExpandedImage = &H20
    esqlDesc = &H40
End Enum

'===========================================================================
' Private: Variables and Declarations
'
Private msKey            As String, _
        mbRecursive      As Boolean, _
        msTableName      As String, _
        msTableCriteria  As String, _
        msTableSort      As String, _
        msIDTag          As String, _
        mlForeColor      As OLE_COLOR, _
        mlBackColor      As OLE_COLOR, _
        mbBold           As Boolean, _
        moFields         As cFields, _
        msParentKey      As String, _
        msSibling        As String, _
        msChild          As String

Private msFilterCriteria As String, _
        mvImage          As Variant, _
        mvSelectedImage  As Variant, _
        mvExpandedImage  As Variant

Private meSQLBuildFields As eSQLBuildFieldTypes

'===========================================================================
' Internal Class Subroutines
'
Private Sub Class_Initialize()
    Set moFields = New cFields
    '
    '## Not a publically creatable collection. This was done as there is no
    '   public flexibity with the ordering/inclusion/exclusion of the
    '   collection.
    '
    With moFields
        .Add("ID").Desc = "ID"
        .Add("Desc").Desc = "Desc"
        .Add("LinkID").Desc = "LinkID"
        .Add("Parent").Desc = ""
        .Add("ForeColor").Desc = ""
        .Add("BackColor").Desc = ""
        .Add("Bold").Desc = ""
        .Add("Image").Desc = ""
        .Add("SelectedImage").Desc = ""
        .Add("ExpandedImage").Desc = ""
    End With
End Sub

Private Sub Class_Terminate()
    Set moFields = Nothing
End Sub

'===========================================================================
' Internally shared properties/methods.
'
Friend Property Get fSibling() As String
    fSibling = msSibling
End Property

Friend Property Let fSibling(ByVal vNewValue As String)
    msSibling = vNewValue
End Property

Friend Property Get fChild() As String
    fChild = msChild
End Property

Friend Property Let fChild(ByVal vNewValue As String)
    msChild = vNewValue
End Property

'===========================================================================
' Public Properties
'
Public Property Get BackColor() As OLE_COLOR
    BackColor = mlBackColor
End Property

Public Property Let BackColor(ByVal vNewValue As OLE_COLOR)
    mlBackColor = vNewValue
End Property

Public Property Get Bold() As Boolean
    Bold = mbBold
End Property

Public Property Let Bold(ByVal vNewValue As Boolean)
    mbBold = vNewValue
End Property

Public Property Get Contents() As Byte()
    '
    '## Returns the contents of the object as a serialized byte array.
    '
    Dim oBag As VBRUN.PropertyBag
    
    Set oBag = New VBRUN.PropertyBag
    With oBag
        .WriteProperty "K", msKey, ""
        .WriteProperty "Re", mbRecursive, False
        .WriteProperty "TN", msTableName, ""
        .WriteProperty "TC", msTableCriteria, ""
        .WriteProperty "TS", msTableSort, ""
        .WriteProperty "ID", msIDTag, ""
        .WriteProperty "FC", mlForeColor, vbWindowText
        .WriteProperty "BC", mlBackColor, vbWindowBackground
        .WriteProperty "Bo", mbBold, False
        .WriteProperty "BF", meSQLBuildFields, 0
        .WriteProperty "Fs", moFields.Contents
        .WriteProperty "PK", msParentKey, ""
        .WriteProperty "Fi", msFilterCriteria, ""
        .WriteProperty "Im", mvImage
        .WriteProperty "SI", mvSelectedImage
        .WriteProperty "EI", mvExpandedImage
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
        msKey = .ReadProperty("K", "")
        mbRecursive = .ReadProperty("Re", False)
        msTableName = .ReadProperty("TN", "")
        msTableCriteria = .ReadProperty("TC", "")
        msTableSort = .ReadProperty("TS", "")
        msIDTag = .ReadProperty("ID", "")
        mlForeColor = .ReadProperty("FC", mlForeColor)
        mlBackColor = .ReadProperty("BC", mlBackColor)
        mbBold = .ReadProperty("Bo", False)
        meSQLBuildFields = .ReadProperty("BF", 0)
        moFields.Contents = .ReadProperty("Fs", "")
        msParentKey = .ReadProperty("PK", "")
        msFilterCriteria = .ReadProperty("Fi", "")
        mvImage = .ReadProperty("Im")
        mvSelectedImage = .ReadProperty("SI")
        mvExpandedImage = .ReadProperty("EI")
    End With

End Property

Public Property Get ExpandedImage() As Variant
    ExpandedImage = mvExpandedImage
End Property

Public Property Let ExpandedImage(ByVal vNewID As Variant)
    mvExpandedImage = vNewID
End Property

Public Property Get Fields() As cFields
    Set Fields = moFields
End Property

Public Property Get FilterCriteria() As String
    FilterCriteria = msFilterCriteria
End Property

Public Property Let FilterCriteria(ByVal vNewValue As String)
    msFilterCriteria = vNewValue
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = mlForeColor
End Property

Public Property Let ForeColor(ByVal vNewValue As OLE_COLOR)
    mlForeColor = vNewValue
End Property

Public Property Get IDTag() As String
    IDTag = msIDTag
End Property

Public Property Let IDTag(ByVal vNewValue As String)
    msIDTag = vNewValue
End Property

Public Property Get Image() As Variant
    Image = mvImage
End Property

Public Property Let Image(ByVal vNewID As Variant)
    mvImage = vNewID
End Property

Public Property Get Key() As String
    Key = msKey
End Property

Public Property Let Key(ByVal vNewValue As String)
    msKey = vNewValue
End Property

Public Property Get ParentKey() As String
    ParentKey = msParentKey
End Property

Public Property Let ParentKey(ByVal vNewValue As String)
    msParentKey = vNewValue
End Property

Public Property Get Recursive() As Boolean
    Recursive = mbRecursive
End Property

Public Property Let Recursive(ByVal vNewValue As Boolean)
    mbRecursive = vNewValue
End Property

Public Property Get SelectedImage() As Variant
    SelectedImage = mvSelectedImage
End Property

Public Property Let SelectedImage(ByVal vNewID As Variant)
    mvSelectedImage = vNewID
End Property

Public Property Get SQL(Optional ByVal UseParent As Boolean = False) As String
    '
    '## Returns a built SQL query string
    '
    Dim sSQL     As String, _
        sFilter  As String, _
        sDescFld As String

    With moFields
        Select Case (meSQLBuildFields And esqlDesc) = esqlDesc
            Case True:  sDescFld = .Item(efldDesc).SQL
            Case False: sDescFld = .Item(efldDesc).Desc
        End Select

        sSQL = "SELECT DISTINCTROW " + .Item(efldID).Desc + ", " + _
                                       sDescFld + _
                                       pSQLBuildFields + _
               " FROM [" + msTableName + "] WHERE "
        If UseParent And (Len(moFields(efldParent)) > 0) Then
           sSQL = sSQL + moFields(efldParent) + "=@@@"
        Else
            sSQL = sSQL + moFields(efldLinkID) + "=@@@"
        End If
        sFilter = Trim$(msTableCriteria)
        If Len(sFilter) Then
            sSQL = sSQL + " AND " + sFilter
        End If
        If Len(Trim$(msTableSort)) Then
            sSQL = sSQL + " ORDER BY " + msTableSort
        End If
    End With

    SQL = sSQL

End Property

Public Property Get SQLBuildFields() As eSQLBuildFieldTypes
    SQLBuildFields = meSQLBuildFields
End Property

Public Property Let SQLBuildFields(ByVal vNewValue As eSQLBuildFieldTypes)
    meSQLBuildFields = vNewValue
End Property

Public Property Get TableName() As String
    TableName = msTableName
End Property

Public Property Let TableName(ByVal vNewValue As String)
    msTableName = vNewValue
End Property

Public Property Get TableCriteria() As String
    TableCriteria = msTableCriteria
End Property

Public Property Let TableCriteria(ByVal vNewValue As String)
    msTableCriteria = vNewValue
End Property

Public Property Get TableSort() As String
    TableSort = msTableSort
End Property

Public Property Let TableSort(ByVal vNewValue As String)
    msTableSort = vNewValue
End Property

'===========================================================================
' General: Private subroutines and functions
'
Private Function pSQLBuildFields() As String

    Dim lCustomSqlType As Long, _
        lLoop          As Long

    With moFields
        For lLoop = efldForeColor To efldExpandedImage
            lCustomSqlType = 2 ^ (lLoop - efldForeColor)
            If (meSQLBuildFields And lCustomSqlType) = lCustomSqlType Then
                pSQLBuildFields = pSQLBuildFields + _
                                   IIf(Len(pSQLBuildFields) And Len(.Item(lLoop).SQL), _
                                       ", ", "") + _
                                   .Item(lLoop).SQL
            End If
        Next
        If Len(pSQLBuildFields) Then pSQLBuildFields = ", " + pSQLBuildFields
    End With

End Function
