VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form fAdoWrapper 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Demo 04 : Advanced - 9 Tables, Grouping, Filtering, Conditional dynamic Node properties"
   ClientHeight    =   7065
   ClientLeft      =   3285
   ClientTop       =   1815
   ClientWidth     =   9405
   Icon            =   "fAdoWrapper.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   9405
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraDetails 
      Caption         =   "Default Images:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Index           =   0
      Left            =   7440
      TabIndex        =   32
      Top             =   4320
      Width           =   1815
      Begin VB.Label lblDetail 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800000&
         Caption         =   "Expanded: "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   300
         Index           =   16
         Left            =   120
         TabIndex        =   35
         Top             =   885
         Width           =   1275
      End
      Begin VB.Label lblDetail 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800000&
         Caption         =   "Standard: "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   300
         Index           =   15
         Left            =   120
         TabIndex        =   34
         Top             =   255
         Width           =   1275
      End
      Begin VB.Label lblDetail 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800000&
         Caption         =   "Selected: "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   300
         Index           =   14
         Left            =   120
         TabIndex        =   33
         Top             =   570
         Width           =   1275
      End
      Begin VB.Image imgDetails 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   2
         Left            =   1455
         Stretch         =   -1  'True
         Top             =   900
         Width           =   240
      End
      Begin VB.Image imgDetails 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   1
         Left            =   1455
         Stretch         =   -1  'True
         Top             =   570
         Width           =   240
      End
      Begin VB.Image imgDetails 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   0
         Left            =   1455
         Stretch         =   -1  'True
         Top             =   240
         Width           =   240
      End
   End
   Begin VB.CommandButton cmdFile 
      Caption         =   "&Load"
      Height          =   375
      Index           =   1
      Left            =   7680
      TabIndex        =   31
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdFile 
      Caption         =   "&Save"
      Height          =   375
      Index           =   0
      Left            =   8520
      TabIndex        =   30
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtSQL 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   2310
      Left            =   5085
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   1470
      Width           =   4215
   End
   Begin VB.CheckBox chkIcons 
      Appearance      =   0  'Flat
      Caption         =   "Show Node Icons"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   120
      TabIndex        =   28
      Top             =   6645
      Value           =   1  'Checked
      Width           =   3495
   End
   Begin MSComctlLib.ImageList imgDialog 
      Left            =   3720
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   22
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fAdoWrapper.frx":27A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fAdoWrapper.frx":2D3C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fAdoWrapper.frx":32D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fAdoWrapper.frx":3870
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fAdoWrapper.frx":3E0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fAdoWrapper.frx":43A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fAdoWrapper.frx":46BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fAdoWrapper.frx":49D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fAdoWrapper.frx":4E2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fAdoWrapper.frx":4F84
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fAdoWrapper.frx":70BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fAdoWrapper.frx":7658
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fAdoWrapper.frx":D27A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fAdoWrapper.frx":D814
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fAdoWrapper.frx":DDAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fAdoWrapper.frx":F540
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fAdoWrapper.frx":14D32
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fAdoWrapper.frx":15184
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fAdoWrapper.frx":155D6
            Key             =   "Co"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fAdoWrapper.frx":15A28
            Key             =   "England"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fAdoWrapper.frx":15B82
            Key             =   "Australia"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fAdoWrapper.frx":15CDC
            Key             =   "USA"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwDialog 
      Height          =   6475
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   3510
      _ExtentX        =   6191
      _ExtentY        =   11430
      _Version        =   393217
      Style           =   7
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame fraDetails 
      Caption         =   "Node Images:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Index           =   1
      Left            =   7440
      TabIndex        =   36
      Top             =   5640
      Width           =   1815
      Begin VB.Image imgDetails 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   3
         Left            =   1455
         Stretch         =   -1  'True
         Top             =   240
         Width           =   240
      End
      Begin VB.Image imgDetails 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   4
         Left            =   1455
         Stretch         =   -1  'True
         Top             =   570
         Width           =   240
      End
      Begin VB.Image imgDetails 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   5
         Left            =   1455
         Stretch         =   -1  'True
         Top             =   900
         Width           =   240
      End
      Begin VB.Label lblDetail 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800000&
         Caption         =   "Selected: "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   300
         Index           =   19
         Left            =   120
         TabIndex        =   39
         Top             =   570
         Width           =   1275
      End
      Begin VB.Label lblDetail 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800000&
         Caption         =   "Standard: "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   300
         Index           =   18
         Left            =   120
         TabIndex        =   38
         Top             =   255
         Width           =   1275
      End
      Begin VB.Label lblDetail 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800000&
         Caption         =   "Expanded: "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   300
         Index           =   17
         Left            =   120
         TabIndex        =   37
         Top             =   885
         Width           =   1275
      End
   End
   Begin VB.Label lblDetail 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000080&
      Caption         =   "Record No.: "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   300
      Index           =   20
      Left            =   7320
      TabIndex        =   41
      Top             =   120
      Width           =   1275
   End
   Begin VB.Label lblDetails 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   14
      Left            =   8685
      TabIndex        =   40
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblDetails 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   13
      Left            =   5085
      TabIndex        =   27
      Top             =   6645
      Width           =   2175
   End
   Begin VB.Label lblDetails 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   12
      Left            =   5085
      TabIndex        =   26
      Top             =   6330
      Width           =   2175
   End
   Begin VB.Label lblDetails 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   11
      Left            =   5085
      TabIndex        =   25
      Top             =   6015
      Width           =   2175
   End
   Begin VB.Label lblDetails 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   10
      Left            =   5085
      TabIndex        =   24
      Top             =   5625
      Width           =   2175
   End
   Begin VB.Label lblDetails 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   9
      Left            =   5085
      TabIndex        =   23
      Top             =   5310
      Width           =   2175
   End
   Begin VB.Label lblDetails 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   8
      Left            =   5085
      TabIndex        =   22
      Top             =   4995
      Width           =   2175
   End
   Begin VB.Label lblDetail 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00800000&
      Caption         =   "Back Colour: "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   300
      Index           =   13
      Left            =   3720
      TabIndex        =   21
      Top             =   6645
      Width           =   1275
   End
   Begin VB.Label lblDetail 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00800000&
      Caption         =   "Fore Colour: "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   300
      Index           =   12
      Left            =   3720
      TabIndex        =   20
      Top             =   6330
      Width           =   1275
   End
   Begin VB.Label lblDetail 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00800000&
      Caption         =   "Bold: "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   300
      Index           =   11
      Left            =   3720
      TabIndex        =   19
      Top             =   6015
      Width           =   1275
   End
   Begin VB.Label lblDetail 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      Caption         =   "Parent Key: "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   300
      Index           =   10
      Left            =   3720
      TabIndex        =   18
      Top             =   5625
      Width           =   1275
   End
   Begin VB.Label lblDetail 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      Caption         =   "Key: "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   300
      Index           =   9
      Left            =   3720
      TabIndex        =   17
      Top             =   5310
      Width           =   1275
   End
   Begin VB.Label lblDetail 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      Caption         =   "ID Tag: "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   300
      Index           =   8
      Left            =   3720
      TabIndex        =   16
      Top             =   4995
      Width           =   1275
   End
   Begin VB.Label lblDetails 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   1
      Left            =   5085
      TabIndex        =   15
      Top             =   435
      Width           =   2175
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblDetails 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   7
      Left            =   5085
      TabIndex        =   14
      Top             =   4680
      Width           =   2175
   End
   Begin VB.Label lblDetails 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   5
      Left            =   5085
      TabIndex        =   13
      Top             =   4320
      Width           =   2175
   End
   Begin VB.Label lblDetails 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Index           =   4
      Left            =   5085
      TabIndex        =   12
      Top             =   3795
      Width           =   4215
   End
   Begin VB.Label lblDetails 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   3
      Left            =   5085
      TabIndex        =   11
      Top             =   1155
      Width           =   2175
   End
   Begin VB.Label lblDetails 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   2
      Left            =   5085
      TabIndex        =   10
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label lblDetails 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   5085
      TabIndex        =   9
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label lblDetail 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00008000&
      Caption         =   "Node Key: "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   300
      Index           =   7
      Left            =   3720
      TabIndex        =   8
      Top             =   435
      Width           =   1275
   End
   Begin VB.Label lblDetail 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      Caption         =   "Recursive: "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   300
      Index           =   6
      Left            =   3720
      TabIndex        =   7
      Top             =   4680
      Width           =   1275
   End
   Begin VB.Label lblDetail 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      Caption         =   "SQL: "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   300
      Index           =   5
      Left            =   3720
      TabIndex        =   6
      Top             =   1470
      Width           =   1275
   End
   Begin VB.Label lblDetail 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      Caption         =   "Table Sort: "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   300
      Index           =   4
      Left            =   3720
      TabIndex        =   5
      Top             =   4320
      Width           =   1275
   End
   Begin VB.Label lblDetail 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      Caption         =   "Filter Criteria: "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   300
      Index           =   3
      Left            =   3720
      TabIndex        =   4
      Top             =   3795
      Width           =   1275
   End
   Begin VB.Label lblDetail 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      Caption         =   "Table Criteria: "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   300
      Index           =   2
      Left            =   3720
      TabIndex        =   3
      Top             =   1155
      Width           =   1275
   End
   Begin VB.Label lblDetail 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      Caption         =   "Table Name: "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   300
      Index           =   1
      Left            =   3720
      TabIndex        =   2
      Top             =   840
      Width           =   1275
   End
   Begin VB.Label lblDetail 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00008000&
      Caption         =   "Node Text: "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   300
      Index           =   0
      Left            =   3720
      TabIndex        =   1
      Top             =   120
      Width           =   1275
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
' Date:         10/07/2003
' Version:      00.01.00
' Description:  Test/Demo 4 - Complex Rules and Multiple Tables made easy!
' Edit History: 00.01.00 10/07/2003 Initial *BETA* Release
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

Private Const clCOLORMAGENTA    As Long = &H800080
Private Const clCOLORBROWN      As Long = &H40C0&

Private Const csAPPSAVE         As String = "\DEMO 04.DS"
Private Const csAPPLOAD         As String = "\DEMO 02.DS"

Private Const csDATABASE        As String = "\..\Demo.mdb"

'===========================================================================
' cADOTreeView Events
'
Private Sub moTreeDB_AfterLoading(ByVal Node As MSComctlLib.Node)
'    Debug.Print "moTreeDB_AfterLoading Node: " + Node.Text
End Sub

Private Sub moTreeDB_BeforeLoading(ByVal Node As MSComctlLib.Node, ByRef Cancel As Boolean)
'    Debug.Print "moTreeDB_BeforeLoading Node: " + Node.Text
End Sub

'===========================================================================
' TreeView Events
'
Private Sub tvwDialog_NodeClick(ByVal Node As MSComctlLib.Node)
    pShowdetails
End Sub

'===========================================================================
' Form Control Events
'
Private Sub chkIcons_Click()
    tvwDialog.Style = IIf(chkIcons.Value = vbChecked, _
                          tvwTreelinesPlusMinusPictureText, _
                          tvwTreelinesPlusMinusText)
End Sub

Private Sub cmdFile_Click(Index As Integer)
    '
    '## NOTE: Below I save/load the datashape structure and Imagelist to a
    '         file. You could have the data stored in a resource file and
    '         load it from there or have a seperate table in your database
    '         or save/load the binary data to/from the registry.
    '
    Select Case Index
        Case 0: moTreeDB.SaveShape App.Path + csAPPSAVE
        Case 1
                With moTreeDB
                    .LoadShape App.Path + csAPPLOAD
                    .DataShape.ConnectString App.Path + csDATABASE, , , , ejvJet4
                    .Reload
                End With
    End Select

End Sub

Private Sub lblDetail_Click(Index As Integer)
    '
    '## Only used to remove all color from the TreeView for Screenshop purposes
    '
    Dim oNode  As MSComctlLib.Node, _
        oImage As VB.Image, _
        oFrame As VB.Frame

    Select Case Index
        Case 0  '## 'Node Text:' label
            For Each oImage In imgDetails
                oImage.Visible = False
            Next

        Case 1  '## 'Table Name:' label
            For Each oNode In tvwDialog.Nodes
                With oNode
                    .ForeColor = vbWindowText
                    .BackColor = vbWindowBackground
                    .Bold = False
                End With
            Next

            With tvwDialog.Nodes(1)
                pShow 12, pColorName(.ForeColor)
                pShow 13, pColorName(.BackColor)
            End With

            For Each oFrame In fraDetails
                oFrame.Visible = False
            Next

    End Select

End Sub

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

    With tvwDialog
        If .Nodes.Count Then
            .SelectedItem = .Nodes(1)   '## Select and show the root node's details
            pShowdetails
        End If
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
        '   |     +- GRPSUP (GU) .................. [Recursive]
        '   |     |     |
        '   |     |     +- SUP (SU)
        '   |     |
        '   |     +- STK (ST)
        '   |     |
        '   |     +- Layer Country (Co) ........... [Recursive]
        '   |           |
        '   |           +- Layer City (Ci)
        '   |                 |
        '   |                 +- Layer Hotel (H)
        '   |
        '   |
        '   +- Layer GroupA (G) ................... [Recursive]
        '         |
        '         +- Layer ProductA (P)
        '
        ' Object         ID  MDB                 Sort   Parent         Build
        ' Key            Tag Table     Recursive Field  Object         Fields
        ' ==============+===+=========+=========+======+==============+==========================
        ' GRPPROD        GP  GroupB       Yes    Desc   ---            Desc, Fore/Back Color
        ' PROD           PR  ProductB     No     Desc   GRPPROD        ---
        ' GRPSTK         GT  GroupB       Yes    Desc   ---            ---
        ' GRPSUP         GU  GroupB       Yes    Desc   GRPSTK         ---
        ' SUP            SU  SupplierB    No     Desc   GRPSUP         ---
        ' STK            ST  StockB       No     Desc   GRPSTK         ---
        ' Layer Country  Co  Country      Yes    Desc   GRPSTK         Image
        ' Layer City     Ci  City         No     Desc   Layer Country  ---
        ' Layer Hotel    H   Hotel        No     Desc   Layer City     ---
        ' Layer GroupA   G   GroupA       Yes    SeqNum ---            ---
        ' Layer ProductA P   ProductA     No     Desc   Layer GroupA   ---
        ' --------------+---+---------+---------+------+--------------+--------------------------
        '
        ' Object         |<-    Required Field/Column Names     ->|
        ' Key             efldID  efldDesc  efldLinkID  efldParent
        ' ===============+=======+=========+===========+==========+==============================
        ' GRPPROD         PkID    KeyDesc   GroupID     ---
        ' PROD            PkID    Desc      GroupID     ---
        ' GRPSTK          PkID    Desc      GroupID     ---
        ' GRPSUP          PkID    Desc      GroupID     ---
        ' SUP             PkID    Desc      GroupID     ---
        ' STK             PkID    Desc      GroupID     ---
        ' Layer Country   PkID    Desc      LinkID      ParentID
        ' Layer City      PkID    Desc      LinkID      ---
        ' Layer Hotel     PkID    Desc      LinkID      ---
        ' Layer GroupA    PkID    Desc      GroupID     ---
        ' Layer ProductA  PkID    Desc      GroupID     ---
        ' ---------------+-------+---------+-----------+----------+------------------------------
        '
        ' Object Key     Property                   SQL Command (Pseudo Syntax)
        ' ==============+==========================+=============================================
        ' GRPPROD        .TableCriteria             Type=0
        ' GRPPROD        .Fields(efldDesc).Sql      (CStr(PkID) + ' - ' + Desc) AS KeyDesc
        ' GRPPROD        .Fields(efldForeColor).Sql (IIf([GroupID] > 9, vbRed, vbMagenta) AS ForeColor
        ' GRPPROD        .Fields(efldBackColor).Sql (IIf([GroupID] > 9, vbYellow, vbCyan) AS BackColor
        ' GRPPROD        .FilterCriteria            [KeyDesc] LIKE '*Be*' OR [KeyDesc] LIKE '*Sp*'
        ' PROD           .FilterCriteria            ([Desc] LIKE '*S*' AND [Desc] LIKE '*c*') OR
        '                                           ([Desc] LIKE '*S*' AND [Desc] LIKE '*t*')
        ' GRPSTK         .TableCriteria             Type=1
        ' GRPSUP         .TableCriteria             Type=4
        ' Layer Country  .Fields(efldImage).Sql     IIf(InStr([Desc],'Hemi'),'',[Desc]) AS NormImage
        ' Layer ProductA .FilterCriteria            [Desc] LIKE '*CR*'
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
        ' 2. In the 'Layer Country' Object, I've used both the 'efldLinkID' & 'efldParent'
        '    Fields. The 'efldLinkID' is used for the recursive (same table) link and the
        '    'efldParent' is the link to its parent object (table) 'GRPSTK'.
        '
        '    Now, 'GRPSUP' also has 'GRPSTK' as its parent object (table) however 'GRPSUP' and
        '    'GRPSTK' share the same TableName. Therefore 'efldLinkID' also acts as the
        '    'efldParent' field. Check the TableCriteria for both objects, look at the raw
        '    table in the database & then run the demo and watch the properties.
        '
        ' 3. If you look closely at GRPPROD, GRPSTK, & GRPSUP, you'll notice that They all use
        '    the same table. Therefore, by using the TableCriteria property of a DataObj to break
        '    the table elements into seperate groups of data, it's possible to set an unlimited
        '    number of definitions against a single table - e.g. Alphabetize Phonebook entries;
        '    color coded ranges; Images as warning indicators; color coded regions, etc...
        '    "... a picture is worth a thousand words."
        '
        ' 4. 'ProductA' table used by 'Layer ProductA' object has 694 records. Comment out the
        '    .FilterCriteria property against the 'Layer ProductA' object to see how quickly
        '    the TreeView still loads.
        '
        ' 5. In the IDE, I've left the DEBUGMODE Conditional flag set to watch the DLL activity
        '    in the IDE 'Immediate Window'. To change, select the vbADOTree project, then goto
        '    Menu 'Project' > 'Properties' > 'Make' > 'Conditional Compilation Arguments' and
        '    set 'DEBUGMODE = 0'. This will speed up the exectution considerablly.
        '
        With .DataShape
            '
            '-----------------------------------------------------------------------------------
            '## Product Group Table (Recursive structure)
            '
            With .Add("GRPPROD", "Groupb", "GP", "", True, vbRed, , True, 3, 5, 4)
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
                '   for custom sort orders without any complicated subclassing or APIs
                '   (refer to 'Retail Product Group Table' below and 'Group A' table in
                '   the database to see an implementation of using custom sorts.
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
                '
                '## Only records containing....
                '   This property is applied to the generated Recordset as a post Filter.
                .FilterCriteria = "[KeyDesc] LIKE '*Be*' OR " + _
                                  "[KeyDesc] LIKE '*Sp*'"
            End With
            '
            '-----------------------------------------------------------------------------------
            '## Product Table (Links to 'Product Group' by Record ID)
            '
            With .Add("PROD", "ProductB", "PR", "GRPPROD", , , , , 9)
                .TableSort = "[Desc]"

                With .Fields
                    .Item(efldID).Desc = "PkID"
                    .Item(efldDesc).Desc = "Desc"
                    .Item(efldLinkID).Desc = "GroupID"
                End With
                '
                '## Only records containing....
                '
                .FilterCriteria = "([Desc] LIKE '*S*' AND [Desc] LIKE '*c*') OR " + _
                                  "([Desc] LIKE '*S*' AND [Desc] LIKE '*t*')"
            End With
            '
            '-----------------------------------------------------------------------------------
            '## Stock Group Table (Recursive structure)
            '
            With .Add("GRPSTK", "GroupB", "GT", "", True, vbBlue, , True, 15)
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
            '## Supplier Group Table (Recursive structure) (Links to 'Stock Group' by Record ID)
            '
            With .Add("GRPSUP", "GroupB", "GU", "GRPSTK", True, vbGreen, , True, 10)
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
            '## Stock Table (Links to 'Stock Group' by Record ID)
            '
            With .Add("STK", "StockB", "ST", "GRPSTK", , , , , 16, 17, 16)
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
            With .Add("SUP", "SupplierB", "SU", "GRPSUP", , , , , 6, 8, 6)
                .TableSort = "[Desc]"

                With .Fields
                    .Item(efldID).Desc = "PkID"
                    .Item(efldDesc).Desc = "Desc"
                    .Item(efldLinkID).Desc = "GroupID"
                End With
            End With
            '
            '-----------------------------------------------------------------------------------
            '## Country Table (Recursive structure) (Links to 'Stock Group' by Record ID)
            '
            With .Add("Layer Country", "Country", "Co", "GRPSTK", True, , , True, "Co")
                .TableSort = "[Desc]"
                .SQLBuildFields = esqlImage

                With .Fields
                    .Item(efldID).Desc = "PkID"
                    .Item(efldDesc).Desc = "Desc"
                    .Item(efldLinkID).Desc = "LinkID"
                    .Item(efldParent).Desc = "ParentID"
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
             With .Add("Layer City", "City", "Ci", "Layer Country", , clCOLORMAGENTA, , , 11)
                .TableSort = "[Desc]"

                With .Fields
                    .Item(efldID).Desc = "PkID"
                    .Item(efldDesc).Desc = "Desc"
                    .Item(efldLinkID).Desc = "LinkID"
                End With
            End With
            '
            '-----------------------------------------------------------------------------------
            '## Hotel Table (Links to 'City' by Record ID)
            '
             With .Add("Layer Hotel", "Hotel", "H", "Layer City", , clCOLORBROWN, , , 18)
                .TableSort = "[Desc]"

                With .Fields
                    .Item(efldID).Desc = "PkID"
                    .Item(efldDesc).Desc = "Desc"
                    .Item(efldLinkID).Desc = "LinkID"
                End With
            End With
            '
            '-----------------------------------------------------------------------------------
            '## Retail Product Group Table (Recursive structure)
            '
            With .Add("Layer GroupA", "GroupA", "G", "", True, clCOLORBROWN, , True, 13)
                .TableSort = "SeqNum" '## Custom (non alphanumeric) sorting

                With .Fields
                    .Item(efldID).Desc = "PkID"
                    .Item(efldDesc).Desc = "Desc"
                    .Item(efldLinkID).Desc = "GroupID"
                End With
            End With
            '
            '-----------------------------------------------------------------------------------
            '## Retail Product Table (Links to 'Retail Product Group' by Record ID)
            '
             With .Add("Layer ProductA", "ProductA", "P", "Layer GroupA", , clCOLORMAGENTA, , , 14)
                .TableSort = "[Desc]"

                With .Fields
                    .Item(efldID).Desc = "PkID"
                    .Item(efldDesc).Desc = "Desc"
                    .Item(efldLinkID).Desc = "GroupID"
                End With
                '
                '## Only records containing....
                '
                .FilterCriteria = "[Desc] LIKE '*CR*'"
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

Private Sub pShowdetails()

    On Error Resume Next

    Dim sParentID  As String, _
        sParentKey As String, _
        oNode      As MSComctlLib.Node

    With moTreeDB
        Set oNode = tvwDialog.SelectedItem
        With oNode
            pShow 0, .Text
            pShow 1, .Key

            Select Case .FullPath = .Text
                Case True:  sParentID = "0"
                Case False: sParentID = CStr(Val(.Parent.Key))
            End Select
        End With

        With pDataObj(oNode.Key)
            '
            '## Get the DataObj relationship - Parent DataObj's Key
            '
            Select Case oNode.FullPath = oNode.Text
                '
                '## No parents to be found
                '
                Case True:  sParentKey = "[ROOT NODE]"
                '
                '## Retrieve the Parent
                '
                Case False
                    sParentKey = pDataObj(oNode.Parent.Key).Key
                    If sParentKey = .Key Then
                        '
                        '## Show that we're in a recursive relationship
                        '
                        sParentKey = "** " + .Key + " **"
                    End If
            End Select
            '
            '## Display important details to user on GUI
            '
            pShow 2, .TableName
            pShow 3, .TableCriteria
            pShow 4, .FilterCriteria
            pShow 5, .TableSort
            pShow 6, Replace(.SQL, "@@@", sParentID)
            pShow 7, CStr(.Recursive)
            pShow 8, .IDTag
            pShow 9, .Key
            pShow 10, sParentKey
            pShow 11, CStr(oNode.Bold)
            pShow 12, pColorName(oNode.ForeColor)
            pShow 13, pColorName(oNode.BackColor)
            pShow 14, Val(oNode.Key)

            If .Image Then
                imgDetails(0).Picture = imgDialog.ListImages(.Image).Picture
            Else
                imgDetails(0).Picture = Nothing
            End If
            If .SelectedImage Then
                imgDetails(1).Picture = imgDialog.ListImages(.SelectedImage).Picture
            Else
                imgDetails(1).Picture = Nothing
            End If
            If .ExpandedImage Then
                imgDetails(2).Picture = imgDialog.ListImages(.ExpandedImage).Picture
            Else
                imgDetails(2).Picture = Nothing
            End If
            
            If oNode.Image Then
                imgDetails(3).Picture = imgDialog.ListImages(oNode.Image).Picture
            Else
                imgDetails(3).Picture = Nothing
            End If
            If oNode.SelectedImage Then
                imgDetails(4).Picture = imgDialog.ListImages(oNode.SelectedImage).Picture
            Else
                imgDetails(4).Picture = Nothing
            End If
            If oNode.ExpandedImage Then
                imgDetails(5).Picture = imgDialog.ListImages(oNode.ExpandedImage).Picture
            Else
                imgDetails(5).Picture = Nothing
            End If

        End With
    End With

End Sub

Private Function pDataObj(ByVal Key As String) As cDataObj
    '
    '## Return a DataObj object from a Node's key.
    '
    With moTreeDB
        Set pDataObj = .DataShape(.Node2ShapeKey(Key))
    End With

End Function

Private Sub pShow(ByVal Index As Long, ByVal Text As String)

    Const vbDblCrLf = vbCrLf + vbCrLf

    If Index = 6 Then
        '
        '## Try and make the SQL command text easier to read
        '
       txtSQL.Text = Trim$(Replace( _
                           Replace( _
                           Replace(Text, " WHERE", vbDblCrLf + " WHERE"), _
                                         " FROM", vbDblCrLf + " FROM"), _
                                         " ORDER BY", vbDblCrLf + " ORDER BY"))
    Else
        lblDetails(Index).Caption = Text
    End If

End Sub

Private Function pColorName(ByVal Value As OLE_COLOR) As String
    '
    '## Convert a color value to a human readable format or name (if known)
    '
    Select Case Value
        '
        '** Visual Basic Color Names
        '
        Case vbBlack:                pColorName = "vbBlack"
        Case vbRed:                  pColorName = "vbRed"
        Case vbGreen:                pColorName = "vbGreen"
        Case vbYellow:               pColorName = "vbYellow"
        Case vbBlue:                 pColorName = "vbBlue"
        Case vbMagenta:              pColorName = "vbMagenta"
        Case vbCyan:                 pColorName = "vbCyan"
        Case vbWhite:                pColorName = "vbWhite"
        '
        '** System Color Names
        '
        Case vbScrollBars:           pColorName = "vbScrollBars"
        Case vbDesktop:              pColorName = "vbDesktop"
        Case vbActiveTitleBar:       pColorName = "vbActiveTitleBar"
        Case vbInactiveTitleBar:     pColorName = "vbInactiveTitleBar"
        Case vbMenuBar:              pColorName = "vbMenuBar"
        Case vbWindowBackground:     pColorName = "vbWindowBackground"
        Case vbWindowFrame:          pColorName = "vbWindowFrame"
        Case vbMenuText:             pColorName = "vbMenuText"
        Case vbWindowText:           pColorName = "vbWindowText"
        Case vbTitleBarText:         pColorName = "vbTitleBarText"
        Case vbActiveBorder:         pColorName = "vbActiveBorder"
        Case vbInactiveBorder:       pColorName = "vbInactiveBorder"
        Case vbApplicationWorkspace: pColorName = "vbApplicationWorkspace"
        Case vbHighlight:            pColorName = "vbHighlight"
        Case vbHighlightText:        pColorName = "vbHighlightText"
        Case vbButtonFace:           pColorName = "vbButtonFace"
        Case vbButtonShadow:         pColorName = "vbButtonShadow"
        Case vbGrayText:             pColorName = "vbGrayText"
        Case vbButtonText:           pColorName = "vbButtonText"
        Case vbInactiveCaptionText:  pColorName = "vbInactiveCaptionText"
        Case vb3DHighlight:          pColorName = "vb3DHighlight"
        Case vb3DDKShadow:           pColorName = "vb3DDKShadow"
        Case vb3DLight:              pColorName = "vb3DLight"
        Case vb3DFace:               pColorName = "vb3DFace"
        Case vb3DShadow:             pColorName = "vb3DShadow"
        Case vbInfoText:             pColorName = "vbInfoText"
        Case vbInfoBackground:       pColorName = "vbInfoBackground"
        '
        '** Convert to hexadecimal instead
        '
        Case Else: pColorName = "0x" + Right$("00000000" + CStr(Hex$(Value)), 8)
    End Select

End Function
