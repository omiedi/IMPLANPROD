VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{02B5E320-7292-11CF-93D5-0020AF99504A}#1.0#0"; "MSCHART.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{CA8A9783-280D-11CF-A24D-444553540000}#1.3#0"; "pdf.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form CTRLSHOW 
   Caption         =   "Controles Active-X"
   ClientHeight    =   4425
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   6960
   LinkTopic       =   "Form1"
   ScaleHeight     =   4425
   ScaleWidth      =   6960
   StartUpPosition =   3  'Windows Default
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   375
      Left            =   5400
      OleObjectBlob   =   "CTRLSHOW.frx":0000
      TabIndex        =   11
      Top             =   3960
      Width           =   1455
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   4440
      OleObjectBlob   =   "CTRLSHOW.frx":007A
      Top             =   3840
   End
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   6960
      _ExtentX        =   12277
      _ExtentY        =   741
      Appearance      =   1
      _Version        =   327682
   End
   Begin PdfLib.Pdf Pdf1 
      Height          =   750
      Left            =   4305
      TabIndex        =   9
      Top             =   3045
      Width           =   1170
      _Version        =   327680
      _ExtentX        =   2064
      _ExtentY        =   1323
      _StockProps     =   0
      SRC             =   ""
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   6120
      Top             =   3120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   855
      Left            =   0
      TabIndex        =   10
      Top             =   1800
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1508
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"CTRLSHOW.frx":02AE
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2205
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2205
      Width           =   1170
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   315
      Left            =   2205
      TabIndex        =   5
      Top             =   1785
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   556
      _Version        =   327681
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBList DBList1 
      Height          =   840
      Left            =   1260
      TabIndex        =   4
      Top             =   1785
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1482
      _Version        =   327681
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Height          =   960
      Left            =   105
      OleObjectBlob   =   "CTRLSHOW.frx":0339
      TabIndex        =   3
      Top             =   3360
      Width           =   1905
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   5985
      Top             =   1785
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   5985
      Top             =   1155
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   327680
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin MSMAPI.MAPISession MAPISession1 
      Left            =   5985
      Top             =   525
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   327680
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5985
      Top             =   2415
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   1695
      Left            =   3990
      TabIndex        =   2
      Top             =   1050
      Width           =   1950
      _ExtentX        =   3440
      _ExtentY        =   2990
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "CTRLSHOW.frx":0D0C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Tab 1"
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Tab 2"
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   1695
      Left            =   0
      TabIndex        =   1
      Top             =   105
      Width           =   2010
      _ExtentX        =   3545
      _ExtentY        =   2990
      _Version        =   393216
   End
   Begin MSChartLib.MSChart MSChart1 
      Height          =   1695
      Left            =   1995
      OleObjectBlob   =   "CTRLSHOW.frx":0D28
      TabIndex        =   0
      Top             =   0
      Width           =   1905
   End
   Begin VB.PictureBox ImageList1 
      BackColor       =   &H80000005&
      Height          =   480
      Left            =   2205
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   7
      Top             =   3465
      Width           =   1200
   End
   Begin VB.OLE OLE1 
      Class           =   "Word.Document.8"
      Height          =   750
      Left            =   2205
      TabIndex        =   6
      Top             =   2625
      Width           =   1065
   End
End
Attribute VB_Name = "CTRLSHOW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
