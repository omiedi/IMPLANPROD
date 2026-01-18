VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5805
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   6600
   LinkTopic       =   "Form1"
   ScaleHeight     =   5805
   ScaleWidth      =   6600
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   540
      Left            =   4830
      TabIndex        =   6
      Top             =   4830
      Width           =   1275
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   645
      Left            =   4935
      TabIndex        =   5
      Top             =   3885
      Width           =   1380
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   "C:\Archivos de programa\DevStudio\VB\Biblio.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      EOFAction       =   2  'Add New
      Exclusive       =   0   'False
      Height          =   345
      Left            =   105
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3150
      Width           =   2010
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2220
      Left            =   5880
      TabIndex        =   2
      Top             =   840
      Width           =   4950
      _ExtentX        =   8731
      _ExtentY        =   3916
      _Version        =   327680
      AllowUserResizing=   1
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "Form8.frx":0000
      Height          =   1905
      Left            =   210
      OleObjectBlob   =   "Form8.frx":0010
      TabIndex        =   1
      Top             =   1050
      Width           =   4425
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   330
      Left            =   3045
      TabIndex        =   0
      Top             =   210
      Width           =   3270
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Archivos de programa\DevStudio\VB\Biblio.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      EOFAction       =   2  'Add New
      Exclusive       =   0   'False
      Height          =   345
      Left            =   210
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Titles"
      Top             =   210
      Width           =   2010
   End
   Begin MSDBGrid.DBGrid DBGrid2 
      Bindings        =   "Form8.frx":09C2
      Height          =   1905
      Left            =   105
      OleObjectBlob   =   "Form8.frx":09D2
      TabIndex        =   4
      Top             =   3780
      Width           =   4425
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   330
      Left            =   5355
      TabIndex        =   3
      Top             =   3360
      Width           =   960
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim DBS As Database
Dim RST As Recordset
'
Set DBS = OpenDatabase("C:\Archivos de programa\DevStudio\VB\Biblio.mdb")
Set RST = DBS.OpenRecordset("SELECT ALL  * FROM TITLES;")
'AUTHORS WHERE AU_ID>1000 AND AU_ID <2000 ORDER BY AUTHOR ASC")
RST.MoveLast
QREG = RST.RecordCount
Qcam = RST.Fields.Count
   
MSFlexGrid1.Cols = Qcam
MSFlexGrid1.Rows = QREG
Label1 = Str(QREG)

'MSFLEXGRID1.CLIP
'Data2.Recordset = RST
'DBGrid2.DataSource = Data2
RST.MoveFirst
'DBGrid2.Col = 0
'Data2.Recordset = RST
'DBGrid2.DataSource = Data2
'DBGrid2.Row = 0
For Q1 = 1 To QREG
  'DBGrid2.Rows = Q1
  'Data2.Recordset = RST
  'DBGrid2.AddNewMode
  MSFlexGrid1.Row = Q1 - 1
  For Q2 = 1 To Qcam
    MSFlexGrid1.Col = Q2 - 1
    On Error Resume Next
    MSFlexGrid1.Text = RST.Fields(Q2 - 1)
  Next Q2
    MSFlexGrid1.Col = Q2 - 1
  RST.MoveNext
Next Q1
MSFlexGrid1.Refresh

'Copyright (c) 1996 Microsoft Corporation
'MsgBox RST.RecordCount
   ' Llena el recordset.
End Sub

Private Sub Command2_Click()
Data2.RecordSource = "SELECT ALL  * FROM TITLES where INSTR(TITLE,'Base')>0 order by title"
Data2.Refresh
'Data2.Recordset.MoveLast
Label1 = Data2.Recordset.RecordCount

End Sub

Private Sub Text1_Change()
'Data1.Recordset.MoveFirst
CRITERIO$ = "INSTR (TITLE,'" & Text1 & "')>0"
Data2.RecordSource = "SELECT ALL  * FROM TITLES where " + CRITERIO$
Data2.Refresh
'Data2.Recordset.MoveLast
Label1 = Data2.Recordset.RecordCount
Label1.Refresh
'Data1.Recordset.FindFirst CRITERIO$
'Data1.Refresh
'DBGrid1.Row = Data1.Recordset..RecordCount

End Sub
