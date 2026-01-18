VERSION 5.00
Begin VB.Form PRESIMPO 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vista Previa de la Importación"
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11055
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   11055
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   435
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2205
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Eliminar Registro"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4620
      TabIndex        =   1
      Top             =   5985
      Width           =   2115
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5460
      ItemData        =   "DATIMP.frx":0000
      Left            =   210
      List            =   "DATIMP.frx":0002
      MultiSelect     =   2  'Extended
      TabIndex        =   0
      Top             =   315
      Width           =   10620
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " F.Costeo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   9765
      TabIndex        =   7
      Top             =   0
      Width           =   1065
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Lote"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7350
      TabIndex        =   6
      Top             =   0
      Width           =   1170
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Costo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8505
      TabIndex        =   5
      Top             =   0
      Width           =   1275
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "U.M."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6720
      TabIndex        =   4
      Top             =   0
      Width           =   645
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Descripción"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2205
      TabIndex        =   3
      Top             =   0
      Width           =   4530
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Código del Artículo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   210
      TabIndex        =   2
      Top             =   0
      Width           =   2010
   End
End
Attribute VB_Name = "PRESIMPO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DBS As Database
Dim RST As Recordset
Dim FLD As Fields

Private Sub Command1_Click()
'
I = 1
Do While I <= List1.ListCount
   If List1.Selected(I - 1) Then
      If I < List1.ListCount Then
         If Not List1.Selected(I) Then
            List1.RemoveItem (I - 1)
            List1.Selected(I - 1) = False
         Else
            List1.RemoveItem (I - 1)
         End If
      Else
         List1.RemoveItem (I - 1)
      End If
   Else
      I = I + 1
   End If
Loop
'
End Sub

Private Sub Form_Load()
On Error GoTo CONTROLERROR
'
Dim XX As String, XXX As String, YY As Integer
Dim LGCOL(1 To 8) As Integer
'
LGCOL(1) = 16: LGCOL(2) = 40
LGCOL(3) = 12: LGCOL(4) = 4
LGCOL(5) = 10: LGCOL(6) = 10
LGCOL(7) = 10: LGCOL(8) = 10
'
Set DBS = OpenDatabase(Trim$(FIMPORTA.Label1.Caption), False, False, "Excel 4.0;HDR=NO;")
Set RST = DBS.OpenRecordset("A1:H10000")
RST.MoveLast
ULTIMO = RST.RecordCount
RST.MoveFirst
'
For I = 1 To ULTIMO
   XXX = ""
   For Z = 1 To 7
      COL = CInt(FIMPORTA.Text1(Z).Text) - 1
      XX = ""
      On Error Resume Next
      XX = RST(COL)
      On Error GoTo 0
      Select Case COL
             Case 0, 1, 2, 3, 4, 7
                  XXX = XXX + AJUSTI$(XX, LGCOL(Z))
             Case 5
                  XXX = XXX + AJUSTD$(XX, LGCOL(Z))
             Case 6
                  XXX = XXX + AJUCEN$(XX, LGCOL(Z))
      End Select
   Next Z
   List1.AddItem XXX
   RST.MoveNext
Next I
'
RST.Close 'Cierra el objeto Recordset
DBS.Close 'Cierra el objeto Database
'
Exit Sub
'
CONTROLERROR:
Call ERRORES(Err.Number)
Exit Sub
End Sub

Private Sub ERRORES(ERROR As Integer)
'
Select Case ERROR
       Case 3011
            MsgBox "No se ha encontrado un archivo de importación"
            
       Case 3274
            MsgBox "El archivo de importación no responde a un" & _
                   "formato válido"
End Select
'
End Sub

