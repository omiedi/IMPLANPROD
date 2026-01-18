VERSION 5.00
Begin VB.Form FIMPORTA 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sistemas FLEXOFT(r) - Importación de Datos"
   ClientHeight    =   6165
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   8625
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   8625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Cancel          =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   7455
      Picture         =   "FIMPORTA.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Cancela - Abandona la Importación"
      Top             =   4840
      Width           =   1065
   End
   Begin VB.Frame Frame5 
      Caption         =   "Ruta de Acceso"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   210
      TabIndex        =   30
      Top             =   2730
      Width           =   8310
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   210
         TabIndex        =   31
         Top             =   315
         Width           =   7890
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Origen de Importación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   3045
      TabIndex        =   26
      Top             =   105
      Width           =   5475
      Begin VB.DirListBox Dir1 
         Height          =   1440
         Left            =   210
         TabIndex        =   29
         Top             =   900
         Width           =   2640
      End
      Begin VB.FileListBox File1 
         Height          =   2040
         Left            =   2940
         TabIndex        =   28
         Top             =   315
         Width           =   2325
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   210
         TabIndex        =   27
         Top             =   315
         Width           =   2640
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Base a Importar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   210
      TabIndex        =   17
      Top             =   105
      Width           =   2640
      Begin VB.OptionButton Option1 
         Caption         =   "Proveedores"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   315
         TabIndex        =   22
         Top             =   1365
         Width           =   1590
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Clientes"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   315
         TabIndex        =   21
         Top             =   1155
         Width           =   1590
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Hojas de Ruta"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   315
         TabIndex        =   20
         Top             =   735
         Width           =   1590
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Formulas de Producto"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   315
         TabIndex        =   19
         Top             =   525
         Width           =   2220
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Maestro de Items"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   315
         TabIndex        =   18
         Top             =   315
         Value           =   -1  'True
         Width           =   2115
      End
   End
   Begin VB.CommandButton Command2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   7455
      Picture         =   "FIMPORTA.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4260
      Width           =   1065
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   7455
      Picture         =   "FIMPORTA.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Inicia el Proceso de Importación"
      Top             =   3675
      Width           =   1065
   End
   Begin VB.Frame Frame2 
      Caption         =   "Formato de la Base a Importar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2430
      Left            =   210
      TabIndex        =   1
      Top             =   3570
      Width           =   7050
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Index           =   10
         Left            =   5880
         TabIndex        =   13
         Text            =   "11"
         ToolTipText     =   "Columna que Contiene el Lote Standard o Económico"
         Top             =   1890
         Width           =   330
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Index           =   9
         Left            =   3675
         TabIndex        =   12
         Text            =   "10"
         ToolTipText     =   "Columna que Contiene el Lote Máximo"
         Top             =   1890
         Width           =   330
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Index           =   0
         Left            =   1785
         TabIndex        =   3
         Text            =   "1"
         ToolTipText     =   "Columna que contiene el Código de Producto"
         Top             =   420
         Width           =   330
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Index           =   8
         Left            =   1785
         TabIndex        =   11
         Text            =   "9"
         ToolTipText     =   "Columna que contiene el Stock Mínimo"
         Top             =   1890
         Width           =   330
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Index           =   7
         Left            =   5880
         TabIndex        =   10
         Text            =   "8"
         ToolTipText     =   "Columna que Contiene el Proveedor Habitual"
         Top             =   1365
         Width           =   330
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Index           =   6
         Left            =   5880
         TabIndex        =   9
         Text            =   "7"
         ToolTipText     =   "Columna que Contiene la Fecha de Costeo"
         Top             =   1050
         Width           =   330
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Index           =   5
         Left            =   5880
         TabIndex        =   8
         Text            =   "6"
         ToolTipText     =   "Columna que Contiene la Moneda de Costeo"
         Top             =   735
         Width           =   330
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Index           =   4
         Left            =   5880
         TabIndex        =   7
         Text            =   "5"
         ToolTipText     =   "Columna que Contiene el Costo Unitario de Reposición"
         Top             =   420
         Width           =   330
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Index           =   3
         Left            =   1785
         TabIndex        =   6
         Text            =   "4"
         ToolTipText     =   "Columna que Contiene lel Tipo de Material"
         Top             =   1365
         Width           =   330
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Index           =   2
         Left            =   1785
         TabIndex        =   5
         Text            =   "3"
         ToolTipText     =   "Columna que contiene la Unidad de Medida"
         Top             =   1050
         Width           =   330
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Index           =   1
         Left            =   1785
         TabIndex        =   4
         Text            =   "2"
         ToolTipText     =   "Columna que Contiene la Descripción del Artículo"
         Top             =   735
         Width           =   330
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Económico:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   10
         Left            =   4410
         TabIndex        =   41
         Top             =   1995
         Width           =   1275
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Máximo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   9
         Left            =   2310
         TabIndex        =   40
         Top             =   1995
         Width           =   1275
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Stock Mínimo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   8
         Left            =   315
         TabIndex        =   39
         Top             =   1995
         Width           =   1275
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Proveedor:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   7
         Left            =   4410
         TabIndex        =   38
         Top             =   1470
         Width           =   1275
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Fecha:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   6
         Left            =   4410
         TabIndex        =   37
         Top             =   1155
         Width           =   1275
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Moneda:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   5
         Left            =   4410
         TabIndex        =   36
         Top             =   840
         Width           =   1275
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Costo:Unitario:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   4
         Left            =   4410
         TabIndex        =   35
         Top             =   525
         Width           =   1275
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Tipo de Material:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   105
         TabIndex        =   34
         Top             =   1470
         Width           =   1485
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Unid.de Medida:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   105
         TabIndex        =   33
         Top             =   1155
         Width           =   1485
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Descripción:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   315
         TabIndex        =   32
         Top             =   840
         Width           =   1275
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Código:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   315
         TabIndex        =   2
         Top             =   525
         Width           =   1275
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Formato"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   210
      TabIndex        =   0
      Top             =   1870
      Width           =   2640
      Begin VB.OptionButton Option2 
         Caption         =   "DBF"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   1890
         TabIndex        =   25
         Top             =   315
         Width           =   700
      End
      Begin VB.OptionButton Option2 
         Caption         =   "TXT"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   1050
         TabIndex        =   24
         Top             =   315
         Width           =   750
      End
      Begin VB.OptionButton Option2 
         Caption         =   "XLS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   210
         TabIndex        =   23
         Top             =   315
         Value           =   -1  'True
         Width           =   750
      End
   End
   Begin VB.Image Image2 
      Height          =   555
      Left            =   7245
      Picture         =   "FIMPORTA.frx":091E
      Stretch         =   -1  'True
      Top             =   5430
      Width           =   5250
   End
End
Attribute VB_Name = "FIMPORTA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'
If VALIDACION Then
   PRESIMPO.Show 1
End If
'
End Sub

Private Sub Command2_Click()
   'shell Label1.Caption
End Sub

Private Sub Command3_Click()
   Call FINAL("")
End Sub

Private Sub Dir1_Change()
  Call MUARCHI
End Sub

Sub MUARCHI()
  RUTA$ = Dir1.Path
  If Right$(RUTA$, 1) <> "\" Then RUTA$ = RUTA$ + "\"
  File1.Path = RUTA$
  EXTEN$ = Option2(0).Caption
  For KKI% = 0 To 2
    If Option2(KKI%).Value = True Then EXTEN$ = Option2(KKI%).Caption
  Next KKI%
  File1.Pattern = "*." + EXTEN$
  Label1.Caption = RUTA$ + File1.filename
End Sub

Private Sub Drive1_Change()
  Dir1.Path = Drive1.Drive + "\"
  Call MUARCHI
End Sub

Private Sub File1_Click()
  Call MUARCHI
End Sub

Private Function VALIDACION() As Boolean
'
Dim VALIDA As Boolean
'
VALIDA = True
'
For I = 1 To 8
  If IsNumeric(Text1(I).TEXT) Then
     For Z = (I + 1) To 8
         If TRIM$(Text1(I).TEXT) <> 0 And TRIM$(Text1(I).TEXT) = TRIM$(Text1(Z).TEXT) Then
            MsgBox "Error - Se ha asignado la misma columna" & vbCrLf _
                   & "a dos datos diferentes"
            VALIDA = False
            Exit For
         End If
     Next Z
     If Not VALIDA Then Exit For
  Else
     If TRIM$(Text1(I).TEXT) = "" Then
        Text1(I).TEXT = "0"
     Else
        MsgBox "El dato ingresado no es un número"
        VALIDA = False
        Exit For
     End If
  End If
Next I
'
VALIDACION = VALIDA
'
End Function

Private Sub Form_Load()
'
'List1.ListIndex = 1
Label1.Caption = Dir1.Path
Call MUARCHI
'
End Sub

Private Sub List1_Click()
'
File1.Pattern = List1.List(List1.ListIndex)
File1.Refresh
'
End Sub

Private Sub Label1_Change()
   '
   LOARCH = 0
   On Error Resume Next
   LOARCH = FileLen(Label1.Caption)
   On Error GoTo 0
   If LOARCH > 0 Then
       Command1.Enabled = True
       Command2.Enabled = True
     Else
       Command1.Enabled = False
       Command2.Enabled = False
   End If
   '
End Sub

Private Sub Option2_Click(Index As Integer)
   Call MUARCHI
End Sub
