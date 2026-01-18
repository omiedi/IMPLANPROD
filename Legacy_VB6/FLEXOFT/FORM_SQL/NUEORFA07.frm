VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form NUEORFA07 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Nueva Orden de Fabricación"
   ClientHeight    =   7800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   0
      Left            =   1800
      ScaleHeight     =   375
      ScaleWidth      =   1110
      TabIndex        =   44
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Operaciones"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2315
      Left            =   120
      TabIndex        =   41
      Top             =   5400
      Width           =   11715
      Begin VB.ListBox List2 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1380
         ItemData        =   "NUEORFA07.frx":0000
         Left            =   120
         List            =   "NUEORFA07.frx":0007
         TabIndex        =   42
         Top             =   645
         Width           =   11490
      End
      Begin VB.Label Label19 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Nro. Operación                Maq.   Hs.Fijo Hs.Var.  %DD   Total     Ops. Eq. Auxiliares"
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   43
         Top             =   360
         Width           =   11490
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Materiales"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   120
      TabIndex        =   38
      Top             =   2090
      Width           =   10815
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2205
         ItemData        =   "NUEORFA07.frx":000E
         Left            =   120
         List            =   "NUEORFA07.frx":0015
         TabIndex        =   39
         Top             =   645
         Width           =   10530
      End
      Begin VB.Label Label18 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Codigo         Descripcion              U-M   U.Unit. Necesidad   St. Neto Faltante"
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   40
         Top             =   360
         Width           =   10530
      End
   End
   Begin VB.PictureBox Picture14 
      BackColor       =   &H00004080&
      Height          =   5205
      Left            =   11025
      ScaleHeight     =   5145
      ScaleWidth      =   1005
      TabIndex        =   8
      Top             =   0
      Width           =   1065
      Begin VB.CommandButton Command2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "NUEORFA07.frx":001C
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "O.K."
         Top             =   3720
         Width           =   650
      End
      Begin VB.CommandButton Command14 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "NUEORFA07.frx":0326
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Configuración de Procesos y Formularios"
         Top             =   2400
         Width           =   650
      End
      Begin VB.CommandButton Command29 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "NUEORFA07.frx":0768
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   1125
         Width           =   650
      End
      Begin VB.CommandButton Command1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "NUEORFA07.frx":0A72
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Salir"
         Top             =   480
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   9
         Top             =   3240
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   10
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -300
         Picture         =   "NUEORFA07.frx":0BBC
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   4650
         Width           =   1515
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Identificación"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   7
      Top             =   240
      Width           =   10815
      Begin VB.CommandButton Command3 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8505
         TabIndex        =   29
         Top             =   420
         Width           =   175
      End
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7440
         TabIndex        =   6
         Text            =   " "
         Top             =   1200
         Width           =   855
      End
      Begin VB.CommandButton Command7 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8280
         TabIndex        =   36
         Top             =   1200
         Width           =   175
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7440
         TabIndex        =   5
         Text            =   " "
         Top             =   840
         Width           =   855
      End
      Begin VB.CommandButton Command6 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8280
         TabIndex        =   33
         Top             =   840
         Width           =   175
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5160
         TabIndex        =   2
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9495
         TabIndex        =   4
         Text            =   " "
         Top             =   420
         Width           =   1065
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7440
         TabIndex        =   3
         Text            =   " "
         Top             =   420
         Width           =   1065
      End
      Begin VB.CommandButton Command5 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   10560
         TabIndex        =   30
         Top             =   420
         Width           =   175
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Height          =   1215
         Left            =   120
         TabIndex        =   17
         Top             =   420
         Width           =   6165
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2745
            TabIndex        =   1
            Text            =   " "
            Top             =   0
            Width           =   315
         End
         Begin VB.CommandButton Command4 
            Caption         =   "."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2880
            TabIndex        =   23
            Top             =   420
            Width           =   175
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   900
            TabIndex        =   0
            Text            =   " "
            Top             =   420
            Width           =   1990
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Pr."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   2340
            TabIndex        =   26
            Top             =   45
            Width           =   375
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cant. "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   4440
            TabIndex        =   31
            Top             =   480
            Width           =   615
         End
         Begin VB.Label Label11 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5190
            TabIndex        =   28
            Top             =   0
            Width           =   945
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   4515
            TabIndex        =   27
            Top             =   45
            Width           =   615
         End
         Begin VB.Label Label7 
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3720
            TabIndex        =   25
            Top             =   420
            Width           =   600
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "U.M."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   3330
            TabIndex        =   24
            Top             =   480
            Width           =   375
         End
         Begin VB.Label Label5 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   900
            TabIndex        =   22
            Top             =   840
            Width           =   5235
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   900
            TabIndex        =   21
            Top             =   0
            Width           =   1095
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Nro."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   -450
            TabIndex        =   20
            Top             =   45
            Width           =   1335
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Código"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   -450
            TabIndex        =   19
            Top             =   495
            Width           =   1335
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Descrip."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   -450
            TabIndex        =   18
            Top             =   915
            Width           =   1335
         End
      End
      Begin VB.Label Label17 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8520
         TabIndex        =   37
         Top             =   1200
         Width           =   2220
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Sector"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6645
         TabIndex        =   35
         Top             =   1260
         Width           =   735
      End
      Begin VB.Label Label15 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8520
         TabIndex        =   34
         Top             =   840
         Width           =   2220
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Destino"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6645
         TabIndex        =   32
         Top             =   900
         Width           =   735
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "F.Ent."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   8955
         TabIndex        =   16
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "F.Lanz."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6645
         TabIndex        =   15
         Top             =   480
         Width           =   735
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   420
      OleObjectBlob   =   "NUEORFA07.frx":2ADE
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "NUEORFA07.frx":2D12
      TabIndex        =   45
      Top             =   315
      Visible         =   0   'False
      Width           =   1590
   End
End
Attribute VB_Name = "NUEORFA07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   '
   Call CIERRARCH("*.*")
   Unload Me
   '
End Sub

Private Sub Command2_Click()
  Command2.Enabled = False
  '1.-VALIDACIONES
  CI1% = CODINT%(Text1.TEXT)
  If CI1% <= 0 Or CI1% > NUMAS% Then
    Call MENSERR(24, "Codigo de Articulo no Válido")
    GoTo 99
  End If
  '
  CANTFAB = Val(Text5)
  If CANTFAB <= 0.005 Then
    Call MENSERR(24, "Cantidad de Orden de Fabricación no Válida")
    GoTo 99
  End If
  '
  FA% = FECHANUM(Label11.Caption) 'fecha generación o/f
  FL% = FECHANUM(Text3.TEXT) 'fecha lanzamiento
  FE% = FECHANUM(Text4.TEXT) 'fecha de entrega proyectada
  '
  If FL% < FA% Then
    Call MENSERR(24, "La Fecha de Lanzamiento no puede ser \Inferior a la Fecha Actual")
    GoTo 99
  End If
  '
  If FE% < FA% Then
    Call MENSERR(24, "La Fecha de Entrega Proyectada no puede ser \Inferior a la Fecha Actual")
    GoTo 99
  End If
  '
  If CONTROL$("", "EXISEPRO") = "SI" Then 'EXIGE SECTOR PRODUCTIVO
    If TRIM$(ECOARCH$("SECPROD", MKS$(Val(Text7)))) = "" Then
      Call MENSERR(24, "Sector Productivo no Declarado o no Válido")
      GoTo 99
    End If
  End If
  '2.-GRABACION DE O/F
  '3.-IMPRESIÓN DE O/F
99 Command2.Enabled = True
End Sub

Private Sub Command29_Click()
  '
  Call OPNOIN("Ayuda FLEXOFT en Línea")
  '
End Sub

Private Sub Command3_Click()
  '
  Call SELECHO("SELFECHA")
  VDEV$ = VALACT1$("SELFECHA")
  If VDEV$ <> "" Then
    Text3.TEXT = VDEV$
  End If
  '
End Sub

Private Sub Command4_Click()
  '
  Call SELECHO("MASTER")
  CODD$ = VALACT1$("MASTER")
  If CODD$ = "" Then
      Exit Sub
  End If
  COD$ = CODD$
  Text1.TEXT = COD$
  CI% = CODINT%(COD$)
  Label5.Caption = DESCRIT$(CI%)
  Label7.Caption = Unimed$(CI%)
  '
End Sub

Private Sub Command5_Click()
  '
  Call SELECHO("SELFECHA")
  VDEV$ = VALACT1$("SELFECHA")
  If VDEV$ <> "" Then
    Text4.TEXT = VDEV$
  End If
  '
End Sub

Private Sub Command6_Click()
  '
  Call SELECHO("DESTFAB")
  COD$ = VALACT1$("DESTFAB")
  Text6.TEXT = COD$
  Label15.Caption = TRIM$(ECOARCH$("DESTFAB", Text6.TEXT))
  '
End Sub

Private Sub Command7_Click()
  '
  Call SELECHO("SECPROD")
  COD$ = VALACT1$("SECPROD")
  Text7.TEXT = COD$
  CS1 = Val(Text7.TEXT)
  Label17.Caption = TRIM$(ECOARCH$("SECPROD", MKS$(CS1)))
  '
End Sub

Private Sub Form_Load()
  '
  Call APLICACIONSKINS(Me, Picture14)
  Call CENTRAFORM(Me)
  '
  Call LIMPIAOF
  '
End Sub

Private Sub Label18_DBLClick()
  '
  Call List1_DblClick
  '
End Sub

Private Sub List1_DblClick()
  '
  CI1% = CODINT%(Text1.TEXT)
  If CI1% <= 0 Then
    Exit Sub
  End If
  '
  Call BLANARCH("!MATEROF")
  '
  JJ& = 0
  For J% = 0 To List1.ListCount - 1
    L1$ = List1.List(J%)
    'LEO DEL LIST1
    CDE1$ = Mid$(L1$, 1, 16)
    CD1% = CODINT%(CDE1$)
    UM1$ = Mid$(L1$, 42, 4)
    USOU1# = Val(Mid$(L1$, 47, 8))
    CAN1 = Val(Mid$(L1$, 57, 8))
    STD1 = Val(Mid$(L1$, 68, 8))
    FALT1 = Val(Mid$(L1$, 77, 8))
    '
    Y$ = REGBLAN$("!MATEROF")
    Call REPLA(Y$, MKI$(CD1%), 1, 2)
    Call REPLA(Y$, UM1$, 3, 2)
    Call REPLA(Y$, MKS$(CAN1), 5, 4)
    Call REPLA(Y$, MKD$(USOU1#), 9, 8)
    Call REPLA(Y$, MKS$(STD1), 21, 4)
    Call REPLA(Y$, MKS$(FALT1), 25, 4)
    JJ& = JJ& + 1
    Call GRAREG("!MATEROF", Y$, JJ&)
  Next J%
  '
  Call SETULTREG("!MATEROF", JJ&)
  Call CIERRARCH("!MATEROF")
  '
  ATRI$ = "BORDEIZQ=W" + TRIM$(str$(Left + 300))
  ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(str$(Top + 1650))
  TITUPA$ = "Materiales"
  VTB% = VENTABU%("MATEROF/" + ATRI$ + "/TITUPAN=" + TITUPA$)
  If VTB% >= 0 Then
    List1.Clear
    COEF = 1
    For JX& = 1 To ULTREG("!MATEROF")
      X$ = REGLEIDO$("!MATEROF", JX&)
      '
      CD1% = CVI(Mid$(X$, 1, 2))
      DE1$ = DESCRIT0$(CD1%)
      UM1$ = Unimed$(CD1%)
      CACONJU1 = CVS(Mid$(X$, 5, 4))
      USOU1# = CVD(Mid$(X$, 9, 8))
      STDIS1 = CVS(Mid$(X$, 21, 4))
      FALTA1 = CVS(Mid$(X$, 25, 4))
      '
      C2$ = CODEXT$(CD1%)
      D2$ = DESCRIT0$(CD1%)
      UNIME2$ = Unimed$(CD1%)
      STDIS1 = STOCKDIS(CD1%)
      FALTA1 = CACONJU1 - STDIS1: If FALTA1 < 0 Then FALTA1 = 0
      'RECALCULO DE EL USO UNITARIO
      On Error Resume Next 'ATAJO DIVISON X 0
      USOU1# = CACONJU1 / Val(Text5)
      On Error GoTo 0
      'USO UNITARIO SI NO TIENE USO UNITARO NOMINAL
'      If USOU1# = 0 Then
'        On Error Resume Next
'        USOU1# = CACONJU1 / Val(Text5)
'        On Error GoTo 0
'      End If
      USOUNI1 = USOU1#
      USOUNI2$ = FORMATNUM(USOUNI1, "F8.3")
      CANESI1 = CACONJU1
      CANESI2$ = FORMATNUM$(CANESI1, "F8.1")
      STDIS2$ = FORMATNUM$(STDIS1, "F8.1")
      FALTA2$ = FORMATNUM$(FALTA1, "F8.1")
      'ARMO EL LIST1
      L2$ = Space$(128)
      Call REPLA(L2$, C2$, 1, 16)
      Call REPLA(L2$, D2$, 17, 24)
      Call REPLA(L2$, UNIME2$, 42, 4)
      Call REPLA(L2$, USOUNI2$, 47, 8)
      Call REPLA(L2$, CANESI2$, 57, 8)
      Call REPLA(L2$, STDIS2$, 68, 8)
      Call REPLA(L2$, FALTA2$, 77, 8)
      List1.AddItem L2$
      '
    Next JX&
  End If
  '
End Sub

Private Sub List1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  'MUESTRA LA DESCRIPCIÓN DEL ITEM SELECCIONADO EN LA LISTA
  CODI1$ = Left$(List1.List(List1.ListIndex), 16)
  CI1% = CODINT%(CODI1$)
  List1.ToolTipText = DESCRIT0$(CI1%)
  '
End Sub

Private Sub Text1_Change()
  '
  CI1% = CODINT%(Text1.TEXT)
  If CI1% > 0 Then
       Label5.Caption = DESCRIT$(CI1%)
       Call RECALCUOF
     Else
       Label5.Caption = ""
  End If
  '
End Sub

Private Sub Text1_DblClick()
  '
  Call SELECHO("MASTER")
  COD$ = VALACT1$("MASTER")
  Text1.TEXT = COD$
  CI% = CODINT%(COD$)
  Label5.Caption = DESCRIT$(CI%)
  Label7.Caption = Unimed$(CI%)
  '
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
  '
  If KeyAscii = 13 Then
    On Error Resume Next
      Text2.SetFocus
    On Error GoTo 0
  End If
  '
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
  '
  If KeyAscii = 13 Then
    On Error Resume Next
      Text5.SetFocus
    On Error GoTo 0
  End If
  '
End Sub

Private Sub Text3_DblClick()
  '
  CI1% = CODINT%(Text1.TEXT)
  If CI1% > 0 Then
    Call Command3_Click
  End If
  '
End Sub

Private Sub Text4_DblClick()
  '
  CI1% = CODINT%(Text1.TEXT)
  If CI1% > 0 Then
    Call Command5_Click
  End If
  '
End Sub
Private Sub Text5_KeyPress(KeyAscii As Integer)
  '
  If KeyAscii = 13 Then
    On Error Resume Next
      Text3.SetFocus
    On Error GoTo 0
  End If
  '
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
  '
  If KeyAscii = 13 Then
    On Error Resume Next
      Text4.SetFocus
    On Error GoTo 0
  End If
  '
End Sub
Private Sub Text4_KeyPress(KeyAscii As Integer)
  '
  If KeyAscii = 13 Then
    On Error Resume Next
      Text6.SetFocus
    On Error GoTo 0
  End If
  '
End Sub

Private Sub Text5_LostFocus()
  '
  T5 = Val(Text5.TEXT)
  Text5.TEXT = FORMATNUM(T5, "F10.1")
  CI1% = CODINT%(Text1.TEXT)
  If CI1% > 0 Then
    Call RECALCUOF 'RECALCULO DE LA OF
  End If
  '
End Sub

Private Sub Text6_Change()
  '
  Label15.Caption = TRIM$(ECOARCH$("DESTFAB", Text6.TEXT))
  '
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
  '
  If KeyAscii = 13 Then
    On Error Resume Next
      Text7.SetFocus
    On Error GoTo 0
  End If
  '
End Sub

Private Sub Text1_GotFocus()
  '
  Text1.SelStart = 0
  Text1.SelLength = Len(Text1)
  '
End Sub

Private Sub Text2_gotfocus()
  '
  Text2.SelStart = 0
  Text2.SelLength = Len(Text2)
  '
End Sub
Private Sub Text3_GotFocus()
  '
  Text3.SelStart = 0
  Text3.SelLength = Len(Text3)
  '
End Sub
Private Sub Text4_GotFocus()
  '
  Text4.SelStart = 0
  Text4.SelLength = Len(Text4)
  '
End Sub
Private Sub Text5_GotFocus()
  '
  Text5.SelStart = 0
  Text5.SelLength = Len(Text5)
  '
End Sub
Private Sub Text6_GotFocus()
  '
  Text6.SelStart = 0
  Text6.SelLength = Len(Text6)
  '
End Sub

Private Sub Text7_Change()
  '
  CS1 = Val(Text7)
  Label17.Caption = TRIM$(ECOARCH$("SECPROD", MKS$(CS1)))
  '
End Sub

Private Sub Text7_GotFocus()
  '
  Text7.SelStart = 0
  Text7.SelLength = Len(Text7)
  '
End Sub
Private Sub Text6_DblClick()
  '
  CI1% = CODINT%(Text1.TEXT)
  If CI1% > 0 Then
    Call Command6_Click
  End If
  '
End Sub

Private Sub Text7_DBLCLICK()
  '
  CI1% = CODINT%(Text1.TEXT)
  If CI1% > 0 Then
    Call Command7_Click
  End If
  '
End Sub
Sub LIMPIAOF() 'LIMPIA FORMULARIO
  '
  Text1.TEXT = ""
  Text2.TEXT = ""
  Text3.TEXT = ""
  Text4.TEXT = ""
  Text5.TEXT = ""
  Text6.TEXT = ""
  Text7.TEXT = ""
  '
  Label4.Caption = ""
  Label5.Caption = ""
  Label7.Caption = ""
  Label11.Caption = ""
  Label15.Caption = ""
  Label17.Caption = ""
  '
  List1.Clear
  List2.Clear
  '
  HOOI% = HOY(HOS$)
  Label11.Caption = HOS$
  Text4.TEXT = HOS$
  Text3.TEXT = HOS$
  'NUMERO DE PROXIMA OF
  NRO& = 1
  Call ABREORFAB
  '
  SQLX$ = "SELECT NumeOrFa FROM ORDEFABR "
  SQLX$ = SQLX$ + "ORDER BY NumeOrFa DESC "
  Dim ORFABTEMP As ADODB.Recordset
  Set ORFABTEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
  'Set ORFABTEMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
  On Error Resume Next
  With ORFABTEMP
     .MoveFirst
     If Err < 1 Then
       NRO& = 1 + !NUMEORFA
     End If
  End With
  On Error GoTo 0
  '
  Label4 = FORMATNUM$(NRO&, "F9.0")
  '
End Sub
Sub RECALCUOF() 'RECALCULO DE LA OF
  '
  List1.Clear
  List2.Clear
  CODICONJ1$ = Text1.TEXT 'CODIGO DEL ARTICULO CONJUNTO
  CANTICO1# = Val(Text5)
  CAPENDE# = CANTICO1#
  Call CARCOMPOF(CODICONJ1$, CANTICO1#) 'CARGA DE COMPONENTES
  Call CAROPEROF(CODICONJ1$, CAPENDE#) 'CARGA DE OPERACIONES
  '
End Sub
Sub CAROPEROF(CODICONJ1$, CAPENDE#) 'CARGA DE OPERACIONES
  '
  Call ABREORFAB
  '
  CIJK% = CODINT%(CODICONJ1$)
  '
  If CIJK% > 0 Then
    SQLX$ = "SELECT * FROM SecOper WITH(NOLOCK) "
    SQLX$ = SQLX$ + " WHERE CodIConj=" & CIJK% & " "
    SQLX$ = SQLX$ + " AND StatOper = 1 ORDER BY NumeOper ASC"
'    Dim SECOPTEMP As ADODB.Recordset
'    Set SECOPTEMP = New ADODB.Recordset
'    SECOPTEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
   Set SECOPTEMP = READSET(SQLX$)
'    Set SECOPTEMP = Articulos.OpenRecordset("SELECT * FROM SecOper WHERE CodIConj=" & CIJK% & " AND StatOper = 1 ORDER BY NumeOper ASC", dbOpenSnapshot)
    On Error Resume Next
    SECOPTEMP.MoveFirst
    If Err < 1 Then
      With SECOPTEMP
        Do While Not .EOF
          NO% = !numeoper
          NO3$ = FORMATNUM(NO%, "F3.0")
          DOPE$ = !descoper
          EQ1$ = !CodMaq0
          EQAUX$ = AJUSTI$(!CodMaq1, 6) + AJUSTI$(!CodMaq2, 6) + AJUSTI$(!CodMaq1, 6)
          LOTESTA = !LOTESTAN: If LOTESTA < 1 Then LOTESTA = 1
          CALOTES = 1
          CAOPS = !CantiOps
            If CAOPS < 0.01 Then CAOPS = 1
          CAOPS1$ = FORMATNUM(CAOPS, "F6.2")
          MINFIJO = 60 * !HorsFijo * CALOTES
          HSFIJO = MINFIJO / 60
            HSFIJO1$ = FORMATNUM(HSFIJO, "F7.1")
          MINVARI = CAPENDE# * !MinuStan
          HSVARI = MINVARI / 60
            HSVARI1$ = FORMATNUM(HSVARI, "F7.1")
          HSTOTA = HSVARI + HSFIJO
            HSTOTA1$ = FORMATNUM(HSTOTA, "F7.1")
          PDEMO = !PorDemor
            PDEMO1$ = FORMATNUM$(PDEMO, "F4.1")
          'ARMO EL LIST2
          L3$ = Space(128)
          Call REPLA(L3$, " ", 1, 1)
          Call REPLA(L3$, NO3$, 2, 3)
          Call REPLA(L3$, DOPE$, 6, 25)
          Call REPLA(L3$, EQ1$, 32, 6)
          Call REPLA(L3$, HSFIJO1$, 39, 7)
          Call REPLA(L3$, HSVARI1$, 47, 7)
          Call REPLA(L3$, PDEMO1$, 55, 4)
          Call REPLA(L3$, HSTOTA1$, 60, 7)
          Call REPLA(L3$, CAOPS1$, 69, 6)
          Call REPLA(L3$, EQAUX$, 76, 18)
          List2.AddItem L3$
        .MoveNext
        Loop
      End With
    End If
    '
    OperFab.Close
    Set OperFab = Nothing
  End If
  '
End Sub
Sub CARCOMPOF(CODICONJ$, CACONJU#) 'CARGA DE COMPONENTES
  Screen.MousePointer = 11
  '
  List1.Clear
  CODPRO$ = CODICONJ$
  COEF = 1
  'ARMA COMPOSICIÓN SEGÚN FORMULA STANDARD
  TITUPA$ = "Consumo por Fórmula Standard"
  Call LEEEXPLOMUL(CODPRO$, 1, 1)
  If CAIT% < 1 Then  'PRODUCTO SIN ESTRUCTURA
    Call MENSERR(24, "No hay Fórmula de Consumo Standard\para este Producto (" + TRIM$(CODPRO$) + ").")
    TITUPA$ = "Consumo Informado"
  End If
  '
  FIN& = CAIT%
  For I1% = 1 To CAIT%
     UI1& = I1%
     Call TRACE(20, UI1&, FIN&)
     If CIJ%(I1%) > 0 Then
       If Abs(USOI(I1%)) > 0 Then
         '
         C2$ = CODEXT$((CIJ%(I1%)))
         D2$ = DESCRIT0$(CIJ%(I1%))
         UNIME2$ = Unimed$(CIJ%(I1%))
         USOUNI1 = USOI(I1%) * COEF
         USOUNI2$ = FORMATNUM(USOUNI1, "F8.3")
         CANESI1 = CACONJU# * USOI(I1%) * COEF
         CANESI2$ = FORMATNUM$(CANESI1, "F8.1")
         STDIS1 = STOCKDIS(CIJ%(I1%))
         STDIS2$ = FORMATNUM$(STDIS1, "F8.1")
         FALTA1 = CANESI1 - STDIS1: If FALTA1 < 0 Then FALTA1 = 0
         FALTA2$ = FORMATNUM$(FALTA1, "F8.1")
         'ARMO EL LIST1
         L2$ = Space$(128)
         Call REPLA(L2$, C2$, 1, 16)
         Call REPLA(L2$, D2$, 17, 24)
         Call REPLA(L2$, UNIME2$, 42, 4)
         Call REPLA(L2$, USOUNI2$, 47, 8)
         Call REPLA(L2$, CANESI2$, 57, 8)
         Call REPLA(L2$, STDIS2$, 68, 8)
         Call REPLA(L2$, FALTA2$, 77, 8)
         List1.AddItem L2$
         '
       End If
     End If
  Next I1%
  Screen.MousePointer = 1
  '
End Sub
