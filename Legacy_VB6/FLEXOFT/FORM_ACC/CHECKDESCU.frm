VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form CHECKDESCU 
   BackColor       =   &H00E2D3C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Verificación de Descuentos Combinados"
   ClientHeight    =   3645
   ClientLeft      =   165
   ClientTop       =   5835
   ClientWidth     =   5565
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   5565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture16 
      Height          =   375
      Left            =   1680
      ScaleHeight     =   315
      ScaleWidth      =   1155
      TabIndex        =   19
      Top             =   3720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox Picture999 
      Height          =   225
      Index           =   0
      Left            =   1470
      ScaleHeight     =   165
      ScaleWidth      =   1110
      TabIndex        =   17
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2D3C0&
      Caption         =   "DATOS DE ENTRADA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Left            =   80
      TabIndex        =   0
      Top             =   120
      Width           =   4500
      Begin VB.CommandButton Command1 
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
         Left            =   4095
         Picture         =   "CHECKDESCU.frx":0000
         TabIndex        =   8
         Top             =   1365
         Width           =   175
      End
      Begin VB.CommandButton Command9 
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
         Left            =   4095
         Picture         =   "CHECKDESCU.frx":030A
         TabIndex        =   6
         Top             =   375
         Width           =   175
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00EEEEDD&
         BackStyle       =   0  'Transparent
         Caption         =   "Código"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -420
         TabIndex        =   11
         Top             =   1440
         Width           =   2430
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2100
         TabIndex        =   10
         Top             =   1365
         Width           =   2010
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   210
         TabIndex        =   9
         Top             =   1785
         Width           =   4110
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00EEEEDD&
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   735
         TabIndex        =   4
         Top             =   450
         Width           =   2430
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3255
         TabIndex        =   2
         Top             =   375
         Width           =   855
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   210
         TabIndex        =   1
         Top             =   795
         Width           =   4110
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00624E28&
      Height          =   6000
      Left            =   4680
      ScaleHeight     =   5940
      ScaleWidth      =   1320
      TabIndex        =   3
      Top             =   0
      Width           =   1380
      Begin VB.CommandButton Command3 
         Height          =   640
         Left            =   105
         Picture         =   "CHECKDESCU.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Cierra y Abandona la Aplicación"
         Top             =   240
         Width           =   640
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2625
      OleObjectBlob   =   "CHECKDESCU.frx":075E
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "CHECKDESCU.frx":0992
      TabIndex        =   18
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEEEDD&
      BackStyle       =   0  'Transparent
      Caption         =   "Modo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   -240
      TabIndex        =   16
      Top             =   3330
      Width           =   840
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   720
      TabIndex        =   15
      Top             =   3240
      Width           =   3570
   End
   Begin VB.Label Label103 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   3225
      TabIndex        =   14
      Top             =   2820
      Width           =   1065
   End
   Begin VB.Label LabeL102 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   285
      TabIndex        =   13
      Top             =   2820
      Width           =   2325
   End
   Begin VB.Label Label101 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEEEDD&
      BackStyle       =   0  'Transparent
      Caption         =   "Porciento"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2250
      TabIndex        =   12
      Top             =   2580
      Width           =   1800
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEEEDD&
      BackStyle       =   0  'Transparent
      Caption         =   "Valor"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   -1080
      TabIndex        =   7
      Top             =   2580
      Width           =   1800
   End
   Begin VB.Menu archivo 
      Caption         =   "Archivo"
      Begin VB.Menu salir 
         Caption         =   "Salir"
      End
   End
End
Attribute VB_Name = "CHECKDESCU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   Call CommSel("MASTER", Command1, Label12)
End Sub

Private Sub Command3_Click()
   Call CIERRARCH("*.*")
   'Hide
   Unload Me
End Sub

Private Sub Command9_Click()
   Call CommSel("DEUDOR1", Command9, Label11)
End Sub

Private Sub Form_Load()
    If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      Picture16.Visible = True
      Picture16.ZOrder 0
      Me.BackColor = &HFCD7B6
      Me.Frame1.BackColor = &HFCD7B6
      MARCOBARRA.Top = 3400
      Set MARCOBARRA.Container = Me.Picture16
      MARCOBARRA.ZOrder 0
    End If

   'Call APLICACIONSKINS(Me, Picture1)
End Sub

Private Sub Label10_Click()
   Label12 = ""
End Sub

Private Sub LabeL102_Change()
   Label103 = FORMATNUM$(PORCEDESCU(LabeL102), "F8.5")
End Sub

Private Sub Label11_Change()
   NCX% = XVALO(Label11)
   If NCX% > 0 Then
       Label8 = rasocli$(NCX%)
     Else
       Label8 = ""
       Label1 = ""
   End If
   '
   CIX% = CODINT%(Label12)
   LabeL102 = DESCOMBINA$(NCX%, CIX%)
   Label1 = MODODESCU$(NCX%, CIX%)

End Sub
Function MODODESCU$(NCX%, CIX%)

   'CARGA DE MODO DE DESCUENTO O BONIFICACION
   MODODESCU$ = "Por Precio"
   CONDI$ = "Nume_Cli = " & NCX%
   CONDI$ = CONDI$ & " AND Cod_Inte =" & CIX%
   MODODESC% = XVALO(VALOBADA("DESCOMBI", "MODO_DESC", CONDI$))
   If MODODESC% > 0 Then
     MODODESCU$ = "Por Producto Bonificado"
   End If
End Function
Private Sub Label12_Change()

   Label1 = ""
   CI% = CODINT%(Label12)
   Label9 = DESCRIT0$(CI%)
   '
   NCX% = XVALO(Label11)
   LabeL102 = DESCOMBINA$(NCX%, CI%)
   Label1 = MODODESCU$(NCX%, CIX%)

End Sub

Private Sub Label5_DblClick()
   Label11 = ""
End Sub

Private Sub Label8_Change()
   If Label8 = "" Then Label8.Caption = "Cualquier Cliente"
End Sub

Private Sub Label9_Change()
   If Label9 = "" Then Label9.Caption = "Cualquier Producto"
End Sub

Private Sub salir_Click()
   Call Command3_Click
End Sub
