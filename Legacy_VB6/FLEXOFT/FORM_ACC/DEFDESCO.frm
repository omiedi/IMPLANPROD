VERSION 5.00
Begin VB.Form DEFDESCO 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Definición de Descuentos Combinados"
   ClientHeight    =   5865
   ClientLeft      =   165
   ClientTop       =   5820
   ClientWidth     =   6015
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
   ScaleHeight     =   5865
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture6 
      Height          =   5895
      Left            =   4935
      ScaleHeight     =   5835
      ScaleWidth      =   1020
      TabIndex        =   27
      Top             =   -20
      Visible         =   0   'False
      Width           =   1080
      Begin VB.Image Image3 
         Height          =   5955
         Left            =   0
         Picture         =   "DEFDESCO.frx":0000
         Stretch         =   -1  'True
         Top             =   -120
         Width           =   1095
      End
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
      Left            =   1155
      TabIndex        =   25
      Top             =   5250
      Width           =   2025
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Registrar"
      Height          =   285
      Left            =   3360
      TabIndex        =   24
      Top             =   5250
      Width           =   1275
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Por Clientes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2325
      Left            =   210
      TabIndex        =   0
      Top             =   210
      Width           =   4500
      Begin VB.CommandButton Command2 
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
         Picture         =   "DEFDESCO.frx":2398
         TabIndex        =   14
         Top             =   1425
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
         Picture         =   "DEFDESCO.frx":26A2
         TabIndex        =   13
         Top             =   375
         Width           =   175
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00EEEEDD&
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo de Clientes:"
         Height          =   225
         Left            =   735
         TabIndex        =   20
         Top             =   1470
         Width           =   2430
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Left            =   210
         TabIndex        =   16
         Top             =   1845
         Width           =   4110
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
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
         Left            =   3255
         TabIndex        =   15
         Top             =   1425
         Width           =   855
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00EEEEDD&
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
         Height          =   225
         Left            =   735
         TabIndex        =   6
         Top             =   435
         Width           =   2430
      End
      Begin VB.Label Label11 
         BorderStyle     =   1  'Fixed Single
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
         Left            =   3255
         TabIndex        =   2
         Top             =   375
         Width           =   855
      End
      Begin VB.Label Label8 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Left            =   210
         TabIndex        =   1
         Top             =   795
         Width           =   4110
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00062A70&
      Height          =   6000
      Left            =   4935
      ScaleHeight     =   5940
      ScaleWidth      =   1320
      TabIndex        =   3
      Top             =   -20
      Width           =   1380
      Begin VB.CommandButton Command6 
         Caption         =   "Check"
         Height          =   690
         Left            =   105
         Picture         =   "DEFDESCO.frx":29AC
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Verificación por Cliente y Producto"
         Top             =   2835
         Width           =   640
      End
      Begin VB.CommandButton Command8 
         Caption         =   "List"
         Height          =   690
         Left            =   105
         Picture         =   "DEFDESCO.frx":2CB6
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Listados de Descuentos"
         Top             =   1050
         Width           =   640
      End
      Begin VB.CommandButton Command3 
         Height          =   640
         Left            =   105
         Picture         =   "DEFDESCO.frx":2FC0
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Cierra y Abandona la Aplicación"
         Top             =   315
         Width           =   640
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   4
         Top             =   4830
         Width           =   640
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   5
            Top             =   105
            Width           =   12000
         End
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "DEFDESCO.frx":310A
         Top             =   5145
         Width           =   1515
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Por Productos "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2325
      Left            =   210
      TabIndex        =   7
      Top             =   2730
      Width           =   4500
      Begin VB.CommandButton Command4 
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
         Picture         =   "DEFDESCO.frx":502C
         TabIndex        =   17
         Top             =   1440
         Width           =   175
      End
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
         Picture         =   "DEFDESCO.frx":5336
         TabIndex        =   11
         Top             =   390
         Width           =   175
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00EEEEDD&
         BackStyle       =   0  'Transparent
         Caption         =   "Linea:"
         Height          =   225
         Left            =   840
         TabIndex        =   21
         Top             =   1515
         Width           =   2430
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
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
         Left            =   3360
         TabIndex        =   19
         Top             =   1440
         Width           =   750
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Left            =   210
         TabIndex        =   18
         Top             =   1860
         Width           =   4110
      End
      Begin VB.Label Label9 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Left            =   210
         TabIndex        =   10
         Top             =   810
         Width           =   4110
      End
      Begin VB.Label Label12 
         BorderStyle     =   1  'Fixed Single
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
         Left            =   2100
         TabIndex        =   9
         Top             =   390
         Width           =   2010
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00EEEEDD&
         BackStyle       =   0  'Transparent
         Caption         =   "Código: "
         Height          =   225
         Left            =   -420
         TabIndex        =   8
         Top             =   465
         Width           =   2430
      End
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEEEDD&
      BackStyle       =   0  'Transparent
      Caption         =   "Valor:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   -735
      TabIndex        =   22
      Top             =   5250
      Width           =   1800
   End
   Begin VB.Menu mnua 
      Caption         =   "Archivo"
      Begin VB.Menu nnus 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu mnulis 
      Caption         =   "Listados"
      Begin VB.Menu numlisdesc 
         Caption         =   "Listados de Descuentos"
      End
   End
End
Attribute VB_Name = "DEFDESCO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   'Call CommSel("MASTER", Command1, Label12)
   Call SELECHO("MASTER")
   Label12 = VALACT1$("MASTER")

End Sub


Private Sub Command2_Click()
   Call CommSel("Grucocli", Command2, Label1)

End Sub

Private Sub Command3_Click()
   Call CIERRARCH("*.*")
   Hide
End Sub

Private Sub Command4_Click()
   Call CommSel("grucoven", Command4, Label4)

End Sub
Private Sub Command6_Click()
    Command6.Enabled = False
    CHECKDESCU.Show 1
    Command6.Enabled = True
End Sub

Private Sub Command5_Click()
   'graba el descuento asociado
   '
   ABUS$ = ""
   NC% = XVALO(Label11)
   GRUP% = XVALO(Label1)
   CODI% = CODINT(TRIM$(Label12))
   LNEA% = XVALO(Label4)
   PORDESC$ = TRIM$(Text1)
   ABUS$ = MKI$(NC%) + MKI$(GRUP%) + MKI$(CODI%) + MKI$(LNEA%)
   Call REPLA(RECORD$, ABUS$, 1, 8)
   Call REPLA(RECORD$, MKI(NC%), 17, 2)
   Call REPLA(RECORD$, MKI(GRUP%), 19, 2)
   Call REPLA(RECORD$, MKI(CODI%), 21, 2)
   Call REPLA(RECORD$, MKI(LNEA%), 23, 2)
   '
   Call REPLA(RECORD$, MKI(NC%), 33, 2)
   Call REPLA(RECORD$, MKI(GRUP%), 35, 2)
   Call REPLA(RECORD$, PORDESC$, 49, 12)
   '
   Call FILTERPUTRECORD("DESCOMBI", 1, ABUS$, RECORD$)
   Call CIERRARCH("DESCOMBI")
End Sub

Private Sub Command8_Click()
   Call COPYSTRU("Descombi", "Ldescomb")
   '
10 OPX% = ALTERNA%("Por Cliente\Por Grupos de Clientes\Por Productos\Por Linea de Productos")
   If OPX% < 1 Or OPX% > 4 Then Exit Sub
   '
   U& = 0
   '
   FIN& = ULTREG&("Descombi")

   If OPX% = 1 Then 'POR CLIENTE
      For I& = 1 To FIN&
        X$ = REGLEIDO$("Descombi", I&)
        If CVI(Mid$(X$, 17, 2)) > 0 Then
        AA = CVI(Mid$(X$, 17, 2))
          U& = U& + 1
          Call GRAREG("Ldescomb", X$, U&)
        End If
      Next I&
      Call SETULTREG("LDESCOMB", U&)
      Call CIERRARCH("Ldescomb")

      
      Call FILESORT("Ldescomb", "", 3, 10, "ASC")
      Call FINAL("*LIDESCLI")
     '
     ElseIf OPX% = 2 Then 'POR GRUPO CLIENTE
      For I& = 1 To FIN&
        X$ = REGLEIDO$("Descombi", I&)
        If CVI(Mid$(X$, 19, 2)) > 0 Then
          U& = U& + 1
          Call GRAREG("Ldescomb", X$, U&)
        End If
      Next I&
      Call SETULTREG("LDESCOMB", U&)
      Call CIERRARCH("Ldescomb")

      Call FILESORT("Ldescomb", "", 5, 13, "ASC")
      Call FINAL("*LIDESGRU")

     '
     ElseIf OPX% = 3 Then 'POR PRODUCTO
      For I& = 1 To FIN&
        X$ = REGLEIDO$("Descombi", I&)
        If CVI(Mid$(X$, 21, 2)) > 0 Then
          U& = U& + 1
          Call GRAREG("Ldescomb", X$, U&)
        End If
      Next I&
      Call SETULTREG("LDESCOMB", U&)
      Call CIERRARCH("Ldescomb")
      
      Call FILESORT("Ldescomb", "", 7, 15, "ASC")
      Call FINAL("*LIDESART")

     ElseIf OPX% = 4 Then 'POR LINEA
      For I& = 1 To FIN&
        X$ = REGLEIDO$("Descombi", I&)
        If CVI(Mid$(X$, 23, 2)) > 0 Then
          U& = U& + 1
          Call GRAREG("Ldescomb", X$, U&)
        End If
      Next I&
      Call SETULTREG("LDESCOMB", U&)
      Call CIERRARCH("Ldescomb")

      Call FILESORT("Ldescomb", "", 9, 15, "ASC")
      Call FINAL("*LIDESLIN")

   End If
   '
   GoTo 10
   '
End Sub

Private Sub Command9_Click()
   Call CommSel("DEUDOR1", Command9, Label11)
End Sub

Private Sub Form_Load()
    If Control("NEWASPEC", "CAMBIOAS") = "SI" Then
      Picture6.Visible = True
      Me.BackColor = &HFCD7B6
      DEFDESCO.Frame1.BackColor = &HFCD7B6
      DEFDESCO.Frame5.BackColor = &HFCD7B6
    End If

End Sub

Private Sub label1_Change()
   NCX% = XVALO(Label1)
   If NCX% > 0 Then
       Label2 = ECOARCH$("GRUCOCLI", MKI$(NCX%))
     Else
       Label2 = "Cualquier Grupo de Clientes"
   End If
   If Label1 <> "" Then
      Label11 = ""
      Label8 = ""
   End If

   Call BUSCAVALOR
End Sub

Private Sub Label10_Click()
   Label12 = ""
End Sub

Private Sub Label11_Change()
   NCX% = XVALO(Label11)
   If NCX% > 0 Then
       Label8 = RASOCLI$(NCX%)
     Else
       Label8 = "Cualquier Cliente"
   End If
   '
   If Label11 <> "" Then
      Label1 = ""
   End If

   Call BUSCAVALOR

End Sub

Private Sub Label12_Change()
   CI% = CODINT%(Label12)
   If CI% > 0 Then
       Label9 = DESCRIT0$(CI%)
     Else
       Label9 = "Todos los Articulos"
   End If
   '
   If Label12 <> "" Then
      Label4 = ""
      Label3 = ""
   End If

   Call BUSCAVALOR

End Sub

Private Sub Label2_Change()
   If Label8 = "" Then Label8.Caption = "Cualquier Grupo de Clientes"
End Sub

Private Sub Label3_Change()
   If Label3 = "" Then Label3.Caption = "Cualquier Linea de Productos"

End Sub

Private Sub Label4_Change()
   NCX% = XVALO(Label4)
   If NCX% > 0 Then
       Label3 = ECOARCH$("GRUCOVEN", MKI$(NCX%))
     Else
       Label3 = "Cualquier Línea"
   End If
   If Label4 <> "" Then
      Label12 = ""
      Label9 = ""
   End If
   Call BUSCAVALOR
End Sub

Private Sub Label5_DblClick()
   Label11 = ""
End Sub

Private Sub Label6_Click()
  Label1 = ""
End Sub

Private Sub Label7_Click()
   Label4 = ""
End Sub

Private Sub Label8_Change()
   If Label8 = "" Then Label8.Caption = "Cualquier Cliente"
End Sub

Private Sub Label9_Change()
   If Label9 = "" Then Label9.Caption = "Cualquier Producto"
End Sub

Sub BUSCAVALOR()
   ABUS$ = ""
   NC% = XVALO(Label11)
   GRUP% = XVALO(Label1)
   CODI% = CODINT(Label12)
   LNEA% = XVALO(Label4)
   '
   ABUS$ = MKI$(NC%) + MKI$(GRUP%) + MKI$(CODI%) + MKI$(LNEA%)
   X$ = FILTERGETRECORD$("DESCOMBI", 1, (ABUS$))
   Text1 = TRIM$(Mid$(X$, 49, 12))
End Sub

Private Sub mnucontrol_Click()
    Call Command6_Click
End Sub

Private Sub nnus_Click()
   Call Command3_Click
End Sub

