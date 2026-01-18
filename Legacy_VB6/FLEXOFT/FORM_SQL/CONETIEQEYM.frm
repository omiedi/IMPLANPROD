VERSION 5.00
Begin VB.Form CONETIEQEYM 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración de Formularios"
   ClientHeight    =   7290
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   6225
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   6225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "T.de Mat. Configurado Para  Producto a la Venta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Left            =   240
      TabIndex        =   20
      Top             =   6000
      Width           =   4920
      Begin VB.TextBox Text6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   23
         Text            =   " "
         Top             =   360
         Width           =   540
      End
      Begin VB.CommandButton Command10 
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
         Height          =   300
         Left            =   4560
         TabIndex        =   22
         Top             =   360
         Width           =   180
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Text            =   " "
         Top             =   720
         Width           =   4575
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Material"
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
         Left            =   1560
         TabIndex        =   24
         Top             =   420
         Width           =   2325
      End
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Impresion de Etiqueta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5610
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   4950
      Begin VB.CommandButton Command11 
         Caption         =   "Abrir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   4095
         Picture         =   "CONETIEQEYM.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   1260
         Width           =   645
      End
      Begin VB.TextBox Text8 
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
         TabIndex        =   32
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1620
         Width           =   3750
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
         Height          =   300
         Left            =   3780
         TabIndex        =   31
         Top             =   1260
         Width           =   180
      End
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2400
         TabIndex        =   30
         Text            =   " "
         Top             =   1275
         Width           =   1380
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2430
         TabIndex        =   28
         Text            =   " "
         Top             =   3240
         Width           =   1380
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
         Height          =   300
         Left            =   3795
         TabIndex        =   27
         Top             =   3240
         Width           =   180
      End
      Begin VB.TextBox Text4 
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
         Left            =   225
         TabIndex        =   26
         TabStop         =   0   'False
         Text            =   " "
         Top             =   3600
         Width           =   3750
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Abrir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   4110
         Picture         =   "CONETIEQEYM.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   3240
         Width           =   645
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Imprimir Etiq. Extras Relacionadas C/N° de Serie"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   200
         TabIndex        =   19
         Top             =   5235
         Width           =   4500
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Asignar Lotes a Equipos x Lectura de Etiqueta "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   200
         TabIndex        =   18
         Top             =   4920
         Width           =   4600
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "NO Utilizar Armadores Externos(Deberá Configurarlos)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   200
         TabIndex        =   17
         ToolTipText     =   "Se Deberá Informar Los Armadores desde Setup de Fabricación"
         Top             =   4515
         Width           =   4600
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Impresión Autómatica Cuando se Genera La O.F."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   200
         TabIndex        =   16
         Top             =   4200
         Width           =   4600
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Abrir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   4095
         Picture         =   "CONETIEQEYM.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   2400
         Width           =   645
      End
      Begin VB.TextBox Text2 
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
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   " "
         Top             =   2760
         Width           =   3750
      End
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
         Height          =   300
         Left            =   3780
         TabIndex        =   10
         Top             =   2400
         Width           =   180
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2415
         TabIndex        =   9
         Text            =   " "
         Top             =   2400
         Width           =   1380
      End
      Begin VB.TextBox Text10 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2400
         TabIndex        =   7
         Text            =   " "
         Top             =   360
         Width           =   1380
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
         Height          =   300
         Left            =   3780
         TabIndex        =   6
         Top             =   345
         Width           =   180
      End
      Begin VB.TextBox Text9 
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
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   " "
         Top             =   705
         Width           =   3750
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Abrir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   4095
         Picture         =   "CONETIEQEYM.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   345
         Width           =   645
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Frm Etiqueta *  C-15 *  Grde.  N° de Serie:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   0
         TabIndex        =   34
         Top             =   1200
         Width           =   2325
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Frm. Impresión Etiq.Peq.(sin Qr) N° Serie:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   1
         Left            =   120
         TabIndex        =   29
         Top             =   3135
         Width           =   2325
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Frm. Impresión Etiq.Peq.(sin Qr) Orden:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   0
         Left            =   105
         TabIndex        =   13
         Top             =   2280
         Width           =   2310
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Frm Etiqueta Grde.  N° de Serie:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   0
         TabIndex        =   8
         Top             =   285
         Width           =   2325
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00404040&
      Height          =   7140
      Left            =   5400
      ScaleHeight     =   7080
      ScaleWidth      =   1140
      TabIndex        =   0
      Top             =   0
      Width           =   1200
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Left            =   100
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   14
         Top             =   3960
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   500
            Left            =   120
            TabIndex        =   15
            Top             =   120
            Width           =   12000
         End
      End
      Begin VB.CommandButton command1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   105
         Picture         =   "CONETIEQEYM.frx":0C28
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Salir"
         Top             =   315
         Width           =   650
      End
      Begin VB.CommandButton Command2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   120
         Picture         =   "CONETIEQEYM.frx":0D72
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Guardar"
         Top             =   840
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "CONETIEQEYM.frx":107C
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   4200
         Width           =   1515
      End
   End
End
Attribute VB_Name = "CONETIEQEYM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub
'\\\OT-05-0723-WAR-22/11/05///

Private Sub Command10_Click()
    Campos$ = "Código/f6;Tipo de Material/A32"
    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, Campos$, FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
    FRMZELECHO.Caption = "TABLA TIPO DE MATERIAL"
    
    Dim obj As New RECORXET
    Set rs1 = obj.RS_T_MATERIAL_SQL()
    Call Carga_MSF_Recordset(rs1, Campos$, FRMZELECHO.msf2, 1)
    On Error Resume Next
    Set obj = Nothing
    rs1.Close
    Set rs1 = Nothing
    On Error GoTo 0
        
    FRMZELECHO.Width = 8000
50  FRMZELECHO.Show 1

    Text6 = TRIM(RESP_ZELE_VECT(1))
    Text3 = TRIM(RESP_ZELE_VECT(2))

End Sub

Private Sub Command11_Click()
     Call CONECRUN("FLFORMS/" + TRIM$(Text7.TEXT), "Formularios de Impresion")
End Sub

Private Sub Command2_Click()
    '
    Call GRACONTROL("", "STIKAPRO", Text10.TEXT)
    Call GRACONTROL("", "STIKAPR2", Text1.TEXT)
    Call GRACONTROL("", "STIKAPR3", Text5.TEXT)
    Call GRACONTROL("", "STIKAC15", Text7.TEXT)
    '
    Call GRACONTROL("", "DEFPROTR", Text6.TEXT)
    '
    IMPRAUTO$ = "NO"
    If Check4.Value = 1 Then IMPRAUTO$ = "SI"
    Call GRACONTROL(IDENTER$, "IMPAUTOM", IMPRAUTO$)
    '
    IMPRAUTO$ = "NO"
    If Check2.Value = 1 Then IMPRAUTO$ = "SI"
    Call GRACONTROL("CONFTRAZ", "ASIGLOTE", IMPRAUTO$)

    MONUME$ = ""
    If Check1.Value = 1 Then MONUME$ = "NO"
    Call GRACONTROL("NUMSERIE", "UTIARMEX", MONUME$)
    '
    MONUME$ = ""
    If Check3.Value = 1 Then MONUME$ = "SI"
    Call GRACONTROL("NUMSERIE", "IMPETIEX", MONUME$) ' IMPRIMIR ETIQUETAS EXTRAS CASO EYM
    '
    
    Call CIERRARCH("*.*")
    Hide
    '
End Sub

Private Sub Command3_Click()
   Call CONECRUN("FLFORMS/" + TRIM$(Text5.TEXT), "Formularios de Impresion")

End Sub
'

Private Sub Command4_Click()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text1.TEXT = VDEVU$
    End If
End Sub


Private Sub Command5_Click()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text5.TEXT = VDEVU$
    End If

End Sub

Private Sub Command6_Click()
   Call CONECRUN("FLFORMS/" + TRIM$(Text1.TEXT), "Formularios de Impresion")
End Sub

Private Sub Command7_Click()
Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text7.TEXT = VDEVU$
    End If
End Sub

Private Sub Command8_Click()
     Call CONECRUN("FLFORMS/" + TRIM$(Text10.TEXT), "Formularios de Impresion")
End Sub
Private Sub Command9_Click()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text10.TEXT = VDEVU$
    End If
End Sub

Private Sub Form_Load()
    '
    '
    Text10.TEXT = Control$("", "STIKAPRO")
    Text7.TEXT = Control$("", "STIKAC15")
    Text9.TEXT = ECOARCH$("INDIFRM", Text10.TEXT)
    Text8.TEXT = ECOARCH$("INDIFRM", Text7.TEXT)
    
    Text1.TEXT = Control$("", "STIKAPR2")
    Text5.TEXT = Control("", "STIKAPR3")
    Check1.Value = 0
    If Control("NUMSERIE", "UTIARMEX") = "NO" Then Check1.Value = 1 ' NO UTILIZA ARMADOR EXTERNO
    '
    Text6.TEXT = Control("", "DEFPROTR")

    Check4.Value = 0
    IMPR$ = TRIM$(Control$(IDENTER$, "IMPAUTOM"))
    If IMPR$ = "SI" Then Check4.Value = 1
    
    Check2.Value = 0
    IMPR$ = Control("CONFTRAZ", "ASIGLOTE")
    If IMPR$ = "SI" Then Check2.Value = 1
    
    Check3.Value = 0
    IMPR$ = Control("NUMSERIE", "IMPETIEX")
    If IMPR$ = "SI" Then Check3.Value = 1
    '
End Sub

Private Sub Text1_Change()
   Text2.TEXT = ECOARCH$("INDIFRM", UCase$(Text1.TEXT))
End Sub

Private Sub Text1_DblClick()
  Call Command4_Click
End Sub

Private Sub Text10_Change()
    Text9.TEXT = ECOARCH$("INDIFRM", UCase$(Text10.TEXT))
End Sub

Private Sub Text10_DblClick()
  Call Command9_Click
End Sub

Private Sub Text5_Change()
   Text4.TEXT = ECOARCH$("INDIFRM", UCase$(Text5.TEXT))
End Sub

Private Sub Text6_Change()

    CONDI$ = "LISTA = 'TABTIMA' AND CAMPO1 = " & XVALO(Text6)
    Text3 = VALOBADA("DATASQL", "CAMPO2", CONDI$, "NOMESS")
End Sub

