VERSION 5.00
Begin VB.Form ETIDESPA 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Impresión de Etiqueta de Despacho"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   6165
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Etiqueta de Despacho"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   5055
      Begin VB.TextBox Text2 
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
         Height          =   300
         Left            =   4200
         MaxLength       =   3
         TabIndex        =   27
         Text            =   " 1"
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox LENTREGA 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   3
         Left            =   1200
         MaxLength       =   32
         TabIndex        =   24
         Top             =   2640
         Width           =   3700
      End
      Begin VB.CommandButton Command24 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1800
         TabIndex        =   22
         Top             =   3000
         Width           =   195
      End
      Begin VB.TextBox Text19 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1200
         TabIndex        =   20
         Top             =   3000
         Width           =   645
      End
      Begin VB.TextBox LENTREGA 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   2040
         MaxLength       =   48
         TabIndex        =   19
         Top             =   3000
         Width           =   2860
      End
      Begin VB.TextBox LENTREGA 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   1200
         MaxLength       =   32
         TabIndex        =   11
         Top             =   1920
         Width           =   3700
      End
      Begin VB.TextBox LENTREGA 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   1200
         MaxLength       =   32
         TabIndex        =   10
         Top             =   2280
         Width           =   3700
      End
      Begin VB.TextBox Text20 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   9
         Top             =   1320
         Width           =   885
      End
      Begin VB.CommandButton Command25 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1080
         TabIndex        =   8
         Top             =   1320
         Width           =   195
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         MaxLength       =   5
         TabIndex        =   7
         Text            =   " "
         Top             =   660
         Width           =   855
      End
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
         Height          =   300
         Left            =   1080
         TabIndex        =   6
         Top             =   660
         Width           =   180
      End
      Begin VB.Label Ltota 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cant. de copias:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   6
         Left            =   2280
         TabIndex        =   28
         Top             =   300
         Width           =   1785
      End
      Begin VB.Label Ltota 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Teléfono:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   4
         Left            =   -10
         TabIndex        =   25
         Top             =   2700
         Width           =   1065
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1320
         TabIndex        =   23
         Top             =   660
         Width           =   3585
      End
      Begin VB.Label Ltota 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Transporte:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   5
         Left            =   -840
         TabIndex        =   21
         Top             =   3050
         Width           =   1905
      End
      Begin VB.Label Ltota 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Localidad:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   3
         Left            =   -840
         TabIndex        =   16
         Top             =   2370
         Width           =   1905
      End
      Begin VB.Label Ltota 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   80
         TabIndex        =   15
         Top             =   420
         Width           =   825
      End
      Begin VB.Label Ltota 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Sucursal:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   105
         TabIndex        =   14
         Top             =   1080
         Width           =   945
      End
      Begin VB.Label Ltota 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Domicilio:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   -840
         TabIndex        =   13
         Top             =   2000
         Width           =   1905
      End
      Begin VB.Label Label27 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1320
         TabIndex        =   12
         Top             =   1320
         Width           =   3585
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00004000&
      Height          =   5775
      Left            =   5280
      ScaleHeight     =   5715
      ScaleWidth      =   1455
      TabIndex        =   1
      Top             =   0
      Width           =   1515
      Begin VB.CommandButton Command3 
         Caption         =   "Exit"
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
         Left            =   120
         Picture         =   "ETIDESPA.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Cerrar y Salir"
         Top             =   120
         Width           =   650
      End
      Begin VB.CommandButton Command5 
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
         Left            =   120
         Picture         =   "ETIDESPA.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Configuración de Funcionamiento"
         Top             =   840
         Width           =   650
      End
      Begin VB.CommandButton Command17 
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
         Left            =   120
         Picture         =   "ETIDESPA.frx":058C
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Imprime Etiqueta"
         Top             =   2040
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   2
         Top             =   2760
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   3
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -300
         Picture         =   "ETIDESPA.frx":0BF6
         Top             =   3000
         Width           =   1515
      End
   End
   Begin VB.ListBox List1 
      Height          =   645
      Left            =   10200
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   5400
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   330
      Left            =   7440
      TabIndex        =   4
      Top             =   1140
      Width           =   1695
   End
End
Attribute VB_Name = "ETIDESPA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command3_Click()
   Unload Me
End Sub

Private Sub Command17_Click()
      '\\\OT-07-0382-FL-12/10/07///
    If Label1 = "" Then
        Call MENSERR(24, "Debe Ingresar un Cliente")
        Exit Sub
    End If

    COPIAS% = Val(Text2.TEXT)
    If COPIAS% < 1 Then
        Call MENSERR(24, "Debe Ingresar un Valor Numérico \Indicando la Cantidad de Copias")
        Exit Sub
    End If
     '\\\OT-07-0382-FIN///
    FORIPRE$ = TRIM$(CONTROL$("", "ETIDESP"))
    If FORIPRE$ = "" Then FORIPRE$ = "ETIQ-EXP"
    PORTUREC$ = TRIM$(CONTROL$(IDENTER$, "PUERDES"))
    If PORTUREC$ = "" Then PORTUREC$ = CONTROL$("*", "PUERDES")
    On Error Resume Next
    Printer.TrackDefault = False
    On Error GoTo 0
    If PORTUREC$ <> "" Then
      NEWPORT$ = PORTUREC$ + "/NOMESS"
      Call SETPRINPORT(NEWPORT$)
      If NEWPORT$ <> "" Then
        Call SETPRINPORT("RESTORE")
      End If
      PRINTERPORT$ = PORTUREC$
    End If
    '
    If EXISTE%(LUDAT$ + FORIPRE$ + ".FRM") Then
      HOII% = HOY(HOS$)
      FECHDOC$ = HOS$
      'NUMEDOC$ = TRIM$(Mid$(OBX$, 109, 16))
      TIPODOC$ = "Etiqueta de Expedición"
      DESTIDOC% = 0
      '
      '\\\OT-07-0382-FL-12-10-07///
      CACOLTAB1% = 0
      CODPARAM$(1) = "TRANSPOR": DOCPARAM$(1) = Label1
      CODPARAM$(2) = "REF-MAT2": DOCPARAM$(2) = Label27
      CODPARAM$(3) = "DOMI-TRA": DOCPARAM$(3) = LENTREGA(0).TEXT
      CODPARAM$(4) = "LOCA-TRA": DOCPARAM$(4) = LENTREGA(1).TEXT
      CODPARAM$(5) = "REF-MAT1": DOCPARAM$(5) = LENTREGA(3).TEXT
      
      CAPARDOC% = 5
      '\\\OT-07-0382-FIN///
      '
      CANCELPRINT% = 0
      NOCONFIRM% = 1
      '\\\OT-07-0382-FL-12/10/07///
      For kk% = 1 To COPIAS%
        Call PRINTDOC("@" + FORIPRE$) ' Etiqueta de Expedición
      Next kk%
      '\\\OT-07-0382-FIN///
      '
    End If
    Call SETPRINPORT("RESTORE")
    PRINTERPORT$ = ""
End Sub

Private Sub Command2_Click()
   Call SELECHO("DEUDOR1")
   NCLIE% = XVALO(TRIM$(VALACT1$("DEUDOR1")))
   Text1.TEXT = TRIM$(Str$(NCLIE%))
   If NCLIE% > 0 Then
     Label1 = TRIM$(RASOCLI$(NCLIE%))
     Call FILTERFILE("SUCENTRE", "", 5, MKI$(NCLIE%))
     For KKI% = 0 To 2
         LENTREGA(KKI%) = LUENTRECLI$(NCLIE%, KKI%)
     Next KKI%
   End If
End Sub

Private Sub Command24_Click()
   Call SELECHO("TRANSPOR")
   NTRAN% = XVALO(VALACT1$("TRANSPOR"))
   Text19 = TRIM$(FORMATNUM(NTRAN%, "I4"))
End Sub

Private Sub Command25_Click()
    Call SELECHO("!SUCENTRE")
    VDEVU$ = VALACT1$("!SUCENTRE")
    Text20 = VDEVU$
End Sub

Private Sub Command5_Click()
    CFGETIDE.Show 1
End Sub

Private Sub Form_Load()
Call BLANFORMU
End Sub

Private Sub LENTREGA_GotFocus(Index As Integer)
    If Index = 2 Then
         If XVALO(Text19) > 0 Then
           Call MENSERR(24, "Campo no Editable")
           On Error Resume Next
           Text19.SetFocus
           On Error GoTo 0
           Exit Sub ' LE AGREGUE ESTA LINEA PARA QUE NO SE EJECUTE EL RESTO DEL CODIGO
         End If
       End If
       LENTREGA(Index).TEXT = TRIM$(LENTREGA(Index).TEXT)
       LENTREGA(Index).SelStart = 0
       LENTREGA(Index).SelLength = Len(LENTREGA(Index).TEXT)
End Sub

Private Sub Text1_Change()
   Call BLANFORMU
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
      Call BLANFORMU
      If XVALO(Text1.TEXT) > 0 Then
        If XVALO(Text1.TEXT) <= 32767 Then
          NCLIE% = XVALO(Text1)
          Label1 = TRIM$(RASOCLI$(NCLIE%))
          Call FILTERFILE("SUCENTRE", "", 5, MKI$(NCLIE%))
          For KKI% = 0 To 2
              LENTREGA(KKI%) = LUENTRECLI$(NCLIE%, KKI%)
          Next KKI%
         End If
       End If
 End If
End Sub

Private Sub Text19_Change()
   LENTREGA(2) = ""
   NTRA = XVALO(Text19)
   If NTRA > 0 Then
      If NTRA < 32767 Then
         NTRAI% = NTRA
         LENTREGA(2) = ECOARCH$("TRANSPOR", MKI$(NTRAI%))
         If LENTREGA(2) <> "" Then
            TRANACT% = NTRAI%
         End If
      End If
   End If
End Sub

Private Sub Text20_Change()
    NSUC& = XVALO(Text20)
    Label27 = ""
    LENTREGA(0) = "": LENTREGA(1) = "": LENTREGA(3) = ""
    If NSUC& > 0 Then
      If NSUC& <= ULTREG&("!SUCENTRE") Then
        XX$ = REGLEIDO$("!SUCENTRE", NSUC&)
        Label27 = TRIM$(Mid$(XX$, 3, 62))
        LENTREGA(0) = TRIM$(Mid$(XX$, 65, 48))
        LENTREGA(1) = TRIM$(Mid$(XX$, 113, 48))
        LENTREGA(3) = TRIM$(Mid$(XX$, 163, 32))
      End If
    End If
End Sub

Private Sub BLANFORMU()
    Label1 = ""
    Text20 = ""
    Call SETULTREG("!SUCENTRE", 0)
    Call FRESHECHO("!SUCENTRE")
    For KKI% = 0 To 3
       LENTREGA(KKI%) = ""
    Next KKI%
    Text19 = ""
End Sub
