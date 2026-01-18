VERSION 5.00
Begin VB.Form ICALREC 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0F0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sistema de Calificación de Proveedores"
   ClientHeight    =   8130
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11325
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
   ScaleHeight     =   8130
   ScaleWidth      =   11325
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Print"
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
      Left            =   10395
      Picture         =   "ICALREC.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   41
      ToolTipText     =   "Imprime Ficha"
      Top             =   4410
      Width           =   750
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   10395
      ScaleHeight     =   75
      ScaleWidth      =   705
      TabIndex        =   39
      Top             =   5355
      Width           =   765
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   40
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Setup"
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
      Left            =   10395
      Picture         =   "ICALREC.frx":066A
      Style           =   1  'Graphical
      TabIndex        =   38
      ToolTipText     =   "Configuración de Funcionamiento"
      Top             =   2835
      Width           =   750
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0F0FF&
      Caption         =   "Anotaciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1755
      Left            =   105
      TabIndex        =   26
      Top             =   6250
      Width           =   11040
      Begin VB.Frame Frame7 
         Caption         =   "Frame7"
         Height          =   1310
         Left            =   10605
         TabIndex        =   27
         Top             =   295
         Width           =   15
      End
      Begin VB.Frame Frame6 
         Caption         =   "Frame4"
         Height          =   1310
         Left            =   8295
         TabIndex        =   28
         Top             =   295
         Width           =   15
      End
      Begin VB.Frame Frame8 
         Caption         =   "Frame3"
         Height          =   1310
         Left            =   1365
         TabIndex        =   29
         Top             =   295
         Width           =   15
      End
      Begin VB.ListBox List3 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   210
         TabIndex        =   30
         Top             =   630
         Width           =   10680
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00E0E0E0&
         Height          =   435
         Left            =   210
         ScaleHeight     =   375
         ScaleWidth      =   10620
         TabIndex        =   31
         Top             =   315
         Width           =   10680
         Begin VB.Label Label15 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha     Detalle                                                   Usuario"
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   105
            TabIndex        =   32
            Top             =   35
            Width           =   10200
         End
      End
      Begin VB.Line Line7 
         X1              =   210
         X2              =   10880
         Y1              =   1590
         Y2              =   1590
      End
      Begin VB.Line Line6 
         X1              =   10890
         X2              =   10890
         Y1              =   320
         Y2              =   1600
      End
      Begin VB.Line Line5 
         X1              =   0
         X2              =   10605
         Y1              =   4410
         Y2              =   4410
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0F0FF&
      Caption         =   "Material Recibido"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2080
      Left            =   105
      TabIndex        =   19
      Top             =   1470
      Width           =   10095
      Begin VB.Frame Frame19 
         Caption         =   "Frame19"
         Height          =   1500
         Left            =   9670
         TabIndex        =   47
         Top             =   420
         Width           =   15
      End
      Begin VB.Frame Frame18 
         Caption         =   "Frame18"
         Height          =   1500
         Left            =   8715
         TabIndex        =   46
         Top             =   420
         Width           =   15
      End
      Begin VB.Frame Frame17 
         Caption         =   "Frame17"
         Height          =   1500
         Left            =   7750
         TabIndex        =   45
         Top             =   420
         Width           =   15
      End
      Begin VB.Frame Frame16 
         Caption         =   "Frame16"
         Height          =   1500
         Left            =   6795
         TabIndex        =   44
         Top             =   420
         Width           =   15
      End
      Begin VB.Frame Frame15 
         Caption         =   "Frame15"
         Height          =   1500
         Left            =   5880
         TabIndex        =   43
         Top             =   420
         Width           =   15
      End
      Begin VB.Frame Frame14 
         Caption         =   "Frame14"
         Height          =   1500
         Left            =   5250
         TabIndex        =   42
         Top             =   420
         Width           =   15
      End
      Begin VB.Frame Frame9 
         Caption         =   "Frame9"
         Height          =   1500
         Left            =   2100
         TabIndex        =   33
         Top             =   420
         Width           =   15
      End
      Begin VB.ListBox List2 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1185
         Left            =   210
         TabIndex        =   20
         Top             =   735
         Width           =   9660
      End
      Begin VB.PictureBox Picture2 
         Height          =   330
         Left            =   210
         ScaleHeight     =   270
         ScaleWidth      =   9600
         TabIndex        =   21
         Top             =   420
         Width           =   9660
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Codigo          Descripción               U-M     Cant.   Apro.   Rech.   Falt."
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   50
            TabIndex        =   22
            Top             =   35
            Width           =   11000
         End
      End
      Begin VB.Line Line4 
         X1              =   210
         X2              =   9870
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Line Line3 
         X1              =   9870
         X2              =   9870
         Y1              =   420
         Y2              =   1920
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0F0FF&
      Caption         =   "Identificación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   105
      TabIndex        =   10
      Top             =   210
      Width           =   7680
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
         Left            =   2100
         TabIndex        =   48
         Top             =   380
         Width           =   175
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
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
         Left            =   0
         TabIndex        =   25
         Top             =   0
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label11 
         BackColor       =   &H00E0E0E0&
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
         Left            =   1260
         TabIndex        =   18
         Top             =   735
         Width           =   6210
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
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
         Height          =   330
         Left            =   210
         TabIndex        =   17
         Top             =   780
         Width           =   960
      End
      Begin VB.Label Label9 
         BackColor       =   &H00E0E0E0&
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
         Left            =   6405
         TabIndex        =   16
         Top             =   375
         Width           =   1065
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
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
         Height          =   330
         Left            =   5355
         TabIndex        =   15
         Top             =   420
         Width           =   960
      End
      Begin VB.Label Label7 
         BackColor       =   &H00E0E0E0&
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
         Left            =   3570
         TabIndex        =   14
         Top             =   375
         Width           =   1800
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nro.Remito:"
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
         Left            =   2205
         TabIndex        =   13
         Top             =   420
         Width           =   1275
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
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
         Left            =   1260
         TabIndex        =   12
         Top             =   380
         Width           =   855
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nro.B.R.:"
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
         Left            =   210
         TabIndex        =   11
         Top             =   420
         Width           =   960
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save"
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
      Left            =   10395
      Picture         =   "ICALREC.frx":0AAC
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Aprueba y Graba Ficha"
      Top             =   1155
      Width           =   750
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0F0FF&
      Caption         =   "I.Calidad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7875
      TabIndex        =   3
      Top             =   210
      Width           =   2325
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Puntaje :"
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
         Left            =   0
         TabIndex        =   8
         Top             =   630
         Width           =   960
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   435
         Left            =   1050
         TabIndex        =   7
         Top             =   525
         Width           =   1065
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0F0FF&
      Caption         =   "Calificación por Indices"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   105
      TabIndex        =   2
      Top             =   3625
      Width           =   10095
      Begin VB.Frame Frame13 
         Caption         =   "Frame13"
         Height          =   1970
         Left            =   6195
         TabIndex        =   37
         Top             =   420
         Width           =   15
      End
      Begin VB.Frame Frame12 
         Caption         =   "Frame12"
         Height          =   1970
         Left            =   5300
         TabIndex        =   36
         Top             =   420
         Width           =   15
      End
      Begin VB.Frame Frame11 
         Caption         =   "Frame11"
         Height          =   1970
         Left            =   4410
         TabIndex        =   35
         Top             =   420
         Width           =   15
      End
      Begin VB.Frame Frame10 
         Caption         =   "Frame10"
         Height          =   1970
         Left            =   630
         TabIndex        =   34
         Top             =   420
         Width           =   15
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
         Height          =   1635
         Left            =   210
         TabIndex        =   6
         Top             =   735
         Width           =   9660
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00E0E0E0&
         Height          =   330
         Left            =   210
         ScaleHeight     =   270
         ScaleWidth      =   9600
         TabIndex        =   4
         Top             =   420
         Width           =   9660
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "N. Indice de Calidad                Pond.  Valor  Observaciones"
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   105
            TabIndex        =   5
            Top             =   35
            Width           =   9405
         End
      End
      Begin VB.Line Line2 
         X1              =   9870
         X2              =   9870
         Y1              =   420
         Y2              =   2370
      End
      Begin VB.Line Line1 
         X1              =   210
         X2              =   9870
         Y1              =   2370
         Y2              =   2370
      End
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Help"
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
      Left            =   10395
      Picture         =   "ICALREC.frx":0DB6
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Ayuda FLEXOFT en Línea"
      Top             =   1995
      Width           =   750
   End
   Begin VB.CommandButton command1 
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
      Height          =   750
      Left            =   10395
      Picture         =   "ICALREC.frx":10C0
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Abandona la Aplicación"
      Top             =   315
      Width           =   750
   End
   Begin VB.Label RECALI 
      Caption         =   "0"
      Height          =   330
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label Label13 
      Caption         =   "5"
      Height          =   330
      Left            =   10395
      TabIndex        =   23
      Top             =   0
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   9795
      Picture         =   "ICALREC.frx":120A
      Top             =   5640
      Width           =   2010
   End
End
Attribute VB_Name = "ICALREC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim OBSEX$(48)
'
Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub
'

Private Sub Command2_Click()
   Command2.Enabled = False
   RFF$ = RECFILT$("OBCALREC", 1, MKS$(Val(Label5)))
   For U& = 1 To Len(RFF$)
     REOB& = CVS(Mid$(RFF$, U&, 4))
     Call GRAREG("OBCALREC", String$(32, 0), REOB&)
   Next U&
   '
   RECA& = Val(RECALI)
   If RECA& < 1 Then RECA& = 1 + ULTREG&("CALIREC")
   '
   HOII% = HOY(HOS$)
   XX$ = REGBLAN$("CALIREC")
   Call REPLA(XX$, MKS$(Val(Label5)), 1, 4)
     FEXI% = FECHANUM(Label9)
   Call REPLA(XX$, MKI$(FEXI%), 5, 2)
     NPRO% = Val(Label14)
   Call REPLA(XX$, MKI$(NPRO%), 7, 2)
   Call REPLA(XX$, TRIM$(Label7), 9, 16)
   Call REPLA(XX$, MKI$(HOII%), 25, 2)
   Call REPLA(XX$, USERNAME$, 27, 22)
   Call REPLA(XX$, MKS$(Val(Label2)), 49, 4)
   '
   YY$ = String$(192, 0)
   ESCALI = Val(Label13)
   REBLAO$ = REGBLAN$("OBCALREC")
   For U& = 1 To List1.ListCount
     TTXX$ = List1.List(U& - 1)
     VALOUN = Val(Mid$(TTXX$, 43, 6)) / ESCALI
     Call REPLA(YY$, MKS$(VALOUN), 4 * U& - 3, 4)
     OBXX$ = TRIM$(Mid$(TTXX$, 51, 24))
     If OBXX$ <> "" Then
       ZZ$ = REBLAO$
       UKI% = U&
       Call REPLA(ZZ$, MKS$(Val(Label5)), 1, 4)
       Call REPLA(ZZ$, MKI$(UKI%), 5, 2)
       Call REPLA(ZZ$, OBXX$, 7, 24)
       Call REGAPP("OBCALREC", ZZ$)
     End If
   Next U&
   Call REPLA(XX$, YY$, 65, 192)
   Call GRAREG("CALIREC", XX$, RECA&)
   Call CIERRARCH("CALIREC")
   Call CIERRARCH("OBCALREC")
   '
   Command2.Enabled = True
   Hide
End Sub
'

Private Sub Command8_Click()
    Call SELECHO("DEUDOR1")
    NCMAX% = Val(VALACT1$("DEUDOR1"))
    If NCMAX% > 0 Then
      Text8.TEXT = TRIM$(Str$(NCMAX%))
    End If
End Sub

Private Sub Command9_Click()
   Command9.Enabled = False
   '
10 Call SELECHO("INDIBOL")
   NBR$ = VALACT1$("INDIBOL")   ' OBJPLA$("SELBOLET", "")
   If NBR$ <> "" Then
     NROBOL& = Val(NBR$)
     If NROBOL& > 0 Then
       Label5 = TRIM$(Str$(NROBOL&))
     End If
   End If
   '
   Command9.Enabled = True
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   If DERACCE%("ACTICAL", "Actualización Indices de Calificación") >= 2 Then
     STPCALPR.Show 1
   End If
   Command3.Enabled = True
End Sub

Private Sub Command4_Click()
   Command4.Enabled = False
   PRINTFORM
   Command4.Enabled = True
End Sub

Private Sub Form_Load()
   '
   If ULTREG&("ICALPROV") < 1 Then
     REBLA$ = REGBLAN$("ICALPROV")
     ZZ$ = REBLA$: Call REPLA(ZZ$, Chr$(1) + "Calidad Propiamente Dicha", 1, 32): Call REGAPP("ICALPROV", ZZ$)
     ZZ$ = REBLA$: Call REPLA(ZZ$, Chr$(2) + "Embalajes", 1, 32): Call REGAPP("ICALPROV", ZZ$)
     ZZ$ = REBLA$: Call REPLA(ZZ$, Chr$(3) + "Identificación", 1, 32): Call REGAPP("ICALPROV", ZZ$)
     ZZ$ = REBLA$: Call REPLA(ZZ$, Chr$(4) + "Entrega de Certificados", 1, 32): Call REGAPP("ICALPROV", ZZ$)
     Call CIERRARCH("ICALPROV")
   End If
   '
   If Val(CONTROL$("CALIPROV", "ESCACALI")) = 10 Then
      Label13 = "10"
   End If
   '
   List1.Clear: PUNMAX = 0
   FIN& = ULTREG&("ICALPROV")
   For U& = 1 To FIN&
     XX$ = REGLEIDO$("ICALPROV", U&)
     TTXX$ = Space$(128)
     Call REPLA(TTXX$, FORMATNUM$(U&, "I2"), 1, 2)
     Call REPLA(TTXX$, Mid$(XX$, 2, 31), 5, 30)
     Call REPLA(TTXX$, FORMATNUM$(CVS(Mid$(XX$, 33, 4)), "F6.2"), 36, 6)
     PUNMAX = PUNMAX + CVS(Mid$(XX$, 33, 4))
     List1.AddItem TTXX$
   Next U&
   '
   Label2.Caption = TRIM$(FORMATNUM$(PUNMAX, "F6.2"))
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
End Sub

Private Sub Text10_Change()
   Label7 = ""
   NCMIN% = Val(Text10.TEXT)
   If Val(NCMIN%) > 0 Then
     Label7 = RASOCLI$(NCMIN%)
   End If
End Sub

Private Sub Text10_DblClick()
    Call SELECHO("DEUDOR1")
    NCMIN% = Val(VALACT1$("DEUDOR1"))
    If NCMIN% > 0 Then
      Text10.TEXT = TRIM$(Str$(NCMIN%))
    End If
End Sub

Private Sub Text8_Change()
   Label8 = ""
   NCMAX% = Val(Text8.TEXT)
   If Val(NCMAX%) > 0 Then
     Label8 = RASOCLI$(NCMAX%)
   End If
End Sub

Private Sub Text8_DblClick()
    Call SELECHO("DEUDOR1")
    NCMAX% = Val(VALACT1$("DEUDOR1"))
    If NCMAX% > 0 Then
      Text8.TEXT = TRIM$(Str$(NCMAX%))
    End If
End Sub

Private Sub Label15_DblClick()
   Call ANOTABOLRE
End Sub

Private Sub Label5_Change()
   NBR& = Val(Label5)
   List2.Clear
   If NBR& > 0 Then
     RFF$ = RECFILT$("BOLRECEP", 1, MKS$(NBR&))
     For U& = 1 To Len(RFF$) Step 4
       REBO& = CVS(Mid$(RFF$, U&, 4))
       XX$ = REGLEIDO$("BOLRECEP", REBO&)
       '
       If List2.ListCount < 1 Then
         NPRO% = CVI(Mid$(XX$, 105, 2))
         Label14 = TRIM$(Str$(NPRO%))
         Label11 = RASOCLI$((-1) * NPRO%)
         Label7 = TRIM$(Mid$(XX$, 119, 16)) ' NUMERO DE REMITO
         FEX = CVI(Mid$(XX$, 81, 2))
         Label9 = FECHATEX$(FEX)
       End If
       '
       TTXX$ = Space$(128)
         CIXI% = CVI(Mid$(XX$, 83, 2))
       Call REPLA(TTXX$, CODEXT$(CIXI%), 1, 15)
       Call REPLA(TTXX$, DESCRIT0$(CIXI%), 17, 25)
         UME$ = TRIM$(Mid$(TTXX$, 85, 4))
         If UME$ = "" Then UME$ = UNIMED$(CIXI%)
       Call REPLA(TTXX$, UME$, 43, 4)
         CANRE = CVD(Mid$(XX$, 89, 8))
         CANAP = CVD(Mid$(XX$, 163, 8))
         CANCH = CVD(Mid$(XX$, 171, 8))
         CANFA = CVD(Mid$(XX$, 179, 8))
         '
         If (Abs(CANAP + CANCH + CANFA - CANRE)) >= 0.05 Then
           Call RECUBORE(REBO&)
           XX$ = REGLEIDO$("BOLRECEP", REBO&)
           CANRE = CVD(Mid$(XX$, 89, 8))
           CANAP = CVD(Mid$(XX$, 163, 8))
           CANCH = CVD(Mid$(XX$, 171, 8))
           CANFA = CVD(Mid$(XX$, 179, 8))
         End If
         '
       Call REPLA(TTXX$, FORMATNUM$(CANRE, "F8.1"), 47, 8)
       Call REPLA(TTXX$, FORMATNUM$(CANAP, "F8.1"), 55, 8)
       Call REPLA(TTXX$, FORMATNUM$(CANCH, "F8.1"), 63, 8)
       Call REPLA(TTXX$, FORMATNUM$(CANFA, "F8.1"), 71, 8)
       List2.AddItem TTXX$
       '
     Next U&
     '
     REPUN$ = String$(192, 0)
     RECALI = ""
     RFF$ = RECFILT$("CALIREC", 1, MKS$(NBR&))
     If Len(RFF$) >= 4 Then
       U1& = CVS(Left$(RFF$, 4))
       XX$ = REGLEIDO$("CALIREC", U1&)
       REPUN$ = Mid$(XX$, 65, 192)
       RECALI = TRIM$(Str$(U1&))
     End If
     '
     ESCALI = Val(Label13)
     '
     For KU% = 1 To 48
       OBSEX$(KU%) = ""
     Next KU%
     RFF$ = RECFILT$("OBCALREC", 1, MKS$(NBR&))
     For U& = 1 To Len(RFF$)
       REOB& = CVS(Mid$(RFF$, U&, 4))
       XX$ = REGLEIDO$("OBCALREC", REOB&)
       IOB% = CVI(Mid$(XX$, 5, 2))
       If IOB% > 0 Then
         If IOB% <= 48 Then
           OBSEX$(IOB%) = Mid$(XX$, 7, 24)
         End If
       End If
     Next U&
     '
     List1.Clear: PUNMAX = 0
     FIN& = ULTREG&("ICALPROV")
     For U& = 1 To FIN&
       XX$ = REGLEIDO$("ICALPROV", U&)
       TTXX$ = Space$(128)
       Call REPLA(TTXX$, FORMATNUM$(U&, "I2"), 1, 2)
       Call REPLA(TTXX$, Mid$(XX$, 2, 31), 5, 30)
       PONDERA = CVS(Mid$(XX$, 33, 4))
       Call REPLA(TTXX$, FORMATNUM$(PONDERA, "F6.2"), 36, 6)
       Call REPLA(TTXX$, FORMATNUM$(ESCALI * CVS(Mid$(REPUN$, 4 * U& - 3, 4)), "F6.1"), 43, 6)
       PUNMAX = PUNMAX + PONDERA * CVS(Mid$(REPUN$, 4 * U& - 3, 4))
       Call REPLA(TTXX$, OBSEX$(U&), 52, 24)
       List1.AddItem TTXX$
     Next U&
     '
     Label2 = FORMATNUM(PUNMAX, "F6.2")
     '
     List3.Clear
     NBR& = Val(Label5)
     RFF$ = RECFILT$("ANOTABOL", 1, MKS$(NBR&))
     For U& = 1 To Len(RFF$) Step 4
       RENOTA& = CVS(Mid$(RFF$, U&, 4))
       XX$ = REGLEIDO$("ANOTABOL", RENOTA&)
       List3.AddItem Mid$(XX$, 5, 92)
     Next U&
     '
   End If
End Sub

Private Sub List1_DblClick()
   '
5  REBLA$ = REGBLAN$("ICALPROV")
   JJ& = 0
   For U& = 1 To List1.ListCount
     TTXX$ = List1.List(U& - 1)
     XX$ = REBLA$
     Call REPLA(XX$, Chr$(U&), 1, 1)
     Call REPLA(XX$, Mid$(TTXX$, 5, 30), 2, 31)
     Call REPLA(XX$, MKS$(Val(Mid$(TTXX$, 36, 6))), 33, 4)
     Call REPLA(XX$, MKS$(Val(Mid$(TTXX$, 43, 6))), 37, 4)
     Call REPLA(XX$, Mid$(TTXX$, 52, 24), 41, 24)
     JJ& = JJ& + 1
     Call GRAREG("!ICALPROV", XX$, JJ&)
   Next U&
   Call CIERRARCH("!ICALPROV")
   '
10 ATRI$ = "/NC"
   ATRI$ = ATRI$ + "/TITUPAN=Calidad de Recepción"
   ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(50 + Top + Frame4.Top))
   ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(50 + Left + Frame4.Left))
   '
   ATX% = 0
   VTB% = VENTABU%("CARICALP" + ATRI$)
   ESCALI = Val(Label13)
   If VTB% >= 0 Then
     List1.Clear
     JJ& = 0: PUNMAX = 0
     For U& = 1 To ULTREG&("!ICALPROV")
       XX$ = REGLEIDO$("!ICALPROV", U&)
       TTXX$ = Space$(128)
       Call REPLA(TTXX$, FORMATNUM$(U&, "I2"), 1, 2)
       Call REPLA(TTXX$, Mid$(XX$, 2, 31), 5, 30)
       Call REPLA(TTXX$, FORMATNUM$(CVS(Mid$(XX$, 33, 4)), "F6.2"), 36, 6)
       PUNTOS = CVS(Mid$(XX$, 37, 4))
       If PUNTOS > ESCALI Then
         PUNTOS = ESCALI
         ATX% = 1
       End If
       Call REPLA(TTXX$, FORMATNUM$(PUNTOS, "F6.1"), 43, 6)
       Call REPLA(TTXX$, Mid$(XX$, 41, 24), 52, 24)
       List1.AddItem TTXX$
       PUNMAX = PUNMAX + CVS(Mid$(XX$, 33, 4)) * CVS(Mid$(XX$, 37, 4)) / ESCALI
     Next U&
     '
     If ATX% > 0 Then
       Call MENSERR(24, "Hay Valores Fuera de Rango.")
       GoTo 5
     End If
     
     Label2.Caption = TRIM$(FORMATNUM$(PUNMAX, "F6.2"))
     Label2.Refresh
   End If
   '
End Sub

Private Sub List3_DblClick()
   Call ANOTABOLRE
End Sub

Sub ANOTABOLRE()
   ANOTABOL.Label16 = Label5
   ANOTABOL.Show 1
   '
   List3.Clear
   NBR& = Val(Label5)
   RFF$ = RECFILT$("ANOTABOL", 1, MKS$(NBR&))
   For U& = 1 To Len(RFF$) Step 4
     RENOTA& = CVS(Mid$(RFF$, U&, 4))
     XX$ = REGLEIDO$("ANOTABOL", RENOTA&)
     List3.AddItem Mid$(XX$, 5, 92)
   Next U&
   '
End Sub

Sub RECUBORE(REGXX&)
  XX$ = REGLEIDO$("BOLRECEP", REGXX&)
  CIXI% = CVI(Mid$(XX$, 83, 2))
  LOTE$ = TRIM$(Mid$(XX$, 69, 12))
  '
  If LOTE$ <> "" Then
    RFF$ = RECFILT$("MOVISTO", 2, MKI$(CIXI%))
    CANRE = 0: CANCH = 0: CANFA = 0
    For UL& = 1 To Len(RFF$) Step 4
      REMO& = CVS(Mid$(RFF$, UL&, 4))
      YY$ = REGLEIDO$("MOVISTO", REMO&)
      If TRIM$(Mid$(YY$, 5, 12)) = LOTE$ Then
        CANTI = CVD(Mid$(YY$, 88, 8)) - CVD(Mid$(YY$, 96, 8))
        REFE$ = UCase$(Mid$(YY$, 43, 30))
        If Left$(REFE$, 7) = "RECHAZO" Then
            CANCH = (-1) * CANTI
          ElseIf Left$(REFE$, 7) = "FALTANT" Then
            CANFA = (-1) * CANTI
          Else
            CANRE = CANTI
        End If
      End If
    Next UL&
    '
    If CANRE >= 0.05 Then
      CANAP = CANRE - CANCH - CANFA
      Call REPLA(XX$, MKD$(CDbl(CANAP)), 163, 8)
      Call REPLA(XX$, MKD$(CDbl(CANCH)), 171, 8)
      Call REPLA(XX$, MKD$(CDbl(CANFA)), 179, 8)
      Call GRAREG("BOLRECEP", XX$, REGXX&)
      Call CIERRARCH("BOLRECEP")
    End If
    '
  End If
  '
End Sub

