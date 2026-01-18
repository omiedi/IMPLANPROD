VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form ICALREC07 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00CDDADA&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sistema de Calificación de Proveedores"
   ClientHeight    =   8130
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11175
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
   ScaleWidth      =   11175
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   330
      Index           =   0
      Left            =   1800
      ScaleHeight     =   270
      ScaleWidth      =   1110
      TabIndex        =   48
      Top             =   240
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00CDDADA&
      Caption         =   "ANOTACIONES"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1755
      Left            =   120
      TabIndex        =   23
      Top             =   6250
      Width           =   11040
      Begin VB.Frame Frame7 
         Caption         =   "Frame7"
         Height          =   1310
         Left            =   10605
         TabIndex        =   24
         Top             =   295
         Width           =   15
      End
      Begin VB.Frame Frame6 
         Caption         =   "Frame4"
         Height          =   1310
         Left            =   8295
         TabIndex        =   25
         Top             =   295
         Width           =   15
      End
      Begin VB.Frame Frame8 
         Caption         =   "Frame3"
         Height          =   1310
         Left            =   1365
         TabIndex        =   26
         Top             =   295
         Width           =   15
      End
      Begin VB.ListBox List3 
         Appearance      =   0  'Flat
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
         Height          =   930
         Left            =   210
         TabIndex        =   27
         Top             =   630
         Width           =   10680
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   210
         ScaleHeight     =   405
         ScaleWidth      =   10650
         TabIndex        =   28
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
            TabIndex        =   29
            ToolTipText     =   "Doble Click para Abrir Anotador"
            Top             =   35
            Width           =   10200
         End
      End
      Begin VB.Line Line5 
         X1              =   0
         X2              =   10605
         Y1              =   4410
         Y2              =   4410
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00CDDADA&
      Caption         =   "MATERIAL RECIBIDO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2080
      Left            =   105
      TabIndex        =   16
      Top             =   1470
      Width           =   10095
      Begin VB.Frame Frame19 
         Caption         =   "Frame19"
         Height          =   1500
         Left            =   9670
         TabIndex        =   40
         Top             =   420
         Width           =   15
      End
      Begin VB.Frame Frame18 
         Caption         =   "Frame18"
         Height          =   1500
         Left            =   8715
         TabIndex        =   39
         Top             =   420
         Width           =   15
      End
      Begin VB.Frame Frame17 
         Caption         =   "Frame17"
         Height          =   1500
         Left            =   7750
         TabIndex        =   38
         Top             =   420
         Width           =   15
      End
      Begin VB.Frame Frame16 
         Caption         =   "Frame16"
         Height          =   1500
         Left            =   6795
         TabIndex        =   37
         Top             =   420
         Width           =   15
      End
      Begin VB.Frame Frame15 
         Caption         =   "Frame15"
         Height          =   1500
         Left            =   5880
         TabIndex        =   36
         Top             =   420
         Width           =   15
      End
      Begin VB.Frame Frame14 
         Caption         =   "Frame14"
         Height          =   1500
         Left            =   5250
         TabIndex        =   35
         Top             =   420
         Width           =   15
      End
      Begin VB.Frame Frame9 
         Caption         =   "Frame9"
         Height          =   1500
         Left            =   2100
         TabIndex        =   30
         Top             =   420
         Width           =   15
      End
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
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
         Height          =   1155
         Left            =   210
         TabIndex        =   17
         Top             =   735
         Width           =   9660
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   210
         ScaleHeight     =   300
         ScaleWidth      =   9630
         TabIndex        =   18
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
            TabIndex        =   19
            Top             =   35
            Width           =   11000
         End
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00CDDADA&
      Caption         =   "IDENTIFICACIÓN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   105
      TabIndex        =   7
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
         TabIndex        =   41
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
         TabIndex        =   22
         Top             =   0
         Visible         =   0   'False
         Width           =   855
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
         Left            =   1260
         TabIndex        =   15
         Top             =   735
         Width           =   6210
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Proveedor"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   210
         TabIndex        =   14
         Top             =   810
         Width           =   960
      End
      Begin VB.Label Label9 
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
         Left            =   6405
         TabIndex        =   13
         Top             =   375
         Width           =   1065
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5595
         TabIndex        =   12
         Top             =   450
         Width           =   720
      End
      Begin VB.Label Label7 
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
         Left            =   3720
         TabIndex        =   11
         Top             =   375
         Width           =   1800
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nro.Remito"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2640
         TabIndex        =   10
         Top             =   450
         Width           =   1035
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
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
         Left            =   1260
         TabIndex        =   9
         Top             =   380
         Width           =   855
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nro.B.R."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   210
         TabIndex        =   8
         Top             =   450
         Width           =   960
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00CDDADA&
      Caption         =   "INDICE CALIDAD"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7875
      TabIndex        =   1
      Top             =   210
      Width           =   2325
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Puntaje"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   0
         TabIndex        =   6
         Top             =   750
         Width           =   960
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
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
         Height          =   435
         Left            =   1050
         TabIndex        =   5
         Top             =   525
         Width           =   1065
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00CDDADA&
      Caption         =   "CALIFICACIÓN POR INDICES"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   105
      TabIndex        =   0
      Top             =   3625
      Width           =   10095
      Begin VB.Frame Frame13 
         Caption         =   "Frame13"
         Height          =   1970
         Left            =   6195
         TabIndex        =   34
         Top             =   420
         Width           =   15
      End
      Begin VB.Frame Frame12 
         Caption         =   "Frame12"
         Height          =   1970
         Left            =   5300
         TabIndex        =   33
         Top             =   420
         Width           =   15
      End
      Begin VB.Frame Frame11 
         Caption         =   "Frame11"
         Height          =   1970
         Left            =   4410
         TabIndex        =   32
         Top             =   420
         Width           =   15
      End
      Begin VB.Frame Frame10 
         Caption         =   "Frame10"
         Height          =   1970
         Left            =   630
         TabIndex        =   31
         Top             =   420
         Width           =   15
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
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
         Height          =   1605
         Left            =   210
         TabIndex        =   4
         Top             =   735
         Width           =   9660
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   210
         ScaleHeight     =   300
         ScaleWidth      =   9630
         TabIndex        =   2
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
            TabIndex        =   3
            Top             =   35
            Width           =   9405
         End
      End
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   6180
      Left            =   10300
      ScaleHeight     =   6150
      ScaleWidth      =   1095
      TabIndex        =   42
      Top             =   0
      Width           =   1125
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H002B9973&
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   120
         ScaleHeight     =   165
         ScaleWidth      =   615
         TabIndex        =   50
         Top             =   5520
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   51
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton command1 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   105
         Picture         =   "ICALREC07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Abandona la Aplicación"
         Top             =   315
         Width           =   650
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Help"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   105
         Picture         =   "ICALREC07.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   1995
         Width           =   650
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   105
         Picture         =   "ICALREC07.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Aprueba y Graba Ficha"
         Top             =   1155
         Width           =   650
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Setup"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   105
         Picture         =   "ICALREC07.frx":075E
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Configuración de Funcionamiento"
         Top             =   2835
         Width           =   650
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   105
         Picture         =   "ICALREC07.frx":0BA0
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Imprime Ficha"
         Top             =   4410
         Width           =   650
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2625
      OleObjectBlob   =   "ICALREC07.frx":120A
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "ICALREC07.frx":143E
      TabIndex        =   49
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label RECALI 
      Caption         =   "0"
      Height          =   330
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label Label13 
      Caption         =   "5"
      Height          =   330
      Left            =   9840
      TabIndex        =   20
      Top             =   0
      Visible         =   0   'False
      Width           =   540
   End
End
Attribute VB_Name = "ICALREC07"
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
   RFF$ = RECFILT$("OBCALREC", 1, MKS$(Val(LABEL5)))
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
   Call REPLA(XX$, MKS$(Val(LABEL5)), 1, 4)
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
       zz$ = REBLAO$
       UKI% = U&
       Call REPLA(zz$, MKS$(Val(LABEL5)), 1, 4)
       Call REPLA(zz$, MKI$(UKI%), 5, 2)
       Call REPLA(zz$, OBXX$, 7, 24)
       Call REGAPP("OBCALREC", zz$)
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
10 Call SELECHO("INDIBOL")
   NBR$ = VALACT1$("INDIBOL")   ' OBJPLA$("SELBOLET", "")
   If NBR$ <> "" Then
     NROBOL& = Val(NBR$)
     If NROBOL& > 0 Then
       LABEL5 = TRIM$(Str$(NROBOL&))
     End If
   End If
   '
   Command9.Enabled = True
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   If DERACCE%("ACTICAL", "Actualización Indices de Calificación") >= 2 Then
     STPCALPR07.Show 1
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
     zz$ = REBLA$: Call REPLA(zz$, Chr$(1) + "Calidad Propiamente Dicha", 1, 32): Call REGAPP("ICALPROV", zz$)
     zz$ = REBLA$: Call REPLA(zz$, Chr$(2) + "Embalajes", 1, 32): Call REGAPP("ICALPROV", zz$)
     zz$ = REBLA$: Call REPLA(zz$, Chr$(3) + "Identificación", 1, 32): Call REGAPP("ICALPROV", zz$)
     zz$ = REBLA$: Call REPLA(zz$, Chr$(4) + "Entrega de Certificados", 1, 32): Call REGAPP("ICALPROV", zz$)
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
   '
   UTILIZA_SKIN% = 1
   Call APLICACIONSKINS(Me)

End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
End Sub

Private Sub Text10_Change()
   Label7 = ""
   NCMIN% = Val(Text10.TEXT)
   If Val(NCMIN%) > 0 Then
     Label7 = rasocli$(NCMIN%)
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
     Label8 = rasocli$(NCMAX%)
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
   NBR& = Val(LABEL5)
   List2.Clear
   If NBR& > 0 Then
     RFF$ = RECFILT$("BOLRECEP", 1, MKS$(NBR&))
     '\\\OT-08-0641-OD-21/08/08///
     Call ABRESTOCKS
     '
     SQLX$ = " SELECT * FROM BOLRECEP "
     SQLX$ = SQLX$ + " WHERE NumeBoRe = " & NBR& & " "
     '
     Dim BRECEPTMP As ADODB.Recordset
     Set BRECEPTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
     With BRECEPTMP

      On Error Resume Next
       
      If Not (.EOF And .BOF) Then
         .MoveFirst
        On Error GoTo 0
        Do While Not (.EOF)
            If List2.ListCount < 1 Then
                NPRO% = !NumProv
                Label14 = TRIM$(Str$(NPRO%))
                Label11 = rasocli$((-1) * NPRO%)
                Label7 = !numrem_prov ' NUMERO DE REMITO
                FEX = CVINT(!fechrecep)
                Label9 = FECHATEX(FEX)
                TxtAnota$ = SAFETEXT$(!anotaorf): LONGI& = Len(TxtAnota$)
            End If
           On Error GoTo 0
         '\\\OT-08-0641-OD-FIN///
          'For U& = 1 To Len(RFF$) Step 4
           'REBO& = CVS(Mid$(RFF$, U&, 4))
           'XX$ = REGLEIDO$("BOLRECEP", REBO&)
           '
    '       If List2.ListCount < 1 Then      Esta informacion la carga de la tabla directamente
    '         NPRO% = CVI(Mid$(XX$, 105, 2))
    '         Label14 = TRIM$(Str$(NPRO%))
    '         Label11 = RASOCLI$((-1) * NPRO%)
    '         Label7 = TRIM$(Mid$(XX$, 119, 16)) ' NUMERO DE REMITO
    '         FEX = CVI(Mid$(XX$, 81, 2))
    '         Label9 = FECHATEX$(FEX)
    '       End If
           '
            TTXX$ = Space$(128)
              CIXI% = XVALO(!COD_INTE)
            Call REPLA(TTXX$, CODEXT$(CIXI%), 1, 15)
            Call REPLA(TTXX$, DESCRIT0$(CIXI%), 17, 25)
              UME$ = SAFETEXT(!Unidad)
              If UME$ = "" Then UME$ = Unimed$(CIXI%)
            Call REPLA(TTXX$, UME$, 43, 4)
              CANRE = XVALO(!cantnom) 'CVD(Mid$(XX$, 89, 8))
              CANAP = XVALO(!CanToApro) 'CVD(Mid$(XX$, 163, 8))
              CANCH = XVALO(!CanToRech) 'CVD(Mid$(XX$, 171, 8))
              CANFA = XVALO(!CanToFalta) 'CVD(Mid$(XX$, 179, 8))
              
              '
'              If (Abs(CANAP + CANCH + CANFA - CANRE)) >= 0.05 Then
'                Call RECUBORE(REBO&)
'                XX$ = REGLEIDO$("BOLRECEP", REBO&)
'                CANRE = CVD(Mid$(XX$, 89, 8))
'                CANAP = CVD(Mid$(XX$, 163, 8))
'                CANCH = CVD(Mid$(XX$, 171, 8))
'                CANFA = CVD(Mid$(XX$, 179, 8))
'              End If
              '
            Call REPLA(TTXX$, FORMATNUM$(CANRE, "F8.1"), 47, 8)
            Call REPLA(TTXX$, FORMATNUM$(CANAP, "F8.1"), 55, 8)
            Call REPLA(TTXX$, FORMATNUM$(CANCH, "F8.1"), 63, 8)
            Call REPLA(TTXX$, FORMATNUM$(CANFA, "F8.1"), 71, 8)
            List2.AddItem TTXX$
            .MoveNext
        Loop
      End If
     End With
     BRECEPTMP.Close
     Set BRECEPTMP = Nothing

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
'     NBR& = Val(Label5)
'     RFF$ = RECFILT$("ANOTABOL", 1, MKS$(NBR&))
'     For U& = 1 To Len(RFF$) Step 4
'       RENOTA& = CVS(Mid$(RFF$, U&, 4))
'       XX$ = REGLEIDO$("ANOTABOL", RENOTA&)
'       List3.AddItem Mid$(XX$, 5, 92)
'     Next U&
     INICIO& = 1
     Do While INICIO& < LONGI&
       FIN& = InStr(INICIO&, TxtAnota$, Chr$(10))
       If FIN& > 0 Then
          TTXX$ = Mid$(TxtAnota$, INICIO&, FIN& - INICIO&)
          List3.AddItem TTXX$
          INICIO& = FIN& + 2
       Else
          TTXX$ = Mid$(TxtAnota$, INICIO&)
          List3.AddItem TTXX$
          Exit Do
       End If
     Loop

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
   ANOTABOL07.Label16 = LABEL5
   ANOTABOL07.Show 1
   '
   List3.Clear
   NBR& = Val(LABEL5)
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

