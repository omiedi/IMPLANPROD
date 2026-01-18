VERSION 5.00
Begin VB.Form FORCONTAC 
   BackColor       =   &H80000016&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Base de Datos de Contactos"
   ClientHeight    =   7695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8865
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   8865
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command7 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7850
      Picture         =   "FORCONTAC.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   39
      ToolTipText     =   "Imprime Ficha del Contacto"
      Top             =   3255
      Width           =   855
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
      Height          =   330
      Left            =   7455
      TabIndex        =   0
      Top             =   420
      Width           =   225
   End
   Begin VB.CommandButton Command3 
      Caption         =   "New"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7850
      Picture         =   "FORCONTAC.frx":066A
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Alta de Nuevo Contacto"
      Top             =   2310
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "O.K."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   7850
      Picture         =   "FORCONTAC.frx":0974
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Aprueba y Graba Datos del Contacto"
      Top             =   1365
      Width           =   855
   End
   Begin VB.CommandButton Command2 
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
      Height          =   855
      Left            =   7850
      Picture         =   "FORCONTAC.frx":0C7E
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Cierra - Abandona la Aplicación"
      Top             =   420
      Width           =   855
   End
   Begin VB.TextBox Text17 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3045
      TabIndex        =   16
      Text            =   " "
      Top             =   5250
      Width           =   4425
   End
   Begin VB.TextBox Text16 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   315
      TabIndex        =   14
      Text            =   " "
      Top             =   5250
      Width           =   2115
   End
   Begin VB.TextBox Text15 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3045
      TabIndex        =   15
      Text            =   " "
      Top             =   4830
      Width           =   4425
   End
   Begin VB.TextBox Text14 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   315
      TabIndex        =   13
      Text            =   " "
      Top             =   4830
      Width           =   2115
   End
   Begin VB.TextBox Text13 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5775
      TabIndex        =   12
      Text            =   " "
      Top             =   3885
      Width           =   1695
   End
   Begin VB.TextBox Text12 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4095
      TabIndex        =   10
      Text            =   " "
      Top             =   3885
      Width           =   750
   End
   Begin VB.TextBox Text11 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1365
      TabIndex        =   9
      Text            =   " "
      Top             =   3885
      Width           =   2220
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
      Height          =   330
      Left            =   5775
      TabIndex        =   11
      Text            =   " "
      Top             =   3465
      Width           =   1695
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
      Height          =   330
      Left            =   4095
      TabIndex        =   8
      Text            =   " "
      Top             =   3465
      Width           =   750
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
      Height          =   330
      Left            =   1365
      TabIndex        =   7
      Text            =   " "
      Top             =   3465
      Width           =   2220
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
      Height          =   330
      Left            =   1365
      TabIndex        =   6
      Text            =   " "
      Top             =   2520
      Width           =   6105
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2835
      MaxLength       =   48
      TabIndex        =   5
      Text            =   " "
      Top             =   2100
      Width           =   4635
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
      Height          =   330
      Left            =   1365
      MaxLength       =   8
      TabIndex        =   4
      Text            =   " "
      Top             =   2100
      Width           =   1275
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
      Height          =   330
      Left            =   1365
      MaxLength       =   48
      TabIndex        =   3
      Text            =   " "
      Top             =   1680
      Width           =   6105
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1365
      MaxLength       =   48
      TabIndex        =   2
      Text            =   " "
      Top             =   1260
      Width           =   6105
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
      Height          =   330
      Left            =   1365
      MaxLength       =   48
      TabIndex        =   1
      Text            =   " "
      Top             =   420
      Width           =   3795
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000016&
      Caption         =   "Empresa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2115
      Left            =   105
      TabIndex        =   21
      Top             =   945
      Width           =   7575
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "R. Social:"
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
         Left            =   210
         TabIndex        =   25
         Top             =   420
         Width           =   960
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Calle y Nro.:"
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
         Left            =   105
         TabIndex        =   24
         Top             =   840
         Width           =   1065
      End
      Begin VB.Label Label4 
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
         Height          =   225
         Left            =   105
         TabIndex        =   23
         Top             =   1260
         Width           =   1065
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "País:"
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
         Left            =   315
         TabIndex        =   22
         Top             =   1680
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000016&
      Caption         =   "Teléfono"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   105
      TabIndex        =   26
      Top             =   3150
      Width           =   7575
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Partic.:"
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
         Left            =   4830
         TabIndex        =   30
         Top             =   420
         Width           =   750
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Celular"
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
         Left            =   4935
         TabIndex        =   29
         Top             =   840
         Width           =   645
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Int:"
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
         Left            =   3255
         TabIndex        =   28
         Top             =   420
         Width           =   645
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Comercial:"
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
         Left            =   105
         TabIndex        =   27
         Top             =   420
         Width           =   1065
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000016&
      Caption         =   "Fax"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   105
      TabIndex        =   31
      Top             =   4515
      Width           =   2535
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H80000016&
      Caption         =   "E-mail"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   2835
      TabIndex        =   32
      Top             =   4515
      Width           =   4845
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H80000016&
      Caption         =   "Clientes Vinculados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   105
      TabIndex        =   33
      Top             =   5880
      Width           =   7575
      Begin VB.CommandButton Command6 
         Caption         =   "Quitar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Left            =   6300
         TabIndex        =   38
         Top             =   945
         Width           =   1065
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Agregar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Left            =   6300
         TabIndex        =   37
         Top             =   315
         Width           =   1065
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
         Height          =   1185
         Left            =   210
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   36
         Top             =   315
         Width           =   5895
      End
   End
   Begin VB.Label REGCONTA 
      Height          =   330
      Left            =   6405
      TabIndex        =   40
      Top             =   0
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   7300
      Picture         =   "FORCONTAC.frx":0DC8
      Top             =   7050
      Width           =   2010
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
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
      Height          =   330
      Left            =   6405
      TabIndex        =   35
      Top             =   420
      Width           =   1065
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Nro.:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5775
      TabIndex        =   34
      Top             =   525
      Width           =   540
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   315
      TabIndex        =   20
      Top             =   525
      Width           =   1065
   End
End
Attribute VB_Name = "FORCONTAC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RECONTANT$
'
Private Sub Command1_Click()
   Call GRACONTAC
End Sub

Private Sub Command2_Click()
   Call CIERRARCH("*.*")
   If App.EXEName = "BACONTAC" Then
       Call FINAL("")
     Else
       Hide
   End If
End Sub

Private Sub Command3_Click()
   Call BLANFORMU
   REGCONTA.Caption = "0"
   RECONTANT$ = ""
   Label7.Caption = TRIM$(Str$(1 + ULTREG&("CONTACT1")))
   Text1.SetFocus
End Sub

Private Sub Command4_Click()
   Call SELECHO("CONTACT1")
   Label7.Caption = VALACT1$("CONTACT1")
   Call CARCONTAC
End Sub

Private Sub Command5_Click()
   '
   NCON& = Val(Label7.Caption)
   If NCON& < 1 Or NCON& > 32767 Then
      Exit Sub
   End If
   NCONI% = NCON&
   If TRIM$(ECOARCH$("CONTACT1", MKI$(NCONI%))) = "" Then
      Exit Sub
   End If
   '
10 Call SELECHO("DEUDOR1")
   NC% = Val(VALACT1$("DEUDOR1"))
   If NC% < 1 Then Exit Sub
   '
   If ECOARCH$("DEUDOR1", MKI$(NC%)) <> "" Then
      NUCLI$ = AJUSTD$("(" + TRIM$(CSTRING$(MKI$(NC%), 1, 8, 0, 2)) + ")", 10)
      RELI$ = Space$(80)
      Call REPLA(RELI$, RASOCLI$(NC%), 1, 32)
      Call REPLA(RELI$, NUCLI$, 33, 10)
      Call REPLA(RELI$, TRIM$(Str$(NC%)), 65, 6)
      List1.AddItem RELI$
      GoTo 10
   End If
   '
End Sub

Private Sub Command6_Click()
   '
10 For U& = 0 To List1.ListCount - 1
      If List1.Selected(U&) = True Then
         List1.RemoveItem U&
         GoTo 10
      End If
   Next U&
   '
End Sub

Private Sub Command7_Click()
   PRINTFORM
End Sub

Private Sub Form_Load()
   '
   If App.EXEName = "BACONTAC" Then
      '
      Call VERJOBID(OKEY%)
      If OKEY% < 1 Then Call FINAL("")
      '
      EPAC$ = TRIM$(EMPREAC$)
      If EPAC$ <> "" Then
         Caption = Caption + "  -  " + EPAC$
      End If
      '
      COMMA$ = Command$
      PPX% = InStr(COMMA$, "#" + JOBID$ + "#")
      If PPX% > 0 Then
         COMMA$ = Left$(COMMA$, PPX% - 1)
      End If
      If COMMA$ <> "" Then
         Label7.Caption = COMMA$
      End If
   End If
   '
End Sub
'
Sub CARCONTAC()
   '
   NCON& = Val(Label7.Caption)
   If NCON& < 1 Or NCON& > 32767 Then
      Call BLANFORMU
      Exit Sub
   End If
   NCONI% = NCON&
   '
   RFF$ = RECFILT$("CONTACT1", 1, MKI$(NCONI%))
   REG& = 0: If Len(RFF$) >= 4 Then REG& = CVS(Left$(RFF$, 4))
   If REG& < 1 Or REG& > ULTREG&("CONTACT1") Then
      Call MENSERR(24, "Numero de Contacto Inexistente")
      Exit Sub
   End If
   '
   Screen.MousePointer = 11
   '
   RECON$ = REGLEIDO$("CONTACT1", REG&)
   Label7.Caption = TRIM$(Str$(NCONI%))
   Text1.TEXT = TRIM$(Mid$(RECON$, 3, 48))
   Text3.TEXT = TRIM$(Mid$(RECON$, 51, 48))
   Text4.TEXT = TRIM$(Mid$(RECON$, 99, 48))
   Text5.TEXT = TRIM$(Mid$(RECON$, 147, 8))
   Text6.TEXT = TRIM$(Mid$(RECON$, 155, 48))
   Text7.TEXT = TRIM$(Mid$(RECON$, 203, 32))
   Text8.TEXT = TRIM$(Mid$(RECON$, 235, 16))
   Text9.TEXT = TRIM$(Mid$(RECON$, 251, 4))
   Text11.TEXT = TRIM$(Mid$(RECON$, 255, 16))
   Text12.TEXT = TRIM$(Mid$(RECON$, 271, 4))
   Text10.TEXT = TRIM$(Mid$(RECON$, 275, 16))
   Text13.TEXT = TRIM$(Mid$(RECON$, 291, 16))
   Text14.TEXT = TRIM$(Mid$(RECON$, 307, 16))
   Text16.TEXT = TRIM$(Mid$(RECON$, 323, 16))
   Text15.TEXT = TRIM$(Mid$(RECON$, 339, 48))
   Text17.TEXT = TRIM$(Mid$(RECON$, 387, 48))
   '
   REGCONTA.Caption = TRIM$(Str$(REG&))
   RECONTANT$ = RECON$
   Call CIERRARCH("CONTACT1")
   '
   List1.Clear
   RNC$ = RECFILT$("CONTACT2", 1, MKI$(NCONI%))
   URN& = ULTREG&("CONTACT2")
   '
   For U& = 1 To Len(RNC$) Step 4
      RENOTA& = CVS(Mid$(RNC$, U&, 4))
      If RENOTA& > 0 Then
         If RENOTA& <= URN& Then
            TTYY$ = Mid$(REGLEIDO$("CONTACT2", RENOTA&), 3, 2)
            NC% = CVI(TTYY$)
            If NC% > 0 Then
               If ECOARCH$("DEUDOR1", MKI$(NC%)) <> "" Then
                  NUCLI$ = AJUSTD$("(" + TRIM$(CSTRING$(MKI$(NC%), 1, 8, 0, 2)) + ")", 10)
                  RELI$ = Space$(80)
                  Call REPLA(RELI$, RASOCLI$(NC%), 1, 32)
                  Call REPLA(RELI$, NUCLI$, 33, 10)
                  Call REPLA(RELI$, TRIM$(Str$(NC%)), 65, 6)
                  List1.AddItem RELI$
               End If
            End If
         End If
      End If
   Next U&
   Call CIERRARCH("CONTACT2")
   '
   Screen.MousePointer = 1
   '
End Sub

Sub GRACONTAC()
   '
   NCON& = Val(Label7.Caption)
   If NCON& < 1 Or NCON& > 32767 Then
      Call MENSERR(24, "Imposible Registrar.\Número de Contacto Fuera de Rango.")
      Exit Sub
   End If
   NCONI% = NCON&
   '
   DECON$ = TRIM$(Text1.TEXT)
   If DECON$ = "" Then
      Call MENSERR(24, "Imposible Registrar.\Falta Nombre del Contacto.")
      Exit Sub
   End If
   '
   Call BLOQARCH("CONTACT1")
   Call BLOQARCH("CONTACT2")
   '
   Screen.MousePointer = 11
   '
   RECON$ = REGBLAN$("CONTACT1")
   Call REPLA(RECON$, MKI$(NCONI%), 1, 2)
   Call REPLA(RECON$, DECON$, 3, 48)
   Call REPLA(RECON$, Text3.TEXT, 51, 48)
   Call REPLA(RECON$, Text4.TEXT, 99, 48)
   Call REPLA(RECON$, Text5.TEXT, 147, 8)
   Call REPLA(RECON$, Text6.TEXT, 155, 48)
   Call REPLA(RECON$, Text7.TEXT, 203, 32)
   Call REPLA(RECON$, Text8.TEXT, 235, 16)
   Call REPLA(RECON$, Text9.TEXT, 251, 4)
   Call REPLA(RECON$, Text11.TEXT, 255, 16)
   Call REPLA(RECON$, Text12.TEXT, 271, 4)
   Call REPLA(RECON$, Text10.TEXT, 275, 16)
   Call REPLA(RECON$, Text13.TEXT, 291, 16)
   Call REPLA(RECON$, Text14.TEXT, 307, 16)
   Call REPLA(RECON$, Text16.TEXT, 323, 16)
   Call REPLA(RECON$, Text15.TEXT, 339, 48)
   Call REPLA(RECON$, Text17.TEXT, 387, 48)
   '
   REG& = Val(REGCONTA.Caption)
   If REG& > 0 And REG& <= ULTREG&("CONTACT1") Then
      XX$ = REGLEIDO$("CONTACT1", REG&)
      If XX$ <> RECONTANT$ Or CVI(Left$(XX$, 2)) <> NCONI% Then
         Call MENSERR(24, "Imposible Grabar Modificaciones - Vuelva a Intentar")
         Call CIERRARCH("CONTACT1")
         Call CIERRARCH("CONTACT2")
         Call CARCONTAC
         GoTo 99
      End If
      Call GRAREG("CONTACT1", RECON$, REG&)
      RECONTANT$ = RECON$
      GoTo 80
   End If
   '
   REG& = 1 + ULTREG&("CONTACT1")
   If REG& > 32767 Then
      Call MENSERR(24, "Imposible Grabar Modificaciones - Vuelva a Intentar")
      Call CIERRARCH("CONTACT1")
      Call CIERRARCH("CONTACT2")
      Call CARCONTAC
      GoTo 99
   End If
   NCONI% = REG&
   Call REPLA(RECON$, MKI$(NCONI%), 1, 2)
   Call REGAPP("CONTACT1", RECON$)
   REGCONTA.Caption = TRIM$(Str$(REG&))
   RECONTANT$ = RECON$
   '
80 URN& = ULTREG&("CONTACT2")
   RNC$ = RECFILT$("CONTACT2", 1, MKI$(NCONI%))
   KK& = (-3)
   '
   For LL% = 0 To List1.ListCount - 1
      NCLI% = Val(Mid$(List1.List(LL%), 65, 6))
      If NCLI% > 0 Then
         If ECOARCH$("DEUDOR1", MKI$(NCLI%)) <> "" Then
            TTYY$ = MKI$(NCONI%) + MKI$(NCLI%)
            KK& = KK& + 4
            If KK& <= Len(RNC$) - 3 Then
                RENOTA& = CVS(Mid$(RNC$, KK&, 4))
              Else
                URN& = URN& + 1
                RENOTA& = URN&
            End If
            Call GRAREG("CONTACT2", TTYY$, RENOTA&)
         End If
      End If
   Next LL%
   '
   While KK& < Len(RNC$) - 3
      KK& = KK& + 4
      RENOTA& = CVS(Mid$(RNC$, KK&, 4))
      Call GRAREG("CONTACT2", MKI$(NCONI%) + MKI$(0), RENOTA&)
   Wend
   '
90 Call CIERRARCH("CONTACT1")
   Call CIERRARCH("CONTACT2")
   Call BAJALDISCO
   '
99 Call FRESHECHO("CONTACT1")
   Screen.MousePointer = 1
   '
End Sub

Sub BLANFORMU()
   '
   Label7.Caption = ""
   Text1.TEXT = ""
   Text3.TEXT = ""
   Text4.TEXT = ""
   Text5.TEXT = ""
   Text6.TEXT = ""
   Text7.TEXT = ""
   Text8.TEXT = ""
   Text9.TEXT = ""
   Text11.TEXT = ""
   Text12.TEXT = ""
   Text10.TEXT = ""
   Text13.TEXT = ""
   Text14.TEXT = ""
   Text16.TEXT = ""
   Text15.TEXT = ""
   Text17.TEXT = ""
   '
   List1.Clear
   '
End Sub
Sub CARESPECI()
   '
   COD$ = Text1.TEXT
   CI% = CODINT%(Text1.TEXT)
   Text3.TEXT = ""
   If CI% > 0 Then
      If CI% <= NUMAS% Then
      End If
   End If
   Text3.Refresh
   '
End Sub
'
Sub GRAESPECI()
   '
End Sub
'

Private Sub Label7_DblClick()
   Call SELECHO("CONTACT1")
   Label7.Caption = VALACT1$("CONTACT1")
   Call CARCONTAC
End Sub

Private Sub List1_DblClick()
   NCLIEN% = Val(Mid$(List1.TEXT, 65, 6))
   Call CONECRUN("AMACLI/" + TRIM$(Str$(NCLIEN%)), "Consulta de Clientes")
End Sub
