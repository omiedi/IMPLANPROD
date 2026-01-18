VERSION 5.00
Begin VB.Form FOSEARIM 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ruta y Archivo de Importación de Pedidos"
   ClientHeight    =   4260
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   7095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   7095
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   770
      Left            =   6050
      Picture         =   "FOSEARIM.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Cancela"
      Top             =   360
      Width           =   855
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00E0E0E0&
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
      Height          =   770
      Left            =   6050
      Picture         =   "FOSEARIM.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Inicia Importación del Pedido de Cliente"
      Top             =   3360
      Width           =   855
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Archivo de Importación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3915
      Left            =   210
      TabIndex        =   0
      Top             =   210
      Width           =   5685
      Begin VB.FileListBox File1 
         Height          =   1845
         Left            =   2940
         TabIndex        =   6
         Top             =   420
         Width           =   2430
      End
      Begin VB.DirListBox Dir1 
         Height          =   1440
         Left            =   315
         TabIndex        =   2
         Top             =   840
         Width           =   2430
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   1470
         TabIndex        =   1
         Top             =   420
         Width           =   1275
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000C&
         X1              =   0
         X2              =   5670
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Unidad:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   315
         TabIndex        =   8
         Top             =   470
         Width           =   1065
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Left            =   2940
         TabIndex        =   7
         Top             =   2775
         Width           =   2430
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Archivo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1785
         TabIndex        =   5
         Top             =   2835
         Width           =   1065
      End
      Begin VB.Label Label11 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   315
         TabIndex        =   4
         Top             =   3210
         Width           =   5055
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Carpeta:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -210
         TabIndex        =   3
         Top             =   2940
         Width           =   1065
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000009&
         X1              =   0
         X2              =   5670
         Y1              =   2520
         Y2              =   2520
      End
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   5520
      Picture         =   "FOSEARIM.frx":0614
      Top             =   2205
      Width           =   2010
   End
End
Attribute VB_Name = "FOSEARIM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub command1_Click()
   Call CIERRARCH("*.*")
   Label1 = ""
   Label11 = ""
   Hide
End Sub

Private Sub Command6_Click()
   Command6.Enabled = False
   HOII% = HOY(HOS$)
   If TRIM$(Label1) = "" Then
     Call MENSERR(24, "Falta Nombre de Archivo")
     GoTo 99
   End If
   '
   If EXISTE%(Label11 + "\" + Label1) < 1 Then
     Call MENSERR(24, "Ruta - Archivo Inexistente")
     GoTo 99
   End If
   '
   ARIMPOR$ = Label11 + "\" + Label1
   N1% = FILEOPEN%(ARIMPOR$, 1, 128)
   On Error Resume Next
   Line Input #N1%, AX$
   If Err Then
     On Error GoTo 0
     Call MENSERR(24, "Importación Imposible - Revise Archivo")
     GoTo 99
   End If
   '
   NCLIEN& = Val(AX$)
   If Err Then
     On Error GoTo 0
     Call MENSERR(24, "Primer Registro no Corresponde\a Número de Cliente Válido.\  \Revise Archivo !")
     GoTo 99
   End If
   '
   On Error GoTo 0
   If NCLIEN& < 1 Or NCLIEN& > 32767 Then
     Call MENSERR(24, "Primer Registro '" + TRIM$(AX$) + "' no Corresponde\a Número de Cliente Válido.\  \Revise Archivo !")
     GoTo 99
   End If
   '
   NC% = NCLIEN&
    If REGICLI&(NC%) < 1 Then
     Call MENSERR(24, "Primer Registro '" + TRIM$(AX$) + "' no Corresponde\a Número de Cliente Válido.\  \Revise Archivo !")
     GoTo 99
   End If
   '
   FORPEDCLI.Text1 = TRIM$(Str$(NC%))
   Call FORPEDCLI.CARDACLI
   FORPEDCLI.MARSELPRESU.Visible = False
   '
   On Error Resume Next
   RECNO& = 0: JJ& = 0
20 If EOF(N1%) = False Then
     Line Input #N1%, AX$
     RECNO& = RECNO& + 1
     AX$ = AX$ + Chr$(124)
     PPXX% = InStr(AX$, Chr$(124))
     If PPXX% > 0 Then
         CODE$ = TRIM$(Left$(AX$, PPXX% - 1))
         AX$ = Mid$(AX$, PPXX% + 1)
         If CODINT%(CODE$) < 1 Then
           Call MENSERR(24, "Registro '" + TRIM$(Str$(RECNO&)) + "' - Código '" + CODE$ + "' Inexistente o no Válido.")
           GoTo 29
         End If
       Else
         Call MENSERR(24, "Imposible Procesar Registro '" + TRIM$(Str$(RECNO&)) + "'.")
         GoTo 29
     End If
     '
     PPXX% = InStr(AX$, Chr$(124))
     If PPXX% > 0 Then
         CANTU = Val(Left$(AX$, PPXX% - 1))
         AX$ = Mid$(AX$, PPXX% + 1)
         If CANTU < 1 Then
           Call MENSERR(24, "Registro '" + TRIM$(Str$(RECNO&)) + "' - Falta Cantidad.")
           GoTo 29
         End If
       Else
         Call MENSERR(24, "Imposible Procesar Registro '" + TRIM$(Str$(RECNO&)) + "'.")
         GoTo 29
     End If
     '
     PPXX% = InStr(AX$, Chr$(124))
     If PPXX% > 0 Then
       FENTSOL% = FECHANUM(Left$(AX$, PPXX% - 1))
     End If
     If FENTSOL% < HOII% Then FENTSOL% = HOII%
     '
     ZZ$ = REGBLAN$("!CARDETPE")
       CIXI% = CODINT(CODE$)
     Call REPLA(ZZ$, MKI$(CIXI%), 1, 2)
       '\\\OT-05-0553-WAR-01/08/05///
       ARCHIPRE$ = TRIM$(FORPEDCLI.TEXT21(0).TEXT)
       While Len(ARCHIPRE$) < 3: ARCHIPRE$ = "0" + ARCHIPRE$: Wend
       ARCHIPRE$ = "LIPRE" + ARCHIPRE$
       PRELI = PRELISTA(ARCHIPRE$, CIXI%)
       '\\\OT-05-0553-WAR-FIN///
       DESCUTEX$ = ""
       CPPP = (100 - COEFDESCU(DESCUTEX$)) / 100
       '
       MONEMA% = MONEMI%
       COEFCON = TICAMONE(MONEMA%) / TICAMONE(MONEMI%)
       CANTIPRES = CANTU
       '
     Call REPLA(ZZ$, MKS$(PRELI * CPPP * COEFCON), 39, 4)
     Call REPLA(ZZ$, MKS$(PRELI * CPPP * COEFCON), 51, 4)
     Call REPLA(ZZ$, MKS$(PRELI * CPPP * COEFCON), 65, 4)
     Call REPLA(ZZ$, MKS$(CANTIPRES), 47, 4)   ' CANTIDAD
     Call REPLA(ZZ$, MKS$(PRELI * CPPP * COEFCON), 73, 4)
     Call REPLA(ZZ$, MKI$(FENTSOL%), 63, 2)
     Call REPLA(ZZ$, MKS$(0), 125, 4)
     '
     JJ& = JJ& + 1
     Call GRAREG("!CARDETPE", ZZ$, JJ&)
     '
29   GoTo 20
   End If
   '
   Call SETULTREG("!CARDETPE", JJ&)
   '
   Command6.Enabled = True
   Call FORPEDCLI.CARDETPEDI(1)
   FORPEDCLI.MARSELPRESU.Visible = False
   FORPEDCLI.MARCONOTA.Visible = False
   Hide
   '
99 Command6.Enabled = True
End Sub

Private Sub Dir1_Change()
   Label1 = ""
   Label11.Caption = Dir1.Path
   File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
   On Error GoTo 20
   Dir1.Path = Drive1.Drive
   GoTo 99
   '
20 Resume 21
21 On Error GoTo 0
   Call MENSERR(24, "Unidad no Disponible.")
   Drive1.Drive = Left$(CurDir$, 2)
   '
99 On Error GoTo 0
End Sub

Private Sub File1_Click()
  Call Dir1_Change
  Label1 = File1.filename
End Sub

Private Sub Form_Activate()
     '
     Label1 = ""
     Label11 = ""
     '
End Sub

Private Sub Form_Load()
    '
    EPAC$ = EMPREAC$
    If EPAC$ <> "" Then
      Caption = Caption + " - " + EMPREAC$
    End If
    '
    File1.Pattern = "*.*"
    Label11 = ""
    '
End Sub

Private Sub Form_Resize()
   If WindowState = 1 Then
     Call CIERRARCH("*.*")
     Call BAJALDISCO
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Hide
End Sub


