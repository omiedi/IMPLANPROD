VERSION 5.00
Begin VB.Form COMLIB04 
   BackColor       =   &H00FFF0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Selección de Items a Comprar"
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   11160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4020
      Left            =   105
      TabIndex        =   3
      Top             =   1680
      Width           =   10935
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0C0C0&
      Height          =   350
      Left            =   105
      ScaleHeight     =   285
      ScaleWidth      =   10875
      TabIndex        =   20
      Top             =   1365
      Width           =   10935
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   $"COMLIB04.frx":0000
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   45
         TabIndex        =   21
         Top             =   60
         Width           =   10830
      End
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
      Height          =   285
      Left            =   420
      MaxLength       =   5
      TabIndex        =   15
      Text            =   " "
      Top             =   555
      Width           =   960
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
      Height          =   285
      Left            =   1365
      TabIndex        =   14
      Top             =   555
      Width           =   195
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1830
      TabIndex        =   13
      TabStop         =   0   'False
      Text            =   " "
      Top             =   210
      Width           =   4485
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1830
      TabIndex        =   12
      TabStop         =   0   'False
      Text            =   " "
      Top             =   540
      Width           =   4485
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1830
      TabIndex        =   11
      TabStop         =   0   'False
      Text            =   " "
      Top             =   870
      Width           =   750
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2670
      TabIndex        =   10
      TabStop         =   0   'False
      Text            =   " "
      Top             =   870
      Width           =   3645
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
      Left            =   8925
      TabIndex        =   9
      Top             =   840
      Width           =   195
   End
   Begin VB.TextBox Text9 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7980
      TabIndex        =   8
      TabStop         =   0   'False
      Text            =   " "
      Top             =   840
      Width           =   960
   End
   Begin VB.TextBox Text11 
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
      Height          =   285
      Left            =   7980
      TabIndex        =   7
      Text            =   " "
      Top             =   210
      Width           =   1120
   End
   Begin VB.PictureBox Picture6 
      BackColor       =   &H00800000&
      Height          =   1170
      Left            =   9520
      ScaleHeight     =   1110
      ScaleWidth      =   1740
      TabIndex        =   0
      Top             =   -30
      Width           =   1800
      Begin VB.CommandButton Command1 
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
         Left            =   105
         Picture         =   "COMLIB04.frx":0091
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Aprueba la Selección de Items a Comprar"
         Top             =   210
         Width           =   650
      End
      Begin VB.CommandButton BOTEXIT 
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
         Left            =   840
         Picture         =   "COMLIB04.frx":039B
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Cancela la Pre-Selección"
         Top             =   210
         Width           =   650
      End
      Begin VB.CommandButton BOTREC 
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
         Picture         =   "COMLIB04.frx":06A5
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Aprueba la Selección de Items Cumplidos"
         Top             =   6555
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   1
         Top             =   6090
         Width           =   640
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   2
            Top             =   105
            Width           =   12000
         End
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Proveedor:"
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
      Left            =   210
      TabIndex        =   19
      Top             =   315
      Width           =   960
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha:"
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
      Left            =   5925
      TabIndex        =   18
      Top             =   900
      Width           =   2010
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Número:"
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
      Left            =   6030
      TabIndex        =   17
      Top             =   285
      Width           =   1905
   End
   Begin VB.Label Label8 
      BackColor       =   &H00EEEEEE&
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
      Height          =   285
      Left            =   420
      TabIndex        =   16
      Top             =   870
      Width           =   1170
   End
End
Attribute VB_Name = "COMLIB04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BOTEXIT_Click()
    '
    Call CIERRARCH("*.*")
    Unload Me
    '
End Sub

Private Sub Command1_Click()
   '
   Command1.Enabled = False
   CRLF$ = Chr$(13) + Chr$(10)
   NPRO% = Val(Text1)
   If REGICLI&((-1) * NPRO%) < 1 Then
     Call MENSERR(24, "Seleccione Proveedor")
     Call Command2_Click
     GoTo 99
   End If
   '
   If List1.ListIndex < 0 Then
     Call MENSERR(24, "Seleccione Item a Comprar")
     GoTo 99
   End If
   '
   XXY$ = List1.TEXT
   NOTRA$ = TRIM$(Left$(XXY$, 16))
   NOGRU$ = TRIM$(Mid$(XXY$, 17, 10))
   NORIT% = Val(Mid$(XXY$, 27, 4))
   CANTU = Val(Mid$(XXY$, 31, 10))
   FENTRESO$ = TRIM$(Mid$(XXY$, 93, 8))
   If FECHANUM(FENTRESO$) < 1 Then
     HOII% = HOY(HOS$)
     FENTRESO$ = HOS$
   End If
   '
   Call ABREORTRA
   With ListaMat
     .FindFirst "Nro_Otrab = '" & NOTRA$ & "' and Nro_Grup = '" & NOGRU$ & "' AND NuOrden = " & NORIT% & " AND CANTID = " & CANTU & " "
     If .NoMatch = False Then
         DECORTA$ = !Descri_Cor$
         DELARGA$ = !Descri_Lar
         CODIMATA$ = !CODEXTE
         NUPLANA$ = !NPlano
         MATERIAX$ = TRIM$(!MATERIAL)
         NUOR$ = CStr(!NuOrden)
       Else
         Call MENSERR(24, "Item no Encontrado")
         GoTo 99
     End If
   End With
   '
   FCOMLIB04.Text1 = Text1
   FCOMLIB04.Text18 = DECORTA$
     If MATERIAX$ <> "" Then FCOMLIB04.Text18 = DECORTA$ & CRLF$ & "Material: " & MATERIAX$
   FCOMLIB04.Text28 = DELARGA$
   FCOMLIB04.Text20 = TRIM$(Str$(CANTU))
   FCOMLIB04.Text22 = "U."
   FCOMLIB04.Label23 = NOTRA$
   FCOMLIB04.Label24 = NOGRU$
   FCOMLIB04.Label25 = ""              ' señal de que no aprobada
   FCOMLIB04.Label29 = CODIMATA$
   FCOMLIB04.Label28 = NUPLANA$
   FCOMLIB04.Label31 = NUOR$
   NEV1$ = NOTRA$
   FCOMLIB04.Text7 = CUEOT(NEV1$) 'CUENTA CONT. DE LA EV
   FCOMLIB04.Text14 = FENTRESO$
   '
   FCOMLIB04.Show 1
   '
   ' LO QUE SIGUE ES PARA MARCAR LOS ITEMS COMPRADOS EN LA LISTA DE MATERIALES
   NOCOMX$ = TRIM$(FCOMLIB04.Label25)
   If NOCOMX$ <> "" Then
     Call ABREORTRA
     With ListaMat
       .FindFirst "Nro_Otrab = '" & NOTRA$ & "' and Nro_Grup = '" & NOGRU$ & "' AND NuOrden = " & NORIT% & " "
       If .NoMatch = False Then
           .Edit
           !Nro_Ocom = NOCOMX$
           .Update
         Else
           Call MENSERR(24, "Registro de Lista de Materiales no Encontrado.\Debe Anular la O/Compra y Re-Emitirla.")
       End If
     End With
     ILISTX& = List1.ListIndex
     If ILISTX& >= 0 Then List1.RemoveItem ILISTX&
   End If
   '
99 Command1.Enabled = True
   '
End Sub
Function CUEOT(NEV1$) As String
   '
   Call ABREORTRA
   SQLX$ = "SELECT * FROM ORDETRAB "
   SQLX$ = SQLX$ + " WHERE Nro_Otrab = '" & NEV1$ & "'"
   Set CUETMP = OTrabajo.OpenRecordset(SQLX$, 4)
   With CUETMP
     CUEOT = ""
     On Error Resume Next
     .MoveFirst
     If Err Then
        On Error GoTo 0
        Exit Function
     End If
     On Error GoTo 0
     If IsNull(!ICONTAB) = True Then
        CUEOT = ""
      Else
        CUEOT = CODCUE(!ICONTAB)
     End If
   End With
   '
End Function
Private Sub Command2_Click()
   '
   Call SELECHO("ACREDOR")
   NCLIE% = Val(TRIM$(VALACT1$("ACREDOR")))
   If NCLIE% > 0 Then
     Text1 = TRIM$(Str$(NCLIE%))
   End If
   '
End Sub

Private Sub Command9_Click()
   Call SELECHO("SELFECHA")
   VDEVU$ = TRIM$(VALACT1$("SELFECHA"))
   If VDEVU$ <> "" Then
      Text9.TEXT = VALACT1$("SELFECHA")
   End If
End Sub

Private Sub Form_Load()
    '
    Call BLANFORMU
    '
End Sub

Private Sub List1_DblClick()
   Call Command1_Click
End Sub

Private Sub Text1_Change()
    '
    NCLIE% = Val(Text1)
    Call BLANFORMU
    If REGICLI&((-1) * NCLIE%) > 0 Then
      Call CARDACLI
    End If
    '
End Sub

Private Sub Text1_DblClick()
   Call Command2_Click
End Sub

Sub CARDACLI()
   '
   Screen.MousePointer = 11
   '
   NC% = (-1) * Val(Text1.TEXT)
   If NC% < 0 Then
     If STATCLI%(NC%) > 0 Then
       Label8 = CODICLI$(NC%)
       Text2.TEXT = RASOCLI$(NC%)
       Text3.TEXT = DOMICLI$(NC%)
       Text4.TEXT = CPOSCLI$(NC%)
       Text5.TEXT = LOCACLI$(NC%)
     End If
   End If
   '
   Screen.MousePointer = 1
   '
End Sub

Sub BLANFORMU()
   '
   Label8 = ""
   Text2.TEXT = ""
   Text3.TEXT = ""
   Text4.TEXT = ""
   Text5.TEXT = ""
   '
   HOII% = HOY(HOS$)
   Text9.TEXT = HOS$
   Text11.TEXT = TRIM$(Str$(PRONUOCO&))
   '
End Sub

Function PRONUOCO&()
   '
   PRONU& = 1
   Call ABRECOMPRAS
   '
   SQLX$ = "SELECT Nume_Ocom FROM OcomEnca "
   SQLX$ = SQLX$ + "ORDER BY Nume_Ocom DESC;"
   Set OEncaTemp = Compras.OpenRecordset(SQLX$, dbOpenSnapshot)
   On Error Resume Next
   With OEncaTemp
     .MoveFirst
     If Err < 1 Then
       PRONU& = 1 + !Nume_Ocom
     End If
   End With
   On Error GoTo 0
   '
   PRONUOCO& = PRONU&
   '
End Function

Private Sub Text11_GotFocus()
   On Error Resume Next 'x err '5' llamada a proc. no valida
    Text1.SetFocus
    Screen.MousePointer = 1
   On Error GoTo 0
End Sub

Private Sub Text9_GotFocus()
   Text1.SetFocus
End Sub
