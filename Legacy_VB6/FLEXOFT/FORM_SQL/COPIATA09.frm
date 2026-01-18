VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form COPIATA09 
   BackColor       =   &H00D0F0F0&
   Caption         =   "Copia de Tareas por Equipo"
   ClientHeight    =   4035
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8265
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   8265
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   330
      Index           =   0
      Left            =   0
      ScaleHeight     =   270
      ScaleWidth      =   1110
      TabIndex        =   25
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D0F0F0&
      Caption         =   "Destino"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1905
      Left            =   120
      TabIndex        =   20
      Top             =   2040
      Width           =   7125
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
         Left            =   1320
         TabIndex        =   8
         Top             =   1320
         Width           =   175
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
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
         Left            =   840
         TabIndex        =   7
         Text            =   "  "
         Top             =   1335
         Width           =   495
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
         Left            =   6750
         TabIndex        =   6
         Top             =   360
         Width           =   175
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
         Height          =   300
         Left            =   5130
         TabIndex        =   5
         Text            =   "  "
         Top             =   360
         Width           =   1650
      End
      Begin VB.Label Label5 
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
         Height          =   300
         Left            =   240
         TabIndex        =   24
         Top             =   840
         Width           =   6670
      End
      Begin VB.Label Label4 
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
         Height          =   300
         Left            =   1680
         TabIndex        =   23
         Top             =   1320
         Width           =   5235
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Parte:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   -555
         TabIndex        =   22
         Top             =   1425
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Equipo:"
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
         Left            =   3720
         TabIndex        =   21
         Top             =   450
         Width           =   1335
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00D0F0F0&
      Caption         =   "Origen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1905
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   7125
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
         Left            =   5130
         TabIndex        =   1
         Text            =   "  "
         Top             =   360
         Width           =   1650
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
         Height          =   300
         Left            =   6750
         TabIndex        =   2
         Top             =   360
         Width           =   175
      End
      Begin VB.TextBox Text7 
         Alignment       =   1  'Right Justify
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
         Left            =   810
         TabIndex        =   3
         Text            =   "  "
         Top             =   1335
         Width           =   495
      End
      Begin VB.CommandButton Command11 
         Caption         =   "."
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
         Left            =   1300
         TabIndex        =   4
         Top             =   1335
         Width           =   175
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Equipo:"
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
         Left            =   3720
         TabIndex        =   19
         Top             =   450
         Width           =   1335
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Parte:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   -555
         TabIndex        =   18
         Top             =   1425
         Width           =   1335
      End
      Begin VB.Label Label13 
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
         Height          =   300
         Left            =   1680
         TabIndex        =   17
         Top             =   1335
         Width           =   5235
      End
      Begin VB.Label Label26 
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
         Height          =   300
         Left            =   240
         TabIndex        =   16
         Top             =   840
         Width           =   6670
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00006060&
      Height          =   10000
      Left            =   7400
      ScaleHeight     =   9945
      ScaleWidth      =   1215
      TabIndex        =   0
      Top             =   0
      Width           =   1275
      Begin VB.CommandButton Command4 
         Caption         =   "Help"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   105
         Picture         =   "COPIATA09.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Ayuda en Línea"
         Top             =   750
         Width           =   650
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   105
         Picture         =   "COPIATA09.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Exit"
         Top             =   105
         Width           =   650
      End
      Begin VB.CommandButton Command99 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "COPIATA09.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Control de Integridad de la Base de Datos"
         Top             =   9000
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   11
         Top             =   2520
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   12
            Top             =   105
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command10 
         Caption         =   "O.K."
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   120
         Picture         =   "COPIATA09.frx":075E
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Guarda Copia de Tareas"
         Top             =   3240
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -255
         Picture         =   "COPIATA09.frx":0A68
         Stretch         =   -1  'True
         Top             =   2760
         Width           =   1410
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   1155
      OleObjectBlob   =   "COPIATA09.frx":298A
      TabIndex        =   26
      Top             =   0
      Visible         =   0   'False
      Width           =   1590
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2730
      OleObjectBlob   =   "COPIATA09.frx":2A04
      Top             =   0
   End
End
Attribute VB_Name = "COPIATA09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   '
   Call SELECHO("INDIEQUI")
   COD$ = TRIM$(VALACT1$("INDIEQUI"))
   If COD$ <> "" Then
     Text1.TEXT = COD$
     Label26 = TRIM$(VALACT2$("INDIEQUI"))
   Else
     Label26 = ""
   End If
   '
End Sub

Private Sub Command10_Click()
    '
    Command10.Enabled = False
    '
    CONDI$ = "Codigo_Equ = '" & Text1.TEXT & "' "
    If CantRegistros("EQUIPOS", CONDI$) < 1 Then
        Call COMUNI("Codigo de Equipo Origen no Valido")
        GoTo 99
    End If
    '
    CONDI$ = "Codigo_Equ = '" & Text2.TEXT & "' "
    If CantRegistros("EQUIPOS", CONDI$) < 1 Then
        Call COMUNI("Codigo de Equipo Destino no Valido")
        GoTo 99
    End If
    '
    CONDI$ = "Codigo_Equ = '" & TRIM$(Text1) & "' and nro_sub = " & XVALO(Text7)
    If CantRegistros("SUBCONJUNTOS", CONDI$) < 1 Then
        Call COMUNI("Codigo de Parte Origen no Valido")
        GoTo 99
    End If
    '
    CONDI$ = "Codigo_Equ = '" & TRIM$(Text2) & "' and nro_sub = " & XVALO(Text3)
    If CantRegistros("SUBCONJUNTOS", CONDI$) < 1 Then
        Call COMUNI("Codigo de Parte Destino no Valido")
        GoTo 99
    End If
    '
    CONDI$ = "codigo_equ = '" & TRIM$(Text2.TEXT) & "' And NRO_SUB = " & XVALO(Text3)
    REEGG& = CantRegistros&("PLAMANPREV", CONDI$)
    If REEGG& > 0 Then
        OPX% = COMALTER%("El Equipo Destino ya Cuenta con Tareas Asignadas\¿Desea Re-Emplazarlas?\\Si, Continuar\No, Cancelar")
        If OPX% <> 1 Then GoTo 99
    End If
    Call ACTUREGISTRO("PLAMANPREV", "Mar_Borra", CONDI$, 1, REG&, "NOMESS")
    '
    SQLX$ = " SELECT * FROM PLAMANPREV "
    SQLX$ = SQLX$ + "where codigo_equ = '" & TRIM$(Text1.TEXT) & "'"
    SQLX$ = SQLX$ + "AND NRO_SUB = " & XVALO(Text7) & ""
    '
    Dim PLAMANTMP As ADODB.Recordset
    Set PLAMANTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    '
    With PLAMANTMP
    Do While Not PLAMANTMP.EOF
        ORDENTAR = !orden_tar
        DESCRITAR$ = !descri_tar
        FRECUTAR$ = !Frecu_tar
        UNIFRECU = !Uni_Frecu
        OBSERTAR$ = !Obse_Tar
        MARBORRA = !Mar_Borra
       
        SQLX$ = "SELECT * FROM PLAMANPREV "
        SQLX$ = SQLX$ & " WHERE codigo_equ = '" & TRIM$(Text2.TEXT) & "' And NRO_SUB = " & XVALO(Text3)
        Dim COPIATMP As ADODB.Recordset
        Set COPIATMP = New ADODB.Recordset
        COPIATMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
        With COPIATMP
           .AddNew   ' YA CONVERTIDO
           On Error Resume Next
           !CodiEmpr = CodiEmp%
           On Error GoTo 0
           !codigo_equ = TRIM$(Text2)
           CONDI$ = "codigo_equ = '" & Text2 & "'"
           CODLARE$ = VALOBADA("Equipos", "codlar_equ", CONDI$, "NOMESS")
           !codlar_equ = TRIM$(CODLARE$)
           !NRO_SUB = XVALO(Text3)
           !orden_tar = ORDENTAR
           !descri_tar = DESCRITAR$
           !Frecu_tar = FRECUTAR$
           !Uni_Frecu = UNIFRECU
           !Obse_Tar = OBSERTAR$
           !Mar_Borra = 0
           .Update
        End With
    PLAMANTMP.MoveNext
    Loop
    End With
    PLAMANTMP.Close
    COPIATMP.Close
    Set PLAMANTMP = Nothing
    Set COPIATMP = Nothing
    '
    CONDI$ = "codigo_equ = '" & TRIM$(Text2.TEXT) & "' And NRO_SUB = " & XVALO(Text3)
    CONDI$ = CONDI$ & " AND MAR_BORRA=1"
    Call BORRAREGISTROS("PLAMANPREV", CONDI$, REG&, "NOMESS")
    '
    Call MENSERR(24, "Proceso de Copia Finalizado")
    '
    Text1 = "": Text2 = ""
    '
99  Command10.Enabled = True
    '
End Sub

Private Sub Command11_Click()
   '
   Call SELECHO("SECSUBC1/Sub-Conjuntos")
   COD$ = TRIM$(VALACT1$("SECSUBC1"))
   If COD$ <> "" Then
     Text7.TEXT = COD$
     Label13 = TRIM$(VALACT2$("SECSUBC1"))
   End If
   '
End Sub

Private Sub Command2_Click()
   '
   Call SELECHO("INDIEQUI")
   COD$ = TRIM$(VALACT1$("INDIEQUI"))
   If COD$ <> "" Then
     Text2.TEXT = COD$
     Label5 = TRIM$(VALACT2$("INDIEQUI"))
   Else
     Label5 = ""
   End If
   '
End Sub

Private Sub Command3_Click()
   Call CIERRARCH("*.*")
   Hide
End Sub

Private Sub Command5_Click()
   Call SELECHO("SECSUBC2/Sub-Conjuntos")
   COD$ = TRIM$(VALACT1$("SECSUBC2"))
   If COD$ <> "" Then
     Text3.TEXT = COD$
     Label4 = TRIM$(VALACT2$("SECSUBC2"))
   End If
End Sub

Private Sub Form_Load()
     Call APLICACIONSKINS(Me)

End Sub

Private Sub Text1_Change()
    '
    Label26 = "": Text7 = ""
    CONDI$ = "Codigo_Equ = '" & Text1.TEXT & "' "
    If CantRegistros("EQUIPOS", CONDI$) > 0 Then
        Call COPYSTRU("SECPROD", "SECSUBC1")
        Call CARGALISEL("SECSUBC1", "SUBCONJUNTOS", "NRO_SUB/F2;DESCRI_SUB/A20", CONDI$, "NOMESS")
        Label26 = VALOBADA("EQUIPOS", "DESCRI_EQU", CONDI$, "NOMESS")
    End If
    '
End Sub

Private Sub Text1_DblClick()
    Call Command1_Click
End Sub

Private Sub Text2_Change()
    '
    Label5 = "": Text3 = ""
    CONDI$ = "Codigo_Equ = '" & TRIM$(Text2.TEXT) & "'"
    If CantRegistros("EQUIPOS", CONDI$) > 0 Then
        Call COPYSTRU("SECPROD", "SECSUBC2")
        Call CARGALISEL("SECSUBC2", "SUBCONJUNTOS", "NRO_SUB/F2;DESCRI_SUB/A20", CONDI$, "NOMESS")
        Label5 = VALOBADA("EQUIPOS", "DESCRI_EQU", CONDI$, "NOMESS")
    End If
    '
End Sub

Private Sub Text2_DblClick()
    Call Command2_Click
End Sub

Private Sub Text3_Change()
    CONDI$ = "Codigo_Equ = '" & TRIM$(Text2) & "' and nro_sub = " & XVALO(Text3)
    Label4 = VALOBADA("SUBCONJUNTOS", "DESCRI_SUB", CONDI$, "NOMESS")
End Sub

Private Sub Text3_DblClick()
    Call Command5_Click
End Sub

Private Sub Text7_Change()
    CONDI$ = "Codigo_Equ = '" & TRIM$(Text1) & "' and nro_sub = " & XVALO(Text7)
    Label13 = VALOBADA("SUBCONJUNTOS", "DESCRI_SUB", CONDI$, "NOMESS")
End Sub

Private Sub Text7_DblClick()
    Call Command11_Click
End Sub
