VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form AjuConsig 
   Caption         =   "Ajuste de Consignaciones a Proveedores"
   ClientHeight    =   7125
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   14175
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   10920
      TabIndex        =   15
      Top             =   120
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   1320
      Left            =   35
      TabIndex        =   0
      Top             =   0
      Width           =   14100
      Begin VB.CommandButton Command52 
         Caption         =   "Consultar"
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
         Left            =   12600
         TabIndex        =   11
         Top             =   720
         Width           =   1380
      End
      Begin VB.TextBox TXTCODIGO 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7320
         MaxLength       =   15
         TabIndex        =   8
         Top             =   240
         Width           =   2565
      End
      Begin VB.CommandButton Command1 
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
         Height          =   330
         Left            =   9885
         Picture         =   "AjuConsig.frx":0000
         TabIndex        =   7
         Top             =   240
         Width           =   175
      End
      Begin VB.TextBox TXTPROV 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1200
         MaxLength       =   15
         TabIndex        =   3
         Top             =   255
         Width           =   2565
      End
      Begin VB.CommandButton Command9 
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
         Height          =   330
         Left            =   3765
         Picture         =   "AjuConsig.frx":030A
         TabIndex        =   2
         Top             =   260
         Width           =   175
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   6240
         TabIndex        =   10
         Top             =   825
         Width           =   1020
      End
      Begin VB.Label LBLDESCRIIPCION 
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
         Height          =   315
         Left            =   7320
         TabIndex        =   9
         Top             =   720
         Width           =   5000
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   6240
         TabIndex        =   6
         Top             =   345
         Width           =   1020
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "R.Social:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   810
         Width           =   1020
      End
      Begin VB.Label LBLRSOCIAL 
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
         Height          =   315
         Left            =   1200
         TabIndex        =   4
         Top             =   705
         Width           =   4215
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Proveedor:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   38
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1020
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSF 
      Height          =   2955
      Left            =   0
      TabIndex        =   13
      Top             =   4200
      Width           =   14100
      _ExtentX        =   24871
      _ExtentY        =   5212
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      FillStyle       =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid MSF2 
      Height          =   2880
      Left            =   0
      TabIndex        =   14
      ToolTipText     =   "Un Click para filtrar grilla inferior"
      Top             =   1320
      Width           =   14100
      _ExtentX        =   24871
      _ExtentY        =   5080
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      FillStyle       =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      Height          =   300
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Menu mnuarchivo 
      Caption         =   "Archivo"
      Visible         =   0   'False
      Begin VB.Menu mnuCopiar 
         Caption         =   "Copiar Celda Seleccionada"
         Visible         =   0   'False
      End
      Begin VB.Menu s 
         Caption         =   "_"
      End
   End
End
Attribute VB_Name = "AjuConsig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim GLOBALFRM_strContenidoCelda$
Private Sub Command1_Click()
   Command1.Enabled = False
   Call SELECHO("MASTER")
   COD$ = TRIM$(VALACT1$("MASTER"))
   '
   If COD$ <> "" Then
     Me.TXTCODIGO = COD$
   End If
99 Command1.Enabled = True
End Sub
Sub MENU_CELDA(MSF As MSFlexGrid, REG&, COL%)
         
        If MSF.Rows <= 1 Then Exit Sub
        
        Call COLOREAR_GRILLA_SOLACELDA(MSF, vbYellow, REG&, Int(COL%))   ' LA PINTA DE AMARILLA
        Call INVICTRL
        GLOBALFRM_strContenidoCelda$ = TRIM$(MSF.TextMatrix(REG&, COL%))  'CELDA SELECCIONADA
        PopupMenu mnuarchivo
        Call INVICTRL
        Call COLOREAR_GRILLA_SOLACELDA(MSF, vbWhite, REG&, COL%)  ' LA PINTA DE AMARILLA
End Sub
Sub INVICTRL()
  '
  mnuCopiar.Visible = Not (mnuCopiar.Visible)
  '
End Sub
Private Sub mnucopiar_Click()
       Clipboard.Clear
       Clipboard.SetText GLOBALFRM_strContenidoCelda$
End Sub
Private Sub Command2_Click()
'   Call AJUSARMADOR
End Sub

Private Sub Command52_Click()
   Command52.Enabled = False
   
   CI1% = CODINT%(Me.TXTCODIGO)
   PROV = XVALO(Me.TXTPROV)
   MSF.Rows = 1: MSF2.Rows = 1
   '
   SQLX$ = "SELECT * FROM MACONSIG WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE COD_INTE > 0"
   If CI1% > 0 Then SQLX$ = SQLX$ + " AND COD_INTE = " & CI1%
   If PROV > 0 Then SQLX$ = SQLX$ + " AND NPRO_OCOM = " & PROV
   SQLX$ = SQLX$ + " AND CANT_RENDID < CANT_ENTREG "
   SQLX$ = SQLX$ + " ORDER BY NRO_OCOM DESC "
   Set RSS = READSET(SQLX$)
   MSF2.Redraw = False
   CAMPOS$ = "Cta./F6;R.Social/A24;Código/A16;Descripción/A24;Rto.Nro./F8;Ord.Compra/A12;Cant.Asig./F10.1;Cant.Ent./F10.1;Cant.Rend./F10.1;Cant.Dev./F10.1"
   INICIALIZO% = 0
   With RSS
     Do While Not .EOF
        NUMCOM$ = TRIM$(SAFETEXT$(!NRO_OCOM))
        TEXTOGRILLA$ = XVALO(!NPRO_OCOM) & "|" & rasocli$(XVALO(!NPRO_OCOM)) & "|" & CODEXT$(XVALO(!Cod_Inte)) & "|"
        TEXTOGRILLA$ = TEXTOGRILLA$ & DESCRIT0$(XVALO(!Cod_Inte)) & "|" & XVALO(!NREMI_CONSIG) & "|" & SAFETEXT$(!NRO_OCOM) & "|"
        TEXTOGRILLA$ = TEXTOGRILLA$ & FORMATNUM$(XVALO(!Cant_Asig), "F10.1") & "|" & FORMATNUM$(XVALO(!Cant_Entreg), "F10.1") & "|"
        TEXTOGRILLA$ = TEXTOGRILLA$ & FORMATNUM$(XVALO(!Cant_Rendid), "F10.1" & "|" & FORMATNUM$(XVALO(!Cant_Devu), "F10.1"))
        Call CARGA_GRILLA_GENERICO(Me, Me.MSF2, CAMPOS$, TEXTOGRILLA$, INICIALIZO%)
      .MoveNext
     Loop
   End With
   MSF2.Redraw = True
   
   CAMPOS$ = "Cta./F6;R.Social/A24;Código/A16;Descripción/A24;Fecha/D8;Ord.Compra/A12;Cant.Consig/F10.1;Cant.Dev./F10.1;Cant.Rend./F10.1"
   Call PRESENTA_INFO_RENCONSIG(Me, Me.MSF, CAMPOS$, CI1%, PROV)
99 Command52.Enabled = True

End Sub
Sub PRESENTA_INFO_RENCONSIG(FORMULARIOX As Form, MSF As MSFlexGrid, CAMPOS$, CI1%, PROV)

   SQLX$ = "SELECT NU_PROVE,CODI_ELEM, FE_RECEP,NRO_OCOM,CANT_RENDID,TIPOMOVIM FROM RENCONSIG WITH(NOLOCK) "
   SQLX$ = SQLX$ + "  WHERE CODI_ELEM > 0 "
   If CI1% > 0 Then SQLX$ = SQLX$ + " AND CODI_ELEM = " & CI1%
   If PROV > 0 Then SQLX$ = SQLX$ + " AND NU_PROVE = " & PROV
   SQLX$ = SQLX$ + " ORDER BY NU_PROVE,CODI_ELEM,NRO_OCOM DESC"
   Set RS = READSET(SQLX$)
   MSF.Redraw = False
   INICIALIZO% = 0

   With RS
     Do While Not .EOF
       FE% = FECHNUM(CVDAT(CVINT(!FE_RECEP)))
       NCOM$ = SAFETEXT$(!NRO_OCOM)
       CANRE = XVALO(!Cant_Rendid)
       TMOVI$ = UCase$(TRIM$(SAFETEXT$(!TIPOMOVIM)))
       CANTICONSI = 0: CANTDEVU = 0: CANTREN = 0
       If TMOVI$ = "CC" Then
         CANTICONSI = CANRE
       ElseIf TMOVI$ = "DV" Or TMOVI$ = "QD" Then
         CANTDEVU = CANRE
       ElseIf TMOVI$ = "QR" Then
         CANTREN = CANRE
       End If
       NPROV = XVALO(!NU_PROVE)
       CODI$ = CODEXT$(XVALO(!CODI_ELEM))
       DESXRIPCION$ = DESCRIT0$(XVALO(!CODI_ELEM))
       RAPRO$ = rasocli$(-1 * XVALO(!NU_PROVE))
       '
       TEXTOGRILLA$ = NPROV & "|" & RAPRO$ & "|" & CODI$ & "|" & DESXRIPCION$ & "|" & FECHATEX$(FE%) & "|" & NCOM$ & "|"
       TEXTOGRILLA$ = TEXTOGRILLA$ & FORMATNUM$(CANTICONSI, "F10.1") & "|" & FORMATNUM$(CANTDEVU, "F10.1") & "|"
       TEXTOGRILLA$ = TEXTOGRILLA$ & FORMATNUM$(CANTREN, "F10.1")
       Call CARGA_GRILLA_GENERICO(FORMULARIOX, MSF, CAMPOS$, TEXTOGRILLA$, INICIALIZO%)
       '
       .MoveNext
     Loop
   End With
   MSF.Redraw = True
   
End Sub

Private Sub Command9_Click()
   Command9.Enabled = False
   Call SELECHO("PROVED1")
   NCLIE% = XVALO(TRIM$(VALACT1$("PROVED1")))
   If NCLIE% > 0 Then
      Me.TXTPROV = NCLIE%
   End If

99 Command9.Enabled = True
End Sub

Private Sub Form_Load()
   CAMPOS$ = "Fecha/D8;Ord.Compra/A12;Cant.Consig/F10.1;Cant.Dev./F10.1;Cant.Rend./F10.1"
End Sub

Private Sub MSF2_Click()
    MSF2.COL = 0
    REG& = MSF2.MouseRow
    If REG& < 1 Then Exit Sub
   PROV = XVALO(MSF2.TextMatrix(REG&, 1))
   CODI$ = MSF2.TextMatrix(REG&, 3)
   CI1% = CODINT%(CODI$)
   CAMPOS$ = "Cta./F6;R.Social/A24;Código/A16;Descripción/A24;Fecha/D8;Ord.Compra/A12;Cant.Consig/F10.1;Cant.Dev./F10.1;Cant.Rend./F10.1"
   Call PRESENTA_INFO_RENCONSIG(Me, Me.MSF, CAMPOS$, CI1%, PROV)
End Sub

Private Sub MSF2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Button = 2 Then
     REG& = MSF2.MouseRow
     COL% = MSF2.MouseCol
     Call MENU_CELDA(MSF2, REG&, COL%)
  End If
End Sub
Private Sub MSF_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Button = 2 Then
     REG& = MSF.MouseRow
     COL% = MSF.MouseCol
     Call MENU_CELDA(MSF, REG&, COL%)
  End If
End Sub

Private Sub TXTCODIGO_Change()
    CI1% = CODINT%(TXTCODIGO)
    If CI1% > 0 Then
      Me.LBLDESCRIIPCION = DESCRIT0$(CI1%)
    Else
      Me.LBLDESCRIIPCION = ""
    End If
    MSF.Rows = 1
End Sub

Private Sub TXTPROV_Change()
    If XVALO(TXTPROV) > 0 Then
      Me.LBLRSOCIAL = rasocli$(XVALO(-1 * Me.TXTPROV))
    Else
      Me.LBLRSOCIAL = ""
    End If
    MSF.Rows = 1
End Sub
