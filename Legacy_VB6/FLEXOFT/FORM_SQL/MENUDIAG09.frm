VERSION 5.00
Begin VB.Form MENUDIAG09 
   Caption         =   "Form1"
   ClientHeight    =   3870
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7590
   LinkTopic       =   "Form1"
   ScaleHeight     =   3870
   ScaleWidth      =   7590
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   7
      Top             =   720
      Width           =   6495
      Begin VB.CommandButton Command6 
         Caption         =   "Rep_Fact"
         Height          =   375
         Left            =   5160
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Recorre el Cajabanc verificando que no existan repeticiones de Facturas en caso  contrario genera un archivo.txt"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   5055
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00004000&
      Height          =   4575
      Left            =   6720
      ScaleHeight     =   4515
      ScaleWidth      =   915
      TabIndex        =   0
      Top             =   0
      Width           =   975
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   540
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   3120
         Width           =   600
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
         Height          =   600
         Left            =   120
         Picture         =   "MENUDIAG09.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Terminar - Salir de la Aplicación"
         Top             =   120
         Width           =   600
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -320
         Picture         =   "MENUDIAG09.frx":014A
         Top             =   3400
         Width           =   1515
      End
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   6495
      Begin VB.CommandButton Command7 
         Caption         =   "Vend vs. Vend"
         Height          =   375
         Left            =   5160
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Genera un archivo informando las facturas que no coinicide el vendedor del pedido con el vendedor de la factura"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   5055
      End
   End
End
Attribute VB_Name = "MENUDIAG09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Call FINAL("")
End Sub

Private Sub Command6_Click()
   '
   'VERIFICACION DE REPETICION DE FACTURAS CON MISMO NUMERO EN CASO DE ENCONTRAR GENERA ARCHIVO INFORMATIVO
   '
   Command6.Enabled = False
   Screen.MousePointer = 11
   '
   Call ABRECABAN
   '
   NARCHI2% = FILEOPEN%("C:\Flexoft\" + "Repet_Facturas.txt", 2, 128) ' ABRE EL ARCHIVO DE ESCRITURA
   CONDI$ = "TipoMovim = 'FVEN'"
   FIN& = CantRegistros("CAJABANC", CONDI$)
   U& = 0
   J& = 0
   '
   SQLX$ = "SELECT * FROM CAJABANC WITH(NOLOCK)"
   SQLX$ = SQLX$ + "WHERE TipoMovim = 'FVEN' "
   SQLX$ = SQLX$ + " ORDER BY CODICOM_LAR"
   Dim CABATEMP As ADODB.Recordset
   Set CABATEMP =  FLEXCONN.EXECUTE (FILTSQL$ (SQLX$))
   With CABATEMP
      Do While Not .EOF
          U& = U& + 1
          Call TRACE(20, U&, FIN&)
          NUFACTU$ = SAFETEXT$(!CodiCom_Lar)
          NC% = XVALO(!NUCLIEN)
          FEEMISI% = CVINT(!FechEmisi)
          FEMIS = FECHATEX$(FEEMISI%)
          CONDI$ = "CODICOM_LAR= '" & NUFACTU$ & "' AND TipoMovim = 'FVEN'"
          CANTFACT% = CantRegistros("CAJABANC", CONDI$)
          If CANTFACT% > 1 Then
             J& = J& + 1
             TX$ = "Factura Número: " & NUFACTU$ & " -Cta: " & NC% & " - " & RASOCLI$(NC%) & " -" & FEMIS
             Print #NARCHI2%, TX$
          End If
      .MoveNext
      Loop
   End With
   '
   CABATEMP.Close
   Set CABATEMP = Nothing
   Close #NARCHI2%
   '
   If J& > 0 Then
     Call COMUNI("Se Detectaron: " & J& & " Repeticiones")
   Else
     Call COMUNI("No Se Detectaron Repeticiones")
   End If
   '
   Screen.MousePointer = 1
   Command6.Enabled = True
   '
End Sub



Private Sub Command7_Click()
   '
   'PROCESO DE VERIFICACION QUE COINCIDA VENDEDOR DEL PEDIDO CON EL VENDEDOR FACTURADO
   '
   Command7.Enabled = False
   Screen.MousePointer = 11
   '
   Call ABRECABAN
   Call ABRESTOCKS
   '
   NARCHI2% = FILEOPEN%("C:\Flexoft\" + "Vend_vs_Vend.txt", 2, 128) ' ABRE EL ARCHIVO DE ESCRITURA
   CONDI$ = "TipoMovim = 'FVEN'"
   FIN& = CantRegistros("CAJABANC", CONDI$)
   U& = 0
   J& = 0
   '
   SQLX$ = "SELECT * FROM CAJABANC WITH(NOLOCK)"
   SQLX$ = SQLX$ + "WHERE TipoMovim = 'FVEN' "
   SQLX$ = SQLX$ + " ORDER BY fechemisi"
   Dim CABATEMP As ADODB.Recordset
   Set CABATEMP =  FLEXCONN.EXECUTE (FILTSQL$ (SQLX$))
   With CABATEMP
      Do While Not .EOF
          U& = U& + 1
          Call TRACE(20, U&, FIN&)
          NUFACTU$ = SAFETEXT$(!CodiCom_Lar)
          VENDFACTU% = XVALO(!Vend_Cli)
          FEEMISI% = CVINT(!FechEmisi)
          FEMIS = FECHATEX$(FEEMISI%)
          '
          CONDI$ = "DOSECLAR= '" & NUFACTU$ & "' AND NUPEDCLI > 0 "
          NPEDMOVISTO& = XVALO(VALOBADA("MOVISTO", "NUPEDCLI", CONDI$, "NOMESS"))
          '
          If NPEDMOVISTO& > 0 Then
            CONDI$ = "NROPED = " & NPEDMOVISTO& & ""
            VENDPED& = XVALO(VALOBADA("PEDENCA", "VENDED", CONDI$, "NOMESS"))
            If VENDPED& > 0 Then
                If VENDFACTU% <> VENDPED& Then
                  J& = J& + 1
                  TX$ = "Factura: " & NUFACTU$ & " -Vend.Fact.: " & VENDFACTU% & " -  Vend.Ped.: " & VENDPED& & " - " & FEMIS
                  Print #NARCHI2%, TX$
                End If
            End If
          End If
      .MoveNext
      Loop
   End With
   '
   CABATEMP.Close
   Set CABATEMP = Nothing
   Close #NARCHI2%
   '
   If J& > 0 Then
     Call COMUNI("Se Detectaron: " & J& & " Inconsistencias")
   Else
     Call COMUNI("No Se Detectaron Inconsistencias")
   End If
   '
   Screen.MousePointer = 1
   Command7.Enabled = True

End Sub


Private Sub Form_Load()
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    Call ABRECONEXION

End Sub
