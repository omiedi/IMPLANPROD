VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form DEFDESC07 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E2D3C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Definición de Descuentos Combinados"
   ClientHeight    =   6420
   ClientLeft      =   165
   ClientTop       =   5535
   ClientWidth     =   5610
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   5610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00624E28&
      ForeColor       =   &H80000008&
      Height          =   6840
      Left            =   4720
      ScaleHeight     =   6810
      ScaleWidth      =   1350
      TabIndex        =   26
      Top             =   0
      Width           =   1380
      Begin VB.CommandButton Command3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         Picture         =   "DEFDESC07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Cierra y Abandona la Aplicación"
         Top             =   240
         Width           =   640
      End
      Begin VB.CommandButton Command8 
         Caption         =   "List"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         Picture         =   "DEFDESC07.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Listados de Descuentos"
         Top             =   930
         Visible         =   0   'False
         Width           =   640
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Check"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         Picture         =   "DEFDESC07.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Verificación por Cliente y Producto"
         Top             =   2310
         Width           =   640
      End
      Begin VB.CommandButton AGenRep 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         Picture         =   "DEFDESC07.frx":075E
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Acceso a Consultas Pre-Configuradas"
         Top             =   1620
         Width           =   650
      End
   End
   Begin VB.PictureBox Picture999 
      Height          =   225
      Index           =   0
      Left            =   1470
      ScaleHeight     =   165
      ScaleWidth      =   1110
      TabIndex        =   24
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E2D3C0&
      Caption         =   "BONIFICACION PRECIO / PRODUCTO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   90
      TabIndex        =   21
      Top             =   5640
      Width           =   4500
      Begin VB.OptionButton Option2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2D3C0&
         Caption         =   "Por Prod.Bonifcado"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2250
         TabIndex        =   23
         Top             =   280
         Width           =   2055
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2D3C0&
         Caption         =   "Por Precio"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   22
         Top             =   280
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   660
      TabIndex        =   20
      Top             =   5250
      Width           =   2025
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Registrar"
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
      Left            =   2865
      TabIndex        =   19
      Top             =   5250
      Width           =   1275
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2D3C0&
      Caption         =   "POR CLIENTES"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   2325
      Left            =   90
      TabIndex        =   0
      Top             =   120
      Width           =   4500
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
         Left            =   4095
         Picture         =   "DEFDESC07.frx":0A68
         TabIndex        =   10
         Top             =   1425
         Width           =   175
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
         Left            =   4100
         Picture         =   "DEFDESC07.frx":0D72
         TabIndex        =   9
         Top             =   375
         Width           =   175
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00EEEEDD&
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo de Clientes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   735
         TabIndex        =   16
         Top             =   1520
         Width           =   2430
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Left            =   210
         TabIndex        =   12
         Top             =   1845
         Width           =   4110
      End
      Begin VB.Label Label1 
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
         Left            =   3255
         TabIndex        =   11
         Top             =   1425
         Width           =   855
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00EEEEDD&
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   735
         TabIndex        =   3
         Top             =   435
         Width           =   2430
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
         Left            =   3255
         TabIndex        =   2
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Left            =   210
         TabIndex        =   1
         Top             =   795
         Width           =   4110
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E2D3C0&
      Caption         =   "POR PRODUCTOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2325
      Left            =   90
      TabIndex        =   4
      Top             =   2640
      Width           =   4500
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
         Height          =   285
         Left            =   4095
         Picture         =   "DEFDESC07.frx":107C
         TabIndex        =   13
         Top             =   1440
         Width           =   175
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
         Height          =   285
         Left            =   4095
         Picture         =   "DEFDESC07.frx":1386
         TabIndex        =   8
         Top             =   390
         Width           =   175
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00EEEEDD&
         BackStyle       =   0  'Transparent
         Caption         =   "LÍnea"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   840
         TabIndex        =   17
         Top             =   1515
         Width           =   2430
      End
      Begin VB.Label Label4 
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
         Left            =   3360
         TabIndex        =   15
         Top             =   1440
         Width           =   750
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Left            =   210
         TabIndex        =   14
         Top             =   1860
         Width           =   4110
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Left            =   210
         TabIndex        =   7
         Top             =   810
         Width           =   4110
      End
      Begin VB.Label Label12 
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
         Left            =   2100
         TabIndex        =   6
         Top             =   390
         Width           =   2010
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00EEEEDD&
         BackStyle       =   0  'Transparent
         Caption         =   "Código"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -420
         TabIndex        =   5
         Top             =   480
         Width           =   2430
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2625
      OleObjectBlob   =   "DEFDESC07.frx":1690
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "DEFDESC07.frx":18C4
      TabIndex        =   25
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEEEDD&
      BackStyle       =   0  'Transparent
      Caption         =   "Valor"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   -1335
      TabIndex        =   18
      Top             =   5325
      Width           =   1800
   End
End
Attribute VB_Name = "DEFDESC07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AGenRep_Click()
    '
    AGenRep.Enabled = False
    Call ACONSUL_PRE(TRIM$(UCase$(Name)))
    AGenRep.Enabled = True
    '
End Sub

Private Sub AGenRep_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
   '
   If Button = 2 Then
      Call CONECTA_CONSUL_PRE(TRIM$(UCase$(Name)))
   End If
   '
End Sub

Private Sub Command1_Click()
   'Call CommSel("MASTER", Command1, Label12)
   Call SELECHO("MASTER")
   Label12 = VALACT1$("MASTER")

End Sub


Private Sub Command2_Click()
   Call CommSel("Grucocli", Command2, Label1)

End Sub

Private Sub Command3_Click()
   Call CIERRARCH("*.*")
   Hide
End Sub

Private Sub Command4_Click()
   Call CommSel("grucoven", Command4, Label4)

End Sub
Private Sub Command6_Click()
    Command6.Enabled = False
    CHECKDESCU.Show 1
    Command6.Enabled = True
End Sub

Private Sub Command5_Click()
   'graba el descuento asociado
   '
   Command5.Enabled = False
   
   Call ABREPEDCLI
   ABUS$ = ""
   NC% = XVALO(Label11)
   GRUP% = XVALO(Label1)
   CODI% = CODINT(TRIM$(Label12))
   LNEA% = XVALO(Label4)
   PORDESC$ = TRIM$(Text1)
'   If PORDESC$ = "" Then
'      COMUNI ("Falta Ingresar Descuento")
'      GoTo 99
'   End If

    
'   ABUS$ = MKI$(NC%) + MKI$(GRUP%) + MKI$(CODI%) + MKI$(LNEA%)
'   Call REPLA(RECORD$, ABUS$, 1, 8)
'   Call REPLA(RECORD$, MKI(NC%), 17, 2)
'   Call REPLA(RECORD$, MKI(GRUP%), 19, 2)
'   Call REPLA(RECORD$, MKI(CODI%), 21, 2)
'   Call REPLA(RECORD$, MKI(LNEA%), 23, 2)
'   '
'   Call REPLA(RECORD$, MKI(NC%), 33, 2)
'   Call REPLA(RECORD$, MKI(GRUP%), 35, 2)
'   Call REPLA(RECORD$, PORDESC$, 49, 12)
'   '
'   Call FILTERPUTRECORD("DESCOMBI", 1, ABUS$, RECORD$)
'   Call CIERRARCH("DESCOMBI")
     
   Condi$ = "Nume_Cli = " & NC%
   Condi$ = Condi$ & " AND GrCo_Cli =" & GRUP%
   Condi$ = Condi$ & " AND Cod_Inte =" & CODI%
   Condi$ = Condi$ & " AND Grucoven =" & LNEA%
   'CONDI$ = CONDI$ & " AND Pordescu ='" & PORDESC$ & "'"
   If TRIM$(Text1.TEXT) = "" Then
      Call BORRAREGISTROS("DESCOMBI", Condi$, REG&, "NOMESS")
      Call COMUNI("Descuento Combinado Eliminado\Para la Combinacion Elegida")
      Text1.TEXT = ""
      GoTo 99
   End If
   '
   If CantRegistros("DESCOMBI", Condi$) < 1 Then
      SQLX$ = "INSERT INTO DESCOMBI (NUME_CLI, GRCO_CLI, COD_INTE, GRUCOVEN) VALUES ("
      SQLX$ = SQLX$ & NC% & ", " & GRUP% & ", " & CODI% & ", " & LNEA% & ")"
       FLEXCONN.Execute (FILTSQL$(SQLX$))
   End If
   Call ACTUREGISTRO("DESCOMBI", "Pordescu", Condi$, PORDESC$, REG&)
   
   'CONTROL DE DESCUENTO POR PRECIO
   VALO% = 0
   If Option2.Value = True Then VALO% = 1
   Call ACTUREGISTRO("DESCOMBI", "Modo_Desc", Condi$, VALO%, REG&)
    
99  Command5.Enabled = True
    
End Sub


Private Sub Command8_Click()
   Call COPYSTRU("Descombi", "Ldescomb")
   '
10 OPX% = ALTERNA%("Por Cliente\Por Grupos de Clientes\Por Productos\Por Linea de Productos")
   If OPX% < 1 Or OPX% > 4 Then Exit Sub
   '
   U& = 0
   '
   FIN& = ULTREG&("Descombi")

   If OPX% = 1 Then 'POR CLIENTE
      For i& = 1 To FIN&
        x$ = REGLEIDO$("Descombi", i&)
        If CVI(Mid$(x$, 17, 2)) > 0 Then
        AA = CVI(Mid$(x$, 17, 2))
          U& = U& + 1
          Call GRAREG("Ldescomb", x$, U&)
        End If
      Next i&
      Call SETULTREG("LDESCOMB", U&)
      Call CIERRARCH("Ldescomb")

      
      Call FILESORT("Ldescomb", "", 3, 10, "ASC")
      Call FINAL("*LIDESCLI")
     '
     ElseIf OPX% = 2 Then 'POR GRUPO CLIENTE
      For i& = 1 To FIN&
        x$ = REGLEIDO$("Descombi", i&)
        If CVI(Mid$(x$, 19, 2)) > 0 Then
          U& = U& + 1
          Call GRAREG("Ldescomb", x$, U&)
        End If
      Next i&
      Call SETULTREG("LDESCOMB", U&)
      Call CIERRARCH("Ldescomb")

      Call FILESORT("Ldescomb", "", 5, 13, "ASC")
      Call FINAL("*LIDESGRU")

     '
     ElseIf OPX% = 3 Then 'POR PRODUCTO
      For i& = 1 To FIN&
        x$ = REGLEIDO$("Descombi", i&)
        If CVI(Mid$(x$, 21, 2)) > 0 Then
          U& = U& + 1
          Call GRAREG("Ldescomb", x$, U&)
        End If
      Next i&
      Call SETULTREG("LDESCOMB", U&)
      Call CIERRARCH("Ldescomb")
      
      Call FILESORT("Ldescomb", "", 7, 15, "ASC")
      Call FINAL("*LIDESART")

     ElseIf OPX% = 4 Then 'POR LINEA
      For i& = 1 To FIN&
        x$ = REGLEIDO$("Descombi", i&)
        If CVI(Mid$(x$, 23, 2)) > 0 Then
          U& = U& + 1
          Call GRAREG("Ldescomb", x$, U&)
        End If
      Next i&
      Call SETULTREG("LDESCOMB", U&)
      Call CIERRARCH("Ldescomb")

      Call FILESORT("Ldescomb", "", 9, 15, "ASC")
      Call FINAL("*LIDESLIN")

   End If
   '
   GoTo 10
   '
End Sub

Private Sub Command9_Click()
   Call CommSel("DEUDOR1", Command9, Label11)
End Sub

Private Sub Form_Load()
  'Call APLICACIONSKINS(Me, Picture1)
  Call ABRECONEXION
  Call label1_Change
  Call Label11_Change
  Call Label12_Change
  Call Label4_Change
  Call BUSCAVALOR
  '
  
End Sub

Private Sub label1_Change()
   NCX% = XVALO(Label1)
   If NCX% > 0 Then
       Label2 = ECOARCH$("GRUCOCLI", MKI$(NCX%))
     Else
       Label2 = "Cualquier Grupo de Clientes"
   End If
   If Label1 <> "" Then
      Label11 = ""
      Label8 = ""
   End If

   Call BUSCAVALOR
End Sub

Private Sub Label10_Click()
   Label12 = ""
End Sub

Private Sub Label11_Change()
   NCX% = XVALO(Label11)
   If NCX% > 0 Then
       Label8 = rasocli$(NCX%)
     Else
       Label8 = "Cualquier Cliente"
   End If
   '
   If Label11 <> "" Then
      Label1 = ""
   End If

   Call BUSCAVALOR

End Sub

Private Sub Label12_Change()
   CI% = CODINT%(Label12)
   If CI% > 0 Then
       Label9 = DESCRIT0$(CI%)
     Else
       Label9 = "Todos los Articulos"
   End If
   '
   If Label12 <> "" Then
      Label4 = ""
      Label3 = ""
   End If

   Call BUSCAVALOR

End Sub

Private Sub Label17_Click()

End Sub

Private Sub Label16_Click()

End Sub

Private Sub Label2_Change()
   If Label8 = "" Then Label8.Caption = "Cualquier Grupo de Clientes"
End Sub

Private Sub Label3_Change()
   If Label3 = "" Then Label3.Caption = "Cualquier Linea de Productos"

End Sub

Private Sub Label4_Change()
   NCX% = XVALO(Label4)
   If NCX% > 0 Then
       Label3 = ECOARCH$("GRUCOVEN", MKI$(NCX%))
     Else
       Label3 = "Cualquier Línea"
   End If
   If Label4 <> "" Then
      Label12 = ""
      Label9 = ""
   End If
   Call BUSCAVALOR
End Sub

Private Sub Label5_DblClick()
   Label11 = ""
End Sub

Private Sub Label6_Click()
  Label1 = ""
End Sub

Private Sub Label7_Click()
   Label4 = ""
End Sub

Private Sub Label8_Change()
   If Label8 = "" Then Label8.Caption = "Cualquier Cliente"
End Sub

Private Sub Label9_Change()
   If Label9 = "" Then Label9.Caption = "Cualquier Producto"
End Sub

Sub BUSCAVALOR()
   ABUS$ = ""
   NC = XVALO(Label11)
   GRUP% = XVALO(Label1)
   CODI% = CODINT(Label12)
   LNEA% = XVALO(Label4)
   '
'   ABUS$ = MKI$(NC%) + MKI$(GRUP%) + MKI$(CODI%) + MKI$(LNEA%)
'   X$ = FILTERGETRECORD$("DESCOMBI", 1, (ABUS$))
'   Text1 = TRIM$(Mid$(X$, 49, 12))
   '
   
   
   Condi$ = "Nume_Cli = " & NC
   Condi$ = Condi$ & " AND GrCo_Cli =" & GRUP%
   Condi$ = Condi$ & " AND Cod_Inte =" & CODI%
   Condi$ = Condi$ & " AND Grucoven =" & LNEA%
   valor$ = ""
   Option1.Value = True
   If CantRegistros("DESCOMBI", Condi$) > 0 Then
      valor$ = SAFETEXT(VALOBADA("DESCOMBI", "Pordescu", Condi$))
      MODODESC% = XVALO(VALOBADA("DESCOMBI", "MODO_DESC", Condi$))
   End If
   '
   'ASIGNO VALORES
   Text1 = valor$
   If MODODESC% > 0 Then Option2.Value = True
   
End Sub

