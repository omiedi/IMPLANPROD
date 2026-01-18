VERSION 5.00
Begin VB.Form FRACLOTIN07 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Apertura de Lotes de Recepción"
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7305
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   7305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00004000&
      Height          =   3500
      Left            =   6405
      ScaleHeight     =   3435
      ScaleWidth      =   1845
      TabIndex        =   10
      Top             =   0
      Width           =   1905
      Begin VB.CommandButton Command2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   610
         Left            =   105
         Picture         =   "FRACLOTIN07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Graba - Guardar Modificaciones"
         Top             =   2100
         Width           =   645
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
         Height          =   610
         Left            =   105
         Picture         =   "FRACLOTIN07.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   315
         Width           =   645
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "FRACLOTIN07.frx":0454
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   1470
         Width           =   1515
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Detalles del Lote"
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
      Left            =   210
      TabIndex        =   0
      Top             =   210
      Width           =   6105
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   5040
         TabIndex        =   20
         Text            =   "Text2"
         Top             =   2205
         Width           =   150
      End
      Begin VB.TextBox Text1 
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
         Height          =   285
         Left            =   4725
         TabIndex        =   9
         Top             =   2055
         Width           =   1170
      End
      Begin VB.Label Label24 
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
         Height          =   285
         Left            =   4725
         TabIndex        =   19
         Top             =   375
         Width           =   1170
      End
      Begin VB.Label Label14 
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
         Height          =   300
         Left            =   3255
         TabIndex        =   18
         Top             =   435
         Width           =   1335
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad:"
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
         Left            =   3255
         TabIndex        =   16
         Top             =   2115
         Width           =   1335
      End
      Begin VB.Label Label7 
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
         Height          =   285
         Left            =   1050
         TabIndex        =   15
         ToolTipText     =   "Id de Lote"
         Top             =   2055
         Width           =   1755
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cant.Origen:"
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
         Left            =   3465
         TabIndex        =   14
         Top             =   855
         Width           =   1125
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
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
         Height          =   285
         Left            =   4725
         TabIndex        =   13
         ToolTipText     =   "Cantidad Real"
         Top             =   795
         Width           =   1170
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Sub-Lote:"
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
         Left            =   -420
         TabIndex        =   8
         Top             =   2055
         Width           =   1335
      End
      Begin VB.Label Label2 
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
         Height          =   285
         Left            =   1050
         TabIndex        =   7
         ToolTipText     =   "Id de Lote"
         Top             =   1635
         Width           =   1755
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Lote:"
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
         Left            =   -420
         TabIndex        =   6
         Top             =   1635
         Width           =   1335
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
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
         Height          =   285
         Left            =   4725
         TabIndex        =   5
         ToolTipText     =   "Cantidad nominal"
         Top             =   1635
         Width           =   1170
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo:"
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
         Left            =   3255
         TabIndex        =   4
         Top             =   1695
         Width           =   1335
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
         Height          =   285
         Left            =   1050
         TabIndex        =   3
         ToolTipText     =   "Código"
         Top             =   795
         Width           =   2175
      End
      Begin VB.Label Label6 
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
         Height          =   285
         Left            =   1050
         TabIndex        =   2
         ToolTipText     =   "Descripción"
         Top             =   1215
         Width           =   4845
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
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
         Left            =   -420
         TabIndex        =   1
         Top             =   855
         Width           =   1335
      End
   End
   Begin VB.Label Label9 
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
      Height          =   285
      Left            =   5040
      TabIndex        =   17
      ToolTipText     =   "Id de Lote"
      Top             =   0
      Visible         =   0   'False
      Width           =   1230
   End
End
Attribute VB_Name = "FRACLOTIN07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
  '
  '\\\OT-07-0066-RG-21/02/07///
  Call CIERRARCH("*.*")
  Hide
  '\\\OT-07-0066-RG-FIN///
  '
End Sub

Private Sub Command2_Click()
  '
  '\\\OT-07-0066-RG-21/02/07///
  If XVALO(Text1) < 1 Then
     Call MENSERR(24, "Falta Ingresar Cantidad del Sub-lote")
     Text1 = ""
     Text1.SetFocus
     Exit Sub
  End If
  '
  If XVALO(Label21.Caption) < 0 Then
     Call MENSERR(24, "Cantidad No Valida")
     Label21.Caption = Label4.Caption
     Text1 = ""
     Text1.SetFocus
     Exit Sub
  End If
  '
  Call APERBASDAT("STOCKS")
  SQLX$ = "SELECT * FROM BOLRECEP"
  SQLX$ = SQLX$ + " WHERE NUMEBORE = " & Int(XVALO(Label9))
  SQLX$ = SQLX$ + " AND IdenLote = '" & Label2 & "'"
  Set BRTEMP = Stocks.OpenRecordset(SQLX, dbOpenDynaset)
  '
  On Error Resume Next
  BRTEMP.MoveFirst
  If Err Then
     On Error GoTo 0
     Call MENSERR(24, "Imposible Fraccionar Este Lote.")
     Exit Sub
  End If
  '
  On Error GoTo 0
  CIXI% = CODINT%(Label27)
  FF% = FECHANUM(Label24)
  FINGRE = CDbl(CVDAT(FF%))
  IDREGIS$ = TRIM$(Label2)
  CANOM# = XVALO(Label21)
  '
  SQLZ$ = " SELECT * FROM MOVISTO "
  SQLZ$ = SQLZ$ + "WHERE Cod_Inte = " & CIXI% & " "
  SQLZ$ = SQLZ$ + "AND IdenLote = '" & IDREGIS$ & "' "
  SQLZ$ = SQLZ$ + "AND CDBL(FechMov) = " & FINGRE & " "
  SQLZ$ = SQLZ$ + "AND CantIngre > 0 "
  Set MOVTEMP = Stocks.OpenRecordset(SQLZ$, dbOpenDynaset)
  '
  On Error Resume Next
  MOVTEMP.MoveFirst
  If Err Then
     On Error GoTo 0
     Call MENSERR(24, "Imposible Fraccionar Este Lote.")
     Exit Sub
  End If
  '
  On Error GoTo 0
  CANFRA = XVALO(Text1)
  CANAP = BRTEMP!CanToApro
  If CANFRA > CANAP + 0.05 Then
     Call MENSERR(24, "La Cantidad del Sub-Lote Supera\la Aprobada (" + TRIM$(FORMATNUM$(CANAP, "F10.1")) + ") del Lote Original.\   \Transacción Cancelada.")
     Exit Sub
  End If
  '
  ' AGREGA REGISTRO DEL SUB-LOTE EN BOLRECEP
  Bolrecep.AddNew   ' YA CONVERTIDO
          ON ERROR RESUME NEXT
  Bolrecep!CodiEmpr = Codiemp%
          ON ERROR GOTO 0
  For KKU% = 1 To BRTEMP.Fields.Count
    Bolrecep.Fields(KKU% - 1).Value = BRTEMP.Fields(KKU% - 1).Value
  Next KKU%
  Bolrecep!IdenLote = Label7
  Bolrecep!CantNom = CANFRA
  Bolrecep!CanToApro = CANFRA
  Bolrecep!CanToRech = 0
  Bolrecep!CanToFalta = 0
  Bolrecep.Update
  '
  ' ACTUALIZA CANTIDAD DEL LOTE ORIGINAL
  BRTEMP.Edit
    CARESTA = BRTEMP!CantNom - CANFRA
  BRTEMP!CantNom = CARESTA
    CARESTA = BRTEMP!CanToApro - CANFRA
  BRTEMP!CanToApro = CARESTA
  BRTEMP.Update
  '
  ' AGREGA REGISTRO DEL SUB-LOTE EN NOVISTO
  Movisto.AddNew   ' YA CONVERTIDO
         ON ERROR RESUME NEXT
  Movisto!CodiEmpr = Codiemp%
         ON ERROR GOTO 0
  For KKU% = 1 To MOVTEMP.Fields.Count
    Movisto.Fields(KKU% - 1).Value = MOVTEMP.Fields(KKU% - 1).Value
  Next KKU%
  Movisto!IdenLote = Label7
  Movisto!CantIngre = CANFRA
  Movisto!FeReMovi = Now
  Movisto.Update
  '
  ' ACTUALIZA CANTIDAD DEL REGISTRO ORIGINAL
  MOVTEMP.Edit
  MOVTEMP!CantIngre = CANOM#
  MOVTEMP.Update
  '
  CIXI% = CODINT%(Label27)
  FF% = FECHANUM(Label24)
  IDREGIS$ = TRIM$(Label2)
  CANOM# = XVALO(Label21)
MsgBox "FALTA DETERMINAR MARCA Y LOTE DE ORIGEN  "
  Call CREALOTIN(FF%, FEFAB%, FEVEN%, CIXI%, IDREGIS$, MARCAX$, LOTORI$, CANOM#, LARGOM, ANCHOMM, DPX%, UBX$)
  '
  IDREGIS$ = Label7
  CANOM# = XVALO(Text1)
  Call CREALOTIN(FF%, FEFAB%, FEVEN%, CIXI%, IDREGIS$, MARCAX$, LOTORI$, CANOM#, LARGOM, ANCHOMM, DPX%, UBX$)
  '
  Screen.MousePointer = 1
  Call COMUNI("Transacción Completada !!!")
  Hide
  '
End Sub

Private Sub Text1_Change()
  '
  On Error Resume Next
  CANORI = XVALO(Label4.Caption)
  CALONUE = XVALO(Text1)
  Label21.Caption = FORMATNUM$(CANORI - CALONUE, "F10.1")
  '
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
  '
  '\\\OT-07-0066-RG-22/02/07///
  If KeyAscii = 13 Then
     Command2.SetFocus
  End If
  '\\\OT-07-0066-RG-FIN///
  '
End Sub

Private Sub text1_lostfocus()
   Text1 = FORMATNUM$(XVALO(Text1), "F10.1")
End Sub
