VERSION 5.00
Begin VB.Form FOCOPOP04 
   BackColor       =   &H00C0FFE0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Copia Interactiva de TMU's"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7485
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   7485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00004000&
      Height          =   5820
      Left            =   6600
      ScaleHeight     =   5760
      ScaleWidth      =   1020
      TabIndex        =   1
      Top             =   -105
      Width           =   1080
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   540
         TabIndex        =   4
         Top             =   4305
         Width           =   600
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   5
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
         Left            =   130
         Picture         =   "FOCOPOP04.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Abandona la Aplicación"
         Top             =   360
         Width           =   600
      End
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
         Height          =   600
         Left            =   130
         Picture         =   "FOCOPOP04.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Ejecuta  la Copia"
         Top             =   960
         Width           =   600
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   -305
         Picture         =   "FOCOPOP04.frx":0614
         Top             =   4575
         Width           =   1515
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0FFE0&
      Caption         =   "Rango de Seleccion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   85
      TabIndex        =   0
      Top             =   90
      Width           =   6420
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0FFE0&
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
         Height          =   2220
         Left            =   150
         TabIndex        =   15
         Top             =   2625
         Width           =   6105
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
            Height          =   315
            Left            =   3015
            TabIndex        =   19
            Top             =   525
            Width           =   175
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
            Height          =   315
            Left            =   1140
            TabIndex        =   18
            Top             =   525
            Width           =   1890
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
            Height          =   315
            Left            =   1155
            TabIndex        =   17
            Top             =   1575
            Width           =   630
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
            Height          =   315
            Left            =   1785
            TabIndex        =   16
            Top             =   1575
            Width           =   175
         End
         Begin VB.Label Label8 
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
            Height          =   315
            Left            =   1140
            TabIndex        =   23
            Top             =   945
            Width           =   4740
         End
         Begin VB.Label Label5 
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
            Height          =   225
            Left            =   -1260
            TabIndex        =   22
            Top             =   630
            Width           =   2325
         End
         Begin VB.Label Label3 
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
            Height          =   315
            Left            =   2100
            TabIndex        =   21
            Top             =   1575
            Width           =   3795
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Operación:"
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
            Left            =   -1260
            TabIndex        =   20
            Top             =   1680
            Width           =   2325
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0FFE0&
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
         Height          =   2220
         Left            =   150
         TabIndex        =   6
         Top             =   315
         Width           =   6105
         Begin VB.CommandButton Command3 
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
            Height          =   315
            Left            =   1785
            TabIndex        =   11
            Top             =   1575
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
            Height          =   315
            Left            =   1155
            TabIndex        =   14
            Top             =   1575
            Width           =   630
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
            Height          =   315
            Left            =   1140
            TabIndex        =   8
            Top             =   525
            Width           =   1890
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
            Height          =   315
            Left            =   3015
            TabIndex        =   7
            Top             =   525
            Width           =   175
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Operación:"
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
            Left            =   -1260
            TabIndex        =   13
            Top             =   1680
            Width           =   2325
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
            Height          =   315
            Left            =   2100
            TabIndex        =   12
            Top             =   1575
            Width           =   3795
         End
         Begin VB.Label Label6 
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
            Height          =   225
            Left            =   -1260
            TabIndex        =   10
            Top             =   630
            Width           =   2325
         End
         Begin VB.Label Label1 
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
            Height          =   315
            Left            =   1140
            TabIndex        =   9
            Top             =   945
            Width           =   4740
         End
      End
   End
End
Attribute VB_Name = "FOCOPOP04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    '
    Unload Me
    '
End Sub

Private Sub Command2_Click()
  '
  If TRIM$(Text10) = "" Or TRIM$(Label1) = "" Then
    Call MENSERR(24, "Código Origen no Válido o Inexistente")
    Exit Sub
  End If
  If TRIM$(Text2) = "" Or TRIM$(Label4) = "" Then
    Call MENSERR(24, "Operación Origen no Válida o Inexistente")
    Exit Sub
  End If
  '
  If TRIM$(Text3) = "" Or TRIM$(Label8) = "" Then
     Call MENSERR(24, "Código Destino no Válido o Inexistente")
     Exit Sub
  End If
  '
  If TRIM$(Text1) = "" Or TRIM$(Label3) = "" Then
    Call MENSERR(24, "Operación Destino no Válida o Inexistente")
    Exit Sub
  End If
  '
  If TRIM$(Text10) = TRIM$(Text3) Then
    If TRIM$(Text2) = TRIM$(Text1) Then
      Call MENSERR(24, "Copia Imposible !!!\Coinciden Origen y Destino de Copia.")
      Exit Sub
    End If
  End If
  '
  Call GRAOPERA
  '
End Sub
Sub GRAOPERA()
  '
  Call ABREMASTER
  Call ABRESECOP
  '
  CID1% = CODINT%(Text3.TEXT) 'CODIGO DESTINO
  OPD1% = XVALO(Text1.TEXT) ' CODIGO DE OPERACION DESTINO
  '
  CIO1% = CODINT%(Text10.TEXT) 'CODIGO ORIGEN
  OPO1% = XVALO(Text2.TEXT) ' CODIGO DE OPERACION ORIGEN
  SQLX$ = " SELECT * FROM TMUSOPER "
  SQLX$ = SQLX$ + " WHERE Cod_Inte = " & CIO1% & " "
  SQLX$ = SQLX$ + " AND Cod_Oper = " & OPO1% & " "
  Set GRAOPERATMP = Articulos.OpenRecordset(SQLX$, 4)
  With GRAOPERATMP
    On Error Resume Next
    .MoveFirst
    If Err Then
        Call COMUNI("Copia Imposible !!!\Código Origen sin Lista de TMU's.")
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0
    TISTOP = 0
    'Si CI% Destino tiene TMU's Pregunta si las BORRO
    With TmusOper
      SQLX$ = "Cod_Inte = " & CID1% & " "
      SQLX$ = SQLX$ + " AND Cod_Oper = " & OPD1% & ""
      .FindFirst SQLX$
      If .NoMatch = False Then
        RTA% = COMALTER%("El Código Destino ya tiene una Lista de TMU'S\ para la Operación Seleccionada. \Puede Eliminar la lista Pre-Existente o\ Agregar a la Lista Pre-Existente.\\Agregar a la Lista\Eliminar Lista")
        Select Case RTA%
          Case 1
            GoTo 5
          Case 2
            Call ELITMUS(CID1%, OPD1%)
            GoTo 5
          Case Else
            Call COMUNI("Copia Cancelada")
            Exit Sub
        End Select
      End If
    End With
    '
5   Do While .EOF = False
      With TmusOper
        .AddNew
          !Cod_Inte = CID1%
          !Cod_Oper = OPD1%
          !Cod_TMU = GRAOPERATMP.Fields("Cod_TMU")
          !Nume_Ord = GRAOPERATMP.Fields("Nume_Ord")
          !Cantidad = GRAOPERATMP.Fields("Cantidad")
          COTMU$ = !Cod_TMU
          CANTI1 = !Cantidad
          MSEG1# = AMASTOK04.MSTMU#(COTMU$) 'ME DEVUELVE LOS MILISEGUNDOS DEL TMU
          TISTOP = TISTOP + (CANTI1 * MSEG1# / 1000)
        .Update
      End With
      .MoveNext
    Loop
    'EDITO EL TIEMPO EN EL MAESTRO
    With Secoper
      .FindFirst "CODICONJ =" & CID1% & " and NumeOper = " & OPD1%
      If .NoMatch = False Then
        .Edit
          !MinuStan = (TISTOP / 60)
        .Update
      Else
        Call COMUNI("Imposible Asignar Tiempo Standard !!!\al Código Destino\Consulte a su Soporte de Sistema.")
      End If
    End With
    '
15  Call COMUNI("Copia Completa")
  End With
  Call Command1_Click
  '
End Sub
Sub ELITMUS(CIN%, OPE1%)
  '
  Call ABREMASTER
  Call ABRESECOP
  '
  With TmusOper
    SQLX$ = "Cod_Inte = " & CIN% & " "
    SQLX$ = SQLX$ + " AND Cod_Oper = " & OPE1% & ""
5   .FindFirst SQLX$
    If .NoMatch = False Then
      .Delete
      .FindFirst SQLX$
      GoTo 5
    End If
  End With
  '
End Sub
Private Sub Command3_Click()
  '
  Call GENECODOPE(Text10.TEXT)
  Call SELECHO("!CODOPE/Operaciones de Codigo: " + Text10.TEXT)
  CODOP1$ = VALACT1$("!CODOPE")
  If CODOP1$ <> "" Then
    Text2.TEXT = CODOP1$
  End If
  '
End Sub

Private Sub Command4_Click()
  '
  Call GENECODOPE(Text3.TEXT)
  Call SELECHO("!CODOPE/Operaciones de Codigo: " + Text3.TEXT)
  CODOP1$ = VALACT1$("!CODOPE")
  If CODOP1$ <> "" Then
    Text1.TEXT = CODOP1$
  End If
  '
End Sub
Private Sub Command5_Click()
  '
  Call SELECHO("ECOMAST")
  CODI1$ = VALACT1$("ECOMAST")
  If CODI1$ <> "" Then
    Text3.TEXT = CODI1$
  End If
  '
End Sub

Private Sub Command9_Click()
  '
  Call SELECHO("ECOMAST")
  CODI1$ = VALACT1$("ECOMAST")
  If CODI1$ <> "" Then
    Text10.TEXT = CODI1$
  End If
  '
End Sub
Private Sub Form_Load()
  '
'  XX$ = Left$(REGLEIDO$("ECOMAST", 1), 16): CI1% = Abs(CODINT%(XX$))
'  XX$ = Left$(REGLEIDO$("ECOMAST", ULTREG&("ECOMAST")), 16): CI2% = Abs(CODINT%(XX$))
'  Text10.TEXT = CODEXT$(CI1%)
'  Text3.TEXT = CODEXT$(CI2%)
  '
End Sub

Private Sub Text1_Change()
  '
  CID1% = CODINT%(Text3.TEXT)
  Label3 = DESCRIOP$(CID1%, XVALO(Text1.TEXT))
  '
End Sub

Private Sub Text10_Change()
  '
  Label1 = ECOARCH$("ECOMAST", Text10.TEXT)
  '
End Sub

Private Sub Text10_GotFocus()
  Text10.SelStart = 0
  Text10.SelLength = Len(Text10)
End Sub

Private Sub Text2_Change()
  '
  CIO1% = CODINT%(Text10.TEXT)
  Label4 = DESCRIOP$(CIO1%, XVALO(Text2.TEXT))
  '
End Sub

Private Sub Text3_Change()
  '
  Label8 = ECOARCH$("ECOMAST", Text3.TEXT)
  '
End Sub

Private Sub Text3_GotFocus()
  Text3.SelStart = 0
  Text3.SelLength = Len(Text3)
End Sub
Private Sub Text1_GotFocus()
  Text1.SelStart = 0
  Text1.SelLength = Len(Text1)
End Sub
Private Sub Text2_GotFocus()
  Text2.SelStart = 0
  Text2.SelLength = Len(Text2)
End Sub
Sub GENECODOPE(CODIGO1$)
  '
  Call SETULTREG("!CODOPE", 0)
  Call FRESHECHO("!CODOPE")
  '
  If TRIM$(CODIGO1$) = "" Then Exit Sub
  CI% = CODINT%(CODIGO1$)
  If CI% <= 0 Then Exit Sub
  '
  Call ABREMASTER
  REBLA1$ = REGBLAN$("CODOPE")
  JJ& = 0
  SQLX$ = "SELECT * FROM Secoper"
  SQLX$ = SQLX$ & " WHERE CodIConj = " & CI%
  SQLX$ = SQLX$ & " AND StatOper >0 "
  SQLX$ = SQLX$ + "ORDER BY NumeOper ASC"
  Set SecopTemp = Articulos.OpenRecordset(SQLX$, dbOpenSnapshot)
  '
  With SecopTemp
    On Error Resume Next
    .MoveFirst
    If Err Then
      On Error GoTo 0
      GoTo 99
    End If
    On Error GoTo 0
    '
    Do Until .EOF = True
      NOPY% = !NumeOper
      DOPE$ = !DescOper
      '
      Z$ = REBLA$
      Call REPLA(Z$, MKI$(NOPY%), 1, 2)
      Call REPLA(Z$, DOPE$, 3, 48)
      JJ& = JJ& + 1
      Call GRAREG("!CODOPE", Z$, JJ&)
      '
      .MoveNext
    Loop
  End With
  '
  Call SETULTREG("!CODOPE", JJ&)
  Call FRESHECHO("!CODOPE")
  '
99 Screen.MousePointer = 1
End Sub
