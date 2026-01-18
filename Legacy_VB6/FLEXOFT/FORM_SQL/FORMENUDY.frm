VERSION 5.00
Begin VB.Form FORMENUDY 
   BackColor       =   &H00C0D0C0&
   Caption         =   "Menues Dinámicos"
   ClientHeight    =   5490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   11820
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   2730
      TabIndex        =   11
      Top             =   1365
      Width           =   15
   End
   Begin VB.Frame Frame3 
      Caption         =   "Frame3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   8715
      TabIndex        =   12
      Top             =   1320
      Width           =   15
   End
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
      Height          =   2220
      Left            =   105
      TabIndex        =   30
      Top             =   1680
      Width           =   10725
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0D0C0&
      Height          =   1120
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Width           =   10695
      Begin VB.TextBox Text3 
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
         Left            =   315
         TabIndex        =   29
         Top             =   700
         Width           =   3135
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00C0D0C0&
         BorderStyle     =   0  'None
         Height          =   435
         Left            =   1785
         ScaleHeight     =   435
         ScaleWidth      =   5895
         TabIndex        =   25
         Top             =   210
         Width           =   5895
         Begin VB.TextBox Text2 
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
            Left            =   1260
            TabIndex        =   27
            Top             =   105
            Width           =   375
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Usuario:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   380
            TabIndex        =   28
            Top             =   130
            Width           =   795
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Clave Padre:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4305
            TabIndex        =   26
            Top             =   130
            Width           =   1455
         End
         Begin VB.Line Line1 
            X1              =   5880
            X2              =   5880
            Y1              =   105
            Y2              =   400
         End
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2625
         TabIndex        =   24
         TabStop         =   0   'False
         Text            =   "Combo1"
         Top             =   315
         Width           =   7785
      End
      Begin VB.TextBox Text4 
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
         Left            =   4095
         TabIndex        =   21
         Top             =   700
         Width           =   6315
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00400040&
      Height          =   7860
      Left            =   10950
      ScaleHeight     =   7800
      ScaleWidth      =   945
      TabIndex        =   13
      Top             =   0
      Width           =   1000
      Begin VB.CommandButton Command12 
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
         Height          =   580
         Left            =   105
         Picture         =   "FORMENUDY.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Cierra y Abandona la Aplicación"
         Top             =   105
         Width           =   650
      End
      Begin VB.CommandButton Command6 
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
         Height          =   590
         Left            =   105
         Picture         =   "FORMENUDY.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   685
         Width           =   650
      End
      Begin VB.CommandButton Command5 
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
         Height          =   645
         Left            =   120
         Picture         =   "FORMENUDY.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Aprueba, Graba e Imprime Pedido"
         Top             =   4680
         Width           =   645
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   14
         Top             =   3720
         Width           =   650
         Begin VB.Frame Frame4 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   15
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   -315
         Picture         =   "FORMENUDY.frx":075E
         Top             =   4080
         Width           =   1515
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0D0C0&
      Caption         =   "Edición"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   4080
      Width           =   10695
      Begin VB.CommandButton Command4 
         Caption         =   "Edit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   245
         Left            =   9375
         TabIndex        =   10
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Del"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   245
         Left            =   9375
         TabIndex        =   9
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Ins"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   245
         Left            =   9375
         TabIndex        =   8
         Top             =   480
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Add"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   245
         Left            =   9375
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox Text8 
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
         Left            =   1425
         TabIndex        =   6
         Top             =   840
         Width           =   6660
      End
      Begin VB.TextBox Text7 
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
         Left            =   5625
         TabIndex        =   5
         Top             =   360
         Width           =   2460
      End
      Begin VB.TextBox Text6 
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
         Left            =   1425
         TabIndex        =   3
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Referencia:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   300
         TabIndex        =   19
         Top             =   880
         Width           =   975
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Imagen:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4800
         TabIndex        =   4
         Top             =   400
         Width           =   735
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Clave Hijo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   400
         Width           =   975
      End
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFC0&
      Height          =   135
      Left            =   7440
      TabIndex        =   23
      Top             =   1300
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFC0&
      Height          =   135
      Left            =   2280
      TabIndex        =   22
      Top             =   1300
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"FORMENUDY.frx":2680
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   10695
   End
End
Attribute VB_Name = "FORMENUDY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub Combo1_Click()
    '
    Call ABRETACON
    CODPA$ = TRIM$(Mid$(Combo1.TEXT, 49, 32))
    USNU% = XVALO(Text2)
    '
10  SQLX$ = "SELECT * FROM USER_MENUS WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE USUARIO = " & USNU%
    SQLX$ = SQLX$ + " AND CodigoHijo = '" & CODPA$ & "'"
    'Set MENUTEMP = BaDaConsul.OpenRecordset(SQLX$, dbOpenSnapshot)
    Set MENUTEMP = READSET(SQLX$)
    '
    DESPA$ = TRIM$(Left$(Combo1.TEXT, 46))
    With MENUTEMP
      If Not (.BOF And .EOF) Then
          DESPA$ = SAFETEXT$(!Descripcion)
        Else
          If USNU% > 0 Then     ' REPITE CON EL MENU BASICO
             MENUTEMP.Close
             USNU% = 0
             GoTo 10
          End If
      End If
    End With
    MENUTEMP.Close
    Set MENUTEMP = Nothing
    Text4 = DESPA$
    '
    USNU% = XVALO(Text2)
    '
20  SQLX$ = "SELECT * FROM USER_MENUS WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE USUARIO = " & USNU%
    SQLX$ = SQLX$ + " AND CodigoPadre = '" & CODPA$ & "'"
    SQLX$ = SQLX$ + " ORDER BY NumeOrden ASC"
    'Set MENUTEMP = BaDaConsul.OpenRecordset(SQLX$, dbOpenSnapshot)
    Set MENUTEMP = READSET(SQLX$)
    '
    List1.Clear
    With MENUTEMP
      If Not (.BOF And .EOF) Then
          Do While Not .EOF
            TTX1$ = AJUSTI$(SAFETEXT$(!CodigoHijo), 24) + " " + AJUSTI$(SAFETEXT$(!Descripcion), 56) + " " + AJUSTI$(SAFETEXT$(!CodigoImagen), 16)
            List1.AddItem TTX1$
          .MoveNext
          Loop
        Else
          If USNU% > 0 Then     ' REPITE CON EL MENU BASICO
             MENUTEMP.Close
             USNU% = 0
             GoTo 20
          End If
      End If
    End With
    MENUTEMP.Close
    Set MENUTEMP = Nothing
    '
End Sub

Private Sub Command1_Click()
    TTX1$ = AJUSTI$(Text6, 24) + " " + AJUSTI$(Text8, 56) + " " + AJUSTI$(Text7, 16)
    List1.AddItem TTX1$
End Sub

Private Sub Command12_Click()
    Hide
    
End Sub

Private Sub Command2_Click()
    Command2.Enabled = False
    TTX1$ = AJUSTI$(Text6, 24) + " " + AJUSTI$(Text8, 56) + " " + AJUSTI$(Text7, 16)
    If List1.ListIndex >= 0 Then
         ILISTA& = List1.ListIndex
         List1.AddItem TTX1$, ILISTA&
      Else
         List1.AddItem TTX1$
    End If
    Command2.Enabled = True
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    If List1.ListIndex >= 0 Then
       ILISTA& = List1.ListIndex
       List1.RemoveItem ILISTA&
    End If
    If ILISTA& < List1.ListCount Then List1.ListIndex = ILISTA&
    Command3.Enabled = True
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    If List1.ListIndex >= 0 Then
       ILISTA& = List1.ListIndex
       TTX1$ = List1.TEXT
       List1.RemoveItem ILISTA&
       Text6 = TRIM$(Left$(TTX1$, 24))
       Text8 = TRIM$(Mid$(TTX1$, 26, 56))
       Text7 = TRIM$(Mid$(TTX1$, 83, 16))
       If ILISTA& < List1.ListCount Then
          List1.ListIndex = ILISTA&
       End If
    End If
    Command4.Enabled = True
End Sub

Private Sub Command5_Click()
    Command5.Enabled = False
    Call ABRETACON
    CODPA$ = TRIM$(Mid$(Combo1.TEXT, 49, 32))
    DESPA$ = Left$(TRIM$(Text4), 64)
    USNU% = XVALO(Text2)
    '
    SQLX$ = "UPDATE USER_MENUS SET Descripcion = '" + DESPA$ + "' "
    SQLX$ = SQLX$ + " WHERE USUARIO = " & USNU%
    SQLX$ = SQLX$ + " AND CodigoHijo = '" & CODPA$ & "'"
    'BaDaConsul.Execute (SQLX$)
    FLEXCONN.Execute (SQLX$)
    '
    SQLX$ = "UPDATE USER_MENUS SET NumeOrden = -1 "
    SQLX$ = SQLX$ + " WHERE USUARIO = " & USNU%
    SQLX$ = SQLX$ + " AND CodigoPadre = '" & CODPA$ & "'"
    'BaDaConsul.Execute (SQLX$)
    FLEXCONN.Execute (SQLX$)
    '
    SQLX$ = "SELECT * FROM USER_MENUS "
    SQLX$ = SQLX$ + " WHERE USUARIO = " & USNU%
    SQLX$ = SQLX$ + " AND CodigoPadre = '" & CODPA$ & "'"
    Dim MENUTEMP As New ADODB.Recordset
    MENUTEMP.Open SQLX$, FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
'    Set MENUTEMP = BaDaConsul.OpenRecordset(SQLX$, dbOpenDynaset)
    '
    NORIT% = 0
    For U& = 1 To List1.ListCount
       CLAHIJ$ = TRIM$(Left$(List1.List(U& - 1), 24))
       DESHIJ$ = TRIM$(Mid$(List1.List(U& - 1), 26, 56))
       IMAGXX$ = TRIM$(Mid$(List1.List(U& - 1), 83, 16))
       With MENUTEMP
          .AddNew   ' YA CONVERTIDO
          On Error Resume Next
          !CODIEMPR = CodiEmp%
          On Error GoTo 0
          !USUARIO = USNU%
          !CodigoPadre = CODPA$
          !CodigoHijo = CLAHIJ$
          !Descripcion = DESHIJ$
          !CodigoImagen = IMAGXX$
             NORIT% = NORIT% + 1
          !NumeOrden = NORIT%
          .Update
       End With
    Next U&
    MENUTEMP.Close
    '
    SQLX$ = "DELETE FROM USER_MENUS "
    SQLX$ = SQLX$ + " WHERE USUARIO = " & USNU%
    SQLX$ = SQLX$ + " AND NumeOrden = -1"
    'BaDaConsul.Execute (SQLX$)
    FLEXCONN.Execute (SQLX$)
    '
    Command5.Enabled = True
End Sub

Private Sub Command7_Click()
    Dim menuflex As Database
    Set menuflex = OpenDatabase("X:\BICI-SQL\FLEXOFT\WINDOWS\MENUFLEX.MDB")
    '
    USNU% = 0
    Call ABRETACON
    SQLX$ = "SELECT * FROM USER_MENUS "
    SQLX$ = SQLX$ + " WHERE USUARIO = " & USNU%
    Dim MENUTEMP As New ADODB.Recordset
    MENUTEMP.Open SQLX$, FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
'    Set MENUTEMP = BaDaConsul.OpenRecordset(SQLX$, dbOpenDynaset)
    '
    For U& = 1 To menuflex.TableDefs.Count
       AAA$ = menuflex.TableDefs(U& - 1).Name
       If Left$(UCase$(AAA$), 4) <> "MSYS" Then
          SQLX$ = "SELECT * FROM " + AAA$
          Set rs = menuflex.OpenRecordset(SQLX$)
          With rs
            If Not (.BOF And .EOF) Then
               NORIT% = 0
               Do While Not .EOF
                 CODPA$ = REPLACAR$(SAFETEXT$(!CodigoPadre), "\", "/")
                 CLAHIJ$ = REPLACAR$(!CodigoHijo, "\", "/")
                 DESHIJ$ = !Descripcion
                 IMAGXX$ = ""
                 With MENUTEMP
                    .AddNew   ' YA CONVERTIDO
                    On Error Resume Next
                    !CODIEMPR = CodiEmp%
                    On Error GoTo 0
                    !USUARIO = USNU%
                    !CodigoPadre = CODPA$
                    !CodigoHijo = CLAHIJ$
                    !Descripcion = DESHIJ$
                    !CodigoImagen = IMAGXX$
                       NORIT% = NORIT% + 1
                    !NumeOrden = NORIT%
                    .Update
                 End With
               .MoveNext
               Loop
            End If
          End With
       End If
    Next U&
    MENUTEMP.Close
    '
    Call COMUNI("Fin de proceso")
    '
End Sub

Private Sub List1_DblClick()
    If List1.ListIndex >= 0 Then
       ILISTA& = List1.ListIndex
       TTX1$ = List1.TEXT
       TTX2$ = AJUSTI$(Mid$(TTX1$, 26, 46), 46) + "  " + AJUSTI$(Left$(TTX1$, 24), 24)
       Combo1.TEXT = TTX2$
       Call Combo1_Click
    End If
End Sub
