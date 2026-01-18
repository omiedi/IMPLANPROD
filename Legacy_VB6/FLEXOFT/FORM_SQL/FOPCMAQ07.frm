VERSION 5.00
Begin VB.Form FOPCMAQ07 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pedido de Cotización"
   ClientHeight    =   8325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8325
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Refrescar 
      BackColor       =   &H00D0F0F0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   8590
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12000
      Begin VB.CommandButton Command14 
         BackColor       =   &H00E0E0E0&
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
         Height          =   950
         Left            =   11025
         Picture         =   "FOPCMAQ07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Aprueba e Imprime Pedido de Cotización"
         Top             =   7245
         Width           =   750
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00D0F0F0&
         Caption         =   "Refrescar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   6480
         TabIndex        =   32
         Top             =   1785
         Width           =   4280
         Begin VB.CommandButton Command4 
            Caption         =   "Selecc. SIM"
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
            Left            =   2280
            TabIndex        =   34
            Top             =   240
            Width           =   1815
         End
         Begin VB.CommandButton Command1 
            Caption         =   "SIM's Pend."
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
            Left            =   160
            TabIndex        =   33
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00D0F0F0&
         Caption         =   "Identificación"
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
         Height          =   1560
         Left            =   6480
         TabIndex        =   23
         Top             =   120
         Width           =   4280
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
            Left            =   2085
            TabIndex        =   29
            Text            =   " "
            Top             =   750
            Width           =   2010
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
            Left            =   3900
            TabIndex        =   25
            Top             =   345
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
            Left            =   2955
            TabIndex        =   24
            TabStop         =   0   'False
            Text            =   " "
            Top             =   345
            Width           =   960
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
            Left            =   2085
            TabIndex        =   41
            Top             =   1110
            Width           =   2010
         End
         Begin VB.Label Label6 
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
            Left            =   945
            TabIndex        =   36
            Top             =   345
            Width           =   1065
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Solicitante:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404000&
            Height          =   330
            Left            =   600
            TabIndex        =   31
            Top             =   1155
            Width           =   1230
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Destino:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404000&
            Height          =   300
            Left            =   360
            TabIndex        =   30
            Top             =   825
            Width           =   1485
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
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404000&
            Height          =   330
            Left            =   2100
            TabIndex        =   27
            Top             =   375
            Width           =   750
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Nº P-C:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404000&
            Height          =   300
            Left            =   -630
            TabIndex        =   26
            Top             =   375
            Width           =   1485
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00D0F0F0&
         Caption         =   "Invitados"
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
         Height          =   2295
         Left            =   2100
         TabIndex        =   20
         Top             =   105
         Width           =   4140
         Begin VB.ListBox List4 
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1500
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   21
            ToolTipText     =   "Doble-Click para Ingresar Proveedores Invitados"
            Top             =   630
            Width           =   3900
         End
         Begin VB.Label Label92 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Pr.            Razon Social       "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404000&
            Height          =   300
            Left            =   120
            TabIndex        =   22
            Top             =   360
            Width           =   3900
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00D0F0F0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   2999
         Left            =   160
         TabIndex        =   10
         Top             =   2880
         Width           =   10580
         Begin VB.Frame Frame1 
            Caption         =   "Frame7"
            Height          =   4530
            Left            =   1080
            TabIndex        =   35
            Top             =   240
            Width           =   15
         End
         Begin VB.Frame Frame6 
            Caption         =   "Frame18"
            Height          =   4530
            Left            =   9240
            TabIndex        =   19
            Top             =   315
            Width           =   15
         End
         Begin VB.Frame Frame3 
            Caption         =   "Frame18"
            Height          =   4530
            Left            =   8160
            TabIndex        =   18
            Top             =   240
            Width           =   15
         End
         Begin VB.Frame Frame7 
            Caption         =   "Frame7"
            Height          =   4530
            Left            =   2880
            TabIndex        =   12
            Top             =   315
            Width           =   15
         End
         Begin VB.Frame Frame19 
            Caption         =   "Frame19"
            Height          =   4530
            Left            =   10290
            TabIndex        =   11
            Top             =   315
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
            Height          =   2370
            ItemData        =   "FOPCMAQ07.frx":030A
            Left            =   0
            List            =   "FOPCMAQ07.frx":030C
            Style           =   1  'Checkbox
            TabIndex        =   15
            Top             =   600
            Width           =   10560
         End
         Begin VB.PictureBox Picture3 
            BackColor       =   &H00C0C0C0&
            Height          =   435
            Left            =   0
            ScaleHeight     =   375
            ScaleWidth      =   10500
            TabIndex        =   13
            Top             =   315
            Width           =   10565
            Begin VB.Label Label22 
               BackStyle       =   0  'Transparent
               Caption         =   $"FOPCMAQ07.frx":030E
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404000&
               Height          =   225
               Left            =   0
               TabIndex        =   14
               Top             =   0
               Width           =   10515
            End
         End
         Begin VB.Line Line4 
            X1              =   10560
            X2              =   10560
            Y1              =   360
            Y2              =   5640
         End
      End
      Begin VB.CommandButton Command10 
         Caption         =   "!"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1680
         TabIndex        =   16
         ToolTipText     =   "Lista de Proveedores con Entregas Programadas"
         Top             =   9520
         Width           =   195
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00D0F0F0&
         Caption         =   "Condiciones / Observaciones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   2200
         Left            =   160
         TabIndex        =   8
         Top             =   5985
         Width           =   10590
         Begin VB.TextBox Text13 
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1710
            Left            =   320
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   9
            ToolTipText     =   "Click para Ingresar las Condiciones / Observaciones"
            Top             =   330
            Width           =   9960
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00004000&
         Height          =   7155
         Left            =   11040
         ScaleHeight     =   7095
         ScaleWidth      =   1005
         TabIndex        =   1
         Top             =   0
         Width           =   1065
         Begin VB.CommandButton Command12 
            BackColor       =   &H00E0E0E0&
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
            Picture         =   "FOPCMAQ07.frx":03C7
            Style           =   1  'Graphical
            TabIndex        =   42
            ToolTipText     =   "Ver por Pantalla y Re-Imprimir"
            Top             =   5400
            Width           =   650
         End
         Begin VB.CommandButton Command7 
            BackColor       =   &H00E0E0E0&
            Height          =   650
            Left            =   105
            Picture         =   "FOPCMAQ07.frx":06D1
            Style           =   1  'Graphical
            TabIndex        =   40
            ToolTipText     =   "Listado de Sollicitudes por  Rango de Fechas"
            Top             =   4620
            Width           =   650
         End
         Begin VB.CommandButton Command6 
            BackColor       =   &H00E0E0E0&
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
            Height          =   690
            Left            =   120
            Picture         =   "FOPCMAQ07.frx":09DB
            Style           =   1  'Graphical
            TabIndex        =   39
            ToolTipText     =   "Modificacion de la SIM"
            Top             =   3840
            Width           =   650
         End
         Begin VB.CommandButton Command8 
            BackColor       =   &H00E0E0E0&
            Height          =   640
            Left            =   105
            Picture         =   "FOPCMAQ07.frx":0CE5
            Style           =   1  'Graphical
            TabIndex        =   37
            ToolTipText     =   "Acceso a Maestro de Items de Stock"
            Top             =   2310
            Width           =   650
         End
         Begin VB.CommandButton Command3 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   640
            Left            =   105
            Picture         =   "FOPCMAQ07.frx":0FEF
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Ayuda FLEXOFT en Línea"
            Top             =   790
            Width           =   650
         End
         Begin VB.CommandButton Command11 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Prove."
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   690
            Left            =   105
            Picture         =   "FOPCMAQ07.frx":12F9
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Acceso a Base de Datos de Proveedores"
            Top             =   1550
            Width           =   650
         End
         Begin VB.CommandButton Command5 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Setup"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   690
            Left            =   105
            Picture         =   "FOPCMAQ07.frx":1603
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Configuración de Funcionamiento de Pedido de Cotización"
            Top             =   3045
            Width           =   650
         End
         Begin VB.CommandButton Command2 
            BackColor       =   &H00E0E0E0&
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
            Height          =   690
            Left            =   105
            Picture         =   "FOPCMAQ07.frx":1A45
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Cierra y Abandona la Aplicación"
            Top             =   100
            Width           =   650
         End
         Begin VB.PictureBox MARCOBARRA 
            BackColor       =   &H000000FF&
            Height          =   135
            Left            =   105
            ScaleHeight     =   75
            ScaleWidth      =   585
            TabIndex        =   2
            Top             =   6360
            Width           =   650
            Begin VB.Frame BARRATRAZA 
               BackColor       =   &H00FF0000&
               BorderStyle     =   0  'None
               Height          =   500
               Left            =   0
               TabIndex        =   3
               Top             =   105
               Width           =   12000
            End
         End
         Begin VB.Image Image1 
            Height          =   390
            Left            =   -315
            Picture         =   "FOPCMAQ07.frx":1B8F
            Top             =   6600
            Width           =   1515
         End
      End
      Begin VB.Line Line7 
         X1              =   165
         X2              =   10680
         Y1              =   5640
         Y2              =   5640
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Por el Presente solicitamos la Cotización de los siguientes Items:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   375
         Left            =   240
         TabIndex        =   28
         Top             =   2550
         Width           =   10455
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "PC"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   750
         Left            =   210
         TabIndex        =   17
         Top             =   210
         Width           =   1695
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00C0C000&
         FillColor       =   &H0080FFFF&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   840
         Shape           =   4  'Rounded Rectangle
         Top             =   375
         Width           =   540
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H0000C000&
         FillColor       =   &H00FFFF80&
         FillStyle       =   0  'Solid
         Height          =   390
         Left            =   735
         Shape           =   4  'Rounded Rectangle
         Top             =   315
         Width           =   750
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H0000C0C0&
         FillColor       =   &H00C0E0FF&
         FillStyle       =   0  'Solid
         Height          =   540
         Left            =   630
         Shape           =   4  'Rounded Rectangle
         Top             =   240
         Width           =   960
      End
   End
End
Attribute VB_Name = "FOPCMAQ07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    List1.Clear
    Call APERBASDAT("COMPRAS")
    SQLS$ = "SELECT  NUME_SOLINT, COD_INTE, Descritem, CANT_SOLINT, FENTRE_SOL, NUORD_ITEM  "
    SQLS$ = SQLS$ + "FROM DETASOLINT "
    SQLS$ = SQLS$ + "WHERE COTI_SOLINT < 1 OR COTI_SOLINT IS NULL "
    SQLS$ = SQLS$ + "ORDER BY NUME_SOLINT ASC , NUORD_ITEM ASC "
    '
    Dim BUSSOLITMP As ADODB.Recordset
    Set BUSSOLITMP = New ADODB.Recordset
    BUSSOLITMP.OPEN FILTSQL$ (SQLS$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
    '
    With BUSSOLITMP
        On Error Resume Next
        .MoveFirst
        If Err < 1 Then
          Do While Not .EOF
              NUSOLI& = !NUME_SOLINT
              SQLS$ = "select * FROM ENCASOLINT WHERE NUME_SOLINT = " & NUSOLI&
              SQLS$ = SQLS$ + " AND  (Status <> 9 OR STATUS IS NULL)"
              SQLS$ = SQLS$ + " AND (AUTORIZA_1 IS NOT NULL AND AUTORIZA_2 IS NOT NULL AND "
              SQLS$ = SQLS$ + "AUTORIZA_3 IS NOT NULL AND AUTORIZA_4 IS NOT NULL)"
              Dim ENCASO As ADODB.Recordset
              Set ENCASO =  FLEXCONN.EXECUTE (FILTSQL$ (SQLS$))
              If ENCASO.BOF And ENCASO.EOF Then GoTo 29
              '
              TTXX$ = Space$(256)
              FEX = CVINT(!FENTRE_SOL)
              Call REPLA(TTXX$, FORMATNUM$(!NUME_SOLINT, "F5"), 1, 5)
              Call REPLA(TTXX$, "/" & TRIM$(Str$(!NUORD_ITEM)), 6, 3)
              Call REPLA(TTXX$, CODEXT$(!COD_INTE), 10, 16)
              Call REPLA(TTXX$, TRIM$(SAFETEXT$(!DESCRITEM)), 27, 45)
              Call REPLA(TTXX$, FORMATNUM$(!CANT_SOLINT, "F10.1"), 76, 10)
              Call REPLA(TTXX$, FECHATEX$(FEX), 88, 8)
              Call REPLA(TTXX$, !NUORD_ITEM, 100, 8)
              List1.AddItem TTXX$
29           .MoveNext
          Loop
        End If
        On Error GoTo 0
    End With
    BUSSOLITMP.Close
    Set BUSSOLITMP = Nothing
    
End Sub

Private Sub Command11_Click()
   Command11.Enabled = False
   Call CONECRUN("ARCHIG07/AMAPRO", "Padron Maestro de Proveedores")
   Command11.Enabled = True
End Sub

Private Sub Command12_Click()
    '
    Command12.Enabled = False
    '
10  If ULTREG&("INCOTIN") > 0 Then ' SI TIENE INFORMACION MUESTRA EL INDICE
      
      Call SELECHO("INCOTIN/Indice General de Solicitud de Cotizaciones")
    Else
      GoTo 99
    End If
    '
    DOCUNU& = XVALO(VALACT1$("INCOTIN")) ' ASIGNO EL NUMERO DEL FORMULARIO IMPRESO  SELECCIONAD0
    If DOCUNU& = 0 Then GoTo 99
    '
    Call MUESTRADOC(DOCUNU&) 'PERMITE VER LOS DOCUMENTOS ANTIGUOS POR PANTALLA
   
    GoTo 10
99  Command12.Enabled = True
End Sub

Private Sub Command14_Click()
  '
  'IMPRIME COTIZACION
  '
  If List4.ListCount <= 0 Then
    Call MENSERR(24, "Debe Elegir Al Menos un Proveedor")
    Exit Sub
  End If
  '
  FECX% = FECHANUM(Text9)
  FF% = HOY(HO$) ' Validacion de que la fecha no sea menor a la del dia
  If FECX% < FF% Then Call MENSERR(24, "La Fecha no Puede ser Menor a la de Hoy"): Exit Sub
  '
  SOLI = XVALO(VALOBADA("DETASOLINT", "MAX(COTI_SOLINT)", "COTI_SOLINT > 0"))
  SOLI = SOLI + 1
  For UU% = 0 To List1.ListCount - 1
    If List1.Selected(UU%) = True Then
    X$ = List1.List(UU%)
         NUMSOL% = XVALO(Mid$(X$, 1, 8))
         NUMITE% = XVALO(Mid$(X$, 100, 8))
         REGIS& = 0
         Call ACTUREGISTRO("DETASOLINT", "COTI_SOLINT", "NUME_SOLINT = " & NUMSOL% & " AND NUORD_ITEM = " & NUMITE%, SOLI, REGI&)
         '
         If NSLI& = XVALO(Mid$(X$, 1, 8)) Then GoTo 10 'Para que no grabe por cada item seleccionado de la misma solicitud
         TTXX$ = "" 'Graba los proveedores a los que se le pidio cotizacion en el anotador de solicitudes
         TTXX$ = "Pedido de Cotización: " & CRLF$
         For I% = 0 To List4.ListCount - 1
            L1$ = List4.List(I%)
            NPRO% = Val(Mid$(L1$, 1, 5))
            PROVE = RASOCLI((-1) * NPRO%)
            TTXX$ = TTXX$ + "- " & PROVE & CRLF$
         Next I%
         NSLI& = XVALO(Mid$(X$, 1, 8))
         Call ANOTASIM07.ANOTASIM(NSLI&, TTXX$)
         CANTITEM% = CANTITEM% + 1
         '
    End If
10 Next UU%
  '
  If CANTITEM% < 1 Then
    Call MENSERR(24, "Debe Elegir Algún Artículo")
    Exit Sub
  End If
  '
  FORIPRE$ = CONTROL$("", "FORMPECO")
  '
  If TRIM$(FORIPRE$) <> "" Then
    '
    For JX% = 0 To List4.ListCount - 1 'UNO POR CADA PROVEEDOR
      L1$ = List4.List(JX%)
      NPRO% = Val(Mid$(L1$, 1, 5))
      Screen.MousePointer = 11
      FECHDOC$ = Text9.TEXT
      NUMEDOC$ = SOLI
      TIPODOC$ = "Pedido de Cotización"
      '
      CODPARAM$(1) = "USUARIO"
      CODPARAM$(2) = "REF-MAT2"
      CODPARAM$(3) = "OBSERVA"
      CAPARDOC% = 3
      '
      DOCPARAM$(1) = USERNAME$
      DOCPARAM$(2) = Text1.TEXT
      DOCPARAM$(3) = Text13.TEXT   'Observaciones
      '
      CODTABU1$(1) = "COD-MPRI"
      CODTABU1$(2) = "DES-MPRI"
      CODTABU1$(3) = "CAN-MPRI"
      CODTABU1$(4) = "FECH-REC"
      CODTABU1$(5) = "PAR-MPRI"
     
      CACOLTAB1% = 5
      '
      DESTIDOC% = NPRO% * (-1)
      CAITAB1% = 0
      CAITCRU% = 0
      '
      NC1% = DESTIDOC%
      FORTELIB$ = TRIM$(UCase$(ATRIFORM$(FORIPRE$, "DES-MPRI/FORMATO"))) ' OBTIENE ANCHO DE LA DESCRIPCION
      ATELIB% = XVALO(Mid$(FORTELIB$, 2))
      '
      For J1% = 0 To List1.ListCount - 1
         If List1.Selected(J1%) = True Then
            X$ = List1.List(J1%)
              NSOLI$ = TRIM$(Mid$(X$, 1, 8))
              MM% = InStr(1, NSOLI$, "/")
            NSOLI$ = Mid$(NSOLI$, 1, MM% - 1) ' OBTIENE EL NUMERO DE SOLICITUD SIN EL DE ORDEN
            CI0% = CODINT(Mid$(X$, 10, 16))
            CANTU = Val(Mid$(X$, 76, 10))
            FEX = FECHANUM(TRIM$(Mid$(X$, 88, 8)))
            NUORD$ = TRIM$(Mid$(X$, 100, 8))
            CAITAB1% = CAITAB1% + 1
            CODDX$ = TRIM$(CODEXT$(CI0%))
            DECRIT$ = VALOBADA("DETASOLINT", "DESCRITEM", "NUME_SOLINT=" & NSOLI$ & " AND NUORD_ITEM= " & NUORD$)
            '
            TETABU1$(1, CAITAB1%) = CODDX$
            TETABU1$(3, CAITAB1%) = TRIM$(Str$(CANTU))
            TETABU1$(4, CAITAB1%) = FECHATEX$(FEX)
            TETABU1$(5, CAITAB1%) = NSOLI$
            '
            Call FRATEXTO(DECRIT$, ATELIB%) ' FRACCIONA SEGUN LO CONFIGURADO EN EL FORMULARIO
            For U& = 1 To CARETEX%
              Z$ = "." + RETEX$(U&)
              TETABU1$(2, CAITAB1%) = Z$
              CAITAB1% = CAITAB1% + 1
            Next U&
         End If
      Next J1%
      '
      MAILPRO$ = TRIM$(EMAILCLI$(NPRO% * -1))
      If MAILPRO$ <> "" Then
        SENDBYMAIL$ = MAILPRO$
      End If
      
      Call PRINTDOC("FORMPECO")   ' Pedido de Cotización
      '
      For KK1% = 1 To CACOLTAB1%
        For KK2% = 1 To CAITAB1%
          TETABU1$(KK1%, KK2%) = ""
        Next KK2%
      Next KK1%
      '
      For KK1% = 1 To CACOLTAB2%
        For KK2% = 1 To CAITAB2%
          TETABU2$(KK1%, KK2%) = ""
        Next KK2%
      Next KK1%
      '
    Next JX%
     '
99  Call CIERRARCH("*.*")
    Call BAJALDISCO
    Screen.MousePointer = 1
    '
    Call GENINCOT
    '
    Call BLANFORMU
    '
   End If
End Sub
Private Sub Command2_Click()
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    Hide
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    Call HELPONLINE("PEDICOTI")
    Command3.Enabled = True
End Sub

Private Sub Command4_Click()
    '
    Command4.Enabled = False
    '
    If DERACCE%("SELECSIM", "Ver todas las S.I.M.s") < 2 Then GoTo 99
    '
10  If ULTREG&("INSOLINT") > 0 Then ' SI TIENE INFORMACION MUESTRA EL INDICE
         Call SELECHO("INSOLINT/Indice General de Solicitud de Materiales")
       Else
         GoTo 99
    End If
    '
    NUDOCU$ = VALACT1$("INSOLINT") ' ASIGNO EL NUMERO DEL FORMULARIO IMPRESO  SELECCIONAD0
    If NUDOCU$ = "" Then GoTo 99
    '
    'Verifica estado de la sim
    RESS% = XVALO(VALOBADA("ENCASOLINT", "STATUS", "NUME_SOLINT = " & XVALO(NUDOCU$) & " AND Status = 9"))
    If RESS% <> 0 Then
        Call MENSERR(24, "Imposible Pedir Cotización\SIM Anulada")
        GoTo 10
    End If
    '
    List1.Clear
    Call APERBASDAT("COMPRAS")
    For I% = 1 To 4
        AUTO$ = "AUTORIZA_" & I%    ' CARGA EL AUTORIZANTE
        XX = (VALOBADA("ENCASOLINT", AUTO$, " NUME_SOLINT = " & NUDOCU$)) 'VERIFICA QUE NINGUN AUTORIZANTE SEA NULO
        If XX = "" Then
            Call MENSERR(24, "Solicitud sin Autorización Completa.")
            GoTo 10
        End If
    Next I%
    Call APERBASDAT("COMPRAS")
    SQLS$ = "SELECT  NUME_SOLINT, COD_INTE, Descritem, CANT_SOLINT, FENTRE_SOL, NUORD_ITEM  "
    SQLS$ = SQLS$ + "FROM DETASOLINT "
    'SQLS$ = SQLS$ + "WHERE NUME_SOLINT = " & NUDOCU$ & " AND (COTI_SOLINT < 1 OR COTI_SOLINT IS NULL) "
    SQLS$ = SQLS$ + "WHERE NUME_SOLINT = " & NUDOCU$ & " "
    SQLS$ = SQLS$ + "ORDER BY NUME_SOLINT ASC , NUORD_ITEM ASC "
    '
    Dim BUSSOLITMP As ADODB.Recordset
    Set BUSSOLITMP = New ADODB.Recordset
    BUSSOLITMP.OPEN FILTSQL$ (SQLS$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
    '
    With BUSSOLITMP
        On Error Resume Next
        .MoveFirst
        If Err < 1 Then
          Do While Not .EOF
              NUSOLI& = !NUME_SOLINT
              SQLS$ = "select * FROM ENCASOLINT WHERE NUME_SOLINT = " & NUSOLI&
              SQLS$ = SQLS$ + " AND (AUTORIZA_1 IS NOT NULL AND AUTORIZA_2 IS NOT NULL AND "
              SQLS$ = SQLS$ + "AUTORIZA_3 IS NOT NULL AND AUTORIZA_4 IS NOT NULL)"
              Dim ENCASO As ADODB.Recordset
              Set ENCASO =  FLEXCONN.EXECUTE (FILTSQL$ (SQLS$))
              If ENCASO.BOF And ENCASO.EOF Then GoTo 29
              '
              TTXX$ = Space$(256)
              FEX = CVINT(!FENTRE_SOL)
              Call REPLA(TTXX$, FORMATNUM$(!NUME_SOLINT, "F5"), 1, 5)
              Call REPLA(TTXX$, "/" & TRIM$(Str$(!NUORD_ITEM)), 6, 3)
              Call REPLA(TTXX$, CODEXT$(!COD_INTE), 10, 16)
              Call REPLA(TTXX$, DESCRIT0$(!COD_INTE), 27, 45)
              Call REPLA(TTXX$, FORMATNUM$(!CANT_SOLINT, "F10.1"), 76, 10)
              Call REPLA(TTXX$, FECHATEX$(FEX), 88, 8)
              Call REPLA(TTXX$, !NUORD_ITEM, 100, 8)
              List1.AddItem TTXX$
29           .MoveNext
          Loop
        End If
        On Error GoTo 0
    End With
    BUSSOLITMP.Close
    Set BUSSOLITMP = Nothing

99  Command4.Enabled = True
End Sub

Private Sub Command5_Click()
  '
  If DERACCE%("CONFPCOT", "Setup Pedido De Cotización") > 1 Then
    OPPCMAQ07.Show 1
  End If
  '
End Sub

Private Sub Command6_Click()
    If DERACCE%("MODISIM", "Modificación de S.I.M.s") > 1 Then
      MODISIM07.Show 1
    End If
    '
End Sub

Private Sub Command7_Click()
    Command7.Enabled = False
    Call FINAL("?LISPEDCOTI")
    Command7.Enabled = True
End Sub

Private Sub Command8_Click()
    Command8.Enabled = False
    Call CIERRARCH("*.*")
    Call CONECRUN("AMASTO07", "Maestro de Items de Stock")
    Command8.Enabled = True
End Sub

Private Sub Command9_Click()
   Call SELECHO("SELFECHA")
   VDEVU$ = TRIM$(VALACT1$("SELFECHA"))
   If VDEVU$ <> "" Then
      Text9.TEXT = VALACT1$("SELFECHA")
   End If
   FF% = FECHANUM(Text9.TEXT)
End Sub

Private Sub Form_Load()
  '
  Call VERJOBID(OKEY%)
  If OKEY% < 1 Then Call FINAL("")
  '
  Call Command1_Click   'Llama al boton SIM's Pendientes
  '
  NSOLI = XVALO(VALOBADA("DETASOLINT", "MAX(COTI_SOLINT)", "COTI_SOLINT > 0")) 'Numero de Solicitud
  Label6.Caption = NSOLI + 1
  '
  Call CENTRAFORM(Me)
  EPAC$ = TRIM$(EMPREAC$)
  If EPAC$ <> "" Then
      Caption = Caption + "  -  " + EPAC$
  End If
  Call GRATITFOR(Caption)
  '
  FECHA1$ = Format(Now, "yymmdd")
  ID1$ = CONTROL$("", "IDENPECO")
  If ID1$ = "" Then ID1$ = "PC"
  'Text6.TEXT = ID1$ + "-" + FECHA1$
  '
  HOII% = HOY(HOS$)
  Text9.TEXT = HOS$
  Label7 = Left$(USERNAME, 20) ' CAMBIO TEXT2 POR Label7
  '
End Sub
Sub BLANFORMU()
  '
  NSOLI = XVALO(VALOBADA("DETASOLINT", "MAX(COTI_SOLINT)", "COTI_SOLINT > 0")) 'Numero de Solicitud
  Label6.Caption = NSOLI + 1
  FECHA1$ = Format(Now, "yymmdd")
  ID1$ = CONTROL$("", "IDENPECO")
  If ID1$ = "" Then ID1$ = "PC"
  'Text6.TEXT = ID1$ + "-" + FECHA1$
  '
  HOII% = HOY(HOS$)
  Text9.TEXT = HOS$
  '
  Text13.TEXT = ""
  List1.Clear
  List4.Clear
  Text1.TEXT = ""
  '
End Sub

Private Sub Label92_dblClick()
  '
  Call BLANARCH("!INVITA2")
  JJ& = 0
  For JX% = 0 To List4.ListCount - 1
    L1$ = List4.List(JX%)
    NPRO% = Mid$(L1$, 1, 5)
    '
    YY$ = REGBLAN$("INVITA2")
    Call REPLA(YY$, MKI$(NPRO%), 1, 2)
    JJ& = JJ& + 1
    Call GRAREG("!INVITA2", YY$, JJ&)
    '
  Next JX%
  Call SETULTREG("!INVITA2", JJ&)
  Call CIERRARCH("!INVITA2")
  '
  PRF$ = "INVITA2"
  ATRI$ = ATRI$ + "/TITUPAN=Lista de Proveedores Invitados /"
  ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + FOPCMAQ07.Left + 1000))
  ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + FOPCMAQ07.Top + 5))
  VTB% = VENTABU%(PRF$ + ATRI$)
  If VTB% >= 0 Then
    '
    List4.Clear
    '
    FIN& = ULTREG&("!INVITA2")
    For J& = 1 To FIN&
      X$ = REGLEIDO$("!INVITA2", J&)
      NPRO1% = CVI(Mid$(X$, 1, 2))
        NPRO2$ = FORMATNUM$(NPRO1%, "F5.0")
      PRO1$ = ECOARCH$("PROVED1", MKI$(NPRO1%))
      'AGREGO AL LIST
      L1$ = Space(128)
      Call REPLA(L1$, NPRO2$, 1, 5)
      Call REPLA(L1$, PRO1$, 8, 20)
      List4.AddItem L1$
    Next J&
    '
  End If
  '
End Sub

Private Sub List1_Click()
    List4.Clear
    CONDI$ = ""
    For UU% = 0 To List1.ListCount - 1
      If List1.Selected(UU%) = True Then
         X$ = List1.List(UU%)
         CODE$ = TRIM$(Mid$(X$, 10, 16))
         CIXI% = CODINT%(CODE$)
         If CIXI% > 0 Then
            If CONDI$ = "" Then
                 CONDI$ = "WHERE COD_INTE = " & CIXI% & " "
               Else
                 CONDI$ = CONDI$ & " OR COD_INTE = " & CIXI% & " "
            End If
         End If
      End If
    Next UU%
    '
    If CONDI$ = "" Then Exit Sub ' VALIDA QUE HAYA ELEGIDO ALGUN ARTICULO
    Call ABRECOMPRAS
    TABLASE$ = "OCOMDETA"
    '
10  SQLX$ = "SELECT DISTINCT NPRO_OCOM FROM " & TABLASE$ & " " & CONDI$
    Set PROSEL =  FLEXCONN.EXECUTE (FILTSQL$ (SQLX$))
    With PROSEL
       Do While Not .EOF
          NPRO1% = XVALO(!NPRO_OCOM)
          For KKI% = 1 To List4.ListCount
             If XVALO(Left$(List4.List(KKI% - 1), 5)) = NPRO1% Then GoTo 17
          Next KKI%
          NPRO2$ = FORMATNUM$(NPRO1%, "F5.0")
          PRO1$ = ECOARCH$("PROVED1", MKI$(NPRO1%))
          L1$ = Space(128)
          Call REPLA(L1$, NPRO2$, 1, 5)
          Call REPLA(L1$, PRO1$, 8, 20)
          List4.AddItem L1$
17     .MoveNext
       Loop
    End With
    PROSEL.Close
    Set PROSEL = Nothing
    '
    If TABLASE$ = "OCOMDETA" Then
       TABLASE$ = "COTIRECI"
       GoTo 10
    End If
    '
End Sub

Private Sub List1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  '
  CODI3$ = Mid$(List1.List(List1.ListIndex), 10, 16) ' MODIFICACION DE LECTURA DE CODEXT
  CODI2% = CODINT%(CODI3$)
  List1.ToolTipText = DESCRIT0(CODI2%)
  '
End Sub

Private Sub List4_Click()
  '
  Call Label92_dblClick
  '
End Sub

Sub GENINCOT() 'CARGA ARCHIVO CENTRAL DE COTIZACIONES
    '
    Call SETULTREG("INCOTIN", 0)
    Call COPYSTRU("PDOCLST", "INCOTIN")
    JJ& = 0
    For UI& = 1 To ULTREG&("PDOCLST")
         X$ = Mid$(REGLEIDO$("PDOCLST", UI&), 1, 60)
         If Mid$(X$, 5, 13) = "Pedido de Cot" Then
            NUMEL& = CVS(Mid$(X$, 1, 4))
            DESC$ = Mid$(X$, 5, 60)
            Call REPLA(X$, MKS(NUMEL&), 1, 4)
            Call REPLA(X$, DESC$, 5, 60)
            JJ& = JJ& + 1
            Call GRAREG("INCOTIN", X$, JJ&)
         End If
    Next UI&
    '
    Call SETULTREG("INCOTIN", JJ&)
    Call CIERRARCH("INCOTIN")
    Call FRESHECHO("INCOTIN")
    '
End Sub


