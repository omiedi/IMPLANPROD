VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form ASIGLOTERUTA 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Agregar Certificados"
   ClientHeight    =   3390
   ClientLeft      =   75
   ClientTop       =   270
   ClientWidth     =   14145
   ForeColor       =   &H00000000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3390
   ScaleWidth      =   14145
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CD1 
      Left            =   12480
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   3375
      Left            =   100
      TabIndex        =   3
      Top             =   0
      Width           =   13935
      Begin VB.TextBox Text9 
         Appearance      =   0  'Flat
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
         Index           =   1
         Left            =   12000
         Locked          =   -1  'True
         TabIndex        =   30
         TabStop         =   0   'False
         Text            =   "Hojas"
         Top             =   1480
         Width           =   550
      End
      Begin VB.TextBox Text6 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   12000
         MaxLength       =   4
         TabIndex        =   29
         ToolTipText     =   "Cantidad de Hojas del Certificado"
         Top             =   1735
         Width           =   550
      End
      Begin VB.CommandButton Command62 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   13200
         Picture         =   "ASIGLOTERUTA.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Tabla de Títulos de Selección"
         Top             =   2500
         Width           =   550
      End
      Begin VB.CommandButton Command9 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   12600
         TabIndex        =   23
         Top             =   2500
         Width           =   550
      End
      Begin VB.CommandButton Command8 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   12000
         TabIndex        =   22
         ToolTipText     =   "Lista de Selección de Títulos de Certificados"
         Top             =   2500
         Width           =   550
      End
      Begin VB.CommandButton Command2 
         Caption         =   "New"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   12600
         Picture         =   "ASIGLOTERUTA.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Agrega una Nueva Ruta"
         Top             =   240
         Width           =   550
      End
      Begin VB.CommandButton Command1 
         Caption         =   "O.K."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   13200
         Picture         =   "ASIGLOTERUTA.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Cargar Items"
         Top             =   240
         Width           =   550
      End
      Begin VB.CommandButton Command4 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   13200
         TabIndex        =   14
         Top             =   1480
         Width           =   550
      End
      Begin VB.CommandButton Command3 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   12600
         TabIndex        =   13
         Top             =   1480
         Width           =   550
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "A.B.M. de Ceritificado x Orden de Fabricación "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1065
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   200
         Width           =   11715
         Begin VB.CommandButton Command7 
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
            Height          =   280
            Left            =   11220
            TabIndex        =   21
            Top             =   260
            Width           =   375
         End
         Begin VB.CommandButton Command6 
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
            Height          =   280
            Left            =   7730
            TabIndex        =   20
            Top             =   260
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.CommandButton Command5 
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
            Height          =   280
            Left            =   9245
            TabIndex        =   19
            Top             =   255
            Visible         =   0   'False
            Width           =   370
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Left            =   9620
            Locked          =   -1  'True
            TabIndex        =   17
            ToolTipText     =   "Doble Click Para Limpiar"
            Top             =   540
            Width           =   1995
         End
         Begin VB.TextBox Text9 
            Appearance      =   0  'Flat
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
            Index           =   0
            Left            =   9620
            Locked          =   -1  'True
            TabIndex        =   16
            TabStop         =   0   'False
            Text            =   "N.Serie"
            Top             =   270
            Width           =   1995
         End
         Begin VB.TextBox Text2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
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
            Left            =   2240
            Locked          =   -1  'True
            TabIndex        =   12
            TabStop         =   0   'False
            Text            =   "Descripción"
            Top             =   270
            Width           =   4100
         End
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Left            =   2240
            Locked          =   -1  'True
            TabIndex        =   11
            TabStop         =   0   'False
            ToolTipText     =   " "
            Top             =   540
            Width           =   4100
         End
         Begin VB.TextBox Text10 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Left            =   8105
            Locked          =   -1  'True
            TabIndex        =   10
            TabStop         =   0   'False
            ToolTipText     =   "Doble Click Para Limpiar"
            Top             =   540
            Width           =   1530
         End
         Begin VB.TextBox Text11 
            Appearance      =   0  'Flat
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
            Height          =   350
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   9
            Top             =   540
            Width           =   2175
         End
         Begin VB.TextBox Text13 
            Appearance      =   0  'Flat
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
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   8
            TabStop         =   0   'False
            Text            =   "Código"
            Top             =   270
            Width           =   2175
         End
         Begin VB.TextBox Text14 
            Appearance      =   0  'Flat
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
            Left            =   8105
            Locked          =   -1  'True
            TabIndex        =   7
            TabStop         =   0   'False
            Text            =   "O/Fab."
            Top             =   270
            Width           =   1530
         End
         Begin VB.TextBox Text9 
            Appearance      =   0  'Flat
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
            Index           =   5
            Left            =   6300
            Locked          =   -1  'True
            TabIndex        =   6
            TabStop         =   0   'False
            Text            =   "B/Recepcion"
            Top             =   270
            Width           =   1840
         End
         Begin VB.TextBox Text15 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Left            =   6300
            Locked          =   -1  'True
            TabIndex        =   5
            ToolTipText     =   "Doble Click Para Limpiar"
            Top             =   540
            Width           =   1840
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ruta de Certificado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         TabIndex        =   25
         Top             =   1250
         Width           =   11775
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Left            =   70
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   26
            TabStop         =   0   'False
            ToolTipText     =   " "
            Top             =   240
            Width           =   11595
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Título"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         TabIndex        =   27
         Top             =   2280
         Width           =   11775
         Begin VB.TextBox Text5 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Left            =   70
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   28
            TabStop         =   0   'False
            ToolTipText     =   " "
            Top             =   240
            Width           =   11595
         End
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFC0C0&
         Height          =   375
         Left            =   12000
         TabIndex        =   32
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
         Height          =   375
         Left            =   12000
         TabIndex        =   31
         Top             =   1080
         Width           =   1575
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   4920
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   0
      Top             =   9720
      Visible         =   0   'False
      Width           =   1170
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   120
      OleObjectBlob   =   "ASIGLOTERUTA.frx":075E
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "ASIGLOTERUTA.frx":0992
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   435
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   1560
   End
End
Attribute VB_Name = "ASIGLOTERUTA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias _
    "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public ORDEX%
'

Private Sub Command1_Click()

    Command1.Enabled = False
    'MODIFICA LOS DATOS DE LA LISTA.
'    If TRIM$(Text10) = "" And TRIM$(Text15) = "" And TRIM$(Text1) = "" Then
'       Call COMUNI("Falta Ingresar Lote-OF. o Ruta de Certificado")
'       GoTo 99
'    End If
    '
    '
    LOTE$ = TRIM$(Text15)   '  BR u OF
    CODI$ = Text11
    CI1% = CODINT(CODI$)
    NORFA$ = TRIM$(Text10)
    NUSER1$ = TRIM$(Text3)
    If LOTE$ <> "" Then
       If CantRegistros&("CERTIFICADOS", "IDENLOTE = '" & LOTE$ & "' AND CODINT = " & CI1%, "NOMESS") < 1 Then
          Call COMUNI("Lote Asignado No Posee Archivo de Certificado")
          ABMCERTI09.Text6 = LOTE$
          ABMCERTI09.Show 1
       End If
    ElseIf NORFA$ <> "" Then
       CONDI$ = "IDENLOTE = '" & NORFA$ & "' AND CODINT = " & CI1%
       If NUSER1$ <> "" Then CONDI$ = CONDI$ + " AND NUMEROSERIE = '" & NUSER1$ & "'"
       If CantRegistros&("CERTIFICADOS", CONDI$, "NOMESS") < 1 Then
           Call COMUNI("Lote Asignado No Posee Archivo de Certificado")
           ABMCERTI09.Text11 = CODI$
           ABMCERTI09.Text4 = DESCRIT$(CODINT%(CODI$))
           ABMCERTI09.Text10 = NORFA$
           ABMCERTI09.Show 1
       End If
    End If
    
    With CERTIFICADOS
        If TRIM$(Text15) <> "" Then LOTE$ = TRIM$(Text15)
        If TRIM$(Text10) <> "" Then LOTE$ = TRIM$(Text10)
       .MSF.TextMatrix(.MSF.Row, 5) = LOTE$ 'LOTE
       .MSF.TextMatrix(.MSF.Row, 6) = Text3 'NS
       .MSF.TextMatrix(.MSF.Row, 7) = Text1 'LINK
       .MSF.TextMatrix(.MSF.Row, 8) = SAFETEXT$(ORDEX%)
       .MSF.TextMatrix(.MSF.Row, 9) = Text6
       'ORDENA
       Call ORDENAR_MSF(CERTIFICADOS.MSF, 1)
       End With
    CERTIFICADOS.CAMBIO% = 1
    Unload Me
99  Command1.Enabled = True
End Sub

Private Sub Command2_Click()
   Command2.Enabled = False
   '
'   If TRIM$(Text10) = "" And TRIM$(Text1) = "" Then
'       Call COMUNI("Falta Ingresar Lote-OF. o Ruta de Certificado")
'       GoTo 99
'   End If
   '
'   If TRIM$(Text15) <> "" Or TRIM$(Text10) <> "" Then
'       Call COMUNI("Imposible Asignar Cant.Hojas Desde Aquí\El Mismo se Asocia Desde El A.B.M. De Certificados")
'       Text6 = ""
'       GoTo 99
'   End If
   '
   If VALI_EXISTE_RUTA%(TRIM$(Text11), TRIM$(Text15), TRIM$(Text10), TRIM$(Text3), TRIM$(Text1)) > 0 Then GoTo 99
   '
   'ACTUALIZAR
   'AGREGA LA RUTA A LA GRILLA
   With CERTIFICADOS
        i& = .MSF.Rows
       .MSF.Rows = i& + 1
       .MSF.TextMatrix(i&, 1) = .MSF.TextMatrix(.MSF.Row, 1)
       .MSF.TextMatrix(i&, 2) = .MSF.TextMatrix(.MSF.Row, 2)
       .MSF.TextMatrix(i&, 3) = .MSF.TextMatrix(.MSF.Row, 3)
       .MSF.TextMatrix(i&, 4) = .MSF.TextMatrix(.MSF.Row, 4)
         If TRIM$(Text15) <> "" Then LOTE$ = TRIM$(Text15)
         If TRIM$(Text10) <> "" Then LOTE$ = TRIM$(Text10)
       .MSF.TextMatrix(i&, 5) = LOTE$
       .MSF.TextMatrix(i&, 6) = Text3
       .MSF.TextMatrix(i&, 7) = Text1
       .MSF.TextMatrix(.MSF.Row, 8) = SAFETEXT$(ORDEX%)
       .MSF.TextMatrix(.MSF.Row, 9) = Text6
       'ORDENA
       .MSF.COL = 1
       .MSF.Sort = 1
    End With

    CERTIFICADOS.CAMBIO% = 1
    Unload Me
99  Command2.Enabled = True
End Sub
Function VALI_EXISTE_RUTA%(CODI$, LOTE$, NORFA$, NUMESERI$, RUTAX$)
   '
   VALI_EXISTE_RUTA% = 0
   If TRIM$(LOTE$) = "" And TRIM$(NUMESERI$) = "" And TRIM$(RUTAX$) = "" And TRIM$(NORFA$) = "" Then
     Call COMUNI("Nó Válido Ingresar - Faltan Datos")
     VALI_EXISTE_RUTA% = 1
     Exit Function
   End If
   '
   With CERTIFICADOS
    'VALIDO QUE LA RUTA DEL CODIGO Y LOTE NO SEA LA MISMA
    For i& = 1 To .MSF.Rows - 1
      If TRIM$(UCase$(.MSF.TextMatrix(i&, 1))) = UCase$(CODI$) Then
        If TRIM$(UCase$(.MSF.TextMatrix(i&, 5))) = UCase$(LOTE$) Or TRIM$(UCase$(.MSF.TextMatrix(i&, 5))) = UCase$(NORFA$) Then
            If TRIM$(UCase$(.MSF.TextMatrix(i&, 6))) = UCase$(NUMESERI$) Then
              If TRIM$(.MSF.TextMatrix(i&, 7)) = "" Then
                Call Command1_Click 'SI EXISTE LA FILA Y ESTA VACIA LA RUTA GRABA NORMALMENTE SIN GENERAR NUEVO REGISTRO.
                VALI_EXISTE_RUTA% = 1 'LE ASIGNA 1 PARA QUE CUANDO VUELVA NO CONTINUE
                Exit For
              End If
              If UCase$(.MSF.TextMatrix(i&, 7)) = UCase$(RUTAX$) Then
                Call COMUNI("La Ruta Seleccionada Yá Existe Para Mismo Código y Lote u O.Fabricación")
                VALI_EXISTE_RUTA% = 1
                Exit For
              End If
            End If
        End If
      End If
    Next i&
   End With

   
End Function


Private Sub Command3_Click()
      Command3.Enabled = False
      If TRIM$(Text15) <> "" Or TRIM$(Text10) <> "" Then
         Call COMUNI("Imposible Asignar Certificado Desde Aquí\El Mismo se Asocia Desde El A.B.M. De Certificados")
         GoTo 99
      End If
      CD1.FILTER = "(*.*)|*.*|"
      CD1.ShowOpen
      Ruta$ = CD1.FileName
      Text1.TEXT = Ruta$
      
99    Command3.Enabled = True
End Sub

Private Sub Command4_Click()
   Text1 = ""
End Sub

Private Sub Command5_Click()
    '
    Command5.Enabled = False
    CODI$ = TRIM$(Text11)
     
    Screen.MousePointer = 11
    SQLX$ = " SELECT  IdSubOr FROM ORDEFABR WITH(NOLOCK) "
    SQLX$ = SQLX$ + " WHERE StatuOrf < 9 "
    SQLX$ = SQLX$ + " AND Cod_Exte = '" & CODI$ & "'"
    SQLX$ = SQLX$ + " ORDER BY IdSubOr "
    '
    CONDI$ = "StatuOrf < 9 "
    CONDI$ = CONDI$ + " AND Cod_Exte = '" & CODI$ & "'"
    '
    If CantRegistros&("ORDEFABR", CONDI$, "NOMESS") < 1 Then
       Screen.MousePointer = 1
       Call COMUNI("No existen Ordenes de Fabricación Para Este Código")
       GoTo 99
    End If
    
    Call CARGALISEL("ORFACODI", "ORDEFABR", "IDSUBOR/A12;COD_INTE/C16+MASTER/A24;CanProy/F8.0", CONDI$, "NOMESS")
    Screen.MousePointer = 1
    Call SELECHO("ORFACODI")
    NORFA$ = TRIM$(VALACT1$("ORFACODI"))
    '
    If NORFA$ = "" Then GoTo 99
    '
    Text10 = NORFA$
    
    
99  Screen.MousePointer = 1
    Command5.Enabled = True
End Sub

Private Sub Command6_Click()
       
        Command6.Enabled = False
        '
        CODI$ = TRIM$(Text11)
        CIXI% = CODINT%(CODI$)
        '
        CONDI$ = "COD_INTE=" & CStr(CIXI%) & " AND (DESTINOCONSUMO='' OR DESTINOCONSUMO IS NULL)"
        Call CARGALISEL("NUMSEDIS", "TRAZARECE", "NUMEROSERIE/I10;IDENLOTE/A18;ORIGEN/A18", CONDI$)
        Call SELECHO("NUMSEDIS/Disponibilidad (N.Serie - Lote - Origen)")
        
'        CONDI$ = " COD_EXTE = '" & CODI$ & "' ORDER BY NumeBoRe"
'
'        Screen.MousePointer = 11
'        Call CARGALISEL("RECECODI", "BOLRECEP", "NumeBoRe/F7;IDENLOTE/A12;FECHRECEP/D8;NumRem_Prov/A16;RaSo_Prov/A24", CONDI$, "NOMESS")
'        Call SELECHO("RECECODI/Recepciones")
'        NUMBRECE& = XVALO(VALACT1$("RECECODI"))
'        If NUMBRECE& < 1 Then GoTo 99
'
'        LOTE$ = Left$(VALACT2$("RECECODI"), 12)
'        Text15 = LOTE$
        
99      Command6.Enabled = True

End Sub

Private Sub Command62_Click()
    '
    Command62.Enabled = False
    '
    Call TRALOCAL("TITUCERT", "")
    TITUPA$ = "Carga de Títulos de Certificados"
    VTB% = VENTABU%("TITUCERT/TITUPAN=" + TITUPA$)
    If VTB% < 0 Then GoTo 99
    '
    J& = 0
    For U& = 1 To ULTREG&("!TITUCERT")
        X$ = REGLEIDO$("!TITUCERT", U&)
        ORDEX% = CVI(Left$(X$, 2))
        If ORDEX% >= 1 Then
            J& = J& + 1
            Call GRAREG("TITUCERT", X$, J&)
        End If
    Next U&
    '
    Call SETULTREG("TITUCERT", J&)
    Call CIERRARCH("TITUCERT")
    Call CIERRARCH("!TITUCERT")
    Call FRESHECHO("TITUCERT")
    '
99  Command62.Enabled = True
    '
End Sub

Private Sub Command7_Click()
    Command7.Enabled = False
    
    If LoteEditable = False Then
        Call COMUNI("Imposible Cambiar.\El Numero de Serie se Encuentra Vinculado con el Codigo de Conjunto")
        GoTo 99
    End If

    CODI$ = TRIM$(Text11)
    CIXI% = CODINT%(CODI$)
    ILO$ = TRIM$(Text15 & Text10)
    '
    CONDI$ = "COD_INTE=" & CStr(CIXI%) & " AND (DESTINO='' OR DESTINO='" & ILO$ & "') AND (DESTINOCONSUMO='' OR DESTINOCONSUMO IS NULL) AND NUMEROSERIE>0 ORDER BY NUMEROSERIE ASC"
    Call CARGALISEL("NUMSEDIS", "TRAZARECE", "NUMEROSERIE/I10;IDENLOTE/A18;ORIGEN/A18", CONDI$)
    Call SELECHO("NUMSEDIS/Disponibilidad (N.Serie - Lote - Origen)")
    '
    NSERIE$ = TRIM$(VALACT1$("NUMSEDIS"))
    If NSERIE$ = "" Then GoTo 99
    
    VALA2$ = TRIM$(VALACT2$("NUMSEDIS"))
    LOTE$ = TRIM$(Mid$(VALA2$, 1, 18))
    '
    Text10 = "": Text15 = ""
    If Mid$(LOTE$, 1, 2) = "BR" Then
        Text15.TEXT = LOTE$ ' RECEPCION
        Text10 = ""
        Text1 = ""
    Else
        Text10.TEXT = LOTE$ ' FABRICACION
        Text15 = ""
        Text1 = ""
    End If
    Text3.TEXT = NSERIE$
    '
99 Command7.Enabled = True
End Sub

Private Sub Command8_Click()
   Command8.Enabled = False
   If TRIM$(Text15) <> "" Or TRIM$(Text10) <> "" Then
      Call COMUNI("Imposible Asignar Título Desde Aquí\El Mismo se Asocia Desde El A.B.M. De Certificados")
      GoTo 99
   End If

   Call SELECHO("TITUCERT")
   ORDEX% = XVALO(VALACT1$("TITUCERT"))
   If ORDEX% < 1 Then GoTo 99
   '
   Text5 = VALACT2$("TITUCERT")
   
99 Command8.Enabled = True
End Sub

Private Sub Command9_Click()
   Text5 = ""
End Sub

Private Sub Form_Load()
    CANTIHOJA% = XVALO(Text6)
      
End Sub

Private Sub Text1_DblClick()
   TX$ = Text1
   Archi$ = TRIM$(TX$)
   If (EXISTE%(Archi$)) <> 1 Then
       Call COMUNI("Archivo Inexistente")
       Exit Sub
     Else
        Call ShellExecute(Me.hwnd, "Open", Archi$, "", "", 1)
    End If
End Sub

Private Sub Text10_DblClick()
    If TRIM$(Text10) <> "" And LoteEditable = True Then
        Text10 = ""
        Text3 = ""
        Text1 = ""
    End If
End Sub

Private Sub Text15_DblClick()
    If TRIM$(Text15) <> "" And LoteEditable = True Then
        Text15 = ""
        Text3 = ""
        Text1 = ""
    End If
End Sub

Private Sub Text3_DblClick()
   Text3 = ""
   Text10 = ""
   Text15 = ""
   Text1 = ""
End Sub

Private Sub Text6_GotFocus()
    If TRIM$(Text15) <> "" Or TRIM$(Text10) <> "" Then
         On Error Resume Next
         Text11.SetFocus
         On Error GoTo 0
    End If

End Sub

Function LoteEditable() As Boolean

    LoteEditable = True
    NUMSE$ = TRIM$(Label2)
    LOTE$ = TRIM$(Text15 & Text10)
    NORFA$ = TRIM$(Label1)
    If LOTE$ <> "" And NUMSE$ <> "" Then
        CONDI$ = "IDENLOTE='" & LOTE$ & "' AND NUMEROSERIE='" & NUMSE$ & "' AND DESTINO='" & NORFA$ & "'"
        If CantRegistros("TRAZARECE", CONDI$) > 0 Then LoteEditable = False
    End If
    '
End Function
