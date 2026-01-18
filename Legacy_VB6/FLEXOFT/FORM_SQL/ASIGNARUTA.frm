VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form ASIGNARUTA 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Edición de Certificados"
   ClientHeight    =   2790
   ClientLeft      =   75
   ClientTop       =   270
   ClientWidth     =   11775
   ForeColor       =   &H00000000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2790
   ScaleWidth      =   11775
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CD1 
      Left            =   9000
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   2655
      Left            =   50
      TabIndex        =   3
      Top             =   120
      Width           =   11700
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
         Left            =   9840
         Locked          =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Text            =   "Hojas"
         Top             =   1320
         Width           =   550
      End
      Begin VB.TextBox Text5 
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
         Left            =   9840
         MaxLength       =   4
         TabIndex        =   22
         ToolTipText     =   "Cantidad de Hojas del Certificado"
         Top             =   1580
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
         Left            =   11040
         Picture         =   "ASIGNARUTA.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Tabla de Títulos de Selección"
         Top             =   1980
         Width           =   550
      End
      Begin VB.CommandButton Command6 
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
         Left            =   10440
         TabIndex        =   20
         Top             =   1980
         Width           =   550
      End
      Begin VB.CommandButton Command5 
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
         Left            =   9840
         TabIndex        =   19
         ToolTipText     =   "Lista de Selección de Títulos de Certificados"
         Top             =   1980
         Width           =   550
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
         Height          =   585
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   " "
         Top             =   1980
         Width           =   9675
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
         Left            =   10440
         Picture         =   "ASIGNARUTA.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   17
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
         Left            =   11040
         Picture         =   "ASIGNARUTA.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Edita la Ruta Seleccionada"
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
         Left            =   11040
         TabIndex        =   15
         Top             =   1320
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
         Left            =   10440
         TabIndex        =   14
         Top             =   1320
         Width           =   550
      End
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
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   " "
         Top             =   1320
         Width           =   9675
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
         Width           =   9675
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
            Left            =   2340
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
            Left            =   2340
            Locked          =   -1  'True
            TabIndex        =   11
            TabStop         =   0   'False
            ToolTipText     =   " "
            Top             =   540
            Width           =   4100
         End
         Begin VB.TextBox Text10 
            Alignment       =   1  'Right Justify
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
            Left            =   6420
            Locked          =   -1  'True
            TabIndex        =   10
            TabStop         =   0   'False
            ToolTipText     =   " "
            Top             =   540
            Width           =   1380
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
            Width           =   2235
         End
         Begin VB.TextBox Text13 
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
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   8
            TabStop         =   0   'False
            Text            =   "Código"
            Top             =   270
            Width           =   2235
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
            Left            =   6420
            Locked          =   -1  'True
            TabIndex        =   7
            TabStop         =   0   'False
            Text            =   "Ord.Fabric."
            Top             =   270
            Width           =   1380
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
            Left            =   7785
            Locked          =   -1  'True
            TabIndex        =   6
            TabStop         =   0   'False
            Text            =   "N.Serie"
            Top             =   270
            Width           =   1755
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
            Left            =   7785
            Locked          =   -1  'True
            TabIndex        =   5
            ToolTipText     =   "Doble Click para Modificar"
            Top             =   540
            Width           =   1755
         End
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
      OleObjectBlob   =   "ASIGNARUTA.frx":075E
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "ASIGNARUTA.frx":0992
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
Attribute VB_Name = "ASIGNARUTA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias _
    "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    
Public ORDEX%

Private Sub Check1_Click(Index As Integer)
      
End Sub

Private Sub Command1_Click()
   'ACTUALIZAR
   Command1.Enabled = False
   '
   If TRIM$(Text10) = "" And TRIM$(Text1) = "" Then
       Call COMUNI("Falta Ingresar Lote-OF. o Ruta de Certificado")
       GoTo 99
   End If
   '
   With ABMCERTI09
     If UCase$(Text14) = "LOTE" Then
       .MSF1.TextMatrix(.MSF1.Row, 4) = Text1
       .MSF1.TextMatrix(.MSF1.Row, 5) = SAFETEXT$(ORDEX%)
       .MSF1.TextMatrix(.MSF1.Row, 6) = Text5
       
     Else
       .MSF.TextMatrix(.MSF.Row, 4) = Text15
       .MSF.TextMatrix(.MSF.Row, 5) = Text1
       .MSF.TextMatrix(.MSF.Row, 6) = SAFETEXT$(ORDEX%)
       .MSF.TextMatrix(.MSF.Row, 7) = Text5
       .MSF.COL = 4
       .MSF.Sort = 1

     End If
   End With
   CAMBIO% = 1

99 Command1.Enabled = True
   Hide

End Sub

Private Sub Command2_Click()
    '
    Command2.Enabled = False
    '
    If VALI_EXISTE_RUTA%(TRIM$(Text11), TRIM$(Text10), TRIM$(Text15), TRIM$(Text1)) > 0 Then GoTo 99
    '
   'ACTUALIZAR
   'AGREGA LA RUTA A LA GRILLA
   With ABMCERTI09
     If UCase$(Text14) = "LOTE" Then
       i& = .MSF1.Rows
       .MSF1.Rows = i& + 1
       .MSF1.TextMatrix(i&, 1) = .MSF1.TextMatrix(.MSF1.Row, 1)
       .MSF1.TextMatrix(i&, 2) = .MSF1.TextMatrix(.MSF1.Row, 2)
       .MSF1.TextMatrix(i&, 3) = .MSF1.TextMatrix(.MSF1.Row, 3)
       .MSF1.TextMatrix(i&, 4) = Text1
       .MSF1.TextMatrix(.MSF1.Row, 5) = SAFETEXT$(ORDEX%)
       .MSF1.TextMatrix(.MSF1.Row, 6) = Text5
       .MSF1.COL = 1
       .MSF1.Sort = 1

     Else
        i& = .MSF.Rows
       .MSF.Rows = i& + 1

       .MSF.TextMatrix(i&, 1) = .MSF.TextMatrix(.MSF.Row, 1)
       .MSF.TextMatrix(i&, 2) = .MSF.TextMatrix(.MSF.Row, 2)
       .MSF.TextMatrix(i&, 3) = .MSF.TextMatrix(.MSF.Row, 3)
       .MSF.TextMatrix(i&, 4) = .MSF.TextMatrix(.MSF.Row, 4)
       .MSF.TextMatrix(i&, 5) = Text1
       .MSF.TextMatrix(i&, 6) = SAFETEXT$(ORDEX%)
       .MSF.TextMatrix(i&, 6) = Text5
       .MSF.COL = 4
       .MSF.Sort = 1
     End If
   End With
    
   CAMBIO% = 1
99 Command2.Enabled = True
   Unload Me
End Sub
Function VALI_EXISTE_RUTA%(CODI$, LOTE$, NUMESERI$, RUTAX$)
   '
   VALI_EXISTE_RUTA% = 0
   If TRIM$(LOTE$) = "" And TRIM$(NUMESERI$) = "" And TRIM$(RUTAX$) = "" Then
     Call COMUNI("Nó Válido Ingresar - Faltan Datos")
     VALI_EXISTE_RUTA% = 1
     Exit Function
   End If
   '
   With ABMCERTI09
        If UCase$(Text14) = "LOTE" Then
            'VALIDO QUE LA RUTA DEL CODIGO Y LOTE NO SEA LA MISMA
            For i& = 1 To .MSF1.Rows - 1
              If UCase$(.MSF1.TextMatrix(i&, 1)) = UCase$(CODI$) Then
                If UCase$(.MSF1.TextMatrix(i&, 3)) = UCase$(LOTE$) Then
                  If TRIM$(.MSF1.TextMatrix(i&, 4)) = "" Then
                     Call Command1_Click 'SI EXISTE LA FILA Y ESTA VACIA GRABA NORMALMENTE SIN GENERAR NUEVO REGISTRO.
                     VALI_EXISTE_RUTA% = 1
                     Exit For
                  End If
                  If UCase$(.MSF1.TextMatrix(i&, 4)) = UCase$(RUTAX$) Then
                     Call COMUNI("La Ruta Seleccionada Yá Existe Para Mismo Código y Lote")
                     VALI_EXISTE_RUTA% = 1
                     Exit For
                  End If
                End If
              End If
            Next i&
        Else 'ORDEN DE FABRICACION
            For i& = 1 To .MSF.Rows - 1
              If UCase$(.MSF.TextMatrix(i&, 1)) = UCase$(CODI$) Then
                If UCase$(.MSF.TextMatrix(i&, 3)) = UCase$(LOTE$) Then
                   If UCase$(.MSF.TextMatrix(i&, 4)) = UCase$(NUMESERI$) Then
                     If TRIM$(.MSF.TextMatrix(i&, 5)) = "" Then
                       Call Command1_Click 'SI EXISTE LA FILA Y ESTA VACIA GRABA NORMALMENTE SIN GENERAR NUEVO REGISTRO.
                       VALI_EXISTE_RUTA% = 1
                       Exit For
                     End If
                     If UCase$(.MSF.TextMatrix(i&, 5)) = UCase$(RUTAX$) Then
                       Call COMUNI("La Ruta Seleccionada Yá Existe Para Mismo Código y O.F. y Número de Serie")
                       VALI_EXISTE_RUTA% = 1
                       Exit For
                     End If
                   End If
                End If
              End If
            Next i&
        End If
         
   End With

End Function
Private Sub Command3_Click()
      Command3.Enabled = False
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
   Command5.Enabled = False
   Call SELECHO("TITUCERT")
   ORDEX% = XVALO(VALACT1$("TITUCERT"))
   If ORDEX% < 1 Then GoTo 99
   '
   Text3 = VALACT2$("TITUCERT")
   '
99 Command5.Enabled = True
   
End Sub

Private Sub Command6_Click()
   Text3 = ""
   ORDEX% = 0
End Sub

Private Sub Command7_Click()
   Text16 = ""
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

Private Sub Text15_DblClick()
     
    Text15 = NORFA$
10  NUSEX$ = InputBox("Número de Serie Nuevo", NORFA$, "")
    Text15 = TRIM$(NUSEX$)
    If NUSEX$ = "" Then Exit Sub
    COFA$ = TRIM$(CODFAMIL$(CODINT%(Text11)))
    CONDI$ = "PACIENTE = '" & COFA$ & "' AND NUMEROSERIE = '" & NUSEX$ & "'"
    If CantRegistros("TANUMSE", CONDI$, "NOMESS") Then
       Call COMUNI("Número de Serie Existente Para Mismo Código Y Familia")
       GoTo 10
    End If
    '
    If VALIDASERIE_ENGRILLA%(Text15) > 0 Then
       Call COMUNI("Número de Serie Existente en Grilla")
       GoTo 10
    End If
    '
'    If VALIDANUMEROSERIE%(NORFA$, NUSEX$) > 0 Then
'       Call COMUNI("Número de Serie Existente")
'       Exit Sub
'    Else
'       OPX% = COMALTER%("Números de Serie\\Realmente Quiere Asignar el Número: " & NUSEX$ & "\Cancelar sin Grabar")
'       If OPX% < 1 Or OPX% > 1 Then Exit Sub
'       If OPX% = 1 Then
'            CI1% = CODINT%(Text11)
'            COFA$ = CODFAMIL$(CI1%)
'            FAMI$ = TRIM$(ECOARCH$("TAFAMIL", COFA$))
'            Call GRABAR_NSEREIE(NORFA$, NUSEX$, FAMI$)
'       End If
'    End If

End Sub
Function VALIDASERIE_ENGRILLA%(NUMEROSER$)
      VALIDASERIE% = 0
      For i% = 1 To ABMCERTI09.MSF.Rows - 1
         NUSX$ = TRIM(ABMCERTI09.MSF.TextMatrix(i%, 4))
         If UCase$(TRIM$(NUMEROSER$)) = NUSX$ Then
           VALIDASERIE% = 1
           Exit For
         End If
      Next i%
      '
      VALIDASERIE_ENGRILLA% = VALIDASERIE%
      '
End Function
