VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form RESERV07 
   BackColor       =   &H00EEEEC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reservas de Materiales"
   ClientHeight    =   6495
   ClientLeft      =   2040
   ClientTop       =   1305
   ClientWidth     =   5565
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   5565
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   0
      Left            =   3240
      ScaleHeight     =   375
      ScaleWidth      =   1110
      TabIndex        =   32
      Top             =   480
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox Picture14 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   6840
      Left            =   4680
      ScaleHeight     =   6810
      ScaleWidth      =   1035
      TabIndex        =   22
      Top             =   0
      Width           =   1065
      Begin VB.CommandButton Command11 
         Caption         =   "C.Dup"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   120
         TabIndex        =   31
         ToolTipText     =   "Control de Duplicados"
         Top             =   4320
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton Command10 
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
         Left            =   120
         Picture         =   "RESERV07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Listado de Componentes Asignados por O.F."
         Top             =   3480
         Width           =   650
      End
      Begin VB.CommandButton Command14 
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
         Left            =   120
         Picture         =   "RESERV07.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Configuración de Procesos y Formularios"
         Top             =   2625
         Width           =   650
      End
      Begin VB.CommandButton Command8 
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
         Left            =   120
         Picture         =   "RESERV07.frx":074C
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   1785
         Width           =   650
      End
      Begin VB.CommandButton Command6 
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
         Left            =   120
         Picture         =   "RESERV07.frx":0A56
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   950
         Width           =   650
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
         Height          =   640
         Left            =   120
         Picture         =   "RESERV07.frx":0D60
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   315
         Width           =   650
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00EEEEC0&
      Caption         =   "ACTUALIZACION"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3225
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   4425
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   1785
         ScaleHeight     =   615
         ScaleWidth      =   2295
         TabIndex        =   27
         Top             =   2310
         Width           =   2325
         Begin VB.CommandButton Command9 
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
            Left            =   0
            Picture         =   "RESERV07.frx":0EAA
            Style           =   1  'Graphical
            TabIndex        =   29
            ToolTipText     =   "Reservas por O/F con Faltantes"
            Top             =   0
            Width           =   600
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Reservas por O/F con Faltantes"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   680
            TabIndex        =   28
            Top             =   105
            Width           =   1590
         End
      End
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   1470
         ScaleHeight     =   615
         ScaleWidth      =   2295
         TabIndex        =   9
         Top             =   1845
         Width           =   2325
         Begin VB.CommandButton Command5 
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
            Left            =   0
            Picture         =   "RESERV07.frx":11B4
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Depuración de la Base de Datos"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   680
            TabIndex        =   11
            Top             =   0
            Width           =   1590
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   1100
         ScaleHeight     =   615
         ScaleWidth      =   2295
         TabIndex        =   18
         Top             =   1370
         Width           =   2325
         Begin VB.CommandButton Command7 
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
            Left            =   0
            Picture         =   "RESERV07.frx":14BE
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Control Según O/Fs en Curso"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   600
            TabIndex        =   20
            Top             =   0
            Width           =   1590
         End
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   735
         ScaleHeight     =   615
         ScaleWidth      =   2295
         TabIndex        =   12
         Top             =   895
         Width           =   2325
         Begin VB.CommandButton Command4 
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
            Left            =   0
            Picture         =   "RESERV07.frx":17C8
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Cancelación de Reservas"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   650
            TabIndex        =   14
            Top             =   0
            Width           =   1590
         End
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   420
         ScaleHeight     =   615
         ScaleWidth      =   2295
         TabIndex        =   15
         Top             =   420
         Width           =   2325
         Begin VB.CommandButton Command3 
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
            Left            =   0
            Picture         =   "RESERV07.frx":1AD2
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Generar  Reservas   por O/F"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   680
            TabIndex        =   17
            Top             =   0
            Width           =   1590
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EEEEC0&
      Caption         =   "CONSULTAS - LISTADOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   3480
      Width           =   4425
      Begin VB.OptionButton Option6 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEC0&
         Caption         =   "Totales por Material Reservado"
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
         Height          =   330
         Left            =   525
         TabIndex        =   7
         Top             =   420
         Width           =   3165
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00EEEEC0&
         Caption         =   "SALIDA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   210
         TabIndex        =   4
         Top             =   2100
         Width           =   4005
         Begin VB.OptionButton Option4 
            Appearance      =   0  'Flat
            BackColor       =   &H00EEEEC0&
            Caption         =   "Impresión"
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
            Height          =   225
            Left            =   1575
            TabIndex        =   5
            Top             =   250
            Width           =   1275
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
            Height          =   490
            Left            =   3340
            Picture         =   "RESERV07.frx":1DDC
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   90
            Width           =   645
         End
         Begin VB.OptionButton Option5 
            Appearance      =   0  'Flat
            BackColor       =   &H00EEEEC0&
            Caption         =   "Pantalla"
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
            Height          =   225
            Left            =   315
            TabIndex        =   6
            Top             =   250
            Value           =   -1  'True
            Width           =   1380
         End
      End
      Begin VB.OptionButton Option3 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEC0&
         Caption         =   "Por Material"
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
         Height          =   330
         Left            =   525
         TabIndex        =   3
         Top             =   1575
         Width           =   2955
      End
      Begin VB.OptionButton Option2 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEC0&
         Caption         =   "Por Número de O/Fabricación "
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
         Height          =   330
         Left            =   525
         TabIndex        =   2
         Top             =   1260
         Width           =   3060
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEC0&
         Caption         =   "Consulta General de Reservas por O/F"
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
         Height          =   330
         Left            =   525
         TabIndex        =   1
         Top             =   945
         Value           =   -1  'True
         Width           =   3795
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   420
      OleObjectBlob   =   "RESERV07.frx":20E6
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "RESERV07.frx":231A
      TabIndex        =   33
      Top             =   315
      Visible         =   0   'False
      Width           =   1590
   End
End
Attribute VB_Name = "RESERV07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Hide
End Sub

Private Sub Command10_Click()
   Command10.Enabled = False
   Call SELECHO("ECORDEP")
   NUORFAB$ = VALACT1$("ECORDEP")
   If TRIM$(NUORFAB$) = "" Then GoTo 99
   If EXISTE%("GENREP07.EXE") > 0 Then
     Call FINAL("?COMPOPOROF(" & NUORFAB$ & ")") ' ABRE LISTADO DE G.R. SEGUN LA O.F.
   End If
99 Command10.Enabled = True

End Sub

Private Sub Command11_Click()
  
   Command11.Enabled = False
   Screen.MousePointer = 11
   '
   Dim NORFAPE$()
   '
   Call ABREORFAB
   '
   SQLX$ = "SELECT DISTINCT IDSUBOR FROM RESERVA WITH(NOLOCK) ORDER BY IDSUBOR"
   
   Dim OFTMP As ADODB.Recordset
   Set OFTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
   '
   NARCHI2% = FILEOPEN%("C:\Flexoft\" + "Codigos_Duplicados.txt", 2, 128) ' ABRE EL ARCHIVO DE ESCRITURA
   '
   HH% = 0
   jj& = 0
   '
   With OFTMP
    Do While Not .EOF
     jj& = jj& + 1
     ReDim Preserve NORFAPE$(jj&)
     NORFAPE$(jj&) = SAFETEXT$(!idsubor)
     '
     CONDI$ = "IDSUBOR  = '" & NORFAPE$(jj&) & "'" 'VALIDO QUE LA ORDEN NO ESTE CERRADA
     ESTADO% = XVALO(VALOBADA("ORDEFABR", "STATUORF", CONDI$, "NOMESS"))
     If ESTADO% >= 3 Then GoTo 60
     '
     SQLX1$ = "SELECT Cod_Inte FROM RESERVA WHERE IdSubOr = '" & NORFAPE$(jj&) & "'"
     SQLX1$ = SQLX1$ + " ORDER BY COD_INTE"
     Dim CODITMP As ADODB.Recordset
     Set CODITMP = FLEXCONN.Execute(FILTSQL$(SQLX1$))
     With CODITMP
     Do While Not .EOF
       CI% = XVALO(!cod_inte)
       CONDI$ = "IDSUBOR ='" & NORFAPE$(jj&) & "' AND COD_INTE = " & CI%
       SIREPETI% = CantRegistros("RESERVA", CONDI$)
       If SIREPETI% > 1 Then
           HH% = HH% + 1
           TX$ = "Códigos Duplicados en Reservas de O/F: " & NORFAPE$(jj&)
           Print #NARCHI2%, TX$
           GoTo 60
       End If
       .MoveNext
     Loop
     End With
    
60   .MoveNext
    Loop
   End With
   OFTMP.Close
   Set OFTMP = Nothing
   '
   Close #NARCHI2%
   '
   Screen.MousePointer = 1
   '
   If HH% > 0 Then
      Call COMUNI("Se Encontraron Repeticiones en " & HH% & " Casos")
   Else
      Call COMUNI("No se Detectaron Repeticiones")
   End If
   '
   Command11.Enabled = True
   '
End Sub

Private Sub Command14_Click()
   Command14.Enabled = False
   Call COMUNI("No hay Opciones de Configuración\para esta Aplicación.")
   Command14.Enabled = True
End Sub

 Sub Command2_Click()
    OPCI% = 1
    If Option2.Value = True Then OPCI% = 2
    If Option3.Value = True Then OPCI% = 3
    If Option6.Value = True Then OPCI% = 4
    If OPCI% = 1 Then
        'Call GENRESDAT
        Call GENRESNEW(OPCI%)
'        Call CONECRUN("*CORESE1", "")
      ElseIf OPCI% = 2 Then
        Call GENRESNEW(OPCI%)
        'Call CONECRUN("*CORESE2", "")
      ElseIf OPCI% = 3 Then
        'Call GENRESDAT
        Call GENRESNEW(OPCI%)
        'Call CONECRUN("*CORESE3", "")
      ElseIf OPCI% = 4 Then
        Call GENRESNEW(OPCI%)
        
        'Call TOMATRES
      Else
        Hide
    End If
    Screen.MousePointer = 1
    '
End Sub

 Sub Command3_Click()
    Command3.Enabled = False
    Call GENRESER("")
    Command3.Enabled = True
End Sub

 Sub Command4_Click()
    Command4.Enabled = False
    Call ANURESER("")
    Command4.Enabled = True
End Sub

 Sub Command5_Click()
    Command5.Enabled = False
    '
    If DERACCE%("DEPURESE", "Depuración de Reservas de Materiales") < 2 Then GoTo 99
    '
    If COMALTER%("Este Proceso Verifica la Base de Datos\de Materiales Reservados y Anula Automaticamente\Todas las Reservas Correspondientes a O/F\Cerradas o Anuladas. Puede Ahora\Cancelar o Continuar\\Cancelar\Continuar") < 2 Then
      GoTo 99
    End If
    Screen.MousePointer = 11
    '
    Call ABREORFAB
    '
    CONDIX$ = " StatuOrf >= 3"
    FIN& = CantRegistros&("ORDEFABR", CONDIX$)
    '
    i& = 0
    SQLX$ = "SELECT IDSUBOR FROM ORDEFABR WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE " & CONDIX$
'    Dim OFTMP As ADODB.Recordset
'    Set OFTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    Set OFTMP = READSET(SQLX$)
    With OFTMP
      Do While Not .EOF
        i& = i& + 1
        Call TRACE(24, i&, FIN&)
        NordFab$ = SAFETEXT$(!idsubor)
        CONDI$ = "IdSubOr = '" & NordFab$ & "'"
        RegAfec& = 0
        Call BORRAREGISTROS("RESERVA", CONDI$, RegAfec&, "NOMESS")
        If RegAfec& > 0 And OPX% < 3 Then ' SI SE ENCONTARARON REGISTROS PARA BORRAR AVISA EL NUMERO DE OF AL CUAL SE LE BORRARON LAS RESERVAS
          OPX% = COMALTER%("Se han Borrado las Reservas de\la Orden de Fabricación Nro: " & NordFab$ & "\\Cancelar\Siguiente\Continuar Sin Visualizar Mensajes Informativos")
          If OPX% <= 1 Then Screen.MousePointer = 1: GoTo 99
        End If
        .MoveNext
      Loop
    End With
    OFTMP.Close
    Set OFTMP = Nothing
    '
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If InStr(EPAC$, "AMBEST") > 0 Then
        CONDI$ = "IdSubOr = 'OF-004895-1'"
        Call BORRAREGISTROS("RESERVA", CONDI$, RegAfec&, "NOMESS")
    End If
    '
    Screen.MousePointer = 1
    '
99  Command5.Enabled = True
End Sub

Private Sub Command6_Click()
    Command6.Enabled = False
    Call HELPONLINE("RESERVA")
    Command6.Enabled = True
End Sub

 Sub Command7_Click()
   Command7.Enabled = False
   Call CHECKRES
   Command7.Enabled = True
End Sub

Private Sub Command8_Click()
  Command8.Enabled = False
  Call ACCDIR("RESERVA")
  Command8.Enabled = True
End Sub

Private Sub Command9_Click()
  '\\\OT-06-0086-WAR-10/03/06///
  Dim OFCUR$()
  Dim STACT#(32767)
  Dim STDISP#(32767)
  '
  Command9.Enabled = False
  Call ABREORFAB
'1.- CARGAR VECTOR DE OF EN CURSO
  SQLX$ = "SELECT IDSUBOR FROM ORDEFABR WITH(NOLOCK) "
  SQLX$ = SQLX$ + " WHERE STATUORF <8 "
  'Set OFTMP = Ofabric.OpenRecordset(SQLX$, 4)
'  Dim OFTMP As ADODB.Recordset
'  Set OFTMP = New ADODB.Recordset
'  OFTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText

  Set OFTMP = READSET(SQLX$)
  With OFTMP
    Screen.MousePointer = 11
    On Error Resume Next
    .MoveLast
    If Err Then
      Call MENSERR(24, "No Hay O/F's Pendientes.")
      Screen.MousePointer = 1
      Command9.Enabled = True
      Exit Sub
    End If
    On Error GoTo 0
    J% = .RecordCount
    FIN& = J%
    ReDim OFCUR$(J%)
    CONTI% = 0
    .MoveFirst
    Do While .EOF = False
      k& = .AbsolutePosition
      Call TRACE(20, k&, FIN&)
      OFCUR$(CONTI%) = !idsubor
      CONTI% = CONTI% + 1
      .MoveNext
    Loop
    Screen.MousePointer = 1
  End With
  OFTMP.Close
  Set OFTMP = Nothing
  '
'2.- PARA CADA OF LEER LAS RESERVAS Y CARGAR ARCHIVO HASTA COLU. QT. RESERVADA
  Y$ = REGBLAN$("RESOFALT")
  Call ABRESERVA
  jj& = 0
  For JX% = 0 To J%
    SQLX$ = "SELECT CantRes, COD_INTE, IDSUBOR  FROM RESERVA WITH(NOLOCK) "
    SQLX$ = SQLX$ + " WHERE IdSubOr = '" & OFCUR$(JX%) & "' "
    SQLX$ = SQLX$ + " AND CantRes-CantCons > 0"
    'Set RESTMP = Stocks.OpenRecordset(SQLX$, 4)
'    Dim RESTMP As ADODB.Recordset
'    Set RESTMP = New ADODB.Recordset
'    RESTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
    Set RESTMP = READSET(SQLX$)
    With RESTMP
      On Error Resume Next
      .MoveFirst
      If Err Then
        On Error GoTo 0
        GoTo 22
      End If
      On Error GoTo 0
      Do While .EOF = False
        CARE1# = !CANTRES
        CI% = !cod_inte
        Call REPLA(Y$, !idsubor, 1, 12)
        Call REPLA(Y$, MKI$(CI%), 13, 2)
        Call REPLA(Y$, MKD$(CARE1#), 15, 8)
        jj& = jj& + 1
        Call GRAREG("RESOFALT", Y$, jj&)
        .MoveNext
      Loop
    End With
    RESTMP.Close
    Set RESTMP = Nothing
22 Next JX%
  Call SETULTREG("RESOFALT", jj&)
  '
'3.- SORT X Nº OF y X CODIGO
  Call FILESORT("RESOFALT", "RESOFALT", 2, 2, "ASC")
  Call FILESORT("RESOFALT", "RESOFALT", 1, 1, "ASC")
  '
'4.- STOCK ACTUAL DE CADA ITEM
  
  For JX% = 1 To NUMAS%
    STACT#(JX%) = STOCKACT(JX%)
    STDISP#(JX%) = STACT#(JX%)
  Next JX%
  Screen.MousePointer = 1
  '
'5.- BARRER EL ARCHIVO
  FIN& = ULTREG&("RESOFALT")
  For JN& = 1 To FIN&
    Call TRACE(24, JN&, FIN&)
    X$ = REGLEIDO$("RESOFALT", JN&)
    CR1# = CVD(Mid$(X$, 15, 8)) 'Q.RESERVADA
    CI% = CVI(Mid$(X$, 13, 2))
'    STD1# = STACT#(CI%) - CR1#
'    If STACT#(CI%) = 0 Then STD1# = 0
    FAL1# = 0
    SOB1# = 0
    If STDISP#(CI%) - CR1# > 0 Then
        SOB1# = STDISP#(CI%) - CR1# 'SOBRANTE
      Else
        FAL1# = Abs(STDISP#(CI%) - CR1#) 'FALTANTE
    End If
    Call REPLA(X$, MKD$(STACT#(CI%)), 23, 8)
    Call REPLA(X$, MKD$(STDISP#(CI%)), 31, 8)
    'AGREGAR FALTANTE Y SOBRANTE AL ARCHIVO
    Call REPLA(X$, MKD$(SOB1#), 39, 8)
    Call REPLA(X$, MKD$(FAL1#), 47, 8)
    Call GRAREG("RESOFALT", X$, JN&)
    STDISP#(CI%) = STDISP#(CI%) - CR1#
    If STDISP#(CI%) < 0 Then STDISP#(CI%) = 0
  Next JN&
  '
  If Option4 Then
    Call FINAL("*RESOFALT")
  ElseIf Option5 Then
    Call CONECRUN("*RESOFAP", "")
  End If
  '
  Command9.Enabled = True
  Screen.MousePointer = 1
  '\\\OT-06-0086-WAR-FIN///
End Sub

Private Sub Form_Load()
   '
   Call TRANSLABELS(Name)
   '
   UTILIZA_SKIN% = 1
   Call APLICACIONSKINS(Me)

   RESERV07.Option1.Value = True
End Sub
'
Sub TOMATRES()
    '
    Dim CARESE#(32767)
    FIN& = ULTREG&("RESERVA")
    If FIN& < 1 Then
      Call MENSERR(24, "No hay Reservas Registradas")
      Exit Sub
    End If
    '
    Screen.MousePointer = 11
    '
    For i& = 1 To FIN&
      Call TRACE(20, i&, FIN&)
      XX$ = REGLEIDO$("RESERVA", i&)
      CI% = CVI(Left$(XX$, 2))
      If CI% > 0 Then
        If CI% <= NUMAS% Then
          TORE# = CVD(Mid$(XX$, 19, 8))
          CACO# = CVD(Mid$(XX$, 27, 8))
          SIGNO% = Sgn(CVS(Mid$(XX$, 75, 4)))
          SARE# = TORE# - CACO#
          If SARE# * SIGNO% < 0 Then
              SARE# = 0
          End If
          Call REPLA(XX$, MKD$(SARE#), 35, 8)
          Call GRAREG("RESERVA", XX$, i&)
          CARESE#(CI%) = CARESE#(CI%) + SARE#
        End If
      End If
    Next i&
    '
    FIN& = ULTREG&("REPUREPA")
    For i& = 1 To FIN&
      Call TRACE(20, i&, FIN&)
      XX$ = REGLEIDO$("REPUREPA", i&)
      CI% = CVI(Mid$(XX$, 7, 2))
      If CI% > 0 Then
        If CI% <= NUMAS% Then
          TORE# = CVS(Mid$(XX$, 21, 4))
          CACO# = CVS(Mid$(XX$, 25, 4))
          '
          If UCase$(Unimed$(CIIX%)) = "GK" Then
            TORE# = TORE# / 1000
            CACO# = CACO# / 1000
          End If
          '
          SARE# = TORE# - CACO#
          If SARE# < 0.05 Then GoTo 29
          '
          NOREPARA& = CVS(Left$(XX$, 4))
          REREP& = REGBUS&("PRESREPA", 1, MKS$(NOREPARA&))
          If REREP& > 0 Then
            YY$ = REGLEIDO$("PRESREPA", REREP&)
            If CVI(Mid$(YY$, 123, 2)) > 0 Then GoTo 29 ' O/Rep Cerrada
            If CVS(Mid$(YY$, 125, 4)) < 1 Then GoTo 29 ' falta O/Venta
          End If
          '
          CARESE#(CI%) = CARESE#(CI%) + SARE#
        End If
      End If
29  Next i&
    '
    jj& = 0
    FIN& = NUMAS%: If FIN& < 1 Then FIN& = 1
    '
    For i& = 1 To NUMAS%
      II$ = REGLEIDO$("INVERT", i& + 1)
      CI% = i&
      If CI% > 0 Then
        If Abs(CARESE#(CI%)) > 0.05 Then
          stock# = STOCKACT(CI%)
          XX$ = REGBLAN$("TOMATRES")
          Call REPLA(XX$, MKI$(CI%), 1, 2)
          Call REPLA(XX$, MKD$(CARESE#(CI%)), 3, 8)
          Call REPLA(XX$, MKD$(stock#), 11, 8)
          FALTA# = CARESE#(CI%) - stock#
          If FALTA# < 0 Then FALTA# = 0
          Call REPLA(XX$, MKD$(FALTA#), 35, 8)
          OBSE$ = "      "
          If FALTA# > 0 Then OBSE$ = "CRIT. "
          If stock# - STOMIN(CI%) < 0 Then OBSE$ = "CRIT. "
          If stock# < 0 Then OBSE$ = "NEG.  "
          Call REPLA(XX$, OBSE$, 43, 6)
          jj& = jj& + 1
          Call GRAREG("TOMATRES", XX$, jj&)
        End If
      End If
    Next i&
    '
    Call SETULTREG("TOMATRES", jj&)
    Call CIERRARCH("TOMATRES")
    Call FILESORT("TOMATRES", "", 1, 1, "ASC")
    '
    Screen.MousePointer = 1
    '
    If Option4.Value = True Then
         Call FINAL("*LITOTRE")       ' impresion
       Else
         Call FINAL("*COTOTRE")       ' pantalla
    End If
    '
End Sub
'

Sub GENRESNEW(OPCI%)
    '
    Call ABRESTOCKS
    Dim obj As New RECORXET
    If OPCI% = 1 Then
        Screen.MousePointer = 11
        Campos$ = "O.F/A16;Código/A20;Descripción/A48;Saldo/F12.1;Uso/F9.1"
        Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, Campos$, GRILLAGRAL, 1)
        Call CARGA_ENCABEZADO(GRILLAGRAL.MSF_2, Campos$, GRILLAGRAL, 1)
        'CARGO EL RECORDSET
        Set rs1 = obj.RS_RESERVAS_ENVGR
        Call Carga_MSF_Recordset(rs1, Campos$, GRILLAGRAL.GRID, 1)
        Set obj = Nothing
        rs1.Close
        Set rs1 = Nothing
        Screen.MousePointer = 1
        '*******************************************
        Screen.MousePointer = 1
        If GRILLAGRAL.GRID.Rows < 2 Then
          Call COMUNI("No se Registraron Reservas")
          Exit Sub
        End If
        TX$ = "Consulta General de Reservas por O/F"
        GRILLAGRAL.Caption = TX$
    ElseIf OPCI% = 2 Then
        Screen.MousePointer = 11
        Campos$ = "Código/A16;Descripción/A48;O.F/A12;Fecha/D8;Cant.Reser/f12.1;Cant.Cons./f12.1;Saldo/f12.1;Lote/A12;Dep./F5.1;Uso/F9.1"
        Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, Campos$, GRILLAGRAL, 1)
        Call CARGA_ENCABEZADO(GRILLAGRAL.MSF_2, Campos$, GRILLAGRAL, 1)
        'CARGO EL RECORDSET
        Set rs1 = obj.RS_RESERVAS_ENV
        Call Carga_MSF_Recordset(rs1, Campos$, GRILLAGRAL.GRID, 1)
        Set obj = Nothing
        rs1.Close
        Set rs1 = Nothing
        Screen.MousePointer = 1
        '*******************************************
        Screen.MousePointer = 1
        If GRILLAGRAL.GRID.Rows < 2 Then
          Call COMUNI("No se Registraron Reservas")
          Exit Sub
        End If
        TX$ = "Por Número de O/Fabricación "
        GRILLAGRAL.Caption = TX$
        '*******************************************
    ElseIf OPCI% = 3 Then
        Screen.MousePointer = 11
        AgregaCliente% = 0
        Campos$ = "Código/A16;Descripción/A48;O.F/A12;Fecha/D8;Cant.Reser/f12.1;Cant.Cons./f12.1;Saldo/f12.1;Lote/A12;Dep./F5.1;Uso/F9.1"
        If InStr(1, EMPREAC$, "ENVAPLAST") > 0 Then
           AgregaCliente% = 1
           Campos$ = "Código/A16;Descripción/A32;O.F/A12;Fecha/D8;Cant.Reser/f12.1;Cant.Cons./f12.1;Saldo/f12.1;Lote/A12;Dep./F5.1;Uso/F9.1;R.Social/A26"
           Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, Campos$, GRILLAGRAL, 1)
           Call CARGA_ENCABEZADO(GRILLAGRAL.MSF_2, Campos$, GRILLAGRAL, 1)
           'CARGO EL RECORDSET
           Set rs1 = obj.RS_RESERVAS_C_CTE_ENV(1) ' POR ARTICULO CON CLIENTE
        Else
           Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, Campos$, GRILLAGRAL, 1)
           Call CARGA_ENCABEZADO(GRILLAGRAL.MSF_2, Campos$, GRILLAGRAL, 1)
           'CARGO EL RECORDSET
           Set rs1 = obj.RS_RESERVAS_ENV(1) ' POR ARTICULO
        End If
        Call Carga_MSF_Recordset(rs1, Campos$, GRILLAGRAL.GRID, 1)
        Set obj = Nothing
        rs1.Close
        Set rs1 = Nothing
        Screen.MousePointer = 1
        '*******************************************
        Screen.MousePointer = 1
        If GRILLAGRAL.GRID.Rows < 2 Then
          Call COMUNI("No se Registraron Reservas")
          Exit Sub
        End If
        TX$ = "Reservados Por Material"
        GRILLAGRAL.Caption = TX$
        '*******************************************
    ElseIf OPCI% = 4 Then
        Screen.MousePointer = 11
        Campos$ = "Código/A16;Descripción/A20;Cant.Reser/f12.1;Stock/f12.1;Falta/f12.1"
        Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, Campos$, GRILLAGRAL, 1)
        Call CARGA_ENCABEZADO(GRILLAGRAL.MSF_2, Campos$, GRILLAGRAL, 1)
        'CARGO EL RECORDSET
        SQLX$ = "select COD_INTE,Cod_Exte ,descrip, sum(CantRes) as reserva FROM RESERVA WITH(NOLOCK)"
        SQLX$ = SQLX$ + " Where COD_INTE > 0 AND CantRes - CantCons > 0"
        SQLX$ = SQLX$ + " group by COD_INTE, Cod_Exte, Descrip "
        SQLX$ = SQLX$ + " order by Cod_Exte,Descrip"
        J& = 0
        Set rs = READSET(SQLX$)
        With rs
          Do While Not .EOF
            J& = J& + 1
            GRILLAGRAL.GRID.Rows = J& + 1
            GRILLAGRAL.GRID.TextMatrix(J&, 1) = SAFETEXT$(!Cod_Exte)
            GRILLAGRAL.GRID.TextMatrix(J&, 2) = SAFETEXT$(!Descrip)
            GRILLAGRAL.GRID.TextMatrix(J&, 3) = FORMATNUM$(XVALO(!Reserva), "F12.1")
            STOCKX = STOCKACT(XVALO(!cod_inte))
            GRILLAGRAL.GRID.TextMatrix(J&, 4) = FORMATNUM$(STOCKX, "F12.1")
            FALTAX = STOCKX - XVALO(!Reserva): If FALTAX >= 0 Then FALTAX = 0
            GRILLAGRAL.GRID.TextMatrix(J&, 5) = FORMATNUM$(Abs(FALTAX), "F12.1")
            .MoveNext
          Loop
        End With
        rs.Close
        Set rs = Nothing
        Screen.MousePointer = 1
        '*******************************************
        Screen.MousePointer = 1
        If GRILLAGRAL.GRID.Rows < 2 Then
          Call COMUNI("No se Registraron Reservas")
          Exit Sub
        End If
        TX$ = "Totales por Material Reservado"
        GRILLAGRAL.Caption = TX$
        '*******************************************
    End If
    GRILLAGRAL.mnuMenu1.Caption = "Imprimir"
    GRILLAGRAL.mnuMenu1.Visible = True
    GRILLAGRAL.GRID.ZOrder 0
    GRILLAGRAL.TXTBUSCAR = ""

    GRILLAGRAL.Show 1
    Dim MSF As msFlexGrid
    If GRILLAGRAL.MENU1 > 0 Then
        Set MSF = GRILLAGRAL.GRID
        If GRILLAGRAL.MSF_2.Rows > 1 Then Set MSF = GRILLAGRAL.MSF_2
        Call CARGA_TABLA_IMPRE(MSF, Campos$, "DEPUNEGA", TX$, "TABLA_IMPRE", RET_NAMECONS$, 1)
        Call FINAL("?" & RET_NAMECONS$)
        GRILLAGRAL.MENU1 = 0
        Screen.MousePointer = 1
    End If
    Set obj = Nothing
End Sub

'CODIGO ANTERIOR
    '
Sub GENRESDAT()
    '
    Call ABRESTOCKS
    '
    SQLX$ = " SELECT COD_INTE, IDSUBOR, COD_INTE, FechRes, CantRes, CANTCONS, Saldo_Entre, IdenLote, Deposito, UsoUnit FROM RESERVA WITH(NOLOCK) "
    SQLX$ = SQLX$ + " WHERE CantRes-CantCons > 0 "
    SQLX$ = SQLX$ + " ORDER BY IdSubOr "
    '
    'Set RESERTMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
'    Dim RESERTMP As ADODB.Recordset
'    Set RESERTMP = New ADODB.Recordset
'    RESERTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
    Set RESERTMP = READSET(SQLX$)
    With RESERTMP
      On Error Resume Next
      .MoveFirst
      If Err Then Call MENSERR(24, " No Existen Reservas efectuadas. ")
      On Error GoTo 0
      jj& = 0
      FIN& = .RecordCount
      Do While (.EOF = False)
        Y$ = REGBLAN("RESERVA")
        '
        Call REPLA(Y$, MKI$(XVALO(!cod_inte)), 1, 2)
        Call REPLA(Y$, SAFETEXT$(!idsubor), 3, 12)
        Call REPLA(Y$, MKI$(XVALO(!cod_inte)), 15, 2)
        Call REPLA(Y$, MKI$(CVINT(!FechRes)), 17, 2)
        Call REPLA(Y$, MKD$(XVALO(!CANTRES)), 19, 8)
        Call REPLA(Y$, MKD$(XVALO(!CANTCONS)), 27, 8)
           TORE# = XVALO(!CANTRES)
           CACO# = XVALO(!CANTCONS)
        Call REPLA(Y$, MKD$(TORE# - CACO#), 35, 8)
        Call REPLA(Y$, SAFETEXT$(!idenlote), 43, 16)
        Call REPLA(Y$, Chr$(XVALO(!DEPOSITO)), 60, 1)
        Call REPLA(Y$, MKS$(XVALO(!UsoUnit)), 75, 4)
        '
        jj& = jj& + 1
        Call GRAREG("RESERVA", Y$, jj&)
        '
      .MoveNext
      Loop
    End With
    RESERTMP.Close
    Set RESERTMP = Nothing
    Call SETULTREG("RESERVA", jj&)
    Call CIERRARCH("RESERVA")
    '
End Sub

    

