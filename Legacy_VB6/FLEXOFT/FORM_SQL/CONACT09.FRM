VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form CONACT09 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Actividades Programadas"
   ClientHeight    =   8250
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11910
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
   ScaleHeight     =   8250
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Actividades Vencidas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2190
      Left            =   60
      TabIndex        =   1
      Top             =   2000
      Width           =   10830
      Begin MSFlexGridLib.MSFlexGrid GRID1 
         Height          =   1880
         Left            =   60
         TabIndex        =   43
         Top             =   250
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   3307
         _Version        =   393216
         Rows            =   23
         Cols            =   4
         FixedCols       =   0
         BackColorFixed  =   8433840
         BackColorSel    =   -2147483643
         ForeColorSel    =   12582912
         BackColorBkg    =   16777215
         Redraw          =   -1  'True
         FocusRect       =   2
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Actividades Proyectadas Pendientes a Futuro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1110
      Left            =   60
      TabIndex        =   0
      Top             =   4200
      Width           =   10830
      Begin MSFlexGridLib.MSFlexGrid GRID2 
         Height          =   800
         Left            =   60
         TabIndex        =   44
         Top             =   250
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   1402
         _Version        =   393216
         Rows            =   23
         Cols            =   4
         FixedCols       =   0
         BackColorFixed  =   8433840
         BackColorSel    =   -2147483643
         ForeColorSel    =   12582912
         BackColorBkg    =   16777215
         Redraw          =   -1  'True
         FocusRect       =   2
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "E-Mails por Responder"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1920
      Left            =   60
      TabIndex        =   38
      Top             =   60
      Width           =   10830
      Begin VB.ListBox List3 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1320
         Left            =   105
         TabIndex        =   39
         Top             =   525
         Width           =   10605
      End
      Begin VB.Label Label16 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"CONACT09.frx":0000
         Height          =   285
         Left            =   105
         TabIndex        =   40
         Top             =   315
         Width           =   10600
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00004000&
      Height          =   5300
      Left            =   11025
      ScaleHeight     =   5235
      ScaleWidth      =   945
      TabIndex        =   31
      Top             =   0
      Width           =   1000
      Begin VB.CommandButton AGenRep 
         Height          =   675
         Left            =   105
         Picture         =   "CONACT09.frx":009E
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Acceso a Consultas Pre-Configuradas"
         Top             =   3675
         Width           =   640
      End
      Begin VB.CommandButton Command8 
         Height          =   675
         Left            =   105
         Picture         =   "CONACT09.frx":03A8
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Programar una Nueva Actividad"
         Top             =   2835
         Width           =   640
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00004000&
         Caption         =   "All"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   180
         TabIndex        =   41
         Top             =   1680
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton Command6 
         Height          =   675
         Left            =   105
         Picture         =   "CONACT09.frx":07EA
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Generar y Enviar Mensaje Interno"
         Top             =   2100
         Width           =   640
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   34
         Top             =   4515
         Width           =   640
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   500
            Left            =   0
            TabIndex        =   35
            Top             =   105
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Exit"
         Height          =   675
         Left            =   105
         Picture         =   "CONACT09.frx":0C2C
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Cierra y Abandona la Aplicación"
         Top             =   180
         Width           =   640
      End
      Begin VB.CommandButton Command1 
         Height          =   675
         Left            =   105
         Picture         =   "CONACT09.frx":0D76
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Refrescar la Lista de Seleccion"
         Top             =   945
         Width           =   640
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "CONACT09.frx":0EC0
         Top             =   4830
         Width           =   1515
      End
   End
   Begin VB.Frame Frame3 
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
      ForeColor       =   &H00808000&
      Height          =   2870
      Left            =   60
      TabIndex        =   2
      Top             =   5300
      Width           =   11800
      Begin VB.CommandButton Command11 
         Caption         =   "."
         Height          =   260
         Left            =   7820
         TabIndex        =   48
         ToolTipText     =   "Lista Completa de Contactos"
         Top             =   315
         Width           =   150
      End
      Begin VB.CommandButton Command10 
         Height          =   580
         Left            =   11080
         Picture         =   "CONACT09.frx":2DE2
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Anulacion de la Actividad (con motivo)"
         Top             =   2160
         Width           =   590
      End
      Begin VB.CommandButton Command9 
         Height          =   580
         Left            =   9880
         Picture         =   "CONACT09.frx":3224
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Generar un Presupuesto"
         Top             =   2160
         Width           =   590
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5565
         Style           =   2  'Dropdown List
         TabIndex        =   42
         Top             =   600
         Width           =   2415
      End
      Begin VB.CommandButton Command5 
         Height          =   580
         Left            =   9280
         Picture         =   "CONACT09.frx":336E
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Redactar y Enviar un Mail"
         Top             =   2160
         Width           =   590
      End
      Begin VB.CommandButton Command4 
         Height          =   580
         Left            =   8680
         Picture         =   "CONACT09.frx":37B0
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Abrir Anotador Historial de Seguimiento"
         Top             =   2160
         Width           =   590
      End
      Begin VB.CommandButton Command3 
         Height          =   580
         Left            =   10480
         Picture         =   "CONACT09.frx":3BF2
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Informar Cumplimiento de la Actividad"
         Top             =   2160
         Width           =   590
      End
      Begin VB.CommandButton Command2 
         Height          =   580
         Left            =   8080
         Picture         =   "CONACT09.frx":4034
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Guardar con Modificaciones"
         Top             =   2160
         Width           =   590
      End
      Begin VB.Frame Frame5 
         Caption         =   "St"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   910
         Left            =   80
         TabIndex        =   25
         Top             =   1850
         Width           =   500
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   510
            Left            =   50
            TabIndex        =   26
            Top             =   315
            Width           =   390
         End
      End
      Begin VB.Frame Frame4 
         Height          =   1870
         Left            =   8085
         TabIndex        =   14
         Top             =   210
         Width           =   3585
         Begin VB.ComboBox Combo3 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   735
            Style           =   2  'Dropdown List
            TabIndex        =   36
            Top             =   850
            Width           =   2730
         End
         Begin VB.CommandButton Command15 
            Caption         =   "."
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   3.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1140
            TabIndex        =   22
            Top             =   1470
            Width           =   150
         End
         Begin VB.TextBox Text12 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   105
            TabIndex        =   21
            Top             =   1470
            Width           =   1030
         End
         Begin VB.ComboBox Combo2 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   735
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   525
            Width           =   2730
         End
         Begin VB.Label Label25 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   2520
            TabIndex        =   24
            Top             =   1470
            Width           =   950
         End
         Begin VB.Label Label15 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1470
            TabIndex        =   23
            Top             =   1470
            Width           =   950
         End
         Begin VB.Label Label14 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   735
            TabIndex        =   20
            Top             =   210
            Width           =   2730
         End
         Begin VB.Label Label11 
            Caption         =   "    F.Proy:           F.Real:         F.Ctrl:"
            Height          =   225
            Left            =   105
            TabIndex        =   19
            Top             =   1250
            Width           =   3360
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            Caption         =   "Ctrl:"
            Height          =   225
            Left            =   50
            TabIndex        =   17
            Top             =   885
            Width           =   650
         End
         Begin VB.Label Label24 
            Alignment       =   1  'Right Justify
            Caption         =   "Resp:"
            Height          =   225
            Left            =   50
            TabIndex        =   16
            Top             =   570
            Width           =   650
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Gen:"
            Height          =   225
            Left            =   50
            TabIndex        =   15
            Top             =   255
            Width           =   650
         End
      End
      Begin VB.TextBox Text8 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   630
         TabIndex        =   4
         Text            =   " "
         Top             =   945
         Width           =   7365
      End
      Begin VB.TextBox Text9 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1470
         Left            =   630
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   1290
         Width           =   7365
      End
      Begin VB.Label Label10 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6230
         TabIndex        =   13
         Top             =   315
         Width           =   1575
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Oport:"
         Height          =   225
         Left            =   5400
         TabIndex        =   12
         Top             =   345
         Width           =   750
      End
      Begin VB.Label Label26 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3825
         TabIndex        =   11
         Top             =   315
         Width           =   1575
      End
      Begin VB.Label Label41 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   630
         TabIndex        =   10
         Top             =   315
         Width           =   990
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Nro:"
         Height          =   225
         Left            =   50
         TabIndex        =   9
         Top             =   345
         Width           =   550
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Prospecto:"
         Height          =   225
         Left            =   2730
         TabIndex        =   8
         Top             =   345
         Width           =   960
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Desc:"
         Height          =   225
         Left            =   50
         TabIndex        =   7
         Top             =   1320
         Width           =   550
      End
      Begin VB.Label Label36 
         Alignment       =   1  'Right Justify
         Caption         =   "Ref:"
         Height          =   225
         Left            =   50
         TabIndex        =   6
         Top             =   990
         Width           =   550
      End
      Begin VB.Label LABEL52 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   630
         TabIndex        =   5
         Top             =   600
         Width           =   4755
      End
   End
End
Attribute VB_Name = "CONACT09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
Private Sub AGenRep_Click()
    'se comento por que daba un error 5 en el procedimiento
    ' text4.gostfocus ***Queda para ver***
    AGenRep.Enabled = False '
    Call ACONSUL_PRE(TRIM$(UCase$(Name)))
    AGenRep.Enabled = True
    '
End Sub

Private Sub AGenRep_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   '
   If Button = 2 Then
      Call CONECTA_CONSUL_PRE(TRIM$(UCase$(Name)))
   End If
   '
End Sub

Sub Command1_Click()
    '
    FLIMI = FETEXCOR1$(DIPOST(HOY(HOS$)))
    USNU% = USNBR% Mod 100
    ' en la condicion falta poner que sean solo los propios o generales
    CONDI1$ = "REFE_ACTI <> '.' AND STAT_ACTI < 3 AND FHOR_PROG < '" & FLIMI & "' order by FHOR_PROG ASC, NUME_ACTI ASC"
       If Check1.Value <> 1 Then CONDI1$ = "USUA_ASIG = " & USNU% & " AND " + CONDI1$
    CONDI2$ = "REFE_ACTI <> '.' AND STAT_ACTI < 3 AND FHOR_PROG >= '" & FLIMI & "' order by FHOR_PROG ASC, NUME_ACTI ASC"
       If Check1.Value <> 1 Then CONDI2$ = "USUA_ASIG = " & USNU% & " AND " + CONDI2$
    Call CARGALISTA(GRID1, "ACTIPROY + BSEGCOM/ACTIPROY.CODI_CLI=BSEGCOM.CODI_CLI", "NUME_ACTI/A8/Nro; FHOR_PROG/D10/Fecha; REFE_ACTI/A30/Descripcion; BSEGCOM.CODI_CLI/A14/Codigo; BSEGCOM.RASO_CLI/A29/Prospecto; USUA_ASIG/F3/Rsp", CONDI1$)
    Call CARGALISTA(GRID2, "ACTIPROY + BSEGCOM/ACTIPROY.CODI_CLI=BSEGCOM.CODI_CLI", "NUME_ACTI/A8/Nro; FHOR_PROG/D10/Fecha; REFE_ACTI/A30/Descripcion; BSEGCOM.CODI_CLI/A14/Codigo; BSEGCOM.RASO_CLI/A29/Prospecto; USUA_ASIG/F3/Rsp", CONDI2$)
    '
    CONDI$ = "USUA_DESTI = " & USNU%
    If Check1.Value = 1 Then CONDI$ = "USUA_DESTI >= 0"
        If USNU% = 1 Then CONDI$ = "( " & CONDI$ & " OR USUA_DESTI IS NULL )"
    CONDI$ = CONDI$ & " AND (CONF_RECEP < '01/01/01' OR NUMAIRESPU = 0 OR NUMAIRESPU IS NULL) "
    CONDI$ = CONDI$ & " ORDER BY FHOR_RECEP ASC "
    Call CARGALISTA(List3, "BANDENTRA", "NUME_MAIL/F5; FHOR_RECEP/D17;ORIG_MAIL/A32;ASUN_MAIL/A34;USUA_DESTI/F5", CONDI$, "NOMESS")
    '
End Sub
            
Private Sub Command10_Click()
    Command3.Enabled = False
    If DERACCE%("ANUACTI", "Anulacion de Actividad Proyectada") = 2 Then
       If Label41 <> "" Then
          ANULACT09.Label41 = Label41
          ANULACT09.Show 1
          Label41 = ""
          Label26 = ""
          Label10 = ""
          Text8 = ""
          Text9 = ""
          Label14 = ""
          Text12 = ""
          Label15 = ""
          Label25 = ""
          Label3 = ""
       End If
    End If
    Command3.Enabled = True
End Sub

Private Sub Command11_Click()
    Call SELECHO("!INDOPORT/INDICE DE OPORTUNIDADES")
    CODIOPO$ = TRIM$(VALACT1$("!INDOPORT"))
    If CODIOPO$ <> "" Then
       Label10 = CODIOPO$
    End If
End Sub



Private Sub Command15_Click()
    Call SELECHO("SELFECHA")
    VDEV$ = VALACT1$("SELFECHA")
    If VDEV$ <> "" Then
        Text12.TEXT = VDEV$
    End If
End Sub

Private Sub Command2_Click()
   '
   Command2.Enabled = False
   If TRIM$(Label41) = "" Then GoTo 99
   '
   USUPRO% = XVALO(Left$(Label14, 3))
   If (USNBR% Mod 100) <> USUPRO% Then
      Call MENSERR(24, "Solo Usuario '" & TRIM$(Label14) & "'\Puede Grabar Datos Modificados.")
      GoTo 99
   End If
   '
   CONDI$ = "NUME_ACTI = '" & Label41 & "'"
   Call ACTUREGISTRO("ACTIPROY", "REFE_ACTI", CONDI$, Text8, RAFE&, "NOMESS")     ' REFERENCIA DE ACTIVIDAD
   Call ACTUREGISTRO("ACTIPROY", "DESC_ACTI", CONDI$, Text9, RAFE&, "NOMESS")     ' DESCRIPCION LARGA DE ACTIVIDAD
   USUASI% = XVALO(Left$(Combo2.TEXT, 3))
   Call ACTUREGISTRO("ACTIPROY", "USUA_ASIG", CONDI$, USUASI%, RAFE&, "NOMESS")   ' USUARIO ASIGNADO
   FEPROG$ = FETEXCOR1$(FECHANUM(Text12))
   Call ACTUREGISTRO("ACTIPROY", "FHOR_PROG", CONDI$, FEPROG$, RAFE&, "NOMESS")   ' FECHA Y HORA PROGRAMADO
   USUCRL% = XVALO(Left$(Combo3.TEXT, 3))
   Call ACTUREGISTRO("ACTIPROY", "USUA_CTRL", CONDI$, USUCRL%, RAFE&, "NOMESS")   ' USUARIO ASIGNADO
   Call ACTUREGISTRO("ACTIPROY", "CODOPORT", CONDI$, COPOR$, RAFE&, "NOMESS")     ' CODIGO DE OPORTUNIDAD
   '
   Call Command1_Click
   '
99 Command2.Enabled = True
   '
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    If Label41 <> "" Then
        APRUACT09.Show 1
    End If
    Command3.Enabled = True
End Sub

Private Sub Command4_Click()
   Command4.Enabled = False
   If Label26 <> "" Then
      ANOTASEG09.Label26 = Label26
      ANOTASEG09.Label7 = Label10
      ANOTASEG09.Show 1
   End If
   Command4.Enabled = True
End Sub

Private Sub Command5_Click()
    Command5.Enabled = False
    CODICLIX$ = Label26
    If CODICLIX$ <> "" Then
       FISECOM09.Label26 = CODICLIX$
       CONTAC$ = UCase$(TRIM$(Combo1.TEXT))
       If CONTAC$ <> "" Then
          For KKI% = 1 To FISECOM09.Text6.Count
             If UCase$(TRIM$(FISECOM09.Text6(KKI% - 1))) = CONTAC$ Then
                Call FISECOM09.Text6_DblClick(KKI% - 1)
                GoTo 99
             End If
          Next KKI%
       End If
       Call FISECOM09.Text5_DblClick
    End If
99  Command5.Enabled = True
End Sub

Private Sub Command6_Click()
    Command6.Enabled = False
    MENSAIN09.Show 1
    Command6.Enabled = True
End Sub

Private Sub Command7_Click()
    Hide
End Sub

Sub Command8_Click()
    '
    Command8.Enabled = False
    PROACT$ = INICRM09.PROXIACTI
    '
    PROYACT09.Label41 = PROACT$
    PROYACT09.Label14 = USERNAME$
    PROYACT09.Show 1
    '
    Command8.Enabled = True
End Sub

Private Sub Command9_Click()
    '
    CONEXI$ = "PRESUP07" ' CONECTA CON PRESUPUESTO Y LE PASA EL NUMERO DE PROSEPCTO
    If TRIM$(Label26) <> "" Then CONEXI$ = CONEXI$ + "/" + TRIM$(Label26)
    Call CONECRUN(CONEXI$, "Presupuestos de Clientes")
    '
End Sub

Private Sub Form_Load()
    Call CARLIUS(Combo2)
    Call CARLIUS(Combo3)
    If USNBR% Mod 100 = 1 Then
       Check1.Visible = True
    End If
End Sub

Private Sub GRID1_DblClick()
    IROW& = GRID1.Row
    NROACT$ = GRID1.TextMatrix(IROW&, 0)
    If NROACT$ <> "" Then
       Label41 = TRIM$(NROACT$)
    End If
End Sub

Private Sub GRID2_DblClick()
    IROW& = GRID2.Row
    NROACT$ = GRID2.TextMatrix(IROW&, 0)
    If NROACT$ <> "" Then
       Label41 = TRIM$(NROACT$)
    End If
End Sub

Private Sub Label26_Change()
    LABEL52 = VALOBADA("BSEGCOM", "RASO_CLI", "CODI_CLI = '" & Label26 & "'")
    CONDI$ = "CodOport <>'' AND CODI_CLI = '" & Label26 & "'"
    Call CARGALISEL("!INDOPORT", "OPORTUNI", "CodOport/A12;RefeOport/A48", CONDI$, "NOMESS")
End Sub

Private Sub Label41_Change()
       CONDI$ = "NUME_ACTI = '" & Label41 & "'"
       Label26 = VALOBADA("ACTIPROY", "Codi_Cli", CONDI$, "NOMESS")     ' CODIGO DE PROSPECTO
       Label10 = VALOBADA("ACTIPROY", "CODOPORT", CONDI$, "NOMESS")     ' CODIGO DE OPORTUNIDAD
         NOMECON$ = VALOBADA("ACTIPROY", "CONTACTO", CONDI$, "NOMESS") ' PERSONA DE CONTACTO
       Call FISECOM09.CARLICONTAC(Combo1, Label26, NOMECON$)
       Text8 = VALOBADA("ACTIPROY", "REFE_ACTI", CONDI$, "NOMESS")      ' REFERENCIA DE ACTIVIDAD
       Text9 = VALOBADA("ACTIPROY", "DESC_ACTI", CONDI$, "NOMESS")      ' DESCRIPCION LARGA DE ACTIVIDAD
         USUPRO% = XVALO(VALOBADA("ACTIPROY", "USUA_PROG", CONDI$, "NOMESS"))    ' USUARIO PROGRAMADO
       Label14 = USERNUNAM$(USUPRO%)
         USUASI% = XVALO(VALOBADA("ACTIPROY", "USUA_ASIG", CONDI$, "NOMESS"))   ' USUARIO ASIGNADO
       Text12 = FECHATEX$(CVINT(VALOBADA("ACTIPROY", "FHOR_PROG", CONDI$, "NOMESS")))    ' FECHA Y HORA PROGRAMADO
       Label15 = FECHATEX$(CVINT(VALOBADA("ACTIPROY", "FHOR_REAL", CONDI$, "NOMESS")))     ' FECHA FIN DE REITERACION
         USUCRL% = XVALO(VALOBADA("ACTIPROY", "USUA_CTRL", CONDI$, "NOMESS"))    ' USUARIO QUE CONTROLO
       Label25 = FECHATEX$(CVINT(VALOBADA("ACTIPROY", "FHOR_CTRL", CONDI$, "NOMESS")))     ' FECHA FIN DE REITERACION
       Label3 = XVALO(VALOBADA("ACTIPROY", "STAT_ACTI", CONDI$, "NOMESS"))    ' STATUS (0 = ACTIVA ; 3 = CERRADA)
       '   Call CREACAMPO("", "ACTIPROY", "COMENTAR", 12, 0)          ' COMENTARIO DE REALIZACION
       '
       Call CARLIUS(Combo2, USERNUNAM$(USUASI%))
       Call CARLIUS(Combo3, USERNUNAM$(USUCRL%))
End Sub

Private Sub List1_Click()
    List2.ListIndex = -1
    Label41 = TRIM$(Left$(List1.TEXT, 5))
End Sub

Private Sub List2_Click()
    List1.ListIndex = -1
    Label41 = TRIM$(Left$(List2.TEXT, 5))
End Sub

Private Sub List3_DblClick()
    NROMAI& = XVALO(Left$(List3.TEXT, 6))
    If NROMAI& > 0 Then
       SHINMAIL09.Label13 = TRIM$(Str$(NROMAI&))
       SHINMAIL09.Show 1
    End If
End Sub
