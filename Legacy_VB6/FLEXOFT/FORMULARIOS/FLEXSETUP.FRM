VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FLEXSETUP 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración General de Funcionamiento FLEXOFT"
   ClientHeight    =   7755
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   6300
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7755
   ScaleWidth      =   6300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame8 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Codigo de Barras"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   210
      TabIndex        =   30
      Top             =   6840
      Width           =   4950
      Begin VB.OptionButton Option6 
         BackColor       =   &H00C0E0FF&
         Caption         =   "BarCode 128"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3000
         TabIndex        =   32
         Top             =   360
         Width           =   1425
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Interleaved 2 of 5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   31
         Top             =   360
         Value           =   -1  'True
         Width           =   1905
      End
   End
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2280
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   28
      Top             =   1320
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Mail"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   5355
      Picture         =   "FLEXSETUP.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Generar y Enviar Mail Circular"
      Top             =   4560
      Width           =   750
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
      Height          =   750
      Left            =   5355
      Picture         =   "FLEXSETUP.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Administrar Licencias"
      Top             =   5460
      Width           =   750
   End
   Begin VB.CommandButton Command5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   5355
      Picture         =   "FLEXSETUP.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Ruteo de Bases de Datos"
      Top             =   6300
      Width           =   750
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Tr"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   5355
      Picture         =   "FLEXSETUP.frx":0896
      TabIndex        =   24
      ToolTipText     =   "Traduccion de Leyendas"
      Top             =   3675
      Width           =   750
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Copia de Respaldo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2205
      Left            =   210
      TabIndex        =   15
      Top             =   4560
      Width           =   4950
      Begin VB.CommandButton Command3 
         Caption         =   "Confeccionar Backup Ahora"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   840
         TabIndex        =   23
         Top             =   1680
         Width           =   3375
      End
      Begin VB.CheckBox Check5 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Suprime Backup Automático en esta Terminal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   420
         TabIndex        =   22
         Top             =   1320
         Width           =   4400
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Ventana"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Left            =   2520
         TabIndex        =   19
         Top             =   240
         Width           =   2220
         Begin VB.OptionButton Option4 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Normal"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   210
            TabIndex        =   21
            Top             =   335
            Value           =   -1  'True
            Width           =   1065
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Minimizada"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   210
            TabIndex        =   20
            Top             =   650
            Width           =   1485
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Utilitarios"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Left            =   210
         TabIndex        =   16
         Top             =   240
         Width           =   2220
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0E0FF&
            Caption         =   "PK-ZIP"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   210
            TabIndex        =   18
            Top             =   650
            Width           =   1065
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0E0FF&
            Caption         =   "ARJ"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   210
            TabIndex        =   17
            Top             =   335
            Value           =   -1  'True
            Width           =   1065
         End
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Estación de Trabajo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1320
      Left            =   210
      TabIndex        =   10
      Top             =   120
      Width           =   4950
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
         Height          =   300
         Left            =   525
         TabIndex        =   12
         Top             =   840
         Width           =   4215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Denominación:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   210
         TabIndex        =   14
         Top             =   525
         Width           =   1695
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ident:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1890
         TabIndex        =   13
         Top             =   460
         Width           =   1695
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
         Height          =   300
         Left            =   3675
         TabIndex        =   11
         Top             =   420
         Width           =   1065
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Entrada al Sistema"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2940
      Left            =   210
      TabIndex        =   3
      Top             =   1560
      Width           =   4950
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Locales (Terminal actual)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1530
         Left            =   210
         TabIndex        =   6
         Top             =   1200
         Width           =   4530
         Begin VB.CheckBox Check4 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Generador de Reportes Residente"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   315
            TabIndex        =   9
            Top             =   1050
            Width           =   3690
         End
         Begin VB.CheckBox Check3 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Revisa Pedidos a Despachar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   315
            TabIndex        =   8
            Top             =   735
            Width           =   3450
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Controla Stocks Negativos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   315
            TabIndex        =   7
            Top             =   420
            Width           =   3555
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Generales"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   210
         TabIndex        =   4
         Top             =   240
         Width           =   4530
         Begin VB.CheckBox Check2 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Presenta Datos de Impresora Activa"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   315
            TabIndex        =   5
            Top             =   420
            Width           =   3690
         End
      End
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   5355
      Picture         =   "FLEXSETUP.frx":0BA0
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2100
      Width           =   750
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   5355
      Picture         =   "FLEXSETUP.frx":0EAA
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1260
      Width           =   750
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
      Height          =   750
      Left            =   5355
      Picture         =   "FLEXSETUP.frx":11B4
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   420
      Width           =   750
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   750
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "FLEXSETUP.frx":12FE
      TabIndex        =   29
      Top             =   0
      Visible         =   0   'False
      Width           =   1590
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   3600
      OleObjectBlob   =   "FLEXSETUP.frx":1378
      Top             =   1200
   End
   Begin VB.Image Image1 
      Height          =   300
      Left            =   3600
      Picture         =   "FLEXSETUP.frx":15AC
      Stretch         =   -1  'True
      ToolTipText     =   "Doble-Click para Bloquear Posición"
      Top             =   4680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   4755
      Picture         =   "FLEXSETUP.frx":16F6
      Top             =   7150
      Width           =   2010
   End
End
Attribute VB_Name = "FLEXSETUP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CAMBIADENO%

Private Sub Command1_Click()
    '
    Call CIERRARCH("*.*")
    Hide
    '
End Sub

Private Sub Command2_Click()
    '
    If CAMBIADENO% > 0 Then
      If COMALTER%("ATENCION !!!\   \Para Efectivizar el Cambio de Denominación\de la Estación de Trabajo\Deberá Re-habilitar la Terminal\\Cancelar Cambio\Registrar") <> 2 Then
        Exit Sub
      End If
    End If
    '
    If TRIM$(Text1) = "" Then
      Call MENSERR(24, "Ingrese Nombre de Terminal")
    End If
    '
    PREIMPRE$ = "NO"
    If Check2.Value = 1 Then PREIMPRE$ = "SI"
    RESIWIN$ = "NO": RSW% = 0
    If Check4.Value = 1 Then RESIWIN$ = "SI": RSW% = 1
    CONEGATI$ = "NO"
    If CHECK1.Value = 1 Then CONEGATI$ = "SI"
    REVIPEPE$ = "NO"
    If Check3.Value = 1 Then REVIPEPE$ = "SI"
    '
    If REVIPEPE$ = "SI" Then
        If EXISTE%("ADMIPED.EXE") < 1 Then
            Call OPNOIN("ADMIPED.EXE - Revisión Pedidos a Despachar")
            Check3.Value = 0
            Exit Sub
        End If
    End If
    '
    Call GRACONTROL("FLEXOFT", "PREIMPRE", PREIMPRE$)
    Call GRACONTROL(IDENTER$, "RESIWIN", RESIWIN$)
    Call GRACONTROL(IDENTER$, "CONEGATI", CONEGATI$)
    Call GRACONTROL(IDENTER$, "REVIPEPE", REVIPEPE$)
    '
    SUPRIBA$ = "NO": If Check5.Value = 1 Then SUPRIBA$ = "SI"
    Call GRACONTROL(IDENTER$, "SUPRIBAK", SUPRIBA$)
    '
    VENBA$ = "NORMAL": If Option3.Value = True Then VENBA$ = "MINIMIZ"
    Call GRACONTROL(IDENTER$, "VENTBAK", VENBA$)
    '
    UTIBA$ = "ARJ": If Option2.Value = True Then UTIBA$ = "PKZIP"
    Call GRACONTROL$(IDENTER$, "UTIBACK", UTIBA$)
    '
    CODBAR$ = "INTE2OF5": If Option6.Value = True Then CODBAR$ = "COBAR128"
    Call GRACONTROL$("", "CODIBAR", CODBAR$)
    '
    RESIWIX$ = "RESIWIN"
    If RSW% < 1 Then
       RESIWIX$ = "       "
       Call SAVEFILE(LUCOM$ + "RCONTROL.DRV", "CLOSETAB")
    End If
    Call SAVEFILE("C:\FLEXOFT\RESIWIN.DRV", RESIWIX$)
    '
    Call GRARAM(306, TRIM$(Str$(RSW%)))
    Call BAJRAM
    '
    If CAMBIADENO% > 0 Then
      NOMTER$ = Text1
      N1% = FILEOPEN%(LUCOM$ + "IDENTER.DRV", 0, 40)
      Dim IDXX As String * 40
      J1$ = Space$(38)
      Call REPLA(J1$, AJUSTI$(Text1, 32), 1, 32)
      Call REPLA(J1$, AJUSTI$(Label1, 6), 33, 6)
      J2$ = ENCRIPTA$(J1$)
      SUDI% = 0
      For i% = 1 To 38
         SUDI% = SUDI% + Asc(Mid$(IDXX$, i%, 1))
      Next i%
      IDXX$ = J2$ + MKI$(SUDI%)
      Put #N1%, 1, IDXX$
      Close #N1%
    End If
    '
    Call CIERRARCH("*.*")
    Hide
    '
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    '
    If CONTROL$(IDENTER$, "SUPRIBAK") = "SI" Then
      SUPRIBA$ = "NO": Call GRACONTROL(IDENTER$, "SUPRIBAK", SUPRIBA$)
      Call GENBACKUP
      SUPRIBA$ = "SI": Call GRACONTROL(IDENTER$, "SUPRIBAK", SUPRIBA$)
    Else
      Call GENBACKUP
    End If
    '
    Command3.Enabled = True
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    TRALECAR.Show 1
    Command4.Enabled = True
End Sub

Private Sub Command5_Click()
   Command5.Enabled = False
   RUTBASES.Show 1
   Command5.Enabled = True
End Sub

Private Sub Command6_Click()
    Command6.Enabled = False
    SELICEN2.Show 1
    Command6.Enabled = True
End Sub

Private Sub Command8_Click()
    '
    Command8.Enabled = False
    '
    If DERACCE%("CFGEMAIL", "Configurar Cuenta de Correo Electronico") < 2 Then
     GoTo 99
    End If
    '
    STPFMAIL.Show 1
    '
99  Command8.Enabled = True
    '
End Sub

Private Sub Form_Load()
    '
    If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
'      Picture16.Visible = True
'      Picture16.ZOrder 0
      Me.BackColor = &H8000000F
      Me.Frame1.BackColor = &H8000000F
      Me.Frame2.BackColor = &H8000000F
      Me.Frame3.BackColor = &H8000000F
      Me.Frame4.BackColor = &H8000000F
      Me.Frame5.BackColor = &H8000000F
      Me.Frame6.BackColor = &H8000000F
      Me.Frame7.BackColor = &H8000000F
      Me.Frame8.BackColor = &H8000000F
      Me.Option1.BackColor = &H8000000F
      Me.Option2.BackColor = &H8000000F
      Me.Option3.BackColor = &H8000000F
      Me.Option4.BackColor = &H8000000F
      Me.Option5.BackColor = &H8000000F
      Me.Option6.BackColor = &H8000000F
      Me.CHECK1.BackColor = &H8000000F
      
      Me.Check2.BackColor = &H8000000F
      Me.Check3.BackColor = &H8000000F
      Me.Check4.BackColor = &H8000000F
      Me.Check5.BackColor = &H8000000F
      Me.Option6.BackColor = &H8000000F
      Me.Image2.Visible = False
'      MARCOBARRA.Top = 3600
'      Set MARCOBARRA.Container = Me.Picture16
'      MARCOBARRA.ZOrder 0
    End If
    '
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me)
    If CONTROL$("FLEXOFT", "PREIMPRE") = "SI" Then
        Check2.Value = 1
    End If
    '
    If CONTROL$(IDENTER$, "RESIWIN") = "SI" Then
        Check4.Value = 1
    End If
    RSWX$ = TRIM$(UCase$(LOADFILE$("C:\FLEXOFT\RESIWIN.DRV")))
    If Left$(RSWX$, 7) = "RESIWIN" Then
        Check4.Value = 1
    End If
    '
    If CONTROL$(IDENTER$, "CONEGATI") = "SI" Then
        CHECK1.Value = 1
    End If
    '
    If CONTROL$(IDENTER$, "REVIPEPE") = "SI" Then
        Check3.Value = 1
    End If
    '
    N1% = FILEOPEN%(LUCOM$ + "IDENTER.DRV", 0, 40)
    Dim IDXX As String * 40
    Get #N1%, 1, IDXX$
    IDENCRIP$ = Left$(IDXX$, 38)
    J1$ = DESENCRI$(Left$(IDXX$, 38))
    Label1 = TRIM$(Mid$(J1$, 33, 6))
    Text1 = TRIM$(Mid$(J1$, 1, 32))
    CAMBIADENO% = 0
    Close #N1%
    '
    For KKI% = 1 To 6
      ICAR% = Asc(Mid$(Label1, KKI%, 1))
      If ICAR% < 48 Or ICAR% > 57 Then
        IDTRX$ = RANDSTRING$(-6)
        Label1 = IDTRX$
        CAMBIADENO% = 1
        GoTo 20
      End If
    Next KKI%
    '
20  If CONTROL$(IDENTER$, "SUPRIBAK") = "SI" Then
       Check5.Value = 1
    End If
    '
    If CONTROL$(IDENTER$, "VENTBAK") = "MINIMIZ" Then
       Option3.Value = True
    End If
    '
    If CONTROL$(IDENTER$, "UTIBACK") = "PKZIP" Then
       Option2.Value = True
    End If
    '
    If CONTROL$("", "CODIBAR") = "COBAR128" Then
       Option6.Value = True
    End If
    '
End Sub

Private Sub Text1_Change()
   CAMBIADENO% = 1
End Sub
