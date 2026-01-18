VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Begin VB.Form FRMPROFESO1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ficha de Profesores"
   ClientHeight    =   6630
   ClientLeft      =   -15
   ClientTop       =   675
   ClientWidth     =   14565
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   14565
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CMD1 
      Left            =   7080
      Top             =   4080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Height          =   6600
      Left            =   7320
      TabIndex        =   28
      Top             =   0
      Width           =   7155
      Begin VB.PictureBox Picture2 
         Height          =   3000
         Left            =   3900
         ScaleHeight     =   2940
         ScaleWidth      =   2940
         TabIndex        =   40
         Top             =   2600
         Width           =   3000
         Begin VB.Image Image1 
            Height          =   3000
            Left            =   0
            Stretch         =   -1  'True
            Top             =   0
            Width           =   3000
         End
      End
      Begin VB.CommandButton Command9 
         Caption         =   "...."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   160
         Picture         =   "FRMPROFESO1.frx":0000
         TabIndex        =   39
         Top             =   5760
         Width           =   540
      End
      Begin VB.TextBox TXTRUTAFIRMA 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         IMEMode         =   3  'DISABLE
         Left            =   720
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   38
         Top             =   5760
         Width           =   6255
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   1
         Left            =   900
         Sorted          =   -1  'True
         TabIndex        =   8
         Text            =   "Combo1"
         Top             =   360
         Width           =   1320
      End
      Begin VB.TextBox Text12 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2250
         TabIndex        =   9
         Top             =   360
         Width           =   4695
      End
      Begin VB.TextBox Text11 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2250
         TabIndex        =   10
         Top             =   720
         Width           =   4695
      End
      Begin VB.TextBox Text10 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2250
         TabIndex        =   11
         Top             =   1080
         Width           =   4695
      End
      Begin VB.TextBox Text9 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2250
         TabIndex        =   13
         Top             =   1800
         Width           =   4695
      End
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2250
         TabIndex        =   12
         Top             =   1440
         Width           =   4695
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   34
         TabStop         =   0   'False
         Text            =   "Télefono:"
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox Text15 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   33
         TabStop         =   0   'False
         Text            =   "Doc.:"
         Top             =   360
         Width           =   555
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   32
         TabStop         =   0   'False
         Text            =   "Celular:"
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox Text16 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   31
         TabStop         =   0   'False
         Text            =   "Contacto:"
         Top             =   1800
         Width           =   1815
      End
      Begin VB.TextBox Text17 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   30
         TabStop         =   0   'False
         Text            =   "E-Mail:"
         Top             =   1440
         Width           =   1815
      End
      Begin VB.TextBox Text18 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2250
         TabIndex        =   14
         Top             =   2160
         Width           =   4695
      End
      Begin VB.TextBox Text16 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   29
         TabStop         =   0   'False
         Text            =   "Celular de Contacto:"
         Top             =   2160
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6600
      Left            =   120
      TabIndex        =   17
      Top             =   0
      Width           =   7155
      Begin VB.CommandButton Command3 
         Appearance      =   0  'Flat
         Caption         =   "Guardar "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   3720
         TabIndex        =   16
         Top             =   5880
         Width           =   3210
      End
      Begin VB.CommandButton Command2 
         Appearance      =   0  'Flat
         Caption         =   "Nuevo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   1800
         TabIndex        =   36
         Top             =   5880
         Width           =   1050
      End
      Begin VB.CommandButton Command4 
         Appearance      =   0  'Flat
         Caption         =   "Baja"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   240
         TabIndex        =   35
         Top             =   5880
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.TextBox TEXT13 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2730
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Top             =   2880
         Width           =   6765
      End
      Begin VB.TextBox TXTDESCRIPROOVINCIA 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   1800
         Width           =   3550
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
         Height          =   315
         Left            =   3070
         TabIndex        =   5
         Top             =   1800
         Width           =   255
      End
      Begin VB.TextBox TXTCODPROVINCIA 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2220
         TabIndex        =   4
         Top             =   1800
         Width           =   795
      End
      Begin VB.TextBox Text15 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   315
         Locked          =   -1  'True
         TabIndex        =   26
         TabStop         =   0   'False
         Text            =   "Provinicia"
         Top             =   1800
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4080
         TabIndex        =   0
         Top             =   240
         Width           =   2595
      End
      Begin VB.CommandButton Command1 
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
         Height          =   315
         Left            =   6705
         TabIndex        =   25
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox Text17 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   6
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         Text            =   "Legajo:"
         Top             =   240
         Width           =   1365
      End
      Begin VB.TextBox Text15 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   6
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Text            =   "Sexo:"
         Top             =   2160
         Width           =   795
      End
      Begin VB.TextBox Text15 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   315
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Text            =   "Fecha de Nacimiento:"
         Top             =   2160
         Width           =   1815
      End
      Begin VB.TextBox Text17 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   315
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Text            =   "Nombres y Apellido:"
         Top             =   690
         Width           =   1815
      End
      Begin VB.TextBox Text16 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   315
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Text            =   "Domicilio:"
         Top             =   1050
         Width           =   1815
      End
      Begin VB.TextBox Text15 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   315
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Text            =   "Localidad:"
         Top             =   1410
         Width           =   1815
      End
      Begin VB.TextBox Text14 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2220
         TabIndex        =   1
         Top             =   690
         Width           =   4695
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2220
         TabIndex        =   2
         Top             =   1050
         Width           =   4695
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2220
         TabIndex        =   3
         Top             =   1410
         Width           =   4695
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   2
         Left            =   5025
         Sorted          =   -1  'True
         TabIndex        =   7
         Text            =   "Combo1"
         Top             =   2160
         Width           =   1920
      End
      Begin ACTIVESKINLibCtl.SkinLabel SKALTA 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FRMPROFESO1.frx":030A
         TabIndex        =   23
         Top             =   300
         Visible         =   0   'False
         Width           =   2475
      End
      Begin MSComCtl2.DTPicker DTP1 
         Height          =   330
         Left            =   2220
         TabIndex        =   6
         Top             =   2160
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   128778241
         CurrentDate     =   41123
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
         Height          =   375
         Left            =   240
         OleObjectBlob   =   "FRMPROFESO1.frx":0388
         TabIndex        =   37
         Top             =   2520
         Width           =   1935
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   10815
      OleObjectBlob   =   "FRMPROFESO1.frx":0400
      Top             =   105
   End
   Begin VB.Menu ABM 
      Caption         =   "ABM"
      Begin VB.Menu MNUALTA 
         Caption         =   "ALTA"
         Shortcut        =   ^A
      End
      Begin VB.Menu MNUBAJA 
         Caption         =   "BAJA"
         Shortcut        =   ^B
      End
      Begin VB.Menu MNUREST 
         Caption         =   "RESTAURAR BAJA"
         Shortcut        =   ^R
      End
   End
End
Attribute VB_Name = "FRMPROFESO1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub CARGA_DATOS_PROFESORES()

   LEGAJO& = XVALO(Text1)
   If LEGAJO& <= 0 Then Exit Sub
   CONDI$ = "NUME_CLI = " & LEGAJO&
   If CantRegistros&("PROFESORES", "NUME_CLI = " & LEGAJO&, "NOMESS") < 1 Then Exit Sub
   '
   SQLX$ = "SELECT * FROM PROFESORES WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE " + CONDI$
   Set rs = FLEXCONN.Execute(SQLX$)
   With rs
        Text14 = SAFETEXT$(!RASO_CLI)
        Text2 = SAFETEXT$(!DOMI_CLI)
        Text3 = SAFETEXT$(!LOCA_CLI)
'        Combo1(0).Text = SAFETEXT$(!PROV_CLI) VER COMO ASIGNARLE LOS DATOS A LOS COMBOS
        Combo1(1).Text = SAFETEXT$(!TIPODOCU_ALUM)
        Combo1(2).Text = SAFETEXT$(!SEXO_CLI)
        Text12.Text = SAFETEXT$(!DOCU_CLI)
        Text11.Text = SAFETEXT$(!Tele_Cli)
        Text10.Text = SAFETEXT$(!CELU_CLI)
        Text8.Text = SAFETEXT$(!MAIL_CLI)
        Text9.Text = SAFETEXT$(!CTAC_CLI)
        Text18.Text = SAFETEXT$(!CELCONTAC_CLI)
        TEXT13 = SAFETEXT$(!nota_Cli)
        Me.TXTRUTAFIRMA = SAFETEXT$(!IMAG_CLI)
        Me.TXTCODPROVINCIA = SAFETEXT$(!PROV_CLI)
        Me.DTP1.Value = CVDAT(CVINT(!FECHNAC))
        
      
    End With
    rs.Close
    Set rs = Nothing


End Sub




Private Sub Combo2_Change()
   TXCOMBO2 = Combo2.Text
End Sub

Private Sub Combo2_Click()
   TXCOMBO2 = Combo2.Text
End Sub

Private Sub Combo3_Change()
   TXCOMBO3 = Combo3.Text
End Sub

Private Sub Combo3_Click()
   TXCOMBO3 = Combo3.Text
End Sub

Private Sub Combo5_Change()
  TXCOMBO5 = Combo5.Text
End Sub

Private Sub Combo5_Click()
  TXCOMBO5 = Combo5.Text
End Sub


Private Sub Combo6_Change()
   TXCOMBO6 = Combo6.Text
End Sub

Private Sub Combo6_Click()
   TXCOMBO6 = Combo6.Text
End Sub

Private Sub Combo7_Change()
   TXCOMBO7 = Combo7.Text
End Sub

Private Sub Combo7_Click()
   TXCOMBO7 = Combo7.Text
End Sub

Private Sub Combo8_Change()
   TXCOMBO8 = Combo8.Text
End Sub

Private Sub Combo8_Click()
   TXCOMBO8 = Combo8.Text
End Sub

Sub creatabla_Profesores()
    Call ABRECONEXION
    Call CREACAMPO("", "PROFESORES", "TIPODOCU_ALUM", 10, 24)
    Call CREACAMPO("", "PROFESORES", "DOCU_CLI", 10, 24)
    Call CREACAMPO("", "PROFESORES", "CELCONTAC_CLI", 10, 48)
    Call CREACAMPO("", "PROFESORES", "ALTA_CLI", 8, 0)
    Call CREACAMPO("", "PROFESORES", "SEXO_CLI", 10, 18)
    Call CREACAMPO("", "PROFESORES", "Nume_Cli", 4, 0, 2)
    Call CREACAMPO("", "PROFESORES", "Raso_Cli", 10, 64, 1)
    Call CREACAMPO("", "PROFESORES", "Stat_Cli", 3, 0, 1)
    Call CREACAMPO("", "PROFESORES", "Domi_Cli", 10, 64)
    Call CREACAMPO("", "PROFESORES", "Cpos_Cli", 10, 16)
    Call CREACAMPO("", "PROFESORES", "Loca_Cli", 10, 48)
    Call CREACAMPO("", "PROFESORES", "Pais_Cli", 10, 48)
    Call CREACAMPO("", "PROFESORES", "Prov_Cli", 10, 2)
    Call CREACAMPO("", "PROFESORES", "Tele_Cli", 10, 32)
    Call CREACAMPO("", "PROFESORES", "Faxi_Cli", 10, 32)
    Call CREACAMPO("", "PROFESORES", "Mail_Cli", 10, 64)
    Call CREACAMPO("", "PROFESORES", "PagWeb", 10, 128)
    Call CREACAMPO("", "PROFESORES", "Piva_Cli", 3, 0)
    Call CREACAMPO("", "PROFESORES", "Cuit_Cli", 10, 24, 1)
    Call CREACAMPO("", "PROFESORES", "Nibr_Cli", 10, 24)
    Call CREACAMPO("", "PROFESORES", "Ctac_Cli", 10, 48)
    Call CREACAMPO("", "PROFESORES", "Nota_Cli", 12, 0)
    Call CREACAMPO("", "PROFESORES", "Obser_Entre", 12, 0)
    Call CREACAMPO("", "PROFESORES", "Recar_Embal", 6, 0)
    Call CREACAMPO("", "PROFESORES", "CELU_CLI", 10, 128)
    Call CREACAMPO("", "PROFESORES", "IMAG_CLI", 10, 512)
    Call CREACAMPO("", "PROFESORES", "FECHNAC", 8, 0)



End Sub


Private Sub Command1_Click()
    Command1.Enabled = False
    '
    If SKALTA.Visible = True Then
      TX$ = "PARA SELECCIONAR PROFESOR EXISTENTE" + vbCrLf
      TX$ = TX$ + "DEBE VOLVER A MODO EDICIÓN (Presionando el Botón 'Edición')"
      MsgBox TX$
      GoTo 99
    End If
    '
    Call ABRECONEXION
    '
    If PROFESORSEL > 0 Then
      Text1 = XVALO(RESP_ZELE_VECT(1))
    End If
    '
99  Command1.Enabled = True
    '
    
End Sub

Private Sub Command2_Click()
   Command2.Enabled = False
   'ALTAX
   If Command2.Caption = "Nuevo" Then
     Command2.Caption = "Edición"
     SKALTA.Visible = True
     NUMPRO& = 1 + MAXNUMCAMPO&("PROFESORES", "NUME_CLI")
     Text1 = NUMPRO&
   Else
     Command2.Caption = "Nuevo"
     SKALTA.Visible = False
     Text1 = ""
   End If

99 Command2.Enabled = True

End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    'VALIDAR SI HAY UN SOLO GRADO O AÑO O SALA DE
    'VALIDAR LA CUOTA.
    Call ABRECONEXION
    JJ& = 0
    '
    'TOMO EL NUMERO DEL LABEL
    LEGAJO& = XVALO(Text1)
    '
    'VALIDACION QUE SE HAYA SELECCIONADO UN CODIGO
    If LEGAJO& < 1 Then
       MsgBox "LEGAJO NO VALIDO"
       GoTo 99
    End If
    '
    If TRIM(Text14) = "" Then
       MsgBox "Debe Ingresar Nombre"
       GoTo 99
    End If
    '
    If TRIM(Me.Combo1(2)) = "" Then
       MsgBox "Falta el Sexo del PROFESOR"
       GoTo 99
    Else
      If Combo1(2).Text <> "Femenino" And Combo1(2) <> "Masculino" Then
         MsgBox "Sexo Seleccionado No Válido"
         GoTo 99
      End If
    End If
    '
    If TRIM(Me.Combo1(1)) = "" Then
       MsgBox "Tipo de Documento No Válido"
       GoTo 99
    Else
      If TRIM$(Combo1(1).Text) <> "DNI" And TRIM$(Combo1(1).Text) <> "CÉDULA" And TRIM$(Combo1(1)) <> "Pasaporte" And TRIM$(Combo1(1)) <> "Otro" Then
         MsgBox "Tipo de Documento Seleccionado No Válido"
         GoTo 99
      End If
    End If
    
    If TRIM$(TXTDESCRIPROOVINCIA) = "" Then
         MsgBox "Ingreso de Provincia Seleccionada No Válido"
         GoTo 99
    End If
    '
    'VALIDAR SI EL DOCUMENTO ES EXISTENTE CON MISMO NUMERO EN OTRO PROFESOR
    If TRIM$(Text12) <> "" Then
        CONDI$ = "DOCU_CLI =  '" & TRIM$(Text12) & "' AND TIPODOCU_ALUM = '" & UCase$(TRIM$(Combo1(1).Text)) & "' AND NUME_CLI <> " & XVALO(Text1)
        VALIDA_LEGAJO& = XVALO(CantRegistros("PROFESORES", CONDI$, "NOMESS"))
        If VALIDA_LEGAJO& > 0 Then
           NOMYAPE$ = TRIM$(VALOADMIN("PROFESORES", "NOMBRE_APELLIDO", CONDI$, "NOMESS"))
           MsgBox "DOCUMENTO YA EXISTENTE " + vbCrLf + " PROFESOR: " & NOMAPE$ & " LEGAJO: " & VALIDA_LEGAJO& & vbCrLf + "FAVOR DE CAMBIAR Y VOLVER A GRABAR"
           GoTo 99
        End If
        If Len(TRIM$(Text12)) < 7 Then
           OPX% = COMALTER%("El Número Ingresado Tiene Menos de 7 Dígitos\\Cancelar\Grabar Igual")
           If OPX% < 2 Then GoTo 99
        End If

    Else
         MsgBox "Falta Nro. de Documento"
         GoTo 99
    End If
    '
    If TRIM$(Text8) <> "" Then
       If InStr(Text8, "@") < 1 Then
          Call COMUNI("Atención !!! Mail Inválido")
          GoTo 99
       End If
    End If

    FF% = HOY(HO$)
    FECHA1 = FETEXCOR1$(FF%)
    '
    'ALTAX
    If SKALTA.Visible = True Then
      LEGAJO& = 1 + MAXNUMCAMPO&("PROFESORES", "NUME_CLI")
    End If
    '
    GRABO% = 0
    '
    'GRABACION , SI NO EXISTE LO GENERA SI NO LO EDITA
    SQLX$ = "SELECT * FROM PROFESORES"
    SQLX$ = SQLX$ + " WHERE NUME_CLI = " & LEGAJO&
    '
    On Error GoTo ERROR_GUARDAR
    Set AUX = New ADODB.Recordset
    AUX.Open (SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    FLEXCONN.BeginTrans
    With AUX
        If (.BOF And .EOF) Then
          .AddNew
          !NUME_CLI = LEGAJO&
          !ALTA_CLI = FECHA1
        Else
          !NUME_CLI = LEGAJO&
        End If
        
        !RASO_CLI = Left$(Text14, 64)
        !DOMI_CLI = Left$(Text2, 64)
        !LOCA_CLI = Left$(Text3, 64)
        '
        !TIPODOCU_ALUM = Left$(Combo1(1).Text, 16)
        !DOCU_CLI = Left$(Text12.Text, 32) 'TEXTO
        !Tele_Cli = Left$(Text11.Text, 64) 'TEXTO
        !CELU_CLI = Left$(Text10.Text, 64)
        !MAIL_CLI = Left$(Text8.Text, 64) 'TEXTO
        '
        !CTAC_CLI = Left$(Text9.Text, 64) 'TEX
        !CELCONTAC_CLI = Left$(Text18.Text, 64) 'TX
        !PROV_CLI = Me.TXTCODPROVINCIA
        !SEXO_CLI = Left$(Combo1(2).Text, 18)
        !IMAG_CLI = Left$(Me.TXTRUTAFIRMA.Text, 512)
        !FECHNAC = Me.DTP1.Value

        '
        !STAT_CLI = 0
        '
        !nota_Cli = TEXT13
      .Update
      '
    End With
    
ERROR_GUARDAR:
    If Err <> 0 Then
      FLEXCONN.RollbackTrans
      MsgBox "ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION" & vbCrLf & Err.Number & " " & Err.Description
      GoTo 99
    Else
      FLEXCONN.CommitTrans
      AUX.Close
      Set AUX = Nothing
      GRABO% = 1
      On Error GoTo 0
    End If
    '
    If SKALTA.Visible = True Then
       TXT$ = "Alta de Profesor"
       Call ANOTAX_PROFESORES(LEGAJO&, TXT$)
    End If

    'ALTAX
    SKALTA.Visible = False
    'REFRESCO PARA EL CASO QUE SE DIO ALTA Y EL NUMERO QUE GRABO ES <> AL QUE QUEDO EN LABEL
    Text1 = LEGAJO&
    Command2.Caption = "Nuevo"
    
    MsgBox "HECHO"
99  Command3.Enabled = True

End Sub

Private Sub Command4_Click()
    Command4.Enabled = False

    Call ABRECONEXION
    JJ& = 0
    '
    'TOMO EL NUMERO DEL LABEL
    LEGAJO& = XVALO(Text1)
    '
    'VALIDACION QUE SE HAYA SELECCIONADO UN CODIGO
    If LEGAJO& < 1 Then
       MsgBox "LEGAJO No Valido"
       GoTo 99
    End If
    '
    CONDI$ = "NUME_CLI = " & LEGAJO&
    '
    'VALIDO SI EXISTE
    CNRG& = CantRegistros&("PROFESORES", CONDI$)
    If CNRG& < 1 Then
      MsgBox "Legajo No Valido"
      GoTo 99
    End If
    '
    TTXX$ = "Realmente desea dar de Baja el Profesor Activo"
    If MsgBox(TTXX$, vbYesNo, "Baja de Profesor") <> vbYes Then GoTo 99
    'GRABACION , SI NO EXISTE LO GENERA SI NO LO EDITA
    SQLX$ = "SELECT * FROM PROFESORES "
    SQLX$ = SQLX$ + " WHERE NUME_CLI = " & LEGAJO&
    '
    Dim RECTMP As ADODB.Recordset
    Set RECTMP = New ADODB.Recordset
    RECTMP.Open SQLX$, FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText

    With RECTMP
        !STAT_CLI = -1
        .Update
    End With
    '
    RECTMP.Close
    Set RECTMP = Nothing
    '
    TXT$ = "Baja de Profesor"
    Call ANOTAX_PROFESORES(LEGAJO&, TXT$)
    '
    Call LIMPIAR
    Text1 = ""
    
99  Command4.Enabled = True

End Sub

Private Sub Command5_Click()
   Command5.Enabled = False
    Call ABRECONEXION
    '
    If PROVINSEL% > 0 Then
      Me.TXTCODPROVINCIA = RESP_ZELE_VECT(2)
    End If
   
   
99 Command5.Enabled = True

End Sub

Private Sub Command9_Click()
   RUTAX$ = SELECCION_Archivo$("*", CMD1)
   If RUTAX$ <> "" Then
        On Error Resume Next
        Image1.Picture = LoadPicture(RUTAX$)
        If Err > 0 Then MsgBox Err.Description
   Else
      GoTo 99
   End If
   TXTRUTAFIRMA = RUTAX$
99 Screen.MousePointer = 1



End Sub

Private Sub Form_Load()
    ''
    Call APLICAR_SKIN1(Me, App.Path & "\SKINS\dogmax.skn")
   'CARGA COMBO DE TIPO DE DOCUMENTO
    Combo1(1).AddItem ""
    Combo1(1).AddItem "DNI"
    Combo1(1).AddItem "CÉDULA"
    Combo1(1).AddItem "Pasaporte"
    Combo1(1).AddItem "Otro"
    Combo1(1).Text = "DNI"
    Combo1(2).Text = "Masculino"
    Combo1(2).AddItem "Femenino"
    Combo1(2).AddItem "Masculino"
    Call creatabla_Profesores
    DTP1.Value = Date
    
End Sub


Sub LIMPIAR()
   Text14 = ""
   Text2 = ""
   Text3 = ""
   Text8 = ""
   Text9 = ""
   Text10 = ""
   Text11 = ""
   Text12 = ""
   TEXT13 = ""
   Text18 = ""
   Text19 = ""
   Text20 = ""
   Text21 = ""
   Text22 = ""
   Text23 = ""
   Text24 = ""
   Text29 = ""
   Text28 = ""
   Text26 = ""
   Text29 = ""
   Text27 = ""
   Text38 = ""
   Text31 = ""
   Text32 = ""
   Text33 = ""
   Text25 = ""
   Me.TXTRUTAFIRMA = ""
   TXTCODPROVINCIA = ""
   Me.TXTDESCRIPROOVINCIA = ""

End Sub

Private Sub MNUALTA_Click()
   Call Command2_Click
End Sub

Private Sub MNUBAJA_Click()
   Call Command4_Click
End Sub

Private Sub MNUREST_Click()
    If PROFESORSEL(1) > 0 Then
      LEGAJO& = XVALO(RESP_ZELE_VECT(1))
    Else
      GoTo 99
    End If
    '
    
    SQLX$ = "SELECT * FROM PROFESORES "
    SQLX$ = SQLX$ + " WHERE NUME_CLI = " & LEGAJO&
    '
    Dim RECTMP As ADODB.Recordset
    Set RECTMP = New ADODB.Recordset
    RECTMP.Open SQLX$, FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    With RECTMP
        !STAT_CLI = 0
        NOMBREYAPE$ = SAFETEXT$(!RASO_CLI)
        .Update
    End With
    '
    RECTMP.Close
    Set RECTMP = Nothing
    '
    MsgBox "PROFESOR: " & NOMBREYAPE$ & " RESTAURADO"
    TXT$ = "Restauración de Profesor"
    Call ANOTAX_ALUMNO(LEGAJO&, TXT$)
    '
    Call LIMPIAR
    Text1 = SAFETEXT$(LEGAJO&)

99
End Sub

Private Sub Text1_Change()
     Call LIMPIAR
     Call CARGA_DATOS_PROFESORES
End Sub
Private Sub Text1_DblClick()
   Call Command1_Click
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
         SendKeys "{TAB}"
   End If
End Sub

Private Sub Text11_Change()
   Text11 = UCase(Text11)
   Text11.SelStart = Len(Text11)

End Sub

Private Sub Text12_Change()
   Text12 = UCase(Text12)
   Text12.SelStart = Len(Text12)

End Sub

Private Sub Text14_Change()
  Text14 = UCase(Text14)
  Text14.SelStart = Len(Text14)

End Sub

Private Sub TEXT14_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
         SendKeys "{TAB}"
   End If

End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
         SendKeys "{TAB}"
   End If

End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
         SendKeys "{TAB}"
   End If

End Sub

Private Sub Text12_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
         SendKeys "{TAB}"
   End If

End Sub

Private Sub Text17_DblClick(Index As Integer)
    If Index = 6 Then Command1_Click
    
End Sub


Private Sub Text18_Change()

   Text18 = UCase(Text18)
   Text18.SelStart = Len(Text18)

End Sub

Private Sub Text19_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
         SendKeys "{TAB}"
   End If

End Sub

Private Sub Text18_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
         SendKeys "{TAB}"
   End If
End Sub

Private Sub Text2_Change()
   Text2 = UCase(Text2)
   Text2.SelStart = Len(Text2)

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
         SendKeys "{TAB}"
   End If

End Sub

Private Sub Text20_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
         SendKeys "{TAB}"
   End If

End Sub

Private Sub Text21_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
         SendKeys "{TAB}"
   End If

End Sub

Private Sub Text22_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
         SendKeys "{TAB}"
   End If

End Sub

Private Sub Text23_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
         SendKeys "{TAB}"
   End If

End Sub

Private Sub Text25_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
         SendKeys "{TAB}"
   End If

End Sub

Private Sub Text26_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
         SendKeys "{TAB}"
   End If

End Sub

Private Sub Text27_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
         SendKeys "{TAB}"
   End If

End Sub

Private Sub Text28_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
         SendKeys "{TAB}"
   End If

End Sub

Private Sub Text29_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
         SendKeys "{TAB}"
   End If

End Sub

Private Sub Text3_Change()
   Text3 = UCase(Text3)
   Text3.SelStart = Len(Text3)

End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
         SendKeys "{TAB}"
   End If

End Sub


Private Sub Text32_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
         SendKeys "{TAB}"
   End If

End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
         SendKeys "{TAB}"
   End If

End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
         SendKeys "{TAB}"
   End If

End Sub

Private Sub Text8_Change()
   Text8 = UCase(Text8)
   Text8.SelStart = Len(Text8)

End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
         SendKeys "{TAB}"
   End If

End Sub

Private Sub Text9_Change()
    Text9 = UCase(Text9)
    Text9.SelStart = Len(Text9)

End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
         SendKeys "{TAB}"
   End If

End Sub

'Private Sub TXCOMBO1_Change(Index As Integer)
'   TXCOMBO1(Index) = UCase(TXCOMBO1(Index))
'End Sub


Private Sub TXTCODPROVINCIA_Change()
     If TRIM$(Me.TXTCODPROVINCIA) <> "" Then
         Me.TXTDESCRIPROOVINCIA = VALOBADA("TAPROVINCIA", "DESCRI_PROV", "CODITEX= '" & Me.TXTCODPROVINCIA & "'", "NOMESS")
     Else
         Me.TXTDESCRIPROOVINCIA = ""
     End If
     TXTDESCRIPROOVINCIA = UCase(TXTDESCRIPROOVINCIA)
     TXTDESCRIPROOVINCIA.SelStart = Len(TXTDESCRIPROOVINCIA)

End Sub


Private Sub TXTRUTAFIRMA_Change()
   Image1.Picture = LoadPicture("")
   If EXISTE%(Me.TXTRUTAFIRMA) > 0 Then
     On Error Resume Next
     Image1.Picture = LoadPicture(TXTRUTAFIRMA)
     On Error GoTo 0
   Else
     
   End If


End Sub


