VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form SEGUIM_COLEGIO1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SEGUIMIENTO"
   ClientHeight    =   9195
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   10755
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
   ScaleHeight     =   9195
   ScaleWidth      =   10755
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Height          =   975
      Left            =   70
      TabIndex        =   20
      Top             =   0
      Width           =   4695
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
         Left            =   1680
         TabIndex        =   23
         Top             =   240
         Width           =   2595
      End
      Begin VB.CommandButton Command1 
         Caption         =   "."
         Height          =   315
         Left            =   4320
         TabIndex        =   22
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
         Left            =   50
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Text            =   "Legajo:"
         Top             =   240
         Width           =   1605
      End
   End
   Begin VB.TextBox TEXT13 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   5640
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   19
      Top             =   100
      Width           =   5050
   End
   Begin VB.Frame Frame3 
      Height          =   975
      Left            =   4800
      TabIndex        =   17
      Top             =   0
      Width           =   870
      Begin VB.CommandButton Command2 
         Height          =   795
         Left            =   40
         Picture         =   "SEGUIM_COLEGIO1.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Actualizar"
         Top             =   120
         Width           =   765
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Anotaciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   65
      TabIndex        =   6
      Top             =   2400
      Width           =   10620
      Begin VB.TextBox Text11 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3855
         Left            =   8040
         TabIndex        =   16
         Text            =   "Text10"
         Top             =   840
         Width           =   90
      End
      Begin VB.TextBox Text10 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3855
         Left            =   1080
         TabIndex        =   15
         Text            =   "Text10"
         Top             =   840
         Width           =   90
      End
      Begin VB.ListBox List1 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3660
         Left            =   0
         TabIndex        =   12
         Top             =   840
         Width           =   10575
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   10515
         TabIndex        =   8
         Top             =   315
         Width           =   10580
         Begin VB.Label Label2 
            BackColor       =   &H00808080&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Fecha     Detalle                                                   Usuario"
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   0
            TabIndex        =   9
            Top             =   0
            Width           =   10560
         End
         Begin VB.Line Line8 
            X1              =   960
            X2              =   1080
            Y1              =   0
            Y2              =   360
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Frame7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4275
         Left            =   10320
         TabIndex        =   7
         Top             =   630
         Width           =   15
      End
      Begin VB.Line Line1 
         X1              =   10580
         X2              =   10580
         Y1              =   320
         Y2              =   6225
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   6120
      OleObjectBlob   =   "SEGUIM_COLEGIO1.frx":0578
      Top             =   1080
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Novedades"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1950
      Left            =   65
      TabIndex        =   0
      Top             =   7080
      Width           =   10620
      Begin VB.TextBox Text9 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   8400
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   840
         Width           =   2000
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   50
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "O.K"
         Height          =   435
         Left            =   8400
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1320
         Width           =   750
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1175
         Left            =   1155
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   735
         Width           =   7140
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   425
         Left            =   0
         ScaleHeight     =   360
         ScaleWidth      =   10515
         TabIndex        =   2
         Top             =   315
         Width           =   10580
         Begin VB.Label Label4 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Fecha     Detalle                                                   Usuario"
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   0
            TabIndex        =   5
            Top             =   0
            Width           =   10560
         End
         Begin VB.Line Line2 
            X1              =   0
            X2              =   1155
            Y1              =   420
            Y2              =   420
         End
         Begin VB.Label Label5 
            Caption         =   "Label5"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   15
            Left            =   0
            TabIndex        =   3
            Top             =   420
            Width           =   1380
         End
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Registrar Anotación"
         Height          =   540
         Left            =   9345
         TabIndex        =   11
         Top             =   1365
         Width           =   960
      End
      Begin VB.Line Line3 
         X1              =   0
         X2              =   10665
         Y1              =   1900
         Y2              =   1900
      End
      Begin VB.Line Line7 
         X1              =   10580
         X2              =   10580
         Y1              =   315
         Y2              =   1900
      End
      Begin VB.Line Line6 
         X1              =   8280
         X2              =   10590
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line Line5 
         X1              =   8295
         X2              =   8295
         Y1              =   735
         Y2              =   1995
      End
      Begin VB.Label Label9 
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
         Height          =   690
         Left            =   0
         TabIndex        =   4
         Top             =   1200
         Width           =   1170
      End
   End
   Begin VB.Frame Frame5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1350
      Left            =   70
      TabIndex        =   24
      Top             =   960
      Width           =   10620
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
         Left            =   75
         Locked          =   -1  'True
         TabIndex        =   43
         TabStop         =   0   'False
         Text            =   "Nombres y Apellido:"
         Top             =   210
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
         Left            =   75
         Locked          =   -1  'True
         TabIndex        =   42
         TabStop         =   0   'False
         Text            =   "Domicilio:"
         Top             =   570
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
         Left            =   75
         Locked          =   -1  'True
         TabIndex        =   41
         TabStop         =   0   'False
         Text            =   "Localidad:"
         Top             =   930
         Width           =   1815
      End
      Begin VB.TextBox Text19 
         Appearance      =   0  'Flat
         BackColor       =   &H00F1F1F1&
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
         Left            =   1980
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   210
         Width           =   8400
      End
      Begin VB.TextBox Text14 
         Appearance      =   0  'Flat
         BackColor       =   &H00F1F1F1&
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
         Left            =   1980
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   570
         Width           =   8400
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         BackColor       =   &H00F1F1F1&
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
         Left            =   1980
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   930
         Width           =   8400
      End
      Begin VB.TextBox TXCOMBO1 
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
         Left            =   690
         Locked          =   -1  'True
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   1365
         Width           =   1230
      End
      Begin VB.TextBox Text12 
         Appearance      =   0  'Flat
         BackColor       =   &H00F1F1F1&
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
         Left            =   1995
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   1365
         Width           =   4950
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         BackColor       =   &H00F1F1F1&
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
         Left            =   1995
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   1725
         Width           =   4950
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         BackColor       =   &H00F1F1F1&
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
         Left            =   1995
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   2085
         Width           =   4950
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
         Left            =   105
         Locked          =   -1  'True
         TabIndex        =   33
         TabStop         =   0   'False
         Text            =   "Télefono:"
         Top             =   1725
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
         Left            =   105
         Locked          =   -1  'True
         TabIndex        =   32
         TabStop         =   0   'False
         Text            =   "Doc.:"
         Top             =   1365
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
         Left            =   105
         Locked          =   -1  'True
         TabIndex        =   31
         TabStop         =   0   'False
         Text            =   "Celular:"
         Top             =   2085
         Width           =   1815
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BackColor       =   &H00F1F1F1&
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
         Left            =   8715
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   1725
         Width           =   4950
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H00F1F1F1&
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
         Left            =   8715
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   1365
         Width           =   4950
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
         Left            =   6825
         Locked          =   -1  'True
         TabIndex        =   28
         TabStop         =   0   'False
         Text            =   "Contacto:"
         Top             =   1725
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
         Left            =   6825
         Locked          =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Text            =   "E-Mail:"
         Top             =   1365
         Width           =   1815
      End
      Begin VB.TextBox Text18 
         Appearance      =   0  'Flat
         BackColor       =   &H00F1F1F1&
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
         Left            =   8715
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   2085
         Width           =   4950
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
         Left            =   6825
         Locked          =   -1  'True
         TabIndex        =   25
         TabStop         =   0   'False
         Text            =   "Celular de Contacto:"
         Top             =   2085
         Width           =   1815
      End
   End
End
Attribute VB_Name = "SEGUIM_COLEGIO1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Private Sub Command1_Click()
    Command1.Enabled = False
    '
    '
    Call ABRECONEXION
    '
    If ALUMONOSEL > 0 Then
      Text2 = XVALO(RESP_ZELE_VECT(1))
    End If
    '
99  Command1.Enabled = True

End Sub

Private Sub Command2_Click()
   Call Text2_Change
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   LGJO& = XVALO(Text2)
   If LGJO& < 1 Then Exit Sub
   '
   Call ANOTAX_ALUMNO(LGJO&, Text1)
   '
   Text1 = ""
   '
   List1.Clear
   Text2 = ""
   Text2.Refresh
   Text2 = LGJO&
   Text2.Refresh
   '
99 Command3.Enabled = True
End Sub




Private Sub Command9_Click()
   Hide
End Sub

Private Sub Form_Load()
    HOO% = HOY(HOS$)
    ''
    Text8 = HOS$
    Text9 = NOMBREUSUARIO$
'    Call APLICAR_SKIN(Me, App.Path & "\SKINS\dogmax.skn")

End Sub

Private Sub Text2_Change()
     Call LIMPIAR
     LGJO& = XVALO(Text2)
     If LGJO& < 1 Then Exit Sub
     
     Call CARGA_DATOS_ALUMNMOS
     Call CARGA_LIST_HISTORIAL
End Sub
Sub CARGA_DATOS_ALUMNMOS()

   LEGAJO& = XVALO(Text2)
   If LEGAJO& <= 0 Then Exit Sub
   CONDI$ = "NUME_CLI = " & LEGAJO&
   
   If CantRegistros&("DEUDOR12", CONDI$, "NOMESS") < 1 Then Exit Sub

   SQLX$ = "SELECT * FROM DEUDOR12 WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE " + CONDI$
   Set ALUMTMP = FLEXCONN.Execute(SQLX$)
   With ALUMTMP
        Text19 = SAFETEXT$(!RASO_CLI)
        Text14 = SAFETEXT$(!DOMI_CLI)
        Text7 = SAFETEXT$(!LOCA_CLI)
        '
        TEXT13 = SAFETEXT$(!nota_Cli)
    End With
    ALUMTMP.Close
    Set ALUMTMP = Nothing

End Sub

Sub CARGA_LIST_HISTORIAL()
     List1.Clear
     LGJO& = XVALO(Text2)
     If LGJO& < 1 Then Exit Sub
     '
     CONDI$ = "LEGAJO = " & LGJO&
     XX$ = VALOADMIN("ANOTA_ALUMNO", "ANOTACION", CONDI$, "NOMESS")
      
     LONGI& = Len(XX$)
     INICIO& = 1
     Do While INICIO& < LONGI&
        FIN& = InStr(INICIO&, XX$, Chr$(10))
        If FIN& > 0 Then
          TTXX$ = Mid$(XX$, INICIO&, FIN& - INICIO&)
          List1.AddItem TTXX$
          
          INICIO& = FIN& + 1
        Else
          TTXX$ = Mid$(XX$, INICIO&)
          List1.AddItem TTXX$
          Exit Do
        End If
     Loop
     '
     ' ESTO ES PARA POSICIONARSE AL FINAL DE LA LISTA
     TPXI& = List1.ListCount - 15
     If TPXI& < 0 Then TPXI& = 0
     List1.TopIndex = TPXI&
     '
     List1.Refresh
     OPENFORMS = DoEvents

   Exit Sub
   

End Sub

Sub LIMPIAR()
List1.Clear
Dim CONTROLX As CONTROL
For Each CONTROLX In SEGUIM_COLEGIO1.Controls ' RECORRO LOS CONTROLES DEL FORM ACTIVO
        If TypeOf CONTROLX Is TextBox Then    ' SI ES CAJA DE TEXTO
           If UCase$(CONTROLX.Name) <> "TEXT2" And (CONTROLX.BackColor = &HF1F1F1 Or CONTROLX.BackColor = &H80000005) Then
              CONTROLX.Text = ""
           End If
        End If
Next
End Sub

