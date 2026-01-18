VERSION 5.00
Begin VB.Form FORDELIBRE 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Notas de Debito de Texto Libre"
   ClientHeight    =   5640
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8200
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   12000
      Begin VB.CommandButton Command21 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   1575
         TabIndex        =   39
         Top             =   5145
         Width           =   180
      End
      Begin VB.TextBox TEXT21 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   1155
         TabIndex        =   38
         Text            =   " "
         Top             =   5145
         Width           =   435
      End
      Begin VB.TextBox Eco21 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   1995
         TabIndex        =   37
         Text            =   " "
         Top             =   5145
         Width           =   3165
      End
      Begin VB.CommandButton Command21 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   7035
         TabIndex        =   36
         Top             =   5115
         Width           =   180
      End
      Begin VB.TextBox TEXT21 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   6615
         TabIndex        =   35
         Text            =   " "
         Top             =   5115
         Width           =   435
      End
      Begin VB.TextBox Eco21 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   7410
         TabIndex        =   34
         Text            =   " "
         Top             =   5115
         Width           =   3270
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H0080FF80&
         Height          =   3060
         Left            =   7410
         ScaleHeight     =   3000
         ScaleWidth      =   3210
         TabIndex        =   24
         Top             =   1890
         Width           =   3270
         Begin VB.Frame Frame2 
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   0  'None
            Caption         =   "Frame2"
            Height          =   2640
            Left            =   0
            TabIndex        =   25
            Top             =   350
            Width           =   3585
            Begin VB.TextBox Text12 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   105
               MaxLength       =   4
               TabIndex        =   43
               Top             =   735
               Width           =   540
            End
            Begin VB.TextBox Text10 
               BeginProperty Font 
                  Name            =   "Courier"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1470
               TabIndex        =   33
               Text            =   " "
               Top             =   210
               Width           =   1590
            End
            Begin VB.Label Label23 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Courier"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1470
               TabIndex        =   42
               Top             =   1260
               Width           =   1590
            End
            Begin VB.Line Line5 
               X1              =   0
               X2              =   3200
               Y1              =   2630
               Y2              =   2630
            End
            Begin VB.Line Line4 
               X1              =   3200
               X2              =   3200
               Y1              =   0
               Y2              =   2625
            End
            Begin VB.Line Line1 
               X1              =   0
               X2              =   3200
               Y1              =   0
               Y2              =   0
            End
            Begin VB.Label Label19 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Tot.Bruto:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   -840
               TabIndex        =   31
               Top             =   2205
               Width           =   2220
            End
            Begin VB.Label Label20 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Courier"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1470
               TabIndex        =   30
               Top             =   2160
               Width           =   1590
            End
            Begin VB.Label Label18 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Courier"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1470
               TabIndex        =   29
               Top             =   735
               Width           =   1590
            End
            Begin VB.Label Label17 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "% IVA:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   210
               TabIndex        =   28
               Top             =   840
               Width           =   1170
            End
            Begin VB.Label Label22 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "IVA-RNI:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   -840
               TabIndex        =   27
               Top             =   1365
               Width           =   2220
            End
            Begin VB.Label Label13 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Tot,Neto:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   -840
               TabIndex        =   26
               Top             =   315
               Width           =   2220
            End
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   $"FORDELIBRE.frx":0000
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   0
            TabIndex        =   32
            Top             =   60
            Width           =   10410
         End
         Begin VB.Line Line2 
            X1              =   3200
            X2              =   3200
            Y1              =   0
            Y2              =   2940
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFC0FF&
         Height          =   3060
         Left            =   105
         ScaleHeight     =   3000
         ScaleWidth      =   7095
         TabIndex        =   21
         Top             =   1890
         Width           =   7155
         Begin VB.TextBox Text8 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2640
            Left            =   -30
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   23
            Top             =   330
            Width           =   7140
         End
         Begin VB.Line Line6 
            X1              =   0
            X2              =   7140
            Y1              =   2970
            Y2              =   2970
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   $"FORDELIBRE.frx":008A
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   0
            TabIndex        =   22
            Top             =   60
            Width           =   10410
         End
      End
      Begin VB.CommandButton Command4 
         Caption         =   "C.Cte."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   10920
         Picture         =   "FORDELIBRE.frx":0114
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Resumen de Cuenta Corriente"
         Top             =   3570
         Width           =   855
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
         Height          =   300
         Left            =   5250
         TabIndex        =   19
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1155
         Width           =   2640
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   6405
         MaxLength       =   1
         TabIndex        =   18
         TabStop         =   0   'False
         Text            =   "X"
         ToolTipText     =   "Configuracion de Funcionamiento - Doble-Click Aquí"
         Top             =   315
         Width           =   500
      End
      Begin VB.CommandButton Command11 
         Caption         =   "New"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   10920
         Picture         =   "FORDELIBRE.frx":041E
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Apertura de Nueva Cuenta de Cliente"
         Top             =   1890
         Width           =   855
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Notas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   10920
         Picture         =   "FORDELIBRE.frx":0860
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Notas de Clientes"
         Top             =   2730
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cancel"
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
         Left            =   10920
         Picture         =   "FORDELIBRE.frx":0B6A
         TabIndex        =   15
         ToolTipText     =   "Abandona la Aplicación"
         Top             =   315
         Width           =   855
      End
      Begin VB.CommandButton BOTNEXT 
         Caption         =   "O.K."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Left            =   10920
         Picture         =   "FORDELIBRE.frx":0E74
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Aprueba - Registra - Imprime Nota de Débito"
         Top             =   4410
         Width           =   855
      End
      Begin VB.TextBox Text11 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   9030
         TabIndex        =   4
         Text            =   " "
         Top             =   1155
         Width           =   1650
      End
      Begin VB.TextBox Text9 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   9345
         TabIndex        =   2
         TabStop         =   0   'False
         Text            =   " "
         Top             =   315
         Width           =   1170
      End
      Begin VB.CommandButton Command9 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   10500
         TabIndex        =   3
         Top             =   315
         Width           =   180
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2310
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1155
         Width           =   2745
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
         Height          =   300
         Left            =   1470
         TabIndex        =   10
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1155
         Width           =   750
      End
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
         Height          =   300
         Left            =   1470
         TabIndex        =   9
         TabStop         =   0   'False
         Text            =   " "
         Top             =   735
         Width           =   3585
      End
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
         Height          =   300
         Left            =   1470
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   " "
         Top             =   315
         Width           =   3585
      End
      Begin VB.CommandButton Command1 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1050
         TabIndex        =   1
         Top             =   1155
         Width           =   180
      End
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
         Height          =   300
         Left            =   210
         TabIndex        =   0
         Text            =   " "
         Top             =   1155
         Width           =   855
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "C.Pago:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   105
         TabIndex        =   41
         Top             =   5145
         Width           =   960
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Vended:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   5565
         TabIndex        =   40
         Top             =   5115
         Width           =   960
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Numero:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   7560
         TabIndex        =   14
         Top             =   1155
         Width           =   1380
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8505
         TabIndex        =   13
         Top             =   315
         Width           =   750
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   315
         TabIndex        =   7
         Top             =   735
         Width           =   960
      End
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   1050
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Pos.IVA:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   960
   End
End
Attribute VB_Name = "FORDELIBRE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LIST2ANT$

Private Sub BOTNEXT_Click()
   '
   OKEY% = 0
   Call GRADEBILIB(OKEY%)
   If OKEY% > 0 Then
      Call PRIDEBILIB
      Call CLEARDEBILIB
      Text1.SetFocus
   End If
   '
End Sub

Private Sub Command1_Click()
    '
    Call SELECHO("DEUDOR1")
    NCXI% = Val(TRIM$(VALACT1$("DEUDOR1")))
    If NCXI% > 0 Then
        Text1.TEXT = TRIM$(Str$(NCXI%))
        NC% = NCXI%
        OKEY% = 0
        Call DEBILIB1(OKEY%)
        If OKEY% <> 1 Then
           Call CLEARDEBILIB
           Exit Sub
        End If
    End If
    '
End Sub

Private Sub Command10_Click()
    Call SELECHO("ECODEP")
    DPX1% = Val(VALACT1$("ECODEP"))
    If DPX1% > 0 Then
        DP1% = DPX1%
        Text10.TEXT = TRIM$(Str$(DP1%))
        Text7.TEXT = ECOARCH$("ECODEP", Chr$(DP1%))
    End If
End Sub

Private Sub Command11_Click()
    Command11.Enabled = False
    Call CIERRARCH("*.*")
    Call CONECRUN("AMACLI/NEW", "Alta de Nueva Cuenta de Cliente")
    Call FRESHECHO("DEUDOR1")
    Command11.Enabled = True
End Sub

Private Sub Command12_Click()
   Call MENSERR(24, "Opción Inserción no Disponible.\Debe Agregarse al Final de la Lista.")
End Sub

Private Sub Command2_Click()
    If Text8.TEXT <> "" Then
          Text8.TEXT = ""
          Text1.TEXT = ""
          Text2.TEXT = ""
          Text3.TEXT = ""
          Text4.TEXT = ""
          Text5.TEXT = ""
          HOII% = HOY(HOS$)
          Text9.TEXT = HOS$
          FEFAC% = HOII%
          '
          Text10.TEXT = ""
          Label18.Caption = ""
          Label20.Caption = ""
          Label23.Caption = ""
          '
          Text1.SetFocus
        Else
          OPFACTURA.Option1.Value = True
          '
          VACON% = (-1)
          Unload FORDELIDIP
          FORFACDIP.Show
          OPFACTURA.Show 1
    End If
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    NCLI$ = TRIM$(Text1.TEXT)
    Call CONECRUN("*NOTACLI/" + NCLI$, "Notas de Clientes")
    Command3.Enabled = True
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    RESCUDEU.Text1.TEXT = Text1.TEXT
    Hide
    PROLLAMA$ = "FORDELIDIP"
    RESCUDEU.Show
    Command4.Enabled = True
End Sub

Private Sub Command9_Click()
    Call SELECHO("SELFECHA")
    VDEVV$ = VALACT1$("SELFECHA")
    If VDEVV$ <> "" Then
        Text9.TEXT = VDEVV$
        FF% = FECHANUM(Text9.TEXT)
        FEFAC% = FF%
    End If
End Sub

Private Sub Form_Load()
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    FF% = HOY(HO$)
    Text9.TEXT = HO$
    FEFAC% = FF%
    '
    Refresh
    VACON% = 0
    '
End Sub

Private Sub Label2_DblClick()
    OPREMIFAC.Show 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Cancel = 1
   If VACON% = (-1) Then
      Cancel = 0
   End If
End Sub

Private Sub Text1_Change()
    Text8.TEXT = ""
    Text10.TEXT = ""
    Label18.Caption = ""
    Label20.Caption = ""
    Label23.Caption = ""
    BOTNEXT.Enabled = True
End Sub

Private Sub Text1_DBLCLICK()
    '
    Call SELECHO("DEUDOR1")
    NCXI% = Val(TRIM$(VALACT1$("DEUDOR1")))
    If NCXI% > 0 Then
        Text1.TEXT = TRIM$(Str$(NCXI%))
        NC% = NCXI%
        OKEY% = 0
        Call DEBILIB1(OKEY%)
        If OKEY% <> 1 Then
           Call CLEARDEBILIB
           Exit Sub
        End If
    End If
    '
End Sub

Private Sub Text1_GotFocus()
    Static CTXX%
    If CTXX% = 0 Then
        Refresh
        CTXX% = 1
    End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       OKEY% = 0
       Call DEBILIB1(OKEY%)
       If OKEY% <> 1 Then
           Call CLEARDEBILIB
           Exit Sub
       End If
    End If
    '
End Sub

Private Sub Text10_Change()
    '
   ALIVABA = Val(Text12.TEXT)
   TOPEDIDO# = Val(Text10.TEXT)
   If PIVAFAC% <> 5 Then FORDELIDIP.Label18.Caption = CSTRING$(MKD$(TOPEDIDO# * ALIVABA / 100), 3, 10, 2, 2)
   If RENOIN% = 1 Then
      FORDELIDIP.Label23.Caption = CSTRING$(MKD$(TOPEDIDO# * (ALIVABA / ALIVA) * ALIVARNI / 100), 3, 10, 2, 2)
   End If
   FORDELIDIP.Label20.Caption = CSTRING$(MKD$(TOPEDIDO# + Val(Label18.Caption) + Val(Label23.Caption)), 3, 10, 2, 2)
   '
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      TOPEDIDO# = Val(Text10.TEXT)
      FORDELIDIP.Text10.TEXT = CSTRING$(MKD$(TOPEDIDO#), 3, 10, 2, 2)
   End If
End Sub

Private Sub Text10_LostFocus()
    TOPEDIDO# = Val(Text10.TEXT)
    FORDELIDIP.Text10.TEXT = CSTRING$(MKD$(TOPEDIDO#), 3, 10, 2, 2)
End Sub

Private Sub Text11_CLICK()
   If MONUMERA$ = "" Then
      MONUMERA$ = "AUTO"
      If TRIM$(CONTROL$("", "NUMFACTU")) = "MANUAL" Then
         MONUMERA$ = "MANUAL"
      End If
   End If
   If MONUMERA$ <> "MANUAL" Then
      Text1.SetFocus
   End If
End Sub

Private Sub Text11_LostFocus()
   NROPRO& = 0
   If Left$(Text11.TEXT, 3) = "ND." Then NROPRO& = Val(Mid$(Text11.TEXT, 4))
   Call NUFACTURA(NROPRO&)
   Text11.TEXT = "ND." + TRIM$(Str$(NROPRO&))
End Sub

Private Sub Text12_Change()
   ALIVABA = Val(Text12.TEXT)
   TOPEDIDO# = Val(Text10.TEXT)
   If PIVAFAC% <> 5 Then FORDELIDIP.Label18.Caption = CSTRING$(MKD$(TOPEDIDO# * ALIVABA / 100), 3, 10, 2, 2)
   If RENOIN% = 1 Then
      FORDELIDIP.Label23.Caption = CSTRING$(MKD$(TOPEDIDO# * (ALIVABA / ALIVA) * ALIVARNI / 100), 3, 10, 2, 2)
   End If
   FORDELIDIP.Label20.Caption = CSTRING$(MKD$(TOPEDIDO# + Val(Label18.Caption) + Val(Label23.Caption)), 3, 10, 2, 2)
End Sub

Private Sub Text2_GotFocus()
    Text1.SetFocus
End Sub

Private Sub Text3_GotFocus()
    Text1.SetFocus
End Sub

Private Sub Text4_GotFocus()
    Text1.SetFocus
End Sub

Private Sub Text5_GotFocus()
    Text1.SetFocus
End Sub

Private Sub Text6_Click()
    Text1.SetFocus
End Sub

Private Sub Text6_DblClick()
    OPFACTUR.Show 1
End Sub

Private Sub Text7_GotFocus()
    Text1.SetFocus
End Sub

Private Sub Text9_DblClick()
    Call SELECHO("SELFECHA")
    VDEVV$ = VALACT1$("SELFECHA")
    If VDEVV$ <> "" Then
        Text9.TEXT = VDEVV$
        FEFAC% = FECHANUM(Text9.TEXT)
    End If
End Sub

Private Sub Text9_LostFocus()
    VDEVV$ = TRIM$(Text9.TEXT)
    If FECHANUM(VDEVV$) > 0 Then
          FEFAC% = FECHANUM(VDEVV$)
        Else
          Call MENSERR(24, "Fecha no Válida")
          FEX = HOY(HO$)
          FEFAC% = FEX
          Text9.TEXT = HO$
          Text9.SetFocus
    End If
End Sub

Private Sub TEXT21_Change(Index As Integer)
    '
    If Index = 1 Then
        Eco21(Index).TEXT = ECOARCH$("CONPAG", MKI$(Val(TEXT21(Index).TEXT)))
      ElseIf Index = 2 Then
        Eco21(Index).TEXT = ECOARCH$("VENDEDOR", Chr$(Val(TEXT21(Index).TEXT)))
    End If
    '
End Sub

Private Sub Command21_Click(Index As Integer)
    '
    If Index = 1 Then
        ARCHEC$ = "CONPAG": LOCA% = 2
      ElseIf Index = 2 Then
        ARCHEC$ = "VENDEDOR": LOCA% = 1
    End If
    ARGU$ = CBIN$(TRIM$(TEXT21(Index).TEXT), 1, LOCA%)
    Call SELECHO(ARCHEC$)
    TEXT21(Index).TEXT = VALACT1$(ARCHEC$)
    '
End Sub


