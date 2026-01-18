VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form MODPED19 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Modificación de Pedidos de Clientes - Formato Lineal"
   ClientHeight    =   8385
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   12150
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8385
   ScaleWidth      =   12150
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2730
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   115
      Top             =   630
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   3360
      Left            =   0
      TabIndex        =   26
      Top             =   5040
      Width           =   11250
      Begin VB.Frame Frame7 
         Caption         =   "Frame7"
         Height          =   1275
         Left            =   4080
         TabIndex        =   56
         Top             =   1800
         Visible         =   0   'False
         Width           =   1380
         Begin VB.TextBox TOTIVA 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   0
            Left            =   0
            TabIndex        =   59
            Text            =   " "
            Top             =   0
            Visible         =   0   'False
            Width           =   870
         End
         Begin VB.TextBox TOTIVA 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   1
            Left            =   0
            TabIndex        =   58
            Text            =   " "
            Top             =   315
            Visible         =   0   'False
            Width           =   870
         End
         Begin VB.TextBox TOTIVA 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   2
            Left            =   0
            TabIndex        =   57
            Text            =   " "
            Top             =   630
            Visible         =   0   'False
            Width           =   870
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00808080&
         Caption         =   "RESUMEN FINAL"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   3180
         Left            =   5670
         TabIndex        =   40
         Top             =   0
         Width           =   5550
         Begin VB.CommandButton Command29 
            Caption         =   "más datos"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   140
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox Text18 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   3240
            MaxLength       =   14
            TabIndex        =   121
            Top             =   1320
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox Text17 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Roboto Mono"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   4125
            TabIndex        =   120
            TabStop         =   0   'False
            Top             =   1320
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.CheckBox Check2 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            Caption         =   "Dt.Rt.Cte."
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
            Left            =   2040
            TabIndex        =   118
            Top             =   1080
            Visible         =   0   'False
            Width           =   1185
         End
         Begin VB.TextBox Text141 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Roboto Mono"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   4440
            TabIndex        =   117
            TabStop         =   0   'False
            Top             =   970
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.Frame Frame8 
            BackColor       =   &H00808080&
            Caption         =   "I.V.A."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   120
            TabIndex        =   84
            Top             =   2205
            Width           =   1485
            Begin VB.Label Label13 
               Alignment       =   1  'Right Justify
               BackColor       =   &H0000C000&
               BackStyle       =   0  'Transparent
               Caption         =   "IVA 0"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   420
               TabIndex        =   90
               Top             =   195
               Width           =   1005
            End
            Begin VB.Label Label14 
               Alignment       =   1  'Right Justify
               BackColor       =   &H0000C000&
               BackStyle       =   0  'Transparent
               Caption         =   "IVA 1"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   420
               TabIndex        =   89
               Top             =   405
               Width           =   1005
            End
            Begin VB.Label Label16 
               BackColor       =   &H0000C000&
               BackStyle       =   0  'Transparent
               Caption         =   "T-1"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   60
               TabIndex        =   88
               Top             =   405
               Width           =   400
            End
            Begin VB.Label Label15 
               BackColor       =   &H0000C000&
               BackStyle       =   0  'Transparent
               Caption         =   "T-0"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   60
               TabIndex        =   87
               Top             =   195
               Width           =   375
            End
            Begin VB.Label Label19 
               BackColor       =   &H0000C000&
               BackStyle       =   0  'Transparent
               Caption         =   "T-2"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   60
               TabIndex        =   86
               Top             =   620
               Width           =   405
            End
            Begin VB.Label Label20 
               Alignment       =   1  'Right Justify
               BackColor       =   &H0000C000&
               BackStyle       =   0  'Transparent
               Caption         =   "IVA 2"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   300
               TabIndex        =   85
               Top             =   615
               Width           =   1125
            End
         End
         Begin VB.TextBox TOTIVA 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Roboto Mono"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   3
            Left            =   3885
            TabIndex        =   65
            Text            =   " "
            Top             =   2040
            Width           =   1500
         End
         Begin VB.TextBox Total 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Roboto Mono"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   3
            Left            =   3885
            TabIndex        =   64
            Text            =   " "
            Top             =   2745
            Width           =   1500
         End
         Begin VB.TextBox Text15 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Roboto Mono"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3885
            TabIndex        =   63
            Top             =   2400
            Width           =   1500
         End
         Begin VB.TextBox TICAMBIS 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   300
            Left            =   840
            MaxLength       =   14
            TabIndex        =   60
            Top             =   1000
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.TextBox Text14 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Roboto Mono"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   2
            Left            =   3885
            Locked          =   -1  'True
            TabIndex        =   54
            TabStop         =   0   'False
            Text            =   "    *******"
            Top             =   970
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.TextBox Text14 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Roboto Mono"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   1
            Left            =   4440
            TabIndex        =   53
            TabStop         =   0   'False
            Text            =   "    *******"
            Top             =   640
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.TextBox Text14 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Roboto Mono"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   0
            Left            =   4440
            TabIndex        =   52
            TabStop         =   0   'False
            Text            =   "    *******"
            Top             =   360
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.TextBox Total 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Roboto Mono"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   0
            Left            =   3885
            Locked          =   -1  'True
            TabIndex        =   43
            Text            =   " "
            Top             =   310
            Width           =   1500
         End
         Begin VB.TextBox Total 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Roboto Mono"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   1
            Left            =   3885
            Locked          =   -1  'True
            TabIndex        =   42
            TabStop         =   0   'False
            Text            =   " "
            Top             =   640
            Width           =   1500
         End
         Begin VB.TextBox PORDESCU 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   2370
            MaxLength       =   14
            TabIndex        =   19
            Top             =   640
            Width           =   1065
         End
         Begin VB.TextBox Total 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   2
            Left            =   3885
            TabIndex        =   41
            TabStop         =   0   'False
            Text            =   " "
            Top             =   970
            Width           =   1500
         End
         Begin VB.Label Label10 
            Caption         =   "Label10"
            Height          =   135
            Left            =   1560
            TabIndex        =   144
            Top             =   240
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label Label100 
            BackStyle       =   0  'Transparent
            Caption         =   "Moneda"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   138
            Top             =   750
            Width           =   795
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo de cambio"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   120
            TabIndex        =   137
            Top             =   900
            Width           =   795
         End
         Begin VB.Label Labeldto 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Dto.Varios"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1875
            TabIndex        =   122
            Top             =   1440
            Visible         =   0   'False
            Width           =   1305
         End
         Begin VB.Label Label31 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   3240
            TabIndex        =   119
            Top             =   960
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Total Bruto Pedido"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   7
            Left            =   2085
            TabIndex        =   68
            Top             =   2850
            Width           =   1665
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Percep. IIBB"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   6
            Left            =   2250
            TabIndex        =   67
            Top             =   2520
            Width           =   1500
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Impuestos (IVA)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   2085
            TabIndex        =   66
            Top             =   2160
            Width           =   1665
         End
         Begin VB.Image Image1 
            BorderStyle     =   1  'Fixed Single
            Height          =   540
            Left            =   120
            Picture         =   "MODPED19.frx":0000
            Top             =   120
            Visible         =   0   'False
            Width           =   540
         End
         Begin VB.Label MONEMIS 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Pesos"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   810
            TabIndex        =   62
            ToolTipText     =   "Doble-Click para 'Dólares'"
            Top             =   645
            Width           =   1125
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Descuento"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   4
            Left            =   1560
            TabIndex        =   61
            Top             =   450
            Width           =   1140
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Total Neto"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   0
            Left            =   1845
            TabIndex        =   46
            Top             =   390
            Width           =   1905
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Sub-Total Neto"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   3
            Left            =   1890
            TabIndex        =   45
            Top             =   1100
            Width           =   1485
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3600
            TabIndex        =   44
            Top             =   735
            Width           =   225
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00808080&
         Caption         =   "CONDICIONES COMERCIALES"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   3180
         Left            =   120
         TabIndex        =   32
         Top             =   0
         Width           =   5350
         Begin VB.Frame Frame6 
            BackColor       =   &H00808080&
            Height          =   390
            Left            =   1160
            TabIndex        =   125
            Top             =   1800
            Width           =   615
            Begin VB.OptionButton Option2 
               BackColor       =   &H00B0D0E0&
               Caption         =   "Option2"
               Height          =   200
               Left            =   330
               TabIndex        =   127
               Top             =   160
               Width           =   255
            End
            Begin VB.OptionButton Option1 
               BackColor       =   &H00B0D0E0&
               Caption         =   "Option1"
               Height          =   200
               Left            =   60
               TabIndex        =   126
               Top             =   160
               Width           =   255
            End
         End
         Begin VB.TextBox TXTFLETE 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1890
            TabIndex        =   123
            Text            =   " "
            ToolTipText     =   "Doble Click Para Borrar"
            Top             =   1890
            Width           =   3270
         End
         Begin VB.CommandButton Command21 
            Caption         =   "."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   3
            Left            =   1575
            TabIndex        =   113
            Top             =   1500
            Width           =   180
         End
         Begin VB.TextBox TEXT21 
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
            Height          =   300
            Index           =   3
            Left            =   1155
            TabIndex        =   112
            Text            =   " "
            Top             =   1500
            Width           =   435
         End
         Begin VB.TextBox Eco21 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   3
            Left            =   1890
            TabIndex        =   111
            Text            =   " "
            Top             =   1500
            Width           =   3270
         End
         Begin VB.TextBox Text44 
            Alignment       =   2  'Center
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
            Height          =   300
            Left            =   4620
            MaxLength       =   4
            TabIndex        =   50
            Top             =   1110
            Width           =   540
         End
         Begin VB.TextBox Text13 
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
            Height          =   795
            Left            =   1155
            MaxLength       =   96
            MultiLine       =   -1  'True
            TabIndex        =   18
            Top             =   2280
            Width           =   4005
         End
         Begin VB.TextBox Eco21 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   2
            Left            =   1890
            TabIndex        =   35
            Text            =   " "
            Top             =   1110
            Width           =   1800
         End
         Begin VB.TextBox TEXT21 
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
            Height          =   300
            Index           =   2
            Left            =   1155
            TabIndex        =   16
            Text            =   " "
            Top             =   1110
            Width           =   435
         End
         Begin VB.CommandButton Command21 
            Caption         =   "."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   2
            Left            =   1575
            TabIndex        =   17
            Top             =   1110
            Width           =   180
         End
         Begin VB.TextBox Eco21 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   1
            Left            =   1890
            TabIndex        =   34
            Text            =   " "
            Top             =   705
            Width           =   1980
         End
         Begin VB.TextBox TEXT21 
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
            Height          =   300
            Index           =   1
            Left            =   1155
            TabIndex        =   14
            Text            =   " "
            Top             =   705
            Width           =   435
         End
         Begin VB.CommandButton Command21 
            Caption         =   "."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   1
            Left            =   1575
            TabIndex        =   15
            Top             =   705
            Width           =   180
         End
         Begin VB.TextBox Eco21 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   0
            Left            =   1890
            TabIndex        =   33
            Text            =   " "
            Top             =   300
            Width           =   3270
         End
         Begin VB.TextBox TEXT21 
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
            Height          =   300
            Index           =   0
            Left            =   1155
            TabIndex        =   12
            Text            =   " "
            Top             =   300
            Width           =   435
         End
         Begin VB.CommandButton Command21 
            Caption         =   "."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   0
            Left            =   1575
            TabIndex        =   13
            Top             =   300
            Width           =   180
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Dto. Rec"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   5
            Left            =   3960
            TabIndex        =   143
            Top             =   840
            Width           =   645
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Flete"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   7
            Left            =   40
            TabIndex        =   124
            Top             =   1920
            Width           =   1050
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo Ingreso"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   6
            Left            =   -3000
            TabIndex        =   114
            Top             =   1575
            Width           =   4110
         End
         Begin VB.Label Label25 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   4620
            TabIndex        =   105
            Top             =   720
            Width           =   510
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "Cant. Items"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   120
            TabIndex        =   70
            Top             =   2640
            Width           =   645
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   600
            TabIndex        =   69
            Top             =   2790
            Width           =   540
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "% Comisión"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   4
            Left            =   3700
            TabIndex        =   51
            Top             =   1185
            Width           =   855
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Observa- ciones"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Index           =   3
            Left            =   120
            TabIndex        =   39
            Top             =   2205
            Width           =   975
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Vendedor"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   2
            Left            =   -3000
            TabIndex        =   38
            Top             =   1185
            Width           =   4110
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cond. de Pago"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   1
            Left            =   -3000
            TabIndex        =   37
            Top             =   780
            Width           =   4110
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Lista de Precio"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Index           =   0
            Left            =   130
            TabIndex        =   36
            Top             =   200
            Width           =   975
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00B0D0E0&
         Caption         =   "ENTREGA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1600
         Left            =   5670
         TabIndex        =   31
         Top             =   1560
         Width           =   5550
         Begin VB.TextBox LENTREGA 
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
            Height          =   300
            Index           =   1
            Left            =   1150
            MaxLength       =   32
            TabIndex        =   102
            Top             =   555
            Width           =   4230
         End
         Begin VB.TextBox LENTREGA 
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
            Index           =   3
            Left            =   1150
            MaxLength       =   32
            TabIndex        =   101
            Top             =   555
            Visible         =   0   'False
            Width           =   3750
         End
         Begin VB.TextBox LENTREGA 
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
            Height          =   300
            Index           =   0
            Left            =   1150
            MaxLength       =   32
            TabIndex        =   96
            Top             =   240
            Width           =   4230
         End
         Begin VB.TextBox Text20 
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
            Height          =   300
            Left            =   1150
            TabIndex        =   95
            Top             =   870
            Width           =   645
         End
         Begin VB.CommandButton Command25 
            Caption         =   "."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1800
            TabIndex        =   94
            Top             =   870
            Width           =   195
         End
         Begin VB.TextBox Text19 
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
            Height          =   300
            Left            =   1150
            TabIndex        =   93
            Top             =   1200
            Width           =   645
         End
         Begin VB.TextBox LENTREGA 
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
            Height          =   300
            Index           =   2
            Left            =   2035
            MaxLength       =   48
            TabIndex        =   92
            Top             =   1200
            Width           =   3345
         End
         Begin VB.CommandButton Command24 
            Caption         =   "."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1800
            TabIndex        =   91
            Top             =   1200
            Width           =   195
         End
         Begin VB.Label Label24 
            Height          =   225
            Left            =   105
            TabIndex        =   104
            Top             =   780
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.Label Label23 
            Height          =   225
            Left            =   105
            TabIndex        =   103
            Top             =   525
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Domicilio"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   2
            Left            =   -800
            TabIndex        =   100
            Top             =   360
            Width           =   1905
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Transporte"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   5
            Left            =   -800
            TabIndex        =   99
            Top             =   1305
            Width           =   1905
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Sucursal"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   13
            Left            =   -800
            TabIndex        =   98
            Top             =   975
            Width           =   1905
         End
         Begin VB.Label Label27 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   2040
            TabIndex        =   97
            Top             =   870
            Width           =   3345
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   8400
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   12240
      Begin VB.PictureBox MARCONOTA 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   3200
         Left            =   120
         ScaleHeight     =   3165
         ScaleWidth      =   11070
         TabIndex        =   129
         Top             =   1755
         Visible         =   0   'False
         Width           =   11100
         Begin VB.CommandButton Command16 
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
            Height          =   570
            Left            =   10440
            Picture         =   "MODPED19.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   135
            ToolTipText     =   "Imprime Ficha de Anotaciones"
            Top             =   2470
            Width           =   570
         End
         Begin VB.CommandButton Command15 
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
            Height          =   570
            Left            =   10440
            Picture         =   "MODPED19.frx":0974
            Style           =   1  'Graphical
            TabIndex        =   134
            ToolTipText     =   "Cancela Pedido de Cliente"
            Top             =   1050
            Width           =   570
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
            Height          =   570
            Left            =   10440
            Picture         =   "MODPED19.frx":0C7E
            Style           =   1  'Graphical
            TabIndex        =   133
            ToolTipText     =   "Continúa Generación del Pedido de Cliente"
            Top             =   420
            Width           =   570
         End
         Begin VB.TextBox Text6 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2800
            Left            =   0
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   132
            Top             =   360
            Width           =   10365
         End
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   435
            Left            =   -30
            ScaleHeight     =   405
            ScaleWidth      =   10380
            TabIndex        =   130
            Top             =   -30
            Width           =   10405
            Begin VB.Label Label8 
               Alignment       =   2  'Center
               BackColor       =   &H00E0E0E0&
               Caption         =   "ANOTACIONES DE CLIENTE"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   390
               Left            =   0
               TabIndex        =   131
               Top             =   120
               Width           =   10545
            End
         End
      End
      Begin MSFlexGridLib.MSFlexGrid MSF 
         Height          =   3195
         Left            =   120
         TabIndex        =   128
         Top             =   1755
         Width           =   11100
         _ExtentX        =   19579
         _ExtentY        =   5636
         _Version        =   393216
         Cols            =   4
         BackColor       =   16777215
         BackColorFixed  =   16501405
         BackColorBkg    =   16777215
         FillStyle       =   1
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.PictureBox Picture6 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   7950
         ScaleHeight     =   330
         ScaleWidth      =   3315
         TabIndex        =   106
         Top             =   600
         Visible         =   0   'False
         Width           =   3315
         Begin VB.CommandButton Command27 
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
            Height          =   300
            Left            =   1200
            TabIndex        =   108
            Top             =   0
            Width           =   180
         End
         Begin VB.TextBox Text16 
            Alignment       =   1  'Right Justify
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
            Height          =   300
            Left            =   840
            MaxLength       =   2
            TabIndex        =   107
            Text            =   " "
            Top             =   0
            Width           =   350
         End
         Begin VB.Label Label28 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   1360
            TabIndex        =   110
            Top             =   0
            Width           =   1900
         End
         Begin VB.Label Label29 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Motivo"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   20
            TabIndex        =   109
            Top             =   90
            Width           =   750
         End
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         ForeColor       =   &H80000008&
         Height          =   8415
         Left            =   11280
         ScaleHeight     =   8385
         ScaleWidth      =   975
         TabIndex        =   71
         Top             =   0
         Width           =   1000
         Begin VB.PictureBox MARCOBARRA 
            Appearance      =   0  'Flat
            BackColor       =   &H002B9973&
            ForeColor       =   &H80000008&
            Height          =   200
            Left            =   105
            ScaleHeight     =   165
            ScaleWidth      =   615
            TabIndex        =   141
            Top             =   6120
            Width           =   650
            Begin VB.Frame BARRATRAZA 
               BackColor       =   &H0048CC9D&
               BorderStyle     =   0  'None
               Height          =   500
               Left            =   0
               TabIndex        =   142
               Top             =   0
               Width           =   12000
            End
         End
         Begin VB.CommandButton Command22 
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
            Height          =   580
            Left            =   105
            Picture         =   "MODPED19.frx":0F88
            Style           =   1  'Graphical
            TabIndex        =   81
            ToolTipText     =   "Anotador de Pedidos"
            Top             =   3380
            Visible         =   0   'False
            Width           =   650
         End
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
            Picture         =   "MODPED19.frx":1292
            Style           =   1  'Graphical
            TabIndex        =   82
            ToolTipText     =   "Consultas y Re-Impresión"
            Top             =   3380
            Width           =   650
         End
         Begin VB.CommandButton Command2 
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
            Height          =   580
            Left            =   105
            Picture         =   "MODPED19.frx":159C
            Style           =   1  'Graphical
            TabIndex        =   80
            ToolTipText     =   "Cierra y Abandona la Aplicación"
            Top             =   105
            Width           =   650
         End
         Begin VB.CommandButton Command4 
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
            Height          =   580
            Left            =   105
            Picture         =   "MODPED19.frx":16E6
            Style           =   1  'Graphical
            TabIndex        =   79
            ToolTipText     =   "Resumen de Cuenta Corriente"
            Top             =   2790
            Width           =   650
         End
         Begin VB.CommandButton Command5 
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
            Height          =   580
            Left            =   105
            Picture         =   "MODPED19.frx":19F0
            Style           =   1  'Graphical
            TabIndex        =   78
            ToolTipText     =   "Configuración de Funcionamiento de Pedidos"
            Top             =   4800
            Width           =   650
         End
         Begin VB.CommandButton Command7 
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
            Height          =   570
            Left            =   105
            Picture         =   "MODPED19.frx":1E32
            Style           =   1  'Graphical
            TabIndex        =   77
            ToolTipText     =   "Acceso a Condiciones Comerciales"
            Top             =   5380
            Width           =   650
         End
         Begin VB.CommandButton Command8 
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
            Height          =   590
            Left            =   105
            Picture         =   "MODPED19.frx":213C
            Style           =   1  'Graphical
            TabIndex        =   76
            ToolTipText     =   "Acceso a Maestro de Artículos"
            Top             =   2000
            Width           =   650
         End
         Begin VB.CommandButton Command11 
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
            Height          =   580
            Left            =   105
            Picture         =   "MODPED19.frx":2446
            Style           =   1  'Graphical
            TabIndex        =   75
            ToolTipText     =   "Acceso a Base de Datos de Clientes"
            Top             =   1420
            Width           =   650
         End
         Begin VB.CommandButton Command13 
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
            Height          =   580
            Left            =   105
            Picture         =   "MODPED19.frx":2888
            Style           =   1  'Graphical
            TabIndex        =   74
            ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
            Top             =   3960
            Width           =   650
         End
         Begin VB.CommandButton Command3 
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
            Height          =   590
            Left            =   105
            Picture         =   "MODPED19.frx":2B92
            Style           =   1  'Graphical
            TabIndex        =   73
            ToolTipText     =   "Ayuda FLEXOFT en Línea"
            Top             =   685
            Width           =   650
         End
         Begin VB.CommandButton Command14 
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
            Height          =   960
            Left            =   120
            Picture         =   "MODPED19.frx":2E9C
            Style           =   1  'Graphical
            TabIndex        =   72
            ToolTipText     =   "Aprueba, Graba e Imprime Pedido"
            Top             =   7200
            Width           =   650
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   960
         TabIndex        =   48
         Top             =   1380
         Width           =   180
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2760
         Left            =   135
         TabIndex        =   47
         Top             =   2205
         Visible         =   0   'False
         Width           =   11115
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   0
         Top             =   0
      End
      Begin VB.CommandButton Command10 
         Height          =   260
         Left            =   420
         TabIndex        =   30
         Top             =   6855
         Width           =   2825
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00E0E0E0&
         Height          =   435
         Left            =   10
         ScaleHeight     =   375
         ScaleWidth      =   10455
         TabIndex        =   28
         Top             =   1785
         Visible         =   0   'False
         Width           =   10515
         Begin VB.Label Label22 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  Cant Código          Descripción            Fecha   Lista  Desc  P.Unit     Importe"
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404000&
            Height          =   370
            Left            =   30
            TabIndex        =   29
            Top             =   0
            Width           =   10430
         End
      End
      Begin VB.TextBox Text12 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1470
         MaxLength       =   40
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1395
         Width           =   3690
      End
      Begin VB.TextBox Text11 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   10095
         TabIndex        =   10
         Text            =   " "
         Top             =   1005
         Width           =   1120
      End
      Begin VB.TextBox Text10 
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
         Height          =   300
         Left            =   9300
         MaxLength       =   16
         TabIndex        =   11
         Top             =   1395
         Width           =   1920
      End
      Begin VB.TextBox Text9 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   10095
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   " "
         Top             =   225
         Width           =   960
      End
      Begin VB.CommandButton Command9 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   11040
         TabIndex        =   9
         Top             =   225
         Width           =   180
      End
      Begin VB.TextBox Text7 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5565
         TabIndex        =   6
         TabStop         =   0   'False
         Text            =   " "
         Top             =   615
         Width           =   1695
      End
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5565
         TabIndex        =   7
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1395
         Width           =   1695
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2310
         TabIndex        =   4
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1005
         Width           =   2850
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1470
         TabIndex        =   3
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1005
         Width           =   750
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1470
         TabIndex        =   2
         TabStop         =   0   'False
         Text            =   " "
         Top             =   615
         Width           =   3690
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1470
         TabIndex        =   1
         TabStop         =   0   'False
         Text            =   " "
         Top             =   225
         Width           =   3690
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         MaxLength       =   6
         TabIndex        =   0
         Text            =   " "
         Top             =   1380
         Width           =   855
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Pedidos de Clientes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   585
         Left            =   0
         TabIndex        =   139
         Top             =   240
         Width           =   1500
      End
      Begin VB.Label TEPRUEBA 
         AutoSize        =   -1  'True
         Caption         =   "Label2"
         Height          =   195
         Left            =   0
         TabIndex        =   136
         Top             =   0
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label LTEXT1 
         Caption         =   "Label19"
         Height          =   330
         Left            =   525
         TabIndex        =   83
         Top             =   0
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Venta"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   330
         Left            =   9510
         TabIndex        =   55
         ToolTipText     =   "Doble-Click para 'Reparación'"
         Top             =   520
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label9 
         Height          =   225
         Left            =   7455
         TabIndex        =   49
         Top             =   0
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Disponible"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5565
         TabIndex        =   27
         Top             =   1150
         Width           =   1695
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Pedido Nro."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   8230
         TabIndex        =   25
         Top             =   1110
         Width           =   1800
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Orden de Compra"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   7680
         TabIndex        =   24
         Top             =   1500
         Width           =   1530
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   9270
         TabIndex        =   23
         Top             =   310
         Width           =   750
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Límite Crédito"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   5565
         TabIndex        =   22
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   120
         TabIndex        =   21
         Top             =   1150
         Width           =   960
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "MODPED19.frx":31A6
      Top             =   1800
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   300
      OleObjectBlob   =   "MODPED19.frx":33DA
      TabIndex        =   116
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   1050
   End
End
Attribute VB_Name = "MODPED19"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ARCHIPRE$

Sub ACTUALIZAOF(NUPE&)
    'SE REALIZO PARA ENVAPLAST PARA VALIDAR SI EL PEDIDO TIENE CANTIDAD ENTREGADA O LA OF VINCULADA TIENE INFORMES DE PROD INFORMADOS
    SQLX$ = "SELECT * FROM PEDDETA WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE NROPED = " & NUPE&
    Set RS = READSET(SQLX$)
    With RS
      Do While Not .EOF
        CI1% = XVALO(!CODIINT)
        CANTI = XVALO(!CanPed)

        CONDI$ = "NROPED = " & NUPE& & " AND COD_INTE = " & CI1%
        NORFA$ = VALOBADA("ORDEFABR", "IDSUBOR", CONDI$, "NOMESS")
        Call ACTUREGISTRO("ORDEFABR", "CANPROY", CONDI$, CANTI, REGAF&, "NOMESS")
        Call FORPED07.GENERES(CI1%, CANTI, NORFA$)
'        Call GENRESER(NORFA$) 'SE REEMPLAZO POR LA DE ARRIBA QUE ES LA QUE UTILIZA CUANDO SE APRUEBA EL PEDIDO Y GENERA LA OF
'        Call ACTUECO(CI1%, CANTI, NORFA$)
        .MoveNext
      Loop
    End With
    RS.Close
    Set RS = Nothing
End Sub

Function YAHAYENTREGAS%(PNroPed)
   'VALIDA SI EL PEDIDO PARAMETRO TIENE ALGUNA ENTREGA
   YAHAYENTREGAS% = 0
   CONDICION$ = " NroPed =  " & PNroPed
   CANTENTREGADA = XVALO(VALOBADA("PEDDETA", "Sum(CantEnt)", CONDICION$, "NOMESS"))
   If CANTENTREGADA > 0 Then YAHAYENTREGAS% = 1
End Function

Private Sub Check2_Click()

   VEND% = XVALO(TEXT21(2))
   VENDED$ = TRIM$(ECOARCH$("VENDEDOR", Chr$(VEND%)))
   If VEND% = 0 Or InStr(UCase$(VENDED$), "GABAL") Then
      If Check2.Value = 1 Then
       Label31 = XVALO(CONTROL("ADMIPED", "DESRTCTE"))
      Else
       GoTo 20
      End If
   Else
20     Label31 = ""
       Check2.Value = 0
   End If


End Sub

Private Sub Command1_Click()
   If List1.ListCount < 1 Then Exit Sub
   'Call SELECHO("DEUDOR1")
   'NCLIE = XVALO(TRIM$(VALACT1$("DEUDOR1")))
   NCLIE = ListaSeleccionClientes()
   If NCLIE > 0 Then
     If REGICLI&(NCLIE) > 0 Then
       Text1.TEXT = TRIM$(Str$(NCLIE))
       Call CARDACLI
     End If
   End If
End Sub

Private Sub Command11_Click()
   Command11.Enabled = False
   Call CIERRARCH("*.*")
   ACONEC$ = "ARCHIG07/AMACLI"
   Call CONECRUN(ACONEC$, "Padron Maestro de Clientes")
   Command11.Enabled = True
'
'   Command11.Enabled = False
'   Call CIERRARCH("*.*")
'   ACONEC$ = "AMACLI01"
'   If EXISTE%("AMACLI02.EXE") > 0 Then ACONEC$ = "AMACLI02"
'   Call CONECRUN(ACONEC$, "Padron Maestro de Clientes")
'   Command11.Enabled = True
End Sub

Private Sub Command12_Click()
       '
    Command12.Enabled = False
    Screen.MousePointer = 11
    '
'    Call CARGALISEL("INDIPEDI", "PEDENCA", "NROPED/F6.0 ; REFERENCIA/A60", "NROPED > 0 ORDER BY NROPED DESC")
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If InStr(EPAC$, "TANIT") < 1 Then
      Call CARGALISEL("INDIPEDI", "PEDENCA", "NROPED/F6.0 ; REFERENCIA/A60", "NROPED > 0 ORDER BY NROPED DESC")
    Else
      Call CARGALISEL("INDIPEDI", "PEDENCA", "NROPED/F6.0 ; REFERENCIA/A60", "NROPED > 199000 ORDER BY NROPED DESC")
    End If
    
'    SQLX$ = "SELECT * FROM PEDENCA ORDER BY NroPed ASC"
'    'Set PEDENCTEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
'    Dim PEDENCTEMP As ADODB.Recordset
'    Set PEDENCTEMP = New ADODB.Recordset
'25  On Error Resume Next
'    PEDENCTEMP.Open SQLX$, FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
'    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
'       On Error GoTo 0
'       Call CapturaErrorOpen
'       GoTo 25
'    End If
'    On Error GoTo 0
'    '
'    JJ& = 0
'    With PEDENCTEMP
'      On Error Resume Next
'      .MoveFirst
'      If Err < 1 Then
'        Do While Not .EOF
'          'ANTIGUO CODIGO ANTES DE OT-05-0682-WAR-18/10/05///
''          NPX& = !NroPed
''          REFEPED$ = !Referencia
''          NCLIE = !NroClie
''          FEX = CVINT(!FECHEMIS)
''          ANULX$ = "": If !Status = 9 Then ANULX$ = "(Anul) "
''          YY$ = MKS$(NPX&) + ANULX$ + AJUSTI$(REFEPED$, 60)
''          JJ& = JJ& + 1
''          Call GRAREG("INDIPEDI", YY$, JJ&)
'            '\\\OT-05-0682-WAR-18/10/05///
'             NRODES& = !NroPed
'             TTXX$ = MKS$(NRODES&)
'               NOCO$ = !NroOcom
'             If NOCO$ <> "" Then
'               TTXX$ = TTXX$ + "O.C." + NOCO$ + " - "
'             End If
'               FEX = CVINT(!FechEmis)
'             TTXX$ = TTXX$ + FECHATEX$(FEX) + " - "
'               NCLIE = !NroClie
'               ANULX$ = "": If !Status = 9 Then ANULX$ = "(Anul) "
'             TTXX$ = TTXX$ + TRIM$(RASOCLI$(NCLIE)) + ANULX$ 'Space$(52)
'             JJ& = JJ& + 1
'             Call GRAREG("INDIPEDI", TTXX$, JJ&)
'             '\\\OT-05-0682-WAR-FIN///
'        .MoveNext
'        Loop
'      End If
'    End With
'    PEDENCTEMP.Close
'    Set PEDENCTEMP = Nothing
'    '
'    Call SETULTREG("INDIPEDI", JJ&)
'    Call CIERRARCH("INDIPEDI")
'    Call FRESHECHO("INDIPEDI")
    Screen.MousePointer = 1
    '
3   Call SELECHO("INDIPEDI/Indice General de Pedidos (Pendientes-Cumplidos-Anulados)")
    NPX& = XVALO(VALACT1$("INDIPEDI"))
    If NPX& < 1 Then GoTo 99
    '
    TIDOCUM$ = "Pedido de Cliente"
    TIDOCUX$ = "Pedido ABIERTO"
    TIDOCUY$ = "Pedido Interno"
    NUDOCU$ = TRIM$(Str$(NPX&))
    While Len(NUDOCU$) < 6: NUDOCU$ = "0" + NUDOCU$: Wend
    NUDOCU$ = NUDOCU$ + " "
    Screen.MousePointer = 11
    '
4   For UI& = ULTREG&("PDOCLST") To 1 Step -1
      XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
      If InStr(XX$, TIDOCUM$) > 4 Or InStr(XX$, TIDOCUX$) > 4 Or InStr(XX$, TIDOCUY$) > 4 Then
        PPXX% = InStr(XX$, "Nro.")
        If PPXX% > 0 Then
          XX$ = Left$(XX$, PPXX% - 1) + " " + Mid$(XX$, PPXX% + 4)
        End If
        If InStr(XX$, NUDOCU$) > 4 Then
          DOCUNU& = CVS(Left$(XX$, 4))
          GoTo 5
        End If
      End If
    Next UI&
    '
    Screen.MousePointer = 1
    Call COMUNI("No se Encontró el Documento Impreso\Correspondiente a este Pedido.\  \Si se Trata de un Pedido Pendiente de Entrega\Puede Refrescarlo con la Opción\'Modificación de Pedido en Curso'.")
    On Error Resume Next
    OPPED07.Command46.SetFocus
    On Error GoTo 0
    On Error Resume Next
    OPPED07.Show 1
    On Error GoTo 0
    GoTo 99
    '
5   For KKI% = 0 To 5
      EDITFORM.Picture1(KKI%).Cls
      EDITFORM.Picture1(KKI%).DrawWidth = 3
      EDITFORM.Picture1(KKI%).DrawStyle = 0
      EDITFORM.Picture1(KKI%).FillStyle = 1
    Next KKI%
    '
    REDOCU$ = REGACT$("PDOCLST")
    URE& = ULTREG&("PDOCSPL")
    LI& = 0: LS& = URE&
6   E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 7
    L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 7
    DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
    If DCC& < DOCUNU& Then LI& = L&: GoTo 6
    If DCC& > DOCUNU& Then LS& = L&: GoTo 6
    GoTo 8
    '
7   Screen.MousePointer = 1
    Call MENSERR(24, "Documento no Registrado")
    GoTo 99
    '
8   TPSP$ = ""
    Screen.MousePointer = 11
    While L& > 1
      DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
      If DCC& < DOCUNU& Then GoTo 9
      L& = L& - 1
    Wend
    '
9   CAPAGINAS% = 0
    For KKL& = L& To URE&
      TBUF$ = REGLEIDO$("PDOCSPL", KKL&)
      If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 95
      If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
        TPSP$ = TPSP$ + Mid$(TBUF$, 5)
      End If
    Next KKL&
    '
95  SPOOLSTRING$ = TPSP$
    For KKI% = 0 To 5
      EDITFORM.Picture1(KKI%).Cls
      EDITFORM.Picture1(KKI%).Top = 0
      EDITFORM.Picture1(KKI%).Refresh
      EDITFORM.Picture1(KKI%).Height = 7170
      EDITFORM.Picture1(KKI%).Width = 11700
      EDITFORM.Picture1(KKI%).DrawWidth = 3
      EDITFORM.Picture1(KKI%).DrawStyle = 0
      EDITFORM.Picture1(KKI%).FillStyle = 1
    Next KKI%
    ALTUPAGINA = 7170
    TOPEPAGINA = 0
    PAGINACTIVA% = 0
    CAPAGINAS% = 0
    '
10  PPXX% = InStr(TPSP$, Chr$(255))
    If PPXX% > 0 Then
      COPRIAT$ = Left$(TPSP$, PPXX% - 1)
      TPSP$ = Mid$(TPSP$, PPXX% + 1)
      Call TRACE(20, FIN& - Len(TPSP$), FIN&)
      '
      If Len(COPRIAT$) > 0 Then
        COPRI% = Asc(Left$(COPRIAT$, 1))
        ATRIP$ = Mid$(COPRIAT$, 2)
        Call SHOWCOMMAND(COPRI%, ATRIP$)
      End If
      GoTo 10
      '
    End If
    '
    Screen.MousePointer = 1
    PAGINACTIVA% = 0
    For KKI% = 0 To 5
       EDITFORM.Picture1(KKI%).Visible = False
    Next KKI%
    EDITFORM.Picture1(PAGINACTIVA%).Visible = True
    EDITFORM.Picture1(PAGINACTIVA%).Refresh
    EDITFORM.VScroll1.Min = 0
    EDITFORM.VScroll1.Value = 0
    EDITFORM.VScroll1.Max = 1
    If ALTUPAGINA > EDITFORM.Frame1.Height Then
       EDITFORM.VScroll1.Max = (ALTUPAGINA - EDITFORM.Frame1.Height) / 100
    End If
    EDITFORM.Command3.Enabled = False
    ' EDITFORM.Command5.Enabled = False
    EDITFORM.Show 1
    '
    GoTo 3
    '
99  Command12.Enabled = True
    '
'
'    '
'    Command12.Enabled = False
'    'Screen.MousePointer = 11
'    '
'    'JJ& = 0
'    'For U& = 1 To ULTREG&("PEDENCA")
'2   '  XX$ = REGLEIDO$("PEDENCA", U&)
'    '  If CVS(Left$(XX$, 4)) > 0 Then
'    '    NPX& = CVS(Left$(XX$, 4))
'    '    NCLIE = CVI(Mid$(XX$, 55, 2))
'    '    FEX = CVI(Mid$(XX$, 53, 2))
'    '    ANULX$ = "": If Mid$(XX$, 267, 1) = Chr$(9) Then ANULX$ = "(Anul) "
'    '    'YY$ = MKS$(NPX&) + ANULX$ + TRIM$(RASOCLI$(NCLIE)) + " - O.C." + TRIM$(Mid$(XX$, 57, 16)) + " - " + FECHATEX$(FEX)
'    '    YY$ = MKS$(NPX&) + ANULX$ + TRIM$(Mid$(XX$, 5, 48))
'    '    JJ& = JJ& + 1
'    '    Call GRAREG("INDIPEDI", YY$, JJ&)
'    '  End If
'    'Next U&
'    'Call SETULTREG("INDIPEDI", JJ&)
'    'Screen.MousePointer = 1
'    '
''3   Call SELECHO("INDIPEDI/Indice General de Pedidos (Pendientes-Cumplidos-Anulados)")
''    NPX& = XVALO(VALACT1$("INDIPEDI"))
'3   Call SELECHO("PEDENCA/Indice General de Pedidos (Pendientes-Cumplidos-Anulados)")
'    NPX& = XVALO(VALACT1$("PEDENCA"))
'    If NPX& < 1 Then GoTo 99
'    '
'    TIDOCUM$ = "Pedido de Cliente"
'    NUDOCU$ = TRIM$(Str$(NPX&))
'    While Len(NUDOCU$) < 6: NUDOCU$ = "0" + NUDOCU$: Wend
'    NUDOCU$ = " " + NUDOCU$ + " "
'    Screen.MousePointer = 11
'    '
'    For UI& = ULTREG&("PDOCLST") To 1 Step -1
'      XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
'      If InStr(XX$, TIDOCUM$) > 4 Then
'        PPXX% = InStr(XX$, "Nro.")
'        If PPXX% > 0 Then
'          XX$ = Left$(XX$, PPXX% - 1) + " " + Mid$(XX$, PPXX% + 4)
'        End If
'        If InStr(XX$, NUDOCU$) > 4 Then
'          DOCUNU& = CVS(Left$(XX$, 4))
'          GoTo 5
'        End If
'      End If
'    Next UI&
'    '
'    Screen.MousePointer = 1
'    Call COMUNI("No se Encontró el Documento Impreso\Correspondiente a este Pedido.\  \Si se Trata de un Pedido Pendiente de Entrega\Puede Refrescarlo con la Opción\'Modificación de Pedido en Curso'.")
'    GoTo 99
'    '
'5   For KKI% = 0 To 5
'      EDITFORM.Picture1(KKI%).Cls
'      EDITFORM.Picture1(KKI%).DrawWidth = 3
'      EDITFORM.Picture1(KKI%).DrawStyle = 0
'      EDITFORM.Picture1(KKI%).FillStyle = 1
'    Next KKI%
'    '
'    REDOCU$ = REGACT$("PDOCLST")
'    URE& = ULTREG&("PDOCSPL")
'    LI& = 0: LS& = URE&
'6   E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 7
'    L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 7
'    DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
'    If DCC& < DOCUNU& Then LI& = L&: GoTo 6
'    If DCC& > DOCUNU& Then LS& = L&: GoTo 6
'    GoTo 8
'    '
'7   Screen.MousePointer = 1
'    Call MENSERR(24, "Documento no Registrado")
'    GoTo 99
'    '
'8   TPSP$ = ""
'    Screen.MousePointer = 11
'    While L& > 1
'      DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
'      If DCC& < DOCUNU& Then GoTo 9
'      L& = L& - 1
'    Wend
'    '
'9   CAPAGINAS% = 0
'    For KKL& = L& To URE&
'      TBUF$ = REGLEIDO$("PDOCSPL", KKL&)
'      If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 95
'      If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
'        TPSP$ = TPSP$ + Mid$(TBUF$, 5)
'      End If
'    Next KKL&
'    '
'95  SPOOLSTRING$ = TPSP$
'    For KKI% = 0 To 5
'      EDITFORM.Picture1(KKI%).Cls
'      EDITFORM.Picture1(KKI%).Top = 0
'      EDITFORM.Picture1(KKI%).Refresh
'      EDITFORM.Picture1(KKI%).Height = 7170
'      EDITFORM.Picture1(KKI%).Width = 11700
'      EDITFORM.Picture1(KKI%).DrawWidth = 3
'      EDITFORM.Picture1(KKI%).DrawStyle = 0
'      EDITFORM.Picture1(KKI%).FillStyle = 1
'    Next KKI%
'    ALTUPAGINA = 7170
'    TOPEPAGINA = 0
'    PAGINACTIVA% = 0
'    CAPAGINAS% = 0
'    '
'10  PPXX% = InStr(TPSP$, Chr$(255))
'    If PPXX% > 0 Then
'      COPRIAT$ = Left$(TPSP$, PPXX% - 1)
'      TPSP$ = Mid$(TPSP$, PPXX% + 1)
'      Call TRACE(20, FIN& - Len(TPSP$), FIN&)
'      '
'      If Len(COPRIAT$) > 0 Then
'        COPRI% = Asc(Left$(COPRIAT$, 1))
'        ATRIP$ = Mid$(COPRIAT$, 2)
'        Call SHOWCOMMAND(COPRI%, ATRIP$)
'      End If
'      GoTo 10
'      '
'    End If
'    '
'    UI& = UI& + 1
'    XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
'    If InStr(XX$, TIDOCUM$) > 4 Then
'      If InStr(XX$, NUDOCU$) > 4 Then
'        DOCUNU& = CVS(Left$(XX$, 4))
'        GoTo 7
'      End If
'    End If
'    '
'    Screen.MousePointer = 1
'    PAGINACTIVA% = 0
'    For KKI% = 0 To 5
'       EDITFORM.Picture1(KKI%).Visible = False
'    Next KKI%
'    EDITFORM.Picture1(PAGINACTIVA%).Visible = True
'    EDITFORM.Picture1(PAGINACTIVA%).Refresh
'    EDITFORM.VScroll1.Min = 0
'    EDITFORM.VScroll1.Value = 0
'    EDITFORM.VScroll1.Max = 1
'    If ALTUPAGINA > EDITFORM.Frame1.Height Then
'       EDITFORM.VScroll1.Max = (ALTUPAGINA - EDITFORM.Frame1.Height) / 100
'    End If
'    EDITFORM.Command3.Enabled = False
'    EDITFORM.Command5.Enabled = False
'    EDITFORM.Show 1
'    '
'    GoTo 3
'    '
'99  Command12.Enabled = True
'    '
End Sub

Private Sub Command13_Click()
   Command3.Enabled = False
   Call ACCDIR("PEDIDO02")
   Command3.Enabled = True
End Sub

 Sub Command14_Click()
   Command14.Enabled = False
   '
   ' VALIDACION SI SE INGRSO UN UN NUMERO DE VENDEDOR VALIDA QUE EL MISMO SEA VALIDO
   VENDED$ = TRIM$(TEXT21(2))
   If XVALO(VENDED$) > 0 Then
      ECOVEN$ = TRIM$(ECOARCH$("VENDEDOR", Chr$(XVALO(VENDED$))))
      If ECOVEN$ = "" Then
        Call MENSERR(24, "Vendedor Ingresado No Válido")
        On Error Resume Next
           TEXT21(2).SelStart = 0
           TEXT21(2).SelLength = Len(TEXT21(2).TEXT)
           TEXT21(2).SetFocus
        On Error GoTo 0
        GoTo 99
      End If
   End If

   EPAC$ = TRIM$(UCase$(EMPREAC$))
   If InStr(EPAC$, "DOSIVA") > 0 Then
     If TRIM$(TEXT21(2)) = "" Or TRIM$(Text44) = "" Then
       Call MENSERR(24, "Falta Vendedor / Comisión")
       GoTo 99
     End If
   End If
   '
   NUPED& = XVALO(Text11)

   'validacion si la orden de compra es existente
   If CONTROL$("ADMIPED", "NOREPOC") = "SI" And TRIM$(Text10) <> "" Then
      NCLIEN = XVALO(Text1)
      NORCOM$ = UCase$(TRIM$(Text10))
      NUPED& = XVALO(Text11)
      If OCOMEXISTE%(NCLIEN, NORCOM$, NUPED&) Then
         Call COMUNI("Orden de Compra Existente")
         Text10 = ""
         GoTo 99
      End If
   End If

   'VALIDACION DE PEDIDO CON MOTIVO
   If XVALO(Text16) > 0 Then
      TX$ = "La Grabación de Este Pedido\Hará que se Deba Volver a Autorizar"
      TX$ = TX$ + "\¿Está Seguro de Continuar?"
      OPX% = COMALTER%(TX$ + "\\No,Cancelar\Si, Confirmar la Modificación")
      If OPX% < 2 Or OPX% > 2 Then GoTo 99
      '
      CONDI$ = "NROPED = " & NUPED&
      Call ACTUREGISTRO("PEDENCA", "OV_AUTORIZA1", CONDI$, "", REGAF&, "NOMESS")
      Call ACTUREGISTRO("PEDENCA", "OV_AUTORIZA2", CONDI$, "", REGAF&, "NOMESS")
   End If
   '
   'VALIDA QUE SE HAYA ELEGID UN TIPO INGRESO SIEMPRE QUE SEA OBLIGATORIO
   If CONTROL$("ADMIPED", "VALTIPIN") = "SI" Then
        If XVALO(TEXT21(3)) < 1 Then
            Call COMUNI("Falta Elegir un Tipo de Ingreso")
            GoTo 99
        End If
   End If
   '
   If CONTROL$("ADMIPED", "OBLICOM") = "SI" Then
   '\\\OT-06-0512-RG-FIN///
    If TRIM$(Text10) = "" Then
       MERRO$ = "Falta Orden de Compra Cliente"
       GoTo 99
    End If
   End If

      MODOAPROB$ = "" ' VARIABLE PARA FORZAR LA APROBACION POSTERIOR EN CASO QUE NO PASE LAS VALIDACIONES
      
   'CONTROL DE VALIDACION DE CONDICIONES COMERCIALES
   'SI SE CAMBIO LA DEL MAESTRO PRESENTARIA MENSAJE
   'CON EXEPCION SI ESTA AUTORIZADA EN CASO QUE ESTE AUTORIZADA SOLO PRESENTA SI NO
   'COINCIDE CON LA DEL MAESTRO NI TAMPOCO CON LA APROBADA
   NC = XVALO(Text1)
   STATUS_PED% = XVALO(VALOBADA("PEDENCA", "STATUS", "NROPED= " & NUPED&, "NOMESS"))
   If TRIM$(CONTROL("ADMIPED", "VALCONPA")) = "SI" Then
       CONDIP% = CPAGCLI%(NC)
       CONDIPAGO_PED% = XVALO(VALOBADA("PEDENCA", "CONPAG", "NROPED= " & NUPED&, "NOMESS"))
       If XVALO(TEXT21(1)) <> CONDIP% Then
'         If STATUS_PED% >= 0 Then
            If XVALO(TEXT21(1)) = CONDIPAGO_PED% Then GoTo 50
'         End If
          OPX% = COMALTER%("Atención !! \Se Ha Cambiado La Condición de Pago Original " & CONDIP% & " Por la " & TEXT21(1) & "\De Aprobarse el Presente Pedido\El Mismo Quedará Pendiente de Aprobación \\Modificar \Continuar ")
          If OPX% = 1 Then
             GoTo 99
           Else
             MODOAPROB$ = "CAMBIO_PAGO"
          End If
       End If
50   End If
   '
   'CONFIGURACION PARA QUE SEA MANDATARIA EL DOMICILIO DE ENTREGA
   If CONTROL("ADMIPED", "DOMIMAND") = "SI" Then
     If TRIM$(LENTREGA(0).TEXT) = "" Then
        Call COMUNI("Falta Ingresar Domicilio de Entrega")
        GoTo 99
     End If
   End If
   
   'VALIDO SI ESTA CONFIGURADA EL CARGO DE FLETE Y SI ES ASI LO EXIGE
   If CONTROL("ADMIPED", "FLETMAND") = "SI" Then
        If TRIM$(Me.TXTFLETE) = "" Then
           Call COMUNI("Falta Ingresar a Quien Corresponde el Cargo del Flete")
           GoTo 99
        End If
   End If
   '
'   'CONTROL DE VALIDACION DE CONDICIONES COMERCIALES
'   If TRIM$(Control("ADMIPED", "VALLISPR")) = "SI" Then
'       LISTAPRECIO% = Str$(LIPRCLI%(NC))
'       LISTAPRECIO_PED% = XVALO(VALOBADA("PEDENCA", "ListPre", "NROPED= " & NUPED&, "NOMESS"))
'       If XVALO(TEXT21(0)) <> LISTAPRECIO% Then
''        If STATUS_PED% >= 0 Then
'          If XVALO(TEXT21(0)) = LISTAPRECIO_PED% Then GoTo 60
''        End If
'          OPX% = COMALTER%("Atención !!\ Se Ha Cambiado La Lista de Precios Original " & LISTAPRECIO% & " Por la " & TEXT21(0) & "\De Aprobarse el Presente Pedido \El Mismo Quedará Registrado Pendiente de Aprobación \\Modificar\Continuar ")
'          If OPX% = 1 Then
'             GoTo 99
'          Else
'             MODOAPROB$ = MODOAPROB$ + " CAMBIO_LISTA"
'          End If
'       End If
'60 End If
   '
''   'VALIDO LOTES EN CASO QUE VENGAN DE PRESUPUESTOS.
''   VENCARPE$ = TRIM$(Control$("", "VENCARPE"))
''   If UCase$(VENCARPE$) = "CARDETLO" Then
''    For U& = 1 To ULTREG&("!CARDETPE")
''      X$ = REGLEIDO$("!CARDETPE", U&)
''      CAN = CVS(Mid$(X$, 47, 4))
''      If CAN > 0.05 Then
''        CIXI% = CVI(Left$(X$, 2))
''        If CIXI% > 0 Then
''           LOTEX$ = TRIM$(Mid$(X$, 3, 16))
''           If LOTEX$ <> "" Then
''            SLP = SALDO_LOTE_PEDIDO(CIXI%, LOTEX$)
''            SLP = SaldoLote(CIXI%, LOTEX$) - SLP
''            If SLP < CAN Then
''              Call COMUNI("Lote: " & LOTEX$ & " Para Código: " & CODEXT$(CIXI%) & "\Cantidad Pedida: " & FORMATNUM$(CAN, "F9.1") & " Supera el Saldo (" & SLP & ")")
''              Call CARGALOTE_PED(NUPED&)
''              Call CARDETPEDI(1)
''              GoTo 99
''            End If
''           Else
''              Call COMUNI("Para Código: " & CODEXT$(CIXI%) & "\Cantidad Pedida: " & FORMATNUM$(CAN, "F9.1") & " No Posee Lote")
''              Call CARGALOTE_PED(NUPED&)
''              Call CARDETPEDI(1)
''              GoTo 99
''           End If
''        End If
''
''      End If
''    Next U&
''    OPX% = COMALTER%("Desea Modificar Lotes del Pedido Activo\\Continuar Sin Modificar\Modificar Lotes")
''    If OPX% = 2 Then Call CARGALOTE_PED(NUPED&): Call CARDETPEDI(1): GoTo 99
''   End If

   'NO CONSIDERO QUE DEBERIA ESTAR PERO CUALQUIER COSA DESCOMENTAR ESTAS VALIDACIONES DE FACTURAS VENCIDAS Y ULTIMA COMPRA.
'   'CONTROL DE VALIDACION DE FACTURAS VENCIDAS
'   If TRIM$(Control("ADMIPED", "FACVENCI")) = "SI" Then
'       FIMISQL$ = FETEXCOR1$(HOY(HO$))
'       CONDI$ = "Nuclien = " & NC & " AND FEVENCIM <= '" & FIMISQL$ & "' AND SaldDebe >= 0.005"
'       FACTUVENCIDAS& = CantRegistros&("SADEUPEN", CONDI$, "NOMESS")
'       If FACTUVENCIDAS& > 0 Then
'          OPX% = COMALTER%("Atención !! Se Han Detectado Facturas Vencidas Para este Cliente\Si se Graba así El Pedido\Quedará Registrado Como Pendiente de Aprobación \\Continuar\Modificar ")
'          If OPX% >= 2 Then
'             GoTo 99
'          Else
'             MODOAPROB$ = MODOAPROB$ + " FACTURA_VENCIDA"
'          End If
'       End If
'   End If
   '
'   DIAS_ULT_COMPRA& = XVALO(Control("ADMIPED", "QULTICOM"))
'   If DIAS_ULT_COMPRA& > 0 Then
'      CANTIFACTURAS& = TIENE_FACTURAS&(NC, DIAS_ULTIMA_COMPRA%, FEULT_FACT%)
'      If CANTIFACTURAS& < 1 Then
'         OPX% = COMALTER%("Atención !! Se Han Detectado que el Cliente Activo\Hace Más de: " & DIAS_ULT_COMPRA& & " Días\Que No Registra Facturas\Última Compra: " & FECHATEX$(FEULT_FACT%) & "\Si se Graba así El Pedido\Quedará Registrado Como Pendiente de Aprobación \\Continuar\Modificar ")
'          If OPX% >= 2 Then
'             GoTo 99
'          Else
'             MODOAPROB$ = MODOAPROB$ + " Fecha Úlitima Factura: " & FECHATEX$(FEULT_FACT%)
'          End If
'      End If
'
'   End If
    CTRLSINCARGO% = ControlItemsSinCargo% ' FUNCION
    '0= TODO BIEN
    '1= TIENE ITEMS SIN CARGO PERO ESTAN APROBADOS
    '-1 TIENE ITEMS SIN CARGO QUE ESTAN EN LA LISTA Y NECESITAN APROBACION
    '-2 TIENE ITEMS  QUE NO ESTAN EN LA LISTA Y LA DIFERENCIA ES MAYOR A LA TOLERADA CONTRA LA LISTA
    '3 TIENE ITEMS CON VALOR 0 QUE NO ESTAN EN LA LISTA (SE CONSIDERA POR SI NO HAY PORCENTAJE DE TOLERANCIA

    If CTRLSINCARGO% < 0 Then
        If CTRLSINCARGO% = (-1) Then
            OPX% = COMALTER%("Control de Importes.\ \Se Han Detectado\" & TEXTOMENSACTRLSCARGO$(CTRLSINCARGO%) & "\\Cancelar\Continuar y Grabar Pedido Pendiente de Aprobación")
            If OPX% < 2 Then
              GoTo 99
            End If
            MODOAPROB$ = "PENDIENTE DE APROBACION POR MATERIALES SIN CARGO"
            'EN GRAPEDIDO NO PRESENTA MENSAJE YA LO PRESENTO EN ESTE IF
        ElseIf CTRLSINCARGO% = (-3) Then
            Call COMUNI("Control de Importes.\ \Imposible Continuar con el Presente Pedidos,\Se Han Detectado\" & TEXTOMENSACTRLSCARGO$(CTRLSINCARGO%))
            GoTo 99
        End If
   End If

   Call GRAPEDIDO(OKX%, MODOAPROB$)
   '
   Call DATMOTIVO
   '
   If OKX% > 0 Then
     '
     If InStr(EPAC$, "ARCON") > 0 Then
       Call OCOMSUCU.GENERAR_ARCHIVO_GENERO_PEDIDO(NC, NUPED&, XVALO(Text10), "MODIPEDI")
     End If
     CONPRIPE.Check7.Value = 1
     If CONTROL$("PEDIDO", "PRINOVEN") = "NO" Then CONPRIPE.Check7.Value = 0
     CONPRIPE.Check8.Value = 0
     If CONTROL$("PEDIDO", "PRINODES") = "SI" Then CONPRIPE.Check8.Value = 1
     CONPRIPE.Check2.Value = 0
     If CONTROL$("PEDCLIEN", "SUPRINT") = "SI" Then CONPRIPE.Check2.Value = 1
40   CONPRIPE.Show 1
   
     If CONPRIPE.Check7.Value = 0 And CONPRIPE.Check1.Value = 1 Then
        Call COMUNI("Imposible Enviar Mail.\No Se Ha Elegido la opción Orden de Venta.")
        GoTo 40
     End If
     
     If CONPRIPE.Label1.Caption = "OK" Then
        '
        CANCELPRINT% = 0
        If CONPRIPE.Check2.Value = 1 Then CANCELPRINT% = 1  ' IMPRESION FANTASMA

        If CONPRIPE.Check7.Value = 1 Then
           Call PRIPEDIDO(1)     ' Imprime Orden de Venta
        End If
        'LA IMPRESION FANTASMA ES SOLO PARA LA ORDEN DE DESPACHO
        If CONPRIPE.Check8.Value = 1 Then
           Call PRIPEDIDO(2)     ' Imprime Orden de Despacho
        End If
        '
        ' VERIFICA SI TIENE QUE REPLICAR LA ORDEN DE VENTA
        If CONPRIPE.Check1.Value = 1 Then
           '
           NCX = XVALO(Text1)
           '
           'PERMITE ENVIAR A PARTIR DE LA LISTA DE MAILS DE CONTACTOS
           COCLI$ = VALOBADA("DEUDOR12", "CODI_CLI", "NUME_CLI = " & NCX, "NOMESS")
           Call CONTACTO_MAIL(COCLI$, NCX) 'AGREGA A SENDBYMAIL LOS MAIL DE LOS CONTACTOS SELECCIONADOS
           If SENDBYMAIL$ = "" Then
              SENDBYMAIL$ = TRIM$(EMAILCLI$(NCX))
           End If
           '
           If SENDBYMAIL$ = "" Or InStr(SENDBYMAIL$, "@") < 1 Then 'TIENE DIRECCION VALIDA
             Call MENSERR(24, "Imposible Enviar por Mail.\   \Falta Direccion de Email o la Existente es Inválida.")
             GoTo 30
           End If
           '
25         CANCELPRINT% = 0
           Call PRIPEDIDO(1)
        End If
        
30      CONPRIPE.Label1 = ""
        CANCELPRINT% = 0
     End If
     '
     Call GENERODES      ' genera orden de despacho
     Call BLANFORMU
     SENDBYMAIL$ = ""
     Call SETULTREG("!CARDETPE", 0)
     Call CIERRARCH("*.*")
     'Call CARGALISEL("INDIPEDI", "PEDENCA", "NROPED/F6.0 ; REFERENCIA/A60", "NROPED > 0 ORDER BY NROPED DESC")
     EPAC$ = TRIM$(UCase$(EMPREAC$))
     If InStr(EPAC$, "TANIT") < 1 Then
      Call CARGALISEL("INDIPEDI", "PEDENCA", "NROPED/F6.0 ; REFERENCIA/A60", "NROPED > 0 ORDER BY NROPED DESC")
     Else
      Call CARGALISEL("INDIPEDI", "PEDENCA", "NROPED/F6.0 ; REFERENCIA/A60", "NROPED > 199000 ORDER BY NROPED DESC")
     End If
     
     Hide
   End If
99 Command14.Enabled = True
End Sub
Sub DATMOTIVO()
'SI EL PEDIDO ES CON MOTIVO PRESENTA LA VENTANA DE CARGA DE DATOS
    NPED& = XVALO(Text11)
    NUMOT% = XVALO(Text16)
    If NUMOT% < 1 Then Exit Sub
    
    DESCMOTI$ = ECOARCH$("TAMOTPRE", Chr$(NUMOT%))
    '
10  FILTX$ = TRIM$(Str$(NPED&)) & ";" & TRIM$(Str$(NUMOT%)) & ";" & TRIM$(DESCMOTI$)
    Call STDFORM("DESMO1", FILTX$)
    If STDSCREEN.APRO <> "OK" Then
       If OPX% = 2 Then GoTo 15
       Exit Sub
    End If
    '

15  ' VER QUE SE VA A VALIDAR ACA
End Sub

Private Sub Command15_Click()
   MARCONOTA.Visible = False
   Text1.TEXT = ""
   Call BLANFORMU
   Call CIERRARCH("*.*")
   Hide
End Sub

Private Sub Command16_Click()
   Call MENSERR(24, "No Disponible por el Momento")
End Sub

Private Sub Command16_LostFocus()
   On Error Resume Next
   Command6.SetFocus
   On Error GoTo 0
End Sub

Private Sub Command2_Click()
    Call SETULTREG("!CARDETPE", 0)
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command22_Click()
    Command22.Enabled = False
    'Screen.MousePointer = 11
    '
    'JJ& = 0
    'For U& = 1 To ULTREG&("PEDENCA")
2   '  XX$ = REGLEIDO$("PEDENCA", U&)
    '  If CVS(Left$(XX$, 4)) > 0 Then
    '    NPX& = CVS(Left$(XX$, 4))
    '    NCLIE = CVI(Mid$(XX$, 55, 2))
    '    FEX = CVI(Mid$(XX$, 53, 2))
    '    ANULX$ = "": If Mid$(XX$, 267, 1) = Chr$(9) Then ANULX$ = "(Anul) "
    '    YY$ = MKS$(NPX&) + ANULX$ + TRIM$(Mid$(XX$, 5, 48))
    '    JJ& = JJ& + 1
    '    Call GRAREG("INDIPEDI", YY$, JJ&)
    '  End If
    'Next U&
    'Call SETULTREG("INDIPEDI", JJ&)
    'Screen.MousePointer = 1
    '
    CONPED07.Show 1
    '
    Command22.Enabled = True
End Sub

Private Sub Command24_Click()
  '\\\OT-6-422-FG-05/10/06
''   Call SELECHO("TRANSPOR")
''   NTRAN% = XVALO(VALACT1$("TRANSPOR"))
''   Text19 = TRIM$(FORMATNUM(NTRAN%, "I4"))
   Call COPYSTRU("TRANSPOR", "TRANSPO1")
   Call CARGALISEL("TRANSPO1", "TRANSPORTES", "CODI_TRA/F5;DENO_TRA/A32", "CODI_TRA > 0 ORDER BY DENO_TRA", "NOMESS")
   Call SELECHO("TRANSPO1")
   NTRAN% = XVALO(VALACT1$("TRANSPO1"))
   If NTRAN% > 0 Then
      Text19 = TRIM$(FORMATNUM(NTRAN%, "I4"))
   End If

  '\\\OT-6-422-FG-FIN
End Sub

Private Sub Command25_Click()
    '\\\OT-6-422-FG-05/10/06
    Call SELECHO("!SUCENTRE")
    VDEVU$ = VALACT1$("!SUCENTRE")
    Text20 = VDEVU$
    '\\\OT-6-422-FG-FIN
End Sub

Private Sub Command27_Click()
    'BOTON DE SELECCION DE MOTIVO
    Command27.Enabled = False
    '
    'CARGO TABLA DE MOTIVOS SI NO ESTABA CARGADA PREVIAMENTE
    If ULTREG&("TAMOTPRE") < 6 Then
        Call GRAREG("TAMOTPRE", Chr$(0) + "Ventas", 1)
        Call GRAREG("TAMOTPRE", Chr$(1) + "Giras/Congresos", 2)
        Call GRAREG("TAMOTPRE", Chr$(2) + "Demostracion", 3)
        Call GRAREG("TAMOTPRE", Chr$(3) + "Prestamo", 4)
        Call GRAREG("TAMOTPRE", Chr$(4) + "Cambio", 5)
        Call GRAREG("TAMOTPRE", Chr$(5) + "Donaciones", 6)
        Call CIERRARCH("TAMOTPRE")
    End If
    '
    'MUESTRO VENTANA DE SELECCION DE TABLA DE MOTIVOS
    Call SELECHO("TAMOTPRE")
    NUMOT% = XVALO(VALACT1$("TAMOTPRE"))
    DESCRIMOT$ = TRIM$(VALACT2$("TAMOTPRE"))
    'VALIDO LOS DATOS PARA SABER SI ELIGIO SALIR O ELIGIO 0 QUE ES VENTAS
    If NUMOT% <= 0 And DESCRIMOT$ = "" Then GoTo 99
    '
    Text16 = Str$(NUMOT%)
    '
99  Command27.Enabled = True

End Sub




Private Sub Command29_Click()
    
    If Command29.Caption = "mas datos" Then
        Command29.Caption = "menos datos"
    Else
        Command29.Caption = "mas datos"
    End If
        
    Call Image1_Click
    
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    Call HELPONLINE("PEDIDO01")
    Command3.Enabled = True
End Sub

Private Sub Command4_Click()
    Call MENSERR(24, "Imposible Acceder al Resumen de Cuenta Corriente\Desde esta Pantalla.")
    Exit Sub
    'Command4.Enabled = False
    'RESCUDEU.Text1.TEXT = Text1.TEXT
    'PROLLAMA$ = "MODPEDCLI"
    'Hide
    'RESCUDEU.Show
    'Command4.Enabled = True
End Sub

Private Sub Command5_Click()
    Command5.Enabled = False
    Call CIERRARCH("*.*")
    If (USNBR% Mod 100) <> 1 Then
      Call COMUNI("La Opción Elegida Requiere\Privilegios de Supervisor.")
      Exit Sub
    End If
    CFGPED07.Show 1
    Command5.Enabled = True
End Sub

Sub Command6_Click()
   MARCONOTA.Visible = False
   Call CARDETPEDI(1)
End Sub
'
Private Sub Command7_Click()
    Command7.Enabled = False
    Call CIERRARCH("*.*")
    MODX% = ALTERNA%("Listas de Precios\Condiciones de Pago")
    If MODX% = 1 Then
         Call CONECRUN("LPREAC07", "Listas de Precios de Venta")
         'Call CONECRUN("ADMIPRE", "Listas de Precios de Venta")
       ElseIf MODX% = 2 Then
         Call CARCONPAGOS
    End If
    Command7.Enabled = True
End Sub

Private Sub Command8_Click()
    Command8.Enabled = False
    Call CIERRARCH("*.*")
    ACONEC$ = "ARCHIG07/AMASTO"
    Call CONECRUN(ACONEC$, "Maestro de Artículos de Stock")
    Command8.Enabled = True
    '
'    Command8.Enabled = False
'    Call CIERRARCH("*.*")
'    ACONEC$ = "AMASTO01"
'    If EXISTE%("AMASTO02.EXE") > 0 Then ACONEC$ = "AMASTO02"
'    Call CONECRUN(ACONEC$, "Maestro de Artículos de Stock")
'    Command8.Enabled = True
End Sub

Private Sub Command9_Click()
   If VALICLIE% < 1 Then
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
      Exit Sub
   End If
End Sub

Private Sub Eco21_GotFocus(Index As Integer)
   On Error Resume Next
   TEXT21(Index).SetFocus
   On Error GoTo 0
End Sub

Private Sub Form_Load()
    '
    Call CENTRAFORM(Me)
'    EPAC$ = TRIM$(EMPREAC$)
'    If EPAC$ <> "" Then
'        Caption = Caption + "  -  " + EPAC$
'    End If
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
       EMPRE_SIN_USUARIO$ = REPLACAR$(EPAC$, USERNAME$, "")
       Caption = Caption + "  -  " + EMPRE_SIN_USUARIO$
    End If
    
    Frame4.Height = 1440
    If CONTROL$("ADMIPED", "MOCOMIS") = "LETRA" Then
       Label21(4) = "Letra:"
    End If
    '
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me)
    'DERECHO DE MODIFICACION DE DESCUENTO EN PANTALLA
    If DERACCE%("MODIDESC/NOMESS", "MODIFICAR DESCUENTO EN PANTALLA") < 2 Then
       PORDESCU.Enabled = False
    End If
    '
    If InStr(UCase$(EPAC$), "DOSIVA") > 0 Or InStr(EPAC$, "MEDITEA") > 0 Then
      Label11.Visible = True
      Command22.Visible = True
      Command12.Visible = False
    End If
    '
    If EXISTE%("REPARA02.EXE") > 0 Then
      Label11.Visible = True
    End If
    '
    CAMPOS$ = "Cant./F6;Código/A16;Descripción/A22;Fecha/D5;Lista/F10.2;Desc/F6.2;P.Unit/F10.2;Importe/F10.2;Lote/A0"
    Call CARGA_ENCABEZADO(Me.MSF, CAMPOS$, Me)
    Call InstalaFuenteAnchoFijo
    For i = 1 To 8
         MSF.COL = i: MSF.CellFontBold = False: MSF.CellFontSize = 9
    Next i
    MSF.Font.Size = 9
    
    'PARA MEDITEA MUESTRA BOTON DE SELECCION DE MOTIVO DEL PEDIDO
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If InStr(EPAC$, "MEDITEA") > 0 Then
       If EXISTE(LUDAT$ + "\TAMOTPRE.RFS") > 0 Then
         Picture6.Visible = True
         Text16 = "0"
       End If
    End If

    MONEMI% = 1  ' PESOS
    MONEMIS.Caption = DEMONECO$(MONEMI%)
    MONEMIS.Visible = True
    CONVERPRECIO% = 1
    TICAMBIS.Visible = True
    Label12.Visible = True
    'Image1.Visible = True
    TICAMBIS = FORMATNUM$(TICAMONE(MONEMI%), "F9.4")
    '
    CMO$ = TRIM$(CONTROL$("", "CODPECLI"))
    If CMO$ <> "" Then Label10.Caption = CMO$
    '
    ECOPRECIO% = 0
    If CONTROL$("ADMIPED", "MUEPREC") <> "SI" Then
      For KKI% = 0 To 2      ' TAPA IMPORTES TOTALES
        Text14(KKI%).Visible = True
      Next KKI%
      ECOPRECIO% = (-1)
    End If
    '
    If ECOPRECIO% < 0 Then
      Label22.Caption = "  Cant Código          Descripción                                           Entrega "
      Image1.Visible = False
    End If
    '
End Sub

 Sub MARCONOTA_Click()

End Sub



'
Sub MSF_DblClick()
   Call CARDETPEDI(2)
End Sub
Sub Image1_Click()
    If ECOPRECIO >= 0 Then
      EPAC$ = TRIM$(UCase$(EMPREAC$))
        If Frame4.Height = 3120 Then
          Frame4.Height = 1440
          If InStr(EPAC$, "GABAL") > 0 Then
            Check2.Visible = False
            Label31.Visible = False
            Labeldto.Visible = False
            Text17.Visible = False
            Text18.Visible = False
            Text141.Visible = False
            Ltota(3).Top = 1155
            Text14(2).Top = 1050
            Total(2).Top = 1050
          End If
        Else
          Frame4.Height = 3120
          If InStr(EPAC$, "GABAL") > 0 Then
            Check2.Visible = True
            Label31.Visible = True
            Labeldto.Visible = True
            Text17.Visible = True
            Text18.Visible = True
            Text141.Visible = True
            Ltota(3).Top = 1805
            Text14(2).Top = 1780
            Total(2).Top = 1700
          End If
        End If
    End If
End Sub

Private Sub Image1_DblClick()
   Static CTXX%
   '
   If ULTREG&("DAPIBCLI") < 1 Then
     Call MENSERR(24, "Imposible desde esta Aplicación")
     Exit Sub
   End If
   '
   NCLIE = XVALO(TRIM$(Text1.TEXT))
   If NCLIE > 0 Then
     If rasocli$(NCLIE) <> "" Then
       '
       If CTXX% < 1 Then
         Call GRAREG("TACATIBC", Chr$(0) + "Exento Percepcion IB", 1)
         Call GRAREG("TACATIBC", Chr$(1) + "Sujeto Percepcion IB", 2)
         Call SETULTREG("TACATIBC", 2)
         CTXX% = 1
       End If
       '
       RECORD$ = FILTERGETRECORD$("DAPIBCLI", 1, MKI$(NCLIE))
       Call REPLA(RECORD$, MKI$(NCLIE), 1, 2)
       CATIB1% = Asc(Mid$(RECORD$, 3, 1))
       If CATIB1% < 0 Or CATIB1% > 1 Then
         Call REPLA(RECORD$, Chr$(1), 3, 1)
         Call REPLA(RECORD$, Space$(32), 9, 32)
       End If
       '
10     OBX$ = OBJPLA$("CARDAPERIB", RECORD$)
       If OBX$ <> "" Then
         RECORD$ = OBX$
         CATIB1% = Asc(Mid$(RECORD$, 3, 1))
         If CATIB1% < 0 Or CATIB1% > 1 Then
           Call MENSERR(24, "Mal Categoría I.Brutos")
           GoTo 10
         End If
         If CATIB1% = 0 Then
           If TRIM$(Mid$(RECORD$, 9, 32)) = "" Then
             Call MENSERR(24, "Mandatorio Indicar Motivo de Exención")
             GoTo 10
           End If
         End If
         If CATIB1% = 1 Then
           Call REPLA(RECORD$, Space$(32), 9, 32)
         End If
         PORCONVE = CVS(Mid$(RECORD$, 5, 4))
         If PORCONVE < 0 Or PORCONVE > 100 Then
           Call MENSERR(24, "% Convenio Multilateral\Fuera de Rango")
           GoTo 10
         End If
         '
         Call REPLA(RECORD$, MKI$(NCLIE), 1, 2)
         Call REPLA(RECORD$, USERNAME$, 41, 20)
         Call REPLA(RECORD$, MKI$(HOY(HOS$)), 61, 2)
         Call REPLA(RECORD$, MKI$(Int(HORANUM(Time$))), 63, 2)
         Call FILTERPUTRECORD("DAPIBCLI", 1, MKI$(NCLIE), RECORD$)
         Call CIERRARCH("DAPIBCLI")
         '
         Call CALPERIB
         '
       End If
     End If
   End If
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   If ECOPRECIO >= 0 Then
'     Frame4.Height = 3120
'   End If
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Frame4.Height = 1440
End Sub

Private Sub Label11_DblClick()
   If TRIM$(UCase$(Label11)) = "VENTA" Then
       Label11.ForeColor = RGB(255, 0, 0)
       Label11 = "Reparación"
       Label11.ToolTipText = "Doble-Click para 'Venta'"
     Else
       Label11.ForeColor = RGB(0, 0, 255)
       Label11 = "Venta"
       Label11.ToolTipText = "Doble-Click para 'Reparación'"
   End If
End Sub

Private Sub Label25_Change()
    Call CALTOPED
End Sub

Private Sub Label31_Change()
    If ROUND#(Label31) < 100 Then
      Label31 = FORMATNUM(XVALO(Label31), "F5.2")
      Call CALTOPED
    Else
      Call COMUNI("Descuento No Válido !!! Verifique")
      Label31 = ""
    End If

End Sub

Private Sub Label8_Click()
   On Error Resume Next
   Text13.SetFocus
   On Error GoTo 0
End Sub

Private Sub LENTREGA_GotFocus(Index As Integer)
   '\\\OT-6-422-FG-05/10/06
   If Index = 2 Then
     If XVALO(Text19) > 0 Then
       Call MENSERR(24, "Campo no Editable")
       On Error Resume Next
       Text19.SetFocus
       On Error GoTo 0
       Exit Sub ' LE AGREGUE ESTA LINEA PARA QUE NO SE EJECUTE EL RESTO DEL CODIGO
     End If
   End If
   '\\\OT-6-422-FG-FIN
   '
   If VALICLIE% < 1 Then
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
      Exit Sub
   End If
   LENTREGA(Index).TEXT = TRIM$(LENTREGA(Index).TEXT)
   LENTREGA(Index).SelStart = 0
   LENTREGA(Index).SelLength = Len(LENTREGA(Index).TEXT)
End Sub

Private Sub List1_DblClick()
   Call CARDETPEDI(2)
End Sub

Private Sub MONEMIS_Change()
Exit Sub
   MONEMA% = MONEMI%
   If Left$(MONEMIS, 1) = "D" Then
        MONEMI% = 2
        MONEMIS.ForeColor = RGB(150, 0, 0)
        MONEMIS.ToolTipText = "Doble-Click para 'Pesos'"
        'Label12.Visible = True
        If TICAMBIS.Visible = True Then TICAMBIS.Enabled = True
        TICAMBIS = FORMATNUM$(TICAMONE(2), "F9.4")
      Else
        MONEMI% = 1
        MONEMIS.ForeColor = RGB(0, 100, 0)
        MONEMIS.ToolTipText = "Doble-Click para 'Dólares'"
        'Label12.Visible = False
        If TICAMBIS.Visible = True Then TICAMBIS.Enabled = False
   End If
   '
   COEFCON = TICAMONE(MONEMA%) / TICAMONE(MONEMI%)
   FIN& = ULTREG&("!CARDETPE")
   For U& = 1 To FIN&
     X$ = REGLEIDO$("!CARDETPE", U&)
     CAN = CVS(Mid$(X$, 47, 4))
     PORDESCUIN = CVS(Mid$(X$, 69, 4))
     PRECARGA = COEFCON * CVS(Mid$(X$, 39, 4))
     PREUNI = PRECARGA * (100 - PORDESCUIN) / 100
     TOT# = CDbl(CAN * PREUNI)
     '
     Call REPLA(X$, MKS$(PRECARGA), 39, 4)
     Call REPLA(X$, MKS$(PRECARGA), 51, 4)
     Call REPLA(X$, MKD$(TOT#), 55, 8)
     Call REPLA(X$, MKS$(PRECARGA), 65, 4)
     Call REPLA(X$, MKS$(PREUNI), 73, 4)
     Call REPLA(X$, MKD$(TOT#), 77, 8)
     Call GRAREG("!CARDETPE", X$, U&)
   Next U&
   '
   Call CARLISPE
   Call CIERRARCH("!CARDETPE")
   Call CALTOPED
   '
End Sub

Private Sub MONEMIS_DblClick()
'   If TRIM$(UCase$(MONEMIS)) = "PESOS" Then
'       MONEMIS = "Dólares"
'     Else
'       MONEMIS = "Pesos"
'   End If
   Static CAMON%, DMONE$(), TCMONE()
   If CAMON% < 1 Then
     CAMON% = ULTREG&("TAMONED")
     If CAMON% < 1 Then
         CAMON% = 1
         ReDim DMONE$(CAMON%), TCMONE(CAMON%)
         DMONE$(1) = "Pesos"
         TCMONE(1) = 1
       Else
         ReDim DMONE$(CAMON%), TCMONE(CAMON%)
         For U& = 1 To CAMON%
           XX$ = REGLEIDO$("TAMONED", U&)
           DMX$ = TRIM$(Left$(XX$, 8)) + " "
           PPXX% = InStr(DMX$, " ")
           DMX$ = Left$(DMX$, PPXX% - 1)
           '
           VHX = CVS(Mid$(XX$, 29, 4)): If VHX < 0.005 Then VHX = 1
           DMONE$(U&) = DMX$
           TCMONE(U) = VHX
         Next U&
     End If
   End If
   '
   MONEMIA% = MONEMI%
15 MONEMI% = MONEMI% + 1
   If MONEMI% > CAMON% Then MONEMI% = 1
   If DMONE$(MONEMI%) = "" Then GoTo 15
   '
   MONEMIS = DMONE$(MONEMI%)
   MONEMIS.ForeColor = RGB(0, 0, 255)
   If MONEMI% > 1 Then
      MONEMIS.ForeColor = RGB(255, 0, 0)
   End If
   MONEMIS.ToolTipText = "T.C.: " + TRIM$(FORMATNUM$(TICAMONE(MONEMI%), "F10.4")) + " - Doble-Click para Cambiar Moneda"
   '
   COEFCON = TICAMONE(MONEMIA%) / TICAMONE(MONEMI%)
   TICAMBIS = FORMATNUM$(TICAMONE(MONEMI%), "F9.4")
   '
   If MONEMI% > 1 Then
     TICAMBIS.Visible = True
     TICAMBIS.Enabled = False
   End If
   '
   FIN& = ULTREG&("!CARDETPE")
   For U& = 1 To FIN&
     X$ = REGLEIDO$("!CARDETPE", U&)
     CAN = CVS(Mid$(X$, 47, 4))
     PORDESCUIN = CVS(Mid$(X$, 69, 4))
     PRECARGA = COEFCON * CVS(Mid$(X$, 39, 4))
     PREUNI = PRECARGA * (100 - PORDESCUIN) / 100
     TOT# = CDbl(CAN * PREUNI)
     '
     Call REPLA(X$, MKS$(PRECARGA), 39, 4)
     Call REPLA(X$, MKS$(PRECARGA), 51, 4)
     Call REPLA(X$, MKD$(TOT#), 55, 8)
     Call REPLA(X$, MKS$(PRECARGA), 65, 4)
     Call REPLA(X$, MKS$(PREUNI), 73, 4)
     Call REPLA(X$, MKD$(TOT#), 77, 8)
     Call GRAREG("!CARDETPE", X$, U&)
   Next U&
   '
   Call CARLISPE
   Call CIERRARCH("!CARDETPE")
   Call CALTOPED
   '
End Sub

Private Sub Option1_Click()
   If Option1.Value = True Then
     TXTFLETE = "Flete A Cargo del Cliente"
   Else
    TXTFLETE = ""
   End If

End Sub

Private Sub Option2_Click()
  If Option2.Value = True Then
     EMPRESAX$ = EMPREAC$
     If Len(EMPREAC$) > 6 Then
       POS1% = InStr(EMPREAC$, " ")
       EMPRESAX$ = Mid$(EMPREAC$, 1, POS1%)
     End If
     TXTFLETE = "Flete A Cargo de " & EMPRESAX$
   Else
     TXTFLETE = ""
   End If
End Sub

Private Sub PORDESCU_Change()
   Call CALTOPED
End Sub

Private Sub PORDESCU_GotFocus()
   If VALICLIE% < 1 Then
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
      Exit Sub
   End If
   '
   If DERACCE%("MODESPE/NOMESS", "Modificacion descuento de pedido") < 2 Then
      On Error Resume Next
      Text13.SetFocus
      On Error GoTo 0
      'If List1.ListCount > 0 Then AA1% = DERACCE%("MODESPE", "Modificacion descuento de pedido")
      If MSF.Rows > 1 Then AA1% = DERACCE%("MODESPE", "Modificacion descuento de pedido")
   End If
   '
End Sub

Private Sub Text1_Change()
   If Text1 = "" Then
        LTEXT1 = ""
      Else
        If Text1 <> LTEXT1 Then
               NROPED = XVALO(Text11)
               If YAHAYENTREGAS%(NROPED) < 1 Then ' en caso que no tenga entregas realizadas se puede cambiar
                    Call CALTOPED ' Actualiza el importe de IIBB al cambiar el cliente
               Else
                    Call COMUNI("Imposible Modificar Cliente.\Este Pedido Tiene Cantidades Entregadas.")
                    Text1 = LTEXT1
               End If
        End If
   End If
   '
   NCLIE = XVALO(Text1.TEXT)
   If NCLIE > 0 Then
      If REGICLI&(NCLIE) > 0 Then
         Call CARDACLI
         CLIPROACTIVO = XVALO(Text1)
      End If
   End If
End Sub

Private Sub Text1_DblClick()
   If List1.ListCount < 1 Then Exit Sub
   'Call SELECHO("DEUDOR1")
   'NCLIE = XVALO(TRIM$(VALACT1$("DEUDOR1")))
   NCLIE = ListaSeleccionClientes()
   If NCLIE > 0 Then
     If REGICLI&(NCLIE) > 0 Then
       Text1.TEXT = TRIM$(Str$(NCLIE))
       Call CARDACLI
     End If
   End If
End Sub

Private Sub Text1_GotFocus()
   If List1.ListCount < 1 Then
      On Error Resume Next
      Command2.SetFocus
      On Error GoTo 0
      Exit Sub
   End If
End Sub

Private Sub Text10_GotFocus()
   Text10.TEXT = TRIM$(Text10.TEXT)
   Text10.SelStart = 0
   Text10.SelLength = Len(Text10.TEXT)
End Sub

Private Sub Text10_LostFocus()
   ' COMENTADDO POR QUE SE VALIDA EN EL BOTON DE APROBACION Y SI QUEDA ESTE
   ' PRESENTA EL MENSAJE DOS VECES
   
'''   If VALICLIE% < 1 Then
'''      On Error Resume Next
'''      Text1.SetFocus
'''      On Error GoTo 0
'''      Exit Sub
'''   End If
'''   '
'''   'VALIDACION DE ORDEN DE COMPRA PREEXISTENTE
'''   NC% = XVALO(Text1)
'''   NORCOM$ = UCase$(TRIM$(Text10))
'''   '
'''   'VALIDO EL CONTROL Y SI EL CASILLERO TIENE INFORMACION
'''   If Control$("ADMIPED", "NOREPOC") = "SI" And TRIM$(Text10) <> "" Then
'''      CONDI$ = "NROCLIE = " & NC% & " And NROOCOM = " & NORCOM$ & " And NROPED <> " & XVALO(Text11)
'''      '
'''      'SI EXISTE LA ORDEN DE COMPRA PRESENTO MENSAJE
'''      EXISTEOCOMPRA% = CantRegistros("PEDENCA", CONDI$)
'''      If EXISTEOCOMPRA% > 0 Then
'''         Call COMUNI("Orden de Compra Existente")
'''            Text10 = ""
'''            Text10.SetFocus
'''      End If
'''   End If
'''
End Sub

Private Sub Text11_GotFocus()
   On Error Resume Next
   Command2.SetFocus
   On Error GoTo 0
End Sub

Private Sub Text12_GotFocus()
   If VALICLIE% < 1 Then
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
   End If
End Sub

Private Sub Text13_GotFocus()
   If VALICLIE% < 1 Then
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
   End If
End Sub

Private Sub Text15_GotFocus()
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Private Sub Text16_Change()
   '
   'VALIDO QUE ESTE VISIBLE EL PICTURE -- EN ESE CASO EXISTE EL TAMOTPRE
   If Picture6.Visible = False Then Exit Sub

   'TEXT DE INGESO DE MOTIVOS
   If Not IsNumeric(Text16) Then Text16 = ""
   
   NUMOT% = XVALO(Text16)
   If NUMOT% >= 0 Then
      Label28 = ECOARCH$("TAMOTPRE", Chr$(NUMOT%))
   Else
      Label28 = ""
   End If

End Sub

Private Sub Text18_Change()
    If ROUND#(Text18) < 100 Then
      Call CALTOPED
    Else
      Call COMUNI("Descuento No Válido !!! Verifique")
      Text18 = ""
    End If
End Sub

Private Sub Text18_GotFocus()
   If DERACCE%("DESPORCA/NOMESS", "Modificacion Descuento por Cantidad") < 2 Then
      On Error Resume Next
      Text13.SetFocus
      On Error GoTo 0
   End If
End Sub


Private Sub Text19_Change()
  '\\\OT-07-0040-RG-02/02/07///
  Label23.Caption = ""
  Label24.Caption = ""
  LENTREGA(2) = ""
  TRANACT% = 0
  '\\\OT-07-0040-RG-FIN///
  '\\\OT-6-422-FG-05/10/06
  NTRA = XVALO(Text19)
  If NTRA > 0 Then
    If NTRA < 32767 Then
      NTRAI% = NTRA
      LENTREGA(2) = DENOTRANS$(NTRAI%)
      If LENTREGA(2) <> "" Then
         TRANACT% = NTRAI%
      End If
      '\\\OT-07-0040-RG-02/02/07///
      Label23.Caption = DOMITRANS$(NTRAI%)
      Label24.Caption = LOCATRANS$(NTRAI%)
      '\\\OT-07-0040-RG-FIN///
    End If
  End If
  '\\\OT-6-422-FG-FIN
End Sub

Private Sub Text2_gotfocus()
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Private Sub Text20_Change()
    '\\\OT-6-422-FG-05/10/06
    NSUC& = XVALO(Text20)
      Label27 = ""
      If NSUC& > 0 Then
'        If NSUC& <= ULTREG&("!SUCENTRE") Then
         'Lectura desde la Tabla SQL.
          CONDI$ = "Nume_Cli = " & Text1 & " and Nume_Suc = " & NSUC
           If CantRegistros("SUCENTRE", CONDI$) > O Then
             Label27 = VALOBADA("SUCENTRE", "Deno_Suc", CONDI$, "NOMES")
             LENTREGA(0) = VALOBADA("SUCENTRE", "Domi_Suc", CONDI$, "NOMES")
             LENTREGA(1) = VALOBADA("SUCENTRE", "Loca_Suc", CONDI$, "NOMES")
             ' Fin de Lectura
           ElseIf NSUC& <= ULTREG&("!SUCENTRE") Then
             XX$ = REGLEIDO$("!SUCENTRE", NSUC&)
             Label27 = TRIM$(Mid$(XX$, 3, 62))
             LENTREGA(0) = TRIM$(Mid$(XX$, 65, 48))
             LENTREGA(1) = TRIM$(Mid$(XX$, 113, 48))
           End If
      End If
      '\\\OT-6-422-FG-FIN
End Sub

Private Sub TEXT21_DblClick(Index As Integer)
    '
    If VALICLIE% < 1 Then
       On Error Resume Next
       Text1.SetFocus
       On Error GoTo 0
       Exit Sub
    End If
    '
    If Index = 0 Then
        ARCHEC$ = "LISTAS": LOCA% = 2
      ElseIf Index = 1 Then
        ARCHEC$ = "CONPAG": LOCA% = 2
      ElseIf Index = 2 Then
        ARCHEC$ = "VENDEDOR": LOCA% = 1
    End If
    ARGU$ = CBIN$(TRIM$(TEXT21(Index).TEXT), 1, LOCA%)
    Call SELECHO(ARCHEC$)
    VDEVU$ = TRIM$(VALACT1$(ARCHEC$))
    If VDEVU$ <> "" Then
       TEXT21(Index).TEXT = VDEVU$
    End If
    '
End Sub

Private Sub Text21_GotFocus(Index As Integer)
   If VALICLIE% < 1 Then
      On Error Resume Next
      Text1.SetFocus
      Exit Sub
   End If
   TEXT21(Index).TEXT = TRIM$(TEXT21(Index).TEXT)
   TEXT21(Index).SelStart = 0
   TEXT21(Index).SelLength = Len(TEXT21(Index).TEXT)
End Sub

Private Sub Text3_GotFocus()
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Private Sub Text4_GotFocus()
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Private Sub Text44_LostFocus()
   If Left$(UCase$(Label21(4)), 5) = "LETRA" Then
      Text44 = UCase$(Left$(Text44, 1))
      Exit Sub
   End If
   '
   If XVALO(Text44) > 99.9 Then Text42 = "99.9"
   If TRIM$(Text44) <> "" Then Text44 = Format$(XVALO(Text44), "#0.0")
End Sub

Private Sub Text5_GotFocus()
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Private Sub Text21_Change(Index As Integer)
    '
    If XVALO(TEXT21(Index).TEXT) < 0 Or XVALO(TEXT21(Index).TEXT) > 255 Then
       TEXT21(Index).TEXT = ""
    End If
    '
    If Index = 0 Then
        If XVALO(TEXT21(Index).TEXT) < 1 Then
            Eco21(Index).TEXT = "CONCERTADOS FIJOS"
            ARCHIPRE$ = "MASTER"
            LIPRACTI% = 0
          Else
            Eco21(Index).TEXT = ECOARCH$("LISTAS", MKI$(XVALO(TEXT21(Index).TEXT)))
            ARCHIPRE$ = TRIM$(TEXT21(0).TEXT): LIPARCTI% = XVALO(TRIM$(ARCHIPRE$))
            While Len(ARCHIPRE$) < 3: ARCHIPRE$ = "0" + ARCHIPRE$: Wend
            ARCHIPRE$ = "LIPRE" + ARCHIPRE$
        End If
        'DESCUENTO EN PEDIDO PARA SETEO POR MONTO Y LISTA 1
        If CONTROL("PEDIDO", "MODODESC") = "MONTO" And XVALO(TEXT21(0)) = 1 Then
           Call FORPED07.DESCUMONTO(Me)
        Else
           PORDESCU.TEXT = DESCUCLI$(XVALO(Text1))
        End If
        
      ElseIf Index = 1 Then
        Eco21(Index).TEXT = ECOARCH$("CONPAG", MKI$(XVALO(TEXT21(Index).TEXT)))
        DESXREX = DESCURECA(XVALO(TEXT21(1)))
        '
        If ROUND#(DESXREX) < 100 Then
          Label25 = FORMATNUM(DESXREX, "F5.1")
        Else
          Call COMUNI("Descuento No Válido !!! Verifique")
          Label25 = ""
        End If

      ElseIf Index = 2 Then
        Eco21(Index).TEXT = ECOARCH$("VENDEDOR", Chr$(XVALO(TEXT21(Index).TEXT)))
        If CONTROL("PEDIDO", "MODODESC") = "MONTO" And XVALO(TEXT21(0)) = 1 Then
           Call FORPED07.DESCUMONTO(Me)
        End If
        Call Check2_Click ' llama al evento donde se valida el vendedor para definir descuento

      
      ElseIf Index = 3 Then
        'TIPO DE INGRESO
        If EXISTE%(LUDAT$ + "TATINGRE.RFS") > 0 Then
           Eco21(Index).TEXT = ECOARCH$("TATINGRE", Chr$(XVALO(TEXT21(Index).TEXT)))
        End If
    End If
    '
End Sub

Private Sub Command21_Click(Index As Integer)
    If VALICLIE% < 1 Then
       On Error Resume Next
       Text1.SetFocus
       On Error GoTo 0
       Exit Sub
    End If
    '
    If Index = 0 Then
        ARCHEC$ = "LISTAS": LOCA% = 2
      ElseIf Index = 1 Then
        ARCHEC$ = "CONPAG": LOCA% = 2
      ElseIf Index = 2 Then
        ARCHEC$ = "VENDEDOR": LOCA% = 1
      ElseIf Index = 3 Then
        If EXISTE%(LUDAT$ + "TATINGRE.RFS") < 1 Then
          Call COMUNI("Tabla de Motivos de Ingreso de Pedidos \No Disponible en la Presente Versión")
          Exit Sub
        End If
        '
        ARCHEC$ = "TATINGRE/Tabla de Ingresos": LOCA% = 1
    End If
    ARGU$ = CBIN$(TRIM$(TEXT21(Index).TEXT), 1, LOCA%)
    Call SELECHO(ARCHEC$)
    VDEVU$ = TRIM$(VALACT1$(ARCHEC$))
    If VDEVU$ <> "" Then
       TEXT21(Index).TEXT = VDEVU$
    End If
    '
End Sub
   '
Function VALICLIE%()
   VCC% = 0
   If XVALO(Text1.TEXT) > 0 Then
     If XVALO(Text1.TEXT) <= 999999 Then
       NCLIE = XVALO(Text1.TEXT)
       If rasocli$(NCLIE) <> "" Then
         VCC% = 1
       End If
     End If
   End If
   '
99 VALICLIE% = VCC%
   '
End Function

Sub BLANFORMU()
   '\\\OT-7-010-RG-12/01/07
   Text1.TEXT = ""
   '\\\OT-7-010-RG-FIN
   Text2.TEXT = ""
   Text3.TEXT = ""
   Text4.TEXT = ""
   Text5.TEXT = ""
   Text7.TEXT = ""
   Text8.TEXT = ""
   Text10.TEXT = ""
   Text12.TEXT = ""
   Text13.TEXT = ""
   Text15 = ""
   Text16 = ""
   Text6.TEXT = ""
   List1.Clear
   Label18 = ""
   '
   '\\\OT-6-422-FG-05/10/06
   Text19 = ""
   Text20 = ""
   Call SETULTREG("!SUCENTRE", 0)
   Call FRESHECHO("!SUCENTRE")
   '\\\OT-6-422-FG-FIN
   '
   TEXT21(0).TEXT = "": Eco21(0) = ""
   TEXT21(1).TEXT = "": Eco21(1) = ""
   TEXT21(2).TEXT = "": Eco21(2) = ""
   TEXT21(3).TEXT = "": Eco21(3) = "" ' TIPO DE INGRESO
   Text44 = ""
   '
   PORDESCU.TEXT = ""
   For KK% = 0 To 3
      Total(KK%).TEXT = ""
      TOTIVA(KK%).TEXT = ""
      LENTREGA(KK%).TEXT = ""
   Next KK%
   '\\\OT-05-0560-WAR-24/11/05///
   Label13 = ""
   Label14 = ""
   Label20 = ""
   '\\\OT-05-0560-WAR-FIN///
   Label11.ForeColor = RGB(0, 0, 255)
   Label11 = "Venta"
   Label11.ToolTipText = "Doble-Click para 'Reparación'"
   '
   Label31 = ""
   Text18 = ""
   Check2.Value = 0
   Frame4.Height = 1440
   TOTIVA(3) = ""
   Text15 = ""
   Total(3) = ""
   '
   MARCONOTA.Visible = False
   '
   Erase CI_VECTOR_TIVA()
   Erase TASAIVAX_CI()

End Sub
'
Sub CARDACLI()
    '
    Screen.MousePointer = 11
    NCLIEN = XVALO(Text1.TEXT)
    If NCLIEN > 0 Then
      If rasocli$(NCLIEN) <> "" Then
        Text2.TEXT = rasocli$(NCLIEN)
        Text3.TEXT = DOMICLI$(NCLIEN)
        Text4.TEXT = CPOSCLI$(NCLIEN)
        Text5.TEXT = LOCACLI$(NCLIEN)
        Text6.TEXT = ANOTACLI$(NCLIEN)
        
        Text7.TEXT = CSTRING$(MKS$(LIMICRE(NCLIEN)), 3, 13, 2, 2)
        Text12.TEXT = TRIM$(POSIVAC$(NCLIEN)) + " - CUIT " + CuitCli$(NCLIEN)
        CREDISPO# = LIMICRE(NCLIEN) - SALDCLI#(NCLIEN) - CHEPEACRE#(NCLIEN)
        Text8.ForeColor = RGB(0, 0, 0)
        If CREDISPO# < 0 Then Text8.ForeColor = RGB(255, 0, 0)
        Text8.TEXT = CSTRING$(MKD$(Abs(CREDISPO#)), 3, 13, 2, 2)
        '
        Text19 = "" ' Necesita ser vaciada para variar por el cambio de cliente
        '
        '
        'CONTROL PARA PRESENTAR DATOS DE ENTREGA
        If CONTROL("NOPREENT", "CONFPEDI") <> "SI" Then
            For KKI% = 0 To 2
              LENTREGA(KKI%) = LUENTRECLI$(NCLIEN, KKI%)
            Next KKI%
            Text19.TEXT = NUTRANCLI%(XVALO(Text1))
        End If
        
       '
        ' CARGA LISTA DE SUCURSALES DEL CLIENTE ACTIVO
        CONDI$ = "Nume_Cli = " & NCLIEN
        If CantRegistros("SUCENTRE", CONDI$) > 0 Then
            Call CARGALISEL("!SUCENTRE", "SUCENTRE", "Nume_Suc/I5;Deno_Suc/A62", CONDI$, "NOMESS")
        Else
            Call FILTERFILE("SUCENTRE", "", 5, MKI$(NCLIEN))
        End If
      End If
    End If
    Screen.MousePointer = 1
    '
End Sub
'
Private Sub Text6_GotFocus()

   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Private Sub Text7_GotFocus()
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Private Sub Text8_GotFocus()
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Private Sub Text9_GotFocus()
   On Error Resume Next
   Command9.SetFocus
   On Error GoTo 0
End Sub

Private Sub TICAMBIS_Change()
   If TICAMBIS.Enabled = True Then
     VH(MONEMI%) = XVALO(TICAMBIS)
   End If
End Sub

Private Sub Timer1_Timer()
   '
   Call CALTOPED
   '
End Sub

Private Sub Total_Change(Index As Integer)
   If Index = 0 Then
     Call CALTOPED
'      DESC$ = PORDESCU.TEXT
'      TTT1 = XVALO(Total(0).TEXT) * COEFDESCU(DESC$) / 100
'      Total(1).TEXT = CSTRING$(MKS$(TTT1), 3, 11, 2, 2)
'      TTT3 = XVALO(Total(0).TEXT) - XVALO(Total(1).TEXT)
'      Total(2).TEXT = CSTRING$(MKS$(TTT3), 3, 11, 2, 2)
   End If
'   If Index < 3 Then
'      TTT1 = XVALO(Total(2).TEXT)
'      TTT2 = XVALO(TOTIVA(3).TEXT)
'      Total(3).TEXT = Str$(TTT1 + TTT2)
'   End If
End Sub
'
Sub PRIPEDIDO(OPCI%)
    '
    FORIPRE$ = ""
    FORIPRE$ = TRIM$(CONTROL$("", "FORPECLI"))
    If FORIPRE$ = "" Then Exit Sub
    '
    'SI TIENE FRM DE DISEÑADOR DE FRM VA POR AHI
    If OPCI% = 1 Then
        ARCSQR$ = FORIPRE$
        MODO$ = "PRINT"
        If CANCELPRINT% > 0 Then MODO$ = ""
        If EXISTE%(LUDAT$ + ARCSQR$ + ".SQR") > 0 Then
           NCLIEN = XVALO(Text1)
           FILTX$ = TRIM$(Str$(XVALO(Text11))) & ";" & TRIM$(Str$(NCLIEN))
           Call STDFORM(ARCSQR$, FILTX$, MODO$)
           Exit Sub
        End If
     End If
'    If OPCI% = 2 Then ' ORDEN DE DESPACHO
'       PUERDESPA$ = Control("*", "PUERDESP")
'       If Control(IDENTER$, "PUERDESP") <> "" Then PUERDESPA$ = Control(IDENTER$, "PUERDESP")
'       FORODES$ = TRIM$(Control$("", "FORDESPA"))
'       If FORODES$ <> "" Then
'          FORIPRE$ = FORODES$
'       End If
'    End If

    PUERDESPA$ = ""
    PUERORVE = ""
    If OPCI% = 1 Then
       PUEROVTA$ = CONTROL("*", "PUERORVE")
       If CONTROL(IDENTER$, "PUERORVE") <> "" Then PUEROVTA$ = CONTROL(IDENTER$, "PUERORVE")
    ElseIf OPCI% = 2 Then ' ORDEN DE DESPACHO
       PUERDESPA$ = CONTROL("*", "PUERDESP")
       If CONTROL(IDENTER$, "PUERDESP") <> "" Then PUERDESPA$ = CONTROL(IDENTER$, "PUERDESP")
       FORODES$ = TRIM$(CONTROL$("", "FORDESPA"))
       If FORODES$ <> "" Then
          FORIPRE$ = FORODES$
       End If
    End If
    '
    If EXISTE%(LUDAT$ + FORIPRE$ + ".FRM") < 1 Then
       Call MENSERR(24, "Imposible Imprimir Pedido de Cliente.\No Existe Formulario Electrónico '" + FORIPRE$ + ".FRM'.")
       Exit Sub
    End If
    '
    If CONTROL$("PEDCLIEN", "SUPRINT") = "SI" And SENDBYMAIL$ = "" Then
       CANCELPRINT% = 1
    End If
    
    Screen.MousePointer = 11
    '
    NCLIEN = XVALO(Text1)
    FECHDOC$ = Text9.TEXT
    NUMEDOC$ = TRIM$(Text11.TEXT)
    While Len(NUMEDOC$) < 6
      NUMEDOC$ = "0" + NUMEDOC$
    Wend
    '
    TIPODOC$ = "Pedido de Cliente"
    If NCLIEN < 1 Then
      TIPODOC$ = "Pedido Interno"
    End If
    If OPCI% = 2 Then
       TIPODOC$ = "Orden de Despacho"
    End If
    '
    CODTABU1$(1) = "COD-MAT"
    CODTABU1$(2) = "DES-MAT"
    CODTABU1$(3) = "CANTIDAD"
    CODTABU1$(4) = "PRE-LIST"
    CODTABU1$(5) = "PORDESCU"
    CODTABU1$(6) = "PRE-NETO"
    CODTABU1$(7) = "NET-ITEM"
    CODTABU1$(8) = "FECH-REC"
    CODTABU1$(9) = "ORD-ITEM"
    CODTABU1$(10) = "Z-PLANO" ' AGREGADOS PARA APAREAR CON FORPED
    CODTABU1$(11) = "DES-MPRI"
    CODTABU1$(12) = "UBI-MPRI"
    CODTABU1$(13) = "CAN-PLCO"
    CODTABU1$(14) = "UNIMED" 'SE AGREGA POR QUE ESTABA FALTANDO
    CODTABU1$(15) = "LAR-MATE"
    CODTABU1$(16) = "NRO-LOTE"
    
    CACOLTAB1% = 16
    '
    'DESTIDOC% = NCLIEN
    DESTIDOCNUE = NCLIEN
    CAITAB1% = 0
    CAITCRU% = 0
    '
    PRIVA$ = CONTROL$("PEDCLIEN", "PRECOIVA")
    DESCPRI% = 0: DESCOMP% = 0: Especif% = 0
    If CONTROL("ADMIPED", "IMPDESPR") = "SI" Then DESCPRI% = 1
    If CONTROL("ADMIPED", "IMPDESCO") = "SI" Then DESCOMP% = 1
    If CONTROL("ADMIPED", "IMPRESPE") = "SI" Then Especif% = 1
    If DESCOMP% + Especif% < 1 Then DESCPRI% = 1 'POR SI NO SE SELECCIONO NADA
    
    '
    For U& = 1 To ULTREG&("!CARDETPE")
      XX$ = REGLEIDO$("!CARDETPE", U&)
      CIPRE% = CVI(Left$(XX$, 2))
      CAPEX$ = Str$(CVS(Mid$(XX$, 47, 4)))
      '
      If CIPRE% > 0 Then
        If CIPRE% <= NUMAS% Then
          If XVALO(CAPEX$) > 0 Then
            '
            PRELIX = PRELISTA#(ARCHIPRE$, CIPRE%)
            PRECARGA = CVS(Mid$(XX$, 39, 4))
            PREUNI = CVS(Mid$(XX$, 73, 4))
            If PRELIX < PRECARGA Then PRELIX = PRECARGA
            If PRELIX < PREUNI Then PRELIX = PREUNI
            LOTEX$ = ""
            If UCase$(TRIM$(VENCARPE$)) = "CARDETLO" Then
              LOTEX$ = TRIM$(Mid$(XX$, 3, 16))
            End If
            '
            PRELI$ = Str$(PRELIX)
            PREUX$ = Str$(CVS(Mid$(XX$, 51, 4)))
            PRESINRED = CVS(Mid$(XX$, 73, 4))
            DESCX$ = "0": On Error Resume Next
                          DESCX$ = Str$(100 * (XVALO(PRELI$) - PRESINRED) / XVALO(PRELI$))
                          On Error GoTo 0
            PIVA# = 1
            If PIVACLI%(NCLIEN) <> 5 And PIVACLI%(NCLIEN) <> 7 Then
            'SI LA POSICION DEL IVA NO ES DE EXPORTACION VERIIFICA SITUACION
               If PRIVA$ = "SI" Then PIVA# = 1 + PORCEIVA(CIPRE%) / 100 ' VERIFICA QUE TIENE SETEADO EN EL SETUP PARA VER SI IMPRIME CON IVA INCLUIDO
            End If
            ITOTX$ = FORMATNUM$((XVALO(CAPEX$)) * (XVALO(PREUX$)) * PIVA#, "F15.4")
            '
            FEX = CVI(Mid$(XX$, 63, 2))
            FENTX$ = FECHATEX$(FEX)
            '
            CAITAB1% = CAITAB1% + 1
            TETABU1$(1, CAITAB1%) = CODEXT$(CIPRE%)
            TETABU1$(3, CAITAB1%) = CAPEX$
            TETABU1$(4, CAITAB1%) = PRELI$
            TETABU1$(5, CAITAB1%) = DESCX$
            PRENETO# = XVALO(PREUX$) * PIVA# ' PRECIO CON IVA
            PREUX$ = FORMATNUM$(PRENETO#, "F15.4")
            TETABU1$(6, CAITAB1%) = PREUX$
            TETABU1$(7, CAITAB1%) = ITOTX$
            TETABU1$(8, CAITAB1%) = FENTX$
            TETABU1$(9, CAITAB1%) = Str$(CAITAB1%)
            TETABU1$(10, CAITAB1%) = NUPLANO$(CIPRE%)
            TETABU1$(14, CAITAB1%) = Unimed$(CIPRE%) 'SE AGREGA POR QUE ESTABA FALTANDO
            
            '
            'LECTURA DE NRO DE PRESUPUESTO Y NORDEN PARA AHP PARA TOMAR EL NUMERO DE PROVEEDOR DE LA TABLA DE PRESUPUESTO
'            NUORIT% = CVI(Mid$(XX$, 123, 2))
            NPROV& = CVI(Mid$(XX$, 123, 2))
            'FIN AHP

            ' LO QUE SIGUE ES PARA IMPRIMIR O/DESPACHO
            If OPCI% = 2 Then
               TETABU1$(4, CAITAB1%) = ""
               TETABU1$(5, CAITAB1%) = ""
               TETABU1$(6, CAITAB1%) = ""
               TETABU1$(7, CAITAB1%) = ""
               TETABU1$(12, CAITAB1%) = UBISTO$(CIPRE%) ' AGREGADO PARA APAREAR CON FORPED
               'CANTIDAD DE BULTOS
               CANTI# = XVALO(CAPEX$)
               CABU = CANBULTOS(CIPRE%, CANTI#, "VENTA")
               TETABU1$(13, CAITAB1%) = TRIM$(FORMATNUM$(CABU, "F7"))
               'AHP TRAE EL NUMERO DE PROVEEDOR DE LA TABLA DE PRESUPUESTO
'                 CONDI$ = " NroPres = " & NUPRESU& & " and NuOrdIt= " & NUORIT%
'                 NPROV& = XVALO(VALOBADA("PRESUVEN", "Num_Prov", CONDI$, "NOMESS"))
                 
                 MARCAX$ = CONDICOM$(-1 * NPROV&)
                 TETABU1$(15, CAITAB1%) = TRIM$(MARCAX$)
                 TETABU1$(16, CAITAB1%) = TRIM$(LOTEX$)
                 'ACUMULO EL PESO
                 PESOUNITARIOX = PESUNI(CIPRE%)
                 ACUMUPESO = ACUMUPESO + (PESOUNITARIOX * CANTI#)
               'FIN AHP

            End If
            ' FIN O/DESPACHO
            
            PHCODI% = XVALO(ATRIFORM$(FORIPRE$, "COD-MAT/POSINHOR"))
            PHDESI% = XVALO(ATRIFORM$(FORIPRE$, "DES-MAT/POSINHOR"))
            PVCODI% = XVALO(ATRIFORM$(FORIPRE$, "COD-MAT/POSINVER"))
            PVDESI% = XVALO(ATRIFORM$(FORIPRE$, "DES-MAT/POSINVER"))
            If PHDESI% <= PHCODI% + 3 Then
               If PVDESI% <= PVCODI% + 1 Then
                  CAITAB1% = CAITAB1% + 1
               End If
            End If
            '
            'leo las 3 descripciones posibles
            DESPRIN$ = DESCRIT0$(CIPRE%)
            DESCOMPLEM$ = DESCRIT2$(CIPRE%)
            ESPECIFICAC$ = ESPECIFIT$(CIPRE%)
            '
            'DESCRIPCION
            DIMPRE$ = ""
            If DESCPRI% > 0 Then
               If DESPRIN$ <> "" Then
                  DIMPRE$ = DESPRIN$
               End If
            End If
            '
            'DESCRIPCION COMPLEMENTARIA
            If DESCOMP% > 0 Then
               If DESCOMPLEM$ <> "" Then
                 If DIMPRE$ <> "" Then
                   DIMPRE$ = DIMPRE$ + vbCrLf
                 End If
                 DIMPRE$ = DIMPRE$ + DESCOMPLEM$
               End If
            End If
            '
            'ESPECIFICACION
            If Especif% > 0 Then
               If ESPECIFICAC$ <> "" Then
                 If DIMPRE$ <> "" Then
                   DIMPRE$ = DIMPRE$ + vbCrLf
                 End If
                 DIMPRE$ = DIMPRE$ + ESPECIFICAC$
               End If
            End If
            '
            'IMPRESION DE LA DESCRIPCION CON TODAS SUS VARIANTES CONFIGURADAS
            FORTELIB$ = TRIM$(UCase$(ATRIFORM$(FORIPRE$, "DES-MAT/FORMATO")))
            ATELIB% = XVALO(Mid$(FORTELIB$, 2))
            Call FRATEXTO(DIMPRE$, ATELIB%)
            For KKU1% = 1 To CARETEX%
                  TETABU1$(2, CAITAB1%) = RETEX$(KKU1%)
                  CAITAB1% = CAITAB1% + 1
            Next KKU1%
            'SE RESTA CAITAB1% PORQUE DEL FOR SALE EN LA PROXIMA LINEA
            'LO CUAL YA SE HACE CUANDO PASA AL PROXIMO CODIGO (VER MAS ARRIBA)
            CAITAB1% = CAITAB1% - 1
            'LINEA DE SEPARACION
            Call CARLINSEP
            '
            ' IMPRIME LOS COMPONENTES
            If OPCI% = 2 Then ' AGREGADO PARA APAREAR CON FORPED
               If DESCOMPO%(CIPRE%) > 0 Then
                   CANTI = CVS(Mid$(XX$, 47, 4))
                   Call FORPED07.PRINCOMPO(CIPRE%, CANTI)
               End If
            End If
            '
            If CVS(Mid$(XX$, 89, 4)) > 0 Then ' AGREGADO PARA APAREAR CON FORPED
               Call FORPED07.IMPRIMEBONIFICACION(U&)
            End If

          End If
        End If
      End If
109 Next U&
    '
    SUMATOTA# = XVALO(Total(0).TEXT)
    IDESCUEN# = XVALO(Total(1).TEXT)
    '
    '\\\OT-05-0853-WAR-24/11/05///
    If PIVACLI%(NCLIEN) = 5 Or PIVACLI%(NCLIEN) = 7 Then
      Label13 = "" 'SI CLIENTE EXTERIOR O PROMOVIDO
      Label14 = "" 'NO DISCRIMINA IVA
      Label20 = ""
    End If
    '\\\OT-05-0853-WAR-FIN///
    '
    CODPARAM$(1) = "IMP-NETO": DOCPARAM$(1) = Str$(SUMATOTA#)
    CODPARAM$(2) = "VENDEDOR":     DOCPARAM$(2) = Eco21(2).TEXT
    CODPARAM$(3) = "COND-PAG":     DOCPARAM$(3) = Eco21(1).TEXT
    CODPARAM$(4) = "NRO-OCOM":     DOCPARAM$(4) = Text10.TEXT
    CODPARAM$(5) = "DOMI-TRA": DOCPARAM$(5) = LENTREGA(0).TEXT
    CODPARAM$(6) = "DESTINO": DOCPARAM$(6) = LENTREGA(1).TEXT
    CODPARAM$(7) = "TRANSPOR": DOCPARAM$(7) = LENTREGA(2).TEXT
    CODPARAM$(8) = "IMP-DESC": DOCPARAM$(8) = Total(1).TEXT
    CODPARAM$(9) = "IMPNEDES": DOCPARAM$(9) = Total(2).TEXT
    CODPARAM$(10) = "IMP-IVA": DOCPARAM$(10) = Label13 ' TOTIVA(3).TEXT
    CODPARAM$(11) = "IMP-TOTA": DOCPARAM$(11) = Total(3).TEXT
    CODPARAM$(12) = "POR-DESC": DOCPARAM$(12) = TRIM$(PORDESCU.TEXT)
    '\\\OT-5-869-28/12/05///
    'CODPARAM$(13) = "MONEDEMI": DOCPARAM$(13) = MONEMIS
    '\\\OT-5-869-FIN///
    CODPARAM$(13) = "PERC-IVA": DOCPARAM$(13) = Text15
    CODPARAM$(14) = "IMP-IVA1": DOCPARAM$(14) = Label14
    '
    '\\\OT-6-422-FG-05/10/06
    CODPARAM$(15) = "NRO-PASE": DOCPARAM$(15) = Label27
    '\\\OT-6-422-FG-FIN
    '
    '\\\OT-07-0007-RG-08/01/07///
    HORA$ = HORAENT$(NCLIEN)
    CODPARAM$(16) = "Z-PLANO": DOCPARAM$(16) = TRIM$(HORA$)
    '\\\OT-07-0007-RG-FIN///
    '
    '\\\OT-07-0010-RG-12/01/07///
    AISUC% = XVALO(Text20)
    TELSUC$ = TELESUC$(NCLIEN, AISUC%)
    CODPARAM$(17) = "REF-MAT1": DOCPARAM$(17) = TRIM$(TELSUC$)
    '\\\OT-07-0010-RG-FIN///
    '
    'imprime domicilio de transportista (expreso)
    NTRA% = XVALO(Text19)
    If Text19.TEXT = "" Then
       If TRIM(NTRANCLI$(NCLIEN)) = TRIM(LENTREGA(2)) Then NTRA% = NUTRANCLI%(NCLIEN)
    End If
    '
    CODPARAM$(18) = "REF-MAT2": DOCPARAM$(18) = DOMITRANS$(NTRA%)
    REG$ = FILTERGETRECORD("CONPAG", 1, Chr$(XVALO(TEXT21(1))))
    
    PORCONPAG = CVS(Mid$(REG$, 44, 4))
    CODPARAM$(19) = "CANT-DIS": DOCPARAM$(19) = TRIM$(FORMATNUM$(PORCONPAG, "F5.1"))
    CODPARAM$(20) = "DES-OPER": DOCPARAM$(20) = TRIM$(Label31) 'DESCUENTO RETIRA CLIENTE GABAL
    CODPARAM$(21) = "REF-PLCO": DOCPARAM$(21) = TRIM$(Text141) 'IMPORTE DE DESCTO RETIRA CLIENTE GABAL
    CODPARAM$(22) = "REF-MAT5": DOCPARAM$(22) = TRIM$(Text18) 'DESCUENTO POR CANTIDAD
    CODPARAM$(23) = "REF-MAT6": DOCPARAM$(23) = TRIM$(Text17) 'IMPORTE DE DESCTO POR CANTIDAD
    CODPARAM$(24) = "LOCA-TRA": DOCPARAM$(24) = LENTREGA(1).TEXT
    CODPARAM$(25) = "PES-CONJ": DOCPARAM$(25) = TRIM$(FORMATNUM$(ACUMUPESO, "F15.4")) 'PESO ACUMULADO
    NTRAI% = NUTRANCLI%(NCLIEN)
    CODPARAM$(26) = "TELE-EMP": DOCPARAM$(26) = TRIM$(TELETRANS$(NTRAI%))
    CODPARAM$(27) = "LOCA-TRA": DOCPARAM$(27) = LENTREGA(1).TEXT
    CODPARAM$(28) = "LAR-CUCO": DOCPARAM$(28) = TRIM$(Me.TXTFLETE)
    '
    CAPARDOC% = 28

    FORIPRE$ = TRIM$(CONTROL$("", "FORPECLI"))
    '
'    If OPCI% = 2 Then ' ORDEN DE DESPACHO
'       If Control("ORDESPA", "PUERTOIM") <> "" Then PUERDESPA$ = Control("ORDESPA", "PUERTOIM")
'       FORODES$ = TRIM$(Control$("", "FORDESPA"))
'       If FORODES$ <> "" Then
'          FORIPRE$ = FORODES$
'       End If
'    End If

    If OPCI% = 2 Then ' ORDEN DE DESPACHO
       PUERDESPA$ = CONTROL("*", "PUERDESP")
       If CONTROL(IDENTER$, "PUERDESP") <> "" Then PUERDESPA$ = CONTROL(IDENTER$, "PUERDESP")
       FORODES$ = TRIM$(CONTROL$("", "FORDESPA"))
       If FORODES$ <> "" Then
          FORIPRE$ = FORODES$
       End If
    End If

    ' LO QUE SIGUE ES PARA QUE EN CASO DE QUE HAYA UN SOLO CASILLERO
    ' DE IVA, TOTALICE TODOS LOS IMPORTES PARCIALES DE IVA EN UNO SOLO
    FORIVASEC$ = TRIM$(UCase$(ATRIFORM$(FORIPRE$, "IMP-IVA1/FORMATO")))
    If FORIVASEC$ = "" Then
      DOCPARAM$(10) = TOTIVA(3).TEXT
      DOCPARAM$(14) = ""
    End If
    '
    ' LO QUE SIGUE ES PARA ORDEN DE DESPACHO
    If OPCI% = 2 Then
       CODPARAM$(1) = ""
       CODPARAM$(8) = ""
       CODPARAM$(9) = ""
       CODPARAM$(10) = ""
       CODPARAM$(11) = ""
       CODPARAM$(12) = ""
       CODPARAM$(13) = ""
       CODPARAM$(14) = ""
    End If
    ' FIN ORDEN DE DESPACHO
    Call BLANARCH("!OBSERVOF")
    Call FRATEXTO(Text13.TEXT, 32)
    If CARETEX% >= 1 Then Call GRAREG("!OBSERVOF", RETEX$(1), 1)
    If CARETEX% >= 2 Then Call GRAREG("!OBSERVOF", RETEX$(2), 2)
    If CARETEX% >= 3 Then Call GRAREG("!OBSERVOF", RETEX$(3), 3)
    URETE& = CARETEX%
    Call SETULTREG("!OBSERVOF", URETE&)
    Call CIERRARCH("!OBSERVOF")
    '
'    If FORIPRE$ <> "" Then
'      Call PRINTDOC("FORPECLI")   ' PEDIDO DE CLIENTE
'    End If
    '

    Call SETPRINPORT("RESTORE")
     If FORIPRE$ <> "" Then
      'Call PRINTDOC("FORPECLI")   ' PEDIDO DE CLIENTE
      If OPCI% = 1 Then Call SETPRINPORT(PUEROVTA$)
      If OPCI% = 2 Then Call SETPRINPORT(PUERDESPA$)
      Call PRINTDOC("@" + FORIPRE$) ' PEDIDO DE CLIENTE
'       If PUERDESPA$ <> "" Then Call SETPRINPORT("RESTORE")
    End If
    Call SETPRINPORT("RESTORE")
    PUERDESPA$ = ""
    PUEROVTA$ = ""


    For KK1% = 1 To CACOLTAB1%
      For KK2% = 1 To CAITAB1%
        TETABU1$(KK1%, KK2%) = ""
      Next KK2%
    Next KK1%
    '


End Sub


Private Sub Total_GotFocus(Index As Integer)
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Sub GRAPEDIDO(OKX%, Optional FUERZAPROBACION$)
   '
   If CONTROL$("ADMIPED", "VERLICRE") = "SI" Then VERLICRE% = 1
   '
   OKX% = 0
   If MARCONOTA.Visible = True Then Exit Sub
   If XVALO(Total(2).TEXT) < 0.01 Then
     If COMALTER%("Este Pedido es SIN CARGO !!!\Está Seguro de Registrarlo Igual ?\\Cancelar Pedido\Registrar Igual") <> 2 Then
       Exit Sub
     End If
   End If
   If ULTREG&("!CARDETPE") < 1 Then Exit Sub
   '
   Call CIERRARCH("PEDENCA")
   Call CIERRARCH("PEDDETA")
   Call CIERRARCH("PEDDETP")
   '
   Call BLOQARCH("PEDENCA")
   Call BLOQARCH("PEDDETA")
   Call BLOQARCH("PEDDETP")
   '
   'REGENCA& = XVALO(Label9.Caption)
   'If REGENCA& < 1 Or REGENCA& > ULTREG&("PEDENCA") Then
   '  Call MENSERR(24, "Imposible Modificar este Pedido.")
   '  Call FINAL("")
   'End If
   '
   'PDC$ = REGLEIDO$("PEDENCA", REGENCA&)
   'If CVS(Left$(PDC$, 4)) <> XVALO(Text11.TEXT) Then
   '  Call MENSERR(24, "Imposible Modificar este Pedido.")
   '  Call FINAL("")
   'End If
   '
   Screen.MousePointer = 11
   If VALICLIE% < 1 Then
      MERRO$ = "Número de Cliente no Válido"
      GoTo 98
   End If
   NCLIEN = XVALO(Text1.TEXT)
   '
   If XVALO(TEXT21(0).TEXT) > 255 Then
      MERRO$ = "Lista de Precios no Válida"
      GoTo 98
   End If
   '
   If XVALO(TEXT21(1).TEXT) > 255 Then
      MERRO$ = "Condición de Pago no Válida"
      GoTo 98
   End If
   '
   If XVALO(TEXT21(2).TEXT) > 255 Then
      MERRO$ = "Vendedor no Válido"
      GoTo 98
   End If
   '
   If XVALO(TEXT21(3).TEXT) > 255 Then
      MERRO$ = "Tipo de Ingreso No Válido"
      GoTo 98
   End If

   HOII% = HOY(HOS$)
   FEMI% = FECHANUM(Text9.TEXT)
   If FEMI% < 1 Or FEMI% > HOII% Then
      MERRO$ = "Fecha de Emisión no Válida"
      GoTo 98
   End If
   '
   ' VERIFICA ITEMS CON CANTIDAD CERO EN SEÑAL DE QUE QUIERE QUITARLO DEL PEDIDO
   ' SOLO ES POSIBLE SI NO EXISTE CANTIDAD ENTREGADA
   ItemAEliminarEntregados = 0
   For U& = 1 To ULTREG&("!CARDETPE")
     X$ = REGLEIDO$("!CARDETPE", U&)
     CodigoInterno% = CVI(Left$(X$, 2))
     CantidadPedida = CVS(Mid$(X$, 47, 4))
     If CantidadPedida < 0.05 Then
        CantidadEntregada = CVS(Mid$(X$, 117, 4))
        If CantidadEntregada > 0.05 Then
            ItemAEliminarEntregados = 1
            Call COMUNI("Imposible Eliminar el Item '" & CODEXT$(CodigoInterno%) & "' (Cantidad = 0).\Posee Entregadas Parciales o Totales.")
        End If
     End If
   Next U&
   If ItemAEliminarEntregados = 1 Then GoTo 98
   
   ' TER - CONTRASTA CANTIDAD PEDIDO CON CANTIDAD DE NUMEROS DE SERIE CREADOS HASTA EL MOMENTO.
   ItemConPedidoMenorAArmado% = 0
   NumeroPedido& = XVALO(Text11.TEXT)
   If InStr(1, EMPREAC$, "TER") > 0 Then
        For U& = 1 To ULTREG&("!CARDETPE")
            X$ = REGLEIDO$("!CARDETPE", U&)
            CodigoInterno% = CVI(Left$(X$, 2))
            CantidadPedida = CVS(Mid$(X$, 47, 4))
            CantidaItemsArmados& = CantRegistros("TANUMSE", "NUMPEDIDO=" & CStr(NumeroPedido&) & " AND COD_INTE=" & CStr(CodigoInterno%) & " AND NUMEROSERIE<>''")
            If CantidadPedida < CantidaItemsArmados& Then
                ItemConPedidoMenorAArmado% = 1
                MERRO$ = "Imposible Grabar Pedido.\ \Se Han Encontrado Articulos Cuya Cantidad Pedida\es Menor a la Cantidad Armada."
            End If
        Next U&
   End If
   If ItemConPedidoMenorAArmado% = 1 Then GoTo 98

   
   NROPE& = XVALO(Text11.TEXT)
   REFE$ = " - " + TRIM$(Text10.TEXT)
   If REFE$ = " - " Then REFE$ = ""
   FEX = FEMI%
   REFE$ = REFE$ + " - " + FECHATEX$(FEX)
   LOMA% = 48 - Len(REFE$)
   REFE$ = Left$(TRIM$(rasocli$(NCLIEN)), LOMA%) + REFE$
   '
   NROPE& = XVALO(Text11.TEXT)
   '
   'ACA ENTRA SI ESTA MARCADALA OPCION DE VALIDAR SI SE REALIZARON CAMBIOS EN LISTAS DE PRECIOS AL CARGAR EL PEDIDO
   OKLISTAPRE% = 1
   If TRIM$(CONTROL("ADMIPED", "VALLISPR")) = "SI" Then
      U& = 0
      VUELTA% = 0
      For U& = 1 To ULTREG&("!CARDETPE")
         X$ = REGLEIDO$("!CARDETPE", U&)
         CIXI% = CVI(Left$(X$, 2))
         LISTA_ITEM% = CVS(Mid$(X$, 85, 4))
         PRE_VENTA# = CVS(Mid$(X$, 73, 4))
         RETORNO% = Vali_Derechos_Lpre(NCLIEN, CIXI%, LISTA_ITEM%, PRE_VENTA#)
         If RETORNO% < 1 Then
            If VUELTA% = 0 Then
              OKLISTAPRE% = 0
              Call REPLA(TXTLPRE$, "********* Item Sin Autorización *********", 11, 56): ACUTXTLPRE$ = ACUTXTLPRE$ + TXTLPRE$ + Chr$(10) + Chr$(13)
              VUELTA% = 1
            End If
            TXTLP$ = Left$(DESCRIT0(CIXI%), 24)
         
            Select Case RETORNO%
              Case 0 'SE CAMBIO LISTA DE PRECIOS SIN DERECHOS
                 TXTLP$ = TXTLP$ & " L.P.(" & LISTA_ITEM% & ")-Camb./Lst. S/Der."
              Case -1 'SE CAMBIO PRECIO SIN DERECHOS
                 TXTLP$ = TXTLP$ & " L.P.(" & LISTA_ITEM% & ")-Camb.Prc. S/Der."
              Case -2 'SE CAMBIO LISTA Y PRECIO SIN DERECHO
                 TXTLP$ = TXTLP$ & " L.P.(" & LISTA_ITEM% & ")-Camb./Lst y Prc. S/Der."
              Case -3  'EL ITEM NO PERTENECE A LA LISTA GRABADA
                 TXTLP$ = TXTLP$ & " L.P.(" & LISTA_ITEM% & ")-No Existe en L.Prc."
            End Select
            Call REPLA(TXTLPRE$, TXTLP$, 11, 56): ACUTXTLPRE$ = ACUTXTLPRE$ + TXTLPRE$ + Chr$(10) + Chr$(13)
         End If

      Next U&
      If OKLISTAPRE% < 1 Then STATUPC% = (-1)
   End If
   
   Call ABREPEDCLI
   'jandy sql
   strSQL = "Select * from Pedenca where NroPed=" & NROPE&
   Dim PedEnca As ADODB.Recordset
   Set PedEnca = New ADODB.Recordset
25 On Error Resume Next
   PedEnca.Open FILTSQL$(strSQL), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
   End If
   On Error GoTo 0
   ANOTANTE$ = ""
   '
   
   With PedEnca
     If (.EOF And .BOF) Then 'si no lo encuentra
       MERRO$ = "Imposible Modificar Este Pedido"
       GoTo 98
     End If
     '
     STATUPE% = XVALO(!Status)
     If FUERZAPROBACION$ <> "" Then STATUPE% = (-1) ' CASO POR QUE ALGUNA CONFIGURACION DE CAMBIOS EN PEDIDO LO DEJE COMO PEND.APROB.
     !Status = STATUPE%
     ANOTANTE$ = SAFETEXT$(!AnotaPed)
     !REFERENCIA = REFE$
     !FECHEMIS = CVDAT(FEMI%)
     !NROCLIE = NCLIEN
     !RaSoClie = rasocli$(NCLIEN)
     !NROOCOM = TRIM$(Text10.TEXT)
     !DOMIENT = LENTREGA(0)
     !LOCAENT = LENTREGA(1)
     !TranspEnt = LENTREGA(2)
     !NTranspo = XVALO(Text19)
     !NSucent = XVALO(Text20)
     !Observa = Text13
     !ListPre = XVALO(TEXT21(0))
     !CondPag = XVALO(TEXT21(1))
     !VENDED = XVALO(TEXT21(2))
     !TIPOINGRE = XVALO(TEXT21(3)) ' tipo de ingreso (telefono , fax etc)
     !PorComVend = XVALO(Text44.TEXT)
        If Left$(UCase$(Label21(4)), 5) = "LETRA" Then
           valor% = 0
           If TRIM$(Text44) <> "" Then valor% = Asc(Text44.TEXT + "A") - 64
           !PorComVend = valor%
        End If
     !Rec_ConPa = TRIM$(Label25)
     !TDescApli = TRIM$(PORDESCU.TEXT)
     !VenRep = Left$(Label11, 1)   ' Venta o Reparacion
     !OV_Motivo = XVALO(Text16)
     !Dto_Ret_Cte = TRIM$(Label31)
     !DtoxCant = TRIM$(Text18)
     !FLETE_TEX = TRIM$(Left$(TXTFLETE, 64))
     .Update
   End With
   PedEnca.Close
   Set PedEnca = Nothing
   '
   If TRIM$(CODCARPE$) <> "" Then ' Anotador de Carpeta de Exportaciones
       Call REPLA(TTXX$, HOS$, 1, 8)
       Call REPLA(TTXX$, USERNAME$, 69, 24)
       NR$ = SAFETEXT(Text11)
       Call REPLA(TTXX$, "Modificacion de Orden de Venta Nº: " + NR$, 11, 56)
       Call AGREGA_ANOTACARPI(CODCARPE$, TTXX$)
       TTXX$ = Space(76)
       Call REPLA(TTXX$, "Cliente: " + Left$(rasocli$(NCLIEN), 64), 11, 56)
       Call AGREGA_ANOTACARPI(CODCARPE$, TTXX$)
       Call AGREGA_ANOTACARPI(CODCARPE$, String$(88, "="))
   End If
   '
    'pedidoss con items sin cargo MEDITEA
   Dim PED_S_CARGO%()
   ReDim PED_S_CARGO%(MSF.Rows)
   If CONTROL("ADMIPED", "ITEMSCGO") = "SI" Then
      For i& = 1 To MSF.Rows - 1
        If MSF.TextMatrix(i&, 0) = "*" And XVALO(MSF.TextMatrix(i&, 8)) <= 0.005 Then
           PED_S_CARGO%(i&) = 1
        End If
      Next i&
   End If
   '
   'jandy sql
   'aca actualizo en Peddeta el status
   strSQL = "Update Peddeta set status=100+status "
   strSQL = strSQL & "Where NroPed=" & NROPE&
   Call EJECUTACOMANDO(strSQL)
   '
   ANOTA$ = ""
   TTXYZ$ = Space$(76)
   HOII% = HOY(HOS$)
   Call REPLA(TTXYZ$, HOS$, 1, 8)
   Call REPLA(TTXYZ$, "Modificación de Orden de Venta:", 11, 56)
   Call REPLA(TTXYZ$, USERTER$, 69, 24)
   ANOTA$ = ANOTA$ + TTXYZ$ + Chr$(10) + Chr$(13)
   '
   CAVER$ = TRIM$(Str$(NROPE&))
   While Len(CAVER$) > 6: CAVER$ = Mid$(CAVER$, 2): Wend
   While Len(CAVER$) < 6: CAVER$ = "0" + CAVER$: Wend
   IKITEM% = 0
   For U& = 1 To ULTREG&("!CARDETPE")
     X$ = REGLEIDO$("!CARDETPE", U&)
     CAN = CVS(Mid$(X$, 47, 4))
     If CAN > 0.05 Then
       '
26     IKITEM% = IKITEM% + 1
       CAVERI$ = TRIM$(Mid$(X$, 93, 16))
       If CAVERI$ = "" Then
          IK$ = TRIM$(Str$(IKITEM%))
          CAVERI$ = CAVER$ + "-" + IK$
          'jandy sql
          strSQL = "Select * from PedDeta where NroPed= " & NROPE&
          strSQL = strSQL & " and  TPedItem = ' " & CAVERI$ & "'"
          Set MIRS = FLEXCONN.Execute(FILTSQL$(strSQL))
          If Not (MIRS.EOF And MIRS.BOF) Then GoTo 26
       End If
       '
       ' ESTO ES PARA PONER EL PRECIO DE LISTA EN LA B.DATOS
       CIXI% = CVI(Left$(X$, 2))
'       If CIXI% > 0 Then  'ojo intervencion
'          If Left$(UCase$(ARCHIPRE$), 5) = "LIPRE" Then
'             PLISTA = PRELISTA#(ARCHIPRE$, CIXI%)
'             Call REPLA(x$, MKS$(PLISTA), 39, 4)
'          End If
'       End If
       If CIXI% > 0 Then 'se cambia para que asigne el valor de la lista correspondiente
'       SI VIENE EL VALOR DE LISTA EN 0 LE ASIGNO DE LA LISTA.
          NLISST% = CVS(Mid$(X$, 85, 4))
          PLISTA = PRELISTA#(SAFETEXT$(NLISST%), CIXI%)
          Call REPLA(X$, MKS$(PLISTA), 39, 4)
       End If

       '
        'jandy sql
        strSQL = "SELECT * FROM PEDDETA "
        strSQL = strSQL & " WHERE NroPed = " & NROPE&
        strSQL = strSQL & " and CodiInt = " & CIXI%
        strSQL = strSQL & " and TPedItem = '" & CAVERI$ & "'"
        
        Dim PEDDETTEMP As ADODB.Recordset
        Set PEDDETTEMP = New ADODB.Recordset
27      On Error Resume Next
        PEDDETTEMP.Open FILTSQL$(strSQL), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
        If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
         On Error GoTo 0
         Call CapturaErrorOpen
         GoTo 27
        End If
        On Error GoTo 0

        With PEDDETTEMP
           If (.EOF And .BOF) Then
              .AddNew   ' YA CONVERTIDO
              On Error Resume Next
              !CodiEmpr = CodiEmp%
              On Error GoTo 0
           End If
           !NROPED = NROPE&
           !FECHEMIS = CVDAT(FEMI%)
           !NROCLIE = NCLIEN
           !RaSoClie = rasocli$(NCLIEN)
           !NROOCOM = TRIM$(Text10.TEXT)
             FENTRESOL% = CVI(Mid$(X$, 63, 2))
           !FeSolEnt = CVDAT(FENTRESOL%)
           If XVALO(!Status) > 90 Then
                STATUI% = !Status - 100
                !Status = STATUI%
              Else
                !Status = STATUPE%   'idem anterior.
           End If
           'CONTEMPLA LA POSIBILIDAD DE QUE UNA MODIFICACION EN LA CANTIDAD
           'RESULTE LA NECESIDAD DE CAMBIAR EL STATUS. SOLO SE APLICA A ARTICULOS
           'CON STATUS PREVIO <= 3, NO A ANULADOS O DADOS POR CUMPLIDO
           CAENTAN = CVS(Mid$(X$, 117, 4))
           If STATUI% <= 3 Then
                If (CAN) > (CAENTAN) Then
                     !Status = 0
                ElseIf (CAN) <= (CAENTAN) Then
                     !Status = 3
                End If
           End If
           
           !ListPre = XVALO(TEXT21(0).TEXT)
           !CondPag = XVALO(TEXT21(1).TEXT)
           !VENDED = XVALO(TEXT21(2).TEXT)
           !PorComVend = XVALO(Text44.TEXT)
             If Left$(UCase$(Label21(4)), 5) = "LETRA" Then
               valor% = 0
               If TRIM$(Text44) <> "" Then valor% = Asc(Text44.TEXT + "A") - 64
               !PorComVend = valor%
             End If
           !NuOrItem = IKITEM%
             CIXI% = CVI(Left$(X$, 2))
           !CODIINT = CIXI%
           !CODIEXT = CODEXT$(CIXI%)
           !SubCod = Mid$(X$, 109, 8)
           !Descrip = DESCRIT0$(CIXI%)
           !TCOLOR = Mid$(X$, 3, 16)
           !TTALLE = Mid$(X$, 19, 8)
           !TLargo = CVS(Mid$(X$, 27, 4))
           !TAncho = CVS(Mid$(X$, 31, 4))
           !TEspesor = CVS(Mid$(X$, 35, 4))
           !VenRep = Left$(Label11, 1)   ' Venta o Reparacion
           !MoneEmis = MONEMI%
           !TIPOCAM = XVALO(TICAMBIS)
           !PreList = CVS(Mid$(X$, 39, 4))
           !CanAsig = CVS(Mid$(X$, 43, 4))
           !CanPed = CAN
           !PreUnid = CVS(Mid$(X$, 73, 4))    ' CVS(Mid$(X$, 51, 4))
           !ImpoTot = CVD(Mid$(X$, 55, 8))
             FULENTAN% = CVI(Mid$(X$, 121, 2)) ' FECHA ULTIMA ENTREGA
           !CANTENT = CAENTAN
           !CanSaldo = CAN - CAENTAN
           !FeUltEnt = CVDAT(FULENTAN%)
           !NroRemi = ""
           !TPedItem = CAVERI$
           !SIN_CARGO = PED_S_CARGO%(U&) 'PEDIDO SIN CARGO  MEDITEA
           '
           'GUARDA-LISTA DE PRECIOS
           NLISST% = CVS(Mid$(X$, 85, 4))
           !lista_de_item = NLISST%
           !PreList = CVS(Mid$(X$, 39, 4))

           'INTERVENCION AHP
           LOTEX1$ = TRIM$(Mid$(X$, 3, 16))
           !IDENLOTE = LOTEX1$
            NPROVX = CVS(Mid$(X$, 31, 4)) '(INT15-372)
           !Num_Prov = NPROVX '(INT15-372)
           'FIN AHP

'           'GUARDA-LISTA DE PRECIOS
'           NLISST% = CVS(Mid$(x$, 85, 4))
'           !lista_de_item = NLISST%
           TTXYZ$ = Space$(76)
           SCG$ = "": If PED_S_CARGO%(U&) > 0 Then SCG$ = "(S/C) "
           Call REPLA(TTXYZ$, "- " + CODEXT$(CIXI%), 11, 18)
           Call REPLA(TTXYZ$, SCG$ & DESCRIT0$(CIXI%), 29, 38)
           Call REPLA(TTXYZ$, CSTRING$(MKS$(CAN), 3, 10, 1, 2), 57, 10)
           ANOTA$ = ANOTA$ + TTXYZ$ + Chr$(10) + Chr$(13)
           '
           'INTERVENCION AHP
           If LOTEX1$ <> "" Then
             'GUARDO EN EL ANTOADOR EL NUMERO DE LOTE
             TELOTE$ = "Lote Reservado: " + LOTEX1$
             TELOTEXX$ = Space$(76)
             Call REPLA(TELOTEXX$, TELOTE$, 29, 38)
             ANOTA$ = ANOTA$ + TELOTEXX$ + Chr$(10) + Chr$(13)
           End If
           'FIN AHP
         .Update
       End With
       PEDDETTEMP.Close
       Set PEDDETTEMP = Nothing
       '
     End If
   Next U&
   '
   'QUITA LOS ITEMS QUE PUDIERON SER QUITADOS DEL DETALLE Y CUYO STATUS
   'QUEDO CON LA MARCA DE STATUS > 100
   strSQL = "Delete PedDeta where NroPed= " & NROPE&
   strSQL = strSQL & " and status BETWEEN 99 and 103"
   
   'VUELVE PARA ATRAS EL STATUS DE AQUELLOS ITEMS CUYO ESTADO
   'ERA DADO POR CUMPLIDO (8) O ANULADO (9)
   strSQL = strSQL & "Update PedDeta set status = status-100"
   strSQL = strSQL & " where NroPed=" & NROPE&
   strSQL = strSQL & " and status BETWEEN 108 and 109"
   FLEXCONN.Execute (FILTSQL$(strSQL))
   '
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   If InStr(EPAC$, "ENVAPLAST") > 0 Then
      Call ACTUALIZAOF(NROPE&)
   End If
      '
'   'BORRO REGISTROS DE BALAREQMO
'   CONDI$ = "NUMECOMPRO=" & NROPE&
'   Call BORRAREGISTROS("BALAREQMO", CONDI$, REGAF&, "NOMESS")
'   '
'   'GRABACION EN BALAREQMO (CELSIUS)
'   FORMATPEDI$ = Right$("000000" + TRIM$(Str$(NROPE&)), 6)
'   SQLX$ = "SELECT * FROM PEDDETA"
'   SQLX$ = SQLX$ + " WHERE NROPED = " & NROPE&
'   Set PEDTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
'   With PEDTMP
'   Do While Not .EOF
'     CI% = XVALO(!CodiInt)
'     If CI% > 0 Then
'       CANTI = XVALO(!CanPed)
'       If CANTI > 0.05 Then
'        ORDITEM% = XVALO(!NUORITEM)
'        DESTIN$ = "OV-" & FORMATPEDI$ & "-1"
'        IDENREQ$ = DESTIN$ & "/" & TRIM$(Str$(ORDITEM%))
'        '
'        CONDI$ = "IDENREQUE = '" & IDENREQ$ & "'"
'        '
'        Call CreaRegistro("BALAREQMO", "IDENREQUE", IDENREQ$, "NOMESS")
'        Call ACTUREGISTRO("BALAREQMO", "NUMECOMPRO", CONDI$, NROPE&, REGAF&, "NOMESS")
'        Call ACTUREGISTRO("BALAREQMO", "CODSUCU", CONDI$, 0, REGAF&, "NOMESS")
'        Call ACTUREGISTRO("BALAREQMO", "NUORITEM", CONDI$, ORDITEM%, REGAF&, "NOMESS")
'        Call ACTUREGISTRO("BALAREQMO", "COD_INTE", CONDI$, CI%, REGAF&, "NOMESS")
'        Call ACTUREGISTRO("BALAREQMO", "IDENLOTE", CONDI$, "", REGAF&, "NOMESS")
'        Call ACTUREGISTRO("BALAREQMO", "CANTIDAD", CONDI$, CANTI, REGAF&, "NOMESS")
'        Call ACTUREGISTRO("BALAREQMO", "KILO_REQUE", CONDI$, CANTI, REGAF&, "NOMESS")
'        Call ACTUREGISTRO("BALAREQMO", "DESTINO", CONDI$, DESTIN$, REGAF&, "NOMESS")
'        Call ACTUREGISTRO("BALAREQMO", "CODICOMPRO", CONDI$, "OV", REGAF&, "NOMESS")
'       End If
'       '
'     End If
'     .MoveNext
'
'   Loop
'   End With
'   PEDTMP.Close
'   Set PEDTMP = Nothing
   'VALIDO SI TIENE MOTIVO DE PEDIDO (MEDITEA) AGREGO EN EL ANOTADOR
   NUMOT% = XVALO(Text16) ' COMPARO POR SI CAMBIO EL MOTIVO EN PANTALLA CONTRA EL DE LA BASE
   If NUMOTIVO%(NROPE&) <> NUMOT% Then ' PARA QUE NO ESCRIBA SIEMPRE O EN CUALQUIER EMPRESA
    DESCMOTI$ = ECOARCH$("TAMOTPRE", Chr$(NUMOT%))
    ANOTA$ = ANOTA$ + Space$(10) + String$(56, "-") + Chr$(10) + Chr$(13)
    TTXYZ$ = ""
    Call REPLA(TTXYZ$, "Motivo: " & NUMOT% & " (" & TRIM$(DESCMOTI$) & ")", 11, 66)
    ANOTA$ = ANOTA$ + TTXYZ$ + Chr$(10) + Chr$(13)
   End If
   '
   ANOTA$ = ANOTA$ + Space$(10) + String$(56, "-") + Chr$(10) + Chr$(13)
   Call FRATEXTO(Text13.TEXT, 32)
   For KKI% = 1 To CARETEX%
     TTXX$ = Space$(76)
     Call REPLA(TTXX$, RETEX$(KKI%), 16, 32)
     If KKI% = 1 Then Call REPLA(TTXX$, "Observaciones:", 1, 15)
     ANOTA$ = ANOTA$ + Space$(10) + TTXX$ + Chr$(10) + Chr$(13)
   Next KKI%
   '
   If InStr(UCase$(FUERZAPROBACION$), "CAMBIO_PAGO") > 0 Then
      CONDIP% = CPAGCLI%(NC)
      CONDIPEDIDO% = XVALO(TEXT21(1))
      TEXTOPAGO$ = "Se Cambió Condición de Pago " & CONDIP% & "  Por la " & CONDIPEDIDO%
      Call REPLA(TXX1$, TEXTOPAGO$, 11, 56): ANOTA$ = ANOTA$ + TXX1$ + Chr$(10) + Chr$(13)
   End If
   '
   If InStr(UCase$(FUERZAPROBACION$), "CAMBIO_LISTA") > 0 Then
      LISTAX% = Str$(LIPRCLI%(NCLIEN))
      LISTAXPEDI% = XVALO(TEXT21(0))
      TEXTOLISTAP$ = "Se Cambió Lista de Precio " & LISTAX% & "  Por la " & LISTAXPEDI%
      Call REPLA(TXX1$, TEXTOLISTAP$, 11, 56): ANOTA$ = ANOTA$ + TXX1$ + Chr$(10) + Chr$(13)
   End If

   If InStr(UCase$(FUERZAPROBACION$), "FACTURA_VENCIDA") > 0 Then Call REPLA(TXX1$, "Se Detectaron Facturas Vencidas", 11, 56): ANOTA$ = ANOTA$ + TXX1$ + Chr$(10) + Chr$(13)
   If InStr(UCase$(FUERZAPROBACION$), "FECHA ULTIMA FACTURA:") > 0 Then Call REPLA(TXX1$, "Supera Fecha Máxima de Última Compra", 11, 56): ANOTA$ = ANOTA$ + TXX1$ + Chr$(10) + Chr$(13)
   If UCase(FUERZAPROBACION$) = "PENDIENTE DE APROBACION POR MATERIALES SIN CARGO" Then Call REPLA(TXX1$, "Items Con Importes Que Difieren de Política S/Cargo", 11, 56): ANOTA$ = ANOTA$ + TXX1$ + Chr$(10) + Chr$(13)
   '
   'ANOTADOR DE DETALLES DE VALIDACION DE DERECHOS SOBRE LISTA DE PRECIOS
   If OKLISTAPRE% < 1 Then
      ANOTA$ = ANOTA$ + ACUTXTLPRE$
   End If
   
   'Anota Cambio de Cliente
   If XVALO(LTEXT1) <> XVALO(Text1) Then
       CAMBCLIE$ = "Cambio de Cliente " & TRIM$(Text1) & " Por " & TRIM$(LTEXT1)
       Call REPLA(TEXCAMBICLIE$, CAMBCLIE$, 11, 56): ANOTA$ = ANOTA$ + TEXCAMBICLIE$ + Chr$(10) + Chr$(13)
       LTEXT1 = Text1
   End If
'   ANOTA$ = ANOTA$ + String$(88, "=") + Chr$(10) + Chr$(13)

   SQLX$ = "SELECT ANOTAPED FROM PEDENCA"
   SQLX$ = SQLX$ + " WHERE NroPed = " & NROPE& & ""
   Dim PEDDDENCA As New ADODB.Recordset
   Set PEDDDENCA = New ADODB.Recordset
28 On Error Resume Next
   PEDDDENCA.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 28
   End If
   On Error GoTo 0
   With PEDDDENCA
      If Not (.BOF And .EOF) Then
         ANOTANTE$ = SAFETEXT$(!AnotaPed)
         ANOTA$ = ANOTANTE$ + ANOTA$
         ANOTA$ = ANOTA$ + String$(88, "=") + Chr$(10) + Chr$(13)
         !AnotaPed = ANOTA$
         .Update
       End If
   End With
   PEDDDENCA.Close
   Set PEDDDENCA = Nothing
   '
   Call ABRECLIEN
   '
   'VALIDACION DE SETUP SI SE GUARDA LOS DATOS DE ENTREGA (ACTUALIZANDO LOS DATOS DEL CLIENTE)
   If CONTROL("PEDIDO", "ACTULEN") <> "NO" Then
        strSQL = "Update Deudor12 set Lentre0_Cli ='" & TRIM$(LENTREGA(0).TEXT) & "'"
        strSQL = strSQL & ", Lentre1_Cli ='" & TRIM$(LENTREGA(1).TEXT) & "'"
        strSQL = strSQL & ", Lentre2_Cli ='" & TRIM$(LENTREGA(2).TEXT) & "'"
        strSQL = strSQL & " where Nume_Cli = " & NCLIEN
        FLEXCONN.Execute (FILTSQL$(strSQL))
   End If
   '
   Call CIERRARCH("*.*")
   Call LIBARCH("*.*")
   Call BAJALDISCO
   '
   Call CIERRARCH("PEDDETA")
   Call CIERRARCH("PEDDETP")
   Call CIERRARCH("DESPAPED")
   Call CIERRARCH("*.*")
   Call LIBARCH("*.*")
   Call BAJALDISCO
   '
'   Call BLOQARCH("INDIPEDI")
'   TTXX$ = MKS$(NROPE&) + REFE$
'   Call FILTERPUTRECORD("INDIPEDI", 1, MKS$(NROPE&), TTXX$)
'   Call FILESORT("INDIPEDI", "", 1, 1, "DESC")
'   Call CIERRARCH("INDIPEDI")
   '
   TPED# = XVALO(Total(3).TEXT)
   OKX% = 1
   If VERLICRE% = 1 Then
     LMCR# = LIMICRE(NCLIEN)
     SDEU# = SALDCLI#(NCLIEN)
     CHPA# = CHEPEACRE#(NCLIEN)
     If SDEU# + CHPA# + TPED# > LMCR# Then
       TLIMI$ = TRIM$(CSTRING$(MKD$(LMCR#), 3, 12, 2, 2))
       If TLIMI$ = "" Then TLIMI$ = "0.00"
       ATTX$ = "ATENCION: Este Pedido Supera\el Límite de Crédito Asignado\para esta Cuenta de $ " + TLIMI$
       If DERACCE%("AUTORENT/NOMESS", "Autorizar Pedidos con Exceso Credito") < 2 Then
           Call COMUNI(ATTX$ + ".\  \Informe al Supervisor de Ventas\para Autorizar Facturación y Despacho.")
         Else
           AUTORENT$ = "NO"
           If COMALTER%(ATTX$ + "\   \Puede Autorizar la Entrega\con Crédito Excedido.\\Continuar\Autorizar Entrega") = 2 Then AUTORENT$ = "SI"
           RFF$ = RECFILT$("AUTORENT", 1, MKS$(NROPE&))
           REGAUT& = 0: If Len(RFF$) >= 4 Then REGAUT& = CVS(Left$(RFF$, 4))
           AUTORI$ = MKS$(NROPE&) + AJUSTI$(USERNAME$, 26) + AUTORENT$
           If REGAUT& > 0 Then
               Call GRAREG("AUTORENT", AUTORI$, REGAUT&)
             Else
               Call REGAPP("AUTORENT", AUTORI$)
           End If
           Call CIERRARCH("AUTORENT")
       End If
     End If
   End If
   '
   If EXEDELIMCC%(NCLIEN, TPED#, DIFER) > 0 Then
       LIMCCTE$ = TRIM$(FORMATNUM$(LCred_CCte(NCLIEN), "f12.2")): If LIMCCTE$ = "" Then LIMCCTE$ = "0.00"
       DifADep$ = TRIM$(FORMATNUM$(DIFER, "f12.2")): If DifADep$ = "" Then DifADep$ = "0.00"
       TOTAPE$ = TRIM$(FORMATNUM$(TPED#, "f12.2")): If TOTAPE$ = "" Then TOTAPE$ = "0.00"
       TX$ = "ATENCION: Este Pedido Supera\el Límite de Crédito Asignado en C.Cte.\para esta Cuenta de $ " + LIMCCTE$
       TX$ = TX$ + "\Puede Ahora Enviar un Mail Solicitando la Cancelación\ del Excedido Incluyendo el Presente Pedido \Por La Diferencia ($ " + DifADep$ + ") \\Enviar Mail\Continuar"
110    If COMALTER%(TX$) <> 2 Then
            MAI_DESTI$ = EMAILCLI$(NCLIEN)
            If TRIM$(MAI_DESTI$) = "" Or InStr(MAI_DESTI$, "@") < 1 Then
120           MAI_DESTI$ = LCase$(TRIM$(InputBox$(" Mail de Cliente", "Intgrese Mail de Destino", MAI_DESTI$)))
              If TRIM$(MAI_DESTI$) = "" Then
                GoTo 110
              ElseIf InStr(MAI_DESTI$, "@") < 1 Then
                  Call COMUNI("Mail Incorrecto")
                  GoTo 120
              End If
            End If
            TXT$ = "Solicitamos realizar el Depósito de $ " & DifADep$
            TXT$ = TXT$ + " Para cubrir la diferencia de su limite de cuenta corriente "
            TXT$ = TXT$ + "de $ " & LIMCCTE$ & " y poder autorizar la entrega del pedido Nro.: " & NROPE&
            TXT$ = TXT$ + " y por  importe de $ " & TOTAPE$
            MAI_DIREL$ = MAI_DESTI$
            MAI_ASUNT$ = "Solicitud de Interdepósito para Pedido Número: " & NROPE&
            MAI_TEXTO$ = TXT$
            '
            Call GENERAMAIL
       End If

   End If
   '
   Screen.MousePointer = 1
   Exit Sub
   '
98 Call CIERRARCH("PEDENCA")
   Call CIERRARCH("PEDDETA")
   Call CIERRARCH("PEDDETP")
   Screen.MousePointer = 1
   Call MENSERR(24, MERRO$)
   '
End Sub
Private Sub Total_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 13 Then
     If Index = 0 Then
       On Error Resume Next
       PORDESCU.SetFocus
       On Error GoTo 0
     End If
   End If
End Sub

Private Sub Total_LostFocus(Index As Integer)
   If Index = 0 Then
     TTT1 = XVALO(Total(0).TEXT)
     Total(0) = CSTRING$(MKS$(TTT1), 3, 11, 2, 2)
   End If
End Sub

Sub CARDETPEDI(MODO%)
   '
   HOII% = HOY(HO$)
   
   NUPEDIDO1& = XVALO(Text11)
   Call ABREPEDCLI
   LIPRACTI% = XVALO(Mid$(ARCHIPRE$, 6, 3))
   ATRI$ = "/NC/ARCHECO(1)=" + ARCHIPRE$
   ATRI$ = ATRI$ + "/TITUPAN=Detalle Pedido de Cliente"
   ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + Frame1.Top + Picture3.Top))
   ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + Frame1.Left + Picture3.Left + 50))
   
   If REFRESCARPED% > 0 Then
     'VARIABLE A NIVEL DE FRM QUE PERMITE REFRESCAR EL PEDIDO SIN MODIFICAR (PARA RECUPERAR IMAGENES DE PEDIDOS ANULADOS O CUMPLIDOS)
     ATRI$ = ATRI$ + "/POSEDIT(1)=2"
     ATRI$ = ATRI$ + "/POSEDIT(3)=2"
     ATRI$ = ATRI$ + "/POSEDIT(4)=2"
     ATRI$ = ATRI$ + "/POSEDIT(5)=2"
     ATRI$ = ATRI$ + "/POSEDIT(6)=2"
     ATRI$ = ATRI$ + "/POINDEL=NO" ' DESHABILITA EL INSERTAR Y EL DEL
     REFRESCARPED% = 0
     GoTo 5
   End If
     
   PEDSER = XVALO(VALOBADA("PEDDETA", "count(NUMEROSERIE)", "(NUMEROSERIE <> '' or (NUMEROSERIE is not null and NUMEROSERIE <> ''))  and NROPED = " & NUPEDIDO1&))
   If PEDSER > 0 Then ' SI TIENE UN NUMERO DE SERIE ASIGNADO PONE NO EDITABLE EL CAMPO 1 Y 4
     ATRI$ = ATRI$ + "/POSEDIT(1)=2"
     ATRI$ = ATRI$ + "/POSEDIT(4)=2"
     ATRI$ = ATRI$ + "/POINDEL=NO" ' DESHABILITA EL INSERTAR Y EL DEL
   End If
   '
   If InStr(1, EMPREAC$, "TER") > 0 Then
        ATRI$ = ATRI$ + "/POSEDIT(1)=2"
        ATRI$ = ATRI$ + "/POSEDIT(3)=2"
        ATRI$ = ATRI$ + "/POSEDIT(5)=2"
        ATRI$ = ATRI$ + "/POSEDIT(6)=2"
        ATRI$ = ATRI$ + "/POINDEL=NO" ' DESHABILITA EL INSERTAR Y EL DEL
   End If
   '
5  Timer1.Enabled = True
   CONLOTIND$ = UCase(TRIM$(CONTROL$("ADMIPED", "VALOIND")))      'CONTROL DE VALIDACION DE VERIFICACION DE LOTE INDIVISIBLE
10 VENCARPE$ = CONTROL$("", "VENCARPE")
   
   'AHP
   If UCase$(TRIM$(VENCARPE$)) = "CARDETLO" Then
     ATRI$ = ATRI$ + "/POSEDIT(1)=2" ' VENTANA CON LOTES AHP
   End If
   '
   If TRIM$(VENCARPE$) = "" Then VENCARPE$ = "CARDETP"
   '
   ' PARA EL CASO QUE SE QUIERA MODIFICAR UN PEDIDO CON NUM.SERIE ASIGNADOS
   '
   PRF$ = VENCARPE$    ' If FORMAPED% = 2 Then PRF$ = "CARDETQ"
   XXX% = VENTABU%(PRF$ + ATRI$)
   '
   '
   Call CALTOPED
   If ECOPRECIO% >= 0 Then
     If XVALO(Total(2).TEXT) < 0.005 Then
       If TRIM$(Text13.TEXT) = "" Then
         'Text13.TEXT = "SIN CARGO !!!"
       End If
     End If
   End If
   Timer1.Enabled = False
   '
   If XXX% < 0 Then
      Call CIERRARCH("!CARDETPE")
      Call BLANARCH("!CARDETPE")
      Call BLANFORMU
      Text1.TEXT = ""
      Call CIERRARCH("*.*")
      Hide
      Exit Sub
   End If
   '
20 If ULTREG&("!CARDETPE") < 1 Then
      Call BLANFORMU
      Text1.TEXT = ""
      Call CIERRARCH("*.*")
      Hide
      Exit Sub
   End If
   '
   Screen.MousePointer = 11   ' Registrando Pedido ...
   List1.Clear
   FIN& = ULTREG&("!CARDETPE")
   If FIN& > 0 Then
     Y$ = REGLEIDO$("!CARDETPE", FIN&)
     CI% = CVI(Left$(Y$, 2))
     If CI% < 1 Or CI% > NUMAS% Then
       Call SETULTREG("!CARDETPE", FIN& - 1)
       Call CIERRARCH("!CARDETPE")
       GoTo 20
     End If
   End If
   '
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   CODSINSTO$ = ""
   VALIREPE% = 1 ' BANDEREA DE VALIDACION DE PRECIOS
   If CONTROL("ADMIPED", "VALIREPE") = "NO" Then VALIREPE% = 0

   VALICODIENCLIENTE% = 0 'BANDERA PARA VALIDAR SI EL CODIGO FUE PEDIDO POR EL CLIENTE
   If CONTROL("ADMIPED", "VACOPED") = "SI" Then VALICODIENCLIENTE% = 1
   
   For U& = 1 To FIN&
     Y$ = REGLEIDO$("!CARDETPE", U&)
     CI% = CVI(Left$(Y$, 2))
     If CI% < 1 Or CI% > NUMAS% Then
       Call MENSERR(24, "Código no Válido")
       Call REPLA(Y$, MKI$(0), 1, 2)
       Call GRAREG("!CARDETPE", Y$, U&)
       GoTo 10
     End If
     '
     FENTRESOL% = CVI(Mid$(Y$, 63, 2))
     If FENTRESOL% < 1 Then
        FENTRESOL% = UFENTRE%
        If FENTRESOL% < 1 Then FENTRESOL% = HOII%
        Call REPLA(Y$, MKI$(FENTRESOL%), 63, 2)
        Call GRAREG("!CARDETPE", Y$, U&)
     End If
     UFENTRE% = FENTRESOL%
     '
     '  verifica codigos repetidos
      If VALIREPE% > 0 Then
       For W& = 1 To U& - 1
         Y$ = REGLEIDO$("!CARDETPE", W&)
         If CVI(Left$(Y$, 2)) = CI% Then
           If CVI(Mid$(Y$, 63, 2)) = FENTRESOL% Then
             Screen.MousePointer = 1
             TTXX$ = "Código " + TRIM$(CODEXT$(CI%)) + "\Duplicado en el Pedido."
             If COMALTER%(TTXX$ + "\\Corregir\Continuar") <> 2 Then
               GoTo 10
             End If
           End If
         End If
       Next W&
     End If
     '
     If VALICODIENCLIENTE% > 0 Then
        NCLIE = XVALO(Text1)
        CONDI$ = "NROCLIE = " & NCLIE & " AND CODIINT = " & CI%
        CODPEDIDO& = CantRegistros&("PEDDETA", CONDI$, "NOMESS")
        If CODPEDIDO& < 1 Then
           TTXX$ = "Código " + TRIM$(CODEXT$(CI%)) + "\" & DESCRIT0$(CI%) & "\"
           TTXX$ = TTXX$ + "Nunca Fue Solicitado Por Este Cliente.\ "
           TTXX$ = TTXX$ + "¿Desea Continuar?"
           If COMALTER%(TTXX$ + "\\Sí, Continuar\No, Cancelar") > 1 Then
              GoTo 10
           End If
        End If
     End If

     X$ = REGLEIDO$("!CARDETPE", U&)
     CI% = CVI(Left$(X$, 2))
     TTXX$ = Space$(96)
     FENTRESOL% = CVI(Mid$(X$, 63, 2))
'     If FENTRESOL% < HOII% Then
'        FENTRESOL% = HOII%
'        Call REPLA(X$, MKI$(FENTRESOL%), 63, 2)
'        Call GRAREG("!CARDETPE", X$, U&)
'     End If
     '
     CAN = CVS(Mid$(X$, 47, 4))
     If CAN > 999999 Then
       Call MENSERR(24, "Cantidad Excesiva - Max 999999")
       Call REPLA(X$, MKS$(0), 47, 4)
       Call GRAREG("!CARDETPE", X$, U&)
       GoTo 10
     End If
'     If CAN < 0.5 Then
'       Call MENSERR(24, "Falta Cantidad")
'       Call REPLA(X$, MKS$(0), 47, 4)
'       Call GRAREG("!CARDETPE", X$, U&)
'       GoTo 10
'     End If
     '
     If CONLOTIND$ = "SI" Then      'SEGUN LA CONFIGRACION REALIZA EL CONTROL DEL LOTE INDIVISIBLE
        If Abs(CAN - REDON_INDIVIS(CI%, CDbl(CAN), "VENTA")) >= 0.05 Then
           TTXX$ = "Cantidad Ajustada a Lotes Indivisibles\en Código '" + CODEXT$(CI%) + "'"
           Call MENSERR(24, TTXX$)
           Call REPLA(X$, MKS$(REDON_INDIVIS(CI%, CDbl(CAN), "VENTA")), 47, 4)
           Call GRAREG("!CARDETPE", X$, U&)
           GoTo 10
        End If
     End If
     '
     PORDESCUIN = CVS(Mid$(X$, 69, 4))
     If PORDESCUIN < 0 Or PORDESCUIN > 99.9 Then
       Call MENSERR(24, "Descuento no Válido")
       Call REPLA(X$, MKS$(0), 69, 4)
       Call GRAREG("!CARDETPE", X$, U&)
       GoTo 10
     End If
     '
     If ECOPRECIO% < 0 Then
       PRECARGA = PRELISTA#(ARCHIPRE$, CI%)
       Call REPLA(X$, MKS$(PRECARGA), 39, 4)
       Call GRAREG("!CARDETPE", X$, U&)
     End If
     PRECARGA = CVS(Mid$(X$, 39, 4))
     PREUNI = PRECARGA * (100 - PORDESCUIN) / 100
     TOT# = CDbl(CAN * PREUNI)
     '
     CADECI% = DECILISTA%(XVALO(Mid$(ARCHIPRE$, 6, 3)))
     VAMINI1 = 0.5 / (10 ^ CADECI%)
     If PRECARGA < VAMINI1 Or PRECARGA > 9999999999.99 Then
       If MODO% < 2 Then
         If ECOPRECIO% >= 0 Then
           Call REPLA(X$, MKS$(0), 39, 4)
           Call GRAREG("!CARDETPE", X$, U&)
           MODO% = 2
           If COMALTER%("Artículo '" + TRIM$(CODEXT$(CI%)) + "' sin Precio.\\Corregir\Continuar") <> 2 Then
              GoTo 10
           End If
         End If
       End If
     End If
     '
     CANTPEDIDA# = CVS(Mid$(X$, 47, 4))
     If CONTROL$("PEDIDO", "VERISTO") = "SI" Then
        If DESCOMPO%(CI%) > 0 Then
          Call AREXPLO(CI%, 1, 1)
          For KKJ% = 1 To CAIT%
             CICO% = CIJ%(KKJ%)
             CACICO = USOI(KKJ%) * CANTPEDIDA#
             If CACICO > STOCKACT(CICO%) Then
                If CODSINSTO$ <> "" Then
                    CODSINSTO$ = CODSINSTO$ & " ; " & CODEXT(CICO%)
                  Else
                    CODSINSTO$ = CODEXT(CICO%)
                End If
             End If
          Next KKJ%
          GoTo 30
        End If
        '
        If CANTPEDIDA# > STOCKACT(CI%) Then
          If CODSINSTO$ <> "" Then
              CODSINSTO$ = CODSINSTO$ & " ; " & CODEXT(CI%)
            Else
              CODSINSTO$ = CODEXT(CI%)
          End If
        End If
     End If
     '
30   Call REPLA(X$, MKS$(PRECARGA), 39, 4)
     Call REPLA(X$, MKS$(PRECARGA), 51, 4)
     Call REPLA(X$, MKD$(TOT#), 55, 8)
     Call REPLA(X$, MKS$(PRECARGA), 65, 4)
     Call REPLA(X$, MKS$(PREUNI), 73, 4)
     Call REPLA(X$, MKD$(TOT#), 77, 8)
     '
     'GUARDA LA LISTA DE PRECIOS, SI YA TIENE MANTIENE SI NO ESCRIBE
     If Abs(CVS(Mid$(X$, 85, 4))) = 0 Then
        LPRE = XVALO(Mid$(ARCHIPRE$, 6, 3))
        If LPRE = 0 Then LPRE = -1
     Else
        LPRE = CVS(Mid$(X$, 85, 4))
     End If
     Call REPLA(X$, MKS$(LPRE), 85, 4)

     Call GRAREG("!CARDETPE", X$, U&)
     '
   Next U&
   '
   If CODSINSTO$ <> "" Then
      OPX% = COMALTER("Los Siguientes Codigos Tienen Stock Insuficiente \" & CODSINSTO$ & "\\Corregir\Continuar")
      If OPX% <> 2 Then GoTo 10
   End If
   '
   'VERIFICA QUE NO SE HAYA QUITADO UN ARTICULO QUE YA FUE ENTREGADO. SE APLICA SOLO PARA FERVI
   If InStr(1, EMPREAC$, "FERVI") > 0 Then
        SQLX$ = "SELECT CODIINT, SUM(CANTENT) AS QENT FROM PEDDETA WITH(NOLOCK) WHERE NROPED=" & NUPEDIDO1& & " GROUP BY CODIINT"
        Set RST = READSET(SQLX$)
        Do While Not RST.EOF
             CIXI% = XVALO(RST!CODIINT)
             QENT# = XVALO(RST!QENT)
             Encontrado% = 0
             If QENT# > 0.05 Then
                 For II& = 1 To ULTREG&("!CARDETPE")
                     Y$ = REGLEIDO$("!CARDETPE", II&)
                     CI% = CVI(Left$(Y$, 2))
                     If CIXI% = CI% Then
                         Encontrado% = 1
                         Exit For
                     End If
                 Next II&
                 If Encontrado% = 0 Then
                     TTXX$ = "Imposible Continuar."
                     TTXX$ = TTXX$ & "\El Código '" & CODEXT$(CIXI%) & " Ha Sido Eliminado"
                     TTXX$ = TTXX$ & "\y Posee Entregas por " & TRIM$(FORMATNUM$(QENT#, "F10.1") & " " & Unimed$(CIXI%))
                     Call COMUNI(TTXX$)
                     GoTo 10
                 End If
             End If
             RST.MoveNext
        Loop
        On Error Resume Next
        RST.Close
        Set RST = Nothing
        On Error GoTo 0
   End If
   '
   
   Call CARLISPE
   Call CIERRARCH("!CARDETPE")
   Call CALTOPED
   '
   If CONTROL("ADMIPED", "ITEMSCGO") = "SI" Then
      For i& = 1 To MSF.Rows - 1
        If XVALO(MSF.TextMatrix(i&, 8)) <= 0.01 Then
           'Call COLOREAR_GRILLA_SOLAFILA(MSF, vbYellow, GLOBALFRM_FILA&, , , 1)
           Call COLOREAR_GRILLA_SOLO_UNA_AFILA(MSF, vbYellow, i&)
           MSF.TextMatrix(i&, 0) = "*"
        Else
           Call COLOREAR_GRILLA_SOLO_UNA_AFILA(MSF, vbWhite, i&)
           MSF.TextMatrix(i&, 0) = ""
        End If
      Next i&
   End If

   'VALIDACION DE CONTROL DE MARCA DE MODO DE DESCUENTO
   'CAMBIA EL DESCUENTO DE ACUERDO A LA TABAL DESMONTO
   If CONTROL("PEDIDO", "MODODESC") = "MONTO" And XVALO(TEXT21(0)) = 1 Then
       Call FORPED07.DESCUMONTO(Me)
   End If
  
   '
   On Error Resume Next
   Command14.SetFocus
   On Error GoTo 0
   Screen.MousePointer = 1
   '
End Sub

Sub CARLISPE()
   List1.Clear
   FIN& = ULTREG&("!CARDETPE")

   For U& = 1 To FIN&
     X$ = REGLEIDO$("!CARDETPE", U&)
     TTXX$ = Space$(96)
     CI% = CVI(Left$(X$, 2))
     LOTE$ = Mid$(X$, 3, 16)
     FENTRESOL% = CVI(Mid$(X$, 63, 2))
     CAN = CVS(Mid$(X$, 47, 4))
     PORDESCUIN = CVS(Mid$(X$, 69, 4))
     PRECARGA = CVS(Mid$(X$, 39, 4))
       DLX% = DECILISTA(XVALO(Mid$(ARCHIPRE$, 6, 3)))    ' DECIMALES DE LA LISTA DE PRECIOS
     If Abs(PORDESCUIN) > 0 Then
       PREUNI = ROUND((PRECARGA * (100 - PORDESCUIN) / 100), DLX%)
       PREUNI_SINRED = PRECARGA * (100 - PORDESCUIN) / 100
     Else
       PREUNI = ROUND(PRECARGA, DLX%)
       PREUNI_SINRED = PRECARGA
     End If
     TOT# = ROUND(CDbl(CAN * PREUNI), 2)
     '
     If CAN <= 9999 Then
       Call REPLA(TTXX$, CSTRING$(MKS$(CAN), 3, 6, 1, 2), 1, 6)
     Else
       Call REPLA(TTXX$, CSTRING$(MKS$(CAN), 3, 6, 0, 2), 1, 6)
     End If
     '
     Call REPLA(TTXX$, CODEXT$(CI%), 8, 15)
     FEX = FENTRESOL%
     If ECOPRECIO% < 0 Then
          Call REPLA(TTXX$, DESCRIT0$(CI%), 24, 52)
          Call REPLA(TTXX$, FECHATEX$(FEX), 78, 8)
        Else
          Call REPLA(TTXX$, DESCRIT0$(CI%), 24, 22)
          Call REPLA(TTXX$, Left$(FECHATEX$(FEX), 5), 47, 5)
     End If
     PRELIX = PRELISTA#(ARCHIPRE$, CI%)
     If PRELIX < PRECARGA Then PRELIX = PRECARGA
     If PRELIX < PREUNI Then PRELIX = PREUNI
     PORDESCUIN = 0: If PRELIX > 0.005 Then PORDESCUIN = 100 * (PRELIX - PREUNI_SINRED) / PRELIX
     ' ESTO SE MODIFICO POR PROBLEMA CON CARGA DE PEDIDOS DE GABAL
     ' PORDESCUIN = 0: If PRELIX > 0.005 Then PORDESCUIN = 100 * (PRELIX - PREUNI) / PRELIX
       DLX% = DECILISTA(XVALO(Mid$(ARCHIPRE$, 6, 3)))    ' DECIMALES DE LA LISTA DE PRECIOS
     If ECOPRECIO% >= 0 Then Call REPLA(TTXX$, CSTRING$(MKS$(PRELIX), 3, 8, DLX%, 2), 52, 8)
     If ECOPRECIO% >= 0 Then Call REPLA(TTXX$, CSTRING$(MKS$(PORDESCUIN), 3, 5, 1, 2), 61, 5)
     If ECOPRECIO% >= 0 Then Call REPLA(TTXX$, CSTRING$(MKS$(PREUNI), 3, 8, DLX%, 2), 66, 8)
     If ECOPRECIO% >= 0 Then Call REPLA(TTXX$, CSTRING$(MKD$(TOT#), 3, 10, 2, 2), 76, 10)
     'List1.AddItem TTXX$
     MSF.Rows = U& + 1
     MSF.TextMatrix(U&, 1) = TRIM$(Mid$(TTXX$, 1, 6))
     MSF.TextMatrix(U&, 2) = TRIM$(Mid$(TTXX$, 8, 15))
     MSF.TextMatrix(U&, 3) = TRIM$(Mid$(TTXX$, 24, 22))
     MSF.TextMatrix(U&, 4) = TRIM$(Mid$(TTXX$, 47, 5))
     MSF.TextMatrix(U&, 5) = TRIM$(FORMATNUM$(PRELIX, "F12." & SAFETEXT$(DLX%)))
     MSF.TextMatrix(U&, 6) = TRIM$(FORMATNUM$(PORDESCUIN, "F7.2"))
     MSF.TextMatrix(U&, 7) = TRIM$(FORMATNUM$(PREUNI, "F12." & SAFETEXT$(DLX%)))
     MSF.TextMatrix(U&, 8) = TRIM$(FORMATNUM$(TOT#, "F14.2"))
     MSF.TextMatrix(U&, 9) = TRIM$(LOTE$)
     '
   Next U&

End Sub
'
Sub CALTOPED()
   '
   Static CLIA$, RNI%, APLIVA%
   If Text1.TEXT <> CLIA$ Then
     CLIA$ = Text1.TEXT: NCLIEN = XVALO(CLIA$)
     RNI% = 0: APLIVA% = 1
     If NCLIEN > 0 Then
       PVCLI% = PIVACLI%(NCLIEN)
       If PVCLI% = 2 Then RNI% = 1
       If PVCLI% = 5 Or PVCLI% = 7 Then APLIVA% = 0
     End If
   End If
   '
   SUTOTPED# = 0: SUIVA# = 0: SUIVARNI# = 0: TOIVA# = 0: TOTOTA# = 0
   '\\\OT-05-0560-WAR-24/11/05///
   Dim IMPTASIVA#()
   ReDim IMPTASIVA#(ULTREG&("GRUPIVA"))
   DESC0$ = TRIM$(PORDESCU.TEXT)
   'SI TIENE DESCUENTO DE RETIRA CLIENTE LO CONCATENA PARA EL CALCULO IMPOSITIVO
   If XVALO(Label31) > 0 Then DESC0$ = TRIM$(Str$(XVALO(Label31))) & "+" & DESC0$
   If XVALO(Text18) > 0 Then DESC0$ = TRIM$(Str$(XVALO(Text18))) & "+" & DESC0$

    
   COEF1# = 1 - COEFDESCU(DESC0$) / 100
   '\\\OT-05-0560-WAR-FIN///
   
   'MODIFICACION DE EYM ITEM CON 2 IVAS (CARGA EL VECTOR CON LOS CODIGOS)
   For i& = 1 To 256
       CI_VECTOR_TIVA(i&) = 0
   Next i&

   For U& = 1 To ULTREG&("!CARDETPE")
     Y$ = REGLEIDO$("!CARDETPE", U&)
     CI_VECTOR_TIVA(U&) = CVI(Left$(Y$, 2))
   Next U&
   Call ACTUALIZAVECTORDEIVA
   
   'FIN MODIFICACION DE EYM ITEM CON 2 IVAS
   TOTITEMS% = 0

   For U& = 1 To ULTREG&("!CARDETPE")
     X$ = REGLEIDO$("!CARDETPE", U&)
     CIXXI% = CVI(Left$(X$, 2))
     CAN = CVS(Mid$(X$, 47, 4))
     PORDESCUIN = CVS(Mid$(X$, 69, 4))
     If PORDESCUIN < 0 Or PORDESCUIN > 99 Then PORDESCUIN = 0
     If CAN > 0.05 Then
       PRELI = CVS(Mid$(X$, 39, 4))
       'If PRELI < 0.01 Then
       '  PRELI = PRELISTA(ARCHIPRE$, CIXXI%)
       'End If
       CDX% = DECILISTA(LIPRACTI%)
       If Abs(PORDESCUIN) > 0 Then
          PREUNI = ROUND(PRELI * (100 - PORDESCUIN) / 100, CDX%)
       Else
          PREUNI = ROUND(PRELI, CDX%)
       End If
       TOT# = ROUND(CDbl(CAN * PREUNI), 2)
       SUTOTPED# = SUTOTPED# + TOT#
       SUIVA# = SUIVA# + (Int(TOT# * PORCEIVA(CIXXI%) + 0.5)) / 100
       '\\\OT-05-0560-WAR-03/08/05///
       TODE1# = TOT# * COEF1#
       GRIVA% = GRUPIVA%(CIXXI%)
       IMPTASIVA#(GRIVA%) = IMPTASIVA#(GRIVA%) + (Int(TODE1# * PORCEIVA(CIXXI%) + 0.5)) / 100
       Label13 = CSTRING$(MKD$(IMPTASIVA#(0)), 3, 13, 2, 2)
       Label14 = CSTRING$(MKD$(IMPTASIVA#(1)), 3, 13, 2, 2)
       Label20 = CSTRING$(MKD$(IMPTASIVA#(2)), 3, 13, 2, 2)
       '\\\OT-05-0560-WAR-FIN///
       Call REPLA(X$, MKS$(PREUNI) + MKD$(TOT#), 51, 12)
       Call GRAREG("!CARDETPE", X$, U&)
       TOTITEMS% = TOTITEMS% + 1
     End If
   Next U&
   '
   Label18 = TRIM$(Str$(TOTITEMS%))
   '
   Total(0).TEXT = CSTRING$(MKD$(SUTOTPED#), 3, 13, 2, 2)
   Total(0).Refresh
   '
   'CALCULO DE RECARGO POR CONDICION DE PAGO
   SIGNO$ = ""
   RECPORPAGO = 0
   RECACON = COEFDESCU(Label25)
   If Abs(XVALO(RECACON)) > 0 Then RECPORPAGO = XVALO(RECACON) * (-1)
   If XVALO(PORDESCU.TEXT) > 0 Then SIGNO$ = "+"
   DESC$ = RECPORPAGO & SIGNO$ & TRIM$(PORDESCU.TEXT)
   '
'   DESC$ = TRIM$(PORDESCU.TEXT)
   TTT1 = XVALO(Total(0).TEXT) * COEFDESCU(DESC$) / 100
   Total(1).TEXT = CSTRING$(MKS$(TTT1), 3, 13, 2, 2)
   '
   'AGREGADO PARA GABAL DESDE ACA
   DESC1 = (XVALO(Total(0).TEXT) - XVALO(Total(1).TEXT)) * XVALO(Label31) / 100
   Text141 = CSTRING$(MKS$(DESC1), 3, 13, 2, 2)
   DESC2 = (XVALO(Total(0).TEXT) - XVALO(Total(1).TEXT)) * XVALO(Text18) / 100
   Text17 = CSTRING$(MKS$(DESC2), 3, 13, 2, 2)

   TTT3 = XVALO(Total(0).TEXT) - XVALO(Total(1).TEXT) - (XVALO(Text141) + XVALO(Text17)) ' ESTA ULTIMA RESTA GABAL
   Total(2).TEXT = CSTRING$(MKS$(TTT3), 3, 13, 2, 2)
   
'   TTT3 = XVALO(Total(0).TEXT) - XVALO(Total(1).TEXT)
'   Total(2).TEXT = CSTRING$(MKS$(TTT3), 3, 13, 2, 2)
   '
   On Error Resume Next
   TOTIVA(0).TEXT = CSTRING$(MKD$(APLIVA% * SUIVA# * XVALO(Total(2).TEXT) / XVALO(Total(0).TEXT)), 4, 13, 2, 2)
   On Error GoTo 0
   TOTIVA(2).TEXT = CSTRING$(MKD$(CDbl(XVALO(TOTIVA(0).TEXT) * 0.5 * RNI%)), 4, 13, 2, 2)
   TOTIVA(3).TEXT = CSTRING$(MKD$(CDbl(XVALO(TOTIVA(0).TEXT) + XVALO(TOTIVA(2).TEXT))), 4, 13, 2, 2)
   '
   Call CALPERIB
   '
   Total(3).TEXT = CSTRING$(MKD$(CDbl(XVALO(Total(2).TEXT) + XVALO(TOTIVA(3).TEXT)) + XVALO(Text15)), 4, 13, 2, 2)
   '
End Sub

Sub GENERODES()
   '
   If CONTROL$("ADMIPED", "GENDESPA") = "REVISTOCK" Then
'     Call GENODESPA(XVALO(Text11), Nodes&, 1, OKKY%)
     If OKKY% = 1 Then
'        Call PRIODESPA(Nodes&)
     End If
   End If
   '
End Sub
  
'
Sub CALPERIB()
   '
   'SE MODIFICA PARA PASARLE EL VALOR DE VARIDOC POR QUE LO ESTABA PASANDO SIN VALOR
   NCLIEN = XVALO(Text1)
   PIVADOC% = PIVACLI%(NCLIEN)
   '
   'SE ASIGNA VALOR DE VARIDOC
   VARIDOC% = 1
   If PIVADOC% = 0 Or PIVADOC% = 3 Or PIVADOC% = 4 Or PIVADOC% = 6 Then
        VARIDOC% = 2
      ElseIf PIVADOC% = 5 Or PIVADOC% = 7 Then
        VARIDOC% = 3
        If LETREXPO$ = "A" Then
          VARIDOC% = 1
        End If
   End If
    
   TASAX = ALIPERBRU#(NCLIEN, MINIPER#, VARIDOC%)
   '
   Text15 = ""
   BIMPON# = XVALO(Total(2))
   If BIMPON# * TICAMONE(MONEMI%) >= MINIPER# Then
      PERSALTA# = BIMPON# * (TASAX / 100)
      Text15 = FORMATNUM$(PERSALTA#, "F11.2")
   End If
   '
End Sub
'
Sub CALPERIB_VIE()
   '
   Static CTTXX%, TASAA, TASAB, TASAEXP, MINIPER
   '
   'If TIFADOC% > 2 Then
   '   Exit Sub
   'End If
   '
   If CTTXX% < 1 Then
     TASAA = XVALO(CONTROL$("PERIBRU", "TASA-A"))
     TASAB = XVALO(CONTROL$("PERIBRU", "TASA-B"))
     TASAEXP = XVALO(CONTROL$("PERIBRU", "TASA-EXP"))
     MINIPER = XVALO(CONTROL$("PERIBRU", "MINIPER"))
     CTTXX% = 1
   End If
   '
   TASAX = TASAA
   NCLIE = XVALO(Text1)
   If NCLIE < 1 Then Exit Sub
   '
   VARIDC% = 1: PVCLX% = PIVACLI%(NCLIE)
   If PVCLX% = 0 Or PVCLX% = 4 Then VARIDC% = 2
   If PVCLX% = 5 Or PVCLX% = 7 Then VARIDC% = 3
   '
   If VARIDC% = 2 Then TASAX = TASAB
   If VARIDC% = 3 Then TASAX = TASAEXP
   '
   Text15 = ""
   REPERCI$ = FILTERGETRECORD$("DAPIBCLI", 1, MKI$(NCLIE))
   CATIB1% = Asc(Mid$(REPERCI$, 3, 1))
   PORCONVE = CVS(Mid$(REPERCI$, 5, 4))
   If PORCONVE > 100 Then PORCONVE = 100
   '
   If CATIB1% > 0 Then         ' NO ESTA EXENTO
     If PORCONVE >= 10.1 Then  ' PARTICIPACION CONVENIO MULTILATERAL
       BIMPON# = XVALO(Total(2))
       If VARIDC% = 2 Then BIMPON# = BIMPON# + XVALO(TOTIVA(3))
       If BIMPON# * TICAMONE(MONEMI%) >= MINIPER Then
         PERSALTA# = (BIMPON# * PORCONVE / 100) * (TASAX / 100)
         Text15 = CSTRING$(MKD$(PERSALTA#), 3, 11, 2, 2)
       End If
     End If
   End If
   '
End Sub
'
Private Sub TOTIVA_GotFocus(Index As Integer)
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Private Sub TXTFLETE_DblClick()
   TXTFLETE = ""
End Sub

