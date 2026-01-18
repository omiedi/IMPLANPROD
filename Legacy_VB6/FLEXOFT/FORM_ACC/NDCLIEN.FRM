VERSION 5.00
Begin VB.Form NDCLIEN 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Nota de Débito de Cliente"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   600
   ClientWidth     =   8220
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   8220
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00400000&
      Height          =   7845
      Left            =   7300
      ScaleHeight     =   7785
      ScaleWidth      =   1155
      TabIndex        =   62
      Top             =   0
      Width           =   1215
      Begin VB.PictureBox Picture16 
         Height          =   6800
         Left            =   0
         ScaleHeight     =   6735
         ScaleWidth      =   1020
         TabIndex        =   69
         Top             =   0
         Visible         =   0   'False
         Width           =   1080
         Begin VB.Image Image3 
            Height          =   6800
            Left            =   0
            Picture         =   "NDCLIEN.frx":0000
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1095
         End
      End
      Begin VB.CommandButton Command14 
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
         Height          =   640
         Left            =   105
         Picture         =   "NDCLIEN.frx":2398
         Style           =   1  'Graphical
         TabIndex        =   68
         ToolTipText     =   "Aprueba y Graba Documento"
         Top             =   5040
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   540
         TabIndex        =   66
         TabStop         =   0   'False
         Top             =   5880
         Width           =   600
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   67
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command3 
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
         Height          =   690
         Left            =   105
         Picture         =   "NDCLIEN.frx":26A2
         Style           =   1  'Graphical
         TabIndex        =   65
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   900
         Width           =   650
      End
      Begin VB.CommandButton Command4 
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
         Height          =   640
         Left            =   105
         Picture         =   "NDCLIEN.frx":29AC
         Style           =   1  'Graphical
         TabIndex        =   64
         ToolTipText     =   "Resumen de Cuenta Corriente"
         Top             =   1890
         Width           =   650
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E0E0E0&
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
         Height          =   690
         Left            =   105
         Picture         =   "NDCLIEN.frx":2CB6
         Style           =   1  'Graphical
         TabIndex        =   63
         ToolTipText     =   "Cierra y Abandona la Aplicación"
         Top             =   210
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "NDCLIEN.frx":2E00
         Top             =   6230
         Width           =   1515
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   8520
      Left            =   -50
      TabIndex        =   0
      Top             =   100
      Width           =   12000
      Begin VB.TextBox Text6 
         Alignment       =   1  'Right Justify
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
         Left            =   5090
         MaxLength       =   4
         TabIndex        =   2
         Top             =   1280
         Width           =   610
      End
      Begin VB.Frame Frame2 
         Caption         =   "Comercial"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004040&
         Height          =   870
         Left            =   210
         TabIndex        =   40
         Top             =   1785
         Width           =   6960
         Begin VB.TextBox DescriTxt 
            BackColor       =   &H00FFFFFF&
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
            Left            =   1890
            MaxLength       =   64
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   360
            Width           =   4800
         End
         Begin VB.CommandButton Command21 
            Caption         =   "."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   0
            Left            =   1575
            TabIndex        =   43
            Top             =   1130
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
            Index           =   0
            Left            =   1155
            TabIndex        =   42
            Text            =   " "
            Top             =   1130
            Width           =   435
         End
         Begin VB.TextBox Eco21 
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
            Height          =   300
            Index           =   0
            Left            =   1890
            TabIndex        =   41
            Text            =   " "
            Top             =   1130
            Width           =   3270
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Referencia:"
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
            Index           =   0
            Left            =   315
            TabIndex        =   45
            Top             =   420
            Width           =   1485
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Vendedor:"
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
            Index           =   0
            Left            =   -3045
            TabIndex        =   44
            Top             =   1205
            Width           =   4110
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Contable"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004040&
         Height          =   870
         Left            =   210
         TabIndex        =   32
         Top             =   2730
         Width           =   6960
         Begin VB.TextBox Eco21 
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
            Height          =   300
            Index           =   3
            Left            =   1890
            TabIndex        =   36
            Text            =   " "
            Top             =   1130
            Width           =   3270
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
            Index           =   3
            Left            =   1155
            TabIndex        =   35
            Text            =   " "
            Top             =   1130
            Width           =   435
         End
         Begin VB.CommandButton Command21 
            Caption         =   "."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   3
            Left            =   1575
            TabIndex        =   34
            Top             =   1130
            Width           =   180
         End
         Begin VB.TextBox Text16 
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
            Left            =   1890
            TabIndex        =   5
            Top             =   345
            Width           =   1170
         End
         Begin VB.CommandButton Command12 
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
            Left            =   3045
            TabIndex        =   33
            Top             =   345
            Width           =   180
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Vendedor:"
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
            Index           =   3
            Left            =   -3045
            TabIndex        =   39
            Top             =   1205
            Width           =   4110
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cuenta:"
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
            Index           =   5
            Left            =   210
            TabIndex        =   38
            Top             =   420
            Width           =   1590
         End
         Begin VB.Label Label16 
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
            Height          =   300
            Left            =   3330
            TabIndex        =   37
            Top             =   345
            Width           =   3375
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Resumen Final"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004040&
         Height          =   2910
         Left            =   210
         TabIndex        =   23
         Top             =   3675
         Width           =   6960
         Begin VB.Frame Frame9 
            BorderStyle     =   0  'None
            Height          =   1065
            Left            =   1890
            TabIndex        =   46
            Top             =   735
            Width           =   4900
            Begin VB.TextBox Total 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Index           =   5
               Left            =   3150
               TabIndex        =   49
               TabStop         =   0   'False
               Text            =   " "
               Top             =   675
               Width           =   1650
            End
            Begin VB.TextBox Total 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Index           =   4
               Left            =   3150
               TabIndex        =   48
               TabStop         =   0   'False
               Text            =   " "
               Top             =   390
               Width           =   1650
            End
            Begin VB.TextBox Total 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Index           =   3
               Left            =   3150
               TabIndex        =   47
               TabStop         =   0   'False
               Text            =   " "
               Top             =   105
               Width           =   1650
            End
            Begin VB.Label Ltota 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "IVA-NI"
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
               Index           =   5
               Left            =   -410
               TabIndex        =   61
               Top             =   690
               Width           =   960
            End
            Begin VB.Label Ltota 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "IVA"
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
               Index           =   4
               Left            =   -525
               TabIndex        =   60
               Top             =   420
               Width           =   1065
            End
            Begin VB.Label Ltota 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "IVA"
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
               Index           =   2
               Left            =   -315
               TabIndex        =   59
               Top             =   135
               Width           =   855
            End
            Begin VB.Label Ltota 
               BackStyle       =   0  'Transparent
               Caption         =   "%"
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
               Index           =   7
               Left            =   1365
               TabIndex        =   58
               Top             =   135
               Width           =   225
            End
            Begin VB.Label Ltota 
               BackStyle       =   0  'Transparent
               Caption         =   "%"
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
               Index           =   8
               Left            =   1365
               TabIndex        =   57
               Top             =   420
               Width           =   225
            End
            Begin VB.Label Ltota 
               BackStyle       =   0  'Transparent
               Caption         =   "%"
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
               Index           =   9
               Left            =   1365
               TabIndex        =   56
               Top             =   690
               Width           =   225
            End
            Begin VB.Label PORIVA 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   0
               Left            =   630
               TabIndex        =   55
               Top             =   135
               Width           =   645
            End
            Begin VB.Label PORIVA 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   1
               Left            =   630
               TabIndex        =   54
               Top             =   420
               Width           =   645
            End
            Begin VB.Label PORIVA 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   2
               Left            =   630
               TabIndex        =   53
               Top             =   690
               Width           =   645
            End
            Begin VB.Label BASIVA 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Index           =   0
               Left            =   1680
               TabIndex        =   52
               Top             =   105
               Width           =   1380
            End
            Begin VB.Label BASIVA 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Index           =   1
               Left            =   1680
               TabIndex        =   51
               Top             =   390
               Width           =   1380
            End
            Begin VB.Label BASIVA 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Index           =   2
               Left            =   1680
               TabIndex        =   50
               Top             =   675
               Width           =   1380
            End
         End
         Begin VB.Frame Frame10 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   1890
            TabIndex        =   28
            Top             =   1785
            Width           =   4860
            Begin VB.TextBox Text15 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   3360
               TabIndex        =   8
               Top             =   210
               Width           =   1440
            End
            Begin VB.TextBox Text17 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   1155
               TabIndex        =   7
               Top             =   210
               Width           =   1440
            End
            Begin VB.Label Label19 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Perc. I-B: "
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   2625
               TabIndex        =   30
               ToolTipText     =   "Doble Click para Configurar Cliente"
               Top             =   60
               Width           =   645
            End
            Begin VB.Label Label23 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Perc. I.V.A.:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Left            =   315
               TabIndex        =   29
               ToolTipText     =   "Doble Click para Configurar Cliente"
               Top             =   105
               Width           =   750
            End
         End
         Begin VB.Frame Frame8 
            BorderStyle     =   0  'None
            Height          =   435
            Left            =   1890
            TabIndex        =   25
            Top             =   405
            Width           =   4845
            Begin VB.TextBox Total 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Index           =   2
               Left            =   3150
               TabIndex        =   6
               TabStop         =   0   'False
               Text            =   " "
               Top             =   30
               Width           =   1650
            End
            Begin VB.Label Ltota 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "ImporteTotal Neto:"
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
               Index           =   10
               Left            =   1050
               TabIndex        =   27
               Top             =   105
               Width           =   2010
            End
            Begin VB.Label MONEMIS 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Pesos"
               BeginProperty Font 
                  Name            =   "Arial Black"
                  Size            =   12
                  Charset         =   0
                  Weight          =   900
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00005500&
               Height          =   330
               Left            =   -105
               TabIndex        =   26
               ToolTipText     =   "Doble-Click para 'Dólares'"
               Top             =   0
               Visible         =   0   'False
               Width           =   1380
            End
         End
         Begin VB.TextBox Total 
            Alignment       =   1  'Right Justify
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
            Height          =   270
            Index           =   6
            Left            =   5040
            TabIndex        =   24
            TabStop         =   0   'False
            Text            =   " "
            Top             =   2445
            Width           =   1650
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Total Documento:"
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
            Index           =   11
            Left            =   3045
            TabIndex        =   31
            Top             =   2520
            Width           =   1905
         End
      End
      Begin VB.CommandButton Command1 
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
         Left            =   1050
         TabIndex        =   16
         Top             =   890
         Width           =   180
      End
      Begin VB.TextBox Text2 
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
         Height          =   300
         Left            =   1470
         TabIndex        =   15
         TabStop         =   0   'False
         Text            =   " "
         Top             =   105
         Width           =   3480
      End
      Begin VB.TextBox Text3 
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
         Height          =   300
         Left            =   1470
         TabIndex        =   14
         TabStop         =   0   'False
         Text            =   " "
         Top             =   490
         Width           =   3480
      End
      Begin VB.TextBox Text4 
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
         Height          =   300
         Left            =   1470
         TabIndex        =   13
         TabStop         =   0   'False
         Text            =   " "
         Top             =   890
         Width           =   750
      End
      Begin VB.TextBox Text5 
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
         Height          =   300
         Left            =   2310
         TabIndex        =   12
         TabStop         =   0   'False
         Text            =   " "
         Top             =   890
         Width           =   2640
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
         Left            =   6990
         TabIndex        =   11
         Top             =   105
         Width           =   180
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
         Height          =   300
         Left            =   6045
         TabIndex        =   10
         TabStop         =   0   'False
         Text            =   " "
         Top             =   105
         Width           =   960
      End
      Begin VB.TextBox Text11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   5775
         MaxLength       =   8
         TabIndex        =   3
         Text            =   " "
         Top             =   1280
         Width           =   1440
      End
      Begin VB.TextBox Text12 
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
         Height          =   300
         Left            =   1470
         MaxLength       =   40
         TabIndex        =   9
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1280
         Width           =   3480
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   0
         Top             =   0
      End
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
         Left            =   210
         MaxLength       =   5
         TabIndex        =   1
         Text            =   " "
         Top             =   890
         Width           =   855
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "ND"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   855
         Left            =   420
         TabIndex        =   21
         Top             =   0
         Width           =   1380
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00004040&
         Height          =   330
         Left            =   5250
         TabIndex        =   22
         Top             =   150
         Width           =   750
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H0000C000&
         FillColor       =   &H0080FF80&
         FillStyle       =   0  'Solid
         Height          =   390
         Left            =   315
         Shape           =   4  'Rounded Rectangle
         Top             =   180
         Width           =   750
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00C0C000&
         FillColor       =   &H00FFFF80&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   420
         Shape           =   4  'Rounded Rectangle
         Top             =   255
         Width           =   540
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5670
         TabIndex        =   20
         Top             =   1260
         Width           =   150
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Nota de Débito de Cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   645
         Left            =   5250
         TabIndex        =   19
         Top             =   735
         Width           =   1800
      End
      Begin VB.Label Label26 
         BackColor       =   &H00EEEEEE&
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
         Height          =   300
         Left            =   210
         TabIndex        =   18
         Top             =   1280
         Width           =   1035
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
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
         Height          =   330
         Left            =   105
         TabIndex        =   17
         Top             =   660
         Width           =   960
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H0000C0C0&
         FillColor       =   &H0080FFFF&
         FillStyle       =   0  'Solid
         Height          =   540
         Left            =   210
         Shape           =   4  'Rounded Rectangle
         Top             =   105
         Width           =   960
      End
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   1050
   End
   Begin VB.Menu archivo 
      Caption         =   "Archivo"
      Begin VB.Menu aprobar 
         Caption         =   "Aprobar"
      End
      Begin VB.Menu s 
         Caption         =   "-"
      End
      Begin VB.Menu salir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu accesodirectos 
      Caption         =   "Accesos Directos"
      Begin VB.Menu cuentacorriente 
         Caption         =   "Cuenta Corriente"
      End
   End
End
Attribute VB_Name = "NDCLIEN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TIFADOC%    ' 2=FACTURA / 3=N.DEBITO / 6=N.CREDITO
Dim MODODOC%    ' 0=MAT.CODIFICADO / 1=TEXTO LIBRE
Dim VARIDOC%
Dim RNIDOC%
Dim DATEDOC%
Dim PIVADOC%
Dim VERLICRE%
Dim MODESCU%
'
Dim CONSICLI%
'
Dim NUOCOMPRA$
Dim NROPEDIDO$
Dim NROREMICO$
'
Dim ARCHIPRE$
Dim DATCLI$(32)
'
Dim MONECO$, TICAMBX#

Private Sub aprobar_Click()
   Call Command14_Click
End Sub

Private Sub BASIVA_DBLClick(Index As Integer)
   '
   If Index <= 1 Then
       NCLIE% = XVALO(Text1)
       RNI% = 0: APLIVA% = 1
       If NCLIE% > 0 Then
         PPVI% = PIVACLI%(NCLIE%)
         If PPVI% = 2 Then RNI% = 1
         If PPVI% = 5 Or PPVI% = 7 Then APLIVA% = 0
       End If
   End If
   '
   Select Case Index
       Case 0 '21%
          If XVALO(BASIVA(0)) < 0.005 Then
            BASIVA(0) = BASIVA(1)
            Total(3).TEXT = CSTRING$(MKD$(XVALO(BASIVA(0).Caption) * XVALO(PORIVA(0).Caption) * APLIVA% / 100), 3, 12, 2, 2)
            BASIVA(1) = ""
            Total(4).TEXT = CSTRING$(MKD$(XVALO(BASIVA(1).Caption) * XVALO(PORIVA(1).Caption) * APLIVA% / 100), 3, 12, 2, 2)
          End If
          Call CALTOFA2
       Case 1 '10.5%
          If XVALO(BASIVA(1)) < 0.005 Then
            BASIVA(1) = BASIVA(0)
            AA1 = XVALO(BASIVA(1))
            aa2 = PORIVA(1)
            aa3 = APLIVA%
            Total(4).TEXT = CSTRING$(MKD$(XVALO(BASIVA(1).Caption) * XVALO(PORIVA(1).Caption) * APLIVA% / 100), 3, 12, 2, 2)
            BASIVA(0) = ""
            Total(3).TEXT = CSTRING$(MKD$(XVALO(BASIVA(0).Caption) * XVALO(PORIVA(0).Caption) * APLIVA% / 100), 3, 12, 2, 2)
          End If
          Call CALTOFA2
    End Select
    '
End Sub

Private Sub BASIVA_LostFocus(Index As Integer)
    BASIVA(Index) = FORMATNUM(BASIVA(Index), "F11.2")
End Sub

Private Sub Command1_Click()
   Call SELECHO("DEUDOR1")
   NCLIE% = XVALO(TRIM$(VALACT1$("DEUDOR1")))
   If NCLIE% > 0 Then
     Text1.TEXT = TRIM$(Str$(NCLIE%))
     If XVALO(Text1.TEXT) > 0 Then
       If XVALO(Text1.TEXT) <= 32767 Then
         NC% = XVALO(Text1)
         'If REGICLI&(NC%) > 0 Then
           Call CARDACLI
         'End If
       End If
     End If
   End If
   End Sub

Private Sub Command12_Click()
   Call SELECHO("ECUECON")
   VDEVU$ = TRIM$(VALACT1$("ECUECON"))
   If VDEVU$ <> "" Then
     Text16 = VDEVU$
   End If
End Sub

Private Sub Command14_Click()
   Command14.Enabled = False
   '
   Screen.MousePointer = 11
   '
   FECHALIM% = XVALO(CONTROL$("", "CIEMEVEN"))
   If FECHANUM(Text9) <= FECHALIM% Then
     Call MENSERR(24, "Fecha no Válida.\  \Ese Período Mensual Está Cerrado.")
     Command14.Enabled = True
     Exit Sub
   End If
   '
   'Valida que el Nº no sea 0. -21/12/04-
   If XVALO(TRIM$(Text11)) <= 0 Then
        Call MENSERR(24, "Imposible Registrar este Documento.\Numero de Nota de Debito Incorrecto. - Consulte.")
        Screen.MousePointer = 1
        Command14.Enabled = True
        Exit Sub
   End If
   '
   If TRIM$(Text6) = "" Or TRIM$(Text11) = "" Then
        Call MENSERR(24, "Imposible Registrar este Documento.\Numero de Nota de Debito Incompleto. - Consulte.")
        Screen.MousePointer = 1
        Command14.Enabled = True
        Exit Sub
   End If
   '
   For J = 0 To 2
    If XVALO(Total(J + 3)) > 0 Then
      If XVALO(BASIVA(J)) < 0.005 Then
        Call MENSERR(24, "Imposible Registrar este Documento.\Base Imponible de Iva de la Nota de Debito Incompleta. - Consulte.")
        Screen.MousePointer = 1
        Command14.Enabled = True
        Exit Sub
      End If
    End If
   Next J
   '
   For J = 0 To 2
    If XVALO(Total(J + 3)) < 0.005 Then
      If XVALO(BASIVA(J)) > 0.005 Then
        Call MENSERR(24, "Imposible Registrar este Documento.\Base Imponible de Iva sin Importe de Iva. - Consulte.")
        Screen.MousePointer = 1
        Command14.Enabled = True
        Exit Sub
      End If
    End If
   Next J
   '
   If TRIM$(DescriTxt.TEXT) = "" Then
    Call MENSERR(24, "Imposible Registrar este Documento.\Descripcion de la Nota de Debito Incompleta. - Consulte.")
    Screen.MousePointer = 1
    Command14.Enabled = True
    Exit Sub
   End If
   '
   If XVALO(Total(6)) <= 0.0005 Then
    Call MENSERR(24, "Imposible Registrar este Documento.\Debe Completar el Importe de la Nota de Debito. - Consulte.")
    Screen.MousePointer = 1
    Command14.Enabled = True
    Exit Sub
   End If
   '
   If TRIM$(Text16.TEXT) = "" Then
    Call MENSERR(24, "Imposible Registrar este Documento.\Cuenta Contable de la Nota de Debito Incompleta. - Consulte.")
    Screen.MousePointer = 1
    Command14.Enabled = True
    Exit Sub
   End If
   '
   If CUECOINT(Text16.TEXT) <= 0 Then
    Call MENSERR(24, "Imposible Registrar este Documento.\Cuenta Contable de la Nota de Debito Inexistente. - Consulte.")
    Screen.MousePointer = 1
    Command14.Enabled = True
    Exit Sub
   End If
   '
   'Valida que no se haya ingresado ya para el mismo cliente la
   'misma nota de debito de cliente. por el Nº de Nota Debito.
   Call APERBASDAT("CABAN")
   NUCLI = XVALO(Text1)
   SQLX$ = " SELECT * FROM CAJABANC "
   SQLX$ = SQLX$ + " WHERE NuClien = " & NUCLI & " "
   SQLX$ = SQLX$ + " AND NumeComp = " & XVALO(Text11) & " "
   SQLX$ = SQLX$ + " AND NumeTalo = " & XVALO(Text6) & " "
   SQLX$ = SQLX$ + " AND TIPOMOVIM = 'NDCLI' "
   SQLX$ = SQLX$ + " AND OpciMovim = 'CI' "
   Set YAREGTMP = CueCorri.OpenRecordset(SQLX, 4)
   With YAREGTMP
     On Error Resume Next
     .MoveFirst
     If Err Then
       On Error GoTo 0
       GoTo 6
     End If
      On Error GoTo 0
      Call MENSERR(24, "Imposible Registrar este Documento.\Nota de Debito Pre-Existente. - Consulte.")
      Screen.MousePointer = 1
      Command14.Enabled = True
      Exit Sub
   End With
   
   '
6 Call REGISNODE

   '
10 Call BLANFORMU
   Call CIERRARCH("*.*")
   Text1.TEXT = ""
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
   '
   Screen.MousePointer = 1
   Command14.Enabled = True
End Sub

Private Sub cuentacorriente_Click()
    Call Command4_Click
End Sub

Private Sub salir_Click()
    Call Command2_Click
End Sub

Private Sub Text15_Change()
   Call CALTOFA2
End Sub

Private Sub Text15_GotFocus()
   If VALICLIE% < 1 Then GoTo 99
   If TIFADOC% < 3 Then GoTo 98
   If XVALO(Total(2)) < 0.5 Then GoTo 98
   Exit Sub
   '
98 On Error Resume Next
   PORDESCU.SetFocus
   On Error GoTo 0
   Exit Sub
   '
99 On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Private Sub Text15_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    Text15 = FORMATNUM$(XVALO(Text15), "F11.2")
  End If
End Sub

Private Sub Text15_LostFocus()
   Text15 = FORMATNUM$(XVALO(Text15), "F11.2")
End Sub

Private Sub Command2_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command2_LostFocus()
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    Call HELPONLINE("FACTURA")
    Command3.Enabled = True
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    RECUEC04.Text1.TEXT = Text1.TEXT
    RECUEC04.Show 1
    Command4.Enabled = True
End Sub

Sub CARCONSIPEN()
   '
   Static DES%, HAS%
   '
   If DES% < 1 Then
2    FEX = HOY(HO$)
     HAS% = FEX
     HASS$ = FECHATEX$(FEX)
     DESS$ = "01" + Mid$(HASS$, 3)
     DES% = FECHANUM(DESS$)
   End If
   '
   VDEF$ = MKI$(DES%) + MKI$(HAS%)
   OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
   If OBX$ = "" Then
     CONSICLI% = 0
     Exit Sub
   End If
   '
   DES% = CVI(Left$(OBX$, 2))
   HAS% = CVI(Mid$(OBX$, 3, 2))
   If DES% > HAS% Then
      GoTo 2
   End If
   FEX = DES%: PERIO$ = FECHATEX$(FEX)
   FEX = HAS%: PERIO$ = PERIO$ + " al " + FECHATEX$(FEX)
   '
   DES1# = CDbl(CVDAT(DES%))
   HAS1# = CDbl(CVDAT(HAS%))
   '
   'LPRE% = XVALO(TEXT21(0).TEXT)
   NCLIE% = XVALO(Text1.TEXT)
   ARCHIPRE$ = TRIM$(Str$(LPRE%))
   While Len(ARCHIPRE$) < 3: ARCHIPRE$ = "0" + ARCHIPRE$: Wend
   ARCHIPRE$ = "LIPRE" + ARCHIPRE$
   If LPRE% < 1 Then ARCHIPRE$ = "MASTER"
   CONSICLI% = 0
   '
   REGODES& = 0
   Call SETULTREG("!CARDEFAC", 0)
   JJ& = 0: CONSICLI% = 0
   '
   CMO$ = TRIM$(CONTROL$("", "CODREFAC"))
   If Len(CMO$) <> 2 Then
     CMO$ = "RT"
   End If
   Screen.MousePointer = 11
   Call APERBASDAT("STOCKS")
   '
   SQLX$ = " SELECT * FROM MOVISTO "
   SQLX$ = SQLX$ + "WHERE Nume_Clie = " & NCLIE% & " "
   SQLX$ = SQLX$ + " AND (CodiMovi = '" & CMO$ & "' "
   SQLX$ = SQLX$ + " OR CodiMovi =  'DV') "
   SQLX$ = SQLX$ + " AND DoPriLar <> '' "
   SQLX$ = SQLX$ + " AND ((ISNULL(DoSecLar) = true) or (DoSecLar = '') OR (DoSecLar LIKE 'P/F*')) "
   SQLX$ = SQLX$ + " AND INT(CDBL(FechMov)) >= " & DES1# & " "
   SQLX$ = SQLX$ + " AND INT(CDBL(FechMov)) <= " & HAS1# & " "
   '
   Set CONSITMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
   With CONSITMP
     On Error Resume Next
     .MoveFirst
     If Err < 1 Then
       RECONSI$ = REGBLAN("CONSICLI")
       J& = 0
       Do While .EOF = False
         CODIMAT% = !Cod_Inte
         CAN = !CANTSALID - !CANTINGRE
         NPEDID& = 0: If IsNull(!NuPedCli) = False Then NPEDID& = !NuPedCli
5        z$ = REGBLAN$("!CARDEFAC")
         Call REPLA(z$, MKS$(CAN), 1, 4)      ' cantidad
           COD$ = CODEXT$(CODIMAT%)
         Call REPLA(z$, MKI$(CODIMAT%), 5, 2)
         Call REPLA(z$, DESCRIT$(CODIMAT%), 7, 64)
           NREMIT$ = !DoPriCor
         Call REPLA(z$, NREMIT$, 387, 16) ' numero de remito
           FEREMIT% = CVINT%(!FechMov)
         Call REPLA(z$, MKI$(FEREMIT%), 419, 2)
         Call REPLA(z$, NUPLANO$(CODIMAT%), 315, 24)
           '
           MONEMA% = !MoneVal
           COEFCON = TICAMONE(MONEMA%) / TICAMONE(MONEMI%)
           PRECIO = !VALOUNIT * COEFCON
           PRELI = PRELISTA#(ARCHIPRE$, CODIMAT%)
           If PRELI < PRECIO Then PRELI = PRECIO
           PORDESCUEN = 0: If PRELI >= 0.0001 Then PORDESCUEN = 100 * (PRELI - PRECIO) / PRELI
           '
         Call REPLA(z$, MKS$(PRELI), 421, 4)
         Call REPLA(z$, MKS$(PORDESCUEN), 425, 4)
         Call REPLA(z$, MKS$(PRECIO), 429, 4)
           IMPO# = CDbl(PRECIO) * CAN
         Call REPLA(z$, MKD$(IMPO#), 433, 8)
         Call REPLA(z$, MKS$(PORCEIVA(CODIMAT%)), 441, 4)
           NORIT% = !NuOrItem
         Call REPLA(z$, MKI$(NORIT%), 503, 2)
         Call REPLA(z$, MKS$(0), 505, 4)
         Call REPLA(z$, MKS$(0), 509, 4)
         Call REPLA(z$, TRIM$(Str$(NPEDID&)), 371, 16)
           NUOCOX$ = NUOCOCLI$(NPEDID&, CODIMAT%)
10       Call REPLA(z$, NUOCOX$, 355, 16)
         '
         Call REGAPP("!CARDEFAC", z$)
         CAITCAR% = CAITCAR% + 1
         '
       .MoveNext
       Loop
12   End If
   End With
   '
   JJ& = 0
   For U& = 1 To ULTREG&("!CARDEFAC")
     z$ = REGLEIDO$("!CARDEFAC", U&)
     If Abs(CVS(Mid$(z$, 1, 4))) >= 0.01 Then
       JJ& = JJ& + 1
       Call GRAREG("!CARDEFAC", z$, JJ&)
     End If
   Next U&
   Call SETULTREG("!CARDEFAC", JJ&)
   Call CIERRARCH("!CARDEFAC")
   '
   CONSICLI% = 1
   If JJ& < 1 Then
     Call MENSERR(24, "No hay Remitos Pendientes de Facturación\del " + PERIO$ + " para la Cuenta Elegida.")
     CONSICLI% = 0
   End If
   '
End Sub
'
Sub CARDICAMBIO()
   '
   Dim DOCABX$(2048), SALPES#(2048), SALDOL#(2048), REDOCA&(2048)
   '
   NC% = XVALO(Text1)
   Call BLANARCH("!CARDEFAC")
   '
   Screen.MousePointer = 11
   CADOCPE% = 0
   RFG$ = RECFILT$("CARBIMON", 2, MKI$(NC%))
   For KUL& = 1 To Len(RFG$) Step 4
     RECABIN& = CVS(Mid$(RFG$, KUL&, 4))
     XXC$ = REGLEIDO$("CARBIMON", RECABIN&)
     DOCABY$ = Mid$(XXC$, 13, 18)
     SALPESY# = CVD(Mid$(XXC$, 65, 8)) - CVD(Mid$(XXC$, 73, 8))
     SALDOLY# = CVD(Mid$(XXC$, 97, 8)) - CVD(Mid$(XXC$, 105, 8))
     '
     For KKI% = 1 To CADOCPE%
       If DOCABX$(KKI%) = DOCABY$ Then GoTo 20
     Next KKI%
     CADOCPE% = CADOCPE% + 1
     DOCABX$(CADOCPE%) = DOCABY$
     REDOCA&(CADOCPE%) = RECABIN&
     KKI% = CADOCPE%
     '
20   SALPES#(KKI%) = SALPES#(KKI%) + SALPESY#
     SALDOL#(KKI%) = SALDOL#(KKI%) + SALDOLY#
     '
   Next KUL&
   '
   Call BLANARCH("SELDOCAP")
   For KKI% = 1 To CADOCPE%
     If TIFADOC% = 6 Then    ' N.CREDITO
       If SALDOL#(KKI%) <= 0.005 Then
         If SALPES#(KKI%) > 0 Then
           XXC$ = REGLEIDO$("CARBIMON", REDOCA&(KKI%))
           FEMI = CVI(Mid$(XXC$, 1, 2))
           z$ = Space$(60) + MKS$(0)
           Call REPLA(z$, DOCABX$(KKI%), 1, 18)
           Call REPLA(z$, FECHATEX$(FEMI), 19, 8)
           Call REPLA(z$, "$" + FORMATNUM$(SALPES#(KKI%), "F11.2"), 29, 12)
           Call REPLA(z$, "u$s" + FORMATNUM$(SALDOL#(KKI%), "F11.2"), 44, 14)
           Call REPLA(z$, MKS$(REDOCA&(KKI%)), 61, 4)
           Call REGAPP("SELDOCAP", z$)
         End If
       End If
     ElseIf TIFADOC% = 3 Then  ' N.DEBITO
       If SALDOL#(KKI%) > 0 Then
         If SALPES#(KKI%) <= 0.005 Then
           XXC$ = REGLEIDO$("CARBIMON", REDOCA&(KKI%))
           FEMI = CVI(Mid$(XXC$, 1, 2))
           z$ = Space$(60) + MKS$(0)
           Call REPLA(z$, DOCABX$(KKI%), 1, 18)
           Call REPLA(z$, FECHATEX$(FEMI), 19, 8)
           Call REPLA(z$, "$" + FORMATNUM$(SALPES#(KKI%), "F11.2"), 29, 12)
           Call REPLA(z$, "u$s" + FORMATNUM$(SALDOL#(KKI%), "F11.2"), 44, 14)
           Call REPLA(z$, MKS$(REDOCA&(KKI%)), 61, 4)
           Call REGAPP("SELDOCAP", z$)
         End If
       End If
     End If
   Next KKI%
   Call CIERRARCH("SELDOCAP")
   '
   Screen.MousePointer = 1
   If ULTREG&("SELDOCAP") < 1 Then
     Exit Sub
   End If
   '
   Call SELECHO("SELDOCAP")
   DOCABY$ = Left$(VALACT1$("SELDOCAP"), 18)
   If TRIM$(DOCABY$) = "" Then Exit Sub
   DOCUDIF = DOCABY$
   BBXX$ = VALACT2$("SELDOCAP")
   '
   MONEMIS = "Pesos"
   IDIFE = XVALO(Mid$(BBXX$, 12, 11))
   If TIFADOC% = 3 Then
     IDIFE = XVALO(Mid$(BBXX$, 29, 11)) * XVALO(TICAMBIS)
   End If
    '
   ZZ$ = REGBLAN$("CARDEFAC")
   Call REPLA(ZZ$, MKS$(1), 1, 4)
   Call REPLA(ZZ$, "Por Diferencia de Cambio", 7, 64)
   Call REPLA(ZZ$, MKS$(IDIFE), 429, 4)
   Call REPLA(ZZ$, MKD$(CDbl(IDIFE)), 433, 8)
   Call REGAPP("!CARDEFAC", ZZ$)
   ZZ$ = REGBLAN$("CARDEFAC")
   Call REPLA(ZZ$, "s/Factura " + DOCABY$, 7, 64)
   Call REGAPP("!CARDEFAC", ZZ$)
   Call CIERRARCH("!CARDEFAC")
   '
End Sub
'
Private Sub Command9_Click()
   If VALICLIE% < 1 Then
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
      Exit Sub
   End If
   Call SELECHO("SELFECHA")
   Text9.TEXT = VALACT1$("SELFECHA")
   FF% = FECHANUM(Text9.TEXT)
End Sub


Private Sub Form_Load()
'    '
    If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      Picture16.Visible = True
      Picture16.ZOrder 0
      Me.BackColor = &HFCD7B6
      Me.Frame1.BackColor = &HFCD7B6
      Me.Frame2.BackColor = &HFCD7B6
      Me.Frame4.BackColor = &HFCD7B6
      Me.Frame6.BackColor = &HFCD7B6
      Me.Frame8.BackColor = &HFCD7B6
      Me.Frame9.BackColor = &HFCD7B6
      Me.Frame10.BackColor = &HFCD7B6
      '
      MARCOBARRA.Top = 5880
      Set MARCOBARRA.Container = Me.Picture16
      MARCOBARRA.ZOrder 0
      Set Command14.Container = Me.Picture16
      Command14.ZOrder 0
      
    End If

    Call CENTRAFORM(Me)
    HOII% = HOY(HOS$)
    Text9.TEXT = HOS$: DATEDOC% = HOII%
    '
    Call FORFACCS.CARGAPUERTOS
    '
    MONEMI% = 1  ' PESOS
    Call LEETASIVA
    '
    Screen.MousePointer = 1
    '
End Sub
'

Private Sub Form_Resize()
   If WindowState = 1 Then
     Call BAJALDISCO
   End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
'   Call CIERRARCH("*.*")
'   Hide
End Sub

Private Sub Label19_DblClick()
   Static CTXX%
   '
   If ULTREG&("DAPIBCLI") < 1 Then
     Call MENSERR(24, "Imposible desde esta Aplicación")
     Exit Sub
   End If
   '
   NCLIE% = XVALO(TRIM$(Text1.TEXT))
   If NCLIE% > 0 Then
     If ECOARCH$("DEUDOR1", MKI$(NCLIE%)) <> "" Then
       '
       If CTXX% < 1 Then
         Call GRAREG("TACATIBC", Chr$(0) + "Exento Percepcion IB", 1)
         Call GRAREG("TACATIBC", Chr$(1) + "Sujeto Percepcion IB", 2)
         Call SETULTREG("TACATIBC", 2)
         CTXX% = 1
       End If
       '
       RECORD$ = FILTERGETRECORD$("DAPIBCLI", 1, MKI$(NCLIE%))
       Call REPLA(RECORD$, MKI$(NCLIE%), 1, 2)
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
         Call REPLA(RECORD$, MKI$(NCLIE%), 1, 2)
         Call REPLA(RECORD$, USERNAME$, 41, 20)
         Call REPLA(RECORD$, MKI$(HOY(HOS$)), 61, 2)
         Call REPLA(RECORD$, MKI$(Int(HORANUM(Time$))), 63, 2)
         Call FILTERPUTRECORD("DAPIBCLI", 1, MKI$(NCLIE%), RECORD$)
         Call CIERRARCH("DAPIBCLI")
         '
         
         '
       End If
     End If
   End If

End Sub

Private Sub Label26_DblClick()
   Call SELECHO("INDICLIE")
   CPROX$ = TRIM$(VALACT1$("INDICLIE"))
   If CPROX$ <> "" Then
     NCXI% = NUMECLI%(CPROX$)
     If NCXI% > 0 Then
       If REGICLI&(NCXI%) > 0 Then
         Label26 = CPROX$
         NC% = NCXI%
         Text1.TEXT = TRIM$(Str$(NC%))
         Call CARDACLI
       End If
     End If
   End If
End Sub

Private Sub Label7_Click()
   '
   NCLIE% = XVALO(Text1.TEXT)
   '
End Sub


Private Sub MONEMIS_Change()
   MONEMA% = MONEMI%
   If Left$(MONEMIS, 1) = "D" Then
        MONEMI% = 2
        MONEMIS.ForeColor = RGB(150, 0, 0)
        MONEMIS.ToolTipText = "Doble-Click para 'Pesos'"
        'Label12.Visible = True
        If TICAMBIS.Visible = True Then TICAMBIS.Enabled = True
        TICAMBIS = FORMATNUM$(TICAMONE(2), "F9.4")
        TICAMBX# = XVALO(TICAMBIS)
      Else
        MONEMI% = 1
        MONEMIS.ForeColor = RGB(0, 100, 0)
        MONEMIS.ToolTipText = "Doble-Click para 'Dólares'"
        'Label12.Visible = False
        If TICAMBIS.Visible = True Then TICAMBIS.Enabled = False
        TICAMBX# = 1
   End If
   '
   COEFCON = TICAMONE(MONEMA%) / TICAMONE(MONEMI%)
   FIN& = ULTREG&("!CARDEFAC")
   For U& = 1 To FIN&
     X$ = REGLEIDO$("!CARDEFAC", U&)
     CAN = CVS(Mid$(X$, 1, 4))
     PRELI = COEFCON * CVS(Mid$(X$, 421, 4))
     PRECIO = COEFCON * CVS(Mid$(X$, 429, 4))
     IMPO# = CDbl(CAN * PRECIO)
     '
     Call REPLA(X$, MKS$(PRELI), 421, 4)
     Call REPLA(X$, MKS$(PRECIO), 429, 4)
     Call REPLA(X$, MKD$(IMPO#), 433, 8)
     Call GRAREG("!CARDEFAC", X$, U&)
   Next U&
   '
   Call CARLISFAC
   Call CIERRARCH("!CARDEFAC")
   Call CALTOFAC
   '
End Sub

Private Sub MONEMIS_DblClick()
   If TRIM$(UCase$(MONEMIS)) = "PESOS" Then
       MONEMIS = "Dólares"
     Else
       MONEMIS = "Pesos"
   End If
End Sub

Private Sub PORDESCU_Change()
   Call CALTOFAC
End Sub

Private Sub PORDESCU_GotFocus()
   If VALICLIE% < 1 Then
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
      Exit Sub
   End If
End Sub



Private Sub Text1_Change()
   Call BLANFORMU
End Sub

Private Sub Text1_DblClick()
  Call SELECHO("DEUDOR1")
  NCLIE% = XVALO(TRIM$(VALACT1$("DEUDOR1")))
  If NCLIE% > 0 Then
    Text1.TEXT = TRIM$(Str$(NCLIE%))
    If XVALO(Text1.TEXT) > 0 Then
      If XVALO(Text1.TEXT) <= 32767 Then
        NC% = XVALO(Text1)
        If REGICLI&(NC%) > 0 Then
          Call CARDACLI
        End If
      End If
    End If
  End If
End Sub

Private Sub Text1_GotFocus()
   Text1.TEXT = TRIM$(Text1.TEXT)
   Text1.SelStart = 0
   Text1.SelLength = Len(Text1.TEXT)

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    If VALICLIE% > 0 Then
      Call BLANFORMU
      If XVALO(Text1.TEXT) > 0 Then
        If XVALO(Text1.TEXT) <= 32767 Then
          NC% = XVALO(Text1)
          If REGICLI&(NC%) > 0 Then

               Call CARDACLI

            On Error Resume Next
            Command6.SetFocus
            On Error GoTo 0
          End If
        End If
      End If
    End If
  End If
End Sub

Private Sub Text11_GotFocus()
   If VALICLIE% < 1 Then
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
      Exit Sub
   End If
   If XVALO(Text11.TEXT) > 9 Then
      On Error Resume Next
      Text10.SetFocus
      On Error GoTo 0
   End If
End Sub

Private Sub Text11_LostFocus()
    '
    On Error Resume Next
    Do While Len(TRIM$(Text11)) < 8
     Text11 = "0" + Text11
    Loop
    On Error GoTo 0
    '
End Sub

Private Sub Text12_GotFocus()
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Private Sub Text16_Change()
   Label16 = Left$(ECOARCH("ECUECON", Text16), 20)
End Sub

Private Sub TEXT17_CHANGE()
   Call CALTOFA2
End Sub
'
Private Sub Text17_GotFocus()
   Text17.SelStart = 0
   Text17.SelLength = Len(Text17)
End Sub

Private Sub Text17_LostFocus()
    '
    Text17 = FORMATNUM$(XVALO(Text17), "F11.2")
    '
End Sub

Private Sub Text2_gotfocus()
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
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

Private Sub Text5_GotFocus()
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub



   '
Function VALICLIE%()
   VCC% = 0
   If XVALO(Text1.TEXT) > 0 Then
     If XVALO(Text1.TEXT) <= 32767 Then
       NCLIE% = XVALO(Text1.TEXT)
       If ECOARCH$("DEUDOR1", MKI$(NCLIE%)) <> "" Then
         VCC% = 1
       End If
     End If
   End If
   '
99 VALICLIE% = VCC%
   '
End Function

Sub BLANFORMU()
   Text2.TEXT = ""
   Text3.TEXT = ""
   Text4.TEXT = ""
   Text5.TEXT = ""
   Text6.TEXT = ""
   Text11.TEXT = ""
   DescriTxt.TEXT = ""
   '
   For J = 2 To 6
    Total(J) = ""
   Next J
   '
   For J = 0 To 2
    BASIVA(J) = ""
   Next J
   '
   Text12.TEXT = ""

   Text15 = ""
   Text16 = ""
   Label16 = ""
   Label26 = ""
   Text17 = "" 'LIMPIA RET-IVA
   '
   DOCUDIF = ""
   '
   HOII% = HOY(HOS$)
   Text9.TEXT = HOS$
   Call SETULTREG("!CARDEFAC", 0)
   If EXISTE%(LUCOM$ + "1CARDEFA.WKF") > 0 Then
     Call SETULTREG("!1CARDEFA", 0)
   End If
   '
   Call CIERRARCH("*.*")
   '
End Sub
'
Sub CARDACLI()
    '
    Screen.MousePointer = 11
    '
    NC% = XVALO(Text1.TEXT): Text1.Refresh
    If NC% > 0 Then
      If ECOARCH$("DEUDOR1", MKI$(NC%)) <> "" Then
        Text2.TEXT = RASOCLI$(NC%): Text2.Refresh
        Text3.TEXT = DOMICLI$(NC%): Text3.Refresh
        Text4.TEXT = CPOSCLI$(NC%): Text4.Refresh
        Text5.TEXT = LOCACLI$(NC%): Text5.Refresh
        Label26 = CODICLI$(NC%)
        Text12.TEXT = TRIM$(POSIVAC$(NC%)) + " - CUIT " + CUITCLI$(NC%)
          Text12.Refresh
        CREDISPO# = LIMICRE(NC%) - SALDCLI#(NC%) - CHEPEACRE#(NC%)

        NCLIE% = XVALO(Text1.TEXT)
        If NCLIE% > 0 Then
        End If
        PIVADOC% = PIVACLI%(NCLIE%)
        '
        PPVI% = PIVADOC%
        If PPVI% < 0 Or PPVI% > 7 Then PPVI% = 1
        Text16 = CONTROL$("SVD451S", "CUEVE-" + TRIM$(Str$(PPVI%)))
        '
        LENTREG$ = Space$(96)
        RFEX$ = RECFILT$("LENTRECL", 1, MKI$(NC%))
        If Len(RFEX$) >= 4 Then
          RELENT& = CVS(Left$(RFEX$, 4))
          If RELENT& > 0 Then
            If RELENT& <= ULTREG&("LENTRECL") Then
              LENTREG$ = Mid$(REGLEIDO$("LENTRECL", RELENT&), 3)
            End If
          End If
        End If
        '
        RNC$ = RECFILT$("NOTACLI", 1, MKI$(NC%))
        URN& = ULTREG&("NOTACLI")

        '
        For U& = 1 To Len(RNC$) Step 4
          RENOTA& = CVS(Mid$(RNC$, U&, 4))
          If RENOTA& > 0 Then
            If RENOTA& <= URN& Then
              TTYY$ = RTrim$(Mid$(REGLEIDO$("NOTACLI", RENOTA&), 3))
              If Text6.TEXT <> "" Then
                Text6.TEXT = Text6.TEXT + Chr$(13) + Chr$(10)
              End If
              Text6.TEXT = Text6.TEXT + RTrim$(TTYY$)
            End If
          End If
        Next U&
        '
        If VERLICRE% = 1 Then
          LMCR# = LIMICRE(NC%)
          SDEU# = SALDCLI#(NC%)
          CHPA# = CHEPEACRE#(NC%)
          If SDEU# + CHPA# >= LMCR# Then
            Call COMUNI("Imposible Facturar a este Cliente.\  \Límite de Crédito Asignado ya está Superado.\  \Informe al Supervisor de Ventas.")
            Call BLANFORMU
            On Error Resume Next
            Text1.SetFocus
            On Error GoTo 0
          End If
        End If
        '
      End If
    End If
    Screen.MousePointer = 1
    '
End Sub
    
Sub CARLETDOC()
   '
   VARIDOC% = 1
   RNIDOC% = 0
   '
   If PIVADOC% = 0 Or PIVADOC% = 3 Or PIVADOC% = 4 Or PIVADOC% = 6 Then

        VARIDOC% = 2
        '
      ElseIf PIVADOC% = 2 Then
        RNIDOC% = 1
        '
      ElseIf PIVADOC% = 5 Or PIVADOC% = 7 Then
        Label13.Caption = "X"
        VARIDOC% = 3
        If TRIM$(LETREXPO$) <> "" Then
          Label13.Caption = LETREXPO$
        End If
        If LETREXPO$ = "A" Then
          VARIDOC% = 1
          Label13.Caption = "A"
        End If
    End If
    '
End Sub
   '

Private Sub Text6_lostFocus()
   On Error Resume Next
   J = Len(TRIM$(Text6))
   Do While Len(TRIM$(Text6)) < 4
    Text6 = "0" + Text6
   Loop
   On Error GoTo 0
End Sub

Private Sub Text9_GotFocus()
   On Error Resume Next
   Command9.SetFocus
   On Error GoTo 0
End Sub

Private Sub TICAMBIS_Change()
   If TICAMBIS.Enabled = True Then
     TICACH(MONEMI%) = XVALO(TICAMBIS)
   End If
   TICAMBX# = 1
   If MONEMI% = 2 Then
     TICAMBX# = XVALO(TICAMBIS)
   End If
End Sub

Private Sub TICAMBIS_LostFocus()
   TICAMBIS = FORMATNUM(XVALO(TICAMBIS), "F9.4")
End Sub

Private Sub Timer1_Timer()
   '
   Call CALTOFAC
   '
End Sub

'Sub PRIFACTURA()
'    '
'    CAHOJAS% = 0
'    FORIPRE$ = TRIM$(CONTROL$("", "FORFAC-" + Label13.Caption))
'    If FORIPRE$ = "" Then
'       Call MENSERR(24, "Falta Definir Controlador de Impresión")
'       Exit Sub
'    End If
'    '
'    If EXISTE%(LUDAT$ + FORIPRE$ + ".FRM") < 1 Then
'       Call MENSERR(24, "Imposible Imprimir Documento.\No Existe Formulario Electrónico '" + FORIPRE$ + ".FRM'.")
'       Exit Sub
'    End If
'    '
'    Screen.MousePointer = 11
'    ITPORHOJA% = XVALO(ATRIFORM$(FORIPRE$, "MAX-ITEM"))
'    If ITPORHOJA% < 1 Then
'      ITPORHOJA% = 32767
'    End If
'    '
'    CANCELPRINT% = 0
'    FECHDOC$ = Text9.TEXT
'    NUMEDOC$ = TRIM$(Text11.TEXT)
'    While Len(NUMEDOC$) < 8
'      NUMEDOC$ = "0" + NUMEDOC$
'    Wend
'    NUMEDOC$ = Label12.Caption + "-" + NUMEDOC$
'    '
'
'    '
''    If InStr(UCase$(TIPODOC$), "FACTURA") > 0 Then'\\\OT-06-0110-WAR-29/03/06///
''      If InStr(UCase$(TIPODOC$), "CONTADO") < 1 Then
''        If CPAG% < 1 Or InStr(UCase$(DECONPA$(CPAG%)), "CONTADO") > 0 Then
''          TIPODOC$ = TIPODOC$ + " CONTADO"
''        End If
''      End If
''    End If
'    '
'    DESTIDOC% = XVALO(Text1.TEXT)
'    '
'    CODTABU1$(1) = "COD-MAT"
'    CODTABU1$(2) = "DES-MAT"
'    CODTABU1$(3) = "CANTIDAD"
'    'CODTABU1$(4) = "PRE-LIST"
'    'CODTABU1$(5) = "PORDESCU"
'    'CODTABU1$(6) = "PRE-NETO"
'    'CODTABU1$(7) = "NET-ITEM"
'    'CODTABU1$(8) = "FECH-REC"
'    CODTABU1$(9) = "ORD-ITEM"
'    CODTABU1$(10) = "UNIMED"
'    CODTABU1$(11) = "REF-MAT1"
'    CODTABU1$(12) = "REF-MAT2"
'    CODTABU1$(13) = "REF-MAT3"
'    CODTABU1$(14) = "N-PLANO"
'    CODTABU1$(15) = "N-ORIGIN"
'    CODTABU1$(16) = "N-OCOMPR"
'    CODTABU1$(17) = "N-REMITO"
'    CODTABU1$(18) = "N-BOLREC"
'    CODTABU1$(19) = "FECH-REC"
'    CODTABU1$(20) = "PRE-LIST"
'    CODTABU1$(21) = "PORDESCU"
'    CODTABU1$(22) = "PRE-NETO"
'    CODTABU1$(23) = "NET-ITEM"
'    CODTABU1$(24) = "ALICUIVA"
'    CODTABU1$(25) = "IVA-ITEM"
'    CODTABU1$(26) = "BRU-ITEM"
'    '
'    CACOLTAB1% = 26
'    '
'    CAITAB1% = 0
'    CAITCRU% = 0
'    '
'    SUMATOTA# = XVALO(Total(0).TEXT)
'    IDESCUEN# = XVALO(Total(1).TEXT)
'    '
'    NroRemi$ = TRIM$(Text10.TEXT)
'    If NroRemi$ <> "" Then
'      While Len(NroRemi$) < 8: NroRemi$ = "0" + NroRemi$: Wend
'      PREFIREM$ = CONTROL$("", "PREFIREM")
'      If PREFIREM$ = "" Then PREFIREM$ = "0001"
'      NroRemi$ = PREFIREM$ + "-" + NroRemi$
'    End If
'    '
'    If CONSICLI% = 1 Then
'      NroRemi$ = NROREMICO$
'    End If
'    '
'    COEFIVAINCLU# = 1
'    If XVALO(Total(2).TEXT) > 0.005 Then
'      COEFIVAINCLU# = (XVALO(Total(2).TEXT) + XVALO(Total(3).TEXT) + XVALO(Total(4).TEXT) + XVALO(Total(5).TEXT)) / XVALO(Total(2).TEXT)
'    End If
'    '
'    If TRIM$(Text14.TEXT) <> "" Then
'      NUOCOMPRA$ = TRIM$(Text14.TEXT)
'    End If
'    '
'    CODPARAM$(1) = "CIVA-CLI": DOCPARAM$(1) = Label13.Caption
'    CODPARAM$(2) = "NRO-REMI": DOCPARAM$(2) = NroRemi$
'    CODPARAM$(3) = "NRO-OCOM":  DOCPARAM$(3) = NUOCOMPRA$
'    CODPARAM$(4) = "NRO-AENT":  DOCPARAM$(4) = NROPEDIDO$
'    CODPARAM$(5) = "IMP-NETO": DOCPARAM$(5) = ""
'    'CODPARAM$(6) = "VENDEDOR": DOCPARAM$(6) = Eco21(2).TEXT
'    'CODPARAM$(7) = "COND-PAG": DOCPARAM$(7) = Eco21(1).TEXT
'    'CODPARAM$(8) = "POR-IVA": DOCPARAM$(8) = PORIVA(0).Caption
'    CODPARAM$(9) = "IMP-IVA": DOCPARAM$(9) = ""
'    'CODPARAM$(10) = "POR-IVA1": DOCPARAM$(10) = PORIVA(1).Caption
'    CODPARAM$(11) = "IMP-IVA1": DOCPARAM$(11) = ""
'    CODPARAM$(12) = "IMP-DESC": DOCPARAM$(12) = ""
'    CODPARAM$(13) = "IMPNEDES": DOCPARAM$(13) = ""
'    'CODPARAM$(4) = "NRO-OCOM": DOCPARAM$(4) = Text10.TEXT
'    CODPARAM$(14) = "IMP-TOTA": DOCPARAM$(14) = ""
'    CODPARAM$(15) = "POR-DESC": DOCPARAM$(15) = TRIM$(PORDESCU.TEXT)
'    CODPARAM$(16) = "IMP-RNI": DOCPARAM$(16) = ""
'    CODPARAM$(17) = "IMP-LET1": DOCPARAM$(17) = ""
'    CODPARAM$(18) = "IMP-LET2": DOCPARAM$(18) = ""
'    CODPARAM$(19) = "IMP-LET3": DOCPARAM$(19) = ""
'    CODPARAM$(23) = "TOTA-DOL": DOCPARAM$(23) = ""
'    '
'   ' CODPARAM$(20) = "DOMI-TRA": DOCPARAM$(20) = LENTREGA(0).TEXT
'   ' CODPARAM$(21) = "DESTINO": DOCPARAM$(21) = LENTREGA(1).TEXT
'   ' CODPARAM$(22) = "TRANSPOR": DOCPARAM$(22) = LENTREGA(2).TEXT
'    CODPARAM$(24) = "MONEDEMI": DOCPARAM$(24) = MONEMIS
'    CODPARAM$(25) = "PERC-BRU": DOCPARAM$(25) = ""
'    CODPARAM$(26) = "SUB-TOTA": DOCPARAM$(26) = ""
'    CODPARAM$(27) = "PERC-IVA": DOCPARAM$(27) = "" 'PERCIVA
'    CAPARDOC% = 27
'    '
'    For KKIL& = 0 To 5
'       Call GRAREG("!OBSERVOF", FINFACTURA.Text1(KKIL&).TEXT, KKIL& + 1)
'    Next KKIL&
'    Call SETULTREG("!OBSERVOF", 6)
'    Call CIERRARCH("!OBSERVOF")
'    '
'    PRINTERPORT$ = PORTUFAC$
'    CACOPDOC% = XVALO(FINFACTURA.Text2(0).TEXT)
'    '
'    NOCONFIRM% = 0
'    If TRIM$(UCase$(CONTROL$("FACTUREM", "IMPRENUM"))) <> "AUTONUME" Then
'      If CONTROL$("FACTUREM", "FORMULAR") <> "HOJASUEL" Then
'        NOCONFIRM% = 1
'      End If
'    End If
'    '
'    HOJASUEL% = 0
'    If CONTROL$("FACTUREM", "FORMULAR") = "HOJASUEL" Then
'       HOJASUEL% = 1
'    End If
'    '
'    FIN& = ULTREG&("!CARDEFAC")
'    SUBTOTA# = 0
'    For U& = 1 To FIN&
'      If (U& Mod ITPORHOJA%) = 1 Then
'        ' PONE ENCABEZADO
'        CAITAB1% = 0
'        CAHOJAS% = CAHOJAS% + 1
'      End If
'      '
'      XXX$ = REGLEIDO$("!CARDEFAC", U&)
'      CIPRE% = CVI(Mid$(XXX$, 5, 2))
'      CAPEX$ = Str$(CVS(Left$(XXX$, 4)))       ' cantidad
'      '
'      CAITAB1% = CAITAB1% + 1
'      TETABU1$(1, CAITAB1%) = CODEXT$(CIPRE%)
'      TETABU1$(2, CAITAB1%) = Mid$(XXX$, 7, 64)
'      TETABU1$(3, CAITAB1%) = CAPEX$
'      'TETABU1$(4, CAITAB1%) = PRELI$
'      'TETABU1$(5, CAITAB1%) = DESCX$
'      'TETABU1$(6, CAITAB1%) = PREUX$
'      'TETABU1$(7, CAITAB1%) = ITOTX$
'      'TETABU1$(8, CAITAB1%) = FENTX$
'      TETABU1$(9, CAITAB1%) = Str$(U&)
'      TETABU1$(10, CAITAB1%) = Mid$(XXX$, 71, 4)                           ' unidad de medida
'      TETABU1$(11, CAITAB1%) = Mid$(XXX$, 75, 64)                ' referencia 1
'      TETABU1$(12, CAITAB1%) = Mid$(XXX$, 139, 64)                ' referencia 1
'      TETABU1$(13, CAITAB1%) = Mid$(XXX$, 203, 64)                ' referencia 1
'      TETABU1$(14, CAITAB1%) = Mid$(XXX$, 315, 24)  ' numero de plano
'      TETABU1$(15, CAITAB1%) = Mid$(XXX$, 339, 16)  ' numero original
'      TETABU1$(16, CAITAB1%) = Mid$(XXX$, 355, 16)  ' numero de o.compra
'      TETABU1$(17, CAITAB1%) = Mid$(XXX$, 387, 16)  ' numero de remito
'      TETABU1$(18, CAITAB1%) = Mid$(XXX$, 403, 16)  ' numero de doc-recepcion
'      TETABU1$(19, CAITAB1%) = FECHATEX$(CVI(Mid$(XXX$, 419, 2))) ' fecha de recepcion
'      '
'      'FACIVITEM = 1: If Label13.Caption = "B" Then FACIVITEM = 1 + (CVS(Mid$(XXX$, 441, 4))) / 100
'      'LPRE% = XVALO(TEXT21(0).TEXT)
'      'If LPRE% > 0 Then
'      '  TETABU1$(20, CAITAB1%) = Str$(FACIVITEM * CVS(Mid$(XXX$, 421, 4))) ' precio de lista
'      '  TETABU1$(21, CAITAB1%) = Str$(CVS(Mid$(XXX$, 425, 4))) ' PORCENTAJE DE DESCUENTO
'      'End If
'      '
'      TETABU1$(22, CAITAB1%) = Str$(FACIVITEM * CVS(Mid$(XXX$, 429, 4))) ' precio unitario
'      TETABU1$(23, CAITAB1%) = Str$(CVD(Mid$(XXX$, 433, 8))) ' total neto
'      TETABU1$(24, CAITAB1%) = Str$(CVS(Mid$(XXX$, 441, 4))) ' alicuota iva
'      TETABU1$(25, CAITAB1%) = Str$(CVD(Mid$(XXX$, 445, 8))) ' importe iva del item
'      TETABU1$(26, CAITAB1%) = Str$(CVD(Mid$(XXX$, 453, 8))) ' total bruto
'      '
'      SUBTOTA# = SUBTOTA# + XVALO(TETABU1$(23, CAITAB1%))
'      '
'      If (U& Mod ITPORHOJA%) = 0 Then
'        If U& < FIN& Then
'          If CACOPDOC% > 0 Then
'            DOCPARAM$(26) = TRIM$(CSTRING$(MKD$(SUBTOTA#), 4, 12, 2, 2))
'            Call PRINTDOC("@" + FORIPRE$)
'          End If
'        End If
'      End If
'109 Next U&
'    '
'    CODPARAM$(5) = "IMP-NETO": DOCPARAM$(5) = Total(0).TEXT
'    CODPARAM$(9) = "IMP-IVA": DOCPARAM$(9) = Total(3).TEXT
'    CODPARAM$(11) = "IMP-IVA1": DOCPARAM$(11) = Total(4).TEXT
'    CODPARAM$(12) = "IMP-DESC": DOCPARAM$(12) = Total(1).TEXT
'    CODPARAM$(13) = "IMPNEDES": DOCPARAM$(13) = Total(2).TEXT
'    CODPARAM$(14) = "IMP-TOTA": DOCPARAM$(14) = Total(6).TEXT
'    CODPARAM$(16) = "IMP-RNI": DOCPARAM$(16) = Total(5).TEXT
'    If TICAMBIO > 0 Then
'       VADOLA# = XVALO(Total(6).TEXT) / TICAMBIO
'       DOCPARAM$(23) = TRIM$(CSTRING$(MKD$(VADOLA#), 4, 12, 2, 2))
'    End If
'    CODPARAM$(25) = "PERC-BRU": DOCPARAM$(25) = Text15
'    DOCPARAM$(26) = ""
'    DOCPARAM$(27) = Text17                          'PARA IMPRIMIR PERC-IVA
'    '
'    If Label13.Caption = "B" Then
'      DOCPARAM$(5) = Str$(COEFIVAINCLU# * XVALO(Total(0).TEXT))
'      DOCPARAM$(12) = Str$(COEFIVAINCLU# * XVALO(Total(1).TEXT))
'      DOCPARAM$(13) = Str$(COEFIVAINCLU# * XVALO(Total(2).TEXT))
'    End If
'    '
'    CODFOR$ = FORIPRE$
'    LOIMPLET% = XVALO(ATRIFORM$(CODFOR$, "LOIMPLET"))
'    If LOIMPLET% < 16 Then LOIMPLET% = 32
'    TTTT# = XVALO(Total(6).TEXT)
'    TOTFAC$ = NUMTEX$(TTTT#)
'    If MONECO$ <> "" Then
'      PPXX% = InStr(TOTFAC$, "PESOS")
'      If PPXX% > 0 Then
'        TOTFAC$ = Left$(TOTFAC$, PPXX% - 1) + UCase$(MONECO$) + Mid$(TOTFAC$, PPXX% + 5)
'      End If
'    End If
'    Call FRATEX(TOTFAC$, NRX1$, NRX2$, NRX3$, LOIMPLET%)
'    CODPARAM$(17) = "IMP-LET1": DOCPARAM$(17) = NRX1$
'    CODPARAM$(18) = "IMP-LET2": DOCPARAM$(18) = NRX2$
'    CODPARAM$(19) = "IMP-LET3": DOCPARAM$(19) = NRX3$
'    '
'    If CACOPDOC% > 0 Then
'      Call PRINTDOC("@" + FORIPRE$)
'    End If
'    '
'    CACOPDOC% = XVALO(FINFACTURA.Text2(1).TEXT)
'    If CACOPDOC% > 0 Then
'      If TRIM$(Text10.TEXT) <> "" Then
'         For KKI% = 0 To CACOLTAB1%
'           For KKJ% = 0 To CAITAB1%
'             TETABU1$(KKI%, KKJ%) = ""
'           Next KKJ%
'         Next KKI%
'         '
'        Call PRIREMIFA
'      End If
'    End If
'    '
'    Call CIERRARCH("*.*")
'    Call BAJALDISCO
'    '
'    If NOCONFIRM% > 0 Then
'       Call COMUNI("Se han Emitido los Documentos Solicitados")
'    End If
'    '
'End Sub

'Sub PRIREMIFA()
'    '
'    FORIPRE$ = TRIM$(CONTROL$("", "FOREMIFA"))
'    If TRIM$(FORIPRE$) <> "" Then
'       '
'       For KKIL& = 0 To 2
'          Call GRAREG("!OBSERVOF", FINFACTURA.Text1(KKIL& + 6).TEXT, KKIL& + 1)
'       Next KKIL&
'       Call SETULTREG("!OBSERVOF", 3)
'       Call CIERRARCH("!OBSERVOF")
'       '
'       PREFIREM$ = CONTROL$("", "PREFIREM")
'       If PREFIREM$ = "" Then PREFIREM$ = "0001"
'       NroRemi$ = TRIM$(Text10.TEXT)
'       While Len(NroRemi$) < 8: NroRemi$ = "0" + NroRemi$: Wend
'       NroRemi$ = PREFIREM$ + "-" + NroRemi$
'       '
'       NC% = XVALO(Text1.TEXT)
'       FECHDOC$ = Text9.TEXT
'       NUMEDOC$ = NroRemi$
'       TIPODOC$ = "Remito"
'       '
'       If TRIM$(Text14.TEXT) <> "" Then
'         NUOCOMPRA$ = TRIM$(Text14.TEXT)
'       End If
'       '
'       CODPARAM$(1) = "TRANSPOR": DOCPARAM$(1) = TRIM$(Mid$(TRXP$, 1, 32))
'       CODPARAM$(2) = "NUBULTOS": DOCPARAM$(2) = TRIM$(FINFACTURA.Text2(2).TEXT)
'       CODPARAM$(3) = "CUIT-TRA": DOCPARAM$(3) = TRIM$(Mid$(TRXP$, 65, 16))
'       CODPARAM$(4) = "VALO-DEC": DOCPARAM$(4) = TRIM$(FINFACTURA.Text2(3).TEXT)
'       CODPARAM$(5) = "COND-TRA": DOCPARAM$(5) = TRIM$(Mid$(TRXP$, 97, 32))
'       CODPARAM$(6) = "DOCUCOND": DOCPARAM$(6) = TRIM$(Mid$(TRXP$, 129, 24))
'       CODPARAM$(7) = "DEPOSAL ": 'DOCPARAM$(7) = TRIM$(DEPOSAL$)
'       CODPARAM$(8) = "NRO-REMI": DOCPARAM$(8) = TRIM$(NroRemi$)
'       CODPARAM$(9) = "NRO-OCOM":  DOCPARAM$(9) = NUOCOMPRA$
'       CODPARAM$(10) = "PATE-TRA":  DOCPARAM$(10) = TRIM$(Mid$(TRXP$, 81, 16))
'       'CODPARAM$(11) = "DOMI-TRA":  DOCPARAM$(11) = LENTREGA(0).TEXT
'       CODPARAM$(12) = "LOCA-TRA":  DOCPARAM$(12) = TRIM$(Mid$(TRXP$, 159, 32))
'       'CODPARAM$(13) = "COND-PAG": DOCPARAM$(13) = Eco21(1).TEXT
'       CODPARAM$(14) = "NRO-AENT":  DOCPARAM$(14) = NROPEDIDO$
'       'CODPARAM$(15) = "DESTINO": DOCPARAM$(15) = LENTREGA(1).TEXT
'       'CODPARAM$(16) = "TRANSPOR": DOCPARAM$(16) = LENTREGA(2).TEXT
'       CAPARDOC% = 16
'       '
'       CODTABU1$(1) = "COD-MAT"
'       CODTABU1$(2) = "DES-MAT"
'       CODTABU1$(3) = "CANTIDAD"
'       CODTABU1$(4) = "ORD-ITEM"
'       CACOLTAB1% = 4
'       '
'       DESTIDOC% = NC%
'       CAITAB1% = 0
'       '
'       If FINFACTURA.Text3 <> "" Then
'         Call FRATEXTO(FINFACTURA.Text3, 40)
'         For KKII% = 1 To CARETEX%
'           CAITAB1% = CAITAB1% + 1
'           TETABU1$(1, CAITAB1%) = ""
'           TETABU1$(2, CAITAB1%) = RETEX$(KKII%)
'           TETABU1$(3, CAITAB1%) = ""
'           TETABU1$(4, CAITAB1%) = Str$(CAITAB1%)
'         Next KKII%
'         FINFACTURA.Text3 = ""
'         FINFACTURA.Frame5.Visible = False
'         GoTo 80
'       End If
'       '
'       For U& = 1 To ULTREG&("!CARDEFAC")
'          XXX$ = REGLEIDO$("!CARDEFAC", U&)
'          CANS$ = Str$(CVS(Left$(XXX$, 4)))       ' cantidad
'          If XVALO(CANS$) >= 0.05 Then
'             If TRIM$(Mid$(XXX$, 387, 16)) = "" Then
'                CAITAB1% = CAITAB1% + 1
'                CDXXS$ = CODEXT$(CVI(Mid$(XXX$, 5, 2)))     ' codigo
'                DDTXT$ = Mid$(XXX$, 7, 64)                  ' descripcion
'                '
'                TETABU1$(1, CAITAB1%) = CDXXS$
'                TETABU1$(2, CAITAB1%) = DDTXT$
'                TETABU1$(3, CAITAB1%) = CANS$
'                TETABU1$(4, CAITAB1%) = Str$(CAITAB1%)
'                '
'             End If
'          End If
'       Next U&
'       '
'80     CACOPDOC% = XVALO(FINFACTURA.Text2(1).TEXT)
'       PRINTERPORT$ = PORTUREM$
'       Call PRINTDOC("FOREMIFA")   ' Remito de Facturacion
'       '
'    End If
'    '
'End Sub

Private Sub Total_Change(Index As Integer)
   '
   If Index <= 2 Then
     Call CALTOFAC
   End If
   Call CALTOFA2
   '
End Sub

'
Private Sub Total_GotFocus(Index As Integer)
   '
   If VALICLIE% < 1 Then GoTo 99
   If Index < 3 Or Index > 4 Then GoTo 98
   If TIFADOC% < 3 Then GoTo 98
   Exit Sub
   '
98 On Error Resume Next
   PORDESCU.SetFocus
   On Error GoTo 0
   Exit Sub
   '
99 On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub
'
Sub GRADEUFA(TIPO%, VA%, NRO, NC%, FFAC%, COPAGI%, TOTFAC#)
   '
   Dim FVI%(6), PLAII%(6), IM#(6), FEX(3)
   '
   LU$ = LUDAT$
   IORIG# = TOTFAC#
   '
   If TIPO% = 0 Then
     If IORIG# < 0 Then
       TIPO% = 7
     End If
   End If
   '
   If TIPO% > 3 Then
     NUPAG% = 1
     FVI%(1) = FFAC%
     IM#(1) = TOTFAC#     ' IORIG#
     GoTo 41
   End If
   '
   ' busca condicion de pago segun COPAGI%
20 Call BUCONPA(LU$, COPAGI%, CoPag$, FREF%, PLA$)
   '
   ' determina vencimientos segun total factura y condicion de pago
30 FF = FFAC%: If FF < 1 Then FF = FECHANUM("01/01/80")
   FFS$ = FECHATEX$(FF)
   PLA$ = Left$(PLA$ + String$(6, 0), 6)

   If (FREF% < 1 Or FREF% > 4) Then
        Call MENSERR(24, "Mal Fecha Referencia en Cond.Pago Nro." + Str$(COPAGI%))
           FREF% = 1
   End If

   If FREF% = 2 Then FREF% = 1   ' AUN NO FUNCIONA SEMANA DE FACTURA
   If FREF% <> 3 Then GoTo 35
   If XVALO(Left$(FFS$, 2)) <= 15 Then FEBAS$ = "15" + Mid$(FFS$, 3)
   If XVALO(Left$(FFS$, 2)) > 15 Then FEBAS$ = "31" + Mid$(FFS$, 3)
35 If FREF% = 1 Then FEBAS$ = FFS$
   If FREF% = 4 Then FEBAS$ = "31" + Mid$(FFS$, 3)
   '
   For I% = 1 To 6
     FVI%(I%) = 0
   Next I%
   '
   FEFA = FECHANUM(FEBAS$)
   For I% = 1 To 6
     PLAI% = Asc(Mid$(PLA$, I%, 1))
     PLAII%(I%) = PLAI%
     If I% > 1 Then
       If (PLAII%(I%) < PLAII%(I% - 1) Or PLAII%(I%) <= 0) Then
         GoTo 31
       End If
     End If
     '
     FEXU = FEFA + PLAI%
     While DIENTRE(FEFA, FEXU) < PLAI%
       FEXU = DIPOST(FEXU)
     Wend
     '
     FVI%(I%) = FEXU
     NUPAG% = I%
   Next I%

31 FVIJ% = FVI%(NUPAG%)
   If NUPAG% < 2 Then FVIJ% = FFAC%
   '
   ' determina importes segun numero de pagos
40 If NUPAG% < 1 Then NUPAG% = 1
   IM#(1) = (Int(100 * IORIG# / NUPAG% + 0.5)) / 100

   For H% = 2 To NUPAG% - 1
     IM#(H%) = IM#(1)
   Next H%
   IM#(NUPAG%) = IORIG#

   For H% = 1 To NUPAG% - 1
     IM#(NUPAG%) = IM#(NUPAG%) - IM#(H%)
   Next H%
   '
   ICANCE# = IORIG# - TOTFAC#
   For H% = 1 To NUPAG%
     If ICANCE# > 0 Then
       If ICANCE# >= IM#(H%) Then
           ICANCE# = ICANCE# - IM#(H%)
           IM#(H%) = 0
         Else
           IM#(H%) = IM#(H%) - ICANCE#
           ICANCE# = 0
       End If
     End If
   Next H%
   '
41 NuorVen% = 0
   SIGX% = 1: If TIPO% > 3 Then SIGX% = (-1)
   For I% = 1 To NUPAG%
     If Abs(IM#(I%)) >= 0.005 Then
       FechEmi = CVDAT(FFAC%)
       IBruTot# = IORIG#
       NuorVen% = NuorVen% + 1
       FeVenc = CVDAT(FVI%(I%))
       ImVenc# = (Abs(IM#(I%))) * SIGX%
       Call GRADEUPEN
     End If
   Next I%
   '
End Sub
'
Sub GRAMOCABE()
   '
   FEMI% = FECHANUM(Text9)
   NCLIE% = XVALO(Text1)
   TIPO% = TIFADOC%
   VA% = VARIDOC%
   NRO& = XVALO(Text11.TEXT)
   NUDOCU$ = FORMDOC$(NRO&, TIPO%, VA%, SUCUVEN%)
   '
   SUCUVEN% = 0
   On Error Resume Next
   SUCUVEN% = XVALO(Label12)
   On Error GoTo 0
   NUDOCU$ = FORMDOC$(NRO&, TIPO%, VA%, SUCUVEN%)
   '
   TOTAPESOS# = TOTAFAC#(1, 8)
   TOTADOLAR# = TOTAPESOS# / TICAMBX#
   '
   RAPLI$ = REGBLAN$("CARBIMON")
   ' documento cabeza
   Call REPLA(RAPLI$, MKI$(FEMI%), 1, 2)
   Call REPLA(RAPLI$, MKI$(NCLIE%), 3, 2)
   Call REPLA(RAPLI$, Chr$(TIPO%), 5, 1)
   Call REPLA(RAPLI$, MKI$(VA%), 6, 1)
   Call REPLA(RAPLI$, MKS$(NRO), 7, 4)
   Call REPLA(RAPLI$, MKI$(SUCUVEN%), 11, 2)
   Call REPLA(RAPLI$, NUDOCU$, 13, 18)
   Call REPLA(RAPLI$, MKI$(MONEMI%), 31, 2)
   '
   ' documento aplicado
   Call REPLA(RAPLI$, MKI$(FEMI%), 33, 2)
   Call REPLA(RAPLI$, MKI$(NCLIE%), 35, 2)
   Call REPLA(RAPLI$, Chr$(TIPO%), 37, 1)
   Call REPLA(RAPLI$, MKI$(VA%), 38, 1)
   Call REPLA(RAPLI$, MKS$(NRO), 39, 4)
   Call REPLA(RAPLI$, MKI$(SUCUVEN%), 43, 2)
   Call REPLA(RAPLI$, NUDOCU$, 45, 18)
   Call REPLA(RAPLI$, MKI$(MONEMI%), 63, 2)
   '
   'importe en pesos
   Call REPLA(RAPLI$, MKD$(TOTAPESOS#), 65, 8)
   Call REPLA(RAPLI$, MKD$(TICAMBX#), 89, 8)
   '
   'importe en dolares
   Call REPLA(RAPLI$, MKD$(TOTADOLAR#), 97, 8)
   '
   Call REGAPP("CARBIMON", RAPLI$)
   Call CIERRARCH("CARBIMON")
   '
End Sub
'


Sub GRAREMITO(OKX%)
   '
   Screen.MousePointer = 11
   If VALICLIE% < 1 Then
      MERRO$ = "Número de Cliente no Válido"
      GoTo 98
   End If
   NCLIE% = XVALO(Text1.TEXT)
   '
   Call BLOQARCH("MOVISTO")
   Call BLOQARCH("CONSICLI")
   '
   AINDI% = 0
   CMO$ = TRIM$(CONTROL$("", "CODREFAC"))
   If Len(CMO$) <> 2 Then
     CMO$ = "RT"
   End If
   '
   REMI$ = "RT." + TRIM$(Text10.TEXT)
   NROO& = XVALO(Mid$(REMI$, 4))
   FF% = FECHANUM(Text9.TEXT)
   DEPOSALI% = DEPOEXPE%
   '
   PREFIREM$ = CONTROL$("", "PREFIREM")
   If PREFIREM$ = "" Then PREFIREM$ = "0001"
   NroRemi$ = TRIM$(Mid$(REMI$, 4))
   While Len(NroRemi$) < 8: NroRemi$ = "0" + NroRemi$: Wend
   NroRemi$ = CMO$ + "." + PREFIREM$ + "-" + NroRemi$
   '
   HO% = HOY(HOS$)
   HP = 100 * XVALO(Left$(Time$, 2)) + XVALO(Mid$(Time$, 4, 2))
   User% = USNBR% Mod 100
   CTRXX$ = MKI$(HO%) + MKI$(Int(HP)) + Chr$(User%) + String$(2, 0)
   '
   RECONSI$ = REGBLAN$("CONSICLI")
   Call REPLA(RECONSI$, String$(12, 0), 109, 12)
   Call REPLA(RECONSI$, MKI$(NCLIE%), 1, 2)
   Call REPLA(RECONSI$, MKI$(FF%), 3, 2)
   Call REPLA(RECONSI$, NroRemi$, 5, 18)
   Call REPLA(RECONSI$, String$(12, 0), 109, 12)
   Call REPLA(RECONSI$, CTRXX$, 122, 7)
   '
   REFE$ = RASOCLI$(NCLIE%)
   '
   YXI% = 0
   FIN& = ULTREG&("!CARDEFAC")
   For I& = 1 To FIN&
     Call TRACE(20, I&, FIN&)
     X$ = REGLEIDO$("!CARDEFAC", I&)
     CANTAA = CVS(Left$(X$, 4))
     If CANTAA > 0 Then
       CI% = CVI(Mid$(X$, 5, 2))
       If CI% > 0 Then
         NPEDID& = XVALO(Mid$(X$, 371, 8))
         COD$ = CODEXT$(CI%)
         PRE# = CDbl(CVS(Mid$(X$, 429, 4)))
         CMOV$ = "RT"
         If Mid$(X$, 461, 1) = "R" Then CMOV$ = "*R" ' REPARACION
         DOPRI$ = REMI$
         DOSEC$ = ""
         LOTE$ = ""
         SIG% = (-1)
         '
         'Call MOVISTO(FF%, CI%, LOTE$, CMOV$, DOPRI$, DOSEC$, REFE$, PRE#, "$", NCLIE%, DEPOSALI%, SIG% * CANTAA)
         '
         YXI% = YXI% + 1
         ZZ$ = RECONSI$
         Call REPLA(ZZ$, TRIM$(REMI$) + "-" + TRIM$(Str$(YXI%)), 23, 12)
         Call REPLA(ZZ$, MKI$(CI%), 35, 2)
         Call REPLA(ZZ$, MKD$(CDbl(Abs(CANTAA))), 37, 8)
         Call REPLA(ZZ$, MKS$(NPEDID&), 109, 4)
         Call REPLA(ZZ$, MKD$(PRE#), 113, 8)
         Call REPLA(ZZ$, Chr$(MONEMI%), 121, 1)
         Call REGAPP("CONSICLI", ZZ$)
         '
       End If
     End If
   Next I&
   '
   Call CIERRARCH("MOVISTO")
   Call CIERRARCH("CONSICLI")
   Screen.MousePointer = 1
   OKX% = 1
   Exit Sub
   '
   '
98 Call CIERRARCH("MOVISTO")
   Call CIERRARCH("CONSICLI")
   Screen.MousePointer = 1
   Call MENSERR(24, MERRO$)
   '
End Sub
'
Private Sub Total_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 13 Then
     If Index = 0 Then
       On Error Resume Next
       PORDESCU.SetFocus
       On Error GoTo 0
      Else
       On Error Resume Next
       Command14.SetFocus
       On Error GoTo 0
     End If
   End If
End Sub

Private Sub Total_LostFocus(Index As Integer)
   If Index >= 2 Then
     TTT1 = XVALO(Total(Index).TEXT)
     Total(Index) = CSTRING$(MKS$(TTT1), 3, 12, 2, 2)
   End If
End Sub

'
Sub CARLISFAC()
    List1.Clear
   FIN& = ULTREG&("!CARDEFAC")
   For U& = 1 To FIN&
      Y$ = REGLEIDO$("!CARDEFAC", U&)
      TTXX$ = Space$(128)
      CI% = CVI(Mid$(Y$, 5, 2))
      PRECIO = CVS(Mid$(Y$, 429, 4))
      PRELI = PRELISTA#(ARCHIPRE$, CI%)
      PORDESCUEN = 0: If PRELI >= 0.0001 Then PORDESCUEN = 100 * (PRELI - PRECIO) / PRELI
      MAREPA$ = UCase$(Mid$(Y$, 461, 1)): If MAREPA$ <> "R" Then MAREPA$ = " "
      NPEDID& = XVALO(Mid$(Y$, 371, 8))
      CAN = CVS(Mid$(Y$, 1, 4))
      '
      Call REPLA(TTXX$, CSTRING$(MKS$(CAN), 3, 6, 0, 2), 1, 6)
      Call REPLA(TTXX$, CODEXT$(CI%), 8, 15)
      Call REPLA(TTXX$, Mid$(Y$, 7, 64), 24, 26)
      PRELI = CVS(Mid$(Y$, 421, 4))
      Call REPLA(TTXX$, CSTRING$(MKS$(PRELI), 3, 8, 2, 2), 50, 8)
      PORDESCUIN = CVS(Mid$(Y$, 425, 4))
      Call REPLA(TTXX$, CSTRING$(MKS$(PORDESCUIN), 3, 5, 1, 2), 59, 5)
      PREUNI = CVS(Mid$(Y$, 429, 4))
      Call REPLA(TTXX$, CSTRING$(MKS$(PREUNI), 3, 8, 2, 2), 64, 8)
      TOT# = CVD(Mid$(Y$, 433, 8))
      Call REPLA(TTXX$, CSTRING$(MKD$(TOT#), 3, 10, 2, 2), 74, 10)
      Call REPLA(TTXX$, Mid$(Y$, 461, 2), 86, 2) ' MARCA REPARACION
      List1.AddItem TTXX$
      '
   Next U&
   List1.Refresh
   '
End Sub
'
Sub CALTOFAC()
   '
   Static CLIA$, RNI%, APLIVA%
   If Text1.TEXT <> CLIA$ Then
     CLIA$ = Text1.TEXT: NCLIE% = XVALO(CLIA$)
     RNI% = 0: APLIVA% = 1
     If NCLIE% > 0 Then
       PPVI% = PIVACLI%(NCLIE%)
       If PPVI% = 2 Then RNI% = 1
       If PPVI% = 5 Or PPVI% = 7 Then APLIVA% = 0
     End If
   End If
   '
   TASA1 = XVALO(PORIVA(0).Caption)
   TASA2 = XVALO(PORIVA(1).Caption)
   TASA3 = XVALO(PORIVA(2).Caption)
   '
   TTTX# = XVALO(Total(2))
   BASIVA(0).Caption = TRIM$(CSTRING$(MKD$(TTTX#), 3, 12, 2, 2))
   Total(3).TEXT = CSTRING$(MKD$(XVALO(BASIVA(0).Caption) * XVALO(PORIVA(0).Caption) * APLIVA% / 100), 3, 12, 2, 2)
   '
   TTTX# = (XVALO(Total(2).TEXT) - XVALO(BASIVA(0).Caption))
   BASIVA(1).Caption = TRIM$(CSTRING$(MKD$(TTTX#), 3, 12, 2, 2))
   Total(4).TEXT = CSTRING$(MKD$(XVALO(BASIVA(1).Caption) * XVALO(PORIVA(1).Caption) * APLIVA% / 100), 3, 12, 2, 2)
   '
   Call CALTOFA2
   '
End Sub
'
Sub CALTOFA2()
   '
   Static CLIA$, RNI%
   If Text1.TEXT <> CLIA$ Then
     CLIA$ = Text1.TEXT: NCLIE% = XVALO(CLIA$)
     RNI% = 0
     If NCLIE% > 0 Then
       PPVI% = PIVACLI%(NCLIE%)
       If PPVI% = 2 Then RNI% = 1
     End If
   End If
   '
   TTTX# = (XVALO(Total(3).TEXT) + XVALO(Total(4).TEXT)) * RNI%
   '
   TTTX# = 0: For KKI% = 2 To 5 'ANTES DE 3 A 5
                 TTTX# = TTTX# + XVALO(Total(KKI%).TEXT)
              Next KKI%
              TTTX# = TTTX# + XVALO(Text15) 'PERCEP IVA
              TTTX# = TTTX# + XVALO(Text17) 'PERCEP I.BRUTOS
   Total(6).TEXT = CSTRING$(MKD$(TTTX#), 3, 12, 2, 2)
   '

End Sub
Sub CALTOFA1()
    '
    
    TTTX# = TTTX# + XVALO(Total(7)) 'TOTAL + NETO
    Total(6).TEXT = CSTRING$(MKD$(TTTX#), 3, 12, 2, 2)
    '
End Sub
Sub LEETASIVA()
   '
   URETA& = ULTREG&("GRUPIVA")
   If URETA& < 2 Then
     Call MENSERR(24, "Falta Definir Grupos de I.V.A.")
     Call CIERRARCH("*.*")
     Call FINAL("")
   End If
   '
   For U& = 1 To 2
      X$ = REGLEIDO$("GRUPIVA", U&)
      TAIVA(U& - 1) = CVS(Mid$(X$, 13, 4))
      PORIVA(U& - 1) = TRIM$(CSTRING$(MKS$(TAIVA(U& - 1)), 3, 6, 2, 2))
   Next U&
   PORIVA(2) = "50.00" ' responsable no inscripto
   Call CIERRARCH("GRUPIVA")
   ALIIVA = ALIVA
   '
End Sub

Sub GRACONSI()
   '
   NUFACLAR$ = CoComLar$
   NUFACORT$ = CoComCor$
   '
   NC1% = XVALO(Text1.TEXT)
   Call APERBASDAT("STOCKS")
   
   FIN& = ULTREG&("!CARDEFAC")
   '
   For I& = 1 To FIN&
     Call TRACE(20, I&, FIN&)
     X$ = REGLEIDO$("!CARDEFAC", I&)
     CANTAA = CVS(Left$(X$, 4))
     If Abs(CANTAA) > 0 Then
       CI% = CVI(Mid$(X$, 5, 2))
       If CI% > 0 Then
         DOPRIC$ = TRIM$(Mid$(X$, 387, 16))
         NORIT% = CVI(Mid$(X$, 503, 2))
         PREUNI = CVS(Mid$(X$, 429, 4))
         NPEDCAR& = XVALO(Mid$(X$, 371, 16))
         '
         SQLX$ = "DoPriCor = '" & DOPRIC$ & "' "
         SQLX$ = SQLX$ + "AND Cod_Inte = " & CI% & " "
         SQLX$ = SQLX$ + "AND Nume_Clie = " & NC1% & " "
         SQLX$ = SQLX$ + "AND NuorItem = " & NORIT% & " "
         '
         With Movisto
           .FindFirst SQLX$
25         If .NoMatch = False Then
             NPEDID& = 0
             If IsNull(!NuPedCli) = False Then NPEDID& = !NuPedCli
             If NPEDID& <> NPEDCAR& Then
               .FindNext SQLX$
               GoTo 25
             End If
             '
             .Edit
             !DoSecCor = NUFACORT$
             !DoSecLar = NUFACLAR$
             !VALOUNIT = PREUNI
             If CANTAA > 0 Then
                 !CANTSALID = CANTAA
                 !CANTINGRE = 0
               Else
                 !CANTSALID = 0
                 !CANTINGRE = Abs(CANTAA)
             End If
             .Update
           End If
         End With
       End If
     End If
   Next I&
   '
End Sub

Sub AJUPREVE()
   '
   NC1% = XVALO(Text1.TEXT)
   Call APERBASDAT("STOCKS")
   
   FIN& = ULTREG&("!DIFPRECO")
   '
   For I& = 1 To FIN&
     Call TRACE(20, I&, FIN&)
     X$ = REGLEIDO$("!DIFPRECO", I&)
     PRENUEVO# = CVD(Mid$(X$, 45, 8))
     If PRENUEVO# > 0.00005 Then
       CI% = CVI(Mid$(X$, 1, 2))
       If CI% > 0 Then
         DOPRIL$ = TRIM$(Mid$(X$, 81, 18))
         NORIT% = CVI(Mid$(X$, 117, 2))
         PREUNI = CVD(Mid$(X$, 45, 8))
         '
         SQLX$ = "DoPriLar = '" & DOPRIL$ & "' "
         SQLX$ = SQLX$ + "AND Cod_Inte = " & CI% & " "
         SQLX$ = SQLX$ + "AND Nume_Clie = " & NC1% & " "
         SQLX$ = SQLX$ + "AND NuorItem = " & NORIT% & " "
         '
         With Movisto
           .FindFirst SQLX$
25         If .NoMatch = False Then
             .Edit
             !VALOUNIT = PREUNI
             .Update
           End If
         End With
       End If
     End If
   Next I&
   '
End Sub

Sub CARTAMONED()
   If DERACCE%("TICAMBIO", "Ajustar Tipo de Cambio Central") < 2 Then
     Exit Sub
   End If
   '
   If ULTREG&("TAMONED") > 6 Then
      Call SETULTREG("TAMONED", 6)
      Call COMUNI("Tabla de Monedas Truncada\al Máximo de 6 Registros")
   End If
   '
   Call TRALOCAL("TAMONED", "")
   VTB% = VENTABU%("TAMONED/NC/TITUPAN=Monedas y Tipo de Cambio")
   If VTB% >= 0 Then
      For U& = 1 To ULTREG&("!TAMONED")
         X$ = REGLEIDO$("!TAMONED", U&)
         Call REPLA(X$, Chr$(U&), 1, 1)
         Call GRAREG("TAMONED", X$, U&)
      Next U&
      FIN& = ULTREG&("!TAMONED")
      Call SETULTREG("TAMONED", FIN&)
   End If
   Call CIERRARCH("TAMONED")
   TICAMBIS = FORMATNUM(TICAMONE(2), "F9.4")
End Sub

Sub SHOWPEDIPEN()
     '
     Static UFEPEDET#
     If CDbl(FileDateTime(LUDAT$ + "PEDDETA.DAT")) <> UFEPEDET# Then
        UFEPEDET# = CDbl(FileDateTime(LUDAT$ + "PEDDETA.DAT"))
        '
        JJ& = 0
        UREPRE& = ULTREG&("PEDDETA")
        Screen.MousePointer = 11
        JJ& = 0
        NROPEDA& = 0: CAPED% = 0
        FINF = 999999: FSUP = 0
        For REPRESU& = 1 To UREPRE&
          Call TRACE(20, REPRESU&, UREPRE&)
          X$ = REGLEIDO$("PEDDETA", REPRESU&)
          If Asc(Mid$(X$, 27, 1)) <> 9 Then ' no esta anulado
            CAN = CVS(Mid$(X$, 79, 4))
            CADESPACH = CVS(Mid$(X$, 95, 4))
            CAPENDENT = CAN - CADESPACH
            If CAPENDENT > 0.5 Then
              NRODES& = CVS(Left$(X$, 4))
              If NRODES& <> NRODESA& Then
                PRIREDETA& = REPRESU&
                NRODESA& = NRODES&
                TTXX$ = MKS$(NRODES&)
                NOCO$ = TRIM$(Mid$(X$, 9, 16))
                If NOCO$ <> "" Then
                  TTXX$ = TTXX$ + "O.C." + NOCO$ + " - "
                End If
                FEX = CVI(Mid$(X$, 5, 2))
                TTXX$ = TTXX$ + FECHATEX$(FEX) + " - "
                NCLIE% = CVI(Mid$(X$, 7, 2))
                TTXX$ = TTXX$ + TRIM$(RASOCLI$(NCLIE%)) + Space$(52)
                REGENCA& = CVS(Mid$(X$, 109, 4))
                JJ& = JJ& + 1
                Call GRAREG("INDIPEP", Left$(TTXX$, 52) + MKS$(REGENCA&), JJ&)
              End If
            End If
          End If
        Next REPRESU&
        Call SETULTREG("INDIPEP", JJ&)
        Call CIERRARCH("INDIPEP")
        '
        Screen.MousePointer = 1
        YAINDI% = 1
     End If
     '
2112 If ULTREG&("INDIPEP") < 1 Then
        Call COMUNI("No hay O/V's Pendientes\que Puedan Consultarse")
        GoTo 9999
     End If
     '
2110 X$ = REGSEL$("INDIPEP/Ordenes de Venta Pendientes de Entrega")
     If Len(X$) <= 4 Then
        GoTo 9999
     End If
     '
     NPX& = CVS(Left$(X$, 4))
     If NPX& < 1 Then GoTo 9999
     '
     TIDOCUM$ = "Orden de Despacho"
     NUDOCU$ = TRIM$(Str$(NPX&))
     While Len(NUDOCU$) < 6: NUDOCU$ = "0" + NUDOCU$: Wend
     NUDOCU$ = NUDOCU$ + " "
     Screen.MousePointer = 11
     '
     For UI& = ULTREG&("PDOCLST") To 1 Step -1
       XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
       If InStr(XX$, TIDOCUM$) > 4 Then
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
     Screen.MousePointer = 1
     Call COMUNI("Documento no Encontrado")
     Exit Sub
     '
5    For KKI% = 0 To 5
       EDITFORM.Picture1(KKI%).Cls
       EDITFORM.Picture1(KKI%).DrawWidth = 3
       EDITFORM.Picture1(KKI%).DrawStyle = 0
       EDITFORM.Picture1(KKI%).FillStyle = 1
     Next KKI%
     '
     REDOCU$ = REGACT$("PDOCLST")
     URE& = ULTREG&("PDOCSPL")
     LI& = 0: LS& = URE&
6    E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 7
     L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 7
     DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
     If DCC& < DOCUNU& Then LI& = L&: GoTo 6
     If DCC& > DOCUNU& Then LS& = L&: GoTo 6
     GoTo 8
     '
7    Screen.MousePointer = 1
     Call MENSERR(24, "Documento no Registrado")
     GoTo 9999
     '
8    TPSP$ = ""
     Screen.MousePointer = 11
     While L& > 1
       DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
       If DCC& < DOCUNU& Then GoTo 9
       L& = L& - 1
     Wend
     '
9    CAPAGINAS% = 0
     For KKL& = L& To URE&
       TBUF$ = REGLEIDO$("PDOCSPL", KKL&)
       If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 95
       If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
         TPSP$ = TPSP$ + Mid$(TBUF$, 5)
       End If
     Next KKL&
     '
95   SPOOLSTRING$ = TPSP$
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
10   PPXX% = InStr(TPSP$, Chr$(255))
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
     'EDITFORM.Command5.Enabled = False
     EDITFORM.Show 1
     '
     GoTo 2110
     '
9999 End Sub
'
Sub REGISNODE()
    'REGISTRAR NOTA DE DEBITO DEL CLIENTE
    '
    TIPOCOMP$ = "ND" 'CI(CREDITO INTERNO) ES NOTA DE DEBITO DEL CLIENTE
    OPCICOMP$ = "A" 'O "B" 'DEPENDIENDO DE LA POSICION DEL IVA DEL CLIENTE?
    '
    NUMEROCOMP$ = XVALO(Mid$(Text11, 1))
    NUMEROTALO$ = XVALO(Mid$(Text6, 1))
    'CORTO
    NCCMP$ = TIPOCOMP$ & "-" & NUMEROCOMP & "-" & OPCICOMP$
    'LARGO
    NLCMP$ = TIPOCOMP$ & "-" & TRIM$(Text6) & "-" & TRIM$(Text11) & "-" & OPCICOMP$
    '
    TOTDOCUM# = XVALO(Total(6))
    '
    NC% = XVALO(Text1)
    CoComCor = NCCMP$
    CoComLar = NLCMP$
    NumeTal = NUMEROTALO$
    NumeCom = NUMEROCOMP$
    Referen = Mid$(TRIM$(DescriTxt), 1, 48)
    FechEmi = Text9
    FeContab = Text9
    FeAcredi = Text9
    TiMovi$ = "FVEN"
    OpMovi = "CI"   ' CREDITO INTERNO
    VaMovi = OPCICOMP$
    NumeProv = 0
    NumeClie = GRUDIVCLI%(NC%)
    CuentaCo% = CUECOCLI%(NC%)
    CodiEmp% = 0
    CodiSuc% = 0
    Referen$ = RASOCLI$(NC%)
    BasImpo# = 0
    ImpoDeb# = TOTDOCUM# * (-1)
    Descrip$ = "N/Debito Cliente " + NLCMP$
    Observa$ = RASOCLI$(NC%)
    StatInit = 0
    Call GRACABAN
    '
    'graba venta neta - Verificar que QQQ sea la IC de combinacion Cliente-Tipo de Comprobante
   TOTANETX# = XVALO(Total(2))
   If TOTANETX# >= 0.005 Then
     QQ1% = CUECOINT%(Text16)
     CuentaCo% = QQ1%   ' TOMA LA QUE APARECE EN PANTALLA
     TiMovi$ = "NDCLI"
     OpMovi$ = "CI"
     VaMovi$ = OPCICOMP$
     StatInit = 0
     ImpoDeb = TOTANETX#
     Call GRACABAN
   End If
   '
   'graba iva inscripto
   For KKI% = 3 To 4
     TOTANETX# = XVALO(Total(KKI%))
     If TOTANETX# >= 0.005 Then
       CUEIVA$ = CONTROL$("SVD451S", "CUEIVA")
       CuentaCo% = CUECOINT%(CUEIVA$)
       TiMovi$ = "IVA"
       OpMovi$ = "DeFi"
       VaMovi$ = PORIVA(KKI% - 3) + " %"
       Descrip$ = "Iva Debito Fiscal " + NCCMP$
       StatInit = 0
       ImpoDeb = TOTANETX#
       BasImpo# = XVALO(BASIVA(KKI% - 3))
       Call GRACABAN
     End If
   Next KKI%
   BasImpo# = 0
   '
   'graba iva NO inscripto
   TOTANETX# = XVALO(Total(5))
   If TOTANETX# >= 0.005 Then
     CUEIVA$ = CONTROL$("SVD451S", "CUEADU")
     CuentaCo% = CUECOINT%(CUEIVA$)
     TiMovi$ = "IVA"
     OpMovi$ = "DeFi"
     VaMovi$ = "RNI"
     Descrip$ = "IVA Resp.No Inscripto " + NCCMP$
     StatInit = 0
     ImpoDeb = TOTANETX#
     BasImpo# = XVALO(BASIVA(2))
     Call GRACABAN
   End If
   BasImpo# = 0
   '
   'graba percepcion I.Brutos
   TOTANETX# = XVALO(Text15)
   If TOTANETX# >= 0.005 Then
     CUEIVA$ = CONTROL$("PERIBRU", "CUECONTA")
     CuentaCo% = CUECOINT%(CUEIVA$)
     TiMovi$ = "IBRU"
     OpMovi$ = "IB-B"
     VaMovi$ = "PERC"
     Descrip$ = "Perc.IB.Bs.As. " + NCCMP$
     StatInit = 0
     ImpoDeb = TOTANETX#
     BasImpo# = XVALO(Total(2))
     Call GRACABAN
   End If
   BasImpo# = 0
   '
   'graba percepcion IVA
   TOTANETX# = XVALO(Text17)
   If TOTANETX# >= 0.005 Then
     CUEPER$ = CONTROL$("SVD451S", "CUERET")
     CuentaCo% = CUECOINT%(CUEPER$)
     TiMovi$ = "IVA"
     OpMovi$ = "DeFi"
     VaMovi$ = "PER-IVA"
     Descrip$ = "Percep.IVA " + NCCMP$
     StatInit = 0
     ImpoDeb = TOTANETX#
     BasImpo# = XVALO(Total(2))
     Call GRACABAN
   End If
   '
   'GRABA DEUDA PENDIENTE
   IBruTot# = TOTDOCUM#
   IIvaTot# = XVALO(Total(3)) + XVALO(Total(4))
   NuorVen% = 0
   FeVenc = FechEmi
   ImVenc# = (-1) * TOTDOCUM
   Call GRADEUPEN
   '
End Sub
