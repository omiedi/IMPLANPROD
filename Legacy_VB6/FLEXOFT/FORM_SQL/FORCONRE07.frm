VERSION 5.00
Begin VB.Form FORCONRE07 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "FORCONRE04"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11625
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   11625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   6045
      Left            =   0
      ScaleHeight     =   5985
      ScaleWidth      =   10980
      TabIndex        =   0
      Top             =   0
      Width           =   11040
      Begin VB.Frame Frame2 
         Caption         =   "Varios"
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
         Left            =   210
         TabIndex        =   68
         Top             =   5040
         Width           =   10620
         Begin VB.Label Label20 
            Caption         =   "Varios(0)"
            Height          =   225
            Left            =   7560
            TabIndex        =   71
            Top             =   315
            Width           =   645
         End
         Begin VB.Label Varios 
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
            Height          =   225
            Left            =   3885
            TabIndex        =   70
            Top             =   315
            Width           =   3420
         End
         Begin VB.Label Label19 
            Caption         =   "Depósito"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   105
            TabIndex        =   69
            Top             =   315
            Width           =   3570
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Transporte"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2430
         Left            =   210
         TabIndex        =   27
         Top             =   2415
         Width           =   10620
         Begin VB.Label Label2 
            Caption         =   "Trans(8)"
            Height          =   225
            Index           =   8
            Left            =   7560
            TabIndex        =   67
            Top             =   1995
            Width           =   645
         End
         Begin VB.Label Label2 
            Caption         =   "Trans(7)"
            Height          =   225
            Index           =   7
            Left            =   7560
            TabIndex        =   66
            Top             =   1785
            Width           =   645
         End
         Begin VB.Label Trans 
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
            Height          =   225
            Index           =   8
            Left            =   3885
            TabIndex        =   65
            Top             =   1995
            Width           =   3420
         End
         Begin VB.Label Trans 
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
            Height          =   225
            Index           =   7
            Left            =   3885
            TabIndex        =   64
            Top             =   1785
            Width           =   3420
         End
         Begin VB.Label Label18 
            Caption         =   "VALO-DEC"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   8610
            TabIndex        =   63
            Top             =   1785
            Width           =   1665
         End
         Begin VB.Label Label17 
            Caption         =   "Conductor"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   105
            TabIndex        =   62
            Top             =   1155
            Width           =   3675
         End
         Begin VB.Label Label16 
            Caption         =   "PATE-TRA"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   8610
            TabIndex        =   61
            Top             =   1575
            Width           =   1665
         End
         Begin VB.Label Label15 
            Caption         =   "DOCUCOND"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   8610
            TabIndex        =   60
            Top             =   1365
            Width           =   1665
         End
         Begin VB.Label Label14 
            Caption         =   "COND-TRA"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   8610
            TabIndex        =   59
            Top             =   1155
            Width           =   1665
         End
         Begin VB.Label Label13 
            Caption         =   "CUIT-TRA"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   8610
            TabIndex        =   58
            Top             =   945
            Width           =   1665
         End
         Begin VB.Label Label12 
            Caption         =   "Patente"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   105
            TabIndex        =   57
            Top             =   1575
            Width           =   3570
         End
         Begin VB.Label Label11 
            Caption         =   "Documento del Conductor"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   105
            TabIndex        =   56
            Top             =   1365
            Width           =   3675
         End
         Begin VB.Label Label10 
            Caption         =   "Cuit del Transportista"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   105
            TabIndex        =   55
            Top             =   945
            Width           =   3675
         End
         Begin VB.Label Label2 
            Caption         =   "Trans(6)"
            Height          =   225
            Index           =   6
            Left            =   7560
            TabIndex        =   50
            Top             =   1575
            Width           =   645
         End
         Begin VB.Label Trans 
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
            Height          =   225
            Index           =   6
            Left            =   3885
            TabIndex        =   49
            Top             =   1575
            Width           =   3420
         End
         Begin VB.Label Label3 
            Caption         =   "Valor Declarado"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   105
            TabIndex        =   48
            Top             =   1785
            Width           =   3570
         End
         Begin VB.Label Label2 
            Caption         =   "Trans(5)"
            Height          =   225
            Index           =   5
            Left            =   7560
            TabIndex        =   47
            Top             =   1365
            Width           =   645
         End
         Begin VB.Label Label2 
            Caption         =   "Trans(4)"
            Height          =   225
            Index           =   4
            Left            =   7560
            TabIndex        =   46
            Top             =   1155
            Width           =   645
         End
         Begin VB.Label Label2 
            Caption         =   "Trans(3)"
            Height          =   225
            Index           =   3
            Left            =   7560
            TabIndex        =   45
            Top             =   945
            Width           =   645
         End
         Begin VB.Label Label2 
            Caption         =   "Trans(2)"
            Height          =   225
            Index           =   2
            Left            =   7560
            TabIndex        =   44
            Top             =   735
            Width           =   645
         End
         Begin VB.Label Label2 
            Caption         =   "Trans(1)"
            Height          =   225
            Index           =   1
            Left            =   7560
            TabIndex        =   43
            Top             =   525
            Width           =   645
         End
         Begin VB.Label Label2 
            Caption         =   "Trans(0)"
            Height          =   225
            Index           =   0
            Left            =   7560
            TabIndex        =   42
            Top             =   315
            Width           =   645
         End
         Begin VB.Label Trans 
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
            Height          =   225
            Index           =   5
            Left            =   3885
            TabIndex        =   41
            Top             =   1365
            Width           =   3420
         End
         Begin VB.Label Trans 
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
            Height          =   225
            Index           =   4
            Left            =   3885
            TabIndex        =   40
            Top             =   1155
            Width           =   3420
         End
         Begin VB.Label Trans 
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
            Height          =   225
            Index           =   3
            Left            =   3885
            TabIndex        =   39
            Top             =   945
            Width           =   3420
         End
         Begin VB.Label Trans 
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
            Height          =   225
            Index           =   2
            Left            =   3885
            TabIndex        =   38
            Top             =   735
            Width           =   3420
         End
         Begin VB.Label Trans 
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
            Height          =   225
            Index           =   1
            Left            =   3885
            TabIndex        =   37
            Top             =   525
            Width           =   3420
         End
         Begin VB.Label Trans 
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
            Height          =   225
            Index           =   0
            Left            =   3885
            TabIndex        =   36
            Top             =   315
            Width           =   3420
         End
         Begin VB.Label Label99 
            Caption         =   "NUBULTOS"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   8610
            TabIndex        =   35
            Top             =   1995
            Width           =   1665
         End
         Begin VB.Label Label97 
            Caption         =   "Numero de Bultos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   105
            TabIndex        =   34
            Top             =   1995
            Width           =   3570
         End
         Begin VB.Label Label63 
            Caption         =   "DOMI-TRA"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   8610
            TabIndex        =   33
            Top             =   525
            Width           =   1665
         End
         Begin VB.Label Label61 
            Caption         =   "Domicilio Transportista"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   105
            TabIndex        =   32
            Top             =   525
            Width           =   3570
         End
         Begin VB.Label Label69 
            Caption         =   "TRANSPOR"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   8610
            TabIndex        =   31
            Top             =   315
            Width           =   1665
         End
         Begin VB.Label Label67 
            Caption         =   "Transportista"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   105
            TabIndex        =   30
            Top             =   315
            Width           =   3675
         End
         Begin VB.Label Label66 
            Caption         =   "DESTINO"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   8610
            TabIndex        =   29
            Top             =   735
            Width           =   1665
         End
         Begin VB.Label Label64 
            Caption         =   "Localidad"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   105
            TabIndex        =   28
            Top             =   750
            Width           =   3675
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Identificación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2115
         Left            =   210
         TabIndex        =   1
         Top             =   105
         Width           =   10620
         Begin VB.Label Label8 
            Caption         =   "NRO-AENT"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   8610
            TabIndex        =   54
            Top             =   1680
            Width           =   1665
         End
         Begin VB.Label Label124 
            Caption         =   "Ident(6)"
            Height          =   225
            Index           =   6
            Left            =   7560
            TabIndex        =   53
            Top             =   1680
            Width           =   1065
         End
         Begin VB.Label Ident 
            BorderStyle     =   1  'Fixed Single
            Height          =   225
            Index           =   6
            Left            =   3885
            TabIndex        =   52
            Top             =   1680
            Width           =   3480
         End
         Begin VB.Label Label5 
            Caption         =   "Número de Orden de Suministro"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   105
            TabIndex        =   51
            Top             =   1680
            Width           =   3675
         End
         Begin VB.Label Label9 
            Caption         =   "N-AUTENT"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   8610
            TabIndex        =   26
            Top             =   1470
            Width           =   1665
         End
         Begin VB.Label Label7 
            Caption         =   "Numero de Pedido"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   105
            TabIndex        =   25
            Top             =   1485
            Width           =   3675
         End
         Begin VB.Label Label6 
            Caption         =   "NRO-OCOM"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   8610
            TabIndex        =   24
            Top             =   1260
            Width           =   1665
         End
         Begin VB.Label Label4 
            Caption         =   "Numero de Orden de Compra"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   105
            TabIndex        =   23
            Top             =   1275
            Width           =   3675
         End
         Begin VB.Label Label124 
            Caption         =   "Ident(5)"
            Height          =   225
            Index           =   5
            Left            =   7560
            TabIndex        =   22
            Top             =   1470
            Width           =   1065
         End
         Begin VB.Label Label124 
            Caption         =   "Ident(4)"
            Height          =   225
            Index           =   4
            Left            =   7560
            TabIndex        =   21
            Top             =   1260
            Width           =   1065
         End
         Begin VB.Label Label124 
            Caption         =   "Ident(3)"
            Height          =   225
            Index           =   3
            Left            =   7560
            TabIndex        =   20
            Top             =   1050
            Width           =   1065
         End
         Begin VB.Label Label124 
            Caption         =   "Ident(2)"
            Height          =   225
            Index           =   2
            Left            =   7560
            TabIndex        =   19
            Top             =   840
            Width           =   1065
         End
         Begin VB.Label Label124 
            Caption         =   "Ident(1)"
            Height          =   225
            Index           =   1
            Left            =   7560
            TabIndex        =   18
            Top             =   630
            Width           =   1065
         End
         Begin VB.Label Label124 
            Caption         =   "Ident(0)"
            Height          =   225
            Index           =   0
            Left            =   7560
            TabIndex        =   17
            Top             =   420
            Width           =   1065
         End
         Begin VB.Label Ident 
            BorderStyle     =   1  'Fixed Single
            Height          =   225
            Index           =   5
            Left            =   3885
            TabIndex        =   16
            Top             =   1470
            Width           =   3480
         End
         Begin VB.Label Label150 
            Caption         =   "Numero de Remito Corto"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   105
            TabIndex        =   15
            Top             =   1065
            Width           =   3570
         End
         Begin VB.Label Ident 
            BorderStyle     =   1  'Fixed Single
            Height          =   225
            Index           =   4
            Left            =   3885
            TabIndex        =   14
            Top             =   1260
            Width           =   3480
         End
         Begin VB.Label Label147 
            Caption         =   "Numero de Remito Largo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   105
            TabIndex        =   13
            Top             =   855
            Width           =   3675
         End
         Begin VB.Label Label146 
            Caption         =   "NRO-REMI"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   8610
            TabIndex        =   12
            Top             =   840
            Width           =   1665
         End
         Begin VB.Label Ident 
            BorderStyle     =   1  'Fixed Single
            Height          =   225
            Index           =   3
            Left            =   3885
            TabIndex        =   11
            Top             =   1050
            Width           =   3480
         End
         Begin VB.Label Ident 
            BorderStyle     =   1  'Fixed Single
            Height          =   225
            Index           =   2
            Left            =   3885
            TabIndex        =   10
            Top             =   840
            Width           =   3480
         End
         Begin VB.Label Ident 
            BorderStyle     =   1  'Fixed Single
            Height          =   225
            Index           =   1
            Left            =   3885
            TabIndex        =   9
            Top             =   630
            Width           =   3480
         End
         Begin VB.Label Label126 
            Caption         =   "Fecha de Emisión"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   105
            TabIndex        =   8
            Top             =   645
            Width           =   3990
         End
         Begin VB.Label Ident 
            BorderStyle     =   1  'Fixed Single
            Height          =   225
            Index           =   0
            Left            =   3885
            TabIndex        =   3
            Top             =   420
            Width           =   3480
         End
         Begin VB.Label Label1 
            Caption         =   "Numero de Cliente"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   105
            TabIndex        =   2
            Top             =   435
            Width           =   3990
         End
      End
   End
   Begin VB.Label Label137 
      Caption         =   "NRO-REMI"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7665
      TabIndex        =   7
      Top             =   0
      Width           =   1665
   End
   Begin VB.Label Label134 
      Caption         =   "NRO-REMI"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7665
      TabIndex        =   6
      Top             =   0
      Width           =   1665
   End
   Begin VB.Label Label131 
      Caption         =   "NRO-REMI"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7665
      TabIndex        =   5
      Top             =   0
      Width           =   1665
   End
   Begin VB.Label Label128 
      Caption         =   "NRO-REMI"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7665
      TabIndex        =   4
      Top             =   0
      Width           =   1665
   End
End
Attribute VB_Name = "FORCONRE07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'\\\OT-06-0405-MH-20/09/06///
Private Sub Form_Load()
      '
       Call APLICACIONSKINS(Me)

'      Label50.Caption = FORFACCS.Label13.Caption 'Cond Iva
'      '
'      REMI$ = "RT." + TRIM$(FORFACCS.Text10.TEXT)
'      NroRemi$ = TRIM$(Mid$(REMI$, 4))
'      While Len(NroRemi$) < 8: NroRemi$ = "0" + NroRemi$: Wend
'      CMO$ = "RT"
'      PREFIREM$ = CONTROL$("", "PREFIREM")
'      If PREFIREM$ = "" Then PREFIREM$ = "0001"
'      NroRemi$ = CMO$ + "." + PREFIREM$ + "-" + NroRemi$
'      '
'      Label2.Caption = NroRemi$                   'nº remito
'      Label5.Caption = NUOCOMPRA$                 'nº Orden de Compra
'      Label8.Caption = NROPEDIDO$                 'nº Pedido
'      IMPO(0).Caption = FORFACCS.Total(0).TEXT    'Total Neto
'      Label14.Caption = FORFACCS.Eco21(2).TEXT    'Vendedor
'      Label17.Caption = FORFACCS.Eco21(1).TEXT    'Cond Pago
'      IMPO(4).Caption = FORFACCS.PORIVA(0).Caption 'Porc IVA Basica
'      IMPO(5).Caption = FORFACCS.Total(3).TEXT    'Impo IVA Basica
'      IMPO(6).Caption = FORFACCS.PORIVA(1).Caption 'Porc IVA Secun
'      IMPO(7).Caption = FORFACCS.Total(4).TEXT    'Impo IVA Secun
'      IMPO(2).Caption = FORFACCS.Total(1).TEXT    'Importe dto
'      IMPO(3).Caption = FORFACCS.Total(2).TEXT  'Importe Neto dto
'      IMPO(12).Caption = FORCACCS.Total(6).TEXT  'Importe Total
'      IMPO(1).Caption = TRIM$(FORFACCS.PORDESCU.TEXT) 'Porc Dto
'      IMPO(8).Caption = FORFACCS.Total(5).TEXT 'Importe iva RNI
'      '
'      TTTT# = XVALO(FORFACCS.Total(6).TEXT)
'      TOTFAC$ = NUMTEX$(TTTT#)
'      Call FRATEX(TOTFAC$, NRX1$, NRX2$, NRX3$, LOIMPLET%)
'      '
'      IMPO(14).Caption = NRX1$  'Importe en letras 1
'      IMPO(15).Caption = NRX2$  'Importe en letras 2
'      IMPO(16).Caption = NRX3$  'Importe en letras 3
'      '
'      IMPO(13).Caption = ""  'Total en dolares
'      Label62.Caption = FORFACCS.LENTREGA(0).TEXT  'Domic transportista
'      Label65.Caption = FORFACCS.LENTREGA(1).TEXT  'Localidad
'      Label68.Caption = FORFACCS.LENTREGA(2).TEXT  'Transportista
'      Label71.Caption = FORFACCS.MONEMIS.Caption   'Moneda emision
'      IMPO(9).Caption = FORFACCS.Text15.TEXT  'Importe percep IB
'      Label77.Caption = ""  ' Subtotal
'      IMPO(10).Caption = FORFACCS.Text17.TEXT 'Importe percep IVA
'      Label83.Caption = Format$(FVENCIM, "dd/mm/yy") 'Fecha vencim
'      On Error Resume Next
'      FECHADOC$ = FORFACCS.Text9.TEXT
'      FEXO = FECHANUM(FECHDOC$)
'      FEXU = CVINT(FVENCIM)
'      CADIVEN = DIENTRE(FEXO, FEXU)
'      On Error GoTo 0
'      Label86.Caption = FORMATNUM$(CADIVEN, "F5.0") 'cant dias para venc
'      IMPO(11).Caption = TRIM(FORFACCS.Text18) ' Valor FOB
'      Label92.Caption = FORMATNUM$(TICAMONE(2), "F7.4") ' Tipo cambio
'      '
'      DESTIDOC% = XVALO(FORFACCS.Text1.TEXT)
'      Label95.Caption = CONTACTCLI$(DESTIDOC%) 'Contacto en cliente
'      Label98.Caption = TRIM$(FINFACTURA.Text2(2)) 'Num bultos
'      Label101.Caption = TRIM$(FINFACTURA.Text2(11)) + " Kg" 'Peso neto
'      Label104.Caption = TRIM$(FINFACTURA.Text2(4)) + " Kg" 'Peso bruto
'      Label107.Caption = TRIM$(FINFACTURA.Text2(5)) 'Banco Exterior
'      Label110.Caption = TRIM$(FINFACTURA.Text2(8)) 'Banco Local
'      Label113.Caption = TRIM$(FINFACTURA.Text2(10)) 'Credito
'      Label116.Caption = TRIM$(FINFACTURA.Text2(9)) 'Guia Importacion
'      Label119.Caption = TRIM$(FINFACTURA.Text2(6)) 'Puerto Embarque
'      Label122.Caption = TRIM$(FINFACTURA.Text2(7)) 'Destino
'  '
End Sub
'\\\OT-06-0405-MH-FIN///

