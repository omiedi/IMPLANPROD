VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form DatoExtRem 
   BackColor       =   &H00E2D3C0&
   ClientHeight    =   5190
   ClientLeft      =   1920
   ClientTop       =   3285
   ClientWidth     =   12075
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5198.828
   ScaleMode       =   0  'User
   ScaleWidth      =   11391.51
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton Command6 
      Caption         =   "Abrir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   650
      Left            =   4080
      Picture         =   "DatoExtRem.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   65
      Top             =   5400
      Width           =   765
   End
   Begin VB.TextBox Text21 
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
      Height          =   330
      Left            =   240
      TabIndex        =   64
      TabStop         =   0   'False
      Text            =   " "
      Top             =   5745
      Width           =   3750
   End
   Begin VB.CommandButton Command4 
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
      Height          =   360
      Left            =   3840
      TabIndex        =   63
      Top             =   5400
      Width           =   175
   End
   Begin VB.TextBox Text14 
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
      Height          =   330
      Left            =   2445
      TabIndex        =   62
      Text            =   " "
      Top             =   5400
      Width           =   1380
   End
   Begin VB.TextBox Text11 
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
      Height          =   945
      Left            =   8400
      MultiLine       =   -1  'True
      TabIndex        =   61
      Text            =   "DatoExtRem.frx":030A
      Top             =   1200
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.TextBox Text7 
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
      Height          =   330
      Left            =   1965
      TabIndex        =   60
      TabStop         =   0   'False
      Top             =   840
      Visible         =   0   'False
      Width           =   9915
   End
   Begin VB.TextBox Text10 
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
      Height          =   330
      Left            =   120
      TabIndex        =   59
      TabStop         =   0   'False
      Text            =   "Facturas Asociadas:"
      Top             =   840
      Visible         =   0   'False
      Width           =   1802
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2D3C0&
      Height          =   4335
      Left            =   120
      TabIndex        =   14
      Top             =   840
      Width           =   11895
      Begin VB.Frame Frame17 
         BackColor       =   &H00E2D3C0&
         Caption         =   "Items"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   10680
         TabIndex        =   57
         Top             =   1920
         Width           =   1095
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
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
            Height          =   270
            Left            =   90
            TabIndex        =   58
            Top             =   240
            Width           =   885
         End
      End
      Begin VB.Frame Frame16 
         BackColor       =   &H00E2D3C0&
         Caption         =   "Orden  de Suministro"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   6390
         TabIndex        =   55
         Top             =   1200
         Width           =   5355
         Begin VB.TextBox Text22 
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
            Height          =   285
            Left            =   120
            MaxLength       =   24
            TabIndex        =   56
            Top             =   300
            Width           =   3940
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00E2D3C0&
         Caption         =   "Valor Declarado de los Items Seleccionados"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   6390
         TabIndex        =   53
         Top             =   1935
         Width           =   4155
         Begin VB.TextBox Text17 
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
            Height          =   285
            Left            =   120
            TabIndex        =   54
            Top             =   250
            Width           =   3930
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00E2D3C0&
         Caption         =   "Observaciones"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1635
         Left            =   6360
         TabIndex        =   43
         Top             =   2580
         Width           =   5355
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
            Height          =   285
            Index           =   4
            Left            =   210
            TabIndex        =   51
            Top             =   2500
            Width           =   3375
         End
         Begin VB.TextBox Text16 
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
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   50
            Top             =   315
            Width           =   5145
         End
         Begin VB.TextBox Text16 
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
            Height          =   285
            Index           =   1
            Left            =   120
            TabIndex        =   49
            Top             =   630
            Width           =   5145
         End
         Begin VB.CommandButton Command2 
            Caption         =   "x"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   470
            Left            =   4995
            TabIndex        =   48
            ToolTipText     =   "Borra información de Caja de Texto"
            Top             =   1080
            Width           =   300
         End
         Begin VB.TextBox TXTFLETE 
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
            Height          =   460
            Left            =   960
            TabIndex        =   47
            Text            =   " "
            Top             =   1080
            Width           =   3975
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00E2D3C0&
            Height          =   580
            Left            =   600
            TabIndex        =   44
            Top             =   980
            Width           =   350
            Begin VB.OptionButton Option1 
               BackColor       =   &H00E2D3C0&
               Caption         =   "Option1"
               Height          =   195
               Left            =   60
               TabIndex        =   46
               Top             =   140
               Width           =   255
            End
            Begin VB.OptionButton Option2 
               BackColor       =   &H00E2D3C0&
               Caption         =   "Option2"
               Height          =   200
               Left            =   60
               TabIndex        =   45
               Top             =   360
               Width           =   255
            End
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Flete:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   7
            Left            =   0
            TabIndex        =   52
            Top             =   1020
            Width           =   495
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E2D3C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3020
         Left            =   0
         TabIndex        =   20
         Top             =   1200
         Width           =   6285
         Begin VB.Frame Frame14 
            BackColor       =   &H00E2D3C0&
            Caption         =   "Conductor"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1170
            Left            =   120
            TabIndex        =   39
            Top             =   1680
            Width           =   2325
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
               Height          =   285
               Left            =   105
               TabIndex        =   41
               Top             =   340
               Width           =   2115
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
               Height          =   285
               Left            =   630
               TabIndex        =   40
               Top             =   735
               Width           =   1590
            End
            Begin VB.Label Label7 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "DNI:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   105
               TabIndex        =   42
               Top             =   800
               Width           =   435
            End
         End
         Begin VB.Frame Frame12 
            BackColor       =   &H00E2D3C0&
            Caption         =   "Orden de Compra"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Left            =   120
            TabIndex        =   37
            Top             =   210
            Width           =   2325
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
               Height          =   285
               Left            =   315
               TabIndex        =   38
               Top             =   315
               Width           =   1695
            End
         End
         Begin VB.Frame Frame11 
            BackColor       =   &H00E2D3C0&
            Caption         =   "Numero de Bultos"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Left            =   120
            TabIndex        =   35
            Top             =   945
            Width           =   2325
            Begin VB.TextBox Text15 
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
               Height          =   285
               Left            =   525
               TabIndex        =   36
               Top             =   315
               Width           =   1275
            End
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H00E2D3C0&
            Caption         =   "Transportista"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2640
            Left            =   2520
            TabIndex        =   21
            Top             =   210
            Width           =   3690
            Begin VB.Frame Frame5 
               BackColor       =   &H00E2D3C0&
               Caption         =   "Domicilio"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1005
               Left            =   105
               TabIndex        =   31
               Top             =   1560
               Width           =   3480
               Begin VB.TextBox DOMITRAN 
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
                  Height          =   285
                  Index           =   3
                  Left            =   120
                  TabIndex        =   34
                  Top             =   240
                  Width           =   2980
               End
               Begin VB.TextBox DOMITRAN 
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
                  Height          =   285
                  Index           =   4
                  Left            =   120
                  TabIndex        =   33
                  Top             =   600
                  Width           =   2980
               End
               Begin VB.CommandButton Command3 
                  Caption         =   "..."
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   645
                  Left            =   3120
                  TabIndex        =   32
                  ToolTipText     =   "Aplicar Domicilio de Transporte Como Domicilio de Entrega"
                  Top             =   240
                  Width           =   300
               End
            End
            Begin VB.Frame Frame10 
               BackColor       =   &H00E2D3C0&
               Caption         =   "Patente"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   620
               Left            =   1995
               TabIndex        =   29
               Top             =   975
               Width           =   1590
               Begin VB.TextBox Text18 
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
                  Height          =   285
                  Left            =   105
                  TabIndex        =   30
                  Top             =   210
                  Width           =   1380
               End
            End
            Begin VB.Frame Frame9 
               BackColor       =   &H00E2D3C0&
               Caption         =   "C.U.I.T."
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   620
               Left            =   105
               TabIndex        =   27
               Top             =   975
               Width           =   1800
               Begin VB.TextBox Text12 
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
                  Height          =   285
                  Left            =   105
                  TabIndex        =   28
                  Top             =   210
                  Width           =   1590
               End
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
               Height          =   285
               Index           =   2
               Left            =   90
               TabIndex        =   26
               Top             =   675
               Width           =   3510
            End
            Begin VB.CommandButton Command19 
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
               Height          =   330
               Left            =   840
               TabIndex        =   25
               Top             =   260
               Width           =   175
            End
            Begin VB.TextBox Text45 
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
               Height          =   330
               Left            =   120
               TabIndex        =   24
               Top             =   255
               Width           =   750
            End
            Begin VB.CommandButton Command16 
               Caption         =   "Nuevo"
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
               Left            =   1200
               Style           =   1  'Graphical
               TabIndex        =   23
               ToolTipText     =   "Agrega Nuevo Transporte a la Base de Datos"
               Top             =   240
               Width           =   1050
            End
            Begin VB.CommandButton Command18 
               Caption         =   "Editar"
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
               Left            =   2520
               Style           =   1  'Graphical
               TabIndex        =   22
               ToolTipText     =   "Modificar Datos del Transporte Activo"
               Top             =   240
               Width           =   1065
            End
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E2D3C0&
         Caption         =   "Datos de Entrega"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   945
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   11595
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
            Height          =   285
            Index           =   1
            Left            =   240
            TabIndex        =   19
            TabStop         =   0   'False
            Text            =   "Domicilio:"
            Top             =   240
            Width           =   1770
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
            Height          =   285
            Index           =   2
            Left            =   240
            TabIndex        =   18
            TabStop         =   0   'False
            Text            =   "Localidad:"
            Top             =   600
            Width           =   1770
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
            Height          =   285
            Index           =   0
            Left            =   2040
            TabIndex        =   17
            Top             =   240
            Width           =   5535
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
            Height          =   285
            Index           =   1
            Left            =   2040
            TabIndex        =   16
            Top             =   600
            Width           =   5535
         End
      End
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
      Left            =   6720
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   120
      Width           =   1170
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Agregar"
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
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   480
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.TextBox Text6 
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
      Left            =   120
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   480
      Width           =   6525
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
      Left            =   9840
      TabIndex        =   10
      TabStop         =   0   'False
      Text            =   "Fecha:"
      Top             =   480
      Width           =   810
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
      Left            =   10770
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   480
      Width           =   1113
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
      Left            =   9960
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   120
      Width           =   1770
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
      Left            =   1920
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   120
      Width           =   1170
   End
   Begin VB.TextBox Text1 
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
      Left            =   8160
      TabIndex        =   6
      TabStop         =   0   'False
      Text            =   "Numero de Remito:"
      Top             =   120
      Width           =   1770
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
      Index           =   0
      Left            =   120
      TabIndex        =   5
      TabStop         =   0   'False
      Text            =   "Numero de Cliente:"
      Top             =   120
      Width           =   1770
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
      Height          =   300
      Left            =   11715
      TabIndex        =   2
      Top             =   120
      Width           =   175
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   720
      OleObjectBlob   =   "DatoExtRem.frx":037C
      Top             =   3000
   End
   Begin MSFlexGridLib.MSFlexGrid GRID 
      Height          =   2445
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "Doble Click Seleccióna Items a Imprimir"
      Top             =   840
      Width           =   11715
      _ExtentX        =   20664
      _ExtentY        =   4313
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      FillStyle       =   1
      GridLines       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Formulario"
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
      Left            =   600
      TabIndex        =   67
      Top             =   5520
      Width           =   1725
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Formulario"
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
      Left            =   360
      TabIndex        =   66
      Top             =   130
      Width           =   1725
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      Height          =   300
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   195
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   855
   End
   Begin VB.Menu mnuarchivo 
      Caption         =   "Archivo"
      Begin VB.Menu mnuGrabarforipre 
         Caption         =   "Grabar Plantilla de Impresión de Formulario"
      End
      Begin VB.Menu mnuImprimir 
         Caption         =   "Imprimir"
      End
      Begin VB.Menu mnuRemOficial 
         Caption         =   "Imprimir Remito Oficial"
         Visible         =   0   'False
      End
      Begin VB.Menu mnAnulRemOficial 
         Caption         =   "Anular Remito Oficial"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRevAnul 
         Caption         =   "Revertir Anulación"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuEdicion 
      Caption         =   "Edición"
      Begin VB.Menu mnuRevertirSeleccion 
         Caption         =   "Revertir Selección"
         Visible         =   0   'False
      End
      Begin VB.Menu s 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuConfiguracion 
      Caption         =   "Configuración"
      Begin VB.Menu mnuConfForm 
         Caption         =   "Configuración de Formulario de Impresion"
      End
   End
End
Attribute VB_Name = "DatoExtRem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Command1_Click()
     Command1.Enabled = False
     If TRIM$(Text3) = "" Then
       Call COMUNI("Imposible Agregar Sin Remito Seleccionado Previamente")
       GoTo 99
     End If
     
     Call CARSELMESA
     
     '
10   Screen.MousePointer = 1
     Call SELECHO("SELMESA")
     MESASEL$ = VALACT1$("SELMESA")
     If TRIM$(MESASEL$) = "" Then
        GoTo 99
     End If
     Call ABRESTOCKS
     '
     CORRESQ$ = MESANO$(MESASEL$, Des%, Has%)
     '
     Desde = FETEXCOR1$(Des%)
     Hasta = FETEXCOR1$(Has%)
     '
     CMO$ = TRIM$(CONTROL$("", "CODREFAC"))
     If Len(CMO$) <> 2 Then
        CMO$ = "RT"
     End If
     CMO1$ = "RP"
     
     CONDI$ = "FechMov >= '" & Desde & "'"
     CONDI$ = CONDI$ + "AND FechMov <= '" & Hasta & "'"
     CONDI$ = CONDI$ + " AND substring(REFERCOR,1,4) <> 'ANUL' " 'PARA EVITAR COMPROBANTES ARRUINADOS
     CONDI$ = CONDI$ + " AND (CodiMovi =  '" & CMO$ & "'"
     CONDI$ = CONDI$ + " OR CodiMovi =  '" & CMO1$ & "')"
     CONDI$ = CONDI$ + " AND substring(DOPRILAR,1,2) = 'RT' "
     CONDI$ = CONDI$ + " AND (CodiMovi =  '" & CMO$ & "'"
     CONDI$ = CONDI$ + " OR CodiMovi =  '" & CMO1$ & "')"
     CONDI$ = CONDI$ + " AND (substring(DOPRILAR,1,2) = 'RT' OR substring(DOPRILAR,1,2) = 'RP') "
    
     CONDI$ = CONDI$ + " AND Nume_Clie = " & XVALO(Text2)
     '
     Screen.MousePointer = 11
     '
     Call COPYSTRU("SELDOCAP", "SELREMIT")
12   Call CARGALISEL("SELREMIT", "MOVISTO", "DOPRILAR/A18;FechMov/D8;REFERCOR/A24", CONDI$, "DISTINCT")
     '
     If ULTREG&("SELREMIT") < 1 Then
        Call MENSERR(24, "No se Encontraron Remitos en el Rango Seleccionado.")
        GoTo 10
     End If

15   Call SELECHO("SELREMIT")
     '
     Screen.MousePointer = 1
     '
     NUDOCU$ = TRIM$(Mid$(VALACT1$("SELREMIT"), 6))
     FECHACBTE$ = Mid$(VALACT2$("SELREMIT"), 1, 8)
     Me.Text4 = FECHACBTE$

     MEXAGE$ = "Remito Existente en Grilla"
     If VALIDA_COD_GRID%(NUDOCU$, Me.GRID, 9, MEXAGE$) < 1 Then
        Text3.TEXT = NUDOCU$
        
     End If
    
    
99   Command1.Enabled = True

End Sub

Private Sub Command16_Click()
   Command16.Enabled = False
   '
   If DERACCE%("ATRANSPO", "Actualizar Base de Datos de Transportes") < 2 Then
      GoTo 99
   End If
   '
   VDEF$ = REGBLAN$("TRANSPOR")
   CONDI$ = "CODI_TRA > 0"
10 PROXI% = 1 + XVALO(VALOBADA("TRANSPORTES", "MAX(CODI_TRA)", CONDI$, "NOMESS"))
   Call REPLA(VDEF$, MKI$(PROXI%), 1, 2)
   '
   OBX$ = OBJPLA$("CARGA-TRANSPOR", VDEF$)
   If OBX$ <> "" Then
     If TRIM$(Mid$(OBX$, 3, 32)) = "" Then
       Call MENSERR(24, "Falta Nombre del Transporte")
       VDEF$ = OBX$
       GoTo 10
     End If
     '
     Call BLOQARCH("TRANSPOR")
     PROXI% = 1 + XVALO(VALOBADA("TRANSPORTES", "MAX(CODI_TRA)", CONDI$, "NOMESS"))
     Call REPLA(OBX$, MKI$(PROXI%), 1, 2)
     Call REGAPP("TRANSPOR", OBX$)
     Call CIERRARCH("TRANSPOR")
     ' Graba en Tabla SQL Transportes.
     COD1% = PROXI%
     DENO$ = TRIM$(Mid$(OBX$, 3, 32))
     RASO$ = TRIM$(Mid$(OBX$, 35, 48))
     DOMI$ = TRIM$(Mid$(OBX$, 83, 48))
     LOCA$ = TRIM$(Mid$(OBX$, 131, 48))
     Cuit$ = TRIM$(Mid$(OBX$, 179, 15))
     PIVA% = Asc(Mid$(OBX$, 194, 1))
     TELE$ = TRIM$(Mid$(OBX$, 195, 48))
     FAXI$ = TRIM$(Mid$(OBX$, 243, 48))
     mail$ = TRIM$(Mid$(OBX$, 291, 48))
     CTAC$ = TRIM$(Mid$(OBX$, 339, 48))
     Dim TRANSPOR As ADODB.Recordset
     Set TRANSPOR = New ADODB.Recordset
     SQLX$ = "Select * from TRANSPORTES"
25   On Error Resume Next
     TRANSPOR.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
     If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
         On Error GoTo 0
         Call CapturaErrorOpen
         GoTo 25
     End If
     With TRANSPOR
          If COD1% > 0 Then
                 .AddNew
                 !CODI_TRA = COD1%
                 !DENO_TRA = DENO$
                 !RASO_TRA = RASO$
                 !DOMI_TRA = DOMI$
                 !LOCA_TRA = LOCA$
                 !CUIT_TRA = Cuit$
                 !PIVA_TRA = PIVA%
                 !TELE_TRA = TELE$
                 !FAXI_TRA = FAXI$
                 !MAIL_TRA = mail$
                 !CTAC_TRA = CTAC$
                 .Update
          End If
     End With
     '
   End If
99 Command16.Enabled = True

End Sub

Private Sub Command18_Click()
   Command18.Enabled = False
   '
   If DERACCE%("ATRANSPO", "Actualizar Base de Datos de Transportes") < 2 Then
      GoTo 99
   End If
   '
   NTRA = XVALO(Text45)
   If NTRA > 0 Then
     If NTRA < 32767 Then
       NTRAI% = NTRA
       'If ECOARCH$("TRANSPOR", MKI$(NTRAI%)) <> "" Then
       
       'MODIFICADO POR QUE EN GABAL SI GRABA EN UNA EMPRESA
       'NO LE DEJA EDITAR POR QUE NO ENCUENTRA EL ECOARCH
       CONDI$ = "CODI_TRA = " & NTRAI%
       If CantRegistros("TRANSPORTES", CONDI$) > 0 Then
         'ASIGNO CON LA FUNICION SI NO ENCUENTRA EN SQL TRAE DE DAT
         '*********************************************************
         CUITRA$ = CUITTRANS$(NTRAI%)
         DENOTRA$ = DENOTRANS$(NTRAI%)
         RASOTRA$ = RASOTRANS$(NTRAI%)
         DOMITRA$ = DOMITRANS$(NTRAI%)
         LOCATRA$ = LOCATRANS$(NTRAI%)
         TELETRA$ = TELETRANS$(NTRAI%)
         FAXITRA$ = FAXITRANS$(NTRAI%)
         MAILTRA$ = MAILTRANS$(NTRAI%)
         CTACTRA$ = CTACTRANS$(NTRAI%)
         PIVATRA% = PIVATRANS%(NTRAI%)
         '
         'ARMO EL STRING
         '**************
         Call REPLA(VDEF$, MKI$(NTRAI%), 1, 2)
         Call REPLA(VDEF$, DENOTRA$, 3, 32)
         Call REPLA(VDEF$, RASOTRA$, 35, 48)
         Call REPLA(VDEF$, DOMITRA$, 83, 48)
         Call REPLA(VDEF$, LOCATRA$, 131, 48)
         Call REPLA(VDEF$, CUITRA$, 179, 15)
         Call REPLA(VDEF$, Chr$(PIVATRA%), 194, 1)
         Call REPLA(VDEF$, TELETRA$, 195, 48)
         Call REPLA(VDEF$, FAXITRA$, 243, 48)
         Call REPLA(VDEF$, MAILTRA$, 291, 48)
         Call REPLA(VDEF$, CTACTRA$, 339, 48)

         'VDEF$ = FILTERGETRECORD$("TRANSPOR", 1, MKI$(NTRAI%))
10       OBX$ = OBJPLA$("CARGA-TRANSPOR", VDEF$)
         If OBX$ <> "" Then
           If TRIM$(Mid$(OBX$, 3, 32)) = "" Then
             Call MENSERR(24, "Falta Nombre del Transporte")
             VDEF$ = OBX$
             GoTo 10
           End If
           '
           Call BLOQARCH("TRANSPOR")
           Call FILTERPUTRECORD("TRANSPOR", 1, MKI$(NTRAI%), OBX$)
           Call CIERRARCH("TRANSPOR")
           ' Edita y Graba en Tabla SQL Transportes.
           COD1% = NTRAI%
           DENO$ = TRIM$(Mid$(OBX$, 3, 32))
           RASO$ = TRIM$(Mid$(OBX$, 35, 48))
           DOMI$ = TRIM$(Mid$(OBX$, 83, 48))
           LOCA$ = TRIM$(Mid$(OBX$, 131, 48))
           Cuit$ = TRIM$(Mid$(OBX$, 179, 15))
           PIVA% = Asc(Mid$(OBX$, 194, 1))
           TELE$ = TRIM$(Mid$(OBX$, 195, 48))
           FAXI$ = TRIM$(Mid$(OBX$, 243, 48))
           mail$ = TRIM$(Mid$(OBX$, 291, 48))
           CTAC$ = TRIM$(Mid$(OBX$, 339, 48))
'           SQLX$ = "Codi_Tra = " & COD1% & ""
           
          Dim RS1 As ADODB.Recordset
          Set RS1 = New ADODB.Recordset
          SQLX$ = "Select * from TRANSPORTES"
          SQLX$ = SQLX$ + " WHERE Codi_Tra = " & COD1%
30        On Error Resume Next
          RS1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText

          If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
            On Error GoTo 0
            Call CapturaErrorOpen
            GoTo 30
          End If
          On Error GoTo 0
        
          '
          With RS1
           Do While Not .EOF
              !DENO_TRA = Left$(TRIM$(DENO$), 128)
              !RASO_TRA = Left$(TRIM$(RASO$), 128)
              !DOMI_TRA = Left$(TRIM$(DOMI$), 128)
              !LOCA_TRA = Left$(TRIM$(LOCA$), 128)
              !CUIT_TRA = Left$(TRIM$(Cuit$), 32)
              !PIVA_TRA = PIVA%
              
              
              !TELE_TRA = Left$(TRIM$(TELE$), 64)
              !FAXI_TRA = Left$(TRIM$(FAXI$), 64)
              !MAIL_TRA = Left$(TRIM$(mail$), 64)
'              !MAIL_TRA = Left$(TRIM$(Text10(3)), 128)
              !CTAC_TRA = Left$(TRIM$(CTAC$), 128)
              .Update
              .MoveNext
           Loop
          End With
          RS1.Close
          Set RS1 = Nothing
           '
           Call Text45_Change
         End If
         GoTo 99
       End If
     End If
   End If
   '
   Call MENSERR(24, "Imposible Editar Datos Transportista")
   '
99 Command18.Enabled = True

End Sub


Private Sub Command19_Click()
   Call COPYSTRU("TRANSPOR", "TRANSPO1")
   Call CARGALISEL("TRANSPO1", "TRANSPORTES", "CODI_TRA/F5;DENO_TRA/A32", "CODI_TRA > 0 ORDER BY CODI_TRA", "NOMESS")
   Call SELECHO("TRANSPO1")
   NTRAN% = XVALO(VALACT1$("TRANSPO1"))
   If NTRAN% > 0 Then
        Text45 = TRIM$(Str$(NTRAN%))
   End If

End Sub

Private Sub Command2_Click()
   Command2.Enabled = False
   TXTFLETE = ""
   Option1.Value = False
   Option2.Value = False
   Command2.Enabled = True

End Sub

Private Sub Command3_Click()
   Me.LENTREGA(0) = Me.DOMITRAN(3)
   Me.LENTREGA(1) = Me.DOMITRAN(4)

End Sub

Private Sub Command4_Click()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text14.TEXT = VDEVU$
        Text21.TEXT = VALACT2$("INDIFRM")
    End If

End Sub

Private Sub Command5_Click()
     Command5.Enabled = False
     Call CARSELMESA
     '
     Text8 = ""
10   Screen.MousePointer = 1
     Call SELECHO("SELMESA")
     MESASEL$ = VALACT1$("SELMESA")
     If TRIM$(MESASEL$) = "" Then
        GoTo 99
     End If
     Call ABRESTOCKS
     '
     CORRESQ$ = MESANO$(MESASEL$, Des%, Has%)
     '
     Desde = FETEXCOR1$(Des%)
     Hasta = FETEXCOR1$(Has%)
     '
     CMO$ = TRIM$(CONTROL$("", "CODREFAC"))
     If Len(CMO$) <> 2 Then
        CMO$ = "RT"
     End If
     CMO1$ = "RP"
     
     CONDI$ = "FechMov >= '" & Desde & "'"
     CONDI$ = CONDI$ + "AND FechMov <= '" & Hasta & "'"
     CONDI$ = CONDI$ + " AND substring(REFERCOR,1,4) <> 'ANUL' " 'PARA EVITAR COMPROBANTES ARRUINADOS
     CONDI$ = CONDI$ + " AND (CodiMovi =  '" & CMO$ & "'"
     CONDI$ = CONDI$ + " OR CodiMovi =  '" & CMO1$ & "')"
     CONDI$ = CONDI$ + " AND (substring(DOPRILAR,1,2) = 'RT' OR substring(DOPRILAR,1,2) = 'RP') "
     '
     Screen.MousePointer = 11
     '
     Call COPYSTRU("SELDOCAP", "SELREMIT")
12   Call CARGALISEL("SELREMIT", "MOVISTO", "DOPRILAR/A18;FechMov/D8;REFERCOR/A24", CONDI$, "DISTINCT")
     '
     If ULTREG&("SELREMIT") < 1 Then
        Call MENSERR(24, "No se Encontraron Remitos en el Rango Seleccionado.")
        GoTo 10
     End If

15   Call SELECHO("SELREMIT")
     '
     Screen.MousePointer = 1
     '
     NUDOCU$ = TRIM$(Mid$(VALACT1$("SELREMIT"), 6))
     If Left$(VALACT1$("SELREMIT"), 2) = "RP" Then
       NUDOCU$ = TRIM$(Mid$(VALACT1$("SELREMIT"), 4))
       Text8 = "RP"
     End If
     FECHACBTE$ = Mid$(VALACT2$("SELREMIT"), 1, 8)
     If NUDOCU$ = "" Then GoTo 10
     
     Me.Text4 = FECHACBTE$
     Me.Text3 = NUDOCU$
99   Command5.Enabled = True
End Sub

Private Sub Command6_Click()
    Call CONECRUN("FLFORMS/" + TRIM$(Text14.TEXT), "Formularios de Impresion")
    If EXISTE%("FLFORMS.EXE") < 1 Then
        Call FINAL("")
    End If

'    Me.Height = 5865
End Sub

Private Sub Form_Load()
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me)
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If InStr(EPAC$, "MEDITEA") > 0 Then
       mnuImprimir.Visible = False
    End If
    
    'PRESEENTA PLANTILLA DE IMPRESION (OCULTO)
    Text14.TEXT = CONTROL$("", "VALEREMI")
    Text21.TEXT = ECOARCH$("INDIFRM", Text14.TEXT)

    
    Campos$ = "CI/F0;Ord./F4;Descripción/A32;Salida/F9.1;Ingreso/f9.1;Nro.Ped./F7;P.Unit./F12.2;N° Serie/A16;Remito/A16"
    Call CARGA_ENCABEZADO(Me.GRID, Campos$, Me)
End Sub

Private Sub GRID_DblClick()
    i& = GRID.MouseRow
    J& = GRID.MouseCol
    If i& = 0 Or i& > GRID.Rows Then Exit Sub
    '
    On Error Resume Next
    GRID.Row = i&
    GRID.COL = J&
    On Error Resume Next

    REG& = i&
    If REG& < 1 Then
        Exit Sub
    End If
    Call COLOREAR_GRILLA_SOLAFILA(Me.GRID, &HFFFF80, REG&)
    Call CalcutotalSeleccionado
End Sub

Private Sub LENTREGA_DblClick(Index As Integer)
    If Index = 0 Then
       CAMPO$ = "Domicilio de Entrega": CAMPOFILTRO$ = "DomiEnt"
    ElseIf Index = 1 Then
       CAMPO$ = "Localidad de Entrega": CAMPOFILTRO$ = "LocaEnt"
    ElseIf Index = 2 And XVALO(Text19) < 1 Then: CAMPOFILTRO$ = "TranspEnt"
        CAMPO$ = "Transporte"
    Else
        Exit Sub
    End If
    
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, CAMPO$ & "/A99;CONSERPARADOR/A0", FRMZELECHO, , , ANTOT)
    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, CAMPO$ & "/A99;CONSERPARADOR/A0", FRMZELECHO, , , ANTOT)
    If ANTOT < 5500 Then ANTOT = 5500
    FRMZELECHO.Width = ANTOT + 1100
    FRMZELECHO.Caption = CAMPO$
    FRMZELECHO.MSF1.Rows = 1
    
    SQLX$ = "SELECT DISTINCT  " & CAMPOFILTRO$ & " AS VALOR1"
    If Index < 2 Then SQLX$ = SQLX$ + " , LocaEnt  AS VALOR2"
    SQLX$ = SQLX$ + " FROM PEDENCA WITH(NOLOCK) WHERE NroClie = " & XVALO(Text2)
    SQLX$ = SQLX$ + " ORDER BY " & CAMPOFILTRO$
    Set rs = READSET(SQLX$)
    With rs
      Do While Not .EOF
        i& = i& + 1
        FRMZELECHO.msf2.Rows = i& + 1
        FRMZELECHO.msf2.TextMatrix(i&, 1) = SAFETEXT$(!VALOR1)
        If Index <= 1 Then
          FRMZELECHO.msf2.TextMatrix(i&, 1) = SAFETEXT$(!VALOR1) & " " & SAFETEXT$(!VALOR2)
          If TRIM$(SAFETEXT$(!VALOR2)) <> "" Then
             FRMZELECHO.msf2.TextMatrix(i&, 2) = SAFETEXT$(!VALOR1) & "|" & SAFETEXT$(!VALOR2)
          End If
        End If
        .MoveNext
      Loop
    End With
    rs.Close
    Set rs = Nothing
    FRMZELECHO.Show 1
    '
    LECTURAX$ = TRIM$(RESP_ZELE_VECT(2))
    If LECTURAX$ = "" Then LECTURAX$ = TRIM$(RESP_ZELE_VECT(1))
    If LECTURAX$ <> "" Then
     If Index <= 1 Then
       Call TEXTOLOTES$(LECTURAX$, "|")
       On Error Resume Next
       LECTURA1$ = CADENATEX$(1)
       LECTURA2$ = CADENATEX$(2)
       On Error GoTo 0
       If Index = 0 Then LENTREGA(0).TEXT = Left$(LECTURA1$, 48)
       LENTREGA(1).TEXT = Left$(LECTURA2$, 48)
     Else
       LENTREGA(Index).TEXT = Left$(LECTURAX$, 48)
     End If
    End If

End Sub


Private Sub mnAnulRemOficial_Click()
10 Campos$ = "Fecha/D8;Remito Oficial/A18;Referencia/A48;ID/A0"
   Call CARGA_ENCABEZADO(FRMZELECHO.msf2, Campos$, FRMZELECHO)
   Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
   FRMZELECHO.Caption = "REMITOS OFICIALES GENERADOS"

   Dim obj As New RECORXET
   Set RS1 = obj.RS_REMITOSOFICIALES(1)
   Call Carga_MSF_Recordset(RS1, Campos$, FRMZELECHO.msf2, 1)
   Set obj = Nothing
   RS1.Close
   Set RS1 = Nothing

   FRMZELECHO.Width = 10000
   FRMZELECHO.Show 1
   NROREMI$ = TRIM$(RESP_ZELE_VECT(2))
   '
   If NROREMI$ = "" Then Exit Sub
   Call ACTUREGISTRO("MOVIOFI", "STATUS", "REMIOFI = '" & NROREMI$ & "'", 9, REGAF&, "NOMESS")
   Call ACTUREGISTRO("MOVIOFI", "REFERENCIA", "REMIOFI = '" & NROREMI$ & "'", "(Anul)" & RESP_ZELE_VECT(3), REGAF&, "NOMESS")

   GoTo 10
   
End Sub

Private Sub mnuConfForm_Click()
    If Me.Height < 6200 Then
       Me.Height = 6915
    Else
       Me.Height = 5955
    End If
End Sub

Private Sub mnuGrabarforipre_Click()
    Call GRACONTROL("", "VALEREMI", Text14.TEXT)
End Sub

Private Sub mnuimprimir_Click()
     Call ReimpresionDeRemito(Text3, OK%)
End Sub

Sub ReimpresionDeRemito(rto$, OK%, Optional TIPODOC1$, Optional NOPRESENTAMENSA%)
    If Me.GRID.Rows < 2 Then Exit Sub
    EPAC$ = TRIM$(UCase$(EMPREAC$))

'    PRINTERPORT$ = TRIM$(CONTROL$(IDENTER$, "PRIREEDI"))
    'FORIPRE$ = TRIM$(CONTROL$("", "FOREMIFA"))
    FORIPRE$ = TRIM$(CONTROL$("", "VALEREMI"))
    
    REPLAORI% = 0
    If CONTROL$("REMSTAN", "REPLAORI") = "SI" Then
       REPLAORI% = 1
    End If

    OK% = 0
    CMO$ = TRIM$(CONTROL$("", "VM"))
    '

    HaySeleccion% = 0
    For i& = 1 To GRID.Rows - 1
         If TRIM$(GRID.TextMatrix(i&, 0)) = "*" Then
            HaySeleccion% = HaySeleccion% + 1
            Exit For
         End If
    Next i&
    '
    If HaySeleccion% < 1 Then
       Call COMUNI("Atención No se Han Seleccionado Items Para Imprimir\Puede Seleccionar Todos desde el Menú > Edición > Revertir Selección\O Realizando Doble Click Sobre los Items a Imprimir")
       GoTo 99
    End If
    '
    
    If TRIM$(FORIPRE$) <> "" Then
      '

      Call LIMPIATETAB
      '
      FF% = CVINT(Text4)
      'FUERZA A CAMBIAR EL NRO CLIENTE EN UNA FUNCION DE DEVOLUCION DEL RECORDSET DE CLIENTE
      A$ = rasocli$(1) 'PQ SI NO NO ESTABA DEVOLVIENDO LOS DATOS DE DICHAS FUNCIONES.
      'FIN FUERZA A CAMBIAR NUMERO DE CLIENTE
      
      NCLI = XVALO(Text2)
      NROREMIT$ = rto$
      '
      FEX = FF%
      FECHDOC$ = FECHATEX$(FEX)
      NUMEDOC$ = NROREMIT$
      TIPODOC$ = "Vale de Consumo de Materiales"
      '
      If TRIM$(DOMITRAN(3)) = "" Then
         TRIM$ (DOMITRAN(3))
      End If
      CODPARAM$(1) = "TRANSPOR": DOCPARAM$(1) = TRIM$(LENTREGA(2))
      CODPARAM$(2) = "NUBULTOS": DOCPARAM$(2) = TRIM$(Text15.TEXT)
      CODPARAM$(3) = "CUIT-TRA": DOCPARAM$(3) = TRIM$(Text12.TEXT)
      CODPARAM$(4) = "VALO-DEC": DOCPARAM$(4) = TRIM$(Text17.TEXT)
      TEXTOENTREGA$ = LENTREGA(0) & " - " & LENTREGA(1)
      CODPARAM$(5) = "DOMI-TRA": DOCPARAM$(5) = TEXTOENTREGA$
       If InStr(EPAC$, "ENVAPLAST") > 0 Or InStr(EPAC$, "CONTEPLAST") > 0 Then
        If TRIM$(LENTREGA(2).TEXT) <> "" Then TEXTOENTREGA$ = TRIM$(LENTREGA(2).TEXT)
        If TRIM$(DOMITRAN(3).TEXT) <> "" Then TEXTOENTREGA$ = TEXTOENTREGA$ & "-" & TRIM$(DOMITRAN(3).TEXT)
        If TRIM$(DOMITRAN(4).TEXT) <> "" Then TEXTOENTREGA$ = TEXTOENTREGA$ & "-" & TRIM$(DOMITRAN(4).TEXT)
        CODPARAM$(5) = "DOMI-TRA": DOCPARAM$(5) = TEXTOENTREGA$
      End If

'      CODPARAM$(5) = "DOMI-TRA": DOCPARAM$(5) = TRIM$(DOMITRAN(3))
      CODPARAM$(6) = "DOCUCOND": DOCPARAM$(6) = TRIM$(Text13.TEXT)
'      CODPARAM$(7) = "DEPOSAL ": DOCPARAM$(7) = TRIM$(DEPOSAL$)
      CODPARAM$(8) = "NRO-REMI": DOCPARAM$(8) = TRIM$(rto$)
      CODPARAM$(9) = "NRO-OCOM": DOCPARAM$(9) = TRIM$(Text20.TEXT)
      CODPARAM$(10) = "PATE-TRA": DOCPARAM$(10) = TRIM$(Text18.TEXT)
'      CODPARAM$(11) = "N-AUTENT": DOCPARAM$(11) = TRIM$(NORVENTA)
      CODPARAM$(12) = "COND-TRA": DOCPARAM$(12) = TRIM$(Text19)
      CODPARAM$(13) = "LOCA-TRA": DOCPARAM$(13) = TRIM$(DOMITRAN(4))
      CODPARAM$(14) = "NRO-AENT": DOCPARAM$(14) = TRIM$(Text22)

'      HORA$ = HORAENT$(NCLI)
      CODPARAM$(15) = "Z-PLANO": DOCPARAM$(15) = ""
      CODPARAM$(16) = "IMP-TOTA": DOCPARAM$(16) = ""
      CODPARAM$(17) = "DOCU-APL": DOCPARAM$(17) = ""
      CODPARAM$(18) = "DESTINO": DOCPARAM$(18) = TRIM$(TXTFLETE)
      CODPARAM$(19) = "OBSERVA": DOCPARAM$(19) = TRIM$(Text16(0)) & " " & TRIM$(Text16(1))
      CODPARAM$(20) = "LAR-CUTE": DOCPARAM$(20) = TRIM$(LENTREGA(0))
      CODPARAM$(21) = "LAR-MATE": DOCPARAM$(21) = TRIM$(LENTREGA(1))
      CODPARAM$(22) = "TELE-CLI": DOCPARAM$(22) = TEL1CLI$(NCLI)
      CODPARAM$(23) = "COND-PAG": DOCPARAM$(23) = ECOARCH$("CONPAG", Chr$(CPAGCLI%(NCLI)))
      
      CAPARDOC% = 23
      '
      CODTABU1$(1) = "COD-MAT"
      CODTABU1$(2) = "DES-MAT"
      CODTABU1$(3) = "CANTIDAD"
      CODTABU1$(4) = "NRO-LOTE"
      CODTABU1$(5) = "UNIMED"
      CODTABU1$(6) = "N-OCOMPR"
      CODTABU1$(7) = "ORD-ITEM"
      CODTABU1$(8) = "REF-MAT1"
      CODTABU1$(9) = "REF-MAT2"
      CODTABU1$(10) = "PRE-NETO"
      CODTABU1$(11) = "NET-ITEM"
      CODTABU1$(12) = "REF-MAT6"
      CODTABU1$(13) = "REF-MAT5"
      CACOLTAB1% = 13
      '
      DESTIDOCNUE = NCLI
      CAITAB1% = 0
      '
      jj& = 0
      '1 "CI/F0;
      '2 Ord./F4;
      '3 Descripción/A32
      '4;Salida/F9.1;
      '5Ingreso/f9.1;
      '6 Nro.Ped./F7;
      '7 P.Unit./F12.2
      '8;N° Serie/A16
      '9;Remito/A16"
      YAPASO% = 0: YAPASO% = 0
      For i& = 1 To GRID.Rows - 1
         If TRIM$(GRID.TextMatrix(i&, 0)) <> "*" Then GoTo 50
         REMI$ = TRIM$(GRID.TextMatrix(i&, 9))
         If REMI_ANT$ <> "" Then
            If REMI_ANT$ <> REMI$ And YAPASO% = 0 Then
               CODPARAM$(8) = "NRO-REMI": DOCPARAM$(8) = "Varios"
               YAPASO% = 1
            End If
         End If
         'GENERO EL STRING CON LOS NUMEROS DE PEDIDOS RELACIONADOS
         NROPED& = XVALO(GRID.TextMatrix(i&, 6))
            If (NROPED_ANT& <> NROPED&) And NROPED& > 0 Then
               PEDIDOSX$ = PEDIDOSX$ & NROPED& & "-"
            End If
         CI1% = XVALO(GRID.TextMatrix(i&, 1))
         jj& = jj& + 1
          CAITAB1% = CAITAB1% + 1
          CODDX$ = CODEXT$(XVALO(GRID.TextMatrix(i&, 1)))
          If REPLAORI% = 1 Then
             CODDX$ = CODORIG$(CI1%, NCLI)
          End If
          NPDX& = XVALO(GRID.TextMatrix(i&, 6))
          CANSAL = XVALO(GRID.TextMatrix(i&, 4))
          CANING = XVALO(GRID.TextMatrix(i&, 5))
          CANS$ = FORMATNUM$(CANSAL - CANING, "F10." & CNTDC$ & "")
          TETABU1$(1, CAITAB1%) = CODDX$
          TETABU1$(2, CAITAB1%) = GRID.TextMatrix(i&, 3)
          TETABU1$(3, CAITAB1%) = CANS$
          NSERIE$ = TRIM$(GRID.TextMatrix(i&, 8))
          If NSERIE$ <> "" And NSERIE$ <> "0" And InStr(EPAC$, "EYM") <= 0 Then
             NSERIE$ = "(N°.Serie: " & NSERIE$ & ")"
          End If

          TETABU1$(4, CAITAB1%) = NSERIE$
          TETABU1$(5, CAITAB1%) = Unimed$(CI1%)
          If InStr(EPAC$, "EYM") > 0 Then
              TETABU1$(6, CAITAB1%) = GRID.TextMatrix(i&, 6)
          End If
          
          'TETABU1$(6, CAITAB1%) = NUOCO$
          TETABU1$(7, CAITAB1%) = Str$(CAITAB1%)
          TETABU1$(8, CAITAB1%) = NUOCOCLI$(NPDX&, CI1%)
          TETABU1$(9, CAITAB1%) = NUPLANO$(CI1%)
'          TETABU1$(10, CAITAB1%) = FORMATNUM$(IMPOUNI#, "F12.2")
'          TETABU1$(11, CAITAB1%) = FORMATNUM$(IMPORTEX#, "F12.2")
'          CORIG$ = TRIM$(CODORIG$(CI%, NCLI))
'          If CORIG$ = CODEXT$(CI1%) Then CORIG$ = ""
'          TETABU1$(12, CAITAB1%) = CORIG$
'          'FIN ALEJPA

          TETABU1$(13, CAITAB1%) = XVALO(GRID.TextMatrix(i&, 2)) ' NUMERO DE ORDEN INGRESADO EN VENTANA DE ARCON NEUQUEN.
          'IMPRESION DE ESPECIFICACION
          If CONTROL("FACTUREM", "INCLUSPE") = "SI" Then ' IMPRIME ESPECIFICACIONES DEL ARTICULO
            XX$ = ESPECIFIT$(CI1%)
            Call PRINTEXT(XX$, 2, "DES-MAT", FORIPRE$)
          End If
          Call CARLINSEP         'Carga linea de separacion entre codigo y codigo
          REMI_ANT$ = REMI$
          NROPED_ANT& = NROPED&
50    Next i&
      CODPARAM$(4) = "VALO-DEC": DOCPARAM$(4) = TRIM$(Text17.TEXT)
      If InStrRev(PEDIDOSX$, "-") > 0 Then PEDIDOSX$ = Mid$(PEDIDOSX$, 1, Len(PEDIDOSX$) - 1)
      CODPARAM$(14) = "NRO-AENT": DOCPARAM$(14) = PEDIDOSX$
      Call PRINTDOC("@" + FORIPRE$)
      OK% = 1
      '
      If TRIM$(PRINTERPORT$) <> "" Then
         Call SETPRINPORT("RESTORE")
      End If
      '
    End If

99
End Sub


Private Sub mnuRemOficial_Click()
      FE% = HOY(HO$)
      User% = USNBR% Mod 100
      
      HaySeleccion% = 0
      For i& = 1 To GRID.Rows - 1
         If TRIM$(GRID.TextMatrix(i&, 0)) = "*" Then
            HaySeleccion% = HaySeleccion% + 1
            Exit For
         End If
      Next i&
      '
      If HaySeleccion% < 1 Then
       Call COMUNI("Atención No se Han Seleccionado Items Para Imprimir\Puede Seleccionar Todos desde el Menú > Edición > Revertir Selección\O Realizando Doble Click Sobre los Items a Imprimir")
       Exit Sub
      End If
      '
      If HaySeleccion% > 0 And HaySeleccion% < GRID.Rows - 1 Then
       OPX% = COMALTER%("Atención Solo se Imprimiran los Items Seleccionados\\Cancelar\Imprimir Items Seleccionados")
       If OPX% < 2 Then Exit Sub
      End If
      
      PREFI$ = TRIM$(InputBox$("Prefijo de Remito:", "Ingrese el Prefijo", PREFI$))
      If PREFI$ = "" Then Exit Sub
      '
      If Len(PREFI$) > 4 Then
         Call COMUNI("Prefijo no Válido (Máximo 4 Dígitos)")
         Exit Sub
      End If
      '
      While Len(PREFI$) < 4: PREFI$ = "0" + PREFI$: Wend
      '
      For U& = 1 To 4
           UI% = Asc(Mid$(PREFI$, U&))
           If UI% < 48 Or UI% > 57 Then
               Call COMUNI("Prefijo no Válido")
               Exit Sub
               Exit Sub
           End If
      Next U&
      '
20    NROREMIX$ = TRIM$(InputBox$("Nro de Remito: (Máximo 8 Caracteres)", "Prefijo " & PREFI$, NROREMIX$))
      If NROREMIX$ = "" Then Exit Sub
      
      If Len(NROREMIX$) > 8 Then
         Call COMUNI("Longitud Ingresada Supera la Cantidad Permitida")
         Exit Sub
      End If
      If XVALO(NROREMIX$) < 1 Then
         Call COMUNI("Debe Ingresar Números")
         Exit Sub
      End If
      REMITOFICIAL$ = "RT-X-" & PREFI$ & "-" & FORMATOCON0$(XVALO(NROREMIX$), 8)
      If CantRegistros&("MOVIOFI", "REMIOFI = '" & REMITOFICIAL$ & "' AND STATUS < 9", "NOMESS") > 0 Then
         Call COMUNI("Imposible Continuar !! Número de Remito Existente\Deberá Anular el Mismo e Intentar Nuevamente")
         GoTo 20
      End If
      
      Dim rs As ADODB.Recordset
      Set rs = New ADODB.Recordset
      
      On Error GoTo ERROR_GUARDAR
      FLEXCONN.BeginTrans
      SQLX$ = "Select * from MOVIOFI WHERE MOVIOFI_ID < 1"
      rs.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
      
      With rs
        For i& = 1 To GRID.Rows - 1
            If TRIM$(GRID.TextMatrix(i&, 0)) <> "*" Then GoTo 50
            .AddNew
            REMI$ = TRIM$(GRID.TextMatrix(i&, 9))
            NUREM& = XVALO(Mid$(REMI$, 6))
            NROPED& = XVALO(GRID.TextMatrix(i&, 6))
            CI1% = XVALO(GRID.TextMatrix(i&, 1))
            CODDX$ = CODEXT$(XVALO(GRID.TextMatrix(i&, 1)))
            If REPLAORI% = 1 Then
               CODDX$ = CODORIG$(CI1%, NCLI)
            End If
            NPDX& = XVALO(GRID.TextMatrix(i&, 6))
            CANSAL = XVALO(GRID.TextMatrix(i&, 4))
            CANING = XVALO(GRID.TextMatrix(i&, 5))
            NSERIE$ = TRIM$(GRID.TextMatrix(i&, 8))
            !COD_INTE = CI1%
            !Cod_Exte = CODDX$
            !FeReMovi = Now
            !FechMov = CVDAT(FE%)
            !Descrip = Left$(DESCRIT0$(CI1%), 64)
            !IDENLOTE = NSERIE$
            !DOPRILAR = REMI$
            !NUPEDCLI = NROPED&
            !NUME_CLIE = Text2
            !CANTINGRE = CANING
            !CantSalid = CANSAL
            !DOPRINUM = NUREM&
            !User_Tran = User%
            !REMIOFI = REMITOFICIAL$
            !REFERENCIA = Left$(rasocli$(XVALO(Text2)), 18) & "-" & FECHATEX$(FE%)
            !Status = 0
            If NPDX& <> NPDX_ANT& And NPDX& > 0 Then
               TTXX$ = Space$(76)
               TXTX$ = "Generación de Remito Oficial N°: " & REMITOFICIAL$ & vbCrLf
               Call AGREGA_ANOTAPED_NEW(NPDX&, TXTX$)
            End If
            NPDX_ANT& = NPDX&
            .Update
50      Next i&
      End With
ERROR_GUARDAR:

      If Err.Number <> 0 Then
        FLEXCONN.RollbackTrans
        Call COMUNI("Se Produjo un Error al Grabar.\No se Ha Realizado Ninguna Transacción.")
      Else
        FLEXCONN.CommitTrans
      End If
      
99
      On Error GoTo 0
      rs.Close
      Set rs = Nothing

      Call ReimpresionDeRemito(REMITOFICIAL$, OK%, "REMITOFI", 1)
      
End Sub

Private Sub mnuReportesVarios_Click()
    Call ACONSUL_PRE(TRIM$(UCase$(Name)))
End Sub

Private Sub mnuRevAnul_Click()

10 Campos$ = "Fecha/D8;Remito Oficial/A18;Referencia/A48;ID/A0"
   Call CARGA_ENCABEZADO(FRMZELECHO.msf2, Campos$, FRMZELECHO)
   Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
   FRMZELECHO.Caption = "REMITOS OFICIALES ANULADOS"

   Dim obj As New RECORXET
   Set RS1 = obj.RS_REMITOSOFICIALES(2)
   Call Carga_MSF_Recordset(RS1, Campos$, FRMZELECHO.msf2, 1)
   Set obj = Nothing
   RS1.Close
   Set RS1 = Nothing

   FRMZELECHO.Width = 10000
   FRMZELECHO.Show 1
   NROREMI$ = TRIM$(RESP_ZELE_VECT(2))
   '
   If NROREMI$ = "" Then Exit Sub
   Call ACTUREGISTRO("MOVIOFI", "STATUS", "REMIOFI = '" & NROREMI$ & "'", 0, REGAF&, "NOMESS")
   Call ACTUREGISTRO("MOVIOFI", "REFERENCIA", "REMIOFI = '" & NROREMI$ & "'", Mid$(RESP_ZELE_VECT(3), 7), REGAF&, "NOMESS")

   GoTo 10

End Sub

Private Sub mnuRevertirSeleccion_Click()
   mnuRevertirSeleccion.Enabled = False
   Screen.MousePointer = 11
   Fin& = GRID.Rows - 1
   GRID.Redraw = False
   For i& = 1 To Fin&
       Call TRACE(20, i&, Fin&)
       Call COLOREAR_GRILLA_SOLAFILA(Me.GRID, &HFFFF80, i&)
   Next i&
   GRID.Redraw = True
   Screen.MousePointer = 1
   mnuRevertirSeleccion.Enabled = True
End Sub


Private Sub mnuVerRemiof_Click()
10 Campos$ = "Fecha/D8;Remito Of./A18;Referencia/A48;ID/A0"
   Call CARGA_ENCABEZADO(FRMZELECHO.msf2, Campos$, FRMZELECHO)
   Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
   FRMZELECHO.Caption = "REMITOS OFICIALES GENERADOS"

   Dim obj As New RECORXET
   Set RS1 = obj.RS_REMITOSOFICIALES(0)
   Call Carga_MSF_Recordset(RS1, Campos$, FRMZELECHO.msf2, 1)
   Set obj = Nothing
   RS1.Close
   Set RS1 = Nothing

   FRMZELECHO.Width = 10000
   FRMZELECHO.Show 1
   NROREMI$ = TRIM$(RESP_ZELE_VECT(2))
   '
   If NROREMI$ = "" Then Exit Sub
   
   DOCUNU& = 0
   TIDOCUM$ = "REMITOFI"

   NUDOCU$ = TRIM$(Str$(NPX&))
   Call MUESTRADOC(DOCUNU&, "", TIDOCUM$, SAFETEXT$(XVALO(Mid$(NROREMI$, 11))), NROREMI$)
   GoTo 10

End Sub

Private Sub mnuVincularReportes_Click()
    Call CONECTA_CONSUL_PRE(TRIM$(UCase$(Name)))
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


Private Sub Text3_Change()
    '
'    Text2 = ""
''    Text3 = ""
'    'Text4 = ""
'    Text6 = ""
    'Text8 = ""
    Text12 = ""
    Text15 = ""
    Text18 = ""
    Text19 = ""
    Text13 = ""
    Text20 = ""
    Text23 = ""
    Me.LENTREGA(0) = ""
    Me.LENTREGA(1) = ""
    Me.LENTREGA(2) = ""
    Text12 = ""
    Text18 = ""
    Me.DOMITRAN(3) = ""
    Me.DOMITRAN(4) = ""
    Me.Text45 = ""
    Text7 = ""
    If Command1.Enabled = True Then
      GRID.Rows = 1
      Text17 = ""
      Label18 = ""
    End If
    '
    CMO$ = "VM"
    '

    If Me.Text3 <> "" Then
        FF% = FECHANUM(Text4)
        CONDI$ = "(CodiMovi = '" & CMO$ & "' )   "
        CONDI$ = CONDI$ + " AND DoPriLAR ='" & Text3 & "' AND CANTSALID > 0 "
        VUELTA% = 0: J& = 0
        SQLX$ = "SELECT COD_INTE,NuorItem,DESCRI_LAR, CantSalid, cantingre, NUPEDCLI, ValoUnitIva, Nume_Clie, FECHMOV,IDENLOTE FROM MOVISTO WITH(NOLOCK)"
        SQLX$ = SQLX$ + " WHERE " & CONDI$ & "  ORDER BY NuOrItem"
        Set rs = READSET(SQLX$)
        With rs
          Do While Not .EOF
             If VUELTA% = 0 Then
               VUELTA% = 1
               Text2 = XVALO(!NUME_CLIE)
               Text6 = rasocli$(XVALO(Text2))
               Text4 = FECHATEX$(CVINT(!FechMov))
             End If
             J& = GRID.Rows
             GRID.Rows = J& + 1
             GRID.TextMatrix(J&, 1) = XVALO(!COD_INTE)
             GRID.TextMatrix(J&, 2) = J&
             GRID.TextMatrix(J&, 3) = SAFETEXT$(!DESCRI_LAR)
             GRID.TextMatrix(J&, 4) = TRIM$(FORMATNUM$(XVALO(!CantSalid), "F9.1"))
             GRID.TextMatrix(J&, 5) = TRIM$(FORMATNUM$(XVALO(!CANTINGRE), "F9.1"))
             GRID.TextMatrix(J&, 6) = XVALO(!NUPEDCLI)
             GRID.TextMatrix(J&, 7) = TRIM$(FORMATNUM$(XVALO(!VALOUNITIVA), "F12.2"))
             IMPORTEX = 0
             If TRIM$(GRID.TextMatrix(J&, 0)) = "*" Then
               IMPORTEX = (XVALO(!CantSalid) - XVALO(!CANTINGRE)) * XVALO(!VALOUNITIVA)
             End If
             GRID.TextMatrix(J&, 8) = SAFETEXT$(!IDENLOTE)
             GRID.TextMatrix(J&, 9) = Text3
             ACUMULATOTAL = ACUMULATOTAL + IMPORTEX
            .MoveNext
          Loop
        End With
        rs.Close
        Set rs = Nothing
        Text17 = TRIM$(FORMATNUM$(XVALO(ACUMULATOTAL) + XVALO(Text17), "F12.2"))
        Label18 = GRID.TextMatrix(GRID.Rows - 1, 2)
        '
        NC1 = XVALO(Text2)
        For i& = 1 To Fin&
            Call TRACE(20, i&, Fin&)
            Call COLOREAR_GRILLA_SOLAFILA(Me.GRID, &HFFFF80, i&)
        Next i&
        
'        If Command1.Enabled = True Then ' SOLO SI ES POR UN SOLO REMITO
'            'BUSCO LAS FACTURAS ASOCIADAS LO HAGO APARTE POR EL ORDEN
'            CONDI$ = CONDI$ + " ORDER BY DoSecLar"
'            J& = 0: FACTASOCI$ = ""
'            SQLX$ = "SELECT DoSecLar FROM MOVISTO WITH(NOLOCK)"
'            SQLX$ = SQLX$ + " WHERE " & CONDI$
'            Set RS = READSET(SQLX$)
'            With RS
'               Do While Not .EOF
'                NUFACT$ = SAFETEXT$(!DOSECLAR)
'                If NUFACT$ <> NUFACT_ANT$ Then
'                  FACTASOCI$ = FACTASOCI$ + NUFACT$ + "; "
'                End If
'                NUFACT_ANT$ = NUFACT$
'                .MoveNext
'               Loop
'            End With
'            PFIN% = InStrRev(FACTASOCI$, "; ")
'            If PFIN% > 0 Then FACTASOCI$ = Mid(FACTASOCI$, 1, Len(FACTASOCI$) - 2)
'            Text7 = Text7 + " " + FACTASOCI$
'        Else
'           Text7 = "Varias"
'        End If
        '
        'MOSTRAR DATOS DE ENTREGA
             SQLX$ = "SELECT DOMI_CLI,LOCA_CLI,CPOS_CLI,NTRANS_CLI FROM DEUDOR12 WITH(NOLOCK)"
             SQLX$ = SQLX$ + " WHERE NUME_CLI = " & NC1
             Set rs = READSET(SQLX$)
             With rs
             Do While Not .EOF
              
                 LENTREGA(0) = SAFETEXT(!DOMI_CLI)
                 LENTREGA(1) = SAFETEXT(!LOCA_CLI)
                 NTRAN% = XVALO(!NTRANS_CLI)
                 If NTRAN% > 0 Then ' SI VIENE CON CODIGO DE TRANSPORTE COMPLETA AUTOMATICAMENTE
                   Text45 = NTRAN%
                 End If
               .MoveNext
             Loop
             End With
20           rs.Close
             Set rs = Nothing
     End If

     Call mnuRevertirSeleccion_Click
End Sub


Sub CalcutotalSeleccionado()
        Text17 = "": ACUMULATOTAL = 0
        For i& = 1 To GRID.Rows - 1
            IMPORTEX = 0
            If TRIM$(GRID.TextMatrix(i&, 0)) = "*" Then
              IMPORTEX = (XVALO(GRID.TextMatrix(i&, 4)) - XVALO(GRID.TextMatrix(i&, 5))) * XVALO(GRID.TextMatrix(i&, 7))
            End If
            ACUMULATOTAL = ACUMULATOTAL + IMPORTEX
        Next i&
        Text17 = FORMATNUM$(ACUMULATOTAL, "F12.2")
End Sub
Sub LIMPIARTEXT(FORMULARIO As Form, LIMPIOTEXT%)
  Dim CONTROLX As CONTROL
  For Each CONTROLX In FORMULARIO.Controls ' RECORRO LOS CONTROLES DEL FORM ACTIVO
        If TypeOf CONTROLX Is TextBox And LIMPIOTEXT% > 0 Then
           CONTROLX.TEXT = ""
        End If
  Next
End Sub

Private Sub Text45_Change()
  Me.LENTREGA(2) = ""
  Text12 = ""
  Text18 = ""
  Me.DOMITRAN(3) = ""
  Me.DOMITRAN(4) = ""
  '
  NTRA = XVALO(Text45)
  If NTRA > 0 Then
    If NTRA < 32767 Then
      NTRAI% = NTRA
      CONDI$ = "CODI_TRA = " & NTRAI%
       If CantRegistros("TRANSPORTES", CONDI$) > 0 Then
        Text12 = CUITTRANS$(NTRAI%)
        LENTREGA(2) = DENOTRANS$(NTRAI%)
        DOMITRAN(3) = DOMITRANS$(NTRAI%)
        DOMITRAN(4) = LOCATRANS$(NTRAI%)
      End If
    End If
  End If

End Sub


