VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FRMMAILS 
   BackColor       =   &H00E2D3C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Datos de Contactos"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   13785
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FRMMAILS.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   13785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.HScrollBar HScroll2 
      Height          =   255
      Left            =   10
      Max             =   370
      TabIndex        =   78
      Top             =   4165
      Width           =   12800
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00E2D3C0&
      Height          =   3700
      Left            =   10
      ScaleHeight     =   3645
      ScaleWidth      =   12720
      TabIndex        =   9
      Top             =   480
      Width           =   12780
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00E2D3C0&
         Height          =   3735
         Left            =   0
         ScaleHeight     =   3675
         ScaleWidth      =   16395
         TabIndex        =   10
         Top             =   0
         Width           =   16455
         Begin VB.Frame Frame8 
            BackColor       =   &H00E2D3C0&
            Caption         =   "CONTACTOS"
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
            Height          =   3375
            Left            =   120
            TabIndex        =   11
            Top             =   100
            Width           =   16140
            Begin VB.Frame Frame14 
               BackColor       =   &H00E2D3C0&
               BorderStyle     =   0  'None
               Caption         =   "Frame14"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2745
               Left            =   45
               TabIndex        =   12
               Top             =   510
               Width           =   15960
               Begin VB.Frame Frame15 
                  BackColor       =   &H00E2D3C0&
                  BorderStyle     =   0  'None
                  Caption         =   "Frame15"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   2805
                  Left            =   0
                  TabIndex        =   13
                  Top             =   0
                  Width           =   15855
                  Begin VB.PictureBox Picture5 
                     BackColor       =   &H00E2D3C0&
                     BorderStyle     =   0  'None
                     Height          =   2700
                     Index           =   4
                     Left            =   13920
                     ScaleHeight     =   2700
                     ScaleWidth      =   2040
                     TabIndex        =   54
                     Top             =   0
                     Width           =   2040
                     Begin VB.TextBox Text30 
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
                        Height          =   285
                        Index           =   8
                        Left            =   105
                        TabIndex        =   63
                        Top             =   2400
                        Width           =   1700
                     End
                     Begin VB.TextBox Text30 
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
                        Height          =   285
                        Index           =   7
                        Left            =   105
                        TabIndex        =   62
                        Top             =   2100
                        Width           =   1700
                     End
                     Begin VB.TextBox Text30 
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
                        Height          =   285
                        Index           =   6
                        Left            =   105
                        TabIndex        =   61
                        Top             =   1800
                        Width           =   1700
                     End
                     Begin VB.TextBox Text30 
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
                        Height          =   285
                        Index           =   5
                        Left            =   105
                        TabIndex        =   60
                        Top             =   1500
                        Width           =   1700
                     End
                     Begin VB.TextBox Text30 
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
                        Height          =   285
                        Index           =   4
                        Left            =   105
                        TabIndex        =   59
                        Top             =   1200
                        Width           =   1700
                     End
                     Begin VB.TextBox Text30 
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
                        Height          =   285
                        Index           =   3
                        Left            =   105
                        TabIndex        =   58
                        Top             =   900
                        Width           =   1700
                     End
                     Begin VB.TextBox Text30 
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
                        Height          =   285
                        Index           =   0
                        Left            =   105
                        TabIndex        =   57
                        Top             =   0
                        Width           =   1695
                     End
                     Begin VB.TextBox Text30 
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
                        Height          =   285
                        Index           =   1
                        Left            =   105
                        TabIndex        =   56
                        Top             =   300
                        Width           =   1700
                     End
                     Begin VB.TextBox Text30 
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
                        Height          =   285
                        Index           =   2
                        Left            =   105
                        TabIndex        =   55
                        Top             =   600
                        Width           =   1700
                     End
                  End
                  Begin VB.PictureBox Picture5 
                     BackColor       =   &H00E2D3C0&
                     BorderStyle     =   0  'None
                     Height          =   2700
                     Index           =   3
                     Left            =   11200
                     ScaleHeight     =   2700
                     ScaleWidth      =   2835
                     TabIndex        =   44
                     Top             =   0
                     Width           =   2835
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
                        Height          =   285
                        Index           =   8
                        Left            =   105
                        TabIndex        =   53
                        Top             =   2400
                        Width           =   2580
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
                        Height          =   285
                        Index           =   7
                        Left            =   105
                        TabIndex        =   52
                        Top             =   2100
                        Width           =   2580
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
                        Height          =   285
                        Index           =   6
                        Left            =   105
                        TabIndex        =   51
                        Top             =   1800
                        Width           =   2580
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
                        Height          =   285
                        Index           =   5
                        Left            =   105
                        TabIndex        =   50
                        Top             =   1500
                        Width           =   2580
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
                        Height          =   285
                        Index           =   4
                        Left            =   105
                        TabIndex        =   49
                        Top             =   1200
                        Width           =   2580
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
                        Height          =   285
                        Index           =   3
                        Left            =   105
                        TabIndex        =   48
                        Top             =   900
                        Width           =   2580
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
                        Height          =   285
                        Index           =   0
                        Left            =   105
                        TabIndex        =   47
                        Top             =   0
                        Width           =   2580
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
                        Height          =   285
                        Index           =   1
                        Left            =   105
                        TabIndex        =   46
                        Top             =   300
                        Width           =   2580
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
                        Height          =   285
                        Index           =   2
                        Left            =   105
                        TabIndex        =   45
                        Top             =   600
                        Width           =   2580
                     End
                  End
                  Begin VB.PictureBox Picture5 
                     BackColor       =   &H00E2D3C0&
                     BorderStyle     =   0  'None
                     Height          =   2700
                     Index           =   2
                     Left            =   5760
                     ScaleHeight     =   2700
                     ScaleWidth      =   5535
                     TabIndex        =   34
                     Top             =   0
                     Width           =   5535
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
                        Height          =   285
                        Index           =   8
                        Left            =   105
                        TabIndex        =   43
                        Top             =   2400
                        Width           =   5300
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
                        Height          =   285
                        Index           =   7
                        Left            =   105
                        TabIndex        =   42
                        Top             =   2100
                        Width           =   5300
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
                        Height          =   285
                        Index           =   6
                        Left            =   105
                        TabIndex        =   41
                        Top             =   1800
                        Width           =   5300
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
                        Height          =   285
                        Index           =   5
                        Left            =   105
                        TabIndex        =   40
                        Top             =   1500
                        Width           =   5300
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
                        Height          =   285
                        Index           =   4
                        Left            =   105
                        TabIndex        =   39
                        Top             =   1200
                        Width           =   5300
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
                        Height          =   285
                        Index           =   3
                        Left            =   105
                        TabIndex        =   38
                        Top             =   900
                        Width           =   5300
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
                        Height          =   285
                        Index           =   0
                        Left            =   105
                        TabIndex        =   37
                        Top             =   0
                        Width           =   5300
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
                        Height          =   285
                        Index           =   1
                        Left            =   105
                        TabIndex        =   36
                        Top             =   300
                        Width           =   5300
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
                        Height          =   285
                        Index           =   2
                        Left            =   105
                        TabIndex        =   35
                        Top             =   600
                        Width           =   5300
                     End
                  End
                  Begin VB.PictureBox Picture5 
                     BackColor       =   &H00E2D3C0&
                     BorderStyle     =   0  'None
                     Height          =   2700
                     Index           =   1
                     Left            =   3180
                     ScaleHeight     =   2700
                     ScaleWidth      =   2655
                     TabIndex        =   24
                     Top             =   0
                     Width           =   2655
                     Begin VB.TextBox Text7 
                        Appearance      =   0  'Flat
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
                        Index           =   2
                        Left            =   105
                        TabIndex        =   33
                        Top             =   600
                        Width           =   2450
                     End
                     Begin VB.TextBox Text7 
                        Appearance      =   0  'Flat
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
                        Index           =   1
                        Left            =   105
                        TabIndex        =   32
                        Top             =   300
                        Width           =   2450
                     End
                     Begin VB.TextBox Text7 
                        Appearance      =   0  'Flat
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
                        Index           =   0
                        Left            =   105
                        TabIndex        =   31
                        Top             =   0
                        Width           =   2450
                     End
                     Begin VB.TextBox Text7 
                        Appearance      =   0  'Flat
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
                        Index           =   3
                        Left            =   105
                        TabIndex        =   30
                        Top             =   900
                        Width           =   2450
                     End
                     Begin VB.TextBox Text7 
                        Appearance      =   0  'Flat
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
                        Left            =   105
                        TabIndex        =   29
                        Top             =   1200
                        Width           =   2450
                     End
                     Begin VB.TextBox Text7 
                        Appearance      =   0  'Flat
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
                        Index           =   5
                        Left            =   105
                        TabIndex        =   28
                        Top             =   1500
                        Width           =   2450
                     End
                     Begin VB.TextBox Text7 
                        Appearance      =   0  'Flat
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
                        Index           =   6
                        Left            =   105
                        TabIndex        =   27
                        Top             =   1800
                        Width           =   2450
                     End
                     Begin VB.TextBox Text7 
                        Appearance      =   0  'Flat
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
                        Index           =   7
                        Left            =   105
                        TabIndex        =   26
                        Top             =   2100
                        Width           =   2450
                     End
                     Begin VB.TextBox Text7 
                        Appearance      =   0  'Flat
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
                        Index           =   8
                        Left            =   105
                        TabIndex        =   25
                        Top             =   2400
                        Width           =   2450
                     End
                  End
                  Begin VB.PictureBox Picture5 
                     BackColor       =   &H00E2D3C0&
                     BorderStyle     =   0  'None
                     Height          =   2700
                     Index           =   0
                     Left            =   350
                     ScaleHeight     =   2700
                     ScaleWidth      =   2880
                     TabIndex        =   14
                     Top             =   0
                     Width           =   2880
                     Begin VB.TextBox Text6 
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
                        Height          =   285
                        Index           =   2
                        Left            =   0
                        TabIndex        =   15
                        Top             =   600
                        Width           =   2800
                     End
                     Begin VB.TextBox Text6 
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
                        Height          =   285
                        Index           =   0
                        Left            =   0
                        TabIndex        =   23
                        Top             =   0
                        Width           =   2800
                     End
                     Begin VB.TextBox Text6 
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
                        Height          =   285
                        Index           =   1
                        Left            =   0
                        TabIndex        =   22
                        Top             =   300
                        Width           =   2800
                     End
                     Begin VB.TextBox Text6 
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
                        Height          =   285
                        Index           =   3
                        Left            =   0
                        TabIndex        =   21
                        Top             =   900
                        Width           =   2800
                     End
                     Begin VB.TextBox Text6 
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
                        Height          =   285
                        Index           =   4
                        Left            =   0
                        TabIndex        =   20
                        Top             =   1200
                        Width           =   2800
                     End
                     Begin VB.TextBox Text6 
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
                        Height          =   285
                        Index           =   5
                        Left            =   0
                        TabIndex        =   19
                        Top             =   1500
                        Width           =   2800
                     End
                     Begin VB.TextBox Text6 
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
                        Height          =   285
                        Index           =   6
                        Left            =   0
                        TabIndex        =   18
                        Top             =   1800
                        Width           =   2800
                     End
                     Begin VB.TextBox Text6 
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
                        Height          =   285
                        Index           =   7
                        Left            =   0
                        TabIndex        =   17
                        Top             =   2100
                        Width           =   2800
                     End
                     Begin VB.TextBox Text6 
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
                        Height          =   285
                        Index           =   8
                        Left            =   0
                        TabIndex        =   16
                        Top             =   2400
                        Width           =   2800
                     End
                  End
                  Begin VB.Label LEFE 
                     Alignment       =   1  'Right Justify
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "3)"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   210
                     Index           =   2
                     Left            =   135
                     TabIndex        =   72
                     Top             =   630
                     Width           =   150
                  End
                  Begin VB.Label LEFE 
                     Alignment       =   1  'Right Justify
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "2)"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   210
                     Index           =   1
                     Left            =   135
                     TabIndex        =   71
                     Top             =   330
                     Width           =   150
                  End
                  Begin VB.Label LEFE 
                     Alignment       =   1  'Right Justify
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "1)"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   210
                     Index           =   0
                     Left            =   135
                     TabIndex        =   70
                     Top             =   45
                     Width           =   150
                  End
                  Begin VB.Label LEFE 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   "6)"
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
                     Index           =   6
                     Left            =   -1300
                     TabIndex        =   69
                     Top             =   1530
                     Width           =   1590
                  End
                  Begin VB.Label LEFE 
                     Alignment       =   1  'Right Justify
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "5)"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   210
                     Index           =   7
                     Left            =   135
                     TabIndex        =   68
                     Top             =   1230
                     Width           =   150
                  End
                  Begin VB.Label LEFE 
                     Alignment       =   1  'Right Justify
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "4)"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   210
                     Index           =   8
                     Left            =   135
                     TabIndex        =   67
                     Top             =   930
                     Width           =   150
                  End
                  Begin VB.Label LEFE 
                     Alignment       =   1  'Right Justify
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "9)"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   210
                     Index           =   9
                     Left            =   135
                     TabIndex        =   66
                     Top             =   2430
                     Width           =   150
                  End
                  Begin VB.Label LEFE 
                     Alignment       =   1  'Right Justify
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "8)"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   210
                     Index           =   10
                     Left            =   135
                     TabIndex        =   65
                     Top             =   2130
                     Width           =   150
                  End
                  Begin VB.Label LEFE 
                     Alignment       =   1  'Right Justify
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "7)"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   210
                     Index           =   11
                     Left            =   135
                     TabIndex        =   64
                     Top             =   1830
                     Width           =   150
                  End
               End
            End
            Begin VB.Label LabeL9 
               BackColor       =   &H00E2D3C0&
               BackStyle       =   0  'Transparent
               Caption         =   "Celular"
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
               Index           =   4
               Left            =   14050
               TabIndex        =   77
               Top             =   270
               Width           =   1005
            End
            Begin VB.Label LabeL9 
               BackColor       =   &H00E2D3C0&
               BackStyle       =   0  'Transparent
               Caption         =   "Telefono"
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
               Index           =   3
               Left            =   11380
               TabIndex        =   76
               Top             =   270
               Width           =   1995
            End
            Begin VB.Label LabeL9 
               BackColor       =   &H00E2D3C0&
               BackStyle       =   0  'Transparent
               Caption         =   "E-Mail"
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
               Index           =   2
               Left            =   5920
               TabIndex        =   75
               Top             =   270
               Width           =   3000
            End
            Begin VB.Label LabeL9 
               BackColor       =   &H00E2D3C0&
               BackStyle       =   0  'Transparent
               Caption         =   "Cargo / Función - Sector"
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
               Index           =   1
               Left            =   3350
               TabIndex        =   74
               Top             =   270
               Width           =   6000
            End
            Begin VB.Label LabeL9 
               BackColor       =   &H00E2D3C0&
               BackStyle       =   0  'Transparent
               Caption         =   "Persona de Contacto"
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
               Index           =   0
               Left            =   405
               TabIndex        =   73
               Top             =   240
               Width           =   6000
            End
         End
      End
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1215
      Left            =   8160
      TabIndex        =   8
      Top             =   1560
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   3720
      Width           =   12800
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   7680
      Top             =   0
   End
   Begin VB.PictureBox Picture999 
      Height          =   225
      Index           =   0
      Left            =   1470
      ScaleHeight     =   165
      ScaleWidth      =   1110
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00624E28&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8355
      Left            =   12840
      ScaleHeight     =   8295
      ScaleWidth      =   1035
      TabIndex        =   1
      Top             =   -120
      Width           =   1095
      Begin VB.CommandButton Command1 
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
         Height          =   600
         Left            =   120
         Picture         =   "FRMMAILS.frx":1272
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Guarda los Cambios Realizados"
         Top             =   920
         Width           =   690
      End
      Begin VB.CommandButton Command3 
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
         Height          =   600
         Left            =   120
         Picture         =   "FRMMAILS.frx":157C
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Cierra y Abandona la Aplicación"
         Top             =   300
         Width           =   690
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2625
      OleObjectBlob   =   "FRMMAILS.frx":16C6
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "FRMMAILS.frx":18FA
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label LabeL9 
      BackColor       =   &H00E2D3C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo de Cuenta"
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
      Index           =   5
      Left            =   9555
      TabIndex        =   3
      Top             =   180
      Width           =   1575
   End
   Begin VB.Label Label26 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   300
      Left            =   11205
      TabIndex        =   2
      Top             =   75
      Width           =   1575
   End
End
Attribute VB_Name = "FRMMAILS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    '
    command1.Enabled = False
    CODICLIX$ = TRIM$(Label26)
    '
    CODICLIX$ = TRIM$(Label26)
    CONDI$ = "CodiCli = '" & CODICLIX$ & "'"
    Call ACTUREGISTRO("CONTACSCOM", "Marborra", CONDI$, 1, RAFE&)
    '
    SQLX$ = "SELECT * FROM CONTACSCOM WHERE " + CONDI$
    Dim Fichanew As ADODB.Recordset
    Set Fichanew = New ADODB.Recordset
    Fichanew.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    For IKI% = 0 To Text6.Count - 1
        TEXTOX$ = Text6(IKI%).TEXT + Text7(IKI%).TEXT + Text8(IKI%).TEXT + Text10(IKI%).TEXT + Text30(IKI%)
        If TRIM$(TEXTOX$) <> "" Then
            With Fichanew
                .AddNew   ' YA CONVERTIDO
                On Error Resume Next
                !CodiEmpr = CodiEmp%
                On Error GoTo 0
                !CODICLI = CODICLIX$
                !ORDI_CTAC = IKI%
                !NOME_CTAC = Text6(IKI%)
                !CARG_CTAC = Text7(IKI%)
                !MAIL_CTAC = Text8(IKI%)
                !TELE_CTAC = Text10(IKI%)
                If TRIM$(Text30(IKI%)) = "" Then 'AGREGADO POR AHP, QUE TIENE UN CARGALISEL CON NOMBRE,TELEFONO Y MAIL
                     Text30(IKI%) = "-"            'SI EL DEL MEDIO NO TIENE INFORMACION NO SE RESPETAN LAS POSICIONES
                End If
                !CELU_CTAC = Text30(IKI%)
                !MARBORRA = 0
                .Update
            End With
        End If
    Next IKI%
    Fichanew.Close
    Set Fichanew = Nothing
    '
    CONDI1$ = CONDI$ + " and MARBORRA = 1"
    Call BORRAREGISTROS("CONTACSCOM", CONDI1$, RAFE&)
    '
    command1.Enabled = True
    Call Command3_Click ' salir
End Sub

Private Sub Command3_Click()
   Call CIERRARCH("*.*")
   Unload Me
End Sub

Private Sub Form_Load()
   UTILIZA_SKIN% = 1
   Call APLICACIONSKINS(Me, Picture3)
   VScroll1.Max = (Frame15.Height - VScroll1.Height) / 100
   VScroll1.Value = 0
   '
   Picture1.Top = 0
   HScroll1.Min = 0
   HScroll1.Max = (Picture1.Width - Picture2.Width) / 10

   For KKI% = 1 To Text6.Count
      Text6(KKI% - 1).FontSize = 8
      Text7(KKI% - 1).FontSize = 8
      Text8(KKI% - 1).FontSize = 8
      Text10(KKI% - 1).FontSize = 8
      Text30(KKI% - 1).FontSize = 8
   Next KKI%
   '
   Call ABRECLIEN
   '
   Call CREACAMPO("", "CONTACSCOM", "CODIEMPR", 3, 0)
   Call CREACAMPO("", "CONTACSCOM", "CODICLI", 10, 12)
   Call CREACAMPO("", "CONTACSCOM", "ORDI_CTAC", 3, 0)
   Call CREACAMPO("", "CONTACSCOM", "NOME_CTAC", 10, 64)
   Call CREACAMPO("", "CONTACSCOM", "CARG_CTAC", 10, 64)
   Call CREACAMPO("", "CONTACSCOM", "MAIL_CTAC", 10, 128)
   Call CREACAMPO("", "CONTACSCOM", "TELE_CTAC", 10, 64)
   Call CREACAMPO("", "CONTACSCOM", "CELU_CTAC", 10, 64)
   Call CREACAMPO("", "CONTACSCOM", "MARBORRA", 3, 0)
   '
   Timer1.Enabled = True
'   For KKI% = 1 To Text6.Count - 1
'      Text6(Index).ToolTipText = Text6(0).ToolTipText
'      Text6(KKI%).Width = 2800
'   Next KKI%
'   '
'   For KKI% = 1 To Text8.Count - 1
'      Text8(KKI%).ToolTipText = Text8(0).ToolTipText
'      Text8(KKI%).Width = 2020
'   Next KKI%
'
'   For KKI% = 1 To Text7.Count - 1
'      Text7(KKI%).Width = 2450
'   Next KKI%
'
'   For KKI% = 1 To Text10.Count - 1
'      Text10(KKI%).Width = 1330
'   Next KKI%
'
'   For KKI% = 1 To Text30.Count - 1
'      Text30(KKI%).Width = 1330
'   Next KKI%

End Sub


Private Sub HScroll2_Change()
   Picture1.Left = (-10) * HScroll2.Value
End Sub

Private Sub HScroll2_Scroll()
   Picture1.Left = (-10) * HScroll2.Value
End Sub

'Private Sub HScroll2_Scroll()
'
'    Picture1.Left = Picture1.Left - HScroll2.Value
'
'End Sub

Private Sub Label26_Change()
   '
   Call CARGAMAILS
   On Error Resume Next
   Call PINTATEXT(Text6(0))
   On Error GoTo 0
   '
End Sub


Private Sub Text6_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
      On Error Resume Next
      Call PINTATEXT(Text7(Index))
      On Error GoTo 0
    End If
End Sub
Private Sub TEXT7_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
      On Error Resume Next
      Call PINTATEXT(Text8(Index))
      On Error GoTo 0
    End If
End Sub
Private Sub TEXT8_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
      On Error Resume Next
      Call PINTATEXT(Text10(Index))
      On Error GoTo 0
    End If
End Sub
Private Sub TEXT10_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
      On Error Resume Next
      Picture1.Left = -2720
      Call PINTATEXT(Me.Text30(Index))
      On Error GoTo 0
    End If
End Sub
Private Sub Text30_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
      On Error Resume Next
      Call PINTATEXT(Text6(Index + 1))
      If Err > 0 Then
         Call PINTATEXT(Text6(0))
      End If
      Picture1.Left = 0
      On Error GoTo 0
    End If
End Sub


Private Sub Timer1_Timer()
  Call PINTATEXT(Text6(0))
  Timer1.Enabled = False
End Sub

Private Sub VScroll1_Change()
   Frame15.Top = (-1) * VScroll1.Value * 100
End Sub
Private Sub VScroll1_Scroll()
   Frame15.Top = (-1) * VScroll1.Value * 100
End Sub



Sub AGREGARENGLONES()
   '
   INTERVER = 25
   NUPIC% = Text6.Count
   Load Text6(NUPIC%)
   Text6(NUPIC%) = ""
   Text6(NUPIC%).Left = Text6(NUPIC% - 1).Left
   Text6(NUPIC%).Top = Text6(NUPIC% - 1).Top + Text6(NUPIC% - 1).Height + INTERVER
   Text6(NUPIC%).Visible = True
   '
   Load Text7(NUPIC%)
   Text7(NUPIC%) = ""
   Text7(NUPIC%).Left = Text7(NUPIC% - 1).Left
   Text7(NUPIC%).Top = Text7(NUPIC% - 1).Top + Text7(NUPIC% - 1).Height + INTERVER
   Text7(NUPIC%).Visible = True
   '
   Load Text8(NUPIC%)
   Text8(NUPIC%) = ""
   Text8(NUPIC%).Left = Text8(NUPIC% - 1).Left
   Text8(NUPIC%).Top = Text8(NUPIC% - 1).Top + Text8(NUPIC% - 1).Height + INTERVER
   Text8(NUPIC%).Visible = True
   '
   Load Text10(NUPIC%)
   Text10(NUPIC%) = ""
   Text10(NUPIC%).Left = Text10(NUPIC% - 1).Left
   Text10(NUPIC%).Top = Text10(NUPIC% - 1).Top + Text10(NUPIC% - 1).Height + INTERVER
   Text10(NUPIC%).Visible = True
   '
   Load Text30(NUPIC%)
   Text30(NUPIC%) = ""
   Text30(NUPIC%).Left = Text30(NUPIC% - 1).Left
   Text30(NUPIC%).Top = Text30(NUPIC% - 1).Top + Text30(NUPIC% - 1).Height + INTERVER
   Text30(NUPIC%).Visible = True
   '
   For KKI% = 0 To Picture5.Count - 1
      If Picture5(KKI%).Height < Text6(NUPIC%).Top + Text6(NUPIC%).Height Then
        Picture5(KKI%).Height = Text6(NUPIC%).Top + Text6(NUPIC%).Height
      End If
   Next KKI%
   If Frame15.Height < Picture5(0).Top + Picture5(0).Height Then
        Frame15.Height = Picture5(0).Top + Picture5(0).Height
   End If
   '
   VScroll1.Max = (Frame15.Height - VScroll1.Height) / 100
   '
End Sub


Sub CARGAMAILS()
   
   COCLI$ = TRIM$(Label26)
   
   SQLX$ = "SELECT * FROM CONTACSCOM"
   SQLX$ = SQLX$ + " WHERE CODICLI = '" & COCLI$ & "' ORDER BY ORDI_CTAC"
   Dim FICHCON As ADODB.Recordset
   Set FICHCON = FLEXCONN.Execute(FILTSQL$(SQLX$))
   With FICHCON
      Do While Not .EOF
        NOME$ = TRIM$(SAFETEXT$(!NOME_CTAC))
        CARG$ = TRIM$(SAFETEXT$(!CARG_CTAC))
        mail$ = TRIM$(SAFETEXT$(!MAIL_CTAC))
        TELE$ = TRIM$(SAFETEXT$(!TELE_CTAC))
        CELU$ = TRIM$(SAFETEXT$(!CELU_CTAC))
        '
'        For JJ2% = 1 To JJ1%
'           If UCase$(Text6(JJ2% - 1)) = UCase$(NOME$) Then
'              If Text7(JJ2% - 1) = "" Then Text7(JJ2% - 1) = CARG$
'              If Text8(JJ2% - 1) = "" Then Text8(JJ2% - 1) = mail$
'              If Text10(JJ2% - 1) = "" Then Text10(JJ2% - 1) = TELE$
'              If Text30(JJ2% - 1) = "" Then Text30(JJ2% - 1) = CELU$
'              GoTo 19
'           End If
'        Next JJ2%
        '
15      If JJ1% >= Text6.Count Then
           Call AGREGARENGLONES
           GoTo 15
        End If
        '
        Text6(JJ1%) = NOME$
        Text7(JJ1%) = CARG$
        Text8(JJ1%) = mail$
        Text10(JJ1%) = TELE$
        Text30(JJ1%) = CELU$
        If Text6(JJ1%) + Text7(JJ1%) + Text8(JJ1%) + Text10(JJ1%) + Text30(JJ1%) <> "" Then
           JJ1% = JJ1% + 1
16         If JJ1% > Text6.Count Then
              Call AGREGARENGLONES
              GoTo 16
           End If
        End If
        '
19    .MoveNext
      Loop
   End With
   FICHCON.Close
   Set FICHCON = Nothing
End Sub



