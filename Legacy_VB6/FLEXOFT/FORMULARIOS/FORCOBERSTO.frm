VERSION 5.00
Begin VB.Form FORCOBERSTO 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Análisis de Cobertura de Stocks"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   10650
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   10650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Consultas - Listados"
      Height          =   4995
      Left            =   4515
      TabIndex        =   16
      Top             =   210
      Width           =   4950
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
         Height          =   285
         Left            =   210
         TabIndex        =   37
         TabStop         =   0   'False
         Text            =   "Maestro de Artículos - Sin Validación"
         Top             =   2415
         Width           =   3330
      End
      Begin VB.CommandButton Command2 
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
         Left            =   3360
         TabIndex        =   36
         Top             =   1995
         Width           =   175
      End
      Begin VB.TextBox Text6 
         Height          =   360
         Left            =   2940
         TabIndex        =   35
         Text            =   " "
         Top             =   1995
         Width           =   435
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Tipo de Reporte"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1230
         Left            =   210
         TabIndex        =   31
         Top             =   2835
         Width           =   4530
         Begin VB.OptionButton Option3 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Generar Estimado Mensual de Entrega"
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
            Left            =   735
            TabIndex        =   34
            Top             =   840
            Width           =   3585
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Analisis de Cobertura de Stocks"
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
            Left            =   735
            TabIndex        =   33
            Top             =   250
            Value           =   -1  'True
            Width           =   3585
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Necesidades de Reposición"
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
            Left            =   735
            TabIndex        =   32
            Top             =   550
            Width           =   3585
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Salida"
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
         Left            =   210
         TabIndex        =   28
         Top             =   4200
         Width           =   4530
         Begin VB.OptionButton Option4 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Impresión"
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
            Left            =   2625
            TabIndex        =   30
            Top             =   250
            Value           =   -1  'True
            Width           =   1275
         End
         Begin VB.OptionButton Option5 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Pantalla"
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
            Left            =   735
            TabIndex        =   29
            Top             =   250
            Width           =   1380
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Meses"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   810
         Left            =   3885
         TabIndex        =   26
         ToolTipText     =   "Meses de Cobertura Requeridos para Necesidad de Reposición"
         Top             =   1050
         Width           =   855
         Begin VB.TextBox Text5 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   210
            TabIndex        =   27
            Text            =   "6"
            Top             =   315
            Width           =   435
         End
      End
      Begin VB.TextBox Text1 
         Height          =   360
         Left            =   2940
         TabIndex        =   23
         Text            =   " "
         Top             =   315
         Width           =   435
      End
      Begin VB.CommandButton Command7 
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
         Left            =   3360
         TabIndex        =   22
         Top             =   315
         Width           =   175
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
         Height          =   285
         Left            =   210
         TabIndex        =   21
         TabStop         =   0   'False
         Text            =   " Todos - Sin Clasificación"
         Top             =   735
         Width           =   3330
      End
      Begin VB.TextBox Text3 
         Height          =   360
         Left            =   2940
         TabIndex        =   20
         Text            =   " "
         Top             =   1155
         Width           =   435
      End
      Begin VB.CommandButton Command5 
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
         Left            =   3360
         TabIndex        =   19
         Top             =   1155
         Width           =   175
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
         Height          =   285
         Left            =   210
         TabIndex        =   18
         TabStop         =   0   'False
         Text            =   "Todos - Sin Clasificación"
         Top             =   1575
         Width           =   3330
      End
      Begin VB.CommandButton Command4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   3885
         Picture         =   "FORCOBERSTO.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   1995
         Width           =   855
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Lista de Precios de Validación:"
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
         Left            =   105
         TabIndex        =   38
         Top             =   2100
         Width           =   2745
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Material:"
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
         Left            =   525
         TabIndex        =   25
         Top             =   420
         Width           =   2325
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Origen Nacional / Importado:"
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
         Left            =   105
         TabIndex        =   24
         Top             =   1260
         Width           =   2745
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Generación - Período:"
      Height          =   1515
      Left            =   210
      TabIndex        =   9
      Top             =   315
      Width           =   4125
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Período:"
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
         Left            =   -1260
         TabIndex        =   15
         Top             =   1035
         Width           =   2220
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1080
         TabIndex        =   14
         Top             =   945
         Width           =   2775
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   320
         Left            =   1080
         TabIndex        =   11
         Top             =   440
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
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
         Index           =   13
         Left            =   -1260
         TabIndex        =   13
         Top             =   525
         Width           =   2220
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Hora:"
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
         Index           =   12
         Left            =   1380
         TabIndex        =   12
         Top             =   525
         Width           =   1590
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   320
         Left            =   3120
         TabIndex        =   10
         Top             =   440
         Width           =   735
      End
   End
   Begin VB.CommandButton Command6 
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
      Height          =   855
      Left            =   9630
      Picture         =   "FORCOBERSTO.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1260
      Width           =   855
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Recálculo - Actualización"
      Height          =   1455
      Left            =   210
      TabIndex        =   3
      Top             =   1995
      Width           =   4110
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   945
         ScaleHeight     =   585
         ScaleWidth      =   2175
         TabIndex        =   4
         Top             =   525
         Width           =   2235
         Begin VB.CommandButton Command3 
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
            Left            =   -20
            Picture         =   "FORCOBERSTO.frx":0614
            Style           =   1  'Graphical
            TabIndex        =   0
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Revisar Situa- ción de Stocks"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   735
            TabIndex        =   5
            Top             =   105
            Width           =   1590
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Control de Proceso"
      Height          =   1635
      Left            =   210
      TabIndex        =   2
      Top             =   3570
      Width           =   4110
      Begin VB.Frame Frame8 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Proceso Activo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   210
         TabIndex        =   39
         Top             =   420
         Width           =   3705
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
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
            Left            =   315
            TabIndex        =   40
            Top             =   315
            Width           =   3165
         End
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   210
         ScaleHeight     =   75
         ScaleWidth      =   3645
         TabIndex        =   7
         Top             =   1300
         Width           =   3705
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   8
            Top             =   0
            Width           =   12000
         End
      End
   End
   Begin VB.CommandButton Command1 
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
      Height          =   855
      Left            =   9630
      Picture         =   "FORCOBERSTO.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   315
      Width           =   855
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   9080
      Picture         =   "FORCOBERSTO.frx":0A68
      Top             =   4670
      Width           =   2010
   End
End
Attribute VB_Name = "FORCOBERSTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub


Private Sub Command2_Click()
    Call SELECHO("LISTAS")
    Text6.TEXT = VALACT1$("LISTAS")
    If TRIM$(Text6.TEXT) = "" Then
        Text7.TEXT = "Maestro de Articulos - Sin Validación"
    End If
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    Call COBERSTO
    Command3.Enabled = True
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    If Option2.Value = True Then
         MODOCOBER% = 1
         Call ANACOBER
       ElseIf Option1.Value = True Then
         MODOCOBER% = 2
         Call ANACOBER
       ElseIf Option3.Value = True Then
         MODOCOBER% = 3
         Call GEPROGEN
    End If
    Command4.Enabled = True
End Sub

Private Sub Command6_Click()
    Call HELPONLINE("COBERSTO")
End Sub

Private Sub Command5_Click()
   Call SELECHO("CLASI004")
   Text3.TEXT = VALACT1$("CLASI004")
   If TRIM$(Text3.TEXT) = "" Then
      Text4.TEXT = "Todos - Sin Clasificación"
   End If
End Sub

Private Sub Command7_Click()
   Call SELECHO("CLASI001")
   Text1.TEXT = VALACT1$("CLASI001")
   If TRIM$(Text1.TEXT) = "" Then
      Text2.TEXT = "Todos - Sin Clasificación"
   End If
End Sub

Private Sub Form_Load()
   Label1.Caption = TRIM$(CORRAR$("COBERST"))
   FEX = FEGEAR%("COBERST")
   Label21.Caption = FECHATEX$(FEX)
   Label22.Caption = HORATEX$(HOGEAR%("COBERST"))
End Sub

Private Sub Text1_Change()
   If Val(Text1.TEXT) < 256 Then
       TIPOMA% = Val(Text1.TEXT)
       If TIPOMA% > 0 Then
          TIPOMAX$ = ECOARCH$("CLASI001", Chr$(TIPOMA%))
          If TIPOMAX$ <> "" Then
             Text2.TEXT = TIPOMAX$
             Exit Sub
          End If
       End If
       Text1.TEXT = ""
       Text2.TEXT = "Todos - Sin Clasificación"
   End If
End Sub

Private Sub Text1_DblClick()
    Call SELECHO("CLASI001")
    Text1.TEXT = VALACT1$("CLASI001")
    If TRIM$(Text1.TEXT) = "" Then
       Text2.TEXT = "Todos - Sin Clasificación"
    End If
End Sub

Private Sub Text1_LostFocus()
    If Val(Text1.TEXT) < 256 Then
       TIPOMA% = Val(Text1.TEXT)
       If TIPOMA% > 0 Then
          TIPOMAX$ = ECOARCH$("CLASI001", Chr$(TIPOMA%))
          If TIPOMAX$ <> "" Then
             Text2.TEXT = TIPOMAX$
             Exit Sub
          End If
       End If
    End If
    Text1.TEXT = ""
    Text2.TEXT = "Todos - Sin Clasificación"
End Sub

Private Sub Text2_CLICK()
    Text1.SetFocus
End Sub

Private Sub Text3_Change()
   If Val(Text3.TEXT) < 256 Then
       ORIGMA% = Val(Text3.TEXT)
       If ORIGMA% > 0 Then
          ORIGMAX$ = ECOARCH$("CLASI004", Chr$(ORIGMA%))
          If ORIGMAX$ <> "" Then
             Text4.TEXT = ORIGMAX$
             Exit Sub
          End If
       End If
       Text3.TEXT = ""
       Text4.TEXT = "Todos - Sin Clasificación"
   End If
End Sub

Private Sub Text3_DblClick()
    Call SELECHO("CLASI004")
    Text3.TEXT = VALACT1$("CLASI004")
    If TRIM$(Text3.TEXT) = "" Then
       Text4.TEXT = "Todos - Sin Clasificación"
    End If
End Sub

Private Sub Text3_LostFocus()
    If Val(Text3.TEXT) < 256 Then
       ORIGMA% = Val(Text3.TEXT)
       If ORIGMA% > 0 Then
          ORIGMAX$ = ECOARCH$("CLASI004", Chr$(ORIGMA%))
          If ORIGMAX$ <> "" Then
             Text4.TEXT = ORIGMAX$
             Exit Sub
          End If
       End If
    End If
    Text3.TEXT = ""
    Text4.TEXT = "Todos - Sin Clasificación"
End Sub

Private Sub Text4_CLICK()
    Text3.SetFocus
End Sub

Private Sub Text6_Change()
   If Val(Text6.TEXT) < 256 Then
       LIPREVA% = Val(Text6.TEXT)
       If LIPREVA% > 0 Then
          LIPREVX$ = ECOARCH$("LISTAS", Chr$(LIPREVA%))
          If LIPREVX$ <> "" Then
             Text7.TEXT = LIPREVX$
             Exit Sub
          End If
       End If
       Text6.TEXT = ""
       Text7.TEXT = "Maestro de Articulos - Sin Validación"
   End If
End Sub

Private Sub Text6_DblClick()
    Call SELECHO("LISTAS")
    Text6.TEXT = VALACT1$("LISTAS")
    If TRIM$(Text6.TEXT) = "" Then
        Text7.TEXT = "Maestro de Articulos - Sin Validación"
    End If
End Sub

Private Sub Text6_LostFocus()
    If Val(Text6.TEXT) < 256 Then
       LIPREVA% = Val(Text6.TEXT)
       If LIPREVA% > 0 Then
          LIPREVX$ = ECOARCH$("LISTAS", Chr$(LIPREVA%))
          If LIPREVX$ <> "" Then
             Text7.TEXT = LIPREVX$
             Exit Sub
          End If
       End If
    End If
    Text6.TEXT = ""
    Text7.TEXT = "Maestro de Articulos - Sin Validación"
End Sub

Private Sub Text7_CLICK()
    Text6.SetFocus
End Sub
