VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form RESCUC07 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E2D3C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Resúmenes de Cuenta Corriente - Clientes"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   8910
   ControlBox      =   0   'False
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
   ScaleHeight     =   3435
   ScaleWidth      =   8910
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture999 
      Height          =   330
      Index           =   0
      Left            =   840
      ScaleHeight     =   270
      ScaleWidth      =   1110
      TabIndex        =   34
      Top             =   360
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E2D3C0&
      Caption         =   "COMPOSICIÓN DE DEUDA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   3960
      TabIndex        =   32
      Top             =   2530
      Width           =   4005
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2D3C0&
         Caption         =   "Reconstrucción Automática"
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
         Height          =   240
         Left            =   630
         TabIndex        =   33
         Top             =   400
         Width           =   2955
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00624E28&
      ForeColor       =   &H80000008&
      Height          =   3900
      Left            =   8040
      ScaleHeight     =   3870
      ScaleWidth      =   1575
      TabIndex        =   29
      Top             =   0
      Width           =   1605
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H002B9973&
         Height          =   200
         Left            =   120
         ScaleHeight     =   135
         ScaleWidth      =   585
         TabIndex        =   36
         Top             =   3000
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   -120
            TabIndex        =   37
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         Picture         =   "RESCUC07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Abandona la Aplicación"
         Top             =   260
         Width           =   650
      End
      Begin VB.CommandButton Command2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   120
         Picture         =   "RESCUC07.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Emite Reporte Solicitado"
         Top             =   2160
         Width           =   650
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E2D3C0&
      Caption         =   "RANGO DE FECHAS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1170
      Left            =   120
      TabIndex        =   22
      Top             =   2160
      Width           =   2115
      Begin VB.TextBox Text5 
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
         Left            =   630
         TabIndex        =   27
         Top             =   735
         Width           =   1065
      End
      Begin VB.CommandButton Command10 
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
         Left            =   1680
         TabIndex        =   26
         Top             =   735
         Width           =   175
      End
      Begin VB.TextBox Text6 
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
         Left            =   630
         TabIndex        =   24
         Top             =   420
         Width           =   1065
      End
      Begin VB.CommandButton Command6 
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
         Left            =   1680
         TabIndex        =   23
         Top             =   420
         Width           =   175
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Hasta"
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
         Left            =   -315
         TabIndex        =   28
         Top             =   790
         Width           =   855
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Desde"
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
         Left            =   -1785
         TabIndex        =   25
         Top             =   480
         Width           =   2325
      End
   End
   Begin VB.ListBox List1 
      Height          =   540
      Left            =   8505
      Sorted          =   -1  'True
      TabIndex        =   21
      Top             =   2730
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E2D3C0&
      Caption         =   "RANGO DE CUENTAS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2010
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   3765
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
         Left            =   2505
         TabIndex        =   18
         Top             =   315
         Width           =   855
      End
      Begin VB.CommandButton Command9 
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
         Left            =   3345
         TabIndex        =   17
         Top             =   315
         Width           =   175
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
         Left            =   195
         TabIndex        =   16
         TabStop         =   0   'False
         Text            =   " "
         Top             =   690
         Width           =   3330
      End
      Begin VB.TextBox Text8 
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
         Left            =   2505
         TabIndex        =   15
         Text            =   " "
         Top             =   1095
         Width           =   855
      End
      Begin VB.CommandButton Command8 
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
         Left            =   3345
         TabIndex        =   14
         Top             =   1095
         Width           =   175
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
         Height          =   300
         Left            =   195
         TabIndex        =   13
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1470
         Width           =   3330
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Desde"
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
         Left            =   90
         TabIndex        =   20
         Top             =   420
         Width           =   2325
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Hasta"
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
         Left            =   90
         TabIndex        =   19
         Top             =   1200
         Width           =   2325
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2D3C0&
      Caption         =   "FILTROS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2380
      Left            =   3960
      TabIndex        =   3
      Top             =   120
      Width           =   4005
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
         Left            =   315
         TabIndex        =   10
         TabStop         =   0   'False
         Text            =   "Todos - Sin Clasificación"
         Top             =   1890
         Width           =   3330
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
         Height          =   300
         Left            =   3465
         TabIndex        =   9
         Top             =   1470
         Width           =   175
      End
      Begin VB.TextBox Text3 
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
         Left            =   3045
         TabIndex        =   8
         Text            =   " "
         Top             =   1470
         Width           =   435
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
         Left            =   315
         TabIndex        =   6
         TabStop         =   0   'False
         Text            =   " Todos - Sin Clasificación"
         Top             =   840
         Width           =   3330
      End
      Begin VB.CommandButton Command3 
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
         Left            =   3465
         TabIndex        =   5
         Top             =   420
         Width           =   175
      End
      Begin VB.TextBox Text1 
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
         Left            =   3045
         TabIndex        =   4
         Text            =   " "
         Top             =   420
         Width           =   435
      End
      Begin VB.Label Label1 
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
         Height          =   225
         Left            =   630
         TabIndex        =   11
         Top             =   1575
         Width           =   2325
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo de Cliente"
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
         Left            =   630
         TabIndex        =   7
         Top             =   525
         Width           =   2325
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E2D3C0&
      Caption         =   "ORDENAMIENTO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1170
      Left            =   2280
      TabIndex        =   0
      Top             =   2160
      Width           =   1600
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2D3C0&
         Caption         =   "Alfabético"
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
         Height          =   240
         Index           =   2
         Left            =   150
         TabIndex        =   2
         Top             =   735
         Width           =   1380
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2D3C0&
         Caption         =   "Numérico"
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
         Height          =   240
         Index           =   1
         Left            =   150
         TabIndex        =   1
         Top             =   420
         Value           =   -1  'True
         Width           =   1380
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2625
      OleObjectBlob   =   "RESCUC07.frx":0454
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "RESCUC07.frx":0688
      TabIndex        =   35
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
End
Attribute VB_Name = "RESCUC07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub
'
Private Sub Command10_Click()
    Call SELECHO("SELFECHA")
    If FECHANUM(VALACT1$("SELFECHA")) > 0 Then
        Text5.TEXT = VALACT1$("SELFECHA")
    End If
End Sub

Private Sub Command2_Click()
    CONTROLGRABA% = 0
    If Check1.Value = 1 Then CONTROLGRABA% = 1
    '
    Hide
    Call FRESHALL
    Call GERESCUECLI
    CONTROLGRABA% = 0
End Sub
'
Private Sub Command3_Click()
    Call SELECHO("GRUCOCLI")
    TEXT1.TEXT = VALACT1$("GRUCOCLI")
    If TRIM$(TEXT1.TEXT) = "" Then
        Text2.TEXT = "Todos - Sin Clasificación"
    End If
End Sub
'
Private Sub Command4_Click()
    Call SELECHO("VENDEDOR")
    Text3.TEXT = VALACT1$("VENDEDOR")
    If TRIM$(Text3.TEXT) = "" Then
        Text4.TEXT = "Todos - Sin Clasificación"
    End If
End Sub


Private Sub Command6_Click()
    Call SELECHO("SELFECHA")
    If FECHANUM(VALACT1$("SELFECHA")) > 0 Then
        Text6.TEXT = VALACT1$("SELFECHA")
    End If
End Sub

Private Sub Command8_Click()
    NCMAXI = ListaSeleccionClientes()
    If NCMAXI > 0 Then
        Text8.TEXT = TRIM$(Str$(NCMAXI))
        Text7.TEXT = rasocli$(NCMAXI)
    End If

'    Call SELECHO("DEUDOR1")
'    NCMAXI = XVALO(VALACT1$("DEUDOR1"))
'    If NCMAXI > 0 Then
'        Text8.TEXT = TRIM$(Str$(NCMAXI))
'        Text7.TEXT = rasocli$(NCMAXI)
'    End If
End Sub

Private Sub Command9_Click()

    NCMINI = ListaSeleccionClientes()
    If NCMINI > 0 Then
        Text10.TEXT = TRIM$(Str$(NCMINI))
        Text9.TEXT = rasocli$(NCMINI)
    End If
    
'    Call SELECHO("DEUDOR1")
'    NCMINI = XVALO(VALACT1$("DEUDOR1"))
'    If NCMINI > 0 Then
'        Text10.TEXT = TRIM$(Str$(NCMINI))
'        Text9.TEXT = rasocli$(NCMINI)
'    End If
End Sub

Private Sub Form_Load()
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
       EMPRE_SIN_USUARIO$ = REPLACAR$(EPAC$, USERNAME$, "")
       Caption = Caption + "  -  " + EMPRE_SIN_USUARIO$
    End If    '
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me, Picture1)
    
End Sub

Private Sub Option1_Click(Index As Integer)
    If Index = 0 Then
        Option1(0) = False
        Option1(1) = True
    End If
End Sub

Private Sub Text1_DblClick()
    Call SELECHO("GRUCOCLI")
    TEXT1.TEXT = VALACT1$("GRUCOCLI")
    If TRIM$(TEXT1.TEXT) = "" Then
        Text2.TEXT = "Todos - Sin Clasificación"
    End If
End Sub

Private Sub Text10_Change()
    NCMINI = XVALO(Text10.TEXT)
    Text9.TEXT = rasocli$(NCMINI)
End Sub

Private Sub Text2_CLICK()
    TEXT1.SetFocus
End Sub

Private Sub Text3_DBLCLICK()
    Call SELECHO("VENDEDOR")
    Text3.TEXT = VALACT1$("VENDEDOR")
    If TRIM$(Text3.TEXT) = "" Then
        Text4.TEXT = "Todos - Sin Clasificación"
    End If
End Sub

Private Sub Text1_LostFocus()
    If XVALO(TEXT1.TEXT) < 256 Then
        GRUCLI% = XVALO(TEXT1.TEXT)
        GRUX$ = ECOARCH$("GRUCOCLI", MKI$(GRUCLI%))
        If GRUX$ <> "" Then
            Text2.TEXT = GRUX$
            Exit Sub
        End If
    End If
    TEXT1.TEXT = ""
    Text2.TEXT = "Todos - Sin Clasificación"
End Sub

Private Sub Text3_LostFocus()
    If XVALO(Text3.TEXT) < 256 Then
        VENDE% = XVALO(Text3.TEXT)
        VENDX$ = ECOARCH$("VENDEDOR", Chr$(VENDE%))
        If VENDX <> "" Then
            Text4.TEXT = VENDX$
            Exit Sub
        End If
    End If
    Text3.TEXT = ""
    Text4.TEXT = "Todos - Sin Clasificación"
End Sub

Private Sub Text4_CLICK()
    Text3.SetFocus
End Sub

Private Sub Text1_Change()
    If XVALO(TEXT1.TEXT) < 256 Then
        GRUCLI% = XVALO(TEXT1.TEXT)
        GRUX$ = ECOARCH$("GRUCOCLI", MKI$(GRUCLI%))
        If GRUX$ <> "" Then
            Text2.TEXT = GRUX$
            Exit Sub
        End If
    End If
End Sub

Private Sub Text3_Change()
    If XVALO(Text3.TEXT) < 256 Then
        VENDE% = XVALO(Text3.TEXT)
        VENDX$ = ECOARCH$("VENDEDOR", Chr$(VENDE%))
        If VENDX <> "" Then
            Text4.TEXT = VENDX$
            Exit Sub
        End If
    End If
End Sub

Private Sub Text5_DblClick()
    Call SELECHO("SELFECHA")
    If FECHANUM(VALACT1$("SELFECHA")) > 0 Then
        Text5.TEXT = VALACT1$("SELFECHA")
    End If
End Sub

Private Sub Text8_Change()
    NCMAXI = XVALO(Text8.TEXT)
    Text7.TEXT = rasocli$(NCMAXI)
End Sub

Private Sub Text10_DblClick()
    NCMINI = ListaSeleccionClientes()
    If NCMINI > 0 Then
        Text10.TEXT = TRIM$(Str$(NCMINI))
        Text9.TEXT = rasocli$(NCMINI)
    End If
End Sub

Private Sub Text6_DblClick()
    Call SELECHO("SELFECHA")
    If FECHANUM(VALACT1$("SELFECHA")) > 0 Then
        Text6.TEXT = VALACT1$("SELFECHA")
    End If
End Sub

Private Sub Text7_CLICK()
    Text8.SetFocus
End Sub

Private Sub Text8_DblClick()
    NCMAXI = ListaSeleccionClientes()
    If NCMAXI > 0 Then
        Text8.TEXT = TRIM$(Str$(NCMAXI))
        Text7.TEXT = rasocli$(NCMAXI)
    End If
End Sub

Private Sub Text9_Click()
    Text10.SetFocus
End Sub

Sub GERESCUECLI()
    '
    Screen.MousePointer = 11
    GRUCLI% = XVALO(TEXT1): GRUX$ = TRIM$(Text2)
    VENDE% = XVALO(Text3): VENDX$ = TRIM$(Text4)
    Des% = FECHANUM(Text6): Has% = FECHANUM(Text5)
    NCMINI = XVALO(Text10): NCMAXI = XVALO(Text8)
    '
    MODORDE% = 1
    If Option1(2) = True Then MODORDE% = 2
    '
    UJ& = 0
    List1.Clear
    Call BLANARCH("!RESCUCLI")
    '
    ConsultaSql$ = "SELECT * FROM DEUDOR12 INNER JOIN CAJABANC WITH(NOLOCK) ON NUME_CLI=NUCLIEN WHERE STAT_CLI > 0 AND DEUDOR12.CODIEMPR >= 0 AND CAJABANC.CODIEMPR=" & CodiEmp%
    '
    If NCMINI > 0 Then CondicionSql$ = " AND DEUDOR12.NUME_CLI >= " & CStr(NCMINI)
    If NCMAXI > 0 Then CondicionSql$ = CondicionSql$ & " AND DEUDOR12.NUME_CLI <= " & CStr(NCMAXI)
    If Des% > 0 Then CondicionSql$ = CondicionSql$ & " AND CAJABANC.FECHCONTA >= '" & FETEXCOR1$(Des%) & "'"
    If Has% > 0 Then CondicionSql$ = CondicionSql$ & " AND CAJABANC.FECHCONTA <= '" & FETEXCOR1$(Has%) & "'"
    If GRUCLI% > 0 Then CondicionSql$ = CondicionSql$ & " AND DEUDOR12.GRCO_CLI = " & CStr(GRUCLI%)
    If VENDE% > 0 Then CondicionSql$ = CondicionSql$ & " AND DEUDOR12.VEND_CLI = " & CStr(VENDE%)
    CondicionSql$ = CondicionSql$ & " AND (CAJABANC.TipoMovim = 'FVEN' OR CAJABANC.TipoMovim = 'RCOB' OR CAJABANC.TipoMovim = 'AJUD')"
    If MODORDE% = 1 Then CondicionSql$ = CondicionSql$ & " ORDER BY DEUDOR12.NUME_CLI, CAJABANC.FECHEMISI, CAJABANC.FEHOREAL"
    If MODORDE% = 2 Then CondicionSql$ = CondicionSql$ & " ORDER BY DEUDOR12.RASO_CLI, CAJABANC.FECHEMISI, CAJABANC.FEHOREAL"

    ConsultaSql$ = ConsultaSql$ & CondicionSql$
    Dim RstCuentaCorriente As ADODB.Recordset
    Set RstCuentaCorriente = READSET(ConsultaSql$)
    '
    NumeroClienteAnterior& = 0
    With RstCuentaCorriente
        Do While Not .EOF
          NumeroCliente& = !nume_cli
            
          If NumeroCliente& <> NumeroClienteAnterior& Then
            If NumeroClienteAnterior& <> 0 Then Call REGAPP("!RESCUCLI", String$(256, "="))

            REBLA$ = REGBLAN$("RESCUCLI")
            CLIE$ = TRIM$(SAFETEXT$(!raso_cli)) + " (" + TRIM$(Str$(NumeroCliente&)) + ")"
            DOMIE$ = TRIM$(SAFETEXT$(!domi_cli))
            LOCAE$ = TRIM$(SAFETEXT$(!Codi_Cli))
            LOCAF$ = SAFETEXT$(!loca_Cli)
            If TRIM$(LOCAF$) <> "" Then
              LOCAE$ = TRIM$(LOCAF$) + " - " + LOCAE$
            End If
            TELEE$ = "Telefax: " + TRIM$(SAFETEXT$(!Tele_Cli))
            '
            Call REPLA(REBLA$, MKI$(NCLIEN), 1, 2)
            Call REPLA(REBLA$, "Sres.", 97, 6)
            Call REPLA(REBLA$, CLIE$, 103, 48)
            Call REPLA(REBLA$, DOMIE$, 151, 48)
            Call REPLA(REBLA$, LOCAE$, 199, 32)
            Call REPLA(REBLA$, TELEE$, 231, 24)
            '
            YATRANSPO% = 0: ITRANSPO# = 0
            NumeroClienteAnterior& = NumeroCliente&
            
            CondicionSql$ = "(CAJABANC.TipoMovim = 'FVEN' OR CAJABANC.TipoMovim = 'RCOB' OR CAJABANC.TipoMovim = 'AJUD') AND FECHCONTA<'" & FETEXCOR1$(Des%) & "' AND NUCLIEN=" & CStr(NumeroCliente&)
            SaldoDebe# = XVALO(VALOBADA("CAJABANC", "SUM(IDEBECOMP)", CondicionSql$))
            SaldoHabe# = XVALO(VALOBADA("CAJABANC", "SUM(IHABECOMP)", CondicionSql$))
            SaldoAnterior# = SaldoDebe# - SaldoHabe#
          End If
            
          If YATRANSPO% = 0 Then
              REGJ$ = REBLA$
              FEX = Des%
              Call REPLA(REGJ$, FECHATEX$(FEX), 3, 8)
              Call REPLA(REGJ$, "Saldo Anterior", 11, 32)
              Call REPLA(REGJ$, MKD$(SaldoAnterior#), 63, 8)
              Call REGAPP("!RESCUCLI", REGJ$)
              Call REGAPP("!RESCUCLI", REBLA$)
              YATRANSPO% = 1
              SaldoAcumulado# = SaldoAnterior#
          End If
          
          COCOLA$ = SAFETEXT$(!codicom_lar)
          TIX$ = Left$(!codicom_lar, 2)
          SG% = (-1)
          TIPOS$ = "Factura": If TIX$ = "AS" Then TIPOS$ = "Ajuste de Saldo"
                              If TIX$ = "AJ" Then TIPOS$ = "Ajuste de Saldo"
                              If TIX$ = "RB" Then TIPOS$ = "Recibo Cobranza"
                              If TIX$ = "VC" Then TIPOS$ = "Venta Contado"
                              If TIX$ = "ND" Then TIPOS$ = "Nota de Debito"
                              If TIX$ = "DC" Then TIPOS$ = "Descuento Caja"
                              If TIX$ = "NC" Then TIPOS$ = "Nota de Credito"
                              If TIX$ = "CC" Then TIPOS$ = "Comp.Ctas.": SG% = 1 '\\\OT-06-0112-WAR-30/03/06///
          ImporteComprobante# = XVALO(!Idebecomp) - XVALO(!IHABECOMP)
          '
890       ImporteCbte# = ROUND(ImporteComprobante#)
          SaldoAcumulado# = SaldoAcumulado# + ImporteCbte#
          FechaEmision% = CVINT%(!FECHEMISI)
          NumeroComprobante& = !NumeComp
          Moneda% = ROUND(!Mone_Emis, 0)
          TIPOCAMBIO# = ROUND(!Tipo_Cam, 5)
          If TIPOCAMBIO# < 0.05 Then TIPOCAMBIO# = 1
          If TIPOCAMBIO# >= 0.01 Then
            ImporteDolar# = ROUND#(ImporteCbte# / TIPOCAMBIO#, 2)
          End If
          '
          Call REPLA(REGJ$, FECHATEX$(FechaEmision%), 3, 8)
          Call REPLA(REGJ$, TIPOS$, 11, 32)
          Call REPLA(REGJ$, MKS$(NumeroComprobante&), 43, 4)
          If ImporteCbte# > 0 Then
            Call REPLA(REGJ$, MKD$(ImporteCbte#), 47, 8)
            Call REPLA(REGJ$, MKD$(0), 55, 8)
          ElseIf ImporteCbte# < 0 Then
            Call REPLA(REGJ$, MKD$(0), 47, 8)
            Call REPLA(REGJ$, MKD$(Abs(ImporteCbte#)), 55, 8)
          Else
            Call REPLA(REGJ$, MKD$(0), 47, 8)
            Call REPLA(REGJ$, MKD$(0), 55, 8)
          End If
          Call REPLA(REGJ$, MKD$(SaldoAcumulado#), 63, 8)
          Call REPLA(REGJ$, SIMBOLPES$(Moneda%), 71, 4)
          Call REPLA(REGJ$, MKD$(ImporteDolar#), 75, 8)
          Call REGAPP("!RESCUCLI", REGJ$)
          '
         .MoveNext
        Loop
990 End With
    RstCuentaCorriente.Close
    Set RstCuentaCorriente = Nothing
    '
    If ULTREG&("!RESCUCLI") > 0 Then
        If NumeroClienteAnterior& <> 0 Then Call REGAPP("!RESCUCLI", String$(256, "="))
        Call TRACENTRAL("RESCUCLI", "")
        Screen.MousePointer = 1
        Call CONECRUN("*RESCCLI", "Listado Resumen de Cuenta Deudores")
    Else
        Screen.MousePointer = 1
        Call COMUNI("No se Ha Encontrado Información para las Opciones Elegidas")
    End If

    
'    FIN& = CantRegistros("deudor12", "stat_cli>0")
'    Dim RST As ADODB.Recordset
'    SQLX$ = "select nume_cli from deudor12 with(nolock) where stat_cli>0"
'    Set RST = READSET(SQLX$)
'    '
'    U& = 0
'    With RST
'        Do While Not .EOF
'           '
'           Call TRACE(20, U&, FIN&)
'           NCLIEN = XVALO(!nume_cli)
'           '
'           If NCLIEN >= NCMINI Then
'             If NCLIEN <= NCMAXI Then
'               If GRUCLI% = 0 Or GRUCOCLI%(NCLIEN) = GRUCLI% Then
'                 If VENDE% = 0 Or VENDCLI%(NCLIEN) = VENDE% Then
'                   '
'                   TLISTA$ = Space$(36)
'                   If MODORDE% = 1 Then
'                        Call REPLA(TLISTA$, CSTRING$(MKS$(NCLIEN), 3, 6, 0, 2), 1, 6)
'                     ElseIf MODORDE% = 2 Then
'                        Call REPLA(TLISTA$, rasocli$(NCLIEN), 1, 30)
'                   End If
'                   Call REPLA(TLISTA$, CSTRING$(MKS$(NCLIEN), 3, 6, 0, 2), 31, 6)
'                   List1.AddItem TLISTA$
'                   '
'                 End If
'               End If
'             End If
'           End If
'           U& = U& + 1
'           '
'          .MoveNext
'        Loop
'     End With
'     On Error Resume Next
'     RST.Close
'     Set RST = Nothing
'     On Error GoTo 0
'
'    '
'    FIN& = List1.ListCount
'    For U& = 1 To FIN&
'      '
'      Call TRACE(20, U&, FIN&)
'      List1.ListIndex = U& - 1
'      NCLIEN = XVALO(Mid$(List1.TEXT, 31, 6))
'      On Error Resume Next
'      RECUEC07.Text1 = TRIM$(Str$(NCLIEN))
'      Call FRESHALL
'      '
'      If CONTROLGRABA% = 1 Then
'        Call RECUEC07.Command4_Click
'      End If
'
'      REBLA$ = REGBLAN$("RESCUCLI")
'      CLIE$ = TRIM$(RECUEC07.Text2) + " (" + TRIM$(Str$(NCLIEN)) + ")"
'      DOMIE$ = TRIM$(RECUEC07.Text3)
'      LOCAE$ = TRIM$(RECUEC07.Text5)
'      LOCAF$ = RECUEC07.Text4
'      If TRIM$(LOCAF$) <> "" Then
'        LOCAE$ = TRIM$(LOCAF$) + " - " + LOCAE$
'      End If
'      TELEE$ = "Telefax: " + TRIM$(RECUEC07.Text7)
'      '
'      Call REPLA(REBLA$, MKI$(NCLIEN), 1, 2)
'      Call REPLA(REBLA$, "Sres.", 97, 6)
'      Call REPLA(REBLA$, CLIE$, 103, 48)
'      Call REPLA(REBLA$, DOMIE$, 151, 48)
'      Call REPLA(REBLA$, LOCAE$, 199, 32)
'      Call REPLA(REBLA$, TELEE$, 231, 24)
'      '
'      YATRANSPO% = 0: ITRANSPO# = 0
'      FIN1& = RECUEC07.List1.ListCount
'      '
'      For V& = 1 To FIN1&
'        '
'        REGI$ = RECUEC07.List1.List(V& - 1)
'        '
'        FECHA$ = Left$(REGI$, 8)
'        FF% = FECHANUM(FECHA$)
'        If FF% < DES% Then
'            ITRANSPO# = XVALO(Mid$(REGI$, 63, 14))
'        End If
'        If FF% >= DES% Then
'          If FF% <= HAS% Then
'            If YATRANSPO% = 0 Then
'                REGJ$ = REBLA$
'                FEX = DES%
'                Call REPLA(REGJ$, FECHATEX$(FEX), 3, 8)
'                Call REPLA(REGJ$, "Saldo Anterior", 11, 32)
'                Call REPLA(REGJ$, MKD$(ITRANSPO#), 63, 8)
'                Call REGAPP("!RESCUCLI", REGJ$)
'                Call REGAPP("!RESCUCLI", REBLA$)
'                YATRANSPO% = 1
'            End If
'            '
'            REGJ$ = REBLA$
'            DOCU$ = Mid$(REGI$, 11, 16)
'            NUDO& = XVALO(Mid$(REGI$, 27, 8))
'            DEBE# = XVALO(Mid$(REGI$, 35, 13))
'            HABER# = XVALO(Mid$(REGI$, 50, 13))
'            SALDO# = XVALO(Mid$(REGI$, 63, 14))
'            SIMPES$ = TRIM$(Mid$(REGI$, 79, 2))
'            IDOLA# = 0
'            TCXX = XVALO(Mid$(REGI$, 81, 9))
'            If TCXX >= 0.01 Then
'              IDOLA# = ROUND#((Abs(DEBE# - HABER#)) / TCXX, 2)
'            End If
'            '
'            Call REPLA(REGJ$, FECHA$, 3, 8)
'            Call REPLA(REGJ$, DOCU$, 11, 32)
'            Call REPLA(REGJ$, MKS$(NUDO&), 43, 4)
'            Call REPLA(REGJ$, MKD$(DEBE#), 47, 8)
'            Call REPLA(REGJ$, MKD$(HABER#), 55, 8)
'            Call REPLA(REGJ$, MKD$(SALDO#), 63, 8)
'            Call REPLA(REGJ$, SIMPES$, 71, 4)
'            Call REPLA(REGJ$, MKD$(IDOLA#), 75, 8)
'            Call REGAPP("!RESCUCLI", REGJ$)
'            '
'          End If
'        End If
'      Next V&
'      '
'      If YATRANSPO% = 1 Then
'          Call REGAPP("!RESCUCLI", String$(256, "="))
'      End If
'      '
'    Next U&
'    On Error GoTo 0
'    '
'    ' SE AGREGA PORQUE SINO EL CONECRUN LISTA EL CENTRAL
'    ' Y LA INFORMACION ANTERIOR QUEDO EN EL LOCAL
'    Call TRACENTRAL("RESCUCLI", "")
'    '
'    Call CONECRUN("*RESCCLI", "Listado Resumen de Cuenta Deudores")
    Screen.MousePointer = 1
    '
End Sub
'


