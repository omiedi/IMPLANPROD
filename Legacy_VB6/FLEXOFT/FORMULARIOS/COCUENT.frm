VERSION 5.00
Begin VB.Form COCUENT 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0E0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control de Cumplimiento de Entregas"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   4980
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   4980
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0E0C0&
      Caption         =   "Indices por Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2865
      Left            =   210
      TabIndex        =   3
      Top             =   210
      Width           =   3585
      Begin VB.Frame Frame6 
         BackColor       =   &H00C0E0C0&
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
         Height          =   705
         Left            =   360
         TabIndex        =   10
         Top             =   1920
         Width           =   2850
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0E0C0&
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
            Left            =   1440
            TabIndex        =   11
            Top             =   300
            Value           =   -1  'True
            Width           =   1170
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0E0C0&
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
            Left            =   240
            TabIndex        =   12
            Top             =   300
            Width           =   1380
         End
      End
      Begin VB.PictureBox Picture8 
         BackColor       =   &H00E0E0E0&
         Height          =   645
         Left            =   735
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   4
         Top             =   1050
         Width           =   2430
         Begin VB.CommandButton Command14 
            Caption         =   "IC-P"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   0
            Picture         =   "COCUENT.frx":0000
            TabIndex        =   5
            ToolTipText     =   "Indice de Cumplimiento de Pedidos Entregados por Numero de Pedido"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Indice de Cumpli- miento por Pedido"
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
            TabIndex        =   6
            Top             =   105
            Width           =   1695
         End
      End
      Begin VB.PictureBox Picture14 
         BackColor       =   &H00E0E0E0&
         Height          =   645
         Left            =   420
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   7
         Top             =   525
         Width           =   2430
         Begin VB.CommandButton Command22 
            Caption         =   "IC-R"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Indice de Cumplimiento de Pedidos Entregados por Remito"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Indice de Cumpli- miento por Remito"
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
            Left            =   630
            TabIndex        =   9
            Top             =   60
            Width           =   1695
         End
      End
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   3990
      ScaleHeight     =   75
      ScaleWidth      =   795
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2220
      Width           =   855
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   12000
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
      Height          =   800
      Left            =   3960
      Picture         =   "COCUENT.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Terminar - Salir de la Aplicación"
      Top             =   315
      Width           =   855
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   3360
      Picture         =   "COCUENT.frx":0294
      Top             =   2550
      Width           =   2010
   End
End
Attribute VB_Name = "COCUENT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    Call FINAL("")
End Sub

Private Sub Command14_Click()
   Command14.Enabled = False
   If Option1.Value = True Then
       Call FINAL("*CCUMPED")
     Else
       Call ICUMCLI(2)
   End If
99 Command14.Enabled = True
End Sub

Private Sub Command22_Click()
   Command22.Enabled = False
   If Option1.Value = True Then
       Call FINAL("*CCUMREM")
     Else
       Call ICUMCLI(1)
   End If
99 Command22.Enabled = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   While WindowState <> 0
     WindowState = 0
     OPENFORMS = DoEvents
   Wend
End Sub

Private Sub Form_Load()
   '
   Call VERJOBID(OKEY%)
   If OKEY% < 1 Then Call FINAL("")
   '
End Sub

Private Sub Form_Resize()
   If WindowState = 1 Then
     Call CIERRARCH("*.*")
     Call BAJALDISCO
   End If
End Sub

Sub ICUMCLI(OPX%)
   '
   HAS% = HOY(HOS$)
   FEX = HAS%
   HASS$ = FECHATEX$(FEX)
   DESS$ = "01" + Mid$(HASS$, 3)
   DES% = FECHANUM(DESS$)
   '
   NCMIN% = 32767: NCMAX% = 0
   IXXX$ = LOADFILE$(LUDAT$ + "DEUDOR1.X01")
   For KKL& = 1 To Len(IXXX$) Step 2
     NCMX% = CVI(Mid$(IXXX$, KKL&, 2))
     If NCMX% > 0 Then
       If NCMX% < NCMIN% Then NCMIN% = NCMX%
       If NCMX% > NCMAX% Then NCMAX% = NCMX%
     End If
   Next KKL&
   '
   VDEF$ = MKI$(DES%) + MKI$(HAS%) + MKI$(NCMIN%) + MKI$(NCMAX%)
   '
10 OBX$ = OBJPLA$("SELRAFECHCLI", VDEF$)
   If OBX$ = "" Then
      Exit Sub
   End If
   DES% = CVI(Left$(OBX$, 2))
   HAS% = CVI(Mid$(OBX$, 3, 2))
   If HAS% < DES% Then GoTo 10
   '
   NC1% = CVI(Mid$(OBX$, 5, 2))
   NC2% = CVI(Mid$(OBX$, 7, 2))
   If NC2% < NC1% Then GoTo 10
   '
   FEX = DES%: PERIO$ = "Periodo: " + FECHATEX$(FEX)
   FEX = HAS%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
   '
   J& = 0
   Call COPYSTRU("COCUENT", "COCUSEL")
   FIN& = ULTREG&("COCUENT")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     X$ = REGLEIDO$("COCUENT", U&)
     FF% = CVI(Mid$(X$, 7, 2))        ' fecha de emision DE REMITO
     If OPX% = 2 Then FF% = CVI(Mid$(X$, 29, 2))
     If FF% >= DES% Then
       If FF% <= HAS% Then
         NP% = CVI(Mid$(X$, 1, 2))        ' CLIENTE
         If NP% >= NCMIN% Then
           If NP% <= NCMAX% Then
             If OPX% = 2 Then
               NREMI& = CVS(Mid$(X$, 3, 4))
               FREMI% = CVI(Mid$(X$, 7, 2))
               NPEDI& = CVS(Mid$(X$, 25, 4))
               FPEDI% = CVI(Mid$(X$, 29, 2))
               Call REPLA(X$, MKS$(NPEDI&) + MKI$(FPEDI%), 3, 6)
               Call REPLA(X$, MKS$(NREMI&) + MKI$(FREMI%), 25, 6)
             End If
             J& = J& + 1
             Call GRAREG("COCUSEL", X$, J&)
           End If
         End If
       End If
     End If
   Next U&
   Call SETULTREG("COCUSEL", J&)
   '
   Call HEADERS("COCUSEL", "")
   Call HEADERS("COCUSEL", PERIO$)
   '
   Call FILESORT("COCUSEL", "", 1, 3)   ' ORDENA POR NUMERO Y DOCUMENTO
   '
   If OPX% = 1 Then
       Call FINAL("*LCUMREM")
     Else
       Call FINAL("*LCUMPED")
   End If
   '
   Screen.MousePointer = 1
   '
End Sub

