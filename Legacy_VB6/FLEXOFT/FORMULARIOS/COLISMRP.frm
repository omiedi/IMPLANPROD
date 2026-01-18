VERSION 5.00
Begin VB.Form COLISMRP 
   BackColor       =   &H00D0F0F0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MRP-II - Complementos, Consultas, Listados"
   ClientHeight    =   5475
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   8955
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   8955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame7 
      BackColor       =   &H00D0F0F0&
      BorderStyle     =   0  'None
      Caption         =   "Frame6"
      Height          =   1455
      Left            =   6480
      TabIndex        =   40
      Top             =   720
      Width           =   735
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00D0F0F0&
      BorderStyle     =   0  'None
      Caption         =   "Frame6"
      Height          =   1455
      Left            =   8760
      TabIndex        =   39
      Top             =   720
      Width           =   735
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00D0F0F0&
      Caption         =   "Necesidades Tipificadas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3120
      Left            =   4560
      TabIndex        =   24
      Top             =   2160
      Width           =   4215
      Begin VB.PictureBox Picture6 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1080
         ScaleHeight     =   585
         ScaleWidth      =   2580
         TabIndex        =   25
         Top             =   1440
         Width           =   2640
         Begin VB.CommandButton Command10 
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
            Left            =   0
            Picture         =   "COLISMRP.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   26
            ToolTipText     =   "Listado de Necesidad Total por Componentes (Neto de Stocks)"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Neto de Stocks por Componente /Fecha"
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
            Left            =   720
            TabIndex        =   27
            Top             =   105
            Width           =   1905
         End
      End
      Begin VB.PictureBox Picture7 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   720
         ScaleHeight     =   585
         ScaleWidth      =   2580
         TabIndex        =   28
         Top             =   945
         Width           =   2640
         Begin VB.CommandButton Command12 
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
            Left            =   0
            Picture         =   "COLISMRP.frx":014A
            Style           =   1  'Graphical
            TabIndex        =   29
            ToolTipText     =   "Listado de Necesidad Total por Componentes Segun Formulas Tipificadas"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Totales Bruto por Componente / Fecha"
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
            Left            =   720
            TabIndex        =   30
            Top             =   60
            Width           =   1905
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00D0F0F0&
         Caption         =   "Formato"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   34
         Top             =   2280
         Width           =   3735
         Begin VB.OptionButton Option4 
            BackColor       =   &H00D0F0F0&
            Caption         =   "Tabla"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2040
            TabIndex        =   36
            Top             =   240
            Width           =   1215
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00D0F0F0&
            Caption         =   "Lista"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   840
            TabIndex        =   35
            Top             =   240
            Value           =   -1  'True
            Width           =   1215
         End
      End
      Begin VB.PictureBox Picture8 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   360
         ScaleHeight     =   585
         ScaleWidth      =   2580
         TabIndex        =   31
         Top             =   420
         Width           =   2640
         Begin VB.CommandButton Command13 
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
            Left            =   0
            Picture         =   "COLISMRP.frx":0454
            Style           =   1  'Graphical
            TabIndex        =   32
            ToolTipText     =   "Listado de Necesidad Componentes Segun Formulas Tipificadas por Destino"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Analitico Bruto por Componente / Fecha"
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
            TabIndex        =   33
            Top             =   45
            Width           =   1905
         End
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   4680
      TabIndex        =   38
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Command7"
      Height          =   495
      Left            =   6000
      TabIndex        =   37
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00D0F0F0&
      Caption         =   "Exportación de Productos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1800
      Left            =   240
      TabIndex        =   17
      Top             =   3480
      Width           =   4215
      Begin VB.PictureBox Picture5 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   945
         ScaleHeight     =   585
         ScaleWidth      =   2580
         TabIndex        =   18
         Top             =   945
         Width           =   2640
         Begin VB.CommandButton Command6 
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
            Left            =   0
            Picture         =   "COLISMRP.frx":075E
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "Listado de Productos Exportados por Rango de Fechas"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Detalle de Productos por Rango de Fechas"
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
            TabIndex        =   20
            Top             =   100
            Width           =   1905
         End
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   630
         ScaleHeight     =   585
         ScaleWidth      =   2580
         TabIndex        =   21
         Top             =   420
         Width           =   2640
         Begin VB.CommandButton Command5 
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
            Left            =   0
            Picture         =   "COLISMRP.frx":0A68
            Style           =   1  'Graphical
            TabIndex        =   22
            ToolTipText     =   "Listado de Remitos de Exportación por Rango de Fechas"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Detalle Remitos por Rango de Fechas"
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
            TabIndex        =   23
            Top             =   45
            Width           =   1905
         End
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00D0F0F0&
      Caption         =   "Programa de Entregas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3120
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   4215
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1080
         ScaleHeight     =   585
         ScaleWidth      =   2580
         TabIndex        =   5
         Top             =   1440
         Width           =   2640
         Begin VB.CommandButton Command8 
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
            Left            =   0
            Picture         =   "COLISMRP.frx":0D72
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Programas Netos de Entregas y Stock"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Listado Neto de Entregas y Stock"
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
            Left            =   840
            TabIndex        =   7
            Top             =   105
            Width           =   1905
         End
      End
      Begin VB.PictureBox Picture11 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   720
         ScaleHeight     =   585
         ScaleWidth      =   2580
         TabIndex        =   11
         Top             =   945
         Width           =   2640
         Begin VB.CommandButton Command9 
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
            Left            =   0
            Picture         =   "COLISMRP.frx":0EBC
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Listados de Programa NETO de Entregas Efectuadas"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Listado Neto de Entregas Efectuadas"
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
            Left            =   720
            TabIndex        =   13
            Top             =   60
            Width           =   1905
         End
      End
      Begin VB.PictureBox Picture12 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   360
         ScaleHeight     =   585
         ScaleWidth      =   2580
         TabIndex        =   14
         Top             =   420
         Width           =   2640
         Begin VB.CommandButton Command15 
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
            Left            =   0
            Picture         =   "COLISMRP.frx":11C6
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Listado de Programa de Entregas BRUTO"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "Listado Anual  Bruto por Programa-Cliente"
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
            TabIndex        =   16
            Top             =   45
            Width           =   1905
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00D0F0F0&
         Caption         =   "Formato"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   8
         Top             =   2280
         Width           =   3735
         Begin VB.OptionButton Option1 
            BackColor       =   &H00D0F0F0&
            Caption         =   "Lista"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   840
            TabIndex        =   10
            Top             =   240
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00D0F0F0&
            Caption         =   "Tabla"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2040
            TabIndex        =   9
            Top             =   240
            Width           =   1215
         End
      End
   End
   Begin VB.CommandButton Command11 
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
      Height          =   720
      Left            =   7200
      Picture         =   "COLISMRP.frx":14D0
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Ayuda FLEXOFT en Línea"
      Top             =   315
      Width           =   720
   End
   Begin VB.CommandButton command1 
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
      Height          =   720
      Left            =   8040
      Picture         =   "COLISMRP.frx":17DA
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Cierra - Abandona la Aplicación"
      Top             =   315
      Width           =   720
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   7440
      ScaleHeight     =   75
      ScaleWidth      =   1245
      TabIndex        =   1
      Top             =   1200
      Width           =   1305
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   2
         Top             =   105
         Width           =   12000
      End
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   6960
      Picture         =   "COLISMRP.frx":1924
      Top             =   1560
      Width           =   2010
   End
End
Attribute VB_Name = "COLISMRP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub command1_Click()
   Call CIERRARCH("*.*")
   Hide
End Sub

Private Sub Command10_Click()
   Command10.Enabled = False
   If Option3.Value = True Then
       Call FINAL("*NECALPR4")
     Else
       Call FINAL("*NETINTA")
   End If
   Command10.Enabled = True
End Sub

Private Sub Command11_Click()
   Command11.Enabled = False
   Call HELPONLINE("MRPSAB04")
   Command11.Enabled = True
End Sub

Private Sub Command12_Click()
   Command12.Enabled = False
   If Option3.Value = True Then
       Call FINAL("*NECALPR3")
     Else
       Call FINAL("*NETIBRA")
   End If
   Command12.Enabled = True
End Sub

Private Sub Command13_Click()
   Command13.Enabled = False
   If Option4.Value = True Then
       Call MENSERR(24, "Formato Tabla no Disponible\para este Listado")
       Exit Sub
     Else
       Call FINAL("*NECALPR2")
   End If
   Command13.Enabled = True
End Sub

Private Sub Command2_Click()
   Command2.Enabled = False
   Call FILESORT("NECALPRO", "", 5, 7)
   Call FINAL("*NECALPR1")
   Command2.Enabled = True
End Sub

Private Sub Command4_Click()
   Command4.Enabled = False
   MsgBox "Ultimo Cálculo: " + TRIM$(CONTROL$("CALCONIT", "FULTICAL"))
10 Call DESHASFECHA(DES%, HAS%, PERIO$)
   If DES% > 0 Then
     If HAS% >= DES% Then
       Screen.MousePointer = 11
       JJ& = 0: FIN& = ULTREG&("EXPOTIP")
       For U& = 1 To FIN&
         Call TRACE(20, U&, FIN&)
         XX$ = REGLEIDO$("EXPOTIP", U&)
         FMXX% = CVI(Left$(XX$, 2))
         If FMXX% >= DES% Then
           If FMXX% <= HAS% Then
             JJ& = JJ& + 1
             Call GRAREG("EXPOTIPF", XX$, JJ&)
           End If
         End If
       Next U&
       Call SETULTREG("EXPOTIPF", JJ&)
       Call CIERRARCH("EXPOTIPF")
       '
       Call HEADERS("EXPOTIPF", "")
       Call HEADERS("EXPOTIPF", "Período: " + PERIO$)
       Call CONECRUN("*EXPOTIPF", "Listado Detalle de Consumos Re-Exportados")
       Screen.MousePointer = 1
       '
     End If
   End If
   Command4.Enabled = True
End Sub

Private Sub Command5_Click()
   '
   Command5.Enabled = False
   MsgBox "Ultimo Cálculo: " + TRIM$(CONTROL$("CALCONIT", "FULTICAL"))
10 Call DESHASFECHA(DES%, HAS%, PERIO$)
   If DES% > 0 Then
     If HAS% >= DES% Then
       Screen.MousePointer = 11
       Call COPYSTRU("MOVISTO", "MOVSTEX1")
       JJ& = 0: FIN& = ULTREG&("MOVSTEX0")
       For U& = 1 To FIN&
         Call TRACE(20, U&, FIN&)
         XX$ = REGLEIDO$("MOVSTEX0", U&)
         FMXX% = CVI(Left$(XX$, 2))
         If FMXX% >= DES% Then
           If FMXX% <= HAS% Then
             JJ& = JJ& + 1
             Call GRAREG("MOVSTEX1", XX$, JJ&)
           End If
         End If
       Next U&
       Call SETULTREG("MOVSTEX1", JJ&)
       Call CIERRARCH("MOVSTEX1")
       '
       Call HEADERS("MOVSTEX1", "")
       Call HEADERS("MOVSTEX1", "Período: " + PERIO$)
       Call CONECRUN("*LMOVEX1", "Listado Detalle de Exportación")
       Screen.MousePointer = 1
       '
     End If
   End If
   Command5.Enabled = True
   '
End Sub

Private Sub Command6_Click()
   '
   Command6.Enabled = False
   MsgBox "Ultimo Cálculo: " + TRIM$(CONTROL$("CALCONIT", "FULTICAL"))
10 Call DESHASFECHA(DES%, HAS%, PERIO$)
   If DES% > 0 Then
     If HAS% >= DES% Then
       Screen.MousePointer = 11
       JJ& = 0: FIN& = ULTREG&("MOVSTEX0")
       For U& = 1 To FIN&
         Call TRACE(20, U&, FIN&)
         XX$ = REGLEIDO$("MOVSTEX0", U&)
         FMXX% = CVI(Left$(XX$, 2))
         If FMXX% >= DES% Then
           If FMXX% <= HAS% Then
             JJ& = JJ& + 1
             Call GRAREG("MOVSTEX1", XX$, JJ&)
           End If
         End If
       Next U&
       Call SETULTREG("MOVSTEX1", JJ&)
       Call CIERRARCH("MOVSTEX1")
       '
       Call FILESORT("MOVSTEX1", "", 2, 2)
       '
       Call HEADERS("MOVSTEX1", "")
       Call HEADERS("NOVSTEX1", "Período: " + PERIO$)
       Call CONECRUN("*LMOVEX2", "Detalle de Exportación por Producto")
       Screen.MousePointer = 1
       '
     End If
   End If
   Command6.Enabled = True
   '
End Sub

Private Sub Command8_Click()
    Command8.Enabled = False
10  Call SELECHO("PROFABRI")
    NPRG% = Val(VALACT1$("PROFABRI"))
    '
    If NPRG% > 0 Then
      Screen.MousePointer = 11
      OPP% = 1: If Option2.Value = True Then OPP% = 2
      If OPP% = 1 Then
          Call FILTERFILE("PRGNETST", "PRGNETFI", 1, MKI$(NPRG%))
          Call HEADERS("PRGNETFI", "")
          Call HEADERS("PRGNETFI", "Programa: " + VALACT2$("PROFABRI"))
          Call CONECRUN("*PRGNETFI", "Listado Programa Anual NETO de Entregas")
        ElseIf OPP% = 2 Then
          DEPROG$ = VALACT2$("PROFABRI")
          Call FILTERFILE("PROGSTKA", "PROGSTKX", 3, MKI$(NPRG%))
          Call FILTERFILE("PROGSTKB", "PROGSTKY", 3, MKI$(NPRG%))
          Call HEADERS("PROGSTKX", "")
          Call HEADERS("PROGSTKX", "Corresponde a: " + DEPROG$ + " - Año 20" + TRIM$(Mid$(HOS$, 7, 2)))
          ASIGUIE$ = TRIM$(Str$(1 + Val(Mid$(HOS$, 7, 2)))): While Len(ASIGUIE$) < 2: ASIGUIE$ = "0" + ASIGUIE$: Wend
          Call HEADERS("PROGSTKY", "")
          Call HEADERS("PROGSTKY", "Corresponde a: " + DEPROG$ + " - Año 20" + TRIM$(ASIGUIE$))
          '
          Call CONECRUN("*PRGSTKX", "Programa de Procuracion")
      End If
      Screen.MousePointer = 1
      GoTo 10
    End If
    Command8.Enabled = True
    '
End Sub

Private Sub Command9_Click()
    Command9.Enabled = False
10  Call SELECHO("PROFABRI")
    NPRG% = Val(VALACT1$("PROFABRI"))
    '
    If NPRG% > 0 Then
      Screen.MousePointer = 11
      OPP% = 1: If Option2.Value = True Then OPP% = 2
      If OPP% = 1 Then
          Call FILTERFILE("PROGENET", "PRGNETFI", 1, MKI$(NPRG%))
          Call HEADERS("PRGNETFI", "")
          Call HEADERS("PRGNETFI", "Programa: " + VALACT2$("PROFABRI"))
          Call CONECRUN("*PRGNETFI", "Listado Programa Anual NETO de Entregas")
        ElseIf OPP% = 2 Then
          DEPROG$ = VALACT2$("PROFABRI")
          Call FILTERFILE("PROGNETA", "PROGNETX", 3, MKI$(NPRG%))
          Call FILTERFILE("PROGNETB", "PROGNETY", 3, MKI$(NPRG%))
          Call HEADERS("PROGNETX", "")
          Call HEADERS("PROGNETX", "Corresponde a: " + DEPROG$ + " - Año 20" + TRIM$(Mid$(HOS$, 7, 2)))
          ASIGUIE$ = TRIM$(Str$(1 + Val(Mid$(HOS$, 7, 2)))): While Len(ASIGUIE$) < 2: ASIGUIE$ = "0" + ASIGUIE$: Wend
          Call HEADERS("PROGNETY", "")
          Call HEADERS("PROGNETY", "Corresponde a: " + DEPROG$ + " - Año 20" + TRIM$(ASIGUIE$))
          '
          Call CONECRUN("*PRGNETX", "Programa de Fabricacion Neto de Entregas")
      End If
      Screen.MousePointer = 1
      GoTo 10
    End If
    Command9.Enabled = True
    Exit Sub
    '
End Sub
'
Sub GENPREXCEL(DEPROG$)
   '
   Dim CODIPRO%(16384), PROGENT$(16384)
   Dim FEPLA%(24)
   '
   HOII% = HOY(HOS$)
   '
   For KU% = 0 To 24
     TTYY$ = TRIM$(Mid$(APROGMRP.Label5, 9 * KU% + 1, 9))
     If KU% > 0 Then
         TTYY$ = MESANO$(TTYY$, DES%, HAS%)
         FEPLA%(KU%) = HAS%
       Else
         FEPLA%(KU%) = DIANTE(FECHANUM("01/01/" + Mid$(HOS$, 7, 2)))
     End If
   Next KU%
   '
   For KKI% = 0 To NUMAS%
     PROGENT$(KKI%) = String$(100, 0)
   Next KKI%
   '
   FIN& = ULTREG("PRGNETFI")
   For I& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     P00$ = REGLEIDO$("PRGNETFI", I&)
     CI% = CVI(Mid$(P00$, 3, 2))
     CANPRO = CVS(Mid$(P00$, 57, 4))
     '
     If CANPRO > 0.05 Then
       If CI% > 0 Then
         If CI% <= NUMAS% Then
           FEXI% = CVI(Mid$(P00$, 5, 2))
           If FEXI% <= FEPLA%(0) Then
              COLUX% = 0
              GoTo 20
            Else
              For KKJ% = 24 To 1 Step -1
                If FEXI% >= FEPLA%(KKJ%) Then
                  COLUX% = KKJ%
                  GoTo 20
                End If
              Next KKJ%
           End If
           GoTo 39
           '
20         For KKI% = 1 To CACODIP%
             If CODIPRO%(KKI%) = CI% Then GoTo 30
           Next KKI%
           CACODIP% = CACODIP% + 1
           CODIPRO%(CACODIP%) = CI%
           KKI% = CACODIP%
           '
30         If COLUX% >= 0 Then
             If COLUX% <= 24 Then
               CANPRO = CANPRO + CVS(Mid$(PROGENT$(KKI%), 1 + 4 * COLUX%, 4))
               Mid$(PROGENT$(KKI%), 1 + 4 * COLUX%, 4) = MKS$(CANPRO)
             End If
           End If
         End If
       End If
     End If
     '
39 Next I&
   '
   KK& = 0
   REBLA$ = REGBLAN$("PROGNETX")
   REBLB$ = REGBLAN$("PROGNETY")
   '
   For KKI% = 1 To CACODIP%
     '
     CIKKI% = CODIPRO%(KKI%)
     TTXA$ = MKI$(CIKKI%)
     '
     SUMA = 0
     For JJ& = 0 To 12
       XX$ = Mid$(PROGENT$(KKI%), 4 * JJ& + 1, 4)
       SUMA = SUMA + CVS(XX$)
       TTXA$ = TTXA$ + XX$
     Next JJ&
     TTXA$ = TTXA$ + MKS$(SUMA) + String$(128, 0)
     '
     TTXB$ = MKI$(CIKKI%) + MKS$(SUMA)
     For JJ& = 13 To 24
       XX$ = Mid$(PROGENT$(KKI%), 4 * JJ& + 1, 4)
       SUMA = SUMA + CVS(XX$)
       TTXB$ = TTXB$ + XX$
     Next JJ&
     TTXB$ = TTXB$ + MKS$(SUMA) + String$(128, 0)
     '
     KK& = KK& + 1
     Call GRAREG("PROGNETX", TTXA$, KK&)
     Call GRAREG("PROGNETY", TTXB$, KK&)
   Next KKI%
   '
   Call SETULTREG("PROGNETX", KK&)
   Call SETULTREG("PROGNETY", KK&)
   Call HEADERS("PROGNETX", "")
   Call HEADERS("PROGNETX", "Corresponde a: " + DEPROG$ + " - Año 20" + TRIM$(Mid$(HOS$, 7, 2)))
   ASIGUIE$ = TRIM$(Str$(1 + Val(Mid$(HOS$, 7, 2)))): While Len(ASIGUIE$) < 2: ASIGUIE$ = "0" + ASIGUIE$: Wend
   Call HEADERS("PROGNETY", "")
   Call HEADERS("PROGNETY", "Corresponde a: " + DEPROG$ + " - Año 20" + TRIM$(ASIGUIE$))
   '
   Call CONECRUN("*PRGNETX", "Programa de Fabricacion Neto de Entregas")

End Sub
'
Sub GENPREXCEL2(DEPROG$)
   '
   Dim CODIPRO%(16384), PROGENT$(16384)
   Dim FEPLA%(24)
   '
   HOII% = HOY(HOS$)
   '
   For KU% = 0 To 24
     TTYY$ = TRIM$(Mid$(APROGMRP.Label5, 9 * KU% + 1, 9))
     If KU% > 0 Then
         TTYY$ = MESANO$(TTYY$, DES%, HAS%)
         FEPLA%(KU%) = HAS%
       Else
         FEPLA%(KU%) = DIANTE(FECHANUM("01/01/" + Mid$(HOS$, 7, 2)))
     End If
   Next KU%
   '
   For KKI% = 0 To NUMAS%
     PROGENT$(KKI%) = String$(100, 0)
   Next KKI%
   '
   FIN& = ULTREG("PRGNETST")
   For I& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     P00$ = REGLEIDO$("PRGNETST", I&)
     CI% = CVI(Mid$(P00$, 3, 2))
     CANPRO = CVS(Mid$(P00$, 57, 4))
     '
     If CANPRO > 0.05 Then
       If CI% > 0 Then
         If CI% <= NUMAS% Then
           FEXI% = CVI(Mid$(P00$, 5, 2))
           If FEXI% <= FEPLA%(0) Then
              COLUX% = 0
              GoTo 20
            Else
              For KKJ% = 24 To 1 Step -1
                If FEXI% >= FEPLA%(KKJ%) Then
                  COLUX% = KKJ%
                  GoTo 20
                End If
              Next KKJ%
           End If
           GoTo 39
           '
20         For KKI% = 1 To CACODIP%
             If CODIPRO%(KKI%) = CI% Then GoTo 30
           Next KKI%
           CACODIP% = CACODIP% + 1
           CODIPRO%(CACODIP%) = CI%
           KKI% = CACODIP%
           '
30         If COLUX% >= 0 Then
             If COLUX% <= 24 Then
               CANPRO = CANPRO + CVS(Mid$(PROGENT$(KKI%), 1 + 4 * COLUX%, 4))
               Mid$(PROGENT$(KKI%), 1 + 4 * COLUX%, 4) = MKS$(CANPRO)
             End If
           End If
         End If
       End If
     End If
     '
39 Next I&
   '
   KK& = 0
   REBLA$ = REGBLAN$("PROGSTKX")
   REBLB$ = REGBLAN$("PROGSTKY")
   '
   For KKI% = 1 To CACODIP%
     '
     CIKKI% = CODIPRO%(KKI%)
     TTXA$ = MKI$(CIKKI%)
     '
     SUMA = 0
     For JJ& = 0 To 12
       XX$ = Mid$(PROGENT$(KKI%), 4 * JJ& + 1, 4)
       SUMA = SUMA + CVS(XX$)
       TTXA$ = TTXA$ + XX$
     Next JJ&
     TTXA$ = TTXA$ + MKS$(SUMA) + String$(128, 0)
     '
     TTXB$ = MKI$(CIKKI%) + MKS$(SUMA)
     For JJ& = 13 To 24
       XX$ = Mid$(PROGENT$(KKI%), 4 * JJ& + 1, 4)
       SUMA = SUMA + CVS(XX$)
       TTXB$ = TTXB$ + XX$
     Next JJ&
     TTXB$ = TTXB$ + MKS$(SUMA) + String$(128, 0)
     '
     KK& = KK& + 1
     Call GRAREG("PROGSTKX", TTXA$, KK&)
     Call GRAREG("PROGSTKY", TTXB$, KK&)
   Next KKI%
   '
   Call SETULTREG("PROGSTKX", KK&)
   Call SETULTREG("PROGSTKY", KK&)
   Call HEADERS("PROGSTKX", "")
   Call HEADERS("PROGSTKX", "Corresponde a: " + DEPROG$ + " - Año 20" + TRIM$(Mid$(HOS$, 7, 2)))
   ASIGUIE$ = TRIM$(Str$(1 + Val(Mid$(HOS$, 7, 2)))): While Len(ASIGUIE$) < 2: ASIGUIE$ = "0" + ASIGUIE$: Wend
   Call HEADERS("PROGSTKY", "")
   Call HEADERS("PROGSTKY", "Corresponde a: " + DEPROG$ + " - Año 20" + TRIM$(ASIGUIE$))
   '
   Call CONECRUN("*PRGSTKX", "Programa de Fabricacion Neto de Stock")

End Sub
'
Private Sub COMMAND15_CLICK()
    Command15.Enabled = False
10  Call SELECHO("PROFABRI")
    NPRG% = Val(VALACT1$("PROFABRI"))
    If NPRG% > 0 Then
      HOII% = HOY(HOS$)
      DEPROG$ = VALACT2$("PROFABRI")
      If Option1.Value = True Then
          Call FILTERFILE("PROGENT", "PROGENFI", 10, MKI$(NPRG%))
          Call FILESORT("PROGENFI", "", 1, 3)
          Call HEADERS("PROGENFI", "")
          Call HEADERS("PROGENFI", "Programa: " + DEPROG$)
          Call CONECRUN("*PROGENFI", "Listado Programa Anual de Entregas")
        Else
          Call FILTERFILE("PRGBRUTA", "PRGBRUTX", 3, MKI$(NPRG%))
          Call FILTERFILE("PRGBRUTB", "PRGBRUTY", 3, MKI$(NPRG%))
          Call HEADERS("PRGBRUTX", "")
          Call HEADERS("PRGBRUTX", "Corresponde a: " + DEPROG$ + " - Año 20" + TRIM$(Mid$(HOS$, 7, 2)))
          ASIGUIE$ = TRIM$(Str$(1 + Val(Mid$(HOS$, 7, 2)))): While Len(ASIGUIE$) < 2: ASIGUIE$ = "0" + ASIGUIE$: Wend
          Call HEADERS("PRGBRUTY", "")
          Call HEADERS("PRGBRUTY", "Corresponde a: " + DEPROG$ + " - Año 20" + TRIM$(ASIGUIE$))
          Call CONECRUN("*PRGBRTX", "Programa Anual de Entregas (Bruto)")
      End If
      GoTo 10
    End If
    Command15.Enabled = True
End Sub
'


