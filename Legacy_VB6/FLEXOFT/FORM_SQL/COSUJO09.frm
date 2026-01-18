VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form COSUJO09 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sueldos y Jornales - Configuracion"
   ClientHeight    =   7035
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5925
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   840
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   28
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Configuración General"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1230
      Left            =   210
      TabIndex        =   23
      Top             =   240
      Width           =   4635
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Actualización - Datos de la Empresa"
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
         Left            =   525
         TabIndex        =   26
         Top             =   315
         Width           =   3500
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Valores de Cambio para Pago"
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
         Left            =   525
         TabIndex        =   25
         Top             =   570
         Width           =   3500
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Calendario de Feriados"
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
         Left            =   525
         TabIndex        =   24
         Top             =   810
         Width           =   3500
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Actualización"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   210
      TabIndex        =   20
      Top             =   5775
      Width           =   4635
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Actualización de Salarios Nominales "
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
         Index           =   10
         Left            =   315
         TabIndex        =   22
         Top             =   630
         Width           =   3855
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Actualización de Parámetros Individuales"
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
         Index           =   9
         Left            =   315
         TabIndex        =   21
         Top             =   360
         Width           =   3855
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Criterios y Tablas "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2730
      Left            =   210
      TabIndex        =   12
      Top             =   1575
      Width           =   4635
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Sectores Productivos"
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
         Index           =   12
         Left            =   525
         TabIndex        =   27
         Top             =   390
         Width           =   3480
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Categorías de Empleados"
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
         Index           =   11
         Left            =   525
         TabIndex        =   19
         Top             =   900
         Width           =   3375
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Grupos de Empleados"
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
         Left            =   525
         TabIndex        =   18
         Top             =   630
         Width           =   3480
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Parámetros Individuales"
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
         Left            =   525
         TabIndex        =   17
         Top             =   1440
         Width           =   3375
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Salarios Nominales Basicos"
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
         Left            =   525
         TabIndex        =   16
         Top             =   1170
         Width           =   3615
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Remuneraciones Ordinarias"
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
         Left            =   525
         TabIndex        =   15
         Top             =   1710
         Width           =   3735
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Retenciones Legales y Aportes"
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
         Left            =   525
         TabIndex        =   14
         Top             =   1980
         Width           =   3855
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Asignaciones y Subsidios Familiares"
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
         Left            =   525
         TabIndex        =   13
         Top             =   2250
         Width           =   3855
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Impresión de Recibo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   210
      TabIndex        =   6
      Top             =   4410
      Width           =   4620
      Begin VB.CommandButton Command8 
         Caption         =   "Abrir"
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
         Left            =   3675
         Picture         =   "COSUJO09.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   315
         Width           =   750
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   210
         TabIndex        =   9
         TabStop         =   0   'False
         Text            =   " "
         Top             =   735
         Width           =   3330
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
         TabIndex        =   8
         Top             =   315
         Width           =   175
      End
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1995
         TabIndex        =   7
         Text            =   " "
         Top             =   315
         Width           =   1380
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Formulario:"
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
         Left            =   -420
         TabIndex        =   11
         Top             =   420
         Width           =   2325
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00008080&
      Height          =   7470
      Left            =   5040
      ScaleHeight     =   7410
      ScaleWidth      =   1320
      TabIndex        =   0
      Top             =   -105
      Width           =   1380
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   4
         Top             =   5955
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   5
            Top             =   105
            Width           =   12000
         End
      End
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
         Height          =   690
         Left            =   105
         Picture         =   "COSUJO09.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Ayuda en Línea"
         Top             =   1800
         Width           =   650
      End
      Begin VB.CommandButton Command2 
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
         Picture         =   "COSUJO09.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Guarda los Cambios Realizados"
         Top             =   1110
         Width           =   650
      End
      Begin VB.CommandButton command1 
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
         Picture         =   "COSUJO09.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Guarda - Cierra - Abandona la Aplicación"
         Top             =   420
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -300
         Picture         =   "COSUJO09.frx":0A68
         Top             =   6225
         Width           =   1515
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   1995
      OleObjectBlob   =   "COSUJO09.frx":298A
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   330
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "COSUJO09.frx":2BBE
      TabIndex        =   29
      Top             =   0
      Visible         =   0   'False
      Width           =   960
   End
End
Attribute VB_Name = "COSUJO09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SBAS(100), PI(100), FORMUSAL$(100)

Private Sub Command1_Click()
  '
  Call CIERRARCH("*.*")
  Unload Me
  '
End Sub

Private Sub Command2_Click()
  '
  Call GRACONTROL("", "RECIHABE", Text7.TEXT)
  '
End Sub

Private Sub Command7_Click()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text7.TEXT = VDEVU$
        Text6.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Command8_Click()
    Call CONECRUN("FLFORMS/" + TRIM$(Text7.TEXT), "Formularios de Impresion")
End Sub

Private Sub Form_Load()
  '
  Call APLICACIONSKINS(Me)
  Call CENTRAFORM(Me)
  COSUJO05.Text7.TEXT = Control$("", "RECIHABE")  'REICBO HABERES
  '
End Sub

Sub Option1_DBLClick(Index As Integer)
  Select Case Index
    Case 3
       Call CARGRUEMP         'GRUPO DE EMPLEADOS
    Case 11
       Call CARCATOPE         'CATEGORIA DE EMPLEADOS
    Case 4
       Call CARSALBAS         'SALARIOS BASICOS
    Case 8
       Call CARPARAINDI       'PARAMETROS INDIVIDUALES
    Case 2
       REMDET05.Show 1        'REMUNERACIONES ORDINARIAS
       Call FRESHALL
       Call HACECONCEP
    Case 1
       RETDET05.Show 1         'RETENCIONES LEGALES Y APORTES
       Call FRESHALL
       Call HACECONCEP
    Case 0
       Call CARASISUB         'ASIGNACIONES FLIARES Y SUBSIDIOS
       Call FRESHALL
       Call HACECONCEP
    Case 7
      Call CARGAFERI          'CARGA DE FERIADOS
    Case 5
      Call CARGEMPRE          'CARGA DE DATOS DE LA EMPRESA
    Case 6
      Call CARGAVALCAM        'CARGA DE VALORES PARA EL CAMBIO
    Case 9
       Call ACTUAPAINDI 'ACTUALIZACION DE PARAMETROS INDIVIDUALES
    Case 10
       Call ACTUASALBAS 'ACTUALIZACION SALARIOS BASICOS
    Case 12
       Call CARSECPROD 'SECTORES PRODUCTIVOS
  End Select
End Sub
Sub CARSECPROD() 'SECTORES PRODUCTIVOS
    '
    Call BLANARCH("!SECPROD")
    J& = 0
    For U& = 1 To ULTREG&("SECPROD")
        X$ = REGLEIDO$("SECPROD", U&)
        COSECC = CVS(Left$(X$, 4))
        If COSECC >= 1 Then
            Call REGAPP("!SECPROD", X$)
        End If
    Next U&
    Call CIERRARCH("SECPROD")
    Call CIERRARCH("!SECPROD")
    '
    ATRI$ = "BORDEIZQ=W" + TRIM$(Str$(Left - 390))
    ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + 850))
    VTB% = VENTABU%("SECPRO2/" + ATRI$) 'XQ SECPROD.PRF ME MUESTRA LOS DEPOSITOS
    If VTB% < 0 Then Exit Sub
    '
    J& = 0
    For U& = 1 To ULTREG&("!SECPROD")
        X$ = REGLEIDO$("!SECPROD", U&)
        COSECC = CVS(Left$(X$, 4))
        If COSECC >= 1 Then
            J& = J& + 1
            Call GRAREG("SECPROD", X$, J&)
        End If
    Next U&
    '
    Call SETULTREG("SECPROD", J&)
    Call CIERRARCH("SECPROD")
    Call CIERRARCH("!SECPROD")
    '
End Sub
Sub CARGAFERI() 'CARGA DE FERIADOS
  '
  Call TRALOCAL("FERIADOS", "FERIADOS")
  '
  ATRI$ = "BORDEIZQ=W" + TRIM$(Str$(Left + 350))
  ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + 150))
  VTB% = VENTABU%("FERIADOS/" + ATRI$)
  If VTB% >= 0 Then
     JJ& = 0
     For U& = 1 To ULTREG&("!FERIADOS")
        X$ = REGLEIDO$("!FERIADOS", U&)
        JJ& = JJ& + 1
        Call GRAREG("FERIADOS", X$, JJ&)
     Next U&
     Call SETULTREG("FERIADOS", JJ&)
     Call CIERRARCH("FERIADOS")
  End If
  '
End Sub
Sub CARGEMPRE()
  '
'  For J& = 1 To ULTREG("DATEMPRE")
'  X$ = REGLEIDO$("DATEMPRE", J&)
'    Select Case J&
'      Case 1
'        RASOEM$ = TRIM$(Mid$(X$, 1, 30))
'        LOCAEM$ = TRIM$(Mid$(X$, 31, 24))
'        CAJUEM$ = TRIM$(Mid$(X$, 55, 16))
'      Case 2
'        FANTEM$ = TRIM$(Mid$(X$, 1, 30))
'        TELEM$ = TRIM$(Mid$(X$, 31, 24))
'        NUIMEM$ = TRIM$(Mid$(X$, 55, 16))
'      Case 3
'        DOMIEM$ = TRIM$(Mid$(X$, 1, 30))
'        DAT1EM$ = TRIM$(Mid$(X$, 31, 24))
'        DAT2EM$ = TRIM$(Mid$(X$, 55, 16))
'    End Select
'  Next J&
'  '
'  With DATEMPRE05
'    .Text25(0) = RASOEM$
'    .Text25(1) = FANTEM$
'    .Text25(2) = DOMIEM$
'    .Text25(3) = LOCAEM$
'    .Text25(4) = TELEM$
'    .Text25(5) = CAJUEM$
'    .Text25(6) = NUIMEM$
'    .Text25(7) = DAT1EM$
'    .Text25(8) = DAT2EM$
'    .MODIFIC = 0
'    .Show 1
'  End With
'  '
'
  For J& = 1 To ULTREG("DATEMPRE")
  X$ = REGLEIDO$("DATEMPRE", J&)
    Select Case J&
      Case 1
        RASOEM$ = TRIM$(Mid$(X$, 1, 30))
        LOCAEM$ = TRIM$(Mid$(X$, 31, 24))
        CAJUEM$ = TRIM$(Mid$(X$, 55, 16))
      Case 2
        FANTEM$ = TRIM$(Mid$(X$, 1, 30))
        TELEM$ = TRIM$(Mid$(X$, 31, 24))
        NUIMEM$ = TRIM$(Mid$(X$, 55, 16))
      Case 3
        DOMIEM$ = TRIM$(Mid$(X$, 1, 30))
        DAT1EM$ = TRIM$(Mid$(X$, 31, 24))
        DAT2EM$ = TRIM$(Mid$(X$, 55, 16))
    End Select
  Next J&
  '
  Call REPLA(VDEF$, RASOEM$, 1, 30)
  Call REPLA(VDEF$, FANTEM$, 31, 30)
  Call REPLA(VDEF$, DOMIEM$, 61, 30)
  Call REPLA(VDEF$, LOCAEM$, 91, 24)
  Call REPLA(VDEF$, TELEM$, 115, 24)
  Call REPLA(VDEF$, CAJUEM$, 139, 16)
  Call REPLA(VDEF$, NUIMEM$, 155, 24)
  Call REPLA(VDEF$, DAT1EM$, 171, 24)
  Call REPLA(VDEF$, DAT2EM$, 195, 24)
  Z$ = OBJPLA$("DATAEMPRE", VDEF$)
  If TRIM$(Z$) <> "" Then
      RASOEM$ = Mid$(Z$, 1, 30)
      FANTEM$ = Mid$(Z$, 31, 30)
      DOMIEM$ = Mid$(Z$, 61, 30)
      LOCAEM$ = Mid$(Z$, 91, 24)
      TELEM$ = Mid$(Z$, 115, 24)
      CAJUEM$ = Mid$(Z$, 139, 16)
      NUIMEM$ = Mid$(Z$, 155, 24)
      DAT1EM$ = Mid$(Z$, 171, 24)
      DAT2EM$ = Mid$(Z$, 195, 24)
      '
      Call BLANARCH("DATEMPRE")
      YY$ = Space$(128)
      Call REPLA(YY$, RASOEM$, 1, 30)
      Call REPLA(YY$, LOCAEM$, 31, 24)
      Call REPLA(YY$, CAJUEM$, 55, 16)
      Call REGAPP("DATEMPRE", YY$)
      '
      YY$ = Space$(128)
      Call REPLA(YY$, FANTEM$, 1, 30)
      Call REPLA(YY$, TELEM$, 31, 24)
      Call REPLA(YY$, NUIMEM$, 55, 16)
      Call REGAPP("DATEMPRE", YY$)
      '
      YY$ = Space$(128)
      Call REPLA(YY$, DOMIEM$, 1, 30)
      Call REPLA(YY$, DAT1EM$, 31, 24)
      Call REPLA(YY$, DAT2EM$, 55, 16)
      Call REGAPP("DATEMPRE", YY$)
      '
      Call CIERRARCH("DATEMPRE")
      '
  End If
  '
End Sub
Sub CARGAVALCAM()
  '
    '
  Dim VALBILL(9)
  '
  For J& = 1 To ULTREG("BILLETES")
  X$ = REGLEIDO$("BILLETES", J&)
    '
    For JJ& = 1 To 40 Step 4
      VAL1 = CVS(Mid$(X$, JJ&, 4)) 'CADA 4 BYTE UN BILLETE
      VALBILL(CONTI1%) = VAL1
      CONTI1% = CONTI1% + 1
      ZJ% = Int(JJ&)
      Call REPLA(VDEF$, MKS$(VAL1), ZJ%, 4)
    Next JJ&
    '
  Next J&
  '
  OBJ1$ = OBJPLA$("VALCAMPA", VDEF$)
    '
  If TRIM(OBJ1$) <> "" Then
    VALBILL(0) = XVALO(CVS(Mid$(OBJ1$, 1, 4)))
    VALBILL(1) = XVALO(CVS(Mid$(OBJ1$, 5, 4)))
    VALBILL(2) = XVALO(CVS(Mid$(OBJ1$, 9, 4)))
    VALBILL(3) = XVALO(CVS(Mid$(OBJ1$, 13, 4)))
    VALBILL(4) = XVALO(CVS(Mid$(OBJ1$, 17, 4)))
    VALBILL(5) = XVALO(CVS(Mid$(OBJ1$, 21, 4)))
    VALBILL(6) = XVALO(CVS(Mid$(OBJ1$, 25, 4)))
    VALBILL(7) = XVALO(CVS(Mid$(OBJ1$, 29, 4)))
    VALBILL(8) = XVALO(CVS(Mid$(OBJ1$, 33, 4)))
    VALBILL(9) = XVALO(CVS(Mid$(OBJ1$, 37, 4)))
    '
    Call BLANARCH("BILLETES")
    YY$ = Space$(40)
    'CADA 4 BYTES GRABO EL VALOR DE UN BILLETE
    For JJ& = 1 To 40 Step 4
      Call REPLA(YY$, MKS$(VALBILL(CONTI)), Int(JJ& * 1), 4)
      CONTI = CONTI + 1
    Next JJ&
    Call REGAPP("BILLETES", YY$)
    '
    Call CIERRARCH("BILLETES")
  End If
  '
  
'  Dim VALBILL(9)
'  '
'  For J& = 1 To ULTREG("BILLETES")
'  X$ = REGLEIDO$("BILLETES", J&)
'    '
'    For JJ& = 1 To 40 Step 4
'      VAL1 = CVS(Mid$(X$, JJ&, 4)) 'CADA 4 BYTE UN BILLETE
'      VALBILL(CONTI1%) = VAL1
'      CONTI1% = CONTI1% + 1
'    Next JJ&
'    '
'  Next J&
'  '
'  With VALCAM05
'    For JJJ% = 0 To 9
'      .Text25(JJJ%) = FORMATNUM$(VALBILL(JJJ%), "F10.2")
'    Next JJJ%
'    .MODIFIC = 0
'    .Show 1
'  End With
'  '
End Sub

Private Sub Text7_Change()
    '
    VDEVU$ = Text7
    If VDEVU$ <> "" Then
        Text7.TEXT = VDEVU$
        Text6.TEXT = ECOARCH$("INDIFRM", Text7.TEXT)
    End If
    '
End Sub
Sub ACTUAPAINDI() 'ACTUALIZACION DE PARAMETROS INDIVIDUALES
  '
  RTA% = COMALTER%("Modalidad de Actualización\\Por Legajo\Por Concepto")
  Select Case RTA%
    Case 1
      Call ACTPARLEG
    Case 2
      Call ACTPARCON
  End Select
99 End Sub
Sub ACTUASALBAS()
  '
9 TIP3% = COMALTER%("Seleccione el Tipo de Consulta\\Actualización\Listado\Cancelar")
  '
  If TIP3% = 1 Then 'Si Actualización
10  RTA% = COMALTER%("Seleccione el Modo de Actualización\\Interactiva\Automática\Cancelar")
    Select Case RTA%
      Case 1
        Call ACTINTER  'Act.Interactiva
        GoTo 10
      Case 2
        Call ACTAUTOM(COD3%)  'Act.Automatica - COD3%= GRUPO EMPLEADOS
        GoTo 10
      Case Else
        GoTo 9
    End Select
  ElseIf TIP3% = 2 Then 'Si Listado
    'RTA1% = COMALTER%("Seleccione el Modo de Listado\\Planilla\Detalle\Cancelar")
    'Select Case RTA1%
    '  Case 1
        'Call LIPLANI
    '  Case 2
        Call LIDETA
        GoTo 9
    '  Case Else
    '    GoTo 9
    'End Select
  End If
  '
99 End Sub
'
Sub LIDETA()
  '
  Call BLANARCH("LIACTSAL")
  Y$ = REGBLAN$("LIACTSAL")
  J2& = 0
  For J& = 1 To ULTREG("PADREMP")
    X$ = REGLEIDO$("PADREMP", J&)
    NLEGAX% = CVI(Mid$(X$, 1, 2))
    REG$ = RECFILT$("PADREMP", 1, MKI$(NLEGAX%))
    REGI1& = CVS(REG$) 'NUMERO DE REGISTRO EN PADRON
    '
    For J1& = 1 To ULTREG&("NOMREM")
      X$ = REGLEIDO$("NOMREM", J1&)
      DESCRI1$ = Mid$(X$, 1, 30)
      VALOACT1 = VALOSAL(J1&, REGI1&) 'VALOR DEL PARAMETRO
      Call REPLA(Y$, MKI$(NLEGAX%), 1, 2)
      Call REPLA(Y$, DESCRI1$, 3, 30)
      Call REPLA(Y$, MKS$(VALOACT1), 33, 4)
      J2& = J2& + 1
      Call GRAREG("LIACTSAL", Y$, J2&)
    Next J1&
  Next J&
  Call SETULTREG("LIACTSAL", J2&)
  '
  Call FINAL("*LIACTSAL")
  '
End Sub
'
Sub ACTAUTOM(GRUP3%)
  'TENGO QUE CARGAR EL ARCHIVOS CON LOS SALARIOS EXISTENTES
  Call BLANARCH("SALACTAU")
  For J& = 1 To ULTREG&("NOMREM")
   X$ = REGLEIDO$("NOMREM", J&)
   COD3$ = "SB" + CStr(J&)
   DENOMI3$ = Mid$(X$, 1, 30)
   MOD3% = Asc(Mid$(X$, 39, 1))
   OBSERVA3$ = ""
   If MOD3% = 2 Then OBSERVA3$ = "FORMULA"
   Y$ = Space$(64)
   Call REPLA(Y$, COD3$, 1, 4)
   Call REPLA(Y$, DENOMI3$, 5, 30)
   Call REPLA(Y$, OBSERVA3$, 43, 14)
   Call REGAPP("SALACTAU", Y$)
  Next J&
  Call CIERRARCH("SALACTAU")
  '
  Call TRALOCAL("SALACTAU", "SALACTAU")
  '
  ATRI$ = "BORDEIZQ=W" + TRIM$(Str$(Left - 750))
  ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + 7))
10  VTB% = VENTABU%("SALACTAU/" + ATRI$)
  If VTB% >= 0 Then
     JJ& = 0
     For U& = 1 To ULTREG&("!SALACTAU")
        X$ = REGLEIDO$("!SALACTAU", U&)
        JJ& = JJ& + 1
        Call GRAREG("SALACTAU", X$, JJ&)
     Next U&
     Call SETULTREG("SALACTAU", JJ&)
     Call CIERRARCH("SALACTAU")
     GoTo 15
  End If
  Exit Sub
  '
15  'RECALCULO DE SALARIOS
  For J& = 1 To ULTREG("PADREMP")
    X$ = REGLEIDO$("PADREMP", J&)
    NUGRU3% = Asc(Mid$(X$, 49, 1))
    If NUGRU3% = GRUP3% Or GRUP3% = 0 Then 'SI FILTRO GRUPO EMP. O TODOS
      XX$ = REGLEIDO$("REMBAS", J&)
      For JJ& = 1 To ULTREG&("SALACTAU")
        XXX$ = REGLEIDO$("SALACTAU", JJ&)
        MOD1% = MOCASAL%(JJ&)
        AUMEN = CVS(Mid$(XXX$, 35, 4))
        VRED = CVS(Mid$(XXX$, 39, 4))
        If MOD1% = 2 And AUMEN > 0.005 Then
          Call MENSERR(24, "No es Posible Actualizar un Salario \con Modo de Recalculo por Fórmula.")
          GoTo 10
        End If
        If AUMEN > 0.005 And MOD1% = 1 Then 'SI %AUM>0
          JJJ& = JJ&
          POSCOM& = (JJJ& - 1) * 4 + 1 'CADA 4 BYTES UN SALARIO
          SALANT = CVS(Mid$(XX$, POSCOM&, 4)) 'LEE POSICION DE SALARIO
          If SALANT > 0 Then 'SI TIENE VALOR ESE SALARIO AUMENTO
            SALAR = SALANT * (1 + AUMEN / 100)
              If VRED < 0.0001 Then
                VRED = 0.0001
              End If
            SALAR = (Int(SALAR / VRED + 0.5)) * VRED
            POSCOM1% = Val(POSCOM)
            Call REPLA(XX$, MKS$(SALAR), POSCOM1%, 4)
            Call GRAREG("REMBAS", XX$, J&) 'ACTUALIZO REMBAS
          End If
        End If
      Next JJ&
    End If
  Next J&
  '
End Sub
Sub ACTINTER() 'ACTUALIZACION INTERACTIVA SALARIO NOMIN BASICOS
  '
10 RTA% = COMALTER%("Modalidad de Actualización\\Por Legajo\Por Salario\Cancelar")
  Select Case RTA%
    Case 1
      Call ACTSALLEG
      GoTo 10
    Case 2
      Call ACTSALNOM
      GoTo 10
    Case 3
      Exit Sub
  End Select
  '
End Sub
Sub ACTSALLEG() 'ATUALIZACION INTERACTIVA SALARIO X LEGAJO
  '
10   Call SELECHO("PADREMP")
  NLEG% = Val(TRIM$(VALACT1$("PADREMP")))
  If NLEG% < 1 Then Exit Sub 'NO ELIGIO NINGUN LEGAJO
  REG$ = RECFILT$("PADREMP", 1, MKI$(NLEG%))
  REGI1& = CVS(REG$) 'NUMERO DE REGISTRO EN PADRON
  'CARGAR PARAMTROS EN  ACTPARLE.DAT Y VALORES
  Call BLANARCH("!ACTSALEG")
  Y$ = REGBLAN$("!ACTSALEG")
  For J& = 1 To ULTREG&("NOMREM")
    X$ = REGLEIDO$("NOMREM", J&)
    FORMUSAL$(J&) = TRIM$(Mid$(X$, 43, 64))
    DESCRI1$ = Mid$(X$, 1, 30)
    VALOACT1 = VALOSAL(J&, REGI1&) 'VALOR DEL PARAMETRO
    Call REPLA(Y$, DESCRI1$, 1, 30)
    Call REPLA(Y$, MKS$(VALOACT1), 31, 4)
    Call REPLA(Y$, MKS$(VALOACT1), 35, 4)
    Call GRAREG("!ACTSALEG", Y$, J&)
  Next J&
  Call SETULTREG("!ACTSALEG", J&)
  '
  'MOSTRA VENTANA ACTSALEG.PRF
15  VTB% = VENTABU%("ACTSALEG/BORDEIZQ=10/BORDESUP=7/NC/TITUPAN=" + "Actualización de Salarios - " + TRIM$(ECOARCH$("PADREMP", MKI$(NLEG%))))
    If VTB% >= 0 Then
    For J& = 1 To ULTREG&("!ACTSALEG")
      'ACA VA EL CODIGO PARA GRABAR EL REMBAS ACTUALIZADO
      X$ = REGLEIDO$("!ACTSALEG", J&)
      MOD1% = MOCASAL%(J&)
      SBAS(J&) = 0
      If MOD1% = 1 Then
        VALOACT2 = CVS(Mid$(X$, 35, 4))
        SBAS(J&) = VALOACT2
      End If
    Next J&
    '
    For J& = 1 To ULTREG&("!ACTSALEG")
      X$ = REGLEIDO$("!ACTSALEG", J&)
      MOD1% = MOCASAL%(J&)
      If MOD1% = 2 Then
        VALOACT2 = SALCALCU(FORMUSAL$(J&))
        SBAS(J&) = VALOACT2
      End If
    Next J&
    '
    For J& = 1 To ULTREG&("!ACTSALEG")
'      If VALOACT2 > 0.005 And MOD1% = 2 Then
'        Call MENSERR(24, "No es Posible Actualizar un Salario \con Modo de Recalculo por Fórmula.")
'        GoTo 15
'      End If
       VALOACT2 = SBAS(J&)
      Call ACTUVALSA(VALOACT2, J&, REGI1&) 'ACTUALIZA VALOR
    Next J&
'    GoTo 10
'  Else
'    GoTo 10
  End If
  GoTo 10 'vuelve a presentar la lista
  '
End Sub

Function SALCALCU(FORMORI4$)
      '
      Dim RI3#(100)
      '
17000 FORMUTRAB$ = FORMORI4$: RI% = 0: RI2# = 0
17005 PINI% = InStr(FORMUTRAB$, "@")
        If PINI% < 1 Then
            GoTo 17025
        End If
17010 GoSub 17060              ' RESOLUCION SUMATORIA
17015 GoTo 17005
17020 '
17025 PINI% = InStr(FORMUTRAB$, "(")
        If PINI% < 1 Then
            GoTo 17045
        End If
17030 GoSub 17155              ' RESOLUCION PARENTESIS
17035 GoTo 17025
17040 '
17045 CALFIN% = 0: GoSub 17220 ' RESOLUCION FORMULA ALGEBRAICA
17050 SALCALCU = RI2#
      Exit Function
17055 '
17060 IZQUI$ = ""
        If PINI% > 1 Then
            IZQUI$ = Left$(FORMUTRAB$, PINI% - 1)
            FORMUTRAB$ = Mid$(FORMUTRAB$, PINI%)
        End If
17065 VARI4$ = Mid$(FORMUTRAB$, 2, 2): FORMUTRAB$ = Mid$(FORMUTRAB$, 4)
17070 LINF% = Val(Left$(FORMUTRAB$, 2)): FORMUTRAB$ = Mid$(FORMUTRAB$, 3)
17075 If Left$(FORMUTRAB$, 1) = "." Then
            FORMUTRAB$ = Mid$(FORMUTRAB$, 2)
            GoTo 17075
      End If
17080 LSUP% = Val(Left$(FORMUTRAB$, 2))
17085 If LSUP% < 10 Then
            DERE$ = Mid$(FORMUTRAB$, 2)
      Else
          DERE$ = Mid$(FORMUTRAB$, 3)
      End If
17090 RI2# = 0
17095 For RANGO% = LINF% To LSUP%
17100 VALOR# = 0
17105 If VARI4$ = "SB" Then
            VALOR# = SBAS(RANGO%)
            GoTo 17125
      End If
17110 If VARI4$ = "PI" Then
            VALOR# = PI(RANGO%)
            GoTo 17125
      End If
17125 RI2# = RI2# + (Int(100 * VALOR# + 0.5)) / 100
17130 Next RANGO%
17135 RI% = RI% + 1: RI3#(RI%) = RI2#: RI4$ = "RI" + Mid$(Str$(RI%), 2)
17140 FORMUTRAB$ = IZQUI$ + RI4$ + DERE$
17145 Return
17150 '
17155 IZQUI$ = ""
        If PINI% > 1 Then
            IZQUI$ = Left$(FORMUTRAB$, PINI% - 1)
        End If
17160 FORMUTRAB$ = Mid$(FORMUTRAB$, PINI% + 1)
17165 PINI% = InStr(FORMUTRAB$, "(")
        If PINI% > 0 Then
            IZQUI$ = IZQUI$ + "(" + Left$(FORMUTRAB$, PINI% - 1)
            FORMUTRAB$ = Mid$(FORMUTRAB$, PINI% + 1)
            GoTo 17165
        End If
17170 PFIN% = InStr(FORMUTRAB$, ")")
        If PFIN% < 1 Then
            PFIN% = 1 + Len(FORMUTRAB$)
        End If
17175 DERE$ = Mid$(FORMUTRAB$, PFIN% + 1)
17180 FORMUTRAB$ = Left$(FORMUTRAB$, PFIN% - 1)
17185 '
17190 CALFIN% = 1: GoSub 17220       ' RESOLUCION FORMULA SIN PARENTESIS
17195 '
17200 RI% = RI% + 1: RI3#(RI%) = RI2#: RI4$ = "RI" + Mid$(Str$(RI%), 2)
17205 FORMUTRAB$ = IZQUI$ + RI4$ + DERE$
17210 Return
17215 '
17220 POPE% = InStr(FORMUTRAB$, "x")
        If POPE% > 0 Then
            OPE% = 1
            GoTo 17285
        End If
17225 POPE% = InStr(FORMUTRAB$, "X")
        If POPE% > 0 Then
            OPE% = 1
            GoTo 17285
        End If
17230 POPE% = InStr(FORMUTRAB$, "*")
        If POPE% > 0 Then
            OPE% = 1
            GoTo 17285
        End If
17235 POPE% = InStr(FORMUTRAB$, "/")
        If POPE% > 0 Then
            OPE% = 2
            GoTo 17285
        End If
17240 '
17245 POPE% = InStr(FORMUTRAB$, "+")
        If POPE% > 0 Then
            OPE% = 3
            GoTo 17285
        End If
17250 '
17255 POPE% = InStr(FORMUTRAB$, "-")
        If POPE% > 0 Then
            OPE% = 4
            GoTo 17285
        End If
17260 '
17265 ' NO QUEDAN OPERACIONES ARITMETICAS POR CONSIDERAR
17270 If Abs(RI2#) < 0.005 Then
            If CALFIN% = 0 Then IZQUIE$ = "": DERIE$ = "": VALO1# = 0: ARGU2$ = FORMUTRAB$: OPE% = 3: CALFIN% = 1: GoTo 17470
      End If
17275 Return
17280 '
17285 ARG1$ = "": ARG2$ = ""
17290 For PP% = POPE% - 1 To 1 Step -1
17295 '
17300 CR$ = Mid$(FORMUTRAB$, PP%, 1)
17305 If CR$ = "x" Or CR$ = "X" Or CR$ = "*" Or CR$ = "/" Or CR$ = "+" Or CR$ = "-" Then
            GoTo 17325
      End If
17310 Next PP%
17315 PP% = 0
17320 '
17325 ARG1$ = Mid$(FORMUTRAB$, PP% + 1, POPE% - PP% - 1)
        IZQUIE$ = ""
        If PP% > 0 Then
            IZQUIE$ = Left$(FORMUTRAB$, PP%)
        End If
17330 '
17335 For PP% = POPE% + 1 To Len(FORMUTRAB$)
17340 CR$ = Mid$(FORMUTRAB$, PP%, 1)
17345 If CR$ = "x" Or CR$ = "X" Or CR$ = "*" Or CR$ = "/" Or CR$ = "+" Or CR$ = "-" Then
            GoTo 17365
      End If
17350 Next PP%
17355 PP% = 1 + Len(FORMUTRAB$)
17360 '
17365 ARG2$ = Mid$(FORMUTRAB$, POPE% + 1, PP% - POPE% - 1): DERIE$ = Mid$(FORMUTRAB$, PP%)
17370 '
17375 ARGU1$ = ""
        For PP% = 1 To Len(ARG1$)
        CR1% = Asc(Mid$(ARG1$, PP%))
        If (CR1% > 47 Or CR1% = 46) Then
            ARGU1$ = ARGU1$ + Chr$(CR1%)
        End If
17380 Next PP%
17385 '
17390 ARGU2$ = ""
        For PP% = 1 To Len(ARG2$)
        CR1% = Asc(Mid$(ARG2$, PP%))
        If (CR1% > 47 Or CR1% = 46) Then
            ARGU2$ = ARGU2$ + Chr$(CR1%)
        End If
17395 Next PP%
17400 '
17405 VALO1# = 0
17410 If Left$(ARGU1$, 2) = "AE" Then
'            VALO1# = AE
            GoTo 17470
      End If
17415 If Left$(ARGU1$, 2) = "AC" Then
            VALO1# = AC
            GoTo 17470
      End If
17425 INDI% = Val(Mid$(ARGU1$, 3))
        If INDI >= 0 Then
            If INDI < 100 Then INDI% = Int(INDI)
        End If
17430 If Left$(ARGU1$, 2) = "RI" Then
            VALO1# = RI3#(INDI%)
            GoTo 17470
      End If
17435 If Left$(ARGU1$, 2) = "SB" Then
            VALO1# = SBAS(INDI%)
            GoTo 17470
      End If
17450 If Left$(ARGU1$, 2) = "PI" Then
            VALO1# = PI(INDI%)
            GoTo 17470
      End If
17455 '
17460 VALO1# = Val(ARGU1$)
17465 '
17470 VALO2# = 0
17475 If Left$(ARGU2$, 2) = "AE" Then
'            VALO2# = AE
            GoTo 17535
      End If
17480 If Left$(ARGU2$, 2) = "AC" Then
            VALO2# = AC
            GoTo 17535
      End If
17490 INDI = Val(Mid$(ARGU2$, 3))
        If INDI >= 0 Then
            If INDI < 100 Then INDI% = Int(INDI)
        End If
17495 If Left$(ARGU2$, 2) = "RI" Then
            VALO2# = RI3#(INDI%)
            GoTo 17535
      End If
17500 If Left$(ARGU2$, 2) = "SB" Then
            VALO2# = SBAS(INDI%)
            GoTo 17535
      End If
17515 If Left$(ARGU2$, 2) = "PI" Then
            VALO2# = PI(INDI%)
            GoTo 17535
      End If
17520 '
17525 VALO2# = Val(ARGU2$)
17530 '
17535 RI2# = 0
17540 If OPE% = 1 Then
            RI2# = VALO1# * VALO2#
            GoTo 17565
      End If
17545 If OPE% = 2 Then
            If Abs(VALO2#) >= 0.005 Then RI2# = VALO1# / VALO2#: GoTo 17565
      End If
17550 If OPE% = 3 Then
            RI2# = VALO1# + VALO2#
            GoTo 17565
      End If
17555 If OPE% = 4 Then
            RI2# = VALO1# - VALO2#
            GoTo 17565
      End If
17560 '
17565 RI% = RI% + 1: RI3#(RI%) = RI2#: RI4$ = "RI" + Mid$(Str$(RI%), 2)
17570 FORMUTRAB$ = IZQUIE$ + RI4$ + DERIE$
17575 GoTo 17220
17580 '

End Function

Function VALOSAL(NUMPAR&, REGI1&)
  'ARGUMENTOS: NºSALARIO - Nº DE EMPLEADO
  'FUNCION: DEVUELVE EL VALOR DEL SALARIO DE EMPLEADO
  XX$ = REGLEIDO$("REMBAS", REGI1&)
  J& = ((NUMPAR& + 1) * 4) - 7
  VALOSAL = CVS(Mid$(XX$, J&, 4))  'DEVUELVE VALOR SALARIO
  '
End Function
Sub ACTUVALSA(VALO1, NUMPAR&, REGI1&)
  'ARGUMENTOS: VALOR - NºSALARIO - Nº EMPLEADO
  'FUNCION: ACTUALIZA EL SALARIO DEL EMPLEADO
  XX$ = REGLEIDO$("REMBAS", REGI1&)
  J% = ((NUMPAR& + 1) * 4) - 7
  Call REPLA(XX$, MKS$(VALO1), J%, 4) 'ACTUALIZA SALARIO
  Call GRAREG("REMBAS", XX$, REGI1&)
  '
End Sub
Sub ACTSALNOM() 'ACTUALIZACION INTERATIVA SALARIO X SALARIO
  'ARMO EL INDICE DE SALARIOS
  Call BLANARCH("INDISALA")
  For J& = 1 To ULTREG("NOMREM")
    XX$ = REGLEIDO$("NOMREM", J&)
    DESCRI1$ = TRIM$(Mid$(XX$, 1, 30))
    Y$ = REGBLAN$("INDISALA")
    NU1% = Int(J&)
    Call REPLA(Y$, MKI$(NU1%), 1, 2)
    Call REPLA(Y$, DESCRI1$, 3, 30)
    Call GRAREG("INDISALA", Y$, J&)
  Next J&
  Call SETULTREG("INDISALA", J&)
  '
10 Call SELECHO("INDISALA")
  '
  NROPAR& = Val(TRIM$(VALACT1$("INDISALA")))
  If NROPAR& <= 0 Then Exit Sub 'NO ELIGIO NINGUN CONCEPTO
  'CARGO EN ACTPARCO LOS EMPLEADOS Y EL VALOR DEL PARAMETRO
  Call BLANARCH("!ACTSALNO")
  For J& = 1 To ULTREG("PADREMP")
    XX$ = REGLEIDO$("PADREMP", J&)
    NROEMP% = CVI(Mid$(XX$, 1, 2))
    NOAP1$ = TRIM$(Mid$(XX$, 3, 30))
    VALO2 = VALOSAL(NROPAR&, J&)
    '
    Y$ = Space(64)
    Call REPLA(Y$, MKI$(NROEMP%), 1, 2)
    Call REPLA(Y$, NOAP1$, 3, 30)
    Call REPLA(Y$, MKS$(VALO2), 33, 4)
    Call REPLA(Y$, MKS$(VALO2), 37, 4)
    '
    Call GRAREG("!ACTSALNO", Y$, J&)
  Next J&
  '
  MOD1% = MOCASAL%(NROPAR&)
  If MOD1% = 2 Then
    Call MENSERR(24, "No es Posible Actualizar un Salario \con Modo de Recalculo por Fórmula.")
    GoTo 10
  End If
  NROPA1% = Val(NROPAR&)
  VTB% = VENTABU("ACTSALNO/BORDEIZQ=8/BORDESUP=7/NC/TITUPAN=" + "Parámetro: " + TRIM$(ECOARCH$("INDISALA", MKI$(NROPA1%))))
  If VTB% >= 0 Then
    'ACA VA EL CODIGO PARA GRABAR EL REMBAS ACTUALIZADO
    For J& = 1 To ULTREG&("!ACTSALNO")
      X$ = REGLEIDO$("!ACTSALNO", J&)
      VALOACT2 = CVS(Mid$(X$, 37, 4))
      Call ACTUVALSA(VALOACT2, NROPAR&, J&) 'ACTUALIZA VALOR
    Next J&
    GoTo 10
    '
  Else
    GoTo 10
  End If
  '
End Sub
Function MOCASAL%(NROSAL&)
  '
  XX$ = REGLEIDO$("NOMREM", NROSAL&)
  MOCASAL% = Asc(Mid$(XX$, 39, 1))
  '
End Function
Sub ACTPARLEG() 'ACTUALIZACION PARAME. INDIVI. POR LEGAJO
  '
10 Call SELECHO("PADREMP")
  NLEG% = Val(TRIM$(VALACT1$("PADREMP")))
  If NLEG% < 1 Then Exit Sub 'NO ELIGIO NINGUN LEGAJO
  REG$ = RECFILT$("PADREMP", 1, MKI$(NLEG%))
  REGI1& = CVS(REG$) 'NUMERO DE REGISTRO EN PADRON
  'CARGAR PARAMTROS EN  ACTPARLE.DAT Y VALORES
  Call BLANARCH("!ACTPARLE")
  Y$ = REGBLAN$("!ACTPARLE")
  For J& = 1 To ULTREG&("PARAME")
    X$ = REGLEIDO$("PARAME", J&)
    DESCRI1$ = Mid$(X$, 1, 30)
    VALOACT1 = VALOPAR(J&, REGI1&) 'VALOR DEL PARAMETRO
    Call REPLA(Y$, DESCRI1$, 1, 30)
    Call REPLA(Y$, MKS$(VALOACT1), 31, 4)
    Call REPLA(Y$, MKS$(VALOACT1), 35, 4)
    Call GRAREG("!ACTPARLE", Y$, J&)
  Next J&
  Call SETULTREG("!ACTPARLE", J&)
  '
  'MOSTRA VENTANA ACTPARLE.PRF
  VTB% = VENTABU%("ACTPARLE/BORDEIZQ=10/BORDESUP=7/NC/TITUPAN=" + "Actualización de Parámetros - " + TRIM$(ECOARCH$("PADREMP", MKI$(NLEG%))))
  If VTB% >= 0 Then
    For J& = 1 To ULTREG&("!ACTPARLE")
      'ACA VA EL CODIGO PARA GRABAR EL REMBAS ACTUALIZADO
      X$ = REGLEIDO$("!ACTPARLE", J&)
      VALOACT2 = CVS(Mid$(X$, 35, 4))
      Call ACTUAVAL(VALOACT2, J&, REGI1&) 'ACTUALIZA VALOR
    Next J&
    GoTo 10
  Else
    GoTo 10
  End If
  '
End Sub
Sub ACTUAVAL(VALO1, NUMPAR&, REGI1&)
  'ARGUMENTOS: VALOR - NºPARAMETRO- Nº EMPLEADO
  'FUNCION: ACTUALIZA EL VALOR DEL PARAMETRO DEL EMPLEADO
  XX$ = REGLEIDO$("REMBAS", REGI1&)
  J% = ((NUMPAR& + 1) * 4) - 7
  Call REPLA(XX$, MKS$(VALO1), 48 + J%, 4) 'ACTUALIZA VALOR
  Call GRAREG("REMBAS", XX$, REGI1&)
  '
End Sub
Function VALOPAR(NUMPAR&, REGI1&)
  'ARGUMENTOS: NºPARAMETRO - Nº DE EMPLEADO
  'FUNCION: DEVUELVE EL VALOR DEL PARAMETRO DEL EMPLEADO
  XX$ = REGLEIDO$("REMBAS", REGI1&)
  J& = ((NUMPAR& + 1) * 4) - 7
  VALOPAR = CVS(Mid$(XX$, 48 + J&, 4)) 'DEVUELVE VALOR
  '
End Function
Sub ACTPARCON() 'ACTUALIZACION PARAME. INDIVI. POR CONCEPTO
  'ARMO EL INDICE DE PARAMETROS
  Call BLANARCH("INDIPARM")
  For J& = 1 To ULTREG("PARAME")
    XX$ = REGLEIDO$("PARAME", J&)
    DESCRI1$ = TRIM$(Mid$(XX$, 1, 30))
    Y$ = REGBLAN$("INDIPARM")
    NU1% = Int(J&)
    Call REPLA(Y$, MKI$(NU1%), 1, 2)
    Call REPLA(Y$, DESCRI1$, 3, 30)
    Call GRAREG("INDIPARM", Y$, J&)
  Next J&
  Call SETULTREG("INDIPARM", J&)
  '
10 Call SELECHO("INDIPARM")
  NROPAR& = Val(TRIM$(VALACT1$("INDIPARM")))
  If NROPAR& <= 0 Then Exit Sub 'NO ELIGIO NINGUN CONCEPTO
  'CARGO EN ACTPARCO LOS EMPLEADOS Y EL VALOR DEL PARAMETRO
  Call BLANARCH("!ACTPARCO")
  For J& = 1 To ULTREG("PADREMP")
    XX$ = REGLEIDO$("PADREMP", J&)
    NROEMP% = CVI(Mid$(XX$, 1, 2))
    NOAP1$ = TRIM$(Mid$(XX$, 3, 30))
    VALO2 = VALOPAR(NROPAR&, J&)
    '
    Y$ = Space(64)
    Call REPLA(Y$, MKI$(NROEMP%), 1, 2)
    Call REPLA(Y$, NOAP1$, 3, 30)
    Call REPLA(Y$, MKS$(VALO2), 33, 4)
    Call REPLA(Y$, MKS$(VALO2), 37, 4)
    '
    Call GRAREG("!ACTPARCO", Y$, J&)
  Next J&
  '
  NROPA1% = Val(NROPAR&)
  VTB% = VENTABU("ACTPARCO/BORDEIZQ=8/BORDESUP=7/NC/TITUPAN=" + "Parámetro: " + TRIM$(ECOARCH$("INDIPARM", MKI$(NROPA1%))))
  If VTB% >= 0 Then
    'ACA VA EL CODIGO PARA GRABAR EL REMBAS ACTUALIZADO
    For J& = 1 To ULTREG&("!ACTPARCO")
      X$ = REGLEIDO$("!ACTPARCO", J&)
      VALOACT2 = CVS(Mid$(X$, 37, 4))
      Call ACTUAVAL(VALOACT2, NROPAR&, J&) 'ACTUALIZA VALOR
    Next J&
    GoTo 10
    '
  Else
    GoTo 10
  End If
  '
End Sub

Sub CARGRUEMP()
  '
  Call TRALOCAL("DATGREMP", "DATGREMP")
  Call TRALOCAL("DATGREM1", "DATGREM1")
  '
  ATRI$ = "BORDEIZQ=W" + TRIM$(Str$(Left - 1800))
  ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + 850))
  VTB% = VENTABU%("DATGREMP/" + ATRI)
  If VTB% >= 0 Then
     JJ& = 0
     For U& = 1 To ULTREG&("!DATGREMP")
        X$ = REGLEIDO$("!DATGREMP", U&)
         JJ& = JJ& + 1
        Call GRAREG("DATGREMP", X$, JJ&)
        '
         Z$ = Space(27)
         Call REPLA(Z$, Chr$(JJ&), 1, 1)
         Call REPLA(Z$, TRIM$(Mid$(X$, 1, 25)), 2, 25)
        Call GRAREG("DATGREM1", Z$, JJ&) 'INIDICE/ECO
        '
     Next U&
     Call SETULTREG("DATGREMP", JJ&)
     Call CIERRARCH("DATGREMP")
     Call SETULTREG("DATGREM1", JJ&) 'INDICE/ECO
     Call CIERRARCH("DATGREM1") 'INDICE/ECO
  End If
  '
End Sub

Sub CARCATOPE()
  '
  Call TRALOCAL("CATEGO", "CATEGO")
  '
  ATRI$ = "BORDEIZQ=W" + TRIM$(Str$(Left - 1400))
  ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + 850))
  VTB% = VENTABU%("CATEGO/" + ATRI$)
  If VTB% >= 0 Then
     APLIBAS% = 0
     If COMALTER%("¿ Desea Aplicar los Nuevos Valores\en Forma Automática a Todo el Personal ?\\No, Continuar\Si, Aplicar") = 2 Then
        APLIBAS% = 1
     End If
     PUCONTROL$ = String$(255, 0)
     JJ& = 0
     For U& = 1 To ULTREG&("!CATEGO")
        X$ = REGLEIDO$("!CATEGO", U&)
        COCAT% = Asc(Left$(X$, 1))
        If COCAT% > 0 Then
          If COCAT% < 256 Then
            If Mid$(PUCONTROL$, COCAT%, 1) = Chr$(0) Then
              Mid$(PUCONTROL$, COCAT%, 1) = Chr$(1)
              JJ& = JJ& + 1
              Call GRAREG("CATEGO", X$, JJ&)
              '
              If APLIBAS% = 1 Then
                 CODICAT$ = TRIM$(Left$(X$, 4))
                 BACAT = CVS(Mid$(X$, 43, 4))
                 Call APLIBASI(CODICAT$, BACAT)
              End If
              '
            End If
          End If
        End If
     Next U&
     Call SETULTREG("CATEGO", JJ&)
     Call CIERRARCH("CATEGO")
  End If
  '
End Sub

Sub APLIBASI(CODICAT$, BACAT)
  '
  Static CTXX1%, FORMUSAL$(100)
  If CTXX1% = 0 Then
    For J& = 1 To ULTREG&("NOMREM")
      X$ = REGLEIDO$("NOMREM", J&)
      FORMUSAL$(J&) = TRIM$(Mid$(X$, 43, 64))
    Next J&
  End If
  '
  If BACAT >= 0.00005 Then
     For KU& = 1 To ULTREG&("PADREMP")
        XX$ = REGLEIDO$("PADREMP", KU&)
        LEGAX% = CVI(Left$(PP0$, 2))
        STAT_LEG% = Asc(Mid$(XX$, 48, 1))     'STATUS
        If STAT_LEG% > 0 Then  ' ESTADO ACTIVO
           CATE_LEG$ = TRIM$(Mid$(XX$, 54, 4))   'CATEGORIA
           If CATE_LEG$ = CODICAT$ Then
              RBPI$ = REGLEIDO$("REMBAS", KU&)
              Call REPLA(RBPI$, MKS$(BACAT), 1, 4)
              Call GRAREG("REMBAS", RBPI$, KU&)
              '
              RB1$ = Left$(RBPI$, 48)
              PI1$ = Mid$(RBPI$, 49, 80)
              '
              For VX% = 1 To 12
                 SBAS(VX%) = CVS(Mid$(RB1$, 4 * VX% - 3))
              Next VX%
              '
              For VX% = 1 To 20
                PI(VX%) = CVS(Mid$(PI1$, 4 * VX% - 3))
              Next VX%
              '
              For J& = 2 To 12
                If MOCASAL%(J&) = 2 Then
                  VALOACT2 = SALCALCU(FORMUSAL$(J&))
                  Call REPLA(RBPI$, MKS$(VALOACT2), 4 * J& - 3, 4)
                  Call GRAREG("REMBAS", RBPI$, KU&)
                End If
              Next J&
              '
           End If
        End If
      Next KU&
      Call CIERRARCH("REMBAS")
      Call CIERRARCH("PADREMP")
   End If
   '
End Sub
'

Sub CARSALBAS()
  '
  Call TRALOCAL("NOMREM", "NOMREM")
  ATRI$ = "BORDEIZQ=W" + TRIM$(Str$(Left - 2600))
  ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + 850))
  VTB% = VENTABU%("NOMREM/" + ATRI$)
  If VTB% >= 0 Then
     JJ& = 0
     For U& = 1 To ULTREG&("!NOMREM")
        X$ = REGLEIDO$("!NOMREM", U&)
        JJ& = JJ& + 1
        Call GRAREG("NOMREM", X$, JJ&)
     Next U&
     Call SETULTREG("NOMREM", JJ&)
     Call CIERRARCH("NOMREM")
  End If
  '
End Sub

Sub CARPARAINDI()
  '
  Call TRALOCAL("PARAME", "PARAME")
  '
  ATRI$ = "BORDEIZQ=W" + TRIM$(Str$(Left - 1900))
  ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + 850))
  VTB% = VENTABU%("PARAME/" + ATRI$)
  If VTB% >= 0 Then
     JJ& = 0
     For U& = 1 To ULTREG&("!PARAME")
        X$ = REGLEIDO$("!PARAME", U&)
        JJ& = JJ& + 1
        Call GRAREG("PARAME", X$, JJ&)
     Next U&
     Call SETULTREG("PARAME", JJ&)
     Call CIERRARCH("PARAME")
  End If
  '
End Sub

Sub CARASISUB()
  '
  Call TRALOCAL("ASIGNA", "ASIGNA")
  '
  ATRI$ = "BORDEIZQ=W" + TRIM$(Str$(Left - 1400))
  ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + 850))
  VTB% = VENTABU%("ASIGNA/" + ATRI$)
  If VTB% >= 0 Then
     JJ& = 0
     For U& = 1 To ULTREG&("!ASIGNA")
        X$ = REGLEIDO$("!ASIGNA", U&)
        JJ& = JJ& + 1
        Call GRAREG("ASIGNA", X$, JJ&)
     Next U&
     Call SETULTREG("ASIGNA", JJ&)
     Call CIERRARCH("ASIGNA")
  End If
  '
End Sub

Sub HACECONCEP() 'CARGA ECO DE CONCEPTOS
    '
    Screen.MousePointer = 11
    Call COPYSTRU("ECONCEP", "ECONPRIN")
    Call SETULTREG("ECONPRIN", 0)
    Screen.MousePointer = 11
    '
    JJ& = 0: KK& = 0
    For AJUI% = 0 To 1
      For J& = 1 To ULTREG("CONCEP")
        X$ = REGLEIDO$("CONCEP", J&)
        NUCOD% = Int(J&)
        DESCRI1$ = TRIM$(Mid$(X$, 1, 30))
        EST% = Asc(Mid$(X$, 31, 1))
        If EST% <> 1 Then GoTo 10 'SI ACTIVO
        If TRIM(DESCRI1$) = "" Then GoTo 10 'SI DESCRIPCION
        Y$ = Space$(32)
        Call REPLA(Y$, MKI$(NUCOD% + 100 * AJUI%), 1, 2)
        Call REPLA(Y$, DESCRI1$, 3, 30)
        If AJUI% = 1 Then Call REPLA(Y$, "AJUSTE " + DESCRI1$, 3, 30)
        JJ& = JJ& + 1
        Call GRAREG("ECONCEP", Y$, JJ&)
        If AJUI% = 0 Then
           KK& = KK& + 1
           Call GRAREG("ECONPRIN", Y$, KK&)
        End If
10    Next J&
    Next AJUI%
    '
    For AJUI% = 0 To 1
      For J& = 1 To ULTREG("RETLEG")
        X1$ = REGLEIDO$("RETLEG", J&)
        NUCOD% = Int(J& + 200)
        DESCRI1$ = TRIM$(Mid$(X1$, 1, 30))
        EST% = Asc(Mid$(X1$, 31, 1))
        If EST% <> 1 Then GoTo 20
        If TRIM(DESCRI1$) = "" Then GoTo 20
        Y1$ = Space$(32)
        Call REPLA(Y1$, MKI$(NUCOD% + 100 * AJUI%), 1, 2)
        Call REPLA(Y1$, DESCRI1$, 3, 30)
        If AJUI% = 1 Then Call REPLA(Y1$, "AJUSTE " + DESCRI1$, 3, 30)
        JJ& = JJ& + 1
        Call GRAREG("ECONCEP", Y1$, JJ&)
        If AJUI% = 0 Then
           KK& = KK& + 1
           Call GRAREG("ECONPRIN", Y1$, KK&)
        End If
20    Next J&
    Next AJUI%
    '
    For AJUI% = 0 To 1
      For J& = 1 To ULTREG("ASIGNA")
        X2$ = REGLEIDO$("ASIGNA", J&)
        NUCOD% = Int(J& + 400)
        DESCRI1$ = TRIM$(Mid$(X2$, 1, 30))
        EST% = Asc(Mid$(X2$, 43, 1))
        If EST% <> 1 Then GoTo 30
        If TRIM(DESCRI1$) = "" Then GoTo 30
        Y2$ = Space$(32)
        Call REPLA(Y2$, MKI$(NUCOD% + 100 * AJUI%), 1, 2)
        Call REPLA(Y2$, DESCRI1$, 3, 30)
        If AJUI% = 1 Then Call REPLA(Y2$, "AJUSTE " + DESCRI1$, 3, 30)
        JJ& = JJ& + 1
        Call GRAREG("ECONCEP", Y2$, JJ&)
        If AJUI% = 0 Then
           KK& = KK& + 1
           Call GRAREG("ECONPRIN", Y2$, KK&)
        End If
30    Next J&
    Next AJUI%
    '
    Call SETULTREG("ECONCEP", JJ&)
    Call CIERRARCH("ECONCEP")
    Call SETULTREG("ECONPRIN", KK&)
    Call CIERRARCH("ECONPRIN")
    '
    YACARFO% = 0            ' PARA QUE REFRESQUE EN LA LIQUIDACION
    '
    Screen.MousePointer = 1
    '
End Sub
'

