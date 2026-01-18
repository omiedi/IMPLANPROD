VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form NUEVAGRILLA 
   Caption         =   "Grilla"
   ClientHeight    =   10650
   ClientLeft      =   60
   ClientTop       =   -330
   ClientWidth     =   19275
   LinkTopic       =   "Form1"
   ScaleHeight     =   10650
   ScaleWidth      =   19275
   Begin VB.TextBox Text10 
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7935
      Left            =   5520
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   25
      Text            =   "NUEVAGRILLA.frx":0000
      Top             =   720
      Visible         =   0   'False
      Width           =   8055
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
      Height          =   285
      Left            =   12120
      Locked          =   -1  'True
      TabIndex        =   24
      Top             =   0
      Width           =   1455
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
      Height          =   270
      Left            =   13560
      Picture         =   "NUEVAGRILLA.frx":0380
      TabIndex        =   23
      Top             =   0
      Width           =   210
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
      Height          =   270
      Left            =   17280
      Picture         =   "NUEVAGRILLA.frx":068A
      TabIndex        =   21
      Top             =   0
      Visible         =   0   'False
      Width           =   210
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
      Height          =   270
      Left            =   15120
      Picture         =   "NUEVAGRILLA.frx":0994
      TabIndex        =   20
      Top             =   0
      Visible         =   0   'False
      Width           =   210
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9360
      Top             =   5040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Fecha2 
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
      Height          =   285
      Left            =   16200
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   0
      Width           =   1095
   End
   Begin VB.TextBox Text8 
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
      Height          =   285
      Left            =   18360
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text7 
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
      Height          =   285
      Left            =   18120
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox FECHA1 
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
      Height          =   285
      Left            =   14040
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   0
      Width           =   1095
   End
   Begin VB.TextBox Text6 
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
      Height          =   285
      Left            =   12120
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   300
      Width           =   7095
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
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   0
      Width           =   9135
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
      Height          =   285
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   0
      Width           =   1095
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
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   300
      Width           =   9135
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
      Height          =   285
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   300
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   26040
      Sorted          =   -1  'True
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   25680
      TabIndex        =   3
      Top             =   240
      Width           =   1485
   End
   Begin VB.TextBox Text11 
      BackColor       =   &H00FFFFC0&
      Height          =   1095
      Left            =   25680
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1320
      Width           =   1485
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
      Left            =   27150
      TabIndex        =   1
      Top             =   270
      Width           =   175
   End
   Begin MSFlexGridLib.MSFlexGrid GRID 
      Height          =   9900
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   19095
      _ExtentX        =   33681
      _ExtentY        =   17463
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   14737632
      BackColorSel    =   -2147483643
      ForeColorSel    =   12582912
      BackColorBkg    =   16777215
      ScrollTrack     =   -1  'True
      FillStyle       =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid GRID2 
      Height          =   375
      Left            =   2280
      TabIndex        =   14
      Top             =   11160
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   661
      _Version        =   393216
      Rows            =   50
      Cols            =   3
      BackColor       =   15724527
      BackColorFixed  =   14737632
      BackColorSel    =   -2147483643
      ForeColorSel    =   12582912
      BackColorBkg    =   16777215
      FocusRect       =   2
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label6 
      Caption         =   "Mes:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11400
      TabIndex        =   22
      Top             =   0
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "Hasta:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   15480
      TabIndex        =   19
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Celda:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   45
      TabIndex        =   17
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Desde:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   13200
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Param."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11400
      TabIndex        =   11
      Top             =   300
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Rep. N°"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   50
      TabIndex        =   8
      Top             =   0
      Width           =   1215
   End
   Begin VB.Menu archivo 
      Caption         =   "Archivo"
      Begin VB.Menu nuevo 
         Caption         =   "Nuevo"
      End
      Begin VB.Menu c1 
         Caption         =   "-"
      End
      Begin VB.Menu seleccionInforme 
         Caption         =   "Seleccionar Informe"
      End
      Begin VB.Menu s 
         Caption         =   "-"
      End
      Begin VB.Menu grabar 
         Caption         =   "Grabar"
      End
      Begin VB.Menu SS 
         Caption         =   "-"
      End
      Begin VB.Menu Ejecutar 
         Caption         =   "Ejecutar Informe"
      End
      Begin VB.Menu s12 
         Caption         =   "-"
      End
      Begin VB.Menu imprimir 
         Caption         =   "Exportar a Excel"
      End
      Begin VB.Menu s6 
         Caption         =   "-"
      End
      Begin VB.Menu salir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu Edi 
      Caption         =   "Edición"
      Begin VB.Menu Borratodo 
         Caption         =   "Limpiar Grilla"
      End
      Begin VB.Menu s5 
         Caption         =   "-"
      End
      Begin VB.Menu eliminainforme 
         Caption         =   "Eliminar Informe"
      End
      Begin VB.Menu sc 
         Caption         =   "-"
      End
      Begin VB.Menu ModoEdición 
         Caption         =   "Modo Edición"
      End
      Begin VB.Menu sn 
         Caption         =   "-"
      End
      Begin VB.Menu quitarespacioslibres 
         Caption         =   "Quitar Espacios Libres"
      End
      Begin VB.Menu mnuInstructivo1 
         Caption         =   "Instructivo Generar Nueva Funcion"
      End
   End
   Begin VB.Menu Tblconfig 
      Caption         =   "Funciones de Usuarios"
      Begin VB.Menu base 
         Caption         =   "En Base a Cuentas Contables"
         Begin VB.Menu tituloinformes 
            Caption         =   "Definición"
         End
         Begin VB.Menu seleccioncuentas 
            Caption         =   "Composición"
         End
      End
   End
End
Attribute VB_Name = "NUEVAGRILLA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NUEVOINFORME% '1 NUEVO
Dim CTRL As Boolean  'CONTROLA SI ESTA APRETADA LA TECLA CONTROL
Dim GRILLA_ACTIVA%
Dim COLACT%
Dim REGACT%
Dim PASO% 'CONTROLA TECLA CONTROL
Dim ESCLICK%
Dim ULTIMAPOS$
Function CALCULABLOQUE(BLOQUEPARACALCULARORIG$)
    BLOQUEPARACALCULAR$ = BLOQUEPARACALCULARORIG$
    'ACA YA VOY A TENER LOS VALORES
    BLOQUEPARACALCULAR$ = TRIM$(REPLACAR$(BLOQUEPARACALCULAR$, " ", "")) 'QUITO LOS ESPACIOS
    'A PARTIR DE ACA CODIGO NUEVO
    'SEPARO DIVISION
'    For I& = 1 To Len(BLOQUEPARACALCULAR$)
'       '10/5*6*10.658+25.6359/-11.55
'       POS1% = InStr(LECTURAPREVIA$, "/")
'       If POS1% > 0 Then
'          LECTURAPREVIA$ = Mid$(BLOQUEPARACALCULAR$, 1, POS1% - 1)
'          IF
'
'
'    Next I&
    'FIN CODIGO NUEVO
    
    Dim ULTIMONEGATIVO$()
    '
    'PRIMERO EJECUTA LAS MULTIPLICACIONES Y LUEGO LAS DIVISIONES Y VA REEMPLAZANDO POR EJEMPLO (10*10) POR 100
40      POS1% = InStr(BLOQUEPARACALCULAR$, "/")
    If POS1% > 0 Then
       LECTURAPREVIA$ = Left$(BLOQUEPARACALCULAR$, POS1% - 1)
       'LECTURAPREVIA$ = REPLACAR$(LECTURAPREVIA$, "-", "|")
       While InStr(LECTURAPREVIA$, "/") > 0
       P1% = InStr(LECTURAPREVIA$, "/"): If P1% > 0 Then LECTURAPREVIA$ = Mid$(LECTURAPREVIA$, P1% + 1)
       Wend
       R& = 0
       While InStr(LECTURAPREVIA$, "-") > 0
       R& = R& + 1
       P1% = InStr(LECTURAPREVIA$, "-"): If P1% > 0 Then LECTURAPREVIA$ = Mid$(LECTURAPREVIA$, P1% + 1)
       ReDim ULTIMONEGATIVO$(R&)
       ULTIMONEGATIVO$(R&) = LECTURAPREVIA$
       Wend
       If P1% > 0 Then SIMBOPREVIO$ = Mid$(ULTIMONEGATIVO$(R& - 1), P1%): If SIMBOPREVIO$ = "-" Then LECTURAPREVIA$ = "-"
       While InStr(LECTURAPREVIA$, "+") > 0
       P1% = InStr(LECTURAPREVIA$, "+"): If P1% > 0 Then LECTURAPREVIA$ = Mid$(LECTURAPREVIA$, P1% + 1)
       Wend
       While InStr(LECTURAPREVIA$, "*") > 0
       P1% = InStr(LECTURAPREVIA$, "*"): If P1% > 0 Then LECTURAPREVIA$ = Mid$(LECTURAPREVIA$, P1% + 1)
       Wend
       'LECTURAPREVIA$ = REPLACAR$(LECTURAPREVIA$, "|", "-") 'el caracter 124 reemplaza al signo - de los negativos
       '
       LECTURAPOST$ = Mid$(BLOQUEPARACALCULAR$, POS1% + 1)
       P1% = InStr(LECTURAPOST$, "/"): If P1% > 1 Then LECTURAPOST$ = Mid$(LECTURAPOST$, 1, P1% - 1) 'CAMBIE DE 0 A 1 POR SI TIENE VALOR NEGATIVO
       P1% = InStr(LECTURAPOST$, "-"): If P1% > 1 Then LECTURAPOST$ = Mid$(LECTURAPOST$, 1, P1% - 1)
       P1% = InStr(LECTURAPOST$, "+"): If P1% > 1 Then LECTURAPOST$ = Mid$(LECTURAPOST$, 1, P1% - 1)
       P1% = InStr(LECTURAPOST$, "*"): If P1% > 1 Then LECTURAPOST$ = Mid$(LECTURAPOST$, 1, P1% - 1)
       'LECTURAPOST$ = REPLACAR$(LECTURAPOST$, "|", "-") 'el caracter 124 reemplaza al signo - de los negativos
A = BLOQUEPARACALCULAR$
       Resultado = 0 'POR SI EL DIVISOR ES 0
       LECTURAPOST1$ = REPLACAR$(LECTURAPOST$, "|", "-"): LECTURAPREVIA1$ = REPLACAR$(LECTURAPREVIA$, "|", "-")
       If XVALO(LECTURAPOST1$) <> 0 Then
        Resultado = XVALO(LECTURAPREVIA1$) / XVALO(LECTURAPOST1$)
        REEMPLARESULTADO$ = TRIM$(FORMATNUM$(Resultado, "F16.6")): If REEMPLARESULTADO$ = "" Then REEMPLARESULTADO$ = "0"
        BLOQUECALCULADO$ = SAFETEXT$(LECTURAPREVIA$) & "/" & SAFETEXT$(LECTURAPOST$)
        'BLOQUEPARACALCULAR$ = REPLACAR$(BLOQUEPARACALCULAR$, "|", "-")
        BLOQUEPARACALCULAR$ = TRIM$(REPLACAR$(BLOQUEPARACALCULAR$, BLOQUECALCULADO$, REEMPLARESULTADO$))
        GoTo 40
       End If
    End If

20  POS1% = InStr(BLOQUEPARACALCULAR$, "*")
    If POS1% > 0 Then
       '10/5*6*10.658+25.6359-11.55
       LECTURAPREVIA$ = Left$(BLOQUEPARACALCULAR$, POS1% - 1)
       While InStr(LECTURAPREVIA$, "/") > 0
       P1% = InStr(LECTURAPREVIA$, "/"): If P1% > 0 Then LECTURAPREVIA$ = Mid$(LECTURAPREVIA$, P1% + 1)
       Wend
       While InStr(LECTURAPREVIA$, "-") > 0
       P1% = InStr(LECTURAPREVIA$, "-"): If P1% > 0 Then LECTURAPREVIA$ = Mid$(LECTURAPREVIA$, P1% + 1)
       Wend
       While InStr(LECTURAPREVIA$, "+") > 0
       P1% = InStr(LECTURAPREVIA$, "+"): If P1% > 0 Then LECTURAPREVIA$ = Mid$(LECTURAPREVIA$, P1% + 1)
       Wend
       While InStr(LECTURAPREVIA$, "*") > 0
       P1% = InStr(LECTURAPREVIA$, "*"): If P1% > 0 Then LECTURAPREVIA$ = Mid$(LECTURAPREVIA$, P1% + 1)
       Wend
       'LECTURAPREVIA$ = REPLACAR$(LECTURAPREVIA$, "|", "-") 'el caracter 124 reemplaza al signo - de los negativos
       '
       
       LECTURAPOST$ = Mid$(BLOQUEPARACALCULAR$, POS1% + 1)
       P1% = InStr(LECTURAPOST$, "/"): If P1% > 1 Then LECTURAPOST$ = Mid$(LECTURAPOST$, 1, P1% - 1)
       P1% = InStr(LECTURAPOST$, "-"): If P1% > 1 Then LECTURAPOST$ = Mid$(LECTURAPOST$, 1, P1% - 1)
       P1% = InStr(LECTURAPOST$, "+"): If P1% > 1 Then LECTURAPOST$ = Mid$(LECTURAPOST$, 1, P1% - 1)
       P1% = InStr(LECTURAPOST$, "*"): If P1% > 1 Then LECTURAPOST$ = Mid$(LECTURAPOST$, 1, P1% - 1)
       'LECTURAPOST$ = REPLACAR$(LECTURAPOST$, "|", "-") 'el caracter 124 reemplaza al signo - de los negativos
       'ASIGNO A OTRA VARIABLE PARA NO TENER PROBLEMAS CUANDO REALIZO EL REEMPLAZO EN LA FORMULA - CON EL TEMA DEL CARACTER 124
       LECTURAPOST1$ = REPLACAR$(LECTURAPOST$, "|", "-"): LECTURAPREVIA1$ = REPLACAR$(LECTURAPREVIA$, "|", "-")
       Resultado = XVALO(LECTURAPREVIA1$) * XVALO(LECTURAPOST1$)
       REEMPLARESULTADO$ = TRIM$(FORMATNUM$(Resultado, "F16.6")): If REEMPLARESULTADO$ = "" Then REEMPLARESULTADO$ = "0"
       BLOQUECALCULADO$ = SAFETEXT$(LECTURAPREVIA$) & "*" & SAFETEXT$(LECTURAPOST$)
       BLOQUEPARACALCULAR$ = REPLACAR$(BLOQUEPARACALCULAR$, "|", "-")
       BLOQUEPARACALCULAR$ = TRIM$(REPLACAR$(BLOQUEPARACALCULAR$, BLOQUECALCULADO$, REEMPLARESULTADO$))
       GoTo 20
    End If
    '
    'AHORA SOLO QUEDAN OPERACIONES CON + O -
    ACUVALOR = 0: VALO$ = "": SIMBOLO$ = ""
    For i& = 1 To Len(BLOQUEPARACALCULAR$)
       CARACTEREX$ = Mid$(BLOQUEPARACALCULAR$, i&, 1)
       CARACTEREX_Y_SIGUIENTE$ = Mid$(BLOQUEPARACALCULAR$, i&, 2)
       If CARACTEREX_Y_SIGUIENTE$ = "--" Then
         CARACTEREX$ = "+"
       End If
       Select Case CARACTEREX$
          Case "+", "-"
            ULTIMO_CARACTER$ = CARACTEREX$
            VALO$ = REPLACAR$(VALO$, "|", "-") 'el caracter 124 reemplaza al signo - de los negativos
            VALORX = XVALO(VALO$)
            If SIMB_ANT$ = "+" Then
               ACUVALOR = ACUVALOR + VALORX
            ElseIf SIMB_ANT$ = "-" Then
               VALO$ = REPLACAR$(VALO$, "|", "-") 'el caracter 124 reemplaza al signo - de los negativos
               ACUVALOR = ACUVALOR - VALORX
            Else
               'ACA DEBERIA ENTRAR SOLO LA PRIIMERA VEZ
               VALO$ = REPLACAR$(VALO$, "|", "-")
               VALORX = XVALO(VALO$)
               ACUVALOR = ACUVALOR + VALORX
            End If
            VALO$ = ""
            SIMB_ANT$ = CARACTEREX$
            If ULTIMO_CARACTER$ = "|" Then SIMB_ANT$ = "-"
          Case Else
            VALO$ = VALO$ & CARACTEREX$
       End Select
       If CARACTEREX_Y_SIGUIENTE$ = "--" Then: i& = i& + 1  ' 'CASO QUE RESTA UN VALOR NEGATIVO

65  Next i&
    If ULTIMO_CARACTER$ = "-" Then
      VALO$ = REPLACAR$(VALO$, "|", "-")
      ACUVALOR = ACUVALOR - XVALO(VALO$)
    ElseIf ULTIMO_CARACTER$ = "+" Then
      VALO$ = REPLACAR$(VALO$, "|", "-")
      ACUVALOR = ACUVALOR + XVALO(VALO$)
    Else
      BLOQUEPARACALCULAR$ = REPLACAR$(BLOQUEPARACALCULAR$, "|", "-") 'el caracter 124 reemplaza al signo - de los negativos
      ACUVALOR = ACUVALOR + XVALO(BLOQUEPARACALCULAR$)
    End If
    CALCULABLOQUE = XVALO(ACUVALOR)
End Function

Sub CARGA_ANCHO_COLUM()
   NUEMEROINFORME% = XVALO(Text4)
   If NUEMEROINFORME% < 1 Then Exit Sub
   '
   'HAGO ESTA VALIDACION POR SI EN OTRA PC SE ESTA POR GRABAR EL MISMO NUMERO
   CONDI$ = "CODIGO = " & NUEMEROINFORME%
   '
   Call ACTUREGISTRO("ANCHCOLU", "MARBORRA", CONDI$, 1, REGAF&)
   'SI EXISTE MARCA DE BORRADO Y ABAJO BORRA, EN CASO QUE NO EXISTSA VA AGREGAR Y NO VA A BORRAR NADA.
   SQLX$ = "SELECT * FROM ANCHCOLU WHERE " & CONDI$
   Dim RS As ADODB.Recordset
   Set RS = New ADODB.Recordset
25 On Error Resume Next
   RS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
   End If
   On Error GoTo 0
   '
   With RS
           For COL& = 1 To GRID.Cols - 1
               .AddNew   ' YA CONVERTIDO
               On Error Resume Next
               !CodiEmpr = CodiEmp%
               On Error GoTo 0
               !Codigo = NUEMEROINFORME%
               !NUMCOLUMNA = COL&
               !ANCHO = GRID.ColWidth(COL&)
               !MARBORRA = 0
               .Update
           Next COL&
    End With
    RS.Close
    Set RS = Nothing
    '
    Call BORRAREGISTROS("ANCHCOLU", CONDI$ & " And MARBORRA = 1", REGAF&)

End Sub

Function COLUMNA&(POS_CELDA_LEIDA$, FILAX&)
   'LA VARIABLE FILAX& LA DEVUELVE COMO PARAMETRO DE SALIDA
   For j% = 1 To Len(POS_CELDA_LEIDA$)  ' RECORRO A6
         LETRAX$ = Mid$(POS_CELDA_LEIDA$, j%, 1) 'A
         If Not IsNumeric(LETRAX$) Then
             If j% = 0 Then
               COLU1& = COLU1& + (Asc(LETRAX$) - 64)
               MULTIP& = COLU1&
             Else
               'EN ESTA VUELTA SE CONSIDERA LAS LETRAS DOBLES EJEMPLO AB
               On Error Resume Next
               COLU1& = COLU1& + (Asc(LETRAX$) - 64) + (25 * MULTIP&) ' LE SUMO 25 POR QUE ES EL TOTAL DE LETRAS - 1
               MULTIP& = COLU1&
               On Error GoTo 0
             End If
             FINLEN% = j%
          End If
      Next j%
      COLUMNA& = COLU1&
      FILAX& = XVALO(Mid$(POS_CELDA_LEIDA$, FINLEN% + 1))
      
End Function

Private Sub Check1_Click()
        COLUM& = COLUMNA&(Text2, FILAX&) 'FILAX& LO DEVUELVE CON VALOR LA FUNCION CON EL NUMERO DE REGISTRO
        REGIS& = FILAX&
        GRID.Row = FILAX&
        GRID.COL = COLUM&
        If Check1.Value = 1 Then
          GRID.CellFontStrikeThrough = True
        Else
          GRID.CellFontStrikeThrough = False
        End If
End Sub

Sub ControlFactura()

'   Call DESHASFECHA(DES%, HAS%, PERIO$)
'   If PERIO$ = "" Then
'      GoTo 99
'   End If
'   '
'   DESA = FETEXCOR1$(DES%)
'   HASA = FETEXCOR1$(HAS%)
'   '
'   '
'   'MATRIZ DONDE SE ALMACENAN LOS COMPROBANTES A RECUPERAR
'   '
'   Dim MATRIZ$()
'   Dim MATRIZ2$()
'   PVTA1% = XVALO(Control("FACTURA", "PUVENELE"))
'   PVTA2% = XVALO(Control("FACTURA", "PUVENCAP"))
'   PVTA3% = XVALO(Control("FACTURA", "PUVEFAEX"))
'   '
'   If PVTA1% + PVTA2% + PVTA3% < 1 Then Exit Sub
'   Screen.MousePointer = 11
'   '
'   'BORRA EL CONTENIDO DE LA TABLA DE FACTURAS NO RECUPERADAS SEGUN PUNTO DE VENTA ELEGIDO
'   '
'   PVTAX$ = "0000" & PVTA%: PVTAX$ = Right(PVTAX$, 4)
'   CONDI1$ = "SUBSTRING(NFACTURA, 4, 4) = '" & PVTAX$ & "'"
'   '
'   'CONDICION DE CONSULTA
'   '
'   CONDI$ = " tipomovim = 'fven' "
'   CONDI$ = CONDI$ + " AND NUMECOMP > 0 "
'   If PVTA1% > 0 Then
'      CONDI$ = CONDI$ + " AND ( NUMETALO =  " & PVTA1%
'   End If
'   If PVTA2% > 0 Then
'      CONDI$ = CONDI$ + " or  NUMETALO =  " & PVTA2%
'   End If
'
'   If PVTA3% > 0 Then
'      CONDI$ = CONDI$ + " or  NUMETALO =  " & PVTA3%
'   End If
'   CONDI$ = CONDI$ + ")"
'   CONDI$ = CONDI$ + " AND FECHEMISI >= '" & DESA & "'"
'   CONDI$ = CONDI$ + " AND FECHEMISI <= '" & HASA & "'"
'
'   '
'   FIN& = CantRegistros("CAJABANC", CONDI$)
'   '
'   'RECORRO CAJABANC VERIFICANDO CORRELATIVIDAD Y GENERO LA MATRIZ CON LOS
'   'DOCUMENTOS A RECUPERAR
'   '
'   SQLX$ = "SELECT * FROM CAJABANC WITH(NOLOCK)"
'   SQLX$ = SQLX$ + " WHERE " + CONDI$
'   SQLX$ = SQLX$ + " ORDER BY OPCIMOVIM, VARIMOVIM, NUMECOMP asc"
'   '
'   I& = 0: j& = 0: VUELTA% = 0: VARI$ = "": OPCI$ = "": U& = 0: PVTA% = 0
'   Dim CABATMP As New ADODB.Recordset
'   Set CABATMP = FLEXCONN.Execute(SQLX$)
'   With CABATMP
'       Do While Not .EOF
'          '
'          U& = U& + 1
'          Call TRACE(20, U&, FIN&)
'          'COMPARA LOS VALORES PARA VER SI CAMBIO EJEMPLO DE 'FC' A 'NC' O LA LETRA
'          If SAFETEXT$(!VARIMOVIM) <> VARI$ Or SAFETEXT$(!OpciMovim) <> OPCI$ Or XVALO(!NUMETALO) <> PVTA% Then
'              I& = XVALO(!NUMECOMP)
'              VARI$ = SAFETEXT$(!VARIMOVIM)
'              OPCI$ = SAFETEXT$(!OpciMovim)
'              PVTA% = SAFETEXT$(!NUMETALO)
'           End If
'          '
'10        If I& < XVALO(!NUMECOMP) Then 'SI SE ROMPE LA CARRELATIVIDAD
''           'CARGA EL VECTOR
'            j& = j& + 1
'            ReDim Preserve MATRIZ$(j&)
'            ReDim Preserve MATRIZ2$(j&)
'            'ARMO EL STRING DE LA FACTURA DEL HUECO
'            NUMCOM$ = "00000000" & (I&): NUMCOM$ = Right(NUMCOM$, 8)
'            'ASIGNO A LA MATRIZ$ LA LECTURA DEL CAE ANTERIOR AL HUECO Y FACTURA DEL HUECO
'            MATRIZ$(j&) = OPCI$ + "-" + PVTAX$ + "-" + NUMCOM$ + "-" + VARI$
'            MATRIZ2$(j&) = FECHATEX$(CVINT(!FECEHEMISI))
'            I& = I& + 1
'            '
'            If I& < XVALO(!NUMECOMP) Then ' SI HAY MAS DE UN HUECO CONSECUTIVO
'              GoTo 10                     ' VUELVE Y COMPARA A VER SI EL SIGUIENTE
'            End If                        ' ES EL QUE SIGUE
'            '
'         End If
'         '
'         I& = I& + 1
'         NUFACTU$ = SAFETEXT$(!CODICOM_LAR)
'         .MoveNext
'        Loop
'    End With
'    '
'    INDIC& = j&
'    If INDIC& > 0 Then ' SI HAY HUECOS EN LAS FACTURAS HABILITO EL TIMER PARA INICIAR PROCESO
'      Timer1.Enabled = True
'    Else
'      Screen.MousePointer = 1
'      Call COMUNI("No Se Encontraron Facturas No Correlativas Para el Punto de Venta: " & PVTA%)
'    End If
'
'50  If UBound(MATRIZ) > 0 Then
'        Call CARGA_ENCABEZADO(MUESTRADATOS.MSF_2, "Fecha Aproximada/A24;Factura/A48", MUESTRADATOS, 1, 0, 0, 10)
'        Call CARGA_ENCABEZADO(MUESTRADATOS.GRID, "Fecha Aproximada/A24;Factura/A48", MUESTRADATOS, 1, 0, 0, 10)
'        For L& = 1 To UBound(MATRIZ)
'           GRID.Rows = L& + 1
'           GRID.TextMatrix(L&, 1) = MATRIZ2$(L&)
'           GRID.TextMatrix(L&, 2) = MATRIZ$(L&)
'        Next L&
'        MUESTRADATOS.Show 1
'        Exit Sub
'   Else
'        Call COMUNI("No Se Encontraron Facturas No Correlativas ")
'
'   End If








End Sub

Function COSTO_CALCULADO(CI1%)
    Des% = FECHANUM(FECHA1)
    FF1$ = FECHATEXLAR$(Des%)
    MESHOY& = Val(Mid$(FF1$, 4, 2))
    ANIOHOY& = Val(Mid$(FF1$, 7, 4))
    IDMESANIO& = (MESHOY& * 10000) + ANIOHOY&

    SQLX$ = "SELECT PrecioVenta "
    SQLX$ = SQLX$ + " FROM HISTOCOSTOS WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE IdPeriodo = " & IDMESANIO&
    SQLX$ = SQLX$ + " AND COD_INTE = " & CI1%
    SQLX$ = SQLX$ + " ORDER BY FECOSCAL DESC "
    Set RS = FLEXCONN.Execute(SQLX$)
    With RS
      COSTOCALADOX = XVALO(!PrecioVenta)
    End With
    RS.Close
    Set RS = Nothing
    
    COSTO_CALCULADO = COSTOCALADOX
    If COSTO_CALCULADO < 0.0005 Then COSTO_CALCULADO = COSCALC#(CI1%)

End Function

Function COSTOSTANDAR_TOTAL(CI1%)
    Des% = FECHANUM(FECHA1)
    FF1$ = FECHATEXLAR$(Des%)
    MESHOY& = Val(Mid$(FF1$, 4, 2))
    ANIOHOY& = Val(Mid$(FF1$, 7, 4))
    IDMESANIO& = (MESHOY& * 10000) + ANIOHOY&

    TOTAL_COSTO_STD = 0
    SQLX$ = "SELECT CostoStdTotal "
    SQLX$ = SQLX$ + " FROM HISTOCOSTOS WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE IdPeriodo = " & IDMESANIO&
    SQLX$ = SQLX$ + " AND COD_INTE = " & CI1%
    SQLX$ = SQLX$ + " ORDER BY FECOSCAL DESC "
    Set RS = FLEXCONN.Execute(SQLX$)
    With RS
      TOTAL_COSTO_STD = XVALO(!CostoStdTotal)
      If TOTAL_COSTO_STD < 1 Then
         Debug.Print CI1%
      End If
    End With
    RS.Close
    Set RS = Nothing
    '
    COSTOSTANDAR_TOTAL = TOTAL_COSTO_STD
    '
End Function

Function COSTO_M_OBRA_REMITIDOS()
'   COSTO_M_OBRA_REMITIDOS
    'DEVUELVE EL COSTO DE MANO DE OBRA DE LOS REMITOS FACTURADOS SEGUN EL MOVISTO SEGUN RANGO DE FECHAS
    If TRIM$(FECHA1) = "" Then
       Call COMUNI("Falta Ingresar Fecha Desde")
       Exit Function
    End If
    '
    If TRIM$(Fecha2) = "" Then
       Call COMUNI("Falta Ingresar Fecha Hasta")
       Exit Function
    End If
    '
    CMO$ = TRIM$(CONTROL$("", "CODREFAC"))
    If Len(CMO$) <> 2 Then
      CMO$ = "RT"
    End If
    '
    Desde = FETEXCOR1$(FECHANUM(FECHA1))
    Hasta = FETEXCOR1$(FECHANUM(Fecha2))
    '
    SQLX$ = " SELECT SUM(CANTINGRE) AS CANTING, SUM(CANTSALID) AS CANSALI, COD_INTE FROM MOVISTO WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE (CodiMovi = '" & CMO$ & "' "
    SQLX$ = SQLX$ + " OR CodiMovi =  'DV') "
    SQLX$ = SQLX$ + " AND DoPriLar <> '' "
    SQLX$ = SQLX$ + " AND (DoSecLar IS NOT NULL) AND (DoSecLar <> '') AND LEFT(DoSecLar,3) <> 'P/F'"
    SQLX$ = SQLX$ + " AND FechMov >= '" & Desde & "' "
    SQLX$ = SQLX$ + " AND FechMov <= '" & Hasta & "' "
    SQLX$ = SQLX$ + " AND LEFT(REFERCOR,4) <> 'ANUL' "
    SQLX$ = SQLX$ + " AND COD_INTE > 0 "
    SQLX$ = SQLX$ + " AND Nume_Clie > 0 "
    SQLX$ = SQLX$ + " GROUP BY COD_INTE"
    Set RS = FLEXCONN.Execute(SQLX$)
    With RS
        Do While Not .EOF
           CI1% = XVALO(!COD_INTE)
           CANTI_INGRE = XVALO(!CANTING)
           CANTI_SALI = XVALO(!CANSALI)
           CANTI = CANTI_SALI - CANTI_INGRE
           COSTOMOBRA = COSTOMOBRA + (CANTI * COSTO_DE_ELABORACION(CI1%))
           .MoveNext
        Loop
    End With
    RS.Close
    Set RS = Nothing
    COSTO_M_OBRA_REMITIDOS = COSTOMOBRA
End Function
Function COSTO_M_OBRA_REMITOS_TODOS()
'   COSTO_M_OBRA_REMITIDOS
    'DEVUELVE EL COSTO DE MANO DE OBRA DE LOS REMITOS SEGUN EL MOVISTO SEGUN RANGO DE FECHAS
    If TRIM$(FECHA1) = "" Then
       Call COMUNI("Falta Ingresar Fecha Desde")
       Exit Function
    End If
    '
    If TRIM$(Fecha2) = "" Then
       Call COMUNI("Falta Ingresar Fecha Hasta")
       Exit Function
    End If
    '
    CMO$ = TRIM$(CONTROL$("", "CODREFAC"))
    If Len(CMO$) <> 2 Then
      CMO$ = "RT"
    End If
    '
    Desde = FETEXCOR1$(FECHANUM(FECHA1))
    Hasta = FETEXCOR1$(FECHANUM(Fecha2))
    '
    
    SQLX$ = "SELECT SUM(CANTINGRE) AS CANTING, SUM(CANTSALID) AS CANSALI, COD_INTE FROM MOVISTO WITH(NOLOCK)"
    SQLX$ = SQLX$ & " WHERE FECHMOV >= '" & Desde & "'"
    SQLX$ = SQLX$ & " AND FECHMOV <= '" & Hasta & "'"
    SQLX$ = SQLX$ + " AND (CodiMovi = '" & CMO$ & "' OR CODIMOVI = 'DV') "
    SQLX$ = SQLX$ + " AND NUME_CLIE > 0 AND LEFT(REFERCOR,4) <> 'ANUL'"
    SQLX$ = SQLX$ + "  AND DOPRILAR <> '' AND (COD_INTE  > 0 OR DESCRI_LAR LIKE 'DESC.:%')"
    SQLX$ = SQLX$ + " GROUP BY COD_INTE"

    'NUME_CILE > 0 AND LEFT(REFERCOR,4) <> 'ANUL' AND (CODIMOVI = 'RT'  OR CODIMOVI = 'DV') AND DOPRILAR <> '' AND (COD_INTE  > 0 OR DESCRILAR LIKE 'DESC.:%')
    Set RS = READSET(SQLX$)
    
    With RS
      Do While Not .EOF
                CI1% = XVALO(!COD_INTE)
                CANREMSAL = XVALO(!CANSALI)
                CANREMIGRE = XVALO(!CANTING)
                CANREM = CANREMSAL - CANREMIGRE
                COSTOMOBRA = COSTOMOBRA + (CANREM * COSTO_DE_ELABORACION(CI1%))
               .MoveNext
      Loop
    End With
    On Error Resume Next
    RS.Close
    Set RS = Nothing
    On Error GoTo 0
    COSTO_M_OBRA_REMITOS_TODOS = COSTOMOBRA

End Function

Function VALOR_VENTA_REMITIDO_MES_ACTIVO()
'
    'DEVUELVE EL VALOR DE VENTA REGISTRADO EN EL MOVISTO DEL CAMPO VALOUNIT QUE ES SIN IVA Y NO SE ESTA CONSIDERANDO SI LA FACTURA POSEE ALGUN DESCUENTO.
    If TRIM$(FECHA1) = "" Then
       Call COMUNI("Falta Ingresar Fecha Desde")
       Exit Function
    End If
    '
    If TRIM$(Fecha2) = "" Then
       Call COMUNI("Falta Ingresar Fecha Hasta")
       Exit Function
    End If
    '
    CMO$ = TRIM$(CONTROL$("", "CODREFAC"))
    If Len(CMO$) <> 2 Then
      CMO$ = "RT"
    End If
    '
    Desde = FETEXCOR1$(FECHANUM(FECHA1))
    Hasta = FETEXCOR1$(FECHANUM(Fecha2))
    '
    VALOVENTA = 0
    SQLX$ = " SELECT SUM(CANTINGRE) AS CANTING, SUM(CANTSALID) AS CANSALI, COD_INTE, VALOUNIT FROM MOVISTO WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE (CodiMovi = '" & CMO$ & "' "
    SQLX$ = SQLX$ + " OR CodiMovi =  'DV') "
    SQLX$ = SQLX$ + " AND DoPriLar <> '' "
    SQLX$ = SQLX$ + " AND (DoSecLar IS NOT NULL) AND (DoSecLar <> '') AND LEFT(DoSecLar,3) <> 'P/F'"
    SQLX$ = SQLX$ + " AND FechMov >= '" & Desde & "' "
    SQLX$ = SQLX$ + " AND FechMov <= '" & Hasta & "' "
    SQLX$ = SQLX$ + " AND LEFT(REFERCOR,4) <> 'ANUL' "
    SQLX$ = SQLX$ + " AND COD_INTE > 0 "
    SQLX$ = SQLX$ + " AND Nume_Clie > 0 "
    SQLX$ = SQLX$ + " GROUP BY COD_INTE,VALOUNIT"
    Set RS = FLEXCONN.Execute(SQLX$)
    With RS
        Do While Not .EOF
           CANTI_INGRE = XVALO(!CANTING)
           CANTI_SALI = XVALO(!CANSALI)
           CANTI = CANTI_SALI - CANTI_INGRE
           
           VALOVENTA = VALOVENTA + (CANTI * XVALO(!VALOUNIT))
           .MoveNext
        Loop
    End With
    RS.Close
    Set RS = Nothing
    '
    VALOR_VENTA_REMITIDO_MES_ACTIVO = VALOVENTA
    '
End Function

Function COSTO_STANDARD_TOTAL()
'   COSTO_M_OBRA_REMITIDOS
    'DEVUELVE EL COSTO DE MANO DE OBRA DE LOS REMITOS FACTURADOS SEGUN EL MOVISTO SEGUN RANGO DE FECHAS
    If TRIM$(FECHA1) = "" Then
       Call COMUNI("Falta Ingresar Fecha Desde")
       Exit Function
    End If
    '
    If TRIM$(Fecha2) = "" Then
       Call COMUNI("Falta Ingresar Fecha Hasta")
       Exit Function
    End If
    '
    CMO$ = TRIM$(CONTROL$("", "CODREFAC"))
    If Len(CMO$) <> 2 Then
      CMO$ = "RT"
    End If
    '
    Desde = FETEXCOR1$(FECHANUM(FECHA1))
    Hasta = FETEXCOR1$(FECHANUM(Fecha2))
    '
    SQLX$ = " SELECT SUM(CANTINGRE) AS CANTING, SUM(CANTSALID) AS CANSALI, COD_INTE FROM MOVISTO WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE (CodiMovi = '" & CMO$ & "' "
    SQLX$ = SQLX$ + " OR CodiMovi =  'DV') "
    SQLX$ = SQLX$ + " AND DoPriLar <> '' "
    SQLX$ = SQLX$ + "  AND DoSecLar <> '' AND LEFT(DoSecLar,3) <> 'P/F'"
    SQLX$ = SQLX$ + " AND FechMov >= '" & Desde & "' "
    SQLX$ = SQLX$ + " AND FechMov <= '" & Hasta & "' "
    SQLX$ = SQLX$ + " AND LEFT(REFERCOR,4) <> 'ANUL' "
    SQLX$ = SQLX$ + " AND COD_INTE > 0 "
    SQLX$ = SQLX$ + " AND Nume_Clie > 0 "
    SQLX$ = SQLX$ + " GROUP BY COD_INTE"
    Set RS = FLEXCONN.Execute(SQLX$)
    With RS
        Do While Not .EOF
           CI1% = XVALO(!COD_INTE)
           CANTI_INGRE = XVALO(!CANTING)
           CANTI_SALI = XVALO(!CANSALI)
           CANTI = CANTI_SALI - CANTI_INGRE
           COSTD = COSTD + (CANTI * COSTOSTANDAR_TOTAL(CI1%))
           .MoveNext
        Loop
    End With
    RS.Close
    Set RS = Nothing
    COSTO_STANDARD_TOTAL = COSTD
End Function


Function COSTO_CALCULADO_REMITIDOS()
'   COSTO_CALCULADO _REMITIDOS
    'DEVUELVE EL COSTO DE MANO DE OBRA DE LOS REMITOS FACTURADOS SEGUN EL MOVISTO SEGUN RANGO DE FECHAS
    If TRIM$(FECHA1) = "" Then
       Call COMUNI("Falta Ingresar Fecha Desde")
       Exit Function
    End If
    '
    If TRIM$(Fecha2) = "" Then
       Call COMUNI("Falta Ingresar Fecha Hasta")
       Exit Function
    End If
    '
    CMO$ = TRIM$(CONTROL$("", "CODREFAC"))
    If Len(CMO$) <> 2 Then
      CMO$ = "RT"
    End If
    '
    Desde = FETEXCOR1$(FECHANUM(FECHA1))
    Hasta = FETEXCOR1$(FECHANUM(Fecha2))
    '
    SQLX$ = " SELECT SUM(CANTINGRE) AS CANTING, SUM(CANTSALID) AS CANSALI, COD_INTE FROM MOVISTO WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE (CodiMovi = '" & CMO$ & "' "
    SQLX$ = SQLX$ + " OR CodiMovi =  'DV') "
    SQLX$ = SQLX$ + " AND DoPriLar <> '' "
    SQLX$ = SQLX$ + " AND (DoSecLar IS NOT NULL) AND (DoSecLar <> '') AND LEFT(DoSecLar,3) <> 'P/F'"
    
    SQLX$ = SQLX$ + " AND FechMov >= '" & Desde & "' "
    SQLX$ = SQLX$ + " AND FechMov <= '" & Hasta & "' "
    SQLX$ = SQLX$ + " AND LEFT(REFERCOR,4) <> 'ANUL' "
    SQLX$ = SQLX$ + " AND COD_INTE > 0 "
    SQLX$ = SQLX$ + " AND Nume_Clie > 0 "
    SQLX$ = SQLX$ + " GROUP BY COD_INTE"
    Set RS = FLEXCONN.Execute(SQLX$)
    With RS
        Do While Not .EOF
           CI1% = XVALO(!COD_INTE)
           CANTI_INGRE = XVALO(!CANTING)
           CANTI_SALI = XVALO(!CANSALI)
           CANTI = CANTI_SALI - CANTI_INGRE
           COSTOCALCU = COSTOCALCU + (COSTO_CALCULADO(CI1%) * CANTI)
           .MoveNext
        Loop
    End With
    RS.Close
    Set RS = Nothing
    COSTO_CALCULADO_REMITIDOS = COSTOCALCU
    
End Function
Function COSTO_SUMINISTROS_REMITIDOS()
    'DEVUELVE EL COSTO DE SUMINISTROS DE LOS REMITOS FACTURADOS SEGUN EL MOVISTO SEGUN RANGO DE FECHAS
    '
    CMO$ = TRIM$(CONTROL$("", "CODREFAC"))
    If Len(CMO$) <> 2 Then
      CMO$ = "RT"
    End If
    '
    Desde = FETEXCOR1$(FECHANUM(FECHA1))
    Hasta = FETEXCOR1$(FECHANUM(Fecha2))
    '
    SQLX$ = " SELECT SUM(CANTINGRE) AS CANTING, SUM(CANTSALID) AS CANSALI, COD_INTE FROM MOVISTO WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE (CodiMovi = '" & CMO$ & "' "
    SQLX$ = SQLX$ + " OR CodiMovi =  'DV') "
    SQLX$ = SQLX$ + " AND DoPriLar <> '' "
    SQLX$ = SQLX$ + " AND (DoSecLar IS NOT NULL) AND (DoSecLar <> '') AND LEFT(DoSecLar,3) <> 'P/F'"
    SQLX$ = SQLX$ + " AND FechMov >= '" & Desde & "' "
    SQLX$ = SQLX$ + " AND FechMov <= '" & Hasta & "' "
    SQLX$ = SQLX$ + " AND LEFT(REFERCOR,4) <> 'ANUL' "
    SQLX$ = SQLX$ + " AND COD_INTE > 0 "
    SQLX$ = SQLX$ + " AND Nume_Clie > 0 "
    SQLX$ = SQLX$ + " GROUP BY COD_INTE"
    Set RS = FLEXCONN.Execute(SQLX$)
    With RS
        Do While Not .EOF
           CI1% = XVALO(!COD_INTE)
           CANTI_INGRE = XVALO(!CANTING)
           CANTI_SALI = XVALO(!CANSALI)
           CANTI = CANTI_SALI - CANTI_INGRE
           COSTOSUMIN = COSTOSUMIN + (COSTO_DE_SUMINISTROS(CI1%) * CANTI)
           .MoveNext
        Loop
    End With
    RS.Close
    Set RS = Nothing
    COSTO_SUMINISTROS_REMITIDOS = COSTOSUMIN
        CODI$ = CODEXT$(2039)

End Function


Function FACTU_ANTIC_DEL_MES_ACTIVO()
    FACTU_ANTIC_DEL_MES_ACTIVO = 0
    Desde = FETEXCOR1$(FECHANUM(FECHA1))
    Hasta = FETEXCOR1$(FECHANUM(Fecha2))

    SQLX$ = " Select SUM(IMPODEBE) AS IDEBE from factuanti with(nolock)"
    SQLX$ = SQLX$ + " where nuclien > 0 "
    SQLX$ = SQLX$ + " and tipomovim = 'VENTA'"
    SQLX$ = SQLX$ + " AND FECHEMISI >= '" & Desde & "'"
    SQLX$ = SQLX$ + " AND FECHEMISI <= '" & Hasta & "'"
    SQLX$ = SQLX$ + " GROUP BY tipomovim"
    '
    Set RS = READSET(SQLX$)
    j& = 0
    With RS
      If Not .EOF Then
           FACTU_ANTIC_DEL_MES_ACTIVO = XVALO(!IDEBE)
      End If
    End With
    RS.Close
    Set RS = Nothing

End Function
Function FACTU_ANTIC_ACUM_ANTERIOR()
    FACTU_ANTIC_ACUM_ANTERIOR = 0
    Desde = FETEXCOR1$(FECHANUM(FECHA1))

    SQLX$ = " Select SUM(IMPODEBE) AS IDEBE from factuanti with(nolock)"
    SQLX$ = SQLX$ + " where nuclien > 0 "
    SQLX$ = SQLX$ + " and tipomovim = 'VENTA'"
    SQLX$ = SQLX$ + " AND FECHEMISI < '" & Desde & "'"
    SQLX$ = SQLX$ + " GROUP BY tipomovim"
    '
    Set RS = READSET(SQLX$)
    j& = 0
    With RS
      If Not .EOF Then
           FACTU_ANTIC_ACUM_ANTERIOR = XVALO(!IDEBE)
      End If
    End With
    RS.Close
    Set RS = Nothing

End Function

Function FACTU_ANTIC_ACUM_A_FIN_MES_ACTIVO()
    FACTU_ANTIC_ACUM_A_FIN_MES_ACTIVO = 0
    Hasta = FETEXCOR1$(FECHANUM(Fecha2))

    SQLX$ = " Select SUM(IMPODEBE) AS IDEBE from factuanti with(nolock)"
    SQLX$ = SQLX$ + " where nuclien > 0 "
    SQLX$ = SQLX$ + " and tipomovim = 'VENTA'"
    SQLX$ = SQLX$ + " AND FECHEMISI <= '" & Hasta & "'"
    SQLX$ = SQLX$ + " GROUP BY tipomovim"
    '
    Set RS = READSET(SQLX$)
    j& = 0
    With RS
      If Not .EOF Then
           FACTU_ANTIC_ACUM_A_FIN_MES_ACTIVO = XVALO(!IDEBE)
      End If
    End With
    RS.Close
    Set RS = Nothing

End Function

Function FACT_ANT_PEND_APLICACION_MES_ACTIVO()
   '=============================================================================================================================
   'ESTO CONSIDERA LAS QUE SE EMITIERON EN EL MES ACTIVO - PUEDE SER QUE HAYA PENDIENTES DE OTROS MESES QUE NO ESTEN APLICADAS.
   '=============================================================================================================================
   
    FACT_ANT_PEND_APLICACION_MES_ACTIVO = 0
    Desde = FETEXCOR1$(FECHANUM(FECHA1))
    Hasta = FETEXCOR1$(FECHANUM(Fecha2))

    SQLX$ = " Select SUM(IMPODEBE) AS IDEBE, SUM(IMPOHABE) AS IHABE from factuanti with(nolock)"
    SQLX$ = SQLX$ + " where nuclien > 0 "
    SQLX$ = SQLX$ + " and (tipomovim = 'VENTA'"
    SQLX$ = SQLX$ + " OR tipomovim = 'AJUS')"
    SQLX$ = SQLX$ + " AND FECHEMISI >= '" & Desde & "'"
    SQLX$ = SQLX$ + " AND FECHEMISI <= '" & Hasta & "'"
    SQLX$ = SQLX$ + " GROUP BY TIPOMOVIM"
    Set RS = READSET(SQLX$)
    j& = 0
    With RS
        Do While Not .EOF
           AntiSaldo# = XVALO(!IDEBE) - XVALO(!IHABE)
          .MoveNext
        Loop
    End With
    RS.Close
    Set RS = Nothing
    FACT_ANT_PEND_APLICACION_MES_ACTIVO = AntiSaldo#
End Function
Function FACT_ANT_PEND_APLICACION_ACUM_ANTERIOR()

    FACT_ANT_PEND_APLICACION_ACUM_ANTERIOR = 0
    Desde = FETEXCOR1$(FECHANUM(FECHA1))
    Hasta = FETEXCOR1$(FECHANUM(Fecha2))

    SQLX$ = " Select SUM(IMPODEBE) AS IDEBE, SUM(IMPOHABE) AS IHABE from factuanti with(nolock)"
    SQLX$ = SQLX$ + " where nuclien > 0 "
    SQLX$ = SQLX$ + " and (tipomovim = 'VENTA'"
    SQLX$ = SQLX$ + " OR tipomovim = 'AJUS')"
    SQLX$ = SQLX$ + " AND FECHEMISI < '" & Desde & "'"
    SQLX$ = SQLX$ + " GROUP BY TIPOMOVIM"
    Set RS = READSET(SQLX$)
    j& = 0
    With RS
        Do While Not .EOF
           AntiSaldo# = XVALO(!IDEBE) - XVALO(!IHABE)
          .MoveNext
        Loop
    End With
    RS.Close
    Set RS = Nothing
    FACT_ANT_PEND_APLICACION_ACUM_ANTERIOR = AntiSaldo#
End Function
Function FACT_ANT_PEND_APLICACION_ACUM_A_FIN_MES_ACTIVO()

    FACT_ANT_PEND_APLICACION_ACUM_A_FIN_MES_ACTIVO = 0
    Hasta = FETEXCOR1$(FECHANUM(Fecha2))

    SQLX$ = " Select SUM(IMPODEBE) AS IDEBE, SUM(IMPOHABE) AS IHABE from factuanti with(nolock)"
    SQLX$ = SQLX$ + " where nuclien > 0 "
    SQLX$ = SQLX$ + " and (tipomovim = 'VENTA'"
    SQLX$ = SQLX$ + " OR tipomovim = 'AJUS')"
    SQLX$ = SQLX$ + " AND FECHEMISI <= '" & Hasta & "'"
    SQLX$ = SQLX$ + " GROUP BY TIPOMOVIM"
    Set RS = READSET(SQLX$)
    j& = 0
    With RS
        Do While Not .EOF
           AntiSaldo# = XVALO(!IDEBE) - XVALO(!IHABE)
          .MoveNext
        Loop
    End With
    RS.Close
    Set RS = Nothing
    FACT_ANT_PEND_APLICACION_ACUM_A_FIN_MES_ACTIVO = AntiSaldo#
End Function

Sub FORMATCOLUM(GRID As MSFlexGrid, Optional NUMEINFO%)
   NUEMEROINFORME% = NUMEINFO%
   If NUEMEROINFORME% < 1 Then NUEMEROINFORME% = XVALO(Text4)
   If NUEMEROINFORME% < 1 Then Exit Sub
   '
   'HAGO ESTA VALIDACION POR SI EN OTRA PC SE ESTA POR GRABAR EL MISMO NUMERO
   CONDI$ = "CODIGO = " & NUEMEROINFORME%
   SQLX$ = "SELECT * FROM ANCHCOLU WITH(NOLOCK) WHERE " & CONDI$
   Set RS = READSET(SQLX$)    '
   With RS
     Do While Not .EOF
      COL& = XVALO(!NUMCOLUMNA)
      ANCHOX = XVALO(!ANCHO)
      If ANCHOX < 1000 Then ANCHOX = 1000
      GRID.ColWidth(COL&) = ANCHOX
      .MoveNext
     Loop
    End With
    RS.Close
    Set RS = Nothing
  
End Sub

Function FUNCIONES_ESPECIFICAS(TXT1$)
   NAMEFUNCION$ = Mid$(TRIM$(UCase$(TXT1$)), 3)
   Select Case NAMEFUNCION$
      Case "COSTO REMITIDO Y FACTURADO DE SUMINISTROS"
           FUNCIONES_ESPECIFICAS = COSTO_SUMINISTROS_REMITIDOS
           
      Case "COSTO REMITIDO Y FACTURADO DE M.OBRA"
           FUNCIONES_ESPECIFICAS = COSTO_M_OBRA_REMITIDOS
           
      Case "COSTO REMITIDO Y FACTURADO STANDARD"
           FUNCIONES_ESPECIFICAS = COSTO_STANDARD_TOTAL
           
      Case "C-CALCALCULADO REMITIDO Y FACTURADO"
           FUNCIONES_ESPECIFICAS = COSTO_CALCULADO_REMITIDOS
           
      Case "FACTU_ANTIC_DEL_MES_ACTIVO"
           FUNCIONES_ESPECIFICAS = FACTU_ANTIC_DEL_MES_ACTIVO
           
      Case "FACTU_ANTIC_ACUM_ANTERIOR"
           FUNCIONES_ESPECIFICAS = FACTU_ANTIC_ACUM_ANTERIOR
           
      Case "FACT_ANT_PEND_APLICACION_MES_ACTIVO"
           FUNCIONES_ESPECIFICAS = FACT_ANT_PEND_APLICACION_MES_ACTIVO
           
      Case "FACT_ANT_PEND_APLICACION_ACUM_ANTERIOR"
           FUNCIONES_ESPECIFICAS = FACT_ANT_PEND_APLICACION_ACUM_ANTERIOR
           
      Case "FACTU_ANTIC_ACUM_A_FIN_MES_ACTIVO"
           FUNCIONES_ESPECIFICAS = FACTU_ANTIC_ACUM_A_FIN_MES_ACTIVO
           
      Case "FACT_ANT_PEND_APLICACION_ACUM_A_FIN_MES_ACTIVO"
           FUNCIONES_ESPECIFICAS = FACT_ANT_PEND_APLICACION_ACUM_A_FIN_MES_ACTIVO
           
      Case "APLICACIONES_DE_FACTU_ANTIC_MES_ACTIVO"
           FUNCIONES_ESPECIFICAS = APLICACIONES_DE_FACTU_ANTIC_MES_ACTIVO
           
      Case "VALOR VENTA REMITIDO MES ACTIVO"
           FUNCIONES_ESPECIFICAS = VALOR_VENTA_REMITIDO_MES_ACTIVO
            
      Case "COSTO REMITIDO DE M.OBRA"
           FUNCIONES_ESPECIFICAS = COSTO_M_OBRA_REMITOS_TODOS
   End Select
   
End Function

Function VALIDA_TIPO_DE_FUNCION$(TXT1$)
   TEXTOX$ = UCase$(TXT1$)
   '
   VALIDA_TIPO_DE_FUNCION$ = ""
   TEXTOCUENTA$ = Mid$(TRIM$(TEXTOX$), 3)

   For KU& = 1 To ULTREG&("SELECUEN")
     XX$ = REGLEIDO$("SELECUEN", KU&)
     NUMTITULO% = CVI(Left$(XX$, 2))
     DESCRICUENTAENTABLA$ = TRIM$(ECOARCH$("TITUDESC", MKI$(NUMTITULO%)))
     If UCase$(DESCRICUENTAENTABLA$) = TRIM$(TEXTOCUENTA$) Then
       FUNCION_CONTABLE% = 1
       Exit For
     End If
   Next KU&
   '
   If FUNCION_CONTABLE% > 0 Then
     VALIDA_TIPO_DE_FUNCION$ = "FUNCION CONTABLE"
   End If
End Function

Private Sub Borratodo_Click()
    Call limpiargrillaytext
    Call ARMARGRILLA(GRID)
End Sub
Sub limpiargrillaytext()
     GRID.Clear
     Text2 = "": Text3 = "": Text4 = "": Text5 = "": Text6 = ""
     Text1 = "": Combo1.Visible = False

End Sub

Private Sub Combo1_Click()
      Text1 = Combo1.TEXT
      Combo1.Visible = False
      Text1.SetFocus
      Call Text1_KeyPress(13)
End Sub

Private Sub Command3_Click()
   Text9 = MESAN$
   FECHA1 = ""
   Fecha2 = ""

   Command3.Enabled = False
   Call FRESHECHO("SELMESA")
   Call SELECHO("SELMESA")
   MESASEL$ = VALACT1$("SELMESA")
   If TRIM$(MESASEL$) <> "" Then
     MESAN$ = MESASEL$
   End If
   Call MESANO$(MESAN$, Des%, Has%)
   '
   FECHA1 = FECHATEX$(Des%)
   Fecha2 = FECHATEX$(Has%)
   Text9 = MESAN$
   Command3.Enabled = True

End Sub

Private Sub Command9_Click()
    Call SELECHO("SELFECHA")
    VDEVV$ = VALACT1$("SELFECHA")
    If VDEVV$ <> "" Then
        FECHA1.TEXT = VDEVV$
    End If

End Sub

Private Sub Ejecutar_Click()
  '
  NUMEROINFORME% = XVALO(Text4)
  If NUMEROINFORME% < 1 Then
    Call COMUNI("No Hay Un Informe Activo")
    Exit Sub
  End If
  
  Borratodo.Enabled = False
  If TRIM$(Text9) = "" Then
     Call COMUNI("Falta Seleccionar el Mes")
     Exit Sub
  End If
  
  If TRIM$(FECHA1) = "" Then
     Call COMUNI("Falta Ingresar Fecha")
     Exit Sub
  End If
  '
  If TRIM$(Fecha2) = "" Then
     Call COMUNI("Falta Ingresar Fecha")
     Exit Sub
  End If

  VUELTA% = 0
  Screen.MousePointer = 11
  Call COPIAR_FORMATO_GRILLA
  For REG& = 1 To GRID.Rows - 1
    For COL& = 1 To GRID.Cols - 1
      
        If VUELTA% < COL& Then
          GRID2.ColWidth(COL&) = GRID.ColWidth(COL&)
          VUELTA% = COL&
          If COL& = GRID.Cols - 1 Then
             VUELTA% = 9999
          End If
        End If
        TXT1$ = GRID.TextMatrix(REG&, COL&)
        If TXT1$ <> "" Then
            If UCase$(Left$(TXT1$, 3)) = "ZPX" Then ' STORED PROCEDURE
              GRID2.TextMatrix(REG&, COL&) = TRIM$(FORMATNUM$(OBTENER_RESULTADO_SP(TXT1$), "F16.2"))
            ElseIf Left$(TXT1$, 2) = "==" Then    'FUNCIONES REGISTRADAS
              If VALIDA_TIPO_DE_FUNCION$(TXT1$) = "FUNCION CONTABLE" Then
                GRID2.TextMatrix(REG&, COL&) = TRIM$(FORMATNUM$(CALCULAR_FUNCION$(TXT1$), "F16.2"))
              Else
                GRID2.TextMatrix(REG&, COL&) = TRIM$(FORMATNUM$(FUNCIONES_ESPECIFICAS(TXT1$), "F16.2"))
                'ACA HAY QUE DEVOLVER EL VALOR DE LA FUNCION.
              End If
            ElseIf Left$(TXT1$, 1) = "=" Then
              'ESTO SE DEBERIA HACER DESPUES
            End If
        End If
    Next COL&
  Next REG&
  GRID.Visible = False
  GRID2.Visible = True
  GRID2.ZOrder 0
  '
  Call control_recursividad(OK%)
  If OK% > 0 Then
    Call ESCRIBIR_RESULTADOS
    Call PONER_ROJO_NEGATIVOS
  End If
  Screen.MousePointer = 1
End Sub  '
Sub PONER_ROJO_NEGATIVOS()
  For REG& = 1 To GRID.Rows - 1
    For COL& = 1 To GRID.Cols - 1
        VALORDEVUELTO# = XVALO(GRID2.TextMatrix(REG&, COL&))
        COLUX% = COL&
        If VALORDEVUELTO# < 0 Then Call COLOREAR_GRILLA_SOLACELDA(GRID2, vbRed, REG&, COLUX%, 1)
    Next COL&
  Next REG&
  
End Sub
Sub COPIAR_FORMATO_GRILLA()
  Call GRID2.Clear
  Call ARMARGRILLA(GRID2)
  GRID2.Rows = GRID.Rows
  GRID2.Cols = GRID.Cols
  GRID2.Height = GRID.Height
  GRID2.Width = GRID.Width
  GRID2.Top = GRID.Top
  GRID2.Left = GRID.Left
  GRID2.BackColor = &HE0E0E0
  GRID2.BackColorFixed = &HC0C0C0
  '
  For REG& = 1 To GRID.Rows - 1
    For COL& = 1 To GRID.Cols - 1
      If IsNumeric(GRID.TextMatrix(REG&, COL&)) Then
         GRID2.TextMatrix(REG&, COL&) = TRIM$(FORMATNUM$(XVALO(GRID.TextMatrix(REG&, COL&)), "F16.2"))
      Else
        GRID2.TextMatrix(REG&, COL&) = GRID.TextMatrix(REG&, COL&)
      End If
    Next COL&
  Next REG&
  '
End Sub

Private Sub eliminainforme_Click()
    Call SELECHO("LISINFOR/Informes Realizados")
    NUMEROINFORME% = XVALO(VALACT1$("LISINFOR"))
    TEXTOINFO$ = VALACT2$("LISINFOR")
    
    If NUMEROINFORME% > 0 Then
          OPX% = COMALTER%("Realmente Desea Eliminar el Informe: " & TEXTOINFO$ & " \\ Cancelar\Eliminar")
          
          If OPX% = 2 Then
             'SI ESTA EN PANTALLA LO QUITA
             If XVALO(Text4) = NUMEROINFORME% Then
                Call Borratodo_Click
             End If
             'BORRA DE LA BASE DE DATOS
             CONDI$ = "CODIGO = " & NUMEROINFORME%
             Call BORRAREGISTROS("ENCAEXCEL", CONDI$, REGAF&)
             Call BORRAREGISTROS("DETAEXCEL", CONDI$, REGAF&)
             'ACTUALIZA LA LISTA DE SELECCION
             Call CARGALISEL("LISINFOR", "ENCAEXCEL", "Codigo/F8;Descripcion/A48", "CODIGO > 0 ORDER BY CODIGO", "MOMESS")
          End If
    End If

End Sub

Private Sub Form_Activate()
    On Error Resume Next
    Text1.SetFocus
    On Error GoTo 0
End Sub

Function ANTICIPOSPENDIENTES()
   If TRIM$(FECHA1) = "" Then
     Call COMUNI("Falta Fecha Desde")
     Exit Function
   ElseIf TRIM$(Fecha2) = "" Then
     Call COMUNI("Falta Fecha Hasta")
     Exit Function
   End If
   
   Desde = FETEXCOR1$(FECHANUM(FECHA1))
   Hasta = FETEXCOR1$(FECHANUM(Fecha2))
   
   SQLX$ = "EXEC ZPXANTICIPOSPENDIENTES '" & Desde & "', '" & Hasta & "'"
   Set RS = FLEXCONN.Execute(SQLX$)
   ANTICIPOSPENDIENTES = RS.Fields(0)
  
End Function

Function ANTICIPOSPEND_ACUM_PREVIOS()
   If TRIM$(Fecha2) = "" Then
     Call COMUNI("Falta Fecha Hasta")
     Exit Function
   End If
   Hasta = FETEXCOR1$(DIANTE(FECHANUM(FECHA1), 1))
   
   SQLX$ = "EXEC ANTICIPOSPEND_ACUM_PREVIOS  " & "'" & Hasta & "'"
   Set RS = FLEXCONN.Execute(SQLX$)
   ANTICIPOSPEND_ACUM_PREVIOS = RS.Fields(0)
End Function

Function ANTICIPOSPEND_ACUMULADOS()
   If TRIM$(Fecha2) = "" Then
     Call COMUNI("Falta Fecha Hasta")
     Exit Function
   End If
   
   Hasta = FETEXCOR1$(FECHANUM(Fecha2))
   
   SQLX$ = "EXEC ZPXANTICIPOSPEND_ACUMULADOS  " & "'" & Hasta & "'"
   Set RS = FLEXCONN.Execute(SQLX$)
   ANTICIPOSPEND_ACUMULADOS = RS.Fields(0)
End Function

Private Sub Form_Load()
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me)
    '
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    '
    Call ABRECONEXION
    Call CREAR_STORED_PROCEDURE
    Call CREARTABLAENCAEXCEL
    Call CREARTABLADETAEXCEL
    Call CREARTABLA_ANCHOCOLUMNA
    
    RESPUESTA = "NOOK"
    Call ARMARGRILLA(GRID)
    Call CARGACOMBO
    Me.KeyPreview = True
    '
    'ACTUALIZO LA FECHA PRIMER DIA DEL MES Y HOY
'    FEX = HOY(HO$)
'    HASS$ = FECHATEX$(FEX)
'    Fecha2 = FECHATEX$(FEX)
'    FECHA1 = "01" + Mid$(HASS$, 3)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Call FINAL("")
End Sub


Sub CARGACOMBO()
    Combo1.Clear
    For U& = 1 To ULTREG&("!TITUDESC")
       X$ = REGLEIDO$("!TITUDESC", U&)
       CODTITU% = CVI(Left(X$, 2))
       DESCRITITULO$ = Mid$(X$, 3)
       Combo1.AddItem "==" & DESCRITITULO$
    Next U&
    
    Set RS = FLEXCONN.OpenSchema(adSchemaProcedures)
    Do Until RS.EOF
            STOR$ = RS!procedure_name
            If UCase$(Left$(STOR$, 3)) <> "DM_" Then
              If UCase$(Left$(STOR$, 3)) <> "FN_" Then
                If UCase$(Left$(STOR$, 3)) <> "SP_" Then
                  If UCase$(Left$(STOR$, 3)) <> "DM_" Then
                    If UCase$(Left$(STOR$, 3)) <> "XP_" Then
                      If Right(TRIM$(STOR$), 2) = ";1" Then STOR$ = Left$(STOR$, Len(STOR$) - 2)
                      Combo1.AddItem STOR$
                    End If
                  End If
                End If
              End If
            End If
    RS.MoveNext
    Loop
    '
    'CARGA FUNCIONES ESPECIALES
    Combo1.AddItem "==" & "COSTO REMITIDO Y FACTURADO DE SUMINISTROS"
    Combo1.AddItem "==" & "COSTO REMITIDO Y FACTURADO DE M.OBRA"
    Combo1.AddItem "==" & "COSTO REMITIDO Y FACTURADO STANDARD"
'    Combo1.AddItem "==" & "C-CALCALCULADO REMITIDO Y FACTURADO"
    Combo1.AddItem "==" & "FACTU_ANTIC_DEL_MES_ACTIVO"
    Combo1.AddItem "==" & "FACTU_ANTIC_ACUM_ANTERIOR"
    'Combo1.AddItem "==" & "FACT_ANT_PEND_APLICACION_MES_ACTIVO"
    Combo1.AddItem "==" & "FACT_ANT_PEND_APLICACION_ACUM_ANTERIOR"
    Combo1.AddItem "==" & "FACTU_ANTIC_ACUM_A_FIN_MES_ACTIVO"
    Combo1.AddItem "==" & "FACT_ANT_PEND_APLICACION_ACUM_A_FIN_MES_ACTIVO"
    Combo1.AddItem "==" & "APLICACIONES_DE_FACTU_ANTIC_MES_ACTIVO"
    Combo1.AddItem "==" & "VALOR VENTA REMITIDO MES ACTIVO"
    Combo1.AddItem "==" & "COSTO REMITIDO DE M.OBRA"




    '
End Sub
Private Sub Command2_Click()
    Call SELECHO("SELFECHA")
    VDEVV$ = VALACT1$("SELFECHA")
    If VDEVV$ <> "" Then
        Fecha2.TEXT = VDEVV$
    End If

End Sub


Private Sub Form_Resize()
    On Error Resume Next
    GRID.Height = Me.Height - 1500
    
    GRID.Width = Me.Width - 235
    If GRID2.Visible = True Then
        GRID2.Height = Me.Height - 1800
        GRID2.Width = Me.Width - 235
    End If
    Text6.Width = Text6.Width + (Me.Width - (Text6.Width + 1000))
    On Error GoTo 0
End Sub

Private Sub grabar_Click()
   NUMEROINFORME% = XVALO(Text4)
   If NUMEROINFORME% < 1 Then Exit Sub
   '
   CONDI$ = "CODIGO = " & NUMEROINFORME%
   MESGRABADO$ = VALOBADA("ENCAEXCEL", "MES", CONDI$, "NOMESS")
   If MESGRABADO$ <> Text9 Then
      OPX% = COMALTER("Realmente Desea Grabar Este Informe Modificando el Mes de Selección de los Datos\\Cancelar Grabación\Si Grabar Pisando el Anterior")
      If OPX% < 2 Then Exit Sub
   End If

   'HAGO ESTA VALIDACION POR SI EN OTRA PC SE ESTA POR GRABAR EL MISMO NUMERO
   YAEXISTEINFORME& = CantRegistros&("ENCAEXCEL", CONDI$, "NOMESS")
   If NUEVOINFORME% > 0 Then
      If YAEXISTEINFORME& > 0 Then
         Text4 = 1 + ULTIMO_INFORME&
         NUMEROINFORME% = XVALO(Text4) 'ACTUALIZO LA VARIABLE NUMEROINFORME%
         'ACTUALIZO LA CONDI
         CONDI$ = "CODIGO = " & NUMEROINFORME%
         
      End If
   End If
   '
   'CREO EL REGISTRO NUEVO EN ENCAEXCEL SI ES NUEVO INFORME
   If NUEVOINFORME% > 0 Then
        Call CreaRegistro("ENCAEXCEL", "CODIGO", NUMEROINFORME%, "NOMESS")
        NUEVOINFORME% = 0
   End If
   Call ACTUREGISTRO("ENCAEXCEL", "DESCRIPCION", CONDI$, Text5, REGAF&, "NOMESS")
   Call ACTUREGISTRO("ENCAEXCEL", "FechaDesde", CONDI$, FETEXCOR1$(CVINT(FECHA1)), REGAF&, "NOMESS")
   Call ACTUREGISTRO("ENCAEXCEL", "FechaHasta", CONDI$, FETEXCOR1$(CVINT(Fecha2)), REGAF&, "NOMESS")
   Call ACTUREGISTRO("ENCAEXCEL", "MES", CONDI$, Text9, REGAF&, "NOMESS")

   '
   Call ACTUREGISTRO("DETAEXCEL", "MARBORRA", CONDI$, 1, REGAF&)
   'CREO LOS REGISTROS EN DETAEXCEL
   'SI EXISTE MARCA DE BORRADO Y ABAJO BORRA, EN CASO QUE NO EXISTSA VA AGREGAR Y NO VA A BORRAR NADA.
   SQLX$ = "SELECT * FROM DETAEXCEL WHERE " & CONDI$
   Dim RS As ADODB.Recordset
   Set RS = New ADODB.Recordset
25 On Error Resume Next
   RS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
   End If
   On Error GoTo 0
   '
   With RS
        For REG& = 1 To GRID.Rows - 1
           For COL& = 1 To GRID.Cols - 1
             FUNCI$ = TRIM$(GRID.TextMatrix(REG&, COL&))
             If FUNCI$ <> "" Then
               .AddNew   ' YA CONVERTIDO
               On Error Resume Next
               !CodiEmpr = CodiEmp%
               On Error GoTo 0
               !Codigo = NUMEROINFORME%
               !FILA = REG&
               !COLUMNA = COL&
               !UBICACION = UBICACIONX$(REG&, COL&)
               !FUNCION = FUNCI$
               !MARBORRA = 0
               .Update
             End If
           Next COL&
       Next REG&
    End With
    RS.Close
    Set RS = Nothing
    '
    Call BORRAREGISTROS("DETAEXCEL", CONDI$ & " And MARBORRA = 1", REGAF&)
    
    Call CARGA_ANCHO_COLUM
    
End Sub
Sub GRID_Click()
    Text2 = ""
    If GRID.MouseRow < 1 Then Exit Sub

    Static ULTIMAPOS$

    If ESCLICK% < 1 Then 'ACA CUANDO VIENE DE ENTER
      COLUM& = COLUMNA&(ULTIMAPOS$, FILAX&) 'FILAX& LO DEVUELVE CON VALOR LA FUNCION CON EL NUMERO DE REGISTRO
      REGIS& = FILAX&
      ULTIMAPOS$ = UBICACIONX$(REGIS& + 1, COLUM&)
      Text7 = ULTIMAPOS$
      GoTo 10
    End If
    If ULTIMAPOS$ = "" Then GoTo 10
    COLUM& = COLUMNA&(ULTIMAPOS$, FILAX&) 'FILAX& LO DEVUELVE CON VALOR LA FUNCION CON EL NUMERO DE REGISTRO
    REGIS& = FILAX&
    If COLUM& < 1 Or FILAX& < 1 Then GoTo 5
    '
5   If (Text1.Top <> GRID.CellTop + GRID.Top - 20) Or (Text1.Left <> GRID.CellLeft) Then
         GRID.TextMatrix(REGIS&, COLUM&) = Text1
    End If
    
    Text1 = ""
10  ESCLICK% = 1
    If GRID.Row > 0 Then
            Text2 = UBICACIONX$(GRID.MouseRow, GRID.MouseCol)
            If CTRL = True Then
               Text3 = Text3 & Text2 & ";"
            Else
               Text3 = ""
            End If
            Text1.Top = GRID.CellTop + GRID.Top
            Text1.Left = GRID.CellLeft + GRID.Left
            Text1.Height = GRID.CellHeight - 10
            Text1.Width = GRID.CellWidth
            Text1.TEXT = TRIM$(GRID.TEXT)
            Text1.Visible = True
            On Error Resume Next
            Command1.SetFocus
            Text1.SetFocus
            On Error GoTo 0
        Else
            GRID.Row = 1
            GRID.COL = 2
            GoTo 10
    End If
    If ESCLICK% > 0 Then
       Text8 = ULTIMAPOS$
       ULTIMAPOS$ = UBICACIONX$(GRID.Row, GRID.COL)
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
        Text8 = PASO% & 0
        If KeyCode = vbKeyControl Then
          If PASO% < 1 Then
            COLACT% = GRID.COL
            REGACT% = GRID.Row
            Text7 = REGACT% & COLACT%
            PASO% = 1
            Text8 = PASO%
            If GRILLA_ACTIVA% > 0 Then CTRL = True
          End If
        End If
        If KeyCode = vbKeyEscape Then
           If Combo1.Visible = True Then Combo1.Visible = False
           If Text1.Visible = True Then Text1.Visible = False
        End If
        If KeyCode = vbKeyF2 Then
        If Text1.Visible = True Then
           Text1.SetFocus
        End If
        If KeyCode = vbKeyEscape Then
          Text1.Visible = False
          Combo1.Visible = False
        End If
     End If

End Sub
  
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   PASO% = 0
   Text8 = PASO%
   If KeyCode = vbKeyControl Then
    CTRL = False
 
    If GRILLA_ACTIVA% > 0 Then
        Text1 = Text3
        If Right(TRIM$(Text1), 1) = ";" Then
           Text1 = Mid$(Text1, 1, Len(TRIM$(Text1)) - 1) & ")"
        End If
        GRID.Row = REGACT%
        GRID.COL = COLACT%
        Call Text1_KeyPress(13)
        GRILLA_ACTIVA% = 0
    End If
   End If
End Sub

Private Sub GRID2_Click()
   Text3 = GRID.TextMatrix(GRID2.Row, GRID2.COL)
   Text2 = UBICACIONX$(GRID2.Row, GRID2.COL)
End Sub

Private Sub imprimir_Click()
   '
   
   If GRID.Visible = False Then
      Dim OBJ_GRID As MSFlexGrid
      Set OBJ_GRID = GRID2
      Call GENEXCEL_GRID(OBJ_GRID)
   Else
      Call COMUNI("Falta Ejecutar La Planilla Para Poder Exportar")
      GoTo 99
   End If
   '
99
End Sub
Sub GENEXCEL_GRID(GRID As MSFlexGrid)
   Screen.MousePointer = 11
   Dim FECHA As Date
   Dim valor As Double
   Dim TEXTO As String
   Dim FORMATONUMERO As String
   Dim objExcel As Excel.Application

   '
   On Error GoTo 20
   Set objExcel = New Excel.Application
   objExcel.SheetsInNewWorkbook = 1
   objExcel.Workbooks.Add
   On Error GoTo 0
   '

   TIPODATO% = 0: COLUM& = 0
   For i& = 1 To GRID.Rows - 1
       RENG1& = i&: COLUM& = 0: CampoConEco% = 0: ColumnaReal = 0: VieneDeUnEco% = 0
       For COLU1& = 1 To GRID.Cols - 1
          If GRID.ColWidth(COLU1&) = 0 Then GoTo 10  ' OBVIA LAS COLUMNAS OCULTAS
          COLUM& = COLUM& + 1
          ColumnaReal = ColumnaReal + 1
          '
          VALORX = GRID.TextMatrix(i&, COLU1)
          If Len(VALORX) = 8 And Mid$(VALORX, 3, 1) = "/" And Mid$(VALORX, 6, 1) = "/" Then
             FEXU$ = VALORX
             VALORX = FETEXCOR1$(Int(FECHANUM(FEXU$)))
             TIPODATO% = 1
          End If
          ' ESTO ES PARA QUE EXPORTE COMO CERO Y NO ESPACIOS
          If CampoConEco% = 0 Then
            FFCC$ = UCase$(Left$(FORCAMPO$(ColumnaReal), 1))
          Else
            ColumnaReal = ColumnaReal
            FFCC$ = UCase$(Left$(FORMAECO$(ColumnaReal), 1)) ' UCase$(Left$(FORCAMPO$(ColumnaFormato), 1))
            CampoConEco% = 0: VieneDeUnEco% = 1
          End If
          If FFCC$ = "I" Or FFCC$ = "F" Or FFCC$ = "G" Or FFCC$ = "$" Then
             If TRIM$(CStr(VALORX)) = "" Then VALORX = 0
             
             FORMATONUMERO = "0" ' FORMATO ENTERO
             If FFCC$ = "F" Or FFCC$ = "G" Or FFCC$ = "$" Then
                MASCARA$ = Mid$(FORCAMPO$(ColumnaReal), 2)
                NN% = InStr(1, MASCARA$, ".")
                PARTEDECIMAL$ = Mid$(MASCARA$, NN% + 1)
                
                JJ% = 1
                If FFCC$ = "G" Or FFCC$ = "$" Then
                    FORMATONUMERO = "#,##0."
                    VALORX = XVALO(REPLACAR(CStr(VALORX), ".", ""))
                ElseIf FFCC$ = "F" Then
                    FORMATONUMERO = "#0."
                End If
                While JJ% <= XVALO(PARTEDECIMAL$)
                    FORMATONUMERO = FORMATONUMERO + "0"
                    JJ% = JJ% + 1
                Wend
             End If
             TIPODATO% = 2
          End If
          If FFCC$ = "C" Or FFCC$ = "A" Then VALORX = "'" & VALORX
          '
          objExcel.ActiveSheet.Cells(RENG1&, COLUM&) = VALORX
          If TIPODATO% = 1 Then XLSHEET.Cells(RENG1&, COLUM&).NumberFormat = "dd/mm/yyyy"
          If TIPODATO% = 2 Then XLSHEET.Cells(RENG1&, COLU1&).NumberFormat = FORMATONUMERO
          'If TIPODATO% = 0 Then XLSHEET.Cells(RENG1&, COLUM&).NumberFormat = "@"
          TIPODATO% = 0
          '
          'If TRIM$(TABLAECO(COLU1&)) <> "" Then CampoConEco% = 1
          If TRIM$(TABLAECO(ColumnaReal)) <> "" And VieneDeUnEco% = 0 Then CampoConEco% = 1
          VieneDeUnEco% = 0
          '
          With objExcel.ActiveSheet.Cells(i&, COLUM&).Font
            If XVALO(VALORX) < 0 Then
              .COLOR = vbRed
            End If
          End With
10     Next COLU1&
   Next i&
   '
   Close #N1%
   On Error GoTo 30
   objExcel.ActiveSheet.Columns("A:Z").AutoFit
   Screen.MousePointer = 1
   objExcel.Visible = True

   On Error GoTo 0
   '
   '
   Screen.MousePointer = 1
   '
Exit Sub
   '
   '
20 Screen.MousePointer = 1
   Resume 21
21 On Error GoTo 0
   MsgBox "Utilitario EXCEL no Instalado"
   On Error GoTo 0
   Exit Sub
   '
30 Resume 31
31 On Error GoTo 0
   Screen.MousePointer = 1
   '
   Call MENSERR(24, "Imposible Crear Planilla Excel.")
   Exit Sub
   '
End Sub


Private Sub Label5_DblClick()
   Call ControlFactura
End Sub


Private Sub mnuInstructivo1_Click()
     If Text10.Visible = True Then
     
        Text10.Visible = False
     Else
        Call COMUNI("Para Ocultar la ventana\Repita los Pasos de Visualizar Nuevamente Desde el Menù")
        Text10.Visible = True
     End If
End Sub

Private Sub ModoEdición_Click()
    If GRID.Visible = True Then
       Call COMUNI("Yá Se Encuentra en Modo Edición")
    Else
       Call Form_KeyDown(vbKeyEscape, 0) ' POR SI ESTA EL TEXT1 VISIBLE O EL COMBO1 VISIBLE EN LA GRILLA
       GRID.Visible = True
       GRID2.Top = GRID.Top + GRID.Height + 800
    End If
    Borratodo.Enabled = True

End Sub

Private Sub nuevo_Click()
    NUMESIGUIENTE% = 1 + ULTIMO_INFORME&
5   DESCRIX$ = TRIM$(InputBox$("Título", "Ingrese Título de Informe", DESCRIX$))
    If DESCRIX$ = "" Then
      Exit Sub
     Else
      Text4 = SAFETEXT$(NUMESIGUIENTE%)
      Text5 = DESCRIX$
      NUEVOINFORME% = 1
      '
      'LIMPIO LA GRILLA
      GRID.Clear
      Text2 = "": Text3 = ""
      Text9 = ""
      FECHA1 = ""
      Fecha2 = ""
      OPX% = COMALTER%("Opciones de Trabajo\\Desde Cero\Por Copia")
      If OPX% = 2 Then
        Call CARGALISEL("LISINFOR", "ENCAEXCEL", "Codigo/F8;Descripcion/A48", "CODIGO > 0 ORDER BY CODIGO", "MOMESS")
        Call SELECHO("LISINFOR/Informes Realizados")
        NUMEROINFORME% = XVALO(VALACT1$("LISINFOR"))
        Call CARGAR_DATOS_ENGRILLA(NUMEROINFORME%)
        Call FORMATCOLUM(GRID, NUMEROINFORME%)
      End If
      
      If GRID2.Visible = True Then
         GRID2.Visible = False
         GRID.Visible = True
'         GRID.ZOrder 0
         Call ARMARGRILLA(GRID)
         Combo1.Visible = False
         Text1.Visible = False
         Call grabar_Click
      End If
    End If
End Sub
Function ULTIMO_INFORME&()
   CONDI$ = "CODIGO > 0 "
   ULTIMO_INFORME& = XVALO(VALOBADA("ENCAEXCEL", "MAX(CODIGO)", CONDI$, "NOMESS"))
End Function
Private Sub seleccionarlista_Click()
End Sub


Private Sub quitarespacioslibres_Click()
   NUMEROINFORME% = XVALO(Text4)

   If NUMEROINFORME% < 1 Then
     Call COMUNI("No hay Ningún Informe Activo")
     Exit Sub
   End If

5  FILASX& = XVALO(InputBox$("Indicar Cantidad de Filas a Subir", "Ingrese Cantidad de Filas", FILASX&))
   If FILASX& < 1 Then Exit Sub
   For i& = 1 To FILASX&
      For j& = 1 To GRID.Cols - 1
         valor$ = TRIM$(GRID.TextMatrix(i&, j&))
         If valor$ <> "" Then
            Call COMUNI("Atención Celdas con Información en Fila: " & i&)
            Exit Sub
         End If
      Next j&
   Next i&
   '
   GRID.Rows = 1
   Call ARMARGRILLA(GRID)
   Call CARGAR_DATOS_ENGRILLA(NUMEROINFORME%, FILASX&)
   Call FORMATCOLUM(GRID, NUMEROINFORME%)
   '
End Sub
Private Sub salir_Click()
  Call FINAL("")
End Sub

Private Sub seleccioncuentas_Click()
       Call TRALOCAL("SELECUEN", "")
10     VTB% = VENTABU%("SELECUEN")
       If VTB% >= 0 Then
          For U& = 1 To ULTREG&("!SELECUEN")
             X$ = REGLEIDO$("!SELECUEN", U&)
             SIMBOLO$ = TRIM$(Mid$(X$, 5, 1))
             If SIMBOLO$ <> "+" And SIMBOLO$ <> "-" Then
                 Call COMUNI("Operador Ingresado (" & SIMBOLO$ & ") En Fila: " & U& & "\No Válido\Operadores Posibles '+' o '-'")
                 GoTo 10
                 Exit For
             End If
          Next U&
          Call TRACENTRAL("SELECUEN", "")
       End If
       Call CARGACOMBO
End Sub

Private Sub seleccionInforme_Click()
    Call CARGALISEL("LISINFOR", "ENCAEXCEL", "Codigo/F8;Descripcion/A48", "CODIGO > 0 ORDER BY CODIGO", "MOMESS")
    Call SELECHO("LISINFOR/Informes Realizados")
    NUMEROINFORME% = XVALO(VALACT1$("LISINFOR"))

    If NUMEROINFORME% > 0 Then
           GRID.Clear
           Call ARMARGRILLA(GRID)
           NUEVOINFORME% = 0
           Text4 = 0
           Text4 = SAFETEXT$(NUMEROINFORME%)
           Text5 = VALACT2$("LISINFOR")
           Borratodo.Enabled = True
           Combo1.Visible = False
           Text1.Visible = False
           If GRID2.Visible = True Then
             GRID2.Visible = False
             GRID.Visible = True
          End If
    End If
    Call FORMATCOLUM(GRID)

End Sub

Private Sub Text1_Change()
   If CTRL <> True Then
     Text3 = Text1
   End If
End Sub

Private Sub Text1_DblClick()
    If GRID.Row < 1 Then GRID.Row = 1

10  If GRID.Row > 0 Then
            
            Combo1.Top = GRID.CellTop + GRID.Top
            Combo1.Left = GRID.CellLeft + GRID.Left
            'Combo1.Height = GRID.CellHeight - 10
            Combo1.Width = GRID.CellWidth
            Combo1.TEXT = TRIM$(GRID.TEXT)
            Combo1.Visible = True
            Combo1.SetFocus
            On Error Resume Next
            Command1.SetFocus
            Combo1.SetFocus
            On Error GoTo 0
        Else
            GRID.Row = 1
            GRID.COL = 2
            GoTo 10
    End If
    
End Sub



Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    FLECHALATERAL% = 0
    If KeyCode = 40 Then
'        GRID.TEXT = TRIM$(Text1.TEXT)
'        GRID.Refresh
'        If GRID.Row >= GRID.Rows - 1 Then Exit Sub
'        GRID.Row = GRID.Row + 1
'        Call GRID_Click
    End If
    If KeyCode = 38 Then
'        GRID.TEXT = TRIM$(Text1.TEXT)
'        GRID.Refresh
'        If GRID.Row <= 1 Then Exit Sub
'
'        GRID.Row = GRID.Row - 1
'        Call GRID_Click
     ElseIf KeyCode = 37 Then
'        If GRID.COL <= 1 Then Exit Sub
'        GRID.TEXT = TRIM$(Text1.TEXT)
'        GRID.Refresh
'
'        GRID.COL = GRID.COL - 1
'        FLECHALATERAL% = 1
'        Call GRID_Click
     ElseIf KeyCode = 39 Then
'        If GRID.COL >= GRID.Cols - 1 Then Exit Sub
'        GRID.TEXT = TRIM$(Text1.TEXT)
'        GRID.Refresh
'        GRID.COL = GRID.COL + 1
'        FLECHALATERAL% = 1
'        Call GRID_Click
     ElseIf KeyCode = vbKeyF2 Then
         Call Text1_DblClick
     End If
     
'        Text2 = ""
'        Text2 = UBICACIONX$(GRID.Row, GRID.COL)

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        GRID.TEXT = TRIM$(Text1.TEXT)
        GRID.Refresh
        GRID.Row = GRID.Row + 1
        ESCLICK% = (-1)
        Call GRID_Click
        Text2 = ""
        Text2 = UBICACIONX$(GRID.Row, GRID.COL)
    End If

End Sub
Private Sub Text1_GotFocus()
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1)
End Sub
'
Sub SETNEXTCOL()
   '
   If YACARDA% < 1 Then Exit Sub
   COLVIE% = GRID.COL
   ROWVIE% = GRID.Row
   '
5  COLNUE = GRID.COL + 1
   ROWNUE = GRID.Row
   UREPOX& = UREPOS&
   '
10 If ROWNUE > UREPOX& Then If APPI% < 1 Then GoTo 99
   If COLNUE >= NUCOLUM% Then
     COLNUE = 0
     If ROWNUE < UREPOX& Then
         ROWNUE = ROWNUE + 1
       Else
         If APPI% < 1 Then GoTo 99
     End If
   End If
   If ROWNUE <= ULTREG&(ARCHILOC$) Then
     If ROWNUE <> INSEROW& Then
       If PE%(COLNUE + 1) <> 1 Or ACL%(COLNUE + 1) < 1 Then
         COLNUE = COLNUE + 1
         GoTo 10
       End If
     End If
   End If
   ' Text1 = ""         ' agregado EPB el 14/11/2011 por el arrastre de valor a otro campo en MEDITEA - BRecepcion
   GRID.COL = COLNUE
   GRID.Row = ROWNUE
   If GRID.CellWidth < 0 Then GoTo 5
   '
99 Exit Sub
End Sub

Function UBICACIONX$(REG&, COL&)
            LETRAX$ = Chr$(COL& + 64)
            NUMBERX$ = SAFETEXT$(REG&)
            UBICACIONX$ = LETRAX$ & NUMBERX$
End Function

Private Sub Text2_Change()
'   Check1.Value = 0
'   If TRIM$(Text2) <> "" Then
'        UBICA$ = TRIM$(Text2)
'        COLUM& = COLUMNA&(UBICA$, FILAX&) 'FILAX& LO DEVUELVE CON VALOR LA FUNCION CON EL NUMERO DE REGISTRO
'        GRID.Row = FILAX&
'        GRID.COL = COLUM&
'        If GRID.CellFontStrikeThrough = True Then
'          Check1.Value = 1
'        Else
'          Check1.Value = 0
'        End If
'        Text1.SetFocus
'   End If
End Sub

Private Sub Text3_Change()
   If TRIM$(Text2) <> "" Then
     Text1 = Text3
     If UCase$(Left$(Text3, 6)) = "=SUMA(" Or UCase$(Left$(Text3, 1)) = "=" Then
       GRILLA_ACTIVA% = 1
     End If
   End If
   '
   'EN EL CASO QUE SEA UN SP MUESTRO LOS PARAMETROS DEL MISMO.
   texto3$ = TRIM$(Text3)
   Text6 = ""
   If texto3$ <> "" Then
    NAME_SP$ = texto3$
    If Left$(NAME_SP$, 3) <> "ZPX" Then Exit Sub
    Set rsrecordset = FLEXCONN.Execute("sp_procedure_params_rowset" & " " & NAME_SP$)
    Do Until rsrecordset.EOF
         If InStr(UCase(rsrecordset!parameter_name), "@RETURN_VALUE") < 1 Then
           Text6 = Text6 + rsrecordset!parameter_name + ", "
         End If
         rsrecordset.MoveNext
    Loop
    Text6 = TRIM$(Text6)
    If InStr(Text6, ",") > 0 Then Text6 = Mid$(Text6, 1, Len(Text6) - 1)
   End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
     GRID.TEXT = TRIM$(Text3.TEXT)
     GRID.Refresh
     Text3 = ""
'     Text1.SetFocus
'     Call Text1_KeyPress(13)
     Text1.Visible = False
     
   End If
End Sub
Sub CREARTABLAENCAEXCEL()
   Call CREACAMPO("", "ENCAEXCEL", "CODIGO", 3, 0, 1)
   Call CREACAMPO("", "ENCAEXCEL", "DESCRIPCION", 10, 128)
   Call CREACAMPO("", "ENCAEXCEL", "FechaDesde", 8, 0)
   Call CREACAMPO("", "ENCAEXCEL", "FechaHasta", 8, 0)
   Call CREACAMPO("", "ENCAEXCEL", "MES", 10, 24)

End Sub
Sub CREARTABLADETAEXCEL()
   Call CREACAMPO("", "DETAEXCEL", "CODIGO", 3, 0, 1)
   Call CREACAMPO("", "DETAEXCEL", "FILA", 3, 0)
   Call CREACAMPO("", "DETAEXCEL", "COLUMNA", 3, 0)
   Call CREACAMPO("", "DETAEXCEL", "UBICACION", 10, 12)
   Call CREACAMPO("", "DETAEXCEL", "FUNCION", 10, 128)
   Call CREACAMPO("", "DETAEXCEL", "MARBORRA", 3, 0, 1)
End Sub
Sub CREARTABLA_ANCHOCOLUMNA()
   Call CREACAMPO("", "ANCHCOLU", "CODIGO", 3, 0, 1)
   Call CREACAMPO("", "ANCHCOLU", "NUMCOLUMNA", 3, 0, 1)
   Call CREACAMPO("", "ANCHCOLU", "ANCHO", 8, 0)
   Call CREACAMPO("", "ANCHCOLU", "MARBORRA", 3, 0)
End Sub

Sub ARMARGRILLA(GRID As MSFlexGrid)
     GRID.ColWidth(0) = 330
     GRID.Rows = 100
     GRID.Cols = 27

     For i& = 1 To GRID.Rows - 1
       GRID.TextMatrix(i&, 0) = i&
       GRID.RowHeight(i&) = 350
    Next i&
    For i& = 1 To GRID.Cols - 1
        GRID.TextMatrix(0, i&) = Chr(i& + 64)
'        GRID.ColWidth(I&) = 1800
    Next i&
End Sub


Private Sub Text4_Change()
    NUMEROINFORME% = XVALO(Text4)
    If CantRegistros&("ENCAEXCEL", "CODIGO = " & NUMEROINFORME%) > 0 Then
       Screen.MousePointer = 11
       CONDI$ = "CODIGO = " & NUMEROINFORME%
       FECHA1 = FECHATEX$(CVINT(VALOBADA("ENCAEXCEL", "FechaDesde", CONDI$, "NOMESS")))
       Fecha2 = FECHATEX$(CVINT(VALOBADA("ENCAEXCEL", "FechaHasta", CONDI$, "NOMESS")))
       Text9 = VALOBADA("ENCAEXCEL", "MES", CONDI$, "NOMESS")
       Call CARGAR_DATOS_ENGRILLA
       Screen.MousePointer = 1
    End If
    
End Sub
'Sub CARGAR_DATOS_ENGRILLA(Optional NUMEROINFO%)
'    '
'    NUMEROINFORME% = NUMEROINFO%
'    If NUMEROINFORME% < 1 Then NUMEROINFORME% = XVALO(Text4)
'    SQLX$ = "SELECT FILA,COLUMNA,FUNCION"
'    SQLX$ = SQLX$ + " FROM DETAEXCEL WITH(NOLOCK)"
'    SQLX$ = SQLX$ + " WHERE CODIGO = " & NUMEROINFORME%
'    Set RS = READSET(SQLX$)
'    With RS
'        Do While Not .EOF
'          REG& = XVALO(!FILA)
'          COL& = XVALO(!COLUMNA)
'          FUNCI$ = SAFETEXT$(!FUNCION)
'          GRID.TextMatrix(REG&, COL&) = FUNCI$
'          .MoveNext
'        Loop
'    End With
'    RS.Close
'    Set RS = Nothing
'    End Sub
Sub CARGAR_DATOS_ENGRILLA(Optional NUMEROINFO%, Optional CANTFILAS_LIBRES&)
    '
    
    NUMEROINFORME% = NUMEROINFO%
    If NUMEROINFORME% < 1 Then NUMEROINFORME% = XVALO(Text4)
    SQLX$ = "SELECT FILA,COLUMNA,FUNCION"
    SQLX$ = SQLX$ + " FROM DETAEXCEL WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE CODIGO = " & NUMEROINFORME%
    Set RS = READSET(SQLX$)
    With RS
        Do While Not .EOF
          REG& = XVALO(!FILA)
          If CANTFILAS_LIBRES& > 0 And REG& > CANTFILAS_LIBRES& Then REG& = REG& - CANTFILAS_LIBRES&
          COL& = XVALO(!COLUMNA)
          FUNCI$ = SAFETEXT$(!FUNCION)
          If FUNCI$ = "Gastos Financieros" Then
            A = 1
          End If
          GRID.TextMatrix(REG&, COL&) = FUNCI$
          .MoveNext
        Loop
    End With
    RS.Close
    Set RS = Nothing
End Sub

Private Sub FECHA1_LostFocus()
   If TRIM$(FECHA1) <> "" Then
        If IsDate(FECHA1) = False Then
           Call COMUNI("Fecha Inválida")
           FECHA1.SetFocus
           Exit Sub
        End If
        If CVINT(FECHA1) < 1 Then
              AAA = CVINT(FECHA1)
              Call COMUNI("Fecha Inválida")
              FECHA1.SetFocus
              Exit Sub
        End If
   End If

   FECHA1 = FECHATEX$(CVINT(FECHA1))
   
End Sub
Private Sub FECHA2_LostFocus()
   If TRIM$(Fecha2) <> "" Then
        If IsDate(Fecha2) = False Then
           Call COMUNI("Fecha Inválida")
           Fecha2.SetFocus
           Exit Sub
        End If
        If CVINT(Fecha2) < 1 Then
              AAA = CVINT(FECHA1)
              Call COMUNI("Fecha Inválida")
              Fecha2.SetFocus
              Exit Sub
        End If
   End If
   Fecha2 = FECHATEX$(CVINT(Fecha2))
   
End Sub

Function CALCULAR_FUNCION$(TEXTOX$)
   '
   SALDO_AFECHA_HASTA% = 0
   TEXTOX$ = UCase$(TEXTOX$)
   If Left(TEXTOX$, 12) = "==ACUMULADO-" Then SALDO_AFECHA_HASTA% = 1
   '
   TEXTOCUENTA$ = Mid$(TRIM$(TEXTOX$), 3)
   Dim CUEI%(1024), SALFIN#(1024), SALINI#(1024), IMPORT#(1024), OPERADOR$(1024)
   CONTA% = 0
   '
   FEDES_1% = DIANTE(FECHANUM(FECHA1), 1)
   FEDES% = FECHANUM(FECHA1)
   FEHAS% = FECHANUM(Fecha2)
   'If FESEL% < COMEJ% Or FESEL% > FINEJ% Then Exit Function
   '
   For KU& = 1 To ULTREG&("SELECUEN")
     XX$ = REGLEIDO$("SELECUEN", KU&)
     NUMTITULO% = CVI(Left$(XX$, 2))
     DESCRICUENTAENTABLA$ = TRIM$(ECOARCH$("TITUDESC", MKI$(NUMTITULO%)))
     If UCase$(DESCRICUENTAENTABLA$) = TRIM$(TEXTOCUENTA$) Then
       CUEII% = CVI(Mid$(XX$, 3, 2))
       CONTA% = CONTA% + 1
       CUEI%(CONTA%) = CUEII%
       OPERADOR$(CONTA%) = TRIM$(Mid$(XX$, 5))
     End If
   Next KU&
   '
   
   'PRIMERO SACO EL SALDO ANTERIOR A LA FECHA DESDE
   For KI% = 1 To CONTA%
     CUEII% = CUEI%(KI%)
     If CUEII% = 891 Then
      A = 1
      DIAXX = FECHATEX$(FEDES_1%)
      End If
      
     IMPORT#(KI%) = SALCUENTA#(CUEII%, FEDES_1%)
     If OPERADOR$(CONTA%) = "-" Then
        SALDO_ANT_DESDE# = SALDO_ANT_DESDE# - IMPORT#(KI%)
     Else
        SALDO_ANT_DESDE# = SALDO_ANT_DESDE# + IMPORT#(KI%)
     End If
   Next KI%
   '
   'LUEGO SACO EL SALDO A LA FECHA HASTA
   For KI% = 1 To CONTA%
     CUEII% = CUEI%(KI%)
     IMPORT#(KI%) = SALCUENTA#(CUEII%, FEHAS%)
     If OPERADOR$(CONTA%) = "-" Then
        SALDO_HASTA# = SALDO_HASTA# - IMPORT#(KI%)
     Else
        SALDO_HASTA# = SALDO_HASTA# + IMPORT#(KI%)
     End If
   Next KI%
   '
   'SE EVALUA SI SOLO ES SALDO A FECHA HASTA O SI ES ENTRE UN RANGO DE FECHAS.
   If SALDO_AFECHA_HASTA% > 0 Then
      TOTALX# = SALDO_HASTA#
   Else
      TOTALX# = SALDO_HASTA# - SALDO_ANT_DESDE#
   End If
   '
'   TOTALX# = SALDO_HASTA# - SALDO_ANT_DESDE#
   'ACA CALCULO EL SALDO HASTA LA FECHA DESDE MENOS EL SALDO A LA FECHA HASTA
   CALCULAR_FUNCION$ = FORMATNUM$(TOTALX#, "F12.2")
   '
End Function


Function SALCUENTA#(CUEII%, FESEL%)
   '
   Call ABRECABAN
   Call ABREPLANCUEN
   '
    ' SUCUI% = UNINEGO%
    SQLX$ = "SELECT SUM(IDEBEPASE) AS SUDEBE, SUM(IHABEPASE) AS SUHABE"
    SQLX$ = SQLX$ + " FROM DETAASIEN WITH(NOLOCK) WHERE "
    SQLX$ = SQLX$ + "FECHASIEN <= '" & FETEXCOR1$(FESEL%, "HASTA") & "' "
    SQLX$ = SQLX$ + "AND CUECOINT = " & CUEII% & " "
    ' SQLX$ = SQLX$ + "AND SUCUASIen = " & SUCUI% & " "
    SQLX$ = SQLX$ + "AND STATPASE > 0; "
    Dim MODIASEL As ADODB.Recordset
    Set MODIASEL = New ADODB.Recordset
    Set MODIASEL = FLEXCONN.Execute(FILTSQL$(SQLX$))
    SALCU# = XVALO(MODIASEL!SUDEBE) - XVALO(MODIASEL!SUHABE)
    GoTo 19
   ' CARGAR MOVIMIENTOS DE DIARIO
   ' SUCUI% = UNINEGO%
   SQLX$ = "SELECT NuIdAsien, IDEBEPASE, IHABEPASE "
   SQLX$ = SQLX$ + " FROM DETAASIEN WHERE "
   SQLX$ = SQLX$ + "FECHASIEN <= '" & CVDAT(FESEL%) & "' "
   SQLX$ = SQLX$ + "AND CUECOINT = " & CUEII% & " "
   ' SQLX$ = SQLX$ + "AND SUCUASIen = " & SUCUI% & " "
   SQLX$ = SQLX$ + "AND STATPASE > 0; "
   '
   'Dim MODIASEL As ADODB.Recordset
   Set MODIASEL = New ADODB.Recordset
   MODIASEL.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText

   
   SALCU# = 0
   With MODIASEL
     On Error Resume Next
     .MoveFirst
     If Err < 1 Then
       On Error GoTo 0
       Do While Not .EOF
'         TIASI% = 0
'         NUAS$ = !NuIdAsien
'         Encasien.FindFirst "NUIDASIEN = '" & NUAS$ & "'"
'         If Encasien.NoMatch = False Then
'           TIASI% = Encasien!TIPOASIEN
'         End If
'         If TIASI% <> 9 Then
          SALCU# = SALCU# + !IDebePase - !IHabePase
'         End If
         .MoveNext
       Loop
     End If
     On Error GoTo 0
   End With
   '
19 MODIASEL.Close
   Set MODIASEL = Nothing
             '
             SQLX$ = "SELECT SUM(IDEBECOMP) AS SUDEBE, SUM(IHABECOMP) AS SUHABE  FROM CAJABANC WITH(NOLOCK) "
             SQLX$ = SQLX$ + "WHERE CueContaI = " & CUEII% & " "
             SQLX$ = SQLX$ + " AND (NuAsidiar = '' or NuAsiDiar is null) "
             SQLX$ = SQLX$ + " AND FechConta <= '" & FETEXCOR1$(FESEL%, "HASTA") & "' "
             Dim CABATEMP As ADODB.Recordset
             Set CABATEMP = New ADODB.Recordset
             Set CABATEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
             SALCU# = SALCU# + XVALO(CABATEMP!SUDEBE) - XVALO(CABATEMP!SUHABE)
             CABATEMP.Close
             Set CABATEMP = Nothing
             GoTo 90
             '
   
   ' CARGAR MOVIMIENTOS DE CAJA Y BANCOS
   SQLX$ = "SELECT * FROM CAJABANC "
   SQLX$ = SQLX$ + "WHERE CueContaI = " & CUEII% & " "
   SQLX$ = SQLX$ + " AND (NuAsidiar = '' or NuAsiDiar is null) "
   SQLX$ = SQLX$ + " AND FechConta <= '" & CVDAT(FESEL%) & "' "
   'Dim CABATEMP As ADODB.Recordset
   Set CABATEMP = New ADODB.Recordset
   CABATEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
   '
   With CABATEMP
     On Error Resume Next
     .MoveFirst
     If Err < 1 Then
       On Error GoTo 0
       Do While Not .EOF
         SALCU# = SALCU# + !IDEBECOMP - !IHABECOMP
         .MoveNext
       Loop
     End If
     On Error GoTo 0
   End With
   Set CABATEMP = Nothing
   '
90 SALCUENTA# = SALCU#
   '
End Function
Private Sub tituloinformes_Click()
    Call TRALOCAL("TITUDESC", "")
    VTB% = VENTABU%("TITUDESC")
    '
    Call TRACENTRAL("TITUDESC", "")
End Sub
Sub control_recursividad(OK%)
    '=suma(D2;D3;D4;D5)
    OK% = 1
    For REG& = 1 To GRID.Rows - 1
      For COL& = 1 To GRID.Cols - 1
            If VALIDA_RECURSIVIDAD%(REG&, COL&) > 0 Then
               OK% = -1
               Exit For
            End If
      Next COL&
    Next REG&
    

End Sub

Function VALIDA_RECURSIVIDAD%(REG&, COL&)
'ejemplo
'recorro y me paro =suma(c4) que esta ubicada en posicion E5 debo recorrer
'y ver que C4 no apunte a E5

   VALIDA_RECURSIVIDAD% = 0
   'DATO ACTIVO =SUMA(A6;B4)
   'LEO EL TEXTO DE LA CELDA QUE QUIERO VALIDAR QUE NO TENGA RECURSIVIDAD
   TEX_CELDA_ACTIVA$ = UCase$(TRIM$(GRID.TextMatrix(REG&, COL&)))
   If Left$(TEX_CELDA_ACTIVA$, 1) = "=" Then
   
        TEX_CELDA_ACTIVA$ = REPLACAR(TEX_CELDA_ACTIVA$, "=SUMA", "")
        TEX_CELDA_ACTIVA$ = REPLACAR(TEX_CELDA_ACTIVA$, "=", "")
        
        TEX_CELDA_ACTIVA$ = REPLACAR(TEX_CELDA_ACTIVA$, "*", ";")
        TEX_CELDA_ACTIVA$ = REPLACAR(TEX_CELDA_ACTIVA$, "-", ";")
        TEX_CELDA_ACTIVA$ = REPLACAR(TEX_CELDA_ACTIVA$, "+", ";")
        TEX_CELDA_ACTIVA$ = REPLACAR(TEX_CELDA_ACTIVA$, "/", ";")
        'LEO A LAS CELDAS A QUE APUNTA PARA VALIDAR QUE NO HAYA NINGUNA QUE APUNTA HACIA LA CELDA QUE ESTOY VALIDANDO
        STRARRAY = Split(TEX_CELDA_ACTIVA$, ";")
        'RECORRO CADA CELDA A QUE APUNTA MI FUNCION
        For i& = LBound(STRARRAY) To UBound(STRARRAY)
           POS_CELDA_LEIDA$ = TRIM(STRARRAY(i&)) 'A6
           POS_CELDA_LEIDA$ = REPLACAR$(POS_CELDA_LEIDA$, "(", "")
           POS_CELDA_LEIDA$ = REPLACAR$(POS_CELDA_LEIDA$, ")", "")
           COLUM& = COLUMNA&(POS_CELDA_LEIDA$, FILAX&) 'FILAX& LO DEVUELVE CON VALOR LA FUNCION CON EL NUMERO DE REGISTRO
           If COLUM& < 1 Or FILAX& < 1 Then Exit For
           REGIS& = FILAX&
           TEX_CELDA_A_VALIDAR$ = GRID.TextMatrix(REGIS&, COLUM&)
           TEX_CELDA_A_VALIDAR$ = REPLACAR(TEX_CELDA_A_VALIDAR$, "=SUMA(", "")
           TEX_CELDA_ACTIVA$ = REPLACAR(TEX_CELDA_ACTIVA$, "=", "")
           TEX_CELDA_A_VALIDAR$ = REPLACAR(TEX_CELDA_A_VALIDAR$, ")", "")
           strArray2 = Split(TEX_CELDA_A_VALIDAR$, ";")
           For JJ& = LBound(strArray2) To UBound(strArray2)
              LEIDA$ = UCase$(TRIM(strArray2(JJ&)))
              CELDACOMPARA$ = UBICACIONX$(REG&, COL&)
              If LEIDA$ = UBICACIONX$(REG&, COL&) Or POS_LEIDA = CELDACOMPARA$ Then
                Call COMUNI("Se ha Detectado un Error de Recursividad \En las Celdas: " & UBICACIONX$(REG&, COL&) & " Y La Celda: " & UBICACIONX$(REGIS&, COLUM&))
                VALIDA_RECURSIVIDAD% = 1
                Exit Function
              End If
           Next JJ&
        Next i&
   End If
End Function



Function VALORCELDA#(UBICAX$)
      
      COL& = COLUMNA&(UBICAX$, FILAX&)
      REG& = FILAX&
      CELDA$ = TRIM$(GRID2.TextMatrix(REG&, COL&))
      CELDA$ = REPLACAR$(CELDA$, " ", "")
      CELDA$ = UCase$(CELDA$)
'      CELDA$ = "=A1+(A2-(A3/A11))"

      If Left$(CELDA$, 1) <> "=" Then
         VALORCELDA# = XVALO(CELDA$)
      Else
        If Left$(UCase$(TRIM$(CELDA$)), 5) = "=SUMA" Then
           CELDA$ = Mid$(CELDA$, 6)
        ElseIf Left$(CELDA$, 1) = "=" Then
           CELDA$ = Mid$(CELDA$, 2)
        End If
        '=suma(E3;E4;E5;E6;E7;E8)
        '=(E3+E4(E5*E6)+ (E75E8))+
        UBICAX1$ = ""
        For i% = 1 To Len(CELDA$)
          CARACTERX$ = Mid$(CELDA$, i%, 1)
          
          Select Case CARACTERX$
            Case "=" 'NO HACE NADA - NO SE ESCRIBE
            Case "(", "+", "-", "*", "/", ")", ";"
              If CARACTERX$ = ";" Then CARACTERX$ = "+"
              ULTIMO_CARACTER$ = CARACTERX$
              ACUMUCARACTERES$ = ACUMUCARACTERES$ + CARACTERX$
              If UBICAX1$ <> "" Then
               VALORINDIVIDUAL$ = UBICAX1$
               If Asc(Left$(UBICAX1$, 1)) >= 65 And Asc(Left$(UBICAX1$, 1)) <= 90 Then ' SI NO ES LETRA (NO ES UNA CELDA SI NO VALOR)
                 VALORINDIVIDUAL$ = LEOCELDA$(UBICAX1$)
               End If
                 If Left$(VALORINDIVIDUAL$, 1) = "-" Then
                    VALORINDIVIDUAL$ = REPLACAR(VALORINDIVIDUAL$, "-", "|") 'el caracter 124 reemplaza al signo - de los negativos
                 End If                                                     'para que no se detecto dos simbolos juntos y romper el funcionamiento programado
                 TEXTOFINAL$ = TEXTOFINAL$ + VALORINDIVIDUAL$
              End If
              UBICAX1$ = ""
              TEXTOFINAL$ = TEXTOFINAL$ & CARACTERX$
            Case Else
              UBICAX1$ = UBICAX1$ + CARACTERX$
              ACUMUCARACTERES$ = ""
            End Select
        Next i%
      End If
      'PARA LA ULTIMA LECTURA VA A PASAR POR ACA
      If TRIM$(UBICAX1$) <> "" Then
        VALORINDIVIDUAL$ = UBICAX1$
        If Asc(Left$(UBICAX1$, 1)) >= 65 And Asc(Left$(UBICAX1$, 1)) <= 90 Then ' SI NO ES LETRA (NO ES UNA CELDA SI NO VALOR)
           VALORINDIVIDUAL$ = LEOCELDA$(UBICAX1$)
        End If
        If Left$(VALORINDIVIDUAL$, 1) = "-" Then
         VALORINDIVIDUAL$ = REPLACAR(VALORINDIVIDUAL$, "-", "|")
        End If
      End If
      TEXTOFINAL$ = TEXTOFINAL$ + VALORINDIVIDUAL$
      'AHORA VOY CALCULANDO CONSIDERANDO LOS PARENTESIS
      'VOY BUSCANDO EL PARENTESIS MAS INTERNO Y DE AHI PARA AFUERA
        '=(E3+E4(E5*E6)+ (E75E8))
      TEXTOFINAL2$ = TEXTOFINAL$
      POS1% = InStr(TEXTOFINAL2$, ")")
      If POS1% > 0 Then
            '(10+(20*30)+40+50)
            For l& = 1 To Len(TEXTOFINAL2$)
              P1% = InStr(TEXTOFINAL2$, "(")
              If P1% > 0 Then
                 TEXTOFINAL2$ = Mid$(TEXTOFINAL2$, P1% + 1)
              Else
                 On Error Resume Next
                 P2% = InStr(TEXTOFINAL2$, ")")
                 TEXTOFINAL2$ = Mid$(TEXTOFINAL2$, 1, P2% - 1)
                 If Err > 0 Then Call COMUNI("Error de Sintaxis en Celda " & UBICAX$): Exit Function
                 On Error GoTo 0

                 BLOQUE_A_CALCULAR$ = TEXTOFINAL2$
                 BLOQUE_A_CALCULAR$ = VALIDO_NEGATIVO(BLOQUE_A_CALCULAR$)
                 VALORCELDA# = XVALO(CALCULABLOQUE(BLOQUE_A_CALCULAR$))
                 TEXTOFINAL$ = REPLACAR$(TEXTOFINAL$, "(" & TEXTOFINAL2$ & ")", TRIM$(Str$(VALORCELDA#)))
                 TEXTOFINAL2$ = TEXTOFINAL$
                 If InStr(TEXTOFINAL2$, "(") < 1 Then
                   'SI YA NO POSEE MAS PARENTESIS PUEDE SER QUE QUEDE UNA OPERACION POR ESO LO PASO DEVUELTA AL VALORCELDA
                   BLOQUE_A_CALCULAR$ = TEXTOFINAL2$
                   VALORCELDA# = XVALO(CALCULABLOQUE(BLOQUE_A_CALCULAR$))
                   Exit For
                 End If
              End If
            Next l&
            
      Else
        BLOQUE_A_CALCULAR$ = TEXTOFINAL2$
        VALORCELDA# = XVALO(CALCULABLOQUE(BLOQUE_A_CALCULAR$))
      End If
End Function
Function VALIDO_NEGATIVO(BLOQUE$)
        VALIDO_NEGATIVO = BLOQUE$
        ACUMUCARACTERES$ = ""
        If InStr(BLOQUE$, "-") < 1 Then Exit Function
        For i% = 1 To Len(BLOQUE$)
          CARACTERX$ = Mid$(BLOQUE$, i%, 1)
          TEXTX$ = ""
          Select Case CARACTERX$
            Case "(", "+", "-", "*", "/", ")", ";"
              CARACTERYSIGUIENTE$ = Mid$(BLOQUE$, i%, 2)
              Select Case CARACTERYSIGUIENTE$
                 Case "*-", "/-", "+-", "--"
                     CARACTERX$ = REPLACAR$(CARACTERX$, "*", "?@")
                     CARACTERX$ = REPLACAR$(CARACTERX$, "/", "¿@")
                     CARACTERX$ = REPLACAR$(CARACTERX$, "+", "!@")
                     CARACTERX$ = REPLACAR$(CARACTERX$, "-", "¡@")
                     i% = i% + 1
              End Select
            Case Else
            End Select
            ACUMUCARACTERES$ = ACUMUCARACTERES$ + CARACTERX$
        
        Next i%
        ACUMUCARACTERES$ = REPLACAR$(ACUMUCARACTERES$, "?@", "*|")
        ACUMUCARACTERES$ = REPLACAR$(ACUMUCARACTERES$, "¿@", "/|")
        ACUMUCARACTERES$ = REPLACAR$(ACUMUCARACTERES$, "!@", "+|")
        ACUMUCARACTERES$ = REPLACAR$(ACUMUCARACTERES$, "¡@", "-|")
        
        
        VALIDO_NEGATIVO = ACUMUCARACTERES$

End Function
 Function CANTICARACTER(TEXTO As String, LETRA As String) As Long
   Palabra = TEXTO$
   Do While Len(Palabra) > 0
      Lugar = InStr(Palabra, LETRA)
      If Lugar = 0 Then Exit Do
      Total = Total + 1
      Palabra = Mid(Palabra, Lugar + 1)
   Loop
   CANTICARACTER = Total
End Function
Function LEOCELDA$(UBICAX$)
      POS_CELDA_LEIDA$ = TRIM$(UBICAX$)
      COL& = COLUMNA&(POS_CELDA_LEIDA$, FILAX&)
      REG& = FILAX&
      On Error Resume Next
      LEOCELDA$ = TRIM$(GRID2.TextMatrix(REG&, COL&))
      On Error GoTo 0
End Function
'
Sub ESCRIBIR_RESULTADOS()
   'PRIMERO SE DEBERAN EJECUTAR LAS FUNCIONES PARAMETRIZADAS CON LAS BASES DE DATOS
   'LUEGO ESTA FUNCION VA A RECORRER TODAS LAS CELDAS DONDE HAYAN FORMULAS
   'TENGO QUE IR RECORRIENDO Y TOMAR LAS QUE YA TIENEN RESULTADOS Y NO FORMULAS
   FUNCIONES_PENDIENTES% = 0
10 For REG& = 1 To GRID2.Rows - 1
      For COL& = 1 To GRID2.Cols - 1
         HAYFUNCION% = 0
         CELDA$ = UCase$(TRIM$(GRID2.TextMatrix(REG&, COL&)))
         If CELDA$ <> "" Then
           If Left$(CELDA$, 1) = "=" Then 'C5 = =SUMA(C3:C4)
              CELDA$ = Mid$(CELDA$, 2) ' LE QUITO EL IGUAL
              UBICAX$ = UBICACIONX$(REG&, COL&)
              If Left$(CELDA$, 4) = "SUMA" Then
                 CELDA$ = Mid$(CELDA$, 5)
              End If
             'ACA LEER EL CONTENIDO DE LA FUNCION  suma(E3;E4;E5;E6;E7;E8)
             'RECORRER LAS POSISIONES E3 , E4 ETC. Y VERIFICAR SI DENTRO DE ESTAS HAY FORMULAS
              REEMPLA_CELDA$ = CELDA$
              REEMPLA_CELDA$ = REPLACAR$(REEMPLA_CELDA$, "+", ";")
              REEMPLA_CELDA$ = REPLACAR$(REEMPLA_CELDA$, "-", ";")
              REEMPLA_CELDA$ = REPLACAR$(REEMPLA_CELDA$, "*", ";")
              REEMPLA_CELDA$ = REPLACAR$(REEMPLA_CELDA$, "/", ";")
              
              STRARRAY = Split(REEMPLA_CELDA$, ";")
             'RECORRO CADA CELDA A QUE APUNTA MI FUNCION
              For i& = LBound(STRARRAY) To UBound(STRARRAY)
                CELDA_LEIDA$ = TRIM(STRARRAY(i&)) 'A6
                CELDA_LEIDA$ = REPLACAR$(CELDA_LEIDA$, "(", "")
                CELDA_LEIDA$ = REPLACAR$(CELDA_LEIDA$, ")", "")
                COL1& = COLUMNA&(CELDA_LEIDA$, FILAX&)
                REG1& = FILAX&
                On Error Resume Next
                TEXTOENCELDA$ = TRIM$(GRID2.TextMatrix(REG1&, COL1&))
                On Error GoTo 0
                If Left$(TEXTOENCELDA$, 1) = "=" Then
                   HAYFUNCION% = 1
                   FUNCIONES_PENDIENTES% = 1
                   Exit For
                End If
50            Next i&
              If HAYFUNCION% < 1 Then
                  GRID2.TextMatrix(REG&, COL&) = TRIM$(FORMATNUM$(VALORCELDA#(UBICAX$), "F16.2"))
              End If
           End If
         End If
      
      Next COL&
   Next REG&
   If FUNCIONES_PENDIENTES% > 0 Then
     FUNCIONES_PENDIENTES% = 0
     GoTo 10
   End If
End Sub

Function OBTENER_RESULTADO_SP(SPNAME$)
   If Mid$(SPNAME$, 4) = "ANTICIPOSPENDIENTES" Then
      OBTENER_RESULTADO_SP = ANTICIPOSPENDIENTES
   ElseIf Mid$(SPNAME$, 4) = "ANTICIPOSPEND_ACUMULADOS" Then
      OBTENER_RESULTADO_SP = ANTICIPOSPEND_ACUMULADOS
   ElseIf Mid$(SPNAME$, 4) = "ANTICIPOSPEND_ACUM_PREVIOS" Then
      OBTENER_RESULTADO_SP = ANTICIPOSPEND_ACUM_PREVIOS
   End If
End Function

Function COSTO_DE_SUMINISTROS(CI1%)
    CODI$ = CODEXT$(CI1%)
    Des% = FECHANUM(FECHA1)
    FF1$ = FECHATEXLAR$(Des%)
    MESHOY& = Val(Mid$(FF1$, 4, 2))
    ANIOHOY& = Val(Mid$(FF1$, 7, 4))
    IDMESANIO& = (MESHOY& * 10000) + ANIOHOY&

    SQLX$ = "SELECT costomp, TTerceros, Piezas "
    SQLX$ = SQLX$ + " FROM HISTOCOSTOS WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE IdPeriodo = " & IDMESANIO&
    SQLX$ = SQLX$ + " AND COD_INTE = " & CI1%
    SQLX$ = SQLX$ + " ORDER BY FECOSCAL DESC "
    Set RS = FLEXCONN.Execute(SQLX$)
    With RS
      COSTOMPX = XVALO(!costomp)
      TTerceroX = XVALO(!TTerceros)
      PiezaX = XVALO(!Piezas)
      
    End With
    RS.Close
    Set RS = Nothing
    
    COSTO_DE_SUMINISTROS = COSTOMPX + TTerceroX + PiezaX

End Function
Function COSTO_DE_ELABORACION(CI1%)
    Des% = FECHANUM(FECHA1)
    FF1$ = FECHATEXLAR$(Des%)
    MESHOY& = Val(Mid$(FF1$, 4, 2))
    ANIOHOY& = Val(Mid$(FF1$, 7, 4))
    IDMESANIO& = (MESHOY& * 10000) + ANIOHOY&

    SQLX$ = "SELECT ManoObra, Equipos "
    SQLX$ = SQLX$ + " FROM HISTOCOSTOS WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE IdPeriodo = " & IDMESANIO&
    SQLX$ = SQLX$ + " AND COD_INTE = " & CI1%
    SQLX$ = SQLX$ + " ORDER BY FECOSCAL DESC "
    Set RS = FLEXCONN.Execute(SQLX$)
    With RS
      MOBRAX = XVALO(!ManoObra)
      EQUIPOX = XVALO(!Equipos)
    End With
    RS.Close
    Set RS = Nothing
    CD$ = CODEXT$(CI1%)
    COSTO_DE_ELABORACION = MOBRAX + EQUIPOX

End Function

Function APLICACIONES_DE_FACTU_ANTIC_MES_ACTIVO()
    'LO QUE SE APLICO DENTRO DEL MES ACTIVO COMO FACTURA ANTICIPO
    APLICACIONES_DE_FACTU_ANTIC_MES_ACTIVO = 0
    Desde = FETEXCOR1$(FECHANUM(FECHA1))
    Hasta = FETEXCOR1$(FECHANUM(Fecha2))

    SQLX$ = " Select SUM(IMPOHABE) AS IHABE from factuanti with(nolock)"
    SQLX$ = SQLX$ + " where nuclien > 0 "
    SQLX$ = SQLX$ + " and tipomovim = 'VENTA'"
    SQLX$ = SQLX$ + " AND FECHEMISI >= '" & Desde & "'"
    SQLX$ = SQLX$ + " AND FECHEMISI <= '" & Hasta & "'"
    SQLX$ = SQLX$ + " GROUP BY tipomovim"
    '
    Set RS = READSET(SQLX$)
    j& = 0
    With RS
      If Not .EOF Then
           APLICACIONES_DE_FACTU_ANTIC_MES_ACTIVO = XVALO(!IHABE)
      End If
    End With
    RS.Close
    Set RS = Nothing

End Function

