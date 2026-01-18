VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form CONCIMAN07 
   BackColor       =   &H0095C5E1&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Situación de Conciliación "
   ClientHeight    =   8940
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8940
   ScaleWidth      =   13215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command9 
      Caption         =   "Selección Multiple"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Left            =   7080
      TabIndex        =   29
      ToolTipText     =   "Permite Seleccionar Varios Movimientos y Conciliar con 1 de los Pendientes del Sistema"
      Top             =   2880
      Width           =   1800
   End
   Begin VB.PictureBox Picture999 
      Height          =   330
      Index           =   0
      Left            =   240
      ScaleHeight     =   270
      ScaleWidth      =   1110
      TabIndex        =   26
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CommandButton Command8 
      Height          =   280
      Index           =   2
      Left            =   6405
      Picture         =   "CONCIMAN07.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Búsqueda por Texto "
      Top             =   5700
      Width           =   650
   End
   Begin VB.CommandButton Command8 
      Height          =   280
      Index           =   1
      Left            =   6405
      Picture         =   "CONCIMAN07.frx":1366
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Búsqueda por Texto "
      Top             =   2850
      Width           =   650
   End
   Begin VB.CommandButton Command8 
      Height          =   280
      Index           =   0
      Left            =   6405
      Picture         =   "CONCIMAN07.frx":26CC
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Búsqueda por Texto "
      Top             =   50
      Width           =   650
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Selección Multiple"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Left            =   7080
      TabIndex        =   19
      ToolTipText     =   "Permite Seleccionar Varios Movimientos y Conciliar con 1 de los Pendientes del Sistema"
      Top             =   5700
      Width           =   1800
   End
   Begin VB.TextBox Text16 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C7D9CC&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   6
      Left            =   10455
      Locked          =   -1  'True
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   8610
      Width           =   1785
   End
   Begin VB.TextBox Text16 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00CDDADA&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   5
      Left            =   10335
      Locked          =   -1  'True
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   5700
      Width           =   1905
   End
   Begin VB.TextBox Text16 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E2D3C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   4
      Left            =   10215
      Locked          =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   2850
      Width           =   2025
   End
   Begin VB.ListBox List1 
      Height          =   450
      Left            =   10605
      Sorted          =   -1  'True
      TabIndex        =   14
      Top             =   2940
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Conciliar Gastos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1935
      TabIndex        =   13
      ToolTipText     =   "Conciliar Gasto Seleccionado de Grilla 2"
      Top             =   8595
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Eliminar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   255
      TabIndex        =   12
      ToolTipText     =   "Quita Registro Seleccionado en Grilla 1 y lo Ubica en Grilla 2 y 3 Según Corresponda"
      Top             =   8595
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Concilia"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3615
      TabIndex        =   11
      ToolTipText     =   "Concilar Item Seleccionados de Grilla 2 y 3"
      Top             =   8595
      Width           =   1575
   End
   Begin VB.TextBox Text16 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   70
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   8520
      Width           =   12180
   End
   Begin VB.TextBox Text16 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C7D9CC&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   2
      Left            =   70
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Text            =   "MOVIMIENTOS DEL BANCO (en el sistema) SIN CONCILIAR"
      Top             =   5720
      Width           =   6250
   End
   Begin VB.TextBox Text16 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00CDDADA&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   1
      Left            =   70
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Text            =   "LISTA DE MOVIMIENTOS NO CONCILIADOS AUTOMATICAMENTE"
      Top             =   2870
      Width           =   6250
   End
   Begin VB.TextBox Text16 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E2D3C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   0
      Left            =   70
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Text            =   "LISTA DE MOVIMIENTOS CONCICLIADOS AUTOMATICAMENTE"
      Top             =   50
      Width           =   6250
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   495
      Left            =   10440
      TabIndex        =   2
      Top             =   4920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H002C3D46&
      FillColor       =   &H00400000&
      ForeColor       =   &H80000008&
      Height          =   8955
      Left            =   12360
      ScaleHeight     =   8925
      ScaleWidth      =   1785
      TabIndex        =   0
      Top             =   0
      Width           =   1815
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H002B9973&
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   110
         ScaleHeight     =   165
         ScaleWidth      =   615
         TabIndex        =   27
         Top             =   8160
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   28
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command6 
         Height          =   560
         Left            =   120
         Picture         =   "CONCIMAN07.frx":3A32
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Conciliación Definitiva "
         Top             =   7440
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   120
         Picture         =   "CONCIMAN07.frx":3D3C
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Terminar - Salir de la Aplicación"
         Top             =   105
         Width           =   615
      End
      Begin VB.CommandButton AGenRep 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   120
         Picture         =   "CONCIMAN07.frx":3E86
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Acceso a Consultas Pre-Configuradas"
         Top             =   1250
         Width           =   615
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSF1 
      Height          =   2505
      Left            =   75
      TabIndex        =   3
      ToolTipText     =   "Click Para Seleccionar/Doble Click Para Ver Documento"
      Top             =   315
      Width           =   12180
      _ExtentX        =   21484
      _ExtentY        =   4419
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   14865344
      BackColorBkg    =   16777215
      AllowUserResizing=   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid MSF2 
      Height          =   2505
      Left            =   75
      TabIndex        =   5
      ToolTipText     =   "Click Para Seleccionar/Doble Click Para Ver Documento"
      Top             =   3150
      Width           =   12180
      _ExtentX        =   21484
      _ExtentY        =   4419
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   13490906
      BackColorBkg    =   16777215
      AllowUserResizing=   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid MSF3 
      Height          =   2505
      Left            =   75
      TabIndex        =   6
      ToolTipText     =   "Click Para Seleccionar/Doble Click Para Ver Documento"
      Top             =   6000
      Width           =   12180
      _ExtentX        =   21484
      _ExtentY        =   4419
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   13097420
      BackColorBkg    =   16777215
      AllowUserResizing=   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "CONCIMAN07.frx":4190
      Top             =   1680
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   60
      OleObjectBlob   =   "CONCIMAN07.frx":43C4
      TabIndex        =   25
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label1 
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
      Left            =   -840
      TabIndex        =   4
      Top             =   960
      Width           =   3480
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      Height          =   195
      Left            =   720
      TabIndex        =   1
      Top             =   16120
      Width           =   870
   End
End
Attribute VB_Name = "CONCIMAN07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public CODICONCI$
Public REGSEL1&
Public REGSEL2&
Public REGSEL3&
Dim ARBUS$, ABUSQUE$

Function CUENTAGASTO%(NUORD%)
   CUENTAGASTO% = XVALO(VALOBADA("CONCILIA", "CUECONTAI", "CODICONCI = '" & CODICONCI$ & "'ORDCONCILIA = " & NUORD%, "NOMESS"))
End Function
Sub ELIMINAR(REG1&, Optional ENTRA%)
   '------------------------------------------------------------------
   'TRASLADA LA INFORMACION SELECCIONADA A LA GRILLA 2
   With MSF1
        J& = msf2.Rows
        msf2.Rows = J& + 1
        msf2.TextMatrix(J&, 1) = ""
        msf2.TextMatrix(J&, 2) = ""
        msf2.TextMatrix(J&, 3) = .TextMatrix(REG1&, 3)
        msf2.TextMatrix(J&, 4) = .TextMatrix(REG1&, 4)
            OrdenConcilia% = XVALO(.TextMatrix(REG1&, 7))
            ImporteDebe# = XVALO(VALOBADA("CONCILIA", "DEBE", "ORDCONCILIA=" & CStr(OrdenConcilia%)))
            ImporteHaber# = XVALO(VALOBADA("CONCILIA", "HABER", "ORDCONCILIA=" & CStr(OrdenConcilia%)))
        msf2.TextMatrix(J&, 5) = TRIM$(FORMATNUM$(ImporteHaber#, "F12.2"))
        msf2.TextMatrix(J&, 6) = TRIM$(FORMATNUM$(ImporteDebe#, "F12.2"))
        msf2.TextMatrix(J&, 7) = .TextMatrix(REG1&, 7)
        msf2.TextMatrix(J&, 8) = .TextMatrix(REG1&, 8)
        msf2.TextMatrix(J&, 9) = .TextMatrix(REG1&, 9)
        
        '
        If TRIM$(UCase$(.TextMatrix(REG1&, 2))) = "GASTOS" Then Exit Sub 'BORRA EL COMPROBANTE GASTOS
        If ENTRA% > 0 Then Exit Sub
        '--------------------------------------------------------------
        'AGREGO A LA LISTA 3
        J& = MSF3.Rows
        MSF3.Rows = J& + 1
        MSF3.TextMatrix(J&, 1) = .TextMatrix(REG1&, 1)
        MSF3.TextMatrix(J&, 2) = .TextMatrix(REG1&, 2)
        MSF3.TextMatrix(J&, 3) = .TextMatrix(REG1&, 3)
           CUEII% = CUENTA_BCO_CONCILIA%(CODICONCI$)
           CONDI$ = "CODICOM_COR = '" & TRIM$(MSF1.TextMatrix(REG1&, 2)) & "' AND CUECONTAI = " & CUEII%
           CONCEP$ = Left$(VALOBADA("CAJABANC", "DESCRIMOV", CONDI$, "NOMESS"), 28)
        MSF3.TextMatrix(J&, 4) = CONCEP$ 'CONCEPTO DEL SISTEMA (DET.REF.)
        MSF3.TextMatrix(J&, 5) = .TextMatrix(REG1&, 5)
        MSF3.TextMatrix(J&, 6) = .TextMatrix(REG1&, 6)
        MSF3.TextMatrix(J&, 7) = "" 'NUNERO DE ORDEN VACIO
        MSF3.TextMatrix(J&, 8) = .TextMatrix(REG1&, 8)
        MSF3.TextMatrix(J&, 9) = .TextMatrix(REG1&, 9)
        '
    End With

End Sub

Function FECHADECONCILIA%(NUORD%)
   FECHADECONCILIA% = CVINT(VALOBADA("CONCILIA", "FECHA", "CODICONCI = '" & CODICONCI$ & "' AND ORDCONCILIA = " & NUORD%, "NOMESS"))
End Function

Sub sortflex(MSF As MSFlexGrid)
  With MSF
    .Redraw = False
    List1.Clear
    For i& = 1 To .Rows - 1
        T1$ = TRIM$(SAFETEXT$(CVINT(.TextMatrix(i&, 3)))) 'INVIERTO PARA QUE ORDENE POR FECHA DEL EXTRACTO/ACREEDITACION
        T2$ = .TextMatrix(i&, 2)
        T3$ = TRIM$(SAFETEXT$(CVINT(.TextMatrix(i&, 1))))
        T4$ = .TextMatrix(i&, 4)
        T5$ = .TextMatrix(i&, 5)
        T6$ = .TextMatrix(i&, 6)
        T7$ = .TextMatrix(i&, 7)
        T8$ = .TextMatrix(i&, 8)
        T9$ = .TextMatrix(i&, 9)
        On Error Resume Next
        T10$ = .TextMatrix(i&, 10)
        On Error GoTo 0
        '
        Call REPLA(TX$, T1$, 1, 10)
        Call REPLA(TX$, T2$, 11, 20)
        Call REPLA(TX$, T3$, 31, 10)
        Call REPLA(TX$, T4$, 41, 70)
        Call REPLA(TX$, T5$, 111, 20)
        Call REPLA(TX$, T6$, 131, 20)
        Call REPLA(TX$, T7$, 151, 10)
        Call REPLA(TX$, T8$, 161, 10)
        Call REPLA(TX$, T9$, 171, 10)
        Call REPLA(TX$, T10$, 181, 10)
        '
        List1.AddItem TX$
    Next i&
    .Rows = 1
    For i& = 1 To List1.ListCount
     TTXX$ = List1.LIST(i& - 1)
     T3$ = Mid$(TTXX$, 1, 10)
     T2$ = Mid$(TTXX$, 11, 20)
     T1$ = Mid$(TTXX$, 31, 10)
     T4$ = Mid$(TTXX$, 41, 70)
     T5$ = Mid$(TTXX$, 111, 20)
     T6$ = Mid$(TTXX$, 131, 20)
     T7$ = Mid$(TTXX$, 151, 10)
     T8$ = Mid$(TTXX$, 161, 10)
     T9$ = Mid$(TTXX$, 171, 10)
     T10$ = Mid$(TTXX$, 181, 10)
     '
     REG& = REG& + 1
     .Rows = REG& + 1
        FECH_EXTR% = XVALO(T1$)
        COMPRBTE$ = T2$
        FECONTA% = XVALO(T3$)
        CONCEP$ = T4$
        IMPODEBEX# = XVALO(T5$)
        IMPOHABEX# = XVALO(T6$)
        NUORD% = XVALO(T7$)
        DETABLA$ = T8$
        CUEII% = XVALO(T9$)
        '
        .TextMatrix(REG&, 1) = FECHATEX$(FECH_EXTR%) 'FECHA DEL COMPROBANTE
        .TextMatrix(REG&, 2) = COMPRBTE$
        .TextMatrix(REG&, 3) = FECHATEX$(FECONTA%)
        .TextMatrix(REG&, 4) = CONCEP$
        .TextMatrix(REG&, 5) = TRIM$(FORMATNUM$(IMPODEBEX#, "F12.2"))
        .TextMatrix(REG&, 6) = TRIM$(FORMATNUM$(IMPOHABEX#, "F12.2"))
        .TextMatrix(REG&, 7) = NUORD%
        .TextMatrix(REG&, 8) = DETABLA$
        .TextMatrix(REG&, 9) = SAFETEXT$(CUEII%)
        On Error Resume Next
        .TextMatrix(REG&, 10) = T10$
        On Error GoTo 0

    Next i&
    MSF.Redraw = True
  End With
End Sub
Sub MULTISELEC(MSF As MSFlexGrid)
  With MSF
    .Redraw = False
    List2.Clear
    For i& = 1 To .Rows - 1
        T1$ = TRIM$(SAFETEXT$(CVINT(.TextMatrix(i&, 3))))
        T2$ = .TextMatrix(i&, 2)
        T3$ = TRIM$(SAFETEXT$(CVINT(.TextMatrix(i&, 1))))
        T4$ = .TextMatrix(i&, 4)
        T5$ = .TextMatrix(i&, 5)
        T6$ = .TextMatrix(i&, 6)
        T7$ = .TextMatrix(i&, 7)
        T8$ = .TextMatrix(i&, 8)
        T9$ = .TextMatrix(i&, 9)
        '
        Call REPLA(TX$, T1$, 1, 10)
        Call REPLA(TX$, T2$, 11, 20)
        Call REPLA(TX$, T3$, 31, 10)
        Call REPLA(TX$, T4$, 41, 70)
        Call REPLA(TX$, T5$, 111, 20)
        Call REPLA(TX$, T6$, 131, 20)
        Call REPLA(TX$, T7$, 151, 10)
        Call REPLA(TX$, T8$, 161, 10)
        Call REPLA(TX$, T9$, 171, 10)
        '
        List1.AddItem TX$
    Next i&
    .Rows = 1
    For i& = 1 To List1.ListCount
     TTXX$ = List1.LIST(i& - 1)
     T3$ = Mid$(TTXX$, 1, 10)
     T2$ = Mid$(TTXX$, 11, 20)
     T1$ = Mid$(TTXX$, 31, 10)
     T4$ = Mid$(TTXX$, 41, 70)
     T5$ = Mid$(TTXX$, 111, 20)
     T6$ = Mid$(TTXX$, 131, 20)
     T7$ = Mid$(TTXX$, 151, 10)
     T8$ = Mid$(TTXX$, 161, 10)
     T9$ = Mid$(TTXX$, 171, 10)
     '
     REG& = REG& + 1
     .Rows = REG& + 1
        FECH_EXTR% = XVALO(T1$)
        COMPRBTE$ = T2$
        FECONTA% = XVALO(T3$)
        CONCEP$ = T4$
        IMPODEBEX# = XVALO(T5$)
        IMPOHABEX# = XVALO(T6$)
        NUORD% = XVALO(T7$)
        DETABLA$ = T8$
        CUEII% = XVALO(T9$)
        '
        .TextMatrix(REG&, 1) = FECHATEX$(FECH_EXTR%) 'FECHA DEL COMPROBANTE
        .TextMatrix(REG&, 2) = COMPRBTE$
        .TextMatrix(REG&, 3) = FECHATEX$(FECONTA%)
        .TextMatrix(REG&, 4) = CONCEP$
        .TextMatrix(REG&, 5) = TRIM$(FORMATNUM$(IMPODEBEX#, "F12.2"))
        .TextMatrix(REG&, 6) = TRIM$(FORMATNUM$(IMPOHABEX#, "F12.2"))
        .TextMatrix(REG&, 7) = NUORD%
        .TextMatrix(REG&, 8) = DETABLA$
        .TextMatrix(REG&, 9) = ""

    Next i&
    MSF.Redraw = True
  End With
End Sub


Sub SUMAMSF(MSF As MSFlexGrid)
  With MSF
    .Redraw = False
    ACUMULO# = 0
    For i& = 1 To .Rows - 1
        DEBEX# = XVALO(.TextMatrix(i&, 5))
        HABEX# = XVALO(.TextMatrix(i&, 6))
        SALDO# = DEBEX# - HABEX#
        ACUMULO# = ACUMULO# + SALDO#
    Next i&
    .Redraw = True
    'TOMO EL VALOR DE LA ULTIMA LETRA DEL MSF1, MSF2 O MSF3
    IND% = 3 + XVALO(Right(.Name, 1))
    Text16(IND%) = FORMATNUM$(ACUMULO#, "F14.2")
 End With
End Sub

Private Sub AGenRep_Click()

    '
    AGenRep.Enabled = False
    Call ACONSUL_PRE(TRIM$(UCase$(Name)))
    AGenRep.Enabled = True
    '
End Sub

Private Sub AGenRep_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   '
   If Button = 2 Then
      Call CONECTA_CONSUL_PRE(TRIM$(UCase$(Name)))
   End If
   '
End Sub



Private Sub Command1_Click()
  Unload Me
End Sub

Private Sub Command2_Click()
    Command2.Enabled = False
    '
    'RECORRE LA LISTA 2 Y VALIDO QUE HAYA SELECCION
    HAY_SELECCION_2% = 0
    
    With msf2
      For i& = 1 To .Rows - 1
          If .TextMatrix(i&, 0) = "*" Then
             HAY_SELECCION_2% = 1
             REG2& = i&
           Exit For
          End If
        Next i&
    End With
    '
    'VALIDO QUE HAYA UNO SELECCIONADO EN LISTA 2
    If HAY_SELECCION_2% < 1 Then
       Call COMUNI("Falta Seleccionar Items en Grilla 2")
       GoTo 99
    End If
    '
    If msf2.TextMatrix(REG2&, 2) <> "GASTOS" Then
        'VARIABLE CODICONCI DE ALCANCE PUBLICO A NIVEL DE FRM
        ESGASTO% = VALIDA_SI_ES_GASTO%(CODICONCI$, msf2.TextMatrix(REG2&, 4))
        If ESGASTO% < 1 Then
          OPX% = COMALTER%("ATENCION: Registro Seleccionado No Corresponde A Gastos.\\Cancelar\Ver Tabla de Gastos")
          If OPX% <> 2 Then GoTo 99
          Call OPCHEQ07.CARGA_TABLA_GASEXTRA
          GoTo 99
        End If
        
    End If
    
    'SUBO LA INFO A LISTA 1
    NUORD% = XVALO(msf2.TextMatrix(REG2&, 7))

    With msf2
           J& = MSF1.Rows
           MSF1.Rows = J& + 1
           MSF1.TextMatrix(J&, 1) = .TextMatrix(REG2&, 1)
           MSF1.TextMatrix(J&, 2) = "GASTOS"
           MSF1.TextMatrix(J&, 3) = FECHATEX$(FECHADECONCILIA%(NUORD%))
           MSF1.TextMatrix(J&, 4) = .TextMatrix(REG2&, 4)
           MSF1.TextMatrix(J&, 5) = .TextMatrix(REG2&, 5)
           MSF1.TextMatrix(J&, 6) = .TextMatrix(REG2&, 6)
           MSF1.TextMatrix(J&, 7) = .TextMatrix(REG2&, 7)
           MSF1.TextMatrix(J&, 8) = .TextMatrix(REG2&, 8)
           MSF1.TextMatrix(J&, 9) = SAFETEXT$(CUENTAGASTO%(NUORD%))
           Call COLOREAR_GRILLASQL_CON_MARCA(CONCIMAN07.msf2, &HC0E0FF, REG2&, "*", 1)
    End With
    '
    NU0RD% = XVALO(msf2.TextMatrix(REG2&, 7))
    Call ACTUREGISTRO("CONCILIA", "COMPROBANTE", "ORDCONCILIA  = " & NU0RD%, "GASTOS", REGAF&, "NOMESS")

    'BORRO LA FILA DE LISTA 2
    Call borrafila(msf2, REG2&)
    '
    'BORRO LA FILA DE LISTA 3
    Call COLOREAR_GRILLASQL_CON_MARCA(CONCIMAN07.MSF3, &HCBEC86, REG3&, "*", 1)
'    Call borrafila(MSF3, REG3&)
    '
    Call sortflex(MSF1)
    Call sortflex(msf2)
    Call SUMAMSF(MSF1)
    Call SUMAMSF(msf2)

    'llevo al final la grilla 1
    TPXI& = MSF1.Rows - 9: If TPXI& < 1 Then TPXI& = 1
    On Error Resume Next
    MSF1.TopRow = TPXI&
    On Error GoTo 0

99  Command2.Enabled = True

End Sub



Private Sub Command24_Click()

End Sub

Private Sub Command3_Click()

   Command3.Enabled = False
    'RECORRE LA LISTA 2 Y VALIDO QUE HAYA SELECCION
    HAY_SELECCION_1% = 0
    With MSF1
      For i& = 1 To .Rows - 1
          If .TextMatrix(i&, 0) = "*" Then
             HAY_SELECCION_1% = 1
             REG1& = i&
           Exit For
          End If
        Next i&
    End With
    '
    'VALIDO QUE HAYA UNO SELECCIONADO EN LISTA 2
    If HAY_SELECCION_1% < 1 Then
       Call COMUNI("Falta Seleccionar Items en Grilla 1")
       GoTo 99
    End If
   '
   'RECORRO LA GRILLA 1 Y ELIMINA LOS COMPROBANTES IGUALES
   COMPROB$ = TRIM$(MSF1.TextMatrix(REG1&, 2))
   OrdenConcilia& = XVALO(MSF1.TextMatrix(REG1&, 7))
   
   ORDENVINCULO& = XVALO(MSF1.TextMatrix(REG1&, 10))
   If ORDENVINCULO& > 0 Then
        For i& = 1 To MSF1.Rows - 1
            If MSF1.TextMatrix(i&, 7) = ORDENVINCULO& Then
                MSF1.Row = i&
                Call COLOREAR_GRILLASQL_CON_MARCA(CONCIMAN07.MSF1, &HFFFFC0, i&, "*", 1, REGSEL1&)
                REGSEL1& = i&
                OrdenConcilia& = ORDENVINCULO&
                COMPROB$ = TRIM$(MSF1.TextMatrix(i&, 2))
                
                Exit For
            End If
        Next i&
        
        RESPUESTA% = COMALTER%("Esta por Anular una Conciliación Múltiple.\Está Seguro que Desea Continuar?\\No, Cancelar\Si, Continuar")
        If RESPUESTA% = 1 Then GoTo 99
   End If
   
   Cuenta% = 0
   If COMPROB$ <> "GASTOS" Then
         'SI NO ES UN GASTO
         'RECORRE Y VA BAJANDO CADA REGISTRO QUE MACHEE CON EL NUMERO DE COMPROBANTE
         For i& = 1 To MSF1.Rows - 1
               ORDENVINCULO& = XVALO(MSF1.TextMatrix(i&, 10))
               If COMPROB$ <> "" And COMPROB$ = TRIM$(MSF1.TextMatrix(i&, 2)) Then
                 Call ELIMINAR(i&, Cuenta%)
                 IMPOACUMULA# = IMPOACUMULA# + (XVALO(MSF1.TextMatrix(i&, 5)) - XVALO(MSF1.TextMatrix(i&, 6)))
                 Cuenta% = 1
               End If
               
               'EN LAS CONCILIACIONES MULTIPLES (1 BANCO A VARIOS CONTABLES), LOS ITEMS NUEVOS
               If ORDENVINCULO& = OrdenConcilia& Then
                    With MSF1
                        J& = MSF3.Rows
                        MSF3.Rows = J& + 1
                        MSF3.TextMatrix(J&, 1) = .TextMatrix(i&, 1)
                        MSF3.TextMatrix(J&, 2) = .TextMatrix(i&, 2)
                        MSF3.TextMatrix(J&, 3) = .TextMatrix(i&, 3)
                           CUEII% = CUENTA_BCO_CONCILIA%(CODICONCI$)
                           CONDI$ = "CODICOM_COR = '" & TRIM$(MSF1.TextMatrix(i&, 2)) & "' AND CUECONTAI = " & CUEII%
                           CONCEP$ = Left$(VALOBADA("CAJABANC", "DESCRIMOV", CONDI$, "NOMESS"), 28)
                        MSF3.TextMatrix(J&, 4) = CONCEP$ 'CONCEPTO DEL SISTEMA (DET.REF.)
                        MSF3.TextMatrix(J&, 5) = .TextMatrix(i&, 5)
                        MSF3.TextMatrix(J&, 6) = .TextMatrix(i&, 6)
                        MSF3.TextMatrix(J&, 7) = "" 'NUNERO DE ORDEN VACIO
                        MSF3.TextMatrix(J&, 8) = .TextMatrix(i&, 8)
                        MSF3.TextMatrix(J&, 9) = .TextMatrix(i&, 9)
                        
                        'IMPOACUMULA# = IMPOACUMULA# + (XVALO(MSF1.TextMatrix(i&, 5)) - XVALO(MSF1.TextMatrix(i&, 6)))
                    End With
               End If
         Next i&
         '
         'RECORRE TODOS LOS ITEMS NUEVAMENTE Y BORRA LAS FILAS QUE SE HAN BAJADO
         For i& = 1 To MSF1.Rows - 1
            ORDEN& = XVALO(MSF1.TextMatrix(i&, 7))
            ORDENVINCULO& = XVALO(MSF1.TextMatrix(i&, 10))
            If REG1& = i& Then
              'CON ESTO BLANQUEA LA FILA PINTADA
              Call COLOREAR_GRILLASQL_CON_MARCA(CONCIMAN07.MSF1, &HFFFFC0, REG1&, "*", 1)
            End If
            'BORRA LOS DATOS DE COMPROBANTE Y FECHA DEL CONCILIA Y
            'BORRA LA FILA DE LA GRILLA 1
            If COMPROB$ <> "" And COMPROB$ = TRIM$(MSF1.TextMatrix(i&, 2)) Then
               Call BORRA_COMPROBANTE_EN_CONCILIA(ORDEN&)
               Call borrafila(MSF1, i&)
               i& = i& - 1
            End If
            'ELIMINA LOS MOVIMIENTOS CREADOS A PARTIR DE LA CONCILIACION MULTIPLE
            If OrdenConcilia& = ORDENVINCULO& Then
                Call BORRAREGISTROS("CONCILIA", "ORDCONCILIA=" & CStr(ORDEN&), REGAF&)
                Call borrafila(MSF1, i&)
                i& = i& - 1
            End If
            
            If i& = MSF1.Rows - 1 Then Exit For
         Next i&
         
   'CASO QUE SEAN GASTOS SOLO BORRA EL SELECCIONADO.
   ElseIf COMPROB$ = "GASTOS" Then
         Call COLOREAR_GRILLASQL_CON_MARCA(CONCIMAN07.MSF1, &HFFFFC0, REG1&, "*", 1)
         Call ELIMINAR(REG1&)
         ORDEN& = XVALO(MSF1.TextMatrix(REG1&, 7))
         Call BORRA_COMPROBANTE_EN_CONCILIA(ORDEN&)
         Call borrafila(MSF1, REG1&)
   End If
   '
   If IMPOACUMULA# > 0 Then
      MSF3.TextMatrix(MSF3.Rows - 1, 5) = FORMATNUM$(Abs(IMPOACUMULA#), "F12.2")
   Else
      MSF3.TextMatrix(MSF3.Rows - 1, 6) = FORMATNUM$(Abs(IMPOACUMULA#), "F12.2")
   End If
   
   Call ActualizaPendientesContables
   
''   '------------------------------------------------------------------
''   'TRASLADA LA INFORMACION SELECCIONADA A LA GRILLA 2
''   With MSF1
''        J& = MSF2.Rows
''        MSF2.Rows = J& + 1
''        MSF2.TextMatrix(J&, 1) = .TextMatrix(REG1&, 1)
''        MSF2.TextMatrix(J&, 2) = ""
''        MSF2.TextMatrix(J&, 3) = ""
''        MSF2.TextMatrix(J&, 4) = .TextMatrix(REG1&, 4)
''        MSF2.TextMatrix(J&, 5) = .TextMatrix(REG1&, 5)
''        MSF2.TextMatrix(J&, 6) = .TextMatrix(REG1&, 6)
''        MSF2.TextMatrix(J&, 7) = .TextMatrix(REG1&, 7)
''        MSF2.TextMatrix(J&, 8) = .TextMatrix(REG1&, 8)
''        MSF2.TextMatrix(J&, 9) = .TextMatrix(REG1&, 9)
''
''        '
''        If TRIM$(UCase$(.TextMatrix(REG1&, 2))) = "GASTOS" Then GoTo 70 'BORRA EL COMPROBANTE GASTOS
''        '--------------------------------------------------------------
''        'AGREGO A LA LISTA 3
''        J& = MSF3.Rows
''        MSF3.Rows = J& + 1
''        MSF3.TextMatrix(J&, 1) = .TextMatrix(REG1&, 1)
''        MSF3.TextMatrix(J&, 2) = .TextMatrix(REG1&, 2)
''        MSF3.TextMatrix(J&, 3) = .TextMatrix(REG1&, 3)
''           NUORD& = XVALO(.TextMatrix(REG1&, 7))
''           CONCEP$ = Left$(VALOBADA("CONCILIA", "CONCEPTO_SISTEMA", "ORDCONCILIA = " & NUORD&), 26)
''        MSF3.TextMatrix(J&, 4) = CONCEP$ 'CONCEPTO DEL SISTEMA (DET.REF.)
''        MSF3.TextMatrix(J&, 5) = .TextMatrix(REG1&, 5)
''        MSF3.TextMatrix(J&, 6) = .TextMatrix(REG1&, 6)
''        MSF3.TextMatrix(J&, 7) = "" 'NUNERO DE ORDEN VACIO
''        MSF3.TextMatrix(J&, 8) = .TextMatrix(REG1&, 8)
''        MSF3.TextMatrix(J&, 9) = .TextMatrix(REG1&, 9)
''        '
        'CON ESTO BLANQUEA LA FILA PINTADA
'        Call COLOREAR_GRILLASQL_CON_MARCA(CONCIMAN07.MSF1, &HFFFFC0, REG1&, "*", 1)
'        ORDEN& = XVALO(MSF1.TextMatrix(REG1&, 7))
''        'BORRO EN EL CONCILIA EL COMPROBANTE Y LA FECHA
'        Call BORRA_COMPROBANTE_EN_CONCILIA(ORDEN&)
'        'BORRO LA FILA DE LISTA
''
''    End With
99  Command3.Enabled = True
    Call sortflex(msf2)
    Call sortflex(MSF3)
    Call SUMAMSF(MSF1)
    Call SUMAMSF(msf2)
    Call SUMAMSF(MSF3)

End Sub

Private Sub ActualizaPendientesContables()

    CUEII% = CUENTA_BCO_CONCILIA%(CODICONCI$)
    DES% = COMEJ%
    Call GESVAL07.GENERA_PENDIENTE_EN_LIMAYO01(CUEII%, DES%, 0)
    '
    'RECORRO EL ARCHIVO LIMAYO01 CON LOS PENDIENTES DEL SISTEMA PARA MOSTRARLOS EN LA GRILLA 3
    With CONCIMAN07.MSF3
        .Redraw = False
        Z$ = Space$(128)
        .Rows = 1
        k& = 0
        For U& = 1 To ULTREG&("!LIMAYOA1")
         X$ = REGLEIDO$("!LIMAYOA1", U&)
         Z$ = Space$(128)
           FEX = CVI(Left$(X$, 2))
           FECONTA1 = FECHATEX$(FEX)
           COCOM$ = TRIM$(Mid$(X$, 3, 12))
           FEX = CVI(Mid$(X$, 15, 2))
           FEACREDIT$ = FECHATEX$(FEX)
           DESCRIMOVIMIENTO$ = TRIM$(Mid$(X$, 17, 32))
           IDEBE# = CVD(Mid$(X$, 49, 8))
           IHABE# = CVD(Mid$(X$, 57, 8))
           TABLA$ = TRIM$(Mid$(X$, 73, 8))
           
           k& = k& + 1
           .Rows = k& + 1
           .TextMatrix(k&, 1) = FECONTA1
           .TextMatrix(k&, 2) = COCOM$
           .TextMatrix(k&, 3) = FEACREDIT$
           .TextMatrix(k&, 4) = DESCRIMOVIMIENTO$
           If IMPODEBEX# <> 0 Then IMPODEBE_PTO = TRIM$(FORMATNUM$(IMPODEBEX#, "F16.5"))
           If IMPOHABEX# <> 0 Then IMPOHABE_PTO = TRIM$(FORMATNUM$(IMPOHABEX#, "F16.5"))
           If IDEBE# > 0 Then
               .TextMatrix(k&, 5) = FORMATNUM$(IDEBE#, "F14.2")
           Else
               .TextMatrix(k&, 6) = FORMATNUM$(IHABE#, "F14.2")
           End If

          .TextMatrix(k&, 7) = ""
          .TextMatrix(k&, 8) = TABLA$
          .TextMatrix(k&, 9) = ""

50     Next U&
       .Redraw = True
   End With
   
   Call CIERRARCH("!LIMAYOA1")
   '
   'ELIMINAR DE PENDIENTES LAS FILAS CONCILIADAS
   Call GESVAL07.DESPEJA_PENDIENTES
   
End Sub

Sub BORRA_COMPROBANTE_EN_CONCILIA(ORDEN&)
       CONDI$ = "ORDCONCILIA = " & ORDEN&
       CONDI$ = CONDI$ + " AND CODICONCI = '" & CODICONCI$ & "'"
       Call ACTUREGISTRO("CONCILIA", "COMPROBANTE", CONDI$, "", REGAF&, "NOMESS")
       Call ACTUREGISTRO("CONCILIA", "FECHACOMPROBANTE", CONDI$, "", REGAF&, "NOMESS")
End Sub


Function borrafila(MSF As MSFlexGrid, REG&)

        For i& = REG& To MSF.Rows - 2
          For J& = 1 To MSF.Cols - 1
            MSF.TextMatrix(i, J) = MSF.TextMatrix(i + 1, J)
          Next
        Next
        MSF.Rows = MSF.Rows - 1
        
End Function

Private Sub Command5_Click()
    Command5.Enabled = False
    '
    'RECORRE LA LISTA 2 Y VALIDO QUE HAYA SELECCION
    HAY_SELECCION_2% = 0
    
    With msf2
      For i& = 1 To .Rows - 1
          If .TextMatrix(i&, 0) = "*" Then
             HAY_SELECCION_2% = 1
             REG2& = i&
           Exit For
          End If
        Next i&
    End With
    '
    'VALIDO QUE HAYA UNO SELECCIONADO EN LISTA 2
    If HAY_SELECCION_2% < 1 Then
       Call COMUNI("Falta Seleccionar Items en Grilla 2")
       GoTo 99
    End If
    '
    'VALIDO QUE HAYA SELECCIONADO EN LISTA 3
    HAY_SELECCION_3% = 0
    With MSF3
      For i& = 1 To .Rows - 1
          If .TextMatrix(i&, 0) = "*" Then
             HAY_SELECCION_3% = 1
             REG3& = i&
           Exit For
          End If
        Next i&
    End With
    '
    'VALIDO QUE HAYA UNO SELECCIONADO EN LISTA 3
    If HAY_SELECCION_3% < 1 Then
       Call COMUNI("Falta Seleccionar Items en Grilla 3")
       GoTo 99
    End If
    '
    'VALIDO QUE NO SE SELECCIONE VALORES CRUZADOS
    DEBEX2# = XVALO(msf2.TextMatrix(REG2&, 5))
    HABEX2# = XVALO(msf2.TextMatrix(REG2&, 6))
    DEBEX3# = XVALO(MSF3.TextMatrix(REG3&, 5))
    HABEX3# = XVALO(MSF3.TextMatrix(REG3&, 6))
    If Abs(DEBEX2# - DEBEX3#) > 0.005 Or Abs(HABEX2# - HABEX3#) > 0.005 Then
      Call COMUNI("Imposible Conciliar Valores Debe y Haber No Coincidente")
      GoTo 99
    End If
    
    'SUBO LA INFO A LISTA 1
    With msf2
           J& = MSF1.Rows
           MSF1.Cols = 11
           msf2.Cols = 10
           MSF1.Rows = J& + 1
           MSF1.TextMatrix(J&, 1) = MSF3.TextMatrix(REG3&, 1)
           MSF1.TextMatrix(J&, 2) = MSF3.TextMatrix(REG3&, 2) 'COMPROBANTE QUE LO TIENE EN LA LISTA 3
           MSF1.TextMatrix(J&, 3) = MSF3.TextMatrix(REG3&, 3) 'FECHA DE ACREDITACION
           MSF1.TextMatrix(J&, 4) = .TextMatrix(REG2&, 4) 'REFERENCIA DEL PENDIENTE DEL SISTEMA
           MSF1.TextMatrix(J&, 5) = .TextMatrix(REG2&, 5)
           MSF1.TextMatrix(J&, 6) = .TextMatrix(REG2&, 6)
           MSF1.TextMatrix(J&, 7) = .TextMatrix(REG2&, 7)
           MSF1.TextMatrix(J&, 8) = .TextMatrix(REG2&, 8)
           MSF1.TextMatrix(J&, 9) = .TextMatrix(REG2&, 9)
           
           
           'CON ESTO BLANQUE LA FILA PINTADA
           Call COLOREAR_GRILLASQL_CON_MARCA(CONCIMAN07.msf2, &HC0E0FF, REG2&, "*", 1)
           'GRABO EN EL CONCILIA EL COMPROBANTE QUE LEO DE LA GRILLA 3
           COMPROBANTE$ = MSF3.TextMatrix(REG3&, 2)
           ' NUMERO DE ORDEN DEL CONCILIA PARA CONDICIONAR EL UPDATE
           ORDEN& = XVALO(msf2.TextMatrix(REG2&, 7))
           FE% = FECHANUM(MSF3.TextMatrix(REG3&, 1))
           Call GRABA_COMPROBANTE_EN_CONCILIA(COMPROBANTE$, FE%, ORDEN&)
    End With
    '
    'BORRO LA FILA DE LISTA 2
    Call borrafila(msf2, REG2&)
    '
    'BORRO LA FILA DE LISTA 3
    'CON ESTO BLANQUE LA FILA PINTADA
    Call COLOREAR_GRILLASQL_CON_MARCA(CONCIMAN07.MSF3, &HCBEC86, REG3&, "*", 1)
    Call borrafila(MSF3, REG3&)
    '
    Call sortflex(MSF1)
    
    Call SUMAMSF(MSF1)
    Call SUMAMSF(msf2)
    Call SUMAMSF(MSF3)
    'llevo al final la grilla 1
    TPXI& = MSF1.Rows - 9: If TPXI& < 1 Then TPXI& = 1
    On Error Resume Next
    MSF1.TopRow = TPXI&
    On Error GoTo 0
    '
99  Command5.Enabled = True

End Sub

Private Sub Command6_Click()
    Command6.Enabled = False
    '
    If msf2.Rows > 1 Then
       Call COMUNI("Se Deberán Conciliar Todos Los Movimientos.")
       GoTo 99
    End If
    '
    RESPUESTA% = COMALTER%("¿Desea Elegir la Fecha del Asiento de Gastos?\\Si, Elegir\No, Tomar la del Día")
    FechaAsientoGastos% = HOY(HO$)
    If RESPUESTA% = 1 Then
        Call SELECHO("SELFECHA")
        FechaAsientoElegida% = FECHANUM(VALACT1$("SELFECHA"))
        If FechaAsientoElegida% > 0 Then
            FechaAsientoGastos% = FechaAsientoElegida%
        End If
    End If
    
    Screen.MousePointer = 11
    CUEII% = CUENTA_BCO_CONCILIA%(CODICONCI$)
    CUEBA$ = CODCUE$(CUEII%)
    '
    FIN& = MSF1.Rows - 1
    If FIN& < 2 Then GoTo 99
    For REG1& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      '
      With MSF1
        FECONTA1 = FETEXCOR1$(CVINT(.TextMatrix(REG1&, 1)))
        IDASIE$ = TRIM$(.TextMatrix(REG1&, 2))
        If UCase$(TRIM$(IDASIE$)) = "GASTOS" Then GoTo 29
        FEACRED% = CVINT(.TextMatrix(REG1&, 3))
        FEACRE = FETEXCOR1$(CVINT(.TextMatrix(REG1&, 3)))
        REFEMOV$ = TRIM$(.TextMatrix(REG1&, 4))
        IDEBE# = XVALO(.TextMatrix(REG1&, 5))
        IHABE# = XVALO(.TextMatrix(REG1&, 6))
        NUORD& = XVALO(.TextMatrix(REG1&, 7))
        ORIMO$ = TRIM$(.TextMatrix(REG1&, 8))
        
        If ORIMO$ = "CABAN" Then
            SQLX$ = "SELECT * FROM CAJABANC "
            SQLX$ = SQLX$ + "WHERE FECHCONTA = '" & FECONTA1 & "' "
            SQLX$ = SQLX$ + "AND CodiCom_COR = '" & IDASIE$ & "' "
            SQLX$ = SQLX$ + "AND CueContaI = " & CUEII% & " "
            '
            Dim CAJJJABANC As ADODB.Recordset
            Set CAJJJABANC = New ADODB.Recordset
27          On Error Resume Next
            CAJJJABANC.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
            If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
               On Error GoTo 0
               Call CapturaErrorOpen
               GoTo 27
            End If
            On Error GoTo 0
            With CAJJJABANC    ' CajaBanc
                CAMOCO% = 0
                Do While Not .EOF
                  If Abs(IDEBE# - !idebecomp) < 0.005 Then
                    If Abs(IHABE# - !ihabecomp) < 0.005 Then
                      If FEACRED% > 0 Then
                          !MARCONSI = "*"
                          !fechacred = FETEXCOR1$(FEACRED%)
                          CAMOCO% = CAMOCO% + 1
                          If CAMOCO% > 1 Then
                             Screen.MousePointer = 1
                             Call COMUNI("Movimientos Adicionales Conciliados\con Igual Importe e Identificador.")
                             Screen.MousePointer = 11
                          End If
                      Else
                          !MARCONSI = " "
                          CAMOCO% = CAMOCO% + 1
                          If CAMOCO% > 1 Then
                             Screen.MousePointer = 1
                             Call COMUNI("Movimientos Adicionales Des-Conciliados\con Igual Importe e Identificador.")
                             Screen.MousePointer = 11
                          End If
                      End If
                      .Update
                    End If
                  End If
                  .MoveNext
                Loop
              On Error GoTo 0
            End With

        ElseIf ORIMO$ = "DIAR" Then
               SQLX$ = "SELECT * FROM DetaAsien "
               SQLX$ = SQLX$ + "WHERE FechAsien = '" & FECONTA1 & "' "
               ' SQLX$ = SQLX$ + "AND SucuAsien = " & SUCUI% & ""
               SQLX$ = SQLX$ + "AND NuIdAsien = '" & IDASIE$ & "' "
               SQLX$ = SQLX$ + "AND CueCoInt = " & CUEII% & " "
               SQLX$ = SQLX$ + "AND LEFT(RefePase, 28) = '" & REFEMOV$ & "' "
               SQLX$ = SQLX$ + "AND StatPase > 0 "
               '
               Dim DETTTASIEN As ADODB.Recordset
               Set DETTTASIEN = New ADODB.Recordset
28             On Error Resume Next
               DETTTASIEN.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
               If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
                  On Error GoTo 0
                  Call CapturaErrorOpen
                  GoTo 28
               End If
               On Error GoTo 0
               With DETTTASIEN
                 On Error Resume Next
                 .MoveFirst
                 If Err < 1 Then
                   Do While Not .EOF
                     On Error GoTo 0
                     If Abs(IDEBE# - !idebepase) < 0.005 Then
                       If Abs(IHABE# - !ihabepase) < 0.005 Then
                         If FEACRED% > 0 Then
                             !MARCONSI = "*"
                         Else
                             !MARCONSI = " "
                         End If
                         .Update
                         GoTo 29
                       End If
                     End If
                     .MoveNext
                   Loop
                 End If
                 On Error GoTo 0
               End With
        End If

      End With
29  Next REG1&
    '--------------------------------------------------------------------------
    'GENERACION DE ASIENTOS DE GASTOS
    '--------------------------------------------------------------------------
    
    NUASI$ = ASIDIAR07.NUPROA$(FechaAsientoGastos%)
    FECHAU% = FechaAsientoGastos%
    REFEASI$ = REFECONCI_AUTOMA$(CODICONCI$)

    CANASIEN& = CantRegistros("CONCILIA", "CODICONCI = '" & CODICONCI$ & "' AND COMPROBANTE = 'GASTOS' ")
    If CANASIEN& < 1 Then GoTo 99
    Dim ENCTMP As ADODB.Recordset
    Set ENCTMP = New ADODB.Recordset
    SQLX$ = "Select * from ENCABEASIEN where NuIDAsien = '" & NUASI$ & "'"
    ENCTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
        With ENCTMP
          .AddNew   ' YA CONVERTIDO
          On Error Resume Next
          !CodiEmpr = CodiEmp%
          On Error GoTo 0
          !nuidasien = NUASI$
          !refeasien = "CA-GASTOS " + Left$(REFEASI$, 54)
          !FECHASIEN = CVDAT(FECHAU%)
             PPXX% = InStr(NUASI$, "-")
          !OrFeAsien = Val(Mid$(NUASI$, PPXX% + 1))
          !SucuAsien = 0
          !NuDfAsien = 0
          !TIPOASIEN = 0
          !StatAsien = 1  ' activo
          !EditAsien = (-1) 'BLOQUEADO PARA SU MODIFICACION
          !QPasAsien = CANASIEN&
          !DebeAsien = 0
          !Observa = " "
          !Nume_Usuar = USNBR%
          !Niv_Clasi = 0
          .Update
        End With
    ENCTMP.Close
    Set ENCTMP = Nothing
     'FIN DE GRABACION DE ENCABEZADO
    '--------------------------------------------------------------------------
    '
    '--------------------------------------------------------------------------
    'DETALLE DE ASIENTO
    '--------------------------------------------------------------------------
    Dim DETATMP As ADODB.Recordset
    Set DETATMP = New ADODB.Recordset
    SQLX$ = "Select * from DetaAsien where NuIDAsien = '" & NUASI$ & "'"
    DETATMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    ORPASE% = 0
    ACUMU_IDEBE# = 0

    REG1& = 0
    For REG1& = 1 To MSF1.Rows - 1
      Call TRACE(20, U&, FIN&)
      If TRIM$(MSF1.TextMatrix(REG1&, 2)) <> "GASTOS" Then GoTo 80
      FECONTA1 = FETEXCOR1$(CVINT(MSF1.TextMatrix(REG1&, 1)))
      FEACRED% = CVINT(MSF1.TextMatrix(REG1&, 3))
      DEPASE$ = MSF1.TextMatrix(REG1&, 4)
      IDEBE# = XVALO(MSF1.TextMatrix(REG1&, 5))
      IHABE# = XVALO(MSF1.TextMatrix(REG1&, 6))
      CTA_GASTO% = XVALO(TRIM$(MSF1.TextMatrix(REG1&, 9)))
      COCUE$ = CODCUE$(CTA_GASTO%)
      '
40      With DETATMP
            .AddNew   ' YA CONVERTIDO
             On Error Resume Next
             !CodiEmpr = CodiEmp%
             On Error GoTo 0
             !nuidasien = NUASI$
               ORPASE% = ORPASE% + 1
             !OrNuPase = ORPASE%
             !RefePase = Left$(DEPASE$, 64)
             !FECHASIEN = CVDAT(FECHAU%)
               PPXX% = InStr(NUASI$, "-")
             !OrFeAsien = XVALO(Mid$(NUASI$, PPXX% + 1))
             !SucuAsien = 0
             !NuDfAsien = 0
             !StatPase = 1
             !CUECOEXT = COCUE$
             !CUECOINT = CTA_GASTO%
             !idebepase = IDEBE#
             !ihabepase = IHABE#
               ACUMU_IDEBE# = ACUMU_IDEBE# + IHABE#
             !MARCONSI = "*"
             !CenCos = 0
             !Niv_Clasi = 0
             .Update
80           If REG1& = FIN& Then
               'SI YA LLEGO AL FINAL GRABA UN REGISTRO MAS CON EL TOTAL Y CUENTA DEL BANCO
               IHABE# = 0
               IDEBE# = ACUMU_IDEBE#
               CTA_GASTO% = CUEII% 'CUENTA DEL BANCO QUE SE ESTA CONCILIANDO
               COCUE$ = CODCUE$(CUEII%)   'CUENTA EXTERNA DEL BANCO QUE SE ESTA CONCILIANDO
               DEPASE$ = DENOCUE$(CUEII%)
               FIN& = FIN& + 1
               GoTo 40
             End If
      End With
  Next REG1&
  '
  DETATMP.Close
  Set DETATMP = Nothing
  '
  'ACTUALIZO EL TOTAL EN DETASIEN
  CONDI$ = "NuIDAsien = '" & NUASI$ & "'"
  Call ACTUREGISTRO("ENCABEASIEN", "DebeAsien", CONDI$, FORMATNUM$(ACUMU_IDEBE#, "F16.5"), REGAF&)
  '
  'CAMBIO EL STADO EN CONCILIA DE LO CONCILIADO
  CONDI$ = "CODICONCI = '" & CODICONCI$ & "'"
  Call ACTUREGISTRO("CONCILIA", "STATUS", CONDI$, 1, REGAF&, "NOMESS")
  '
  'ACTUALIZA EL INDICE DE ASIENTOS
  Call ASIDIAR07.GENINDIAS
  '
  Screen.MousePointer = 1
  If FIN& > 1 Then
   Call COMUNI("Se Ha Completado la Conciliacion Bancaria")
  End If
  '
99 Screen.MousePointer = 1
   Command6.Enabled = True

End Sub

Private Sub Command7_Click()
    Command7.Enabled = False
    'valido que haya un item seleccionado
    HAY_SELECCION_1% = 0
    With CONCIMAN07.MSF3
    For i& = 1 To .Rows - 1
          If .TextMatrix(i&, 0) = "*" Then
             HAY_SELECCION_1% = 1
             Exit For
          End If
        Next i&
    End With
    '
    If HAY_SELECCION_1% > 0 Then
       DEBEX# = XVALO(MSF3.TextMatrix(REGSEL3&, 5))
       HABEX# = XVALO(MSF3.TextMatrix(REGSEL3&, 6))
       MULITISEL09.Label8 = FORMATNUM$(DEBEX# - HABEX#, "F12.2")
       MULITISEL09.Show 1
    Else
       Call COMUNI("Falta Seleccionar \Movimientos Bancarios (en el sistema)\Pendientes de Conciliación")
       GoTo 99
    End If
    Screen.MousePointer = 11
    Call SUMAMSF(MSF1)
    Call SUMAMSF(msf2)
    Call SUMAMSF(MSF3)
    Screen.MousePointer = 1

99  Command7.Enabled = True
End Sub

Private Sub Command8_Click(Index As Integer)
    '
    Select Case Index + 1
     Case 1: Call BUSQUEDA(MSF1)
     Case 2: Call BUSQUEDA(msf2)
     Case 3: Call BUSQUEDA(MSF3)
    End Select
    '
   
   
End Sub
  Sub BUSQUEDA(GRID As MSFlexGrid)
    ABUSQUE$ = UCase$(TRIM$(InputBox$("Busqueda por Texto", "Ingrese el Texto de Búsqueda", ARBUS$)))
    If ABUSQUE$ <> "" Then
      ARBUS$ = ABUSQUE$
      For U& = 1 To GRID.Rows - 1
       For i& = 0 To GRID.Cols - 1
        If InStr(UCase$(GRID.TextMatrix(U&, i&)), ABUSQUE$) > 0 Then
         GRID.Row = U&: GRID.COL = i&
         If GRID.TopRow > GRID.Row Then GRID.TopRow = GRID.Row
         On Error Resume Next
         VISIROWS = (GRID.Height / GRID.RowHeight(J&)) - 4
         If Err > 0 Then
          VISIROWS = 1
         End If
         On Error GoTo 0
          If VISIROWS > 0 Then
             If GRID.TopRow < GRID.Row - VISIROWS Then
               GRID.TopRow = GRID.Row - VISIROWS
             End If
          End If
         GRID.SetFocus
         Exit Sub
        End If
        Next i&
      Next U&
      '
      Call MENSERR(24, "Argumento no Encontrado.")
    End If
  End Sub

Private Sub Command9_Click()
    Command9.Enabled = False
    'valido que haya un item seleccionado
    HAY_SELECCION_1% = 0
    With msf2
    For i& = 1 To .Rows - 1
          If .TextMatrix(i&, 0) = "*" Then
             HAY_SELECCION_1% = 1
             Exit For
          End If
        Next i&
    End With
    '
    If HAY_SELECCION_1% > 0 Then
       DEBEX# = XVALO(msf2.TextMatrix(REGSEL2&, 5))
       HABEX# = XVALO(msf2.TextMatrix(REGSEL2&, 6))
       MULITISEL091.Label8 = FORMATNUM$(DEBEX# - HABEX#, "F12.2")
       MULITISEL091.Show 1
    Else
       Call COMUNI("Falta Seleccionar Movimiento\de la Lista de No Conciliados Automáticamente.")
       GoTo 99
    End If
    Screen.MousePointer = 11
    
    Call SUMAMSF(MSF1)
    Call SUMAMSF(msf2)
    Call SUMAMSF(MSF3)
    Screen.MousePointer = 1

99  Command9.Enabled = True

End Sub

Private Sub MSF1_DblClick()
   If MSF1.Rows < 2 Then Exit Sub
   'ASIGNO POSICION SELECCIONADA A VARIABLES
   i& = MSF1.MouseRow
   J& = MSF1.MouseCol
   If i& = 0 Or i& > MSF1.Rows Then Exit Sub
   REGSEL1& = i& 'GUARDO EL REGISTRO SELECCIONADO COMO VARIABLE GLOBAL, CUANDO VUELVE A ENTRAR
   COMPROBANTE$ = TRIM$(MSF1.TextMatrix(REGSEL1&, 2))
   If COMPROBANTE$ <> "" Then
     Call VERCOMPROIMPRESO(COMPROBANTE$)
  End If

End Sub

Private Sub MSF1_KeyDown(KeyCode As Integer, Shift As Integer)

   If KeyCode = 114 Then ' BUSQUEDA  POR TEXTO CON F3
       If ARBUS$ <> "" Then
          VUELTA% = 0
          J& = MSF1.Row

          h& = 0: If MSF1.COL < MSF1.Cols Then h& = MSF1.COL + 1
10        For U& = J& To MSF1.Rows - 1
             HI& = h&: If U& > J& Then HI& = 0
             For i& = HI& To MSF1.Cols - 1
              If InStr(UCase$(MSF1.TextMatrix(U&, i&)), ABUSQUE$) > 0 Then
                MSF1.Row = U&: MSF1.COL = i&
                On Error Resume Next
                 VISIROWS = (MSF1.Height / MSF1.RowHeight(J&)) - 4
                If Err > 0 Then
                 VISIROWS = 1
                End If
                On Error GoTo 0
                 If VISIROWS > 0 Then
                    If MSF1.TopRow < MSF1.Row - VISIROWS Then
                      MSF1.TopRow = MSF1.Row - VISIROWS
                    End If
                 End If
                MSF1.SetFocus
                Exit Sub
              End If
              Next i&
          Next U&
          '
          Call MENSERR(24, "Argumento no Encontrado.")
       End If
   End If

End Sub
Private Sub MSF2_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 114 Then ' BUSQUEDA  POR TEXTO CON F3
       If ARBUS$ <> "" Then
          VUELTA% = 0
          J& = msf2.Row

          h& = 0: If msf2.COL < msf2.Cols Then h& = msf2.COL + 1
10        For U& = J& To msf2.Rows - 1
             HI& = h&: If U& > J& Then HI& = 0
             For i& = HI& To msf2.Cols - 1
              If InStr(UCase$(msf2.TextMatrix(U&, i&)), ABUSQUE$) > 0 Then
                msf2.Row = U&: msf2.COL = i&
                On Error Resume Next
                 VISIROWS = (msf2.Height / msf2.RowHeight(J&)) - 4
                If Err > 0 Then
                 VISIROWS = 1
                End If
                On Error GoTo 0
                 If VISIROWS > 0 Then
                    If msf2.TopRow < msf2.Row - VISIROWS Then
                      msf2.TopRow = msf2.Row - VISIROWS
                    End If
                 End If
                msf2.SetFocus
                Exit Sub
              End If
              Next i&
          Next U&
          '
          Call MENSERR(24, "Argumento no Encontrado.")
       End If
   End If

End Sub

Private Sub MSF3_DblClick()
   If MSF3.Rows < 2 Then Exit Sub
   'ASIGNO POSICION SELECCIONADA A VARIABLES
   i& = MSF3.MouseRow
   J& = MSF3.MouseCol
   If i& = 0 Or i& > MSF3.Rows Then Exit Sub
   REGSEL3& = i& 'GUARDO EL REGISTRO SELECCIONADO COMO VARIABLE GLOBAL, CUANDO VUELVE A ENTRAR
   COMPROBANTE$ = TRIM$(MSF3.TextMatrix(REGSEL3&, 2))
   If COMPROBANTE$ <> "" Then
     Call VERCOMPROIMPRESO(COMPROBANTE$)
  End If
End Sub
  
  Sub VERCOMPROIMPRESO(COMPROBANTE$)
     'ORDEN DE PAGO
     '
     Screen.MousePointer = 11
     If Left$(COMPROBANTE$, 3) = "OP-" Then
        NPX& = XVALO(Mid$(COMPROBANTE$, 4, 6))
        TIDOCUM$ = "Orden de Pago"
        NUDOCU$ = " " + TRIM$(Str$(NPX&)) + " -"
        Call MUESTRADOC(DOCUNU&, "", TIDOCUM$, Str$(NPX&), NUDOCU$)
     'RECIBO DE COBRANZA
     ElseIf Left$(COMPROBANTE$, 3) = "RB-" Then
        NPX& = XVALO(Mid$(COMPROBANTE$, 4, 6))
        TIDOCUM$ = "Recibo Nro."
        NUDOCU$ = TRIM$(Str$(NPX&))
        Call MUESTRADOC(DOCUNU&, "", TIDOCUM$, Str$(NPX&), NUDOCU$)
    
    'DEPOSITO BANCARIO
     ElseIf Left$(COMPROBANTE$, 3) = "DB-" Then
        NPX& = XVALO(Mid$(COMPROBANTE$, 4, 6))
        TIDOCUM$ = "Deposito Bancario"
        NUDOCU$ = Str$(NPX&)
        Call MUESTRADOC(DOCUNU&, "", TIDOCUM$, Str$(NPX&), COMPROBANTE$)
    End If
    '
    Screen.MousePointer = 1
    '
End Sub
Private Sub MSF3_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 114 Then ' BUSQUEDA  POR TEXTO CON F3
       If ARBUS$ <> "" Then
          VUELTA% = 0
          J& = MSF3.Row

          h& = 0: If MSF3.COL < MSF3.Cols Then h& = MSF3.COL + 1
10        For U& = J& To MSF3.Rows - 1
             HI& = h&: If U& > J& Then HI& = 0
             For i& = HI& To MSF3.Cols - 1
              If InStr(UCase$(MSF3.TextMatrix(U&, i&)), ABUSQUE$) > 0 Then
                MSF3.Row = U&: MSF3.COL = i&
                On Error Resume Next
                 VISIROWS = (MSF3.Height / MSF3.RowHeight(J&)) - 4
                If Err > 0 Then
                 VISIROWS = 1
                End If
                On Error GoTo 0
                 If VISIROWS > 0 Then
                    If MSF3.TopRow < MSF3.Row - VISIROWS Then
                      MSF3.TopRow = MSF3.Row - VISIROWS
                    End If
                 End If
                MSF3.SetFocus
                Exit Sub
              End If
              Next i&
          Next U&
          '
          Call MENSERR(24, "Argumento no Encontrado.")
       End If
   End If

End Sub


Private Sub Form_Load()
   '
   UTILIZA_SKIN% = 1
   Call APLICACIONSKINS(Me)
   '\\\OT-08-0990-OD-FIN///
   '
   TEPRUEBA.Font = MSF1.Font
   TEPRUEBA.FontSize = MSF1.FontSize
   MSF1.Cols = 10
   NUCOLU% = 1: TEPRUEBA = "99": MSF1.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF1.ColAlignment(NUCOLU% - 1) = 1: MSF1.TextMatrix(0, NUCOLU% - 1) = " "
   NUCOLU% = 2: TEPRUEBA = "99/99/99M": MSF1.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF1.ColAlignment(NUCOLU% - 1) = 1: MSF1.TextMatrix(0, NUCOLU% - 1) = "F.Cont"
   NUCOLU% = 3: TEPRUEBA = "1234567890123456789000000": MSF1.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF1.ColAlignment(NUCOLU% - 1) = 1: MSF1.TextMatrix(0, NUCOLU% - 1) = "Comprobante"
   NUCOLU% = 4: TEPRUEBA = "99/99/99M": MSF1.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF1.ColAlignment(NUCOLU% - 1) = 1: MSF1.TextMatrix(0, NUCOLU% - 1) = "F.Acred."
   NUCOLU% = 5: TEPRUEBA = "MM-0000-00000000000000000000000000": MSF1.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF1.ColAlignment(NUCOLU% - 1) = 1: MSF1.TextMatrix(0, NUCOLU% - 1) = "Det.Ref."
   NUCOLU% = 6: TEPRUEBA = "ABCDEFABCDEF": MSF1.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF1.ColAlignment(NUCOLU% - 1) = 6: MSF1.TextMatrix(0, NUCOLU% - 1) = "Debe"
   NUCOLU% = 7: TEPRUEBA = "MM-0000-0000": MSF1.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF1.ColAlignment(NUCOLU% - 1) = 6: MSF1.TextMatrix(0, NUCOLU% - 1) = "Haber"
   NUCOLU% = 8: TEPRUEBA = "MM-00 ": MSF1.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF1.ColAlignment(NUCOLU% - 1) = 6: MSF1.TextMatrix(0, NUCOLU% - 1) = "N.O." 'NUMERO DE ORDEN
   NUCOLU% = 9: TEPRUEBA = "": MSF1.ColWidth(NUCOLU% - 1) = 0: MSF1.ColAlignment(NUCOLU% - 1) = 1: MSF1.TextMatrix(0, NUCOLU% - 1) = "" 'SI VIENE DE DETAASIEN O CAJABANC
   NUCOLU% = 10: TEPRUEBA = "": MSF1.ColWidth(NUCOLU% - 1) = 0: MSF1.ColAlignment(NUCOLU% - 1) = 1: MSF1.TextMatrix(0, NUCOLU% - 1) = "" 'CUENTA CONTABLE DE GASTOS
   '
   msf2.Cols = 10
   NUCOLU% = 1: TEPRUEBA = "99": msf2.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: msf2.ColAlignment(NUCOLU% - 1) = 1: msf2.TextMatrix(0, NUCOLU% - 1) = " "
   NUCOLU% = 2: TEPRUEBA = "99/99/99M": msf2.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: msf2.ColAlignment(NUCOLU% - 1) = 1: msf2.TextMatrix(0, NUCOLU% - 1) = "F.Cont"
   NUCOLU% = 3: TEPRUEBA = "1234567890123456789000000": msf2.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: msf2.ColAlignment(NUCOLU% - 1) = 1: msf2.TextMatrix(0, NUCOLU% - 1) = "Comprobante"
   NUCOLU% = 4: TEPRUEBA = "99/99/99M": msf2.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: msf2.ColAlignment(NUCOLU% - 1) = 1: msf2.TextMatrix(0, NUCOLU% - 1) = "F.Acred."
   NUCOLU% = 5: TEPRUEBA = "MM-0000-00000000000000000000000000": msf2.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: msf2.ColAlignment(NUCOLU% - 1) = 1: msf2.TextMatrix(0, NUCOLU% - 1) = "Det.Ref."
   NUCOLU% = 6: TEPRUEBA = "ABCDEFABCDEF": msf2.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: msf2.ColAlignment(NUCOLU% - 1) = 6: msf2.TextMatrix(0, NUCOLU% - 1) = "Debe"
   NUCOLU% = 7: TEPRUEBA = "MM-0000-0000": msf2.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: msf2.ColAlignment(NUCOLU% - 1) = 6: msf2.TextMatrix(0, NUCOLU% - 1) = "Haber"
   NUCOLU% = 8: TEPRUEBA = "MM-00 ": msf2.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: msf2.ColAlignment(NUCOLU% - 1) = 6: msf2.TextMatrix(0, NUCOLU% - 1) = "N.O." 'NUMERO DE ORDEN
   NUCOLU% = 9: TEPRUEBA = "": msf2.ColWidth(NUCOLU% - 1) = 0: msf2.ColAlignment(NUCOLU% - 1) = 6: msf2.TextMatrix(0, NUCOLU% - 1) = "" 'SI VIENE DE DETAASIEN O CAJABANC
   NUCOLU% = 10: TEPRUEBA = "MM-00 ": msf2.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: msf2.ColAlignment(NUCOLU% - 1) = 6: msf2.TextMatrix(0, NUCOLU% - 1) = "N.O." 'CUENTA CONTABLE DE GASTOS
   '
   MSF3.Cols = 10
   NUCOLU% = 1: TEPRUEBA = "99": MSF3.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF3.ColAlignment(NUCOLU% - 1) = 1: MSF3.TextMatrix(0, NUCOLU% - 1) = " "
   NUCOLU% = 2: TEPRUEBA = "99/99/99M": MSF3.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF3.ColAlignment(NUCOLU% - 1) = 1: MSF3.TextMatrix(0, NUCOLU% - 1) = "F.Cont"
   NUCOLU% = 3: TEPRUEBA = "1234567890123456789000000": MSF3.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF3.ColAlignment(NUCOLU% - 1) = 1: MSF3.TextMatrix(0, NUCOLU% - 1) = "Comprobante"
   NUCOLU% = 4: TEPRUEBA = "99/99/99M": MSF3.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF3.ColAlignment(NUCOLU% - 1) = 1: MSF3.TextMatrix(0, NUCOLU% - 1) = "F.Acred."
   NUCOLU% = 5: TEPRUEBA = "MM-0000-00000000000000000000000000": MSF3.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF3.ColAlignment(NUCOLU% - 1) = 1: MSF3.TextMatrix(0, NUCOLU% - 1) = "Det.Ref."
   NUCOLU% = 6: TEPRUEBA = "ABCDEFABCDEF": MSF3.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF3.ColAlignment(NUCOLU% - 1) = 6: MSF3.TextMatrix(0, NUCOLU% - 1) = "Debe"
   NUCOLU% = 7: TEPRUEBA = "MM-0000-0000": MSF3.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF3.ColAlignment(NUCOLU% - 1) = 6: MSF3.TextMatrix(0, NUCOLU% - 1) = "Haber"
   NUCOLU% = 8: TEPRUEBA = "": MSF3.ColWidth(NUCOLU% - 1) = 0: MSF3.ColAlignment(NUCOLU% - 1) = 6: MSF3.TextMatrix(0, NUCOLU% - 1) = ""
   NUCOLU% = 9: TEPRUEBA = "": MSF3.ColWidth(NUCOLU% - 1) = 0: MSF3.ColAlignment(NUCOLU% - 1) = 6: MSF3.TextMatrix(0, NUCOLU% - 1) = "" 'SI VIENTE DE DETAASIEN O CAJABANC
   NUCOLU% = 10: TEPRUEBA = "": MSF3.ColWidth(NUCOLU% - 1) = 0: MSF3.ColAlignment(NUCOLU% - 1) = 6: msf2.TextMatrix(0, NUCOLU% - 1) = "" 'CUENTA CONTABLE DE GASTOS
   '
   REGSEL1& = 0
   REGSEL2& = 0
   REGSEL3& = 0
   
End Sub



Private Sub MSF1_Click()
   '
   If MSF1.Rows < 2 Then Exit Sub
   'ASIGNO POSICION SELECCIONADA A VARIABLES
   i& = MSF1.MouseRow
   J& = MSF1.MouseCol
   If i& = 0 Or i& > MSF1.Rows Then Exit Sub

   Call COLOREAR_GRILLASQL_CON_MARCA(CONCIMAN07.MSF1, &HFFFFC0, i&, "*", 1, REGSEL1&)
   REGSEL1& = i& 'GUARDO EL REGISTRO SELECCIONADO COMO VARIABLE GLOBAL, CUANDO VUELVE A ENTRAR
   'BORRA PRIMERO ESTE REGISTRO Y LUEGO PINTA EL PASADO EN I&,
End Sub

Private Sub msf2_Click()
   '
   If msf2.Rows < 2 Then Exit Sub
   'ASIGNO POSICION SELECCIONADA A VARIABLES
   i& = msf2.MouseRow
   J& = msf2.MouseCol
   If i& = 0 Or i& > msf2.Rows Then Exit Sub

   Call COLOREAR_GRILLASQL_CON_MARCA(CONCIMAN07.msf2, &HC0E0FF, i&, "*", 1, REGSEL2&)
   REGSEL2& = i& 'GUARDO EL REGISTRO SELECCIONADO COMO VARIABLE GLOBAL, CUANDO VUELVE A ENTRAR
   'BORRA PRIMERO ESTE REGISTRO Y LUEGO PINTA EL PASADO EN I&,

End Sub
Private Sub MSF3_Click()
   '
   If MSF3.Rows < 2 Then Exit Sub
   'ASIGNO POSICION SELECCIONADA A VARIABLES
   i& = MSF3.MouseRow
   J& = MSF3.MouseCol
   If i& = 0 Or i& > MSF3.Rows Then Exit Sub
   Call COLOREAR_GRILLASQL_CON_MARCA(CONCIMAN07.MSF3, &HCBEC86, i&, "*", 1, REGSEL3&)
   REGSEL3& = i& 'GUARDO EL REGISTRO SELECCIONADO COMO VARIABLE GLOBAL, CUANDO VUELVE A ENTRAR
   'BORRA PRIMERO ESTE REGISTRO Y LUEGO PINTA EL PASADO EN I&,

End Sub

Sub COLOREAR_GRILLASQL_CON_MARCA(MSF As MSFlexGrid, COLOR As Variant, REG&, Optional MARCA$, Optional UNICASELECCION%, Optional REGISTRO_a_BLANQUEAR&)
   'CON UN CLIC PINTA TODO EL REGISTRO DEL COLOR PARAMETRIZADO Y LE PONE UN ASTERIZCO EN COLUMNA 0
   'EN EL SEGUNDO CLIC SI YA ESTA PINTADA BORRA EL COLOR Y TAMBIEN EL ASTERIZCO
   '
   FILA_MARCADA% = 0
   With MSF
     
        ASTERI$ = ""
        If TRIM$(MARCA$) <> "" Then ASTERI$ = MARCA$
        '
        'SI ES UNICA SELECCION CUANDO MARCA UNA BORRA LA ANTERIOR PARA QUE SIEMPRE QUEDE SOLO UNA MARCADA
        If UNICASELECCION% > 0 Then
          'SI TIENE VALOR EL 'REGISTROABLANQUEAR' BLANQUEA EL REGISTRO PASADO COMO PARAMETRO
          If REGISTRO_a_BLANQUEAR& > 0 And REGISTRO_a_BLANQUEAR& < MSF.Rows Then
             MSF.Row = REGISTRO_a_BLANQUEAR&
             For J& = 1 To MSF.Cols - 1
                 .COL = J&
                 .CellBackColor = vbWhite
                 If .TextMatrix(REGISTRO_a_BLANQUEAR&, 0) = ASTERI$ Then FILA_MARCADA% = 1
                 .TextMatrix(REGISTRO_a_BLANQUEAR&, 0) = ""
             Next J&
             If FILA_MARCADA% = 1 And REGISTRO_a_BLANQUEAR& = REG& Then Exit Sub
          End If
          GoTo 30 'VA A PINTAR LA CELDA
          
'ESTE BLOQUE DE SOLO UNICA ESTARIA QUEDANDO INUTIL CONSIDERANDO QUE EN ESTE FORMULARIO
'SE ESTA GURDANDO DE MANERA GLOBAL EL REGISTRO SELECCIONADO Y CUANDO CLICKEA EN OTRO
'LIMIPIA SOLAMENTE ESTE.
          For i& = 1 To .Rows - 1
            .Row = i&
            If MSF.TextMatrix(i&, 0) = MARCA$ Then
                For J& = 1 To MSF.Cols - 1
                  If i& <> REG& Then
                    .COL = J&
                    .CellBackColor = vbWhite
                    .TextMatrix(i&, 0) = ""
                  End If
                Next J&
                GoTo 30

            End If
          Next i&
        End If
'FIN SOLO UNICA
        '
        'PINTA LA FILA SELECCIONADA
        
30      If REG& < 1 Then Exit Sub
        .Row = REG&
        .COL = 1
        '
        'SI ES EL MISMO COLOR LA BLANQUEA Y QUITA EL AZTERIZCO CAMBIANDO EL VALOR DE LAS VARIABLES
        If .CellBackColor = COLOR Then
           COLOR = vbWhite: ASTERI$ = ""
        End If
        
        For iCol% = 1 To .Cols - 1 ' DENTRO DEL RENGLON RECORRE CADA COLUMNA
          .COL = iCol%
          .CellBackColor = COLOR
        Next iCol%
        '
        'AGREGA O QUITA EL AZTERIZCO
        .TextMatrix(REG&, 0) = ASTERI$
        
   End With
End Sub


Function VALIDA_SI_ES_GASTO%(CODICONCI$, CONCEPTOMSF2$)

    'LA FUNCION VALIDA SI ES UN GASTO PARA ELLO RECORRE LA LISTA DE
    'REFERENCIA DE GASTOS DE LA TABLA GASEXTRA Y SI EXISTE COMPARADO CON EL PASADO COMO PARAMETRO
    'GRABA EN EL CONCILIA (LOS QUE POSEAN ESTE TEXTO ) EL NUMER DE CUENTA Y EL COMPROBANTE GASTOS

    CUEI% = CUENTA_BCO_CONCILIA%(CODICONCI$)
    CUENTA_SELECCIONADA$ = CODCUE$(CUEI%)
    VALIDA_SI_ES_GASTO% = 0
    '
    RFF$ = RECFILT$("GASEXTRA", 1, AJUSTI$(CUENTA_SELECCIONADA$, 12))
    For U& = 1 To Len(RFF$) Step 4
        REG& = CVS(Mid$(RFF$, U&, 4))
        XX$ = REGLEIDO$("GASEXTRA", REG&)
        CONCEP$ = TRIM$(Mid$(XX$, 13, 128))
        CUENTACONT% = CUEINT%(TRIM$(Mid$(XX$, 141, 12)))
        If InStr(UCase$(CONCEP$), UCase$(TRIM$(CONCEPTOMSF2$))) > 0 Then
          CONDI$ = "(COMPROBANTE = '' OR COMPROBANTE IS NULL) AND CONCEPTO LIKE '%" & CONCEP$ & "%' AND STATUS < 1 "
          CONDI$ = CONDI$ + " AND CODICONCI = '" & CODICONCI$ & "'"
          Call ACTUREGISTRO("CONCILIA", "CUECONTAI", CONDI$, CUENTACONT%, REGAF&, "NOMESS")
          Call ACTUREGISTRO("CONCILIA", "COMPROBANTE", CONDI$, "GASTOS", REGAF&, "NOMESS")
          VALIDA_SI_ES_GASTO% = 1
        End If
    Next U&

End Function

   Sub GRABA_COMPROBANTE_EN_CONCILIA(COMPROBANTE$, FE%, ORDEN&)
       CONDI$ = "ORDCONCILIA = " & ORDEN&
       CONDI$ = CONDI$ + " AND CODICONCI = '" & CODICONCI$ & "'"
       Call ACTUREGISTRO("CONCILIA", "COMPROBANTE", CONDI$, COMPROBANTE$, REGAF&, "NOMESS")
       Call ACTUREGISTRO("CONCILIA", "FECHACOMPROBANTE", CONDI$, FETEXCOR1$(FE%), REGAF&, "NOMESS")
   End Sub
   
Sub ORDENA_MSF2()
    '
    SQLX$ = "SELECT * FROM CONCILIA WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE CODICONCI = '" & CODICONCI$ & "'"
    SQLX$ = SQLX$ + " AND (COMPROBANTE = '' OR COMPROBANTE IS NULL)"
    SQLX$ = SQLX$ + " AND status < 1"
    SQLX$ = SQLX$ + " ORDER BY ORDCONCILIA"
    '
    CONCIMAN07.msf2.Rows = 1
    '
    CONCIMAN07.msf2.Redraw = False
    Dim RSTTMP2 As ADODB.Recordset
    Set RSTTMP2 = READSET(SQLX$)
    With RSTTMP2
      Do While Not .EOF
         IMPODEBEX# = XVALO(!DEBE)
         IMPOHABEX# = XVALO(!HABER)
         CUEII% = XVALO(!CUECONTAI)
         IDREFE$ = SAFETEXT$(ID_REFERENCIA)
         FECONTA% = CVINT(!FechaComprobante)
         FECH1 = FETEXCOR1$(FECONTA%)
         FECH_EXTR% = CVINT(!FECHA)
         CONCEP$ = SAFETEXT$(!CONCEPTO)
         REFECON$ = SAFETEXT$(!ID_REFERENCIA)
         COMPRBTE$ = SAFETEXT$(!COMPROBANTE)
         NUORD% = XVALO(!ORDCONCILIA)
         DETABLA$ = SAFETEXT$(!COMPROB_DE_TABLA)
         With CONCIMAN07.msf2
            REG& = REG& + 1
            .Rows = REG& + 1
            .TextMatrix(REG&, 1) = FECHATEX$(FECH_EXTR%) 'FECHA DEL COMPROBANTE
            .TextMatrix(REG&, 2) = COMPRBTE$
            .TextMatrix(REG&, 3) = FECHATEX$(FECONTA%)
            .TextMatrix(REG&, 4) = CONCEP$
            .TextMatrix(REG&, 6) = TRIM$(FORMATNUM$(IMPODEBEX#, "F12.2"))
            .TextMatrix(REG&, 5) = TRIM$(FORMATNUM$(IMPOHABEX#, "F12.2"))
            .TextMatrix(REG&, 7) = NUORD%
            .TextMatrix(REG&, 8) = DETABLA$
                TX$ = FILTERGETRECORD$("GASEXTRA", 2, AJUSTI$(CONCEP$, 128))
                CODICUENTA$ = TRIM$(Mid$(TX$, 141, 12))
            .TextMatrix(REG&, 9) = CODICUENTA$
         End With
      .MoveNext
      Loop
    End With
    CONCIMAN07.msf2.Redraw = True
    RSTTMP2.Close
    Set RSTTMP2 = Nothing

End Sub

   

