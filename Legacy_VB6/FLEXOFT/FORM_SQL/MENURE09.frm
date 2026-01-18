VERSION 5.00
Begin VB.Form MENURE09 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ordenes de Reparación"
   ClientHeight    =   7020
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5265
   Icon            =   "MENURE09.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7020
   ScaleWidth      =   5265
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command20 
      Caption         =   "Realizar Consumo de una Repara Cerrada"
      Height          =   375
      Left            =   480
      TabIndex        =   42
      ToolTipText     =   "Genera el consumo de una reparacion que se cerro pero que no se hizo el consumo"
      Top             =   3360
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Ficha / Limpieza"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1110
      Left            =   240
      TabIndex        =   33
      Top             =   1275
      Width           =   3900
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   400
         ScaleHeight     =   585
         ScaleWidth      =   2985
         TabIndex        =   34
         Top             =   330
         Width           =   3045
         Begin VB.CommandButton Command14 
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
            Left            =   2445
            Picture         =   "MENURE09.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   37
            ToolTipText     =   "O/Trabajo Por Limpieza y Desarme"
            Top             =   0
            Width           =   560
         End
         Begin VB.CommandButton Command17 
            Height          =   600
            Left            =   0
            Picture         =   "MENURE09.frx":0614
            Style           =   1  'Graphical
            TabIndex        =   36
            ToolTipText     =   "Ficha de Reparacion"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Ficha de Reparacion / OF Limpieza"
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
            Left            =   645
            TabIndex        =   35
            Top             =   105
            Width           =   1800
         End
      End
   End
   Begin VB.PictureBox Picture14 
      BackColor       =   &H00400000&
      Height          =   7050
      Left            =   4320
      ScaleHeight     =   6990
      ScaleWidth      =   1155
      TabIndex        =   22
      Top             =   -45
      Width           =   1215
      Begin VB.CommandButton Command170 
         Caption         =   "Rem."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   620
         Left            =   135
         Picture         =   "MENURE09.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "VisualizarRemitos"
         Top             =   4080
         Width           =   650
      End
      Begin VB.CommandButton Command18 
         Caption         =   "Q.R."
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
         Left            =   135
         TabIndex        =   39
         ToolTipText     =   "Quita Reservas de Ordenes de Reparación Cerradas"
         Top             =   6045
         Width           =   645
      End
      Begin VB.CommandButton Command7 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   620
         Left            =   135
         Picture         =   "MENURE09.frx":0C28
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Abrir Anotador de Seguimiento de Reparaciones"
         Top             =   2130
         Width           =   650
      End
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
         Height          =   620
         Left            =   135
         Picture         =   "MENURE09.frx":0F32
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Visualizar Comprobante de Recepción"
         Top             =   2775
         Width           =   650
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
         Height          =   620
         Left            =   135
         Picture         =   "MENURE09.frx":123C
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Visualizar Presupuestos"
         Top             =   3420
         Width           =   650
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
         Height          =   620
         Left            =   135
         Picture         =   "MENURE09.frx":1546
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Visualizar Ordenes de Reparación"
         Top             =   4750
         Width           =   650
      End
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
         Height          =   620
         Left            =   135
         Picture         =   "MENURE09.frx":1850
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Materiales Suministrados por el Cliente"
         Top             =   5400
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   135
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   26
         Top             =   6375
         Width           =   645
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   27
            Top             =   105
            Width           =   12000
         End
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
         Height          =   620
         Left            =   135
         Picture         =   "MENURE09.frx":1C92
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Cierra - Abandona la Aplicación"
         Top             =   135
         Width           =   650
      End
      Begin VB.CommandButton Command11 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   620
         Left            =   135
         Picture         =   "MENURE09.frx":1DDC
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Ayuda FLEXOFT en línea"
         Top             =   780
         Width           =   650
      End
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
         Height          =   620
         Left            =   135
         Picture         =   "MENURE09.frx":20E6
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Configuración de Funcionamiento"
         Top             =   1425
         Width           =   650
      End
      Begin VB.Image Image3 
         Height          =   390
         Left            =   -300
         Picture         =   "MENURE09.frx":2528
         Top             =   6570
         Width           =   1515
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Cotización"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1110
      Left            =   210
      TabIndex        =   17
      Top             =   2420
      Width           =   3900
      Begin VB.PictureBox Picture5 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   675
         ScaleHeight     =   585
         ScaleWidth      =   2580
         TabIndex        =   18
         Top             =   330
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
            Picture         =   "MENURE09.frx":444A
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Generar Presupuesto de Reparación"
            Top             =   -15
            Width           =   560
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Generación de Presu- puesto de Reparación"
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
            Left            =   615
            TabIndex        =   19
            Top             =   60
            Width           =   1905
         End
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Pedido de Cotización"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1110
      Left            =   210
      TabIndex        =   13
      Top             =   105
      Width           =   3900
      Begin VB.PictureBox Picture6 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   360
         ScaleHeight     =   585
         ScaleWidth      =   3045
         TabIndex        =   14
         Top             =   330
         Width           =   3105
         Begin VB.CommandButton Command19 
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
            Left            =   2500
            Picture         =   "MENURE09.frx":4754
            Style           =   1  'Graphical
            TabIndex        =   41
            ToolTipText     =   "Anulación de Recepción / Revertir Anulación"
            Top             =   0
            Width           =   560
         End
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
            Picture         =   "MENURE09.frx":4A5E
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Recepcion Equipos para Cotizar Reparación"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Recepción de Equipos a Reparar"
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
            TabIndex        =   16
            Top             =   105
            Width           =   1800
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Expedición / Despacho"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1110
      Left            =   210
      TabIndex        =   10
      Top             =   5760
      Width           =   3900
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   615
         ScaleHeight     =   585
         ScaleWidth      =   2820
         TabIndex        =   11
         Top             =   330
         Width           =   2880
         Begin VB.CommandButton Command16 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   2280
            Picture         =   "MENURE09.frx":4D68
            Style           =   1  'Graphical
            TabIndex        =   38
            ToolTipText     =   "Acceso a Maestro de Artículos"
            Top             =   0
            Width           =   560
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
            Height          =   600
            Left            =   0
            Picture         =   "MENURE09.frx":5072
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "Conexión con Remito"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Generar Remito de Expedición"
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
            TabIndex        =   12
            Top             =   105
            Width           =   1485
         End
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Reparaciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2145
      Left            =   210
      TabIndex        =   0
      Top             =   3600
      Width           =   3900
      Begin VB.PictureBox Picture4 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   930
         ScaleHeight     =   585
         ScaleWidth      =   2580
         TabIndex        =   1
         Top             =   1365
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
            Picture         =   "MENURE09.frx":537C
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Anulación de Orden en Curso en Curso"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Anulación de Orden de Reparación"
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
            TabIndex        =   3
            Top             =   105
            Width           =   1695
         End
      End
      Begin VB.PictureBox Picture11 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   615
         ScaleHeight     =   585
         ScaleWidth      =   2580
         TabIndex        =   4
         Top             =   840
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
            Picture         =   "MENURE09.frx":5686
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Cerrar Orden en Curso"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Cerrar Orden de Reparación en Curso"
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
            Top             =   50
            Width           =   1905
         End
      End
      Begin VB.PictureBox Picture12 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   300
         ScaleHeight     =   585
         ScaleWidth      =   2580
         TabIndex        =   7
         Top             =   315
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
            Left            =   15
            Picture         =   "MENURE09.frx":5990
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Generación a Partir de un Presupuesto Previo"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "Generar Orden de Reparación"
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
            Left            =   945
            TabIndex        =   9
            Top             =   45
            Width           =   1485
         End
      End
   End
End
Attribute VB_Name = "MENURE09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub CONSUMO_MAT_SUM_CLIENTE(TALONX&, Optional NROREMI&)
             Campos$ = "ci/f0;Código/A16;Descripción/A20;Cantidad/f10.1;Consumo/F10.1;NCLIENTE/F0"
             Call CARGA_ENCABEZADO(BRECPAHP.MSF, Campos$, BRECPAHP)
             Call CARGA_ENCABEZADO(BRECPAHP.MSFFORMULA, Campos$, BRECPAHP)
             FORMATOS$ = GENERA_FORMATOS$(Campos$) ' DEVUELVE ESTO = " /A16; /A20; /A3; /f10.1; /D8; /F10.2; /F12.2; /F10.2; /F7; /F10.2; /F0, /F0; /D8"
             FORMATOSCONBARRA$ = GENERA_FORMATOS_CON_BARRA$(Campos$)
             Call CARGA_ENCABEZADO(BRECPAHP.MSFTOTALES, FORMATOSCONBARRA$, BRECPAHP)
        
             BRECPAHP.MSFFORMULA.Rows = 2  'DEBEN IR EN FILA 1
'            BRECPAHP.MSFFORMULA.TextMatrix(1, 7) = "4;*;6"
'            BRECPAHP.MSFFORMULA.TextMatrix(1, 9) = "=PORCENT_ENTRE_2_VALORES;10;6" 'DEBE LLEVAR SIMBOLO =
             Call AGREGAFILA(BRECPAHP.MSFFORMULA, FORMATOS$, 2) 'DEBEN IR EN LA FILA 2
             '
             'EDITABLES
             BRECPAHP.EDITABLES.Caption = "5"
             '
             BRECPAHP.Caption = "Materiales Suministrados por el Cliente (Talón: " & TALONX& & ")"
             TALONX& = XVALO(TALONX&)
             If TALONX& < 1 Then Exit Sub
         
             SQLX$ = "SELECT  COD_INTE,COD_EXTE,(SUM(CANTINGRE) - SUM(CANTSALID)) AS INGRE,NUME_CLIE "
             SQLX$ = SQLX$ + " FROM DETAMASU WITH(NOLOCK)"
             SQLX$ = SQLX$ + " WHERE TALON =  " & TALONX&
             SQLX$ = SQLX$ + " GROUP BY COD_INTE,COD_EXTE,NUME_CLIE"
             SQLX$ = SQLX$ + " HAVING (SUM(CANTINGRE) - SUM(CANTSALID)) > 0"
             Set rs = READSET(SQLX$)
             With rs
               Do While Not .EOF
                  i& = i& + 1
                  BRECPAHP.MSF.Rows = i& + 1
                  CIXI% = XVALO(!cod_inte)
                  BRECPAHP.MSF.TextMatrix(i&, 1) = CIXI%
                  BRECPAHP.MSF.TextMatrix(i&, 2) = CODEXT$(CIXI%)
                  BRECPAHP.MSF.TextMatrix(i&, 3) = DESCRIT0$(CIXI%)
                  BRECPAHP.MSF.TextMatrix(i&, 4) = FORMATNUM$(XVALO(!INGRE), "F10.1")
                  BRECPAHP.MSF.TextMatrix(i&, 5) = ""
                  BRECPAHP.MSF.TextMatrix(i&, 6) = XVALO(!NUME_CLIE)
                  .MoveNext
               Loop
             End With
        
             rs.Close
             Set rs = Nothing
             'SI NO HAY ITEMS PARA CONSUMIR SE VA.
             If BRECPAHP.MSF.Rows <= 1 Then Exit Sub
       
             'SI LA LLAMADA A ESTA RUTINA VIENE DE LA GENERACION DE REMITO PREGUNTA SI SE QUIERE INFORMAR CONSUMO O DEVOLUCION DE MATERIALES
             'SI LLEGA ACA ES POR QUE HAY MATERIALES PENDIENTES DE CONSUMO O DE DEVOLUCION
             If NROREMI& > 0 Then
                OPX% = COMALTER%("Atención Se Han Detectado Materiales Suministrados por El Cliente Pendientes de Consumo o Devolución\\Informar Consumo\Continuar")
                If OPX% > 1 Then Exit Sub
             End If
              
              
             'PRESENTA LA VENTANA CON EL DETALLE DE LOS PENDIENTES MSC
             Screen.MousePointer = 1
80           BRECPAHP.Show 1
'SI APRUEBA VENTANA
             'VALIDACION QUE LA CANTIDAD A CONSUMIR NO SUPERE LA CANTIDAD ORIGINAL
             If BRECPAHP.APROBADO% > 0 Then
                 With BRECPAHP
                   For i& = 1 To .MSF.Rows - 1
                     CANTI = XVALO(.MSF.TextMatrix(i&, 4))
                     CANTICONS = XVALO(.MSF.TextMatrix(i&, 5))
                     If CANTICONS > CANTI Then
                        Call COMUNI("Imposible Continuar !!!\Cantidad a Consumir en Fila " & i& & "\ Supera la Cantidad Recibida")
                        GoTo 80
                     End If
                   Next i&
                 End With

                'FIN VALIDACION
                '
'RECORRE LA GRILLA Y SI HAY CONSUMO GENERA EL MOVIMIEINTO
                'GENERA EL CONSUMO
                NEWID& = PROXID_DETAMASU
                CONTITULO% = 1
                FE1% = HOY(HO$)
                With BRECPAHP
                   For i& = 1 To .MSF.Rows - 1
                     CANTICONS = XVALO(.MSF.TextMatrix(i&, 5))
                     If CANTICONS > 0 Then
                       CIXI% = XVALO(.MSF.TextMatrix(i&, 1))
                       NCLIE = XVALO(.MSF.TextMatrix(i&, 6))
                       Call MOVI_SUMINISTRADOS(NEWID&, CIXI%, "CF", CANTICONS * (-1), NCLIE, TALONX&, FE1%, "Consumo de Fábrica")
                       Call FICHREP09.ANOTAMATSUM(TALONX&, CODEXT$(CIXI%), CANTICONS, DESCRIT0$(CIXI%), "Detalle de Consumo de Material Suministrado por el Cliente:", CONTITULO%)
                       CONTITULO% = 0
                     End If
                   Next i&
                 End With
                'FIN CONSUMO
                'SI ESTA VARIABLE QUE SE INICIALIZA EN 1 ESTA EN 0 ES POR QUE ENTRO A CONSUMIR ALGUN ITEM
                If CONTITULO% < 1 Then
                  'IMPRIME CONSUMO
                  Call PRESRE09.AGREGA_ANOTAREPARA(TALONX&, String$(88, "="))
                  FILTX$ = TRIM$(Str$(TALONX&)) & ";" & TRIM$(NEWID&)
                  Call STDFORM("CONS-ARC", FILTX$)
                End If
                
                'SI LA LLAMADA A ESTA RUTINA VIENE DE LA GENERACION DE REMITO PREGUNTA SI SE QUIERE INFORMAR CONSUMO O DEVOLUCION DE MATERIALES
          
                If NROREMI& > 0 Then
                    With BRECPAHP
                       For i& = 1 To .MSF.Rows - 1
                         CANTI = XVALO(.MSF.TextMatrix(i&, 4))
                         CANTICONS = XVALO(.MSF.TextMatrix(i&, 5))
                         If CANTICONS < CANTI Then
                         'HASTA AQUI LLEGUE SI SE DA ESTA CONDICION SE DEBERA PREGUNTAR SI SE QUIERE DEVOLVER LA DIFERENCIA.
                            OPX% = COMALTER%("Atención Se Han Detectado Materiales Suministrados Sin Consumir\\Continuar Sin Generar Devolución\Realizar Devolución Automática De lo No Consumido")
                            If OPX% <= 1 Then Exit Sub
                            Exit For
                         End If
                       Next i&
                    End With
                    'SI SIGUE REALIZA LA DEVOLUCION
        
                    NEWID& = PROXID_DETAMASU
                    VUELTA% = 0: CONTITULO% = 1
                    With BRECPAHP
                      For i& = 1 To .MSF.Rows - 1
                         CIXI% = XVALO(.MSF.TextMatrix(i&, 1))
                         Codigo$ = .MSF.TextMatrix(i&, 2)
                         CANTI = XVALO(.MSF.TextMatrix(i&, 4))
                         CANTICONS = XVALO(.MSF.TextMatrix(i&, 5))
                         NCLIE = XVALO(.MSF.TextMatrix(i&, 6))
                         If CANTICONS < CANTI Then
                           DEVOLUCION = CANTI - CANTICONS
                           Call MOVI_SUMINISTRADOS(NEWID&, CIXI%, "DV", DEVOLUCION * (-1), NCLIE, TALONX&, HOY(HO$), "Devolución de Suministros")
                           Call FICHREP09.ANOTAMATSUM(TALONX&, Codigo$, DEVOLUCION, DESCRI$, "Devolución de Materiales Suministrados por el Cliente:", CONTITULO%)
                           CONTITULO% = 0
                         End If
                      Next i&
                    End With
                    '
                    'IMPRIME DEVOLUCION DE MATERIALES
                    Call PRESRE09.AGREGA_ANOTAREPARA(TALONX&, String$(88, "="))
                    FILTX$ = TRIM$(Str$(TALONX&)) & ";" & TRIM$(NEWID&)
                    Call STDFORM("REDV-ARC", FILTX$)

                End If
             End If
            

End Sub


Private Sub Command1_Click()
   '
   Call CIERRARCH("*.*")
   Call FINAL("")
   '
End Sub

Private Sub Command10_Click()
    '
    Command10.Enabled = False
    '
    Call RECEBO09.verRecepcion
    GoTo 99
    
    
    Call COPYSTRU("PRESREPA", "PREPAEMI")
    '
'3   Call SELECHO("PRESREPA/Indice Recepciones de Reparación")
'    NPX& = Val(VALACT1$("PRESREPA"))
    '
3   Call SELECHO("PREPAEMI/Indice Recepciones de Reparación")
    NPX& = Val(VALACT1$("PREPAEMI"))
    '
    If NPX& < 1 Then GoTo 99
    '
    NUDOCU$ = TRIM$(Str$(NPX&))
'    For UI& = ULTREG&("PRESREPA") To 1 Step -1
'      XX$ = REGLEIDO$("PRESREPA", UI&)
'      If CVS(Left$(XX$, 4)) = NPX& Then GoTo 4
'    Next UI&
'    Call COMUNI("Error de Indices - Consulte.")
'    GoTo 99
    '
4   NUDOCU$ = TRIM$(Str$(NPX&))
    Screen.MousePointer = 11
    TIDOCUM$ = "Recepción de Reparación " + NUDOCU$
    Screen.MousePointer = 11
    '
    For UI& = ULTREG&("PDOCLST") To 1 Step -1
      XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
      If InStr(XX$, TIDOCUM$) > 4 Then
        DOCUNU& = CVS(Left$(XX$, 4))
        GoTo 5
      End If
    Next UI&
    '
    Screen.MousePointer = 1
    Call COMUNI("No se Encontró el Documento Impreso")
    On Error Resume Next
    GoTo 99
    '
5   Call MUESTRADOC(DOCUNU&)
'5   For KKI% = 0 To 5
'      EDITFORM.Picture1(KKI%).Cls
'      EDITFORM.Picture1(KKI%).DrawWidth = 3
'      EDITFORM.Picture1(KKI%).DrawStyle = 0
'      EDITFORM.Picture1(KKI%).FillStyle = 1
'    Next KKI%
'    '
'    REDOCU$ = REGACT$("PDOCLST")
'    URE& = ULTREG&("PDOCSPL")
'    LI& = 0: LS& = URE&
'6   E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 7
'    L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 7
'    DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
'    If DCC& < DOCUNU& Then LI& = L&: GoTo 6
'    If DCC& > DOCUNU& Then LS& = L&: GoTo 6
'    GoTo 8
'    '
'7   Screen.MousePointer = 1
'    Call MENSERR(24, "Documento no Registrado")
'    Exit Sub
'    '
'8   TPSP$ = ""
'    Screen.MousePointer = 11
'    While L& > 1
'      DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
'      If DCC& < DOCUNU& Then GoTo 9
'      L& = L& - 1
'    Wend
'    '
'9   CAPAGINAS% = 0
'    For KKL& = L& To URE&
'      TBUF$ = REGLEIDO$("PDOCSPL", KKL&)
'      If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 95
'      If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
'        TPSP$ = TPSP$ + Mid$(TBUF$, 5)
'      End If
'    Next KKL&
'    '
'95  SPOOLSTRING$ = TPSP$
'    For KKI% = 0 To 5
'      EDITFORM.Picture1(KKI%).Cls
'      EDITFORM.Picture1(KKI%).Top = 0
'      EDITFORM.Picture1(KKI%).Refresh
'      EDITFORM.Picture1(KKI%).Height = 7170
'      EDITFORM.Picture1(KKI%).Width = 11700
'      EDITFORM.Picture1(KKI%).DrawWidth = 3
'      EDITFORM.Picture1(KKI%).DrawStyle = 0
'      EDITFORM.Picture1(KKI%).FillStyle = 1
'    Next KKI%
'    ALTUPAGINA = 7170
'    TOPEPAGINA = 0
'    PAGINACTIVA% = 0
'    CAPAGINAS% = 0
'    '
'50  PPXX% = InStr(TPSP$, Chr$(255))
'    If PPXX% > 0 Then
'       COPRIAT$ = Left$(TPSP$, PPXX% - 1)
'       TPSP$ = Mid$(TPSP$, PPXX% + 1)
'       '
'       If Len(COPRIAT$) > 0 Then
'         COPRI% = Asc(Left$(COPRIAT$, 1))
'         ATRIP$ = Mid$(COPRIAT$, 2)
'         Call SHOWCOMMAND(COPRI%, ATRIP$)
'       End If
'       GoTo 50
'       '
'    End If
'    '
'    Screen.MousePointer = 1
'    PAGINACTIVA% = 0
'    For KKI% = 0 To 5
'      EDITFORM.Picture1(KKI%).Visible = False
'      If KKI% = PAGINACTIVA% Then EDITFORM.Picture1(KKI%).Visible = True
'    Next KKI%
'    EDITFORM.Picture1(PAGINACTIVA%).Refresh
'    EDITFORM.VScroll1.Min = 0
'    EDITFORM.VScroll1.Value = 0
'    EDITFORM.VScroll1.Max = 1
'    If ALTUPAGINA > EDITFORM.Frame1.Height Then
'       EDITFORM.VScroll1.Max = (ALTUPAGINA - EDITFORM.Frame1.Height) / 100
'    End If
'    EDITFORM.Command3.Enabled = False
''   EDITFORM.Command5.Enabled = False
'    EDITFORM.Show 1
'    '
    GoTo 3
    '
99  Command10.Enabled = True
    '
End Sub

Private Sub Command12_Click()
   Command12.Enabled = False
   Call CONECRUN("MASUC09", "Material Suministrado por el Cliente")
   Command12.Enabled = True
End Sub

Private Sub Command13_Click()
   Command13.Enabled = False
   CONFRE09.Show 1
   Command13.Enabled = True
End Sub

Private Sub Command14_Click()
    '
    Command14.Enabled = False
    '
10  OPX% = COMALTER%("Seleccion de Opciones:\\Generar Nuevo O-F\Modificar O-F\Cerrar O-F")
    If OPX% = 1 And OPX% = 2 And OPX = 3 Then Exit Sub
    Select Case OPX%
        Case 1: Call GENEOFLIM: GoTo 10
        Case 2: Call MODIFOFLIM: GoTo 10
        Case 3: Call CONECRUN("ORFABR07", "Ordenes de Fabricacion")
    End Select
    '
    Command14.Enabled = True
End Sub

Sub GENEOFLIM()
    '
    VDEF$ = ""
5   OBX$ = OBJPLA$("GENOFLIM1", VDEF$)
    If OBX$ = "" Then Exit Sub
    '
    VDEF$ = OBX$
    NPX& = XVALO(CVS(Mid$(OBX$, 1, 4)))
    MANOBRA% = XVALO(CVI(Mid$(OBX$, 5, 2)))
    '
    If NPX& < 1 Then
        Call COMUNI("Falta Seleccionar Orden de Reparacion")
        GoTo 5
    End If
    If MANOBRA% < 1 Then
        Call COMUNI("Falta Ingresar Mano de Obra")
        GoTo 5
    End If
    '
    DL$ = NPX&
    While Len(DL$) < 6: DL$ = "0" + DL$: Wend
    DL$ = "DL-" + DL$ + "-1"
    CONDI$ = "Idsubor='" & DL$ & "'"
    If CantRegistros("ORDEFABR", CONDI$, "NOMESS") > 0 Then
        Call COMUNI("O/F de Limpieza Ya Generada para la O/R Elegida")
        GoTo 99
    End If
    '
    Call SETULTREG("!BOLREDET", 0)
    '
    PRF$ = "INSUFAB"
10  XXX% = VENTABU%(PRF$ + ATRI$)
    '
    If XXX% < 0 Then GoTo 5
    ' Validacion que contenga cantidad para continuar
    For YX& = 1 To ULTREG&("!BOLREDET")
        YY$ = REGLEIDO$("!BOLREDET", YX&)
        Ci% = CVI(Mid$(YY$, 83, 2))
        CAN# = CVD(Mid$(YY$, 89, 8))
        If CAN# < 0.05 Then
            Call COMUNI("Falta Ingresar Cantidad en Codigo '" & CODEXT$(Ci%) & "'")
            GoTo 10
        End If
    Next YX&
    'Grabacion en Tablas

    'ORDEFABR
20  NUMOF& = XVALO(RANDSTRING$(-6))
    If CantRegistros("ORDEFABR", "NUMEORFA = " & NUMOF&, "NOMESS") > 0 Then GoTo 20
    Call CreaRegistro("ORDEFABR", "NUMEORFA", NUMOF&, "NOMESS")
    '
    CONDI$ = "NUMEORFA = " & NUMOF&
    Call ACTUREGISTRO("ORDEFABR", "REFEREN", CONDI$, "O/F de Limpieza para O/Rep: nnn", RAFE&, "NOMESS")
       DL$ = NPX&
       While Len(DL$) < 6: DL$ = "0" + DL$: Wend
       DL$ = "DL-" + DL$ + "-1"
    Call ACTUREGISTRO("ORDEFABR", "IDSUBOR", CONDI$, DL$, RAFE&, "NOMESS")
       FECIE% = HOY(HOS$)
       CODI$ = Control$("REPARA", "CODGENER")
       CIXI% = CODINT(CODI$)
    Call ACTUREGISTRO("ORDEFABR", "FECHGEN", CONDI$, FECIE%, RAFE&, "NOMESS")
    Call ACTUREGISTRO("ORDEFABR", "COD_INTE", CONDI$, CIXI%, RAFE&, "NOMESS")
    Call ACTUREGISTRO("ORDEFABR", "COD_EXTE", CONDI$, CODEXT$(CIXI%), RAFE&, "NOMESS")
    Call ACTUREGISTRO("ORDEFABR", "DESCRIP", CONDI$, "", RAFE&, "NOMESS")
    Call ACTUREGISTRO("ORDEFABR", "UNIDAD", CONDI$, "U.", RAFE&, "NOMESS")
    Call ACTUREGISTRO("ORDEFABR", "CANPROY", CONDI$, 1, RAFE&, "NOMESS")
    Call ACTUREGISTRO("ORDEFABR", "DESTINO", CONDI$, "STOCKS", RAFE&, "NOMESS")
    Call ACTUREGISTRO("ORDEFABR", "FECHENSOL", CONDI$, FECIE%, RAFE&, "NOMESS")
    Call ACTUREGISTRO("ORDEFABR", "PRIORID", CONDI$, 0, RAFE&, "NOMESS")
    Call ACTUREGISTRO("ORDEFABR", "STATUORF", CONDI$, 0, RAFE&, "NOMESS")
    Call ACTUREGISTRO("ORDEFABR", "NUORDPROD", CONDI$, "", RAFE&, "NOMESS")
    Call ACTUREGISTRO("ORDEFABR", "FECHLAN", CONDI$, FECIE%, RAFE&, "NOMESS")
    Call ACTUREGISTRO("ORDEFABR", "SECCIPRO", CONDI$, 0, RAFE&, "NOMESS")
    Call ACTUREGISTRO("ORDEFABR", "CANAPRO", CONDI$, 0, RAFE&, "NOMESS")
    Call ACTUREGISTRO("ORDEFABR", "CANRECH", CONDI$, 0, RAFE&, "NOMESS")
    Call ACTUREGISTRO("ORDEFABR", "NUMEORFA", CONDI$, 0, RAFE&, "NOMESS")
    '
    NORFA$ = DL$
    DESCRI$ = "O-R por Limpieza"
    CANS$ = " - (" + TRIM$(Str$(1)) + ")"
    DESCRI1$ = Left$(DESCRI$, 52 - Len(CANS$)) + CANS$
    ECORDEF$ = AJUSTI$(NORFA$, 12) + DESCRI1$
    '
    Call FILTERPUTRECORD("ECORDEF", 1, NORFA$, ECORDEF$)
    'OPERFAB
30       NUMOF2& = XVALO(RANDSTRING$(-6))
    If CantRegistros("OPERFAB", "DESCRIP = " & NUMOF2&, "NOMESS") > 0 Then GoTo 30
    Call CreaRegistro("OPERFAB", "DESCRIP", NUMOF2&, "NOMESS")
    '
    Condi1$ = "DESCRIP ='" & NUMOF2& & "'"
    Call ACTUREGISTRO("OPERFAB", "IDSUBOR", Condi1$, DL$, RAFE&, "NOMESS")
    Call ACTUREGISTRO("OPERFAB", "COD_INTE", Condi1$, 0, RAFE&, "NOMESS")
    Call ACTUREGISTRO("OPERFAB", "COD_EXTE", Condi1$, "", RAFE&, "NOMESS")
    Call ACTUREGISTRO("OPERFAB", "NUMEOPER", Condi1$, 1, RAFE&, "NOMESS")
    Call ACTUREGISTRO("OPERFAB", "CANTPROG", Condi1$, 1, RAFE&, "NOMESS")
    Call ACTUREGISTRO("OPERFAB", "TIEMVARI", Condi1$, MANOBRA%, RAFE&, "NOMESS")
    Call ACTUREGISTRO("OPERFAB", "DESCRIP", Condi1$, "", RAFE&, "NOMESS")
    'RESERVA
    SQLX$ = "SELECT * FROM RESERVA"
    Dim RESER As ADODB.Recordset
    Set RESER = New ADODB.Recordset
    RESER.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    '
    With RESER
        '
        For YX& = 1 To ULTREG&("!BOLREDET")
             YY$ = REGLEIDO$("!BOLREDET", YX&)
             Ci% = CVI(Mid$(YY$, 83, 2))
             CAN = CVD(Mid$(YY$, 89, 8))
             .AddNew   ' YA CONVERTIDO
             !cod_inte = XVALO(Ci%)
             !CANTRES = XVALO(CAN)
             !UsoUnit = XVALO(CAN)
             !IdSubOr = SAFETEXT$(DL$)
             !CodiEmpr = CodiEmp%
             .Update
        Next YX&
    End With
    RESER.Close
    Set RESER = Nothing
    'REPARA
    CONDI$ = "NUMREPA = " & NPX&
    Call ACTUREGISTRO("REPARA", "IDSUBOR", CONDI$, DL$, RAFE&, "NOMESS")
    '
    Call IMPRIMEOFLIMPIEZA(NPX&)
    '
    MANOBR$ = HORATEX$(MANOBRA%)
    TTXYZ$ = Space$(92)
    NROREP& = NPX&
    HOII% = HOY(HOS$)
    Call REPLA(TTXYZ$, HOS$, 1, 8)
    Call REPLA(TTXYZ$, USERTER$, 69, 24)
    Call REPLA(TTXYZ$, "O/F por Limpieza: " + DL$, 11, 45)
    Call PRESRE09.AGREGA_ANOTAREPARA(NROREP&, TTXYZ$)
    '
    TTXYZ$ = Space$(92)
    Call REPLA(TTXYZ$, Space$(56), 11, 56)
    Call PRESRE09.AGREGA_ANOTAREPARA(NROREP&, TTXYZ$)
    Call REPLA(TTXYZ$, "Mano de Obra: " + TRIM$(MANOBR$), 11, 30)
    Call PRESRE09.AGREGA_ANOTAREPARA(NROREP&, TTXYZ$)
    '
    TTXYZ$ = Space$(92)
    Call REPLA(TTXYZ$, Space$(56), 11, 56)
    Call PRESRE09.AGREGA_ANOTAREPARA(NROREP&, TTXYZ$)
    Call REPLA(TTXYZ$, "Lista de Insumos:", 11, 56)
    Call PRESRE09.AGREGA_ANOTAREPARA(NROREP&, TTXYZ$)
    '
    TTXYZ$ = ""
    For KKI& = 1 To ULTREG&("!BOLREDET")
        XX$ = REGLEIDO$("!BOLREDET", KKI&)
        Ci% = CVI(Mid$(XX$, 83, 2))
        CAN = CVD(Mid$(XX$, 89, 8))
        Call REPLA(TTXYZ$, CODEXT(Ci%), 1, 16)
        Call REPLA(TTXYZ$, DESCRIT0$(Ci%), 13, 32)
        Call REPLA(TTXYZ$, CSTRING$(MKS$(CAN), 3, 6, 1, 2), 50, 52)
        Call PRESRE09.AGREGA_ANOTAREPARA(NROREP&, Space$(10) + TTXYZ$)
    Next KKI&
    '
    Call PRESRE09.AGREGA_ANOTAREPARA(NROREP&, String$(88, "="))
    '
    Call COMUNI("O-F Limpieza '" & DL$ & "' Generada Exitosamente")
    '
    
99 Exit Sub
End Sub

Sub MODIFOFLIM()
    '
    CONDI$ = "REPARA.NuOrdRep = 0 AND ORDEFABR.STATUORF < 3 ORDER BY REPARA.NUMREPA"
    Call CARGALISEL("OFABLIM", "REPARA+ORDEFABR/REPARA.IDSUBOR=ORDEFABR.IDSUBOR", "REPARA.NumRepa/I10;REPARA.Referen/A64;REPARA.IDSUBOR/A12", CONDI$)
3   Call SELECHO("OFABLIM/Ordenes de Reparacion con O/F de Limpieza")
    '
    NPX& = XVALO(VALACT1$("OFABLIM"))
    If NPX& = 0 Then GoTo 99
    '
    DL$ = NPX&
    While Len(DL$) < 6: DL$ = "0" + DL$: Wend
    DL$ = "DL-" + DL$ + "-1"
    '
    CONDI$ = "Idsubor='" & DL$ & "'"
    MANOBRA% = XVALO(VALOBADA("OPERFAB", "TIEMVARI", CONDI$, "NOMESS"))
    '
    VDEF$ = MKS$(NPX&) + MKI$(MANOBRA%)
    If NPX& = 0 Then GoTo 99
5   OBX$ = OBJPLA$("GENOFLIM2", VDEF$)
    If OBX$ = "" Then GoTo 99
'    '
    VDEF$ = OBX$
    NPX& = XVALO(CVS(Mid$(OBX$, 1, 4)))
    MANOBRA% = XVALO(CVI(Mid$(OBX$, 5, 2)))
    '
    If NPX& < 1 Then
        Call COMUNI("Falta Seleccionar Orden de Reparacion")
        GoTo 5
    End If
    If MANOBRA% < 1 Then
        Call COMUNI("Falta Ingresar Mano de Obra")
        GoTo 5
    End If
    '
    Call ACTUREGISTRO("OPERFAB", "TIEMVARI", CONDI$, MANOBRA%, RAFE&, "NOMESS")
    '
    Call SETULTREG("!BOLREDET", 0)
    SQLX$ = "SELECT * FROM RESERVA "
    SQLX$ = SQLX$ + " WHERE Idsubor='" & DL$ & "'"
    Dim REPU As ADODB.Recordset
    Set REPU = FLEXCONN.Execute(FILTSQL$(SQLX$))
    '
    JJ& = 0
    With REPU
      On Error GoTo 0
      'CARGO LOS REPUESTOS
      Do While .EOF = False
         CAN# = XVALO(!CANTRES)
         Ci% = XVALO(!cod_inte)
         XX$ = REGBLAN$("BOLREDET")
         
         Call REPLA(XX$, MKI$(Ci%), 83, 2)
         Call REPLA(XX$, MKD$(CAN#), 89, 8)
         JJ& = JJ& + 1
         Call GRAREG("!BOLREDET", XX$, JJ&)
      .MoveNext
      Loop
    End With
    '
    PRF$ = "INSUFAB"
10      XXX% = VENTABU%(PRF$ + ATRI$)
    '
    If XXX% < 0 Then GoTo 5
    '
    CONDI2$ = "IDSUBOR='" & DL$ & "'"
    Call ACTUREGISTRO("RESERVA", "MARBORRA", CONDI2$, 1, REG&, "NOMESS")
    '
    SQLX$ = "SELECT * FROM RESERVA"
    Dim RESER As ADODB.Recordset
    Set RESER = New ADODB.Recordset
    RESER.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    '
    With RESER
        '
        For YX& = 1 To ULTREG&("!BOLREDET")
             YY$ = REGLEIDO$("!BOLREDET", YX&)
             Ci% = CVI(Mid$(YY$, 83, 2))
             CAN = CVD(Mid$(YY$, 89, 8))
             .AddNew   ' YA CONVERTIDO
             !cod_inte = XVALO(Ci%)
             !CANTRES = XVALO(CAN)
             !UsoUnit = XVALO(CAN)
             !IdSubOr = SAFETEXT$(DL$)
             !CodiEmpr = CodiEmp%
             .Update
        Next YX&
    End With
    RESER.Close
    Set RESER = Nothing
    '
    CONDI3$ = CONDI2$ + " AND MARBORRA = 1"
    Call BORRAREGISTROS("RESERVA", CONDI3$, REG&, "NOMESS")
    '
    Call IMPRIMEOFLIMPIEZA(NPX&)
    '
    MANOBR$ = HORATEX$(MANOBRA%)
    TTXYZ$ = Space$(92)
    NROREP& = NPX&
    HOII% = HOY(HOS$)
    Call REPLA(TTXYZ$, HOS$, 1, 8)
    Call REPLA(TTXYZ$, USERTER$, 69, 24)
    Call REPLA(TTXYZ$, "Modificacion O/F por Limpieza: " + DL$, 11, 45)
    Call PRESRE09.AGREGA_ANOTAREPARA(NROREP&, TTXYZ$)
    '
    TTXYZ$ = Space$(92)
    Call REPLA(TTXYZ$, Space$(56), 11, 16)
    Call PRESRE09.AGREGA_ANOTAREPARA(NROREP&, TTXYZ$)
    Call REPLA(TTXYZ$, "Mano de Obra:" + " " + TRIM$(MANOBR$), 11, 30)
    Call PRESRE09.AGREGA_ANOTAREPARA(NROREP&, TTXYZ$)
    '
    TTXYZ$ = Space$(92)
    Call REPLA(TTXYZ$, Space$(56), 11, 16)
    Call PRESRE09.AGREGA_ANOTAREPARA(NROREP&, TTXYZ$)
    Call REPLA(TTXYZ$, "Lista de Insumos:", 11, 56)
    '
    Call PRESRE09.AGREGA_ANOTAREPARA(NROREP&, TTXYZ$)
    '
    TTXYZ$ = ""
    For KKI& = 1 To ULTREG&("!BOLREDET")
        XX$ = REGLEIDO$("!BOLREDET", KKI&)
        Ci% = CVI(Mid$(XX$, 83, 2))
        CAN = CVD(Mid$(XX$, 89, 8))
        Call REPLA(TTXYZ$, CODEXT(Ci%), 1, 16)
        Call REPLA(TTXYZ$, DESCRIT0$(Ci%), 13, 32)
        Call REPLA(TTXYZ$, CSTRING$(MKS$(CAN), 3, 6, 1, 2), 50, 52)
        Call PRESRE09.AGREGA_ANOTAREPARA(NROREP&, Space$(10) + TTXYZ$)
    Next KKI&
    '
    Call PRESRE09.AGREGA_ANOTAREPARA(NROREP&, String$(88, "="))
    '
    Call COMUNI("O-F Limpieza '" & DL$ & "' Modificada Exitosamente")

99 Exit Sub
End Sub
Sub CIERROFLIM()
   '
   CONDI$ = "IdSubOr <> ''"
   If CantRegistros("REPARA", CONDI$, "NOMESS") < 1 Then
        Call COMUNI("No Hay Ordenes de Reparación en Curso\que Puedan Cerrarse.")
        GoTo 99
   End If
   '
   Call CARGALISEL("OFLIMPIE", "REPARA", "IdSubOr/A12;Referen/A64", CONDI$)
   '
   Call SELECHO("OFLIMPIE/Indice de O/F por Limpieza")
   NOREP$ = VALACT1$("OFLIMPIE")
   If NOREP$ = "" Then GoTo 99
   'Lectura de Materiales Reservados
   '
   Call SELECHO("TADEPOSI")
   DP2% = XVALO(VALACT1$("TADEPOSI"))
   If DP2% < 1 Then Exit Sub
   DEPOCON$ = VALACT2$("TADEPOSI")
   '
   SQLX$ = "SELECT * FROM RESERVA "
   SQLX$ = SQLX$ + " WHERE IdSubOr = '" & UCase(NOREP$) & "' "
   Dim RESER As ADODB.Recordset
   Set RESER = FLEXCONN.Execute(FILTSQL$(SQLX$))
   '
   JJ& = 0
   With RESER
     On Error GoTo 0
     Do While .EOF = False
         CAN# = XVALO(!CANTRES)
         Ci% = XVALO(!cod_inte)
         XX$ = REGBLAN$("BOLREDET")
         '
         Call REPLA(XX$, MKI$(Ci%), 83, 2)
         Call REPLA(XX$, MKD$(CAN#), 89, 8)
         JJ& = JJ& + 1
         Call GRAREG("!BOLREDET", XX$, JJ&)
        .MoveNext
     Loop
   End With
   '
   XXX% = VENTABU%("BOLREDET")
   JG% = 0: FF% = HOY(HO$): NC% = 1
   ' Validacion que contenga cantidad para continuar
   For YX& = 1 To ULTREG&("!BOLREDET")
        YY$ = REGLEIDO$("!BOLREDET", YX&)
        Ci% = CVI(Mid$(YY$, 83, 2))
        CAN# = CVD(Mid$(YY$, 89, 8))
        If CAN# > 0 And Ci% > 0 Then
           JG% = JG% + 1
           Call MOVISTOK(FF%, Ci%, "", "CR", NOREP$, NOREP$, NOREP$, NOREP$, "O-R Consumo de Limpieza", JG%, 0, 1, NC%, DP2%, CAN#)
        End If
    Next YX&
    '

99 End Sub

Sub ANULAFLIM()

End Sub
 
Sub Command15_Click()
   '
   Command15.Enabled = False
'3  Screen.MousePointer = 11
'   Call COPYSTRU("ORDEREPA", "PREPAEMI")
'   JJ& = 0
'   For U& = 1 To ULTREG&("ORDEREPA")
'     XX$ = REGLEIDO$("ORDEREPA", U&)
'     If CVI(Mid$(XX$, 65, 2)) < 3 Then
'       JJ& = JJ& + 1
'       Call GRAREG("PREPAEMI", XX$, JJ&)
'     End If
'   Next U&
'   Call SETULTREG("PREPAEMI", JJ&)
'   Call CIERRARCH("PREPAEMI")
'   '
'   Screen.MousePointer = 1
   '
   CONDI$ = "NuOrdRep > 0 And Status < 3"
   JJ& = CantRegistros("REPARA", CONDI$)
   If JJ& < 1 Then
     Call COMUNI("No Hay Ordenes de Reparación en Curso\que Puedan Cerrarse.")
     GoTo 99
   End If
   '
   CONDI$ = "NuOrdRep > 0 And Status < 3 ORDER BY NuOrdRep"
   Call CARGALISEL("PREPAEM", "REPARA", "NuOrdRep/I10;RefOrdRep/A64", CONDI$)
   '
   Call SELECHO("PREPAEM/Cierre de Orden de Reparacion")
   NOREP& = Val(VALACT1$("PREPAEM"))
   If NOREP& < 1 Then GoTo 99
   '
   NOREX$ = TRIM$(Str$(NOREP&))
   'While Len(NOREX$) < 6: NOREX$ = "0" + NOREX$: Wend
   NOREX$ = TRIM$("OR-" + NOREX$)
   '
   Call SELECHO("TADEPOSI")
   DP2% = XVALO(VALACT1$("TADEPOSI"))
   If DP2% < 1 Then GoTo 99
   DEPOCON$ = VALACT2$("TADEPOSI")
   '
'   VACON% = 1
'   CODIGENERI% = CODINT%(Control$("REPARA", "CODGENER"))
'   COD$ = CODEXT$(CODIGENERI%)
'   Call CARCOMPO(COD$, NOREX$, 1, VACON%, 2)
'   If VACON% < 0 Then GoTo 99
   '
   'BUSCA EL NUMERO DE REPARACION EN BASE AL NUMERO DE RECEPCION.
   CONDI$ = "NuOrdRep = " & NOREP&
   TALONX& = VALOBADA("REPARA", "NumRepa", CONDI$, "NOMESS")
   '
   Call SETULTREG("!MATEROF", 0)
   SQLX$ = "SELECT * FROM REPUREPA"
   SQLX$ = SQLX$ + " WHERE NumRepa = '" & TALONX& & "' AND ORIGEN=1" ' ORIGEN=0 REPUESTOS DE TECNICA Y ORIGEN=1 DE COMERCIAL
   Dim REPU As ADODB.Recordset
   Set REPU = FLEXCONN.Execute(FILTSQL$(SQLX$))
   '
   JJ& = 0
   With REPU
     On Error GoTo 0
     'CARGO LOS REPUESTOS
     Do While .EOF = False
        Ci% = XVALO(!cod_inte)
        CAN# = XVALO(!CANT_SOL)
        XX$ = REGBLAN$("MATEROF")
        '
        Call REPLA(XX$, MKI$(Ci%), 1, 2)
        Call REPLA(XX$, MKS$(CAN#), 5, 4)
        '
        JJ& = JJ& + 1
        Call GRAREG("!MATEROF", XX$, JJ&)
     .MoveNext
     Loop
   End With
   '
   VTB% = VENTABU%("MATEROF")
   '
   If COMALTER%("Confirma Cierre de " + NOREX$ + "\Consumiendo Componentes de Depósito '" + DEPOCON$ + "'\\Si, Cerrar\No, Cancelar") <> 1 Then
      GoTo 99
   End If
   '
   Screen.MousePointer = 11
   CONDI$ = "NuOrdRep = " & NOREP&
   REFERE$ = SAFETEXT$(VALOBADA("REPARA", "RefOrdRep", CONDI$))
   REFOP$ = Left$("(cerr) " + REFERE$, 64)
   Call ACTUREGISTRO("REPARA", "RefOrdRep", CONDI$, REFOP$, REG&)
   Call ACTUREGISTRO("REPARA", "Status", CONDI$, 3, REG&)
   '
'   NUPRESU& = 0
'   For U& = 1 To ULTREG&("ORDEREPA")
'     X$ = REGLEIDO$("ORDEREPA", U&)
'     If CVS(Left$(X$, 4)) = NOREP& Then
'       NUPRESU& = CVS(Mid$(X$, 125, 4))
'       REFERE$ = Mid$(X$, 86, 9)
'       REFOP$ = "(cerr) " + Mid$(X$, 5, 60)
'       CODREPA$ = TRIM$(Mid$(X$, 497, 16))
'       Call REPLA(X$, REFOP$, 5, 60)
'       Call REPLA(X$, MKI$(3), 65, 2) ' MARCA CERRADA
'       Call GRAREG("ORDEREPA", X$, U&)
'       GoTo 20
'     End If
'   Next U&
   '
'20 NOREX$ = TRIM$(Str$(NOREP&))
'   While Len(NOREX$) < 6: NOREX$ = "0" + NOREX$: Wend
'   NOREX$ = "OR-" + NOREX$
'   '
'   RFF$ = RECFILT$("RESERVA", 3, NOREX$)
'   RCC$ = ""
'   URE& = ULTREG&("RESERVA")
'   For KK& = 1 To Len(RFF$) Step 4
'     RGG& = CVS(Mid$(RFF$, KK&, 4))
'     If RGG& > 0 Then
'       If RGG& <= URE& Then
'         XX$ = REGLEIDO$("RESERVA", RGG&)
'         If TRIM$(Mid$(XX$, 3, 12)) = NOREX$ Then
'           Call GRAREG("RESERVA", String$(128, 0), RGG&)
'         End If
'       End If
'     End If
'   Next KK&
   '
   FF% = HOY(HO$)
   '
   'FORMATEO EL NUMERO DE REPARACION SEGUN SE GUARDA EN LA TABLA RESERVA
   NROREPA$ = TRIM$(Str$(NOREP&))
   While Len(NROREPA$) < 6: NROREPA$ = "0" + NROREPA$: Wend
   NROREPA$ = TRIM$("OR-" + NROREPA$)
   'BORRO LA RESERVA
   CONDI$ = "IdSubOr ='" & NROREPA$ & "'"
   Call BORRAREGISTROS("RESERVA", CONDI$, REG&)
   'FIN QUITAR RESERVAS
   '
   'METALES ARCON - CREA VALORES POR DEFAULT DE LOS REGISTROS DEL ENTRECOM
   NROO& = Val(Mid$(NROREPA$, 4))
   RECONSI$ = REGBLAN$("ENTRECOM")
   Call REPLA(RECONSI$, MKI$(FF%), 1, 2)
   Call REPLA(RECONSI$, NROREPA$, 3, 14)
     TIPODO$ = "Cons. Reparacion"
   Call REPLA(RECONSI$, TIPODO$, 17, 16)
     CONDI$ = "NuOrdRep = " & NOREP&
     NC% = XVALO(VALOBADA("REPARA", "Nume_Cli", CONDI$))
   Call REPLA(RECONSI$, MKI$(NC%), 33, 2)
   Call REPLA(RECONSI$, rasocli$(NC%), 35, 30)
   'Call REPLA(RECONSI$, Label8, 65, 32)
   Call REPLA(RECONSI$, MKI$(1), 191, 2)
   Call REPLA(RECONSI$, MKI$(0), 193, 2)
   IDOCUM& = 10000& * (FF% Mod 1000) + 100& * (NROO& Mod 100)
   
20: If CantRegistros("ENTRECOM", "IDOCUM=" & IDOCUM&) > 0 Then
       IDOCUM& = IDOCUM& + 1
       GoTo 20
    End If
   '
   FECIE% = HOY(HOS$)
   NUORIT% = 0
   For U& = 1 To ULTREG&("!MATEROF")
     zz$ = REGLEIDO$("!MATEROF", U&)
     '
     CI2% = CVI(Mid$(zz$, 1, 2))
     CAN# = (-1) * CVS(Mid$(zz$, 5, 4))
     '
     USOUNI# = CAN#
     If CI2% > 0 Then
       If CI2% <= NUMAS% Then
         If Abs(CAN#) >= 0.05 Then
           LOTE$ = ""
           CMOV$ = "RP"
           '
           DOPRI$ = NOREX$
           DOSEC$ = DOPRI$
           REMI$ = DOPRI$
           DEPX% = 1
           NUORIT% = NUORIT% + 1
           REFE$ = "Reparación " + NOREX$
           If Abs(CAN#) >= 0.05 Then
                DEPX% = DP2%: If DP2% = 99 Then DEPX% = DEPOPRE%(CI2%)
                Call MOVISTOK(FF%, CI2%, LOTE$, CMOV$, REMI$, REMI$, DOSEC$, DOSEC$, REFE$, NUORIT%, 0, 0, 0, DP2%, CAN#)
                AA$ = CODEXT$(CI2%)
                '
                If InStr(1, EMPREAC$, "NEUQUEN") > 0 Then
                    'GUARDA ARCHIVO ENTRECOM
                    RTRASY$ = RECONSI$
                    IDOCUM& = IDOCUM& + 1
                    Call REPLA(RTRASY$, MKS$(IDOCUM&), 97, 4)
                    Call REPLA(RTRASY$, Space(16), 101, 16)
                    Call REPLA(RTRASY$, MKI$(CI2%), 117, 2)
                    Call REPLA(RTRASY$, CODEXT$(CI2%), 119, 16)
                    Call REPLA(RTRASY$, DESCRIT0$(CI2%), 135, 48)
                    CAN# = CAN# * (-1)
                    Call REPLA(RTRASY$, MKD$(CAN#), 183, 8)
                    Call REGAPP("ENTRECOM", RTRASY$)
                    '
                    XX$ = RTRASY$
                    FF% = CVI(Mid$(XX$, 1, 2))
                    REMI$ = Mid$(XX$, 3, 14)
                    TIPODO$ = Mid$(XX$, 17, 16)
                    NC% = CVI(Mid$(XX$, 33, 2))
                    IDOCUM& = CVS(Mid$(XX$, 97, 4))
                    OCOMPRA$ = Mid$(XX$, 101, 16)
                    Ci% = CVI(Mid$(XX$, 117, 2))
                    Codigo$ = Mid$(XX$, 119, 16)
                    '
                    SQLX$ = "SELECT * FROM ENTRECOM WHERE IDOCUM < 1"
                    Dim RST1 As ADODB.Recordset
                    Set RST1 = New ADODB.Recordset
26:                 RST1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
                    If Err Then
                       On Error GoTo 0
                       Call CapturaErrorOpen
                       GoTo 26
                    End If
                    On Error GoTo 0
                    '
                    With RST1
                        .AddNew
                        !FECHA = CVDAT(FF%)
                        !documento = REMI$
                        !TIPODOC = TIPODO$
                        !nume_cli = NC%
                        !IDOCUM = IDOCUM&
                        !cod_inte = Ci%
                        !Cod_Exte = Codigo$
                        !Cantidad = CAN#
                        !Status = 1
                        !VECESENVIADO = 0
                        .Update
                    End With
                    '
                    On Error Resume Next
                    RST1.Close
                    Set RST1 = Nothing
                    On Error GoTo 0
                End If
           End If
         End If
       End If
     End If
   Next U&
   '
'   Call CIERRARCH("MOVISTO")
'   Call CIERRARCH("PRESREPA")
'   Call CIERRARCH("ORDEREPA")
'   Call CIERRARCH("RESERVA")
   '
   'ANOTACION EN EL ANOTADOR DE CIERRE DE O.R.
   TTXYZ$ = Space$(92)
   HOII% = HOY(HOS$)
   Call REPLA(TTXYZ$, HOS$, 1, 8)
   Call REPLA(TTXYZ$, USERTER$, 69, 24)
   Call REPLA(TTXYZ$, "Generación de Cierre de Reparación", 11, 56)
   Call PRESRE09.AGREGA_ANOTAREPARA(TALONX&, TTXYZ$)
   Call PRESRE09.AGREGA_ANOTAREPARA(TALONX&, String$(88, "="))

   
   Call CONSUMO_MAT_SUM_CLIENTE(TALONX&)
   
   Call COMUNI("Cierre " + NOREX$ + " Completo")
   
   Screen.MousePointer = 1

99 Command15.Enabled = True

End Sub
Private Sub Command16_Click()
    '
    Command16.Enabled = False
    CONDI$ = "NuOrdRep > 0 And Status = 3"
    JJ& = CantRegistros("REPARA", CONDI$)
    If JJ& < 1 Then
       Call COMUNI("No Hay Orden de Reparación para completar la Lista.")
       GoTo 99
    End If
    '
    CONDI$ = "NuOrdRep > 0 And Status = 3 ORDER BY NuOrdRep"
    Call CARGALISEL("PREPAEMI", "REPARA", "NuOrdRep/I10;RefOrdRep/A128", CONDI$)
    Call FRESHECHO("PREPAEMI")
    '
    Call SELECHO("PREPAEMI/Indice de Reparaciones Pendientes")
    NUOREP& = XVALO(VALACT1$("PREPAEMI"))
    FF$ = FECHATEX$(HOY(HO$))
    If NUOREP& < 1 Then GoTo 99
    '
    FILTX$ = TRIM$(Str$(NUOREP&)) & ";" & TRIM$(FF$)
    Call STDFORM("LIST-ACN", FILTX$)
    '
99  Command16.Enabled = True
End Sub

Private Sub Command17_Click()
    '
    Command17.Enabled = False
    FICHREP09.Show 1
    Command17.Enabled = True
    '
End Sub

Private Sub Command170_Click()
    Command170.Enabled = False
    CONDI$ = "CODIMOVI = 'RT' and COD_EXTE = 'REPAGEN' AND LEFT(DOPRILAR,9) = 'RT-X-0001'  ORDER BY DOPRILAR"
    Call CARGALISEL("LISTREMI", "MOVISTO", "DOPRILAR/A18;FechMov/D8;REFERCOR/A24", CONDI$, "NOMESS/DISTINCT")
15  Call SELECHO("LISTREMI")
    NUDOCU$ = TRIM$(Mid$(VALACT1$("LISTREMI"), 6))
    If NUDOCU$ = "" Then GoTo 99
    TIDOCUM$ = "Remito"
    nubus = XVALO(Mid$(NUDOCU$, 6))
    '
    Screen.MousePointer = 11
     '
    DOCUNU& = 0
    Call MUESTRADOC(DOCUNU&, "", TIDOCUM$, Str(nubus), NUDOCU$, "NOMESS")
    Screen.MousePointer = 1
    If DOCUNU& > 0 Then GoTo 15
    
99   Command170.Enabled = True

End Sub


Private Sub Command18_Click()
    Command18.Enabled = False
    '
    If CLACONTRA% < 1 Then Exit Sub
    '
    SQLX$ = "SELECT * FROM REPARA WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE STATUS = 3 "
    SQLX$ = SQLX$ + " AND NUORDREP > 0 "
    Set REPATMP = READSET(SQLX$)
    
    With REPATMP
        Do While Not .EOF
           NREPARA$ = TRIM$(Str$(XVALO(!NUORDREP)))
           While Len(NREPARA$) < 6: NREPARA$ = "0" + NREPARA$: Wend
           NREPARA$ = "OR-" + NREPARA$
           CONDI$ = "IdSubOr ='" & NREPARA$ & "'"
           Call BORRAREGISTROS("RESERVA", CONDI$, REG&)
          .MoveNext
        Loop
    End With
    REPATMP.Close
    Set REPATMP = Nothing
    Call COMUNI("PROCESO FINALIZADO")
    Command18.Enabled = True
End Sub

Private Sub Command19_Click()

    Command9.Enabled = False
    
    Respuesta% = COMALTER%("Opciones en Anulación:\\Anular Recepción\Revertir Anulación")
    If Respuesta% = 1 Then
        CondicionSql$ = "NuOrdRep = 0 AND STATUS < 9 ORDER BY NUMREPA"
        Call CARGALISEL("RECEQUIP", "REPARA", "NumRepa/I10;Referen/A128", CondicionSql$)
        Call FRESHECHO("RECEQUIP")
        Call SELECHO("RECEQUIP/Recepciones de Equipos a Reparar")
        '
        NumeroReparacion& = XVALO(VALACT1$("RECEQUIP"))
        If NumeroReparacion& > 0 Then
           Respuesta% = COMALTER%("¿Está Seguro que Quiere Eliminar la Reparación '" & CStr(NumeroReparacion&) & "?'\\No, Cancelar\Si, Continuar")
           If Respuesta% = 2 Then
             CondicionSql$ = "NUMREPA=" & CStr(NumeroReparacion&)
             Call ACTUREGISTRO("REPARA", "STATUS", CondicionSql$, 9, REGAF&)
             Call COMUNI("Reparación '" & CStr(NumeroReparacion&) & "' Anulada")
           End If
        End If
    Else
        CondicionSql$ = "STATUS = 9 ORDER BY NUMREPA"
        Call CARGALISEL("EQUIANU", "REPARA", "NumRepa/I10;Referen/A128", CondicionSql$)
        Call FRESHECHO("EQUIANU")
        Call SELECHO("EQUIANU/Recepciones de Equipos Anuladas")
        '
        NumeroReparacion& = XVALO(VALACT1$("EQUIANU"))
        If NumeroReparacion& > 0 Then
            CondicionSql$ = "NUMREPA=" & CStr(NumeroReparacion&)
            Call ACTUREGISTRO("REPARA", "STATUS", CondicionSql$, 0, REGAF&)
            Call COMUNI("Reparación '" & CStr(NumeroReparacion&) & "' Activa")
        End If
    End If
    
    Command9.Enabled = True
    
End Sub

Private Sub Command2_Click()
   Command2.Enabled = False
   '
   CONDI$ = "NuOrdRep > 0 And Status = 3"
   JJ& = CantRegistros("REPARA", CONDI$)
   If JJ& < 1 Then
     Call COMUNI("No Hay Presupuestos Pendientes\para Emitir Orden de Reparación.")
     GoTo 99
   End If
   '
   CONDI$ = "NuOrdRep > 0 And Status = 3 ORDER BY NuOrdRep"
   Call CARGALISEL("PREPAEMI", "REPARA", "NuOrdRep/I10;RefOrdRep/A128", CONDI$)
   Call FRESHECHO("PREPAEMI")
   
   Call SELECHO("PREPAEMI/Indice de Reparaciones Pendientes")
   NUOREP& = Val(VALACT1$("PREPAEMI"))
   If NUOREP& < 1 Then GoTo 99
   NPX& = XVALO(VALOBADA("REPARA", "NumRepa", "NuOrdRep = " & NUOREP&))
   '
   FOREMI09.REPRESUP = NUOREP&
   CONDI$ = "NumRepa =" & NPX& & ""
   TTXX$ = VALOBADA("REPARA", "NumeroSerie", CONDI$, "NOMESS")
   TTXY$ = VALOBADA("REPARA", "Referen2", CONDI$, "NOMESS")
   SEREF$ = "S/No. " + TTXX$ + " - Ref: " + TTXY$
   NCLIE% = VALOBADA("REPARA", "Nume_Cli", CONDI$, "NOMESS")
   FOREMI09.Text1 = TRIM$(Str$(NCLIE%))
   FOREMI09.Text1.TEXT = TRIM$(Str$(NCLIE%))
   FOREMI09.Text2.TEXT = rasocli$(NCLIE%)
   FOREMI09.Text3.TEXT = DOMICLI$(NCLIE%)
   FOREMI09.Text4.TEXT = CPOSCLI$(NCLIE%)
   FOREMI09.Text5.TEXT = LOCACLI$(NCLIE%)
   Call CARTRANSPOR
'  FOREMI09.Text22 = TRIM$(Mid$(XX$, 129, 368)) + Chr$(13) + Chr$(10) + SEREF$
   FOREMI09.Text22 = TRIM$(VALOBADA("REPARA", "Observa", CONDI$, "NOMESS")) + Chr$(13) + Chr$(10) + SEREF$
   FOREMI09.Show 1
   '
99 Command2.Enabled = True
End Sub

Private Sub Command20_Click()
  Command20.Enabled = False

   '
   'CONDI$ = "NuOrdRep > 0 And Status < 3"
   CONDI$ = "NuOrdRep > 0 "
   JJ& = CantRegistros("REPARA", CONDI$)
   If JJ& < 1 Then
     Call COMUNI("No Hay Ordenes de Reparación en Curso\que Puedan Cerrarse.")
     GoTo 99
   End If
   '
   'CONDI$ = "NuOrdRep > 0 And Status < 3 ORDER BY NuOrdRep"
   CONDI$ = "NuOrdRep > 0  ORDER BY NuOrdRep"
   Call CARGALISEL("PREPAEM", "REPARA", "NuOrdRep/I10;RefOrdRep/A64", CONDI$)
   '
   Call SELECHO("PREPAEM/Cierre de Orden de Reparacion")
   NOREP& = Val(VALACT1$("PREPAEM"))
   If NOREP& < 1 Then GoTo 99
   '
   NOREX$ = TRIM$(Str$(NOREP&))
   'While Len(NOREX$) < 6: NOREX$ = "0" + NOREX$: Wend
   NOREX$ = TRIM$("OR-" + NOREX$)
   '
   Call SELECHO("TADEPOSI")
   DP2% = XVALO(VALACT1$("TADEPOSI"))
   If DP2% < 1 Then GoTo 99
   DEPOCON$ = VALACT2$("TADEPOSI")
   '

   MsgBox TALONX&
   
   NPX& = XVALO(InputBox$("NUMERO DE TALON (NO el de Reparacion, EL QUE FIGURA EN LA FICHA)", "Ingrese el N° TALON"))
   If NPX& < 1 Then Exit Sub
   TALONX& = NPX&
   
   Call SETULTREG("!MATEROF", 0)
   SQLX$ = "SELECT * FROM REPUREPA"
   SQLX$ = SQLX$ + " WHERE NumRepa = '" & TALONX& & "' AND ORIGEN=0" ' ORIGEN=0 REPUESTOS DE TECNICA Y ORIGEN=1 DE COMERCIAL
   Dim REPU As ADODB.Recordset
   Set REPU = FLEXCONN.Execute(FILTSQL$(SQLX$))
   '
   JJ& = 0
   With REPU
     On Error GoTo 0
     'CARGO LOS REPUESTOS
     Do While .EOF = False
        Ci% = XVALO(!cod_inte)
        CAN# = XVALO(!CANT_SOL)
        XX$ = REGBLAN$("MATEROF")
        '
        Call REPLA(XX$, MKI$(Ci%), 1, 2)
        Call REPLA(XX$, MKS$(CAN#), 5, 4)
        '
        JJ& = JJ& + 1
        Call GRAREG("!MATEROF", XX$, JJ&)
     .MoveNext
     Loop
   End With
   '
'   NOREP1$ = TRIM$(Str$(NOREP&))
'   While Len(NOREP1$) < 6: NOREP1$ = "0" + NOREP1$: Wend
'   NOREP1$ = "OR-" + NOREP1$
'   Call GENERARESERVA_VERSION2(NOREP1$)
   '
   VTB% = VENTABU%("MATEROF")
   '
   If COMALTER%("Confirma Cierre de " + NOREX$ + "\Consumiendo Componentes de Depósito '" + DEPOCON$ + "'\\Si, Cerrar\No, Cancelar") <> 1 Then
      GoTo 99
   End If
   '
   Screen.MousePointer = 11
   CONDI$ = "NuOrdRep = " & NOREP&
   REFERE$ = SAFETEXT$(VALOBADA("REPARA", "RefOrdRep", CONDI$))
   REFOP$ = Left$("(cerr) " + REFERE$, 64)
   Call ACTUREGISTRO("REPARA", "RefOrdRep", CONDI$, REFOP$, REG&)
   Call ACTUREGISTRO("REPARA", "Status", CONDI$, 3, REG&)
   '

   '
   FF% = HOY(HO$)
   '
   'FORMATEO EL NUMERO DE REPARACION SEGUN SE GUARDA EN LA TABLA RESERVA
   NROREPA$ = TRIM$(Str$(NOREP&))
   While Len(NROREPA$) < 6: NROREPA$ = "0" + NROREPA$: Wend
   NROREPA$ = TRIM$("OR-" + NROREPA$)
   'BORRO LA RESERVA
   CONDI$ = "IdSubOr ='" & NROREPA$ & "'"
   Call BORRAREGISTROS("RESERVA", CONDI$, REG&)
   'FIN QUITAR RESERVAS
   '
   'METALES ARCON - CREA VALORES POR DEFAULT DE LOS REGISTROS DEL ENTRECOM
   NROO& = Val(Mid$(NROREPA$, 4))
   RECONSI$ = REGBLAN$("ENTRECOM")
   Call REPLA(RECONSI$, MKI$(FF%), 1, 2)
   Call REPLA(RECONSI$, NROREPA$, 3, 14)
     TIPODO$ = "Cons. Reparacion"
   Call REPLA(RECONSI$, TIPODO$, 17, 16)
     CONDI$ = "NuOrdRep = " & NOREP&
     NC% = XVALO(VALOBADA("REPARA", "Nume_Cli", CONDI$))
   Call REPLA(RECONSI$, MKI$(NC%), 33, 2)
   Call REPLA(RECONSI$, rasocli$(NC%), 35, 30)
   'Call REPLA(RECONSI$, Label8, 65, 32)
   Call REPLA(RECONSI$, MKI$(1), 191, 2)
   Call REPLA(RECONSI$, MKI$(0), 193, 2)
   IDOCUM& = 10000& * (FF% Mod 1000) + 100& * (NROO& Mod 100)
   
20: If CantRegistros("ENTRECOM", "IDOCUM=" & IDOCUM&) > 0 Then
       IDOCUM& = IDOCUM& + 1
       GoTo 20
    End If
   '
   FECIE% = HOY(HOS$)
   NUORIT% = 0
   For U& = 1 To ULTREG&("!MATEROF")
     zz$ = REGLEIDO$("!MATEROF", U&)
     '
     CI2% = CVI(Mid$(zz$, 1, 2))
     CAN# = (-1) * CVS(Mid$(zz$, 5, 4))
     '
     USOUNI# = CAN#
     If CI2% > 0 Then
       If CI2% <= NUMAS% Then
         If Abs(CAN#) >= 0.05 Then
           LOTE$ = ""
           CMOV$ = "RP"
           '
           DOPRI$ = NOREX$
           DOSEC$ = DOPRI$
           REMI$ = DOPRI$
           DEPX% = 1
           NUORIT% = NUORIT% + 1
           REFE$ = "Reparación " + NOREX$
           If Abs(CAN#) >= 0.05 Then
                DEPX% = DP2%: If DP2% = 99 Then DEPX% = DEPOPRE%(CI2%)
                Call MOVISTOK(FF%, CI2%, LOTE$, CMOV$, REMI$, REMI$, DOSEC$, DOSEC$, REFE$, NUORIT%, 0, 0, 0, DP2%, CAN#)
                AA$ = CODEXT$(CI2%)
                '
                If InStr(1, EMPREAC$, "NEUQUEN") > 0 Then
                    'GUARDA ARCHIVO ENTRECOM
                    RTRASY$ = RECONSI$
                    IDOCUM& = IDOCUM& + 1
                    Call REPLA(RTRASY$, MKS$(IDOCUM&), 97, 4)
                    Call REPLA(RTRASY$, Space(16), 101, 16)
                    Call REPLA(RTRASY$, MKI$(CI2%), 117, 2)
                    Call REPLA(RTRASY$, CODEXT$(CI2%), 119, 16)
                    Call REPLA(RTRASY$, DESCRIT0$(CI2%), 135, 48)
                    CAN# = CAN# * (-1)
                    Call REPLA(RTRASY$, MKD$(CAN#), 183, 8)
                    Call REGAPP("ENTRECOM", RTRASY$)
                    '
                    XX$ = RTRASY$
                    FF% = CVI(Mid$(XX$, 1, 2))
                    REMI$ = Mid$(XX$, 3, 14)
                    TIPODO$ = Mid$(XX$, 17, 16)
                    NC% = CVI(Mid$(XX$, 33, 2))
                    IDOCUM& = CVS(Mid$(XX$, 97, 4))
                    OCOMPRA$ = Mid$(XX$, 101, 16)
                    Ci% = CVI(Mid$(XX$, 117, 2))
                    Codigo$ = Mid$(XX$, 119, 16)
                    '
                    SQLX$ = "SELECT * FROM ENTRECOM WHERE IDOCUM < 1"
                    Dim RST1 As ADODB.Recordset
                    Set RST1 = New ADODB.Recordset
26:                 RST1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
                    If Err Then
                       On Error GoTo 0
                       Call CapturaErrorOpen
                       GoTo 26
                    End If
                    On Error GoTo 0
                    '
                    With RST1
                        .AddNew
                        !FECHA = CVDAT(FF%)
                        !documento = REMI$
                        !TIPODOC = TIPODO$
                        !nume_cli = NC%
                        !IDOCUM = IDOCUM&
                        !cod_inte = Ci%
                        !Cod_Exte = Codigo$
                        !Cantidad = CAN#
                        !Status = 1
                        !VECESENVIADO = 0
                        .Update
                    End With
                    '
                    On Error Resume Next
                    RST1.Close
                    Set RST1 = Nothing
                    On Error GoTo 0
                End If
           End If
         End If
       End If
     End If
   Next U&
   '
'   Call CIERRARCH("MOVISTO")
'   Call CIERRARCH("PRESREPA")
'   Call CIERRARCH("ORDEREPA")
'   Call CIERRARCH("RESERVA")
   '
   'ANOTACION EN EL ANOTADOR DE CIERRE DE O.R.
   TTXYZ$ = Space$(92)
   HOII% = HOY(HOS$)
   Call REPLA(TTXYZ$, HOS$, 1, 8)
   Call REPLA(TTXYZ$, USERTER$, 69, 24)
   Call REPLA(TTXYZ$, "Consumo Forzado de Cierre de Reparación", 11, 56)
   Call PRESRE09.AGREGA_ANOTAREPARA(TALONX&, TTXYZ$)
   Call PRESRE09.AGREGA_ANOTAREPARA(TALONX&, String$(88, "="))

   
   'Call CONSUMO_MAT_SUM_CLIENTE(TALONX&)
   
   Call COMUNI("Cierre " + NOREX$ + " Completo")
   
   Screen.MousePointer = 1

99 Command20.Enabled = True
End Sub

'Sub GENERARESERVA_VERSION2(NROREP$)
'    '
'    List1.Clear
'    Call SETULTREG("!BOLREDET", 0)
'    NPX& = XVALO(InputBox$("NUMERO DE TALON", "Ingrese el N° TALON"))
'    If NPX& < 1 Then Exit Sub
'
'    NPX& = XVALO(NPX&)
'    IMGTOT% = CantRegistros("REPUREPA", "NUMREPA = '" & UCase(NPX&) & "'")
'    If IMGTOT% = 0 Then
'      GoTo 99
'    End If
'
'    SQLX$ = "SELECT * FROM REPUREPA "
'    SQLX$ = SQLX$ + " WHERE NUMREPA = '" & UCase(NPX&) & "' AND ORIGEN=0"
'    Dim REPU As ADODB.Recordset
'    Set REPU = FLEXCONN.Execute(FILTSQL$(SQLX$))
'    '
'    JJ& = 0
'    With REPU
'      On Error GoTo 0
'      'CARGO LOS REPUESTOS
'      Do While .EOF = False
'         CAN# = XVALO(!CANT_SOL)
'         Ci% = XVALO(!cod_inte)
'         XX$ = REGBLAN$("BOLREDET")
'
'         Call REPLA(XX$, MKI$(Ci%), 83, 2)
'         Call REPLA(XX$, MKD$(CAN#), 89, 8)
'         JJ& = JJ& + 1
'         Call GRAREG("!BOLREDET", XX$, JJ&)
'      .MoveNext
'      Loop
'      List1.Refresh
'    End With
'
'    Call CARGALISTAREPVERSION2
'
'
'
'
'    Call BORRAREGISTROS("RESERVA", "IDSUBOR = '" & NROREP$ & "'", REGAF&, "NOMESS")
'
'    SQLX1$ = "SELECT * FROM RESERVA"
'    SQLX1$ = SQLX1$ + " WHERE COD_INTE < 1 "
'    Dim RESETMP As ADODB.Recordset
'    Set RESETMP = New ADODB.Recordset
'    RESETMP.Open FILTSQL$(SQLX1$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
'    With RESETMP
'      KK& = 0
'      For UI% = 1 To List1.ListCount
'         zz$ = List1.List(UI% - 1)
'         CODD$ = TRIM$(Left$(zz$, 16))
'         CIIX% = CODINT%(CODD$)
'         If CIIX% > 0 Then
'           If CIIX% <= NUMAS% Then
'             CANTX# = Val(Mid$(zz$, 59, 10))
'             .AddNew
'             On Error Resume Next
'             !CODIEMPR = CodiEmp%
'             On Error GoTo 0
'             !cod_inte = CIIX%
'             !Cod_Exte = CODEXT$(CIIX%)
'             !Descrip = DESCRIT$(CIIX%)
'             !IdSubOr = NROREP$
'             !FechRes = Now
'             !CANTRES = CANTX#
'             !CantCons = 0
'             !Saldo_Entre = CANTX#
'             !IDENLOTE = ""
'             !DEPOSITO = 0
'             !UsoUnit = CANTX#
'                KK& = KK& + 1
'             !NuOrItem = KK&
'             .Update
'           End If
'         End If
'      Next UI%
'   End With
'   '
'   On Error Resume Next
'   RESETMP.Close
'   Set RESETMP = Nothing
'   On Error GoTo 0
'   '
'99
'End Sub
'
Sub CARGALISTAREPVERSION2()
   List1.Clear
   For YX& = 1 To ULTREG&("!BOLREDET")
     YY$ = REGLEIDO$("!BOLREDET", YX&)
     Ci% = CVI(Mid$(YY$, 83, 2))
     CAN = CVD(Mid$(YY$, 89, 8))
     Call REPLA(TTXX$, CODEXT$(Ci%), 1, 16)
     Call REPLA(TTXX$, DESCRIT0$(Ci%), 14, 52)
     Call REPLA(TTXX$, MKS$(CAN), 59, 10)
     List1.AddItem TTXX$
   Next YX&
End Sub

Private Sub Command3_Click()
    '
    Command3.Enabled = False
    '
    If DERACCE%("VEPREPAR", "Visualizar Presupuestos de Reparación") < 2 Then
         GoTo 99
    End If

3   CONDI$ = "Nume_Cli > 0 ORDER BY NUMREPA"
    Call CARGALISEL("PREPAEMI", "REPARA", "NumRepa/I10;Referen/A128", CONDI$)
51  Call SELECHO("PREPAEMI/Indice Presupuestos de Reparación Pendientes")
    NPX& = Val(VALACT1$("PREPAEMI"))
    If NPX& < 1 Then GoTo 99
    '
    NUDOCU$ = SAFETEXT$(NPX&)
    OPX% = COMALTER%("Documentos a Visualizar\\Ver Presupuesto de Reparación Interno\Ver Presupuesto de Reparación Externo")
    If OPX% < 1 Then GoTo 51
    If OPX% = 1 Then TIDOCUM$ = TRIM$("Presupuesto de Reparación " + NUDOCU$)
    If OPX% = 2 Then TIDOCUM$ = TRIM$("Presup.Reparación (Externo) " + NUDOCU$)
  
     Call MUESTRADOC(DOCUNU&, "", TIDOCUM$, NUDOCU$, NUDOCU$, "NOMESS")
     GoTo 99
'    NUDOCU$ = TRIM$(Str$(NPX&))
'    For UI& = ULTREG&("PRESREPA") To 1 Step -1
'      XX$ = REGLEIDO$("PRESREPA", UI&)
'      If CVS(Left$(XX$, 4)) = NPX& Then GoTo 4
'    Next UI&
'    Call COMUNI("Error de Indices - Consulte.")
'    GoTo 99
    '
4   NUDOCU$ = TRIM$(Str$(NPX&))
    Screen.MousePointer = 11
'    TIDOCUM$ = "Presupuesto de Reparación " + NUDOCU$
'    TIDOCU2$ = "Operaciones de Reparación PR-" + NUDOCU$
    Screen.MousePointer = 11
    '
    DOCUNU& = 0
    For UI& = ULTREG&("PDOCLST") To 1 Step -1
      XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
      If InStr(XX$, TIDOCUM$) > 4 Then
        DOCUNU& = CVS(Left$(XX$, 4))
        GoTo 5
      End If
    Next UI&
    '
    Screen.MousePointer = 1
    Call COMUNI("No se Ha Encontrado el Documento Impreso")
    GoTo 99
    '
5   Call MUESTRADOC(DOCUNU&)
'5   For KKI% = 0 To 5
'      EDITFORM.Picture1(KKI%).Cls
'      EDITFORM.Picture1(KKI%).DrawWidth = 3
'      EDITFORM.Picture1(KKI%).DrawStyle = 0
'      EDITFORM.Picture1(KKI%).FillStyle = 1
'    Next KKI%
'    '
'    REDOCU$ = REGACT$("PDOCLST")
'    URE& = ULTREG&("PDOCSPL")
'    LI& = 0: LS& = URE&
'6   E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 7
'    L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 7
'    DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
'    If DCC& < DOCUNU& Then LI& = L&: GoTo 6
'    If DCC& > DOCUNU& Then LS& = L&: GoTo 6
'    GoTo 8
'    '
'7   Screen.MousePointer = 1
'    Call MENSERR(24, "Documento no Registrado")
'    Exit Sub
'    '
'8   TPSP$ = ""
'    Screen.MousePointer = 11
'    While L& > 1
'      DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
'      If DCC& < DOCUNU& Then GoTo 9
'      L& = L& - 1
'    Wend
'    '
'9   CAPAGINAS% = 0
'    For KKL& = L& To URE&
'      TBUF$ = REGLEIDO$("PDOCSPL", KKL&)
'      If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 10
'      If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
'        TPSP$ = TPSP$ + Mid$(TBUF$, 5)
'      End If
'    Next KKL&
'    '
'10  For UI& = ULTREG&("PDOCLST") To 1 Step -1
'      XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
'      If InStr(XX$, TIDOCU2$) > 4 Then
'        DOCUNU& = CVS(Left$(XX$, 4))
'        GoTo 15
'      End If
'    Next UI&
'    '
'15  REDOCU$ = REGACT$("PDOCLST")
'    URE& = ULTREG&("PDOCSPL")
'    LI& = 0: LS& = URE&
'16  E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 17
'    L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 17
'    DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
'    If DCC& < DOCUNU& Then LI& = L&: GoTo 16
'    If DCC& > DOCUNU& Then LS& = L&: GoTo 16
'    GoTo 18
'    '
'17  Screen.MousePointer = 1
'    Call MENSERR(24, "Documento no Registrado")
'    Exit Sub
'    '
'18  While L& > 1
'      DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
'      If DCC& < DOCUNU& Then GoTo 19
'      L& = L& - 1
'    Wend
'    '
'19  For KKL& = L& To URE&
'      TBUF$ = REGLEIDO$("PDOCSPL", KKL&)
'      If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 95
'      If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
'        TPSP$ = TPSP$ + Mid$(TBUF$, 5)
'      End If
'    Next KKL&
'    '
'95  SPOOLSTRING$ = TPSP$
'    For KKI% = 0 To 5
'      EDITFORM.Picture1(KKI%).Cls
'      EDITFORM.Picture1(KKI%).Top = 0
'      EDITFORM.Picture1(KKI%).Refresh
'      EDITFORM.Picture1(KKI%).Height = 7170
'      EDITFORM.Picture1(KKI%).Width = 11700
'      EDITFORM.Picture1(KKI%).DrawWidth = 3
'      EDITFORM.Picture1(KKI%).DrawStyle = 0
'      EDITFORM.Picture1(KKI%).FillStyle = 1
'    Next KKI%
'    ALTUPAGINA = 7170
'    TOPEPAGINA = 0
'    PAGINACTIVA% = 0
'    CAPAGINAS% = 0
'    '
'50  PPXX% = InStr(TPSP$, Chr$(255))
'    If PPXX% > 0 Then
'       COPRIAT$ = Left$(TPSP$, PPXX% - 1)
'       TPSP$ = Mid$(TPSP$, PPXX% + 1)
'       '
'       If Len(COPRIAT$) > 0 Then
'         COPRI% = Asc(Left$(COPRIAT$, 1))
'         ATRIP$ = Mid$(COPRIAT$, 2)
'         Call SHOWCOMMAND(COPRI%, ATRIP$)
'       End If
'       GoTo 50
'       '
'    End If
'    '
'    Screen.MousePointer = 1
'    PAGINACTIVA% = 0
'    For KKI% = 0 To 5
'      EDITFORM.Picture1(KKI%).Visible = False
'      If KKI% = PAGINACTIVA% Then EDITFORM.Picture1(KKI%).Visible = True
'    Next KKI%
'    EDITFORM.Picture1(PAGINACTIVA%).Refresh
'    EDITFORM.VScroll1.Min = 0
'    EDITFORM.VScroll1.Value = 0
'    EDITFORM.VScroll1.Max = 1
'    If ALTUPAGINA > EDITFORM.Frame1.Height Then
'       EDITFORM.VScroll1.Max = (ALTUPAGINA - EDITFORM.Frame1.Height) / 100
'    End If
'    EDITFORM.Command3.Enabled = False
''    EDITFORM.Command5.Enabled = False
'    EDITFORM.Show 1
    '
    GoTo 51
    '
99  Command3.Enabled = True
    '
End Sub

Private Sub Command4_Click()
    '
    Command4.Enabled = False
    '
3   CONDI$ = "NuOrdRep > 0 ORDER BY NuOrdRep"
    Call CARGALISEL("PREPAEMI", "REPARA", "NuOrdRep/I10;RefOrdRep/A128", CONDI$)
    Call SELECHO("PREPAEMI/Indice Presupuestos de Reparación Pendientes")
    NPX& = Val(VALACT1$("PREPAEMI"))
    If NPX& < 1 Then GoTo 99
    '
4   NUDOCU$ = TRIM$(Str$(NPX&))
    TIDOCUM$ = "Orden de Reparación " + NUDOCU$
    TIDOCU2$ = "Procedimiento de Reparación OR-" + NUDOCU$
    Screen.MousePointer = 11
    '
    For UI& = ULTREG&("PDOCLST") To 1 Step -1
      XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
      If InStr(XX$, TIDOCUM$) > 4 Then
        DOCUNU& = CVS(Left$(XX$, 4))
        GoTo 5
      End If
    Next UI&
    '
    Screen.MousePointer = 1
    Call COMUNI("No se Encontró el Documento Impreso")
    On Error Resume Next
    GoTo 99
    '
5   Call MUESTRADOC(DOCUNU&)
'5   For KKI% = 0 To 5
'      EDITFORM.Picture1(KKI%).Cls
'      EDITFORM.Picture1(KKI%).DrawWidth = 3
'      EDITFORM.Picture1(KKI%).DrawStyle = 0
'      EDITFORM.Picture1(KKI%).FillStyle = 1
'    Next KKI%
'    '
'    REDOCU$ = REGACT$("PDOCLST")
'    URE& = ULTREG&("PDOCSPL")
'    LI& = 0: LS& = URE&
'6   E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 7
'    L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 7
'    DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
'    If DCC& < DOCUNU& Then LI& = L&: GoTo 6
'    If DCC& > DOCUNU& Then LS& = L&: GoTo 6
'    GoTo 8
'    '
'7   Screen.MousePointer = 1
'    Call MENSERR(24, "Documento no Registrado")
'    GoTo 99
'    '
'8   TPSP$ = ""
'    Screen.MousePointer = 11
'    While L& > 1
'      DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
'      If DCC& < DOCUNU& Then GoTo 9
'      L& = L& - 1
'    Wend
'    '
'9   CAPAGINAS% = 0
'    For KKL& = L& To URE&
'      TBUF$ = REGLEIDO$("PDOCSPL", KKL&)
'      If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 10
'      If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
'        TPSP$ = TPSP$ + Mid$(TBUF$, 5)
'      End If
'    Next KKL&
'    '
'10  For UI& = ULTREG&("PDOCLST") To 1 Step -1
'      XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
'      If InStr(XX$, TIDOCU2$) > 4 Then
'        DOCUNU& = CVS(Left$(XX$, 4))
'        GoTo 15
'      End If
'    Next UI&
'    '
'15  REDOCU$ = REGACT$("PDOCLST")
'    URE& = ULTREG&("PDOCSPL")
'    LI& = 0: LS& = URE&
'16  E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 17
'    L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 17
'    DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
'    If DCC& < DOCUNU& Then LI& = L&: GoTo 16
'    If DCC& > DOCUNU& Then LS& = L&: GoTo 16
'    GoTo 18
'    '
'17  Screen.MousePointer = 1
'    Call MENSERR(24, "Documento no Registrado")
'    Exit Sub
'    '
'18  While L& > 1
'      DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
'      If DCC& < DOCUNU& Then GoTo 19
'      L& = L& - 1
'    Wend
'    '
'19  For KKL& = L& To URE&
'      TBUF$ = REGLEIDO$("PDOCSPL", KKL&)
'      If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 95
'      If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
'        TPSP$ = TPSP$ + Mid$(TBUF$, 5)
'      End If
'    Next KKL&
'    '
'95  SPOOLSTRING$ = TPSP$
'    For KKI% = 0 To 5
'      EDITFORM.Picture1(KKI%).Cls
'      EDITFORM.Picture1(KKI%).Top = 0
'      EDITFORM.Picture1(KKI%).Refresh
'      EDITFORM.Picture1(KKI%).Height = 7170
'      EDITFORM.Picture1(KKI%).Width = 11700
'      EDITFORM.Picture1(KKI%).DrawWidth = 3
'      EDITFORM.Picture1(KKI%).DrawStyle = 0
'      EDITFORM.Picture1(KKI%).FillStyle = 1
'    Next KKI%
'    ALTUPAGINA = 7170
'    TOPEPAGINA = 0
'    PAGINACTIVA% = 0
'    CAPAGINAS% = 0
'    '
'    FIN& = Len(TPSP$)
'50  PPXX% = InStr(TPSP$, Chr$(255))
'    If PPXX% > 0 Then
'      COPRIAT$ = Left$(TPSP$, PPXX% - 1)
'      TPSP$ = Mid$(TPSP$, PPXX% + 1)
'      Call TRACE(20, FIN& - Len(TPSP$), FIN&)
'      '
'      If Len(COPRIAT$) > 0 Then
'        COPRI% = Asc(Left$(COPRIAT$, 1))
'        ATRIP$ = Mid$(COPRIAT$, 2)
'        Call SHOWCOMMAND(COPRI%, ATRIP$)
'      End If
'      GoTo 50
'      '
'    End If
'    '
'    Screen.MousePointer = 1
'    PAGINACTIVA% = 0
'    For KKI% = 0 To 5
'       EDITFORM.Picture1(KKI%).Visible = False
'    Next KKI%
'    EDITFORM.Picture1(PAGINACTIVA%).Visible = True
'    EDITFORM.Picture1(PAGINACTIVA%).Refresh
'    EDITFORM.VScroll1.Min = 0
'    EDITFORM.VScroll1.Value = 0
'    EDITFORM.VScroll1.Max = 1
'    If ALTUPAGINA > EDITFORM.Frame1.Height Then
'       EDITFORM.VScroll1.Max = (ALTUPAGINA - EDITFORM.Frame1.Height) / 100
'    End If
'    EDITFORM.Command3.Enabled = False
''   EDITFORM.Command5.Enabled = False
'    EDITFORM.Show 1
    '
    GoTo 3
    '
99  Command4.Enabled = True
    '
End Sub

 Sub Command5_Click()
   Command5.Enabled = False
'3  Screen.MousePointer = 11
'   Call COPYSTRU("ORDEREPA", "PREPAEMI")
'   JJ& = 0
'   For U& = 1 To ULTREG&("ORDEREPA")
'     XX$ = REGLEIDO$("ORDEREPA", U&)
'     If CVI(Mid$(XX$, 65, 2)) < 3 Then
'       JJ& = JJ& + 1
'       Call GRAREG("PREPAEMI", XX$, JJ&)
'     End If
'   Next U&
'   Call SETULTREG("PREPAEMI", JJ&)
'   Call CIERRARCH("PREPAEMI")
'   Screen.MousePointer = 1
   '
   CONDI$ = "NuOrdRep > 0 And Status < 3"
   JJ& = CantRegistros("REPARA", CONDI$)
   If JJ& < 1 Then
     Call COMUNI("No Hay Ordenes de Reparación en Curso\que Puedan Anularse.")
     GoTo 99
   End If
   '
   CONDI$ = "NuOrdRep > 0 And Status < 3 ORDER BY NuOrdRep"
   Call CARGALISEL("PREPAEM", "REPARA", "NuOrdRep/I10;RefOrdRep/A64", CONDI$)
   '
   Call SELECHO("PREPAEM")
   NOREP& = Val(VALACT1$("PREPAEM"))
   If NOREP& < 1 Then GoTo 99
   '
   If COMALTER%("Confirma Anulación de O/R " + TRIM$(Str$(NOREP&)) + "\" + TRIM$(VALACT2$("PREPAEMI")) + "\\Cancelar\Anular O/R") <> 2 Then
      GoTo 99
   End If
   '
   Screen.MousePointer = 11
   
   NOREX$ = TRIM$(Str$(NOREP&))
   While Len(NOREX$) < 6: NOREX$ = "0" + NOREX$: Wend
   NOREX$ = TRIM$("OR-" + NOREX$)
   
   CONDI$ = "NuOrdRep = " & NOREP&
   Call ACTUREGISTRO("REPARA", "Status", CONDI$, 9, REG&)
   'BORRO LA LISTA DE REPUESTOS DE DICHA REPARACION
   CONDI$ = "NumRepa = " & NOREP&
   Call BORRAREGISTROS("REPUREPA", CONDI$, REG&)
   '
   CONDI$ = "IdSubOr ='" & NOREX$ & "'"
   Call BORRAREGISTROS("RESERVA", CONDI$, REG&)
'   NUPRESU& = 0
'   For U& = 1 To ULTREG&("ORDEREPA")
'     X$ = REGLEIDO$("ORDEREPA", U&)
'     If CVS(Left$(X$, 4)) = NOREP& Then
'       REFOP$ = "(anul) " + Mid$(X$, 5, 60)
'       NUPRESU& = CVS(Mid$(X$, 125, 4))
'       Call REPLA(X$, REFOP$, 5, 60)
'       Call REPLA(X$, MKI$(9), 65, 2) ' MARCA ANULADA
'       Call GRAREG("ORDEREPA", X$, U&)
'       GoTo 20
'     End If
'   Next U&
'   '
'20 If NUPRESU& > 0 Then
'     For U& = 1 To ULTREG&("PRESREPA")
'       X$ = REGLEIDO$("PRESREPA", U&)
'       If CVS(Left$(X$, 4)) = NUPRESU& Then
'         Call REPLA(X$, MKS$(0), 125, 4) ' QUITA NUMERO DE O/R
'         Call GRAREG("PRESREPA", X$, U&)
'         GoTo 30
'       End If
'     Next U&
'   End If
'   '
'30 NOREX$ = TRIM$(Str$(NOREP&))
'   While Len(NOREX$) < 6: NOREX$ = "0" + NOREX$: Wend
'   NOREX$ = "OR-" + NOREX$
'   '
'   RFF$ = RECFILT$("RESERVA", 3, NOREX$)
'   RCC$ = ""
'   URE& = ULTREG&("RESERVA")
'   For KK& = 1 To Len(RFF$) Step 4
'     RGG& = CVS(Mid$(RFF$, KK&, 4))
'     If RGG& > 0 Then
'       If RGG& <= URE& Then
'         XX$ = REGLEIDO$("RESERVA", RGG&)
'         If TRIM$(Mid$(XX$, 3, 12)) = NOREX$ Then
'           Call GRAREG("RESERVA", String$(128, 0), RGG&)
'         End If
'       End If
'     End If
'   Next KK&
'   '
'   Call CIERRARCH("PRESREPA")
'   Call CIERRARCH("ORDEREPA")
'   Call CIERRARCH("RESERVA")
   Screen.MousePointer = 1
   '
   Call COMUNI("Anulación Completa")
   '
99 Command5.Enabled = True
End Sub

Private Sub Command6_Click()
   '
   Command6.Enabled = False
   '
   Call COPYSTRU("PRESREPA", "PREPAEMI")
'   JJ& = 0
'   For U& = 1 To ULTREG&("PRESREPA")
'     XX$ = REGLEIDO$("PRESREPA", U&)
'     If CVS(Mid$(XX$, 125, 4)) < 0.5 Then
'       JJ& = JJ& + 1
'       Call GRAREG("PREPAEMI", XX$, JJ&)
'     End If
'   Next U&
'   Call SETULTREG("PREPAEMI", JJ&)
'   Call CIERRARCH("PREPAEMI")
   '
   CONDI$ = "NROPED < 0.5 AND STATUS < 9"
   JJ& = CantRegistros("REPARA", CONDI$)
   If JJ& < 1 Then
     Call COMUNI("No Hay Pendientes\para Emitir Presupuestos de Reparación.")
     GoTo 99
   End If
   '
   CONDI$ = CONDI$ + " ORDER BY NUMREPA"
   Call CARGALISEL("PREPAEMI", "REPARA", "NumRepa/I10;Referen/A128", CONDI$)
   Call FRESHECHO("PREPAEMI")
   '
2  VDEF$ = MKS$(0) + Space$(16)
3  OBX$ = OBJPLA$("ENCAREPA", VDEF$)
   If OBX$ = "" Then GoTo 99
   '
   VDEF$ = OBX$
   NPX& = CVS(Left$(OBX$, 4))
   If NPX& < 1 Or ECOARCH$("PREPAEMI", MKS$(NPX&)) = "" Then GoTo 3
   NUPRESUP& = NPX&
   '
   COD$ = TRIM$(Mid$(OBX$, 5, 16))
   If COD$ = "" Or ECOARCH$("INDIREP", COD$) = "" Then GoTo 3
   '
'   For UI& = ULTREG&("PRESREPA") To 1 Step -1
'     XX$ = REGLEIDO$("PRESREPA", UI&)
'     If CVS(Left$(XX$, 4)) = NPX& Then GoTo 10
'   Next UI&
'   Call COMUNI("Error de Indices - Consulte.")
'   GoTo 99
'   '
'   If CVS(Mid$(XX$, 125, 4)) > 0.5 Then
'      Call COMUNI("Imposible Generar / Modificar Presupuesto.\Orden de Reparación ya Emitida.")
'      GoTo 3
'   End If
'10 NC% = CVI(Mid$(XX$, 67, 2))
   '
   CONDI$ = "NumRepa=" & Str(NPX&)
   NPP& = XVALO(VALOBADA("REPARA", "NROPED", CONDI$))
   If NPP& > 0 Then
      Call COMUNI("Imposible Generar / Modificar Presupuesto.\Orden de Reparación ya Emitida.")
      GoTo 3
   End If
   
10 NC% = VALOBADA("REPARA", "Nume_Cli", CONDI$, "NOMESS")
   If NC% > 0 Then
     If rasocli$(NC%) <> "" Then
       PRESRE09.Label18 = TRIM$(Str$(NPX&))
       PRESRE09.Label11 = TRIM$(Str$(NC%))
       PRESRE09.Label10 = COD$
       '
       Screen.MousePointer = 11
       RNC$ = RECFILT$("DESCREP", 1, COD$)
       URN& = ULTREG&("DESCREP")
       '
       For U& = 1 To Len(RNC$) Step 4
         RENOTA& = CVS(Mid$(RNC$, U&, 4))
         If RENOTA& > 0 Then
           If RENOTA& <= URN& Then
             TTYY$ = RTrim$(Mid$(REGLEIDO$("DESCREP", RENOTA&), 17))
             If TTYZ$ <> "" Then
               TTYZ$ = TTYZ$ + Chr$(13) + Chr$(10)
             End If
             TTYZ$ = TTYZ$ + RTrim$(TTYY$)
           End If
         End If
       Next U&
       PRESRE09.Label8 = TTYZ$
       '
       coefi = Val(Control$("REPARA", "COEFPREA"))
       If coefi < 1 Then coefi = 1
       CUHOR = Val(Control$("REPARA", "VALHORA"))
       PRESRE09.Label20 = CSTRING$(MKS$(CUHOR), 4, 8, 2, 2)
       '
       If CantRegistros("REPUREPA", "NumRepa = " & NPX&) > 0 Then
            OPX% = COMALTER%("La Reparacion Elegida Posee Repuestos.\¿Desea Actualizarlos En Base Al Codigo de Bomba Elegido?\\No, Conservar\Si, Refrescar")
            If OPX% <= 1 Or OPX% > 2 Then GoTo 98
       End If
       '
       PRESRE09.List1.Clear
       SUREPUES = 0
       Call SETULTREG("!MATEROF", 0)
       '
       CODD$ = COD$
       Screen.MousePointer = 11
       RFF$ = RECFILT$("REPUESRE", 1, CODD$)
       For IK& = 1 To Len(RFF$) Step 4
          RERE& = CVS(Mid$(RFF$, IK&, 4))
          If RERE& > 0 Then
            If RERE& <= ULTREG&("REPUESRE") Then
              XX$ = REGLEIDO$("REPUESRE", RERE&)
              CIY% = CVI(Mid$(XX$, 17, 2))
              CODXX$ = CODEXT$(CIY%)
              DEXX$ = DESCRIT$(CIY%)
              COSUN = coefi * COSUNI(CIY%)
              COSUX$ = CSTRING$(MKS$(COSUN), 4, 10, 2, 2)
              Uso = CVS(Mid$(XX$, 19, 4))
              USOX$ = CSTRING$(MKS$(Uso), 4, 10, 1, 2)
              COSTO = COSUN * Uso
              COSTX$ = CSTRING$(MKS$(COSTO), 4, 12, 2, 2)
              SUREPUES = SUREPUES + COSTO
              '
              XXX$ = Space$(64)
              Call REPLA(XXX$, CODXX$, 1, 16)
              Call REPLA(XXX$, DEXX$, 17, 38)
              Call REPLA(XXX$, USOX$, 54, 10)
              Call REPLA(XXX$, COSUX$, 64, 10)
              Call REPLA(XXX$, COSTX$, 74, 12)
              PRESRE09.List1.AddItem XXX$
            End If
          End If
       Next IK&
       PRESRE09.SUMAREP = CSTRING$(MKS$(SUREPUES), 4, 12, 2, 2)
       '
       SUTIEREP = 0
       RFFX$ = RECFILT$("SECOPREP", 1, CODD$)
       For I1& = 1 To Len(RFFX$) Step 4
         i& = CVS(Mid$(RFFX$, I1&, 4))
         X$ = REGLEIDO$("SECOPREP", i&)
         '
         CAOPS = CVS(Mid$(X$, 97, 4)): If CAOPS < 1 Then CAOPS = 1
         SUTIEREP = SUTIEREP + CAOPS * CVS(Mid$(X$, 73, 4))
       Next I1&
       PRESRE09.Label19 = CSTRING$(MKS$(SUTIEREP), 4, 8, 1, 2)
       '
98     Screen.MousePointer = 1
       '
       PRESRE09.Label18 = TRIM$(Str$(NUPRESUP&))
       PRESRE09.Show 1
       GoTo 2
     End If
   End If
   '
99 Command6.Enabled = True

End Sub

Private Sub Command7_Click()
   Command7.Enabled = False
   ANOREP09.Show 1
   Command7.Enabled = True
End Sub

Private Sub Command8_Click()
   Command8.Enabled = False
   RECEBO09.Show 1
   Command8.Enabled = True
End Sub

 Sub Command9_Click()
   Command9.Enabled = False
   '
3  Screen.MousePointer = 11
   Call COPYSTRU("PRESREPA", "PREPAEMI")
   JJ& = 0
'   For u& = 1 To ULTREG&("PRESREPA")
'     XX$ = REGLEIDO$("PRESREPA", u&)
'     If CVI(Mid$(XX$, 121, 2)) > 0 Then
'       If CVS(Mid$(XX$, 125, 4)) < 0.5 Then
'         JJ& = JJ& + 1
'         Call GRAREG("PREPAEMI", XX$, JJ&)
'       End If
'     End If
'   Next u&
'   Call SETULTREG("PREPAEMI", JJ&)
'   Call CIERRARCH("PREPAEMI")
   '
   CONDI$ = "NROPED > 0 AND STATUS < 9 AND NuOrdRep < 0.5"
   If InStr(1, EMPREAC$, "NEUQUEN") > 1 Then CONDI$ = "NuOrdRep < 0.5 AND STATUS < 9"
   JJ& = CantRegistros("REPARA", CONDI$)
   If JJ& < 1 Then
     Call COMUNI("No Hay Ordenes de Venta Pendientes\para Emitir Orden de Reparación.")
     GoTo 99
   End If
   
   'CONDI$ = "NROPED > 0 AND NuOrdRep < 0.5 ORDER BY NUMREPA"
   CONDI$ = CONDI$ + " ORDER BY NUMREPA"
   Call CARGALISEL("PREPAEMI", "REPARA", "NumRepa/I10;Referen/A128", CONDI$)
   
   Call SELECHO("PREPAEMI/Indice Ordenes de Venta de Reparación Pendientes")
   NPX& = Val(VALACT1$("PREPAEMI"))
   If NPX& < 1 Then GoTo 99
   '
'   For UI& = ULTREG&("PRESREPA") To 1 Step -1
'     XX$ = REGLEIDO$("PRESREPA", UI&)
'     If CVS(Left$(XX$, 4)) = NPX& Then GoTo 10
'   Next UI&
'   Call COMUNI("Error de Indices - Consulte.")
'
'   Exit Sub
   '
'10 REPRESUP = TRIM$(Str$(UI&))
'   XX$ = REGLEIDO$("PRESREPA", Val(REPRESUP))
'   If CVI(Mid$(XX$, 121, 2)) < 1 Then
'      Call COMUNI("Falta Generar y Emitir Presupuesto.")
'      GoTo 99
'   End If
   '
'   If CVS(Mid$(XX$, 125, 4)) > 0.5 Then
'      Call COMUNI("Orden de Reparación ya Emitida.")
'      GoTo 3
'   End If
   '
   CONDI$ = "NumRepa = " & NPX&
'   FECHAPRESU = CVINT(VALOBADA("REPARA", "FechPres", CONDI$))
'   If FECHAPRESU < 1 Then
'      Call COMUNI("Falta Generar y Emitir Presupuesto.")
'      GoTo 99
'   End If
   '
   NPP& = XVALO(VALOBADA("REPARA", "NroPed", CONDI$))
'   If NPP& > 0 Then
'      Call COMUNI("Orden de Reparación ya Emitida.")
'      GoTo 3
'   End If
   '
   ORREPA09.Label18 = TRIM$(Str$(NPX&))
   ORREPA09.Label2 = TRIM$(Str$(NUORDREP&))
   '
'   NPP& = 1
'   FIN& = ULTREG&("PEDENCA")
'   If FIN& > 0 Then
'      For U& = FIN& To FIN& - 16 Step -1
'         If U& < 1 Then GoTo 99
'         XX$ = REGLEIDO$("PEDENCA", U&)
'         NPPX& = CVS(Mid$(XX$, 1, 4))
'         If NPPX& >= NPP& Then
'            NPP& = 1 + NPPX& ' Nro de Pedido
'         End If
'      Next U&
'   End If
'   CONDI$ = "NroPed > 0"
'   NPP& = XVALO(VALOBADA("PEDENCA", "Max(NroPed)", CONDI$)) + 1
   ORREPA09.Label25 = TRIM$(Str$(NPP&))
   '
   ORREPA09.Show 1
   GoTo 3
   '
99 Command9.Enabled = True
   Screen.MousePointer = 1
   '
End Sub

Function NUORDREP&()
   '
   NPP& = XVALO(VALOBADA("REPARA", "Max(NuOrdRep)", "NumRepa > 0")) + 1
'   NPP& = 1
'   '
'   FIN& = ULTREG&("ORDEREPA")
'   If FIN& > 0 Then
'      For U& = FIN& To FIN& - 128 Step -1
'         If U& < 1 Then GoTo 99
'         XX$ = REGLEIDO$("ORDEREPA", U&)
'         NPPX& = CVS(Mid$(XX$, 1, 4))
'         If NPPX& >= NPP& Then
'            NPP& = 1 + NPPX&
'          End If
'      Next U&
'   End If
'   '
99 NUORDREP& = NPP&
   '
End Function
'

Private Sub Form_Load()
    '
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    'If HOY(HOS$) > FECHANUM("15/02/03") Then
    '  Call CONECRUN("REPABO02", "")
    '  Call FINAL("")
    'End If
    '
    CODGENER$ = TRIM$(Control$("REPARA", "CODGENER"))
    If CODINT%(CODGENER$) < 1 Then
       Call MENSERR(24, "Falta Código Genérico Reparaciones.")
       CONFRE09.Show 1
    End If
    '
    Call APERBASDAT("ORFAB")
    '
    Call PoneEnCeroNull("REPUREPA", "ORIGEN", REG&)
    '
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub

Sub CARTRANSPOR()
    '
    NCLIE% = Val(FOREMI09.Text1.TEXT)
    LENTREG$ = Space$(96)
    RFEX$ = RECFILT$("LENTRECL", 1, MKI$(NC%))
    If Len(RFEX$) >= 4 Then
       RELENT& = CVS(Left$(RFEX$, 4))
       If RELENT& > 0 Then
         If RELENT& <= ULTREG&("LENTRECL") Then
           LENTREG$ = Mid$(REGLEIDO$("LENTRECL", RELENT&), 3)
         End If
       End If
    End If
    FOREMI09.Text8.TEXT = TRIM$(Mid$(LENTREG$, 65, 32))
    '
End Sub

Sub IMPRIMEOFLIMPIEZA(NROREP&)
    
    FORIPRE1$ = Control$("", "FORMORFA")
    If FORIPRE1$ = "" Then
       Call MENSERR(24, "Imposible Imprimir O/F Limpieza.\Falta Definir Formulario Electrónico.")
       Exit Sub
    End If
    If EXISTE%(LUDAT$ + FORIPRE1$ + ".FRM") < 1 Then
       Call MENSERR(24, "Imposible Imprimir O/F Limpieza.\No existe Formulario Electrónico '" + FORIPRE$ + ".FRM'.")
       Exit Sub
    End If
    '
    FORIPRE2$ = Control$("", "FORMORFB")
    If FORIPRE2$ = "" Then
       Call MENSERR(24, "Imposible Imprimir Vale de O/F Limpieza.\Falta Definir Formulario Electrónico.")
       Exit Sub
    End If
    If EXISTE%(LUDAT$ + FORIPRE2$ + ".FRM") < 1 Then
       Call MENSERR(24, "Imposible Imprimir Vale de O/F Limpieza.\No existe Formulario Electrónico '" + FORIPRE$ + ".FRM'.")
       Exit Sub
    End If
    '
10  NC% = Val(Label11)
    CONDI$ = "NUMREPA=" & NROREP&
    NC% = XVALO(VALOBADA("REPARA", "NUME_CLI", CONDI$, "NOMES"))
    DL$ = NROREP&
    While Len(DL$) < 6: DL$ = "0" + DL$: Wend
    DL$ = "DL-" + DL$ + "-1"
    CONDI$ = "IDSUBOR='" & DL$ & "'"
    TIEMPO$ = HORATEX$(XVALO(VALOBADA("OPERFAB", "TIEMVARI", CONDI$)))
    Ci% = XVALO(VALOBADA("ORDEFABR", "COD_INTE", CONDI$, "NOMES"))
    
    TIPODOC$ = "O/F de Limpieza"
    DESTIDOC% = NC%
    NUMEDOC$ = TRIM$(DL$)
    FF$ = FECHATEX$(HOY(HO$))
    FECHDOC$ = TRIM$(FF$)
    '
    CODPARAM$(1) = "COD-CONJ": DOCPARAM$(1) = CODEXT$(Ci%) ' CODIGO
    CODPARAM$(2) = "NUME-COM": DOCPARAM$(2) = DL$ ' NUMERO
    CODPARAM$(3) = "FECH-EMI": DOCPARAM$(3) = FECHDOC$ ' FECHA
    CODPARAM$(4) = "CAN-CONJ": DOCPARAM$(4) = "1" ' CANTIDAD"
    CODPARAM$(5) = "DES-CONJ": DOCPARAM$(5) = DESCRIT0$(Ci%) ' DESCRIPCION
    CODPARAM$(6) = "DES-MPRI": DOCPARAM$(6) = "Según Vale de Materiales"
    CAPARDOC% = 6
    '
    CODTABU1$(1) = "NOP-OPER"
    CODTABU1$(2) = "OBS-OPER"
    CACOLTAB1% = 2
    '
    CAITAB1% = 1
    TETABU1$(1, CAITAB1%) = "10"
    TETABU1$(2, CAITAB1%) = "Reserva de Horas " & TIEMPO$
    '
    Call PRINTDOC("@" + FORIPRE1$)
    '
    If ULTREG&("!BOLREDET") > 0 Then
        TIPODOC$ = "Vale de Materiales"
        CODPARAM$(1) = "COD-CONJ": DOCPARAM$(1) = CODEXT$(Ci%) ' CODIGO
        CODPARAM$(2) = "NUME-COM": DOCPARAM$(2) = DL$ ' NUMERO
        CODPARAM$(3) = "FECH-EMI": DOCPARAM$(3) = FECHDOC$ ' FECHA
        CODPARAM$(4) = "CAN-CONJ": DOCPARAM$(4) = "1" ' CANTIDAD"
        CODPARAM$(5) = "DES-CONJ": DOCPARAM$(5) = DESCRIT0$(Ci%) ' DESCRIPCION
        CAPARDOC% = 5
        
        CODTABU1$(1) = "COD-MPRI"
        CODTABU1$(2) = "DES-MPRI"
        CODTABU1$(3) = "UNIMED"
        CODTABU1$(4) = "USO-MPRI"
        CODTABU1$(5) = "CAN-MPRI"
        CACOLTAB1% = 5
        '
        CAITAB1% = 0
        For II& = 1 To ULTREG&("!BOLREDET")
            XX$ = REGLEIDO$("!BOLREDET", II&)
            Ci% = CVI(Mid$(XX$, 83, 2))
            DESCRI$ = DESCRIT0$(Ci%)
            UNI$ = Unimed(Ci%)
            CAN# = CVD(Mid$(XX$, 89, 8))
        
            CAITAB1% = CAITAB1% + 1
            TETABU1$(1, CAITAB1%) = CODEXT$(Ci%)
            TETABU1$(2, CAITAB1%) = DESCRI$
            TETABU1$(3, CAITAB1%) = UNI$
            TETABU1$(5, CAITAB1%) = FORMATNUM$(CAN#, "F12.1")
        Next II&
        '
        Call PRINTDOC("@" + FORIPRE2$)
    End If
End Sub





