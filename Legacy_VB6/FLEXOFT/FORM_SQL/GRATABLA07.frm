VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form GRATABLA 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Costo por Tipo de Ventas"
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11955
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   11955
   Begin VB.CommandButton Command3 
      Caption         =   "Exportar a Excel"
      Height          =   255
      Left            =   6120
      TabIndex        =   27
      Top             =   3960
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Imprimir"
      Height          =   260
      Left            =   4080
      Picture         =   "GRATABLA07.frx":0000
      TabIndex        =   26
      Top             =   3960
      Width           =   1800
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3900
      Index           =   1
      Left            =   105
      ScaleHeight     =   3900
      ScaleWidth      =   11715
      TabIndex        =   13
      Top             =   4305
      Width           =   11715
      Begin VB.Frame Frame7 
         Caption         =   "Frame7"
         Height          =   3060
         Index           =   1
         Left            =   11285
         TabIndex        =   15
         Top             =   630
         Width           =   15
      End
      Begin VB.Frame Frame6 
         Caption         =   "Frame6"
         Height          =   3060
         Index           =   1
         Left            =   10135
         TabIndex        =   16
         Top             =   630
         Width           =   15
      End
      Begin VB.Frame Frame5 
         Caption         =   "Frame5"
         Height          =   3165
         Index           =   1
         Left            =   9135
         TabIndex        =   17
         Top             =   630
         Width           =   15
      End
      Begin VB.Frame Frame4 
         Caption         =   "Frame4"
         Height          =   3165
         Index           =   1
         Left            =   7560
         TabIndex        =   18
         Top             =   630
         Width           =   15
      End
      Begin VB.Frame Frame1 
         Caption         =   "Frame1"
         Height          =   3165
         Index           =   1
         Left            =   5985
         TabIndex        =   21
         Top             =   630
         Width           =   15
      End
      Begin VB.Frame Frame3 
         Caption         =   "Frame3"
         Height          =   3165
         Index           =   1
         Left            =   4305
         TabIndex        =   19
         Top             =   630
         Width           =   15
      End
      Begin VB.Frame Frame8 
         Caption         =   "Frame8"
         Height          =   3060
         Index           =   1
         Left            =   3360
         TabIndex        =   14
         Top             =   630
         Width           =   15
      End
      Begin VB.Frame Frame2 
         Caption         =   "Frame2"
         Height          =   3165
         Index           =   1
         Left            =   2415
         TabIndex        =   20
         Top             =   630
         Width           =   15
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2310
         Index           =   1
         Left            =   120
         TabIndex        =   22
         Top             =   945
         Width           =   11530
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   105
         TabIndex        =   25
         Top             =   105
         Width           =   11565
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   2
         Height          =   330
         Index           =   1
         Left            =   105
         Top             =   75
         Width           =   11565
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   105
         X2              =   11655
         Y1              =   3675
         Y2              =   3675
      End
      Begin VB.Shape Shape2 
         BorderWidth     =   2
         Height          =   3140
         Index           =   1
         Left            =   105
         Top             =   630
         Width           =   11565
      End
      Begin VB.Line Line2 
         Index           =   1
         X1              =   105
         X2              =   11655
         Y1              =   945
         Y2              =   945
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "GRUPO CLIENTE         Cant      %         VENTA        COSTO     MARGEN B.   % C/V  % MB/VTA"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   170
         TabIndex        =   24
         Top             =   670
         Width           =   11565
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "GRUPO CLIENTE         Cant      %         VENTA        COSTO     MARGEN B.   % C/V  % MB/VTA"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   170
         TabIndex        =   23
         Top             =   3360
         Width           =   11565
      End
      Begin VB.Line Line3 
         Index           =   1
         X1              =   105
         X2              =   11655
         Y1              =   3255
         Y2              =   3255
      End
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3900
      Index           =   0
      Left            =   105
      ScaleHeight     =   3900
      ScaleWidth      =   11715
      TabIndex        =   0
      Top             =   105
      Width           =   11715
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   0
         Top             =   3480
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Frame Frame8 
         Caption         =   "Frame8"
         Height          =   3060
         Index           =   0
         Left            =   3360
         TabIndex        =   12
         Top             =   630
         Width           =   15
      End
      Begin VB.Frame Frame7 
         Caption         =   "Frame7"
         Height          =   3060
         Index           =   0
         Left            =   11285
         TabIndex        =   10
         Top             =   630
         Width           =   15
      End
      Begin VB.Frame Frame6 
         Caption         =   "Frame6"
         Height          =   3060
         Index           =   0
         Left            =   10135
         TabIndex        =   9
         Top             =   630
         Width           =   15
      End
      Begin VB.Frame Frame5 
         Caption         =   "Frame5"
         Height          =   3165
         Index           =   0
         Left            =   9135
         TabIndex        =   8
         Top             =   630
         Width           =   15
      End
      Begin VB.Frame Frame4 
         Caption         =   "Frame4"
         Height          =   3165
         Index           =   0
         Left            =   7560
         TabIndex        =   7
         Top             =   630
         Width           =   15
      End
      Begin VB.Frame Frame3 
         Caption         =   "Frame3"
         Height          =   3165
         Index           =   0
         Left            =   4305
         TabIndex        =   5
         Top             =   630
         Width           =   15
      End
      Begin VB.Frame Frame2 
         Caption         =   "Frame2"
         Height          =   3165
         Index           =   0
         Left            =   2415
         TabIndex        =   4
         Top             =   630
         Width           =   15
      End
      Begin VB.Frame Frame1 
         Caption         =   "Frame1"
         Height          =   3165
         Index           =   0
         Left            =   5985
         TabIndex        =   3
         Top             =   630
         Width           =   15
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2310
         Index           =   0
         Left            =   90
         TabIndex        =   2
         Top             =   945
         Width           =   11530
      End
      Begin VB.Line Line3 
         Index           =   0
         X1              =   105
         X2              =   11655
         Y1              =   3255
         Y2              =   3255
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "GRUPO CLIENTE         Cant      %         VENTA        COSTO     MARGEN B.   % C/V  % MB/VTA"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   170
         TabIndex        =   11
         Top             =   3360
         Width           =   11565
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "GRUPO CLIENTE         Cant      %         VENTA        COSTO     MARGEN B.   % C/V  % MB/VTA"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   170
         TabIndex        =   6
         Top             =   670
         Width           =   11565
      End
      Begin VB.Line Line2 
         Index           =   0
         X1              =   105
         X2              =   11655
         Y1              =   945
         Y2              =   945
      End
      Begin VB.Shape Shape2 
         BorderWidth     =   2
         Height          =   3140
         Index           =   0
         Left            =   105
         Top             =   630
         Width           =   11565
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   105
         X2              =   11655
         Y1              =   3675
         Y2              =   3675
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   2
         Height          =   330
         Index           =   0
         Left            =   105
         Top             =   75
         Width           =   11565
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   105
         TabIndex        =   1
         Top             =   105
         Width           =   11565
      End
   End
End
Attribute VB_Name = "GRATABLA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   Hide
End Sub

Private Sub Command2_Click()
   '
   ATENTI% = ATENTI% + 1
   Hide
   '
End Sub

Private Sub Command3_Click()
   CommonDialog1.CancelError = True
   CommonDialog1.DialogTitle = "Exportar a Excel"
       ' cancelado, si no da error - CommonDialog1.Flags = &H2& + &H4& + &H8 + &H800 + &H2000 + &H8000 'cdlOFNCreatePrompt
   CommonDialog1.Filter = "Excel Files(*.XLS)|*.XLS|"
   CommonDialog1.FilterIndex = 1
   On Error GoTo 99
   CommonDialog1.ShowOpen
   On Error GoTo 0
   '
   DoEvents
   Screen.MousePointer = 11
   AREXPO$ = UCase$(CommonDialog1.FileName)
   If TRIM$(AREXPO$) <> "" Then
      Call GENEXCEL(AREXPO$)
   End If
99:
End Sub

Private Sub Form_Load()
    Call CENTRAFORM(Me)
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
End Sub

Sub GENEXCEL(NOMBREDEARCHIVO$)
   Dim XLAPP As Excel.Application
   Dim XLBOOK As Excel.Workbook
   Dim XLSHEET As Excel.Worksheet
   Dim XLSHEET1 As Excel.Worksheet
   '
   Dim FECHA As Date
   Dim VALOR As Double
   Dim TEXTO As String
   '
   On Error GoTo 20
   Set XLAPP = New Excel.Application
   Set XLBOOK = XLAPP.Workbooks.Add
   Set XLSHEET = XLBOOK.Worksheets.Add
   On Error GoTo 0
   '
   TIPODATO% = 0: COLUM& = 0
   '
   ' PRIMER REPORTE
   '
   Set XLSHEET1 = XLBOOK.Worksheets.Add
   XLSHEET1.Cells(1, 1) = Label1(0) ' TITULO DEL REPORTE
   XLSHEET1.Cells(3, 1) = "Grupo de Cliente"
   XLSHEET1.Cells(3, 2) = "Cantidad"
   XLSHEET1.Cells(3, 3) = "%"
   XLSHEET1.Cells(3, 4) = "Venta"
   XLSHEET1.Cells(3, 5) = "Costo"
   XLSHEET1.Cells(3, 6) = "Margen Bruto"
   XLSHEET1.Cells(3, 7) = "% C/V"
   XLSHEET1.Cells(3, 8) = "% MB/VTA"
   For I& = 4 To (List1(0).ListCount - 1) + 4
        TXT$ = List1(0).List(I& - 4)
        VALORX$ = Mid$(TXT$, 1, 20)
        XLSHEET1.Cells(I&, 1) = VALORX$
        VALORX$ = Mid$(TXT$, 21, 6)
        XLSHEET1.Cells(I&, 2) = XVALO(VALORX$)
        VALORX$ = Mid$(TXT$, 27, 8)
        XLSHEET1.Cells(I&, 3) = XVALO(VALORX$)
        VALORX$ = Mid$(TXT$, 36, 13)
        XLSHEET1.Cells(I&, 4) = XVALO(VALORX$)
        VALORX$ = Mid$(TXT$, 49, 13)
        XLSHEET1.Cells(I&, 5) = XVALO(VALORX$)
        VALORX$ = Mid$(TXT$, 62, 13)
        XLSHEET1.Cells(I&, 6) = XVALO(VALORX$)
        VALORX$ = Mid$(TXT$, 75, 8)
        XLSHEET1.Cells(I&, 7) = XVALO(VALORX$)
        VALORX$ = Mid$(TXT$, 85, 8)
        XLSHEET1.Cells(I&, 8) = XVALO(VALORX$)
   Next I&
   '
   NN% = (List1(0).ListCount - 1) + 4 + 2
   TXT$ = Label3(0)
   VALORX$ = "Total": XLSHEET1.Cells(NN%, 1) = VALORX$
   VALORX$ = Mid$(TXT$, 21, 6): XLSHEET1.Cells(NN%, 2) = XVALO(VALORX$)
   VALORX$ = Mid$(TXT$, 27, 6): XLSHEET1.Cells(NN%, 3) = XVALO(VALORX$)
   VALORX$ = Mid$(TXT$, 36, 13): XLSHEET1.Cells(NN%, 4) = XVALO(VALORX$)
   VALORX$ = Mid$(TXT$, 49, 13): XLSHEET1.Cells(NN%, 5) = XVALO(VALORX$)
   VALORX$ = Mid$(TXT$, 62, 13): XLSHEET1.Cells(NN%, 6) = XVALO(VALORX$)
   VALORX$ = Mid$(TXT$, 75, 8): XLSHEET1.Cells(NN%, 7) = XVALO(VALORX$)
   VALORX$ = Mid$(TXT$, 85, 8): XLSHEET1.Cells(NN%, 8) = XVALO(VALORX$)
   '
   ' SEGUNDO REPORTE
   '
   XLSHEET.Cells(1, 1) = Label1(1) ' TITULO DEL REPORTE
   XLSHEET.Cells(3, 1) = "Grupo de Cliente"
   XLSHEET.Cells(3, 2) = "Cantidad"
   XLSHEET.Cells(3, 3) = "%"
   XLSHEET.Cells(3, 4) = "Venta"
   XLSHEET.Cells(3, 5) = "Costo"
   XLSHEET.Cells(3, 6) = "Margen Bruto"
   XLSHEET.Cells(3, 7) = "% C/V"
   XLSHEET.Cells(3, 8) = "% MB/VTA"
   For I& = 4 To (List1(1).ListCount - 1) + 4
        TXT$ = List1(1).List(I& - 4)
        VALORX$ = Mid$(TXT$, 1, 20)
        XLSHEET.Cells(I&, 1) = VALORX$
        VALORX$ = Mid$(TXT$, 21, 6)
        XLSHEET.Cells(I&, 2) = XVALO(VALORX$)
        VALORX$ = Mid$(TXT$, 27, 8)
        XLSHEET.Cells(I&, 3) = XVALO(VALORX$)
        VALORX$ = Mid$(TXT$, 36, 13)
        XLSHEET.Cells(I&, 4) = XVALO(VALORX$)
        VALORX$ = Mid$(TXT$, 49, 13)
        XLSHEET.Cells(I&, 5) = XVALO(VALORX$)
        VALORX$ = Mid$(TXT$, 62, 13)
        XLSHEET.Cells(I&, 6) = XVALO(VALORX$)
        VALORX$ = Mid$(TXT$, 75, 8)
        XLSHEET.Cells(I&, 7) = XVALO(VALORX$)
        VALORX$ = Mid$(TXT$, 85, 8)
        XLSHEET.Cells(I&, 8) = XVALO(VALORX$)
   Next I&
   '
   NN% = (List1(1).ListCount - 1) + 4 + 2
   TXT$ = Label3(1)
   VALORX$ = "Total": XLSHEET.Cells(NN%, 1) = VALORX$
   VALORX$ = Mid$(TXT$, 21, 6): XLSHEET.Cells(NN%, 2) = XVALO(VALORX$)
   VALORX$ = Mid$(TXT$, 27, 6): XLSHEET.Cells(NN%, 3) = XVALO(VALORX$)
   VALORX$ = Mid$(TXT$, 36, 13): XLSHEET.Cells(NN%, 4) = XVALO(VALORX$)
   VALORX$ = Mid$(TXT$, 49, 13): XLSHEET.Cells(NN%, 5) = XVALO(VALORX$)
   VALORX$ = Mid$(TXT$, 62, 13): XLSHEET.Cells(NN%, 6) = XVALO(VALORX$)
   VALORX$ = Mid$(TXT$, 75, 8): XLSHEET.Cells(NN%, 7) = XVALO(VALORX$)
   VALORX$ = Mid$(TXT$, 85, 8): XLSHEET.Cells(NN%, 8) = XVALO(VALORX$)
   '
   On Error GoTo 30
   XLSHEET.SaveAs NOMBREDEARCHIVO$
   On Error GoTo 0
   '
   XLAPP.Quit      ' cierra EXCEL que se esta ejecutando minimizado
   Set XLAPP = Nothing     ') libera memoria
   Set XLBOOK = Nothing    ') utilizada por
   Set XLSHEET = Nothing   ') las variables objeto
   '
   Screen.MousePointer = 1
   Call COMUNI("Planilla Excel Creada Exitosamente:\ \" + REPLACAR$(NOMBREDEARCHIVO$, "\", "/"))
   '
Exit Sub
   '
   '
20 Resume 21
21 On Error GoTo 0
   MsgBox "Utilitario EXCEL no Instalado"
   On Error GoTo 0
   Exit Sub
   '
30 Resume 31
31 On Error GoTo 0
   Screen.MousePointer = 1
   '
   ' Cierra el objeto Book, el cual es utilizado al crear
   ' una nueva hoja en el archivo Excel.
   ' El valor False es para que no pregunte si quiere
   ' guardar el archivo.
   XLBOOK.Close (False)
   XLAPP.Quit ' Cierra Excel
   Set XLAPP = Nothing
   Set XLBOOK = Nothing
   Set XLSHEET = Nothing
   '
   Call MENSERR(24, "Imposible Crear Planilla Excel.")
   Exit Sub
   '
End Sub
