VERSION 5.00
Begin VB.Form SEPRIN 
   Caption         =   "Form1"
   ClientHeight    =   3225
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   8730
   LinkTopic       =   "Form1"
   ScaleHeight     =   3225
   ScaleWidth      =   8730
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   315
      TabIndex        =   1
      Top             =   210
      Width           =   7470
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   435
      Left            =   6825
      TabIndex        =   0
      Top             =   2520
      Width           =   1170
   End
End
Attribute VB_Name = "SEPRIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FUFIJA$(128)
Dim TIFUE$(128)

Private Sub Command1_Click()
    '
    List1.Clear
    NOMIMPRE$ = Trim$(Printer.DeviceName)
    If NOMIMPRE$ = "" Then
        Close: End
        ' Call MENSERR(24, "Falta Seleccionar Impresora\en la Configuración de Windows")
        ' Call FINAL("")
    End If
    '
    YAESTA% = 0
    'Call BLANARCH("!PRINTERS")
    'NOMI$ = AJUSTI$(NOMIMPRE$, 48)
    'NX% = FILEOPEN("C:\FLEXOFT\PRINTERS.DRV", 0, 128)
    'For IPRIN& = 1 To LOF(NX%) / 128
    '    Get #NX%, IPRIN&, PPZZ$
    '    PPXX$ = PPZZ$
    '    If Trim$(Left$(PPXX$, 48)) = NOMIMPRE$ Then
    '        Call REGAPP("!PRINTERS", PPXX$)
    '        YAESTA% = 1
    '    End If
    'Next IPRIN&
    '
    'Close #NX%
    'If YAESTA% = 1 Then
    '    PPXX$ = Left$(PPXX$, 52)
    '    GoTo 50
    'End If
    '
    'Printer.ScaleMode = 1
    'ANCHOJMM = Int(Printer.ScaleWidth / 56.7)
    'Call REPLA(PPXX$, MKS$(ANCHOJMM), 49, 4)
    '
    'CAFUFIJ% = 0
    For I% = 1 To Printer.FontCount
        Printer.FontName = Printer.Fonts(I% - 1)
        If CAFUFIJ% < 128 Then
            CAFUFIJ% = CAFUFIJ% + 1
            FUFIJA$(CAFUFIJ%) = LTrim$(RTrim$(Printer.Fonts(I% - 1)))
            MsgBox FUFIJA$(CAFUFIJ%)
            A1 = Printer.TextWidth(String$(80, "l"))
            A2 = Printer.TextWidth(String$(80, "M"))
            If A1 = A2 Then
                TIFUE$(CAFUFIJ%) = " (fija)"
            End If
        End If
    Next I%
    '
    'If CAFUFIJ% < 1 Then
    '    Call MENSERR(24, "No se Encontraron Fuentes de Ancho Fijo\para la Impresora Seleccionada.\  \No es Posible Emitir Listado por Impresora.\  \Consulte a su Soporte de Sistemas.")
    '    Call FINAL("")
    'End If
    '
    'Call BLANARCH("!PRINTERS")
    TALEBASI = 24
    TEPRUEBA$ = String$(128, "*")
    Printer.FontBold = False
    Printer.FontItalic = False
    For I% = 1 To CAFUFIJ%
        Printer.FontName = FUFIJA$(I%)
        If Printer.FontName <> FUFIJA$(I%) Then GoTo 19
        ATEXTA = 0
        FUENTEMIN = 1024
        FUENTEMAX = 0
        '
        TALEBA = 24.4
        While TALEBA > 2
            TALEBA = TALEBA - 0.4
            Printer.FontSize = TALEBA
            If Printer.FontSize > TALEBA Then GoTo 13
            ATEXTO = Printer.TextWidth(TEPRUEBA$)
            If ATEXTO <> ATEXTA Then
                List1.AddItem FUFIJA$(I%) + TIFUE$(I%) + Str$(Printer.FontSize)
                ATEXTA = ATEXTO
            End If
            TALEBA = Printer.FontSize
13      Wend
19  Next I%
    '
End Sub
