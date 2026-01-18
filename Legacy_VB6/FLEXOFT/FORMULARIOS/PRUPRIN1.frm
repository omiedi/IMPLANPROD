VERSION 5.00
Begin VB.Form PRUPRIN1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1095
      Left            =   1200
      TabIndex        =   0
      Top             =   600
      Width           =   2055
   End
End
Attribute VB_Name = "PRUPRIN1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
               TTXX$ = TTTT$ + Chr$(10) + Chr$(13)
               TTXX$ = TTXX$ + "SCALEMODE: " & Printer.ScaleMode & Chr$(10) & Chr$(13)
               TTXX$ = TTXX$ + "SCALEWIDTH: " & Printer.ScaleWidth & Chr$(10) & Chr$(13)
               TTXX$ = TTXX$ + "SCALELEFT: " & Printer.ScaleLeft & Chr$(10) & Chr$(13)
               TTXX$ = TTXX$ + "SCALEHEIGHT: " & Printer.ScaleHeight & Chr$(10) & Chr$(13)
               TTXX$ = TTXX$ + "HEIGHT: " & Printer.Height & Chr$(10) & Chr$(13)
               TTXX$ = TTXX$ + "SCALETOP: " & Printer.ScaleTop & Chr$(10) & Chr$(13)
               TTXX$ = TTXX$ + "PAPERSIZE: " & Printer.PaperSize
               MsgBox TTXX$
   Printer.Height = 15000
               TTXX$ = TTTT$ + Chr$(10) + Chr$(13)
               TTXX$ = TTXX$ + "SCALEMODE: " & Printer.ScaleMode & Chr$(10) & Chr$(13)
               TTXX$ = TTXX$ + "SCALEWIDTH: " & Printer.ScaleWidth & Chr$(10) & Chr$(13)
               TTXX$ = TTXX$ + "SCALELEFT: " & Printer.ScaleLeft & Chr$(10) & Chr$(13)
               TTXX$ = TTXX$ + "SCALEHEIGHT: " & Printer.ScaleHeight & Chr$(10) & Chr$(13)
               TTXX$ = TTXX$ + "HEIGHT: " & Printer.Height & Chr$(10) & Chr$(13)
               TTXX$ = TTXX$ + "SCALETOP: " & Printer.ScaleTop & Chr$(10) & Chr$(13)
               TTXX$ = TTXX$ + "PAPERSIZE: " & Printer.PaperSize
               MsgBox TTXX$
   
End Sub
