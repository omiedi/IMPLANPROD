VERSION 5.00
Begin VB.Form IMAGEN 
   Caption         =   "Form1"
   ClientHeight    =   8625
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   11715
   LinkTopic       =   "Form1"
   ScaleHeight     =   8625
   ScaleWidth      =   11715
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      Height          =   5790
      Left            =   6405
      ScaleHeight     =   5730
      ScaleWidth      =   4680
      TabIndex        =   3
      Top             =   525
      Width           =   4740
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   7470
      Left            =   210
      TabIndex        =   6
      Top             =   210
      Width           =   5475
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   7470
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   5415
      End
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      Height          =   2745
      Left            =   6405
      ScaleHeight     =   2685
      ScaleWidth      =   4680
      TabIndex        =   5
      Top             =   3675
      Width           =   4740
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   195
      Left            =   3150
      TabIndex        =   4
      Top             =   3360
      Width           =   1170
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   330
      Left            =   8085
      TabIndex        =   2
      Top             =   0
      Width           =   1170
   End
   Begin VB.TextBox Text1 
      Height          =   3375
      Left            =   3360
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "IMAGEN.frx":0000
      Top             =   4200
      Width           =   5685
   End
   Begin VB.PictureBox Picture1 
      Height          =   2640
      Left            =   420
      ScaleHeight     =   2580
      ScaleWidth      =   5550
      TabIndex        =   0
      Top             =   525
      Width           =   5610
   End
End
Attribute VB_Name = "IMAGEN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
   TEXTO$ = Text1.Text + Chr$(10)
   MsgBox TEXTO$
10 If Len(TEXTO$) < 1 Then Stop
   TTXX$ = ""
   For U& = 1 To Len(TEXTO$)
      If Asc(Mid$(TEXTO$, U&, 1)) < 32 Then
         If Len(TTXX$) > 0 Then
            MsgBox TTXX$
         End If
         TEXTO$ = Mid$(TEXTO$, U& + 1)
         GoTo 10
        Else
         TTXX$ = TTXX$ + Mid$(TEXTO$, U&, 1)
      End If
   Next U&

End Sub

Private Sub Command1_Click()
   Picture2.PaintPicture Image1.Picture, 0, 0, 25000, 25000
End Sub

Private Sub Form_Load()


   'Picture2.Picture = LoadPicture("A:S1.JPG")
   Image1.Picture = LoadPicture("\CLIENTES\AMBEST\MONTAJE1.JPG")
   
   
'   AAA = Picture2.Picture.Width
 '  AAA = Picture2.Height
  ' BBB = Picture2.Width
   Show
   'Picture1.Picture = Clipboard.GetDATA
   'Picture2.PaintPicture Picture2.Picture, 0, 0, 2000, 3000
   'Picture2.Refresh
   'Printer.EndDoc
   'For KK% = 0 To 4
   '  For X% = 0 To BBB Step 10
   '    For Y% = 0 To AAA Step 10
   '      Picture3.PSet (Y% + KK%, X% + KK%), Picture2.Point(X% + KK%, Y% + KK%)
   '    Next Y%
'Picture3.Refresh
 '    Next X%
   'Next KK%
End Sub

