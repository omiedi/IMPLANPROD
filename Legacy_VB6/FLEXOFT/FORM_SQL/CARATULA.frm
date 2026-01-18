VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form CARATULA 
   BackColor       =   &H00FEE7D6&
   Caption         =   "Databook"
   ClientHeight    =   10935
   ClientLeft      =   75
   ClientTop       =   285
   ClientWidth     =   13635
   ForeColor       =   &H00000000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   10935
   ScaleWidth      =   13635
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6720
      Top             =   5280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FEE7D6&
      Height          =   10935
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   13635
      Begin VB.CommandButton COMMAND3 
         Caption         =   "O.K."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   11160
         Picture         =   "CARATULA.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Grabar Carátula"
         Top             =   10200
         Width           =   975
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   12240
         Picture         =   "CARATULA.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Imprime Databook Completo"
         Top             =   10200
         Width           =   1200
      End
      Begin VB.PictureBox Picture1 
         Height          =   4030
         Left            =   120
         ScaleHeight     =   3975
         ScaleWidth      =   13275
         TabIndex        =   5
         Top             =   240
         Width           =   13335
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   4560
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   30
            TabStop         =   0   'False
            ToolTipText     =   " "
            Top             =   1530
            Width           =   1090
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   4560
            MultiLine       =   -1  'True
            TabIndex        =   28
            TabStop         =   0   'False
            Text            =   "CARATULA.frx":0974
            Top             =   2160
            Width           =   1095
         End
         Begin VB.TextBox TEXT33 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   640
            Left            =   10755
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   22
            TabStop         =   0   'False
            Text            =   "CARATULA.frx":0976
            Top             =   1500
            Width           =   2520
         End
         Begin VB.TextBox Text16 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   9480
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   9
            TabStop         =   0   'False
            Text            =   "CARATULA.frx":0978
            ToolTipText     =   " "
            Top             =   2130
            Width           =   1275
         End
         Begin VB.TextBox Text17 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   660
            Left            =   9480
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   10
            TabStop         =   0   'False
            Text            =   "CARATULA.frx":0985
            ToolTipText     =   " "
            Top             =   1485
            Width           =   1275
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   660
            Left            =   9480
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   17
            TabStop         =   0   'False
            Text            =   "CARATULA.frx":099C
            ToolTipText     =   " "
            Top             =   885
            Width           =   1275
         End
         Begin VB.TextBox Text2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   900
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   26
            Top             =   0
            Width           =   3075
         End
         Begin VB.CommandButton Command2 
            Caption         =   "."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   570
            Left            =   4320
            TabIndex        =   25
            Top             =   2160
            Width           =   255
         End
         Begin VB.TextBox Text22 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   3060
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   6
            TabStop         =   0   'False
            Text            =   "CARATULA.frx":09AD
            ToolTipText     =   " "
            Top             =   2715
            Width           =   2595
         End
         Begin VB.TextBox Text31 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   5640
            MultiLine       =   -1  'True
            TabIndex        =   20
            TabStop         =   0   'False
            Text            =   "CARATULA.frx":09DB
            Top             =   2720
            Width           =   7640
         End
         Begin VB.TextBox Text32 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   660
            Left            =   5640
            MultiLine       =   -1  'True
            TabIndex        =   24
            TabStop         =   0   'False
            Text            =   "CARATULA.frx":09DF
            Top             =   3300
            Width           =   7640
         End
         Begin VB.TextBox Text30 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   610
            Left            =   5640
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   19
            TabStop         =   0   'False
            Text            =   "CARATULA.frx":09E1
            Top             =   2130
            Width           =   3855
         End
         Begin VB.TextBox Text29 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   610
            Left            =   5640
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   18
            TabStop         =   0   'False
            Text            =   "CARATULA.frx":09E3
            Top             =   1530
            Width           =   3855
         End
         Begin VB.TextBox Text18 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   20.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   660
            Left            =   3060
            Locked          =   -1  'True
            TabIndex        =   16
            TabStop         =   0   'False
            Text            =   "DATA BOOK"
            ToolTipText     =   " "
            Top             =   885
            Width           =   6435
         End
         Begin VB.TextBox Text28 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   660
            Left            =   3060
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   15
            TabStop         =   0   'False
            Text            =   "CARATULA.frx":09E5
            ToolTipText     =   " "
            Top             =   1485
            Width           =   1515
         End
         Begin VB.TextBox Text27 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   660
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   14
            Text            =   "ARCON"
            Top             =   885
            Width           =   3075
         End
         Begin VB.TextBox Text26 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   27.75
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   900
            Left            =   3060
            Locked          =   -1  'True
            TabIndex        =   13
            TabStop         =   0   'False
            Text            =   "METALURGICASIAM S.A."
            ToolTipText     =   " "
            Top             =   0
            Width           =   10220
         End
         Begin VB.TextBox Text23 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   3060
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   12
            TabStop         =   0   'False
            Text            =   "CARATULA.frx":09F5
            ToolTipText     =   " "
            Top             =   2130
            Width           =   1515
         End
         Begin VB.TextBox Text21 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   660
            Left            =   3060
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   11
            TabStop         =   0   'False
            Text            =   "CARATULA.frx":0A0B
            ToolTipText     =   " "
            Top             =   3300
            Width           =   2600
         End
         Begin VB.TextBox Text24 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1585
            Left            =   0
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   8
            Text            =   "CARATULA.frx":0A20
            Top             =   2370
            Width           =   3075
         End
         Begin VB.TextBox Text25 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   27.75
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   825
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   7
            Text            =   "SIAM"
            Top             =   1560
            Width           =   3075
         End
         Begin VB.TextBox Text19 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   10755
            MultiLine       =   -1  'True
            TabIndex        =   21
            TabStop         =   0   'False
            Text            =   "CARATULA.frx":0A9D
            Top             =   2150
            Width           =   2520
         End
         Begin VB.TextBox Text20 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   10755
            MultiLine       =   -1  'True
            TabIndex        =   23
            TabStop         =   0   'False
            Text            =   "CARATULA.frx":0A9F
            Top             =   900
            Width           =   2520
         End
      End
      Begin MSFlexGridLib.MSFlexGrid MSF 
         Bindings        =   "CARATULA.frx":0AA1
         Height          =   5655
         Left            =   120
         TabIndex        =   4
         Top             =   4440
         Width           =   13305
         _ExtentX        =   23469
         _ExtentY        =   9975
         _Version        =   393216
         BackColor       =   16777215
         WordWrap        =   -1  'True
         FocusRect       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   1200
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   0
      Top             =   10440
      Visible         =   0   'False
      Width           =   1170
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   120
      OleObjectBlob   =   "CARATULA.frx":0AB5
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "CARATULA.frx":0CE9
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   435
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.Menu AGREGAR_RUTA 
      Caption         =   "Agregar Ruta"
      Visible         =   0   'False
      Begin VB.Menu mnuedi 
         Caption         =   "Editar"
      End
   End
End
Attribute VB_Name = "CARATULA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" _
                                    Alias "ShellExecuteA" _
                                    (ByVal hwnd As Long, _
                                    ByVal lpOperation As String, _
                                    ByVal lpFile As String, _
                                    ByVal lpParameters As String, _
                                    ByVal lpDirectory As String, _
                                    ByVal nShowCmd As Long) _
                                    As Long

Private Sub Command1_Click()
Dim oWord As Word.Application
    Dim oDoc As Word.Document
    Dim oTable As Word.Table
    Dim oPara1 As Word.Paragraph, oPara2 As Word.Paragraph
    Dim oPara3 As Word.Paragraph, oPara4 As Word.Paragraph
    Dim oRng As Word.Range
    Dim oShape As Word.InlineShape
    Dim oChart As Object
    Dim Pos As Double

    'Start Word and open the document template.
    Set oWord = CreateObject("Word.Application")
    oWord.Visible = True
    Set oDoc = oWord.Documents.Add
    
    'Insert a paragraph at the beginning of the document.
    Set oPara1 = oDoc.Content.Paragraphs.Add
    oPara1.Range.TEXT = "Heading 1"
    oPara1.Range.Font.Bold = True
    oPara1.Format.SpaceAfter = 24    '24 pt spacing after paragraph.
    oPara1.Range.InsertParagraphAfter
    
    'Insert a paragraph at the end of the document.
    '** \endofdoc is a predefined bookmark.
    Set oPara2 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks("\endofdoc").Range)
    oPara2.Range.TEXT = "Heading 2"
    oPara2.Format.SpaceAfter = 6
    oPara2.Range.InsertParagraphAfter
    
    'Insert another paragraph.
    Set oPara3 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks("\endofdoc").Range)
    oPara3.Range.TEXT = "This is a sentence of normal text. Now here is a table:"
    oPara3.Range.Font.Bold = False
    oPara3.Format.SpaceAfter = 24
    oPara3.Range.InsertParagraphAfter
    
    'Insert a 3 x 5 table, fill it with data and make the first row
    'bold,italic.
    Dim R As Integer, c As Integer
    Set oTable = oDoc.Tables.Add(oDoc.Bookmarks("\endofdoc").Range, 3, 5)
    oTable.Range.ParagraphFormat.SpaceAfter = 6
    For R = 1 To 3
        For c = 1 To 5
            oTable.Cell(R, c).Range.TEXT = "r" & R & "c" & c
        Next
    Next
    oTable.Rows(1).Range.Font.Bold = True
    oTable.Rows(1).Range.Font.Italic = True
    
    'Add some text after the table.
    'oTable.Range.InsertParagraphAfter
    Set oPara4 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks("\endofdoc").Range)
    oPara4.Range.InsertParagraphBefore
    oPara4.Range.TEXT = "And here's another table:"
    oPara4.Format.SpaceAfter = 24
    oPara4.Range.InsertParagraphAfter
    
    'Insert a 5 x 2 table, fill it with data and change the column widths.
    Set oTable = oDoc.Tables.Add(oDoc.Bookmarks("\endofdoc").Range, 5, 2)
    oTable.Range.ParagraphFormat.SpaceAfter = 6
    For R = 1 To 5
        For c = 1 To 2
            oTable.Cell(R, c).Range.TEXT = "r" & R & "c" & c
        Next
    Next
    oTable.Columns(1).Width = oWord.InchesToPoints(2)   'Change width of columns 1 & 2.
    oTable.Columns(2).Width = oWord.InchesToPoints(3)
    
    'Keep inserting text. When you get to 7 inches from top of the
    'document, insert a hard page break.
    Pos = oWord.InchesToPoints(7)
    oDoc.Bookmarks("\endofdoc").Range.InsertParagraphAfter
    Do
        Set oRng = oDoc.Bookmarks("\endofdoc").Range
        oRng.ParagraphFormat.SpaceAfter = 6
        oRng.InsertAfter "A line of text"
        oRng.InsertParagraphAfter
    Loop While Pos >= oRng.Information(wdVerticalPositionRelativeToPage)
    oRng.Collapse (wdCollapseEnd)
    oRng.InsertBreak wdPageBreak
    oRng.Collapse wdCollapseEnd
    oRng.InsertAfter "We're now on page 2. Here's my chart:"
    oRng.InsertParagraphAfter
    
    'Insert a chart and change the chart.
    Set oShape = oDoc.Bookmarks("\endofdoc").Range.InlineShapes.AddOLEObject( _
        ClassType:="MSGraph.Chart.8", FileName _
        :="", LinkToFile:=False, DisplayAsIcon:=False)
    Set oChart = oShape.OLEFormat.Object
    oChart.ChartType = 4 'xlLine = 4
    oChart.Application.Update
    oChart.Application.QUIT
    '... If desired, you can proceed from here using the Microsoft Graph
    'Object model on the oChart object to make additional changes to the
    'chart.
    oShape.Width = oWord.InchesToPoints(6.25)
    oShape.Height = oWord.InchesToPoints(3.57)
    
    'Add text after the chart.
    Set oRng = oDoc.Bookmarks("\endofdoc").Range
    oRng.InsertParagraphAfter
    oRng.InsertAfter "THE END."

    'All done. Unload this form.
    Unload Me

End Sub

Private Sub Command2_Click()
   Command2.Enabled = False
   Call SELECHO("DEUDOR1")
   NCLIE% = XVALO(TRIM$(VALACT1$("DEUDOR1")))
   Text3.TEXT = SAFETEXT$(NCLIE%)
   Command2.Enabled = True
End Sub

Private Sub Command3_Click()
       Command3.Enabled = False
       NORFA$ = TRIM$(Text2)
       NUSER1$ = TRIM$(Text33)
       If NORFA$ = "" Or NUSER1$ = "" Then GoTo 99
       
       SQLX$ = "SELECT * FROM CARATULA "
       SQLX$ = SQLX$ + " WHERE IDSUBOR = '" & NORFA$ & "'"
       SQLX$ = SQLX$ + " AND NUMEROSERIE  = '" & NUSER1$ & "'"
       Dim RS As ADODB.Recordset
       Set RS = New ADODB.Recordset
25     On Error Resume Next
       RS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText     ' SE MODIFICO EL adOpenDynamic POR EL adOpenKeyset
       '
       If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
         On Error GoTo 0
         Call CapturaErrorOpen
         GoTo 25
       End If
       On Error GoTo 0
       '
       With RS
            If .EOF Then
              .AddNew
            End If
            !NUCLIEN = XVALO(Text3)
            !idsubor = NORFA$
            !NUMEROSERIE = NUSER1$
            !DOMICLI = Text31
            !DESTINO = Text32
            !O_VENTA = Left$(Text19, 64)
            !PAGINAS = XVALO(Text20)
            '    Text4 = XVALO(!NUCLIEN)
'    Text31 = SAFETEXT$(!DOMICLI)
'    Text32 = SAFETEXT$(!DESTINO)
'    Text19 = SAFETEXT$(!DESTINO)
'    Text20 = XVALO(!HOJAS)

            .Update
        End With
   
99    Command3.Enabled = True
End Sub

Private Sub Command8_Click()
   Call GENERAPRESUWORD
End Sub
Sub GENERAPRESUWORD()

On Error GoTo errsub

    CommonDialog1.CancelError = True
    CommonDialog1.DialogTitle = "Ubicacion del Presupuesto"
    'CommonDialog1.FILTER = "Hojas de Cálculo Excel(*.XLS)|*.XLS|"
    CommonDialog1.FilterIndex = 1
    On Error Resume Next
    CommonDialog1.ShowOpen
    On Error GoTo 0
    RUTAWORD$ = TRIM$(CommonDialog1.FileName)
    If RUTAWORD$ = "" Then
        Call MENSERR(24, "Ubicacion del Presupuesto Erronea")
        Exit Sub
    End If
    '
    'Variable de tipo Word Aplication
    Dim o_Word As Word.Application
    'Variable de objeto para hacer referencia al documento
    Dim Documento As Word.Document
    'Variable de tipo tabla para poder hacer referencia a la tabla dentro del Word
    Dim Parrafo As Table
    'La F es para referenciar a las filas y la C para las columnas de la tabla de Word
    Dim F As Long
    Dim c As Integer
    Dim DATOSEMPRESA(4) As String
    
    'nuevo objeto de tipo Word Application
    Set o_Word = New Word.Application
    
    'hacemos visible el documento
    o_Word.Visible = True
    
    'Añadimos un Nuevo documento de word
    Set Documento = o_Word.Documents.Add

'    'ARMADO DEL ENCABEZADO
'    '*********************
    TXX1$ = Text26.TEXT & " (" & Text27.TEXT & " - " & Text25.TEXT & ")"
    Call AgregaLinea(Documento, TXX1$, Text26.Font, 15, True, wdAlignParagraphCenter)
    '
    Call REPLA(TX$, "Dr.A.Lavarello", 1, 16)
    Call REPLA(TX$, "1870 - Avellaneda - ", 17, 20)
    Call REPLA(TX$, "Prov. Bs.As.- Argentina", 35, 24)
    Call REPLA(TX2$, "Te. 54-11-4203-0011/6577", 1, 24)
    Call REPLA(TX2$, " - Email: info@ metales - arcon.com.ar", 25, 38)
    Call AgregaLinea(Documento, TX$, Text25.Font, 11, True, wdAlignParagraphLeft)
    Call AgregaLinea(Documento, TX2$, Text25.Font, 11, True, wdAlignParagraphLeft)
    '

    'CREA LA TABLA Y ESTABLECE EL ENCABEZADO DE LAS COLUMNAS
    Set Parrafo2 = Documento.Paragraphs.Add
    Set Rango = Parrafo2.Range.Next
    Set Parrafo2 = Documento.Tables.Add(Rango, 5, 4, wdWord9TableBehavior, AutoFitBehavior)
    For i = 1 To 5
      For J = 1 To 4
        Parrafo2.Cell(i, J).Range.Font.Size = 10
        Parrafo2.Cell(i, J).Range.Font.Bold = Negrita
      Next J
    Next i

    Parrafo2.Cell(1, 1).Range.Font.Size = 15
    Parrafo2.Cell(1, 1).Range.Font.Bold = Negrita
    Parrafo2.Cell(1, 1).Range.InsertAfter Text18
    Parrafo2.Cell(1, 1).Shading.BackgroundPatternColor = &HE0E0E0

    Parrafo2.Cell(1, 2).Range.InsertAfter ""
    Parrafo2.Cell(1, 3).Range.InsertAfter Text1
    Parrafo2.Cell(1, 4).Range.InsertAfter Text20
    
    Parrafo2.Cell(2, 1).Shading.BackgroundPatternColor = &HE0E0E0
    Parrafo2.Cell(2, 3).Shading.BackgroundPatternColor = &HE0E0E0
    Parrafo2.Cell(2, 1).Range.InsertAfter Text28
    Parrafo2.Cell(2, 2).Range.InsertAfter Text29
    Parrafo2.Cell(2, 3).Range.InsertAfter Text17
    Parrafo2.Cell(2, 4).Range.InsertAfter Text33
    '
    Parrafo2.Cell(3, 1).Shading.BackgroundPatternColor = &HE0E0E0
    Parrafo2.Cell(3, 3).Shading.BackgroundPatternColor = &HE0E0E0
    Parrafo2.Cell(3, 1).Range.InsertAfter Text23
    Parrafo2.Cell(3, 2).Range.InsertAfter Text30
    Parrafo2.Cell(3, 3).Range.InsertAfter Text16
    Parrafo2.Cell(3, 4).Range.InsertAfter Text19
    '
    Parrafo2.Cell(4, 1).Shading.BackgroundPatternColor = &HE0E0E0
    Parrafo2.Cell(4, 3).Shading.BackgroundPatternColor = &HE0E0E0
    Parrafo2.Cell(4, 1).Range.InsertAfter Text22
    Parrafo2.Cell(4, 2).Range.InsertAfter Text31
  
    Parrafo2.Cell(5, 1).Shading.BackgroundPatternColor = &HE0E0E0
    Parrafo2.Cell(5, 1).Range.InsertAfter Text21
    Parrafo2.Cell(5, 2).Range.InsertAfter Text32
    '
    'HASTA AQUI GRILLA DE ENCABEZADO
    '
    For H& = 0 To MSF.Rows - 1
      If H& >= MSF.Rows - 1 Then CANTREG& = CANTREG& + 1: Exit For
      If MSF.TextMatrix(H, 1) <> MSF.TextMatrix(H + 1, 1) Then
        CANTREG& = CANTREG& + 1
      End If
    Next H&
    Set Parrafo2 = Documento.Paragraphs.Add
    Set Rango = Parrafo2.Range.Next
    Set Parrafo2 = Documento.Tables.Add(Rango, CANTREG&, 3, wdWord9TableBehavior, AutoFitBehavior)
    For L& = 0 To MSF.Rows - 1
       familia$ = MSF.TextMatrix(L&, 1)
       TITULOX$ = MSF.TextMatrix(L&, 2)
       PAXINAS$ = MSF.TextMatrix(L&, 3)
       For k& = L& To MSF.Rows - 1
         If k& >= MSF.Rows - 1 Then GoTo 50
         If familia$ = MSF.TextMatrix(k& + 1, 1) Then
            TITULOX$ = TITULOX$ & vbCrLf & MSF.TextMatrix(k& + 1, 2)
            PAXINAS$ = PAXINAS$ & vbCrLf & MSF.TextMatrix(k& + 1, 3)
         Else
50         L& = k&
           Exit For
         End If
       Next k&
           Fila& = Fila& + 1
           If Fila& = 1 Then
             Parrafo2.Cell(Fila&, 1).Shading.BackgroundPatternColor = &HE0E0E0
             Parrafo2.Cell(Fila&, 2).Shading.BackgroundPatternColor = &HE0E0E0
             Parrafo2.Cell(Fila&, 3).Shading.BackgroundPatternColor = &HE0E0E0
           End If
           Parrafo2.Cell(Fila&, 1).Range.InsertAfter familia$
           Parrafo2.Cell(Fila&, 2).Range.InsertAfter TITULOX$
           Parrafo2.Cell(Fila&, 3).Range.InsertAfter PAXINAS$

    Next L&
10  Documento.SaveAs (RUTAWORD$ & ".DOC")
    Set o_Word = Nothing
    Set Documento = Nothing
    Set Parrafo = Nothing

Exit Sub

errsub:

MsgBox Err.Description, vbCritical
    On Error Resume Next
    Set Parrafo = Nothing
    Set Documento = Nothing
    Set o_Word = Nothing
End Sub

Sub AgregaLinea(Documento, Texto$, Fuente As Font, Tamanio, Negrita, alineacion)

    Set Parrafo = Documento.Paragraphs.Add
    Set Rango = Parrafo.Range.Next
    Rango.ParagraphFormat.Alignment = alineacion
    Rango.Font.Size = Tamanio
    Rango.Font.Bold = Negrita
    Rango.InsertBefore Texto$
'    Parrafo.LineSpacing = 0.7
End Sub


Function docCombinarCeldas(numceldas As Integer)
  Word.Selection.MoveRight Unit:=wdCharacter, Count:=numceldas, Extend:=wdExtend
  Word.Selection.Cells.Merge

End Function

Private Sub Form_Load()
   Call CARGA_ENCABEZADO(Me.MSF, "Familia/A32.0;Título Certificado/A64;Paginas/A12", Me)

End Sub

Sub Text2_Change()

       NORFA$ = TRIM$(Text2)
       NUSER1$ = TRIM$(Text33)
       '
       If NUSER1$ = "" Then Exit Sub
       SQLX$ = "SELECT * FROM CARATULA WITH(NOLOCK)"
       SQLX$ = SQLX$ + " WHERE IDSUBOR = '" & NORFA$ & "'"
       SQLX$ = SQLX$ + " AND NUMEROSERIE  = '" & NUSER1$ & "'"
       Set RS = READSET(SQLX$)       '
       With RS
            Text3 = XVALO(!NUCLIEN)
            Text31 = SAFETEXT$(!DOMICLI)
            Text32 = SAFETEXT$(!DESTINO)
            Text19 = SAFETEXT$(!O_VENTA)
            Text20 = XVALO(!PAGINAS)
       End With
       RS.Close
       Set RS = Nothing

End Sub

Private Sub Text3_Change()
   NCLIE% = XVALO(Text3)
   Text30.TEXT = rasocli$(NCLIE%)
End Sub

Sub TEXT33_Change()

End Sub
