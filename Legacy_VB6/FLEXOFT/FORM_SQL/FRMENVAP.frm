VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FRMENVAP 
   BackColor       =   &H00E2D3C0&
   Caption         =   "Ficha Técnica"
   ClientHeight    =   9525
   ClientLeft      =   11475
   ClientTop       =   1170
   ClientWidth     =   15720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9525
   ScaleWidth      =   15720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command444 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Left            =   7920
      TabIndex        =   221
      Top             =   780
      Width           =   175
   End
   Begin VB.TextBox TXTFUELLE 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7560
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   220
      Top             =   780
      Width           =   345
   End
   Begin VB.TextBox TXTLABEL1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   3
      Left            =   6840
      TabIndex        =   219
      TabStop         =   0   'False
      Text            =   "FUELLE"
      Top             =   780
      Width           =   690
   End
   Begin VB.TextBox TXTALERTA 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   3960
      TabIndex        =   218
      TabStop         =   0   'False
      Text            =   "Revise aquí"
      Top             =   2040
      Visible         =   0   'False
      Width           =   2460
   End
   Begin VB.TextBox TXTUSO 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   29
      Left            =   11350
      TabIndex        =   217
      TabStop         =   0   'False
      Top             =   9180
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.TextBox TXTUSO 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   28
      Left            =   11350
      TabIndex        =   216
      TabStop         =   0   'False
      Top             =   8880
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.TextBox TXTUSO 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   27
      Left            =   11350
      TabIndex        =   215
      TabStop         =   0   'False
      Top             =   8580
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.TextBox TXTUSO 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   26
      Left            =   11350
      TabIndex        =   214
      TabStop         =   0   'False
      Top             =   8280
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.TextBox TXTUSO 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   25
      Left            =   11350
      TabIndex        =   213
      TabStop         =   0   'False
      Top             =   7980
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.TextBox TXTUSO 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   24
      Left            =   11350
      TabIndex        =   212
      TabStop         =   0   'False
      Top             =   7680
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.TextBox TXTUSO 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   23
      Left            =   11350
      TabIndex        =   211
      TabStop         =   0   'False
      Top             =   7380
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.TextBox TXTUSO 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   22
      Left            =   11350
      TabIndex        =   210
      TabStop         =   0   'False
      Top             =   7080
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.TextBox TXTUSO 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   21
      Left            =   11350
      TabIndex        =   209
      TabStop         =   0   'False
      Top             =   6780
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.TextBox TXTUSO 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   20
      Left            =   11350
      TabIndex        =   208
      TabStop         =   0   'False
      Top             =   6480
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.TextBox TXTUSO 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   19
      Left            =   11350
      TabIndex        =   207
      TabStop         =   0   'False
      Top             =   6180
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.TextBox TXTUSO 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   18
      Left            =   11350
      TabIndex        =   206
      TabStop         =   0   'False
      Top             =   5880
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.TextBox TXTUSO 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   17
      Left            =   11350
      TabIndex        =   205
      TabStop         =   0   'False
      Top             =   5580
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.TextBox TXTUSO 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   16
      Left            =   11350
      TabIndex        =   204
      TabStop         =   0   'False
      Top             =   5280
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.TextBox TXTUSO 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   15
      Left            =   11350
      TabIndex        =   203
      TabStop         =   0   'False
      Top             =   4980
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.TextBox TXTUSO 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   14
      Left            =   11350
      TabIndex        =   202
      TabStop         =   0   'False
      Top             =   4680
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.TextBox TXTUSO 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   13
      Left            =   11350
      TabIndex        =   201
      TabStop         =   0   'False
      Top             =   4380
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.TextBox TXTUSO 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   12
      Left            =   11350
      TabIndex        =   200
      TabStop         =   0   'False
      Top             =   4080
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.TextBox TXTUSO 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   11
      Left            =   11350
      TabIndex        =   199
      TabStop         =   0   'False
      Top             =   3780
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.TextBox TXTUSO 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   10
      Left            =   11350
      TabIndex        =   198
      TabStop         =   0   'False
      Top             =   3480
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.TextBox TXTUSO 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   9
      Left            =   11350
      TabIndex        =   197
      TabStop         =   0   'False
      Top             =   3180
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.TextBox TXTUSO 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   8
      Left            =   11350
      TabIndex        =   196
      TabStop         =   0   'False
      Top             =   2880
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.TextBox TXTUSO 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   7
      Left            =   11350
      TabIndex        =   195
      TabStop         =   0   'False
      Top             =   2580
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.TextBox TXTUSO 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   6
      Left            =   11350
      TabIndex        =   194
      TabStop         =   0   'False
      Top             =   2280
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.TextBox TXTUSO 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   5
      Left            =   11350
      TabIndex        =   193
      TabStop         =   0   'False
      Top             =   1980
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.TextBox TXTUSO 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   4
      Left            =   11350
      TabIndex        =   192
      TabStop         =   0   'False
      Top             =   1680
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.TextBox TXTUSO 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   3
      Left            =   11350
      TabIndex        =   191
      TabStop         =   0   'False
      Top             =   1380
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.TextBox TXTUSO 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   2
      Left            =   11350
      TabIndex        =   190
      TabStop         =   0   'False
      Top             =   1080
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.TextBox TXTUSO 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   1
      Left            =   11350
      TabIndex        =   189
      TabStop         =   0   'False
      Top             =   780
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.TextBox TXTUSO 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   0
      Left            =   11350
      TabIndex        =   188
      TabStop         =   0   'False
      Top             =   360
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.TextBox TXTUNIMED 
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
      Height          =   280
      Left            =   14400
      Locked          =   -1  'True
      TabIndex        =   187
      TabStop         =   0   'False
      Text            =   "M."
      Top             =   6120
      Width           =   435
   End
   Begin VB.TextBox TXTLABEL1 
      Alignment       =   1  'Right Justify
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
      Height          =   280
      Index           =   170
      Left            =   12600
      Locked          =   -1  'True
      TabIndex        =   186
      TabStop         =   0   'False
      Text            =   "CORTE:"
      Top             =   6120
      Width           =   795
   End
   Begin VB.TextBox TXTCORTE 
      Alignment       =   1  'Right Justify
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
      Height          =   280
      Left            =   13440
      Locked          =   -1  'True
      TabIndex        =   185
      TabStop         =   0   'False
      Top             =   6120
      Width           =   885
   End
   Begin VB.TextBox TXTLABEL 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   10
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   183
      TabStop         =   0   'False
      Text            =   "COLA 3 "
      Top             =   3480
      Width           =   1800
   End
   Begin VB.TextBox TXTDESC 
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
      Height          =   280
      Index           =   10
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   182
      TabStop         =   0   'False
      Top             =   3480
      Width           =   3450
   End
   Begin VB.TextBox TXTCAJA 
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
      Height          =   280
      Index           =   10
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   181
      TabStop         =   0   'False
      Top             =   3480
      Width           =   2000
   End
   Begin VB.CommandButton Command4 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   10
      Left            =   3915
      TabIndex        =   180
      Top             =   3480
      Width           =   175
   End
   Begin VB.TextBox TXTOBS 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   10
      Left            =   7560
      TabIndex        =   179
      TabStop         =   0   'False
      Top             =   3480
      Width           =   3800
   End
   Begin VB.TextBox TXTLABEL 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   9
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   178
      TabStop         =   0   'False
      Text            =   "COLA 2 "
      Top             =   3180
      Width           =   1800
   End
   Begin VB.TextBox TXTDESC 
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
      Height          =   285
      Index           =   9
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   177
      TabStop         =   0   'False
      Top             =   3180
      Width           =   3450
   End
   Begin VB.TextBox TXTCAJA 
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
      Height          =   285
      Index           =   9
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   176
      TabStop         =   0   'False
      Top             =   3180
      Width           =   2000
   End
   Begin VB.CommandButton Command4 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   9
      Left            =   3915
      TabIndex        =   175
      Top             =   3180
      Width           =   175
   End
   Begin VB.TextBox TXTOBS 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   9
      Left            =   7560
      TabIndex        =   174
      TabStop         =   0   'False
      Top             =   3180
      Width           =   3800
   End
   Begin VB.TextBox TXTDESCRICOD 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   11640
      TabIndex        =   173
      TabStop         =   0   'False
      Top             =   450
      Width           =   4005
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   7560
      Top             =   4440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FRAME17 
      BackColor       =   &H00E2D3C0&
      Caption         =   "Adjuntar Archivos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Index           =   7
      Left            =   11640
      TabIndex        =   169
      Top             =   6480
      Width           =   3975
      Begin VB.CommandButton Command47 
         Caption         =   "Agregar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   1440
         TabIndex        =   172
         Top             =   1605
         Width           =   1185
      End
      Begin VB.CommandButton Command47 
         Caption         =   "Quitar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   2700
         TabIndex        =   171
         Top             =   1605
         Width           =   1125
      End
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1290
         ItemData        =   "FRMENVAP.frx":0000
         Left            =   240
         List            =   "FRMENVAP.frx":0002
         TabIndex        =   170
         Top             =   240
         Width           =   3630
      End
   End
   Begin VB.TextBox TXTMED 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   3160
      MaxLength       =   10
      TabIndex        =   162
      Top             =   780
      Width           =   950
   End
   Begin VB.TextBox TXTLABEL1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   5310
      TabIndex        =   168
      TabStop         =   0   'False
      Text            =   "Corte M."
      Top             =   780
      Width           =   615
   End
   Begin VB.TextBox TXTLABEL1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   4060
      TabIndex        =   167
      TabStop         =   0   'False
      Text            =   "X"
      Top             =   780
      Width           =   300
   End
   Begin VB.TextBox TXTLABEL1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   2860
      TabIndex        =   166
      TabStop         =   0   'False
      Text            =   "X"
      Top             =   780
      Width           =   300
   End
   Begin VB.TextBox TXTMED 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   5925
      MaxLength       =   10
      TabIndex        =   164
      Top             =   780
      Width           =   950
   End
   Begin VB.TextBox TXTMED 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   4360
      MaxLength       =   10
      TabIndex        =   163
      Top             =   780
      Width           =   950
   End
   Begin VB.TextBox TXTMED 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   1920
      MaxLength       =   10
      TabIndex        =   161
      Top             =   780
      Width           =   950
   End
   Begin VB.TextBox TXTDESC 
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
      Height          =   280
      Index           =   1
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   75
      TabStop         =   0   'False
      Top             =   780
      Width           =   4960
   End
   Begin VB.CommandButton Command44 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   11880
      TabIndex        =   160
      Top             =   50
      Width           =   175
   End
   Begin VB.TextBox TXTCODINT 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   1800
      TabIndex        =   159
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.TextBox TXTOBS 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   32
      Left            =   13440
      TabIndex        =   158
      TabStop         =   0   'False
      Top             =   5760
      Width           =   2205
   End
   Begin VB.TextBox TXTOBS 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   31
      Left            =   945
      Locked          =   -1  'True
      TabIndex        =   157
      TabStop         =   0   'False
      Top             =   50
      Width           =   900
   End
   Begin VB.TextBox TXTLABEL 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   31
      Left            =   105
      Locked          =   -1  'True
      TabIndex        =   156
      TabStop         =   0   'False
      Text            =   "CLIENTE"
      Top             =   50
      Width           =   840
   End
   Begin VB.TextBox TXTRAZOCLI 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   155
      TabStop         =   0   'False
      Top             =   50
      Width           =   5940
   End
   Begin VB.TextBox TXTLABEL 
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
      Height          =   280
      Index           =   32
      Left            =   11640
      Locked          =   -1  'True
      TabIndex        =   154
      TabStop         =   0   'False
      Text            =   "CÓDIGO DE CLIENTE"
      Top             =   5760
      Width           =   1755
   End
   Begin VB.TextBox TXTCODIGO 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   12120
      TabIndex        =   153
      TabStop         =   0   'False
      Top             =   50
      Width           =   3525
   End
   Begin VB.TextBox TXTLABEL1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   29
      Left            =   10560
      TabIndex        =   152
      TabStop         =   0   'False
      Text            =   "CÓDIGO"
      Top             =   50
      Width           =   1485
   End
   Begin VB.TextBox TXTFECHAPROBO 
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
      Height          =   280
      Left            =   14640
      Locked          =   -1  'True
      TabIndex        =   151
      TabStop         =   0   'False
      Top             =   9180
      Width           =   1000
   End
   Begin VB.TextBox TXTFECHREVISO 
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
      Height          =   280
      Left            =   14640
      Locked          =   -1  'True
      TabIndex        =   150
      TabStop         =   0   'False
      Top             =   8880
      Width           =   1000
   End
   Begin VB.TextBox TXTLABEL1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   35
      Left            =   11640
      Locked          =   -1  'True
      TabIndex        =   149
      TabStop         =   0   'False
      Text            =   "PREPARÓ"
      Top             =   8580
      Width           =   1000
   End
   Begin VB.TextBox TXTLABEL1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   34
      Left            =   11640
      Locked          =   -1  'True
      TabIndex        =   148
      TabStop         =   0   'False
      Text            =   "REVISÓ"
      Top             =   8880
      Width           =   1000
   End
   Begin VB.TextBox TXTLABEL1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   31
      Left            =   11640
      Locked          =   -1  'True
      TabIndex        =   147
      TabStop         =   0   'False
      Text            =   "APROBÓ"
      Top             =   9180
      Width           =   1000
   End
   Begin VB.TextBox TXTFECHPREPARO 
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
      Height          =   280
      Left            =   14640
      Locked          =   -1  'True
      TabIndex        =   146
      TabStop         =   0   'False
      Top             =   8580
      Width           =   1000
   End
   Begin VB.TextBox TXTPREPARO 
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
      Height          =   280
      Left            =   12600
      Locked          =   -1  'True
      TabIndex        =   145
      TabStop         =   0   'False
      Top             =   8580
      Width           =   2000
   End
   Begin VB.TextBox TXTREVISO 
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
      Height          =   280
      Left            =   12600
      Locked          =   -1  'True
      TabIndex        =   144
      TabStop         =   0   'False
      Top             =   8880
      Width           =   2000
   End
   Begin VB.TextBox TXTAPROBO 
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
      Height          =   280
      Left            =   12600
      Locked          =   -1  'True
      TabIndex        =   143
      TabStop         =   0   'False
      ToolTipText     =   "La Aprobación Actualiza la Estructura del Item."
      Top             =   9180
      Width           =   2000
   End
   Begin VB.TextBox TXTOBS 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   29
      Left            =   7560
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   9180
      Width           =   3800
   End
   Begin VB.TextBox TXTOBS 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   1
      Left            =   8160
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   780
      Width           =   3195
   End
   Begin VB.TextBox TXTOBS 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   2
      Left            =   7560
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1080
      Width           =   3800
   End
   Begin VB.TextBox TXTOBS 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   3
      Left            =   7560
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1380
      Width           =   3800
   End
   Begin VB.TextBox TXTOBS 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   4
      Left            =   7560
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1680
      Width           =   3800
   End
   Begin VB.TextBox TXTOBS 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   5
      Left            =   7560
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1980
      Width           =   3800
   End
   Begin VB.TextBox TXTOBS 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   6
      Left            =   7560
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2280
      Width           =   3800
   End
   Begin VB.TextBox TXTOBS 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   7
      Left            =   7560
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2580
      Width           =   3800
   End
   Begin VB.TextBox TXTOBS 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   8
      Left            =   7560
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2880
      Width           =   3800
   End
   Begin VB.TextBox TXTOBS 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   11
      Left            =   7560
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   3780
      Width           =   3800
   End
   Begin VB.TextBox TXTOBS 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   12
      Left            =   7560
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   4080
      Width           =   3800
   End
   Begin VB.TextBox TXTOBS 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   13
      Left            =   7560
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   4380
      Width           =   3800
   End
   Begin VB.TextBox TXTOBS 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   14
      Left            =   7560
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   4680
      Width           =   3800
   End
   Begin VB.TextBox TXTOBS 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   15
      Left            =   7560
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   4980
      Width           =   3800
   End
   Begin VB.TextBox TXTOBS 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   16
      Left            =   7560
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   5280
      Width           =   3800
   End
   Begin VB.TextBox TXTLABEL1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   169
      Left            =   7560
      TabIndex        =   142
      TabStop         =   0   'False
      Text            =   "COMENTARIO"
      Top             =   480
      Width           =   4000
   End
   Begin VB.TextBox TXTOBS 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   17
      Left            =   7560
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   5580
      Width           =   3800
   End
   Begin VB.TextBox TXTOBS 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   18
      Left            =   7560
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   5880
      Width           =   3800
   End
   Begin VB.TextBox TXTOBS 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   19
      Left            =   7560
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   6180
      Width           =   3800
   End
   Begin VB.TextBox TXTOBS 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   20
      Left            =   7560
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   6480
      Width           =   3800
   End
   Begin VB.TextBox TXTOBS 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   21
      Left            =   7560
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   6780
      Width           =   3800
   End
   Begin VB.TextBox TXTOBS 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   22
      Left            =   7560
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   7080
      Width           =   3800
   End
   Begin VB.TextBox TXTOBS 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   23
      Left            =   7560
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   7380
      Width           =   3800
   End
   Begin VB.TextBox TXTOBS 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   24
      Left            =   7560
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   7680
      Width           =   3800
   End
   Begin VB.TextBox TXTOBS 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   25
      Left            =   7560
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   7980
      Width           =   3800
   End
   Begin VB.TextBox TXTOBS 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   26
      Left            =   7560
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   8280
      Width           =   3800
   End
   Begin VB.TextBox TXTOBS 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   27
      Left            =   7560
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   8580
      Width           =   3800
   End
   Begin VB.TextBox TXTOBS 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   28
      Left            =   7560
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   8880
      Width           =   3800
   End
   Begin VB.CommandButton Command4 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   26
      Left            =   3920
      TabIndex        =   129
      Top             =   8280
      Width           =   175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   24
      Left            =   3920
      TabIndex        =   128
      Top             =   7680
      Width           =   175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   29
      Left            =   3920
      TabIndex        =   127
      Top             =   9180
      Width           =   175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   28
      Left            =   3920
      TabIndex        =   126
      Top             =   8880
      Width           =   175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   27
      Left            =   3920
      TabIndex        =   125
      Top             =   8580
      Width           =   175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   25
      Left            =   3920
      TabIndex        =   124
      Top             =   7980
      Width           =   175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   23
      Left            =   3920
      TabIndex        =   123
      Top             =   7380
      Width           =   175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   22
      Left            =   3920
      TabIndex        =   122
      Top             =   7080
      Width           =   175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   21
      Left            =   3920
      TabIndex        =   121
      Top             =   6780
      Width           =   175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   20
      Left            =   3920
      TabIndex        =   120
      Top             =   6480
      Width           =   175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   19
      Left            =   3920
      TabIndex        =   119
      Top             =   6180
      Width           =   175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   18
      Left            =   3920
      TabIndex        =   118
      Top             =   5880
      Width           =   175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   17
      Left            =   3920
      TabIndex        =   117
      Top             =   5580
      Width           =   175
   End
   Begin VB.TextBox TXTCAJA 
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
      Height          =   280
      Index           =   29
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   141
      TabStop         =   0   'False
      Top             =   9180
      Width           =   2000
   End
   Begin VB.TextBox TXTCAJA 
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
      Height          =   280
      Index           =   28
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   140
      TabStop         =   0   'False
      Top             =   8880
      Width           =   2000
   End
   Begin VB.TextBox TXTCAJA 
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
      Height          =   280
      Index           =   27
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   139
      TabStop         =   0   'False
      Top             =   8580
      Width           =   2000
   End
   Begin VB.TextBox TXTCAJA 
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
      Height          =   280
      Index           =   26
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   138
      TabStop         =   0   'False
      Top             =   8280
      Width           =   2000
   End
   Begin VB.TextBox TXTCAJA 
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
      Height          =   280
      Index           =   25
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   137
      TabStop         =   0   'False
      Top             =   7980
      Width           =   2000
   End
   Begin VB.TextBox TXTCAJA 
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
      Height          =   280
      Index           =   24
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   136
      TabStop         =   0   'False
      Top             =   7680
      Width           =   2000
   End
   Begin VB.TextBox TXTCAJA 
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
      Height          =   280
      Index           =   23
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   135
      TabStop         =   0   'False
      Top             =   7380
      Width           =   2000
   End
   Begin VB.TextBox TXTCAJA 
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
      Height          =   280
      Index           =   22
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   134
      TabStop         =   0   'False
      Top             =   7080
      Width           =   2000
   End
   Begin VB.TextBox TXTCAJA 
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
      Height          =   280
      Index           =   21
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   133
      TabStop         =   0   'False
      Top             =   6780
      Width           =   2000
   End
   Begin VB.TextBox TXTCAJA 
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
      Height          =   280
      Index           =   20
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   132
      TabStop         =   0   'False
      Top             =   6480
      Width           =   2000
   End
   Begin VB.TextBox TXTCAJA 
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
      Height          =   280
      Index           =   19
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   131
      TabStop         =   0   'False
      Top             =   6180
      Width           =   2000
   End
   Begin VB.TextBox TXTCAJA 
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
      Height          =   280
      Index           =   18
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   130
      TabStop         =   0   'False
      Top             =   5880
      Width           =   2000
   End
   Begin VB.CommandButton Command4 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   16
      Left            =   3920
      TabIndex        =   116
      Top             =   5280
      Width           =   175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   15
      Left            =   3920
      TabIndex        =   115
      Top             =   4980
      Width           =   175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   14
      Left            =   3920
      TabIndex        =   114
      Top             =   4680
      Width           =   175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   13
      Left            =   3920
      TabIndex        =   113
      Top             =   4380
      Width           =   175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   12
      Left            =   3920
      TabIndex        =   112
      Top             =   4080
      Width           =   175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   11
      Left            =   3920
      TabIndex        =   111
      Top             =   3780
      Width           =   175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   8
      Left            =   3920
      TabIndex        =   110
      Top             =   2880
      Width           =   175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   7
      Left            =   3920
      TabIndex        =   109
      Top             =   2580
      Width           =   175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   6
      Left            =   3920
      TabIndex        =   108
      Top             =   2280
      Width           =   175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   5
      Left            =   3920
      TabIndex        =   107
      Top             =   1980
      Width           =   175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   4
      Left            =   3920
      TabIndex        =   106
      Top             =   1680
      Width           =   175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   3
      Left            =   3920
      TabIndex        =   105
      Top             =   1380
      Width           =   175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   2
      Left            =   3920
      TabIndex        =   104
      Top             =   1080
      Width           =   175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   1
      Left            =   3920
      TabIndex        =   103
      Top             =   780
      Width           =   175
   End
   Begin VB.TextBox TXTDESC 
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
      Height          =   280
      Index           =   29
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   102
      TabStop         =   0   'False
      Top             =   9180
      Width           =   3450
   End
   Begin VB.TextBox TXTDESC 
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
      Height          =   280
      Index           =   28
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   101
      TabStop         =   0   'False
      Top             =   8880
      Width           =   3450
   End
   Begin VB.TextBox TXTDESC 
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
      Height          =   280
      Index           =   27
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   100
      TabStop         =   0   'False
      Top             =   8580
      Width           =   3450
   End
   Begin VB.TextBox TXTDESC 
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
      Height          =   280
      Index           =   26
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   99
      TabStop         =   0   'False
      Top             =   8280
      Width           =   3450
   End
   Begin VB.TextBox TXTDESC 
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
      Height          =   280
      Index           =   25
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   98
      TabStop         =   0   'False
      Top             =   7980
      Width           =   3450
   End
   Begin VB.TextBox TXTDESC 
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
      Height          =   280
      Index           =   24
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   97
      TabStop         =   0   'False
      Top             =   7680
      Width           =   3450
   End
   Begin VB.TextBox TXTDESC 
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
      Height          =   280
      Index           =   23
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   96
      TabStop         =   0   'False
      Top             =   7380
      Width           =   3450
   End
   Begin VB.TextBox TXTDESC 
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
      Height          =   280
      Index           =   22
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   95
      TabStop         =   0   'False
      Top             =   7080
      Width           =   3450
   End
   Begin VB.TextBox TXTDESC 
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
      Height          =   280
      Index           =   21
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   94
      TabStop         =   0   'False
      Top             =   6780
      Width           =   3450
   End
   Begin VB.TextBox TXTDESC 
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
      Height          =   280
      Index           =   20
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   93
      TabStop         =   0   'False
      Top             =   6480
      Width           =   3450
   End
   Begin VB.TextBox TXTDESC 
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
      Height          =   280
      Index           =   19
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   92
      TabStop         =   0   'False
      Top             =   6180
      Width           =   3450
   End
   Begin VB.TextBox TXTDESC 
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
      Height          =   280
      Index           =   18
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   91
      TabStop         =   0   'False
      Top             =   5880
      Width           =   3450
   End
   Begin VB.TextBox TXTDESC 
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
      Height          =   280
      Index           =   17
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   90
      TabStop         =   0   'False
      Top             =   5580
      Width           =   3450
   End
   Begin VB.TextBox TXTLABEL1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   33
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   89
      TabStop         =   0   'False
      Text            =   "DESCRIPCIÓN"
      Top             =   480
      Width           =   3450
   End
   Begin VB.TextBox TXTDESC 
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
      Height          =   280
      Index           =   16
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   88
      TabStop         =   0   'False
      Top             =   5280
      Width           =   3450
   End
   Begin VB.TextBox TXTDESC 
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
      Height          =   280
      Index           =   15
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   87
      TabStop         =   0   'False
      Top             =   4980
      Width           =   3450
   End
   Begin VB.TextBox TXTDESC 
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
      Height          =   280
      Index           =   14
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   86
      TabStop         =   0   'False
      Top             =   4680
      Width           =   3450
   End
   Begin VB.TextBox TXTDESC 
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
      Height          =   280
      Index           =   13
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   85
      TabStop         =   0   'False
      Top             =   4380
      Width           =   3450
   End
   Begin VB.TextBox TXTDESC 
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
      Height          =   280
      Index           =   12
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   84
      TabStop         =   0   'False
      Top             =   4080
      Width           =   3450
   End
   Begin VB.TextBox TXTDESC 
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
      Height          =   280
      Index           =   11
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   83
      TabStop         =   0   'False
      Top             =   3780
      Width           =   3450
   End
   Begin VB.TextBox TXTDESC 
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
      Height          =   280
      Index           =   8
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   82
      TabStop         =   0   'False
      Top             =   2880
      Width           =   3450
   End
   Begin VB.TextBox TXTDESC 
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
      Height          =   280
      Index           =   7
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   81
      TabStop         =   0   'False
      Top             =   2580
      Width           =   3450
   End
   Begin VB.TextBox TXTDESC 
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
      Height          =   280
      Index           =   6
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   80
      TabStop         =   0   'False
      Top             =   2280
      Width           =   3450
   End
   Begin VB.TextBox TXTDESC 
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
      Height          =   280
      Index           =   5
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   79
      TabStop         =   0   'False
      Top             =   1980
      Width           =   3450
   End
   Begin VB.TextBox TXTDESC 
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
      Height          =   280
      Index           =   4
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   78
      TabStop         =   0   'False
      Top             =   1680
      Width           =   3450
   End
   Begin VB.TextBox TXTDESC 
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
      Height          =   280
      Index           =   3
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   77
      TabStop         =   0   'False
      Top             =   1380
      Width           =   3450
   End
   Begin VB.TextBox TXTDESC 
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
      Height          =   280
      Index           =   2
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   76
      TabStop         =   0   'False
      Top             =   1080
      Width           =   3450
   End
   Begin VB.TextBox TXTOBS 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4440
      Index           =   30
      Left            =   11640
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   74
      Top             =   1200
      Width           =   4000
   End
   Begin VB.TextBox TXTLABEL 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   30
      Left            =   11640
      Locked          =   -1  'True
      TabIndex        =   73
      TabStop         =   0   'False
      Text            =   "OBSERVACIONES"
      Top             =   840
      Width           =   4000
   End
   Begin VB.TextBox TXTLABEL 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   29
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   72
      TabStop         =   0   'False
      Text            =   "CERT. CALIDAD"
      Top             =   9180
      Width           =   1800
   End
   Begin VB.TextBox TXTLABEL 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   28
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   71
      TabStop         =   0   'False
      Text            =   "BOLSAS X PALLET"
      Top             =   8880
      Width           =   1800
   End
   Begin VB.TextBox TXTLABEL 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   27
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   70
      TabStop         =   0   'False
      Text            =   "BOLSAS X PAQ."
      Top             =   8580
      Width           =   1800
   End
   Begin VB.TextBox TXTLABEL 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   25
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   69
      TabStop         =   0   'False
      Text            =   "COLOR 6 "
      Top             =   7980
      Width           =   1800
   End
   Begin VB.TextBox TXTLABEL 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   26
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   68
      TabStop         =   0   'False
      Text            =   "PAQ. ENFUNDADOS "
      Top             =   8280
      Width           =   1800
   End
   Begin VB.TextBox TXTLABEL 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   24
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   67
      TabStop         =   0   'False
      Text            =   "COLOR 5 "
      Top             =   7680
      Width           =   1800
   End
   Begin VB.TextBox TXTLABEL 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   23
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   66
      TabStop         =   0   'False
      Text            =   "COLOR 4 "
      Top             =   7380
      Width           =   1800
   End
   Begin VB.TextBox TXTLABEL 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   21
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   65
      TabStop         =   0   'False
      Text            =   "COLOR 2 "
      Top             =   6780
      Width           =   1800
   End
   Begin VB.TextBox TXTLABEL 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   22
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   64
      TabStop         =   0   'False
      Text            =   "COLOR 3 "
      Top             =   7080
      Width           =   1800
   End
   Begin VB.TextBox TXTLABEL 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   20
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   63
      TabStop         =   0   'False
      Text            =   "COLOR 1 "
      Top             =   6480
      Width           =   1800
   End
   Begin VB.TextBox TXTLABEL 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   19
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   62
      TabStop         =   0   'False
      Text            =   "COSIDO "
      Top             =   6180
      Width           =   1800
   End
   Begin VB.TextBox TXTLABEL 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   17
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   61
      TabStop         =   0   'False
      Text            =   "ABREFACIL "
      Top             =   5580
      Width           =   1800
   End
   Begin VB.TextBox TXTLABEL 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   18
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   60
      TabStop         =   0   'False
      Text            =   "POLIETILENO "
      Top             =   5880
      Width           =   1800
   End
   Begin VB.TextBox TXTCAJA 
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
      Height          =   280
      Index           =   17
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   59
      TabStop         =   0   'False
      Top             =   5580
      Width           =   2000
   End
   Begin VB.TextBox TXTLABEL 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   15
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   58
      TabStop         =   0   'False
      Text            =   "PABILO "
      Top             =   4980
      Width           =   1800
   End
   Begin VB.TextBox TXTLABEL 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   16
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   57
      TabStop         =   0   'False
      Text            =   "CREPPE "
      Top             =   5280
      Width           =   1800
   End
   Begin VB.TextBox TXTCAJA 
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
      Height          =   280
      Index           =   16
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   5280
      Width           =   2000
   End
   Begin VB.TextBox TXTCAJA 
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
      Height          =   280
      Index           =   15
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   55
      TabStop         =   0   'False
      Top             =   4980
      Width           =   2000
   End
   Begin VB.TextBox TXTLABEL 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   14
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   54
      TabStop         =   0   'False
      Text            =   "HILADO "
      Top             =   4680
      Width           =   1800
   End
   Begin VB.TextBox TXTCAJA 
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
      Height          =   280
      Index           =   14
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   53
      TabStop         =   0   'False
      Top             =   4680
      Width           =   2000
   End
   Begin VB.TextBox TXTLABEL 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   13
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   52
      TabStop         =   0   'False
      Text            =   "PLIEGOS PEGADOS "
      Top             =   4380
      Width           =   1800
   End
   Begin VB.TextBox TXTCAJA 
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
      Height          =   280
      Index           =   13
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   4380
      Width           =   2000
   End
   Begin VB.TextBox TXTLABEL 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   11
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   50
      TabStop         =   0   'False
      Text            =   "VALVULAS "
      Top             =   3780
      Width           =   1800
   End
   Begin VB.TextBox TXTLABEL 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   12
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   49
      TabStop         =   0   'False
      Text            =   "MANGA "
      Top             =   4080
      Width           =   1800
   End
   Begin VB.TextBox TXTCAJA 
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
      Height          =   280
      Index           =   12
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   4080
      Width           =   2000
   End
   Begin VB.TextBox TXTCAJA 
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
      Height          =   280
      Index           =   11
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   3780
      Width           =   2000
   End
   Begin VB.TextBox TXTLABEL 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   8
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   46
      TabStop         =   0   'False
      Text            =   "COLA 1 "
      Top             =   2880
      Width           =   1800
   End
   Begin VB.TextBox TXTCAJA 
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
      Height          =   280
      Index           =   8
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   2880
      Width           =   2000
   End
   Begin VB.TextBox TXTLABEL 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   7
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   44
      TabStop         =   0   'False
      Text            =   "MICROPERFORADO "
      Top             =   2580
      Width           =   1800
   End
   Begin VB.TextBox TXTCAJA 
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
      Height          =   280
      Index           =   7
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   2580
      Width           =   2000
   End
   Begin VB.TextBox TXTLABEL 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   5
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   42
      TabStop         =   0   'False
      Text            =   "PLIEGO 4 "
      Top             =   1980
      Width           =   1800
   End
   Begin VB.TextBox TXTLABEL 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   6
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   41
      TabStop         =   0   'False
      Text            =   "PLIEGO 5 "
      Top             =   2280
      Width           =   1800
   End
   Begin VB.TextBox TXTCAJA 
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
      Height          =   280
      Index           =   6
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   2280
      Width           =   2000
   End
   Begin VB.TextBox TXTCAJA 
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
      Height          =   280
      Index           =   5
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   1980
      Width           =   2000
   End
   Begin VB.TextBox TXTLABEL 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   4
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   38
      TabStop         =   0   'False
      Text            =   "PLIEGO 3 "
      Top             =   1680
      Width           =   1800
   End
   Begin VB.TextBox TXTCAJA 
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
      Height          =   280
      Index           =   4
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   1680
      Width           =   2000
   End
   Begin VB.TextBox TXTLABEL 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   3
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   36
      TabStop         =   0   'False
      Text            =   "PLIEGO 2 "
      Top             =   1380
      Width           =   1800
   End
   Begin VB.TextBox TXTCAJA 
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
      Height          =   280
      Index           =   3
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   1380
      Width           =   2000
   End
   Begin VB.TextBox TXTLABEL 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   1
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Text            =   "MEDIDAS (Cm) "
      Top             =   780
      Width           =   1800
   End
   Begin VB.TextBox TXTLABEL 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   2
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Text            =   "PLIEGO Exterior "
      Top             =   1080
      Width           =   1800
   End
   Begin VB.TextBox TXTCAJA 
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
      Height          =   280
      Index           =   2
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1080
      Width           =   2000
   End
   Begin VB.TextBox TXTCAJA 
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
      Height          =   280
      Index           =   1
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   780
      Width           =   2000
   End
   Begin VB.TextBox TXTCAJA 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   0
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Text            =   "CÓDIGO / DATO"
      Top             =   480
      Width           =   2000
   End
   Begin VB.TextBox TXTLABEL 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   0
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Text            =   "ITEM"
      Top             =   480
      Width           =   1800
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   840
      Top             =   5280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2625
      OleObjectBlob   =   "FRMENVAP.frx":0004
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   3240
      OleObjectBlob   =   "FRMENVAP.frx":0238
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4155
      Left            =   90
      TabIndex        =   184
      Top             =   20
      Visible         =   0   'False
      Width           =   7450
   End
   Begin VB.Label Label1 
      Caption         =   "X"
      Height          =   255
      Left            =   3000
      TabIndex        =   165
      Top             =   840
      Width           =   375
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      Height          =   300
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Menu mnuarchivo 
      Caption         =   "Archivo"
      Begin VB.Menu mnuAltaxCopia 
         Caption         =   "Alta x Copia"
      End
      Begin VB.Menu sep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGuardar 
         Caption         =   "Guardar"
      End
      Begin VB.Menu s1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuimprimir 
         Caption         =   "Imprimir"
      End
      Begin VB.Menu s2 
         Caption         =   "-"
      End
      Begin VB.Menu mnusalir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu mnuedicion 
      Caption         =   "Edición"
      Begin VB.Menu mnuActUso 
         Caption         =   "Actualizar Uso en Formula (Solo Materiales Sin Marca de Consumo x Corte)"
      End
   End
   Begin VB.Menu mnuVer 
      Caption         =   "Ver"
      Begin VB.Menu menverestruc 
         Caption         =   "Ver Estrctura Codificada"
      End
   End
   Begin VB.Menu mnuConfig 
      Caption         =   "Configuración"
      Begin VB.Menu mnuNivelconfig 
         Caption         =   "Configurar Niveles de Control (1 Prepara.- 2 Revisa - 3  Aprueba)"
      End
      Begin VB.Menu mnuCarpetaImagenes 
         Caption         =   "Configurar Carpeta de Archivos de Imagen"
      End
   End
   Begin VB.Menu mnuAD 
      Caption         =   "Accesos Directos"
      Begin VB.Menu mnuAltaMP 
         Caption         =   "Alta de Materias Primas"
      End
   End
End
Attribute VB_Name = "FRMENVAP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias _
    "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Dim NOCERRARFRM%



Sub ACTUALIZA_TXTUSO()
   corte = XVALO(TXTCORTE)
   For L% = 1 To Me.TXTUSO.UBound
      
     CI2% = CODINT%(TXTCAJA(L%))
     If CI2% > 0 Then
       CODXX$ = CODEXT$(CI2%)
       DEXX$ = DESCRIT0$(CI2%)
       UNXX$ = Unimed$(CI12%)
       USOX$ = FORMATNUM$(XVALO(Me.TXTUSO(L%)), "F12.4")
       PDX$ = "0"
       VALASIG = XVALO(TALLE$(CI2%))
       If VALASIG <= 0.0005 Then VALASIG = 1
       If ConsuSegunCorte%(CI2%) > 0 Then
          'SI TIENE MARCA CALCULA SEGUN CORTE.
          USOX$ = FORMATNUM$((corte * VALASIG), "F12.4")
       Else
          USOX$ = VALASIG
       End If
       TXTUSO(L%).TEXT = USOX$
     End If
   Next L%
End Sub

Sub AgregarCodigoReplacoc(CI1%, NC, CODCLIENTE$)
   Call TRALOCAL("REPLACOC", "")
   
   RFF$ = RECFILT$("!REPLACOC", 1, MKI$(NC))
   UREP& = ULTREG&("!REPLACOC")
   For KKI% = 1 To Len(RFF$) Step 4
     REREP& = CVS(Mid$(RFF$, KKI%, 4))
     If REREP& > 0 Then
       If REREP& <= UREP& Then
         XXYY$ = REGLEIDO$("!REPLACOC", REREP&)
         If CVI(Mid$(XXYY$, 3, 2)) = CI1% Then
           Call REPLA(XXYY$, CODCLIENTE$, 5, 28)
             Call GRAREG("!REPLACOC", XXYY$, REREP&)
             GoTo 99
         End If
       End If
     End If
   Next KKI%
   
   'SI NO ENTRO A GRAREG LLEGA ACA LO DA DE ALTA
   
   Call REPLA(RECORD$, MKI$(NC), 1, 2)
   Call REPLA(RECORD$, MKI$(CI1%), 3, 2)
   Call REPLA(RECORD$, CODCLIENTE$, 5, 28)
   Call REGAPP("!REPLACOC", RECORD$)
   Call SETULTREG("!REPLACOC", ULTREG("!REPLACOC"))
99 Call TRACENTRAL("REPLACOC", "")

'    FIN& = ULTREG("!REPLACOC")
'    For REG& = 1 To FIN&
'      x$ = REGLEIDO$("!REPLACOC", REG&)
'      A = CVI(Mid$(x$, 1, 2))
'      B = CVI(Mid$(x$, 3, 2))
'      c = Mid$(x$, 5, 28)
'    Next
End Sub

Sub ControlCodigoCliente()
    'NECESITA PASAR EL CODIGO DEL CLIENTE QUE ESTA EN EL CAMPO COMENTARIO EN EL ID 32 DE CADA CODIGO
    'RECORRO AGRUPANDO POR CODIGO
    'HAGO UNA CONASULTA SEGUN CODIGO ORDENADA DESCENDENTE QUE VA A DEVOLVER LA POSICION 32 POR QUE ES EL ULTIMO QUE GRABA
    J& = 0
    If CantRegistros("ESPECIFIC", "CODCLIENTE <> ''", "NOMESS") > 0 Then Exit Sub
    SQLX$ = "SELECT COD_INTE FROM ESPECIFIC WITH(NOLOCK) WHERE COD_INTE > 0 GROUP BY COD_INTE"
    Set RS = READSET(SQLX$)
    With RS
       Do While Not .EOF
         CI1% = XVALO(!cod_inte)
         SQLX$ = "SELECT TOP 1 COMENTARIO FROM ESPECIFIC WITH(NOLOCK) WHERE COD_INTE = " & CI1%
         SQLX$ = SQLX$ + " ORDER BY ID DESC"
         Set rs1 = READSET(SQLX$)
         With rs1
           If Not .EOF Then
              CODCLIENTE$ = TRIM$(SAFETEXT$(!COMENTARIO))
              If CODCLIENTE$ <> "" Then
                 J& = J& + 1
                 Call ACTUREGISTRO("ESPECIFIC", "CODCLIENTE", "COD_INTE = " & CI1%, CODCLIENTE$, REGAF&, "NOMESS")
              End If
           End If
         End With
         .MoveNext
       Loop
    End With
    Call COMUNI("Se han Actualizado: " & J& & " Registros")
End Sub

Sub COPIAR_ESTRUCTURA(CI1%, CI1NUEVO%)
        Dim ITEMX(), CODIDATOX(), INDIOBJX(), UsoBrutoX(), CORTEX(), OBSERVACIONX(), COMENTARIOX$()
        ReDim Preserve ITEMX(1), CODIDATOX(1), INDIOBJX(1), UsoBrutoX(1), CORTEX(1), OBSERVACIONX(1), COMENTARIOX$(1)
        IND& = 0
        SQLX$ = "SELECT * FROM ESPECIFIC WITH(NOLOCK) where COD_INTE= " & CI1%
        SQLX$ = SQLX$ + " ORDER BY INDIOBJ ASC"
        Set RS = READSET(SQLX$)
        HAYESTRUCTURA% = 0
        With RS
          Do While Not .EOF
             IND& = IND& + 1
             ReDim Preserve ITEMX(IND&), CODIDATOX(IND&), INDIOBJX(IND&), UsoBrutoX(IND&), CORTEX(IND&), OBSERVACIONX(IND&), COMENTARIOX$(IND&)
             ITEMX(IND&) = SAFETEXT$(!Item)
             CODIDATOX(IND&) = SAFETEXT$(!CODIDATO)
             INDIOBJX(IND&) = SAFETEXT$(!INDIOBJ)
             UsoBrutoX(IND&) = SAFETEXT$(!usobruto)
             CORTEX(IND&) = XVALO(!corte)
             OBSERVACIONX(IND&) = SAFETEXT$(!OBSERVACION)
             COMENTARIOX$(IND&) = SAFETEXT$(!COMENTARIO)
             HAYESTRUCTURA% = 1
            .MoveNext
          Loop
        End With
        RS.Close: Set RS = Nothing
        
        If HAYESTRUCTURA% < 1 Then Exit Sub
        
        COD$ = CODEXT$(CI1NUEVO%)
        
        Dim rs1 As ADODB.Recordset
        Set rs1 = New ADODB.Recordset
        SQLX$ = "Select * from ESPECIFIC  where COD_INTE=" & CI1NUEVO%
        rs1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
        With rs1
         For J& = 1 To UBound(ITEMX)
           
           .AddNew
           !ID = J&
           !CODIEMPR = CodiEmp%
           !cod_inte = CI1NUEVO%
           !Codigo = COD$
           !Descripcion = ""
           If J& = 1 Then !Descripcion = Left$(Me.TXTDESC(J&), 64)
           !Descripcion = Descr$
           !Item = ITEMX(J&)
           !CODIDATO = CODIDATOX(J&)
           !PREPARO = ""
           !REVISO = ""
           !APROBO = ""
           !FEPREPARO = CVDAT(0)
           !FEPREPARO = CVDAT(0)
           !FEPREPARO = CVDAT(0)
           !CODCLIENTE = ""
           !MARBORRA = 0
           !INDIOBJ = INDIOBJX(J&)
           !usobruto = UsoBrutoX(J&)
           !corte = CORTEX(J&)
           If J& = 30 Then
             !OBSERVACION = TXTOBS(J&).TEXT
            Else
             !OBSERVACION = ""
             !COMENTARIO = COMENTARIOX$(J&)
           End If

           .Update
          Next
          End With
          rs1.Close
          Set rs1 = Nothing
          
'           CACAM% = RMaster.Fields.Count
'
'           For KU% = 1 To CACAM%
'             DATO$ = SAFETEXT$(RMaster.Fields(KU% - 1).Value)
'             NOCAMPO$ = TRIM$(UCase$(RMaster.Fields(KU% - 1).Name))
'
'             If NOCAMPO$ <> "CODIGO" Then
'                If NOCAMPO$ <> "DESCRIPCION" Then
'                   If NOCAMPO$ <> "CODINT" Then
'                      If NOCAMPO$ <> "ALIAS" Then
'                         If NOCAMPO$ <> "UMEDIDA" Then
'                           CONDI$ = "CODINT=" & CI_NEW%
'                           Call ACTUREGISTRO("MASTER", NOCAMPO$, CONDI$, DATO$, REGAF&, "NOMESS")
'                         End If
'                      End If
'                   End If
'                End If
'              End If
'           Next KU%
'        End If
'       End If
'
'    SQLX$ = "INSERT INTO ESPECIFIC(ID"
'    SQLX$ = SQLX$ + ",1" & CodiEmp%
'    SQLX$ = SQLX$ + "," & CI1NUEVO%
'    SQLX$ = SQLX$ + ",'" & CODEXT$(CI1NUEVO%) & "'"
'    SQLX$ = SQLX$ + ",'" & DESCRIT0$(CI1NUEVO%) & "'"
'    SQLX$ = SQLX$ + ", Item, CODIDATO, '','','' "
'    SQLX$ = SQLX$ + ",'" & CVDAT(0) & "'"
'    SQLX$ = SQLX$ + " ,'" & CVDAT(0) & "'"
'    SQLX$ = SQLX$ + ",'" & CVDAT(0) & "'"
'    SQLX$ = SQLX$ + ", COMENTARIO, CODCLIENTE , OBSERVACION, MARBORRA, INDIOBJ, UsoBruto, CORTE)"
'    SQLX$ = SQLX$ + " SELECT ID, CodiEmpr, COD_INTE, Codigo, Descripcion, ITEM, CODIDATO, PREPARO, REVISO, APROBO, FEPREPARO, FEREVISO, FEAPROBO, COMENTARIO,"
'    SQLX$ = SQLX$ + "CODCLIENTE , OBSERVACION, MARBORRA, INDIOBJ, UsoBruto, CORTE FROM ESPECIFIC WHERE COD_INTE = " & CI1%
'    FLEXCONN.Execute SQLX$
End Sub

Sub GrabarCodigoClienteEnPlacoc()
    'ESTA SUB LE ESTA GURDANDO EN LA TABLA DAT REPLACOC EL CODIGO DEL CLIENTE.
    SQLX$ = "SELECT COD_INTE,CODCLIENTE,COMENTARIO FROM ESPECIFIC WITH(NOLOCK) WHERE COD_INTE > 0 AND ITEM = 'CLIENTE' ORDER BY COD_INTE"
    Set RS = READSET(SQLX$)
    With RS
       Do While Not .EOF
         CODICLIENTEX$ = TRIM$(SAFETEXT$(!CODCLIENTE))
         CI1% = XVALO(!cod_inte)
         NC = XVALO(!COMENTARIO)
         
         
         Call GuardarCodigoCliente(CI1%, NC, CODICLIENTEX$)
         .MoveNext
       Loop
    End With

End Sub
Sub LIMPIACASILLEROS()
  For i& = 1 To TXTCAJA.UBound
     TXTCAJA(i&) = ""
  Next i&
  
  For i& = 1 To TXTOBS.UBound
     TXTOBS(i&) = ""
  Next i&
  
  For i& = 1 To TXTDESC.UBound
     TXTDESC(i&) = ""
  Next i&
  For i& = 0 To TXTMED.UBound
     TXTMED(i&) = ""
  Next i&
  Me.TXTPREPARO = "": Me.TXTFECHREVISO = ""
  Me.TXTREVISO = "": Me.TXTFECHREVISO = ""
  Me.TXTAPROBO = "": Me.TXTFECHAPROBO = ""
  TXTFUELLE.TEXT = ""
  List1.Clear
End Sub

Private Sub Command4_Click(Index As Integer)
    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, "Código/a16;Descricpción/A48;Cons.-x-Corte /F16.0;Val.Asig/f16.4", FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "Código/a16;Descricpción/A48;Cons.-x-Corte /F16.0;Val.Asig/f16.4", FRMZELECHO)
    FRMZELECHO.Caption = "Listado de Materias Primas"

    Dim obj As New RECORXET
    Set rs1 = obj.RS_MasterMateriasPrimas()
    Call Carga_MSF_Recordset(rs1, "Código/a16;Descricpción/A48;Cons.-x-Corte /F16.0;Val.Asig/f16.4", FRMZELECHO.msf2, 1)
    Set obj = Nothing
    rs1.Close
    Set rs1 = Nothing
    '
    FRMZELECHO.Width = 12000
    FRMZELECHO.Show 1
    COD$ = RESP_ZELE_VECT(1)
    If COD$ <> "" Then
       TXTCAJA(Index) = COD$
       TXTDESC(Index) = DESCRIT0$(CODINT%(COD$))
    End If
    '
End Sub

Private Sub Command44_Click()

     Call SELECHO("MASTER")
     COD$ = TRIM$(VALACT1$("MASTER"))
     If COD$ <> "" Then
       TXTCODIGO = COD$
     End If

End Sub


Private Sub Command444_Click()
   X = LISTASELECGENERAL("TABLASI_NO", , , 1)
   TXTFUELLE.TEXT = TRIM$(RESP_ZELE_VECT(2))
End Sub

Private Sub Command47_Click(Index As Integer)
   Command47(Index).Enabled = False
    
    'BOTON ADD
    CARPETAX$ = Control("RUTAIMAN", "RUTAIMAG")
    If CARPETAX$ = "" Then
       Call COMUNI("Imposible Continuar\Falta Seleccionar Carpeta de Imagenes")
       GoTo 99
     End If
     
    On Error Resume Next
    X = GetAttr(CARPETAX$)
    If Err Then
        Call COMUNI("Imposible Continuar\Falta Seleccionar Carpeta de Imagenes o Ruta Inexistente")
        GoTo 99
    End If
    On Error GoTo 0
    Select Case Index
     Case 0: CARGARUTA
     'BOTON QUITAR
     Case 1
       If List2.ListIndex >= 0 Then
          List2.RemoveItem (List2.ListIndex)
       End If
   End Select
   '
   '
99 Command47(Index).Enabled = True

End Sub

Private Sub Form_Load()
   Call CREATABLA_ARTICULOS
   Call GENERATABLA_ESPECIFIC
   Call ControlCodigoCliente
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   
   If InStr(EPAC$, "CONTEPLAST") > 0 Then Call ControlIvaCliente

   
'   Call GrabarCodigoClienteEnPlacoc
''''   Call ProcesoAgregarCodigoIntaCpomp
End Sub
Sub ControlIvaCliente()
    'ESTA RUTINA LE AGREGA DE WALDBOT LA POSICION DE IVA A LA TABLA FLEXOFT.
    'TAMBIEN LE AGREGA EL CODIGO DE CLIENTE (TEXTO) QUE NO LO TENIA INGRESADO
    On Error Resume Next
    Call abreBaDaWaldClientes
    Dim env As New Envap
    If Err > 0 Then Call COMUNI("Imposible Conectar con Base de Datos"): GoTo 99
    On Error GoTo 0
   
   SQLX$ = "SELECT NUME_CLI,Piva_Cli FROM DEUDOR12 WITH(NOLOCK) ORDER BY NUME_CLI"
   Set RS0 = READSET(SQLX$)
   With RS0
    Do While Not .EOF
      NC& = XVALO(!nume_cli)
      PVCTE% = XVALO(!Piva_Cli)
      CONDI$ = "NUMERO = '" & FORMATOCON0$(NC&, 7) & "'"
      SQLX$ = "SELECT IMPUESTO FROM CLIENTES WHERE " & CONDI$
      Dim RS As New ADODB.Recordset
      Set RS = BASEWALDCLIENTE.Execute(SQLX$)
      PIVX$ = TRIM$(SAFETEXT$(RS!IMPUESTO))
      If PIVX$ <> "" Then
        PIVXFLEX% = env.PosIvaFlexVSWaldbott%(PIVX$)
        If PVCTE% <> PIVXFLEX% Then
          CONDI2$ = " NUME_CLI = " & NC&
          Call ACTUREGISTRO("DEUDOR12", "Piva_Cli", CONDI2$, PIVXFLEX%, REGAF&, "NOMESS")
          Call ACTUREGISTRO("DEUDOR12", "CODI_CLI", CONDI2$, "CLI-" & NC&, REGAF&, "NOMESS")
        End If
      End If
      .MoveNext
    Loop
   End With
   RS0.Close
   Set RS0 = Nothing
   RS.Close
   Set RS = Nothing
   BASEWALDCLIENTE.Close
   
   Set env = Nothing
99
End Sub
Sub ProcesoAgregarCodigoIntaCpomp()
   SQLX$ = "select CODIDATO,CODINT_COMP, * from ESPECIFIC "
   SQLX$ = SQLX$ + " WHERE CODIDATO <> '' AND (CODINT_COMP < 1 OR CODINT_COMP IS NULL) AND CODIDATO <> 'SI' AND CODIDATO <> 'NO' AND CODIDATO <> 'X'"
   Set RS = New ADODB.Recordset
29 On Error Resume Next
   RS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 29
   End If
   On Error GoTo 0
   With RS
     Do While Not .EOF
         CI1% = CODINT%(SAFETEXT$(!CODIDATO))
         If CI1% > 0 Then
            !CODINT_COMP = CI1%
            Debug.Print SAFETEXT$(!CODIDATO)
            .Update
         End If
         .MoveNext
     Loop
   End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If Cancel = 0 Then
     If NOCERRARFRM% = 1 Then
       Cancel = 1
     End If
   End If
End Sub

Private Sub List1_DblClick()
   List1.Visible = False
End Sub


Private Sub List2_Click()
   List2.ToolTipText = TRIM$(Mid$(List2.TEXT, 25))
End Sub

Private Sub menverestruc_Click()
      If List1.Visible = True Then List1.Visible = False: Exit Sub
      
      Ci% = XVALO(TXTCODINT)
      CODD$ = Me.TXTCODIGO
      If Ci% < 1 Then Exit Sub
      '
      Screen.MousePointer = 11
      PORMILLAR% = 0
      Call LEEEXPLO(CODD$, 1, 1)
      Campos$ = "Código/A18;Descripición/A48;U.Stocks/A12;Uso Unitario/F14.4"

      For i% = 1 To CAIT%
         XXX$ = EXPLOLIN$(i%)
         CODI$ = Mid$(XXX$, 1, 16)
         Descr$ = Mid$(XXX$, 17, 28)
         un$ = Mid$(XXX$, 47, 2)
         Uso = Mid$(XXX$, 49, 12)
TEXTOGRILLA$ = CODI$ & "|" & Descr$ & "|" & un$ & "|" & Uso
'Call CARGA_GRILLA_SHOWREP(SHOWREP07, SHOWREP07.GRID, Campos$, TEXTOGRILLA$, INICIALIZO%)
         
      Next i%
      '
'MUESTRA FRM DE GENREP07
SHOWREP07.Label7 = SHOWREP07.GRID.Rows - 1
SHOWREP07.Frame1.Visible = False
SHOWREP07.Label3 = "Estructura de Materiales Códificados de Código: " & CODD$
Screen.MousePointer = 1
Call COLOREAR_GRILLASQL(SHOWREP07.GRID, &HC0C0C0, 1)

SHOWREP07.Show 1
End Sub

Private Sub mnuActUso_Click()
  '
  CIXI% = XVALO(Me.TXTCODINT)
  
  If CIXI% <= 0 Then
      Exit Sub
  End If
  '
  If XVALO(Me.TXTCORTE) <= 0 Then
      Call COMUNI("Falta Ingresar Valor de Corte")
      Exit Sub
  End If


  LOTE# = LOTESTAN(CIXI%)
  If LOTE# < 1 Then
     LOTE# = 1
   End If
  TITU$ = "Formula de Código:  " & Me.TXTCODIGO & " - " & Unimed(Ci%) & "  -  " & DESCRIT0(Ci%)
  Call SETULTREG("!MATEROF", 0)
   REG& = 0: J& = 1
   Do While J& <= TXTCAJA.UBound
     CODIDAT$ = TXTCAJA(J&)
     CI1% = CODINT%(CODIDAT$)
     IDOBJ& = XVALO(J&)
     If CI1% > 0 And ConsuSegunCorte%(CI1%) < 1 Then 'MUESTRA SOLO LOS COMPONENTES QUE NO ESTEN MARCADOS POR CONSUMO X CORTE
        AAXX$ = REGBLAN("MATEROF")
        CODIELE% = CI1%
        Unidad$ = Unimed$(CODIELE%)
        Uso# = XVALO(Me.TXTUSO(J&))
        Operacion% = 0
        '
        REG& = REG& + 1
        Call REPLA(AAXX$, MKI(CODIELE%), 1, 2)
        Call REPLA(AAXX$, Unidad$, 3, 2)
        Call REPLA(AAXX$, MKD(Uso#), 9, 8)
        Call REPLA(AAXX$, MKS(IDOBJ&), 25, 4)
        Call REPLA(AAXX$, MKI(Operacion%), 29, 2)
        '
        Call GRAREG("!MATEROF", AAXX$, REG&)
     End If
     J& = J& + 1

   Loop
     '
     ATRI$ = "/TITUPAN = " & TITU$
     ATRI$ = ATRI$ + "/POSEDIT(1)=2"
     ATRI$ = ATRI$ + "/POSEDIT(3)=2"
'     ATRI$ = ATRI$ + "/POSEDIT(4)=2"
     ATRI$ = ATRI$ + "/POSEDIT(5)=2"
     ATRI$ = ATRI$ + "/POSEDIT(6)=2"
     ATRI$ = ATRI$ + "/POINDEL=NO" ' DESHABILITA EL INSERTAR Y EL DEL
     
     
10    XXX% = VENTABU%("FORPATRO" + ATRI$)

     If XXX% < 0 Then
        Exit Sub
     End If
     '
     
     REGACTI& = 1
     REGFINAL% = ULTREG("!Materof")
     Do While REGACTI& <= REGFINAL%
          AAXX$ = REGLEIDO("!MATEROF", REGACTI&)
          CODIELE% = CVI(Mid$(AAXX$, 1, 2))
          Unidad$ = Mid$(AAXX$, 3, 2)
          Uso# = CVD(Mid$(AAXX$, 9, 8))
          IDOBJ& = CVS(Mid$(AAXX$, 25, 4))
          OPERA% = CVI(Mid$(AAXX$, 29, 2))
          REGACTI& = REGACTI& + 1
          '
          TXTUSO(IDOBJ&) = FORMATNUM$(Uso#, "F12.4")
      Loop
      '
      Call TXTAPROBO_DblClick
      
     
99

End Sub

Private Sub mnuAltaMP_Click()
'    AMASTOK07.Command21_Click
    Call CONECRUN("AMASTO07/ALTACODIGO", "ALTA DE NUEVO CODIGO")
    
End Sub

Private Sub mnuAltaxCopia_Click()
   CI1% = XVALO(TXTCODINT)
   If CI1% < 1 Then
      Call COMUNI("Debe Seleccionar Un Código")
      Exit Sub
   End If
   '
   OPX% = COMALTER%("Opciones de Trabajo\\Nueva Versión de Código Activo\Nueva Bolsa")
   If OPX% < 1 Then Exit Sub

   If OPX% = 1 Then
     Call DARDEALTA(0, CI1%, 1)
   End If
   
   If OPX% = 2 Then Call DARDEALTA(0, CI1%)
   
End Sub

Sub DARDEALTA(Optional NCLIE, Optional CII%, Optional NUEVAVERSION%)
    NUEVOCODI.TXTCLIENTE = XVALO(NCLIE)
    Erase CADENATEX$
    If NUEVAVERSION% > 0 Then
       NUEVAVER$ = GENERAR_NUEVA_VERSION$(CII%)
       If NUEVAVER$ = "" Then NUEVAVERSION% = 0 'NO PUDO GENERAR LA NUEVA VERSION
    End If
    Call TEXTOLOTES(NUEVAVER$, "-")
    On Error Resume Next
    For i& = 1 To UBound(CADENATEX$)
         Select Case i&
           Case 1: NUEVOCODI.TXTTIPOBOLSA = CADENATEX$(i&)
           Case 2: NUEVOCODI.TXTCLIENTE = CADENATEX$(i&)
           Case 3: NUEVOCODI.TXTORDEN = CADENATEX$(i&)
           Case 4: NUEVOCODI.TXTVERSION = CADENATEX$(i&)
         End Select
    Next i&
    On Error GoTo 0
20  NUEVOCODI.Show 1
    ESPORCOPIA% = 0
    If NUEVOCODI.TXTCODIGO = "" Then Exit Sub
    
       VDEF$ = String$(128, 0)
       '
       'PASA LOS DATOS DEL
       LONCODI% = XVALO(Control$("MASTER", "LOMAXCOD"))
       COD$ = TRIM$(NUEVOCODI.TXTCODIGO)
       EXISTECOD% = CODINT%(COD$)
       If EXISTECOD% > 0 Then
          Call COMUNI("Imposible dar de Alta - Código Existente")
          GoTo 20
       End If
       CI_NEW% = CODINT_NEW
       Call REPLA(VDEF$, MKI$(CI_NEW%), 1, 2)
       If CII% > 0 Then 'ES POR COPIA
        ESPORCOPIA% = 1
        Call REPLA(VDEF$, MKI$(GRUCOVEI%(CII%)), 3, 2)
        Call REPLA(VDEF$, DESCRIT0$(CII%), 5, 64)
        Call REPLA(VDEF$, COD$, 77, LONCODI%)
        Call REPLA(VDEF$, Unimed$(CII%), 109, 2)
        Call REPLA(VDEF$, Chr$(5), 101, 1)
       Else
        Call REPLA(VDEF$, COD$, 77, LONCODI%)
        Call REPLA(VDEF$, "U.", 109, 2)
        Call REPLA(VDEF$, Chr$(1), 101, 1)
       End If
       '
       OKX% = 0
       
       Call ALTACODIGO(VDEF$, OKX%)
       
       If LONCODI% < 15 Or LONCODI% > 24 Then LONCODI% = 15
       If OKX% = 1 Then
          CODINUE$ = Mid$(VDEF$, 77, LONCODI%)
       End If
       '
       Screen.MousePointer = 11
       Call BLOQUEAR_CONTROLES(Me, 2, 1, 1, 1)
       NOCERRARFRM% = 1 ' NO PERMITE CERRAR EL FORMULARIO
       If OKX% = 1 And ESPORCOPIA% > 0 Then
        Dim RMaster As ADODB.Recordset
        Set RMaster = New ADODB.Recordset
        strSQL$ = "Select * from Master WITH(NOLOCK) where codint=" & CII%
        Set RMaster = READSET(strSQL$)
        If Not (RMaster.EOF And RMaster.BOF) Then
           CACAM% = RMaster.Fields.Count
           For KU% = 1 To CACAM%
             Dato$ = SAFETEXT$(RMaster.Fields(KU% - 1).Value)
             NOCAMPO$ = TRIM$(UCase$(RMaster.Fields(KU% - 1).Name))
             If NOCAMPO$ <> "CODIGO" Then
                If NOCAMPO$ <> "DESCRIPCION" Then
                   If NOCAMPO$ <> "CODINT" Then
                      If NOCAMPO$ <> "ALIAS" Then
                         If NOCAMPO$ <> "UMEDIDA" Then
                           CONDI$ = "CODINT=" & CI_NEW%
                           Call ACTUREGISTRO("MASTER", NOCAMPO$, CONDI$, Dato$, REGAF&, "NOMESS")
                         End If
                      End If
                   End If
                End If
              End If
           Next KU%
        End If
       End If
       'ACTUALIZO EL MASTER.
       CONDI$ = " CODIGO <> '' AND STATUS > 0 ORDER BY CODIGO ASC"
       If Control$("", "IDSECUN") = "ALIAS" Then
           Call CARGALISEL("ECOMAST", "MASTER", "CODIGO/A24;DESCRIPCION/A36;NPLANO/A24", CONDI$, "NOMESS")
       Else
           Call CARGALISEL("ECOMAST", "MASTER", "CODIGO/A24;DESCRIPCION/A50", CONDI$, "NOMESS")
       End If
        '
        Call FRMENVAP.COPIAR_ESTRUCTURA(CII%, CI_NEW%)

        On Error Resume Next
        Unload NUEVOCODI
        On Error GoTo 0
        
        NCLIEN = XVALO(Mid$(COD$, 4, 5))
        FRMENVAP.TXTOBS(31) = NCLIEN
        FRMENVAP.TXTCODIGO = COD$
        FRMENVAP.TXTRAZOCLI = rasocli$(NCLIEN)
        FRMENVAP.TXTPREPARO = USERNAME$
        FRMENVAP.TXTFECHPREPARO = FECHATEX$(HOY(HO$))
        Screen.MousePointer = 1
        Call BLOQUEAR_CONTROLES(Me, 1, 1, 1, 1)
        NOCERRARFRM% = 0
End Sub







Private Sub mnuCarpetaImagenes_Click()
    RUTAIMAG.Show 1
End Sub

Private Sub mnuGuardar_Click()
   If XVALO(TXTCODINT) < 1 Then Exit Sub
   
   If TRIM$(TXTREVISO) <> "" Then
      Call COMUNI("Imposible Grabar  - Estructura Revisada")
'      Call ALERTA(TXTREVISO, "Imposible Grabar  - Estructura Revisada")
'      Call PINTATEXT(TXTREVISO)
      Exit Sub
   End If
   If TRIM$(TXTAPROBO) <> "" Then
      Call COMUNI("Imposible Grabar  - Estructura Aprobada")
      Exit Sub
   End If
    '
   If TRIM$(Me.TXTAPROBO) <> "" Then
      Call COMUNI("Imposible Guardar\Especificación Técnica Aprobada\Deberá Generar una Nueva Versión")
     Exit Sub
   End If
   '
   If XVALO(Me.TXTCORTE) < 0 Then
      Call COMUNI("Falta Ingresar Valor de Corte")
      Exit Sub
   End If

   Call GENERATABLA_ESPECIFIC
   '
   Call GUARDARFRMENVAP
   
End Sub
Sub GUARDARFRMENVAP(Optional NOPRESENTARMENSAJE%)
   CONDI$ = "COD_INTE = " & XVALO(Me.TXTCODINT)
   Call ACTUREGISTRO("ESPECIFIC", "MARBORRA", CONDI$, 1, REGAF&, "NOMESS")

   MAXID = XVALO(VALOBADA("ESPECIFIC", "ID", "ID > 0 ORDER BY ID DESC", "NOMESS"))
   
   SQLX$ = "SELECT TOP 1 * FROM ESPECIFIC WHERE COD_INTE = " & XVALO(Me.TXTCODINT) & "   ORDER BY ID DESC"
   Set RS = New ADODB.Recordset
29 On Error Resume Next
   RS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 29
   End If
   On Error GoTo 0
   With RS
    IDX = MAXID
    For i& = 1 To TXTLABEL.UBound
        .AddNew   ' YA CONVERTIDO
        On Error Resume Next
        !CODIEMPR = CodiEmp%
        On Error GoTo 0
        ITEMX$ = TRIM$(UCase$(TXTLABEL(i&).TEXT))
        CODIDATO$ = "": USOBR = 0
        If i& <= TXTCAJA.UBound Then CODIDATO$ = TRIM$(UCase$(TXTCAJA(i&).TEXT)) 'P/Q HAY MENOS CAJAS QUE EL TOTAL
        If i& <= Me.TXTUSO.UBound Then USOBR = TRIM$(UCase$(TXTUSO(i&).TEXT))  'P/Q HAY MENOS CAJAS QUE EL TOTAL
        
        COMENTARIOX$ = TRIM$(UCase$(TXTOBS(i&).TEXT))
        !ID = IDX + i&
        !cod_inte = XVALO(TXTCODINT)
        !Codigo = Me.TXTCODIGO
        !Item = ITEMX$
        !CODIDATO = CODIDATO$
        !CODINT_COMP = 0
        'GUARDO EL CODIGO INTERNO DEL CODIGO COMPONENTE
        If CODIDATO$ <> "" Then
          If CODIDATO$ <> "X" And CODIDATO$ <> "SI" And CODIDATO$ <> "NO" Then
             CI_COMP% = CODINT%(CODIDATO$)
             If CI_COMP% > 0 Then
               !CODINT_COMP = CI_COMP%
             End If
          End If
        End If
        !usobruto = XVALO(USOBR)
       
        !corte = ROUND(XVALO(Me.TXTCORTE), 4)
        !FUELLE = TRIM$(TXTFUELLE)
        !Descripcion = ""
        If i& = 1 Then !Descripcion = Left$(Me.TXTDESC(i&), 64)
        If TRIM$(Me.TXTPREPARO) = "" Then TXTPREPARO = USERNAME$
        !PREPARO = Me.TXTPREPARO
        !REVISO = Me.TXTREVISO
        !APROBO = Me.TXTAPROBO
        FEPREPA% = CVINT(Me.TXTFECHPREPARO): If FEPREPA% < 1 Then FEPREPA% = HOY(HO$): Me.TXTFECHPREPARO = FECHATEX$(FEPREPA%)
        !FEPREPARO = CVDAT(CVINT(Me.TXTFECHPREPARO))
        If TRIM$(Me.TXTREVISO) <> "" Then !FEREVISO = CVDAT(CVINT(Me.TXTFECHREVISO))
        If TRIM$(Me.TXTAPROBO) <> "" Then !FEAPROBO = CVDAT(CVINT(Me.TXTFECHAPROBO))
        !MARBORRA = 0
        !INDIOBJ = i&
        If i& = 30 Then
           !OBSERVACION = TXTOBS(i&).TEXT
        Else
           !OBSERVACION = ""
           !COMENTARIO = COMENTARIOX$
        End If
        
        .Update
    Next i&

  End With
  RS.Close
  Set RS = Nothing
  Call ACTUREGISTRO("ESPECIFIC", "CODCLIENTE", "COD_INTE = " & XVALO(Me.TXTCODINT), TXTOBS(32), REGAF&, "NOMESS")
  
  'BORRO LO ANTERIOR
  CONDI$ = "COD_INTE = " & XVALO(Me.TXTCODINT) & " AND MARBORRA = 1"
  Call BORRAREGISTROS("ESPECIFIC", CONDI$, REGAF&, "NOMESS")
  '
     'GRAABAR P.O.E.`S
   '****************
   'MARCA DE BORRADO
   CI1% = XVALO(TXTCODINT)
   Call CREATABLA_ARTICULOS
   Call ACTUREGISTRO("DOCUREL", "MARBORRA", "COD_INTE = " & CI1%, 1, REG&)
   '
   'RECORRO EL LIST PARA GRABAR
   For i& = 1 To List2.ListCount
      Z$ = List2.List(i& - 1)
      DESCRI$ = TRIM$(Mid$(Z, 1, 24))
      RUTAX$ = TRIM$(Mid$(Z, 32, 512))
      
      SQLX$ = "INSERT INTO DOCUREL  "
      SQLX$ = SQLX$ + "(COD_INTE ,DESCRIPCION,RUTA,MARBORRA)"
      SQLX$ = SQLX$ + " VALUES(" & CI1%
      SQLX$ = SQLX$ + ",'" & Left$(DESCRI$, 24) & "'"
      SQLX$ = SQLX$ + ",'" & Left$(RUTAX$, 524) & "'"
      SQLX$ = SQLX$ + ",0"
      SQLX$ = SQLX$ + ")"
      
      FLEXCONN.Execute (SQLX$)
   Next i&
   '
   'BORRO LOS REGISTROS MARCADOS CON 1
   Call BORRAREGISTROS("DOCUREL", "MARBORRA = 1 AND COD_INTE = " & CI1%, REGAF&, "NOMESS")
   Call AgregarCodigoReplacoc(CI1%, XVALO(TXTOBS(31)), TXTOBS(32))
   Screen.MousePointer = 1
   If NOPRESENTARMENSAJE% < 1 Then
     Call COMUNI("Detalle Actualizado")
   End If
End Sub

Sub GUARDARFRMENVAP_MODIFICADO(Optional NOPRESENTARMENSAJE%)
   Call CREATABLA_ARTICULOS
   
   CONDI$ = "COD_INTE = " & XVALO(Me.TXTCODINT)
   SQLX$ = "DELETE FROM ESPECIFIC WHERE " & CONDI$
   FLEXCONN.Execute (SQLX$)
'   Call ACTUREGISTRO("ESPECIFIC", "MARBORRA", CONDI$, 1, REGAF&, "NOMESS")

'   MAXID = XVALO(VALOBADA("ESPECIFIC", "ID", "ID > 0 ORDER BY ID DESC", "NOMESS"))
   
   SQLX$ = "SELECT TOP 1 * FROM ESPECIFIC WHERE COD_INTE = " & XVALO(Me.TXTCODINT) & "   ORDER BY ESPECIFIC_ID DESC"
   Set RS = New ADODB.Recordset
29 On Error Resume Next
   RS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 29
   End If
   On Error GoTo 0
   With RS
    IDX = XVALO(!ESPECIFIC_ID)
    For i& = 1 To TXTLABEL.UBound
        .AddNew   ' YA CONVERTIDO
        On Error Resume Next
        !CODIEMPR = CodiEmp%
        On Error GoTo 0
        ITEMX$ = TRIM$(UCase$(TXTLABEL(i&).TEXT))
        CODIDATO$ = "": USOBR = 0
        If i& <= TXTCAJA.UBound Then CODIDATO$ = TRIM$(UCase$(TXTCAJA(i&).TEXT)) 'P/Q HAY MENOS CAJAS QUE EL TOTAL
        If i& <= Me.TXTUSO.UBound Then USOBR = TRIM$(UCase$(TXTUSO(i&).TEXT))  'P/Q HAY MENOS CAJAS QUE EL TOTAL
        
        COMENTARIOX$ = TRIM$(UCase$(TXTOBS(i&).TEXT))
        !ID = IDX + i&
        !cod_inte = XVALO(TXTCODINT)
        !Codigo = Me.TXTCODIGO
        !Item = ITEMX$
        !CODIDATO = CODIDATO$
        !CODINT_COMP = 0
        'GUARDO EL CODIGO INTERNO DEL CODIGO COMPONENTE
        If CODIDATO$ <> "" Then
          If CODIDATO$ <> "X" And CODIDATO$ <> "SI" And CODIDATO$ <> "NO" Then
             CI_COMP% = CODINT%(CODIDATO$)
             If CI_COMP% > 0 Then
               !CODINT_COMP = CI_COMP%
             End If
          End If
        End If
        !usobruto = XVALO(USOBR)
       
        !corte = ROUND(XVALO(Me.TXTCORTE), 4)
        !Descripcion = ""
        If i& = 1 Then !Descripcion = Left$(Me.TXTDESC(i&), 64)
        If TRIM$(Me.TXTPREPARO) = "" Then TXTPREPARO = USERNAME$
        !PREPARO = Me.TXTPREPARO
        !REVISO = Me.TXTREVISO
        !APROBO = Me.TXTAPROBO
        FEPREPA% = CVINT(Me.TXTFECHPREPARO): If FEPREPA% < 1 Then FEPREPA% = HOY(HO$): Me.TXTFECHPREPARO = FECHATEX$(FEPREPA%)
        !FEPREPARO = CVDAT(CVINT(Me.TXTFECHPREPARO))
        If TRIM$(Me.TXTREVISO) <> "" Then !FEREVISO = CVDAT(CVINT(Me.TXTFECHREVISO))
        If TRIM$(Me.TXTAPROBO) <> "" Then !FEAPROBO = CVDAT(CVINT(Me.TXTFECHAPROBO))
        !MARBORRA = 0
        !INDIOBJ = i&
        If i& = 30 Then
           !OBSERVACION = TXTOBS(i&).TEXT
        Else
           !OBSERVACION = ""
           !COMENTARIO = COMENTARIOX$
        End If
        .Update
    Next i&

  End With
  RS.Close
  Set RS = Nothing
  
'  'BORRO LO ANTERIOR
'  CONDI$ = "COD_INTE = " & XVALO(Me.TXTCODINT) & " AND MARBORRA = 1"
'  Call BORRAREGISTROS("ESPECIFIC", CONDI$, REGAF&, "NOMESS")
  '
     'GRAABAR P.O.E.`S
   '****************
   'MARCA DE BORRADO
   CI1% = XVALO(TXTCODINT)
'   Call ACTUREGISTRO("DOCUREL", "MARBORRA", "COD_INTE = " & CI1%, 1, REG&)
   '
   CONDI$ = "COD_INTE = " & CI1%
   SQLX$ = "DELETE FROM DOCUREL WHERE " & CONDI$
   FLEXCONN.Execute (SQLX$)

   'RECORRO EL LIST PARA GRABAR
   For i& = 1 To List2.ListCount
      Z$ = List2.List(i& - 1)
      DESCRI$ = TRIM$(Mid$(Z, 1, 24))
      RUTAX$ = TRIM$(Mid$(Z, 32, 512))
      
      SQLX$ = "INSERT INTO DOCUREL  "
      SQLX$ = SQLX$ + "(COD_INTE ,DESCRIPCION,RUTA,MARBORRA)"
      SQLX$ = SQLX$ + " VALUES(" & CI1%
      SQLX$ = SQLX$ + ",'" & Left$(DESCRI$, 24) & "'"
      SQLX$ = SQLX$ + ",'" & Left$(RUTAX$, 524) & "'"
      SQLX$ = SQLX$ + ",0"
      SQLX$ = SQLX$ + ")"
      
      FLEXCONN.Execute (SQLX$)
   Next i&
   '
'   'BORRO LOS REGISTROS MARCADOS CON 1
'   Call BORRAREGISTROS("DOCUREL", "MARBORRA = 1 AND COD_INTE = " & CI1%, REGAF&, "NOMESS")
   
   Screen.MousePointer = 1
   If NOPRESENTARMENSAJE% < 1 Then
     Call COMUNI("Detalle Actualizado")
   End If
End Sub


Private Sub mnuimprimir_Click()
      If XVALO(TXTCODINT) < 1 Then
        Call COMUNI("Debe Ingresar un Código Existente")
        Exit Sub
      End If
         
      SQLX$ = "SELECT CORTE,Descripcion"
      SQLX$ = SQLX$ + " FROM ESPECIFIC WITH(NOLOCK) "
      SQLX$ = SQLX$ + " WHERE COD_INTE = " & XVALO(TXTCODINT) & " ORDER BY ID"
      J& = 0
      Set RS = READSET(SQLX$)
      With RS
        Do While Not .EOF
            J& = J& + 1
    
            If J& = 1 Then
              textox$ = SAFETEXT$(!Descripcion)
              X = TEXTOLOTES$(textox$, "X")
              On Error Resume Next
              For k% = 0 To UBound(CADENATEX$)
                If k% < TXTMED.UBound Then
                  If k% > 0 Then
                    MEDI$ = MEDI$ & " x " & CADENATEX$(k% + 1)
                  Else
                    MEDI$ = CADENATEX$(k% + 1)
                  End If
                End If
              Next k%
              corte$ = TRIM$(FORMATNUM$(XVALO(!corte), "F10.2"))
              On Error GoTo 0
           End If
           .MoveNext
        Loop
      End With
      FILTX$ = TRIM$(TXTCODINT) & ";" & TRIM$(MEDI$) & ";" & corte$
      Call STDFORM("RCTA-ENV", FILTX$)
      
End Sub


Private Sub mnuNivelconfig_Click()
  Call Mostrar_Datusers
End Sub

Private Sub mnuSalir_Click()
   Unload Me
End Sub

Private Sub TXTALERTA_DblClick()
   TXTALERTA.Visible = False
End Sub


Private Sub TXTAPROBO_DblClick()
   CI1% = CODINT%(Me.TXTCODIGO)
   If CI1% < 1 Then
      Call COMUNI("Código Inexistente o Inválido")
      Exit Sub
   End If
   '
   If XVALO(Me.TXTCORTE) <= 0 Then
     Call COMUNI("Falta ingresar Valor de Corte")
     Exit Sub
   End If
   
   If VALIDA_ESTRUC_ENVAP% > 0 Then Exit Sub


   FIRMAELEC.Text1 = ""
   FIRMAELEC.Text2 = ""
   FIRMAELEC.Label1 = ""

   FIRMAELEC.Show 1

   If FIRMAELEC.ESTADO > 0 Then
      If NIVCONFI%(XVALO(FIRMAELEC.Label1)) < 3 Then
         Call COMUNI("Usuario Con Nivel de Clasificación Inválido Para Esta Operación")
         Exit Sub
      End If
      
      OPX% = COMALTER%("Atención Al Grabarse la Revisión\Se Actualizará Toda la Información Tal Cual Esta en Pantalla\\Mantener Información Original\Actualizar")
      If OPX% < 2 Then Exit Sub
      Screen.MousePointer = 11
      Me.TXTAPROBO = USERNAM$(XVALO(FIRMAELEC.Label1))
      Me.TXTFECHAPROBO = FECHATEX$(HOY(HO$))
      
      '
'      On Error GoTo ERROR_GUARDAR
'      FLEXCONN.BeginTrans
      Call GUARDARFRMENVAP(1) ' parametro 1 opcional para que no presente el mensaje de terminado que esta dentro de esta rutina
      Call ACTUALIZA_ESTRUCTURA(CI1%, XVALO(Me.TXTCORTE))
      Call GuardarCodigoCliente(CI1%, XVALO(TXTOBS(31)), TXTOBS(32))
      
      Call TXTCODIGO_Change
      
      Screen.MousePointer = 1
      Call COMUNI("Detalle y Estructura Actualizado")
   End If
   Unload FIRMAELEC
   Screen.MousePointer = 1
'   Exit Sub
'
'ERROR_GUARDAR:
'    If Err <> 0 Then
'       A = Err.Description
'       On Error Resume Next
'       FLEXCONN.RollbackTrans
'       On Error GoTo 0
'       Call COMUNI("ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION")
'       Screen.MousePointer = 1
'       Exit Sub
'    Else
'      Call COMUNI("Proceso Realizado Exitosamente")
'      FLEXCONN.CommitTrans
'    End If
 

End Sub
Function VALIDA_ESTRUC_ENVAP%()
    VALIDA_ESTRUC_ENVAP% = 0
    For L% = 1 To FRMENVAP.TXTCAJA.UBound
         CI2% = CODINT%(FRMENVAP.TXTCAJA(L%))
         If CI2% > 0 Then
            If XVALO(TALLE$(CI2%)) <= 0 Then
               Call COMUNI("Se Han Detectado Códigos Sin Valor de Uso")
               VALIDA_ESTRUC_ENVAP% = L%
               Call ALERTA(FRMENVAP.TXTCAJA(L%), "")
               Call PINTATEXT(FRMENVAP.TXTCAJA(L%))
               Exit For
            End If
         End If
    Next L%
    Exit Function
    

End Function
Sub ALERTA(Objeto As Object, TEXMEN$)
    OLEFT = Objeto.Left
    OANCHO = Objeto.Width
    DE_IZQ = OLEFT + OANCHO
    MAX_IZQ = DE_IZQ + TXTALERTA.Width
    ANCHOFORM = Me.Width
    If TEXMEN$ <> "" Then TXTALERTA = TEXMEN$
    If MAX_IZQ > ANCHOFORM Then
       TXTALERTA.Visible = True
       TXTALERTA.Left = OLEFT - (Objeto.Width + 300)
       TXTALERTA.Top = Objeto.Top
    Else
       TXTALERTA.Visible = True
       TXTALERTA.Left = DE_IZQ
       TXTALERTA.Top = Objeto.Top
    End If
    For i% = 1 To 3
    Call T_Espera(1)
        RECUERDO$ = TXTALERTA.TEXT
        TXTALERTA.TEXT = ""
        Call T_Espera(1 / 10)
        TXTALERTA.TEXT = RECUERDO$
        Call T_Espera(1)
    Next i%
End Sub
Sub ACTUALIZA_ESTRUCTURA(CI1%, corte)
   '
   If CI1% < 1 Or CI1% > NUMAS% Then Exit Sub
   '
   If DERACCE%("MODISTRU", "Modificación Fórmula de Conjunto") < 2 Then
      Exit Sub
   End If
   '
   '
   For L% = 1 To TXTCAJA.UBound
     CI2% = CODINT%(TXTCAJA(L%))
     If CI2% > 0 Then
       CODXX$ = CODEXT$(CI2%)
       DEXX$ = DESCRIT0$(CI2%)
       UNXX$ = Unimed$(CI12%)
       USOX$ = FORMATNUM$(XVALO(Me.TXTUSO(L%)), "F12.4")
       PDX$ = "0"
       VALASIG = XVALO(TALLE$(CI2%))
       If VALASIG <= 0.0005 Then VALASIG = 1
       If ConsuSegunCorte%(CI2%) > 0 Then
          'SI TIENE MARCA CALCULA SEGUN CORTE.
          USOX$ = FORMATNUM$((corte * VALASIG), "F12.4")
       Else
          USOX$ = VALASIG
       End If
'       If L% <= Me.TXTUSO.UBound Then TXTUSO(L%) = USOX$
       '
       XXX$ = Space$(128)
       Call REPLA(XXX$, CODXX$, 1, 16)
       Call REPLA(XXX$, DEXX$, 17, 28)
       Call REPLA(XXX$, UNXX$, 47, 2)
       Call REPLA(XXX$, USOX$, 49, 12)
       Call REPLA(XXX$, PDX$, 61, 4)
       Call REPLA(XXX$, USOX$, 89, 9)
       CAIT% = CAIT% + 1
       List1.AddItem XXX$
       List1.ListIndex = List1.ListCount - 1
     End If
   Next L%

   Ci% = CI1%
   '
   '
   Screen.MousePointer = 11
   CAIT% = List1.ListCount
   '
   Dim EXPLOLI1$(512)
   CAIT1% = CAIT%
   For JJI% = 1 To 512
      EXPLOLI1$(JJI%) = ""
      If JJI% <= List1.ListCount Then
         EXPLOLI1$(JJI%) = List1.List(JJI% - 1)
      End If
   Next JJI%
   '
   For i& = 1 To CAIT%
     EXPLOLIN(i&) = List1.List(i& - 1)
     CICOMPO% = CODINT%(Left$(EXPLOLIN(i&), 16))
     Call EXPLOMUL(CICOMPO%)
     For JJI% = 1 To CONTAMUL%
       If CIK%(JJI%) = Ci% Then
         Call MENSERR(24, "Imposible Registrar Esta Fórmula.\Error de Recursividad en Código " + TRIM$(CODEXT$(Ci%)) + "\" + DESCRIT0$(Ci%) + ".")
         GoTo 99
       End If
     Next JJI%
   Next i&
   '
   CAIT% = CAIT1%
   For JJI% = 0 To 512
      EXPLOLIN$(JJI%) = EXPLOLI1$(JJI%)
   Next JJI%
   '
   CODD$ = CODEXT$(CI1%)
   CABAS = 1
   CAIT% = List1.ListCount
   '
   If CABAS > 0.01 Then
      REVIFORMU = HOY(HOS$)
      Call EXPLOGR(CODD$, CABAS)
   End If
   '
99 Screen.MousePointer = 1

End Sub



 Sub TXTCAJA_Change(Index As Integer)
' SE PONE MUY LENTO
'     Me.TXTDESC(index) = DESCRIT0$(CODINT%(TXTCAJA(index)))
End Sub

Private Sub TXTCAJA_DblClick(Index As Integer)

   If TXTCAJA(Index) = "X" Then
      TXTCAJA(Index) = "SI"
   ElseIf TXTCAJA(Index) = "SI" Then
      TXTCAJA(Index) = "NO"
   ElseIf TXTCAJA(Index) = "NO" Then
      TXTCAJA(Index) = "X"
   Else
      TXTCAJA(Index) = "X"
   End If
End Sub

Private Sub TXTCAJA_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 2 Then
     If Index >= 26 And Index <= 28 Then
10       INGREXO& = XVALO(InputBox$("Solo Válido Ingreso de Números", "Ingrese Cantidad"))
         If INGREXO& < 0 Then Call COMUNI("No Válido el Ingreso Negativo"): GoTo 10
         TXTCAJA(Index) = TRIM(CStr(INGREXO&))
         TXTCAJA(Index).Alignment = 1
     End If
  End If
End Sub

Private Sub TXTCODIGO_Change()
   If TRIM$(TXTCODIGO) = "" Then Exit Sub
   Call GENERATABLA_ESPECIFIC
   Call LIMPIACASILLEROS
   TXTCODINT = CODINT(Me.TXTCODIGO)
   If XVALO(TXTCODINT) < 1 Then Exit Sub
   TXTLABELMAX = TXTLABEL.UBound
   Me.TXTDESCRICOD = DESCRIT0$(TXTCODINT)
   TXTCAJAMAX = TXTCAJA.UBound
   TXTDESCMAX = TXTDESC.UBound
   'Me.TXTUNIMED = UNIMED$(TXTCODINT)
   '***********************************************************
   '
   Dim VectArticulos As New articulos
   VectArticulos.cargarvector
   SQLX$ = "SELECT CODINT_COMP,Item,PREPARO,REVISO,APROBO,CORTE,FEPREPARO,FEREVISO,FEAPROBO,CODIDATO,COD_INTE,USOBRUTO,Observacion,COMENTARIO,Descripcion"
   SQLX$ = SQLX$ + " , FUELLE FROM ESPECIFIC WITH(NOLOCK) "
   SQLX$ = SQLX$ + " WHERE COD_INTE = " & XVALO(TXTCODINT) & " ORDER BY ID"
   J& = 0
   Set RS = READSET(SQLX$)
   With RS
     Do While Not .EOF
       it$ = TRIM$(SAFETEXT$(!Item))
       J& = J& + 1
       If J& = 1 Then
          Me.TXTPREPARO = SAFETEXT$(!PREPARO)
          
          Me.TXTREVISO = SAFETEXT$(!REVISO)
          Me.TXTAPROBO = SAFETEXT$(!APROBO)
          Me.TXTCORTE = FORMATNUM$(XVALO(!corte), "F12.4")
          Me.TXTFECHPREPARO = CVDAT(CVINT(!FEPREPARO)):
          Me.TXTFECHREVISO = CVDAT(CVINT(!FEREVISO)): If TXTFECHREVISO = "01/01/80" Then TXTFECHREVISO = ""
          Me.TXTFECHAPROBO = CVDAT(CVINT(!FEAPROBO)): If TXTFECHAPROBO = "01/01/80" Then TXTFECHAPROBO = ""
          TXTFUELLE.TEXT = SAFETEXT$(!FUELLE)
          
       End If
       
       If J& <= TXTLABELMAX Then
         AAA = TRIM$(UCase$(TXTLABEL(J&)))
         Me.Caption = AAA
         If it$ = TRIM$(UCase$(TXTLABEL(J&))) Then
'            AAA = SAFETEXT$(!CODIDATO) 'PROBANDO
            If J& <= TXTCAJAMAX Then TXTCAJA(J&) = SAFETEXT$(!CODIDATO)
            If J& <= TXTCAJAMAX Then TXTUSO(J&) = SAFETEXT$(!usobruto)
            AA = SAFETEXT$(!usobruto)
            CI1% = XVALO(!CODINT_COMP)
            If CI1% > 0 And J& <= TXTDESCMAX Then
                 TXTDESC(J&) = VectArticulos.Descripcion(CI1%)
            End If

'
            If J& = 30 Then
               TXTOBS(J&) = SAFETEXT$(!OBSERVACION) 'GUARDA EN EL CAMPO OBSERVACION POR QUE ES MEMO
            Else
               TXTOBS(J&) = SAFETEXT$(!COMENTARIO)

            End If
         End If
       End If
        '
       If J& = 1 Then
          textox$ = SAFETEXT$(!Descripcion)
          
          X = TEXTOLOTES$(textox$, "X")
          On Error Resume Next
          For k% = 0 To UBound(CADENATEX$)
            If k% <= TXTMED.UBound Then
              TXTMED(k%) = CADENATEX$(k% + 1)
            End If
          Next k%
          On Error GoTo 0
       End If
       .MoveNext
     Loop
   End With '
   'CARGAR ARCHIVO ADJUNTOS
   List2.Clear
   SQLX$ = "SELECT * FROM DOCUREL WITH(NOLOCK) WHERE COD_INTE = " & XVALO(TXTCODINT)
   Dim DOCTMP As ADODB.Recordset
   Set DOCTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
   With DOCTMP
       Do While Not .EOF
          DESCRI$ = SAFETEXT$(!Descripcion)
          RUTAX$ = SAFETEXT$(!Ruta$)
          Call REPLA(TX$, DESCRI$, 1, 24)
          Call REPLA(TX$, RUTAX$, 32, 512)
          List2.AddItem TX$
         .MoveNext
       Loop
   End With
   'VALIDO SI EL NUMERO DE CLIENTE INMERSO EN EL CODIGO ES VALIDO
   NCLIEN = XVALO(Mid$(Me.TXTCODIGO, 4, 5))
   If CantRegistros("DEUDOR12", "Nume_Cli = " & NCLIEN, "MOMESS") < 1 Then
      OPX1% = COMALTER%("Validación de Número de Cliente Incorporado en Código No Válido o Inexistente\\Cancelar\Continuar")
      If OPX1% < 2 Then
        Call LIMPIACASILLEROS
        Me.TXTCODIGO = ""
        Exit Sub
      End If
   Else
      TXTOBS(31) = NCLIEN
   End If
   TXTOBS(32) = VALOBADA("ESPECIFIC", "CODCLIENTE", "COD_INTE = " & XVALO(TXTCODINT), "NOMESS")
   Set VectArticulos = Nothing
   Me.Caption = "FICHA TÉCNICA"
   
End Sub


Private Sub TXTCORTE_Change()
   Call ACTUALIZA_TXTUSO 'SI CAMBIA EL CORTE ACUTALIZA EL USO EN LA CAJA DE TEXTO TXTUSO SEGUN  EN EL NUEVO VALOR DE USO SEGUN EL CORTE.

End Sub

Private Sub TXTCORTE_LostFocus()
   TXTCORTE = FORMATNUM$(XVALO(TXTCORTE), "F10.2")
End Sub

Private Sub TXTMED_Change(Index As Integer)
  Call ARMAMEDIDA
  TXTCORTE = FORMATNUM$(XVALO(TXTMED(3)), "F10.2")
  
End Sub
Sub ARMAMEDIDA()
   'GRABA  EN EL CASILLERO QUE ESTA DE BAJO OCULTO
   TXTDESC(1) = ""
   TXTDESC(1) = XVALO(Me.TXTMED(0)) & " X " & XVALO(Me.TXTMED(1)) & " X " & XVALO(Me.TXTMED(2)) & " X " & XVALO(Me.TXTMED(3))
End Sub


Private Sub TXTMED_KeyPress(Index As Integer, KeyAscii As Integer)
   
   If KeyAscii = 13 Then
     If Index >= 3 Then Call PINTATEXT(TXTOBS(2)): Exit Sub
      If Index < TXTMED.UBound Then
         Call PINTATEXT(TXTMED(Index + 1))
      Else
         Call PINTATEXT(TXTMED(TXTMED.LBound))
      End If
   End If
End Sub

Private Sub TXTOBS_Change(Index As Integer)
   If Index = 31 And XVALO(Me.TXTOBS(31)) > 0 Then
      Me.TXTRAZOCLI = rasocli$(XVALO(Me.TXTOBS(31)))
   End If
End Sub

Private Sub TXTOBS_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Index >= 30 Then Exit Sub ' OBSERVACIONES PARA QUE NO SE VAYA A OTRA CAJA CUANDO DA ENTER
      If Index < TXTOBS.UBound Then
         Call PINTATEXT(TXTOBS(Index + 1))
      Else
         Call PINTATEXT(TXTOBS(TXTOBS.LBound))
      End If
   End If
End Sub

Private Sub TXTREVISO_DblClick()
   If TRIM$(TXTAPROBO) <> "" Then
      Call COMUNI("Esta Estructura Está Aprobada\Imposible Modificar\Solo Podrá Modificarla quien Aprobó")
      Exit Sub
   End If
   FIRMAELEC.Text1 = ""
   FIRMAELEC.Text2 = ""
   FIRMAELEC.Label1 = ""
   FIRMAELEC.Show 1
   
   If FIRMAELEC.ESTADO > 0 Then
        If NIVCONFI%(XVALO(FIRMAELEC.Label1)) < 2 Then
         Call COMUNI("Usuario Con Nivel de Clasificación Inválido Para Esta Operación")
         Exit Sub
        End If
        '
        OPX% = COMALTER%("Atención Al Grabarse la Revisión\Se Actualizará Toda la Información Tal Cual Esta en Pantalla\\Mantener Información Original\Actualizar")
        If OPX% < 2 Then Exit Sub
        TXTREVISO = USERNAM$(XVALO(FIRMAELEC.Label1))
        Me.TXTFECHREVISO = FECHATEX$(HOY(HO$))

        Call GUARDARFRMENVAP
        Call TXTCODIGO_Change
   End If
   
   Unload FIRMAELEC

End Sub

Sub CARGARUTA()
      'PERMITE AGREGAR NUEVOS POE'S
      '
      CODI$ = TRIM$(TXTCODIGO.TEXT)
      If CODI$ = "" Then
         Call COMUNI("Imposible Continuar !!\Debe Seleccionar Un Código")
         Exit Sub
      End If
      
      CD1.FILTER = "(*.*)|*.*|"
      CD1.ShowOpen
      Ruta$ = CD1.FileName
      If TRIM(Ruta$) = "" Then Exit Sub
      '
20    DESCRI$ = TRIM$(InputBox$("Ingrese Texto ", "Guardar Como", CODI$))
      If DESCRI$ = "" Then Exit Sub
      If Len(DESCRI$) > 24 Then
        Call COMUNI("Longitud de Descripción \Supera La Cantidad Máxima de Caracteres (24)")
        GoTo 20
      End If
      On Error Resume Next
      CARPETA_ORIGEN$ = Ruta$
      POS1% = InStrRev(Ruta$, "\")
      POS2% = InStrRev(Ruta$, ".")
      CARPETA_DESTINO$ = Control("RUTAIMAN", "RUTAIMAG") & "\" & TRIM$(DESCRI$) & Mid$(Ruta$, POS2%)
      FileCopy CARPETA_ORIGEN$, CARPETA_DESTINO$
      If Err > 0 Then
       A = 1
      End If
      On Error GoTo 0
      'AGREGAR INFORMACION
      Call REPLA(TX$, DESCRI$, 1, 24)
      Call REPLA(TX$, CARPETA_DESTINO$, 32, 512)
      List2.AddItem TX$

End Sub

Private Sub List2_DblClick()
   Linea$ = List2.List(List2.ListIndex)
   Archi$ = TRIM$(Mid$(Linea$, 25))
   If (EXISTE%(Archi$)) <> 1 Then
       Call COMUNI("Archivo Inexistente")
       Exit Sub
     Else
        Call ShellExecute(Me.hwnd, "Open", Archi$, "", "", 1)
    End If
End Sub

