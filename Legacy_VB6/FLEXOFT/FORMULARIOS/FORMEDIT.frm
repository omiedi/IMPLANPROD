VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FORMEDIT 
   BackColor       =   &H00800000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   11070
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   11070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2280
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   66
      Top             =   1320
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.TextBox ENTRA 
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
      Index           =   31
      Left            =   0
      TabIndex        =   65
      Text            =   " "
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton BOTSEL 
      Caption         =   "..."
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
      Index           =   31
      Left            =   1890
      Picture         =   "FORMEDIT.frx":0000
      TabIndex        =   64
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.TextBox ENTRA 
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
      Index           =   30
      Left            =   0
      TabIndex        =   63
      Text            =   " "
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton BOTSEL 
      Caption         =   "..."
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
      Index           =   30
      Left            =   1890
      Picture         =   "FORMEDIT.frx":030A
      TabIndex        =   62
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.TextBox ENTRA 
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
      Index           =   29
      Left            =   0
      TabIndex        =   61
      Text            =   " "
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton BOTSEL 
      Caption         =   "..."
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
      Index           =   29
      Left            =   1890
      Picture         =   "FORMEDIT.frx":0614
      TabIndex        =   60
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.TextBox ENTRA 
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
      Index           =   28
      Left            =   0
      TabIndex        =   59
      Text            =   " "
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton BOTSEL 
      Caption         =   "..."
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
      Index           =   28
      Left            =   1890
      Picture         =   "FORMEDIT.frx":091E
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.TextBox ENTRA 
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
      Index           =   27
      Left            =   0
      TabIndex        =   57
      Text            =   " "
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton BOTSEL 
      Caption         =   "..."
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
      Index           =   27
      Left            =   1890
      Picture         =   "FORMEDIT.frx":0C28
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.TextBox ENTRA 
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
      Index           =   26
      Left            =   0
      TabIndex        =   55
      Text            =   " "
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton BOTSEL 
      Caption         =   "..."
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
      Index           =   26
      Left            =   1890
      Picture         =   "FORMEDIT.frx":0F32
      TabIndex        =   54
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.TextBox ENTRA 
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
      Index           =   25
      Left            =   0
      TabIndex        =   53
      Text            =   " "
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton BOTSEL 
      Caption         =   "..."
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
      Index           =   25
      Left            =   1890
      Picture         =   "FORMEDIT.frx":123C
      TabIndex        =   52
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.TextBox ENTRA 
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
      Index           =   24
      Left            =   0
      TabIndex        =   51
      Text            =   " "
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton BOTSEL 
      Caption         =   "..."
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
      Index           =   24
      Left            =   1890
      Picture         =   "FORMEDIT.frx":1546
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.TextBox ENTRA 
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
      Index           =   23
      Left            =   0
      TabIndex        =   49
      Text            =   " "
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton BOTSEL 
      Caption         =   "..."
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
      Index           =   23
      Left            =   1890
      Picture         =   "FORMEDIT.frx":1850
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.TextBox ENTRA 
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
      Index           =   22
      Left            =   0
      TabIndex        =   47
      Text            =   " "
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton BOTSEL 
      Caption         =   "..."
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
      Index           =   22
      Left            =   1890
      Picture         =   "FORMEDIT.frx":1B5A
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.TextBox ENTRA 
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
      Index           =   21
      Left            =   0
      TabIndex        =   45
      Text            =   " "
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton BOTSEL 
      Caption         =   "..."
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
      Index           =   21
      Left            =   1890
      Picture         =   "FORMEDIT.frx":1E64
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.TextBox ENTRA 
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
      Index           =   20
      Left            =   0
      TabIndex        =   43
      Text            =   " "
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton BOTSEL 
      Caption         =   "..."
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
      Index           =   20
      Left            =   1890
      Picture         =   "FORMEDIT.frx":216E
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.TextBox ENTRA 
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
      Index           =   19
      Left            =   0
      TabIndex        =   41
      Text            =   " "
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton BOTSEL 
      Caption         =   "..."
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
      Index           =   19
      Left            =   1890
      Picture         =   "FORMEDIT.frx":2478
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.TextBox ENTRA 
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
      Index           =   18
      Left            =   0
      TabIndex        =   39
      Text            =   " "
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton BOTSEL 
      Caption         =   "..."
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
      Index           =   18
      Left            =   1890
      Picture         =   "FORMEDIT.frx":2782
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.TextBox ENTRA 
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
      Index           =   17
      Left            =   0
      TabIndex        =   37
      Text            =   " "
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton BOTSEL 
      Caption         =   "..."
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
      Index           =   17
      Left            =   1890
      Picture         =   "FORMEDIT.frx":2A8C
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.TextBox ENTRA 
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
      Index           =   16
      Left            =   0
      TabIndex        =   35
      Text            =   " "
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton BOTSEL 
      Caption         =   "..."
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
      Index           =   16
      Left            =   1890
      Picture         =   "FORMEDIT.frx":2D96
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.TextBox ENTRA 
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
      Index           =   15
      Left            =   0
      TabIndex        =   33
      Text            =   " "
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton BOTSEL 
      Caption         =   "..."
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
      Index           =   15
      Left            =   1890
      Picture         =   "FORMEDIT.frx":30A0
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.TextBox ENTRA 
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
      Index           =   14
      Left            =   0
      TabIndex        =   31
      Text            =   " "
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton BOTSEL 
      Caption         =   "..."
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
      Index           =   14
      Left            =   1890
      Picture         =   "FORMEDIT.frx":33AA
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.TextBox ENTRA 
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
      Index           =   13
      Left            =   0
      TabIndex        =   29
      Text            =   " "
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton BOTSEL 
      Caption         =   "..."
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
      Index           =   13
      Left            =   1890
      Picture         =   "FORMEDIT.frx":36B4
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.TextBox ENTRA 
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
      Left            =   0
      TabIndex        =   27
      Text            =   " "
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton BOTSEL 
      Caption         =   "..."
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
      Index           =   12
      Left            =   1890
      Picture         =   "FORMEDIT.frx":39BE
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.TextBox ENTRA 
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
      Left            =   0
      TabIndex        =   25
      Text            =   " "
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton BOTSEL 
      Caption         =   "..."
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
      Index           =   11
      Left            =   1890
      Picture         =   "FORMEDIT.frx":3CC8
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.TextBox ENTRA 
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
      Left            =   0
      TabIndex        =   23
      Text            =   " "
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton BOTSEL 
      Caption         =   "..."
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
      Index           =   10
      Left            =   1890
      Picture         =   "FORMEDIT.frx":3FD2
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.TextBox ENTRA 
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
      Left            =   0
      TabIndex        =   21
      Text            =   " "
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton BOTSEL 
      Caption         =   "..."
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
      Index           =   9
      Left            =   1890
      Picture         =   "FORMEDIT.frx":42DC
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.TextBox ENTRA 
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
      Left            =   0
      TabIndex        =   19
      Text            =   " "
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton BOTSEL 
      Caption         =   "..."
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
      Index           =   8
      Left            =   1890
      Picture         =   "FORMEDIT.frx":45E6
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.TextBox ENTRA 
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
      Left            =   0
      TabIndex        =   17
      Text            =   " "
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton BOTSEL 
      Caption         =   "..."
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
      Index           =   7
      Left            =   1890
      Picture         =   "FORMEDIT.frx":48F0
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.TextBox ENTRA 
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
      Left            =   0
      TabIndex        =   15
      Text            =   " "
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton BOTSEL 
      Caption         =   "..."
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
      Index           =   6
      Left            =   1890
      Picture         =   "FORMEDIT.frx":4BFA
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.TextBox ENTRA 
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
      Left            =   0
      TabIndex        =   13
      Text            =   " "
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton BOTSEL 
      Caption         =   "..."
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
      Index           =   5
      Left            =   1890
      Picture         =   "FORMEDIT.frx":4F04
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.TextBox ENTRA 
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
      Left            =   0
      TabIndex        =   11
      Text            =   " "
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton BOTSEL 
      Caption         =   "..."
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
      Index           =   4
      Left            =   1890
      Picture         =   "FORMEDIT.frx":520E
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.TextBox ENTRA 
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
      Left            =   0
      TabIndex        =   9
      Text            =   " "
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton BOTSEL 
      Caption         =   "..."
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
      Index           =   3
      Left            =   1890
      Picture         =   "FORMEDIT.frx":5518
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.TextBox ENTRA 
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
      Left            =   0
      TabIndex        =   7
      Text            =   " "
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton BOTSEL 
      Caption         =   "..."
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
      Index           =   2
      Left            =   1890
      Picture         =   "FORMEDIT.frx":5822
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.TextBox ENTRA 
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
      Left            =   0
      TabIndex        =   5
      Text            =   " "
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton BOTSEL 
      Caption         =   "..."
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
      Left            =   1890
      Picture         =   "FORMEDIT.frx":5B2C
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.TextBox ENTRA 
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
      Left            =   3465
      TabIndex        =   3
      Text            =   " "
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton BOTSEL 
      Caption         =   "..."
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
      Left            =   5355
      Picture         =   "FORMEDIT.frx":5E36
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.CommandButton Command2 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   8085
      TabIndex        =   1
      Top             =   0
      Width           =   435
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7665
      TabIndex        =   0
      Top             =   0
      Width           =   435
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   750
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "FORMEDIT.frx":6140
      TabIndex        =   67
      Top             =   0
      Visible         =   0   'False
      Width           =   1590
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   3600
      OleObjectBlob   =   "FORMEDIT.frx":61BA
      Top             =   1200
   End
End
Attribute VB_Name = "FORMEDIT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    '
    CONTROLGRABA% = 0
    Call GRAENTRA
    MODIFI% = 1
    ' WEDITAB.LITAB.Refresh
    YAPRE% = 0
    If CONTROLGRABA% = 1 Then
       Hide
       Call FRESHALL
    End If
    '
End Sub
'

Private Sub Command2_Click()
    '
    Call CANCELENTRA
    '
End Sub
'

Private Sub ENTRA_Change(Index As Integer)
    '
    ARCHEC$ = TRIM$(AE$(Index + 1))
    VNUE$ = ENTRA(Index).TEXT
    COLOG% = Index + 1
    If CE%(COLOG%) > 0 Then
        XXZ$ = CBIN$(VNUE$, TVL%(CFT%(Index% + 1)), LF%(CFT%(Index% + 1)))
        If TVL%(CFT%(Index% + 1)) = 6 Then XXZ$ = TRIM$(XXZ$)
        XXE$ = ""
        If XXZ$ <> "" Then
            XXE$ = ECOARCH$(ARCHEC$, XXZ$)
        End If
        ENTRA(CE%(COLOG%) - 1).TEXT = XXE$
        ENTRA(CE%(COLOG%) - 1).Refresh
        If Left$(ARCHEC$, 5) = "LIPRE" Then
          NLIS% = Val(Mid$(ARCHEC$, 6, 3))
          If NLIS% > 0 Then
            PRELIX$ = TRIM$(CSTRING$(MKD$(PRELISTA#(ARCHEC$, CVI(XXZ$))), 3, 12, 4, 2))
            ENTRA(CE%(COLOG%)).TEXT = AJUSTD$(PRELIX$, lC%(1 + CE%(COLOG%)))
          End If
        End If
    End If
    '
End Sub
'

Private Sub ENTRA_DBLCLICK(Index As Integer)
    '
    If PE%(Index + 1) <> 1 Then
       If (MODEDI% = 4 Or PE%(Index + 1) = 3) Then
          Exit Sub
       End If
    End If
    '
    ARCHEC$ = TRIM$(AE$(Index + 1))
    If TVL%(CFT%(Index + 1)) = 5 Then ARCHEC$ = "SELFECHA"
    If ARCHEC$ <> "" Then
        VVIE$ = ENTRA(Index).TEXT
        SELECONTROL% = 1
        Call SELECHO(ARCHEC$)
        SELECONTROL% = 0
        VNUE$ = VALACT1$(ARCHEC$)
        If TRIM$(VNUE$) = "" Then VNUE$ = VVIE$
        ENTRA(Index).TEXT = VNUE$
    End If
    Call SETNEXTFOC(Index)
End Sub
'

Private Sub ENTRA_GotFocus(Index As Integer)
    '
    COLUAC% = Index
    If PE%(Index + 1) <> 1 Then
       If (MODEDI% = 4 Or PE%(Index + 1) = 3) Then
          Call SETNEXTFOC(Index)
       End If
    End If
    '
End Sub

Private Sub ENTRA_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
       If PE%(Index + 1) <> 1 Then
          If (MODEDI% = 4 Or PE%(Index + 1) = 3) Then
             Exit Sub
          End If
       End If
       '
       ARCHEC$ = TRIM$(AE$(Index + 1))
       If TVL%(CFT%(Index + 1)) = 5 Then ARCHEC$ = "SELFECHA"
       If ARCHEC$ <> "" Then
          VVIE$ = ENTRA(Index).TEXT
          Call SELECHO(ARCHEC$)
          VNUE$ = VALACT1$(ARCHEC$)
          If TRIM$(VNUE$) = "" Then VNUE$ = VVIE$
          ENTRA(Index).TEXT = VNUE$
       End If
       Call SETNEXTFOC(Index)
    End If
End Sub

Private Sub ENTRA_KeyPress(Index As Integer, KeyAscii As Integer)
    '
    If KeyAscii = 13 Then
       Call SETNEXTFOC(Index + 1)
       'If Index < NCL% - 1 Then
       '    ENTRA(Index + 1).SetFocus
       '    Exit Sub
       'End If
       'Command1.SetFocus
     ElseIf KeyAscii = 27 Then
       Call CANCELENTRA
    End If
    '
End Sub

Private Sub ENTRA_LostFocus(Index As Integer)
    '
    If PE%(Index + 1) <> 1 Then
       If (MODEDI% = 4 Or PE%(Index + 1) = 3) Then
          Exit Sub
       End If
    End If
    '
    IX0% = CFT%(Index + 1)
    XXZ$ = CBIN$(TRIM$(ENTRA(Index).TEXT), TVL%(IX0%), LF%(IX0%))
    XXY$ = CSTRING$(XXZ$, TVL%(IX0%), lC%(Index% + 1), LDL%(IX0%), JT%(Index + 1))
    If MG%(Index + 1) > 0 Then
       If JT%(Index + 1) = 1 Then
           XXY$ = Left$(Space$(MG%(Index + 1)) + XXY$, lC%(Index + 1))
         ElseIf JT%(Index + 1) = 2 Then
           If TRIM$(Left$(XXY$, MG%(Index + 1))) = "" Then
               XXY$ = Mid$(XXY$, 1 + MG%(Index + 1)) + Space$(MG%(Index + 1))
           End If
       End If
    End If
    '
    If Left$(TRIM$(XXY$), 1) <> "*" Then
        ENTRA(Index).TEXT = XXY$
    End If
    '
End Sub

Private Sub Form_Activate()

    If COLUAC% = 0 Then
        On Error Resume Next
        ENTRA(0).SetFocus
        On Error GoTo 0
      Else
        Call SETNEXTFOC(COLUAC%)
    End If
End Sub
