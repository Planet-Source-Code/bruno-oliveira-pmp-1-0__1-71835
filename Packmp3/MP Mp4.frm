VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "Packmp3  1.0(Beta) "
   ClientHeight    =   8820
   ClientLeft      =   4425
   ClientTop       =   615
   ClientWidth     =   6735
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "MP Mp4.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8820
   ScaleWidth      =   6735
   Begin VB.PictureBox Picture5 
      Height          =   4095
      Left            =   360
      ScaleHeight     =   4035
      ScaleWidth      =   4275
      TabIndex        =   100
      Top             =   4080
      Visible         =   0   'False
      Width           =   4335
      Begin VB.Label Label27 
         Alignment       =   2  'Center
         Caption         =   "Wait While ordering the list..."
         Height          =   735
         Left            =   960
         TabIndex        =   101
         Top             =   1200
         Width           =   2895
      End
   End
   Begin VB.CommandButton Command17 
      BackColor       =   &H000080FF&
      Caption         =   "Order By Title"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   99
      ToolTipText     =   $"MP Mp4.frx":044A
      Top             =   1320
      Width           =   4455
   End
   Begin VB.CheckBox Check9 
      BackColor       =   &H80000012&
      Caption         =   "Time Remaining"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   165
      Left            =   2160
      TabIndex        =   98
      Top             =   720
      Width           =   2175
   End
   Begin VB.CheckBox Check8 
      BackColor       =   &H80000012&
      Caption         =   "Time Elapsed"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   165
      Left            =   360
      TabIndex        =   97
      Top             =   720
      Width           =   1815
   End
   Begin VB.Timer Timer9 
      Interval        =   1
      Left            =   7440
      Top             =   1440
   End
   Begin VB.CheckBox Check7 
      BackColor       =   &H80000009&
      Caption         =   "Random Mode"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   4920
      TabIndex        =   95
      Top             =   7680
      Width           =   1455
   End
   Begin VB.CommandButton Command16 
      BackColor       =   &H000080FF&
      Caption         =   "Cd Player"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   94
      Top             =   6120
      Width           =   1335
   End
   Begin VB.CheckBox Check6 
      BackColor       =   &H80000009&
      Caption         =   "Intro 30 seg"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   4920
      TabIndex        =   93
      Top             =   7920
      Width           =   1455
   End
   Begin VB.CommandButton Command15 
      BackColor       =   &H000080FF&
      Caption         =   "Internet Radio"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   90
      Top             =   5880
      Width           =   1335
   End
   Begin VB.CheckBox system 
      Caption         =   "System"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4320
      TabIndex        =   75
      Top             =   1200
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CheckBox Hidden 
      Caption         =   "Hidden"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   4320
      TabIndex        =   74
      Top             =   1440
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CheckBox archive 
      Caption         =   "Archive"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4320
      TabIndex        =   73
      Top             =   1680
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CheckBox read 
      Caption         =   "Read Only"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4320
      TabIndex        =   72
      Top             =   1920
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H000080FF&
      Caption         =   "Speed  -"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   5400
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6840
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   240
      TabIndex        =   69
      Text            =   "Text1"
      Top             =   1920
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00004040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   3600
      TabIndex        =   38
      Top             =   3000
      Width           =   735
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   4
         Left            =   480
         ScaleHeight     =   75
         ScaleWidth      =   75
         TabIndex        =   68
         Top             =   360
         Width           =   135
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   5
         Left            =   480
         ScaleHeight     =   75
         ScaleWidth      =   75
         TabIndex        =   65
         Top             =   240
         Width           =   135
         Begin VB.PictureBox Picture3 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   15
            Index           =   1
            Left            =   0
            ScaleHeight     =   15
            ScaleWidth      =   255
            TabIndex        =   67
            Top             =   240
            Width           =   255
         End
         Begin VB.PictureBox Picture4 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   15
            Index           =   1
            Left            =   0
            ScaleHeight     =   15
            ScaleWidth      =   375
            TabIndex        =   66
            Top             =   240
            Width           =   375
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   6
         Left            =   480
         ScaleHeight     =   75
         ScaleWidth      =   75
         TabIndex        =   64
         Top             =   120
         Width           =   135
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   7
         Left            =   480
         ScaleHeight     =   75
         ScaleWidth      =   75
         TabIndex        =   63
         Top             =   0
         Width           =   135
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00C00000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   8
         Left            =   360
         ScaleHeight     =   75
         ScaleWidth      =   75
         TabIndex        =   62
         Top             =   360
         Width           =   135
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00C00000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   9
         Left            =   360
         ScaleHeight     =   75
         ScaleWidth      =   75
         TabIndex        =   61
         Top             =   240
         Width           =   135
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00C00000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   10
         Left            =   360
         ScaleHeight     =   75
         ScaleWidth      =   75
         TabIndex        =   58
         Top             =   120
         Width           =   135
         Begin VB.PictureBox Picture4 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   15
            Index           =   2
            Left            =   0
            ScaleHeight     =   15
            ScaleWidth      =   375
            TabIndex        =   60
            Top             =   240
            Width           =   375
         End
         Begin VB.PictureBox Picture3 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   15
            Index           =   2
            Left            =   0
            ScaleHeight     =   15
            ScaleWidth      =   255
            TabIndex        =   59
            Top             =   240
            Width           =   255
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00C00000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   11
         Left            =   360
         ScaleHeight     =   75
         ScaleWidth      =   75
         TabIndex        =   57
         Top             =   0
         Width           =   135
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00C00000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   0
         Left            =   600
         ScaleHeight     =   75
         ScaleWidth      =   75
         TabIndex        =   56
         Top             =   360
         Width           =   135
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00C00000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   3
         Left            =   600
         ScaleHeight     =   75
         ScaleWidth      =   75
         TabIndex        =   55
         Top             =   0
         Width           =   135
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00C00000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   2
         Left            =   600
         ScaleHeight     =   75
         ScaleWidth      =   75
         TabIndex        =   52
         Top             =   120
         Width           =   135
         Begin VB.PictureBox Picture3 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   15
            Index           =   0
            Left            =   0
            ScaleHeight     =   15
            ScaleWidth      =   255
            TabIndex        =   54
            Top             =   240
            Width           =   255
         End
         Begin VB.PictureBox Picture4 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   15
            Index           =   0
            Left            =   0
            ScaleHeight     =   15
            ScaleWidth      =   375
            TabIndex        =   53
            Top             =   240
            Width           =   375
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00C00000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   1
         Left            =   600
         ScaleHeight     =   75
         ScaleWidth      =   75
         TabIndex        =   51
         Top             =   240
         Width           =   135
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   12
         Left            =   240
         ScaleHeight     =   75
         ScaleWidth      =   75
         TabIndex        =   50
         Top             =   360
         Width           =   135
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   13
         Left            =   240
         ScaleHeight     =   75
         ScaleWidth      =   75
         TabIndex        =   49
         Top             =   240
         Width           =   135
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   14
         Left            =   240
         ScaleHeight     =   75
         ScaleWidth      =   75
         TabIndex        =   46
         Top             =   120
         Width           =   135
         Begin VB.PictureBox Picture4 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   15
            Index           =   3
            Left            =   0
            ScaleHeight     =   15
            ScaleWidth      =   375
            TabIndex        =   48
            Top             =   240
            Width           =   375
         End
         Begin VB.PictureBox Picture3 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   15
            Index           =   3
            Left            =   0
            ScaleHeight     =   15
            ScaleWidth      =   255
            TabIndex        =   47
            Top             =   240
            Width           =   255
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   15
         Left            =   240
         ScaleHeight     =   75
         ScaleWidth      =   75
         TabIndex        =   45
         Top             =   0
         Width           =   135
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00C00000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   16
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   75
         TabIndex        =   44
         Top             =   360
         Width           =   135
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00C00000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   17
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   75
         TabIndex        =   41
         Top             =   240
         Width           =   135
         Begin VB.PictureBox Picture3 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   15
            Index           =   4
            Left            =   0
            ScaleHeight     =   15
            ScaleWidth      =   255
            TabIndex        =   43
            Top             =   240
            Width           =   255
         End
         Begin VB.PictureBox Picture4 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   15
            Index           =   4
            Left            =   0
            ScaleHeight     =   15
            ScaleWidth      =   375
            TabIndex        =   42
            Top             =   240
            Width           =   375
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00C00000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   18
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   75
         TabIndex        =   40
         Top             =   120
         Width           =   135
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00C00000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   19
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   75
         TabIndex        =   39
         Top             =   0
         Width           =   135
      End
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   35
      Text            =   "0"
      ToolTipText     =   "Song Frequency"
      Top             =   360
      Width           =   495
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   34
      Text            =   "0"
      ToolTipText     =   "Song Bitrate"
      Top             =   360
      Width           =   495
   End
   Begin VB.Timer Timer8 
      Interval        =   83
      Left            =   7080
      Top             =   1440
   End
   Begin VB.Timer Timer7 
      Interval        =   25
      Left            =   7440
      Top             =   1080
   End
   Begin VB.Timer Timer6 
      Interval        =   45
      Left            =   7080
      Top             =   1080
   End
   Begin VB.Timer Timer5 
      Interval        =   50
      Left            =   7080
      Top             =   720
   End
   Begin VB.Timer Timer4 
      Interval        =   120
      Left            =   7440
      Top             =   720
   End
   Begin VB.CommandButton Command14 
      BackColor       =   &H000080FF&
      Caption         =   "Directory"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   7800
      Top             =   240
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H000080FF&
      Caption         =   "Speed +"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   5160
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2880
      Picture         =   "MP Mp4.frx":04F0
      ScaleHeight     =   555
      ScaleWidth      =   435
      TabIndex        =   27
      Top             =   1680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H80000009&
      Caption         =   "Repeat all"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   4920
      TabIndex        =   25
      ToolTipText     =   "Select to repeat all songs"
      Top             =   7440
      UseMaskColor    =   -1  'True
      Width           =   1575
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H80000009&
      Caption         =   "Repeat 1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4920
      TabIndex        =   23
      ToolTipText     =   "Select to repeat one song"
      Top             =   7200
      Width           =   1335
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Century"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   4680
      TabIndex        =   22
      Top             =   360
      Width           =   255
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Century"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   4680
      TabIndex        =   21
      Top             =   1080
      Value           =   -1  'True
      Width           =   255
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   17
      Top             =   8565
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5821
            MinWidth        =   5821
            Text            =   "Copyright © 2008"
            TextSave        =   "Copyright © 2008"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "13-03-2009"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "14:40"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H000080FF&
      Caption         =   "Get Playlist"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4440
      Width           =   1335
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H000080FF&
      Caption         =   "Save Playlist"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Timer Timer3 
      Interval        =   1
      Left            =   7440
      Top             =   240
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   7080
      Top             =   240
   End
   Begin MSComctlLib.Slider Slider2 
      Height          =   255
      Left            =   2280
      TabIndex        =   8
      Top             =   960
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Min             =   -5000
      Max             =   5000
      TickStyle       =   3
   End
   Begin MSComctlLib.Slider Slider1 
      CausesValidation=   0   'False
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   960
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      LargeChange     =   10
      Max             =   2000
      TickStyle       =   3
      TickFrequency   =   10
      TextPosition    =   1
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H000080FF&
      Caption         =   "Clean Playlist"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4920
      MaskColor       =   &H80000010&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4920
      UseMaskColor    =   -1  'True
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H000080FF&
      Caption         =   "Remove"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4200
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4920
      MaskColor       =   &H80000014&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3960
      Width           =   1335
   End
   Begin MSComctlLib.Slider Slider3 
      Height          =   270
      Left            =   120
      TabIndex        =   26
      Top             =   2400
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   476
      _Version        =   393216
      SelectRange     =   -1  'True
      SelLength       =   10
      TickStyle       =   3
      TickFrequency   =   3
   End
   Begin MSComDlg.CommonDialog cd2 
      Left            =   6840
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   6840
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00004040&
      Caption         =   "Controlos:"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   975
      Left            =   120
      TabIndex        =   10
      Top             =   2640
      Width           =   6495
      Begin VB.CommandButton Command13 
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   3.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1080
         Picture         =   "MP Mp4.frx":093A
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Pause's the song"
         Top             =   240
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton Command10 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Random"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Plays a random song."
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   3.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         Picture         =   "MP Mp4.frx":108C
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Skip to previous"
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   3.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2280
         Picture         =   "MP Mp4.frx":1819
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Skip To Next"
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   3.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1680
         Picture         =   "MP Mp4.frx":1FDA
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Stop's the song"
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   3.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1080
         Picture         =   "MP Mp4.frx":26FF
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Play's a song"
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         BackColor       =   &H00004000&
         Caption         =   "Next"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   135
         Left            =   2280
         TabIndex        =   87
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         BackColor       =   &H00004000&
         Caption         =   "Stop"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   135
         Left            =   1800
         TabIndex        =   86
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         BackColor       =   &H00004000&
         Caption         =   "Play/Pause"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   135
         Left            =   1080
         TabIndex        =   85
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BackColor       =   &H00004000&
         Caption         =   "Previous"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   135
         Left            =   480
         TabIndex        =   84
         Top             =   720
         Width           =   615
      End
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H80000009&
      Caption         =   "No Sound"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4920
      MaskColor       =   &H000000FF&
      TabIndex        =   24
      ToolTipText     =   "Select to put mute"
      Top             =   6480
      Width           =   1455
   End
   Begin VB.CheckBox Check4 
      BackColor       =   &H80000009&
      Caption         =   "Reset"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   4920
      TabIndex        =   81
      ToolTipText     =   "Select to return to the beginning of the playlist"
      Top             =   6960
      Width           =   855
   End
   Begin VB.CheckBox Check5 
      BackColor       =   &H80000009&
      Caption         =   "Countinous"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   4920
      TabIndex        =   82
      ToolTipText     =   "Select to play all the songs until the end of the playlist"
      Top             =   6720
      Width           =   1335
   End
   Begin VB.ListBox List2 
      BackColor       =   &H00808000&
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   4050
      ItemData        =   "MP Mp4.frx":2E92
      Left            =   360
      List            =   "MP Mp4.frx":2E94
      TabIndex        =   19
      Top             =   4080
      Width           =   4335
   End
   Begin VB.ListBox List1 
      BackColor       =   &H80000001&
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3900
      Left            =   360
      TabIndex        =   9
      Top             =   4080
      Width           =   1575
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000007&
      Caption         =   "Label8"
      Height          =   255
      Left            =   1320
      TabIndex        =   80
      Top             =   360
      Width           =   135
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      Caption         =   "0:00       0:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   360
      OLEDropMode     =   1  'Manual
      TabIndex        =   76
      ToolTipText     =   "Song Time"
      Top             =   360
      UseMnemonic     =   0   'False
      Width           =   2055
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label26 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "0:00       0:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   360
      TabIndex        =   96
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "(No File)"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   2040
      TabIndex        =   92
      Top             =   8160
      Width           =   1095
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   "File Extension:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   480
      TabIndex        =   91
      Top             =   8160
      Width           =   1455
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Caption         =   "(No File)"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   3840
      TabIndex        =   89
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "Playing:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2760
      TabIndex        =   88
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   2040
      TabIndex        =   78
      ToolTipText     =   "Number of songs"
      Top             =   3840
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   480
      TabIndex        =   11
      Top             =   3840
      Width           =   1815
   End
   Begin VB.Label Label15 
      BackColor       =   &H80000007&
      Height          =   1455
      Left            =   6480
      TabIndex        =   83
      Top             =   1800
      Width           =   375
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   4575
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   3840
      Width           =   6495
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   615
      Left            =   120
      TabIndex        =   79
      Top             =   1680
      Width           =   18495
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   4800
      TabIndex        =   77
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "See Playlist"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   5040
      TabIndex        =   71
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Hide Playlist"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   5040
      TabIndex        =   70
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "kb's"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   3000
      TabIndex        =   37
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "kh'z"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   3840
      TabIndex        =   36
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   32
      Top             =   1440
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1920
      TabIndex        =   28
      Top             =   1800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1080
      TabIndex        =   18
      Top             =   3480
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   3000
      TabIndex        =   13
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   1080
      TabIndex        =   12
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   14
      Top             =   1320
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   120
      Picture         =   "MP Mp4.frx":2E96
      Stretch         =   -1  'True
      Top             =   120
      Width           =   4395
   End
   Begin VB.Menu mnupopup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu EI 
         Caption         =   "Fullscreen Videos"
      End
      Begin VB.Menu minvid 
         Caption         =   "Minimize Videos"
      End
      Begin VB.Menu restvid 
         Caption         =   "Restaure Videos"
      End
      Begin VB.Menu mnufich 
         Caption         =   "File"
         Begin VB.Menu mnuabr 
            Caption         =   "Open"
         End
      End
      Begin VB.Menu mnuac 
         Caption         =   "Functions"
         Begin VB.Menu mnutoc 
            Caption         =   "Play"
         End
         Begin VB.Menu mnupar 
            Caption         =   "Stop"
         End
         Begin VB.Menu mnupas 
            Caption         =   "Pause"
         End
         Begin VB.Menu mnuseg 
            Caption         =   "Next"
         End
         Begin VB.Menu mnuant 
            Caption         =   "Previous"
         End
      End
      Begin VB.Menu mnucon 
         Caption         =   "Sound Control"
      End
      Begin VB.Menu mnucor 
         Caption         =   "Backcolor"
         Begin VB.Menu mnubranc 
            Caption         =   "White"
         End
         Begin VB.Menu mnupret 
            Caption         =   "Black"
         End
         Begin VB.Menu verde 
            Caption         =   "Green"
         End
         Begin VB.Menu Vermelho 
            Caption         =   "Red"
         End
         Begin VB.Menu Azul 
            Caption         =   "Blue"
         End
         Begin VB.Menu Laranja 
            Caption         =   "Orange"
         End
      End
      Begin VB.Menu mnusp 
         Caption         =   "-"
      End
      Begin VB.Menu mnusair 
         Caption         =   "Close"
      End
   End
   Begin VB.Menu mnuteste 
      Caption         =   "Img"
      Visible         =   0   'False
      Begin VB.Menu mnusbr 
         Caption         =   "About"
      End
      Begin VB.Menu mnsep 
         Caption         =   "-"
      End
      Begin VB.Menu mnutsair 
         Caption         =   "Close"
      End
   End
   Begin VB.Menu mnuinf 
      Caption         =   "Informations"
      Visible         =   0   'False
      Begin VB.Menu mnuid3 
         Caption         =   "Id3"
      End
      Begin VB.Menu mnudes 
         Caption         =   "Description"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2

Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4

Private Const WM_LeBUTTONDBLCLCK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_RBUTTONUP = &H205

Private Type NOTIFYICONDATA
cbSize As Long
hwnd As Long
uId As Long
uFlags As Long
ucallbackMessage As Long
hIcon As Long
szTip As String * 64
End Type

Private st1 As String, st2 As String, sti As Integer


Dim img As NOTIFYICONDATA
Dim paused As Boolean
Dim t1 As Integer
Dim a As Integer
'Dim sp As New SpeechLib.SpVoice





Private Sub Azul_Click()
Form1.BackColor = &HC00000
Option1.BackColor = &HC00000
Option2.BackColor = &HC00000
Option1.ForeColor = black
Option2.ForeColor = black
Shape2.FillColor = black
Check1.BackColor = black
Check2.BackColor = black
Check3.BackColor = black
Check4.BackColor = black
Check5.BackColor = black
Check6.BackColor = black
Check7.BackColor = black
'Check1.ForeColor = &H8000000E
'Check2.ForeColor = &H8000000E
'Check3.ForeColor = &H8000000E
'Check4.ForeColor = &H8000000E
'Check5.ForeColor = &H8000000E
'Label1.ForeColor = &H8000000E
Label13.ForeColor = black
Label14.ForeColor = black
Label15.BackColor = &HC00000
Form1.Refresh
End Sub

'Dim allow_stop As Boolean
'Dim allow_pause As Boolean
Private Sub Check1_Click()
If Check1.Value = 1 Then
Form7.MediaPlayer1.Mute = True
Else
Form7.MediaPlayer1.Mute = False
End If
End Sub



Private Sub Check2_Click()
If Check2.Value = 1 Then
Check3.Enabled = False
Check4.Enabled = False
Check7.Enabled = False
Else
Check3.Enabled = True
Check4.Enabled = True
Check7.Enabled = True
End If
End Sub

Private Sub Check3_Click()
If Check3.Value = 1 Then
Check2.Enabled = False
Check4.Enabled = False
Check7.Enabled = False
Else
Check2.Enabled = True
Check4.Enabled = True
Check7.Enabled = True
End If
End Sub

Private Sub Check4_Click()
If Check4.Value = 1 Then
Check2.Enabled = False
Check3.Enabled = False
Check7.Enabled = False
Else
Check2.Enabled = True
Check3.Enabled = True
Check7.Enabled = True
End If
End Sub

Private Sub Check5_Click()
If Check5.Value = 1 Then
Check2.Enabled = True
Check3.Enabled = True
Check4.Enabled = True
Check7.Enabled = True
Else
Check2.Value = 0
Check3.Value = 0
Check4.Value = 0
Check7.Value = 0
Check2.Enabled = False
Check3.Enabled = False
Check4.Enabled = False
Check7.Enabled = False
End If
End Sub

Private Sub Check6_Click()
On Error Resume Next
If Check6.Value = 1 Then
Slider3.Enabled = False
End If
If Check6.Value = 0 Then
Slider3.Enabled = True
End If
End Sub

Private Sub Check7_Click()
If Check7.Value = 1 Then
Check4.Enabled = False
Check3.Enabled = False
Check2.Enabled = False
Else
Check4.Enabled = True
Check3.Enabled = True
Check2.Enabled = True
End If
End Sub

Private Sub Check8_Click()
If Check8.Value = 1 Then
Label5.Visible = True
Label26.Visible = False
Check9.Value = 0
End If
If Check8.Value = 0 Then
Check9.Value = 1
End If
End Sub

Private Sub Check9_Click()
If Check9.Value = 1 Then
Label5.Visible = False
Label26.Visible = True
Check8.Value = 0
End If
If Check9.Value = 0 Then
Check8.Value = 1
End If
End Sub

Private Sub Command1_Click()
On Error GoTo hell
List1.ListIndex = List2.ListIndex
If Check1.Value = 1 Then
Form7.MediaPlayer1.Mute = True
Else
Form7.MediaPlayer1.Mute = False
End If
If List1.Text <> "" Then
If paused = False Then
Check6.Enabled = False
Text1.Text = List2.Text
Label6.Caption = "Now Playing:" + " " + Text1.Text
Label10.Caption = Text1.Text
Form7.MediaPlayer1.FileName = List1.Text
Form7.MediaPlayer1.Play
Label25.Caption = Right(List1.Text, 3)
Text2.Locked = False
Text3.Locked = False
Text2.Text = (Round((FileLen(Form7.MediaPlayer1.FileName) / Form7.MediaPlayer1.SelectionEnd) / 125))
'get bitrate
Text3.Text = "49"
'Text3.Text = Int(FileLen(MediaPlayer1.FileName) / 100) & " Kb" 'get file length (bytes)
'Text4.Text = Int((MediaPlayer1.SelectionEnd / 60 * 100)) / 100 & " mins" 'get length (mins)
paused = False
Command1.Visible = False
Command13.Visible = True
mnutoc.Enabled = False
mnupas.Enabled = True
Text2.Locked = True
Text3.Locked = True
If Form7.MediaPlayer1.ImageSourceHeight = 0 Then
Label23.Caption = "(Audio)"
Form7.Hide
Else
Label23.Caption = "(Video)"
Load Form7
Form7.Show
End If
Else
Check6.Enabled = False
Label6.Caption = "Now Playing:" + " " + Text1.Text
Label10.Caption = Text1.Text
Form7.MediaPlayer1.Play
Label25.Caption = Right(List1.Text, 3)
vol = Abs(Slider1.Value) - 2000
Form7.MediaPlayer1.Volume = vol
paused = False
'allow_stop = True
'allow_pause = True
Command1.Visible = False
Command13.Visible = True
mnutoc.Enabled = False
mnupas.Enabled = True
If Form7.MediaPlayer1.ImageSourceHeight = 0 Then
Label23.Caption = "(Audio)"
Form7.Hide
Else
Label23.Caption = "(Video)"
Load Form7
Form7.Show
End If
End If
End If
Exit Sub
hell:
'sp.speak "The music" + List2.Text + " isn't supported"
'If List1.ListIndex > "-1 " Then
End Sub

Private Sub Command10_Click()
On Error Resume Next
If List1.ListCount > 0 And List2.ListCount > 0 Then
Randomize
List1.ListIndex = Int(List1.ListCount * Rnd)
List2.ListIndex = List1.ListIndex
Form7.MediaPlayer1.FileName = List1.Text
Text1.Text = List2.Text
Label6.Caption = "Now Playing:" + " " + Text1.Text
Label10.Caption = Text1.Text
Label25.Caption = Right(List1.Text, 3)
Text2.Locked = False
Text2.Text = (Int((FileLen(Form7.MediaPlayer1.FileName) / Form7.MediaPlayer1.SelectionEnd) / 125))  'get bitrate
Text2.Locked = True
Text3.Locked = False
Text3.Text = 49
Command1.Visible = False
Command13.Visible = True
mnutoc.Enabled = False
mnupas.Enabled = True
Form7.MediaPlayer1.Play
Check6.Enabled = False
If Form7.MediaPlayer1.ImageSourceHeight = 0 Then
Label23.Caption = "(Audio)"
Form7.Hide
Else
Label23.Caption = "(Video)"
Load Form7
Form7.Show
End If
End If
End Sub

Private Sub Command11_Click()
On Error Resume Next
Form7.MediaPlayer1.Rate = Label9.Caption
If Label9.Caption < 1.1 Then
Label9.Caption = Label9.Caption + 0.1
Form7.MediaPlayer1.Rate = Label9.Caption
End If
End Sub

Private Sub Command12_Click()
On Error Resume Next
Form7.MediaPlayer1.Rate = Label9.Caption
If Label9.Caption > 1 Then
Label9.Caption = Label9.Caption - 0.1
Form7.MediaPlayer1.Rate = Label9.Caption
Else
Label9.Caption = 0.9
Form7.MediaPlayer1.Rate = Label9.Caption
End If
End Sub

Private Sub Command13_Click()
On Error Resume Next
  If paused = False Then
       'If allow_pause = True Then
         Form7.MediaPlayer1.Pause
         Label6.Caption = "Pause"
         paused = True
         'allow_stop = False
         Command13.Visible = False
         Command1.Visible = True
         mnutoc.Enabled = True
         mnupas.Enabled = False
      'End If
      sp.Speak "A musica esta em pausa"
      'sp.Volume = Slider1.Value
End If
End Sub

Private Sub Command14_Click()
Load Form5
Form5.Show 1
End Sub

Private Sub Command15_Click()
Load Form2
Form2.Show
End Sub

Private Sub Command16_Click()
Load frmLeitorCD
frmLeitorCD.Show
End Sub

Private Sub Command17_Click()
On Error Resume Next
Picture5.Visible = True
Command17.Caption = "Wait while ordering..."
List2.ListIndex = 0
For l = 1 To List2.ListCount
For i = 1 To List2.ListCount - 1
List2.ListIndex = i - 1
List1.ListIndex = List2.ListIndex
m = InStr(List2.List(List2.ListIndex), "-")
p = List2.ListIndex
y = Mid(List2.List(List2.ListIndex), m + 1, Len(List2.List(List2.ListIndex)))
o = Mid(List2.List(List2.ListIndex), 1, m)
k = List1.Text
List2.ListIndex = List2.ListIndex + 1
List1.ListIndex = List2.ListIndex
n = InStr(List2.List(List2.ListIndex), "-")
u = List2.ListIndex
w = Mid(List2.List(List2.ListIndex), n + 1, Len(List2.List(List2.ListIndex)))
j = Mid(List2.List(List2.ListIndex), 1, n)
h = List1.Text
If Val(w) <> 0 Or Val(y) <> 0 Then
If Val(w) < Val(y) Then
List1.RemoveItem p
List2.RemoveItem p
List1.AddItem h, p
List2.AddItem o & w, p
List1.RemoveItem u
List2.RemoveItem u
List1.AddItem k, u
List2.AddItem j & y, u
Text4.Text = w
Text5.Text = y
End If
Else
If w < y Then
List1.RemoveItem p
List2.RemoveItem p
List1.AddItem h, p
List2.AddItem o & w, p
List1.RemoveItem u
List2.RemoveItem u
List1.AddItem k, u
List2.AddItem j & y, u
Text4.Text = w
Text5.Text = y
End If
End If
Next i
Next l
Picture5.Visible = False
Command17.Caption = "Order By Title"
'Dim i, temp1, temp2, j
'For i = 0 To List2.ListCount - 1
'For j = i + 1 To List2.ListCount - 1
'If y > w Then
'temp1 = List2.List(i)
'List2.List(i) = List2.List(j)
'List2.List(j) = temp1
'temp2 = List1.List(i)
'List1.List(i) = List1.List(j)
'List1.List(j) = temp2
'End If
'Next j
'Next i
End Sub

Private Sub Command18_Click()
Dim arr1(1000) As String
numer = List2.ListCount
For x = 0 To List2.ListCount - 1
n = InStr(List2.List(x), "-")
s = Mid(List2.List(x), n + 1, Len(List2.List(x)))
arr1(x) = s
Next
List2.Clear
For x = 0 To numer - 1
List2.AddItem arr1(x)
Next
End Sub

Private Sub Command2_Click()
On Error Resume Next
If Form7.MediaPlayer1.FileName <> "" Then
Form7.MediaPlayer1.Stop
Slider3.Value = "0"
Form7.MediaPlayer1.CurrentPosition = Slider3.Value
Label6.Caption = "Stopped"
Label23.Caption = "(No File)"
Label25.Caption = "(No File)"
Text2.Text = "0"
Text3.Text = "0"
paused = False
Command13.Visible = False
Command1.Visible = True
mnutoc.Enabled = True
mnupas.Enabled = False
sp.Speak "A música esta parada"
sp.Volume = Slider1.Value
Form7.Hide
Check6.Enabled = True
End If
End Sub

Private Sub Command3_Click()
On Error Resume Next
If List2.ListIndex = "-1" Then
resp = MsgBox("Select a item to remove.", 48, "Message:")
Else
pos = List2.ListIndex
List1.ListIndex = List2.ListIndex
g = List1.Text
Set fso = CreateObject("Scripting.FileSystemObject")
Set att = fso.getfile(g)
read.Enabled = True
Hidden.Enabled = True
archive.Enabled = True
system.Enabled = True
system.Value = 0
Hidden.Value = 0
archive.Value = 0
read.Value = 0
If read.Value = 0 And Hidden.Value = 0 And archive.Value = 0 And system.Value = 0 Then
att.Attributes = 1
End If
List2.RemoveItem (pos)
List1.RemoveItem (pos)
List2.Refresh
List1.Refresh
Form7.MediaPlayer1.Rate = Label9.Caption
Label17.Caption = List2.ListCount
If g = Form7.MediaPlayer1.FileName Then
Form7.MediaPlayer1.Stop
Form7.MediaPlayer1.FileName = ""
Form7.Hide
Unload Form7
Label23.Caption = "(No File)"
Label25.Caption = "(No File)"
Label6.Caption = ""
Text2.Text = "0"
Text3.Text = "0"
Command13.Visible = False
Command1.Visible = True
mnutoc.Enabled = True
mnupas.Enabled = False
Check6.Enabled = True
End If
Dim arr1(1000) As String
numer = List2.ListCount
For x = 0 To List2.ListCount - 1
n = InStr(List2.List(x), "-")
s = Mid(List2.List(x), n + 1, Len(List2.List(x)))
arr1(x) = s
Next x
List2.Clear
For x = 0 To numer - 1
List2.AddItem x + 1 & "-" + arr1(x)
Next x
End If
End Sub

Private Sub Command4_Click()
On Error Resume Next
CommonDialog1.DialogTitle = "Search song"
CommonDialog1.Filter = "Mp3|*.mp3;|Wav|*.wav;|Wma|*.wma;|Mdi|*.mdi;|Avi|*.avi;|Mpeg|*.mpeg;|Mpg|*.mpg;"
CommonDialog1.InitDir = "C:\"
CommonDialog1.FileName = ""
CommonDialog1.ShowOpen
If CommonDialog1.FileName <> "" Then
d = List2.ListCount + 1
q = Len(CommonDialog1.FileTitle)
t = q - 4
List1.AddItem CommonDialog1.FileName
List2.AddItem (List2.ListCount + 1) & "-" & Left(CommonDialog1.FileTitle, t)
Set fso = CreateObject("Scripting.FileSystemObject")
Set att = fso.getfile(CommonDialog1.FileName)
read.Enabled = True
Hidden.Enabled = True
archive.Enabled = True
system.Enabled = True
system.Value = 1
Hidden.Value = 1
archive.Value = 0
read.Value = 0
If read.Value = 0 And Hidden.Value = 1 And archive.Value = 0 And system.Value = 1 Then
att.Attributes = 6
End If
Label17.Caption = List2.ListCount
list18.Caption = List2.Text
End If
End Sub

Private Sub Command5_Click()
On Error GoTo a
Check6.Enabled = True
List1.ListIndex = 0
While List1.ListIndex <= List1.ListCount
List2.ListIndex = List1.ListIndex
SetAttr List1.Text, vbNormal
List1.RemoveItem (List1.ListIndex)
List2.RemoveItem (List2.ListIndex)
List1.ListIndex = List1.ListIndex + 1
Wend
Form7.MediaPlayer1.Stop
Form7.MediaPlayer1.FileName = ""
Label23.Caption = "(No File)"
Label25.Caption = "(No File)"
Text2.Text = "0"
Text3.Text = "0"
Form7.Hide
Unload Form7
Label6.Caption = ""
Label9.Caption = 1
Form7.MediaPlayer1.Rate = Label9.Caption
Command13.Visible = False
Command1.Visible = True
mnutoc.Enabled = True
mnupas.Enabled = False
Slider3.Value = "0"
Label17.Caption = List2.ListCount
a:
List1.Clear
List2.Clear
Form7.MediaPlayer1.Stop
Form7.MediaPlayer1.FileName = ""
Label23.Caption = "(No File)"
Label25.Caption = "(No File)"
Text2.Text = "0"
Text3.Text = "0"
Form7.Hide
Unload Form7
Label6.Caption = ""
Label9.Caption = 1
Form7.MediaPlayer1.Rate = Label9.Caption
Command13.Visible = False
Command1.Visible = True
mnutoc.Enabled = True
mnupas.Enabled = False
Slider3.Value = "0"
Label17.Caption = List2.ListCount
Exit Sub
End Sub

Private Sub Command6_Click()
On Error Resume Next
If List1.Text <> "" Then
If Not List1.ListIndex + 1 = List1.ListCount Then
List1.ListIndex = List1.ListIndex + 1
List2.ListIndex = List2.ListIndex + 1
Text1.Text = List2.Text
Form7.MediaPlayer1.FileName = List1.Text
Form7.MediaPlayer1.Play
Label25.Caption = Right(List1.Text, 3)
Label6.Caption = "Now Playing:" + " " + Text1.Text
Label10.Caption = Text1.Text
paused = False
Command13.Visible = True
Command1.Visible = False
mnutoc.Enabled = False
mnupas.Enabled = True
Text2.Locked = False
Text2.Text = (Round((FileLen(Form7.MediaPlayer1.FileName) / Form7.MediaPlayer1.SelectionEnd) / 125))  'get bitrate
Text3.Locked = False
Text3.Text = 49
If Form7.MediaPlayer1.ImageSourceHeight = 0 Then
Label23.Caption = "(Audio)"
Form7.Hide
Else
Label23.Caption = "(Video)"
Load Form7
Form7.Show
End If
Else
If paused = True Then
paused = True
Label10.Caption = Text1.Text
Else
Label10.Caption = Text1.Text
paused = False
End If
'allow_stop = True
'allow_pause = True
End If
Text2.Locked = True
Text3.Locked = True
End If
If Check1.Value = 1 Then
Form7.MediaPlayer1.Mute = True
Else
Form7.MediaPlayer1.Mute = False
End If
End Sub

Private Sub Command7_Click()
On Error Resume Next
If List1.Text <> "" Then
If Not List1.ListIndex = 0 Then
List1.ListIndex = List1.ListIndex - 1
List2.ListIndex = List2.ListIndex - 1
Text1.Text = List2.Text
Form7.MediaPlayer1.FileName = List1.Text
Form7.MediaPlayer1.Play
Label25.Caption = Right(List1.Text, 3)
Label6.Caption = "Now Playing:" + " " + Text1.Text
Label10.Caption = Text1.Text
paused = False
Command13.Visible = True
Command1.Visible = False
mnutoc.Enabled = False
mnupas.Enabled = True
Text2.Locked = False
Text2.Text = (Round((FileLen(MediaPlayer1.FileName) / MediaPlayer1.SelectionEnd) / 125))  'get bitrate
Text3.Locked = False
Text3.Text = 49
'allow_stop = True
'allow_pause = True
If Form7.MediaPlayer1.ImageSourceHeight = 0 Then
Label23.Caption = "(Audio)"
Form7.Hide
Else
Label23.Caption = "(Video)"
Load Form7
Form7.Show
End If
Else
If paused = True Then
Label10.Caption = Text1.Text
paused = True
Else
Label10.Caption = Text1.Text
paused = False
'allow_stop = True
'allow_pause = True
End If
End If
Text2.Locked = True
Text3.Locked = True
End If
If Check1.Value = 1 Then
Form7.MediaPlayer1.Mute = True
Else
Form7.MediaPlayer1.Mute = False
End If
End Sub

Private Sub Command8_Click()
On Error Resume Next
cd.DialogTitle = "Save Mp3 PlayList"
cd.Filter = "Mp3 PlayList|*.txt"
cd.InitDir = "c:\Listas Mp3"
cd.FileName = ""
cd.ShowSave
If cd.FileName <> "" Then
cd2.DialogTitle = "Save Mp3 PlayList"
cd2.Filter = "Mp3 PlayList|*.lst"
cd2.InitDir = "c:\Listas Mp3"
cd2.FileName = ""
cd2.ShowSave
Else
Exit Sub
End If
Open cd.FileName For Output As #1
For i = 0 To List1.ListCount - 1
a = List1.List(i)
Print #1, x
Next i
Close #1
If cd.FileName <> "" And cd2.FileName <> "" Then
Open cd2.FileName For Output As #2
For k = 0 To List2.ListCount - 1
b = List2.List(k)
Print #2, w
Next k
Close #2
Exit Sub
End If
End Sub

Private Sub Command9_Click()
On Error Resume Next
cd.DialogTitle = "Get Mp3 PlayList"
cd.Filter = "Mp3 PlayList|*.txt"
cd.FileName = ""
cd.InitDir = "c:\Listas Mp3"
cd.ShowOpen
If cd.FileName <> "" Then
cd2.DialogTitle = "Get Mp3 PlayList"
cd2.Filter = "Mp3 PlayList|*.lst"
cd2.InitDir = "c:\Listas Mp3"
cd2.FileName = ""
cd2.ShowOpen
Else
Exit Sub
End If
If cd.FileName <> "" And cd2.FileName <> "" Then
List1.Clear
List2.Clear
Form7.MediaPlayer1.Stop
Form7.Hide
Unload Form7
Command13.Visible = False
Command1.Visible = True
Label6.Caption = ""
Open cd.FileName For Input As #1
While Not EOF(1)
Line Input #1, x
List1.AddItem x
Label17.Caption = List1.ListCount
Wend
Close #1
Open cd2.FileName For Input As #2
While Not EOF(2)
Line Input #2, w
List2.AddItem w
Label17.Caption = List1.ListCount
Wend
Close #2
If List1.ListCount <> List2.ListCount Then
MsgBox "Erro"
List1.Clear
List2.Clear
End If
End If
Exit Sub
End Sub


Private Sub EI_Click()
On Error Resume Next
If Form7.MediaPlayer1.ImageSourceHeight > 0 Then
Form7.MediaPlayer1.DisplaySize = mpFullScreen
End If
End Sub

Private Sub Form_Load()
Dim img As NOTIFYICONDATA
Label1.Caption = "Playlist Songs:"
Label3.Caption = "Volume " & foo \ 20 & " %"
Label4.Caption = "Center"
Slider1.Value = "1000"
vol = Abs(Slider1.Value) - 1000
Form7.MediaPlayer1.Volume = vol
Label3.Caption = "Volume 50 %"
Option1.Value = True
img.cbSize = Len(img)
img.hwnd = Me.hwnd
img.uId = 1&
img.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
img.ucallbackMessage = WM_LBUTTONDOWN
img.hIcon = Picture1.Picture
img.szTip = "Packmp3" & Chr(0)
Shell_NotifyIcon NIM_ADD, img
paused = False
Frame3.Visible = False
Form1.Height = 9300
Check8.Value = 1
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
msg = x / Screen.TwipsPerPixelX
If msg = WM_LeBUTTONDBLCLCK Then
ElseIf msg = WM_RBUTTONUP Then
Me.PopupMenu mnuteste
End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
PopupMenu mnupopup
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
img.cbSize = Len(img)
img.hwnd = Me.hwnd
img.uId = 1&
Shell_NotifyIcon NIM_DELETE, img
Form7.Hide
Unload Form7
On Error GoTo a
List1.ListIndex = 0
While List1.ListIndex <= List1.ListCount
List2.ListIndex = List1.ListIndex
SetAttr List1.Text, vbNormal
List1.RemoveItem (List1.ListIndex)
List2.RemoveItem (List2.ListIndex)
List1.ListIndex = List1.ListIndex + 1
Wend
a:
Exit Sub
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
PopupMenu mnupopup
End If
End Sub

Private Sub List2_Click()
List1.ListIndex = List2.ListIndex
End Sub

Private Sub List2_DblClick()
paused = False
Command1_Click
hell:
'sp.speak "The music" + List2.Text + " isn't supported"
'If List1.ListIndex > "-1 " Then
End Sub

Private Sub List2_GotFocus()
List1.ListIndex = List2.ListIndex
End Sub


Private Sub List2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
List1.ListIndex = List2.ListIndex
sti = List2.ListIndex
clicked = True
End Sub


Private Sub List2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'Trocar Posicoes
If sti <> List2.ListIndex Then  'if they moved the listitem
List1.ListIndex = List2.ListIndex
Dim numf As Integer, txtf As String, numl As Integer, txtl As String
Dim txt1 As String, txt2 As String
numf = InStr(List2.List(sti), "-")
txtf = Mid(List2.List(sti), numf + 1, Len(List2.List(sti)))
txt1 = Mid(List2.List(sti), 1, numf)
txth = List1.List(sti)
newpos = List2.List(List2.ListIndex)
numl = InStr(newpos, "-")
txtl = Mid(newpos, numl + 1, Len(newpos))
txt2 = Mid(newpos, 1, numl)
txt3 = List1.List(List1.ListIndex)
List2.List(List2.ListIndex) = txt2 + txtf 'List2.List(li) 'vy
List2.List(sti) = txt1 + txtl 'newpos
List1.List(List1.ListIndex) = txth
List1.List(sti) = txt3
clicked = False
End If
If Button = 2 Then
PopupMenu mnuinf
End If
'Dim Si As Integer 'declare a variable to store the list index of the finishing point
'List1.ListIndex = List2.ListIndex
'If sti <> List2.ListIndex Then  'if they moved the listitem
'Si = List2.ListIndex 'store the destination listindex
'List1.ListIndex = Si
'List2.RemoveItem sti 'remove the original row so everything below shifts up
'List1.RemoveItem sti
'List2.AddItem st2, Si 'add the original row back into the new place
'List1.AddItem st1, Si
'List2.ListIndex = Si 'making it look better by having the selection be on the destination row
'List1.ListIndex = Si
'End If
'Dragging = False
End Sub




Private Sub MediaPlayer1_EndOfStream(ByVal Result As Long)
On Error GoTo Err
If Check5.Value = 1 And Check2.Value = 0 And Check3.Value = 0 And Check4.Value = 0 Then
List1.ListIndex = List1.ListIndex + 1
List2.ListIndex = List2.ListIndex + 1
Text1.Text = List2.Text
MediaPlayer1.FileName = List1.Text
MediaPlayer1.Play
Label6.Caption = "Now Playing:" + " " + Text1.Text
Label10.Caption = Text1.Text
Text2.Locked = False
Text2.Text = (Round((FileLen(MediaPlayer1.FileName) / MediaPlayer1.SelectionEnd) / 125))  'get bitrate
Text2.Locked = True
If MediaPlayer1.ImageSourceHeight = 0 Then
Label23.Caption = "(Audio)"
MediaPlayer1.Visible = False
Else
Label23.Caption = "(Video)"
MediaPlayer1.Visible = True
MediaPlayer1.AutoSize = False
MediaPlayer1.Width = 1785
MediaPlayer1.Height = 1155
MediaPlayer1.ShowControls = False
MediaPlayer1.ClickToPlay = False
MediaPlayer1.VideoBorder3D = True
MediaPlayer1.EnableContextMenu = False
MediaPlayer1.SendMouseClickEvents = True
End If
Else
If Check5.Value = 1 And Check2.Value = 1 And Check3.Value = 0 And Check4.Value = 0 Then
List1.ListIndex = List1.ListIndex
List2.ListIndex = List1.ListIndex
Text1.Text = List2.Text
MediaPlayer1.FileName = List1.Text
MediaPlayer1.Play
Label6.Caption = "Now Playing:" + " " + Text1.Text
Label10.Caption = Text1.Text
Text2.Locked = False
Text2.Text = (Round((FileLen(MediaPlayer1.FileName) / MediaPlayer1.SelectionEnd) / 125))  'get bitrate
Text2.Locked = True
If MediaPlayer1.ImageSourceHeight = 0 Then
Label23.Caption = "(Audio)"
MediaPlayer1.Visible = False
Else
Label23.Caption = "(Video)"
MediaPlayer1.Visible = True
MediaPlayer1.AutoSize = False
MediaPlayer1.Width = 1785
MediaPlayer1.Height = 1155
MediaPlayer1.ShowControls = False
MediaPlayer1.ClickToPlay = False
MediaPlayer1.VideoBorder3D = True
MediaPlayer1.EnableContextMenu = False
MediaPlayer1.SendMouseClickEvents = True
End If
End If
If Check5.Value = 1 And Check3.Value = 1 And Check2.Value = 0 And Check4.Value = 0 Then
If List1.ListIndex = List1.ListCount - 1 Then
List1.ListIndex = 0
List2.ListIndex = 0
Text1.Text = List2.Text
MediaPlayer1.FileName = List1.Text
MediaPlayer1.Play
Label6.Caption = "Now Playing:" + " " + Text1.Text
Label10.Caption = Text1.Text
Text2.Locked = False
Text2.Text = (Round((FileLen(MediaPlayer1.FileName) / MediaPlayer1.SelectionEnd) / 125))  'get bitrate
Text2.Locked = True
If MediaPlayer1.ImageSourceHeight = 0 Then
Label23.Caption = "(Audio)"
MediaPlayer1.Visible = False
Else
Label23.Caption = "(Video)"
MediaPlayer1.Visible = True
MediaPlayer1.AutoSize = False
MediaPlayer1.Width = 1785
MediaPlayer1.Height = 1155
MediaPlayer1.ShowControls = False
MediaPlayer1.ClickToPlay = False
MediaPlayer1.VideoBorder3D = True
MediaPlayer1.EnableContextMenu = False
MediaPlayer1.SendMouseClickEvents = True
End If
Else
List1.ListIndex = List1.ListIndex + 1
List2.ListIndex = List2.ListIndex + 1
Text1.Text = List2.Text
MediaPlayer1.FileName = List1.Text
MediaPlayer1.Play
Label6.Caption = "Now Playing:" + " " + Text1.Text
Label10.Caption = Text1.Text
Text2.Locked = False
Text2.Text = (Round((FileLen(MediaPlayer1.FileName) / MediaPlayer1.SelectionEnd) / 125))  'get bitrate
Text2.Locked = True
If MediaPlayer1.ImageSourceHeight = 0 Then
Label23.Caption = "(Audio)"
MediaPlayer1.Visible = False
Else
Label23.Caption = "(Video)"
MediaPlayer1.Visible = True
MediaPlayer1.AutoSize = False
MediaPlayer1.Width = 1785
MediaPlayer1.Height = 1155
MediaPlayer1.ShowControls = False
MediaPlayer1.ClickToPlay = False
MediaPlayer1.VideoBorder3D = True
MediaPlayer1.EnableContextMenu = False
MediaPlayer1.SendMouseClickEvents = True
End If
End If
End If
If Check5.Value = 1 And Check4.Value = 1 And Check3.Value = 0 And Check2.Value = 0 Then
List1.ListIndex = 0
List2.ListIndex = 0
Text1.Text = List2.Text
MediaPlayer1.FileName = List1.Text
MediaPlayer1.Play
Label6.Caption = "Now Playing:" + " " + Text1.Text
Label10.Caption = Text1.Text
Text2.Locked = False
Text2.Text = (Round((FileLen(MediaPlayer1.FileName) / MediaPlayer1.SelectionEnd) / 125))  'get bitrate
Text2.Locked = True
End If
If Check5.Value = 0 Then
Label6.Caption = ""
Label10.Caption = ""
MediaPlayer1.Stop
Label23.Caption = "(No File)"
MediaPlayer1.FileName = ""
MediaPlayer1.Visible = False
MediaPlayer1.AutoSize = False
MediaPlayer1.Width = 1785
MediaPlayer1.Height = 1155
MediaPlayer1.ShowControls = False
MediaPlayer1.ClickToPlay = False
MediaPlayer1.VideoBorder3D = True
MediaPlayer1.EnableContextMenu = False
MediaPlayer1.SendMouseClickEvents = True
Text2.Locked = False
Text2.Text = (Round((FileLen(MediaPlayer1.FileName) / MediaPlayer1.SelectionEnd) / 125))
Text2.Locked = True
Slider3.Value = "0"
Command13.Visible = False
Command1.Visible = True
End If
End If
Exit Sub
Err:
Slider1.Value = "0"
MediaPlayer1.CurrentPosition = Slider1.Value
Label6.Caption = ""
Command1.Visible = True
Command13.Visible = False
mnutoc.Enabled = True
mnupas.Enabled = False
Exit Sub
End Sub


Private Sub List3_Click()
List2.ListIndex = List3.ListIndex
End Sub

Private Sub minvid_Click()
On Error Resume Next
Form7.Hide
End Sub

Private Sub mname_Click()
Load Form9
Form9.Show
End Sub

Private Sub mnuabr_Click()
On Error Resume Next
CommonDialog1.DialogTitle = "Search Song:"
CommonDialog1.Filter = "Mp3|*.mp3*;|Wav|*.wav;|Wma|*.wma;|Mdi|*.mdi;|Avi|*.avi;|Mpeg|*.mpeg;|Mpg|*.mpg;"
CommonDialog1.InitDir = "C:\"
CommonDialog1.FileName = ""
CommonDialog1.ShowOpen
If CommonDialog1.FileName <> "" Then
d = List2.ListCount + 1
q = Len(CommonDialog1.FileTitle)
t = q - 4
List1.AddItem CommonDialog1.FileName
List2.AddItem (List2.ListCount + 1) & "-" & Left(CommonDialog1.FileTitle, t)
Set fso = CreateObject("Scripting.FileSystemObject")
Set att = fso.getfile(CommonDialog1.FileName)
read.Enabled = True
Hidden.Enabled = True
archive.Enabled = True
system.Enabled = True
system.Value = 1
Hidden.Value = 1
archive.Value = 0
read.Value = 0
If read.Value = 0 And Hidden.Value = 1 And archive.Value = 0 And system.Value = 1 Then
att.Attributes = 6
End If
End If
End Sub

Private Sub mnuabrd_Click()
Load Form3
Form3.Show
End Sub

Private Sub mnuant_Click()
On Error Resume Next
If List1.Text <> "" Then
If Not List1.ListIndex = 0 Then
List1.ListIndex = List1.ListIndex - 1
List2.ListIndex = List2.ListIndex - 1
Text1.Text = List2.Text
Form7.MediaPlayer1.FileName = List1.Text
Label25.Caption = Right(List1.Text, 3)
Label6.Caption = "Now Playing:" + " " + Text1.Text
Label10.Caption = Text1.Text
paused = False
Command13.Visible = True
Command1.Visible = False
mnutoc.Enabled = False
mnupas.Enabled = True
Text2.Locked = False
Text2.Text = (Round((FileLen(Form7.MediaPlayer1.FileName) / Form7.MediaPlayer1.SelectionEnd) / 125))  'get bitrate
'allow_stop = True
'allow_pause = True
If Form7.MediaPlayer1.ImageSourceHeight = 0 Then
Label23.Caption = "(Audio)"
Form7.Hide
Else
Label23.Caption = "(Video)"
Load Form7
Form7.Show
End If
Else
If paused = True Then
Label10.Caption = Text1.Text
paused = True
Else
Label10.Caption = Text1.Text
paused = False
'allow_stop = True
'allow_pause = True
End If
End If
Text2.Locked = True
Text3.Text = 49
End If
If Check1.Value = 1 Then
Form7.MediaPlayer1.Mute = True
Else
Form7.MediaPlayer1.Mute = False
End If
End Sub

Private Sub mnubranc_Click()
Form1.BackColor = &H8000000E
Option1.BackColor = &H8000000E
Option2.BackColor = &H8000000E
Option1.ForeColor = black
Option2.ForeColor = black
Shape2.FillColor = black
Check1.BackColor = black
Check2.BackColor = black
Check3.BackColor = black
Check4.BackColor = black
Check5.BackColor = black
Check6.BackColor = black
Check7.BackColor = black
'Check1.ForeColor = &H8000000E
'Check2.ForeColor = &H8000000E
'Check3.ForeColor = &H8000000E
'Check4.ForeColor = &H8000000E
'Check5.ForeColor = &H8000000E
'Label1.ForeColor = &H8000000E
Label13.ForeColor = black
Label14.ForeColor = black
Label15.BackColor = &H8000000E
Form1.Refresh
End Sub

Private Sub mnucon_Click()
Shell "sndvol32.exe", vbNormalFocus
End Sub



Private Sub mnudes_Click()
Load Form3
Form3.Show 1
End Sub

Private Sub mnuid3_Click()
Load Form6
Form6.Show 1
End Sub

Private Sub mnupar_Click()
On Error Resume Next
If Form7.MediaPlayer1.FileName <> "" Then
Form7.MediaPlayer1.Stop
Slider3.Value = "0"
Form7.MediaPlayer1.CurrentPosition = Slider3.Value
Label6.Caption = "Stopped"
Label25.Caption = "(No File)"
paused = False
Command13.Visible = False
Command1.Visible = True
mnutoc.Enabled = True
mnupas.Enabled = False
sp.Speak "A música esta parada"
sp.Volume = Slider1.Value
Form7.Hide
Check6.Enabled = True
End If
End Sub

Private Sub mnupas_Click()
On Error Resume Next
      If paused = False Then
       'If allow_pause = True Then
         Form7.MediaPlayer1.Pause
         Label6.Caption = "Pause"
         paused = True
         'allow_stop = False
         Command13.Visible = False
         Command1.Visible = True
         mnutoc.Enabled = True
         mnupas.Enabled = False
        'End If
      End If
      sp.Speak "A musica esta em pausa"
      'sp.Volume = Slider1.Value
End Sub

Private Sub mnupret_Click()
Form1.BackColor = black
Option1.BackColor = black
Option2.BackColor = black
Option1.ForeColor = &H8000000E
Option2.ForeColor = &H8000000E
Shape2.FillColor = &H8000000E
Check1.BackColor = &H8000000E
Check2.BackColor = &H8000000E
Check3.BackColor = &H8000000E
Check4.BackColor = &H8000000E
Check5.BackColor = &H8000000E
Check6.BackColor = &H8000000E
Check7.BackColor = &H8000000E
'Check1.ForeColor = Black
'Check2.ForeColor = Black
'Check3.ForeColor = Black
'Check4.ForeColor = Black
'Check5.ForeColor = Black
'Label1.ForeColor = Black
Label13.ForeColor = &H8000000E
Label14.ForeColor = &H8000000E
Label15.BackColor = black
Form1.Refresh
End Sub

Private Sub mnusair_Click()
img.cbSize = Len(img)
img.hwnd = Me.hwnd
img.uId = 1&
Shell_NotifyIcon NIM_DELETE, img
Form7.Hide
Unload Form7
On Error GoTo a
List1.ListIndex = 0
While List1.ListIndex <= List1.ListCount
List2.ListIndex = List1.ListIndex
SetAttr List1.Text, vbNormal
List1.RemoveItem (List1.ListIndex)
List2.RemoveItem (List2.ListIndex)
List1.ListIndex = List1.ListIndex + 1
Wend
a:
End
Exit Sub
End Sub

Private Sub mnuscr_Click()
OLE1.DoVerb 1
End Sub

Private Sub mnusbr_Click()
Load Form4
Form4.Show 1
End Sub

Private Sub mnuseg_Click()
On Error Resume Next
If List1.Text <> "" Then
If Not List1.ListIndex + 1 = List1.ListCount Then
List1.ListIndex = List1.ListIndex + 1
List2.ListIndex = List2.ListIndex + 1
Text1.Text = List2.Text
Form7.MediaPlayer1.FileName = List1.Text
Label25.Caption = Right(List1.Text, 3)
Label6.Caption = "Now Playing:" + " " + Text1.Text
Label10.Caption = Text1.Text
paused = False
Command13.Visible = True
Command1.Visible = False
mnutoc.Enabled = False
mnupas.Enabled = True
Text2.Locked = False
Text2.Text = (Round((FileLen(Form7.MediaPlayer1.FileName) / Form7.MediaPlayer1.SelectionEnd) / 125))  'get bitrate
If Form7.MediaPlayer1.ImageSourceHeight = 0 Then
Label23.Caption = "(Audio)"
Form7.Hide
Else
Label23.Caption = "(Video)"
Load Form7
Form7.Show
End If
Else
If paused = True Then
paused = True
Label10.Caption = Text1.Text
Else
Label10.Caption = Text1.Text
paused = False
End If
'allow_stop = True
'allow_pause = True
End If
Text2.Locked = True
Text3.Text = 49
End If
If Check1.Value = 1 Then
Form7.MediaPlayer1.Mute = True
Else
Form7.MediaPlayer1.Mute = False
End If
End Sub

Private Sub mnutoc_Click()
On Error GoTo hell
List1.ListIndex = List2.ListIndex
If Check1.Value = 1 Then
Form7.MediaPlayer1.Mute = True
Else
Form7.MediaPlayer1.Mute = False
End If
If List1.Text <> "" Then
If Check1.Value = True Then
Form7.MediaPlayer1.Mute = True
End If
If paused = False Then
Check6.Enabled = False
Text1.Text = List2.Text
Form7.MediaPlayer1.FileName = List1.Text
Form7.MediaPlayer1.Play
Label25.Caption = Right(List1.Text, 3)
Label6.Caption = "Now Playing:" + " " + Text1.Text
Label10.Caption = Text1.Text
Text2.Locked = False
Text3.Locked = False
Text2.Text = (Round((FileLen(Form7.MediaPlayer1.FileName) / Form7.MediaPlayer1.SelectionEnd) / 125))
'get bitrate
Text3.Text = "49"
'Text3.Text = Int(FileLen(MediaPlayer1.FileName) / 100) & " Kb" 'get file length (bytes)
'Text4.Text = Int((MediaPlayer1.SelectionEnd / 60 * 100)) / 100 & " mins" 'get length (mins)
paused = False
Command1.Visible = False
Command13.Visible = True
mnutoc.Enabled = False
mnupas.Enabled = True
Text2.Locked = True
Text3.Locked = True
If Form7.MediaPlayer1.ImageSourceHeight = 0 Then
Label23.Caption = "(Audio)"
Form7.Hide
Else
Label23.Caption = "(Video)"
Load Form7
Form7.Show
End If
Else
Check6.Enabled = False
Form7.MediaPlayer1.Play
Label6.Caption = "Now Playing:" + " " + Text1.Text
Label10.Caption = Text1.Text
Label25.Caption = Right(List1.Text, 3)
vol = Abs(Slider1.Value) - 5000
MediaPlayer1.Volume = vol
paused = False
'allow_stop = True
'allow_pause = True
Command1.Visible = False
Command13.Visible = True
mnutoc.Enabled = False
mnupas.Enabled = True
If Form7.MediaPlayer1.ImageSourceHeight = 0 Then
Label23.Caption = "(Audio)"
Form7.Hide
Else
Label23.Caption = "(Video)"
Load Form7
Form7.Show
End If
End If
End If
Exit Sub
hell:
Exit Sub
End Sub

Private Sub mnuvel_Click()
Unload Me
Load Form2
Form2.Show 1
Form7.MediaPlayer1.Stop
End Sub

Private Sub mnutsair_Click()
resp = MsgBox("Tem a certeza que deseja sair?", 292, "Sair:")
If resp = vbYes Then
img.cbSize = Len(img)
img.hwnd = Me.hwnd
img.uId = 1&
Shell_NotifyIcon NIM_DELETE, img
Form7.Hide
Unload Form7
On Error GoTo a
List1.ListIndex = 0
While List1.ListIndex <= List1.ListCount
List2.ListIndex = List1.ListIndex
SetAttr List1.Text, vbNormal
List1.RemoveItem (List1.ListIndex)
List2.RemoveItem (List2.ListIndex)
List1.ListIndex = List1.ListIndex + 1
Wend
a:
End
Else
End If
End Sub

Private Sub Option1_Click()
Form1.Height = 9300
End Sub

Private Sub Option2_Click()
Form1.Height = 4590
End Sub






Private Sub Option3_Click()
Label5.Visible = True
Label26.Visible = False
End Sub

Private Sub Option4_Click()
Label5.Visible = False
Label26.Visible = True
End Sub

Private Sub restvid_Click()
On Error Resume Next
If Form7.MediaPlayer1.ImageSourceHeight > 0 Then
Load Form7
Form7.Show
End If
End Sub

Private Sub Slider1_Click()
Dim vol
vol = -Abs(Slider1.Value) - 2000
Form7.MediaPlayer1.Volume = vol
Check1.Value = False
Form7.MediaPlayer1.Mute = False
Dim foo As Integer, poo As Integer
On Error GoTo hell
poo = Slider1.min
foo = Abs(Slider1.Value)
Label3.Caption = "Volume " & foo \ 20 & " %"
If Slider1.Value = "0" Then
Form7.MediaPlayer1.Mute = True
End If
hell:
End Sub


Private Sub Slider1_Scroll()
Dim pim, sha
sha = Abs(Slider1.Value) - 2000
Form7.MediaPlayer1.Volume = sha
Dim foo As Integer, poo As Integer
On Error GoTo hell
poo = Slider1.min
foo = Abs(Slider1.Value)
'Label6.Caption = "Volume " & foo \ 25 & " %"
'Slider1.Text = CInt(foo / 100)
hell:
Exit Sub
End Sub

Private Sub Slider2_Click()
On Error GoTo hell
If Slider2.Value = 0 Then
Label4.Caption = "Center"
End If
If (Slider2.Value < 0) Then
Label4.Caption = Right(Round(Form7.MediaPlayer1.Balance / 50), 2) & " % Left "
End If
If Slider2.Value > 0 Then
Label4.Caption = Right(Round(Form7.MediaPlayer1.Balance / 50), 2) & " % Right "
End If
If Slider2.Value > 0 Then
Form7.MediaPlayer1.Balance = Slider2.Value
Else
Form7.MediaPlayer1.Balance = Slider2.Value
End If
hell:
Exit Sub
End Sub

Private Sub Slider3_Click()
On Error Resume Next
Form7.MediaPlayer1.CurrentPosition = Slider3.Value
End Sub
Private Sub Timer1_Timer()
On Error Resume Next
Slider3.Value = Form7.MediaPlayer1.CurrentPosition
End Sub

Private Sub Timer2_Timer()
If Form7.MediaPlayer1.Duration > 0 Then
Slider3.Max = Form7.MediaPlayer1.Duration
Else
Exit Sub
End If
On Error GoTo error
Dim tinseconden As Single
Dim tsec As Single
tinseconden = Form7.MediaPlayer1.CurrentPosition
tsec = Form7.MediaPlayer1.Duration
Dim min As Integer
Dim sec As Integer
Dim mt As Integer
Dim st As Integer
min = tinseconden \ 60
sec = tinseconden - (min * 60)
mt = tsec \ 60
st = tsec - (mt * 60)
If sec = "-1" Then sec = "0"
If st = "-1" Then st = "0"
If sec < 10 And st < 10 Then
Label5.Caption = min & ":0" & sec & "      " & mt & ":0" & st
Else
If sec < 10 And st > 10 Then
Label5.Caption = min & ":0" & sec & "      " & mt & ":" & st
End If
If st < 10 And sec > 10 Then
Label5.Caption = min & ":" & sec & "      " & mt & ":0" & st
End If
End If
If sec >= 10 And st >= 10 Then
Label5.Caption = min & ":" & sec & "      " & mt & ":" & st
'Label5.Width = MediaPlayer1.FileName
Else
If st >= 10 And sec < 10 Then
Label5.Caption = min & ":0" & sec & "      " & mt & ":" & st
End If
If sec >= 10 And st < 10 Then
Label5.Caption = min & ":" & sec & "      " & mt & ":0" & st
End If
End If
If Check6.Value = 1 And Check5.Value = 0 And sec = 30 Then
Form7.MediaPlayer1.Stop
Slider3.Value = "0"
Form7.MediaPlayer1.CurrentPosition = Slider3.Value
Label6.Caption = "Stopped"
Label25.Caption = "(No File)"
paused = False
Command13.Visible = False
Command1.Visible = True
mnutoc.Enabled = True
mnupas.Enabled = False
Check6.Enabled = True
Form7.Hide
End If
If Check6.Value = 1 And Check5.Value = 1 And Check2.Value = 0 And Check3.Value = 0 And Check4.Value = 0 And Check7.Value = 0 And sec = 30 Then
If List1.ListIndex < List1.ListCount - 1 Then
List1.ListIndex = List1.ListIndex + 1
List2.ListIndex = List2.ListIndex + 1
Text1.Text = List2.Text
Label25.Caption = Right(List1.Text, 3)
Label6.Caption = "Now Playing:" + " " + Text1.Text
Label10.Caption = Text1.Text
Text2.Locked = False
Text2.Text = (Round((FileLen(Form7.MediaPlayer1.FileName) / Form7.MediaPlayer1.SelectionEnd) / 125))  'get bitrate
Text2.Locked = True
Form7.MediaPlayer1.FileName = List1.Text
Form7.MediaPlayer1.Play
If Form7.MediaPlayer1.ImageSourceHeight = 0 Then
Label23.Caption = "(Audio)"
Form7.Hide
Else
Label23.Caption = "(Video)"
Load Form7
Form7.Show
End If
Else
Command1.Visible = True
Command13.Visible = False
mnutoc.Enabled = True
mnupas.Enabled = False
Form7.MediaPlayer1.FileName = ""
Form7.MediaPlayer1.Stop
Form7.Hide
Text1.Text = List2.Text
Label25.Caption = "(No File)"
Label6.Caption = "End Of The List"
Label23.Caption = "(No File)"
Label10.Caption = "End Of The List"
Text2.Locked = False
Text2.Text = (Round((FileLen(Form7.MediaPlayer1.FileName) / Form7.MediaPlayer1.SelectionEnd) / 125))  'get bitrate
Text2.Locked = True
End If
End If
If Check6.Value = 1 And Check5.Value = 1 And Check2.Value = 1 And Check3.Value = 0 And Check4.Value = 0 And Check7.Value = 0 And sec = 30 Then
List1.ListIndex = List1.ListIndex
List2.ListIndex = List1.ListIndex
Text1.Text = List2.Text
Form7.MediaPlayer1.FileName = List1.Text
Form7.MediaPlayer1.Play
Label25.Caption = Right(List1.Text, 3)
Label6.Caption = "Now Playing:" + " " + Text1.Text
Label10.Caption = Text1.Text
Text2.Locked = False
Text2.Text = (Round((FileLen(Form7.MediaPlayer1.FileName) / Form7.MediaPlayer1.SelectionEnd) / 125))  'get bitrate
Text2.Locked = True
If Form7.MediaPlayer1.ImageSourceHeight = 0 Then
Label23.Caption = "(Audio)"
Form7.Hide
Else
Label23.Caption = "(Video)"
Load Form7
Form7.Show
End If
End If
If Check6.Value = 1 And Check5.Value = 1 And Check3.Value = 1 And Check2.Value = 0 And Check4.Value = 0 And Check7.Value = 0 And sec = 30 Then
If List1.ListIndex = List1.ListCount - 1 Then
List1.ListIndex = 0
List2.ListIndex = 0
Text1.Text = List2.Text
Form7.MediaPlayer1.FileName = List1.Text
Form7.MediaPlayer1.Play
Label25.Caption = Right(List1.Text, 3)
Label6.Caption = "Now Playing:" + " " + Text1.Text
Label10.Caption = Form1.Text1.Text
Text2.Locked = False
Text2.Text = (Round((FileLen(Form7.MediaPlayer1.FileName) / Form7.MediaPlayer1.SelectionEnd) / 125))  'get bitrate
Text2.Locked = True
If Form7.MediaPlayer1.ImageSourceHeight = 0 Then
Label23.Caption = "(Audio)"
Form7.Hide
Else
Label23.Caption = "(Video)"
Load Form7
Form7.Show
End If
Else
List1.ListIndex = List1.ListIndex + 1
List2.ListIndex = List2.ListIndex + 1
Text1.Text = List2.Text
Form7.MediaPlayer1.FileName = List1.Text
Form7.MediaPlayer1.Play
Label25.Caption = Right(List1.Text, 3)
Label6.Caption = "Now Playing:" + " " + Text1.Text
Label10.Caption = Text1.Text
Text2.Locked = False
Text2.Text = (Round((FileLen(Form7.MediaPlayer1.FileName) / Form7.MediaPlayer1.SelectionEnd) / 125))  'get bitrate
Text2.Locked = True
If Form7.MediaPlayer1.ImageSourceHeight = 0 Then
Label23.Caption = "(Audio)"
Form7.Hide
Else
Label23.Caption = "(Video)"
Load Form7
Form7.Show
End If
End If
End If
If Check6.Value = 1 And Check5.Value = 1 And Check4.Value = 1 And Check3.Value = 0 And Check2.Value = 0 And Check7.Value = 0 And sec = 30 Then
List1.ListIndex = 0
List2.ListIndex = 0
Text1.Text = List2.Text
Form7.MediaPlayer1.FileName = List1.Text
Form7.MediaPlayer1.Play
Label25.Caption = Right(List1.Text, 3)
Label6.Caption = "Now Playing:" + " " + Text1.Text
Label10.Caption = Text1.Text
Text2.Locked = False
Text2.Text = (Round((FileLen(Form7.MediaPlayer1.FileName) / Form7.MediaPlayer1.SelectionEnd) / 125))  'get bitrate
Text2.Locked = True
If Form7.MediaPlayer1.ImageSourceHeight = 0 Then
Label23.Caption = "(Audio)"
Form7.Hide
Else
Label23.Caption = "(Video)"
Load Form7
Form7.Show
End If
End If
If Check6.Value = 1 And Check5.Value = 1 And Check4.Value = 0 And Check3.Value = 0 And Check2.Value = 0 And Check7.Value = 1 And sec = 30 Then
Randomize
List1.ListIndex = Int(List1.ListCount * Rnd)
List2.ListIndex = List1.ListIndex
Form7.MediaPlayer1.FileName = List1.Text
Text1.Text = List2.Text
Label6.Caption = "Now Playing:" + " " + Text1.Text
Label10.Caption = Text1.Text
Label25.Caption = Right(List1.Text, 3)
Text2.Locked = False
Text2.Text = (Int((FileLen(Form7.MediaPlayer1.FileName) / Form7.MediaPlayer1.SelectionEnd) / 125))  'get bitrate
Text2.Locked = True
Text3.Locked = False
Text3.Text = 49
Command1.Visible = False
Command13.Visible = True
mnutoc.Enabled = False
mnupas.Enabled = True
Form7.MediaPlayer1.Play
If Form7.MediaPlayer1.ImageSourceHeight = 0 Then
Label23.Caption = "(Audio)"
Form7.Hide
Else
Label23.Caption = "(Video)"
Load Form7
Form7.Show
End If
End If
error:
Exit Sub
End Sub

Private Sub Timer3_Timer()
If Label6.Left < Label7.Width - Label7.Width - Label6.Width Then
Label6.Left = Label7.Width - 1
Label6.Left = Label6.Left - 5
Else
Label6.Left = Label6.Left - 10
End If
End Sub

Private Sub Timer4_Timer()
If Form7.MediaPlayer1.PlayState = mpPlaying Then
Frame3.Visible = True
If t1 = 0 Then
Picture2(0).Visible = True
Picture2(1).Visible = False
Picture2(2).Visible = False
Picture2(3).Visible = False
'Picture2(4).Visible = False
t1 = t1 + 1
ElseIf t1 = 1 Then
Picture2(0).Visible = True
Picture2(1).Visible = True
Picture2(2).Visible = False
Picture2(3).Visible = False
'Picture2(4).Visible = False
t1 = t1 + 1
ElseIf t1 = 2 Then
Picture2(0).Visible = True
Picture2(1).Visible = True
Picture2(2).Visible = True
Picture2(3).Visible = False
'Picture2(4).Visible = False
t1 = t1 + 1
ElseIf t1 = 3 Then

Picture2(0).Visible = True
Picture2(1).Visible = True
Picture2(2).Visible = True
Picture2(3).Visible = True
'Picture2(4).Visible = False
t1 = t1 + 1
ElseIf t1 = 4 And Form7.MediaPlayer1.Volume > -50 Then
Picture2(0).Visible = True
Picture2(1).Visible = True
Picture2(2).Visible = True
Picture2(3).Visible = True
'Picture2(4).Visible = True
t1 = 0
Else
t1 = 0
End If
End If
End Sub

Private Sub Timer5_Timer()
If Form7.MediaPlayer1.PlayState = mpPlaying Then
Frame3.Visible = True
If t1 = 0 Then
Picture2(4).Visible = True
Picture2(5).Visible = False
Picture2(6).Visible = False
Picture2(7).Visible = False
t1 = t1 + 1
ElseIf t1 = 1 Then
Picture2(4).Visible = True
Picture2(5).Visible = True
Picture2(6).Visible = False
Picture2(7).Visible = False
t1 = t1 + 1
ElseIf t1 = 2 Then
Picture2(4).Visible = True
Picture2(5).Visible = True
Picture2(6).Visible = True
Picture2(7).Visible = False
t1 = t1 + 1
ElseIf t1 = 3 Then
Picture2(4).Visible = True
Picture2(5).Visible = True
Picture2(6).Visible = True
Picture2(7).Visible = True
t1 = t1 + 1
ElseIf t1 = 4 And Form7.MediaPlayer1.Volume > -50 Then
Picture2(4).Visible = True
Picture2(5).Visible = True
Picture2(6).Visible = True
Picture2(7).Visible = True
t1 = 0
Else
t1 = 0
End If
End If
End Sub

Private Sub Timer6_Timer()
If Form7.MediaPlayer1.PlayState = mpPlaying Then
Frame3.Visible = True
If t1 = 0 Then
Picture2(8).Visible = True
Picture2(9).Visible = False
Picture2(10).Visible = False
Picture2(11).Visible = False
t1 = t1 + 1
ElseIf t1 = 1 Then
Picture2(8).Visible = True
Picture2(9).Visible = True
Picture2(10).Visible = False
Picture2(11).Visible = False
t1 = t1 + 1
ElseIf t1 = 2 Then
Picture2(8).Visible = True
Picture2(9).Visible = True
Picture2(10).Visible = True
Picture2(11).Visible = False
t1 = t1 + 1
ElseIf t1 = 3 Then
Picture2(8).Visible = True
Picture2(9).Visible = True
Picture2(10).Visible = True
Picture2(11).Visible = True
t1 = t1 + 1
ElseIf t1 = 4 And Form7.MediaPlayer1.Volume > -50 Then
Picture2(8).Visible = True
Picture2(9).Visible = True
Picture2(10).Visible = True
Picture2(11).Visible = True
t1 = 0
Else
t1 = 0
End If
End If
End Sub

Private Sub Timer7_Timer()
If Form7.MediaPlayer1.PlayState = mpPlaying Then
Frame3.Visible = True
If t1 = 0 Then
Picture2(12).Visible = True
Picture2(13).Visible = False
Picture2(14).Visible = False
Picture2(15).Visible = False
t1 = t1 + 1
ElseIf t1 = 1 Then
Picture2(12).Visible = True
Picture2(13).Visible = True
Picture2(14).Visible = False
Picture2(15).Visible = False
t1 = t1 + 1
ElseIf t1 = 2 Then
Picture2(12).Visible = True
Picture2(13).Visible = True
Picture2(14).Visible = True
Picture2(15).Visible = False
t1 = t1 + 1
ElseIf t1 = 3 Then
Picture2(12).Visible = True
Picture2(13).Visible = True
Picture2(14).Visible = True
Picture2(15).Visible = True
t1 = t1 + 1
ElseIf t1 = 4 And Form7.MediaPlayer1.Volume > -50 Then
Picture2(12).Visible = True
Picture2(13).Visible = True
Picture2(14).Visible = True
Picture2(15).Visible = True
t1 = 0
Else
t1 = 0
End If
End If
End Sub

Private Sub Timer8_Timer()
If Form7.MediaPlayer1.PlayState = mpPlaying Then
Frame3.Visible = True
If t1 = 0 Then
Picture2(12).Visible = True
Picture2(13).Visible = False
Picture2(14).Visible = False
Picture2(15).Visible = False
t1 = t1 + 1
ElseIf t1 = 1 Then
Picture2(16).Visible = True
Picture2(17).Visible = True
Picture2(18).Visible = False
Picture2(19).Visible = False
t1 = t1 + 1
ElseIf t1 = 2 Then
Picture2(16).Visible = True
Picture2(17).Visible = True
Picture2(18).Visible = True
Picture2(19).Visible = False
t1 = t1 + 1
ElseIf t1 = 3 Then
Picture2(16).Visible = True
Picture2(17).Visible = True
Picture2(18).Visible = True
Picture2(19).Visible = True
t1 = t1 + 1
ElseIf t1 = 4 And Form7.MediaPlayer1.Volume > -50 Then
Picture2(16).Visible = True
Picture2(17).Visible = True
Picture2(18).Visible = True
Picture2(19).Visible = True
t1 = 0
Else
t1 = 0
End If
End If
End Sub




Private Sub Timer9_Timer()
If Form7.MediaPlayer1.Duration > 0 Then
Slider3.Max = Form7.MediaPlayer1.Duration
Else
Exit Sub
End If
On Error GoTo error
Dim tinseconden As Single
Dim tsec As Single
tinseconden = Form7.MediaPlayer1.CurrentPosition
tsec = Form7.MediaPlayer1.Duration
Dim min As Integer
Dim sec As Integer
Dim mt As Integer
Dim st As Integer
min = tinseconden \ 60
sec = tinseconden - (min * 60)
mt = tsec \ 60
st = tsec - (mt * 60)
sr = st - sec
mr = mt - min
If sec = "-1" Then sec = "0"
If st = "-1" Then st = "0"
If sec = "-1" Then sec = "0"
If st = "-1" Then st = "0"
If sr > 0 Then
If sr < 10 And st < 10 Then
Label26.Caption = mr & ":0" & sr & "      " & mt & ":0" & st
Else
If sr < 10 And st > 10 Then
Label26.Caption = mr & ":0" & sr & "      " & mt & ":" & st
End If
If sr < 10 And sec > 10 Then
Label26.Caption = mr & ":" & sr & "      " & mt & ":0" & st
End If
End If
If sr >= 10 And st >= 10 Then
Label26.Caption = mr & ":" & sr & "      " & mt & ":" & st
'Label5.Width = MediaPlayer1.FileName
Else
If st >= 10 And sr < 10 Then
Label26.Caption = mr & ":0" & sr & "      " & mt & ":" & st
End If
If sr >= 10 And st < 10 Then
Label26.Caption = mr & ":" & sr & "      " & mt & ":0" & st
End If
End If
Else
mr = mr - 1
sr = (59 + st) - sec
If mr = -1 Then
mr = 0
sr = 0
End If
If sr < 10 And st < 10 Then
Label26.Caption = mr & ":0" & sr & "      " & mt & ":0" & st
Else
If sr < 10 And st > 10 Then
Label26.Caption = mr & ":0" & sr & "      " & mt & ":" & st
End If
If sr < 10 And sec > 10 Then
Label26.Caption = mr & ":" & sr & "      " & mt & ":0" & st
End If
End If
If sr >= 10 And st >= 10 Then
Label26.Caption = mr & ":" & sr & "      " & mt & ":" & st
'Label5.Width = MediaPlayer1.FileName
Else
If st >= 10 And sr < 10 Then
Label26.Caption = mr & ":0" & sr & "      " & mt & ":" & st
End If
If sr >= 10 And st < 10 Then
Label26.Caption = mr & ":" & sr & "      " & mt & ":0" & st
End If
End If
End If
If Check6.Value = 1 And Check5.Value = 0 And sec = 30 Then
Form7.MediaPlayer1.Stop
Slider3.Value = "0"
Form7.MediaPlayer1.CurrentPosition = Slider3.Value
Label6.Caption = "Stopped"
Label25.Caption = "(No File)"
paused = False
Command13.Visible = False
Command1.Visible = True
mnutoc.Enabled = True
mnupas.Enabled = False
Check6.Enabled = True
Form7.Hide
End If
If Check6.Value = 1 And Check5.Value = 1 And Check2.Value = 0 And Check3.Value = 0 And Check4.Value = 0 And Check7.Value = 0 And sec = 30 Then
If List1.ListIndex < List1.ListCount - 1 Then
List1.ListIndex = List1.ListIndex + 1
List2.ListIndex = List2.ListIndex + 1
Text1.Text = List2.Text
Label25.Caption = Right(List1.Text, 3)
Label6.Caption = "Now Playing:" + " " + Text1.Text
Label10.Caption = Text1.Text
Text2.Locked = False
Text2.Text = (Round((FileLen(Form7.MediaPlayer1.FileName) / Form7.MediaPlayer1.SelectionEnd) / 125))  'get bitrate
Text2.Locked = True
Form7.MediaPlayer1.FileName = List1.Text
Form7.MediaPlayer1.Play
If Form7.MediaPlayer1.ImageSourceHeight = 0 Then
Label23.Caption = "(Audio)"
Form7.Hide
Else
Label23.Caption = "(Video)"
Load Form7
Form7.Show
End If
Else
Command1.Visible = True
Command13.Visible = False
mnutoc.Enabled = True
mnupas.Enabled = False
Form7.MediaPlayer1.FileName = ""
Form7.MediaPlayer1.Stop
Form7.Hide
Text1.Text = List2.Text
Label25.Caption = "(No File)"
Label6.Caption = "End Of The List"
Label23.Caption = "(No File)"
Label10.Caption = "End Of The List"
Text2.Locked = False
Text2.Text = (Round((FileLen(Form7.MediaPlayer1.FileName) / Form7.MediaPlayer1.SelectionEnd) / 125))  'get bitrate
Text2.Locked = True
End If
End If
If Check6.Value = 1 And Check5.Value = 1 And Check2.Value = 1 And Check3.Value = 0 And Check4.Value = 0 And Check7.Value = 0 And sec = 30 Then
List1.ListIndex = List1.ListIndex
List2.ListIndex = List1.ListIndex
Text1.Text = List2.Text
Form7.MediaPlayer1.FileName = List1.Text
Form7.MediaPlayer1.Play
Label25.Caption = Right(List1.Text, 3)
Label6.Caption = "Now Playing:" + " " + Text1.Text
Label10.Caption = Text1.Text
Text2.Locked = False
Text2.Text = (Round((FileLen(Form7.MediaPlayer1.FileName) / Form7.MediaPlayer1.SelectionEnd) / 125))  'get bitrate
Text2.Locked = True
If Form7.MediaPlayer1.ImageSourceHeight = 0 Then
Label23.Caption = "(Audio)"
Form7.Hide
Else
Label23.Caption = "(Video)"
Load Form7
Form7.Show
End If
End If
If Check6.Value = 1 And Check5.Value = 1 And Check3.Value = 1 And Check2.Value = 0 And Check4.Value = 0 And Check7.Value = 0 And sec = 30 Then
If List1.ListIndex = List1.ListCount - 1 Then
List1.ListIndex = 0
List2.ListIndex = 0
Text1.Text = List2.Text
Form7.MediaPlayer1.FileName = List1.Text
Form7.MediaPlayer1.Play
Label25.Caption = Right(List1.Text, 3)
Label6.Caption = "Now Playing:" + " " + Text1.Text
Label10.Caption = Form1.Text1.Text
Text2.Locked = False
Text2.Text = (Round((FileLen(Form7.MediaPlayer1.FileName) / Form7.MediaPlayer1.SelectionEnd) / 125))  'get bitrate
Text2.Locked = True
If Form7.MediaPlayer1.ImageSourceHeight = 0 Then
Label23.Caption = "(Audio)"
Form7.Hide
Else
Label23.Caption = "(Video)"
Load Form7
Form7.Show
End If
Else
List1.ListIndex = List1.ListIndex + 1
List2.ListIndex = List2.ListIndex + 1
Text1.Text = List2.Text
Form7.MediaPlayer1.FileName = List1.Text
Form7.MediaPlayer1.Play
Label25.Caption = Right(List1.Text, 3)
Label6.Caption = "Now Playing:" + " " + Text1.Text
Label10.Caption = Text1.Text
Text2.Locked = False
Text2.Text = (Round((FileLen(Form7.MediaPlayer1.FileName) / Form7.MediaPlayer1.SelectionEnd) / 125))  'get bitrate
Text2.Locked = True
If Form7.MediaPlayer1.ImageSourceHeight = 0 Then
Label23.Caption = "(Audio)"
Form7.Hide
Else
Label23.Caption = "(Video)"
Load Form7
Form7.Show
End If
End If
If Check6.Value = 1 And Check5.Value = 1 And Check4.Value = 1 And Check3.Value = 0 And Check2.Value = 0 And Check7.Value = 0 And sec = 30 Then
List1.ListIndex = 0
List2.ListIndex = 0
Text1.Text = List2.Text
Form7.MediaPlayer1.FileName = List1.Text
Form7.MediaPlayer1.Play
Label25.Caption = Right(List1.Text, 3)
Label6.Caption = "Now Playing:" + " " + Text1.Text
Label10.Caption = Text1.Text
Text2.Locked = False
Text2.Text = (Round((FileLen(Form7.MediaPlayer1.FileName) / Form7.MediaPlayer1.SelectionEnd) / 125))  'get bitrate
Text2.Locked = True
If Form7.MediaPlayer1.ImageSourceHeight = 0 Then
Label23.Caption = "(Audio)"
Form7.Hide
Else
Label23.Caption = "(Video)"
Load Form7
Form7.Show
End If
End If
If Check6.Value = 1 And Check5.Value = 1 And Check4.Value = 0 And Check3.Value = 0 And Check2.Value = 0 And Check7.Value = 1 And sec = 30 Then
Randomize
List1.ListIndex = Int(List1.ListCount * Rnd)
List2.ListIndex = List1.ListIndex
Form7.MediaPlayer1.FileName = List1.Text
Text1.Text = List2.Text
Label6.Caption = "Now Playing:" + " " + Text1.Text
Label10.Caption = Text1.Text
Label25.Caption = Right(List1.Text, 3)
Text2.Locked = False
Text2.Text = (Int((FileLen(Form7.MediaPlayer1.FileName) / Form7.MediaPlayer1.SelectionEnd) / 125))  'get bitrate
Text2.Locked = True
Text3.Locked = False
Text3.Text = 49
Command1.Visible = False
Command13.Visible = True
mnutoc.Enabled = False
mnupas.Enabled = True
Form7.MediaPlayer1.Play
If Form7.MediaPlayer1.ImageSourceHeight = 0 Then
Label23.Caption = "(Audio)"
Form7.Hide
Else
Label23.Caption = "(Video)"
Load Form7
Form7.Show
End If
End If
End If
error:
Exit Sub
End Sub

Private Sub verde_Click()
Form1.BackColor = &HC000&
Option1.BackColor = &HC000&
Option2.BackColor = &HC000&
Option1.ForeColor = black
Option2.ForeColor = black
Shape2.FillColor = black
Check1.BackColor = black
Check2.BackColor = black
Check3.BackColor = black
Check4.BackColor = black
Check5.BackColor = black
Check6.BackColor = black
Check7.BackColor = black
'Check1.ForeColor = &H8000000E
'Check2.ForeColor = &H8000000E
'Check3.ForeColor = &H8000000E
'Check4.ForeColor = &H8000000E
'Check5.ForeColor = &H8000000E
'Label1.ForeColor = &H8000000E
Label13.ForeColor = black
Label14.ForeColor = black
Label15.BackColor = &HC000&
Form1.Refresh
End Sub

Private Sub Vermelho_Click()
Form1.BackColor = &HFF&
Option1.BackColor = &HFF&
Option2.BackColor = &HFF&
Option1.ForeColor = black
Option2.ForeColor = black
Shape2.FillColor = black
Check1.BackColor = black
Check2.BackColor = black
Check3.BackColor = black
Check4.BackColor = black
Check5.BackColor = black
Check6.BackColor = black
Check7.BackColor = black
'Check1.ForeColor = &H8000000E
'Check2.ForeColor = &H8000000E
'Check3.ForeColor = &H8000000E
'Check4.ForeColor = &H8000000E
'Check5.ForeColor = &H8000000E
'Label1.ForeColor = &H8000000E
Label13.ForeColor = black
Label14.ForeColor = black
Label15.BackColor = &HFF&
Form1.Refresh
End Sub

