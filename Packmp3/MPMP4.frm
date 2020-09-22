VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "msdxm.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Packmp3  1.0(Beta) "
   ClientHeight    =   8220
   ClientLeft      =   3675
   ClientTop       =   540
   ClientWidth     =   6750
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
   Icon            =   "MPMP4.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   6750
   Begin MSComDlg.CommonDialog Cd2 
      Left            =   6840
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog Cd1 
      Left            =   6840
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6840
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   81
      Top             =   7965
      Width           =   6750
      _ExtentX        =   11906
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.Slider Slider3 
      Height          =   255
      Left            =   120
      TabIndex        =   80
      Top             =   2400
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   450
      _Version        =   327682
      TickStyle       =   3
   End
   Begin ComctlLib.Slider Slider2 
      Height          =   255
      Left            =   2400
      TabIndex        =   79
      Top             =   960
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   450
      _Version        =   327682
      Min             =   -2500
      Max             =   2500
      TickStyle       =   3
   End
   Begin ComctlLib.Slider Slider1 
      Height          =   255
      Left            =   360
      TabIndex        =   78
      Top             =   960
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   450
      _Version        =   327682
      Min             =   -2500
      Max             =   0
      TickStyle       =   3
   End
   Begin VB.ListBox List2 
      BackColor       =   &H00808000&
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   2940
      Left            =   360
      TabIndex        =   16
      Top             =   4080
      Width           =   4335
   End
   Begin VB.Frame Frame2 
      Caption         =   "MP"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3000
      TabIndex        =   73
      Top             =   1560
      Visible         =   0   'False
      Width           =   975
      Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
         Height          =   615
         Left            =   480
         TabIndex        =   82
         Top             =   360
         Width           =   615
         AudioStream     =   -1
         AutoSize        =   0   'False
         AutoStart       =   -1  'True
         AnimationAtStart=   -1  'True
         AllowScan       =   -1  'True
         AllowChangeDisplaySize=   -1  'True
         AutoRewind      =   0   'False
         Balance         =   0
         BaseURL         =   ""
         BufferingTime   =   5
         CaptioningID    =   ""
         ClickToPlay     =   -1  'True
         CursorType      =   0
         CurrentPosition =   -1
         CurrentMarker   =   0
         DefaultFrame    =   ""
         DisplayBackColor=   0
         DisplayForeColor=   16777215
         DisplayMode     =   0
         DisplaySize     =   4
         Enabled         =   -1  'True
         EnableContextMenu=   -1  'True
         EnablePositionControls=   -1  'True
         EnableFullScreenControls=   0   'False
         EnableTracker   =   -1  'True
         Filename        =   ""
         InvokeURLs      =   -1  'True
         Language        =   -1
         Mute            =   0   'False
         PlayCount       =   1
         PreviewMode     =   0   'False
         Rate            =   1
         SAMILang        =   ""
         SAMIStyle       =   ""
         SAMIFileName    =   ""
         SelectionStart  =   -1
         SelectionEnd    =   -1
         SendOpenStateChangeEvents=   -1  'True
         SendWarningEvents=   -1  'True
         SendErrorEvents =   -1  'True
         SendKeyboardEvents=   0   'False
         SendMouseClickEvents=   0   'False
         SendMouseMoveEvents=   0   'False
         SendPlayStateChangeEvents=   -1  'True
         ShowCaptioning  =   0   'False
         ShowControls    =   -1  'True
         ShowAudioControls=   -1  'True
         ShowDisplay     =   0   'False
         ShowGotoBar     =   0   'False
         ShowPositionControls=   -1  'True
         ShowStatusBar   =   0   'False
         ShowTracker     =   -1  'True
         TransparentAtStart=   0   'False
         VideoBorderWidth=   0
         VideoBorderColor=   0
         VideoBorder3D   =   0   'False
         Volume          =   -600
         WindowlessVideo =   0   'False
      End
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
      Left            =   4440
      TabIndex        =   71
      Top             =   1920
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
      Left            =   4440
      TabIndex        =   70
      Top             =   1920
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
      Left            =   4440
      TabIndex        =   69
      Top             =   1920
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
      Left            =   4440
      TabIndex        =   68
      Top             =   1920
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H000080FF&
      Caption         =   "Velocidade -"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   5520
      Width           =   1335
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
      TabIndex        =   65
      Text            =   "Text1"
      Top             =   1800
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
      TabIndex        =   34
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
         TabIndex        =   64
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
         TabIndex        =   61
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
            TabIndex        =   63
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
            TabIndex        =   62
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
         TabIndex        =   60
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
         TabIndex        =   59
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
         TabIndex        =   58
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
         TabIndex        =   57
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
         TabIndex        =   54
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
            TabIndex        =   56
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
            TabIndex        =   55
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
         TabIndex        =   53
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
         TabIndex        =   52
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
         TabIndex        =   51
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
         TabIndex        =   48
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
            TabIndex        =   50
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
            TabIndex        =   49
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
         TabIndex        =   47
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
         TabIndex        =   46
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
         TabIndex        =   45
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
         TabIndex        =   42
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
            TabIndex        =   44
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
            TabIndex        =   43
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
         TabIndex        =   41
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
         TabIndex        =   40
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
         TabIndex        =   37
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
            TabIndex        =   39
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
            TabIndex        =   38
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
         TabIndex        =   36
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
         TabIndex        =   35
         Top             =   0
         Width           =   135
      End
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   31
      ToolTipText     =   "Frequência da música"
      Top             =   360
      Width           =   375
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   30
      ToolTipText     =   "Bitrate da música"
      Top             =   360
      Width           =   375
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
      Caption         =   "Directório"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   5760
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   7800
      Top             =   240
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H000080FF&
      Caption         =   "Velocidade +"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   5280
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
      Left            =   2040
      Picture         =   "MPMP4.frx":044A
      ScaleHeight     =   555
      ScaleWidth      =   435
      TabIndex        =   23
      Top             =   1680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H80000009&
      Caption         =   "Repetir"
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
      Left            =   4920
      TabIndex        =   22
      Top             =   6840
      UseMaskColor    =   -1  'True
      Width           =   1335
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H80000009&
      Caption         =   "Sem Som"
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
      Left            =   4920
      TabIndex        =   21
      Top             =   6480
      Width           =   1335
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H80000009&
      Caption         =   "Reset"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   4920
      TabIndex        =   20
      Top             =   6240
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
      Left            =   4800
      TabIndex        =   19
      Top             =   480
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
      Left            =   4800
      TabIndex        =   18
      Top             =   960
      Width           =   255
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H000080FF&
      Caption         =   "Buscar Lista"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4560
      Width           =   1335
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H000080FF&
      Caption         =   "Guardar Lista"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4800
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
   Begin VB.CommandButton Command5 
      BackColor       =   &H000080FF&
      Caption         =   "Limpar Tudo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4920
      MaskColor       =   &H80000010&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5040
      UseMaskColor    =   -1  'True
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H000080FF&
      Caption         =   "Remover"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4320
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "Adicionar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4920
      MaskColor       =   &H80000014&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4080
      Width           =   1335
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
      Height          =   1260
      Left            =   1440
      TabIndex        =   7
      Top             =   4560
      Width           =   2055
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
      TabIndex        =   8
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
         Picture         =   "MPMP4.frx":0894
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Pausar"
         Top             =   360
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton Command10 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Aleatório"
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
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Toca uma musica a sorte"
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
         Picture         =   "MPMP4.frx":0FE6
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Anterior"
         Top             =   360
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
         Picture         =   "MPMP4.frx":1773
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Seguinte"
         Top             =   360
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
         Picture         =   "MPMP4.frx":1F34
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Parar"
         Top             =   360
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
         Picture         =   "MPMP4.frx":2659
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Tocar"
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000007&
      Caption         =   "Label8"
      Height          =   255
      Left            =   1200
      TabIndex        =   77
      Top             =   360
      Width           =   15
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   7200
      TabIndex        =   76
      Top             =   1920
      Width           =   10695
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
      Height          =   255
      Left            =   1920
      TabIndex        =   75
      ToolTipText     =   "Número de músicas"
      Top             =   3840
      Width           =   1575
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
      TabIndex        =   74
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Ver Lista"
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
      TabIndex        =   67
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Esconder Lista"
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
      Height          =   615
      Left            =   5040
      TabIndex        =   66
      Top             =   480
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
      Left            =   2760
      TabIndex        =   33
      Top             =   360
      Width           =   615
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
      TabIndex        =   32
      Top             =   360
      Width           =   495
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
      Left            =   360
      TabIndex        =   28
      Top             =   1560
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
      Left            =   1800
      TabIndex        =   24
      Top             =   1800
      Visible         =   0   'False
      Width           =   255
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
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   480
      TabIndex        =   9
      Top             =   3840
      Width           =   1815
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
      TabIndex        =   15
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
      TabIndex        =   11
      Top             =   960
      Width           =   975
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
      Left            =   1200
      TabIndex        =   10
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
      Left            =   3960
      TabIndex        =   12
      Top             =   1560
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   3495
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   3840
      Width           =   6495
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      Caption         =   "0:00    0:00"
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
      TabIndex        =   72
      ToolTipText     =   "Tempo da música"
      Top             =   360
      UseMnemonic     =   0   'False
      Width           =   1695
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   120
      Picture         =   "MPMP4.frx":2DEC
      Stretch         =   -1  'True
      Top             =   120
      Width           =   4440
   End
   Begin VB.Menu mnupopup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnufich 
         Caption         =   "Ficheiro"
         Begin VB.Menu mnuabr 
            Caption         =   "Abrir"
         End
      End
      Begin VB.Menu mnuac 
         Caption         =   "Acçoes"
         Begin VB.Menu mnutoc 
            Caption         =   "Tocar"
         End
         Begin VB.Menu mnupar 
            Caption         =   "Parar"
         End
         Begin VB.Menu mnupas 
            Caption         =   "Pausar"
         End
         Begin VB.Menu mnuseg 
            Caption         =   "Seguinte"
         End
         Begin VB.Menu mnuant 
            Caption         =   "Anterior"
         End
      End
      Begin VB.Menu mnucon 
         Caption         =   "Controlo Volume"
      End
      Begin VB.Menu mnucor 
         Caption         =   "Cor de Fundo"
         Begin VB.Menu mnubranc 
            Caption         =   "Branco"
         End
         Begin VB.Menu mnupret 
            Caption         =   "Preto"
         End
         Begin VB.Menu verde 
            Caption         =   "Verde"
         End
         Begin VB.Menu Vermelho 
            Caption         =   "Vermelho"
         End
         Begin VB.Menu Azul 
            Caption         =   "Azul"
         End
         Begin VB.Menu Laranja 
            Caption         =   "Laranja"
         End
      End
      Begin VB.Menu mnusp 
         Caption         =   "-"
      End
      Begin VB.Menu mnusair 
         Caption         =   "Sair"
      End
   End
   Begin VB.Menu mnuteste 
      Caption         =   "Img"
      Visible         =   0   'False
      Begin VB.Menu mnusbr 
         Caption         =   "Sobre"
      End
      Begin VB.Menu mnsep 
         Caption         =   "-"
      End
      Begin VB.Menu mnutsair 
         Caption         =   "Sair"
      End
   End
   Begin VB.Menu mnuinf 
      Caption         =   "Informaçoes"
      Visible         =   0   'False
      Begin VB.Menu mnuid3 
         Caption         =   "Id3"
      End
      Begin VB.Menu mnudes 
         Caption         =   "Descrição"
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
'Dim sp As New SpeechLib.SpVoice





Private Sub Azul_Click()
Form1.BackColor = &HC00000
Option1.BackColor = &HC00000
Option2.BackColor = &HC00000
Option1.ForeColor = BLACK
Option2.ForeColor = BLACK
Shape2.FillColor = BLACK
Check1.BackColor = BLACK
Check2.BackColor = BLACK
Check3.BackColor = BLACK
Check1.ForeColor = &H8000000E
Check2.ForeColor = &H8000000E
Check3.ForeColor = &H8000000E
Label1.ForeColor = &H8000000E
Label13.ForeColor = BLACK
Label14.ForeColor = BLACK
Form1.Refresh
End Sub

'Dim allow_stop As Boolean
'Dim allow_pause As Boolean



Private Sub Check1_Click()
If Check1.Value = 1 Then
MediaPlayer1.Mute = True
Else
MediaPlayer1.Mute = False
End If
End Sub



Private Sub Command1_Click()
On Error GoTo hell
List1.ListIndex = List2.ListIndex
If List1.Text <> "" Then
If Check1.Value = True Then
MediaPlayer1.Mute = True
End If
If paused = False Then
Text1.Text = List2.Text
Label6.Caption = "Now Playing:" + " " + Text1.Text
Label10.Caption = Text1.Text
MediaPlayer1.FileName = List1.Text
MediaPlayer1.Play
Text2.Locked = False
Text3.Locked = False
Text2.Text = (Round((FileLen(MediaPlayer1.FileName) / MediaPlayer1.SelectionEnd) / 125))
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
Else
Label6.Caption = "Now Playing:" + " " + Text1.Text
Label10.Caption = Text1.Text
MediaPlayer1.Play
vol = Abs(Slider1.Value) - 2500
MediaPlayer1.Volume = vol
paused = False
'allow_stop = True
'allow_pause = True
Command1.Visible = False
Command13.Visible = True
mnutoc.Enabled = False
mnupas.Enabled = True
End If
End If
Exit Sub
hell:
sp.Speak "The music" + List2.Text + " isn't supported"
'If List1.ListIndex > "-1 " Then
End Sub

Private Sub Command10_Click()
On Error Resume Next
If List1.ListCount > 0 And List2.ListCount > 0 Then
Randomize
List1.ListIndex = Int(List1.ListCount * Rnd)
List2.ListIndex = List1.ListIndex
MediaPlayer1.FileName = List1.Text
Text1.Text = List2.Text
Label6.Caption = "Now Playing:" + " " + Text1.Text
Label10.Caption = Text1.Text
Text2.Locked = False
Text2.Text = (Int((FileLen(MediaPlayer1.FileName) / MediaPlayer1.SelectionEnd) / 125))  'get bitrate
Text2.Locked = True
Text3.Locked = False
Text3.Text = 49
Command1.Visible = False
Command13.Visible = True
mnutoc.Enabled = False
mnupas.Enabled = True
End If
End Sub

Private Sub Command11_Click()
On Error Resume Next
MediaPlayer1.Rate = Label9.Caption
If Label9.Caption < 8 Then
Label9.Caption = Label9.Caption + 1
MediaPlayer1.Rate = Label9.Caption
End If
End Sub

Private Sub Command12_Click()
On Error Resume Next
MediaPlayer1.Rate = Label9.Caption
If Label9.Caption > 1 Then
Label9.Caption = Label9.Caption - 1
MediaPlayer1.Rate = Label9.Caption
End If
End Sub

Private Sub Command13_Click()
On Error Resume Next
  If paused = False Then
       'If allow_pause = True Then
         MediaPlayer1.Pause
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

Private Sub Command2_Click()
On Error Resume Next
If MediaPlayer1.FileName <> "" Then
MediaPlayer1.Stop
Slider3.Value = "0"
MediaPlayer1.CurrentPosition = Slider3.Value
Label6.Caption = "Stopped"
paused = False
Command13.Visible = False
Command1.Visible = True
mnutoc.Enabled = True
mnupas.Enabled = False
sp.Speak "A música esta parada"
sp.Volume = Slider1.Value
End If
End Sub

Private Sub Command3_Click()
On Error Resume Next
If List2.ListIndex = "-1" Then
resp = MsgBox("Seleccione um item para remover.", 48, "Mensagem:")
Else
pos = List2.ListIndex
List1.ListIndex = List2.ListIndex
a = List1.Text
Set fso = CreateObject("Scripting.FileSystemObject")
Set att = fso.getfile(a)
read.Enabled = True
Hidden.Enabled = True
archive.Enabled = True
system.Enabled = True
system.Value = 0
Hidden.Value = 0
archive.Value = 0
read.Value = 0
If read.Value = 0 And Hidden.Value = 0 And archive.Value = 0 And system.Value = 0 Then
att.Attributes = 0
End If
List2.RemoveItem (pos)
List1.RemoveItem (pos)
List2.Refresh
List1.Refresh
Label9.Caption = 1
MediaPlayer1.Rate = Label9.Caption
Label17.Caption = List2.ListCount
If a = MediaPlayer1.FileName Then
MediaPlayer1.Stop
MediaPlayer1.FileName = ""
Label6.Caption = ""
Command13.Visible = False
Command1.Visible = True
mnutoc.Enabled = True
mnupas.Enabled = False
End If
End If
End Sub

Private Sub Command4_Click()
On Error Resume Next
CommonDialog1.DialogTitle = "Procurar musica"
CommonDialog1.Filter = "Mp3|*.mp3*;|Wav|*.wav;|Wma|*.wma;|Mdi|*.mdi;|Avi|*.avi;|Mpeg|*.mpeg;|Mpg|*.mpg;"
CommonDialog1.InitDir = "C:\"
CommonDialog1.FileName = ""
CommonDialog1.ShowOpen
If CommonDialog1.FileName <> "" Then
List1.AddItem CommonDialog1.FileName
List2.AddItem CommonDialog1.FileTitle
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
List1.ListIndex = 0
While List1.ListIndex <= List1.ListCount
List2.ListIndex = List1.ListIndex
SetAttr List1.Text, vbNormal
List1.RemoveItem (List1.ListIndex)
List2.RemoveItem (List2.ListIndex)
List1.ListIndex = List1.ListIndex + 1
Wend
MediaPlayer1.Stop
MediaPlayer1.FileName = ""
Label6.Caption = ""
Label9.Caption = 1
MediaPlayer1.Rate = Label9.Caption
Command13.Visible = False
Command1.Visible = True
mnutoc.Enabled = True
mnupas.Enabled = False
Slider3.Value = "0"
Label17.Caption = List2.ListCount
a:
List1.Clear
List2.Clear
MediaPlayer1.Stop
MediaPlayer1.FileName = ""
Label6.Caption = ""
Label9.Caption = 1
MediaPlayer1.Rate = Label9.Caption
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
MediaPlayer1.FileName = List1.Text
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
If Check1.Value = True Then
MediaPlayer1.Mute = True
End If
End Sub

Private Sub Command7_Click()
On Error Resume Next
If List1.Text <> "" Then
If Not List1.ListIndex = 0 Then
List1.ListIndex = List1.ListIndex - 1
List2.ListIndex = List2.ListIndex - 1
Text1.Text = List2.Text
MediaPlayer1.FileName = List1.Text
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
If Check1.Value = True Then
MediaPlayer1.Mute = True
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
Cd2.DialogTitle = "Save Mp3 PlayList"
Cd2.Filter = "Mp3 PlayList|*.lst"
Cd2.InitDir = "c:\Listas Mp3"
Cd2.FileName = ""
Cd2.ShowSave
Else
Exit Sub
End If
Open cd.FileName For Output As #1
For i = 0 To List1.ListCount - 1
a = List1.List(i)
Print #1, a
Next i
Close #1
If cd.FileName <> "" And Cd2.FileName <> "" Then
Open Cd2.FileName For Output As #2
For k = 0 To List2.ListCount - 1
b = List2.List(k)
Print #2, b
Next k
Close #2
Exit Sub
End If
End Sub

Private Sub Command9_Click()
On Error Resume Next
cd.DialogTitle = "Save Mp3 PlayList"
cd.Filter = "Mp3 PlayList|*.txt"
cd.FileName = ""
cd.InitDir = "c:\Listas Mp3"
cd.ShowOpen
If cd.FileName <> "" Then
Cd2.DialogTitle = "Save Mp3 PlayList"
Cd2.Filter = "Mp3 PlayList|*.lst"
Cd2.InitDir = "c:\Listas Mp3"
Cd2.FileName = ""
Cd2.ShowOpen
Else
Exit Sub
End If
If cd.FileName <> "" And Cd2.FileName <> "" Then
List1.Clear
List2.Clear
MediaPlayer1.Stop
Command13.Visible = False
Command1.Visible = True
Label6.Caption = ""
Open cd.FileName For Input As #1
While Not EOF(1)
Line Input #1, a
List1.AddItem a
Label17.Caption = List1.ListCount
Wend
Close #1
Open Cd2.FileName For Input As #2
While Not EOF(2)
Line Input #2, b
List2.AddItem b
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


Private Sub Form_Load()
Dim img As NOTIFYICONDATA
Label1.Caption = "Lista pessoal:"
Label3.Caption = "Volume " & foo \ 25 & " %"
Label4.Caption = "Center"
Slider1.Value = "1250"
vol = Abs(Slider1.Value) - 2500
MediaPlayer1.Volume = vol
Label3.Caption = "Volume 50 %"
Option2.Value = True
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
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
msg = X / Screen.TwipsPerPixelX
If msg = WM_LeBUTTONDBLCLCK Then
ElseIf msg = WM_RBUTTONUP Then
Me.PopupMenu mnuteste
End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
PopupMenu mnupopup
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
img.cbSize = Len(img)
img.hwnd = Me.hwnd
img.uId = 1&
Shell_NotifyIcon NIM_DELETE, img
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



Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
PopupMenu mnupopup
End If
End Sub



Private Sub List2_Click()
List1.ListIndex = List2.ListIndex
End Sub

Private Sub List2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
List1.ListIndex = List2.ListIndex
sti = List2.ListIndex
st2 = List2.Text
st1 = List1.Text
End Sub

Private Sub List2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button = 2 Then
 PopupMenu mnuinf
 End If
 Dim Si As Integer 'declare a variable to store the list index of the finishing point
    List1.ListIndex = List2.ListIndex
    If sti <> List2.ListIndex Then  'if they moved the listitem
        Si = List2.ListIndex 'store the destination listindex
        List1.ListIndex = Si
        List2.RemoveItem sti 'remove the original row so everything below shifts up
        List1.RemoveItem sti
        List2.AddItem st2, Si 'add the original row back into the new place
        List1.AddItem st1, Si
        List2.ListIndex = Si 'making it look better by having the selection be on the destination row
        List1.ListIndex = Si
    End If
    Dragging = False
End Sub

Private Sub MediaPlayer1_EndOfStream(ByVal Result As Long)
On Error GoTo Err
c = 1
If Check2.Value = 1 Then
List1.ListIndex = 0
List2.ListIndex = 0
Text1.Text = List2.Text
MediaPlayer1.FileName = List1.Text
Label6.Caption = "Now Playing:" + " " + Text1.Text
Label10.Caption = Text1.Text
MediaPlayer1.Play
Text2.Locked = False
Text2.Text = (Round((FileLen(MediaPlayer1.FileName) / MediaPlayer1.SelectionEnd) / 125))  'get bitrate
Text2.Locked = True
Else
a:
If Check3.Value = 1 Then
List1.ListIndex = List1.ListIndex
Text1.Text = List2.Text
Label6.Caption = "Now Playing:" + " " + Text1.Text
Label10.Caption = Text1.Text
MediaPlayer1.FileName = List1.Text
Text2.Locked = False
Text2.Text = (Round((FileLen(MediaPlayer1.FileName) / MediaPlayer1.SelectionEnd) / 125)) 'get bitrate
Text2.Locked = True
Else
List1.ListIndex = List1.ListIndex + 1
List2.ListIndex = List2.ListIndex + 1
Text1.Text = List2.Text
Label6.Caption = "Now Playing:" + " " + Text1.Text
Label10.Caption = Text1.Text
MediaPlayer1.FileName = List1.Text
Text2.Locked = False
Text2.Text = (Round((FileLen(MediaPlayer1.FileName) / MediaPlayer1.SelectionEnd) / 125))  'get bitrate
Text2.Locked = True
End If
Exit Sub
End If
If c = 1 Then
List1.ListIndex = 0
List2.ListIndex = 0
Text1.Text = List2.Text
Label6.Caption = "Now Playing:" + " " + Text1.Text
Label10.Caption = Text1.Text
MediaPlayer1.FileName = List1.Text
c = c + 1
Text2.Locked = False
'Text2.Text = (Round((FileLen(MediaPlayer1.FileName) / MediaPlayer1.SelectionEnd) / 125))  'get bitrate
Text2.Locked = True
Else
GoTo a
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
End Sub


Private Sub mnuabr_Click()
On Error Resume Next
CommonDialog1.DialogTitle = "Procurar musica"
CommonDialog1.Filter = "Mp3|*.mp3*;|Wav|*.wav;|Wma|*.wma;|Mdi|*.mdi;|Avi|*.avi;|Mpeg|*.mpeg;|Mpg|*.mpg;"
CommonDialog1.InitDir = "C:\"
CommonDialog1.FileName = ""
CommonDialog1.ShowOpen
If CommonDialog1.FileName <> "" Then
List1.AddItem CommonDialog1.FileName
List2.AddItem CommonDialog1.FileTitle
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
MediaPlayer1.FileName = List1.Text
Label6.Caption = "Now Playing:" + " " + Text1.Text
Label10.Caption = Text1.Text
paused = False
Command13.Visible = True
Command1.Visible = False
mnutoc.Enabled = False
mnupas.Enabled = True
Text2.Locked = False
Text2.Text = (Round((FileLen(MediaPlayer1.FileName) / MediaPlayer1.SelectionEnd) / 125))  'get bitrate
'allow_stop = True
'allow_pause = True
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
If Check1.Value = True Then
MediaPlayer1.Mute = True
End If
End Sub

Private Sub mnubranc_Click()
Form1.BackColor = &H8000000E
Option1.BackColor = &H8000000E
Option2.BackColor = &H8000000E
Option1.ForeColor = BLACK
Option2.ForeColor = BLACK
Shape2.FillColor = BLACK
Check1.BackColor = BLACK
Check2.BackColor = BLACK
Check3.BackColor = BLACK
Check1.ForeColor = &H8000000E
Check2.ForeColor = &H8000000E
Check3.ForeColor = &H8000000E
Label1.ForeColor = &H8000000E
Label13.ForeColor = BLACK
Label14.ForeColor = BLACK
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
If MediaPlayer1.FileName <> "" Then
MediaPlayer1.Stop
Slider3.Value = "0"
MediaPlayer1.CurrentPosition = Slider3.Value
Label6.Caption = "Stopped"
paused = False
Command13.Visible = False
Command1.Visible = True
mnutoc.Enabled = True
mnupas.Enabled = False
sp.Speak "A música esta parada"
sp.Volume = Slider1.Value
End If
End Sub

Private Sub mnupas_Click()
On Error Resume Next
      If paused = False Then
       'If allow_pause = True Then
         MediaPlayer1.Pause
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
Form1.BackColor = BLACK
Option1.BackColor = BLACK
Option2.BackColor = BLACK
Option1.ForeColor = &H8000000E
Option2.ForeColor = &H8000000E
Shape2.FillColor = &H8000000E
Check1.BackColor = &H8000000E
Check2.BackColor = &H8000000E
Check3.BackColor = &H8000000E
Check1.ForeColor = BLACK
Check2.ForeColor = BLACK
Check3.ForeColor = BLACK
Label1.ForeColor = BLACK
Label13.ForeColor = &H8000000E
Label14.ForeColor = &H8000000E
Form1.Refresh
End Sub

Private Sub mnusair_Click()
img.cbSize = Len(img)
img.hwnd = Me.hwnd
img.uId = 1&
Shell_NotifyIcon NIM_DELETE, img
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
MediaPlayer1.FileName = List1.Text
Label6.Caption = "Now Playing:" + " " + Text1.Text
Label10.Caption = Text1.Text
paused = False
Command13.Visible = True
Command1.Visible = False
mnutoc.Enabled = False
mnupas.Enabled = True
Text2.Locked = False
Text2.Text = (Round((FileLen(MediaPlayer1.FileName) / MediaPlayer1.SelectionEnd) / 125))  'get bitrate
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
If Check1.Value = True Then
MediaPlayer1.Mute = True
End If
End Sub

Private Sub mnutoc_Click()
On Error GoTo hell
List1.ListIndex = List2.ListIndex
If List1.Text <> "" Then
If Check1.Value = True Then
MediaPlayer1.Mute = True
End If
If paused = False Then
Text1.Text = List2.Text
MediaPlayer1.FileName = List1.Text
MediaPlayer1.Play
Label6.Caption = "Now Playing:" + " " + Text1.Text
Label10.Caption = Text1.Text
Text2.Locked = False
Text3.Locked = False
Text2.Text = (Round((FileLen(MediaPlayer1.FileName) / MediaPlayer1.SelectionEnd) / 125))
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
Else
MediaPlayer1.Play
Label6.Caption = "Now Playing:" + " " + Text1.Text
Label10.Caption = Text1.Text
vol = Abs(Slider1.Value) - 2500
MediaPlayer1.Volume = vol
paused = False
'allow_stop = True
'allow_pause = True
Command1.Visible = False
Command13.Visible = True
mnutoc.Enabled = False
mnupas.Enabled = True
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
MediaPlayer1.Stop
End Sub

Private Sub mnutsair_Click()
resp = MsgBox("Tem a certeza que deseja sair?", 292, "Sair:")
If resp = vbYes Then
img.cbSize = Len(img)
img.hwnd = Me.hwnd
img.uId = 1&
Shell_NotifyIcon NIM_DELETE, img
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
Form1.Height = 8200
End Sub

Private Sub Option2_Click()
Form1.Height = 4590
End Sub




Private Sub Slider1_Click()
Dim vol
vol = Abs(Slider1.Value) - 2500
MediaPlayer1.Volume = vol
Check1.Value = False
MediaPlayer1.Mute = False
Dim foo As Integer, poo As Integer
On Error GoTo hell
poo = Slider1.min
foo = Abs(Slider1.Value)
Label3.Caption = "Volume " & foo \ 25 & " %"
hell:
End Sub


Private Sub Slider1_Scroll()
Dim pim, sha
sha = Abs(Slider1.Value) - 2500
MediaPlayer1.Volume = sha
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
If Slider2.Value > -500 And Slider2.Value < 500 Then
Label4.Caption = "Center"
End If
If Slider2.Value < -500 Then
Label4.Caption = "Balance:" & -(MediaPlayer1.Balance / 50) & " % Left "
End If
If Slider2.Value > 500 Then
Label4.Caption = "Balance :" & MediaPlayer1.Balance / 50 & " % Right "
End If
MediaPlayer1.Balance = Slider2.Value
hell:
Exit Sub
End Sub

Private Sub Slider3_Click()
MediaPlayer1.CurrentPosition = Slider3.Value
End Sub



Private Sub Timer1_Timer()
Slider3.Value = MediaPlayer1.CurrentPosition
End Sub

Private Sub Timer2_Timer()
If MediaPlayer1.Duration > 0 Then
Slider3.Max = MediaPlayer1.Duration
Else
Exit Sub
End If
On Error GoTo error
Dim tinseconden As Single
Dim tsec As Single
tinseconden = MediaPlayer1.CurrentPosition
tsec = MediaPlayer1.Duration
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
Label5.Caption = min & ":0" & sec & "    " & mt & ":0" & st
Else
If sec < 10 And st > 10 Then
Label5.Caption = min & ":0" & sec & "    " & mt & ":" & st
End If
If st < 10 And sec > 10 Then
Label5.Caption = min & ":" & sec & "    " & mt & ":0" & st
End If
End If
If sec >= 10 And st >= 10 Then
Label5.Caption = min & ":" & sec & "    " & mt & ":" & st
'Label5.Width = MediaPlayer1.FileName
Else
If st >= 10 And sec < 10 Then
Label5.Caption = min & ":0" & sec & "    " & mt & ":" & st
End If
If sec >= 10 And st < 10 Then
Label5.Caption = min & ":" & sec & "    " & mt & ":0" & st
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
If MediaPlayer1.PlayState = mpPlaying Then
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
ElseIf t1 = 4 And MediaPlayer1.Volume > -50 Then
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
If MediaPlayer1.PlayState = mpPlaying Then
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
ElseIf t1 = 4 And MediaPlayer1.Volume > -50 Then
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
If MediaPlayer1.PlayState = mpPlaying Then
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
ElseIf t1 = 4 And MediaPlayer1.Volume > -50 Then
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
If MediaPlayer1.PlayState = mpPlaying Then
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
ElseIf t1 = 4 And MediaPlayer1.Volume > -50 Then
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
If MediaPlayer1.PlayState = mpPlaying Then
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
ElseIf t1 = 4 And MediaPlayer1.Volume > -50 Then
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


Private Sub verde_Click()
Form1.BackColor = &HC000&
Option1.BackColor = &HC000&
Option2.BackColor = &HC000&
Option1.ForeColor = BLACK
Option2.ForeColor = BLACK
Shape2.FillColor = BLACK
Check1.BackColor = BLACK
Check2.BackColor = BLACK
Check3.BackColor = BLACK
Check1.ForeColor = &H8000000E
Check2.ForeColor = &H8000000E
Check3.ForeColor = &H8000000E
Label1.ForeColor = &H8000000E
Label13.ForeColor = BLACK
Label14.ForeColor = BLACK
Form1.Refresh
End Sub

Private Sub Vermelho_Click()
Form1.BackColor = &HFF&
Option1.BackColor = &HFF&
Option2.BackColor = &HFF&
Option1.ForeColor = BLACK
Option2.ForeColor = BLACK
Shape2.FillColor = BLACK
Check1.BackColor = BLACK
Check2.BackColor = BLACK
Check3.BackColor = BLACK
Check1.ForeColor = &H8000000E
Check2.ForeColor = &H8000000E
Check3.ForeColor = &H8000000E
Label1.ForeColor = &H8000000E
Label13.ForeColor = BLACK
Label14.ForeColor = BLACK
Form1.Refresh
End Sub
