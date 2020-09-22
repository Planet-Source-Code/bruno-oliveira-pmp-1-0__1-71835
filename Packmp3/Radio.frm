VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form2 
   BackColor       =   &H80000012&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Internet Radio"
   ClientHeight    =   1470
   ClientLeft      =   3765
   ClientTop       =   5595
   ClientWidth     =   4545
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1470
   ScaleWidth      =   4545
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.Slider Slider1 
      Height          =   855
      Left            =   4200
      TabIndex        =   6
      Top             =   480
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   1508
      _Version        =   393216
      BorderStyle     =   1
      OLEDropMode     =   1
      Orientation     =   1
      Min             =   -100
      Max             =   0
      TickStyle       =   3
      TextPosition    =   1
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stop"
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Play"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Text            =   "http://radio.com/listen.asx"
      Top             =   480
      Width           =   3975
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000012&
      Caption         =   "Volume:0%"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   3600
      TabIndex        =   7
      Top             =   240
      Width           =   975
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmp 
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Visible         =   0   'False
      Width           =   3975
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   7011
      _cy             =   873
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      Caption         =   "Url:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
If Text1.Text <> "" Then
wmp.Url = Text1.Text
Label5.Caption = wmp.currentPlaylist.Name
End If
End Sub

Private Sub Command2_Click()
On Error Resume Next
wmp.Controls.Stop
End Sub

Private Sub Command3_Click()
On Error Resume Next
Unload Form2
Form2.Hide
End Sub

Private Sub Slider1_Click()
On Error Resume Next
wmp.settings.Volume = (-2) * ((Slider1.Value) / 2)
Label2.Caption = "Volume:" & wmp.settings.Volume & "%"
End Sub
