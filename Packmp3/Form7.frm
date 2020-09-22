VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "msdxm.ocx"
Begin VB.Form Form7 
   BackColor       =   &H80000007&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   " Videos"
   ClientHeight    =   6990
   ClientLeft      =   60
   ClientTop       =   165
   ClientWidth     =   8430
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6990
   ScaleWidth      =   8430
   ShowInTaskbar   =   0   'False
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
      CausesValidation=   0   'False
      Height          =   6765
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   8265
      AudioStream     =   0
      AutoSize        =   0   'False
      AutoStart       =   0   'False
      AnimationAtStart=   0   'False
      AllowScan       =   0   'False
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   0   'False
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   0   'False
      EnablePositionControls=   0   'False
      EnableFullScreenControls=   0   'False
      EnableTracker   =   0   'False
      Filename        =   ""
      InvokeURLs      =   0   'False
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
      SendOpenStateChangeEvents=   0   'False
      SendWarningEvents=   0   'False
      SendErrorEvents =   0   'False
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   -1  'True
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   0   'False
      ShowCaptioning  =   0   'False
      ShowControls    =   0   'False
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   0   'False
      ShowStatusBar   =   0   'False
      ShowTracker     =   0   'False
      TransparentAtStart=   -1  'True
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -600
      WindowlessVideo =   0   'False
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Resize()
On Error Resume Next
MediaPlayer1.Height = (Form7.Height - 80)
MediaPlayer1.Width = (Form7.Width - 130)
End Sub

Private Sub MediaPlayer1_DblClick(Button As Integer, ShiftState As Integer, x As Single, y As Single)
MediaPlayer1.DisplaySize = mpFullScreen
End Sub

Private Sub MediaPlayer1_EndOfStream(ByVal Result As Long)
Dim j As Long
j = Form1.List1.ListCount - 1
On Error GoTo Err
If Form1.Check5.Value = 1 And Form1.Check7.Value = 1 And Form1.Check2.Value = 0 And Form1.Check3.Value = 0 And Form1.Check4.Value = 0 Then
Randomize
MediaPlayer1.Visible = False
Form1.List1.ListIndex = Int(Form1.List1.ListCount * Rnd)
Form1.List2.ListIndex = Form1.List1.ListIndex
MediaPlayer1.FileName = Form1.List1.Text
Form1.Text1.Text = Form1.List2.Text
Form1.Label6.Caption = "Now Playing:" + " " + Form1.Text1.Text
Form1.Label10.Caption = Form1.Text1.Text
Form1.Label25.Caption = Right(Form1.List1.Text, 3)
Form1.Text2.Locked = False
Form1.Text2.Text = (Int((FileLen(MediaPlayer1.FileName) / MediaPlayer1.SelectionEnd) / 125))  'get bitrate
Form1.Text2.Locked = True
Form1.Text3.Locked = False
Form1.Text3.Text = 49
Form1.Command1.Visible = False
Form1.Command13.Visible = True
Form1.mnutoc.Enabled = False
Form1.mnupas.Enabled = True
MediaPlayer1.Play
MediaPlayer1.Visible = True
If MediaPlayer1.ImageSourceHeight = 0 Then
Form1.Label23.Caption = "(Audio)"
Form7.Hide
Else
Form1.Label23.Caption = "(Video)"
Load Form7
Form7.Show
End If
End If
If Form1.Check5.Value = 1 And Form1.Check2.Value = 0 And Form1.Check3.Value = 0 And Form1.Check4.Value = 0 And Form1.Check7.Value = 0 Then
If Form1.List1.ListIndex < j Then
MediaPlayer1.Visible = False
Form1.List1.ListIndex = Form1.List1.ListIndex + 1
Form1.List2.ListIndex = Form1.List2.ListIndex + 1
Form1.Text1.Text = Form1.List2.Text
Form1.Label25.Caption = Right(Form1.List1.Text, 3)
Form1.Label6.Caption = "Now Playing:" + " " + Form1.Text1.Text
Form1.Label10.Caption = Form1.Text1.Text
Form1.Text2.Locked = False
Form1.Text2.Text = (Round((FileLen(MediaPlayer1.FileName) / MediaPlayer1.SelectionEnd) / 125))  'get bitrate
Form1.Text2.Locked = True
MediaPlayer1.FileName = Form1.List1.Text
MediaPlayer1.Play
MediaPlayer1.Visible = True
If MediaPlayer1.ImageSourceHeight = 0 Then
Form1.Label23.Caption = "(Audio)"
Form7.Hide
Else
Form1.Label23.Caption = "(Video)"
Load Form7
Form7.Show
End If
Else
MediaPlayer1.FileName = ""
MediaPlayer1.Stop
Form7.Hide
Form1.Text1.Text = Form1.List2.Text
Form1.Label25.Caption = "(No File)"
Form1.Label6.Caption = "End Of The List"
Form1.Label23.Caption = "(No File)"
Form1.Label10.Caption = "End Of The List"
Form1.Text2.Locked = False
Form1.Text2.Text = (Round((FileLen(MediaPlayer1.FileName) / MediaPlayer1.SelectionEnd) / 125))  'get bitrate
Form1.Text2.Locked = True
Form1.Command1.Visible = True
Form1.Command13.Visible = False
Form1.mnutoc.Enabled = True
Form1.mnupas.Enabled = False
End If
Else
If Form1.Check5.Value = 1 And Form1.Check2.Value = 1 And Form1.Check3.Value = 0 And Form1.Check4.Value = 0 And Form1.Check7.Value = 0 Then
MediaPlayer1.Visible = False
Form1.List1.ListIndex = Form1.List1.ListIndex
Form1.List2.ListIndex = Form1.List1.ListIndex
Form1.Text1.Text = Form1.List2.Text
MediaPlayer1.FileName = Form1.List1.Text
MediaPlayer1.Play
Form1.Label25.Caption = Right(Form1.List1.Text, 3)
Form1.Label6.Caption = "Now Playing:" + " " + Form1.Text1.Text
Form1.Label10.Caption = Form1.Text1.Text
Form1.Text2.Locked = False
Form1.Text2.Text = (Round((FileLen(MediaPlayer1.FileName) / MediaPlayer1.SelectionEnd) / 125))  'get bitrate
Form1.Text2.Locked = True
MediaPlayer1.Visible = True
If MediaPlayer1.ImageSourceHeight = 0 Then
Form1.Label23.Caption = "(Audio)"
Form7.Hide
Else
Form1.Label23.Caption = "(Video)"
Load Form7
Form7.Show
End If
End If
If Form1.Check5.Value = 1 And Form1.Check3.Value = 1 And Form1.Check2.Value = 0 And Form1.Check4.Value = 0 And Form1.Check7.Value = 0 Then
If Form1.List1.ListIndex = Form1.List1.ListCount - 1 Then
MediaPlayer1.Visible = False
Form1.List1.ListIndex = 0
Form1.List2.ListIndex = 0
Form1.Text1.Text = Form1.List2.Text
MediaPlayer1.FileName = Form1.List1.Text
MediaPlayer1.Play
Form1.Label25.Caption = Right(Form1.List1.Text, 3)
Form1.Label6.Caption = "Now Playing:" + " " + Form1.Text1.Text
Form1.Label10.Caption = Form1.Text1.Text
Form1.Text2.Locked = False
Form1.Text2.Text = (Round((FileLen(MediaPlayer1.FileName) / MediaPlayer1.SelectionEnd) / 125))  'get bitrate
Form1.Text2.Locked = True
MediaPlayer1.Visible = True
If MediaPlayer1.ImageSourceHeight = 0 Then
Form1.Label23.Caption = "(Audio)"
Form7.Hide
Else
Form1.Label23.Caption = "(Video)"
Load Form7
Form7.Show
End If
Else
MediaPlayer1.Visible = False
Form1.List1.ListIndex = Form1.List1.ListIndex + 1
Form1.List2.ListIndex = Form1.List2.ListIndex + 1
Form1.Text1.Text = Form1.List2.Text
MediaPlayer1.FileName = Form1.List1.Text
MediaPlayer1.Play
Form1.Label25.Caption = Right(Form1.List1.Text, 3)
Form1.Label6.Caption = "Now Playing:" + " " + Form1.Text1.Text
Form1.Label10.Caption = Form1.Text1.Text
Form1.Text2.Locked = False
Form1.Text2.Text = (Round((FileLen(MediaPlayer1.FileName) / MediaPlayer1.SelectionEnd) / 125))  'get bitrate
Form1.Text2.Locked = True
MediaPlayer1.Visible = True
If MediaPlayer1.ImageSourceHeight = 0 Then
Form1.Label23.Caption = "(Audio)"
Form7.Hide
Else
Form1.Label23.Caption = "(Video)"
Load Form7
Form7.Show
End If
End If
End If
If Form1.Check5.Value = 1 And Form1.Check4.Value = 1 And Form1.Check3.Value = 0 And Form1.Check2.Value = 0 And Form1.Check7.Value = 0 Then
MediaPlayer1.Visible = False
Form1.List1.ListIndex = 0
Form1.List2.ListIndex = 0
Form1.Text1.Text = Form1.List2.Text
MediaPlayer1.FileName = Form1.List1.Text
MediaPlayer1.Play
Form1.Label25.Caption = Right(Form1.List1.Text, 3)
Form1.Label6.Caption = "Now Playing:" + " " + Form1.Text1.Text
Form1.Label10.Caption = Form1.Text1.Text
Form1.Text2.Locked = False
Form1.Text2.Text = (Round((FileLen(MediaPlayer1.FileName) / MediaPlayer1.SelectionEnd) / 125))  'get bitrate
Form1.Text2.Locked = True
MediaPlayer1.Visible = True
If MediaPlayer1.ImageSourceHeight = 0 Then
Form1.Label23.Caption = "(Audio)"
Form7.Hide
Else
Form1.Label23.Caption = "(Video)"
Load Form7
Form7.Show
End If
End If
If Form1.Check5.Value = 0 Then
h:
Form1.Label6.Caption = ""
Form1.Label10.Caption = ""
Form7.Hide
MediaPlayer1.Stop
Form1.Check6.Enabled = True
Form1.Label25.Caption = "(No File)"
Form1.Label23.Caption = "(No File)"
MediaPlayer1.FileName = ""
Form1.Text2.Locked = False
Form1.Text2.Text = "0"
Form1.Text3.Text = "0"
Form1.Text2.Locked = True
Form1.Slider3.Value = "0"
Form1.Command13.Visible = False
Form1.Command1.Visible = True
MediaPlayer1.CurrentPosition = Form1.Slider3.Value
MediaPlayer1.ClickToPlay = False
MediaPlayer1.EnableContextMenu = False
MediaPlayer1.SendMouseClickEvents = True
MediaPlayer1.ShowTracker = False
MediaPlayer1.ShowControls = False
End If
End If
Exit Sub
Err:
Form1.Slider1.Value = "0"
Form1.Check6.Enabled = True
MediaPlayer1.CurrentPosition = Form1.Slider1.Value
MediaPlayer1.ClickToPlay = False
MediaPlayer1.EnableContextMenu = False
MediaPlayer1.SendMouseClickEvents = True
MediaPlayer1.ShowTracker = False
MediaPlayer1.ShowControls = False
Form1.Label6.Caption = ""
Form1.Label10.Caption = ""
Form1.Label25.Caption = "(No File)"
Form1.Label23.Caption = "(No File)"
MediaPlayer1.FileName = ""
Form1.Text2.Locked = False
Form1.Text2.Text = "0"
Form1.Text3.Text = "0"
Form1.Text2.Locked = True
Form1.Slider3.Value = "0"
Form1.Command1.Visible = True
Form1.Command13.Visible = False
Form1.mnutoc.Enabled = True
Form1.mnupas.Enabled = False
Form7.Hide
MediaPlayer1.Stop
MediaPlayer1.CurrentPosition = Form1.Slider3.Value
Exit Sub
End Sub


