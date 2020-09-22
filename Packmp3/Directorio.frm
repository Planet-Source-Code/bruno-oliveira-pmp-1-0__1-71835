VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H80000008&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Add Directory"
   ClientHeight    =   4680
   ClientLeft      =   4890
   ClientTop       =   315
   ClientWidth     =   3285
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   3285
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000009&
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Adds to the playlist all the songs of the up paste"
      Top             =   3840
      Width           =   2895
   End
   Begin VB.DirListBox Dir1 
      BackColor       =   &H80000013&
      Height          =   3015
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   2895
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H8000000D&
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
   Begin VB.FileListBox File1 
      Height          =   1065
      Left            =   240
      Pattern         =   "|*.wma;*.wav;*.mp3;*.mpeg;*.mpg;*.avi|"
      TabIndex        =   2
      Top             =   1920
      Width           =   2895
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
If File1.ListCount <> 0 Then
    For a = 1 To File1.ListCount
        File1.ListIndex = a - 1
        b = Len(File1.FileName)
        t = b - 4
        c = Left(File1.FileName, t)
        Form1.List2.AddItem (Form1.List2.ListCount + 1 & "-" & c)
        Form1.List1.AddItem Dir1.Path & "\" & File1.FileName
        z = Dir1.Path & "\" & File1.FileName
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set att = fso.getfile(z)
        If Drive1.Drive <> "f:" And Drive1.Drive <> "e:" Then
        With Form1
        .read.Enabled = True
        .Hidden.Enabled = True
        .archive.Enabled = True
        .system.Enabled = True
        .system.Value = 1
        .Hidden.Value = 1
        .archive.Value = 0
        .read.Value = 0
         If .read.Value = 0 And .Hidden.Value = 1 And .archive.Value = 0 And .system.Value = 1 Then
          att.Attributes = 6
         End If
        End With
        End If
       Form1.Label17.Caption = Form1.List2.ListCount
    Next a
    Unload Me
Else
MsgBox "No files were found in specific folder", vbOKOnly, "Error"
Unload Me
End If
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
On Error Resume Next
Dir1.Path = Drive1.Drive
Exit Sub
End Sub

