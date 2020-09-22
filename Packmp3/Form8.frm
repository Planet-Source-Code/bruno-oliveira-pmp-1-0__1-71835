VERSION 5.00
Begin VB.Form Form8 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "                      Password"
   ClientHeight    =   1920
   ClientLeft      =   5055
   ClientTop       =   540
   ClientWidth     =   4875
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   4875
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text2 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1320
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1080
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   360
      Width           =   3255
   End
   Begin VB.Label Label2 
      Caption         =   "Pass"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "User"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   735
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private DynPass As Recordset
Private DynConn As Connection
Dim a As Integer

Private Sub Form_Load()
Dim strName As String
Dim lngBuffer As Long
  strName = String$(255, 0)
  'lngBuffer = GetUserName(strName, Len(strName))
  Text1.Text = strName
x = LoadDbase()
If x <> True Then MsgBox "Error: In loading USername and Password Data -  " & error

'lngI = SetFocuses(Me.hwnd)
End Sub

Private Function LoadDbase() As Boolean
On Error Resume Next
' Loads the Database
Set DynConn = New Connection ' note instead  of Entering Dim names as new object, it's good practice to Declare _
Dim on the top of the line. With out the new statement. And use Set Dclared = new Object
'Ex. :
'Dim Connections as connection
' Set Connectiosn = new connections.


Set DynPass = New Recordset

DynConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.3.51;Data Source=C:\Documents and Settings\Bruno\Os meus documentos\Códigos do VB\Os Meus Códigos\Packmp3\Packmp3\ExUsers.mdb;"
DynConn.Open

DynPass.Open "Select * from DatAccount", DynConn, adOpenDynamic, adLockPessimistic

If DynConn.State <> adStateOpen Then LoadDbase = False
If DynPass.State <> adStateOpen Then LoadDbase = False
If DynConn.State = adStateOpen And DynPass.State = adStateOpen Then LoadDbase = True

End Function

Private Function FindUser(ByVal UName As String) As String
On Error Resume Next
DynPass.Find "strUser = '" & UName & "'"
If DynPass.EOF = True Or DynPass.BOF = True Then
   DynPass.MoveFirst
   DynPass.AddNew
   DynPass!strUser = UName
   DynPass!strPass = Text2.Text
   DynPass.Update
   MsgBox "New Password Set"
   FindUser = Text2
   
   
   
   Exit Function
   Else
   FindUser = DynPass!strPass
End If


If Err > 0 Then MsgBox "Error In database: " & vbCrLf & error




End Function

Private Sub text2_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
If FindUser(Text1.Text) = "NOTSET" Then
MsgBox "Username not found"
End

End If
     a = a + 1
     If Text2.Text = FindUser(Text1.Text) Then
     Load Form1
     Form1.Show
     Me.Hide
     Unload Me
     Else
     If a < 4 Then
     e = 4 - a
     MsgBox "Username / Password Restam-lhe" & " " & e & " " & "tentativas"
     Text2.Text = ""
     Else
     End
     End If
     End If
     
End If
End Sub
