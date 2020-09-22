Attribute VB_Name = "ID3"
'
'Prepare the variables for holding the tag information
Public HTag As Boolean
'Set the tag holding strings to the right length to get the right file info.
Public Tagged As String * 3
Public Title As String * 30
Public Artist As String * 30
Public Album As String * 30
Public Dated As String * 4
Public Comment As String * 30
'Call this function to check a tag
Public Function GetTag(File)
'
'Prepare TagFile to act as our file handler
Dim TagFile As Integer
TagFile = FreeFile
On Error GoTo hasnt
'Open the file
Open File For Binary As #TagFile
'Check to see if it has a tag - i.e. Tagged will equal "TAG"
Get #TagFile, FileLen(File) - 127, Tagged
If Tagged = "TAG" Then GoTo has Else GoTo hasnt 'go to the appropriate section

has:
'Set HTag to show a Tag exists
HTag = True
'Set the other values to show the Title, Artist, Album, Date and Comment of the track - Get the info. from the openned file
Get #TagFile, , Title
Get #TagFile, , Artist
Get #TagFile, , Album
Get #TagFile, , Dated
Get #TagFile, , Comment
'Finished getting info

Title = Trim(Title)
Artist = Trim(Artist)
Album = Trim(Album)
Dated = Trim(Dated)
Comment = Trim(Comment)

GoTo done

hasnt:
'Set HTag to show a Tag does not exist
HTag = False
'Set values to null to show no tag was found
Title = ""
Artist = ""
Album = ""
Dated = ""
Comment = ""
'Finished
GoTo done

done:
'Close the openned file
Close #TagFile
End Function






