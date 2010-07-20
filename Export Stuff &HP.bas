Option Explicit

' This script creates a text file from IPTC information for all currently selected files
' If you want to change the output of this script, search for PrintField and read the
' comments above this command


Dim gblFH As Long
Dim gblStream As New IPTCStream

Sub Main
	' Get the image selection from the active folder
	Dim imgs As Images
	Dim db As Database
	Dim DriveLetter As String
	Set db = Application.ActiveDatabase
	If db Is Nothing Then
		MsgBox "Please open a database first!"
		Exit Sub
	End If

	Set imgs = db.ActiveSelection
	If imgs Is Nothing Then
		MsgBox "Please select some images first!"
		Exit Sub
	End If

	DriveLetter = db.ActiveFolder.Path()
	DriveLetter = Left$(DriveLetter,1)

	' Open a standard file dialog box and let the user select or enter a file name'
	Dim filename As String
	'filename = GetFilePath("*.txt","txt", , "Please choose a file", 3+4)

	' If the dialog box is aborted, we do nothing and return
	'If (filename = "") Then
	'	Exit Sub
	'End If

	filename = DriveLetter & ":\Users\Horst\Pictures\DB\Scripts\HardLink\link.txt"

	' Open/Create the file
	gblFH = FreeFile
	Open filename For Output As gblFH
    Print #gblFH, DriveLetter;vbCrLf;

	' Enable the percentage status bar
	Application.StatusBarShowPercentage 0, imgs.Count
	Application.StatusBarSetText "Exporting images"

	' Iterate over all images
	Dim i As Image
	Dim imgcnt As Long
	imgcnt = 0

	Dim res As IMIPTCResult

	Dim cat As Category
	Dim catstr As String
	Dim r As Integer
	Dim l As String

	For Each i In imgs
		Print #gblFH, i.FileName;vbTab;
		catstr = ""
		For Each cat In i.Categories
			catstr += cat.FullName & ","
		Next
    	Print #gblFH, catstr;vbTab;

		r = IMatch.ActiveDatabase.GetRating(i)
		Print #gblFH, Str(r);vbTab;

		l = IMatch.ActiveDatabase.GetLabel(i)
		Print #gblFH, l;

        imgcnt = imgcnt+1
	    Application.StatusBarSetPercentage imgcnt
		Print #gblFH, ""
	Next


	Application.StatusBarHidePercentage

	' Close the file
	Close gblFH

	' Open the file with notepad
	Shell Environ$("COMSPEC") & " /k " & "C:\Python26\python.exe " & DriveLetter & ":\Users\Horst\Pictures\DB\Scripts\HardLink\iMatchHardlink.py" & " " & DriveLetter, vbNormalFocus
    'Shell  DriveLetter & ":\Users\Horst\Pictures\DB\Scripts\HardLink\hl.bat", vbNormalFocus

	'Shell "notepad.exe " & filename, vbNormalFocus

End Sub


Sub PrintField (FieldName As String)

	If Not gblStream.Fields(FieldName) Is Nothing Then
		Dim values As Variant
		values = gblStream.Fields(FieldName)
		Dim vi As Integer
		For vi = LBound(values) To UBound(values)
			Print #gblFH, values(vi);
		Next vi
	End If
	Print #gblFH, vbTab;
End Sub

