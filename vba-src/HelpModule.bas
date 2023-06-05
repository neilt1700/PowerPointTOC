Attribute VB_Name = "HelpModule"
Option Explicit

Private Const msMODULE As String = "HelpModule"

Public Sub ShowHelp()

  Const sSOURCE As String = "ShowHelp"
  On Error GoTo ErrorHandler

  Dim stTempFilePath As String
  stTempFilePath = Environ$("TEMP") & "\Commtap Table Of Contents Help.txt"
 
  Dim fso, tempFile
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set tempFile = fso.CreateTextFile(FileName:=stTempFilePath, overwrite:=True, unicode:=True)
  
  tempFile.WriteLine HelpInfo

  tempFile.Close
  
  Shell Environ$("windir") & "\notepad.exe " & stTempFilePath, 1
     
  ' Notepad opens a 'copy' of the file - it does not 'lock' it
  ' so the file can be deleted without problems
  Kill stTempFilePath
  
  Exit Sub
    
ErrorHandler:
  If bCentralErrorHandler(msMODULE, sSOURCE) Then
    Stop
    Resume
  End If

End Sub

Private Function HelpInfo() As String

  Dim hlp As String
  
  hlp = "HOW TO CREATE A TABLE OF CONTENTS"
  hlp = hlp & vbCrLf & ""
  hlp = hlp & vbCrLf & "(1) Add titles into the title placeholder for each slide you want to be in the table of contents"
  hlp = hlp & vbCrLf & "(2) Make sure you are on the slide where you want a table of contents to appear"
  hlp = hlp & vbCrLf & "(3) Run the TableOfContents macro"
  hlp = hlp & vbCrLf & "(4) A table of contents will appear on the page"
  hlp = hlp & vbCrLf & ""
  hlp = hlp & vbCrLf & "UPDATING A TABLE OF CONTENTS"
  hlp = hlp & vbCrLf & ""
  hlp = hlp & vbCrLf & "(1) Make sure you are on a slide which already has a table of contents"
  hlp = hlp & vbCrLf & "(2) Click on ""Table Of Contents"""
  hlp = hlp & vbCrLf & ""
  hlp = hlp & vbCrLf & "CHANGING THE POSITION AND FORMATTING OF A TABLE OF CONTENTS"
  hlp = hlp & vbCrLf & ""
  hlp = hlp & vbCrLf & "(1) Select a table of contents"
  hlp = hlp & vbCrLf & "(2) Change position, size, font-size, and/or font"
  hlp = hlp & vbCrLf & "(3) Run the TableOfContents macro"
  hlp = hlp & vbCrLf & "Note, for text, the updated table will use the font and font size for the first cell in the original table."
  hlp = hlp & vbCrLf & ""
  hlp = hlp & vbCrLf & "HAVING MORE THAN ONE TABLE OF CONTENTS"
  hlp = hlp & vbCrLf & ""
  hlp = hlp & vbCrLf & "You might want to have more than one table of contents - for example a general table of contents and another one for appendices."
  hlp = hlp & vbCrLf & ""
  hlp = hlp & vbCrLf & "(1) Add in all the slides where you want a table of contents to appear"
  hlp = hlp & vbCrLf & "(2) On the last slide where you want a table of contents to appear, click on ""Table of Contents"". The contents list will start from the one after the current slide."
  hlp = hlp & vbCrLf & "(3) Do the same on an earlier slide where you want a Table of Contents. You will get a table of contents starting from the next slide and up to and including the next table of contents slide."
  hlp = hlp & vbCrLf & ""
  hlp = hlp & vbCrLf & "NOTES"
  hlp = hlp & vbCrLf & ""
  hlp = hlp & vbCrLf & "- If more than one slide has the same title, slides following the first slide with that title will not appear in the table."
  hlp = hlp & vbCrLf & ""
  
  
  
  HelpInfo = hlp


End Function
