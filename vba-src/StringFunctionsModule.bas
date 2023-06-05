Attribute VB_Name = "StringFunctionsModule"
Option Explicit

Private Const msMODULE As String = "StringFunctionsModule"


Public Function StartsWith(ByVal stSubject As String, ByVal stSearch As String) As Boolean

  Const sSOURCE As String = "StartsWith"
  On Error GoTo ErrorHandler

  If Mid$(stSubject, 1, Len(stSearch)) = stSearch Then
    StartsWith = True
  End If
  
  Exit Function
    
ErrorHandler:
  ' Run simple clean-up code here
  If bCentralErrorHandler(msMODULE, sSOURCE) Then
    Stop
    Resume
  End If
  
End Function
