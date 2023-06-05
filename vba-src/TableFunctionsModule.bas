Attribute VB_Name = "TableFunctionsModule"
Option Explicit

Private Const msMODULE As String = "TableFunctionsModule"

Public Function TableWidth(ByRef tblTable As Table) As Double

  Const sSOURCE As String = "TableWidth"
  On Error GoTo ErrorHandler
  
  Dim oColumn As Column
  
  For Each oColumn In tblTable.Columns
    TableWidth = TableWidth + oColumn.Width
  Next oColumn

Exit Function
    
ErrorHandler:
  ' Run simple clean-up code here
  If bCentralErrorHandler(msMODULE, sSOURCE) Then
    Stop
    Resume
  End If


End Function

