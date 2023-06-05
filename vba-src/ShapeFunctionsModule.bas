Attribute VB_Name = "ShapeFunctionsModule"
Option Explicit

Private Const msMODULE As String = "ShapeFunctionsModule"

Public Function HasShape(ByVal sShapeName As String, oSl As Slide) As Boolean

  Const sSOURCE As String = "HasShape"
  On Error GoTo ErrorHandler

  Dim oSh As Shape
  For Each oSh In oSl.Shapes
    If oSh.Name = sShapeName Then
      HasShape = True
      Exit Function
    End If
  Next
  
  Exit Function
    
ErrorHandler:
  ' Run simple clean-up code here
  If bCentralErrorHandler(msMODULE, sSOURCE) Then
    Stop
    Resume
  End If
  
End Function

' Returns a shape having a name sShapeName on slide oSl
Public Function ShapeNamed(ByVal sShapeName As String, oSl As Slide) As Shape

  Const sSOURCE As String = "ShapeNamed"
  On Error GoTo ErrorHandler

  Dim oSh As Shape
  For Each oSh In oSl.Shapes
    If oSh.Name = sShapeName Then
      Set ShapeNamed = oSh
      Exit Function
    End If
  Next
  
  Exit Function
    
ErrorHandler:
  ' Run simple clean-up code here
  If bCentralErrorHandler(msMODULE, sSOURCE) Then
    Stop
    Resume
  End If
  
End Function

Public Function TableShapeNamed(ByVal sShapeName As String, oSl As Slide) As Shape

  Const sSOURCE As String = "TableShapeNamed"
  On Error GoTo ErrorHandler

  Dim oSh As Shape
  For Each oSh In oSl.Shapes
    If oSh.HasTable Then
      If oSh.Name = sShapeName Then
        Set TableShapeNamed = oSh
      End If
    End If
  Next
  
  Exit Function
    
ErrorHandler:
  ' Run simple clean-up code here
  If bCentralErrorHandler(msMODULE, sSOURCE) Then
    Stop
    Resume
  End If
  
End Function


Public Function RemoveTableShapesNamed(ByVal sShapeName As String, oSl As Slide) As Table

  Const sSOURCE As String = "RemoveTableShapesNamed"
  On Error GoTo ErrorHandler

  Dim oSh As Shape
  For Each oSh In oSl.Shapes
    If oSh.HasTable Then
      If oSh.Name = sShapeName Then
        oSh.Delete
      End If
    End If
  Next
  
  Exit Function
    
ErrorHandler:
  ' Run simple clean-up code here
  If bCentralErrorHandler(msMODULE, sSOURCE) Then
    Stop
    Resume
  End If
  
End Function

