Attribute VB_Name = "Main"
Option Explicit

Private Const msMODULE As String = "ShapeFunctionsModule"
Public Const constTableOfContentsName As String = "TableOfContents"

' Entry point
Public Sub TableOfContents()

  Const sSOURCE As String = "TableOfContents"
  On Error GoTo ErrorHandler
     
  Dim stErrorMessage As String
  
  Dim slActiveSlide As Slide
  Set slActiveSlide = SelectedSlide(stErrorMessage)
  
  If stErrorMessage <> vbNullString Then
    MsgBox stErrorMessage, vbExclamation, "Table of Contents"
    Exit Sub
  End If
  
  Dim mydcContents As CMyDictionary
  Set mydcContents = TitlesAndSlideNumbers(slActiveSlide)
  
  PlaceTOC slActiveSlide, mydcContents
  
ErrorExit:
  Exit Sub

ErrorHandler:
  If bCentralErrorHandler(msMODULE, sSOURCE, , True) Then
    Stop
    Resume
  Else
    Resume ErrorExit
  End If

End Sub

Private Function SelectedSlide(ByRef stErrorMessage As String) As Slide

  Const sSOURCE As String = "SelectedSlide"
  On Error GoTo ErrorHandler
   
   With Application.ActiveWindow
    If .View.Type <> ppViewNormal Then
      stErrorMessage = _
        "No slide is active: please go to ""Normal View"" and select a slide."
      Exit Function
    End If
  End With
  
  On Error GoTo NoSlideError
  
  Set SelectedSlide = Application.ActiveWindow.View.Slide
  
  On Error GoTo ErrorHandler
  
  Exit Function
  
NoSlideError:

  stErrorMessage = _
    "No slide is active. Please select a slide where you would like the " & _
    "table of contents to appear, or where you would like to update an " & _
    "existing table of contents."
        
  Exit Function
    
ErrorHandler:
  ' Run simple clean-up code here
  If bCentralErrorHandler(msMODULE, sSOURCE) Then
    Stop
    Resume
  End If

End Function


' Entry point
Public Sub Help()

  Const sSOURCE As String = "Help"
  On Error GoTo ErrorHandler
   
  HelpModule.ShowHelp
   
ErrorExit:
  Exit Sub

ErrorHandler:
  If bCentralErrorHandler(msMODULE, sSOURCE, , True) Then
    Stop
    Resume
  Else
    Resume ErrorExit
  End If

End Sub

Public Function TitlesAndSlideNumbers(ByRef slActiveSlide As Slide) As CMyDictionary

  Const sSOURCE As String = "TitlesAndSlideNumbers"
  On Error GoTo ErrorHandler

  Dim i As Long
  Dim stTitle As String
  Dim stPreviousTitle As String
  
  Dim slSlide As Slide
  
  Dim mydcContents As CMyDictionary
  Set mydcContents = Factory.CreateCMyDictionary
  
  With Application.ActivePresentation
    For i = slActiveSlide.SlideIndex + 1 To .Slides.Count
      Set slSlide = .Slides(i)
      With .Slides(i).Shapes
      If .HasTitle Then
        If .Title.TextFrame.HasText Then
          stTitle = .Title.TextFrame.TextRange.Text
          If stTitle <> stPreviousTitle Then
            mydcContents.Add CStr(i), stTitle
          End If
          
          ' Creates table of contents for slides up until the next table of
          ' contents.
          If HasShape(constTableOfContentsName, slSlide) Then
            Exit For
          End If
          
          stPreviousTitle = .Title.TextFrame.TextRange.Text
        End If
      End If
      End With
    Next i
  End With

  Set TitlesAndSlideNumbers = mydcContents

  Exit Function
    
ErrorHandler:
  ' Run simple clean-up code here
  If bCentralErrorHandler(msMODULE, sSOURCE) Then
    Stop
    Resume
  End If

End Function


Public Sub PlaceTOC(ByRef slSlide As Slide, ByRef mydcContents As CMyDictionary)

  Const sSOURCE As String = "PlaceTOC"
  On Error GoTo ErrorHandler
  
  Dim arrContentsKeys
  arrContentsKeys = mydcContents.Keys
  
  Dim tblTOC As Table
  Set tblTOC = TOCTable(slSlide, mydcContents.Count, 2)

  If Not tblTOC Is Nothing Then

    Dim i As Long
    For i = 0 To UBound(arrContentsKeys)
      tblTOC.Rows(i + 1).Cells(1).Shape.TextFrame.TextRange.Text = mydcContents.Item(arrContentsKeys(i))
      HyperLinkToSlide _
        tblTOC.Rows(i + 1).Cells(1).Shape, _
        arrContentsKeys(i), _
        tblTOC.Rows(i + 1).Cells(1).Shape.TextFrame.TextRange.Text, _
        slSlide
      tblTOC.Rows(i + 1).Cells(2).Shape.TextFrame.TextRange.Text = arrContentsKeys(i)
    Next i
  
  End If
  
  Exit Sub
    
ErrorHandler:
  ' Run simple clean-up code here
  If bCentralErrorHandler(msMODULE, sSOURCE) Then
    Stop
    Resume
  End If

End Sub

Public Sub HyperLinkToSlide(ByRef shTableCellShape As Shape, _
                            ByVal lngSlideToLinkTo As Long, _
                            ByVal stSlideTitle As String, _
                            ByVal slTOCSlide As Slide)

  Const sSOURCE As String = "AddHyperLinkToSlide"
  On Error GoTo ErrorHandler

  Dim lngSlideID As Long
  With Application.ActivePresentation.Slides(lngSlideToLinkTo)
    lngSlideID = .SlideID
  End With
  
  Dim dblCellHeight As Double
  Dim dblCellWidth As Double
  Dim dblCellTop As Double
  Dim dblCellLeft As Double
  
  With shTableCellShape
    dblCellHeight = .Height
    dblCellWidth = .Width
    dblCellTop = .Top
    dblCellLeft = .Left
  End With
  
  Dim shHyperLinkShape As Shape
  Set shHyperLinkShape = slTOCSlide.Shapes.AddShape( _
    msoShapeRectangle, _
    dblCellLeft, _
    dblCellTop, _
    dblCellWidth, _
    dblCellHeight)
  
  With shHyperLinkShape
    .Fill.Visible = msoFalse
    .Line.Visible = msoFalse
  End With
  
  With shHyperLinkShape.ActionSettings(ppMouseClick).Hyperlink
    .SubAddress = lngSlideID & "," & lngSlideToLinkTo & "," & stSlideTitle
  End With
  
  Exit Sub
    
ErrorHandler:
  ' Run simple clean-up code here
  If bCentralErrorHandler(msMODULE, sSOURCE) Then
    Stop
    Resume
  End If
  
End Sub


Public Function TOCTable(ByRef slSlide As Slide, _
                         ByVal numRows As Long, _
                         ByVal numColumns As Long) As Table

  Const sSOURCE As String = "TOCTable"
  
  On Error GoTo ErrorHandler
  
  If numRows < 1 Or numColumns < 1 Then
    Exit Function
  End If
  
  Dim shTOCShape As Shape
  Set shTOCShape = ShapeFunctionsModule.TableShapeNamed(constTableOfContentsName, slSlide)
  
  ' Start with some defaults for positioning the TOC
  Dim dblLeft As Double
  dblLeft = DefaultTOCLeft
  Dim dblTop As Double
  dblTop = DefaultTOCTop
  Dim dblWidth As Double
  dblWidth = DefaultTOCWidth
  
  Dim mydcCellProperties As CMyDictionary
  Set mydcCellProperties = Factory.CreateCMyDictionary
  
  With mydcCellProperties
    .Add "FontName", DefaultFontName()
    .Add "FontSize", DefaultFontSize()
    .Add "WidthColumn1", 0
    .Add "WidthColumn2", 0
  End With
  
  If Not shTOCShape Is Nothing Then
    dblLeft = shTOCShape.Left
    dblTop = shTOCShape.Top
    dblWidth = shTOCShape.Width
    
    If shTOCShape.Table.Rows.Count > 0 Then
      With shTOCShape.Table.Rows(1).Cells(1).Shape.TextFrame.TextRange.Font
        mydcCellProperties.Item("FontName") = .Name
        mydcCellProperties.Item("FontSize") = .Size
      End With
      
      Dim dblTotalWidth As Double
      dblTotalWidth = TableWidth(shTOCShape.Table)

      mydcCellProperties.Item("WidthColumn1") = shTOCShape.Table.Columns(1).Width
      If shTOCShape.Table.Rows(1).Cells.Count > 1 Then
        mydcCellProperties.Item("WidthColumn2") = shTOCShape.Table.Columns(2).Width
      End If
    End If
       
  End If
  
  ShapeFunctionsModule.RemoveTableShapesNamed constTableOfContentsName, slSlide
  
  Dim shTOC As Shape
  Set shTOC = slSlide.Shapes.AddTable(numRows, numColumns, dblLeft, dblTop, dblWidth)
  shTOC.Name = constTableOfContentsName
  
  Dim tblTOC As Table
  Set tblTOC = shTOC.Table
  
  tblTOC.FirstRow = False
  SetTableFormatting tblTOC, mydcCellProperties
  
  Set TOCTable = tblTOC
  
  Exit Function
    
ErrorHandler:
  ' Run simple clean-up code here
  If bCentralErrorHandler(msMODULE, sSOURCE) Then
    Stop
    Resume
  End If


End Function

Public Sub SetTableFormatting(ByRef tblTable As Table, ByRef mydcCellProperties As CMyDictionary)

  Const sSOURCE As String = "SetTableFormatting"
  On Error GoTo ErrorHandler

  Dim oRow As Row
  Dim oCell As Cell

  Dim dblTotalWidth As Double
  dblTotalWidth = TableWidth(tblTable)

  If mydcCellProperties.Item("WidthColumn1") > 0 Then
    tblTable.Columns(1).Width = mydcCellProperties.Item("WidthColumn1")
  Else
    tblTable.Columns(1).Width = 0.8 * dblTotalWidth
  End If
  
  If mydcCellProperties.Item("WidthColumn2") > 0 Then
    tblTable.Columns(2).Width = mydcCellProperties.Item("WidthColumn2")
  Else
    tblTable.Columns(2).Width = 0.2 * dblTotalWidth
  End If

  For Each oRow In tblTable.Rows
    For Each oCell In oRow.Cells
      oCell.Shape.Fill.Transparency = 1
      With oCell.Shape.TextFrame.TextRange.Font
        .Name = mydcCellProperties.Item("FontName")
        .Size = mydcCellProperties.Item("FontSize")
      End With
    Next oCell
    
    oRow.Height = 1
    oRow.Cells(2).Shape.TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignRight
    
  Next oRow
    
  Exit Sub
    
ErrorHandler:
  ' Run simple clean-up code here
  If bCentralErrorHandler(msMODULE, sSOURCE) Then
    Stop
    Resume
  End If

End Sub


