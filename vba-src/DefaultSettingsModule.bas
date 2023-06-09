Attribute VB_Name = "DefaultSettingsModule"
Option Explicit

' Copyright Commtap CIC 2023

Private Const msMODULE As String = "DefaultSettingsModule"
Public Const constTableOfContentsName As String = "TableOfContents"

Public Property Get DefaultFontName() As String
  DefaultFontName = "Cavolini"
End Property

Public Property Get DefaultFontSize() As Double
  DefaultFontSize = 12
End Property

Public Property Get DefaultTOCLeft() As Double
  DefaultTOCLeft = 51.56 ' For A4 portrait
End Property

Public Property Get DefaultTOCTop() As Double
  DefaultTOCTop = 118.63 ' For A4 portrait
End Property

Public Property Get DefaultTOCWidth() As Double
  DefaultTOCWidth = 436.88 ' For A4 portrait
End Property

