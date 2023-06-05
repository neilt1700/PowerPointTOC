Attribute VB_Name = "Factory"
Option Explicit


Public Function CreateCMyDictionary() As CMyDictionary
  Set CreateCMyDictionary = New CMyDictionary
  CreateCMyDictionary.InitiateProperties
End Function
