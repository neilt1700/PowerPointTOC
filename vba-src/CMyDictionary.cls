VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMyDictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("CommtapSymboliser.DataClasses")

Option Explicit

' Commtap Symboliser - add pictures/symbols to text in a PowerPoint document
' Copyright (C) 2017 - 2022 Commtap CIC
' If you would like to bundle this software with your symbols product, please
' contact us at tap@commtap.org.

' Dictionary Wrapper

' Methods
' =======
'
' Method       Description
' ------       -----------
' Add          Adds a new key/item pair to a Dictionary object.
' Exists       Returns a Boolean value that indicates whether a specified key
'              exists in the Dictionary object.
' Items        Returns an array of all the items in a Dictionary object.
' Keys         Returns an array of all the keys in a Dictionary object.
' Remove       Removes one specified key/item pair from the Dictionary object.
' RemoveAll    Removes all the key/item pairs in the Dictionary object.
'
' Properties
' ==========
'
' Property     Description
' --------     -----------
' CompareMode  Sets or returns the comparison mode for comparing keys in a
'              Dictionary object.
' Count        Returns the number of key/item pairs in a Dictionary object.
' Item         Sets or returns the value of an item in a Dictionary object.
' Key          Sets a new key value for an existing key value in a Dictionary
'              object.

Private Const msMODULE As String = "CMyDictionary"

Private p_dcDictionary As Object

Public Property Let CompareMode(ByVal Compare As VbCompareMethod)

  Const sSOURCE As String = "CompareMode"
  On Error GoTo ErrorHandler

  p_dcDictionary.CompareMode = Compare
  
  Exit Property
    
ErrorHandler:
  ' Run simple clean-up code here
  If bCentralErrorHandler(msMODULE, sSOURCE) Then
    Stop
    Resume
  End If
  
End Property
Public Property Get CompareMode() As VbCompareMethod
  CompareMode = p_dcDictionary.CompareMode
End Property

Public Property Get Count() As Long
  Count = p_dcDictionary.Count
End Property

Public Property Let Item(ByVal Key As Variant, ByVal newitem As Variant)

  Const sSOURCE As String = "Item"
  On Error GoTo ErrorHandler

  p_dcDictionary.Item(Key) = newitem
  
  Exit Property
  
ErrorHandler:
  ' Run simple clean-up code here
  If bCentralErrorHandler(msMODULE, sSOURCE) Then
    Stop
    Resume
  End If
  
End Property

Public Property Get Item(ByVal Key As Variant) As Variant

  Const sSOURCE As String = "Item"
  On Error GoTo ErrorHandler

  If VarType(p_dcDictionary.Item(Key)) = vbObject Then
    Set Item = p_dcDictionary.Item(Key)
  Else
    Item = p_dcDictionary.Item(Key)
  End If
  
  Exit Property
  
ErrorHandler:
  ' Run simple clean-up code here
  If bCentralErrorHandler(msMODULE, sSOURCE) Then
    Stop
    Resume
  End If

End Property

Public Property Let Key(ByVal Key As Variant, ByVal newkey As Variant)

  Const sSOURCE As String = "Key"
  On Error GoTo ErrorHandler

  p_dcDictionary.Key(Key) = newkey
  
  Exit Property
  
ErrorHandler:
  ' Run simple clean-up code here
  If bCentralErrorHandler(msMODULE, sSOURCE) Then
    Stop
    Resume
  End If
  
End Property

Public Sub Add(Key As Variant, Item As Variant)

  Const sSOURCE As String = "Add"
  On Error GoTo ErrorHandler

  p_dcDictionary.Add Key, Item
  
  Exit Sub
  
ErrorHandler:
  ' Run simple clean-up code here
  If bCentralErrorHandler(msMODULE, sSOURCE) Then
    Stop
    Resume
  End If
  
End Sub

Public Function Exists(Key As Variant) As Boolean

  Const sSOURCE As String = "Exists"
  On Error GoTo ErrorHandler

  Exists = p_dcDictionary.Exists(Key)
  
  Exit Function
  
ErrorHandler:
  ' Run simple clean-up code here
  If bCentralErrorHandler(msMODULE, sSOURCE) Then
    Stop
    Resume
  End If
  
End Function

Public Function Items(Optional ByVal Index As Long = -1)

  Const sSOURCE As String = "Items"
  On Error GoTo ErrorHandler

  If Index = -1 Then
    Items = p_dcDictionary.Items
  Else
    Dim arrItems As Variant
    arrItems = p_dcDictionary.Items
    If IsObject(arrItems(Index)) Then
      Set Items = arrItems(Index)
    Else
      Items = arrItems(Index)
    End If
  End If
  
  Exit Function
  
ErrorHandler:
  ' Run simple clean-up code here
  If bCentralErrorHandler(msMODULE, sSOURCE) Then
    Stop
    Resume
  End If
  
End Function

Public Function Keys(Optional ByVal Index As Long = -1)
  
  Const sSOURCE As String = "Keys"
  On Error GoTo ErrorHandler

  If Index = -1 Then
    Keys = p_dcDictionary.Keys
  Else
    Dim arrKeys As Variant
    arrKeys = p_dcDictionary.Keys
    If IsObject(arrKeys(Index)) Then
      Set Keys = arrKeys(Index)
    Else
      Keys = arrKeys(Index)
    End If
  End If
  
  Exit Function
  
ErrorHandler:
  ' Run simple clean-up code here
  If bCentralErrorHandler(msMODULE, sSOURCE) Then
    Stop
    Resume
  End If
  
End Function

Public Sub Remove(Key As Variant)

  Const sSOURCE As String = "Remove"
  On Error GoTo ErrorHandler

  p_dcDictionary.Remove Key
  
  Exit Sub
  
ErrorHandler:
  ' Run simple clean-up code here
  If bCentralErrorHandler(msMODULE, sSOURCE) Then
    Stop
    Resume
  End If
  
End Sub


Public Sub InitiateProperties()
  Set p_dcDictionary = CreateObject("Scripting.Dictionary")
End Sub
