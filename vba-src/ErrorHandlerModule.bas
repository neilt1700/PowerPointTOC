Attribute VB_Name = "ErrorHandlerModule"
'@Folder("CommtapSymboliser.ErrorHandling")

'
' Description:  This module contains the central error
'               handler and related constant declarations.
'
' Authors:      Rob Bovey, www.appspro.com
'               Stephen Bullen, www.oaltd.co.uk
'
' Chapter Change Overview
' Ch#   Comment
' --------------------------------------------------------------
' 15    Initial version
'
Option Explicit
Option Private Module

' **************************************************************
' Global Constant Declarations Follow
' **************************************************************
Public Const gbDEBUG_MODE As Boolean = True    ' True enables debug mode, False disables it.
Public Const glHANDLED_ERROR As Long = 9999     ' Run-time error number for our custom errors.
Public Const glUSER_CANCEL As Long = 18         ' The error number generated when the user cancels program execution.


' **************************************************************
' Module Constant Declarations Follow
' **************************************************************
Private Const msSILENT_ERROR As String = "UserCancel"   ' Used by the central error handler to bail out silently on user cancel.
'Private Const msFILE_ERROR_LOG As String = "CommtapSymboliser"  ' The name of the file where error messages will be logged to.


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments: This is the central error handling procedure for the
'           program. It logs and displays any run-time errors
'           that occur during program execution.
'
' Arguments:    sModule         The module in which the error occured.
'               sProc           The procedure in which the error occured.
'               sFile           (Optional) For multiple-workbook
'                               projects this is the name of the
'                               workbook in which the error occured.
'               bEntryPoint     (Optional) True if this call is
'                               being made from an entry point
'                               procedure. If so, an error message
'                               will be displayed to the user.
'
' Returns:      Boolean         True if the program is in debug
'                               mode, False if it is not.
'
' Date          Developer       Chap    Action
' --------------------------------------------------------------
' 03/30/08      Rob Bovey       Ch15    Initial version
'
Public Function bCentralErrorHandler( _
       ByVal sModule As String, _
       ByVal sProc As String, _
       Optional ByVal sFile As String, _
       Optional ByVal bEntryPoint As Boolean, _
       Optional ByVal bReThrow As Boolean = True) As Boolean

  Static sErrMsg As String
    
  Dim lngFile As Long
  Dim lErrNum As Long
  Dim sFullSource As String
  Dim stErrorDescription As String
    
  ' Grab the error info before it's cleared by
  ' On Error Resume Next below.
  lErrNum = Err.Number
  ' If this is a user cancel, set the silent error flag
  ' message. This will cause the error to be ignored.
  If lErrNum = glUSER_CANCEL Then
    sErrMsg = msSILENT_ERROR
  End If
  ' If this is the originating error, the static error
  ' message variable will be empty. In that case, store
  ' the originating error message in the static variable.
  If Len(sErrMsg) = 0 Then
    sErrMsg = Err.Description
  End If
    
  ' We cannot allow errors in the central error handler.
  On Error Resume Next
    
  ' Load the default filename if required.
  If Len(sFile) = 0 Then
    sFile = ActivePresentation.Name
  End If
    
  ' Construct the fully-qualified error source name.
  sFullSource = "[" & sFile & "]" & sModule & "." & sProc
                           
  ' Do not display or debug silent errors.
  If sErrMsg <> msSILENT_ERROR Then
    
    ' Show the error message when we reach the entry point
    ' procedure or immediately if we are in debug mode.
    If bEntryPoint Or gbDEBUG_MODE Then
      MsgBox sErrMsg & " Source: " & sFullSource, vbOKCancel + vbCritical, gsAPP_NAME
      ' Clear the static error message variable once
      ' we've reached the entry point so that we're ready
      ' to handle the next error.
      sErrMsg = vbNullString
    End If
        
    ' The return value is the debug mode status.
    bCentralErrorHandler = gbDEBUG_MODE
        
  Else
    ' If this is a silent error, clear the static error
    ' message variable when we reach the entry point.
    If bEntryPoint Then sErrMsg = vbNullString
    bCentralErrorHandler = False
  End If
    
  ' If we're using re-throw error handling, this is not the entry point and
  ' we're not debugging. Re-raise the error to be caught in the next
  ' procedure up the call stack. Procedures that handle their own errors can
  ' call the central error handler with bReThrow:=False to log the error but
  ' not re-raise it.
  If bReThrow Then
    If Not bEntryPoint And Not gbDEBUG_MODE Then
      On Error GoTo 0
      Err.Raise lErrNum, sFullSource, sErrMsg
    End If
  Else
    ' Error is being logged and handled, so clear the static error message
    ' variable.
    sErrMsg = vbNullString
  End If
    
End Function
