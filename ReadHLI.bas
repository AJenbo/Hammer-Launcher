Attribute VB_Name = "ReadHLI"
'set exe variable
Public MyDIR As String, HLIFile As String, x As Long, sSection As String, sString As String

Declare Function GetPrivateProfileString Lib "kernel32" Alias _
   "GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
   ByVal lpKeyName As Any, ByVal lpDefault As String, _
   ByVal lpReturnedString As String, ByVal nSize As Long, _
   ByVal lpFileName As String) As Long

Declare Function WritePrivateProfileString Lib "kernel32" Alias _
   "WritePrivateProfileStringA" (ByVal lpApplicationName As String, _
   ByVal lpKeyName As Any, ByVal lpString As Any, _
   ByVal lpFileName As String) As Long
