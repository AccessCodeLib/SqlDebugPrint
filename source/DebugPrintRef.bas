Attribute VB_Name = "DebugPrintRef"
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>%AppFolder%/source/DebugPrintRef.bas</file>
'</codelib>
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
'----------------------------------------------------------------------------------
' Modul     : DebugPrintRef
' Purpose   : hält Konstanten und die Referenz der Klasse "SqlDebugger" vor
'----------------------------------------------------------------------------------

Public Const MSGBOXTITLE As String = "SQL-Debugger"
Public Const SqlDebugPrintFactoryModuleName As String = "SqlDebugPrintFactory"
Public Const TEMPQUERYNAME As String = "qTempSqlDebugPrint"

Public Const SqlDebugPrintVersion As String = "3.1.0"
'3.1.0: Umgestellt auf accda (3.0 war noch mda)

#If VBA7 Then

Private Declare PtrSafe Function WNetGetConnection Lib "mpr.dll" Alias "WNetGetConnectionA" ( _
            ByVal lpszLocalName As String, _
            ByVal lpszRemoteName As String, _
            cbRemoteName As Long) As Long
#Else
Private Declare Function WNetGetConnection Lib "mpr.dll" Alias "WNetGetConnectionA" ( _
            ByVal lpszLocalName As String, _
            ByVal lpszRemoteName As String, _
            cbRemoteName As Long) As Long
#End If

Public SqlDebuggerTransferRef As SqlDebugger

Public Function NewTransferModule() As TransferCodeModule
    
    Set NewTransferModule = New TransferCodeModule
    
End Function

Public Function NewSqlDebugger() As SqlDebugger
    
    Set NewSqlDebugger = New SqlDebugger
    
End Function

Public Function StartSqlDebugPrint()
  
    DoCmd.OpenForm "frmStart"
    
End Function

Public Function UncPath( _
                ByVal Path As String, _
                Optional ByVal IgnoreErrors As Boolean = True) As String
   
   Dim UNC As String * 512
   
   If Len(Path) = 1 Then Path = Path & ":"
   
   If WNetGetConnection(Left$(Path, 2), UNC, Len(UNC)) Then
   
      ' API-Routine gibt Fehler zurück:
      If IgnoreErrors Then
         UncPath = Path
      Else
         Err.Raise 5 ' Invalid procedure call or argument
      End If
   Else
      ' Ergebnis zurückgeben:
      UncPath = Left$(UNC, InStr(UNC, vbNullChar) - 1) & Mid$(Path, 3)
   End If
   
End Function
