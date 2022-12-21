VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ErrorTrappingObserver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>%AppFolder%/source/ErrorTrappingObserver.cls</file>
'</codelib>
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit

Public Enum ErrorTrapping
    BreakOnAllErrors = 0
    BreakInClassModule = 1
    BreakOnUnhandledErrors = 2
End Enum

Private m_InitialErrorTrapping As ErrorTrapping
Private m_CurrentErrorTrapping As ErrorTrapping

Private Sub Class_Initialize()
   m_InitialErrorTrapping = Application.GetOption("Error Trapping")
End Sub

Private Sub Class_Terminate()
   RestoreErrorTrapping
End Sub

Public Property Get ErrorTrapping() As ErrorTrapping
   m_CurrentErrorTrapping = Application.GetOption("Error Trapping")
   ErrorTrapping = m_CurrentErrorTrapping
End Property

Public Property Let ErrorTrapping(ByVal NewValue As ErrorTrapping)
    If ErrorTrapping <> NewValue Then
        Application.SetOption "Error Trapping", NewValue
    End If
End Property

Private Sub RestoreErrorTrapping()
    If m_InitialErrorTrapping <> m_CurrentErrorTrapping Then
        Application.SetOption "Error Trapping", m_InitialErrorTrapping
    End If
End Sub
