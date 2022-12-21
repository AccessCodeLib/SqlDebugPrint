VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TransferCodeModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>%AppFolder%/source/TransferCodeModule.cls</file>
'</codelib>
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit

Const ErrNumber91_ObjectVariableOrWithBlockVariableNotSet As Long = 91

#Const EARLYBINDING = 0
#Const DEBUGMODE = 0

#If EARLYBINDING = 1 Then
#Else

   Public Enum vbext_ComponentType
      vbext_ct_StdModule = 1
      vbext_ct_ClassModule = 2
      vbext_ct_MSForm = 3
      vbext_ct_ActiveXDesigner = 11
      vbext_ct_Document = 100
   End Enum
   
#End If

#If EARLYBINDING = 1 Then
Private Property Get CurrentVbProject() As VBIDE.VBProject
#Else
Private Property Get CurrentVbProject() As Object
#End If

#If EARLYBINDING = 1 Then
   Dim Proj As VBProject
   Dim Temp As VBProject
#Else
   Dim Proj As Object
   Dim Temp As Object
#End If
   Dim strCurrentDbName As String

   Set Proj = VBE.ActiveVBProject
   'Prüfen, ob das richtige VbProject gewählt wurde (muss das von CurrentDb sein)
   strCurrentDbName = UncPath(CurrentDb.Name)
   If Proj.filename <> strCurrentDbName Then
      Set Proj = Nothing
      For Each Temp In VBE.VBProjects
         If Temp.filename = strCurrentDbName Then
            Set Proj = Temp
            Exit For
         End If
      Next
   End If
   
   Set CurrentVbProject = Proj

End Property

Public Sub CreateModuleInApp()

On Error GoTo Fehler

#If EARLYBINDING = 1 Then
    Dim cm As CodeModule
    Dim vbc As VBComponent
    Dim vbp As VBProject
#Else
    Dim cm As Object
    Dim vbc As Object
    Dim vbp As Object
#End If
    
    If CodeModuleExists(SqlDebugPrintFactoryModuleName) Then
        Set vbc = CurrentVbProject.VBComponents(SqlDebugPrintFactoryModuleName)
    Else 'Modul anlegen und benennen
        Set vbc = CurrentVbProject.VBComponents.Add(vbext_ct_StdModule)
        vbc.Name = SqlDebugPrintFactoryModuleName
    End If

   'Code einfügen
    Set cm = vbc.CodeModule
    cm.DeleteLines 1, cm.CountOfLines
    cm.AddFromString GetModuleString
    
    'Leerzeilen am Ende säubern
    Do While Len(Trim$(cm.Lines(cm.CountOfLines, 1))) = 0 And cm.CountOfLines > 1
        cm.DeleteLines cm.CountOfLines
    Loop
    
    Call SaveModule
    
Ende_Cleanup:
    On Error Resume Next
    Set cm = Nothing
    Set vbc = Nothing
    Set vbp = Nothing
    Exit Sub

Fehler:
    If Err.Number = ErrNumber91_ObjectVariableOrWithBlockVariableNotSet Then
        RaiseError ErrNumber91_ObjectVariableOrWithBlockVariableNotSet, "TransferCodeModule.CreateModuleInApp", Err.Source, "Keine Datenbank geöffnet!"
    Else
        RaiseError Err.Number, "TransferCodeModule.CreateModuleInApp", Err.Source, Err.Description
    End If
    Resume Ende_Cleanup
    
End Sub

Private Sub SaveModule()
    DoCmd.Save acModule, SqlDebugPrintFactoryModuleName
    DoCmd.Close acModule, SqlDebugPrintFactoryModuleName
End Sub

Public Sub DeleteModuleFromApp()

On Error GoTo Fehler

#If EARLYBINDING = 1 Then
    Dim vbc As VBComponent
#Else
    Dim vbc As Object
#End If
    
    Set vbc = CurrentVbProject.VBComponents(SqlDebugPrintFactoryModuleName)
    CurrentVbProject.VBComponents.Remove vbc
    
Ende_Cleanup:
    On Error Resume Next
    Set vbc = Nothing
    Exit Sub

Fehler:
    If Err.Number = ErrNumber91_ObjectVariableOrWithBlockVariableNotSet Then
        RaiseError ErrNumber91_ObjectVariableOrWithBlockVariableNotSet, "TransferCodeModule.DeleteModuleFromApp", Err.Source, "Keine Datenbank geöffnet!"
    Else
        RaiseError Err.Number, "TransferCodeModule.DeleteModuleFromApp", Err.Source, Err.Description
    End If
    Resume Ende_Cleanup

End Sub

Public Function CodeModuleExists( _
                ByVal CodeModulName As String) As Boolean

#If EARLYBINDING = 1 Then
   Dim VbcCol As VBIDE.VBComponents
   Dim vbc As VBIDE.VBComponent
#Else
   Dim VbcCol As Object
   Dim vbc As Object
#End If
   
   Set VbcCol = CurrentVbProject.VBComponents
   For Each vbc In VbcCol
      If vbc.Name = CodeModulName Then
         CodeModuleExists = True
         Exit For
      End If
   Next
   
End Function

Private Sub RaiseError( _
                    ByVal ErrNumber As Long, _
                    ByVal ErrSource As String, _
                    ByVal ErrSubSource As String, _
                    ByVal ErrDescription As String)

    If ErrSubSource <> "sqlDebugPrint" Then
        ErrSource = ErrSource & "->" & ErrSubSource
    End If
    Err.Raise ErrNumber, ErrSource, ErrDescription

End Sub
                    
Private Function GetModuleString() As String

    Dim str As String
    
    str = "Option Compare Database" & vbCrLf _
        & "Option Explicit" & vbCrLf _
        & "'----------------------------------------------------------------------------------" & vbCrLf _
        & "' Modul     : " & SqlDebugPrintFactoryModuleName & vbCrLf _
        & "' Date      : 11.08.2017" & vbCrLf _
        & "' Purpose   : dieses Modul muss in die Anwendung in der das Tool verwendet wird" & vbCrLf _
        & "'----------------------------------------------------------------------------------" & vbCrLf & vbCrLf

    str = str _
        & "Private m_SqlDebugger As Object 'Muss leider Latebinding sein" & vbCrLf _
        & "Private Const SqlDebuggerWizardFileName As String = ""SQLDebugPrint""" & vbCrLf & vbCrLf _
        
    str = str _
        & "Public Function SqlDebugger() As Object" & vbCrLf _
        & "    If m_SqlDebugger Is Nothing Then" & vbCrLf _
        & "        Set m_SqlDebugger = Application.Run(GetAddInLocation & SqlDebuggerWizardFileName & "".NewSqlDebugger"")" & vbCrLf _
        & "    End If" & vbCrLf _
        & "    Set SqlDebugger = m_SqlDebugger" & vbCrLf _
        & "End Function" & vbCrLf & vbCrLf
        
    str = str _
        & "Public Function SqlDebug( _" & vbCrLf _
        & "                       ByVal SqlText As String, _" & vbCrLf _
        & "                       Optional WindowMode As AcWindowMode = acWindowNormal)" & vbCrLf & vbCrLf _
        & "    SqlDebugger.ShowSql SqlText, WindowMode" & vbCrLf & vbCrLf _
        & "End Function" & vbCrLf & vbCrLf

    str = str _
        & "Public Function CriteriaDebug( _" & vbCrLf _
        & "                        ByVal Source As String, _" & vbCrLf _
        & "                        ByVal Criteria As String, _" & vbCrLf _
        & "                        Optional WindowMode As AcWindowMode = acWindowNormal)" & vbCrLf & vbCrLf _
        & "    Dim SqlText As String" & vbCrLf & vbCrLf _
        & "    If Not (Trim(Source) Like ""Select *"") Then" & vbCrLf _
        & "        Source = ""SELECT * from "" & Source" & vbCrLf _
        & "    End If" & vbCrLf & vbCrLf _
        & "    SqlText = Source" & vbCrLf & vbCrLf _
        & "    If Len(Criteria) > 0 Then" & vbCrLf _
        & "       SqlText = SqlText & "" WHERE "" & Criteria" & vbCrLf _
        & "    End If" & vbCrLf & vbCrLf _
        & "    SqlDebugger.ShowSql SqlText, WindowMode" & vbCrLf & vbCrLf _
        & "End Function" & vbCrLf & vbCrLf
    
    str = str _
        & "Private Function GetAddInLocation() As String" & vbCrLf _
        & "   GetAddInLocation = GetAppDataLocation & ""\Microsoft\AddIns\""" & vbCrLf _
        & "End Function" & vbCrLf & vbCrLf
        
    str = str _
        & "Public Function GetAppDataLocation()" & vbCrLf & vbCrLf _
        & "   With CreateObject(""WScript.Shell"")" & vbCrLf _
        & "      GetAppDataLocation = .ExpandEnvironmentStrings(""%APPDATA%"") & """"" & vbCrLf _
        & "   End With" & vbCrLf & vbCrLf _
        & "End Function"
        
    GetModuleString = str

End Function
