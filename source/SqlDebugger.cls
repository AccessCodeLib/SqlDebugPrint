VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SqlDebugger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>%AppFolder%/source/SqlDebugger.cls</file>
'</codelib>
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit

Public Enum QueryTyp
    Kriterium = 0
    AuswahlAbfrage = 1
    AktionsAbfrage = 2
End Enum

Public Enum SqlCheckInfo
    [_KeineAngabe] = 0
    SqlIstKorrekt = 1
    SqlIstFehlerhaft = 2
End Enum

Public Enum ViewOption
    DataView = acViewNormal
    SqlView = acViewDesign
End Enum

Public Sub ShowSql( _
                    ByVal SqlText As Variant, _
                    Optional WindowMode As AcWindowMode = acWindowNormal)

   If WindowMode = acIcon Or WindowMode = acHidden Then WindowMode = acWindowNormal

   If WindowMode = acDialog Then
      Set SqlDebuggerTransferRef = Me
      DoCmd.OpenForm "frmDebugPrint", acNormal, , , , acDialog, SqlText
      Set SqlDebuggerTransferRef = Nothing
   Else
      OpenNewFormInstance SqlText
   End If

End Sub

Private Sub OpenNewFormInstance(ByVal SqlText As Variant)

   Dim frm As Form_frmDebugPrint

   Set frm = New Form_frmDebugPrint

   Set frm.SqlDebugger = Me
   frm.SqlText = SqlText

End Sub

Public Function SQLPruefen( _
                            ByVal SqlText As String, _
                            Optional ByRef InfoText As String = vbNullString) As SqlCheckInfo

On Error GoTo Fehler

    Dim rstTemp As DAO.Recordset

    With New ErrorTrappingObserver
      .ErrorTrapping = BreakOnUnhandledErrors

    DBEngine.BeginTrans
    With CurrentDb.CreateQueryDef("", SqlText)
        If GetQueryTyp(SqlText) = QueryTyp.AktionsAbfrage Then
            .ReturnsRecords = False
            .Execute
        Else
            Set rstTemp = .OpenRecordset
            rstTemp.Close
        End If
    End With
    DBEngine.Rollback
    InfoText = "Die SQL-Anweisung/das Kriterium ist ohne Fehler ausführ- bzw. anwendbar."
    SQLPruefen = SqlCheckInfo.SqlIstKorrekt

   End With

ExitHere:
    Set rstTemp = Nothing
    Exit Function

Fehler:
    DBEngine.Rollback
    Dim i As Long
    Dim ErrDesc As String
    For i = 0 To (DBEngine.Errors.Count - 1)
        ErrDesc = ErrDesc & DBEngine.Errors(i).Description & vbNewLine
    Next
    InfoText = ErrDesc
    SQLPruefen = SqlCheckInfo.SqlIstFehlerhaft

    Resume ExitHere

End Function

Public Function GetQueryTyp( _
                        ByVal varString As String) As QueryTyp

   Dim TestSql As String
   
   TestSql = Trim(Replace(varString, vbNewLine, " "))
   
   If TextStartWith(TestSql, "PARAMETERS ") Then
      'Parameter-Abschnitt abschneiden
      TestSql = Trim(Mid(TestSql, InStr(1, TestSql, ";") + 1))
   End If
   
   If TextStartWith(TestSql, "SELECT ", "TRANSFORM ") Then
      If TextStartWith(TestSql, "SELECT INTO ") Then
         GetQueryTyp = AktionsAbfrage
      Else
         GetQueryTyp = AuswahlAbfrage
      End If
   ElseIf TextStartWith(TestSql, "UPDATE ", "INSERT INTO ", "DELETE ") Then
      GetQueryTyp = AktionsAbfrage
   Else
      GetQueryTyp = Kriterium
   End If

End Function

Private Function TextStartWith(ByVal TextToCheck As String, ParamArray TextToFind() As Variant) As Boolean

   Dim i As Long
   Dim ReturnValue As Boolean
   
   For i = LBound(TextToFind) To UBound(TextToFind)
      If StrComp(Left(TextToCheck, Len(TextToFind(i))), TextToFind(i), vbTextCompare) = 0 Then
         ReturnValue = True
         Exit For
      End If
   Next
   
   TextStartWith = ReturnValue
   
End Function
