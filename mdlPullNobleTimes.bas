Attribute VB_Name = "mdlPullNobleTimes"
Option Explicit

Function createNobleTimes(strTable As String, strPath As String, strBaseTable As String) As Boolean


Dim tdf As TableDef
Dim strConnect As String
Dim fRetval As Boolean
Dim myDB As DAO.Database

    DoCmd.SetWarnings False
    Set myDB = CurrentDb
    Set tdf = myDB.CreateTableDef(strTable)
    
    With tdf
        .Connect = strPath
        .SourceTableName = strBaseTable
    End With
    
    myDB.TableDefs.Append tdf
    
    fRetval = True
        
    DoCmd.SetWarnings True

CreateNobleTimesExit:
    createNobleTimes = fRetval
    Exit Function

CreateNobleTimesError:
    If Err = 3110 Then
        Resume CreateNobleTimesExit
    Else
        If Err = 3011 Then
            Resume Next
        End If
    End If
    
End Function

Public Function NobleTIMES()

Dim dbs As DAO.Database
Dim rst As Recordset

Set dbs = CurrentDb()

Dim SQL, SQL1, SQL1a, SQL2 As String

Dim RunDate As Variant
Dim MonthChecker As Integer

RunDate = InputBox("Enter date in M/DD/YYYY format only.", "Noble Hour Counts")

If Len(RunDate) = 9 Then
'3/16/2016
RunDate = Left(RunDate, 2) & Mid(RunDate, 3, 2) + 1 & Right(RunDate, 5)
End If

If Len(RunDate) > 9 Then
'11/16/2015
RunDate = Left(RunDate, 3) & Mid(RunDate, 4, 2) + 1 & Right(RunDate, 5)
End If



Dim datestring As String
    datestring = "tsktsrhst" & Replace(Format(RunDate, "mm/dd/yy"), "/", "")
    
Call createNobleTimes(datestring, "ODBC;DSN=emhst;DBQ=/usr/task/task_hst", datestring)

DoCmd.SetWarnings False

'Delete Data From Stage Tables
SQL1 = "DELETE * FROM t_AgentReport_TEMP;"
SQL2 = "DELETE * FROM t_AgentReport_TEMP_01;"

'Start t_AgentReport_TempPart
SQL = "INSERT INTO t_AgentReport_TEMP ( call_date, NobleID, appl, tot_calls, Contacts, Calls, totn, totb, totd, totu, SumOftot_1, SumOftot_2," & _
      "SumOftot_3, SumOftot_4, SumOftot_5, SumOftot_6, SumOftot_7, SumOftot_8, SumOftot_9, SumOftot_10, timeconnect, timepause, timewait, timedeassign," & _
      "timeacw, tothours ) SELECT " & datestring & ".call_date, CVar([tsr]) AS NobleID, " & datestring & ".appl, " & datestring & ".tot_calls, Sum(([tot_1]+[tot_2]" & _
      "+[tot_3]+[tot_4]+[tot_6]+[tot_7]+[tot_8])) AS " & _
      "Contacts, Sum(([tot_1]+[tot_2]+[tot_3]+[tot_4]+[tot_5]+[tot_6]+[tot_7]+[tot_8]+[tot_9]+[tot_10]+[tot_n]+[tot_d]+[tot_b]+[tot_u])) AS Calls," & _
      "Sum(" & datestring & ".tot_n) AS totn, Sum(" & datestring & ".tot_b) AS totb, Sum(" & datestring & ".tot_d) AS totd, Sum(" & datestring & ".tot_u)" & _
      "AS totu, Sum(" & datestring & ".tot_1) AS SumOftot_1, Sum(" & datestring & ".tot_2) AS SumOftot_2, Sum(" & datestring & ".tot_3) AS SumOftot_3, Sum(" & datestring & ".tot_4) AS SumOftot_4, Sum(" & datestring & ".tot_5) AS SumOftot_5, Sum(" & datestring & ".tot_6) AS SumOftot_6, Sum(" & datestring & ".tot_7) AS " & _
      "SumOftot_7, Sum(" & datestring & ".tot_8) AS SumOftot_8, Sum(" & datestring & ".tot_9) AS SumOftot_9, Sum(" & datestring & ".tot_10) AS SumOftot_10," & _
      "Sum(" & datestring & ".time_connect) AS timeconnect, Sum(" & datestring & ".time_paused) AS timepause, Sum(" & datestring & ".time_waiting) AS timewait, Sum(" & datestring & ".time_deassigned) AS timedeassign, Sum(" & datestring & ".time_acw) AS timeacw, Sum([time_connect]+[time_paused]+[time_waiting]+[time_deassigned]+[time_acw]) AS tothours FROM " & datestring & " GROUP BY " & datestring & ".call_date, CVar([tsr]), " & datestring & ".appl, " & datestring & ".tot_calls;"



'Run SQL
DoCmd.RunSQL (SQL1)
DoCmd.RunSQL (SQL2)
DoCmd.RunSQL (SQL)

'Into new group stage table
DoCmd.OpenQuery "qryNOBLEAgentReport_02"

'Do some converting for math / numbers / percents
DoCmd.OpenQuery "qryNOBLEAgentReport_03"

'Update Agent Name from NOBLEID
DoCmd.OpenQuery "qryNOBLEAgentReport_UpdateAgentName_04"

'Delete Mortgage (RMP*)
DoCmd.OpenQuery "qryNOBLEAgentReport_DeleteMortgage_05"

'Delete Noble Daily Linked Table
dbs.TableDefs.Delete datestring


DoCmd.SetWarnings True




End Function


