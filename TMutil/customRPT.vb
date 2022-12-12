'Imports System.ComponentModel
'Imports System.Threading
Imports Microsoft.Office.Interop '.Interop.Excel

Public Class customRPT



    Public Sub dump2TXT(ByRef myXLS3d As Object, ByVal numRows As Long, rArgs As reportingArgs)
        Dim F As New Collection
        F = rArgs.someColl

        Dim hugE$ = ""
        Dim roW As Long = 0
        Dim coL As Integer = 0

        Dim tempL$ = ""

        For coL = 0 To F.Count - 1
            tempL += F(coL + 1)
            If coL = F.Count - 1 Then tempL += vbCrLf Else tempL += ","
        Next
        hugE += tempL

        For roW = 0 To numRows - 1
            tempL = ""
            For coL = 0 To F.Count - 1
                tempL += csvObj(myXLS3d(roW, coL))
                If coL = F.Count - 1 Then tempL += vbCrLf Else tempL += ","
            Next
            hugE += tempL
        Next

        Dim fileN$ = rArgs.s1

        Call streamWriterTxt(fileN, hugE)
        hugE = ""

    End Sub

    Private Function csvObj(ByRef a$) As String
        If a = "0" Or a = "" Or Val(a) > 0 Then
            csvObj = a
        Else
            a = Replace(Replace(a, Chr(34), ""), ",", "_")
            csvObj = Chr(34) + a + Chr(34)
        End If
    End Function
    Public Sub dump2XLS(ByRef myXLS3d As Object, ByVal numRows As Long, rArgs As reportingArgs, Optional dontFreeze As Boolean = False, Optional ByVal doPivot As Boolean = False)
        Dim F As New Collection
        F = rArgs.someColl


        If rArgs.booL1 = False Then
            dump2TXT(myXLS3d, numRows, rArgs)
            Exit Sub
        End If

        'MainUI.addLOG("Opening Excel")

        Dim appXL As New Excel.Application
        Dim xlWB As Excel.Workbook
        Dim xlWS As Excel.Worksheet

        appXL.Visible = True

        xlWB = appXL.Workbooks.Add
        xlWS = xlWB.Sheets(1)

        xlWS.Activate()

        Dim colNUM As Integer = 0

        'column titles
        For Each field In F
            colNUM += 1
            xlWS.Cells(1, colNUM) = field
        Next

        Dim startRow$ = "A2"
        'MainUI.addLOG("Copying 3D to Excel")

        xlWS.Range(startRow + ":" + xlsColName(F.Count) + Trim(Str(numRows + 1))).Value = myXLS3d

        xlWB.RefreshAll()
        xlWS.Columns.AutoFit()

        ' not available first convert to .netcore
        '   For colNUM = 1 To F.Count
        '       If xlWS.Columns(colNUM).columnwidth > 50 Then xlWS.Columns(colNUM).columnwidth = 50
        '  Next

        appXL.Visible = True

        '        If dontFreeze = False Then
        '            xlWS.Rows("1:1").
        '            With appXL.ActiveWindow
        '                .SplitColumn = 0
        '                .SplitRow = 1
        '            End With
        '            appXL.ActiveWindow.FreezePanes = True
        '        End If


        If doPivot = True Then
            xlWS.Name = "RAW_DATA"
            xlWS = xlWB.Sheets.Add
            xlWS.Name = "Pivot"
            xlWS.Activate()
            Call xlWSdoPivot(xlWS, numRows, F, "TF", appXL)
            '      If rArgs.s3 = "Scans" Then Call xlWSdoPivot(xlWS, numRows, F, CxDATAtype.Scans, appXL)
            '      If rArgs.s3 = "Results" Then Call xlWSdoPivot(xlWS, numRows, F, CxDATAtype.Results, appXL)
        End If

        Dim numCopies As Integer = 0

        Dim a$ = ""

        'xlWS.Rows("1:1").Select


        appXL.DisplayAlerts = True

        '        appXL.Visible = True

        '        Do Until Dir(rArgs.s1) = ""
        '            numCopies += 1
        '            rArgs.s1 = Replace(rArgs.s1, ".xlsx", "") + numCopies.ToString + ".xlsx"
        '        Loop

        '        xlWB.SaveAs(rArgs.s1)
        '       xlWB.Close()
        '       appXL.Quit()

        xlWS = Nothing
        xlWB = Nothing
        appXL = Nothing
        myXLS3d = Nothing
        GC.Collect()

    End Sub





    Public Sub xlWSdoPivot(ByRef xlWS As Excel.Worksheet, ByVal rowNum As Long, ByRef fieldS As Collection, tData$, ByRef appXL As Excel.Application)

        xlWS.PivotTableWizard(Excel.XlPivotTableSourceType.xlDatabase, SourceData:="RAW_DATA!R1C1:R" + Trim(Str(rowNum - 1)) + "C" + Trim(Str(fieldS.Count)), TableDestination:=xlWS.Range("A5"), TableName:="ResultsSummary")

        ' + Trim(Str(fieldS.Count)), TableDestination:=xlWS.Range("A5"), TableName:="ResultsSummary")
        xlWS.Select()

        Dim PT As Excel.PivotTable = xlWS.PivotTables("ResultsSummary")


        Select Case tData$
            Case "TF"
                '                .Add("Component")field
                '                .Add("ComponentType")
                '                .Add("Labels")
                '                .Add("Threat")
                '                .Add("Severity")
                '                .Add("Label%")
                '                .Add("Security Requirement")
                For Each fielD In fieldS
                    With PT.PivotFields(fielD)
                        If InStr(fielD, "#_") Then
                            .Orientation = Excel.XlPivotFieldOrientation.xlDataField
                        Else
                            .Orientation = Excel.XlPivotFieldOrientation.xlRowField
                            If InStr(fielD, "Component") Then .subtotals(1) = True Else .subtotals(1) = False
                        End If

                    End With
                Next


            Case "Activity"
                With PT.PivotFields("# Scans")
                    .Orientation = Excel.XlPivotFieldOrientation.xlDataField
                End With
                With PT.PivotFields("Origin")
                    .Orientation = Excel.XlPivotFieldOrientation.xlColumnField
                    .subtotals(1) = False
                End With
                With PT.PivotFields("Team")
                    .Orientation = Excel.XlPivotFieldOrientation.xlRowField
                End With
                With PT.PivotFields("ProjName")
                    .Orientation = Excel.XlPivotFieldOrientation.xlRowField
                    .subtotals(1) = False
                End With

                xlWS.Rows("7:7").select : appXL.ActiveWindow.FreezePanes = True

            Case "Devices"
                With PT.PivotFields("# Projects")
                    .Orientation = Excel.XlPivotFieldOrientation.xlDataField
                End With
                With PT.PivotFields("Team")
                    .Orientation = Excel.XlPivotFieldOrientation.xlRowField
                End With
                With PT.PivotFields("ProjName")
                    .Orientation = Excel.XlPivotFieldOrientation.xlRowField
                    .subtotals(1) = False
                End With

                xlWS.Rows("7:7").select : appXL.ActiveWindow.FreezePanes = True

            Case "Alerts"
                With PT.PivotFields("# Results")
                    .Orientation = Excel.XlPivotFieldOrientation.xlDataField
                End With
                With PT.PivotFields("Team")
                    .Orientation = Excel.XlPivotFieldOrientation.xlRowField
                End With
                With PT.PivotFields("ProjName")
                    .Orientation = Excel.XlPivotFieldOrientation.xlRowField
                    .subtotals(1) = False
                End With
                With PT.PivotFields("ScanId")
                    .Orientation = Excel.XlPivotFieldOrientation.xlRowField
                    .subtotals(1) = False
                End With


                xlWS.Rows("7:7").select : appXL.ActiveWindow.FreezePanes = True

        End Select
        '       For Each F In fieldS
        '            If grpNDX(colFields, F) Then
        '                PT.PivotFields(F).orientation = Excel.XlPivotFieldOrientation.xlColumnField
        '                PT.PivotFields(F).position = 1
        '                PT.PivotFields(F).subtotals(1) = False
        '            ElseIf grpNDX(rowFields, F) Then
        '                PT.PivotFields(F).orientation = Excel.XlPivotFieldOrientation.xlRowField
        '                PT.PivotFields(F).subtotals(1) = False

        '            ElseIf 
        '      If InStr(F, "# ") Then PT.PivotFields(F).orientation = Excel.XlPivotFieldOrientation.xlDataField
        '  PT.PivotFields(F).subtotals(1) = False
        '            End If

        '     Next

        '        For Each F In valFields
        '            If grpNDX(valFields, F) Then
        '                'PT.PivotTables("ResultsSummary").AddDataField(PT.PivotTables("ResultsSummary").PivotFields(F.ToString), "Avg of " + F.ToString, -4106)
        '                '       PT.PivotFields(F).orientation = Excel.XlPivotFieldOrientation.xlDataField
        '
        '        PT.AddDataField(PT.PivotFields(F), "Sum of " + F, -4157)
        '        PT.PivotFields(F).subtotals(1) = False
        '    End If
        'Next

        'PT.PivotFields("Sum of VULN_CT").Orientation = Excel.XlPivotFieldOrientation.xlHidden
        'PT.PivotFields("ScanDate").AutoSort xlDescending, "ScanDate"
        'PT.PivotFields("ScanDate").autosort(2, "ScanDate")
        PT.RowGrand = True
        PT.TableStyle2 = "PivotStyleDark2"
        xlWS.Columns.AutoFit()

    End Sub
End Class

Public Class reportingArgs
    Public rptName$
    Public booL1 As Boolean = False
    Public booL2 As Boolean = False
    Public s1$ = ""
    Public s2$ = ""
    Public s3$ = ""
    Public numeriC As Integer = 0
    Public numeriC2 As Integer = 0
    Public numeriC3 As Integer = 0
    Public someColl As Collection
End Class


