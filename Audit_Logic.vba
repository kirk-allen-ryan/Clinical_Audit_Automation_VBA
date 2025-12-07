'##############################################################################################
'Create a short UUID for file_cluster Identification
'##############################################################################################

Function GenerateShortID(length As Integer) As String
    Dim randomNumber As Long
    Dim base64Chars As String
    Dim shortID As String
    Dim i As Integer
    
    ' Characters for base64 representation
    base64Chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-_"
    
    ' Generate a random number
    Randomize
    Do While Len(shortID) < length
        randomNumber = CLng(Rnd * 64)
        shortID = shortID & Mid(base64Chars, randomNumber + 1, 1)
    Loop
    
    GenerateShortID = shortID
End Function
'##############################################################################################

Sub ReThreader()

'##############################################################################################
'Declare variables
'##############################################################################################
    
Dim isFound As Boolean
Dim currentDate As Date
Dim elapsedTime, endTime, startTime As Double
Dim cv, expt, i, j, k, LastRow, nextRow, nn_count, status, rowCount, colCount As Integer   
Dim newColumn, verColumn1, verColumn2 As ListColumn
Dim staffKeyTable, tbl, tbl2, sourceTable, targetTable, finalTable, verTbl As ListObject
Dim newRow, tblRow, verRow As listRow
Dim olTask, OutApp, OutMail As Object
Dim myDelegate As Outlook.recipient
Dim pTable As PivotTable
Dim cell, f2, pendingRange, rng, rng2, sourceRange, targetRange, finalRange, transferData, finalCopy, coverSheet As Range
Dim btn As Shape
Dim auditBotPath, cellValue2, columnName, installed, job1, newname, permPath, PERM_PATH, recipient, manager, reviewTemplatePath, abName, abPath, calcName, calcPath, logName, parName As String
Dim sourceFilePath, staffKeyPath, targetWorkbookName, template, admin, UserName, windowname, logPath, routePath, randomID, revPath, skName, skPath, , revName, parPath As String    
Dim values(), worksheetNames() As Variant
Dim newWorkbook, sourceBook, targetBook, wb, wbStaffKey, wbTemplate, calc, sk, log, rev, par, ab As Workbook
Dim lastWs, logz, PRN, ws, ws2, siFoc, sourceSheet, targetSheet, finalSheet, accVer As Worksheet
    
'############################################################################################## 
'Preliminary: verifiy root path for file management
'Set paths for helper file templates
'Check for verified admin
'Set emial addresses for admin and manager roles from the lauch-file dictionary    
'##############################################################################################    

        Dim verver As String: verver = ThisWorkbook.Worksheets("Sheet5").Range("C10").Value
        If verver <> "True" Then
        MsgBox "Verify Root Path before continuing..."
        Exit Sub
        End If
              
        Application.ScreenUpdating = False

        startTime = Timer
        
        UserName = Environ("USERNAME")
              
        abName = "AuditBot.xlsm"
        
        abPath = PERM_PATH & abName
        
        Set ab = Workbooks(abName)
        
        PERM_PATH = ab.Sheets("DICT").Range("B2").Value
        
        calcName = "IP_Calc.Thread.xlsm"
        
        calcPath = PERM_PATH & "Files\INPATIENT\Templates\" & calcName
        
        logName = "Run_Logs.xlsm"
        
        logPath = PERM_PATH & "Logs\" & logName
        
        mlogName = "Message_Logs.xlsm"
        
        mlogPath = PERM_PATH & "Logs\" & logName
        
        parName = "Pain_Audit_Report.xlsx"
        
        parPath = Environ("USERPROFILE") & "\Downloads\" & parName
        
        revName = "IP_Review.Template.xlsm"
        
        revPath = PERM_PATH & "Files\INPATIENT\Templates\" & revName
        
        skName = "IP_Staff_Key.xlsm"
        
        skPath = PERM_PATH & "Files\INPATIENT\Templates\" & skName
        
        Application.EnableEvents = False
        
        Set siFoc = ab.Worksheets("Sheet5")

        cellValue2 = siFoc.Range("C10").Value
                
         Set siFoc = ab.Worksheets("DICT")
         cellValue2 = siFoc.Range("B3").Value
         
    If Len(cellValue2) < 1 Then
        MsgBox "Bummer - the admin parameter needs to be verified - you can't execute these scripts without a verified admin..."
        Exit Sub
        Else: admin = cellValue2
    End If
    
        cellValue2 = siFoc.Range("b4").Value
        
    If Len(cellValue2) < 1 Then
        MsgBox "Bummer - the IP Review Account parameter needs to be verified - you can't execute these scripts without a verified Review Account..."
        Exit Sub
    Else: manager = cellValue2
    End If
  
    '_____If  job1 = TEST, then send to test account!_____
    
    On Error Resume Next

'##############################################################################################                        
'Open Pain_Audit_Report (raw data) from admin's download folder - abort if not there
'##############################################################################################   
                        
    Set par = Workbooks.Open(parPath)
    
    On Error GoTo 0
    
    If par Is Nothing Then
        MsgBox "I didn't find a Pain Audit Report in your downloads...my job here is done...ping me when you have the raw data ready", vbExclamation, "File Not Found"
    Exit Sub
    
    Else
        MsgBox "I found your latest Pain Audit Report - this will only take a few seconds..."
        
    End If
                                
'##############################################################################################                        
'Format raw data file (BO delivers everything as text)
'check for test-mode (emails sent to admin), etc. 
'delete Pharmacy report sheets - moved to independent process (module-10)  
'##############################################################################################                                
    
    For Each ws In par.Worksheets
        
        Set rng = ws.Cells

        rng.NumberFormat = "General"
        
    Next ws
    
    Set siFoc = par.Worksheets("PRN")
    
    job1 = siFoc.Range("Z100000").Value 'job1 = 'TEST', send any messages to admin, if job1="", send message to manager

    worksheetNames = Array("Pharm_PRN", "Assessments_Pharm", "Order Comments", "OR Case")
        
    For k = LBound(worksheetNames) To UBound(worksheetNames)
        On Error Resume Next
        Set ws = par.Sheets(worksheetNames(k))
        On Error GoTo 0
        If Not ws Is Nothing Then
            Application.DisplayAlerts = False
            ws.Delete
            Application.DisplayAlerts = True
        End If
    Next k
    
    Erase worksheetNames

'##############################################################################################                        
'Open StaffKey file
'Check for raw-data staff names NOT present in the key - send out the key to fix if found
'Delete raw-data rows linked to exempt staff
'Check for ALL staff being exempt - abort and log run if no eligible raw-data survives
'Create coversheet with UUID, attach to Staff Key when sending to manager/admin to log return
'Create Outlook Task for MANAGER ONLY as needed                                
'##############################################################################################    
    
    Set sk = Workbooks.Open(skPath)
        
    Sheets("Staff_Key").Copy After:=par.Sheets(4)
    
    sk.Close SaveChanges:=False

    Set ws = par.Sheets("Staff_Key")
    Dim parSK As ListObject: Set parSK = ws.ListObjects("StaffKey")
    Dim parSKSource As ListColumn: Set parSKSource = parSK.ListColumns(5)
    
    Dim rowsToDelete As Collection
  
    Set rowsToDelete = New Collection

    If job1 <> "TEST" Then
    For i = parSK.ListRows.Count To 1 Step -1
        If parSK.ListRows(i).Range.Cells(1, parSKSource.Index).Value = "Test_Data" Then
            parSK.ListRows(i).Delete
        End If
    Next i
    End If
          
        Set siFoc = par.Sheets("PRN")
        
        siFoc.Rows(1).RowHeight = 20
        siFoc.Rows(1).EntireRow.Insert xlDown
        siFoc.Rows(1).EntireRow.Insert xlDown
        Set rng = siFoc.Range("A4").CurrentRegion
        Set tbl = siFoc.ListObjects.Add(xlSrcRange, rng, , xlYes)
        tbl.Name = "prn"
        Set newColumn = tbl.ListColumns.Add(Position:=tbl.ListColumns.Count + 1)
        newColumn.DataBodyRange.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-2],StaffKey,4,FALSE),""PENDING"")"
        newColumn.Name = "keep"
        
        siFoc.Range("O1").FormulaR1C1 = "=COUNTIF(R[4]C:R[4999]C, ""APPLY"")"
        siFoc.Range("O2").FormulaR1C1 = "=COUNTIF(R[3]C:R[4998]C,""EXEMPT"")"
        siFoc.Range("O3").FormulaR1C1 = "=COUNTIF(R[2]C:R[4997]C, ""PENDING"")"
        siFoc.Range("E1").FormulaR1C1 = "=TEXT(MIN(R[4]C:R[4999]C), ""mm/dd/yy"")"
        siFoc.Range("E2").FormulaR1C1 = "=TEXT(MAX(R[3]C:R[4998]C),""mm/dd/yy"")"
        siFoc.Range("F1").FormulaR1C1 = "=CONCATENATE(LEFT(RC[-1],2) & ""_"" & RIGHT(RC[-1],2))"
        
        fn = siFoc.Range("F1").Value
        cv = siFoc.Range("O2").Value
        expt = siFoc.Range("O1").Value + siFoc.Range("O3").Value
        status = siFoc.Range("O1").Value + siFoc.Range("O2").Value
 
        If cv > 0 Then
        tbl.Range.AutoFilter Field:=15, Criteria1:="EXEMPT"
        With tbl.Range
            Set rng = .Offset(1).Resize(.Rows.Count - 1).SpecialCells(xlCellTypeVisible)
            .AutoFilter
            rng.Delete
        End With
        End If

        If expt < 1 Then
        MsgBox "Congrats! There are no flag-eligible staff detected in this report - no flags, nothing more to do!"
        par.Close SaveChanges:=False
        
            Set log = Workbooks.Open(logPath)
            Set siFoc = log.Worksheets("LOG")
            Set tbl = siFoc.ListObjects("log")
            endTime = Timer
            elapsedTime = endTime - startTime
            
            Set newRow = tbl.ListRows.Add
            With newRow.Range
                .Cells(1, 1).Value = Now
                .Cells(1, 2).Value = "IP"
                .Cells(1, 3).Value = UserName
                .Cells(1, 4).Value = elapsedTime
                .Cells(1, 6).Value = "COMPLETE"
            End With
            log.Close SaveChanges:=True
            Exit Sub
        
        End If
        
        cv = siFoc.Range("O3").Value
        
        If cv > 0 Then
        
           If job1 = "TEST" Then
               MsgBox "There are staff names on the report without Staff_Key flag status - the Staff_Key will be sent to " & admin, vbExclamation, "Warning"
           ElseIf job1 <> "TEST" Then
               MsgBox "There are staff names on the report without Staff_Key flag status - the Staff_Key will be sent to " & manager, vbExclamation, "Warning"
           End If

               Set siFoc = par.Sheets.Add
               par.PivotCaches.Create(SourceType:=xlDatabase, SourceData:="prn" _
                   , Version:=6).CreatePivotTable TableDestination:="Sheet2!R3C1", TableName _
                   :="PivotTable1", DefaultVersion:=6
               Set pTable = siFoc.PivotTables("PivotTable1")
               With pTable.PivotFields("STAFF")
                   .Orientation = xlRowField
                   .Position = 1
               End With
               With pTable
                   .ColumnGrand = False
                   .RowGrand = False
               End With
            
            If status > 0 Then
                
                With pTable.PivotFields("keep")
                    .Orientation = xlPageField
                    .Position = 1
                End With
                pTable.PivotFields("keep").CurrentPage = "(All)"
                With pTable.PivotFields("keep")
                    .PivotItems("APPLY").Visible = False
                End With
                pTable.PivotFields("keep").EnableMultiplePageItems = True
                
            End If

                siFoc.Range("D1").FormulaR1C1 = "=SUBTOTAL(3,R[3]C[-3]:R[99]C[-3])"
                nn_count = siFoc.Range("D1").Value
                Set rng2 = siFoc.Range("A4:A" & (3 + nn_count))

                Set targetSheet = par.Worksheets("Staff_Key")
            
                Set staffKeyTable = targetSheet.ListObjects("StaffKey")
            
                Set targetRange = staffKeyTable.ListColumns(1).DataBodyRange.Cells(staffKeyTable.ListRows.Count + 1, 1)
            
                targetRange.Resize(nn_count, 1).Value = rng2.Value
                
                LastRow = staffKeyTable.ListRows.Count
                Set pendingRange = staffKeyTable.ListColumns(4).DataBodyRange.Cells((LastRow - nn_count) + 1, 1).Resize(nn_count)
                pendingRange.Value = "PENDING"
                
                LastRow = staffKeyTable.ListRows.Count
                Set pendingRange = staffKeyTable.ListColumns(5).DataBodyRange.Cells((LastRow - nn_count) + 1, 1).Resize(nn_count)
           
                Set f2 = targetSheet.Range("F2")
                
            If job1 = "TEST" Then
                
                f2.Value = UserName
                pendingRange.Value = "Test_Data"
            Else
                f2.Value = ""
                pendingRange.Value = "Live_Data"
            End If

'##############################################################################################
'Create coversheet from template, populate, attach, and hide - send created objects as needed
'Coversheet contains UUID, log-in path to recipient client-folders, etc                                                                                                                                
'##############################################################################################
                                                                                                                                
                'fill out your par.coverSheet for staff key here...
                
                randomID = GenerateShortID(12)
                
                recipient = IIf(job1 = "TEST", admin, manager)
                    Set accVer = ab.Worksheets("ACCOUNTS_VER")
                    Set verTbl = accVer.ListObjects("verified")
                    Set verColumn1 = verTbl.ListColumns(2)
                    Set verColumn2 = verTbl.ListColumns(3)

                For Each verRow In verTbl.ListRows
                
                    If verRow.Range(verColumn1.Index).Value = recipient Then
                    routePath = verRow.Range(verColumn2.Index).Value
                    End If
                Next verRow
                
                Set coverSheet = targetSheet.Range("coversheet")
                coverSheet.Cells(3, 1).Value = Now
                coverSheet.Cells(3, 2).Value = recipient
                coverSheet.Cells(3, 3).Value = admin
                coverSheet.Cells(3, 4).Value = ""
                coverSheet.Cells(3, 5).Value = randomID
                coverSheet.Cells(3, 6).Value = skName
                coverSheet.Cells(3, 7).Value = "IP_SK_SEND"
                coverSheet.Cells(3, 8).Value = routePath
                coverSheet.Cells(3, 9).Value = ""
                coverSheet.Cells(3, 10).Value = PERM_PATH
                              
                Set accVer = Nothing
                Set verTbl = Nothing
                Set verColumn1 = Nothing
                Set verColumn2 = Nothing
                       
                Set newWorkbook = Workbooks.Add
                targetSheet.Copy Before:=newWorkbook.Sheets(1)
                Application.DisplayAlerts = False
                newWorkbook.Sheets("Sheet1").Delete
                   
                newWorkbook.SaveAs skPath, FileFormat:=xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False
                
                newWorkbook.Close SaveChanges:=False
                Application.DisplayAlerts = True
               
                Set OutApp = CreateObject("Outlook.Application")
                Set OutMail = OutApp.CreateItem(0)
                
            With OutMail
                .To = recipient
                .Subject = "AuditBot: Review Inpatient Staff_Key"
                .Body = "I found some PENDINIG names in a quick scan of the most recent data - can you APPLY or EXEMPT these names (APPLY means flags WILL BE created) and send the file back to me?"
                .Attachments.Add skPath
                .Send
            End With
              
                Set olTask = OutApp.CreateItem(3)
            
                If job1 <> "TEST" Then
 
                olTask.Assign
                Set myDelegate = olTask.Recipients.Add(recipient)
                myDelegate.Resolve
                
                End If
                
                olTask.Subject = "AuditBot: IP staff names detected without Flag Status"
                olTask.Body = "The system has just sent you the IP_StaffKey for you to update - please resolve any PENDING rows with either APPLY or EXEMPT, and click APPROVE & RETURN when finished"
                olTask.DueDate = Date + 7
                olTask.Importance = 2
                olTask.Categories = "AuditBot_IP"
            
                olTask.Save
                
                If job1 <> "TEST" Then
            
                olTask.Send
                
                End If
    
                Set OutMail = Nothing
                Set OutApp = Nothing

                par.Close SaveChanges:=False

                Exit Sub

        End If
'##############################################################################################
'Format raw-data into tables, arrange, clean, sort, prep for transfer to category templates
'Calculate infusion end timestamp
'Transfer raw data to category (pre-combine) sheets - this is where dynamic event IDs are created                                                                                                                                                                                           
'##############################################################################################
                                                                                                                                                                                            
        Set siFoc = par.Sheets("SCORES")
        
        siFoc.Rows(1).Delete Shift:=xlUp
        siFoc.Columns("A").Delete Shift:=xlToLeft
        Set rng = siFoc.Range("A1").CurrentRegion
        Set tbl = siFoc.ListObjects.Add(xlSrcRange, rng, , xlYes)
        tbl.Name = "ass"
       
        Set siFoc = par.Sheets("INF_END")
        
        siFoc.Rows("1:3").Delete Shift:=xlUp
        siFoc.Columns("A").Delete Shift:=xlToLeft
        Set ws = ActiveSheet
        Set rng = siFoc.Range("A1").CurrentRegion
        Set tbl = siFoc.ListObjects.Add(xlSrcRange, rng, , xlYes)
        tbl.Name = "inf_end"

        tbl.Sort.SortFields.clear
        tbl.Sort.SortFields.Add Key:=tbl.ListColumns(2).Range, SortOn:=xlSortOnValues, Order:=xlAscending
        tbl.Sort.SortFields.Add Key:=tbl.ListColumns(4).Range, SortOn:=xlSortOnValues, Order:=xlAscending
        tbl.Sort.SortFields.Add Key:=tbl.ListColumns(3).Range, SortOn:=xlSortOnValues, Order:=xlDescending
        tbl.Sort.Apply
        
        tbl.ListColumns(1).DataBodyRange.Formula = "=IF(NOT(EXACT(RC[1],R[1]C[1])),RC[1],"""")"
        
        Set rng = tbl.ListColumns(1).DataBodyRange
        rng.Value = rng.Value
        
        tbl.Sort.SortFields.clear
        tbl.Sort.SortFields.Add Key:=tbl.ListColumns(2).Range, SortOn:=xlSortOnValues, Order:=xlAscending
        tbl.Sort.SortFields.Add Key:=tbl.ListColumns(1).Range, SortOn:=xlSortOnValues, Order:=xlDescending
        tbl.Sort.Apply
     
        Set siFoc = par.Sheets("INFUSION")
        
        siFoc.Rows(1).Delete Shift:=xlUp
        Set rng = siFoc.Range("A1").CurrentRegion
        Set tbl = siFoc.ListObjects.Add(xlSrcRange, rng, , xlYes)
        tbl.Name = "inf"
        tbl.Sort.SortFields.clear
        tbl.Sort.SortFields.Add Key:=tbl.ListColumns(2).Range, SortOn:=xlSortOnValues, Order:=xlAscending
        tbl.Sort.SortFields.Add Key:=tbl.ListColumns(5).Range, SortOn:=xlSortOnValues, Order:=xlAscending
        tbl.Sort.SortFields.Add Key:=tbl.ListColumns(9).Range, SortOn:=xlSortOnValues, Order:=xlDescending
        tbl.Sort.Apply
        Set newColumn = tbl.ListColumns("ENDTIME")

        newColumn.DataBodyRange.FormulaR1C1 = _
            "=IF(OR(NOT(ISERROR(SEARCH(""Begin"",RC[-6])))," & _
                "NOT(ISERROR(SEARCH(""mL/hr"",RC[-6])))),RC[-10]," & _
            "IF(EXACT(RC[-14],R[1]C[-14]),R[1]C[-10]," & _
            "VLOOKUP(RC[-13],inf_end,4,FALSE)))"   

        Set calc = Workbooks.Open(calcPath)
        
        Set sourceSheet = par.Sheets("PRN")
        Set targetSheet = calc.Sheets("PRN")
        Set sourceTable = sourceSheet.ListObjects("prn")
        Set targetTable = targetSheet.ListObjects("meds")
        Set sourceRange = sourceTable.DataBodyRange.Resize(ColumnSize:=sourceTable.ListColumns.Count - 1)
        Set targetRange = targetSheet.Range("B2")
        Set finalSheet = calc.Worksheets("PRECOMBINE")
      
        rowCount = sourceRange.Rows.Count
        colCount = sourceRange.Columns.Count
        
        targetTable.Resize targetTable.Range.Resize(rowCount, targetTable.ListColumns.Count)
        targetRange.Resize(rowCount, colCount).Value = sourceRange.Value
        
        Application.Calculation = xlCalculationManual
        
        Set newColumn = targetTable.ListColumns.Add(Position:=17)
        newColumn.DataBodyRange.FormulaR1C1 = _
            "=IF(ISBLANK(RC[-5]), """", " & _
                "IF(NOT(ISERROR(SEARCH(""IV"",RC[-5]))), " & _
                    """A"", " & _
                    """B""" & _
                ")" & _
            ")"
        newColumn.Name = "TYPE"
        Application.Calculate
        Set rng = newColumn.DataBodyRange
        rng.Value = rng.Value
        
        Set newColumn = targetTable.ListColumns.Add(Position:=18)
        newColumn.DataBodyRange.FormulaR1C1 = _
            "=IF(ISBLANK(RC[-15]), """", " & _
                "IF(AND(" & _
                    "NOT(EXACT(R[-1]C3,RC3)), " & _
                    "RC17=R1C" & _
                "), 1, " & _
                "IF(AND(" & _
                    "NOT(EXACT(R[-1]C3,RC3)), " & _
                    "RC17<>R1C" & _
                "), 0, " & _
                "IF(AND(" & _
                    "EXACT(R[-1]C3,RC3), " & _
                    "RC17<>R1C" & _
                "), R[-1]C, " & _
                "IF(AND(" & _
                    "EXACT(R[-1]C3,RC3), " & _
                    "RC17=R1C" & _
                "), R[-1]C+1, " & _
                """)))))"
        newColumn.Name = "A"
        Application.Calculate
        Set rng = newColumn.DataBodyRange
        rng.Value = rng.Value
        
        Set newColumn = targetTable.ListColumns.Add(Position:=19)
        newColumn.DataBodyRange.FormulaR1C1 = _
            "=IF(ISBLANK(RC[-16]), """", " & _
                "IF(AND(" & _
                    "NOT(EXACT(R[-1]C3,RC3)), " & _
                    "RC17=R1C" & _
                "), 1, " & _
                "IF(AND(" & _
                    "NOT(EXACT(R[-1]C3,RC3)), " & _
                    "RC17<>R1C" & _
                "), 0, " & _
                "IF(AND(" & _
                    "EXACT(R[-1]C3,RC3), " & _
                    "RC17<>R1C" & _
                "), R[-1]C, " & _
                "IF(AND(" & _
                    "EXACT(R[-1]C3,RC3), " & _
                    "RC17=R1C" & _
                "), R[-1]C+1, " & _
                """)))))"
        newColumn.Name = "B"
        Application.Calculate
        Set rng = newColumn.DataBodyRange
        rng.Value = rng.Value
        
        Set newColumn = targetTable.ListColumns.Add(Position:=20)
        newColumn.DataBodyRange.FormulaR1C1 = _
            "=IF(ISBLANK(RC[-17]),"""", " & _
            "CONCATENATE(R1C, " & _
                "IF(RC[-17]=""I"", " & _
                    "IF(RC[-9]=R1C19, " & _
                        "CONCATENATE(R1C18,R1C2), " & _
                        "IF(RC[-9]=R1C20,CONCATENATE(R1C18,R1C2),"""") " & _
                    "), " & _
                "IF(RC[-17]=""P"", " & _
                    "IF(RC[-9]=R1C21, " & _
                        "CONCATENATE(R1C18,R1C2), " & _
                        "IF(RC[-9]=R1C22,CONCATENATE(R1C18,R1C2),"""")" & _
                    ")," & _
                    """)))"
        newColumn.Name = "ID"
        Application.Calculate
        Set rng = newColumn.DataBodyRange
        rng.Value = rng.Value
        Dim prnEvents As Long
        prnEvents = Application.WorksheetFunction.CountA(rng)
        
        Set newColumn = targetTable.ListColumns(11)
        Set rng = newColumn.DataBodyRange
                                                                                                                                                                                                                                                                                                                    
'##############################################################################################
'Log the event counts - need the i for stats later
'##############################################################################################
                                                                                                                                                                                                                                                                                                                    
        Dim narcEvents As Long: narcEvents = 0
        Dim nonnarcEvents As Long: nonnarcEvents = 0
                
        For Each cell In rng
            If cell.Value = "Narc" Then
            narcEvents = narcEvents + 1
        ElseIf cell.Value = "Non-Narc" Then
            nonnarcEvents = nonnarcEvents + 1
        End If
        Next cell
    
        For k = 17 To 19
            targetTable.ListColumns(k).Range.EntireColumn.Hidden = True
        Next k
        
        k = targetTable.DataBodyRange.Columns(20).Rows.Count
        
        For i = 1 To k
        targetTable.DataBodyRange.Cells(i, 1).Value = targetTable.DataBodyRange.Cells(i, 20).Value
        Next i               
        Set sourceSheet = par.Sheets("SCORES")
        Set targetSheet = calc.Sheets("SCORES")
        Set sourceTable = sourceSheet.ListObjects("ass")
        Set targetTable = targetSheet.ListObjects("scores")
        Set sourceRange = sourceTable.DataBodyRange
        Set targetRange = targetSheet.Range("B2")
        
        rowCount = sourceRange.Rows.Count
        colCount = sourceRange.Columns.Count
        
        targetRange.Resize(sourceRange.Rows.Count, sourceRange.Columns.Count).Value = sourceRange.Value
'##############################################################################################
'Build out the precombine helper columns - add the completion string the the scores                                                                                                                                                                                                                                                                                                                                
'Add the assessment status based on the string and score value
                                                                                                                                                                                                                                                                                                                           
'##############################################################################################
                                                                                                                                                                                                                                                                                                                                
        Set newColumn = targetTable.ListColumns.Add(Position:=17)
        newColumn.DataBodyRange.FormulaR1C1 = _
            "=IF(ISBLANK(RC[-14]), """", " & _
                "IF(NOT(ISFORMULA(R[-1]C)), " & _
                    "RC[-11], " & _
                "IF(AND(" & _
                    "ISFORMULA(R[-1]C), " & _
                    "ROUND(((RC[-11]-R[-1]C)*1440), 0) <= 5, " & _
                    "EXACT(RC[-14],R[-1]C[-14])" & _
                "), " & _
                    "R[-1]C, " & _
                    "RC[-11]" & _
                "))" & _
            ")"
        newColumn.Name = "OPEN"
        
        Set newColumn = targetTable.ListColumns.Add(Position:=18)
        newColumn.DataBodyRange.FormulaR1C1 = _
            "=IFERROR(" & _
                "IF(" & _
                    "VLOOKUP(RC[-10],ADTA,5,FALSE) = 0, " & _
                    """""", " & _
                    "VLOOKUP(RC[-10],ADTA,5,)" & _
                "), " & _
            """""")"
        newColumn.Name = "ADTA"
        
        Application.Calculate
        
        targetTable.Sort.SortFields.clear
        targetTable.Sort.SortFields.Add Key:=targetTable.ListColumns(2).Range, SortOn:=xlSortOnValues, Order:=xlAscending
        targetTable.Sort.SortFields.Add Key:=targetTable.ListColumns(6).Range, SortOn:=xlSortOnValues, Order:=xlAscending
        targetTable.Sort.SortFields.Add Key:=targetTable.ListColumns(18).Range, SortOn:=xlSortOnValues, Order:=xlAscending
        targetTable.Sort.Apply
              
        Set newColumn = targetTable.ListColumns.Add(Position:=19)
        newColumn.DataBodyRange.FormulaR1C1 = _
            "=IF( " & _
                "NOT(EXACT( " & _
                    "CONCATENATE(RC[-16]&RC[-2]), " & _
                    "CONCATENATE(R[-1]C[-16]&R[-1]C[-2])" & _
                ")), " & _
                "RC[-1], " & _
                "CONCATENATE(R[-1]C&RC[-1])" & _
            ")"
        newColumn.Name = "STRING"
  
        Set newColumn = targetTable.ListColumns.Add(Position:=20)
        newColumn.DataBodyRange.FormulaR1C1 = _
            "=IFERROR(" & _
                "IF(OR(" & _
                    "RC[-2] = ""w"", " & _
                    "RC[-2] = ""x"", " & _
                    "RC[-2] = ""y""" & _
                "), " & _
                    "VALUE(TRIM(LEFT(RC[-10],2))), " & _
                    """"" " & _
                "), " & _
            """NO SCORE"")"
        newColumn.Name = "Row-Score"
       
        Set newColumn = targetTable.ListColumns.Add(Position:=21)
        newColumn.DataBodyRange.FormulaR1C1 = _
            "=IF(AND(" & _
                "ISNUMBER(RC[-1])," & _
                "NOT(EXACT(" & _
                    "CONCATENATE(RC[-18]&RC[-4])," & _
                    "CONCATENATE(R[1]C[-18]&R[1]C[-4])" & _
                "))" & _
            "), " & _
                "RC[-1], " & _
                """"" " & _
            ")"
        newColumn.Name = "SCORE"
      
        Set newColumn = targetTable.ListColumns.Add(Position:=22)
        newColumn.DataBodyRange.FormulaR1C1 = _
            "=IF(NOT(ISNUMBER(RC[-1])), """", " & _
                "IF(AND(ISNUMBER(RC[-1]), RC[-1]=0), " & _
                    """COMPLETE"", " & _
                "IF(AND(ISNUMBER(RC[-1]), RC[-1]>0, OR(" & _
                    "ISERROR(SEARCH(""e"",RC[-3])), " & _
                    "ISERROR(SEARCH(""f"",RC[-3])), " & _
                    "ISERROR(SEARCH(""i"",RC[-3])), " & _
                    "ISERROR(SEARCH(""j"",RC[-3])), " & _
                    "ISERROR(SEARCH(""k"",RC[-3])), " & _
                    "ISERROR(SEARCH(""l"",RC[-3]))" & _
                ")), " & _
                    """PARTIAL"", " & _
                "IF(AND(ISNUMBER(RC[-1]), RC[-1]>0, AND(" & _
                    "NOT(ISERROR(SEARCH(""e"",RC[-3]))), " & _
                    "NOT(ISERROR(SEARCH(""f"",RC[-3]))), " & _
                    "NOT(ISERROR(SEARCH(""i"",RC[-3]))), " & _
                    "NOT(ISERROR(SEARCH(""j"",RC[-3]))), " & _
                    "NOT(ISERROR(SEARCH(""k"",RC[-3]))), " & _
                    "NOT(ISERROR(SEARCH(""l"",RC[-3]))))" & _
                "), " & _
                    """COMPLETE"", " & _
                    """))" & _
            ")"
        newColumn.Name = "STATUS"
           
        Set newColumn = targetTable.ListColumns.Add(Position:=23)
        newColumn.Name = "PRN_ID"

        Set newColumn = targetTable.ListColumns.Add(Position:=24)
       newColumn.DataBodyRange.FormulaR1C1 = _
            "=IF(AND(" & _
                "NOT(ISNUMBER(RC[-3])), " & _
                "NOT(EXACT(RC[-21],R[-1]C[-21]))" & _
            "), """", " & _
            "IF(AND(" & _
                "ISNUMBER(RC[-3]), " & _
                "OR(" & _
                    "LEN(R[-1]C)<1, " & _
                    "NOT(EXACT(RC[-21],R[-1]C[-21]))" & _
                ")" & _
            "), " & _
                "CONCATENATE(RC[-21] & ""-D1""), " & _
            "IF(AND(" & _
                "NOT(ISNUMBER(RC[-3])), " & _
                "EXACT(RC[-21],R[-1]C[-21])" & _
            "), " & _
                "R[-1]C, " & _
            "IF(AND(" & _
                "ISNUMBER(RC[-3]), " & _
                "EXACT(RC[-21],R[-1]C[-21])" & _
            "), " & _
                "CONCATENATE(" & _
                    "RC[-21]&""-D""&" & _
                    "SUM(RIGHT(R[-1]C,LEN(R[-1]C)-SEARCH(""D"",R[-1]C))+1)" & _
                ")" & _
            "))))"
        newColumn.Name = "SCORE_ID"
        
        Application.Calculate
               
        Set newColumn = targetTable.ListColumns(1)
        newColumn.DataBodyRange.FormulaR1C1 = "=RC[23]"
        Set rng = newColumn.DataBodyRange
        rng.Value = rng.Value
       
            For k = 17 To 24
            Set rng = targetTable.ListColumns(k).DataBodyRange
            rng.Value = rng.Value
        Next k
                
        targetTable.Range.AutoFilter Field:=22, Criteria1:="<>"
        
        For k = 17 To 22
            targetTable.ListColumns(k).Range.EntireColumn.Hidden = True
        Next k
    
        Set sourceSheet = par.Sheets("INFUSION")
        Set targetSheet = calc.Sheets("INF")
        Set sourceTable = sourceSheet.ListObjects("inf")
        Set targetTable = targetSheet.ListObjects("inf")
        Set sourceRange = sourceTable.DataBodyRange.Resize(ColumnSize:=sourceTable.ListColumns.Count - 2)
        Set targetRange = targetSheet.Range("B2")
        
        rowCount = sourceRange.Rows.Count
        colCount = sourceRange.Columns.Count
        
        targetRange.Resize(sourceRange.Rows.Count, sourceRange.Columns.Count).Value = sourceRange.Value
        
        Set newColumn = targetTable.ListColumns.Add(Position:=19)
        newColumn.DataBodyRange.FormulaR1C1 = _
       newColumn.DataBodyRange.FormulaR1C1 = _
            "=IF(ISBLANK(RC[-16]), """", " & _
                "IF(AND(" & _
                    "NOT(EXACT(RC[-16],R[-1]C[-16])), " & _
                    "NOT(EXACT(RC[-16],R[1]C[-16]))" & _
                "), " & _
                    """Z"", " & _
                "IF(NOT(EXACT(RC[-16],R[-1]C[-16])), " & _
                    """W"", " & _
                "IF(AND(" & _
                    "EXACT(RC[-16],R[-1]C[-16]), " & _
                    "EXACT(RC[-16],R[1]C[-16]), " & _
                    "NOT(ISERROR(SEARCH(""mL/hr"",RC[-9])))" & _
                "), " & _
                    """X"", " & _
                "IF(NOT(EXACT(RC[-16],R[1]C[-16])), " & _
                    """Y"", " & _
                    """"" " & _
                "))" & _
            "))"
        newColumn.Name = "TYPE"
        Application.Calculate
        Set rng = newColumn.DataBodyRange
        rng.Value = rng.Value
        
        Set newColumn = targetTable.ListColumns.Add(Position:=20)
        newColumn.DataBodyRange.FormulaR1C1 = _
        newColumn.DataBodyRange.FormulaR1C1 = _
            "=IF(NOT(EXACT(RC[-17],R[-1]C[-17])), """", " & _
                "IF(AND(" & _
                    "ISERROR(SEARCH(""X"",RC[-1])), " & _
                    "LEN(R[-1]C) < 1" & _
                "), """", " & _
                "IF(AND(" & _
                    "NOT(ISERROR(SEARCH(""X"",RC[-1]))), " & _
                    "LEN(R[-1]C) < 1" & _
                "), 1, " & _
                "IF(AND(" & _
                    "RC[-1]<>""X"", " & _
                    "ISFORMULA(R[-1]C), " & _
                    "LEN(R[-1]C) > 0" & _
                "), " & _
                    "R[-1]C, " & _
                "IF(AND(" & _
                    "NOT(ISERROR(SEARCH(""X"",RC[-1]))), " & _
                    "LEN(R[-1]C) > 0" & _
                "), " & _
                    "R[-1]C+1, " & _
                    """)))" & _
            "))"
        newColumn.Name = "X"
        Application.Calculate
        Set rng = newColumn.DataBodyRange
        rng.Value = rng.Value
        
        Set newColumn = targetTable.ListColumns.Add(Position:=21)
        newColumn.DataBodyRange.FormulaR1C1 = _
       newColumn.DataBodyRange.FormulaR1C1 = _
            "=IF(LEN(RC[-2])<1,""""," & _
                "CONCATENATE(" & _
                    "RC[-18]&""-""&RC[-2], " & _
                    "IF(RC[-2]=""X"", RC[-1], """")" & _
                ")" & _
            ")"
        newColumn.Name = "ID"
        Application.Calculate
        Set rng = newColumn.DataBodyRange
        rng.Value = rng.Value
        Dim infEvents As Long
        infEvents = Application.WorksheetFunction.CountA(rng)
        
        Set newColumn = targetTable.ListColumns(1)
        newColumn.DataBodyRange.FormulaR1C1 = "=RC[20]"
        Set rng = newColumn.DataBodyRange
        rng.Value = rng.Value

        targetTable.ListColumns(19).Range.EntireColumn.Hidden = True
        targetTable.ListColumns(20).Range.EntireColumn.Hidden = True

        targetTable.Range.AutoFilter Field:=21, Criteria1:="<>"
        
        Sheets("PRN").Select
        Range("B2:T2").Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.SpecialCells(xlCellTypeVisible).Select
        Application.CutCopyMode = False
        Selection.Copy
        Sheets("PRECOMBINE").Select
        Range("A2").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
       
        Sheets("SCORES").Select
        Range("B1:X1").Offset(1, 0).Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.SpecialCells(xlCellTypeVisible).Select
        Application.CutCopyMode = False
        Selection.Copy
        Sheets("PRECOMBINE").Select
        Range("A1").Select
        
        Selection.End(xlDown).Offset(1, 0).Select
    
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
            
        Range("A1").Select
        Selection.End(xlDown).Offset(1, 0).Select
    
        Sheets("INF").Select
        Range("B2:U2").Select
        Range(Selection, Selection.End(xlDown)).Select
        
        Selection.SpecialCells(xlCellTypeVisible).Select
        Application.CutCopyMode = False
        Selection.Copy
        Sheets("PRECOMBINE").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Application.CutCopyMode = False
        
        Set rng = finalSheet.Range("A1").Resize(finalSheet.Cells(finalSheet.Rows.Count, 1).End(xlUp).row, 18)

        Set finalTable = finalSheet.ListObjects.Add(xlSrcRange, rng, , xlYes)
        finalTable.Name = "calcs"
        
        finalTable.Sort.SortFields.clear
        finalTable.Sort.SortFields.Add Key:=finalTable.ListColumns(2).Range, SortOn:=xlSortOnValues, Order:=xlAscending
        finalTable.Sort.SortFields.Add Key:=finalTable.ListColumns(5).Range, SortOn:=xlSortOnValues, Order:=xlAscending
        finalTable.Sort.SortFields.Add Key:=finalTable.ListColumns(7).Range, SortOn:=xlSortOnValues, Order:=xlDescending
        finalTable.Sort.Apply
        
        Set newColumn = finalTable.ListColumns.Add(Position:=19)
        newColumn.DataBodyRange.FormulaR1C1 = _
        "=IF(AND(" & _
            "NOT(EXACT(RC2,R[-1]C2)), " & _
            "ISBLANK(RC[-3])" & _
        "), """", " & _
        "IF(NOT(ISBLANK(RC[-3])), " & _
            "RC[-3], " & _
            "IF(AND(LEN(R[-1]C)<1, ISBLANK(RC[-3])), " & _
                """""", " & _
            "IF(AND(LEN(R[-1]C)>1, ISBLANK(RC[-3])), " & _
                "R[-1]C, " & _
                """)))" & _
        ")"

        newColumn.Name = "CARRY_PRN"
        Application.Calculate
        Set rng = newColumn.DataBodyRange
        rng.Value = rng.Value
        
        Set newColumn = finalTable.ListColumns.Add(Position:=20)
        newColumn.DataBodyRange.FormulaR1C1 = _
        "=IF(AND(" & _
                "NOT(EXACT(RC2,R[-1]C2)), " & _
                "ISBLANK(RC[-3])" & _
            "), """", " & _
            "IF(NOT(ISBLANK(RC[-3])), " & _
                "RC[-3], " & _
                "IF(AND(LEN(R[-1]C)<1, ISBLANK(RC[-3])), " & _
                    """""", " & _
                "IF(AND(LEN(R[-1]C)>1, ISBLANK(RC[-3])), " & _
                    "R[-1]C, " & _
                    """)))" & _
        ")" 

        newColumn.Name = "CARRY_SCORE"
        Application.Calculate
        Set rng = newColumn.DataBodyRange
        rng.Value = rng.Value

        Set newColumn = finalTable.ListColumns.Add(Position:=21)
        newColumn.DataBodyRange.FormulaR1C1 = _
        "=IF(AND(" & _
                "NOT(EXACT(RC2,R[-1]C2)), " & _
                "ISBLANK(RC[-3])" & _
            "), """", " & _
            "IF(NOT(ISBLANK(RC[-3])), " & _
                "RC[-3], " & _
                "IF(AND(LEN(R[-1]C)<1, ISBLANK(RC[-3])), " & _
                    """""", " & _
                "IF(AND(LEN(R[-1]C)>1, ISBLANK(RC[-3])), " & _
                    "R[-1]C, " & _
                    """)))" & _
        ")"

        newColumn.Name = "CARRY_INF"
        Application.Calculate
        Set rng = newColumn.DataBodyRange
        rng.Value = rng.Value
            
        Set newColumn = finalTable.ListColumns.Add(Position:=22)
                newColumn.DataBodyRange.FormulaR1C1 = _
            "=IFERROR(" & _
                "IF(OR(" & _
                    "LEN(RC[-4]) < 1, " & _
                    "AND(" & _
                        "EXACT(RC[-20],R[1]C[-20])," & _
                        "EXACT(RC[-17],R[1]C[-17])," & _
                        "NOT(ISERROR(SEARCH(""D"",R[1]C[-5])))" & _
                    ")" & _
                "), " & _
                    """""", " & _
                "IF(AND(" & _
                    "NOT(ISERROR(SEARCH(""X"",RC[-4]))), " & _
                    "((RC[-17] - VLOOKUP(RC[-2],scores,6,FALSE)) * 1440) > 45 " & _
                "), " & _
                    "CONCATENATE(" & _
                        "VLOOKUP(RC[-2],scores,7,FALSE) & "" charted @ "" & " & _
                        "TEXT(VLOOKUP(RC[-2],scores,6,FALSE),""mm/dd/yy hh:mm:ss"") & "" was not charted within 45 min (30-policy + 15-grace period) BEFORE the "" & " & _
                        "RC[-11] &"" rate-change to ("" & VLOOKUP(RC[-4],inf,10,FALSE) & "") charted @ "" & " & _
                        "TEXT(RC[-17], ""mm/dd/yy hh:mm:ss"") & "" ||"" & RC[-9] & ""||""" & _
                    "), " & _
                "IF(AND(" & _
                    "NOT(ISERROR(SEARCH(""Y"",RC[-4]))), " & _
                    "((RC[-17] - VLOOKUP(RC[-2],scores,6,FALSE)) * 1440) > 45 " & _
                "), " & _
                    "CONCATENATE(" & _
                        "VLOOKUP(RC[-2],scores,7,FALSE) & "" charted @ "" & " & _
                        "TEXT(VLOOKUP(RC[-2],scores,6,FALSE),""mm/dd/yy hh:mm:ss"") & "" was not charted within 45 min (30-policy + 15-grace period) BEFORE the final "" & " & _
                        "RC[-11] & "" record charted @ "" & TEXT(RC[-17], ""mm/dd/yy hh:mm:ss"") & " & _
                        "CHAR(10) & ""Check for accurate d/c time."" & "" ||"" & RC[-9] & ""||""" & _
                    "), " & _
                    """"" " & _
                ")" & _
            "))), " & _
            """"" " & _
            ")"

        newColumn.Name = "IRF1"
        Application.Calculate
        Set rng = newColumn.DataBodyRange
        rng.Value = rng.Value
        
        
        Set newColumn = finalTable.ListColumns.Add(Position:=23)
        newColumn.DataBodyRange.FormulaR1C1 = _
            "=IF(OR(" & _
                "LEN(RC[-5]) < 1, " & _
                "LEN(RC[-3]) > 1 " & _
            "), """", " & _
            "IF(AND(" & _
                "NOT(ISERROR(SEARCH(""X"",RC[-5]))), " & _
                "LEN(RC[-3]) < 1" & _
            "), " & _
                "CONCATENATE(" & _
                    """No pain-score observed prior to the "" & " & _
                    "RC[-15] & "" rate-change to ("" & RC[-14] & "") charted @ "" & " & _
                    "TEXT(RC[-18], ""mm/dd/yy hh:mm:ss"") & "" ||"" & RC[-10] & ""||""" & _
                "), " & _
            "IF(AND(" & _
                "NOT(ISERROR(SEARCH(""Y"",RC[-5]))), " & _
                "LEN(RC[-3]) < 1" & _
            "), " & _
                "CONCATENATE(" & _
                    """No pain-score observed prior to the final "" & " & _
                    "RC[-15] & "" record charted @ "" & " & _
                    "TEXT(RC[-18], ""mm/dd/yy hh:mm:ss"") & "" ||"" & RC[-10] & ""||""" & _
                "), " & _
                """)))"        
        newColumn.Name = "IRF2"
        Application.Calculate
        Set rng = newColumn.DataBodyRange
        rng.Value = rng.Value
        
        Set newColumn = finalTable.ListColumns.Add(Position:=24)
        
         newColumn.DataBodyRange.FormulaR1C1 = _
            "=IFERROR(" & _
                "IF(OR(" & _
                    "LEN(RC[-6]) < 1, " & _
                    "LEN(RC[-4]) < 1," & _
                    "VLOOKUP(RC[-4],scores,21,FALSE)<>""PARTIAL""," & _
                    "AND(ISERROR(SEARCH(""X"",RC[-6])),ISERROR(SEARCH(""Y"",RC[-6])))" & _
                "), """"," & _
                "CONCATENATE(" & _
                    """The previous pain assessment ("" & " & _
                    "TEXT(VLOOKUP(RC[-4],scores,6,FALSE),""mm/dd/yy hh:mm:ss"") & "") observed before "" & RC[-13] & "" rate/status change charted @ "" & " & _
                    "TEXT(RC[-19],""mm/dd/yy hh:mm:ss"") & "" was incomplete."") & CHAR(10) & "" ||"" & RC[-11] & ""||"" & CHAR(10), " & _
                    "IF(ISERROR(SEARCH(CodeKey!R2C5,VLOOKUP(RC[-4],scores,18,FALSE))),""'""&CodeKey!R2C3&"" ' is missing""&CHAR(10),"""") & " & _
                    "IF(ISERROR(SEARCH(CodeKey!R3C5,VLOOKUP(RC[-4],scores,18,FALSE))),""'""&CodeKey!R3C3&"" ' is missing""&CHAR(10),"""") & " & _
                    "IF(ISERROR(SEARCH(CodeKey!R4C5,VLOOKUP(RC[-4],scores,18,FALSE))),""'""&CodeKey!R4C3&"" ' is missing""&CHAR(10),"""") & " & _
                    "IF(ISERROR(SEARCH(CodeKey!R5C5,VLOOKUP(RC[-4],scores,18,FALSE))),""'""&CodeKey!R5C3&"" ' is missing""&CHAR(10),"""") & " & _
                    "IF(ISERROR(SEARCH(CodeKey!R6C5,VLOOKUP(RC[-4],scores,18,FALSE))),""'""&CodeKey!R6C3&"" ' is missing""&CHAR(10),"""") & " & _
                    "IF(ISERROR(SEARCH(CodeKey!R7C5,VLOOKUP(RC[-4],scores,18,FALSE))),""'""&CodeKey!R7C3&"" ' is missing""&CHAR(10),"""") & " & _
                    "IF(ISERROR(SEARCH(CodeKey!R8C5,VLOOKUP(RC[-4],scores,18,FALSE))),""'""&CodeKey!R8C3&"" ' is missing""&CHAR(10),"""")" & _
                "), " & _
                """"" " & _
            ")"

        newColumn.Name = "IRF3"
        Application.Calculate
        Set rng = newColumn.DataBodyRange
        rng.Value = rng.Value
            
        Set newColumn = finalTable.ListColumns.Add(Position:=25)
        
        newColumn.DataBodyRange.FormulaR1C1 = _
            "=IFERROR(" & _
                "IF(OR(" & _
                    "LEN(RC[-7]) < 1, " & _
                    "AND(" & _
                        "ISERROR(SEARCH(""X"",RC[-7])), " & _
                        "ISERROR(SEARCH(""Y"",RC[-7])), " & _
                        "ISERROR(SEARCH(""Z"",RC[-7]))" & _
                    ")" & _
                "), """"," & _
                "IF(((VLOOKUP(" & _
                    "CONCATENATE(RC[-23]&""-D""& SUM(RIGHT(RC[-5], LEN(RC[-5]) - SEARCH(""D"",RC[-5]))+1))," & _
                    "scores,6,FALSE) - RC[-20]) * 1440) > 75, " & _
                "CONCATENATE(" & _
                    """The follow-up pain assessment charted after "" & RC[-14] & "" rate/status change charted @ "" & " & _
                    "TEXT(RC[-20],""mm/dd/yy hh:mm:ss"") & "" was "" & " & _
                    "ROUND((((VLOOKUP(" & _
                        "CONCATENATE(RC[-23]&""-D""& SUM(RIGHT(RC[-5], LEN(RC[-5]) - SEARCH(""D"",RC[-5]))+1))," & _
                        "scores,6,FALSE) - RC[-20]) * 1440) - 75),0) & "" minutes late."" & CHAR(10) & ""The "" & " & _
                    "VLOOKUP(" & _
                        "CONCATENATE(RC[-23]&""-D""& SUM(RIGHT(RC[-5], LEN(RC[-5]) - SEARCH(""D"",RC[-5]))+1))," & _
                        "scores,7,FALSE) & "" was due within 75min (60-policy + 15-grace period) but was not charted until "" & " & _
                    "TEXT(" & _
                        "VLOOKUP(CONCATENATE(RC[-23]&""-D""& SUM(RIGHT(RC[-5], LEN(RC[-5]) - SEARCH(""D"",RC[-5]))+1))," & _
                        "scores,6,FALSE), ""mm/dd/yy hh:mm:ss"") & "" ||""& RC[-12]& ""||""" & _
                "), """")" & _
            "))"
        newColumn.Name = "IRF4"
        Application.Calculate
        Set rng = newColumn.DataBodyRange
        rng.Value = rng.Value
               
        Set newColumn = finalTable.ListColumns.Add(Position:=26)
        
        newColumn.DataBodyRange.FormulaR1C1 = _
            "=IFERROR(" & _
                "IF(OR(" & _
                    "LEN(RC[-8]) < 1, " & _
                    "MONTH(RC[-23]) > MONTH(RC[-21]), " & _
                    "AND(" & _
                        "ISERROR(SEARCH(""X"",RC[-8])), " & _
                        "ISERROR(SEARCH(""Y"",RC[-8])), " & _
                        "ISERROR(SEARCH(""Z"",RC[-8]))" & _
                    ")" & _
                "), """", " & _
                "IF(ISERROR(VLOOKUP(" & _
                    "CONCATENATE(RC[-24]&""-D""& SUM(RIGHT(RC[-6], LEN(RC[-6]) -SEARCH(""D"",RC[-6]))+1))," & _
                    "scores,6,FALSE)" & _
                "), " & _
                    "CONCATENATE(" & _
                        """No follow-up pain assessment observed after "" & RC[-15] & " & _
                        "IF(NOT(ISERROR(SEARCH(""X"", RC[-8]))), " & _
                            """ rate change charted @ "", " & _
                            """ final infusion record charted @ """ & _
                        ") & " & _
                        "TEXT(RC[-21],""mm/dd/yy hh:mm:ss"") & CHAR(10) & ""Patient discharged @ "" & " & _
                        "TEXT(RC[-23], ""mm/dd/yy hh:mm:ss"") & CHAR(10) & "" ||"" & RC[-13] & ""||""" & _
                    "), " & _
                    """"" " & _
                "))" & _
            ")" 
        newColumn.Name = "IRF5"
        Application.Calculate
        Set rng = newColumn.DataBodyRange
        rng.Value = rng.Value
                        
        Set newColumn = finalTable.ListColumns.Add(Position:=27)
        
        newColumn.DataBodyRange.FormulaR1C1 = _
            "=IF(OR(" & _
                "LEN(RC[-11]) < 1, " & _
                "AND(" & _
                    "EXACT(RC[-25],R[1]C[-25])," & _
                    "EXACT(RC[-22],R[1]C[-22])," & _
                    "NOT(ISERROR(SEARCH(""D"",R[1]C[-10])))" & _
                ")" & _
            "), """", " & _
            "IF( ( (RC[-22] - VLOOKUP(RC[-7],scores,6,FALSE)) * 1440) > 45, " & _
                "CONCATENATE(" & _
                    "VLOOKUP(RC[-7],scores,7,FALSE) & "" charted @ "" & " & _
                    "TEXT(VLOOKUP(RC[-7],scores,6,FALSE),""mm/dd/yy hh:mm:ss"") & "" was not charted within 45 min (30-policy + 15-grace period) BEFORE the "" & " & _
                    "RC[-16] & "" "" & RC[-21] & "" charted @ "" & " & _
                    "TEXT(RC[-22], ""mm/dd/yy hh:mm:ss"") & "" ||"" & RC[-14] & ""||""" & _
                "), " & _
                """"" " & _
            "))"        
        newColumn.Name = "PRF1"
        Application.Calculate
        Set rng = newColumn.DataBodyRange
        rng.Value = rng.Value
            
        Set newColumn = finalTable.ListColumns.Add(Position:=28)
        
        newColumn.DataBodyRange.FormulaR1C1 = _
            "=IF(OR(" & _
                "LEN(RC[-12]) < 1, " & _
                "LEN(RC[-8]) > 1 " & _
            "), """", " & _
            "IF(AND(" & _
                "NOT(ISERROR(SEARCH(""A"",RC[-12]))), " & _
                "LEN(RC[-8]) < 1 " & _
            "), " & _
                "CONCATENATE(" & _
                    """No pain-score observed prior to the "" & " & _
                    "RC[-17] & "" "" & RC[-22] & "" charted @ "" & " & _
                    "TEXT(RC[-23], ""mm/dd/yy hh:mm:ss"") & "" ||"" & RC[-15] & ""||""" & _
                "), " & _
                """"" " & _
            "))"

        newColumn.Name = "PRF2"
        Application.Calculate
        Set rng = newColumn.DataBodyRange
        rng.Value = rng.Value
        
        Set newColumn = finalTable.ListColumns.Add(Position:=29)
        
        newColumn.DataBodyRange.FormulaR1C1 = _
            "=IFERROR(" & _
                "IF(OR(" & _
                    "LEN(RC[-13]) < 1, " & _
                    "LEN(RC[-9]) < 1, " & _
                    "VLOOKUP(RC[-9],scores,21,FALSE)<>""PARTIAL"" " & _
                "), """"," & _
                "CONCATENATE(" & _
                    """The previous pain assessment observed before "" & RC[-18] & "" "" & RC[-23] & "" charted @ "" & " & _
                    "TEXT(RC[-24],""mm/dd/yy hh:mm:ss"") & "" was incomplete."") & CHAR(10) & "" ||"" & RC[-16] & ""||"" & CHAR(10), " & _
                    "IF(ISERROR(SEARCH(CodeKey!R2C5,VLOOKUP(RC[-9],scores,18,FALSE))), ""'"" & CodeKey!R2C3 & "" ' is missing"" & CHAR(10), """") & " & _
                    "IF(ISERROR(SEARCH(CodeKey!R3C5,VLOOKUP(RC[-9],scores,18,FALSE))), ""'"" & CodeKey!R3C3 & "" ' is missing"" & CHAR(10), """") & " & _
                    "IF(ISERROR(SEARCH(CodeKey!R4C5,VLOOKUP(RC[-9],scores,18,FALSE))), ""'"" & CodeKey!R4C3 & "" ' is missing"" & CHAR(10), """") & " & _
                    "IF(ISERROR(SEARCH(CodeKey!R5C5,VLOOKUP(RC[-9],scores,18,FALSE))), ""'"" & CodeKey!R5C3 & "" ' is missing"" & CHAR(10), """") & " & _
                    "IF(ISERROR(SEARCH(CodeKey!R6C5,VLOOKUP(RC[-9],scores,18,FALSE))), ""'"" & CodeKey!R6C3 & "" ' is missing"" & CHAR(10), """") & " & _
                    "IF(ISERROR(SEARCH(CodeKey!R7C5,VLOOKUP(RC[-9],scores,18,FALSE))), ""'"" & CodeKey!R7C3 & "" ' is missing"" & CHAR(10), """") & " & _
                    "IF(ISERROR(SEARCH(CodeKey!R8C5,VLOOKUP(RC[-9],scores,18,FALSE))), ""'"" & CodeKey!R8C3 & "" ' is missing"" & CHAR(10), """")" & _
                "), " & _
                """"" " & _
            ")"        
        newColumn.Name = "PRF3"
        Application.Calculate
        Set rng = newColumn.DataBodyRange
        rng.Value = rng.Value
                
        Set newColumn = finalTable.ListColumns.Add(Position:=30)
        
        newColumn.DataBodyRange.FormulaR1C1 = _
            "=IFERROR(" & _
                "IF(OR(" & _
                    "LEN(RC[-14]) < 1, " & _
                    "ISERROR(SEARCH(""A"",RC[-14])) " & _
                "), """", " & _
                "IF(((VLOOKUP(" & _
                    "CONCATENATE(RC[-28]&""-D""& SUM(RIGHT(RC[-10], LEN(RC[-10]) - SEARCH(""D"",RC[-10]))+1))," & _
                    "scores,6,FALSE) - RC[-25]) * 1440) > 45, " & _
                "CONCATENATE(" & _
                    """The follow-up pain assessment charted after "" & RC[-19] & "" "" & RC[-24] & "" charted @ "" & " & _
                    "TEXT(RC[-25],""mm/dd/yy hh:mm:ss"") & "" was "" & " & _
                    "ROUND((((VLOOKUP(" & _
                        "CONCATENATE(RC[-28]&""-D""& SUM(RIGHT(RC[-10], LEN(RC[-10]) - SEARCH(""D"",RC[-10]))+1))," & _
                        "scores,6,FALSE) - RC[-25]) * 1440) - 45),0) & "" minutes late."" & CHAR(10) & ""The "" & " & _
                    "VLOOKUP(" & _
                        "CONCATENATE(RC[-28]&""-D""& SUM(RIGHT(RC[-10], LEN(RC[-10]) - SEARCH(""D"",RC[-10]))+1))," & _
                        "scores,7,FALSE) & "" was due within 45min (30-policy + 15-grace period) but was not charted until "" & " & _
                    "TEXT(" & _
                        "VLOOKUP(CONCATENATE(RC[-28]&""-D""& SUM(RIGHT(RC[-10], LEN(RC[-10]) - SEARCH(""D"",RC[-10]))+1))," & _
                        "scores,6,FALSE), ""mm/dd/yy hh:mm:ss"") & "" ||"" & RC[-17] & ""||""" & _
                "), """")" & _
            "))"                
        newColumn.Name = "PRF4"
        Application.Calculate
        Set rng = newColumn.DataBodyRange
        rng.Value = rng.Value
                
        Set newColumn = finalTable.ListColumns.Add(Position:=31)
        
        newColumn.DataBodyRange.FormulaR1C1 = _
            "=IFERROR(" & _
                "IF(OR(" & _
                    "LEN(RC[-15]) < 1, " & _
                    "ISERROR(SEARCH(""B"",RC[-15])) " & _
                "), """", " & _
                "IF(((VLOOKUP(" & _
                    "CONCATENATE(RC[-29]&""-D""& SUM(RIGHT(RC[-11], LEN(RC[-11]) - SEARCH(""D"",RC[-11]))+1))," & _
                    "scores,6,FALSE) - RC[-26]) * 1440) > 75, " & _
                "CONCATENATE(" & _
                    """The follow-up pain assessment charted after "" & RC[-20] & "" "" & RC[-25] & "" charted @ "" & " & _
                    "TEXT(RC[-26],""mm/dd/yy hh:mm:ss"") & "" was "" & " & _
                    "ROUND((((VLOOKUP(" & _
                        "CONCATENATE(RC[-29]&""-D""& SUM(RIGHT(RC[-11], LEN(RC[-11]) - SEARCH(""D"",RC[-11]))+1))," & _
                        "scores,6,FALSE) - RC[-26]) * 1440) - 75),0) & "" minutes late."" & CHAR(10) & ""The "" & " & _
                    "VLOOKUP(" & _
                        "CONCATENATE(RC[-29]&""-D""& SUM(RIGHT(RC[-11], LEN(RC[-11]) - SEARCH(""D"",RC[-11]))+1))," & _
                        "scores,7,FALSE) & "" was due within 75min (60-policy + 15-grace period) but was not charted until "" & " & _
                    "TEXT(" & _
                        "VLOOKUP(CONCATENATE(RC[-29]&""-D""& SUM(RIGHT(RC[-11], LEN(RC[-11]) - SEARCH(""D"",RC[-11]))+1))," & _
                        "scores,6,FALSE), ""mm/dd/yy hh:mm:ss"") & "" ||"" & RC[-18] & ""||""" & _
                "), """")" & _
            "))"                
        newColumn.Name = "PRF5"
        Application.Calculate
        Set rng = newColumn.DataBodyRange
        rng.Value = rng.Value
          
        Set newColumn = finalTable.ListColumns.Add(Position:=32)
        
        newColumn.DataBodyRange.FormulaR1C1 = _
            "=IFERROR(" & _
                "IF(LEN(RC[-16]) < 1, """", " & _
                "IF(ISERROR(VLOOKUP(" & _
                    "CONCATENATE(RC[-30]&""-D""& SUM(RIGHT(RC[-12], LEN(RC[-12]) - SEARCH(""D"",RC[-12]))+1))," & _
                    "scores,6,FALSE)" & _
                "), " & _
                    "CONCATENATE(" & _
                        """No follow-up pain assessment observed after "" & RC[-21] & "" "" & RC[-26] & "" charted @ "" & " & _
                        "TEXT(RC[-27],""mm/dd/yy hh:mm:ss"") & CHAR(10) & ""Patient discharged @ "" & " & _
                        "TEXT(RC[-29], ""mm/dd/yy hh:mm:ss"") & CHAR(10) & "" ||"" & RC[-19] & ""||""" & _
                    "), " & _
                    """"" " & _
                "))" & _
            ")"                
        newColumn.Name = "PRF6"
        Application.Calculate
        Set rng = newColumn.DataBodyRange
        rng.Value = rng.Value
            
        Set newColumn = finalTable.ListColumns.Add(Position:=33)
        
        newColumn.DataBodyRange.FormulaR1C1 = _
            "=IFERROR(" & _
                "IF(OR(" & _
                    "LEN(RC[-17]) < 1, " & _
                    "RIGHT(RC[-13],2) = ""D1"", " & _
                    "NOT(RC[-28] < VLOOKUP(" & _
                        "CONCATENATE(RC[-31]&""-D""& SUM(RIGHT(RC[-13], LEN(RC[-13]) - SEARCH(""D"",RC[-13]))+1))," & _
                        "scores,6,FALSE)" & _
                    "), " & _
                    "VALUE(VLOOKUP(RC[-13], scores,10,FALSE)) = 0 " & _
                "), """"," & _
                "IF((VLOOKUP(RC[-13], scores,10,FALSE) - VLOOKUP(" & _
                    "CONCATENATE(RC[-31]&""-D""& SUM(RIGHT(RC[-13], LEN(RC[-13]) - SEARCH(""D"",RC[-13]))+1))," & _
                    "scores,10,FALSE)" & _
                ") < 1, " & _
                "CONCATENATE(" & _
                    """The "" & VLOOKUP(RC[-13], scores,7,FALSE) & "" charted at "" & " & _
                    "TEXT(VLOOKUP(RC[-13], scores, 6, FALSE), ""mm/dd/yy hh:mm:ss"") & "" = "" & VLOOKUP(RC[-13], scores, 10, FALSE) & CHAR(10) & " & _
                    """The "" & VLOOKUP(" & _
                        "CONCATENATE(RC[-31]&""-D""& SUM(RIGHT(RC[-13], LEN(RC[-13]) - SEARCH(""D"",RC[-13]))+1))," & _
                        "scores,7,FALSE)" & _
                    " & "" charted at "" & " & _
                    "TEXT(" & _
                        "VLOOKUP(CONCATENATE(RC[-31]&""-D""& SUM(RIGHT(RC[-13], LEN(RC[-13]) - SEARCH(""D"",RC[-13]))+1))," & _
                        "scores,6,FALSE), ""mm/dd/yy hh:mm:ss"") & "" = "" & " & _
                    "VLOOKUP(" & _
                        "CONCATENATE(RC[-31]&""-D""& SUM(RIGHT(RC[-13], LEN(RC[-13]) - SEARCH(""D"",RC[-13]))+1))," & _
                        "scores,10,FALSE)" & _
                    " & CHAR(10) & ""The "" & RC[-22] & "" "" & RC[-27] & "" given at "" & " & _
                    "TEXT(RC[-28], ""mm/dd/yy hh:mm:ss"") & "" was ineffective.""" & _
                "), """")" & _
            ")"                
        newColumn.Name = "PRF7"
        Application.Calculate
        Set rng = newColumn.DataBodyRange
        rng.Value = rng.Value
        
        Set newColumn = finalTable.ListColumns.Add(Position:=34)
        
        newColumn.DataBodyRange.FormulaR1C1 = _
            "=IFERROR(" & _
                "IF(AND(" & _
                    "NOT(ISERROR(VLOOKUP(CONCATENATE(RC[-32] & ""-W""),inf,6,FALSE))), " & _
                    "LEN(RC[-13]) > 1, " & _
                    "((RC[-29] - VLOOKUP(CONCATENATE(RC[-32] & ""-W""),inf,6,FALSE)) * 1440) < 1440, " & _
                    "((RC[-29] - VLOOKUP(CONCATENATE(RC[-32] & ""-W""),inf,6,FALSE)) * 1440) > 135, " & _
                    "RC[-29] < VLOOKUP(CONCATENATE(RC[-32]&""-Y""),inf,6,FALSE), " & _
                    "LEN(RC[-17])>1, " & _
                    "(RC[-29] - VLOOKUP(" & _
                        "CONCATENATE(RC[-32] & ""-D"" & SUM(RIGHT(RC[-14], LEN(RC[-14])-SEARCH(""D"",RC[-14]))-1))," & _
                        "scores,6,FALSE))*1440 >135" & _
                "), " & _
                "CONCATENATE(" & _
                    "ROUND((RC[-29] - VLOOKUP(" & _
                        "CONCATENATE(RC[-32] & ""-D"" & SUM(RIGHT(RC[-14], LEN(RC[-14])-SEARCH(""D"",RC[-14]))-1))," & _
                        "scores,6,FALSE))*1440,0) & "" minutes elapsed between pain scores while patient has active "" & " & _
                    "VLOOKUP(CONCATENATE(RC[-32] & ""-W""),inf, 12, FALSE) & "" within first 24 hours."" & CHAR(10) & " & _
                    """First infusion record charted @ "" & TEXT(VLOOKUP(CONCATENATE(RC[-32] & ""-W""),inf,6,FALSE), ""mm/dd/yy hh:mm:ss"") & CHAR(10) & " & _
                    """Previous score charted @ "" & TEXT(" & _
                        "VLOOKUP(CONCATENATE(RC[-32] & ""-D"" & SUM(RIGHT(RC[-14], LEN(RC[-14])-SEARCH(""D"",RC[-14]))-1))," & _
                        "scores,6,FALSE), ""mm/dd/yy hh:mm:ss"") & CHAR(10) & " & _
                    """Flagged score charted @ "" & TEXT(RC[-29], ""mm/dd/yy hh:mm:ss"") & "")"" & CHAR(10) & "" ||"" & RC[-21] & ""||""" & _
                "), " & _
                """"" " & _
            "))"                
        newColumn.Name = "SRF1"
        Application.Calculate
        Set rng = newColumn.DataBodyRange
        rng.Value = rng.Value
           
        Set newColumn = finalTable.ListColumns.Add(Position:=35)
        
        newColumn.DataBodyRange.FormulaR1C1 = _
            "=IFERROR(" & _
                "IF(AND(" & _
                    "NOT(ISERROR(VLOOKUP(CONCATENATE(RC[-33] & ""-W""),inf,6,FALSE))), " & _
                    "LEN(RC[-14]) > 1, " & _
                    "((RC[-30] - VLOOKUP(CONCATENATE(RC[-33] & ""-W""),inf,6,FALSE)) * 1440) > 1440, " & _
                    "RC[-30] < VLOOKUP(CONCATENATE( RC[-33]&""-Y""),inf,6,FALSE), " & _
                    "LEN(RC[-18])>1, " & _
                    "(RC[-30] - VLOOKUP(" & _
                        "CONCATENATE(RC[-33] & ""-D"" & SUM(RIGHT(RC[-15], LEN(RC[-15])-SEARCH(""D"",RC[-15]))-1))," & _
                        "scores,6,FALSE))*1440 >255" & _
                "), " & _
                "CONCATENATE(" & _
                    "ROUND((RC[-30] - VLOOKUP(" & _
                        "CONCATENATE(RC[-33] & ""-D"" & SUM(RIGHT(RC[-15], LEN(RC[-15])-SEARCH(""D"",RC[-15]))-1))," & _
                        "scores,6,FALSE))*1440,0) & "" minutes elapsed between pain scores while patient has active "" & " & _
                    "VLOOKUP(CONCATENATE(RC[-33] & ""-W""),inf, 12, FALSE) & "" after the first 24 hours."" & CHAR(10) & " & _
                    """First infusion record charted @ "" & TEXT(VLOOKUP(CONCATENATE(RC[-33] & ""-W""),inf,6,FALSE), ""mm/dd/yy hh:mm:ss"") & CHAR(10) & " & _
                    """Previous score charted @ "" & TEXT(" & _
                        "VLOOKUP(CONCATENATE(RC[-33] & ""-D"" & SUM(RIGHT(RC[-15], LEN(RC[-15])-SEARCH(""D"",RC[-15]))-1))," & _
                        "scores,6,FALSE), ""mm/dd/yy hh:mm:ss"") & CHAR(10) & " & _
                    """Flagged score charted @ "" & TEXT(RC[-30], ""mm/dd/yy hh:mm:ss"") & "")"" & CHAR(10) & "" ||"" & RC[-22] & ""||""" & _
                "), " & _
                """"" " & _
            "))"                
        newColumn.Name = "SRF2"
        Application.Calculate
        Set rng = newColumn.DataBodyRange
        rng.Value = rng.Value
        
        Set newColumn = finalTable.ListColumns.Add(Position:=36)
        
        newColumn.DataBodyRange.FormulaR1C1 = _
            "=IFERROR(" & _
                "IF(AND(" & _
                    "RIGHT(RC[-16],2)<>""D1"", " & _
                    "((RC[-31] - VLOOKUP(" & _
                        "CONCATENATE(RC[-34]&""-D""& SUM(RIGHT(RC[-16], LEN(RC[-16]) - SEARCH(""D"",RC[-16]))-1))," & _
                        "scores,6,FALSE)) * 1440) > 735, " & _
                    "LEN(RC[-19])>1 " & _
                "), " & _
                "CONCATENATE(" & _
                    "ROUND((RC[-31] - VLOOKUP(" & _
                        "CONCATENATE(RC[-34] & ""-D"" & SUM(RIGHT(RC[-16], LEN(RC[-16])-SEARCH(""D"",RC[-16]))-1))," & _
                        "scores,6,FALSE))*1440,0) & "" minutes elapsed between pain scores, exceeding the 12-hr per-shift limit "" & CHAR(10) & " & _
                    """Previous score charted @ "" & TEXT(" & _
                        "VLOOKUP(CONCATENATE(RC[-34] & ""-D"" & SUM(RIGHT(RC[-16], LEN(RC[-16])-SEARCH(""D"",RC[-16]))-1))," & _
                        "scores,6,FALSE), ""mm/dd/yy hh:mm:ss"") & CHAR(10) & " & _
                    """Flagged score charted @ "" & TEXT(RC[-31], ""mm/dd/yy hh:mm:ss"") & "")"" & CHAR(10) & "" ||"" & RC[-23] & ""||""" & _
                "), " & _
                """"" " & _
            ")"                
        newColumn.Name = "SRF3"
        Application.Calculate
        Set rng = newColumn.DataBodyRange
        rng.Value = rng.Value
          
        Set newColumn = finalTable.ListColumns.Add(Position:=37)
        
        newColumn.DataBodyRange.FormulaR1C1 = _
            "=CONCATENATE(" & _
                "IF(LEN(RC[-14])>1, CONCATENATE(RC[-14] & CHAR(10)),"""") & " & _
                "IF(LEN(RC[-13])>1, CONCATENATE(RC[-13] & CHAR(10)),"""") & " & _
                "IF(LEN(RC[-11])>1, CONCATENATE(RC[-11] & CHAR(10)),"""") & " & _
                "IF(LEN(RC[-9])>1, CONCATENATE(RC[-9] & CHAR(10)),"""") & " & _
                "IF(LEN(RC[-8])>1, CONCATENATE(RC[-8] & CHAR(10)),"""") & " & _
                "IF(LEN(RC[-5])>1, CONCATENATE(RC[-5] & CHAR(10)),"""")" & _
            ")"                
        newColumn.Name = "COMP"
        Application.Calculate
        Set rng = newColumn.DataBodyRange
        rng.Value = rng.Value
            
        Set newColumn = finalTable.ListColumns.Add(Position:=38)
        
        newColumn.DataBodyRange.FormulaR1C1 = _
            "=CONCATENATE(" & _
                "IF(LEN(RC[-16])>1, CONCATENATE(RC[-16] & CHAR(10)),"""") & " & _
                "IF(LEN(RC[-13])>1, CONCATENATE(RC[-13] & CHAR(10)),"""") & " & _
                "IF(LEN(RC[-11])>1, CONCATENATE(RC[-11] & CHAR(10)),"""") & " & _
                "IF(LEN(RC[-8])>1, CONCATENATE(RC[-8] & CHAR(10)),"""") & " & _
                "IF(LEN(RC[-7])>1, CONCATENATE(RC[-7] & CHAR(10)),"""") & " & _
                "IF(LEN(RC[-4])>1, CONCATENATE(RC[-4] & CHAR(10)),"""") & " & _
                "IF(LEN(RC[-3])>1, CONCATENATE(RC[-3] & CHAR(10)),"""") & " & _
                "IF(LEN(RC[-2])>1, CONCATENATE(RC[-2] & CHAR(10)),"""")" & _
            ")"                
        newColumn.Name = "TIMING"
        Application.Calculate
        Set rng = newColumn.DataBodyRange
        rng.Value = rng.Value
             
        Set newColumn = finalTable.ListColumns.Add(Position:=39)
        
        newColumn.DataBodyRange.FormulaR1C1 = "=RC[-6]"
                
        newColumn.Name = "EFF"
        Application.Calculate
        Set rng = newColumn.DataBodyRange
        rng.Value = rng.Value
          
        Set newColumn = finalTable.ListColumns.Add(Position:=40)
        
        newColumn.DataBodyRange.FormulaR1C1 = _
            "=IF(OR(" & _
                "LEN(RC[-3])>1," & _
                "LEN(RC[-2])>1, " & _
                "LEN(RC[-1])>1" & _
            "), 1, """")"
               
        newColumn.Name = "ANY"
        Application.Calculate
        Set rng = newColumn.DataBodyRange
        rng.Value = rng.Value
        
        finalTable.Range.AutoFilter Field:=40, Criteria1:="<>"
        
        finalSheet.Range("C:AJ").EntireColumn.Hidden = True

        finalSheet.Range("AN:AN").EntireColumn.Hidden = True
        
        staffKeyPath = PERM_PATH & "Files\INPATIENT\Templates\IP_Staff_Key.xlsm"
        reviewTemplatePath = PERM_PATH & "Files\INPATIENT\Templates\IP_Review.Template.xlsm"
       
        Set wbTemplate = Workbooks.Open(reviewTemplatePath)
        Set wbStaffKey = Workbooks.Open(staffKeyPath)
        Set wsStaffKey = wbStaffKey.Sheets("Staff_Key")
        Set lastWs = wbTemplate.Sheets(wbTemplate.Sheets.Count)
        wsStaffKey.Copy After:=lastWs
        wbStaffKey.Close SaveChanges:=False
           
        Windows("IP_Review.Template.xlsm").Activate
        Sheets("Staff_Key").Select
        
        Set ws = ActiveWorkbook.Sheets("Staff_Key")
      
        For Each btn In ws.Shapes
            If Left(btn.Name, 6) = "Button" Then
                btn.Delete
            End If
        Next btn
        
        ws.Visible = xlSheetHidden
        
        Set wsStaffKey = Nothing
        Set lastWs = Nothing
        Set wbStaffKey = Nothing
        Set wbTemplate = Nothing

        Set sourceBook = Workbooks("IP_Calc.Thread.xlsm")
        Set targetBook = Workbooks("IP_Review.Template.xlsm")
        Set sourceSheet = sourceBook.Worksheets("CodeKey")
        Set targetSheet = targetBook.Worksheets(targetBook.Worksheets.Count)
        sourceSheet.Copy After:=targetSheet
        Set sourceSheet = targetBook.Worksheets("CodeKey")
        sourceSheet.Visible = xlSheetHidden

        Set sourceBook = Nothing
        Set targetBook = Nothing
        Set sourceSheet = Nothing
        Set targetSheet = Nothing
             
        Windows("IP_Review.Template.xlsm").Activate

        Set ws = ActiveWorkbook.Worksheets("DATA2")
        ws.Visible = xlSheetVisible
               
        Windows("IP_Calc.Thread.xlsm").Activate
        
        Range("calcs[[#Headers],[NAME]:[EFF]]").Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.SpecialCells(xlCellTypeVisible).Select
        Application.CutCopyMode = False
        Selection.Copy
        
        Windows("IP_Review.Template.xlsm").Activate

        Sheets("DATA2").Select
        ActiveSheet.Range("C8").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Application.CutCopyMode = False
        ws.Visible = xlSheetHidden
        
        Set ws = ActiveWorkbook.Worksheets("DATA")
        Set rng = ws.Range("G1")
      
        ws.Visible = xlSheetVisible
        rng.Value = admin
        ws.Visible = xlSheetHidden

        Application.DisplayAlerts = False
        
        Windows("Pain_Audit_Report.xlsx").Activate
        
        Dim jparName, jcalName, jrevName As String
        
        If jop1 = "TEST" Then
        
        jparName = fn & "_RAW_TEST.xlsm"
        jcalName = fn & "_CALC_TEST.xlsm"
        jrevName = fn & "_REVIEW_TEST.xlsm"
        
        Else
        
        jparName = fn & "_RAW.xlsm"
        jcalName = fn & "_CALC.xlsm"
        jrevName = fn & "_REVIEW.xlsm"
        
        End If
        Dim killPath As String: killPath = parPath
        
        randomID = GenerateShortID(12)
        
        parPath = PERM_PATH & "Files\INPATIENT\Reports\RAW\" & randomID & ".xlsm"
        calcPath = PERM_PATH & "Files\INPATIENT\Reports\CALC\" & randomID & ".xlsm"
        revPath = PERM_PATH & "Files\INPATIENT\Reports\REVIEW\" & randomID & ".xlsm"
    
        ActiveWorkbook.SaveAs Filename:= _
            parPath, FileFormat:=xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False
        ActiveWorkbook.Close
        
        Kill killPath
        
        Windows("IP_Calc.Thread.xlsm").Activate
        
        ActiveWorkbook.SaveAs Filename:= _
            calcPath, FileFormat:=xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False
        ActiveWorkbook.Close
        
        Windows("IP_Review.Template.xlsm").Activate
        ActiveWorkbook.Sheets("REVIEW").Activate
        ActiveSheet.Range("C2").Select
        
        recipient = IIf(job1 = "TEST", admin, manager)
            Set accVer = ab.Worksheets("ACCOUNTS_VER")
            Set verTbl = accVer.ListObjects("verified")
            Set verColumn1 = verTbl.ListColumns(2)
            Set verColumn2 = verTbl.ListColumns(3)

        For Each verRow In verTbl.ListRows
        
            If verRow.Range(verColumn1.Index).Value = recipient Then
            routePath = verRow.Range(verColumn2.Index).Value
            End If
        Next verRow
        
        Dim cs2 As Worksheet: Set cs2 = ActiveWorkbook.Sheets("COVERSHEET")
        
        Set coverSheet = cs2.Range("coversheet")
        coverSheet.Cells(1, 2).Value = Now
        coverSheet.Cells(2, 2).Value = recipient
        coverSheet.Cells(3, 2).Value = admin
        coverSheet.Cells(4, 2).Value = ""
        coverSheet.Cells(5, 2).Value = randomID
        coverSheet.Cells(6, 2).Value = jrevName
        coverSheet.Cells(7, 2).Value = "IP_REVTEMP_SEND"
        coverSheet.Cells(8, 2).Value = routePath
        coverSheet.Cells(9, 2).Value = ""
        coverSheet.Cells(10, 2).Value = PERM_PATH
        coverSheet.Cells(11, 2).Value = UserName
      
        ActiveWorkbook.SaveAs Filename:= _
            revPath, _
            FileFormat:=xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False
        ActiveWorkbook.Close
            
   Set OutApp = CreateObject("Outlook.Application")
            Set OutMail = OutApp.CreateItem(0)
            
       recipient = IIf(job1 = "TEST", admin, manager)
  
        With OutMail
            .To = recipient
            .Subject = "AuditBot: Inpatient Flagging Report ready for review"
            .Body = "The latest Inpatient Audit Flagging Report (" & jrevName & ") is attached and ready for your review! Thanks!"
            .Attachments.Add revPath
            .Send
        End With
            
            Set olTask = OutApp.CreateItem(3)
        
            If job1 <> "TEST" Then
            olTask.Assign
            Set myDelegate = olTask.Recipients.Add(recipient)
            myDelegate.Resolve
            
            End If
            
            olTask.Subject = "AuditBot: Review & Approve Inpatient Flagging Report"
            olTask.Body = "The latest Inpatient Audit Flagging Report (" & jrevName & ") is attached and ready for your review! Thanks!"
            olTask.DueDate = Date + 7
            olTask.Importance = 2
            olTask.Categories = "AuditBot_IP"
        
            olTask.Save
            
            If job1 <> "TEST" Then
        
            olTask.Send
            
            End If

            Set OutMail = Nothing
            Set OutApp = Nothing
        
            logPath = PERM_PATH & "Logs\Message_Logs.xlsm"
            Set wb = Workbooks.Open(logPath)
            Set logz = wb.Sheets("LOG")
            Set tbl = logz.ListObjects("log")
            Set tblRow = tbl.ListRows.Add
            
        With tblRow.Range
            .Columns("A").Value = UserName
            .Columns("B").Value = Now
            .Columns("C").Value = recipient
            .Columns("D").Value = "REQUEST"
            .Columns("E").Value = IIf(job1 = "TEST", "IP_Test", "IP_Review")
            .Columns("F").Value = revPath
        End With
        
        ActiveWorkbook.Save
        ActiveWorkbook.Close
     
    Application.DisplayAlerts = True
       
    auditBotPath = PERM_PATH & "AuditBot.xlsm"
    
    On Error Resume Next
    Set wb = GetObject(auditBotPath)
    On Error GoTo 0
    
    If wb Is Nothing Then
        
        Set wb = Workbooks.Open(auditBotPath)
    Else
        wb.Activate
    End If

    Set lastWs = ActiveWorkbook.Worksheets("IP_HIST")
    Set tbl = lastWs.ListObjects("ip_hx")
    
    Set newRow = tbl.ListRows.Add
    newRow.Range(1, 1).Value = Now
    newRow.Range(1, 2).Value = randomID
    newRow.Range(1, 3).Value = jparName
    newRow.Range(1, 4).Value = jcalName
    newRow.Range(1, 5).Value = jrevName

    newRow.Range(1, 3).Hyperlinks.Add _
        Anchor:=newRow.Range(1, 3), _
        Address:=parPath, _
        TextToDisplay:=newRow.Range(1, 3).Value
        
    newRow.Range(1, 4).Hyperlinks.Add _
        Anchor:=newRow.Range(1, 4), _
        Address:=calcPath, _
        TextToDisplay:=newRow.Range(1, 4).Value
        
    newRow.Range(1, 5).Hyperlinks.Add _
        Anchor:=newRow.Range(1, 5), _
        Address:=revPath, _
        TextToDisplay:=newRow.Range(1, 5).Value

        Workbooks.Open Filename:= _
            PERM_PATH & "Logs\Run_Logs.xlsm"
        Set ws = ActiveWorkbook.Worksheets("LOG")
        Set tbl = ws.ListObjects("log")
        tbl.ListRows.Add
        Range("A1").Select
        Selection.End(xlDown).Offset(1, 0).Select
        ActiveCell.Value = currentDate
        Range("B1").Select
        Selection.End(xlDown).Offset(1, 0).Select
        ActiveCell.Value = "Inpatient"
        Range("C1").Select
        Selection.End(xlDown).Offset(1, 0).Select
        ActiveCell.Value = UserName
        endTime = Timer
        elapsedTime = endTime - startTime
        Range("D1").Select
        Selection.End(xlDown).Offset(1, 0).Select
        ActiveCell.Value = elapsedTime
        ActiveWorkbook.Save
        ActiveWorkbook.Close
        
        MsgBox "Elapsed Time: " & elapsedTime & "   (not too shabby...anyway, links to your raw data, calculations, and review worksheets are in the IP History Summary on the next tab..."
        
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        
End Sub
