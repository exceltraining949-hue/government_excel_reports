Attribute VB_Name = "HighlightDuplicates"
Option Explicit

'==============================================================================
' HIGHLIGHT & SUMMARIZE DUPLICATES  (reviewed & corrected version)
'------------------------------------------------------------------------------
' Duplicate key : Name (B) + Age (C) + Department (D) + Salary (E)
'                 Trim() ke baad EXACT match, 1 se zyada baar = duplicate.
'                 Name akela duplicate NAHI hota. Case-sensitive.
' Kya hota hai   : Duplicate rows A:J tak yellow highlight hoti hain (values
'                 delete/modify NAHI hote), aur "Duplicates" sheet me har
'                 unique combination ki ek summary row banti hai:
'                 Name, Age, Department, Salary, Duplicate Count, Total Salary.
' Total Salary   = Salary x Duplicate Count   (Salary key ka hissa hai, is
'                 liye group ki tamam rows ka salary same hota hai).
'
' ORIGINAL CODE ME JO BUGS FIX KIYE GAYE:
'  1. Headers A1:J1 par likhe ja rahe the (6 values, 10 cells) -> A1:F1 kiya.
'  2. NumberFormat "#,##0" D:F par lag raha tha -> isme Duplicate Count (E)
'     bhi aa jata. Ab SIRF Salary (D) aur Total Salary (F) par lagta hai.
'  3. Total Salary pehli baar salary likh kar baad mein += kar ke jama hota
'     tha -> text salary par Type Mismatch error (run beech me toot jata).
'     Ab spec ke mutabiq ek hi baar: CDbl(Salary) x Count. Non-numeric salary
'     ho to "n/a" likha jata hai aur message me honestly report hota hai.
'  4. Required columns (B1:E1 headers) ka koi validation nahi tha ->
'     ab Name/Age/Department/Salary check hota hai, missing ho to macro
'     ruk jata hai aur saaf batata hai kaun se columns chahiye.
'  5. "No duplicates" case me bhi "0 rows found" wala message aa raha tha ->
'     ab clear "No duplicate records were found."
'  6. Protected sheet/workbook par macro beech me fail hota tha ->
'     ab pehle check hota hai aur wajah batata hai (claim success nahi karta).
'  7. ActiveSheet chart sheet ho, ya macro khud "Duplicates" sheet par chalaye
'     -> guard add kiya.
' 8. Run ke BAAD self-validation: highlighted rows ki recount, summary rows
'     ki recount, har summary row ka Total Salary = Salary x Count verify.
'     Koi check fail ho to message me saaf likha jata hai.
'  9. ScreenUpdating/DisplayAlerts ab hamesha restore hote hain (error par bhi).
' 10. Option Explicit + proper error handler (partial run ki honesty ke saath).
'
' NOTE: Poori tarah blank rows (B/C/D/E sab khali) aapas me "duplicate" count
' hongi (key "|||") — ye di gayi logic ka natural behaviour hai.
'==============================================================================

Private Const DUP_SHEET As String = "Duplicates"

Public Sub HighlightAndSummarizeDuplicates()
    Dim ws As Worksheet, dupWs As Worksheet, wb As Workbook
    Dim lastRow As Long, i As Long, outputRow As Long, r As Long
    Dim key As String
    Dim dict As Object, summaryDict As Object
    Dim duplicateCount As Long
    Dim nonNumeric As Long
    Dim screenWasOn As Boolean

    ' ---------- guards (honest error handling) ----------
    If TypeName(ActiveSheet) <> "Worksheet" Then
        MsgBox "Active sheet ek worksheet nahi hai (chart sheet lagti hai)." & vbCrLf & _
               "Data worksheet par click kar ke macro dobara chalaein.", _
               vbExclamation, "Highlight & Summarize Duplicates"
        Exit Sub
    End If
    Set ws = ActiveSheet
    Set wb = ws.Parent

    If ws.Name = DUP_SHEET Then
        MsgBox "Aap abhi '" & DUP_SHEET & "' (summary) sheet par hain." & vbCrLf & _
               "Macro DATA wali sheet par chalaein.", _
               vbExclamation, "Highlight & Summarize Duplicates"
        Exit Sub
    End If

    ' required columns: B=Name, C=Age, D=Department, E=Salary (headers row 1)
    Dim reqHeaders As Variant, missing As String
    reqHeaders = Array("NAME", "AGE", "DEPARTMENT", "SALARY")
    For i = 0 To 3
        If UCase$(Trim$(CStr(ws.Cells(1, 2 + i).Value))) <> CStr(reqHeaders(i)) Then
            missing = missing & vbCrLf & "  Column " & Chr$(66 + i) & " : " & reqHeaders(i)
        End If
    Next i
    If Len(missing) > 0 Then
        MsgBox "Required columns missing ya headers match nahi kar rahe:" & vbCrLf & _
               "Expected ->" & missing & vbCrLf & vbCrLf & _
               "Row 1 me headers rakhein (B=Name, C=Age, D=Department, E=Salary) " & _
               "aur dobara chalaein.", _
               vbExclamation, "Required Columns Missing"
        Exit Sub
    End If

    If ws.ProtectContents Then
        MsgBox "Worksheet '" & ws.Name & "' protected hai." & vbCrLf & _
               "Highlighting add nahi ho sakti jab tak sheet unprotect na ho." & vbCrLf & _
               "Review tab se Unprotect kar ke dobara chalaein.", _
               vbExclamation, "Worksheet Protected"
        Exit Sub
    End If
    If wb.ProtectStructure Then
        MsgBox "Workbook ki structure protected hai — '" & DUP_SHEET & _
               "' sheet delete/create nahi ho sakti." & vbCrLf & _
               "Review > Protect Workbook se unprotect kar ke dobara chalaein.", _
               vbExclamation, "Workbook Protected"
        Exit Sub
    End If

    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "Column B me header (row 1) ke ilawa koi data row nahi mili.", _
               vbExclamation, "No Data"
        Exit Sub
    End If

    screenWasOn = Application.ScreenUpdating
    On Error GoTo CleanFail
    Application.ScreenUpdating = False

    ' ---------- 3) purana highlight hatayein ----------
    ws.Range("A2:J" & lastRow).Interior.Pattern = xlNone

    ' ---------- 4) purani Duplicates sheet delete ----------
    Application.DisplayAlerts = False
    On Error Resume Next
    wb.Worksheets(DUP_SHEET).Delete
    On Error GoTo CleanFail
    Application.DisplayAlerts = True

    ' ---------- 5) nayi Duplicates sheet + 6) headers ----------
    Set dupWs = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    dupWs.Name = DUP_SHEET
    dupWs.Range("A1:F1").Value = Array("Name", "Age", "Department", "Salary", _
                                       "Duplicate Count", "Total Salary")
    outputRow = 2
    duplicateCount = 0
    nonNumeric = 0

    ' ==================================================
    ' FIRST PASS - har combination kitni baar aaya
    ' ==================================================
    Set dict = CreateObject("Scripting.Dictionary")
    For i = 2 To lastRow
        key = DupKey(ws, i)
        If dict.Exists(key) Then
            dict(key) = dict(key) + 1
        Else
            dict.Add key, 1
        End If
    Next i

    ' ==================================================
    ' SECOND PASS - highlight + sirf EK summary row per combo
    ' ==================================================
    Set summaryDict = CreateObject("Scripting.Dictionary")
    For i = 2 To lastRow
        key = DupKey(ws, i)
        If dict(key) > 1 Then
            ' highlight (A se J tak, yellow)
            ws.Range("A" & i & ":J" & i).Interior.Color = RGB(255, 255, 0)

            If Not summaryDict.Exists(key) Then
                summaryDict.Add key, outputRow
                dupWs.Cells(outputRow, "A").Value = ws.Cells(i, "B").Value
                dupWs.Cells(outputRow, "B").Value = ws.Cells(i, "C").Value
                dupWs.Cells(outputRow, "C").Value = ws.Cells(i, "D").Value
                dupWs.Cells(outputRow, "D").Value = ws.Cells(i, "E").Value
                dupWs.Cells(outputRow, "E").Value = dict(key)
                ' SALARY RULE: Total = Salary x Count (ek hi baar, accumulation nahi)
                If IsNumeric(ws.Cells(i, "E").Value) And Not IsEmpty(ws.Cells(i, "E").Value) Then
                    dupWs.Cells(outputRow, "F").Value = CDbl(ws.Cells(i, "E").Value) * dict(key)
                Else
                    dupWs.Cells(outputRow, "F").Value = "n/a"
                    nonNumeric = nonNumeric + dict(key)
                End If
                outputRow = outputRow + 1
            End If
            duplicateCount = duplicateCount + 1
        End If
    Next i

    ' ==================================================
    ' FORMAT - bold white-on-blue headers, AutoFit, #,##0
    ' ==================================================
    With dupWs.Range("A1:F1")
        .Font.Bold = True
        .Interior.Color = RGB(0, 112, 192)
        .Font.Color = RGB(255, 255, 255)
    End With
    dupWs.Columns("A:F").AutoFit
    dupWs.Columns("D").NumberFormat = "#,##0"   ' Salary
    dupWs.Columns("F").NumberFormat = "#,##0"   ' Total Salary  (E = count par NAHI)

    ' ==================================================
    ' SELF-VALIDATION (honest - fail ho to bataya jata hai)
    ' ==================================================
    Dim highlighted As Long, summaryRows As Long, badTotal As Long
    highlighted = 0
    For r = 2 To lastRow
        If ws.Cells(r, "A").Interior.Color = RGB(255, 255, 0) Then highlighted = highlighted + 1
    Next r
    summaryRows = outputRow - 2
    badTotal = 0
    For r = 2 To 1 + summaryRows
        If IsNumeric(dupWs.Cells(r, "D").Value) And IsNumeric(dupWs.Cells(r, "E").Value) _
           And IsNumeric(dupWs.Cells(r, "F").Value) Then
            If Abs(CDbl(dupWs.Cells(r, "D").Value) * CDbl(dupWs.Cells(r, "E").Value) _
                   - CDbl(dupWs.Cells(r, "F").Value)) > 0.005 Then badTotal = badTotal + 1
        End If
    Next r

    Dim vNote As String
    If highlighted <> duplicateCount Then
        vNote = vNote & vbCrLf & "WARNING: highlighted rows (" & highlighted & _
                ") count (" & duplicateCount & ") se different hain - please review karein."
    End If
    If summaryRows <> summaryDict.Count Then
        vNote = vNote & vbCrLf & "WARNING: summary rows (" & summaryRows & _
                ") unique combos (" & summaryDict.Count & ") se different hain."
    End If
    If badTotal > 0 Then
        vNote = vNote & vbCrLf & "WARNING: " & badTotal & " summary row(s) ka Total Salary = Salary x Count nahi hai."
    End If
    If nonNumeric > 0 Then
        vNote = vNote & vbCrLf & "NOTE: " & nonNumeric & " row(s) ki Salary numeric nahi thi - Total Salary 'n/a' hai."
    End If

    Application.ScreenUpdating = screenWasOn

    ' ==================================================
    ' FINAL MESSAGE
    ' ==================================================
    If duplicateCount = 0 Then
        MsgBox "No duplicate records were found." & vNote, _
               vbInformation, "Highlight & Summarize Duplicates"
    Else
        MsgBox "Duplicate analysis completed." & vbCrLf & vbCrLf & _
               duplicateCount & " duplicate rows found." & vbCrLf & _
               summaryDict.Count & " unique duplicate entries created." & vbCrLf & vbCrLf & _
               "A summary has been created in the '" & DUP_SHEET & "' sheet." & vNote, _
               vbInformation, "Highlight & Summarize Duplicates"
    End If
    Exit Sub

CleanFail:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Macro poora complete NAHI hua: " & Err.Description & " (Error " & Err.Number & ")." & vbCrLf & _
           "Jo highlight ab tak hua hai wo hai; '" & DUP_SHEET & "' sheet adhoori ho sakti hai. " & _
           "Data ki values delete/modify nahi hui hain - sirf formatting ka asar ho sakta hai.", _
           vbCritical, "Highlight & Summarize Duplicates"
End Sub

'------------------------------------------------------------------------------
' Duplicate key: Trim(CStr()) ke baad Name|Age|Department|Salary (exact, case-sensitive)
'------------------------------------------------------------------------------
Private Function DupKey(ws As Worksheet, ByVal i As Long) As String
    DupKey = Trim$(CStr(ws.Cells(i, "B").Value)) & "|" & _
             Trim$(CStr(ws.Cells(i, "C").Value)) & "|" & _
             Trim$(CStr(ws.Cells(i, "D").Value)) & "|" & _
             Trim$(CStr(ws.Cells(i, "E").Value))
End Function
