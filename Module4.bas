Attribute VB_Name = "Module4"
'Copy data from the Export File and the Entered Information into the corresponding files
Public Sub Copy_Information_CWI()

Dim fldr As FileDialog
Dim strFilePath As String 'Path of file to be selected
Dim strFileName As String 'Name of file selected: Must be the roster export
Dim strTargetPath As String 'Working directory where all 5 folders are
Dim strInstructor As String 'Instructor
Dim strStartDate As String 'Starting date of seminar
Dim strEndDate As String 'Ending date of seminar
Dim strDate As String 'Combination of both dates
Dim strSiteCode As String 'Site code.
Dim strState As String 'State
Dim strFacility As String 'Facility
Dim strAddress As String 'Address
Dim strCity As String 'City
Dim strZip As String 'Zip code
Dim strPhone As String 'Phone number
Dim intAttendees As Integer 'Number of attendees


strPhone = Range("C10") 'Get phone number from user input
strZip = Range("C9") 'Get zip from user input
strCity = Range("C7") 'Get city from user input
strAddress = Range("C6") 'Get address from user input
strFacility = Range("C5") 'Get facility from user input
strState = Range("C8") 'Get state from user input
strSiteCode = Range("C3") 'Get site code from user input


'Obtaining the Date string
strStartDate = Range("C11")
strEndDate = Range("C12")
    'Check if months are equal; if not, then use both month names for Date string
    If MonthName(Month(strStartDate)) = MonthName(Month(strEndDate)) Then
        strDate = MonthName(Month(strStartDate)) & " " & Day(strStartDate) & "-" & Day(strEndDate) & ", " & Year(strStartDate)
        Else
        strDate = MonthName(Month(strStartDate)) & " " & Day(strStartDate) & "-" & MonthName(Month(strEndDate)) & " " & Day(strEndDate) & ", " & Year(strStartDate)
        End If



strInstructor = Range("C4") 'Get instructor name from user entered info
strTargetPath = Range("C14") 'Get target folder from user entered info



    'User selects the Export file, if no selection is made then exit procedure
    Set fldr = Application.FileDialog(msoFileDialogFilePicker)
    With fldr
        .AllowMultiSelect = False
        .Title = "Select Seminar Export File"
    End With
    
    If fldr.Show = -1 Then
        strFilePath = fldr.SelectedItems(1)
    Else
        MsgBox "No file has been selected"
        Exit Sub
    End If

'Copy names from export

    
    'Delete rows with "Online" type of seminar
    Dim rngFee As Range
    
    strFileName = Right(strFilePath, Len(strFilePath) - InStrRev(strFilePath, "\"))
    
    Workbooks.Open strFilePath
    Set rngFee = Workbooks(strFileName).Worksheets(1).Range("J6", Range("J6").End(xlDown))
     
    For i = rngFee.Cells.Count To 1 Step -1
        If InStr(LCase(rngFee.Item(i).Value), LCase("online")) > 0 Then
            rngFee.Item(i).EntireRow.Delete
        End If
    Next i
    Workbooks(strFileName).Worksheets(1).Range("E6", Range("E6").End(xlDown)).copy
    intAttendees = Workbooks(strFileName).Worksheets(1).Range("E6", Range("E6").End(xlDown)).Count

    
'Mail merge file
    
    'Paste names into mail merge
    Workbooks.Open (strTargetPath & "\3. Pre-Course\" & "Sample Mail Merge List.xlsx")
    Workbooks("Sample Mail Merge List.xlsx").Worksheets(1).Range("A2").Select
    Workbooks("Sample Mail Merge List.xlsx").Worksheets(1).Paste
    
    'ActiveSheet.Paste
    
    'Add instrutor name to mail merge
    Dim intLastRow As Integer
    intLastRow = (Cells(Rows.Count, 1).End(xlUp).Row)
    Range(Cells(2, 2), Cells(intLastRow, 2)).Value = strInstructor
    
    'Add date to mail merge
    Range(Cells(2, 3), Cells(intLastRow, 3)).Value = strDate
    
    Workbooks("Sample Mail Merge List.xlsx").Close SaveChanges:=True
    

    
'Roster file

    'Copy roster sheet to roster file
    
    Workbooks.Open (strTargetPath & "\3. Pre-Course\" & "Seminar Roster.xlsx")
    Application.DisplayAlerts = False
    
        For Each Sheet In Workbooks("Seminar Roster.xlsx").Sheets
            If LCase(Sheet.Name) = LCase("Seminar roster-" & strSiteCode) Then
                Workbooks("Seminar Roster.xlsx").Worksheets("Seminar roster-" & strSiteCode).Delete
                
            End If
        Next Sheet
        
    Workbooks(strFileName).Worksheets(1).copy Before:=Workbooks("Seminar Roster.xlsx").Sheets(1)
    Workbooks("Seminar Roster.xlsx").Sheets(1).Name = "Seminar roster-" & strSiteCode
    
    
    'Copy names in Sign in Sheet
    Workbooks("Seminar Roster.xlsx").Worksheets(1).Range("E6", Range("E6").End(xlDown)).copy
    Workbooks("Seminar Roster.xlsx").Worksheets("Sign-In Sheet").Range("B3").PasteSpecial Paste:=xlPasteValues
    
    'Copy course in Sign in sheet
    Workbooks("Seminar Roster.xlsx").Worksheets(1).Range("J6", Range("J6").End(xlDown)).copy
    Workbooks("Seminar Roster.xlsx").Worksheets("Sign-In Sheet").Range("C3").PasteSpecial Paste:=xlPasteValues
    
    'Change header name
    Workbooks("Seminar Roster.xlsx").Worksheets("Sign-In Sheet").Range("B1").Value = "Sign-in Sheet- " & strState & ", " & strSiteCode & " - " & strDate
    
    'Fit rows and columns
    Workbooks("Seminar Roster.xlsx").Worksheets("Sign-In Sheet").Cells.EntireColumn.AutoFit
    Workbooks("Seminar Roster.xlsx").Worksheets("Sign-In Sheet").Cells.EntireRow.AutoFit

    'Copy names in Book Status
    Workbooks("Seminar Roster.xlsx").Worksheets(1).Range("E6", Range("E6").End(xlDown)).copy
    Workbooks("Seminar Roster.xlsx").Worksheets("Book Status").Range("A3").PasteSpecial Paste:=xlPasteValues
    
    'Copy Address in Book Status
    Workbooks("Seminar Roster.xlsx").Worksheets(1).Range("I6", Range("I6").End(xlDown)).copy
    Workbooks("Seminar Roster.xlsx").Worksheets("Book Status").Range("B3").PasteSpecial Paste:=xlPasteValues
    
    'Change header name
    Workbooks("Seminar Roster.xlsx").Worksheets("Book Status").Range("A1").Value = "Books in Advance Status " & strSiteCode
    
    'Fit rows and columns
    Workbooks("Seminar Roster.xlsx").Worksheets("Book Status").Cells.EntireColumn.AutoFit
    Workbooks("Seminar Roster.xlsx").Worksheets("Book Status").Cells.EntireRow.AutoFit
    
    Workbooks("Seminar Roster.xlsx").Close SaveChanges:=True
    

    
'TSS file

    Workbooks.Open (strTargetPath & "\3. Pre-Course\" & "CWI Packing List - TSS.xlsx")
    Workbooks("CWI Packing List - TSS.xlsx").Worksheets(1).Range("A5") = strSiteCode
    Workbooks("CWI Packing List - TSS.xlsx").Worksheets(1).Range("B5") = strCity & ", " & strState
    Workbooks("CWI Packing List - TSS.xlsx").Worksheets(1).Range("D5") = strStartDate
    Workbooks("CWI Packing List - TSS.xlsx").Worksheets(1).Range("E5") = strEndDate
    Workbooks("CWI Packing List - TSS.xlsx").Worksheets(1).Range("F5") = strInstructor
    Workbooks("CWI Packing List - TSS.xlsx").Worksheets(1).Range("H5") = "TO: " & strInstructor & ", Arriving guest" & vbNewLine & _
                                                                                    strFacility & vbNewLine & strAddress & vbNewLine _
                                                                                    & strCity & ", " & strState & " " & strZip _
                                                                                    & vbNewLine & strPhone
    
    Workbooks("CWI Packing List - TSS.xlsx").Close SaveChanges:=True
    

    
'AWS file
                                                                                    
    Workbooks.Open (strTargetPath & "\3. Pre-Course\" & "CWI Packing List - AWS.xlsx")
    Workbooks("CWI Packing List - AWS.xlsx").Worksheets(1).Range("A4") = strSiteCode
    Workbooks("CWI Packing List - AWS.xlsx").Worksheets(1).Range("B4") = strCity & ", " & strState
    Workbooks("CWI Packing List - AWS.xlsx").Worksheets(1).Range("D4") = strStartDate
    Workbooks("CWI Packing List - AWS.xlsx").Worksheets(1).Range("E4") = strEndDate
    Workbooks("CWI Packing List - AWS.xlsx").Worksheets(1).Range("F4") = strInstructor
    Workbooks("CWI Packing List - AWS.xlsx").Worksheets(1).Range("H4") = "TO: " & strInstructor & ", Arriving guest" & vbNewLine & _
                                                                                    strFacility & vbNewLine & strAddress & vbNewLine _
                                                                                    & strCity & ", " & strState & " " & strZip _
                                                                                    & vbNewLine & strPhone
    
    Workbooks("CWI Packing List - AWS.xlsx").Close SaveChanges:=True
    

    
'Book return form

    Workbooks.Open (strTargetPath & "\3. Pre-Course\" & "CWI Book Return Form.xlsx")
    Workbooks("CWI Book Return Form.xlsx").Worksheets(1).Range("A5") = strSiteCode
    Workbooks("CWI Book Return Form.xlsx").Worksheets(1).Range("B5") = strCity & ", " & strState
    Workbooks("CWI Book Return Form.xlsx").Worksheets(1).Range("D5") = strStartDate
    Workbooks("CWI Book Return Form.xlsx").Worksheets(1).Range("E5") = strEndDate
    Workbooks("CWI Book Return Form.xlsx").Worksheets(1).Range("F5") = strInstructor
    Workbooks("CWI Book Return Form.xlsx").Worksheets(1).Range("H5") = "TO: " & strInstructor & ", Arriving guest" & vbNewLine & _
                                                                                    strFacility & vbNewLine & strAddress & vbNewLine _
                                                                                    & strCity & ", " & strState & " " & strZip _
                                                                                    & vbNewLine & strPhone
   
   Workbooks("CWI Book Return Form.xlsx").Close SaveChanges:=True
   

    
'Facility evaluation form

    Workbooks.Open (strTargetPath & "\3. Pre-Course\" & "Facility Evaluations.xlsx")
    Workbooks("Facility Evaluations.xlsx").Worksheets(1).Range("D3") = strFacility
    Workbooks("Facility Evaluations.xlsx").Worksheets(1).Range("J10") = strDate
    Workbooks("Facility Evaluations.xlsx").Worksheets(1).Range("J12") = strSiteCode
    Workbooks("Facility Evaluations.xlsx").Worksheets(1).Range("D5") = strCity & ", " & strState
    Workbooks("Facility Evaluations.xlsx").Worksheets(1).Range("D12") = strCity & ", " & strState
    Workbooks("Facility Evaluations.xlsx").Worksheets(1).Range("D10") = strInstructor
    
    Workbooks("Facility Evaluations.xlsx").Close SaveChanges:=True
       
       
'Shipping confirmation file


    Workbooks.Open (strTargetPath & "\3. Pre-Course\" & "Shipping Confirmation.xlsx")
    
    'Fill Location, Date, and Site Code
    Workbooks("Shipping Confirmation.xlsx").Worksheets(1).Range("B4") = strCity & ", " & strState
    Workbooks("Shipping Confirmation.xlsx").Worksheets(1).Range("B5") = strDate
    Workbooks("Shipping Confirmation.xlsx").Worksheets(1).Range("B6") = strSiteCode
    
    'Fill address information
    Workbooks("Shipping Confirmation.xlsx").Worksheets(1).Range("A11") = "TO: " & strInstructor & ", Arriving guest" & vbNewLine & _
                                                                                        strFacility & vbNewLine & strAddress & vbNewLine _
                                                                                        & strCity & ", " & strState & " " & strZip _
                                                                                        & vbNewLine & strPhone
                                                                                        
     
    If intAttendees >= 1 And _
        intAttendees <= 10 Then
            Workbooks("Shipping Confirmation.xlsx").Worksheets(1).Range("A17") = "Materials & 1 Cases"
    ElseIf intAttendees > 10 And _
            intAttendees <= 22 Then
                Workbooks("Shipping Confirmation.xlsx").Worksheets(1).Range("A17") = "Materials & 2 Cases"
    ElseIf intAttendees > 22 And _
            intAttendees <= 34 Then
                Workbooks("Shipping Confirmation.xlsx").Worksheets(1).Range("A17") = "Materials & 3 Cases"
    ElseIf intAttendees > 34 And _
            intAttendees <= 46 Then
                Workbooks("Shipping Confirmation.xlsx").Worksheets(1).Range("A17") = "Materials & 4 Cases"
    ElseIf intAttendees > 46 And _
            intAttendees <= 58 Then
                Workbooks("Shipping Confirmation.xlsx").Worksheets(1).Range("A17") = "Materials & 5 Cases"
                
    
    
    End If
    
    Workbooks("Shipping Confirmation.xlsx").Close SaveChanges:=True
       
    
'Close export file
    Workbooks(strFileName).Close SaveChanges:=True
    
MsgBox "All information has been copied without problems!"
    
End Sub
