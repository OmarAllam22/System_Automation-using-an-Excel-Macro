Sub Automated_OFAC_Commitment_Folders()


    Dim Rng As Range
    ' this is to get the data from what we select in the excel sheet.
    Set Rng = Selection
    Dim maxRows, maxCols, r, pointer As Integer
    maxRows = Rng.Rows.Count
    r = 9
    
    
    Do While r <= maxRows + 8
    
        '----> code to load our website called "OFAC"
        
        Dim IE As Object
        Dim doc As HTMLDocument
        Set IE = CreateObject("InternetExplorer.Application")
        IE.Visible = True
        IE.navigate "https://sanctionssearch.ofac.treas.gov/"
        
        
        
        '----> this code is to wait until the web_page is loaded
        
        Do While IE.Busy
            Application.Wait DateAdd("s", 1, Now)
        Loop
        
        
        
        '----> this code is to automate the data Entry on the website "OFAC" getting it from this excel sheet.
        
        Set doc = IE.document
        ' filling the name
        doc.getElementById("ctl00_MainContent_txtLastName").Value = ThisWorkbook.Sheets("Sheet1").Cells(r, 2).Value
        ' filling the ID
        doc.getElementById("ctl00_MainContent_txtID").Value = ThisWorkbook.Sheets("Sheet1").Cells(r, 4).Value
        ' filling the city
        doc.getElementById("ctl00_MainContent_txtCity").Value = "Egypt"
        
        
        
        '----> this code is to click the "search" button in the website and proceed to the rest of the code only if certian message appears.
              
        doc.getElementById("ctl00_MainContent_btnSearch").Click
        
        Do While IE.Busy
            Application.Wait DateAdd("s", 1, Now)
        Loop
    
        If InStr(doc.getElementById("ctl00_MainContent_lblMessage").innerText, "Your search has not returned any results.") = 0 Then
            MsgBox "Fail"   'the site won't be closed until you click "ok" on the MsgBox
            IE.Quit
        Else
                        
            ' here our needed message appears so we will:
            '   1) create folders for each name in our sheet.
            '   2) make a word file for each student and locate it in the folder of this student.
            '   3) print the "OFAC" webpage of this student and locate in the same folder of the student.


            '----> this code is to do the first step "create a Folder for each student"
            
                If Len(Dir(ActiveWorkbook.Path & "\" & Rng(r - 8, 1), vbDirectory)) = 0 Then
                    MkDir (ActiveWorkbook.Path & "\" & Rng(r - 8, 1))
        
            '----> this code is to do the first step "fill the student data in a word tempelate"
                    
                    Dim Wordapp As Word.Application
                    Set Wordapp = New Word.Application
                    Dim wdDOC As Word.document
                    Dim WdRng As Word.Range
                    Dim WdRng1 As Word.Range
                    Dim WdRng2 As Word.Range
                    Dim WdRng3 As Word.Range
                    Dim WdRng4 As Word.Range
                    Dim WdRng5 As Word.Range
                    Dim WdRng6 As Word.Range
                    Dim WdRng7 As Word.Range
                    
                ' here we open our word tempelate
                
                    Wordapp.Visible = True
                    Wordapp.Activate
                    Set wdDOC = Wordapp.Documents.Open(ActiveWorkbook.Path & "\new_commitment.docx")
            
                ' here we copy data from this excel sheet and paste it in their certain location in word tempelate (based on bookmarks we put in this word tempelate)
                
                    With Wordapp.Selection
                        Cells(1, 10).Copy
                    End With
                    Set WdRng4 = wdDOC.bookmarks("Round_Name").Range
                    WdRng4.PasteAndFormat Type:=wdFormatPlainText
                    
                    
                    With Wordapp.Selection
                        Cells(1, 11).Copy
                    End With
                    Set WdRng5 = wdDOC.bookmarks("Round_Number").Range
                    WdRng5.PasteAndFormat Type:=wdFormatPlainText
                    
                    
                    With Wordapp.Selection
                        Cells(1, 12).Copy
                    End With
                    Set WdRng6 = wdDOC.bookmarks("Start").Range
                    WdRng6.PasteAndFormat Type:=wdFormatPlainText
                    
                    
                    With Wordapp.Selection
                        Cells(1, 13).Copy
                    End With
                    Set WdRng7 = wdDOC.bookmarks("End").Range
                    WdRng7.PasteAndFormat Type:=wdFormatPlainText
                    
                   
                    With Wordapp.Selection
                        Cells(r, 2).Copy
                    End With
                    Set WdRng = wdDOC.bookmarks("Name").Range
                    WdRng.PasteAndFormat Type:=wdFormatPlainText
                        
                        
                    With Wordapp.Selection
                        Cells(r, 9).Copy
                    End With
                    Set WdRng1 = wdDOC.bookmarks("Faculty").Range
                    WdRng1.PasteAndFormat Type:=wdFormatPlainText
                    
                    
                    With Wordapp.Selection
                        Cells(r, 8).Copy
                    End With
                    Set WdRng2 = wdDOC.bookmarks("Year").Range
                    WdRng2.PasteAndFormat Type:=wdFormatPlainText
            
                    
                    With Wordapp.Selection
                        Cells(r, 2).Copy
                    End With
                    Set WdRng3 = wdDOC.bookmarks("Name2").Range
                    WdRng3.PasteAndFormat Type:=wdFormatPlainText
            
                ' here we save the word file in the folder of student whose data was filled right now
                
                    Wordapp.ActiveDocument.SaveAs2 ActiveWorkbook.Path & "\" & Rng(r - 8, 1) & "\" & ThisWorkbook.Sheets("Sheet1").Cells(r, 2) & ".docx"
                    Wordapp.ActiveDocument.Close
                    Wordapp.Quit
                
                    On Error Resume Next
                    End If
                
    
 
            '----> this code is to print the webpage of the "OFAC" and save it in the folder of student whose data was filled right now.
             
             
                ' this code to click the print button in the webpage.
            
                    doc.getElementById("ctl00_MainContent_ImageButton2").Click


                    
                ' this block of code is to proceed to the step after the print only if the file is printed and saved in its location
                ' so the following if statement is to check the existence of the printed file in the specified location
                                       
                                       
                    Dim fso As Object
                    Dim file_path As String
                    
                    file_path = ActiveWorkbook.Path & "\" & Rng(r - 8, 1) & "\OFAC.pdf"
                    Set fso = CreateObject("Scripting.FileSystemObject")
                    pointer = 0
                    
                    Do While pointer <> 1
                        If fso.FileExists(file_path) Then
                            pointer = 1
                            IE.Quit
                            GoTo Next_iteration:
                        Else
                            Application.Wait DateAdd("s", 1, Now)  ' this to wait if the page is not loaded and still busy.
                        End If
                    Loop
                              
                
Next_iteration:     ' this "Next_iteration" is a mark which has a reference in the if statement above.
        End If
        r = r + 1
    Loop
    MsgBox "Automation is Done"
    
End Sub