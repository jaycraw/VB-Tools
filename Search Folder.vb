Sub fileCheck(keyword As String, folderstr As String, lo As ListObject)
    'Given a search keyword, folder, and table, add a row for each file containing the keyword containing file name,  date last modified, and a link to the file 
    Dim fso As Object
    Set fso = CreateObject("scripting.filesystemobject")
    
    'yell at user if folder doesn't exist
    If fso.FolderExists(folderstr) = False Then
        MsgBox folderstr & " doesn't exist"
    End If
    
    
    Dim lr As ListRow
    strfile = Dir(folderstr)
    
    'iterate through each file in the folder. If it contains the desired search term, add a row to the output table with file name, date modified, and a hyperlink
    Do While strfile <> ""
        If InStr(1, strfile, keyword) Then
            Set lr = lo.ListRows.Add
            lr.Range(1).Value = strfile
            lr.Range(2).Value = fso.GetFile(folderstr & strfile).DateLastModified
            lr.Range(3).Value = "Link"
            ThisWorkbook.Worksheets("File Check").Hyperlinks.Add lr.Range(3), folderstr & strfile
            End If
        strfile = Dir()
    Loop
    
End Sub