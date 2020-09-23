Attribute VB_Name = "FileListTasks"
Public Sub LoadListFromFile(xPath As String, LListBox As ListBox)
    'This Sub will load any list of information
    'from your hard drive into a ListBox control.
    'xPath represents the file to be loaded.
    Dim TheFile As Integer
    Dim OurBuffer As String
    If Dir(xPath) = vbNullString Then Exit Sub 'MsgBox "File Not Found!", 16, "Error: File Not Found": Exit Sub
    TheFile = FreeFile() 'This will assure that the
    'File number is a good one
    
    Open xPath For Input As #TheFile
    Do While Not EOF(TheFile)
        Line Input #TheFile, OurBuffer
        'If the line being read is empty, do not add it
        If OurBuffer = vbNullString Then GoTo SkipAdd
        LListBox.AddItem OurBuffer
        SkipAdd: 'If the item is empty, instructions jump here
        DoEvents 'Allow other proccesses to run
    Loop
    Close #TheFile
End Sub
Public Sub SaveList(xPath As String, LListBox As ListBox)
    'This Sub saves the contents of a ListBox control to a
    'file on the hard drive
    Dim TheFile, ListCount As Integer
    TheFile = FreeFile() 'This assures that our file number
    'is not already in use by another file
    
    Open xPath For Output As #TheFile
    For ListCount = 0 To LListBox.ListCount - 1
        Print #TheFile, LListBox.List(ListCount)
        DoEvents
    Next ListCount
    Close #TheFile
End Sub
