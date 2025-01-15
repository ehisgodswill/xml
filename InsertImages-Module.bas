Attribute VB_Name = "Module1"
Sub InsertImages()
Dim ws As Worksheet
Dim imgFolder As String
Dim cell As Range
Dim lastRow As Long
Dim uuid As String

' Set the worksheet
Set ws = ThisWorkbook.Sheets("Sheet1") ' Change to your sheet name

' Set the folder path where images are stored
imgFolder = "C:\Users\USER\Downloads\83cfcef720ca41b9a4d884d79c766224\" ' Change to your folder path

' Get the last row in the column where filenames are listed
lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row ' Assuming filenames are in column A

' Loop through each cell
For Each cell In ws.Range("O2:O" & lastRow)
    uuid = cell.Value & "\"
    
    ' Check if the file exists
    If uuid <> "\" And Dir(imgFolder & uuid) <> "" Then
        ' Insert the image
        Set target = cell.Offset(0, -1)
        Set picture = Application.ActiveSheet.Shapes.AddPicture(imgFolder & uuid & target.Value, False, True, target.Left, target.Top, target.Width, 100)
        'target.Value = ""
        
        Set target2 = cell.Offset(0, -2)
        Set picture2 = Application.ActiveSheet.Shapes.AddPicture(imgFolder & uuid & target2.Value, False, True, target2.Left, target2.Top, target2.Width, 100)
        'target2.Value = ""
    End If
Next cell

End Sub
