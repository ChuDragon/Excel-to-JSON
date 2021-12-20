Attribute VB_Name = "Module1"
Option Base 1
Sub toJSON()
Attribute toJSON.VB_ProcData.VB_Invoke_Func = " \n14"
' Converts an Excel table to JSON
'
Dim FSO As New FileSystemObject, JsonTS As TextStream
Dim dict As Scripting.Dictionary, dict2 As Scripting.Dictionary

Dim row As Integer, col As Integer, col2 As Integer, keys As Variant, cell As Variant
Dim tblObj As ListObject, jsonItem As String

ActiveWorkbook.Sheets("Data").Activate
Set tblObj = getTableObject(ActiveSheet) 'function to obtain a table/listObject

' Copy keys from rows 1-2 of table into `keys` array
ReDim keys(1 To 2, 1 To tblObj.ListColumns.Count - 1)
For row = 1 To 2
    For col = 1 To tblObj.ListColumns.Count - 1
        keys(row, col) = tblObj.Range.Cells(row, col + 1)
    Next col
Next row

' Open a .json file for writing (home dir typically "C:\Users\<userName>\Documents")
If Not FSO.FolderExists(".\JSON output") Then FSO.CreateFolder (".\JSON output")
Set JsonTS = FSO.OpenTextFile(".\JSON output\timesheets.json", ForWriting, True)
JsonTS.Write ("[")
jsonItem = ""
' Excel table is defined with 1 header row, we have 2 keys, then data takes ListRows.Count-1 rows
For row = 1 To tblObj.ListRows.Count - 1

    ' Create a new level-1 object
    Set dict = New Scripting.Dictionary
    'Update `dict` object with row data for each key
    For col = 1 To tblObj.ListColumns.Count - 1
        
    If keys(2, col) = "" Then
        cell = tblObj.Range.Cells(row + 2, col + 1) 'rows 1-2 are headers
        dict(keys(1, col)) = cell 'Automatically adds key if it doesn't exist, in 1st loop step
    Else
    ' Create a new level-2 object
    Set dict2 = New Scripting.Dictionary
        
        For col2 = col To col + 1 'level-2 objects are always made of 2 elements
            cell = tblObj.Range.Cells(row + 2, col2 + 1)
            dict2(keys(2, col2)) = cell
        Next col2
        
        dict.Add keys(1, col), dict2
        col = col + 1 'now catch up col counter
    End If
    
    Next col
    
    'Convert to JSON using JsonConverter library, source: https://github.com/VBA-tools/VBA-JSON
    jsonItem = jsonItem + JsonConverter.ConvertToJson(dict) & IIf(row < tblObj.ListRows.Count - 1, ",", "")
Next row

'tblObj.Delete 'delte the table listObject
JsonTS.Write (jsonItem) ' write result to .json file
JsonTS.Write ("]")
JsonTS.Close
End Sub

Private Function getTableObject(ActiveSheet) As ListObject

With ActiveSheet
    If .ListObjects.Count = 0 Or .ListObjects.Count = Null Then
        ' if table listObject doesn't exist, create it from range
        .Range("A5").Select
        .Range(Selection, Selection.End(xlToRight)).Select
        .Range(Selection, Selection.End(xlDown)).Select
        Set getTableObject = .ListObjects.Add(xlSrcRange, Selection, , xlYes)
    Else
        Set getTableObject = .ListObjects(1)  'VBA uses 1-based list here
        If .ListObjects.Count > 1 Then
            MsgBox ("Warning: more than 1 table in " & ActiveSheet.Name & "sheet, using 1st")
        End If
    End If
End With

End Function
