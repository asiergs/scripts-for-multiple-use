Sub download_RFR_EIOPA()

Dim myURL, filename As String
Dim wb As Workbook

Set wb = Application.ThisWorkbook

file = InputBox("Selecciona fecha formato YYYYMMDD", "Input fecha", "20221231")
file = "eiopa_rfr_" & file

filename = file & ".zip"
myURL = "https://www.eiopa.europa.eu/sites/default/files/risk_free_interest_rate/" & filename

Dim WinHttpReq As Object
Set WinHttpReq = CreateObject("Microsoft.XMLHTTP")
WinHttpReq.Open "GET", myURL, False, "username", "password"
WinHttpReq.send

If WinHttpReq.Status = 200 Then
    Set oStream = CreateObject("ADODB.Stream")
    oStream.Open
    oStream.Type = 1
    oStream.Write WinHttpReq.responseBody
    oStream.SaveToFile ThisWorkbook.Path & "\" & filename, 2 ' 1 = no overwrite, 2 = overwrite
    oStream.Close
End If
    Dim Fname As Variant
    Fname = ThisWorkbook.Path & "\" & filename

    strFileName = Fname
    strFileExists = Dir(strFileName)
    If strFileExists = "" Then
        MsgBox ("No existe archivo, revisar fecha." & vbCrLf & _
        "El dia siempre debe ser el final de mes. Tener en cuenta que el primer archivo fue en enero de 2016")
        Exit Sub
    End If

    Dim FSO As Object
    Dim oApp As Object
    
    Dim FileNameFolder As Variant
    Dim DefPath As String
    Dim strDate As String
    
    If Fname = False Then
        'Do nothing
    Else
        'Root folder for the new folder.
        'You can also use DefPath = "C:\Users\Ron\test\"
        DefPath = ThisWorkbook.Path
        If Right(DefPath, 1) <> "\" Then
            DefPath = DefPath & "\"
        End If

        'Create the folder name
        FileNameFolder = DefPath & file & "\"
        FileNameFolder_del = DefPath & file

        'Make the normal folder in DefPath
        MkDir FileNameFolder

        'Extract the files into the newly created folder
        Set oApp = CreateObject("Shell.Application")
    
        oApp.Namespace(FileNameFolder).CopyHere oApp.Namespace(Fname).items

        'If you want to extract only one file you can use this:
        'oApp.Namespace(FileNameFolder).CopyHere _
        'oApp.Namespace(Fname).items.Item("test.txt")

        On Error Resume Next
        Set FSO = CreateObject("scripting.filesystemobject")
        FSO.DeleteFolder Environ("Temp") & "\Temporary Directory*", True
    End If

Dim objWorkbook As Workbook
Dim time_noVA_VA(1 To 150, 1 To 3) As Double
Dim i As Integer
Application.DisplayAlerts = False
Set objWorkbook = Workbooks.Open(FileNameFolder & "\" & file & "_term_structures.xlsx")
For i = 1 To 150
    time_noVA_VA(i, 1) = objWorkbook.Worksheets("RFR_spot_no_VA").Cells(i + 10, 2)
    time_noVA_VA(i, 2) = objWorkbook.Worksheets("RFR_spot_no_VA").Cells(i + 10, 3)
    time_noVA_VA(i, 3) = objWorkbook.Worksheets("RFR_spot_with_VA").Cells(i + 10, 3)
Next i
'close the workbook
objWorkbook.Close
Application.DisplayAlerts = True

'delete files
Dim DeleteFile As String
DeleteFile = Fname
If Len(Dir$(DeleteFile)) > 0 Then
SetAttr DeleteFile, vbNormal
Kill DeleteFile
End If
Set FSO = CreateObject("Scripting.FileSystemObject")
 
'Delete specified folder
FSO.DeleteFolder FileNameFolder_del

Selection = file
Selection.Offset(1, 0) = "Year"
Selection.Offset(1, 1) = "No VA"
Selection.Offset(1, 2) = "VA"

For i = 1 To 150
    Selection.Offset(i + 1, 0) = time_noVA_VA(i, 1)
    Selection.Offset(i + 1, 1) = time_noVA_VA(i, 2)
    Selection.Offset(i + 1, 2) = time_noVA_VA(i, 3)
Next i

End Sub


