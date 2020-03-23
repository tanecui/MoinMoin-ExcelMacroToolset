Attribute VB_Name = "MoinModule"
Private Function ConvertNewLine(content As String) As String
    Dim resultContent As String
    'MsgBox content
    resultContent = Replace(content, vbLf, "<<BR>>")
    
    'MsgBox resultContent
    ConvertNewLine = resultContent
End Function

Sub MoinFormat()
'
' MoinFormat Macro
' The macro will convert the excel data to moin format
'

'
    Dim result, cLine As String
    Dim selectedRange, cRow As Range
    Dim cCell As Object
    Dim cCellText As String
    
    Set selectedRange = Selection
    result = ""
    Dim isTitle As Boolean
    isTitle = True
    
    For Each cRow In selectedRange.Rows
        
        cLine = "||"
        For Each cCell In cRow.Cells
            cCellText = ConvertNewLine(cCell.Text)
            If isTitle Then
                cLine = cLine & "'''" & cCellText & "'''||"
            Else
                cLine = cLine & cCellText & "||"
            End If
        Next cCell
        isTitle = False
        result = result & cLine & Chr(13)
    Next cRow
    SetClipboard (result)
    MsgBox "Moin Format has been set to clipboard"
    'result.PutInClipboard
End Sub
