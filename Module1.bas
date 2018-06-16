Attribute VB_Name = "Module1"
Sub Button1_Click()
Dim i As Long
Dim mySheet As Worksheet

For Each xWs In Application.ActiveWorkbook.Worksheets
    If xWs.Name <> "Data Sources" Then
        xWs.Delete 'Delete all sheets other than our Data Sources Sheet
    End If
Next


Set mySheet = ActiveSheet

For i = 2 To mySheet.Rows.Count
If Not IsEmpty(mySheet.Cells(i, 1).Value) Then
    Dim urlSource As String
    Dim sheetName As String
    urlSource = mySheet.Cells(i, 1).Value
    sheetName = mySheet.Cells(i, 2).Value
    Call LoadURL(urlSource, sheetName)
End If
Next i
End Sub

Sub LoadURL(ByVal urlSource As String, ByVal sheetName As String)
Dim ws As Worksheet
With ThisWorkbook
    Set ws = .Sheets.Add(After:=.Sheets(.Sheets.Count))
    ws.Name = sheetName
    
    Dim xhr As MSXML2.ServerXMLHTTP
    Dim doc As MSHTML.HTMLDocument
    Dim tables As MSHTML.IHTMLElementCollection
    Dim tableRows As MSHTML.IHTMLElementCollection
    Dim tableCells As MSHTML.IHTMLElementCollection
    Dim headerRows As MSHTML.IHTMLElementCollection
    Dim headerRow As MSHTML.HTMLTableRow
    Dim headerCells As MSHTML.IHTMLElementCollection
    Dim headerCell As MSHTML.HTMLTableCell
    Dim curTable As MSHTML.HTMLTable
    Dim curRow As MSHTML.HTMLTableRow
    Dim curCell As MSHTML.HTMLTableCell
    Dim headerValue As Variant
    Dim cellValue As Variant
    Dim tableCount As Integer
    Dim rowCount As Integer
    Dim cellCount As Integer
        
    tableCount = 0
    
    Set xhr = CreateObject("MSXML2.ServerXMLHttp")
    xhr.Open "GET", urlSource, False
    xhr.send
    If xhr.readyState = 4 And xhr.Status = 200 Then
        Set doc = New MSHTML.HTMLDocument
        doc.body.innerHTML = xhr.responseText
    Else
        MsgBox "Error" & vbNewLine & "Ready state: " & xhr.readyState & _
        vbNewLine & "HTTP request status: " & xhr.Status
    End If
    ws.Cells(1, 1).Value = "League"
    ws.Cells(1, 2).Value = "Owner"
    ws.Cells(1, 3).Value = "Player"
    Set tables = doc.getElementsByTagName("table")
    For Each curTable In tables
        If curTable.Width = 210 Then 'get the table header from here
        Set headerRows = curTable.getElementsByTagName("tr")
        For Each headerRow In headerRows
            Set headerCells = headerRow.getElementsByTagName("td")
            For Each headerCell In headerCells
                Dim valignStyle As String
                valignStyle = headerCell.getAttribute("valign")
                If valignStyle = "top" Then
                    headerValue = headerCell.innerHTML
                    Dim charPos As Integer
                    charPos = InStr(headerValue, "<TABLE")
                    headerValue = Left(headerValue, charPos - 1)
                    headerValue = GetText(headerValue)
                End If
                
                If headerValue <> "" Then GoTo ParseChildTable
            Next
        Next
        End If
    
ParseChildTable:
        'There is no id on the table but the important ones appear to have a width of 240
        If curTable.Width = 240 Then
            rowCount = 1
            cellCount = 1
            Set tableRows = curTable.getElementsByTagName("tr")
            ws.Cells((tableCount * 39) + (rowCount + 1), cellCount + 2).Value = headerValue
            For Each curRow In tableRows
                cellCount = 1
                Set tableCells = curRow.getElementsByTagName("td")
                For Each curCell In tableCells
                    cellValue = curCell.innerText
                    ws.Cells((tableCount * 39) + (rowCount + 2), cellCount + 2).Value = cellValue
                    cellCount = cellCount + 1
                Next
                rowCount = rowCount + 1
            Next
            tableCount = tableCount + 1
            headerValue = ""
        End If

    Next
    Columns("C:C").Select
    Selection.Cut
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight
End With
End Sub


Public Function GetText(inputHtml As Variant)
Set odoc = CreateObject("htmlfile")
odoc.Open
odoc.write inputHtml
odoc.Close
GetText = odoc.body.innerText
End Function
