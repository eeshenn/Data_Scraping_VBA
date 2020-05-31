Attribute VB_Name = "Updating_data"
Sub update()

Dim ie As InternetExplorer
Dim htmldoc As HTMLDocument
Dim htmltable As HTMLDTElement
Dim htmlrows As Variant


Set ie = New InternetExplorer
ie.Visible = False
ie.navigate "https://data.gov.sg/dataset/bunker-sales-monthly/resource/44da3191-6c57-4d4a-8268-8e2c418d4b43/view/d8d07ae9-ab58-4ae0-84f7-7bf6e6ee93f9"

Dim t As Date, ele As Object
    Const MAX_WAIT_SEC As Long = 1 '<==Adjust wait time

    While ie.Busy Or ie.readyState < 4: DoEvents: Wend
    t = Timer
    Do
        DoEvents
    On Error Resume Next
    Set ele = ie.document.getElementById("firstname")
    If Timer - t > MAX_WAIT_SEC Then Exit Do
    On Error GoTo 0
    Loop While ele Is Nothing

    If Not ele Is Nothing Then
    'do something
    End If


Set htmldoc = ie.document
Set htmltable = htmldoc.getElementById("resource_table")
Set hbody = htmltable.getElementsByTagName("tbody")
Set htmlrows = hbody(0).getElementsByTagName("tr")

Website_Latest_Date = CDate(htmlrows(0).Children(0).innerText)
Excel_Latest_Date = CDate(Sheets(2).Cells(Rows.Count, 1).End(xlUp).Value)


'Check if lastest date of data on excel is the same as the one on website
If Website_Latest_Date = Excel_Latest_Date Then
    MsgBox "Data is already updated"

'If not the same, then transfer the latest data over to excel from website
ElseIf Website_Latest_Date <> Excel_Latest_Date Then
    input_date = Format(Website_Latest_Date, "mmm-yyyy")
    Set cell = Sheets(2).Cells(Rows.Count, 1).End(xlUp).Offset(1, 0)
    cell.Value = input_date
    For i = 0 To 11
        cell.Offset(0, i + 1).Value = htmlrows(i).Children(2).innerText
    Next
    
End If



End Sub
