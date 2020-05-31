Attribute VB_Name = "Scraping_historical_data"
Sub scrape()
Dim ie As InternetExplorer
Dim htmldoc As HTMLDocument
Dim htmltable As HTMLDTElement
Dim htmlrows As Variant
Dim htmlrow As HTMLDTElement


Set ie = New InternetExplorer

ie.Visible = False
ie.navigate "https://data.gov.sg/dataset/bunker-sales-monthly/resource/44da3191-6c57-4d4a-8268-8e2c418d4b43/view/d8d07ae9-ab58-4ae0-84f7-7bf6e6ee93f9"

x = 1

For i = 0 To 26 'Adjust accordingly to the number of pages to go back

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
    Set htmloption = htmldoc.getElementById("resource_table_next")
    Set htmloptionlink = htmloption.getElementsByTagName("a")(0) 'element of the button to go to the next page
    Set htmltable = htmldoc.getElementById("resource_table")
    Set hbody = htmltable.getElementsByTagName("tbody")
    Set htmlrows = hbody(0).getElementsByTagName("tr")

    'transfer of raw data from website to excel
    For Each htmlrow In htmlrows
        Sheets(1).Cells(x, 1) = htmlrow.Children(0).innerText
        Sheets(1).Cells(x, 2) = htmlrow.Children(2).innerText
        x = x + 1
    Next

    htmloptionlink.Click 'function to click on the button to go to the next page
Next

transform 'After raw data is transfered over, transform the raw data to structured table form

End Sub

Sub transform()
Dim wks1 As Worksheet
Set wks1 = Sheets(1)
Dim wks2 As Worksheet
Set wks2 = Sheets(2)
Dim cell As Range
Dim cell2 As Range
Dim cellcheck As String

Set cell = wks1.Range("A1")

Do While cell.Value <> "2018-12" 'Adjust accordingly to the date before the latest date you want to transfer over
    cellcheck = cell.Value
    wks2.Range("A5") = Format(cell.Value, "mmm-yyyy")
    Set cell2 = cell.Offset(0, 1)
    For i = 0 To 11
        wks2.Range("B5").Offset(0, i) = cell2.Offset(i, 0)
    Next
    Set cell = cell.Offset(12, 0)
    If cell.Value <> cellcheck And cell.Value <> "2018-12" Then
        wks2.Rows(5).Insert shift:=xlShiftDown
    End If
Loop




End Sub
