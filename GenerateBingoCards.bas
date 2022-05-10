Option Explicit

 Sub GenerateBingoCard()
 '
 '  Import this into a new blank Excel workbook or copy and
 '     paste into a new module.
 '
 '  On Sheet1, list all the items you wish to randomly generate
 '     within a bingo card in Column A starting at A1
 '
 '  Sheet1 must contain at least 24 items in Column A
 '  to generate a bingo card.
 '

 On Error Resume Next
 Dim i, k, x, y As Integer
 Dim j As Double
 Dim notenough As Boolean
 Dim BingoCard() As String
 ReDim BingoCard(1 To 24) As String

 With Application
     .ScreenUpdating = False
     .Interactive = False
     .EnableEvents = False
 End With

 notenough = False

 With ThisWorkbook

     ' Create Sheet2 if it doesn't already exist
     If Not sheetExists("Sheet2") Then
         Call gensheet2
     End If

     ' Randomly select items for Bingo Card
     With .Worksheets("Sheet1")
         If .Cells(sheet1.Rows.Count, "A").End(xlUp).Row < 24 Then
             notenough = True
             GoTo notenoughitems
         End If
         y = .Cells(sheet1.Rows.Count, "A").End(xlUp).Row - 1
         .Range("B:B").ClearContents
         For i = 1 To 24
             j = CInt(1 + Rnd * y)
             If .Range("B" & j) <> "USED" Then
                 .Range("B" & j) = "USED"
                 BingoCard(i) = CStr(.Range("A" & j).Value)
             Else
                 i = i - 1
             End If
         Next i
         .Range("B:B").ClearContents
     End With

     ' Insert the randomly selected items into bingo card
     x = 1
     With .Worksheets("Sheet2")
         .Range("A:Q").ClearContents
         For k = 1 To 5
             For i = 1 To 5
                 .Range("B" & i + 1).Offset(0, k - 1) = _
                     CStr(BingoCard(x))
                 x = x + 1
             Next i
         Next k
         .Range("F6") = .Range("D4")
         .Range("D4") = "FREE SPACE"
     End With
 End With
 notenoughitems:
     With Application
         .ScreenUpdating = True
         .Interactive = True
         .EnableEvents = True
     End With
     If notenough = True Then
         MsgBox "Please enter at least 24 items in Sheet1" & _
             "column A, starting at A1!", , "Not enough items!"
     End If
 End Sub

 Sub gensheet2()
 ' Creates Sheet2 and formats cells for a basic bingo card
 On Error Resume Next
 With Application
     .ScreenUpdating = False
     .Interactive = False
     .EnableEvents = False
 End With

 ThisWorkbook.Worksheets.Add.Name = "Sheet2"

 With ThisWorkbook.Worksheets("Sheet2")
     .Columns("B:F").ColumnWidth = 13.57
     .Rows("2:6").RowHeight = 75
     .Rows("1:1").RowHeight = 22.5
     .Columns("A:A").ColumnWidth = 3.57
     With .Range("B2:F6")
         .HorizontalAlignment = xlCenter
         .VerticalAlignment = xlCenter
         With .Borders
             .LineStyle = xlContinuous
             .Weight = xlMedium
         End With
         With .Borders(xlInsideVertical)
             .LineStyle = xlContinuous
             .Weight = xlThin
         End With
         With .Borders(xlInsideHorizontal)
             .LineStyle = xlContinuous
             .Weight = xlThin
         End With
     End With
 End With
 End Sub

 Function sheetExists(sheetToFind As  String, _
     Optional InWorkbook As Workbook) As Boolean

     If InWorkbook Is Nothing Then Set InWorkbook = ThisWorkbook
     On Error Resume Next
     sheetExists = Not InWorkbook.Sheets(sheetToFind) Is Nothing

 End Function
