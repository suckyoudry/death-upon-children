Imports Excel = Microsoft.Office.Interop.Excel
Public Class Form1

    Dim IndexDates(4, 100) As String
    Dim IndexOpen(4, 100) As Decimal
    Dim IndexHigh(4, 100) As Decimal
    Dim IndexLow(4, 100) As Decimal
    Dim IndexClosingPrice(4, 100) As Decimal
    Dim IndexAdjClose(4, 100) As Decimal
    Dim IndexVolume(4, 100) As Decimal
    Dim sharename(10) As String
    Dim Sharedates(10, 100) As String
    Dim ClosingPrice(10, 100) As Decimal
    Dim Highprice(10, 100) As Decimal
    Dim Lowprice(10, 100) As Decimal
    Dim Openprice(10, 100) As Decimal
    Dim xpos As Integer
    Dim temparea As String

    'the excel commands for the excel applications & workbooks
    Dim xlApp As Excel.Application
    Dim xlWorkBook As Excel.Workbook
    Dim xlWorkSheet As Excel.Worksheet



    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load

        'Open Worksheet for each index then read into data arrays each excel cell required for each index
        'selecting the correct file & worksheet
        xlApp = New Excel.Application
        xlWorkBook = xlApp.Workbooks.Open("Z:\Fullbrook A-Levels\Computer Science\computing project\Index Closing Prices.xlsx")

        For n = 1 To 4
            xlWorkSheet = xlWorkBook.Worksheets("Index" + CStr(n))

            If n < 2 Then ' this is making the first index which is formatted differently work in the same loop
                For i = 2 To 100
                    IndexDates(n, i - 1) = (xlWorkSheet.Cells(i, 1).value)
                    temparea = (xlWorkSheet.Cells(i, 3).value) 'i had to use a temp area because the code didn't work trying to search for "-" as a block of code
                    If temparea = "-" Then
                        IndexOpen(n, i - 1) = IndexOpen(n, i - 2)
                    Else
                        IndexOpen(n, i - 1) = Val(temparea)
                    End If
                    temparea = (xlWorkSheet.Cells(i, 5).value)
                    If temparea = "-" Then
                        IndexLow(n, i - 1) = IndexLow(n, i - 2)
                    Else
                        IndexLow(n, i - 1) = Val(temparea)
                    End If
                    temparea = (xlWorkSheet.Cells(i, 4).value)
                    If temparea = "-" Then
                        IndexHigh(n, i - 1) = IndexHigh(n, i - 2)
                    Else
                        IndexHigh(n, i - 1) = Val(temparea)
                    End If
                    temparea = (xlWorkSheet.Cells(i, 2).value)
                    If temparea = "-" Then
                        IndexClosingPrice(n, i - 1) = IndexClosingPrice(n, i - 2)
                    Else
                        IndexClosingPrice(n, i - 1) = Val(temparea)
                    End If
                Next

            Else
                For i = 2 To 100
                    IndexDates(n, i - 1) = (xlWorkSheet.Cells(i, 1).value)
                    temparea = (xlWorkSheet.Cells(i, 2).value)
                    If temparea = "-" Then
                        IndexOpen(n, i - 1) = IndexOpen(n, i - 2)
                    Else
                        IndexOpen(n, i - 1) = Val(temparea)
                    End If
                    temparea = (xlWorkSheet.Cells(i, 3).value)
                    If temparea = "-" Then
                        IndexLow(n, i - 1) = IndexLow(n, i - 2)
                    Else
                        IndexLow(n, i - 1) = Val(temparea)
                    End If
                    temparea = (xlWorkSheet.Cells(i, 4).value)
                    If temparea = "-" Then
                        IndexHigh(n, i - 1) = IndexHigh(n, i - 2)
                    Else
                        IndexHigh(n, i - 1) = Val(temparea)
                    End If
                    temparea = (xlWorkSheet.Cells(i, 5).value)
                    If temparea = "-" Then
                        IndexClosingPrice(n, i - 1) = IndexClosingPrice(n, i - 2)
                    Else
                        IndexClosingPrice(n, i - 1) = Val(temparea)
                    End If
                Next
            End If
        Next

        'Open worksheet for each share and then read into data arrays each excel cell required for each share

        xlApp = New Excel.Application
        xlWorkBook = xlApp.Workbooks.Open("Z:\Fullbrook A-Levels\Computer Science\graphing\database for markets.xlsx")

        For n = 1 To 10
            xlWorkSheet = xlWorkBook.Worksheets("Share" + CStr(n))
            For i = 2 To 101
                Sharedates(n, i - 1) = (xlWorkSheet.Cells(i, 1).value)
                temparea = (xlWorkSheet.Cells(i, 2).value)
                If temparea = "-" Then
                    Openprice(n, i - 1) = Openprice(n, i - 2)
                Else
                    Openprice(n, i - 1) = Val(temparea)
                End If
                temparea = (xlWorkSheet.Cells(i, 3).value)
                If temparea = "-" Then
                    Lowprice(n, i - 1) = Lowprice(n, i - 2)
                Else
                    Lowprice(n, i - 1) = Val(temparea)
                End If
                temparea = (xlWorkSheet.Cells(i, 4).value)
                If temparea = "-" Then
                    Highprice(n, i - 1) = Highprice(n, i - 2)
                Else
                    Highprice(n, i - 1) = Val(temparea)
                End If
                temparea = (xlWorkSheet.Cells(i, 5).value)
                If temparea = "-" Then
                    ClosingPrice(n, i - 1) = ClosingPrice(n, i - 2)
                Else
                    ClosingPrice(n, i - 1) = Val(temparea)
                End If
            Next
        Next

        For i = 1 To 10
            xlWorkSheet = xlWorkBook.Worksheets("Sharename")
            sharename(i) = (xlWorkSheet.Cells(i, 1).value)
        Next


        xlWorkBook.Close()
        xlApp.Quit()

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        xpos = 1

        For i = 99 To 1 Step -1
            Me.Chart1.Series("Closing Price").Points.AddXY(xpos, ClosingPrice(1, i))
            xpos = xpos + 1
        Next

        MsgBox(sharename(4))

    End Sub
End Class
