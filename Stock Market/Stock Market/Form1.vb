Imports Excel = Microsoft.Office.Interop.Excel
Public Class Form1


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
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
        'opening the file and correct worksheet
        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet

        'since FTSE Index excel download was formatted differently have to seperate it from the rest of the Indexes

        xlApp = New Excel.Application
        xlWorkBook = xlApp.Workbooks.Open("Z:\Fullbrook A-Levels\Computer Science\computing project\Index Closing Prices.xlsx")

        For n = 1 To 1
            xlWorkSheet = xlWorkBook.Worksheets("Index" + CStr(n))
            For i = 2 To 100
                IndexDates(n, i - 1) = (xlWorkSheet.Cells(i, 1).value)
                temparea = (xlWorkSheet.Cells(i, 3).value)
                If temparea = "-" Then
                    IndexOpen(n, i - 1) = Openprice(n, i - 2)
                Else
                    IndexOpen(n, i - 1) = Val(temparea)
                End If
                temparea = (xlWorkSheet.Cells(i, 5).value)
                If temparea = "-" Then
                    IndexLow(n, i - 1) = Openprice(n, i - 2)
                Else
                    IndexLow(n, i - 1) = Val(temparea)
                End If
                temparea = (xlWorkSheet.Cells(i, 4).value)
                If temparea = "-" Then
                    IndexHigh(n, i - 1) = Openprice(n, i - 2)
                Else
                    IndexHigh(n, i - 1) = Val(temparea)
                End If
                temparea = (xlWorkSheet.Cells(i, 2).value)
                If temparea = "-" Then
                    IndexClosingPrice(n, i - 1) = Openprice(n, i - 2)
                Else
                    IndexClosingPrice(n, i - 1) = Val(temparea)
                End If
            Next
        Next

        'Open Worksheet for each index then read into data arrays each excel cell required for each index

        xlApp = New Excel.Application
        xlWorkBook = xlApp.Workbooks.Open("Z:\Fullbrook A-Levels\Computer Science\computing project\Index Closing Prices.xlsx")

        For n = 2 To 4
            xlWorkSheet = xlWorkBook.Worksheets("Index" + CStr(n))
            For i = 2 To 100
                IndexDates(n, i - 1) = (xlWorkSheet.Cells(i, 1).value)
                temparea = (xlWorkSheet.Cells(i, 2).value)
                If temparea = "-" Then
                    IndexOpen(n, i - 1) = Openprice(n, i - 2)
                Else
                    IndexOpen(n, i - 1) = Val(temparea)
                End If
                temparea = (xlWorkSheet.Cells(i, 3).value)
                If temparea = "-" Then
                    IndexLow(n, i - 1) = Openprice(n, i - 2)
                Else
                    IndexLow(n, i - 1) = Val(temparea)
                End If
                temparea = (xlWorkSheet.Cells(i, 4).value)
                If temparea = "-" Then
                    IndexHigh(n, i - 1) = Openprice(n, i - 2)
                Else
                    IndexHigh(n, i - 1) = Val(temparea)
                End If
                temparea = (xlWorkSheet.Cells(i, 5).value)
                If temparea = "-" Then
                    IndexClosingPrice(n, i - 1) = Openprice(n, i - 2)
                Else
                    IndexClosingPrice(n, i - 1) = Val(temparea)
                End If
            Next
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
                    Lowprice(n, i - 1) = Openprice(n, i - 2)
                Else
                    Lowprice(n, i - 1) = Val(temparea)
                End If
                temparea = (xlWorkSheet.Cells(i, 4).value)
                If temparea = "-" Then
                    Highprice(n, i - 1) = Openprice(n, i - 2)
                Else
                    Highprice(n, i - 1) = Val(temparea)
                End If
                temparea = (xlWorkSheet.Cells(i, 5).value)
                If temparea = "-" Then
                    ClosingPrice(n, i - 1) = Openprice(n, i - 2)
                Else
                    ClosingPrice(n, i - 1) = Val(temparea)
                End If
            Next
        Next

        xpos = 1

        For i = 99 To 1 Step -1
            Me.Chart1.Series("Closing Price").Points.AddXY(xpos, IndexClosingPrice(1, i))
            xpos = xpos + 1
        Next

        xlWorkBook.Close()
        xlApp.Quit()

    End Sub
End Class
