Imports Excel = Microsoft.Office.Interop.Excel
Public Class Form1
    Dim UserSelection As Integer

    Dim IndexDates(4, 21) As String
    Dim IndexOpen(4, 21) As Decimal
    Dim IndexHigh(4, 21) As Decimal
    Dim IndexLow(4, 21) As Decimal
    Dim IndexClosingPrice(4, 21) As Decimal
    Dim IndexAdjClose(4, 21) As Decimal
    Dim IndexVolume(4, 21) As Decimal

    Dim sharename(10, 10) As String
    Dim Sharedates(10, 100) As String
    Dim ClosingPrice(10, 100) As Decimal
    Dim Highprice(10, 100) As Decimal
    Dim Lowprice(10, 100) As Decimal
    Dim Openprice(10, 100) As Decimal

    Dim ShareAVG(10) As Decimal
    Dim FiveDayShareTrend(10, 5) As Decimal

    Dim FiveDayFTSETrend(1, 5) As Decimal
    Dim FTSEAVG(1) As Decimal

    Dim FTSETrend As Decimal
    Dim DJTrend As Decimal
    Dim HSTrend As Decimal
    Dim NTrend As Decimal

    Dim Prediction As Decimal
    Dim PredictedShare As Decimal

    Dim xpos As Integer
    Dim temparea As String

    'the excel commands for the excel applications & workbooks
    Dim xlApp As Excel.Application
    Dim xlWorkBook As Excel.Workbook
    Dim xlWorkSheet As Excel.Worksheet



    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load

        'Open Worksheet for each index then read into data arrays each excel cell required for each index
        'selecting the correct file & worksheet.  The top (header) line is ignored hence for i=2 to 100
        xlApp = New Excel.Application
        xlWorkBook = xlApp.Workbooks.Open("Z:\Fullbrook A-Levels\Computer Science\computing project\Index Closing Prices.xlsx")

        For n = 1 To 4
            xlWorkSheet = xlWorkBook.Worksheets("Index" + CStr(n))

            If n < 2 Then ' this is making the first index which is formatted differently work in the same loop
                For i = 2 To 21
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
                For i = 2 To 21
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

        xlWorkBook.Close()
        xlApp.Quit()

        'Open worksheet for each share and then read into data arrays each excel cell required for each share
        'The top (header) line is ignored hence for i=2 to 100

        xlApp = New Excel.Application
        xlWorkBook = xlApp.Workbooks.Open("Z:\Fullbrook A-Levels\Computer Science\computing project\database for markets.xlsx")

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

        'gives a list of the share names
        For i = 1 To 10
            For n = 1 To 2 'the second column is going to get the hyperlink for the share names
                xlWorkSheet = xlWorkBook.Worksheets("Sharename")
                sharename(i, n) = (xlWorkSheet.Cells(i, n).value)
            Next
        Next

        xlWorkBook.Close()
        xlApp.Quit()

        ' loads the options for the listbox

        For i = 1 To 10
            ListBox1.Items.Add(sharename(i, 1))
        Next



        FiveDayFTSEAVG()
        FiveDayShareAVG()

        FTSEOvernightTrend()
        DowJonesOvernightTrend()
        HangSengOvernightTrend()
        NikkeiOvernightTrend()

        FTSEPredictor()






    End Sub
    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged
        'Add 1 to the list box selection given this starts at zero
        UserSelection = ListBox1.SelectedIndex + 1

        SharePredictor()

        xpos = 1

        For i = 99 To 1 Step -1
            Me.Chart1.Series("Closing Price").Points.AddXY(xpos, ClosingPrice(UserSelection, i))
            xpos = xpos + 1
        Next

        Me.ShareNameText.Text = (sharename(UserSelection, 1))
        Me.HyperlinkText.Text = (sharename(UserSelection, 2))

        Me.SharePercentage.Text = "Share %"
        Me.FTSEPercentage.Text = "FTSE %"

        Me.ShareBox1.Text = (FiveDayShareTrend(UserSelection, 5))
        Me.ShareBox2.Text = (FiveDayShareTrend(UserSelection, 4))
        Me.ShareBox3.Text = (FiveDayShareTrend(UserSelection, 3))
        Me.ShareBox4.Text = (FiveDayShareTrend(UserSelection, 2))
        Me.ShareBox5.Text = (FiveDayShareTrend(UserSelection, 1))

        Me.FTSEBox1.Text = (FiveDayFTSETrend(1, 5))
        Me.FTSEBox2.Text = (FiveDayFTSETrend(1, 4))
        Me.FTSEBox3.Text = (FiveDayFTSETrend(1, 3))
        Me.FTSEBox4.Text = (FiveDayFTSETrend(1, 2))
        Me.FTSEBox5.Text = (FiveDayFTSETrend(1, 1))

        Me.ShareAverage.Text = (ShareAVG(UserSelection))
        Me.FTSEAverage.Text = (FTSEAVG(1))

        Me.FTSEPrediction.Text = (Prediction)
        Me.SharePrediction.Text = (PredictedShare)




    End Sub

    Function FiveDayShareAVG() As Decimal

        ' this functions is going to get the 5 day average percentage change of the chosen share
        For i = 1 To 10
            For n = 1 To 5

                FiveDayShareTrend(i, n) = ((ClosingPrice(i, n) - ClosingPrice(i, n + 1)) / ClosingPrice(i, n + 1)) * 100

                ShareAVG(i) += FiveDayShareTrend(i, n)

            Next

            ShareAVG(i) = ShareAVG(i) / 5
        Next

        Return FiveDayShareAVG
    End Function

    Function FiveDayFTSEAVG() As Decimal
        ' going to give the average percentage change of the FTSE index over the course of 5 days

        For n = 1 To 5

            FiveDayFTSETrend(1, n) = ((IndexClosingPrice(1, n) - IndexClosingPrice(1, n + 1)) / IndexClosingPrice(1, n + 1)) * 100

            FTSEAVG(1) += FiveDayFTSETrend(1, n)

        Next

        FTSEAVG(1) = FTSEAVG(1) / 5


        Return FiveDayFTSEAVG
    End Function

    Function FTSEOvernightTrend() As Decimal
        'calculates the overnight trend for the FTSE index
        FTSETrend = ((IndexClosingPrice(1, 1) - IndexClosingPrice(1, 2)) / IndexClosingPrice(1, 2)) * 100

        Return FTSEOvernightTrend
    End Function

    Function DowJonesOvernightTrend() As Decimal
        'calculates the overnight trend for the Dow Jones index
        DJTrend = ((IndexClosingPrice(2, 1) - IndexClosingPrice(2, 2)) / IndexClosingPrice(2, 2)) * 100

        Return DowJonesOvernightTrend
    End Function

    Function HangSengOvernightTrend() As Decimal
        'calculates the overnight trend for the Hang Seng index
        HSTrend = ((IndexClosingPrice(3, 1) - IndexClosingPrice(3, 2)) / IndexClosingPrice(3, 2)) * 100

        Return HangSengOvernightTrend
    End Function

    Function NikkeiOvernightTrend() As Decimal
        'calculates the overnight trend for the Nikkei index
        NTrend = ((IndexClosingPrice(4, 1) - IndexClosingPrice(4, 2)) / IndexClosingPrice(4, 2)) * 100

        Return NikkeiOvernightTrend
    End Function

    Function FTSEPredictor() As Decimal
        'calculates a prediction for the FTSE index
        Prediction = (FTSETrend + DJTrend + NTrend + HSTrend) / 4

        Return FTSEPredictor
    End Function

    Function SharePredictor() As Decimal


        'calculates a prediction for the chosen share
        PredictedShare = (ShareAVG(UserSelection) / FTSEAVG(1)) * Prediction

        Return SharePredictor
    End Function



    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        End

    End Sub

End Class
