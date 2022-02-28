<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim ChartArea1 As System.Windows.Forms.DataVisualization.Charting.ChartArea = New System.Windows.Forms.DataVisualization.Charting.ChartArea()
        Dim Legend1 As System.Windows.Forms.DataVisualization.Charting.Legend = New System.Windows.Forms.DataVisualization.Charting.Legend()
        Dim Series1 As System.Windows.Forms.DataVisualization.Charting.Series = New System.Windows.Forms.DataVisualization.Charting.Series()
        Me.Chart1 = New System.Windows.Forms.DataVisualization.Charting.Chart()
        Me.ShareAverage = New System.Windows.Forms.TextBox()
        Me.FTSEAverage = New System.Windows.Forms.TextBox()
        Me.FTSEPrediction = New System.Windows.Forms.TextBox()
        Me.ShareBox1 = New System.Windows.Forms.TextBox()
        Me.ShareBox2 = New System.Windows.Forms.TextBox()
        Me.ShareBox3 = New System.Windows.Forms.TextBox()
        Me.ShareBox4 = New System.Windows.Forms.TextBox()
        Me.ShareBox5 = New System.Windows.Forms.TextBox()
        Me.FTSEBox1 = New System.Windows.Forms.TextBox()
        Me.FTSEBox2 = New System.Windows.Forms.TextBox()
        Me.FTSEBox3 = New System.Windows.Forms.TextBox()
        Me.FTSEBox4 = New System.Windows.Forms.TextBox()
        Me.FTSEBox5 = New System.Windows.Forms.TextBox()
        Me.SharePrediction = New System.Windows.Forms.TextBox()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.ListBox1 = New System.Windows.Forms.ListBox()
        Me.ShareNameText = New System.Windows.Forms.TextBox()
        Me.HyperlinkText = New System.Windows.Forms.TextBox()
        Me.SharePercentage = New System.Windows.Forms.TextBox()
        Me.FTSEPercentage = New System.Windows.Forms.TextBox()
        CType(Me.Chart1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Chart1
        '
        ChartArea1.Name = "ChartArea1"
        Me.Chart1.ChartAreas.Add(ChartArea1)
        Legend1.Name = "Legend1"
        Me.Chart1.Legends.Add(Legend1)
        Me.Chart1.Location = New System.Drawing.Point(234, 59)
        Me.Chart1.Margin = New System.Windows.Forms.Padding(2)
        Me.Chart1.Name = "Chart1"
        Series1.ChartArea = "ChartArea1"
        Series1.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line
        Series1.Legend = "Legend1"
        Series1.Name = "Closing Price"
        Me.Chart1.Series.Add(Series1)
        Me.Chart1.Size = New System.Drawing.Size(717, 372)
        Me.Chart1.TabIndex = 1
        Me.Chart1.Text = "Chart1"
        '
        'ShareAverage
        '
        Me.ShareAverage.Location = New System.Drawing.Point(77, 462)
        Me.ShareAverage.Name = "ShareAverage"
        Me.ShareAverage.Size = New System.Drawing.Size(62, 20)
        Me.ShareAverage.TabIndex = 2
        '
        'FTSEAverage
        '
        Me.FTSEAverage.Location = New System.Drawing.Point(145, 462)
        Me.FTSEAverage.Name = "FTSEAverage"
        Me.FTSEAverage.Size = New System.Drawing.Size(62, 20)
        Me.FTSEAverage.TabIndex = 3
        '
        'FTSEPrediction
        '
        Me.FTSEPrediction.Location = New System.Drawing.Point(745, 462)
        Me.FTSEPrediction.Name = "FTSEPrediction"
        Me.FTSEPrediction.Size = New System.Drawing.Size(100, 20)
        Me.FTSEPrediction.TabIndex = 4
        '
        'ShareBox1
        '
        Me.ShareBox1.Location = New System.Drawing.Point(77, 333)
        Me.ShareBox1.Name = "ShareBox1"
        Me.ShareBox1.Size = New System.Drawing.Size(62, 20)
        Me.ShareBox1.TabIndex = 5
        '
        'ShareBox2
        '
        Me.ShareBox2.Location = New System.Drawing.Point(77, 307)
        Me.ShareBox2.Name = "ShareBox2"
        Me.ShareBox2.Size = New System.Drawing.Size(62, 20)
        Me.ShareBox2.TabIndex = 6
        '
        'ShareBox3
        '
        Me.ShareBox3.Location = New System.Drawing.Point(77, 281)
        Me.ShareBox3.Name = "ShareBox3"
        Me.ShareBox3.Size = New System.Drawing.Size(62, 20)
        Me.ShareBox3.TabIndex = 7
        '
        'ShareBox4
        '
        Me.ShareBox4.Location = New System.Drawing.Point(77, 255)
        Me.ShareBox4.Name = "ShareBox4"
        Me.ShareBox4.Size = New System.Drawing.Size(62, 20)
        Me.ShareBox4.TabIndex = 8
        '
        'ShareBox5
        '
        Me.ShareBox5.Location = New System.Drawing.Point(77, 229)
        Me.ShareBox5.Name = "ShareBox5"
        Me.ShareBox5.Size = New System.Drawing.Size(62, 20)
        Me.ShareBox5.TabIndex = 9
        '
        'FTSEBox1
        '
        Me.FTSEBox1.Location = New System.Drawing.Point(145, 333)
        Me.FTSEBox1.Name = "FTSEBox1"
        Me.FTSEBox1.Size = New System.Drawing.Size(62, 20)
        Me.FTSEBox1.TabIndex = 10
        '
        'FTSEBox2
        '
        Me.FTSEBox2.Location = New System.Drawing.Point(145, 307)
        Me.FTSEBox2.Name = "FTSEBox2"
        Me.FTSEBox2.Size = New System.Drawing.Size(62, 20)
        Me.FTSEBox2.TabIndex = 11
        '
        'FTSEBox3
        '
        Me.FTSEBox3.Location = New System.Drawing.Point(145, 281)
        Me.FTSEBox3.Name = "FTSEBox3"
        Me.FTSEBox3.Size = New System.Drawing.Size(62, 20)
        Me.FTSEBox3.TabIndex = 12
        '
        'FTSEBox4
        '
        Me.FTSEBox4.Location = New System.Drawing.Point(145, 255)
        Me.FTSEBox4.Name = "FTSEBox4"
        Me.FTSEBox4.Size = New System.Drawing.Size(62, 20)
        Me.FTSEBox4.TabIndex = 13
        '
        'FTSEBox5
        '
        Me.FTSEBox5.Location = New System.Drawing.Point(145, 229)
        Me.FTSEBox5.Name = "FTSEBox5"
        Me.FTSEBox5.Size = New System.Drawing.Size(62, 20)
        Me.FTSEBox5.TabIndex = 14
        '
        'SharePrediction
        '
        Me.SharePrediction.Location = New System.Drawing.Point(851, 462)
        Me.SharePrediction.Name = "SharePrediction"
        Me.SharePrediction.Size = New System.Drawing.Size(100, 20)
        Me.SharePrediction.TabIndex = 15
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(992, 10)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(75, 23)
        Me.Button2.TabIndex = 16
        Me.Button2.Text = "Exit"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'ListBox1
        '
        Me.ListBox1.FormattingEnabled = True
        Me.ListBox1.Location = New System.Drawing.Point(77, 13)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.Size = New System.Drawing.Size(130, 173)
        Me.ListBox1.TabIndex = 17
        '
        'ShareNameText
        '
        Me.ShareNameText.Location = New System.Drawing.Point(234, 13)
        Me.ShareNameText.Name = "ShareNameText"
        Me.ShareNameText.Size = New System.Drawing.Size(183, 20)
        Me.ShareNameText.TabIndex = 18
        '
        'HyperlinkText
        '
        Me.HyperlinkText.Location = New System.Drawing.Point(423, 13)
        Me.HyperlinkText.Name = "HyperlinkText"
        Me.HyperlinkText.Size = New System.Drawing.Size(200, 20)
        Me.HyperlinkText.TabIndex = 19
        Me.HyperlinkText.Text = " "
        '
        'SharePercentage
        '
        Me.SharePercentage.Location = New System.Drawing.Point(77, 203)
        Me.SharePercentage.Name = "SharePercentage"
        Me.SharePercentage.Size = New System.Drawing.Size(62, 20)
        Me.SharePercentage.TabIndex = 20
        '
        'FTSEPercentage
        '
        Me.FTSEPercentage.Location = New System.Drawing.Point(145, 203)
        Me.FTSEPercentage.Name = "FTSEPercentage"
        Me.FTSEPercentage.Size = New System.Drawing.Size(62, 20)
        Me.FTSEPercentage.TabIndex = 21
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Menu
        Me.ClientSize = New System.Drawing.Size(1111, 609)
        Me.Controls.Add(Me.FTSEPercentage)
        Me.Controls.Add(Me.SharePercentage)
        Me.Controls.Add(Me.HyperlinkText)
        Me.Controls.Add(Me.ShareNameText)
        Me.Controls.Add(Me.ListBox1)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.SharePrediction)
        Me.Controls.Add(Me.FTSEBox5)
        Me.Controls.Add(Me.FTSEBox4)
        Me.Controls.Add(Me.FTSEBox3)
        Me.Controls.Add(Me.FTSEBox2)
        Me.Controls.Add(Me.FTSEBox1)
        Me.Controls.Add(Me.ShareBox5)
        Me.Controls.Add(Me.ShareBox4)
        Me.Controls.Add(Me.ShareBox3)
        Me.Controls.Add(Me.ShareBox2)
        Me.Controls.Add(Me.ShareBox1)
        Me.Controls.Add(Me.FTSEPrediction)
        Me.Controls.Add(Me.FTSEAverage)
        Me.Controls.Add(Me.ShareAverage)
        Me.Controls.Add(Me.Chart1)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.Name = "Form1"
        Me.Text = "Form1"
        CType(Me.Chart1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Chart1 As DataVisualization.Charting.Chart
    Friend WithEvents ShareAverage As TextBox
    Friend WithEvents FTSEAverage As TextBox
    Friend WithEvents FTSEPrediction As TextBox
    Friend WithEvents ShareBox1 As TextBox
    Friend WithEvents ShareBox2 As TextBox
    Friend WithEvents ShareBox3 As TextBox
    Friend WithEvents ShareBox4 As TextBox
    Friend WithEvents ShareBox5 As TextBox
    Friend WithEvents FTSEBox1 As TextBox
    Friend WithEvents FTSEBox2 As TextBox
    Friend WithEvents FTSEBox3 As TextBox
    Friend WithEvents FTSEBox4 As TextBox
    Friend WithEvents FTSEBox5 As TextBox
    Friend WithEvents SharePrediction As TextBox
    Friend WithEvents Button2 As Button
    Friend WithEvents ListBox1 As ListBox
    Friend WithEvents ShareNameText As TextBox
    Friend WithEvents HyperlinkText As TextBox
    Friend WithEvents SharePercentage As TextBox
    Friend WithEvents FTSEPercentage As TextBox
End Class
