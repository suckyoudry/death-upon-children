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
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Chart1 = New System.Windows.Forms.DataVisualization.Charting.Chart()
        Me.ShareAverage = New System.Windows.Forms.TextBox()
        Me.FTSEAverage = New System.Windows.Forms.TextBox()
        Me.FTSETrendText = New System.Windows.Forms.TextBox()
        CType(Me.Chart1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(9, 10)
        Me.Button1.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(128, 49)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Chart1
        '
        ChartArea1.Name = "ChartArea1"
        Me.Chart1.ChartAreas.Add(ChartArea1)
        Legend1.Name = "Legend1"
        Me.Chart1.Legends.Add(Legend1)
        Me.Chart1.Location = New System.Drawing.Point(163, 10)
        Me.Chart1.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
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
        Me.ShareAverage.Location = New System.Drawing.Point(9, 348)
        Me.ShareAverage.Name = "ShareAverage"
        Me.ShareAverage.Size = New System.Drawing.Size(62, 20)
        Me.ShareAverage.TabIndex = 2
        '
        'FTSEAverage
        '
        Me.FTSEAverage.Location = New System.Drawing.Point(77, 348)
        Me.FTSEAverage.Name = "FTSEAverage"
        Me.FTSEAverage.Size = New System.Drawing.Size(60, 20)
        Me.FTSEAverage.TabIndex = 3
        '
        'FTSETrendText
        '
        Me.FTSETrendText.Location = New System.Drawing.Point(579, 444)
        Me.FTSETrendText.Name = "FTSETrendText"
        Me.FTSETrendText.Size = New System.Drawing.Size(100, 20)
        Me.FTSETrendText.TabIndex = 4
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Menu
        Me.ClientSize = New System.Drawing.Size(978, 549)
        Me.Controls.Add(Me.FTSETrendText)
        Me.Controls.Add(Me.FTSEAverage)
        Me.Controls.Add(Me.ShareAverage)
        Me.Controls.Add(Me.Chart1)
        Me.Controls.Add(Me.Button1)
        Me.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.Name = "Form1"
        Me.Text = "Form1"
        CType(Me.Chart1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Button1 As Button
    Friend WithEvents Chart1 As DataVisualization.Charting.Chart
    Friend WithEvents ShareAverage As TextBox
    Friend WithEvents FTSEAverage As TextBox
    Friend WithEvents FTSETrendText As TextBox
End Class
