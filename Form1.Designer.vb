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
        Me.Unzip_b = New System.Windows.Forms.Button()
        Me.csv_xls_b = New System.Windows.Forms.Button()
        Me.MacroBox = New System.Windows.Forms.CheckBox()
        Me.DXFBox = New System.Windows.Forms.CheckBox()
        Me.STPBox = New System.Windows.Forms.CheckBox()
        Me.SuspendLayout()
        '
        'Unzip_b
        '
        Me.Unzip_b.Location = New System.Drawing.Point(125, 81)
        Me.Unzip_b.Name = "Unzip_b"
        Me.Unzip_b.Size = New System.Drawing.Size(101, 51)
        Me.Unzip_b.TabIndex = 0
        Me.Unzip_b.Text = "Unpack ZIP"
        Me.Unzip_b.UseVisualStyleBackColor = True
        '
        'csv_xls_b
        '
        Me.csv_xls_b.Location = New System.Drawing.Point(267, 81)
        Me.csv_xls_b.Name = "csv_xls_b"
        Me.csv_xls_b.Size = New System.Drawing.Size(101, 51)
        Me.csv_xls_b.TabIndex = 1
        Me.csv_xls_b.Text = "csv to BOM"
        Me.csv_xls_b.UseVisualStyleBackColor = True
        '
        'MacroBox
        '
        Me.MacroBox.AutoSize = True
        Me.MacroBox.Location = New System.Drawing.Point(403, 81)
        Me.MacroBox.Name = "MacroBox"
        Me.MacroBox.Size = New System.Drawing.Size(83, 17)
        Me.MacroBox.TabIndex = 5
        Me.MacroBox.Text = "Run macros"
        Me.MacroBox.UseVisualStyleBackColor = True
        '
        'DXFBox
        '
        Me.DXFBox.AutoSize = True
        Me.DXFBox.Location = New System.Drawing.Point(403, 116)
        Me.DXFBox.Name = "DXFBox"
        Me.DXFBox.Size = New System.Drawing.Size(102, 17)
        Me.DXFBox.TabIndex = 6
        Me.DXFBox.Text = "Filter DXF <1Mb"
        Me.DXFBox.UseVisualStyleBackColor = True
        '
        'STPBox
        '
        Me.STPBox.AutoSize = True
        Me.STPBox.Location = New System.Drawing.Point(403, 153)
        Me.STPBox.Name = "STPBox"
        Me.STPBox.Size = New System.Drawing.Size(109, 17)
        Me.STPBox.TabIndex = 7
        Me.STPBox.Text = "STP files to folder"
        Me.STPBox.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(542, 251)
        Me.Controls.Add(Me.STPBox)
        Me.Controls.Add(Me.DXFBox)
        Me.Controls.Add(Me.MacroBox)
        Me.Controls.Add(Me.csv_xls_b)
        Me.Controls.Add(Me.Unzip_b)
        Me.Name = "Form1"
        Me.Text = "Unpacker"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Unzip_b As Button
    Friend WithEvents csv_xls_b As Button
    Friend WithEvents MacroBox As CheckBox
    Friend WithEvents DXFBox As CheckBox
    Friend WithEvents STPBox As CheckBox
End Class
