<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form remplace la méthode Dispose pour nettoyer la liste des composants.
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

    'Requise par le Concepteur Windows Form
    Private components As System.ComponentModel.IContainer

    'REMARQUE : la procédure suivante est requise par le Concepteur Windows Form
    'Elle peut être modifiée à l'aide du Concepteur Windows Form.  
    'Ne la modifiez pas à l'aide de l'éditeur de code.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.NsTheme1 = New WindowsApplication1.NSTheme()
        Me.NsLabel2 = New WindowsApplication1.NSLabel()
        Me.NsLabel1 = New WindowsApplication1.NSLabel()
        Me.NsSeperator1 = New WindowsApplication1.NSSeperator()
        Me.NsProgressBar1 = New WindowsApplication1.NSProgressBar()
        Me.NsTheme1.SuspendLayout()
        Me.SuspendLayout()
        '
        'NsTheme1
        '
        Me.NsTheme1.AccentOffset = 42
        Me.NsTheme1.BackColor = System.Drawing.Color.FromArgb(CType(CType(50, Byte), Integer), CType(CType(50, Byte), Integer), CType(CType(50, Byte), Integer))
        Me.NsTheme1.BorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.NsTheme1.Colors = New WindowsApplication1.Bloom(-1) {}
        Me.NsTheme1.Controls.Add(Me.NsLabel2)
        Me.NsTheme1.Controls.Add(Me.NsLabel1)
        Me.NsTheme1.Controls.Add(Me.NsSeperator1)
        Me.NsTheme1.Controls.Add(Me.NsProgressBar1)
        Me.NsTheme1.Customization = ""
        Me.NsTheme1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.NsTheme1.Font = New System.Drawing.Font("Verdana", 8.0!)
        Me.NsTheme1.Image = Nothing
        Me.NsTheme1.Location = New System.Drawing.Point(0, 0)
        Me.NsTheme1.Movable = True
        Me.NsTheme1.Name = "NsTheme1"
        Me.NsTheme1.NoRounding = False
        Me.NsTheme1.Sizable = True
        Me.NsTheme1.Size = New System.Drawing.Size(800, 450)
        Me.NsTheme1.SmartBounds = True
        Me.NsTheme1.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultLocation
        Me.NsTheme1.TabIndex = 1
        Me.NsTheme1.Text = "NsTheme1"
        Me.NsTheme1.TransparencyKey = System.Drawing.Color.Empty
        Me.NsTheme1.Transparent = False
        '
        'NsLabel2
        '
        Me.NsLabel2.Font = New System.Drawing.Font("Segoe UI", 10.0!, System.Drawing.FontStyle.Bold)
        Me.NsLabel2.Location = New System.Drawing.Point(714, 422)
        Me.NsLabel2.Name = "NsLabel2"
        Me.NsLabel2.Size = New System.Drawing.Size(75, 23)
        Me.NsLabel2.TabIndex = 3
        Me.NsLabel2.Text = "NsLabel2"
        Me.NsLabel2.Value1 = ""
        Me.NsLabel2.Value2 = ""
        '
        'NsLabel1
        '
        Me.NsLabel1.Font = New System.Drawing.Font("Segoe UI", 11.25!, System.Drawing.FontStyle.Bold)
        Me.NsLabel1.Location = New System.Drawing.Point(609, 423)
        Me.NsLabel1.Name = "NsLabel1"
        Me.NsLabel1.Size = New System.Drawing.Size(154, 23)
        Me.NsLabel1.TabIndex = 2
        Me.NsLabel1.Text = "NsLabel1"
        Me.NsLabel1.Value1 = "Temps "
        Me.NsLabel1.Value2 = " Ecoulé"
        '
        'NsSeperator1
        '
        Me.NsSeperator1.Location = New System.Drawing.Point(3, 415)
        Me.NsSeperator1.Name = "NsSeperator1"
        Me.NsSeperator1.Size = New System.Drawing.Size(794, 23)
        Me.NsSeperator1.TabIndex = 1
        Me.NsSeperator1.Text = "NsSeperator1"
        '
        'NsProgressBar1
        '
        Me.NsProgressBar1.Location = New System.Drawing.Point(268, 207)
        Me.NsProgressBar1.Maximum = 100
        Me.NsProgressBar1.Minimum = 0
        Me.NsProgressBar1.Name = "NsProgressBar1"
        Me.NsProgressBar1.Size = New System.Drawing.Size(244, 23)
        Me.NsProgressBar1.TabIndex = 0
        Me.NsProgressBar1.Text = "NsProgressBar1"
        Me.NsProgressBar1.Value = 0
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.NsTheme1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.NsTheme1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents NsProgressBar1 As NSProgressBar
    Friend WithEvents NsTheme1 As NSTheme
    Friend WithEvents NsLabel2 As NSLabel
    Friend WithEvents NsLabel1 As NSLabel
    Friend WithEvents NsSeperator1 As NSSeperator
End Class
