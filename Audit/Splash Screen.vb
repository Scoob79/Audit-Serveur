Public Class Form3

    Private Sub Form3_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles Me.KeyPress
        '        If e.KeyChar ChrW(Keys.Enter) Then Licence.Show()
        'verifie que la valeur de la touche frappée est 0 ou 1 ou backspace ou delete 
        If e.KeyChar = ChrW(99) Then Licence.Show()
    End Sub

    Private Sub Form3_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Top = (My.Computer.Screen.WorkingArea.Height \ 2) - (Me.Height \ 2)
        Me.Left = (My.Computer.Screen.WorkingArea.Width \ 2) - (Me.Width \ 2)
        Me.Refresh()
    End Sub


    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick
        Principale.Show()
        Timer2.Enabled = False
        Me.Dispose()

    End Sub
End Class