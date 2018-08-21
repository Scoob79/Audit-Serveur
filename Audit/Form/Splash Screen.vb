Public Class Form3

    '   █████╗ ██╗   ██╗██████╗ ██╗████████╗     ██████╗     ██╗
    '  ██╔══██╗██║   ██║██╔══██╗██║╚══██╔══╝    ██╔═████╗   ███║
    '  ███████║██║   ██║██║  ██║██║   ██║       ██║██╔██║   ╚██║
    '  ██╔══██║██║   ██║██║  ██║██║   ██║       ████╔╝██║    ██║
    '  ██║  ██║╚██████╔╝██████╔╝██║   ██║       ╚██████╔╝██╗ ██║
    '  ╚═╝  ╚═╝ ╚═════╝ ╚═════╝ ╚═╝   ╚═╝        ╚═════╝ ╚═╝ ╚═╝
    '                                                           
    '    Traitement des données collectées par le script et gestion d'alarmes sur le serveurs
    '    Copyright (C) 2018 KASPAR Olivier
    '
    '    Ce programme est un logiciel libre: vous pouvez le redistribuer
    '    et/ou le modifier selon les termes de la "GNU General Public
    '    License", tels que publiés par la "Free Software Foundation"; soit
    '    la version 2 de cette licence ou (à votre choix) toute version
    '    ultérieure.
    '
    '    Ce programme est distribué dans l'espoir qu'il sera utile, mais
    '    SANS AUCUNE GARANTIE, ni explicite ni implicite; sans même les
    '    garanties de commercialisation ou d'adaptation dans un but spécifique.
    '
    '    Se référer à la "GNU General Public License" pour plus de détails.
    '
    '    Vous devriez avoir reçu une copie de la "GNU General Public License"
    '    en même temps que ce programme; sinon, écrivez a la "Free Software
    '    Foundation, Inc., 675 Mass Ave, Cambridge, MA 02139, USA".    
    Private Sub Form3_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles Me.KeyPress
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

    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label1.Click

    End Sub
End Class