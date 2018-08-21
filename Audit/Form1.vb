Imports System.IO
Public Class Form1
    Public N As Integer

    Const Redhat = 0
    Const Suze = 1
    Const Win2000 = 2
    Const Ubuntu = 3
    Const Win7 = 4
    Const Win8 = 5
    Const Win10 = 6
    Const Win98 = 7
    Const WinServeur2 = 8
    Const WinServeur2003 = 9
    Const WinServeur2016 = 10
    Const WinServeur = 11
    Const WinXP = 12

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        Dim i As Integer
        For i = 0 To My.Computer.FileSystem.GetFiles("c:\temp").Count - 1
            If InStr(1, My.Computer.FileSystem.GetFiles("c:\temp").Item(i), ComboBox1.Text) <> 0 Then ComboBox2.Items.Add(My.Computer.FileSystem.GetFiles("c:\temp").Item(i))
        Next i
    End Sub
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ComboBox2.Items.Clear()

    End Sub
    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged
        Dim Fichier As New StreamReader(ComboBox2.Text)
        Dim Ligne As String, Lecteur As String, NS As String, Tipe As String, SysFic As String, EspLibre As String, EspTotal As String
        Dim Nom As String, Description As String, Statut As String, Etat As String, CodeSortie As String
        Do
            Ligne = Fichier.ReadLine
            If Ligne = "[POSTE]" Then
                Ligne = Fichier.ReadLine
                Ligne = Fichier.ReadLine
                TextBox1.Text = Ligne.Substring(InStr(Ligne, "="))
                Ligne = Fichier.ReadLine
                TextBox2.Text = Ligne.Substring(InStr(Ligne, "="))
                Ligne = Fichier.ReadLine
                TextBox3.Text = Ligne.Substring(InStr(Ligne, "="))
                Ligne = Fichier.ReadLine
                TextBox4.Text = Ligne.Substring(InStr(Ligne, "="))
                Ligne = Fichier.ReadLine
                TextBox5.Text = Ligne.Substring(InStr(Ligne, "="))
                Ligne = Fichier.ReadLine
                TextBox6.Text = Ligne.Substring(InStr(Ligne, "="))
                Ligne = Fichier.ReadLine
                TextBox7.Text = Ligne.Substring(InStr(Ligne, "="))
                Ligne = Fichier.ReadLine
                TextBox8.Text = Ligne.Substring(InStr(Ligne, "="))
                If InStr(TextBox3.Text, "7") > 0 Then PictureBox1.Image = ImageList1.Images(Win7)
                If InStr(TextBox3.Text, "2016") > 0 Then PictureBox1.Image = ImageList1.Images(WinServeur2016)
                If InStr(TextBox3.Text, "2003") > 0 Then PictureBox1.Image = ImageList1.Images(WinServeur2003)
                If InStr(TextBox3.Text, "98") > 0 Then PictureBox1.Image = ImageList1.Images(Win98)
                If InStr(TextBox3.Text, "XP") > 0 Then PictureBox1.Image = ImageList1.Images(WinXP)
                If InStr(TextBox3.Text, "8") > 0 Then PictureBox1.Image = ImageList1.Images(Win8)
                If InStr(TextBox3.Text, "10") > 0 Then PictureBox1.Image = ImageList1.Images(Win10)
                Ligne = Fichier.ReadLine
                Exit Do
            End If
        Loop Until Ligne Is Nothing

        Do
            Ligne = Fichier.ReadLine
            If Ligne = "[CARTE-MERE]" Then
                Ligne = Fichier.ReadLine
                Ligne = Fichier.ReadLine
                TextBox11.Text = Ligne.Substring(InStr(Ligne, "="))
                Ligne = Fichier.ReadLine
                TextBox10.Text = Ligne.Substring(InStr(Ligne, "="))
                Ligne = Fichier.ReadLine
                TextBox9.Text = Ligne.Substring(InStr(Ligne, "="))
                Exit Do
            End If
        Loop Until Ligne Is Nothing

        Do
            Ligne = Fichier.ReadLine
            If Ligne = "[PROCESSEUR]" Then
                Ligne = Fichier.ReadLine
                Ligne = Fichier.ReadLine
                TextBox14.Text = Ligne.Substring(InStr(Ligne, "="))
                Ligne = Fichier.ReadLine
                TextBox13.Text = Ligne.Substring(InStr(Ligne, "="))
                Ligne = Fichier.ReadLine
                TextBox12.Text = Ligne.Substring(InStr(Ligne, "="))
                Ligne = Fichier.ReadLine
                TextBox15.Text = Ligne.Substring(InStr(Ligne, "="))
                Ligne = Fichier.ReadLine
                TextBox16.Text = Ligne.Substring(InStr(Ligne, "="))
                Exit Do
            End If
        Loop Until Ligne Is Nothing

        Do
            Ligne = Fichier.ReadLine
            If Ligne = "[MEMOIRE]" Then
                Ligne = Fichier.ReadLine
                TextBox17.Text = Ligne.Substring(InStr(Ligne, "="))
                Exit Do
            End If
        Loop Until Ligne Is Nothing

        DataGridView1.Rows.Clear()
        Do
            Ligne = Fichier.ReadLine
            If Ligne = "[HDD]" Then
                Ligne = Fichier.ReadLine
                Do While Not Ligne = "[RESEAU]"
                    Lecteur = Fichier.ReadLine
                    If Lecteur = "[RESEAU]" Then Exit Do
                    NS = Fichier.ReadLine
                    If NS = "[RESEAU]" Then Exit Do
                    Tipe = Fichier.ReadLine
                    If Tipe = "[RESEAU]" Then Exit Do
                    SysFic = Fichier.ReadLine
                    If SysFic = "[RESEAU]" Then Exit Do
                    EspLibre = Fichier.ReadLine
                    If EspLibre = "[RESEAU]" Then Exit Do
                    EspTotal = Fichier.ReadLine : Ligne = EspTotal
                    If EspTotal = "[RESEAU]" Then Exit Do
                    DataGridView1.Rows.Add(Lecteur, NS, Tipe, SysFic, EspLibre, EspTotal)
                    Ligne = Fichier.ReadLine
                Loop
                Exit Do
            End If
        Loop Until Ligne Is Nothing

        N = -1
        DataGridView2.Rows.Clear()
        Do
            Ligne = Fichier.ReadLine
            If InStr(Ligne, "NomCarte") <> 0 Then DataGridView2.Rows.Add(Ligne.Substring(9)) : N = N + 1
            If InStr(Ligne, "TypeCarte") <> 0 Then DataGridView2.Rows.Item(N).Cells.Item(1).Value = Ligne.Substring(10)
            If InStr(Ligne, "Description") <> 0 Then DataGridView2.Rows.Item(N).Cells.Item(2).Value = Ligne.Substring(12)
            If InStr(Ligne, "@MAC") <> 0 Then DataGridView2.Rows.Item(N).Cells.Item(3).Value = Ligne.Substring(5)
            If InStr(Ligne, "VitesseMAX") <> 0 Then DataGridView2.Rows.Item(N).Cells.Item(4).Value = Ligne.Substring(11)
            If InStr(Ligne, "@IP") <> 0 Then
                If Ligne.Substring(4) = "" Then
                    N = -1
                Else
                    DataGridView2.Rows.Item(N).Cells.Item(5).Value = Ligne.Substring(5)
                End If
            End If
            If InStr(Ligne, "MSR") <> 0 Then DataGridView2.Rows.Item(N).Cells.Item(6).Value = Ligne.Substring(5)
            If InStr(Ligne, "DHCP") <> 0 And InStr(Ligne, "@") = 0 Then DataGridView2.Rows.Item(N).Cells.Item(7).Value = Ligne.Substring(6)
            If InStr(Ligne, "@DHCP") <> 0 Then DataGridView2.Rows.Item(N).Cells.Item(8).Value = Ligne.Substring(7)
            If InStr(Ligne, "@DNS") <> 0 Then DataGridView2.Rows.Item(N).Cells.Item(9).Value = Ligne.Substring(6) : N = -1

            If Ligne = "[UTILISATEURS]" Then Exit Do
        Loop Until Ligne Is Nothing

        ListBox1.Items.Clear()
        ListBox2.Items.Clear()
        ListBox3.Items.Clear()
        ListBox4.Items.Clear()

        Do
            Ligne = Fichier.ReadLine
            If Ligne <> "" And Ligne <> "[GROUPES]" Then ListBox1.Items.Add(Ligne)
        Loop While Not Ligne = "[GROUPES]"

        Do
            Ligne = Fichier.ReadLine
            If Ligne <> "" And Ligne <> "[STRATEGIE]" Then ListBox2.Items.Add(Ligne)
        Loop While Not Ligne = "[STRATEGIE]"

        Do
            If Ligne = "[STRATEGIE]" Then
                Ligne = Fichier.ReadLine
                Ligne = Fichier.ReadLine
                TextBox25.Text = Ligne.Substring(InStr(Ligne, "="))
                Ligne = Fichier.ReadLine
                TextBox24.Text = Ligne.Substring(InStr(Ligne, "="))
                Ligne = Fichier.ReadLine
                TextBox23.Text = Ligne.Substring(InStr(Ligne, "="))
                Ligne = Fichier.ReadLine
                TextBox22.Text = Ligne.Substring(InStr(Ligne, "="))
                Ligne = Fichier.ReadLine
                TextBox21.Text = Ligne.Substring(InStr(Ligne, "="))
                Ligne = Fichier.ReadLine
                TextBox20.Text = Ligne.Substring(InStr(Ligne, "="))
                Ligne = Fichier.ReadLine
                TextBox19.Text = Ligne.Substring(InStr(Ligne, "="))
                Ligne = Fichier.ReadLine
                TextBox18.Text = Ligne.Substring(InStr(Ligne, "="))
                Ligne = Fichier.ReadLine
                TextBox26.Text = Ligne.Substring(InStr(Ligne, "="))
                Ligne = Fichier.ReadLine
                Ligne = Fichier.ReadLine
                Ligne = Fichier.ReadLine
                Exit Do
            End If
        Loop Until Ligne Is Nothing

        Do
            Ligne = Fichier.ReadLine
            If Ligne <> "" And Ligne <> "[PILOTES]" Then ListBox3.Items.Add(Ligne)
        Loop While Not Ligne = "[PILOTES]"

        Do
            Ligne = Fichier.ReadLine
            If Ligne <> "" And Ligne <> "[SERVICES]" Then ListBox4.Items.Add(Ligne)
        Loop While Not Ligne = "[SERVICES]"

        DataGridView3.Rows.Clear()
        Do
            If Ligne = "[SERVICES]" Then
                Ligne = Fichier.ReadLine
                Do While Not Ligne = "[MAJ]"
                    Nom = Fichier.ReadLine
                    If Nom = "[MAJ]" Then Exit Do
                    Description = Fichier.ReadLine
                    If Description = "[MAJ]" Then Exit Do
                    Statut = Fichier.ReadLine
                    If Statut = "[MAJ]" Then Exit Do
                    Etat = Fichier.ReadLine
                    If Etat = "[MAJ]" Then Exit Do
                    CodeSortie = Fichier.ReadLine
                    If CodeSortie = "[MAJ]" Then Exit Do
                    DataGridView3.Rows.Add(Nom.Substring(4), Description.Substring(12), Statut.Substring(7), Etat.Substring(8), CodeSortie.Substring(11))
                    Ligne = Fichier.ReadLine
                Loop
                Exit Do
            End If
        Loop Until Ligne Is Nothing

    End Sub

    Private Sub Label23_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label23.Click

    End Sub
End Class


