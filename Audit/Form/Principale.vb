
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

Imports System.IO
Imports System.Data.OleDb
Imports System.Data.SqlClient

Public Class Principale
    Public P As Principale
    Public N As Integer
    Public Todo As String
    Public Champ As String
    Public boucle As Boolean = False
    Public PingEC As Boolean = False
    Public Const ChaineDeConnexion As String = "Provider=microsoft.jet.oledb.4.0;Data Source=D:\Users\u165147\source\repos\Audit\Audit\BDD\BDD.mdb"
    
    Dim text1 As New TextBox, labele1 As New Label, bouton1 As New Button, RadBt1 As New RadioButton, RadBt2 As New RadioButton, RadBt3 As New RadioButton

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

    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim Config As New StreamReader("c:\temp\ConfCompar.ini"), Ligne As String
        Dim Serveur As New StreamReader("c:\temp\serveur.ini")

        Charge_Alarme()
        On Error Resume Next
        DateTimePicker1.Value = Today.AddDays(-1)
        ComboBox2.Items.Clear()
        Do
            Ligne = Config.ReadLine
            If InStr(Ligne, "{") > 0 Then ComboBox3.Items.Add(Ligne.Substring(0, InStr(Ligne, "{") - 1))
        Loop Until Ligne Is Nothing

        Do
            Ligne = Serveur.ReadLine
            ComboBox1.Items.Add(Ligne)
        Loop Until Ligne Is Nothing
        Form1.Left = (Screen.PrimaryScreen.WorkingArea.Width - Me.Size.Width) / 2
        Form1.Top = (Screen.PrimaryScreen.WorkingArea.Height - Me.Size.Height) / 2
    End Sub

    Public Sub Exec()
        ' Exécute le recherche demandée par l'utilistateur en fonction des paramètres connu dans le fichier ConfCompar.ini et en affiche le résultat dans l'onglet Comparaison
        Dim Serveur As New StreamReader("c:\temp\serveur.ini")
        Dim Nom As String, Ligne As String, Hier As Date, Res As Integer, Section As String, Cle As String, Valeur As String, Ajout As String
        On Error Resume Next
        If boucle Then Exit Sub

        If Not RadBt1.Checked And Not RadBt2.Checked And Not RadBt3.Checked Then RadBt2.Checked = True

        DataGridView5.Rows.Clear()
        DataGridView5.Columns.Clear()
        DataGridView5.Columns.Add(0, "Champ")
        DataGridView5.Rows.Add()
        DataGridView5.Refresh()
        Do
            Nom = Serveur.ReadLine
            Hier = DateTimePicker1.Value
            Section = ComboBox3.Text.Substring(0, InStr(ComboBox3.Text, "\") - 1)
            Cle = ComboBox3.Text.Substring(InStr(ComboBox3.Text, "\"))

            If My.Computer.FileSystem.FileExists("c:\temp\audit_" & Nom & "_" & Replace(Hier, "/", "") & ".txt") Then
                Dim Audit As New StreamReader("c:\temp\audit_" & Nom & "_" & Replace(Hier, "/", "") & ".txt")

                If Todo = "rech" Then ' Si on demande une recherche
                    Do
                        Do
                            Ligne = Audit.ReadLine
                            If InStr(Ligne, Section) > 0 Then Exit Do
                        Loop Until Ligne Is Nothing

                        Do
                            Ligne = Audit.ReadLine
                            If InStr(Ligne, "=") > 0 And Len(Cle) < Len(Ligne) Then
                                If Ligne.Substring(0, Len(Cle)) = Cle Then
                                    Valeur = Ligne.Substring(InStr(Ligne, "="))
                                    If RadBt1.Checked Then
                                        If Val(Valeur) < text1.Text Then
                                            Res += 1
                                            DataGridView5.Columns.Add(0, Nom)
                                            DataGridView5.Rows(0).Cells(0).Value = Cle
                                            DataGridView5.Rows(0).Cells(Res).Value = Ligne.Substring(InStr(Ligne, "="))
                                            Exit Do
                                        End If
                                    End If
                                    If RadBt2.Checked Then
                                        If Valeur = text1.Text Then
                                            Res += 1
                                            DataGridView5.Columns.Add(0, Nom)
                                            DataGridView5.Rows(0).Cells(0).Value = Cle
                                            DataGridView5.Rows(0).Cells(Res).Value = Ligne.Substring(InStr(Ligne, "="))
                                            Exit Do
                                        End If
                                    End If
                                    If RadBt3.Checked Then
                                        If Val(Valeur) > text1.Text Then
                                            Res += 1
                                            DataGridView5.Columns.Add(0, Nom)
                                            DataGridView5.Rows(0).Cells(0).Value = Cle
                                            DataGridView5.Rows(0).Cells(Res).Value = Ligne.Substring(InStr(Ligne, "="))
                                            Exit Do
                                        End If
                                    End If

                                End If
                            End If
                        Loop Until Ligne Is Nothing
                    Loop Until Ligne Is Nothing
                End If

                If Todo.Substring(0, 7) = "rechadd" Then ' Si on veut une recherche mais avec une chaine supplémentaire à la fin
                    Ajout = Replace(Todo.Substring(7), Chr(34), "")
                    Do
                        Do
                            Ligne = Audit.ReadLine
                            If InStr(Ligne, Section) > 0 Then Exit Do
                        Loop Until Ligne Is Nothing

                        Do
                            Ligne = Audit.ReadLine
                            If InStr(Ligne, "=") > 0 And Len(Cle) < Len(Ligne) Then
                                If Ligne.Substring(0, Len(Cle)) = Cle Then
                                    Valeur = Ligne.Substring(InStr(Ligne, "="))
                                    If RadBt1.Checked And Valeur <> "" Then
                                        If Val(Valeur.Substring(0, Len(Valeur) - Len(Ajout))) < text1.Text Then
                                            Res += 1
                                            DataGridView5.Columns.Add(0, Nom)
                                            DataGridView5.Rows(0).Cells(0).Value = Cle
                                            DataGridView5.Rows(0).Cells(Res).Value = Ligne.Substring(InStr(Ligne, "="))
                                            Exit Do
                                        End If
                                    End If
                                    If RadBt2.Checked And Valeur <> "" Then
                                        If Valeur = text1.Text & Ajout Then
                                            Res += 1
                                            DataGridView5.Columns.Add(0, Nom)
                                            DataGridView5.Rows(0).Cells(0).Value = Cle
                                            DataGridView5.Rows(0).Cells(Res).Value = Ligne.Substring(InStr(Ligne, "="))
                                            Exit Do
                                        End If
                                    End If
                                    If RadBt3.Checked And Valeur <> "" Then
                                        If Val(Valeur.Substring(0, Len(Valeur) - Len(Ajout))) > text1.Text Then
                                            Res += 1
                                            DataGridView5.Columns.Add(0, Nom)
                                            DataGridView5.Rows(0).Cells(0).Value = Cle
                                            DataGridView5.Rows(0).Cells(Res).Value = Ligne.Substring(InStr(Ligne, "="))
                                            Exit Do
                                        End If
                                    End If

                                End If
                            End If
                        Loop Until Ligne Is Nothing
                    Loop Until Ligne Is Nothing
                End If




            End If
        Loop Until Nom Is Nothing
        boucle = True
    End Sub



    Private Sub AProposToolStripMenuItem_Click_1(sender As Object, e As EventArgs)
    End Sub



    Private Sub LicenceToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Licence.Show()

    End Sub

    Public Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        ' Timer qui se lance au démarrage de l'application puis toutes les 60 secondes. Il gére le lancement du Thread Ping
        Dim ThreadPing As New System.Threading.Thread(
        AddressOf Ping_serveur)
        ThreadPing.Start() ' Démarrer le nouveau thread.
        'Timer1.Interval = 600000
    End Sub

    Private Sub ActiverToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    End Sub
    Public Sub Ping_serveur()
        ' Ce Thread va depuis le fichier Serveur.ini pinguer tous les serveurs et retourner le résultat dans l'onglet Surveillance
        Dim Serveur As New StreamReader("c:\temp\serveur.ini"), Ligne As String, Ping_serveur As Boolean, N As String
        Dim connect As New OleDbConnection(ChaineDeConnexion), Trouvé As Boolean, Nowdate As String, Nowheure As String
        Dim cmd As New OleDbCommand
        Dim Res As OleDbDataReader
        N = 0
        If PingEC Then Exit Sub
        PingEC = True

        Do
            Ligne = Serveur.ReadLine()
            If Ligne Is Nothing Then Exit Do

            Try
                Ping_serveur = My.Computer.Network.Ping(Ligne)
                ' vérifie si l'alarme existe
                connect.Open()
                cmd.Connection = connect
                cmd.CommandText = "SELECT Serveur.ID, Serveur.Serveur FROM Serveur WHERE ('" & Ligne & "')"
                Res = cmd.ExecuteReader()
                While Res.Read()
                    If Res.Item(1).ToString = Ligne Then Trouvé = True
                End While
                Trouvé = True
                connect.Close()
                If Trouvé Then
                    connect.Open()
                    cmd.Connection = connect
                    cmd.CommandText = "DELETE Serveur.Serveur FROM Serveur WHERE (((Serveur.Serveur)='" & Ligne & "'));"
                    cmd.ExecuteNonQuery()
                    connect.Close()
                    'Invoke(New MethodInvoker(Sub() DataGridView1.Rows.Remove(0)))

                End If
            Catch pex As System.Net.NetworkInformation.PingException
                Try ' Connexion à la base de données
                    connect.Open()
                    cmd.Connection = connect

                    ' vérifie si l'alarme existe
                    cmd.CommandText = "SELECT Serveur.serveur FROM Serveur WHERE ('" & Ligne & "')"
                    Res = cmd.ExecuteReader()
                    While Res.Read()
                        If Res.Item(0).ToString = Ligne Then Trouvé = True : Exit While
                    End While
                    connect.Close()

                    ' Si l'alarme n'existe pas on l'ajoute
                    N += 1
                    connect.Open()
                    cmd.Connection = connect
                    If Not Trouvé Then
                        cmd.Connection = connect
                        Nowdate = Format(Now, "dd/MM/yyyy")
                        Nowheure = Format(Now, "hh:mm")
                        cmd.CommandText = "INSERT INTO Serveur ( ID, Serveur, Descritpion, [Date], Heure, Niveau, [Action] ) values ('" & N & "','" & Ligne & "','Le serveur " & Ligne & " n''est pas en ligne.','" & Nowdate & "','" & Nowheure & "','Critique','Ping')"
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "INSERT INTO Archive ( ID, Serveur, Descritpion, [Date], Heure, Niveau, [Action] ) values ('" & N & "','" & Ligne & "','Le serveur " & Ligne & " n''est pas en ligne.','" & Nowdate & "','" & Nowheure & "','Critique','Ping')"
                        cmd.ExecuteNonQuery()
                    Else ' Si l'alarme existe on la met à jour
                        Nowdate = Format(Now, "dd/MM/yyyy")
                        Nowheure = Format(Now, "hh:mm")
                        cmd.CommandText = "UPDATE Serveur SET Serveur.[Date] = '" & Nowdate & "',Serveur.[Heure] = '" & Nowheure & "' WHERE (((Serveur.Serveur)='" & Ligne & "'));"
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "INSERT INTO Archive ( ID, Serveur, Descritpion, [Date], Heure, Niveau, [Action] ) values ('" & N & "','" & Ligne & "','Le serveur " & Ligne & " n''est pas en ligne.','" & Nowdate & "','" & Nowheure & "','Critique','Ping')"
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "SELECT Serveur.ID, Serveur.serveur FROM Serveur WHERE (((Serveur.Serveur)='" & Ligne & "'));"
                        Res = cmd.ExecuteReader()

                        If Res.Read Then N = Val(Res.Item(0).ToString)
                    End If
                    connect.Close()

                    ' Gestion d'erreur
                Catch ex As OleDbException
                    MsgBox("Erreur lors de l'ajout de données dans le base : " & vbCrLf & ex.Message)
                    connect.Close()
                Catch ex As InvalidOperationException
                    MsgBox("Erreur lors de la connexion à la base : " & vbCrLf & ex.Message)
                    connect.Close()
                Catch ex As Exception
                    MsgBox("Erreur lors lors de la manipulation de la base de données : " & vbCrLf & ex.Message)
                    connect.Close()
                End Try

                If Not Trouvé Then ' Si l'alarme n'existe pas on l'ajoute
                    Invoke(New MethodInvoker(Sub() DataGridView6.Rows.Add(N, "Le serveur " & Ligne & " n'est pas en ligne.", Format(Now, "dd/MM/yyyy"), Format(Now, "hh:mm"), "Critique")))
                Else
                    Invoke(New MethodInvoker(Sub() DataGridView6.Rows(N - 1).Cells(2).Value = Nowdate))
                    Invoke(New MethodInvoker(Sub() DataGridView6.Rows(N - 1).Cells(3).Value = Nowheure))
                End If
                connect.Close()
            End Try
            Trouvé = False
        Loop Until Ligne Is Nothing
        PingEC = False
        Invoke(New MethodInvoker(Sub() DataGridView6.Refresh()))
    End Sub

    Private Function Int(getString As Func(Of Integer, String)) As String
        Throw New NotImplementedException()
    End Function

    Private Sub NsTheme1_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub NsButton1_Click(sender As Object, e As EventArgs) Handles NsButton1.Click
        Form2.Show()

    End Sub

    Private Sub NsButton2_Click(sender As Object, e As EventArgs) Handles NsButton2.Click
        Licence.Show()

    End Sub

    Private Sub SurveillanceToolStripMenuItem_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub NsOnOffBox1_CheckedChanged(sender As Object) Handles NsOnOffBox1.CheckedChanged
        If NsOnOffBox1.Checked Then
            Timer1.Enabled = False
        Else
            Timer1.Enabled = True
        End If

    End Sub

    Private Sub NsButton3_Click(sender As Object, e As EventArgs) Handles NsButton3.Click
        End
    End Sub

    Private Sub NsTheme1_Click_1(sender As Object, e As EventArgs) Handles NsTheme1.Click

    End Sub

    Private Sub ComboBox2_SelectedIndexChanged_1(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        ' En fonction du fichier sélectionné lit le fichier un complète tous les champs de l'onglet Audit
        Dim Fichier As New StreamReader(ComboBox2.Text)
        Dim Ligne As String, Lecteur As String, NS As String, Tipe As String, SysFic As String, EspLibre As String, EspTotal As String, Pourcentage As Integer
        Dim Nom As String, Description As String, Statut As String, Etat As String, CodeSortie As String, X As Integer, Data As String, ColData(12) As String
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
                Ligne = Fichier.ReadLine
                TextBox17.Text = Format(Val(Ligne.Substring(InStr(Ligne, "="))) / 1024, "# ### ###.00 Mo")
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
                    If Val(EspLibre.Substring(InStr(EspLibre, "="))) > 1024 Then
                        EspLibre = Format(Val(EspLibre.Substring(InStr(EspLibre, "="))) / 1024, "# ### ###.00 Go")
                    Else
                        EspLibre = EspLibre.Substring(InStr(EspLibre, "=")) & " Mo"
                    End If
                    If Val(EspTotal.Substring(InStr(EspTotal, "="))) > 1024 Then
                        EspTotal = Format(Val(EspTotal.Substring(InStr(EspTotal, "="))) / 1024, "# ### ###.00 Go")
                    Else
                        EspTotal = EspTotal.Substring(InStr(EspTotal, "=")) & " Mo"
                    End If
                    Pourcentage = Val(EspLibre.Substring(InStr(EspLibre, "="))) / (Val(EspTotal.Substring(InStr(EspLibre, "="))) / 100)
                    DataGridView1.Rows.Add(Lecteur.Substring(InStr(Lecteur, "=")), NS.Substring(InStr(NS, "=")), Tipe.Substring(InStr(Tipe, "=")), SysFic.Substring(InStr(SysFic, "=")),
                                           EspLibre, Pourcentage, EspTotal)
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

        Ligne = Fichier.ReadLine
        Ligne = Fichier.ReadLine
        Do
            Data = Fichier.ReadLine
            If Data = "" Then Exit Do
            Do
                X = X + 1
                ColData(X) = Data.Substring(0, InStr(Data, "  "))
                Data = LTrim(Data.Substring(Len(ColData(X))))
            Loop While Not Len(Data) = 0
            DataGridView4.Rows.Add(ColData(1), ColData(2), ColData(3), ColData(4), ColData(5), ColData(6))
            X = 0
        Loop
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        ' En fonction de l'option demandé charge les contrôle prédéfinit dans le fihcier ConfCompar.ini
        Dim Fichier As New StreamReader("c:\temp\ConfCompar.ini"), Ligne As String, Instruction As String
        NsTabControl2.TabPages(1).Controls.Remove(labele1)
        NsTabControl2.TabPages(1).Controls.Remove(text1)
        NsTabControl2.TabPages(1).Controls.Remove(bouton1)
        NsTabControl2.TabPages(1).Controls.Remove(RadBt1)
        NsTabControl2.TabPages(1).Controls.Remove(RadBt2)
        NsTabControl2.TabPages(1).Controls.Remove(RadBt3)

        Do
            Ligne = Fichier.ReadLine()

            If Ligne <> "" Then
                If InStr(Ligne, "{") = 0 Or InStr(Ligne, "{") = 0 Or InStr(Ligne, "{") = 0 Then MsgBox("Erreur dans le fichier de configuration.", vbExclamation) : End
            End If

            Do
                If InStr(Ligne, ComboBox3.Text) > 0 Then
                    Instruction = Ligne.Substring(InStr(Ligne, "{"))
                    Champ = Instruction.Substring(0, InStr(Instruction, "|") - 1)
                    Todo = Instruction.Substring(InStr(Instruction, "|"))
                    Todo = Todo.Substring(0, Len(Todo) - 1)
                End If
                Ligne = Fichier.ReadLine()
            Loop Until Ligne Is Nothing

            If Champ = "text" Then
                NsTabControl2.TabPages(1).Controls.Add(text1)
                NsTabControl2.TabPages(1).Controls.Add(labele1)
                NsTabControl2.TabPages(1).Controls.Add(bouton1)
                NsTabControl2.TabPages(1).Controls.Add(RadBt1)
                NsTabControl2.TabPages(1).Controls.Add(RadBt2)
                NsTabControl2.TabPages(1).Controls.Add(RadBt3)
                labele1.Location = New Drawing.Point(340, 22)
                labele1.Text = "Recherche"
                text1.Location = New Drawing.Point(400, 20)
                text1.Width = 250
                bouton1.Location = New Drawing.Point(780, 20)
                bouton1.Height = 20
                bouton1.Text = "Lancer !!!"
                RadBt1.Location = New Drawing.Point(660, 20) : RadBt1.Text = "<" : RadBt1.Width = 30
                RadBt2.Location = New Drawing.Point(700, 20) : RadBt2.Text = "=" : RadBt2.Width = 30 : RadBt2.Checked = True
                RadBt3.Location = New Drawing.Point(740, 20) : RadBt3.Text = ">" : RadBt3.Width = 30
                AddHandler bouton1.Click, AddressOf Exec
                Exit Do

            End If
        Loop Until Ligne Is Nothing
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        ' Lorsque le Combo change les Combo2 est replis en concéquence (Liste des fichiers audit concernant ce serveur)
        Dim i As Integer
        ComboBox2.Items.Clear()
        For i = 0 To My.Computer.FileSystem.GetFiles("c:\temp").Count - 1
            If InStr(1, My.Computer.FileSystem.GetFiles("c:\temp").Item(i), ComboBox1.Text) <> 0 Then ComboBox2.Items.Add(My.Computer.FileSystem.GetFiles("c:\temp").Item(i))
        Next i

    End Sub

    Private Sub Principale_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        End
    End Sub

    Public Overloads Function Open()
        Dim connect As New OleDbConnection(ChaineDeConnexion)
        connect.Open()
        Return connect
    End Function
    Public Overloads Function Close()
        Dim connect As New OleDbConnection(ChaineDeConnexion)
        connect.Close()
        Return connect
    End Function

    Public Sub Charge_Alarme()
        Dim connect As New OleDbConnection(ChaineDeConnexion)
        Dim cmd As New OleDbCommand
        DataGridView6.Rows.Clear()
        Try ' Connexion à la base de données
            connect.Open()
            cmd.Connection = connect

            ' vérifie si l'alarme existe
            cmd.CommandText = "SELECT Serveur.ID, Serveur.Descritpion, Serveur.Date, Serveur.Heure, Serveur.Niveau FROM Serveur"
            Dim Res As OleDbDataReader = cmd.ExecuteReader()
            While Res.Read()
                DataGridView6.Rows.Add(Res.Item(0).ToString, Res.Item(1).ToString, Res.Item(2).ToString, Res.Item(3).ToString, Res.Item(4).ToString)
            End While
            connect.Close()
        Catch ex As Exception
            MsgBox("Une erreur s'est produite pendant le chargement de la base Alarme : " & vbCrLf & ex.Message)
        End Try
    End Sub
End Class

