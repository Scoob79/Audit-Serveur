
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


Public Class Principale
    Public P As Principale
    Public N As Integer
    Public Todo As String
    Public Champ As String
    Public boucle As Boolean = False
    Public PingEC As Boolean = False
    Public Const ChaineDeConnexion As String = "Provider=microsoft.jet.oledb.4.0;Data Source=D:\Users\u165147\source\repos\Audit\Audit\BDD\BDD.mdb"
    Declare Function GetTickCount Lib "kernel32" Alias "GetTickCount" () As Long

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
        'TODO: cette ligne de code charge les données dans la table 'BDDDataSet.ServeurLst'. Vous pouvez la déplacer ou la supprimer selon les besoins.
        Me.ServeurLstTableAdapter.Fill(Me.BDDDataSet.ServeurLst)
        'TODO: cette ligne de code charge les données dans la table 'Database1DataSet.Table'. Vous pouvez la déplacer ou la supprimer selon les besoins.
        Dim Config As New StreamReader("c:\temp\ConfCompar.ini"), Ligne As String

        Charge_Alarme()

        Dim ThreadCollect As New System.Threading.Thread(AddressOf Collect)
        ThreadCollect.Priority = Threading.ThreadPriority.Highest
        ThreadCollect.Start() ' Démarrer le nouveau thread.

        On Error Resume Next
        DateTimePicker1.Value = Today.AddDays(-1)
        ComboBox2.Items.Clear()
        Do
            Ligne = Config.ReadLine
            If InStr(Ligne, "{") > 0 Then ComboBox3.Items.Add(Ligne.Substring(0, InStr(Ligne, "{") - 1))
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
    Private Sub LicenceToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Licence.Show()
    End Sub
    Public Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        ' Timer qui se lance au démarrage de l'application puis toutes les 60 secondes. Il gére le lancement du Thread Ping
        Dim ThreadPing As New System.Threading.Thread(
        AddressOf Ping_serveur)
        ThreadPing.Start() ' Démarrer le nouveau thread.
        Timer1.Interval = 600000
    End Sub
    Public Sub Ping_serveur()
        ' Ce Thread va depuis le fichier Serveur.ini pinguer tous les serveurs et retourner le résultat dans l'onglet Surveillance
        Dim Serveur As New StreamReader("c:\temp\serveur.ini"), Ligne As String, Ping_serveur As Boolean
        Dim connect As New OleDbConnection(ChaineDeConnexion), Trouvé As Boolean, Nowdate As String, Nowheure As String
        Dim cmd As New OleDbCommand
        Dim Res As OleDbDataReader
        Dim NetSql As New SQL
        N = 0
        If PingEC Or NsOnOffBox1.Checked = False Then Exit Sub
        PingEC = True

        Do
            Ligne = Serveur.ReadLine()
            If Ligne Is Nothing Then Exit Do
            Nowdate = Format(Now, "dd/MM/yyyy")
            Nowheure = Format(Now, "hh:mm")

            Try
                Ping_serveur = My.Computer.Network.Ping(Ligne)
                ' vérifie si l'alarme existe
                connect.Open()
                cmd.Connection = connect
                cmd.CommandText = "SELECT ID, Serveur FROM Alarme WHERE Serveur='" & Ligne & "';"
                Res = cmd.ExecuteReader()
                While Res.Read()
                    If Res.Item(0).ToString = Ligne Then Trouvé = True
                End While
                connect.Close()
                If Trouvé Then
                    NetSql.Requete(ChaineDeConnexion, "DELETE Serveur FROM Alarme WHERE Serveur='" & Ligne & "';")
                    'Invoke(New MethodInvoker(Sub() DataGridView1.Rows.Remove(0)))

                End If
                NetSql.Requete(ChaineDeConnexion, "INSERT INTO Archive (ID, Serveur, Descritpion, Jours, Heure, Niveau, [Action]) VALUES ('" & N & "','" & Ligne & "','Le serveur " & Ligne & " est en ligne.','" & Nowdate & "','" & Nowheure & "','Information','Ping')")
            Catch pex As System.Net.NetworkInformation.PingException
                Try ' Connexion à la base de données
                    connect.Open()
                    cmd.Connection = connect

                    ' vérifie si l'alarme existe
                    cmd.CommandText = "SELECT serveur FROM Alarme WHERE Serveur='" & Ligne & "'"
                    Res = cmd.ExecuteReader()
                    While Res.Read()
                        If Res.Item(0).ToString = Ligne Then Trouvé = True : Exit While
                    End While
                    connect.Close()

                    ' Si l'alarme n'existe pas on l'ajoute

                    connect.Open()
                    cmd.Connection = connect
                    cmd.CommandText = "SELECT COUNT(*) FROM Alarme;"
                    Res = cmd.ExecuteReader()

                    If Res.Read Then
                        N = Val(Res.Item(0).ToString)
                    End If
                    connect.Close()

                    N += 1
                    If Not Trouvé Then
                        NetSql.Requete(ChaineDeConnexion, "INSERT INTO Alarme (ID, Serveur, Descritpion, Jours, Heure, Niveau, [Action]) VALUES ('" & N & "','" & Ligne & "','Le serveur " & Ligne & " n''est pas en ligne.','" & Nowdate & "','" & Nowheure & "','Critique','Ping')")
                        NetSql.Requete(ChaineDeConnexion, "INSERT INTO Archive (ID, Serveur, Descritpion, Jours, Heure, Niveau, [Action]) VALUES ('" & N & "','" & Ligne & "','Le serveur " & Ligne & " n''est pas en ligne.','" & Nowdate & "','" & Nowheure & "','Critique','Ping')")
                    Else ' Si l'alarme existe on la met à jour
                        NetSql.Requete(ChaineDeConnexion, "UPDATE Alarme SET Jours = '" & Nowdate & "',Heure = '" & Nowheure & "' WHERE Serveur='" & Ligne & "';")
                        NetSql.Requete(ChaineDeConnexion, "INSERT INTO Archive (ID, Serveur, Descritpion, Jours, Heure, Niveau, [Action]) VALUES ('" & N & "','" & Ligne & "','Le serveur " & Ligne & " n''est pas en ligne.','" & Nowdate & "','" & Nowheure & "','Critique', 'Ping');")

                        connect.Open()
                        cmd.Connection = connect
                        cmd.CommandText = "SELECT ID, serveur FROM Alarme WHERE Serveur='" & Ligne & "';"
                        Res = cmd.ExecuteReader()
                        If Res.Read Then N = Val(Res.Item(0).ToString)
                        connect.Close()

                    End If

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
    Private Sub NsButton1_Click(sender As Object, e As EventArgs) Handles NsButton1.Click
        Form2.Show()
    End Sub
    Private Sub NsButton2_Click(sender As Object, e As EventArgs) Handles NsButton2.Click
        Licence.Show()
    End Sub
    Private Sub NsOnOffBox1_CheckedChanged(sender As Object) Handles NsOnOffBox1.CheckedChanged
        If NsOnOffBox1.Checked Then
            Timer1.Enabled = True
        Else
            Timer1.Enabled = False
        End If
    End Sub
    Private Sub NsButton3_Click(sender As Object, e As EventArgs) Handles NsButton3.Click
        End
    End Sub
    Private Sub ComboBox2_SelectedIndexChanged_1(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        'On Error GoTo GestErr
        ' En fonction du fichier sélectionné lit le fichier un complète tous les champs de l'onglet Audit
        Dim Fichier As New StreamReader(ComboBox2.Text)
        Dim Ligne As String, Lecteur As String, NS As String, Tipe As String, SysFic As String, EspLibre As String, EspTotal As String, Pourcentage As Integer
        Dim Nom As String, Description As String, Statut As String, Etat As String, CodeSortie As String, X As Integer, Data As String, ColData(12) As String

        Dim connect As New OleDbConnection(ChaineDeConnexion)
        Dim cmd As New OleDbCommand, Res As OleDbDataReader, Serveur As OleDbDataReader, Serveur_HDD As OleDbDataReader, Serveur_Reseau As OleDbDataReader, Serveur_Services As OleDbDataReader
        Dim Serveur_MAJ As OleDbDataReader
        Try ' Connexion à la base de données
            connect.Open()
            cmd.Connection = connect
            ComboBox2.Items.Clear()
            ' Récupère l'intégralité des données contenue dans la base serveur dont le champ POSTE_NomPoste commence par <Nom du serveur>_<Date selectionnée>
            cmd.CommandText = "SELECT * FROM Serveur WHERE POSTE_NomPoste LIKE '" & ComboBox1.Text & "_" & Format(ComboBox2.Text, "ddMMyyyy") & "%'"
            Serveur = cmd.ExecuteReader()
            ' Récupère l'intégralité des données contenue dans la base Serveur_HDD dont le champ POSTE_NomPoste commence par <Nom du serveur>_<Date selectionnée>
            cmd.CommandText = "SELECT * FROM Serveur_HDD WHERE POSTE_NomPoste LIKE '" & ComboBox1.Text & "_" & Format(ComboBox2.Text, "ddMMyyyy") & "%'"
            Serveur_HDD = cmd.ExecuteReader()
            ' Récupère l'intégralité des données contenue dans la base Serveur_Reseau dont le champ POSTE_NomPoste commence par <Nom du serveur>_<Date selectionnée>
            cmd.CommandText = "SELECT * FROM Serveur_Reseau WHERE POSTE_NomPoste LIKE '" & ComboBox1.Text & "_" & Format(ComboBox2.Text, "ddMMyyyy") & "%'"
            Serveur_Reseau = cmd.ExecuteReader()
            ' Récupère l'intégralité des données contenue dans la base <Nom du serveur>_Services dont le champ POSTE_NomPoste commence par <Nom du serveur>_<Date selectionnée>
            cmd.CommandText = "SELECT * FROM " & ComboBox1.Text & "_Services WHERE POSTE_NomPoste LIKE '" & ComboBox1.Text & "_" & Format(ComboBox2.Text, "ddMMyyyy") & "%'"
            Serveur_Services = cmd.ExecuteReader()
            ' Récupère l'intégralité des données contenue dans la base <Nom du serveur>_MAJ dont le champ POSTE_NomPoste commence par <Nom du serveur>_<Date selectionnée>
            cmd.CommandText = "SELECT * FROM " & ComboBox1.Text & "_MAJ WHERE POSTE_NomPoste LIKE '" & ComboBox1.Text & "_" & Format(ComboBox2.Text, "ddMMyyyy") & "%'"
            Serveur_MAJ = cmd.ExecuteReader()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        ' ========================================== POSTE ==========================================
        TextBox1.Text = Serveur.Item(0)
        TextBox2.Text = Serveur.Item(1)
        TextBox3.Text = Serveur.Item(2)
        TextBox4.Text = Serveur.Item(3)
        NsTextBox1.Text = Serveur.Item(4)
        TextBox5.Text = Serveur.Item(5)
        TextBox6.Text = Serveur.Item(6)
        TextBox7.Text = Serveur.Item(7)
        If InStr(TextBox3.Text, "7") > 0 Then PictureBox1.Image = ImageList1.Images(Win7)
        If InStr(TextBox3.Text, "2016") > 0 Then PictureBox1.Image = ImageList1.Images(WinServeur2016)
        If InStr(TextBox3.Text, "2003") > 0 Then PictureBox1.Image = ImageList1.Images(WinServeur2003)
        If InStr(TextBox3.Text, "98") > 0 Then PictureBox1.Image = ImageList1.Images(Win98)
        If InStr(TextBox3.Text, "XP") > 0 Then PictureBox1.Image = ImageList1.Images(WinXP)
        If InStr(TextBox3.Text, "8") > 0 Then PictureBox1.Image = ImageList1.Images(Win8)
        If InStr(TextBox3.Text, "10") > 0 Then PictureBox1.Image = ImageList1.Images(Win10)
        TextBox8.Text = Serveur.Item(8)

        ' ========================================== CARTE-MERE ==========================================

        TextBox11.Text = Serveur.Item(9)
        TextBox10.Text = Serveur.Item(10)
        TextBox9.Text = Serveur.Item(11)

        ' ========================================== PROCESSEUR ==========================================

        TextBox14.Text = Serveur.Item(11)
        TextBox13.Text = Serveur.Item(12)
        TextBox12.Text = Serveur.Item(13)
        TextBox15.Text = Serveur.Item(14)
        TextBox16.Text = Serveur.Item(15)

        ' ========================================== MEMOIRE ==========================================

        TextBox17.Text = Format(Val(Serveur.Item(16)) / 1024, "# ### ###.00 Mo")

        ' ========================================== HDD ==========================================

        DataGridView1.Rows.Clear()
        N = -1

#Disable Warning BC42104 ' La variable est utilisée avant de se voir attribuer une valeur
        If Val(Serveur_HDD.Item(5)) > 1024 Then
            EspLibre = Format(Val(Serveur_HDD.Item(5)) / 1024, "# ### ###.00 Go")
        Else
                EspLibre = EspLibre.Substring(InStr(EspLibre, "=")) & " Mo"
                    End If
        If Val(Serveur_HDD.Item(6)) > 1024 Then
            EspTotal = Format(Val(EspTotal.Substring(InStr(EspTotal, "="))) / 1024, "# ### ###.00 Go")
        Else
            EspTotal = EspTotal.Substring(InStr(EspTotal, "=")) & " Mo"
                    End If
                    Pourcentage = Val(EspLibre.Substring(InStr(EspLibre, "="))) / (Val(EspTotal.Substring(InStr(EspLibre, "="))) / 100)
                    DataGridView1.Rows.Add()
                    N += 1
        If Lecteur IsNot Nothing Then DataGridView1.Rows(N).Cells(0).Value = Serveur_HDD.Item(1)
        If NS IsNot Nothing Then DataGridView1.Rows(N).Cells(1).Value = Serveur_HDD.Item(2)
            If Tipe IsNot Nothing Then DataGridView1.Rows(N).Cells(2).Value = Serveur_HDD.Item(3)
            If SysFic IsNot Nothing Then DataGridView1.Rows(N).Cells(3).Value = Serveur_HDD.Item(4)
            If EspLibre IsNot Nothing Then DataGridView1.Rows(N).Cells(4).Value = EspLibre
                    If EspLibre IsNot Nothing And EspTotal IsNot Nothing Then DataGridView1.Rows(N).Cells(5).Value = Pourcentage
                    If EspTotal IsNot Nothing Then DataGridView1.Rows(N).Cells(6).Value = EspTotal
#Enable Warning BC42104 ' La variable est utilisée avant de se voir attribuer une valeur


        ' ========================================== RESEAU ========================================== <====== J'EN SUIS LA !!!!!

        N = -1
        DataGridView2.Rows.Clear()
        Do

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

            If Ligne <> "" And Ligne <> "[GROUPES]" Then ListBox1.Items.Add(Ligne)
        Loop While Not Ligne = "[GROUPES]"

        Do

            If Ligne <> "" And Ligne <> "[STRATEGIE]" Then ListBox2.Items.Add(Ligne)
        Loop While Not Ligne = "[STRATEGIE]"

        Do
            If Ligne = "[STRATEGIE]" Then


                If InStr(Ligne, "Expiration") > 0 Then TextBox25.Text = Ligne.Substring(InStr(Ligne, "="))

                If InStr(Ligne, "MDPVieMin") > 0 Then TextBox24.Text = Ligne.Substring(InStr(Ligne, "="))

                If InStr(Ligne, "MDPVieMax") > 0 Then TextBox23.Text = Ligne.Substring(InStr(Ligne, "="))

                If InStr(Ligne, "MDPLongueur") > 0 Then TextBox22.Text = Ligne.Substring(InStr(Ligne, "="))

                If InStr(Ligne, "MDPAnterieur") > 0 Then TextBox21.Text = Ligne.Substring(InStr(Ligne, "="))

                If InStr(Ligne, "SeuilVerrou") > 0 Then TextBox20.Text = Ligne.Substring(InStr(Ligne, "="))

                If InStr(Ligne, "DureeVerrou") > 0 Then TextBox19.Text = Ligne.Substring(InStr(Ligne, "="))

                If InStr(Ligne, "FenObsVerrou") > 0 Then TextBox18.Text = Ligne.Substring(InStr(Ligne, "="))

                If InStr(Ligne, "RolePoste") > 0 Then TextBox26.Text = Ligne.Substring(InStr(Ligne, "="))



                Exit Do
            End If
        Loop Until Ligne Is Nothing

        Do

            If Ligne <> "" And Ligne <> "[PILOTES]" Then ListBox3.Items.Add(Ligne)
        Loop While Not Ligne = "[PILOTES]"

        Do

            If Ligne <> "" And Ligne <> "[SERVICES]" Then ListBox4.Items.Add(Ligne)
        Loop While Not Ligne = "[SERVICES]"

        DataGridView3.Rows.Clear()
        Do
            If Ligne = "[SERVICES]" Then

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

                Loop
                Exit Do
            End If
        Loop Until Ligne Is Nothing



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
        connect.Close()
        Exit Sub
GestErr:
        MsgBox("Le fichier Decimal donnée du serveur est incorrect. Merci Decimal le vérifier." & vbCrLf & ComboBox2.Text)
        connect.Close()
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
        Dim connect As New OleDbConnection(ChaineDeConnexion)
        Dim cmd As New OleDbCommand, Res As OleDbDataReader
        Try ' Connexion à la base de données
            connect.Open()
            cmd.Connection = connect
            ComboBox2.Items.Clear()
            ' vérifie si l'alarme existe
            cmd.CommandText = "SELECT Date_collecte FROM Serveur WHERE POSTE_NomPoste LIKE '" & ComboBox1.Text & "%'"
            Res = cmd.ExecuteReader()
            While Res.Read()
                ComboBox2.Items.Add(Res.Item(0))
            End While
            connect.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
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
            cmd.CommandText = "SELECT ID, Descritpion, Jours, Heure, Niveau FROM Alarme"
            Dim Res As OleDbDataReader = cmd.ExecuteReader()
            While Res.Read()
                DataGridView6.Rows.Add(Res.Item(0).ToString, Res.Item(1).ToString, Res.Item(2).ToString, Res.Item(3).ToString, Res.Item(4).ToString)
            End While
            connect.Close()
        Catch ex As Exception
            MsgBox("Une erreur s'est produite pendant le chargement de la base Alarme : " & vbCrLf & ex.Message)
        End Try
    End Sub
    Public Sub Collect()
        Dim NetSql As New SQL, Ligne As String, Ena(10) As String, Ret As String, RequeteQ As OleDb.OleDbDataReader
        Dim ServeurLst As New StreamReader("C:\varsoft\chksys\.enable_win.lst")
        Dim connect As New OleDbConnection(ChaineDeConnexion)
        Dim cmd As New OleDbCommand, Debut As Date, M As Integer, S As Integer
        Debut = Now
        NsProgressBar1.Maximum = 1
        NsLabel38.Value1 = "Récupération de la liste des serveurs"
        Do
            Ligne = ServeurLst.ReadLine
            If Ligne IsNot Nothing Then
                Ena = Ligne.Split(":")
                Try
                    Ret = NetSql.Requete(ChaineDeConnexion, "INSERT INTO ServeurLst (Etat, Nom, IP, DerniereFoisVu) VALUES ('" & Ena(0) & "', '" & Ena(1) & "', '" & Ena(2) & "', '" & Ena(5) & "')")
                    If Ret <> "" Then Ret = NetSql.Requete(ChaineDeConnexion, "UPDATE ServeurLst SET Etat='" & Ena(0) & "', Nom='" & Ena(1) & "', IP='" & Ena(2) & "', DerniereFoisVu='" & Ena(5) & "' WHERE Nom='" & Ena(1) & "'")
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try
                NsProgressBar1.Maximum += 1
            End If
        Loop Until Ligne Is Nothing
        NsProgressBar1.Maximum -= 1
        NsLabel38.Value1 = "Chargement de la base. Merci de patienter avant de vous servir de la comparaison"

        Try
            connect.Open()
            cmd.Connection = connect
            cmd.CommandText = "SELECT Nom FROM ServeurLst"
            RequeteQ = cmd.ExecuteReader()
            While RequeteQ.Read
                NetSql.ChargementFichier("c:\varsoft\chksys\" & RequeteQ(0).ToString() & "\winaudit.txt")
                NsProgressBar1.Value += 1
            End While
        Catch ex As Exception
            connect.Close()
        End Try
        connect.Close()

        M = 0
        M = DateDiff(DateInterval.Minute, Debut, Now)
        S = Val(DateDiff(DateInterval.Second, Debut, Now)) - (M * 60)
        NsLabel38.Value1 = "Base chargée en " & M & ":" & S & " minutes."
        Invoke(New MethodInvoker(Sub() NsProgressBar1.Visible = False))
    End Sub
End Class

