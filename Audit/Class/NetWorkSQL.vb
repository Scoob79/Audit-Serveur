Imports System.IO
Imports System.Data.OleDb
Imports System.Data.SqlClient
Public Class SQL
    Public Sub Requete(Chaine As String, ChaineDeConnexion As String)
        Dim connect As New OleDbConnection(ChaineDeConnexion)
        Dim cmd As New OleDbCommand
        connect.Open()
        cmd.Connection = connect
        cmd.CommandText = Chaine
        cmd.ExecuteNonQuery()
        connect.Close()
    End Sub

    Public Sub ChargementFichier(FichierAudit As String)
        Dim Fichier As New StreamReader(FichierAudit)
        Dim Ligne As String, Lecteur As String, NS As String, Tipe As String, SysFic As String, EspLibre As String, EspTotal As String, Pourcentage As Integer
        Dim Nom As String, Description As String, Statut As String, Etat As String, CodeSortie As String, X As Integer, Data As String, ColData(12) As String
        Dim NomPoste As String, DescPoste As String, OS As String, Version As String, DateInstall As String, NumDernierSPMa As String, NumDernierSPMi As String, Fabricant As String
        Dim Model As String, Manufacturier As String, Modèle As String, TypeProc As String, NomProc As String, DescProc As String, VitesseACT As String
        Dim VitesseMAX As String, Taille As String, N As String, NomCarte As String, TypeCarte As String, MAC As String, RxVitesseMAX As String, IP(20) As String
        Dim MSR As String, DHCP As String, AddDHCP As String, DNS As String, Utilisateurs(100) As String, Groupes(100) As String, Expiration As String, MDPVieMin As String
        Dim MDPVieMax As String, MDPLongueur As String, MDPAnterieur As String, SeuilVerrou As String, DureeVerrou As String, FenObsVerrou As String, RolePoste As String, Logiciels(500) As String
        Dim Pilotes(100) As String, service(500, 5) As String, MAJ(1000, 6) As String

        ' Création de la table dans la base


        Do
            Ligne = Fichier.ReadLine
            If Ligne = "[POSTE]" Then
                Ligne = Fichier.ReadLine
                Ligne = Fichier.ReadLine
                If InStr(Ligne, "NomPoste") > 0 Then NomPoste = Ligne.Substring(InStr(Ligne, "="))
                Ligne = Fichier.ReadLine
                If InStr(Ligne, "DescPoste") > 0 Then DescPoste = Ligne.Substring(InStr(Ligne, "="))
                Ligne = Fichier.ReadLine
                If InStr(Ligne, "OS") > 0 Then OS = Ligne.Substring(InStr(Ligne, "="))
                Ligne = Fichier.ReadLine
                If InStr(Ligne, "Version") > 0 Then Version = Ligne.Substring(InStr(Ligne, "="))
                Ligne = Fichier.ReadLine
                If InStr(Ligne, "DateInstall") > 0 Then DateInstall = Ligne.Substring(InStr(Ligne, "="))
                Ligne = Fichier.ReadLine
                If InStr(Ligne, "NumDernierSPMa") > 0 Then NumDernierSPMa = Ligne.Substring(InStr(Ligne, "="))
                Ligne = Fichier.ReadLine
                If InStr(Ligne, "NumDernierSPMi") > 0 Then NumDernierSPMi = Ligne.Substring(InStr(Ligne, "="))
                Ligne = Fichier.ReadLine
                If InStr(Ligne, "Fabricant") > 0 Then Fabricant = Ligne.Substring(InStr(Ligne, "="))
                Ligne = Fichier.ReadLine
                If InStr(Ligne, "Model") > 0 Then Model = Ligne.Substring(InStr(Ligne, "="))
                Ligne = Fichier.ReadLine
                Exit Do
            End If
        Loop Until Ligne Is Nothing

        Do
            Ligne = Fichier.ReadLine
            If Ligne = "[CARTE-MERE]" Then
                Ligne = Fichier.ReadLine
                Ligne = Fichier.ReadLine
                If InStr(Ligne, "Nom") > 0 Then Nom = Ligne.Substring(InStr(Ligne, "="))
                Ligne = Fichier.ReadLine
                If InStr(Ligne, "Modèle") > 0 Then Modèle = Ligne.Substring(InStr(Ligne, "="))
                Ligne = Fichier.ReadLine
                If InStr(Ligne, "Manufacturier") > 0 Then Manufacturier = Ligne.Substring(InStr(Ligne, "="))
                Exit Do
            End If
        Loop Until Ligne Is Nothing

        Do
            Ligne = Fichier.ReadLine
            If Ligne = "[PROCESSEUR]" Then
                Ligne = Fichier.ReadLine
                Ligne = Fichier.ReadLine
                If InStr(Ligne, "TypeProc") > 0 Then TypeProc = Ligne.Substring(InStr(Ligne, "="))
                Ligne = Fichier.ReadLine
                If InStr(Ligne, "NomProc") > 0 Then NomProc = Ligne.Substring(InStr(Ligne, "="))
                Ligne = Fichier.ReadLine
                If InStr(Ligne, "DescProc") > 0 Then DescProc = Ligne.Substring(InStr(Ligne, "="))
                Ligne = Fichier.ReadLine
                If InStr(Ligne, "VitesseACT") > 0 Then VitesseACT = Ligne.Substring(InStr(Ligne, "="))
                Ligne = Fichier.ReadLine
                If InStr(Ligne, "VitesseMAX") > 0 Then VitesseMAX = Ligne.Substring(InStr(Ligne, "="))
                Exit Do
            End If
        Loop Until Ligne Is Nothing

        Do
            Ligne = Fichier.ReadLine
            If Ligne = "[MEMOIRE]" Then
                Ligne = Fichier.ReadLine
                Ligne = Fichier.ReadLine
                If InStr(Ligne, "Taille") > 0 Then Taille = Format(Val(Ligne.Substring(InStr(Ligne, "="))) / 1024, "# ### ###.00 Mo")
                Exit Do
            End If
        Loop Until Ligne Is Nothing


        N = -1
        Do
            Ligne = Fichier.ReadLine
            If Ligne = "[HDD]" Then
                Ligne = Fichier.ReadLine
                Do While Not Ligne = "[RESEAU]"
                    Ligne = Fichier.ReadLine
                    If InStr(Ligne, "Lecteur") > 0 Then Lecteur = Ligne
                    If Ligne = "[RESEAU]" Then Exit Do

                    Ligne = Fichier.ReadLine
                    If InStr(Ligne, "NS") > 0 Then NS = Ligne
                    If Ligne = "[RESEAU]" Then Exit Do

                    Ligne = Fichier.ReadLine
                    If InStr(Ligne, "Type") > 0 Then Tipe = Ligne
                    If Ligne = "[RESEAU]" Then Exit Do

                    Ligne = Fichier.ReadLine
                    If InStr(Ligne, "SysFic") > 0 Then SysFic = Ligne
                    If Ligne = "[RESEAU]" Then Exit Do

                    Ligne = Fichier.ReadLine
                    If InStr(Ligne, "EspLibre") > 0 Then EspLibre = Ligne
                    If Ligne = "[RESEAU]" Then Exit Do

                    Ligne = Fichier.ReadLine
                    If InStr(Ligne, "EspTotal") > 0 Then EspTotal = Ligne
                    If Ligne = "[RESEAU]" Then Exit Do

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
                    N += 1
                    Ligne = Fichier.ReadLine
                Loop
                Exit Do
            End If
        Loop Until Ligne Is Nothing
        N = -1

        Do
            Ligne = Fichier.ReadLine
            If InStr(Ligne, "NomCarte") <> 0 Then NomCarte = Ligne
            If InStr(Ligne, "TypeCarte") <> 0 Then TypeCarte = Ligne
            If InStr(Ligne, "Description") <> 0 Then Description = Ligne
            If InStr(Ligne, "@MAC") <> 0 Then MAC = Ligne
            If InStr(Ligne, "VitesseMAX") <> 0 Then RxVitesseMAX = Ligne
            If InStr(Ligne, "@IP") <> 0 Then
                If Ligne.Substring(4) = "" Then
                    N = -1
                Else
                    IP(N) = Ligne.Substring(5)
                End If
            End If
            If InStr(Ligne, "MSR") <> 0 Then MSR = Ligne.Substring(5)
            If InStr(Ligne, "DHCP") <> 0 And InStr(Ligne, "@") = 0 Then DHCP = Ligne.Substring(6)
            If InStr(Ligne, "@DHCP") <> 0 Then AddDHCP = Ligne.Substring(7)
            If InStr(Ligne, "@DNS") <> 0 Then DNS = Ligne.Substring(6) : N = -1

            If Ligne = "[UTILISATEURS]" Then Exit Do
        Loop Until Ligne Is Nothing

        N = 0
        Do
            Ligne = Fichier.ReadLine
            If Ligne <> "" And Ligne <> "[GROUPES]" Then N += 1 : Utilisateurs(N) = Ligne
        Loop While Not Ligne = "[GROUPES]"

        N = 0
        Do
            Ligne = Fichier.ReadLine
            If Ligne <> "" And Ligne <> "[STRATEGIE]" Then N += 1 : Groupes(N) = Ligne
        Loop While Not Ligne = "[STRATEGIE]"

        Do
            If Ligne = "[STRATEGIE]" Then
                Ligne = Fichier.ReadLine
                Ligne = Fichier.ReadLine
                If InStr(Ligne, "Expiration") > 0 Then Expiration = Ligne.Substring(InStr(Ligne, "="))
                Ligne = Fichier.ReadLine
                If InStr(Ligne, "MDPVieMin") > 0 Then MDPVieMin = Ligne.Substring(InStr(Ligne, "="))
                Ligne = Fichier.ReadLine
                If InStr(Ligne, "MDPVieMax") > 0 Then MDPVieMax = Ligne.Substring(InStr(Ligne, "="))
                Ligne = Fichier.ReadLine
                If InStr(Ligne, "MDPLongueur") > 0 Then MDPLongueur = Ligne.Substring(InStr(Ligne, "="))
                Ligne = Fichier.ReadLine
                If InStr(Ligne, "MDPAnterieur") > 0 Then MDPAnterieur = Ligne.Substring(InStr(Ligne, "="))
                Ligne = Fichier.ReadLine
                If InStr(Ligne, "SeuilVerrou") > 0 Then SeuilVerrou = Ligne.Substring(InStr(Ligne, "="))
                Ligne = Fichier.ReadLine
                If InStr(Ligne, "DureeVerrou") > 0 Then DureeVerrou = Ligne.Substring(InStr(Ligne, "="))
                Ligne = Fichier.ReadLine
                If InStr(Ligne, "FenObsVerrou") > 0 Then FenObsVerrou = Ligne.Substring(InStr(Ligne, "="))
                Ligne = Fichier.ReadLine
                If InStr(Ligne, "RolePoste") > 0 Then RolePoste = Ligne.Substring(InStr(Ligne, "="))
                Ligne = Fichier.ReadLine
                Ligne = Fichier.ReadLine
                Ligne = Fichier.ReadLine
                Exit Do
            End If
        Loop Until Ligne Is Nothing

        N = 0
        Do
            Ligne = Fichier.ReadLine
            If Ligne <> "" And Ligne <> "[PILOTES]" Then N += 1 : Logiciels(N) = Ligne
        Loop While Not Ligne = "[PILOTES]"

        N = 0
        Do
            Ligne = Fichier.ReadLine
            If Ligne <> "" And Ligne <> "[SERVICES]" Then N += 1 : Pilotes(N) = Ligne
        Loop While Not Ligne = "[SERVICES]"

        N = 0
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
                    N += 1
                    service(N, 1) = Nom.Substring(4)
                    service(N, 2) = Description.Substring(12)
                    service(N, 3) = Statut.Substring(7)
                    service(N, 4) = Etat.Substring(8)
                    service(N, 5) = CodeSortie.Substring(11))
                    Ligne = Fichier.ReadLine
                Loop
                Exit Do
            End If
        Loop Until Ligne Is Nothing

        N = 0
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
            N += 1
            MAJ(N, 1) = ColData(1)
            MAJ(N, 2) = ColData(2)
            MAJ(N, 3) = ColData(3)
            MAJ(N, 4) = ColData(4)
            MAJ(N, 5) = ColData(5)
            MAJ(N, 6) = ColData(6)
            X = 0
        Loop

    End Sub
End Class
